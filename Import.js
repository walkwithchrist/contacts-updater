
/**
 * Imports Contacts from Exported_Contacts sheet. If given trigger information will also set up a trigger to do it automatically
 * @param {Object.String[]} [ImportObj.black_list_emails = []] Emails to stop from running the import script. We'd recommend the one that is being exported from
 * @param {TriggerObj} [TriggerObj = {}] Object with fields of trigger_handler, and trigger frequency. If left out, it won't create a trigger
 * @param {TriggerObj.String} TriggerObj.trigger_handler Name of function to create a trigger for, needs to be a function made by user, not library
 * @param {TriggerObj.number} [TriggerObj.trigger_frequency=1] How frequent you want the trigger to run, defaults to 1 which is every day 
 */
function Import({ black_list_emails = []}, {trigger_handler, trigger_frequency = 1 } = {}) {
	userLock(() => {
		const importer = new Importer({ black_list_emails} , {trigger_handler, trigger_frequency})
		// If the user is in black_list_emails it won't return an importer object
		if(importer instanceof Importer) importer.Import()
	})
}

class Importer {

	/**
   * Creates an instance of Importer.
   * @param {Object.String[]} [ImportObj.black_list_emails = []] Emails to stop from running the import script. We'd recommend the one that is being exported from
   * @param {TriggerObj} [TriggerObj = {}] Object with fields of trigger_handler, and trigger frequency. If left out, it won't create a trigger
   * @param {TriggerObj.String} TriggerObj.trigger_handler Name of function to create a trigger for, needs to be a function made by user, not library
   * @param {TriggerObj.number} [TriggerObj.trigger_frequency=1] How frequent you want the trigger to run, defaults to 1 which is every day 
   * @memberof Importer
   */
	constructor({black_list_emails = []} = {}, {trigger_frequency=1, trigger_handler}={}) {
		if (Array.isArray(black_list_emails) && black_list_emails.includes(Session.getActiveUser().getEmail())) {
			return SpreadsheetApp.getUi().alert('Please run this script on another account, not ' + Session.getActiveUser().getEmail())
		}

		if (trigger_handler && typeof (trigger_handler) == 'string') AddTrigger({trigger_handler, trigger_frequency})

		this.sheet = GetSheet('Exported_Contacts')

		this.groups_to_members = { 'Delete Group': [] }
		this.group_names_to_ids = {}
		this.orig_r_names = []
		this.people_to_import = []

		this.group_helper = new GroupHelper()

		// matches - contacts that match Ex. [{old_contact:  {...},new_contact: {...}}]
		this.matches = []

		// contacts_to_remove -  All people in the groups to update that no longer match Ex. [{...}, {...}, ...]
		this.contacts_rnames_to_remove = []

		// contacts_to_update_photos -  contacts that need a photo update Ex. [{resourceName: '', photoBytes: , contact: {}}, ...]
		this.contacts_to_update_photos = []

		// Bind this to all functions so they can use class properties
		this.Import = this.Import.bind(this)
		this.PullSheetData = this.PullSheetData.bind(this)
		this.MatchHelper = this.MatchHelper.bind(this)
		this.CreateBatchContacts = this.CreateBatchContacts.bind(this)
		this.CreateContacts = this.CreateContacts.bind(this)
		this.UpdateMatches = this.UpdateMatches.bind(this)
		this.UpdateBatchPhotos = this.UpdateBatchPhotos.bind(this)
		this.UpdatePhotos = this.UpdatePhotos.bind(this)
		this.AddToGroupList = this.AddToGroupList.bind(this)
		this.FilterContactsWithPhotos = this.FilterContactsWithPhotos.bind(this)
		this.UpdateGroup = this.UpdateGroup.bind(this)
		this.RemoveGroup = this.RemoveGroup.bind(this)
		this.RemoveFailedMatches = this.RemoveFailedMatches.bind(this)
	}

	/**
   * Import contacts into the ActiveUsers email. It will update contacts, instead of adding them, 
   * if they have a matching value in person.miscKeywords.
   * If there are non matching contacts with the same groups as the imported contacts it will remove the ones that don't match
   * @memberof Importer
   */
	Import() {
		Logger.log('User Running the script ' + Session.getActiveUser().getEmail())

		this.PullSheetData()

		this.group_names_to_ids = this.group_helper.GetAllGroupNames()

		const all_contacts = this.group_helper.GetMembersInGroups({ group_names: Object.keys(this.groups_to_members) })

		// Compare old contacts vs new, if they have a matching miscKeyword of old resource names, make them a match. 
		// Updates this.matches. this.groups_to_members,this.contacts_to_update_photos
		all_contacts.forEach(this.MatchHelper)

		Logger.log({
			need_photo_update: this.contacts_to_update_photos.length, matches: this.matches.length,
			contacts_to_remove: this.groups_to_members['Delete Group'].length, people_to_create: this.people_to_import.length
		})

		// Create contacts for anyone that wasn't already on the users account
		this.contacts_to_update_photos.push(...this.CreateBatchContacts())

		// Update all the contacts that matched
		this.UpdateMatches()

		// Update Groups so all contacts are in the right group, max of updating 1000 users to one group
		Object.keys(this.groups_to_members).forEach(group_name => {
			if (this.groups_to_members[group_name].length) this.UpdateGroup(group_name)
		})

		// Delete all bad contacts if they exist. Deleting groups is much faster than batch deleting contacts
		if (this.groups_to_members['Delete Group'].length)
			this.RemoveGroup('Delete Group')

		// Update the photos, and format the contacts
		this.RemoveFailedMatches(this.UpdateBatchPhotos())
	}

	/**
   * Pull all data from Exported_Contacts and put it into class variables
   * @memberof Importer
   */
	PullSheetData() {
		const sheet_data = this.sheet.getRange(1, 1, this.sheet.getLastRow(), this.sheet.getLastColumn()).getValues()

		sheet_data.forEach(row => {
			// Exported sheet is formatted as  ['Groups', ....] etc. so split it up
			const [type, ...values] = row
			if (type == 'Groups')
				values.forEach(value => { if (value) this.groups_to_members[value] = [] })
			else if (type == 'Contacts') {
				const people = values.reduce((total, contact) => contact ? total.concat(JSON.parse(contact)) : total, [])
				this.people_to_import.push(...people)
			}
		})

		// Create an array of all the new resource names to compare old and new contacts for faster comparison
		this.people_to_import.forEach(person => {
			this.orig_r_names.push(...person.miscKeywords.reduce((matches, id) => {
				return id.type == 'OUTLOOK_USER' ? matches.concat([id.value]) : matches
			}, []))
		})

	}

	/**
   * Reduces contacts into lists of matches and contacts to remove, as well as contacts needing photos updated 
   * @param {Object} contact - A Person object to decide if it needs to be updated, created, or removed
   */
	MatchHelper(contact) {
		//If the contact doesn't have miscKeywords its old
		if (!contact.miscKeywords)
			return this.groups_to_members['Delete Group'].push(contact.resourceName)

		// Check for the old resource id, and see if it's in the incoming values
		let old_id = contact.miscKeywords.reduce((value, id) => id.type == 'OUTLOOK_USER' ? id.value : value, '')
		let newIndex = this.orig_r_names.indexOf(old_id)

		// newIndex is the index of the person in orignal_resourec_names and people
		if (newIndex >= 0) {

			// copy metadata from the contact on the user to the incoming contact
			this.CopyMetadata(contact, this.people_to_import[newIndex])
			this.people_to_import[newIndex].etag = contact.etag

			// Add the user to the list for all the groups it belongs to
			this.people_to_import[newIndex].miscKeywords.forEach(this.AddToGroupList(contact))

			// Add the match to the matches.
			this.matches.push({ old_contact: contact, new_contact: this.people_to_import[newIndex] })

			// Add info to this.contacts_to_update_photos if the contact photo needs to be updated.
			let photoToUpdate = this.PhotosToUpdate(contact, this.people_to_import[newIndex])
			if (photoToUpdate)
				this.contacts_to_update_photos.push(photoToUpdate)

			// Remove the person from these arrays, since the match was found
			this.orig_r_names.splice(newIndex, 1)
			this.people_to_import.splice(newIndex, 1)
		}
		else {
			//The match wasn't found, so list the person to be removed
			this.groups_to_members['Delete Group'].push(contact.resourceName)
		}
	}

	/**
   * Function that will take an array of People objects and create contacts for each one. It also makes sure you don't run into the limit of creating more than 90 a minute.
   * See {@link https://developers.google.com/people/api/rest/v1/people/batchCreateContacts|batchCreateContacts} for information on batchContactCreation
   * @param {Object[]} this.people_to_import Array of People Objects ready to be created
   * @returns {Object[]} Array of Objects mapping resource names to photo_urls for contacts that need photos updated
   */
	CreateBatchContacts() {
		Logger.log('Creating Contacts')

		// Array for all contacts that will need their photos updated as well
		const need_photo_update = []
		while (this.people_to_import.length) {
			// Start a timer, to make sure we don't try to make more than 90 contacts in a minute, which is the max quota per user
			const startTime = new Date().getTime()

			// Create the first 89 contacts
			need_photo_update.push(...this.CreateContacts(this.people_to_import.splice(0, 89)))

			// Check if a minute has passed, if not wait until the end of the minute
			let timeElasped = new Date().getTime() - startTime
			if (timeElasped < 61000 && this.people_to_import.length)
				Utilities.sleep(61000 - timeElasped)
		}

		return need_photo_update
	}

	/**
   * Function that will take an array of People objects and create contacts for each one. If the 
   * array its given is longer than 90 contacts it may hit a write limit of 90 contacts in a minute
   * See {@link https://developers.google.com/people/api/rest/v1/people/batchCreateContacts|batchCreateContacts} for information on batchContactCreation
   * @param {Object[]} contacts_to_create - Array of People Objects ready to be created
   * @returns {Object[]} Array of Objects mapping resource names to photo_urls for contacts that need photos updated
   */
	CreateContacts(contacts_to_create) {
		// Array to store all new contacts
		let newContacts = []
		const contacts = contacts_to_create.map(contactPerson => { return { contactPerson } })
		while (contacts.length) {
			// Create contacts, batchCreateContacts only takes 10 at a time
			const people = People.People.batchCreateContacts(
				{ contacts: contacts.splice(0, 10) },
				{ readMask: 'miscKeywords' })

			newContacts.push(...people.createdPeople)
		}
		// returns all contacts that have miscKeywords of OUTLOOK_SUBJECT which has the photo_urls 
		return newContacts.reduce(this.FilterContactsWithPhotos, [])
	}

	/**
   * A reducer function that takes a contact, and checks if it has a nondefault photo that needs to be uploaded.
   * If so, it adds it to the running total.
   * @param {Object} needs_photos reducer array that has contacts that need a photo update ex [{photoBytes: '', resourceName: ''}, ...]
   * @param {Object} contact (Person) Object to check if the photo needs updating
   * @returns {Object[]} Array of Objects with keys of resourceNames and photoBytes
   */
	FilterContactsWithPhotos(needs_photos, contact) {

		const photo_keyword = contact.person.miscKeywords.find(keyword => keyword.type == 'OUTLOOK_KEYWORD')
		// make sure the field exists
		if (contact.person.miscKeywords && photo_keyword) {

			// Turn the photo_url into encoded photo data. If its an invalid url photo_data will be null
			const photoBytes = encodePhotoData64(photo_keyword.value)

			// If there is photo data, add this to the list of contacts needing photo updates
			if (photoBytes)
				needs_photos.push({ resourceName: contact.person.resourceName, photoBytes })

			// Also, put the user into a list of all the groups they should be a part of
			contact.person.miscKeywords.forEach(this.AddToGroupList(contact.person))
		}
		return needs_photos
	}

	/**
   * Updates all current contacts with matching incoming contacts
   * @param {Object[]} matches Array of objects (Person) objects formated as [{old_contact: {...}, new_contact: {...}}, ...]
   * @param {string} updateMask mask for determining what fields to update. Ex 'names,addresses,miscKeywords'
   */
	UpdateMatches(matches = [], updateMask = 'names,addresses,emailAddresses,phoneNumbers,biographies,miscKeywords') {
		// Turn matches into an array of objects containing 200 matches each Ex. [{'person/121fdasd': {}, 'person/da7dafa32': {}, ...}, ...]
		Logger.log('Updating Contacts')
		if (!matches.length) matches = this.matches
		// Update all contacts that already exist with new info, max of 200 contacts per batchUpdate
		matches.reduce(this.SplitMatches, []).forEach(contacts => {
			if (Object.keys(contacts).length) People.People.batchUpdateContacts({ contacts, updateMask })
		})
	}

	/**
   * Reducer that splits up the matches into chunks of 200, which is the max for batchUpdateContacts
   * @param {Object[]} total An array of objects with keys of resource names to objects (Person)
   * @param {Object} match The current pair to add Ex. {old_contact: {...}, new_contact: {...}}
   * @param {number} i index in the array, used to track chunks of 200
   * @returns {Object[]} an array of objects, that contain up to 200 resourceName keys with values of people Objects
   */
	SplitMatches(total, { new_contact, old_contact }, i) {
		// If our array doesn't have the index yet, add it
		if (!total[Math.floor(i / 199)])
			total.push({})

		new_contact.resourceName = old_contact.resourceName
		new_contact.etag = old_contact.etag
		delete new_contact.metadata

		// Add the value to the array object formatted {'person/1ac3b3615': {...}}
		total[Math.floor(i / 199)][old_contact.resourceName] = new_contact
		return total
	}

	/**
   * Copies metadata fields from one object to another
   * @param {Object} copy_from object to pull data from
   * @param {Object} copy_to object to copy data into
   * @param {String[]} [fields=['names', 'addresses', 'emailAddresses', 'phoneNumbers', 'biographies']] fields on the object to copy data from
   */
	CopyMetadata(copy_from, copy_to, fields = ['names', 'addresses', 'emailAddresses', 'phoneNumbers', 'biographies']) {
		// Run through the object, for each of the field names and see if they exist on both objects, if so copy the data over
		if (typeof (copy_to) == 'object' && typeof (copy_from) == 'object') {
			fields.forEach(field_name => {
				if (copy_to[field_name] && copy_from[field_name])
					copy_to[field_name].forEach((field, i) => {
						if (copy_from[field_name][i]) field.metadata = copy_from[field_name][i].metadata
					})
			})
		} else Logger.log('Error in CopyMetadata, typeof(copy_to) == ' + typeof (copy_to) + ', typeof(copy_from) == ' + typeof (copy_from))
	}

	/**
   * Checks the current contact with the new_contact to see if the photo needs updating.
   * It will check for the miscKeyword type of OUTLOOK_KEYWORD for a photo_url
   * @param {Object} old_contact person Object with miscKeywords field on it.
   * @param {Object} new_contact person Object with miscKeywords field on it.
   * @returns {{undefined|Object}} Either undefined if it doesn't need updating or an object with resourceName and photoBytes keys
   */
	PhotosToUpdate(old_contact, new_contact) {
		// Check for both the contacts photos
		let contact_photo_list = undefined
		const new_photo_keyword = new_contact.miscKeywords ? new_contact.miscKeywords.find(key => key.type == 'OUTLOOK_KEYWORD') : {}
		const old_photo_keyword = old_contact.miscKeywords ? old_contact.miscKeywords.find(key => key.type == 'OUTLOOK_KEYWORD') : {}

		const old_photo = old_photo_keyword ? old_photo_keyword.value : undefined
		const new_photo = new_photo_keyword ? new_photo_keyword.value : undefined

		// Compare the urls, if they are not the same return an object with the resourceName and photoBytes of the photo
		if (old_photo != new_photo) {
			let photoBytes = encodePhotoData64(new_photo)
			contact_photo_list = photoBytes ? { resourceName: old_contact.resourceName, photoBytes } : undefined
		}

		return contact_photo_list
	}

	/**
   * Updates photos for contacts while not hitting the limit of updating more than 60 photos in one minute
   * @param {Object[]} this.contacts_to_update_photos Array of objects
   * @param {string} contacts_to_update_photos[].resourceName Resource Name of contact to update
   * @param {number} contacts_to_update_photos[].photoBytes photo_url thats been base64 encoded
   */
	UpdateBatchPhotos(contacts_to_update_photos) {
		Logger.log('Updating Photos')
		if (!Array.isArray(contacts_to_update_photos)) contacts_to_update_photos = this.contacts_to_update_photos

		let failedPhotos = []
		// Keep updating until all contacts are updated
		while (contacts_to_update_photos.length) {
			// Time to check if we went under 1 minute
			const startTime = new Date().getTime()

			// Update the first 58 photos
			failedPhotos.push(...this.UpdatePhotos(contacts_to_update_photos.splice(0, 58)))

			// If it hasn't been 1 minute, wait then start again
			let timeElasped = new Date().getTime() - startTime
			if (timeElasped < 63000 && contacts_to_update_photos.length)
				Utilities.sleep(63000 - timeElasped)
		}
		return failedPhotos
	}

	/**
   * Updates photos for contacts while not hitting the limit of updating more than 60 photos in one minute
   * @param {Object[]} photos_to_update Array of objects
   * @param {string} photos_to_update[].resourceName Resource Name of contact to update
   * @param {number} photos_to_update[].photoBytes photo_url thats been base64 encoded
   */
	UpdatePhotos(photos_to_update) {
		let failedPhotos = []
		//If photos_to_update is greater than 60 use the UpdateBatchPhotos function
		if (photos_to_update.length > 60)
			return this.UpdateBatchPhotos(photos_to_update)

		// for each entry, attempt to update the photo
		while (photos_to_update.length) {
			let { resourceName, photoBytes } = photos_to_update.pop()

			// It will error if you reach over the limit of updating 60 photos per user in one minute
			try {
				People.People.updateContactPhoto({ photoBytes }, resourceName)
			} catch (err) {
				//If there is an error, attempt again, with a short delay to help it
				if (err.name != 'GoogleJsonResponseException') Logger.log(err)
				Utilities.sleep(2000)
				try {
					People.People.updateContactPhoto({ photoBytes }, resourceName)
				} catch (err) {
					Logger.log('Couldn\'t Update photo for ' + resourceName)
					failedPhotos.push({ resourceName, photoBytes })
				}
			}
		}

		return failedPhotos
	}

	/**
   * Returns a forEach helper function that takes an object (Person) and adds them 
   * to all the groups they were attached to stored in miscKeywords under the type of 'OUTLOOK_SUBJECT'
   * @param {Object} contact A Person object with miscKeywords field
   * @returns {function} forEachHelper helper function that takes the miscKeywod and adds the contact to a group if it has the OUTLOOK_SUBJECT keyword
   */
	AddToGroupList(contact) {
		return keyword => {
			if (keyword.type == 'OUTLOOK_SUBJECT' && this.groups_to_members[keyword.value]) {
				this.groups_to_members[keyword.value].push(contact.resourceName)
			}
		}
	}

	/**
   * Creates a group and put all the members into it from a list of resource names.
   * @param {string} group_name the name of the group to create
   * @param {boolean} first_try A variable used to retry adding members, if the modify failed due to modifying too soon after the group is created
   */
	UpdateGroup(group_name, first_try = true) {
		let group_resource_name = this.group_names_to_ids[group_name]
		const resourceNamesToAdd = this.groups_to_members[group_name]
		if (!group_resource_name) {
			try {
				Logger.log(`Creating Group ${group_name}`)
				//Create the new group
				group_resource_name = People.ContactGroups.create({ contactGroup: { name: group_name } }).resourceName
			} catch (err) {
				Logger.log(err)
				Utilities.sleep(2000)
				group_resource_name = People.ContactGroups.create({ contactGroup: { name: group_name } }).resourceName
			} finally {
				Utilities.sleep(2000)
			}
		}

		//Add the new group to our maps
		this.group_names_to_ids[group_name] = group_resource_name

		try {
			Logger.log(`Adding contacts to ${group_name}`)
			//Add all the members to the group by their resourceNames
			if (resourceNamesToAdd.length)
				People.ContactGroups.Members.modify({ resourceNamesToAdd }, group_resource_name)
		} catch (err) {
			Logger.log(err)
			Utilities.sleep(2000)
			// If this is the first time, run it again, because floating members without groups sucks.
			if (first_try) this.UpdateGroup(group_name, !first_try)
		}
	}

	/**
   * Removes a group and all the members in it
   * @param {string} group_name the name of the group to remove
   */
	RemoveGroup(group_name) {
		Logger.log('Deleting Contacts')

		//Sleep to make sure if the group was just created it shows up, and members are added
		Utilities.sleep(1000)
		let group_resource_name = this.group_names_to_ids[group_name] || this.group_helper.GetAllGroupNames(group_name)[group_name]
		Logger.log('Deleting Group: ' + group_resource_name)
		try {
			// if the group exists remove it and its values from the resourceName maps
			if (group_resource_name) {
				People.ContactGroups.remove(group_resource_name, { deleteContacts: true })
				if (this.group_names_to_ids[group_name]) delete this.group_names_to_ids[group_name]
			}
		} catch (err) {
			Logger.log('Group Delete Failed, trying again')
			// The group exists, but it might be still adding members etc
			Utilities.sleep(2000)
			People.ContactGroups.remove(group_resource_name, { deleteContacts: true })
			if (this.group_names_to_ids[group_name]) delete this.group_names_to_ids[group_name]
		}
	}

	/**
   * Removes the miscKeywords for contacts that failed their photoUpdate
   * @param {Object[]} failedMatches - Array of objects formatted {resourceName, photoBytes}
   * @memberof Importer
   */
	RemoveFailedMatches(failedMatches) {
		if (failedMatches.length) {
			const bad_photos = this.group_helper.GetMembersByResourceNames(failedMatches.map(value => value.resourceName))
			failedMatches = []

			// Strip the miscKeywords and then formate it for updateMatches to update
			bad_photos.forEach(person => {
				person.miscKeywords = person.miscKeywords.filter(key => key.type != 'OUTLOOK_KEYWORD')
				failedMatches.push({ old_contact: person, new_contact: person })
			})

			this.UpdateMatches(failedMatches, 'miscKeywords')
		}
	}
}