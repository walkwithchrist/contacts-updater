/**
 * Exports contacts from the owner script
 * @param {Object} ExportObj - Object containing keys of source_groups, mission_groups, exclude_groups, and remove_duplicate_numbers
 * @param {Object.String[]} [ExportObj.source_groups = ['ICE', 'Static', 'IMOS Roster']] - Groups to pull contacts from
 * @param {Object.String[]} [ExportObj.exclude_groups = ['IMOS Roster']] - A group that contacts are pulled from, but one you don't want the exported contacts to use
 * @param {Object.String} [ExportObj.mission_group = 'Mission Contacts'] - Group for all exported contacts to belong to 
 * @param {Object.boolean} [ExportObj.remove_duplicate_numbers = true] - Tells the exporter if you want to keep multiple contacts with duplicate numbers
 */
function Export({
	source_groups = ['ICE', 'Static', 'IMOS Roster'],
	mission_group = 'Mission Contacts',
	exclude_groups = ['IMOS Roster'],
	remove_duplicate_numbers = true
} = {}) {

	// Run the Export, but make sure it is only running once at a time
	userLock(() => {
		const exportObj = new Exporter({
			source_groups,
			mission_group,
			exclude_groups,
			remove_duplicate_numbers
		})
		if (exportObj instanceof Exporter) exportObj.ExportContacts()
	})
}

class Exporter {

	/**
   * Creates an instance of Exporter.
   * @param {Object} ExportObj - Object containing keys of source_groups, mission_groups, exclude_groups, and remove_duplicate_numbers
   * @param {Object.String[]} [ExportObj.source_groups = ['ICE', 'Static', 'IMOS Roster']] - Groups to pull contacts from
   * @param {Object.String[]} [ExportObj.exclude_groups = ['IMOS Roster']] - A group that contacts are pulled from, but one you don't want the exported contacts to use
   * @param {Object.String} [ExportObj.mission_group = 'Mission Contacts'] - Group for all exported contacts to belong to 
   * @param {Object.boolean} [ExportObj.remove_duplicate_numbers = true] - Tells the exporter if you want to keep multiple contacts with duplicate numbers
   * @memberof Exporter
   */
	constructor({
		source_groups = ['ICE', 'Static', 'IMOS Roster'],
		mission_group = 'Mission Contacts',
		exclude_groups = ['IMOS Roster'],
		remove_duplicate_numbers = true
	} = {}) {

		if (typeof (mission_group) == 'string') this.mission_group = mission_group
		else throw new Error('Please provide a valid string for mission_group Ex \'Mission Contacts\'')

		// Groups to pull contacts from, if remove_duplicate_contacts is true, if there is a contact in the first groups of the array with the number, 
		// any future contacts found with the same number will not be added
		if (Array.isArray(source_groups)) this.source_groups = source_groups
		else throw new Error('Please provide a valid array for source_groups Ex [\'ICE\', \'Static\', \'IMOS Roster\']')

		if (Array.isArray(exclude_groups)) this.exclude_groups = exclude_groups
		else throw new Error('Please provide a valid array for exclude_groups Ex [\'IMOS Roster\']')

		this.sheet = GetSheet('Exported_Contacts')

		this.remove_duplicate_numbers = remove_duplicate_numbers

		this.group_helper = new GroupHelper()

		// dictionary of all possible groups we want the new contacts to be a part of, format {'group/1jhadg13': 'NDBM Contacts', ...}
		this.group_ids = swap(this.group_helper.GetAllGroupNames(this.source_groups, this.exclude_groups))

		// Array for all filtered contacts
		this.final_contact_list = []

		// Array containing all contacts already in the list
		this.filter_arr = []

		this.ExportContacts = this.ExportContacts.bind(this)
		this.FilterFinalContacts = this.FilterFinalContacts.bind(this)
		this.AddMiscKeywords = this.AddMiscKeywords.bind(this)
		this.FormatContactsHelper = this.FormatContactsHelper.bind(this)
	}

	/**
   * Takes all the contacts in given groups and store them in a specified spreadsheet as Json data
   * @return {Array[]} Returns all the data it added to the Exported_Contacts  
   * @memberof Exporter
   */
	ExportContacts() {
		// Pull contacts from each of the group names.
		//  If there are duplicate numbers in different contacts, 
		//  the contact in a group near the beginning of the array with the same number will override the rest
		this.source_groups.forEach(this.FilterFinalContacts)

		// Get the spreadsheet we want and clean it for new export
		this.sheet.clear()

		const exportData = []

		// Add the group names to the first row, so we don't have to filter them all again
		exportData.push(['Groups', this.mission_group, ...Object.values(this.group_ids)])

		// this.sheet.appendRow(['Groups', this.mission_group, ...groups])
		const contact_row = ['Contacts']

		// Prepare every contact to be inserted into the sheet
		this.final_contact_list.forEach(contact => {
			this.AddMiscKeywords(contact, this.group_ids)
			contact_row.push(JSON.stringify(contact))
		})
 
		exportData.push(contact_row)
		exportData.push(['Total', contact_row.length - 1])

		// Fill out the array with empty values to use setValues function 
		const final_export = fillOutRange(exportData)

		// Use setValues because it is faster than making individual api calls
		this.sheet.getRange(`R1C1:R${exportData.length}C${final_export[0].length}`).setValues(final_export)
		return exportData
	}

	/**
   * FilterFinalContacts function
   * Returns:
   *  A helper function for an array.foreach, group_name refers to the group name like 'NDBM Contacts'
   */

	/**
   * Goes through a contact group and add them to the final export list, filtering out duplicates
   * @param {String} group_name - The name of the group to pull contacts from
   * @memberof Exporter
   */
	FilterFinalContacts(group_name) {
		// Get all contacts from the group stored into the array contacts
		const contacts = this.group_helper.GetMembersInGroup({
			group_name
		}).map(contact => contact.person)

		//filter out all duplicate contacts based on resource name and phone number
		contacts.forEach(contact => {

			// Format phone numbers so they are all the same to compare
			this.FormatContacts(contact)

			// if the contact info isn't already in the array, put it in
			if (!this.filter_arr.find(this.ContactFilter(contact))) {

				// Add the contact to the final list
				this.final_contact_list.push(contact)

				//Add the resourceName and phoneNumber to the filter array
				this.filter_arr.push(contact.resourceName)

				//If it has a number, add it to the fliter arr, since we don't want two contacts with the same number
				if (this.remove_duplicate_numbers && contact.phoneNumbers && contact.phoneNumbers.length)
					contact.phoneNumbers.forEach(numberObj => this.filter_arr.push(numberObj.value))
			}
		})
	}

	/**
   * ContactFilter function
   * Purpose: returns an Array.filter helper function to filter out contacts with the same resource name or phone number
   * Parameters:
   *  contact: the contact we are comparing
   * Returns:
   *   A function that takes in a value from an array, and compares it to the contacts value
   */

	/**
   * Filters contacts to make sure we don't add duplicates
   * @param {Object} contact - A Contact to compare with previously added contacts to add or not
   * @return {boolean} Returns true if the contact has already been added  
   * @memberof Exporter
   */
	ContactFilter(contact) {
		return (res) => {
			//If the contacts resource name and/or number has already been added return true, meaning we found it
			if ((res == contact.resourceName) || (this.remove_duplicate_numbers && contact.phoneNumbers && contact.phoneNumbers.find(personNum => personNum.value == res)))
				return true
			return false
		}
	}

	/**
   * Stores important metadata on contacts, under the miscKeywords section
   * @param {Object} contact - An Object (Person) to add the miscKeywords to
   * @memberof Exporter
   */
	AddMiscKeywords(contact) {
		//The Format is weird because you can't use any type you want on miscKeywords
		//  Outlook_keyword references photos
		//  Outlook_user is the contact resource name
		//  Outlook_subject is groups the contact belongs to
		const miscKeywordsDefault = [{
			type: 'OUTLOOK_KEYWORD',
			value: contact.photos[0].url
		}, {
			type: 'OUTLOOK_USER',
			value: contact.resourceName
		}]

		if (this.mission_group) miscKeywordsDefault.push({
			type: 'OUTLOOK_SUBJECT',
			value: this.mission_group
		})

		//Add any additional groups on the member to the miscKeywords
		if (contact.memberships && contact.memberships.length) {
			contact.memberships.forEach(membership => {
				if (membership.contactGroupMembership) {
					let g_resource_name = membership.contactGroupMembership.contactGroupResourceName

					//group_dict has a master list of groups we want, if its not in there, don't add it
					if (this.group_ids[g_resource_name]) {
						miscKeywordsDefault.push({
							type: 'OUTLOOK_SUBJECT',
							value: this.group_ids[g_resource_name]
						})
					}
				}
			})
		}

		contact.miscKeywords = miscKeywordsDefault

		// The import function can't use this data, so we delete it
		delete contact.memberships
		delete contact.resourceName
		delete contact.photos
	}

	/**
   * Prepares the contact data for export and remove unnecessary data
   * @param {Object} contact - The (Person) object to format
   * @memberof Exporter
   */
	FormatContacts(contact) {
		//Remove excess names and photos, import breaks if there are more than one
		if (contact.names && contact.names.length > 1) contact.names = contact.names.filter(name => name.metadata.primary)
		if (contact.photos && contact.photos.length > 1) contact.photos = contact.photos.filter(photo => photo.metadata.primary)
		if (contact.etag) delete contact.etag

		//Format all numbers to be '**********' for accurate comparisons
		if (contact.phoneNumbers && contact.phoneNumbers.length)
			contact.phoneNumbers.forEach(numberObj => {
				numberObj.value = numberObj.value.replace(/[^\d]/g, '').replace(/^.*(\d{10})$/, '$1')
			})

		this.FormatContactsHelper(contact)
	}

	/**
   * Recursive function to format the person object so createContact can work. sourceIds in metadata will cause createContact to fail.
   * @param {Object} field - An object to filter through all keys and check for bad data
   * @memberof Exporter
   */
	FormatContactsHelper(field) {
		// Loop through every key in the object
		for (let key in field) {

			//Make sure the key exists, or else errors will be thrown
			if (field.hasOwnProperty(key)) {

				// If it's an object we will need to go through its children too
				if (typeof field[key] == 'object') {

					if (key == 'metadata')
						delete field[key]
					else
						this.FormatContactsHelper(field[key])
				}
			}
		}
	}
}