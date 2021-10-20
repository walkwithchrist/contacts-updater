const DEFAULT_PERSON_FIELDS = ['names', 'addresses', 'emailAddresses', 'miscKeywords', 'names', 'phoneNumbers', 'biographies', 'photos', 'memberships']

/**
 * @class GroupHelper
 */
class GroupHelper {
	/**
   * Creates an instance of GroupHelper.
   * @param {String[]} [person_fields=DEFAULT_PERSON_FIELDS] - Array of fields to pull when pulling data from contacts
   * @memberof GroupHelper
   */
	constructor(person_fields = DEFAULT_PERSON_FIELDS) {
		if (!Array.isArray(person_fields)) throw new Error('GroupHelper person_fields should be an array of strings')

		this.person_fields = person_fields

		this.GetMembersInGroup = this.GetMembersInGroup.bind(this)
		this.GetMembersInGroups = this.GetMembersInGroups.bind(this)
		this.GetGroupMemberResourceNames = this.GetGroupMemberResourceNames.bind(this)
		this.GetAllGroupNames = this.GetAllGroupNames.bind(this)
	}


	/**
   * Gets all the members in the given groups and returns the (Person) Objects for those members
   * @param {Object} group_info - Object that contains names or group_resource_names to pull members from ex. { group_names = [], group_resource_names = [] }
   * @param {String[]} [person_fields=this.person_fields] - Array of fields to pull info for on the contacts
   * @return {Object[]} array of (Person) Objects  
   * @memberof GroupHelper
   */
	GetMembersInGroups({ group_names = [], group_resource_names = [] }, person_fields = this.person_fields) {
		group_resource_names.push(...Object.values(this.GetAllGroupNames(group_names)))

		const group_members = group_resource_names.reduce((total, group_resource_name) => {
			return total.concat(...this.GetMembersInGroup({ group_resource_name }, person_fields))
		}, [])

    
		const [formatted_members] = group_members.reduce((total, contact) => {

			// Filters contacts so duplicate People aren't pull
			if (!total[1].includes(contact.requestedResourceName)) {
				total[1].push(contact.requestedResourceName)
				total[0].push(contact.person)
			}
			return total
		}, [[], []])
		return formatted_members
	}

	/**
   * Gets all the members in the given group and returns the (Person) Objects for those members
   * @param {Object} group_info - Object that contains the group name or group_resource_name to pull members from ex. { group_name, group_resource_name }
   * @param {String[]} [person_fields=this.person_fields] - Array of fields to pull info for on the contacts
   * @return {Object[]} array of (Person) Objects  
   * @memberof GroupHelper
   */
	GetMembersInGroup({ group_name, group_resource_name }, personFields = this.person_fields) {
		const contacts = this.GetGroupMemberResourceNames({ group_name, group_resource_name })

		const group_members = []
		while (contacts && contacts.length) { //use while loop because "getBatchGet" only allows up to 50 requests at a time  
			try {
				group_members.push(...People.People.getBatchGet({
					resourceNames: contacts.splice(0, 50), personFields
				}).responses)
			} catch (err) {
				// Very rarely the getBatchGet will fail for no reason, so run it again
				Logger.log(err)
				Logger.log({ contacts, personFields })
				Utilities.sleep(1000)
        
				group_members.push(...People.People.getBatchGet({
					resourceNames: contacts.splice(0, 50), personFields
				}).responses)
			}
		}
		return group_members
	}

	/**
   *
   * Get the Resource Names (looks like: people/a123456789456213) of all of the members in a group, given the group name
   * @param {Object} group_info - Object that contains the group name or group_resource_name to pull members from ex. { group_name, group_resource_name }
   * @return {string[]} returns an a rray of all resource names  
   * @memberof GroupHelper
   */
	GetGroupMemberResourceNames({ group_name, group_resource_name }) {
		let member_resource_names = []
		if (!group_name && !group_resource_name)
			return member_resource_names

		// Get the group resource name
		group_resource_name = group_resource_name ? group_resource_name : this.GetAllGroupNames([group_name])[group_name]

		if (group_resource_name) {
			let names = {}
			try {
				names = People.ContactGroups.get(group_resource_name, {
					maxMembers: 1000,
					groupFields: 'name'
				})
			} catch (err) {
				// Very rarely the contactGroups.get will fail for no reason, so run it again
				Logger.log(err)
				Logger.log({ contacts, personFields })
				Utilities.sleep(1000)

				names = People.ContactGroups.get(group_resource_name, {
					maxMembers: 1000,
					groupFields: 'name'
				})
			}

			//get resource names of group members
			if (names.memberResourceNames != null)
				member_resource_names = names.memberResourceNames
		}

		return member_resource_names
	}

	/**
   * @param {string[]} [specific_groups=[]] groups to pull resourceNames from, if empty it will pull all from the active user
   * @param {string[]} [exclude_groups=[]] groups to specifically exclude from request
   * @return {Object} Object mapping group names to their resource Names ex {'NDBM Contacts': 'group/s81jdaas', ....}
   * @memberof GroupHelper
   */
	GetAllGroupNames(specific_groups = [], exclude_groups = []) {
		// All default groups, some are depricated, but we don't need them muddying up our data
		const default_exclude = ['chatBuddies', 'all', 'myContacts', 'friends', 'family', 'blocked', 'coworkers']

		//Pull out any specified groups from the exclude list, but if there is a duplicate exclude and source group it will be excluded
		const exclude_list = default_exclude.filter(name => !specific_groups.includes(name)).concat(...exclude_groups)
		const groups = People.ContactGroups.list({
			pageSize: 1000
		})
		// Pull all the groups, reducing them to the format of {'NDBM Contacts': 'group/s81jdaas', ....}
		const group = groups.contactGroups.reduce((total, group) => {
			if (!exclude_list.includes(group.name)) {
				if (specific_groups.length) {
					if (specific_groups.includes(group.name))
						total[group.name] = group.resourceName
				}
				else {
					total[group.name] = group.resourceName
				}
			}
			return total
		}, {})
		return group
	}

	/**
   * Gets all contacts for the given resourceNames
   * @param {string[]} [resourceNames=[]] - Array of resource names to pull contacts from
   * @param {string[]} [personFields=DEFAULT_PERSON_FIELDS] - Specific fields to pull for the contact
   * @return {Object[]} Array of (People) objects 
   * @memberof GroupHelper
   */
	GetMembersByResourceNames(resourceNames = [], personFields = DEFAULT_PERSON_FIELDS) {
		const people = People.People.getBatchGet({ resourceNames, personFields }).responses
		return people.map(person => person.person)
	}
}