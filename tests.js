function doGet(e) {
	QUnit.urlParams(e.parameter)
	QUnit.config({ title: 'Unit tests for my project' })
	QUnit.load(myTests)
	return QUnit.getHtml()
}

// Imports the following functions:
// ok, equal, notEqual, deepEqual, notDeepEqual, strictEqual,
// notStrictEqual, throws, module, test, asyncTest, expect
QUnit.helpers(this)

function myTests() {
	module('Test Contact Updater')

	// Set the Active Spreadsheet to a test sheet
	SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById('1tn0b4Rv4ACCPEXr11kRVXDpmbSapHLbq5tj6f-Ussek'))

	const test_groupHelper = new GroupHelper()

	test('Export', 8, function (assert) {
		const myExporter = new Exporter(test_export_params)
		assert.equal(myExporter.source_groups, test_export_params.source_groups)
		assert.equal(myExporter.mission_group, test_export_params.mission_group)
		assert.equal(myExporter.exclude_groups, test_export_params.exclude_groups)
		assert.equal(myExporter.remove_duplicate_numbers, test_export_params.remove_duplicate_numbers)

		const exportData = myExporter.ExportContacts()
		const sheetData = myExporter.sheet.getDataRange().getValues().map(row => row.filter(value => value))
		assert.equal(sheetData.length, exportData.length)
		assert.equal(exportData[0].length, sheetData[0].length)
		assert.equal(sheetData[1].length, exportData[1].length)

		// This tests if the total number of groups were equal
		assert.equal(sheetData[2][1], exportData[2][1])
	})

	module('Import Tests')
	test('CreateImporter', 1, function (assert) {
		if (Session.getActiveUser().getEmail() == test_import_params.black_list_emails[0])
			assert.throws(function () { new Importer(test_import_params) }, Error)
		else
			assert.ok(new Importer(test_import_params))
	})

	test('Failing to Update Photos', function (assert) {
		const importer = new Importer()

		let contacts = test_groupHelper.GetMembersInGroups({ group_names: ['Test Contacts'] })
		let test_contacts_to_update_photos = contacts.map((contact, i) => {
			return { resourceName: contact.resourceName, photoBytes: test_photoBytes[i % test_photoBytes.length] }
		})

		// Test Failing contactPhotoUpdate
		let failedMatches = importer.UpdateBatchPhotos(test_contacts_to_update_photos)
		assert.ok(true)
	})

}
