/**
 * Adds a trigger to run, and will clear any old triggers with the same handler for the user
 * @param {TriggerObj} [TriggerObj = {}] - Object with keys of trigger_handler and trigger_frequency
 * @param {TriggerObj.String} [TriggerObj.trigger_handler='Import'] Function name to create the trigger. This function needs to be on the users script 
 * @param {TriggerObj.number} [TriggerObj.trigger_frequency=1] Frequency of when to run the script. Every x number of days. Defaults to every day
 */
function AddTrigger({trigger_handler = 'Import', trigger_frequency=1}={}) {
	const ss = SpreadsheetApp.getActiveSpreadsheet()
	const userTriggers = ScriptApp.getUserTriggers(ss)

	// If the trigger already exists, delete the old one
	userTriggers.forEach(x => {
		if (x.getHandlerFunction().includes(trigger_handler))
			ScriptApp.deleteTrigger(x)
	})

	ScriptApp.newTrigger(trigger_handler)
		.timeBased()
		.atHour(5)
		.nearMinute(30)
		.everyDays(trigger_frequency)
		.create()
}

// Clear all triggers for the current user
function clearTriggers() {
	const ss = SpreadsheetApp.getActiveSpreadsheet()
	ScriptApp.getUserTriggers(ss).forEach(x => ScriptApp.deleteTrigger(x))
}