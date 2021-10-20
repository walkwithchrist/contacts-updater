// Swap {a: 1, b: 2, c: 3} to {1: a, 2: b, 3: c}
function swap(dict) {
	var ret = {}
	for (var key in dict) {
		ret[dict[key]] = key
	}
	return ret
}

/**
 * encodePhotoData function
 * Purpose: to take a photo url and encode it to base64 to be used as a contact photo
 * Parameter:
 *  photo_url: The url of the photo you want to encode
 * Returns:
 *  base64 encoded photo
 */
/**
 *
 *
 * @param {String} photo_url - Url of photo to encode
 * @return {null | string} - Either Returns null if it can't encode the photo, or returns the encoded photo bytes  
 */
function encodePhotoData64(photo_url) {
	try {
		var photoBlob = UrlFetchApp.fetch(photo_url)
	} catch (err) {
		Logger.log(err)
		return null
	}

	const photoBytes = photoBlob.getBlob().getAs('image/png').getBytes()
	if (photoBytes) return Utilities.base64EncodeWebSafe(photoBytes)
	return null
}


/**
 * Returns a sheet object to interact with. Will create a sheet with the given name if it doesn't exist
 * @param {String} sheet_endpoint - Name of sheet to return, or create
 * @return {Sheet} Google apps script sheet object   
 */
function GetSheet(sheet_endpoint) {
	let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_endpoint)
	if (!sheet)
		sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheet_endpoint)
	return sheet
}

/**
 * Fills each row array to the right with selected value to match the largest row in the dataset. 
 * @param {array} range: 2d array of data 
 * @param {string} fillItem: (optional) String containg the value you want to add to fill out your array.  
 * @returns 2d array with all rows of equal length.
 */
function fillOutRange(range, fillItem = '') {

	//Get the max row length out of all rows in range.
	const maxRowLen = range.reduce((acc, cur) => Math.max(acc, cur.length), 0)

	//Fill shorter rows to match max with selecte value.
	const filled = range.map((row) => {
		const dif = maxRowLen - row.length
		if (dif > 0) row = row.concat(new Array(dif).fill(fillItem))
		return row
	})

	return filled
}


/**
 * Limits the running of the given function to only one per user
 * @param {function} func - Function to run inside the lock
 */
function userLock(func) {
	const timeoutMs = 1000  // wait 0.1 seconds before giving up

	const lock = LockService.getUserLock()
	lock.tryLock(timeoutMs)

	try {
		// Run the function if it is not currently being run by this user
		if (lock.hasLock()) func()
		else Logger.log(`Skipped execution after ${timeoutMs} miliseconds; ${Session.getActiveUser().getEmail()} ran script multiple times.`)
	}
	catch (err) {
		SpreadsheetApp.flush()
		lock.releaseLock()

		throw new Error(err)
	}
}