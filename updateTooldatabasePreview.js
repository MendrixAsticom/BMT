function copyDataToNewSheet() {
	const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	const sourceSheet = sourceSpreadsheet.getSheetByName("Tool database");

	// Get the data and formats starting from the second row
	const sourceData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
	const sourceFormats = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getNumberFormats();

	// Replace 'YOUR_NEW_SPREADSHEET_ID' with the actual ID of the new Google Sheet
	const targetSpreadsheetId = "1xgROcPRmkI6_dLfZmadbCI0Yv_0s_D7S23PIdNnkA40";
	const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
	const targetSheet = targetSpreadsheet.getActiveSheet();

	// Clear the target sheet before updating it, but keep the first row intact
	targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();

	// Paste data into the new sheet starting from the second row
	targetSheet.getRange(2, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

	// Set the formats in the target sheet
	targetSheet.getRange(2, 1, sourceFormats.length, sourceFormats[0].length).setNumberFormats(sourceFormats);
}

function createTrigger() {
	// Create a trigger that runs every minute
	ScriptApp.newTrigger("copyDataToNewSheet").timeBased().everyMinutes(1).create();
}
