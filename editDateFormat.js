function formatDates() {
	// Open the spreadsheet by its name
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");

	// Get the data range for columns F, G, and I
	var rangeF = sheet.getRange("F2:F"); // Assuming the data starts from row 2
	var rangeG = sheet.getRange("G2:G"); // Assuming the data starts from row 2
	var rangeI = sheet.getRange("I2:I"); // Assuming the data starts from row 2

	// Get the values from the columns
	var valuesF = rangeF.getValues();
	var valuesG = rangeG.getValues();
	var valuesI = rangeI.getValues();

	// Function to parse and format date string
	function formatDateString(dateStr) {
		var date = new Date(dateStr);

		// Check if the date is valid
		if (isNaN(date.getTime())) {
			return ""; // Return an empty string if the date is invalid
		}

		return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
	}

	// Process and format dates in column F
	for (var i = 0; i < valuesF.length; i++) {
		var formattedDate = formatDateString(valuesF[i][0]);
		rangeF.getCell(i + 1, 1).setValue(formattedDate);
	}

	// Process and format dates in column G
	for (var j = 0; j < valuesG.length; j++) {
		var formattedDate = formatDateString(valuesG[j][0]);
		rangeG.getCell(j + 1, 1).setValue(formattedDate);
	}

	// Process and format dates in column I
	for (var k = 0; k < valuesI.length; k++) {
		var formattedDate = formatDateString(valuesI[k][0]);
		rangeI.getCell(k + 1, 1).setValue(formattedDate);
	}
}
