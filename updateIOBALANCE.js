function updateIOBalance() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();

	// Get sheets
	var toolDatabaseSheet = ss.getSheetByName("Tool database");
	var bacUpdateSheet = ss.getSheetByName("BAC Update");

	// Get headers for Tool database and BAC Update sheets
	var toolHeaders = toolDatabaseSheet.getRange(1, 1, 1, toolDatabaseSheet.getLastColumn()).getValues()[0];
	var bacHeaders = bacUpdateSheet.getRange(1, 1, 1, bacUpdateSheet.getLastColumn()).getValues()[0];

	// Find the index of the columns based on header names
	var ioNumberColTool = toolHeaders.indexOf("IO Number") + 1;
	var ioBalanceColTool = toolHeaders.indexOf("IO BALANCE") + 1;
	var ioNumberColBAC = bacHeaders.indexOf("IO #") + 1;
	var availableColBAC = bacHeaders.indexOf("Available") + 1;

	if (ioNumberColTool == 0 || ioBalanceColTool == 0 || ioNumberColBAC == 0 || availableColBAC == 0) {
		Logger.log("Error: Could not find one or more column headers.");
		return;
	}

	// Get the data from the Tool database sheet
	var ioDataRange = toolDatabaseSheet.getRange(2, ioNumberColTool, toolDatabaseSheet.getLastRow() - 1, 1);
	var ioNumbers = ioDataRange.getValues(); // Array of IO Numbers

	// Get the IO and Balance data from BAC Update sheet
	var bacDataRange = bacUpdateSheet.getRange(2, ioNumberColBAC, bacUpdateSheet.getLastRow() - 1, availableColBAC);
	var bacData = bacDataRange.getValues(); // Array of IO Numbers and their balances from BAC Update

	// Iterate over each IO Number in the Tool database sheet
	for (var i = 0; i < ioNumbers.length; i++) {
		var ioNumber = ioNumbers[i][0];

		if (ioNumber) {
			// Ensure IO Number is not empty
			var ioBalance = findBalanceForIO(bacData, ioNumber, ioNumberColBAC, availableColBAC);

			if (ioBalance !== null) {
				// Write the balance to the "IO BALANCE" column in Tool database
				toolDatabaseSheet.getRange(i + 2, ioBalanceColTool).setValue(ioBalance);
			}
		}
	}
}

function findBalanceForIO(bacData, ioNumber, ioColIndex, balanceColIndex) {
	// Iterate over each row in the BAC Update data
	for (var i = 0; i < bacData.length; i++) {
		var bacIONumber = bacData[i][0];
		var availableBalance = bacData[i][1];

		if (bacIONumber == ioNumber) {
			return availableBalance; // Return the balance if IO Number matches
		}
	}
	return null; // Return null if no match is found
}
