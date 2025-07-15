function updateCEPaymentStatus() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");

	// Define column indices
	const costEstimateNoColIndex = getColumnIndex(sheet, "Cost Estimate No.");
	const ceBalanceColIndex = getColumnIndex(sheet, "CE BALANCE");
	const cePaymentStatusColIndex = getColumnIndex(sheet, "CE Payment Status");

	// Check if column indices are valid
	if (costEstimateNoColIndex < 1 || ceBalanceColIndex < 1 || cePaymentStatusColIndex < 1) {
		Logger.log("One or more column indices are invalid. Please check your column names.");
		return;
	}

	// Get the last row with data
	const lastRow = sheet.getLastRow();

	// If no rows, exit
	if (lastRow < 2) return;

	// Get necessary columns
	const costEstimateNoCol = sheet.getRange(2, costEstimateNoColIndex, lastRow - 1).getValues(); // Cost Estimate No. column
	const ceBalanceCol = sheet.getRange(2, ceBalanceColIndex, lastRow - 1).getValues(); // CE BALANCE column
	const cePaymentStatusCol = sheet.getRange(2, cePaymentStatusColIndex, lastRow - 1).getValues(); // CE Payment Status column

	// Loop through each row of data
	for (let i = 0; i < costEstimateNoCol.length; i++) {
		const costEstimateNo = costEstimateNoCol[i][0];
		const ceBalance = ceBalanceCol[i][0]; // CE Balance
		let status = "Not Yet Fully Paid"; // Default status

		// Process rows that have a Cost Estimate No.
		if (costEstimateNo) {
			// Check if CE BALANCE is a number and not empty
			if (ceBalance !== "" && ceBalance !== null && !isNaN(ceBalance)) {
				status = ceBalance <= 0 ? "Paid" : "Not Yet Fully Paid";
			}

			// Update CE Payment Status if it's different from the existing status
			if (status !== cePaymentStatusCol[i][0]) {
				cePaymentStatusCol[i][0] = status;
			}
		}
	}

	// Set the updated values back into the sheet
	sheet.getRange(2, cePaymentStatusColIndex, cePaymentStatusCol.length).setValues(cePaymentStatusCol); // Write back to CE Payment Status column
}

function getColumnIndex(sheet, columnName) {
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	const index = headers.indexOf(columnName) + 1; // Adding 1 because index is 0-based
	if (index <= 0) {
		Logger.log(`Column "${columnName}" not found.`);
	}
	return index;
}
