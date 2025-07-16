function validateCostEstimates() {
	// Open the spreadsheet
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	// Access the "Tool database" and "Accrual Report" sheets
	const toolDbSheet = ss.getSheetByName("Tool database");
	const accrualReportSheet = ss.getSheetByName("Accrual Report");

	// Get the data range for both sheets
	const toolDbData = toolDbSheet.getDataRange().getValues();
	const accrualReportData = accrualReportSheet.getDataRange().getValues();

	// Find the column indexes for the relevant fields
	const ceNumberColToolDb = toolDbData[0].indexOf("Cost Estimate No."); // Cost Estimate No. column in Tool database
	const accrualStatusColToolDb = toolDbData[0].indexOf("Accrued?"); // Accrual Status column in Tool database

	const ceNumberColAccrual = accrualReportData[0].indexOf("CE NO."); // CE No. column in Accrual Report
	const accrualNumberColAccrual = accrualReportData[0].indexOf("Accrual Number"); // Accrual Number column in Accrual Report

	if (ceNumberColToolDb === -1 || accrualStatusColToolDb === -1 || ceNumberColAccrual === -1 || accrualNumberColAccrual === -1) {
		throw new Error("One or more required columns are missing.");
	}

	// Iterate through the "Tool database" sheet starting from the second row (assuming first row is headers)
	for (let i = 1; i < toolDbData.length; i++) {
		const ceNumberToolDb = toolDbData[i][ceNumberColToolDb]; // Get CE Number from Tool database

		if (ceNumberToolDb) {
			let accrualStatus = "NO"; // Default status is 'NO'

			// Search for the CE Number in the Accrual Report sheet
			for (let j = 1; j < accrualReportData.length; j++) {
				const ceNumberAccrual = accrualReportData[j][ceNumberColAccrual]; // Get CE Number from Accrual Report
				const accrualNumber = accrualReportData[j][accrualNumberColAccrual]; // Get Accrual Number from Accrual Report

				// If CE Number matches and Accrual Number is not empty
				if (ceNumberToolDb === ceNumberAccrual && accrualNumber) {
					accrualStatus = "YES"; // Set status to 'YES' if CE Number and Accrual Number are found
					break; // Exit loop once a match is found
				}
			}

			// Update the Accrual Status column in the Tool database
			toolDbSheet.getRange(i + 1, accrualStatusColToolDb + 1).setValue(accrualStatus); // i+1 because of header row
		}
	}
}
