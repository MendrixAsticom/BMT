function processAccruals() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var toolDatabaseSheet = ss.getSheetByName("Tool database");
	var accrualReportSheet = ss.getSheetByName("Accrual Report");

	var lastRowToolDB = toolDatabaseSheet.getLastRow();

	// Define the column names and get their indices from the first row of Tool database
	var toolHeaders = toolDatabaseSheet.getRange(1, 1, 1, toolDatabaseSheet.getLastColumn()).getValues()[0];

	// Function to get column index safely from Tool database
	function getColumnIndex(headerName, headers) {
		var index = headers.indexOf(headerName) + 1;
		if (index === 0) {
			throw new Error("Header '" + headerName + "' not found.");
		}
		return index;
	}

	try {
		var projectStartDateCol = getColumnIndex("CE Start Date", toolHeaders);
		var projectEndDateCol = getColumnIndex("CE End Date", toolHeaders);
		var ceBalanceCol = getColumnIndex("CE BALANCE", toolHeaders);
		var processedCol = getColumnIndex("Checking of Unpaid CEs", toolHeaders);
		var glDescriptionCol = getColumnIndex("GL Code", toolHeaders);
		var costCenterCol = getColumnIndex("Cost Center", toolHeaders);
		var ioNumberCol = getColumnIndex("IO Number", toolHeaders);
		var ceNoCol = getColumnIndex("Cost Estimate No.", toolHeaders);
		var vendorNameCol = getColumnIndex("Vendor Name", toolHeaders);
		var costVatExCol = getColumnIndex("Total Cost Estimate Amount (Vat-ex)", toolHeaders);
		var projectNameCol = getColumnIndex("Program Name", toolHeaders);

		// Accrual Report headers (based on your provided headers)
		var accrualHeaders = accrualReportSheet.getRange(1, 1, 1, accrualReportSheet.getLastColumn()).getValues()[0];

		// Map the corresponding columns in "Accrual Report" based on your provided headers
		var accrualStartDateCol = getColumnIndex("PROJECT START DATE", accrualHeaders);
		var accrualEndDateCol = getColumnIndex("PROJECT END DATE", accrualHeaders);
		var accrualGLCodeCol = getColumnIndex("GL CODE", accrualHeaders);
		var accrualCostCenterCol = getColumnIndex("COST CENTER", accrualHeaders);
		var accrualIONumberCol = getColumnIndex("IO NO.", accrualHeaders);
		var accrualCENoCol = getColumnIndex("CE NO.", accrualHeaders);
		var accrualAmountCol = getColumnIndex("AMOUNT", accrualHeaders);
		var accrualVendorNameCol = getColumnIndex("VENDOR NAME", accrualHeaders);
		var accrualProjectNameCol = getColumnIndex("PROGRAM NAME", accrualHeaders);

		// Find the next empty row in the Accrual Report (starting from row 2)
		var lastRowAccrual = accrualReportSheet.getLastRow();
		var nextEmptyRow = lastRowAccrual < 2 ? 2 : lastRowAccrual + 1;

		// Loop through all rows in the Tool database
		for (var i = 2; i <= lastRowToolDB; i++) {
			var projectStartDate = toolDatabaseSheet.getRange(i, projectStartDateCol).getValue();
			var projectEndDate = toolDatabaseSheet.getRange(i, projectEndDateCol).getValue();
			var ceBalance = toolDatabaseSheet.getRange(i, ceBalanceCol).getValue();
			var alreadyProcessed = toolDatabaseSheet.getRange(i, processedCol).getValue();

			// Skip rows that are already processed or if CE BALANCE is not greater than 0
			if (alreadyProcessed === "Yes" || ceBalance <= 0) {
				continue; // Skip this row and move to the next one
			}

			// Data extraction from Tool database
			var glDescription = toolDatabaseSheet.getRange(i, glDescriptionCol).getValue();
			var costCenter = toolDatabaseSheet.getRange(i, costCenterCol).getValue();
			var ioNumber = toolDatabaseSheet.getRange(i, ioNumberCol).getValue();
			var ceNo = toolDatabaseSheet.getRange(i, ceNoCol).getValue();
			var vendorName = toolDatabaseSheet.getRange(i, vendorNameCol).getValue();
			var costVatEx = toolDatabaseSheet.getRange(i, costVatExCol).getValue();
			var projectName = toolDatabaseSheet.getRange(i, projectNameCol).getValue(); // Get Project Name

			// Extract only the first 6 digits of GL Code
			var glCode = glDescription.toString().substring(0, 6);

			// Format the amount with commas and convert it to a string
			var formattedAmount = Number(costVatEx).toLocaleString("en-US", { style: "decimal", minimumFractionDigits: 2, maximumFractionDigits: 2 });

			// Insert data into the correct columns in Accrual Report
			accrualReportSheet.getRange(nextEmptyRow, accrualStartDateCol).setValue(projectStartDate);
			accrualReportSheet.getRange(nextEmptyRow, accrualEndDateCol).setValue(projectEndDate);
			accrualReportSheet.getRange(nextEmptyRow, accrualGLCodeCol).setValue(glCode);
			accrualReportSheet.getRange(nextEmptyRow, accrualCostCenterCol).setValue(costCenter);
			accrualReportSheet.getRange(nextEmptyRow, accrualIONumberCol).setValue(ioNumber);
			accrualReportSheet.getRange(nextEmptyRow, accrualCENoCol).setValue(ceNo);
			accrualReportSheet.getRange(nextEmptyRow, accrualAmountCol).setValue(formattedAmount);
			accrualReportSheet.getRange(nextEmptyRow, accrualVendorNameCol).setValue(vendorName);
			accrualReportSheet.getRange(nextEmptyRow, accrualProjectNameCol).setValue(projectName); // Insert Project Name

			// Increment the nextEmptyRow for the next iteration
			nextEmptyRow++;

			// Mark the row as processed in Tool database
			toolDatabaseSheet.getRange(i, processedCol).setValue("Yes");
		}
	} catch (error) {
		Logger.log("Error: " + error.message);
	}
}
