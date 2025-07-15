function applyFormulaToTable() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tool database');
  
  // Define the starting row (2) to skip the header
  const startRow = 2;

  // Get the headers (assuming headers are in row 1)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the column indexes for all 20 "Invoice Amount (VAT EX)" columns
  const invoiceAmountColumns = [];
  for (let i = 1; i <= 20; i++) {
    const colName = `Invoice Amount (VAT EX) ${i}`;
    const colIndex = headers.indexOf(colName) + 1;
    if (colIndex > 0) {
      invoiceAmountColumns.push(colIndex);
    }
  }

  // Find the column indexes for other required columns
  const colInvoiceTotal = headers.indexOf('INVOICE TOTAL') + 1;
  const colTotalCostEstimate = headers.indexOf('Total Cost Estimate Amount (Vat-ex)') + 1;
  const colCEBalance = headers.indexOf('CE BALANCE') + 1;

  // Get the last row with data in the sheet
  const lastRow = sheet.getLastRow();
  
  // Get the data from the relevant columns for invoice amounts and total cost estimate
  const invoiceAmountRanges = invoiceAmountColumns.map(col => sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues());
  const rangeTotalCostEstimate = sheet.getRange(startRow, colTotalCostEstimate, lastRow - startRow + 1, 1).getValues();

  // Prepare arrays to hold the Invoice Total and CE Balance results
  const invoiceTotalArray = [];
  const ceBalanceArray = [];

  // Loop through each row and calculate the sums and differences
  for (let i = 0; i < invoiceAmountRanges[0].length; i++) {
    let invoiceTotal = 0;

    // Sum all 20 invoice amounts
    for (let j = 0; j < invoiceAmountColumns.length; j++) {
      const amount = invoiceAmountRanges[j][i][0] || 0;
      invoiceTotal += amount;
    }

    invoiceTotalArray.push([invoiceTotal]);

    // Get the total cost estimate for the row
    const totalCostEstimate = rangeTotalCostEstimate[i][0] || 0;

    // Calculate the CE Balance only if Total Cost Estimate Amount (Vat-ex) is present
    const ceBalance = totalCostEstimate ? totalCostEstimate - invoiceTotal : '';

    ceBalanceArray.push([ceBalance]);
  }

  // Set the Invoice Total in the appropriate column
  sheet.getRange(startRow, colInvoiceTotal, invoiceTotalArray.length, 1).setValues(invoiceTotalArray);

  // Set the CE Balance in the appropriate column
  sheet.getRange(startRow, colCEBalance, ceBalanceArray.length, 1).setValues(ceBalanceArray);

  // Optional: Log to confirm success
  Logger.log('Formulas applied to INVOICE TOTAL and CE BALANCE columns for all rows.');
}