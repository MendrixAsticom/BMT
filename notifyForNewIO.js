function checkSpreadsheet() {
	var sheetName = "Spend Plan update(backend db)";
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

	if (!sheet) {
		Logger.log("Sheet not found: " + sheetName);
		return;
	}

	var dataRange = sheet.getDataRange();
	var data = dataRange.getValues();

	// Get the header row (first row)
	var headers = data[0];

	// Find indices of the relevant columns
	var vlookupIndex = headers.indexOf("vlookup");
	var ioIndex = headers.indexOf("MONTHLY");
	var checkerIndex = headers.indexOf("CHECKER");

	for (var i = 1; i < data.length; i++) {
		var vlookupValue = data[i][vlookupIndex];
		var ioValue = data[i][ioIndex];
		var checkerValue = data[i][checkerIndex];

		if (vlookupValue === "NEW" && !checkerValue) {
			Logger.log("Processing row: " + (i + 1));
			processInput(ioValue, i + 1); // Row index adjusted for zero-based array
		}
	}
}

function processInput(ioValue, row) {
	var costElementSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cost Element Owners");
	var spendPlanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Spend Plan update(backend db)");
	var checkerIndex = 23; // Column W corresponds to index 23

	// Columns to check in Cost Element Owners
	var ioColumnsRange = costElementSheet.getRange("AD:AZ").getValues(); // IO Columns

	Logger.log("Processing IO: " + ioValue);

	var first7CharsIO = ioValue.substring(0, 7).trim();
	var first6CharsIO = ioValue.substring(0, 6).trim();
	var matchedColumn = null;

	// Find matching IO column
	function findMatchingIOColumn(ioPrefix) {
		for (var col = 0; col < ioColumnsRange[0].length; col++) {
			for (var row = 0; row < ioColumnsRange.length; row++) {
				var ioElement = String(ioColumnsRange[row][col]).trim();
				if (ioElement && ioElement.substring(0, ioPrefix.length) === ioPrefix) {
					return col + 30; // AD starts from the 30th column
				}
			}
		}
		return null;
	}

	matchedColumn = findMatchingIOColumn(first7CharsIO) || findMatchingIOColumn(first6CharsIO);

	if (matchedColumn !== null) {
		// Always fetch the email from the 3rd row of the matched column
		var emailRange = costElementSheet.getRange(3, matchedColumn);
		var emailCellValue = emailRange.getValue();
		Logger.log("Email value from cell: " + emailCellValue);

		// Split emails based on space, comma, newline, or semicolon
		var emails = emailCellValue
			.split(/[\s,;\n]+/)
			.map(function (email) {
				return email.trim();
			})
			.filter(Boolean); // Remove empty strings

		if (emails.length > 0) {
			var subject = "Request for CE File Upload";
			var message = generateEmailMessage(ioValue);

			var emailSentSuccess = false;
			emails.forEach(function (email) {
				try {
					MailApp.sendEmail({
						to: email,
						subject: subject,
						htmlBody: message,
					});
					Logger.log("Email sent to: " + email);
					emailSentSuccess = true;
				} catch (error) {
					Logger.log("Error sending email to " + email + ": " + error.message);
				}
			});

			if (emailSentSuccess) {
				spendPlanSheet.getRange(row, checkerIndex).setValue(true);
			}
		} else {
			Logger.log("No email addresses found.");
		}
	} else {
		Logger.log("No matching IO column found for IO: " + ioValue);
	}
}

// Generate Email Template for IO Owner
function generateEmailMessage(ioValue) {
	return `
    Dear IO Owner,<br><br>

    We are pleased to inform you that the budget for your project has been approved. The corresponding IO number is <strong>${ioValue}</strong>.<br><br>

    <span style='color: red; font-weight: bold;'>As the Budget owner, you are accountable for all budget expenditures, including preventing unauthorized spending, proper tracking, and ensuring timely payment for satisfactory services.</span><br><br>

    <span style='color: red; font-weight: bold;'>Please share this IO number to your execution partners with discretion to execution partners who need to know this information so they may proceed with procurement services/materials for your Program/Campaign. Be mindful that anyone who knows your IO number may charge expenses to this budget.</span><br><br>

    <span style='color: red; font-weight: bold;'>NO Award Document, No Work.</span> All Purchases must have a confirmed Award Document, approved by an authorized employee with the appropriate LOA. Failure to comply will result in non-payment of the unauthorized services and HR sanctions to the concerned employee as stated in Globe's Code of Conduct.<br><br>

    <table border="1" style="border-collapse: collapse; width: 100%;">
      <tr>
        <th>Award Document</th>
        <th>Sample Transactions in Marketing</th>
      </tr>
      <tr>
        <td>Purchase Order (via Ariba)</td>
        <td>Event Agency, Merchandising Materials, Managed Services</td>
      </tr>
      <tr>
        <td>Cost Estimate (approval signature must be on the CE document)</td>
        <td>Media buys</td>
      </tr>
      <tr>
        <td>Contracts (approved via Docusign)</td>
        <td>Agencies (Creative, PR, Media, etc); Ambassadors/Influencers</td>
      </tr>
      <tr>
        <td>Conforme (approval signature must be on the document)</td>
        <td>Partnerships, Sponsorships</td>
      </tr>
    </table><br><br>

    <strong>FOR Non-PO Purchases:</strong><br>
    Upload APPROVED Cost Estimate, Contract, or Conforme in the <a href="https://script.google.com/a/macros/globe.com.ph/s/AKfycbzKuhoJwjQeoc8VWTOVZ2PZuOjocUcdqy7pRzMRStOOgBLuYLdiEeE5txanUfP-17RR/exec">Budget Tool</a>.<br><br>

    <span style='color: red; font-weight: bold;'>File Name Convention:</span> <strong>IOnumber_Vendor Name_CENumberv</strong> (e.g., MKTG5D12344_ABC Corp_CE00123) for easy identification and organization.<br><br>

    <strong>Basis for Payment:</strong> The uploaded CE or Contract and Invoices will serve as the basis for payment. It is crucial to upload the correct and final version to avoid any issues during the payment process.<br><br>

    Request for PR filing through the <a href="https://script.google.com/a/macros/globe.com.ph/s/AKfycbxM4a0ITeWjkIyXHbahfyomdw0e4Hf9dCL7-oBUw_Q/dev">Budget Tool</a>. The Budget Management team will be the one to file in Ariba. Attachments are approved PWP, Contract/Cost Estimate/Conforme/Proposal.<br><br>

    <span style='color: red; font-weight: bold;'>Delays in submitting Cost Estimates, Contracts, or Purchase Orders</span> will cause the corresponding amount to be tagged as <span style='color: red; font-weight: bold;'>unutilized and will result in budget reallocation</span>.<br><br>

    Thank you for your cooperation.<br><br>

    Should you have any questions or concerns, please do not hesitate to contact <a href='mailto:marketingsettlement@globe.com.ph'>marketingsettlement@globe.com.ph</a>.<br><br>

    Sincerely,<br>
    Marketing Budget Team
  `;
}
