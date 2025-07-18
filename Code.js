const PREVIEW_FOLDER_ID = "17JGBt2AUzG5NSAa9BhhJCRbq8IVpCp4K";
const PRPO_FOLDER_ID = "1E2DX-t6mYEcwL7ZHFQKa2o5WiB09_vQe";
const CE_FOLDER_ID = "15D-Iw5DipcotqFZtFKyvQZQKhvzn1zzo";
const CE_INVOICE_FOLDER_ID = "1rGoow6auDco7X9lPZn5kSTeAFdcNI0w6";

function doGet() {
	return (
		HtmlService.createTemplateFromFile("Index")
			.evaluate()
			.addMetaTag("viewport", "width=device-width, initial-scale=1")
			//.setFaviconUrl('https://www.flaticon.com/free-icon/internet_10453396?term=website+logo&page=5&position=42&origin=tag&related_id=10453396')
			.setTitle("Budget Tool")
			.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
	);
}

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

//FOR LOG IN PAGE ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function authenticate(email) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ADMINaoMarch4");
	var startRow = 2; // Starting from row 2
	var emailColumn = 3; // Column C "Email"
	var roleColumn = 5; // Column E "Role"

	// Get all email addresses in column C, starting from row 2
	var data = sheet.getRange(startRow, emailColumn - 1, sheet.getLastRow() - startRow + 1, 3).getValues();

	// Check if the username exists in the email column
	var dataRow = data.find(function (row) {
		return row[1] === email;
	});

	if (!dataRow) {
		return false;
	}
	return dataRow; //  return ['username','email','login type'] for valid emails
}

function validateSSO() {
	const user = Session.getActiveUser();
	const auth = authenticate(user.getEmail()); //authenticate user email

	if (auth) {
		const userData = {
			firstName: auth[0], // name col
			email: auth[1], // email col
			loginType: auth[2], // role col
		};
		Logger.log(auth, userData);
		return userData;
	} else {
		return false;
	}
}

function validateLogin(username) {
	var validationResult = authenticate(username);
	return validationResult === "valid" ? validationResult : "invalid";
}

function getUserNameByUsername(username) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ADMINaoMarch4");
	var data = sheet.getDataRange().getValues();
	for (var i = 1; i < data.length; i++) {
		if (data[i][2] === username) {
			// Email is in column C (index 2)
			return {
				firstName: data[i][1].split(" ")[0], // Assuming first name is the first part of the full name
				email: data[i][2],
			};
		}
	}
	return null;
}

//FOR UPLOAD COST ESTIMATE PAGE --------------------------------------------------------------------------------------

function uploadFile(fileData, fileName, ioNumber) {
	var fileName = `${fileName.split(".")[0]}_${Date.now()}.pdf`;
	const folderId = PREVIEW_FOLDER_ID; // Preview Folder
	const folder = DriveApp.getFolderById(folderId);
	const files = folder.getFilesByName(fileName);

	//validateIO
	let allIO = fetchIO();
	allIO = allIO.flat().slice(1);
	const index = allIO.indexOf(ioNumber);
	if (index == -1) {
		Logger.log(`IO number ${ioNumber} not found`);
		return {
			error: "IO Number does not exist. Contact Marketing Investments Team: Sherann Barrameda",
		};
	}

	if (files.hasNext()) {
		return {
			error: "File with the same name already exists. Please upload a new CE File with file name in this format: IO_PartnerName_CENumber_Corrected.pdf or .xls/.xlsx.",
		};
	}

	// Determine MIME type based on file extension
	const extension = fileName.split(".").pop().toLowerCase();
	let mimeType;

	if (extension === "pdf") {
		mimeType = MimeType.PDF;
	} else if (extension === "xls") {
		mimeType = MimeType.MICROSOFT_EXCEL;
	} else if (extension === "xlsx") {
		mimeType = MimeType.MICROSOFT_EXCEL_LEGACY;
	} else {
		return {
			error: "Invalid file type. Please upload a PDF, XLS, or XLSX file.",
		};
	}

	const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
	const file = folder.createFile(blob);

	file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
	return {
		fileId: file.getId(),
		fileName: fileName, // Return file name for autofill
		fileUrl: file.getUrl(),
	};
}

// Fetch all vendor names from the "Vendor DB" sheet
function searchVendors() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendor DB");
	if (!sheet) {
		throw new Error("Vendor DB sheet not found");
	}

	// Get all vendor names from column A (assuming vendor names are in column A)
	const data = sheet.getRange("A:A").getValues().flat(); // Flatten to 1D array

	// Filter out empty or invalid values
	const vendors = data
		.filter((vendor) => typeof vendor === "string" && vendor.trim() !== "") // Filter out non-string or empty values
		.map((vendor) => vendor.trim()); // Trim whitespace

	return vendors; // Return all vendor names
}

function getGLInfo(vendorName) {
	console.log("getGLInfo called with vendorName:", vendorName);

	if (!vendorName || typeof vendorName !== "string") {
		console.log("Invalid vendorName provided.");
		return { glCode: null, glDescription: null };
	}

	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendor DB");
	const data = sheet.getDataRange().getValues(); // Fetch all data
	if (data.length < 2) {
		console.log("No data available in the sheet.");
		return { glCode: null, glDescription: null };
	}

	const headers = data[0]; // Get the header row
	const vendorIndex = headers.indexOf("VENDOR");
	const glCodeIndex = headers.indexOf("GL Code");
	const glDescriptionIndex = headers.indexOf("GL Description");

	if (vendorIndex === -1 || glCodeIndex === -1 || glDescriptionIndex === -1) {
		console.log("Required columns not found.");
		return { glCode: null, glDescription: null };
	}

	console.log(`Column indices - Vendor: ${vendorIndex}, GL Code: ${glCodeIndex}, GL Description: ${glDescriptionIndex}`);

	for (let i = 1; i < data.length; i++) {
		// Start from row 1 (skip headers)
		let currentVendor = data[i][vendorIndex];

		if (typeof currentVendor === "string" && currentVendor.trim() !== "") {
			if (currentVendor.toLowerCase() === vendorName.toLowerCase()) {
				console.log(`Vendor found at row ${i + 1}`);
				console.log("GL Code:", data[i][glCodeIndex]);
				console.log("GL Description:", data[i][glDescriptionIndex]);

				return {
					glCode: data[i][glCodeIndex] || null,
					glDescription: data[i][glDescriptionIndex] || null,
				};
			}
		}
	}

	console.log("Vendor not found.");
	return { glCode: null, glDescription: null };
}

function checkCostEstimateNo(costEstimateNo) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");
	var dataRange = sheet.getDataRange(); // Get the entire data range
	var data = dataRange.getValues(); // Get all values as a 2D array

	// Find the index of the "Cost Estimate No." column
	var headers = data[0]; // Get the header row
	var costEstimateNoColumnIndex = headers.indexOf("Cost Estimate No."); // Find the index based on column name

	if (costEstimateNoColumnIndex === -1) {
		throw new Error("Cost Estimate No. column not found");
	}

	// Check for duplicates in the "Cost Estimate No." column
	var isDuplicate = data.slice(1).some(function (row) {
		return row[costEstimateNoColumnIndex] === costEstimateNo; // Check for duplicates
	});

	return isDuplicate; // Return true if duplicate found, otherwise false
}

function getIOBalance(ioNumber) {
	const ss = SpreadsheetApp.getActiveSpreadsheet(); // Use the active spreadsheet
	const sheet = ss.getSheetByName("Spend Plan update(backend db)");
	const data = sheet.getDataRange().getValues();

	// Get headers from the first row
	const headers = data[0];

	// Find the indices of the relevant columns
	const ioNumberIndex = headers.indexOf("MONTHLY"); // Column name for IO number
	const ioBalanceIndex = headers.indexOf("IO Balance"); // Column name for IO balance

	// Find IO balance based on the IO number
	for (let i = 1; i < data.length; i++) {
		if (data[i][ioNumberIndex] === ioNumber) {
			// Check if IO number matches
			return data[i][ioBalanceIndex]; // Return the corresponding IO balance
		}
	}
	return null; // Return null if IO number not found
}

//Get the Link of the File
function getUploadedFileUrl(fileId) {
	try {
		const folderId = CE_FOLDER_ID; //Update CE Folder
		const storageFolder = DriveApp.getFolderById(folderId);
		const fileToMove = DriveApp.getFileById(fileId).moveTo(storageFolder);
		const fileUrl = fileToMove.getUrl();

		return fileUrl;
	} catch (e) {
		Logger.log("Error fetching file URL: " + e.toString());
	}
	return null;
}

function checkDuplicateCostEstimateNo(costEstimateNo) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");
	var data = sheet.getDataRange().getValues();
	var headers = data[0];
	var ceNoIndex = headers.indexOf("Cost Estimate No.");

	if (ceNoIndex === -1) return false;

	for (var i = 1; i < data.length; i++) {
		if (data[i][ceNoIndex] === costEstimateNo) {
			return true;
		}
	}
	return false;
}

function saveDataToSpreadsheet(data) {
	Logger.log("Received data to save: " + JSON.stringify(data));

	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");

	// Find the first empty row
	var timestampColumn = 2; // Column B for Timestamp
	var lastRow = sheet.getLastRow();
	var range = sheet.getRange(2, timestampColumn, lastRow);
	var values = range.getValues();

	var startRow =
		values.findIndex(function (row) {
			return row[0] === "";
		}) + 2;

	if (startRow === 1) {
		startRow = lastRow + 1;
	}

	// Get headers and prepare new row data
	var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	var newRowData = new Array(headers.length).fill("");

	// Map data to columns
	headers.forEach(function (header, index) {
		if (data[header] !== undefined) {
			newRowData[index] = data[header];
		}
	});

	// Handle file upload if present
	// Ensure 'Uploaded CE File' is included in the data
	var uploadedCEFileIndex = headers.indexOf("Uploaded CE File");
	if (uploadedCEFileIndex !== -1 && data["Uploaded CE File"]) {
		var fileUrl = getUploadedFileUrl(data["Uploaded CE File"]); // Fetch the file URL
		if (fileUrl) {
			sheet.getRange(startRow, uploadedCEFileIndex + 1).setFormula(`=HYPERLINK("${fileUrl}", "${data["Uploaded CE File"]}")`);
		} else {
			newRowData[uploadedCEFileIndex] = data["Uploaded CE File"];
		}
	}

	// Write the new row
	sheet.getRange(startRow, 1, 1, newRowData.length).setValues([newRowData]);

	// Run automations
	validateCostEstimates();
	updateCEPaymentStatus();
	applyFormulaToTable();

	Logger.log("Data saved in row: " + startRow);
	Logger.log("Saved row data: " + JSON.stringify(newRowData));

	return {
		success: true,
		message: "Data saved successfully",
		savedData: newRowData,
	};
}

//test
function testSaveData() {
	var testData = {
		"Time Stamp (Tool Generated)": new Date().toISOString(), // Adds timestamp
		"Vendor Name": "Test Vendor",
		"Cost Estimate No.": "CE12345",
		"Uploaded CE File": {
			url: "https://drive.google.com/file/d/EXAMPLE/view",
			name: "TestFile",
		},
	};
	saveDataToSpreadsheet(testData);
}

//Automated Email for New CE Saved in Tool Database
function sendNotificationCEEmail(costCenter) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auto Email");
	if (!sheet) {
		Logger.log("Email sheet not found.");
		return;
	}

	const headers = sheet
		.getRange(1, 1, 1, sheet.getLastColumn())
		.getValues()[0]
		.map((h) => h.trim());
	const emailColumnIndex = headers.indexOf("New CE Saved") + 1;
	if (emailColumnIndex === 0) {
		Logger.log("New CE Saved column not found.");
		return;
	}

	const emailColumn = sheet
		.getRange(2, emailColumnIndex, sheet.getLastRow() - 1, 1)
		.getValues()
		.flat()
		.filter(String);
	if (emailColumn.length === 0) {
		Logger.log("No emails found in the New CE Saved column.");
		return;
	}

	// Get emails based on Cost Center
	const costCenterEmails = getEmailsByCostCenter(costCenter);
	const allEmails = [...new Set([...emailColumn, ...costCenterEmails])]; // Combine and remove duplicates

	// Log the Cost Center and email recipients
	Logger.log(`Cost Center: ${costCenter}`);
	Logger.log(`Email recipients: ${allEmails.join(", ")}`);

	const subject = "New CE Saved in Tool Database";
	const body = "Dear Budget Team,\n\nA new CE has been saved in the tool database for your checking.\n\nBest regards,\nBudget Tool";

	MailApp.sendEmail({
		to: allEmails.join(","),
		subject: subject,
		body: body,
	});
}

function sendCEConfirmationEmail(userEmail, costEstimateNo) {
	const subject = "Your CE Number Has Been Saved Successfully";
	const body = `
        <p>Dear User,</p>
        <p>Your CE Number: <strong>${costEstimateNo}</strong> has been saved successfully in the Tool database.</p>
        <p>To review your entry, please click the link below:</p>
        <p>
            <a href="https://docs.google.com/spreadsheets/d/1xgROcPRmkI6_dLfZmadbCI0Yv_0s_D7S23PIdNnkA40/edit?usp=sharing">
                View Tool Database
            </a>
        </p>
        <p>Best regards,<br>Your Budget Tool Team</p>
    `;

	MailApp.sendEmail({
		to: userEmail,
		subject: subject,
		htmlBody: body, // Use htmlBody for formatted content
	});
}

//GLOBAL FUNCTIONS ------------------------------------------------------

function getGLDropdownOptions() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cost Element Owners");
	const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); // Get GL Code (A) and GL Description (B) starting from row 2

	// Concatenate GL Code and GL Description, excluding the header row
	const glOptions = data.map((row) => `${row[0]}-${row[1]}`);
	return glOptions; // Return the cADMINanated options to the client
}

function getCostCenterOptions() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName("ADMINaoMarch4"); // Change "Admin" to your sheet name if needed
	const costCenterRange = sheet.getRange("G2:G"); // Assuming data starts from G2 and goes down
	const costCenterValues = costCenterRange.getValues().filter(String); // Filter out empty values
	return costCenterValues.map((row) => row[0]); // Return as a flat array
}

function getEmailsByCostCenter(costCenter) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email based on CC");
	if (!sheet) {
		Logger.log("Email based on CC sheet not found.");
		return [];
	}

	const data = sheet.getDataRange().getValues();
	const headers = data[0];
	const costCenterIndex = headers.indexOf("Cost Center");
	const emailIndex = headers.indexOf("Email");

	if (costCenterIndex === -1 || emailIndex === -1) {
		Logger.log("Required columns (Cost Center or Email) not found.");
		return [];
	}

	const emails = [];
	let matchedCostCenter = null;

	// Log the Cost Center being searched
	Logger.log(`Searching for Cost Center: ${costCenter}`);

	// Try partial matching
	for (let i = 1; i < data.length; i++) {
		const sheetCostCenter = data[i][costCenterIndex].toString().trim();
		const partialMatchLength = Math.min(costCenter.length, sheetCostCenter.length);

		// Check for partial match (first 7 characters, then up to 10)
		for (let j = 7; j <= 10; j++) {
			if (j > partialMatchLength) break; // Stop if we exceed the length of either string

			const partialCostCenter = costCenter.substring(0, j);
			const partialSheetCostCenter = sheetCostCenter.substring(0, j);

			if (partialCostCenter === partialSheetCostCenter) {
				matchedCostCenter = sheetCostCenter;

				// Split the email cell content by commas, spaces, or new lines
				const emailCellContent = data[i][emailIndex].toString().trim();
				const emailList = emailCellContent.split(/[\s,\n]+/).filter((email) => email.trim() !== "");

				// Add all valid emails to the emails array
				emails.push(...emailList);
				Logger.log(`Partial match found: ${partialCostCenter} (Cost Center: ${sheetCostCenter})`);
				break;
			}
		}

		if (matchedCostCenter) break; // Stop after the first match
	}

	if (!matchedCostCenter) {
		Logger.log(`No match found for Cost Center: ${costCenter}`);
	}

	return emails;
}

//REQUEST FOR PR/PO PAGE --------------------------------------------------------------------------------------

function getUploadedFileUrlPRPO(fileId) {
	try {
		// const previewFolderId = '17JGBt2AUzG5NSAa9BhhJCRbq8IVpCp4K'
		const storageFolder = DriveApp.getFolderById(PRPO_FOLDER_ID);
		const fileToMove = DriveApp.getFileById(fileId).moveTo(storageFolder);
		const fileUrl = fileToMove.getUrl();

		return fileUrl;
	} catch (e) {
		Logger.log("Error in getUploadedFileUrl: " + e.toString());
		return null;
	}
}

function fetchIO() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UniqueIO");
	const data = sheet.getDataRange().getValues();
	return data;
}

function uploadRequestFile(fileData, fileName, ioNumber) {
	try {
		//validateIO
		let allIO = fetchIO();
		allIO = allIO.flat().slice(1);
		const index = allIO.indexOf(ioNumber);
		if (index == -1) {
			Logger.log(`IO number ${ioNumber} not foud`);
			return {
				error: "IO Number does not exist. Contact Marketing Investments Team: Sherann Barrameda",
			};
		}

		// Determine MIME type based on file extension
		const extension = fileName.split(".").pop().toLowerCase();
		let mimeTypePR;

		if (extension === "pdf") {
			mimeTypePR = MimeType.PDF;
		} else if (extension === "xls") {
			mimeTypePR = MimeType.MICROSOFT_EXCEL;
		} else if (extension === "xlsx") {
			mimeTypePR = MimeType.MICROSOFT_EXCEL_LEGACY;
		} else {
			return {
				error: "Invalid file type. Please upload a PDF, XLS, or XLSX file.",
			};
		}

		var fileName = `${fileName.split(".")[0]}_${Date.now()}.${extension}`;
		const folderId = PREVIEW_FOLDER_ID; // Preview Folder
		const folder = DriveApp.getFolderById(folderId);
		const files = folder.getFilesByName(fileName);

		if (files.hasNext()) {
			return {
				error: "File with the same name already exists. Please upload a new PR/PO File with file name in this format: IO_PartnerName_CENumber_Corrected.pdf or .xls/.xlsx",
			};
		}

		console.log("before blob1");
		const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeTypePR, fileName);
		console.log(blob);
		const file = folder.createFile(blob);
		file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
		console.log("after blob");
		return {
			fileId: file.getId(),
			fileName: fileName,
			fileUrl: file.getUrl(),
		};
	} catch (error) {
		return { error: error.message };
	}
}

//end

// Function to process and extract data from the uploaded file
function processUploadedRequestFile(fileId) {
	try {
		const file = DriveApp.getFileById(fileId);
		const doc = DocumentApp.openById(fileId);
		const body = doc.getBody().getText(); // Retrieve text from the document

		// Log the document text for debugging purposes
		Logger.log(body);

		// Process the text to extract relevant data
		const extractedData = extractDataFromText(body);

		// Log the extracted data for debugging purposes
		Logger.log(extractedData);

		// Return the extracted data to the client-side
		return extractedData;
	} catch (error) {
		Logger.log("Error processing the file: " + error.message);
		throw new Error("Error processing the file: " + error.message);
	}
}

function getFileNameFromId(fileId) {
	try {
		const file = DriveApp.getFileById(fileId);
		if (file) {
			// Make sure the file is accessible via link
			file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
			return {
				fileName: file.getName(),
				fileUrl: file.getUrl(),
			};
		}
		return null;
	} catch (e) {
		Logger.log("Error in getFileNameFromId: " + e.toString());
		return null;
	}
}

function saveDataToSpreadsheetPRPO(data) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PR-PO format");
	var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	var newRowValues = new Array(headers.length).fill("");

	// Find the first empty row
	var lastRow = sheet.getLastRow();
	var firstEmptyRow = lastRow + 1;

	// Map data to columns based on headers
	headers.forEach(function (header, index) {
		if (data[header] !== undefined) {
			// Check if this is a hyperlink formula
			if (data[header].toString().startsWith("=HYPERLINK")) {
				newRowValues[index] = data[header];
			} else {
				newRowValues[index] = data[header];
			}
		}
	});

	// Write the new row
	sheet.getRange(firstEmptyRow, 1, 1, headers.length).setValues([newRowValues]);

	return { success: true };
}

// Automated Email for New PRPO
function sendNotificationCEEmailPRPO(costCenter) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auto Email");
	if (!sheet) {
		Logger.log("Email sheet not found.");
		return;
	}

	const headers = sheet
		.getRange(1, 1, 1, sheet.getLastColumn())
		.getValues()[0]
		.map((h) => h.trim());
	const emailColumnIndex = headers.indexOf("New PRPO Added") + 1;
	if (emailColumnIndex === 0) {
		Logger.log("New PRPO Added column not found.");
		return;
	}

	const emailAddresses = sheet
		.getRange(2, emailColumnIndex, sheet.getLastRow() - 1, 1)
		.getValues()
		.flat()
		.filter(String);
	if (emailAddresses.length === 0) {
		Logger.log("No emails found in the New PRPO Added column.");
		return;
	}

	// Get emails based on Cost Center
	const costCenterEmails = getEmailsByCostCenter(costCenter);
	const allEmails = [...new Set([...emailAddresses, ...costCenterEmails])]; // Combine and remove duplicates

	// Log the Cost Center and email recipients
	Logger.log(`Cost Center: ${costCenter}`);
	Logger.log(`Email recipients: ${allEmails.join(", ")}`);

	const subject = "New PR/PO Saved in PR-PO Report Sheet";
	const body = "Dear Budget Team,\n\nA new PR/PO has been saved in the PR-PO Report Sheet for your checking.\n\nBest regards,\nBudget Tool";

	MailApp.sendEmail({
		to: allEmails.join(","),
		subject: subject,
		body: body,
	});
}

function sendPRPOConfirmationEmail(currentUserEmail, costIONo) {
	const subject = "Your CE Number Has Been Saved Successfully";
	const body = `
        <p>Dear User,</p>
        <p>Your IO Number: <strong>${costIONo}</strong> has been saved successfully in the PR/PO database</p>

        <p>Best regards,<br>Your Budget Tool Team</p>
    `;

	MailApp.sendEmail({
		to: currentUserEmail,
		subject: subject,
		htmlBody: body, // Use htmlBody for formatted content
	});
}

//Bulk Invoice Page --------------------------------------------------------------------------------------

function newSaveUploadedFile(base64Data, fileName) {
	try {
		var fileName = `${fileName.split(".")[0]}_${Date.now()}.pdf`;
		const folder = DriveApp.getFolderById(PREVIEW_FOLDER_ID); // Update Preview Folder
		const extension = fileName.split(".").pop().toLowerCase();
		let mimeType;

		if (extension === "pdf") {
			mimeType = MimeType.PDF;
		} else if (extension === "xls") {
			mimeType = MimeType.MICROSOFT_EXCEL_LEGACY;
		} else if (extension === "xlsx") {
			mimeType = MimeType.MICROSOFT_EXCEL;
		} else {
			return { error: "Invalid file type. Please upload a PDF, XLS, or XLSX file." };
		}

		// Decode base64 data
		const decodedData = Utilities.base64Decode(base64Data);
		const blob = Utilities.newBlob(decodedData, mimeType, fileName);
		const file = folder.createFile(blob);

		file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

		return {
			fileUrl: file.getUrl(),
			fileName: file.getName(),
			fileId: file.getId(),
		};
	} catch (error) {
		return { error: error.message };
	}
}

// Function to search for CE Number, IO Number, and Vendor Name
function searchCENumber(ceNumber, ioNumber = null, vendorName = null) {
	try {
		Logger.log("Searching for CE Number: " + ceNumber);

		const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");
		if (!sheet) {
			Logger.log("Error: 'Tool database' sheet not found");
			return JSON.stringify({ error: "Sheet 'Tool database' not found" });
		}

		const dataRange = sheet.getDataRange();
		const values = dataRange.getValues();
		Logger.log("Total rows in sheet: " + values.length);

		const headers = values[0]; // assuming the headers are in the first row

		// Find the column indices for CE Number, IO Number, and Vendor Name
		const colCENumber = headers.indexOf("Cost Estimate No.");
		const colIONumber = headers.indexOf("IO Number");
		const colVendorName = headers.indexOf("Vendor Name");

		if (colCENumber === -1 || colIONumber === -1 || colVendorName === -1) {
			Logger.log("Error: Required columns not found");
			return JSON.stringify({ error: "Required columns not found" });
		}

		let result = { found: false, data: {}, isUnique: true, nextSearch: null };

		// Filter rows based on CE Number (case-insensitive and trimmed)
		let filteredRows = values.filter((row) => {
			const rowCENumber = row[colCENumber].toString().trim().toLowerCase();
			return rowCENumber === ceNumber.trim().toLowerCase();
		});

		if (filteredRows.length === 0) {
			Logger.log("No data found for CE Number: " + ceNumber);
			return JSON.stringify({ found: false, error: "No data found for CE Number: " + ceNumber });
		}

		// If IO Number is provided, filter further (case-insensitive and trimmed)
		if (ioNumber) {
			filteredRows = filteredRows.filter((row) => {
				const rowIONumber = row[colIONumber].toString().trim().toLowerCase();
				return rowIONumber === ioNumber.trim().toLowerCase();
			});
			if (filteredRows.length === 0) {
				Logger.log("No data found for IO Number: " + ioNumber);
				return JSON.stringify({ found: false, error: "No data found for IO Number: " + ioNumber });
			}
		}

		// If Vendor Name is provided, filter further (case-insensitive and trimmed)
		if (vendorName) {
			filteredRows = filteredRows.filter((row) => {
				const rowVendorName = row[colVendorName].toString().trim().toLowerCase();
				return rowVendorName === vendorName.trim().toLowerCase();
			});
			if (filteredRows.length === 0) {
				Logger.log("No data found for Vendor Name: " + vendorName);
				return JSON.stringify({ found: false, error: "No data found for Vendor Name: " + vendorName });
			}
		}

		// Check if the result is unique
		if (filteredRows.length > 1) {
			result.isUnique = false;
			if (!ioNumber) {
				result.nextSearch = "IO Number";
			} else if (!vendorName) {
				result.nextSearch = "Vendor Name";
			}
		}

		// If unique, return the first row's data
		if (filteredRows.length === 1) {
			result.found = true;
			result.data = {
				"Time Stamp (Tool Generated)": filteredRows[0][headers.indexOf("Time Stamp (Tool Generated)")].toString(),
				"Email Address": filteredRows[0][headers.indexOf("Email Address")].toString(),
				"CE Start Date": filteredRows[0][headers.indexOf("CE Start Date")].toString(),
				"CE End Date": filteredRows[0][headers.indexOf("CE End Date")].toString(),
				"Date of Issue": filteredRows[0][headers.indexOf("Date of Issue")].toString(),
				"Program Name": filteredRows[0][headers.indexOf("Program Name")].toString(),
				"Vendor Name": filteredRows[0][headers.indexOf("Vendor Name")].toString(),
				"Signed by Vendor?": filteredRows[0][headers.indexOf("Signed by Vendor?")].toString(),
				"Globe Authorize Signatory": filteredRows[0][headers.indexOf("Globe Authorize Signatory")].toString(),
				"Payment Terms": filteredRows[0][headers.indexOf("Payment Terms")].toString(),
				"Cost Estimate No.": filteredRows[0][headers.indexOf("Cost Estimate No.")].toString(),
				Currency: filteredRows[0][headers.indexOf("Currency")].toString(),
				"Total Cost Estimate Amount (Vat-ex)": filteredRows[0][headers.indexOf("Total Cost Estimate Amount (Vat-ex)")].toString(),
				"GL Description": filteredRows[0][headers.indexOf("GL Description")].toString(),
				"IO Number": filteredRows[0][headers.indexOf("IO Number")].toString(),
				"IO BALANCE": filteredRows[0][headers.indexOf("IO BALANCE")].toString(),
				"Cost Center": filteredRows[0][headers.indexOf("Cost Center")].toString(),
				"Accrued?": filteredRows[0][headers.indexOf("Accrued?")].toString(),
				"CE Payment Status": filteredRows[0][headers.indexOf("CE Payment Status")].toString(),
				"Uploaded CE File": filteredRows[0][headers.indexOf("Uploaded CE File")].toString(),
				"SAP Ariba Reference Code 1": filteredRows[0][headers.indexOf("SAP Ariba Reference Code 1")].toString(),
				"SAP Ariba Reference Code 2": filteredRows[0][headers.indexOf("SAP Ariba Reference Code 2")].toString(),
				"SAP Ariba Reference Code 3": filteredRows[0][headers.indexOf("SAP Ariba Reference Code 3")].toString(),
				"SAP Ariba Reference Code 4": filteredRows[0][headers.indexOf("SAP Ariba Reference Code 4")].toString(),
				"SAP Ariba Reference Code 5": filteredRows[0][headers.indexOf("SAP Ariba Reference Code 5")].toString(),
			};
		}

		Logger.log("Returning result: " + JSON.stringify(result));
		return JSON.stringify(result);
	} catch (error) {
		Logger.log("Error in searchCENumber: " + error.toString());
		return JSON.stringify({ error: error.toString() });
	}
}

function saveBulkInvoiceData(fileName, ioNumber, data) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");
	if (!sheet) {
		throw new Error("Bulk Invoice sheet not found");
	}

	// Get all data from the sheet
	const dataRange = sheet.getDataRange();
	const values = dataRange.getValues();
	const headers = values[0]; // First row contains headers

	// Find the column indices for 'Uploaded CE File' and 'IO Number'
	const fileColumnIndex = headers.indexOf("Uploaded CE File");
	const ioNumberColumnIndex = headers.indexOf("IO Number");

	if (fileColumnIndex === -1 || ioNumberColumnIndex === -1) {
		throw new Error("Required columns (Uploaded CE File or IO Number) not found");
	}

	// Log the fileName and ioNumber being searched
	console.log("Searching for File Name:", fileName, "and IO Number:", ioNumber);

	// Find the row index matching the file name and IO Number
	let fileRowIndex = -1;
	for (let i = 1; i < values.length; i++) {
		const rowFileName = values[i][fileColumnIndex];
		const rowIONumber = values[i][ioNumberColumnIndex];

		// Log the values in the spreadsheet for debugging
		console.log(`Row ${i + 1}: File Name: "${rowFileName}", IO Number: "${rowIONumber}"`);

		// Ensure rowFileName and rowIONumber are treated as strings
		const rowFileNameStr = String(rowFileName || "")
			.trim()
			.toLowerCase();
		const rowIONumberStr = String(rowIONumber || "")
			.trim()
			.toLowerCase();
		const fileNameStr = String(fileName || "")
			.trim()
			.toLowerCase();
		const ioNumberStr = String(ioNumber || "")
			.trim()
			.toLowerCase();

		// Check if the file name and IO Number match (case-insensitive and trimmed)
		if (rowFileNameStr === fileNameStr && rowIONumberStr === ioNumberStr) {
			fileRowIndex = i + 1; // Rows are 1-indexed in Google Sheets
			console.log("Match found at row:", fileRowIndex);
			break;
		}
	}

	if (fileRowIndex === -1) {
		throw new Error("File name and IO Number combination not found in the sheet");
	}

	// Define the invoice fields up to 20 slots
	const invoiceFields = [];
	for (let i = 1; i <= 20; i++) {
		invoiceFields.push({
			email: `Invoice Email Address ${i}`,
			vendorName: `Invoice Vendor Name ${i}`,
			number: `Invoice Number ${i}`,
			amount: `Invoice Amount (VAT EX) ${i}`,
			timestamp: `INVOICE SUBMISSION TIME STAMP ${i}`,
			accrualRef: `Accrual Reference Doc ${i}`,
			uploadedFile: `Uploaded Invoice File ${i}`,
			invoiceDate: `Invoice Date ${i}`,
			invoiceDueDate: `Invoice Due Date ${i}`,
		});
	}
	console.log(fileRowIndex);
	// Find the first available slot
	let nextSlotIndex = 0;
	for (let i = 0; i < invoiceFields.length; i++) {
		const emailCol = headers.indexOf(invoiceFields[i].email) + 1;
		const emailValue = sheet.getRange(fileRowIndex, emailCol).getValue();
		if (!emailValue) {
			nextSlotIndex = i;
			break;
		}
	}

	if (nextSlotIndex >= invoiceFields.length) {
		throw new Error("No more invoice slots available");
	}

	// Save each row of data to the next available invoice slot
	data.forEach((row, index) => {
		const field = invoiceFields[nextSlotIndex];

		// Save the row's data to the corresponding invoice slot
		const emailCol = headers.indexOf(field.email) + 1;
		const vendorNameCol = headers.indexOf(field.vendorName) + 1;
		const numberCol = headers.indexOf(field.number) + 1;
		const amountCol = headers.indexOf(field.amount) + 1;
		const timestampCol = headers.indexOf(field.timestamp) + 1;
		const accrualRefCol = headers.indexOf(field.accrualRef) + 1;
		const uploadedFileCol = headers.indexOf(field.uploadedFile) + 1;
		const invoiceDateCol = headers.indexOf(field.invoiceDate) + 1;
		const invoiceDueDateCol = headers.indexOf(field.invoiceDueDate) + 1;
		console.log("accrualRefCol: ", accrualRefCol);
		console.log("invoiceDateCol: ", invoiceDateCol);
		console.log("invoiceDueDateCol: ", invoiceDueDateCol);

		// Save the user's email address (hidden in modal but saved in spreadsheet)
		sheet.getRange(fileRowIndex, emailCol).setValue(row.invoiceEmail);

		// Save the Invoice Vendor Name
		sheet.getRange(fileRowIndex, vendorNameCol).setValue(row.invoiceVendorName);

		// Save the rest of the data
		sheet.getRange(fileRowIndex, numberCol).setValue(row.invoiceNumber);
		sheet.getRange(fileRowIndex, amountCol).setValue(row.invoiceAmount);
		sheet.getRange(fileRowIndex, timestampCol).setValue(row.invoiceTimestamp);
		sheet.getRange(fileRowIndex, accrualRefCol).setValue(row.accrualRefDoc || "");
		sheet.getRange(fileRowIndex, invoiceDateCol).setValue(row.invoiceDate);
		sheet.getRange(fileRowIndex, invoiceDueDateCol).setValue(row.invoiceDueDate);

		// Save the uploaded file as a hyperlink
		if (row.uploadedFile) {
			try {
				// Extract the file ID from the URL
				const fileId = row.uploadedFile.match(/[-\w]{25,}/);
				if (fileId) {
					const storageFolder = DriveApp.getFolderById(CE_INVOICE_FOLDER_ID);
					const file = DriveApp.getFileById(fileId[0]).moveTo(storageFolder);

					const fileName = file.getName();
					const url = file.getUrl();

					// Remove file extensions
					const fileNameWithoutExtension = fileName.replace(/\.(pdf|xls|xlsx)$/i, "");

					// Create a hyperlink formula
					const hyperlinkFormula = `=HYPERLINK("${url}", "${fileNameWithoutExtension}")`;
					sheet.getRange(fileRowIndex, uploadedFileCol).setFormula(hyperlinkFormula);
				} else {
					// Fallback: just store the URL as is
					sheet.getRange(fileRowIndex, uploadedFileCol).setValue(row.uploadedFile);
				}
			} catch (e) {
				// If there's any error accessing the file, just store the URL as is
				sheet.getRange(fileRowIndex, uploadedFileCol).setValue(row.uploadedFile);
			}
		}

		// Move to the next available slot
		nextSlotIndex++;
	});

	return { success: true };
}

// Function to check if the vendor name exists in the Vendor DB
function checkVendorName(vendorName) {
	try {
		const vendorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendor DB");
		if (!vendorSheet) {
			throw new Error("Vendor DB sheet not found");
		}

		const vendorData = vendorSheet.getDataRange().getValues();
		const vendorNames = vendorData.map((row) => row[0].toString().trim().toLowerCase()); // Assuming vendor names are in the first column

		// Check if the vendor name exists in the Vendor DB
		const vendorExists = vendorNames.includes(vendorName.trim().toLowerCase());

		return vendorExists;
	} catch (error) {
		Logger.log("Error in checkVendorName: " + error.toString());
		return false;
	}
}

// Function to search for IO Number directly
function searchIONumberDirect(ioNumber) {
	try {
		const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");
		if (!sheet) {
			throw new Error("Sheet 'Tool database' not found");
		}

		const dataRange = sheet.getDataRange();
		const values = dataRange.getValues();
		const headers = values[0];

		// Find the column indices for IO Number, Vendor Name, Program Name, and CE Number
		const colIONumber = headers.indexOf("IO Number");
		const colVendorName = headers.indexOf("Vendor Name");
		const colProgramName = headers.indexOf("Program Name");
		const colCENumber = headers.indexOf("Cost Estimate No.");

		if (colIONumber === -1 || colVendorName === -1 || colProgramName === -1 || colCENumber === -1) {
			throw new Error("Required columns not found");
		}

		// Filter rows based on IO Number (case-insensitive and trimmed)
		const filteredRows = values.filter((row) => row[colIONumber].toString().trim().toLowerCase() === ioNumber.trim().toLowerCase());

		if (filteredRows.length === 0) {
			return JSON.stringify({ found: false, error: "No data found for IO Number: " + ioNumber });
		}

		// If IO Number is unique, return the first row's data
		if (filteredRows.length === 1) {
			const row = filteredRows[0];
			const result = {
				"Time Stamp (Tool Generated)": row[headers.indexOf("Time Stamp (Tool Generated)")].toString(),
				"Email Address": row[headers.indexOf("Email Address")].toString(),
				"CE Start Date": row[headers.indexOf("CE Start Date")].toString(),
				"CE End Date": row[headers.indexOf("CE End Date")].toString(),
				"Date of Issue": row[headers.indexOf("Date of Issue")].toString(),
				"Program Name": row[headers.indexOf("Program Name")].toString(),
				"Vendor Name": row[headers.indexOf("Vendor Name")].toString(),
				"Signed by Vendor?": row[headers.indexOf("Signed by Vendor?")].toString(),
				"Globe Authorize Signatory": row[headers.indexOf("Globe Authorize Signatory")].toString(),
				"Payment Terms": row[headers.indexOf("Payment Terms")].toString(),
				"Cost Estimate No.": row[headers.indexOf("Cost Estimate No.")].toString(),
				Currency: row[headers.indexOf("Currency")].toString(),
				"Total Cost Estimate Amount (Vat-ex)": row[headers.indexOf("Total Cost Estimate Amount (Vat-ex)")].toString(),
				"GL Description": row[headers.indexOf("GL Description")].toString(),
				"IO Number": row[headers.indexOf("IO Number")].toString(),
				"IO BALANCE": row[headers.indexOf("IO BALANCE")].toString(),
				"Cost Center": row[headers.indexOf("Cost Center")].toString(),
				"Accrued?": row[headers.indexOf("Accrued?")].toString(),
				"CE Payment Status": row[headers.indexOf("CE Payment Status")].toString(),
				"Uploaded CE File": row[headers.indexOf("Uploaded CE File")].toString(),
				"SAP Ariba Reference Code 1": row[headers.indexOf("SAP Ariba Reference Code 1")].toString(),
				"SAP Ariba Reference Code 2": row[headers.indexOf("SAP Ariba Reference Code 2")].toString(),
				"SAP Ariba Reference Code 3": row[headers.indexOf("SAP Ariba Reference Code 3")].toString(),
				"SAP Ariba Reference Code 4": row[headers.indexOf("SAP Ariba Reference Code 4")].toString(),
				"SAP Ariba Reference Code 5": row[headers.indexOf("SAP Ariba Reference Code 5")].toString(),
			};
			return JSON.stringify({ found: true, data: result });
		} else {
			// If IO Number is not unique, return the list of non-unique rows
			const nonUniqueRows = filteredRows.map((row) => ({
				ioNumber: row[colIONumber].toString(),
				vendorName: row[colVendorName].toString(),
				programName: row[colProgramName].toString(),
				ceNumber: row[colCENumber].toString(),
			}));
			return JSON.stringify({ found: false, nonUniqueRows: nonUniqueRows });
		}
	} catch (error) {
		Logger.log("Error in searchIONumberDirect: " + error.toString());
		return JSON.stringify({ error: error.toString() });
	}
}

// Function to get row data by IO Number, Vendor Name, Program Name, and CE Number
function getRowDataByIONumberAndDetails(ioNumber, vendorName, programName, ceNumber) {
	try {
		const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tool database");
		if (!sheet) {
			throw new Error("Sheet 'Tool database' not found");
		}

		const dataRange = sheet.getDataRange();
		const values = dataRange.getValues();
		const headers = values[0];

		// Find the column indices for IO Number, Vendor Name, Program Name, and CE Number
		const colIONumber = headers.indexOf("IO Number");
		const colVendorName = headers.indexOf("Vendor Name");
		const colProgramName = headers.indexOf("Program Name");
		const colCENumber = headers.indexOf("Cost Estimate No.");

		if (colIONumber === -1 || colVendorName === -1 || colProgramName === -1 || colCENumber === -1) {
			throw new Error("Required columns not found");
		}

		// Find the row that matches the IO Number, Vendor Name, Program Name, and CE Number
		const row = values.find(
			(row) =>
				row[colIONumber].toString().trim().toLowerCase() === ioNumber.trim().toLowerCase() &&
				row[colVendorName].toString().trim().toLowerCase() === vendorName.trim().toLowerCase() &&
				row[colProgramName].toString().trim().toLowerCase() === programName.trim().toLowerCase() &&
				row[colCENumber].toString().trim().toLowerCase() === ceNumber.trim().toLowerCase()
		);

		if (!row) {
			throw new Error("Row not found");
		}

		const result = {
			"Time Stamp (Tool Generated)": row[headers.indexOf("Time Stamp (Tool Generated)")].toString(),
			"Email Address": row[headers.indexOf("Email Address")].toString(),
			"CE Start Date": row[headers.indexOf("CE Start Date")].toString(),
			"CE End Date": row[headers.indexOf("CE End Date")].toString(),
			"Date of Issue": row[headers.indexOf("Date of Issue")].toString(),
			"Program Name": row[headers.indexOf("Program Name")].toString(),
			"Vendor Name": row[headers.indexOf("Vendor Name")].toString(),
			"Signed by Vendor?": row[headers.indexOf("Signed by Vendor?")].toString(),
			"Globe Authorize Signatory": row[headers.indexOf("Globe Authorize Signatory")].toString(),
			"Payment Terms": row[headers.indexOf("Payment Terms")].toString(),
			"Cost Estimate No.": row[headers.indexOf("Cost Estimate No.")].toString(),
			Currency: row[headers.indexOf("Currency")].toString(),
			"Total Cost Estimate Amount (Vat-ex)": row[headers.indexOf("Total Cost Estimate Amount (Vat-ex)")].toString(),
			"GL Description": row[headers.indexOf("GL Description")].toString(),
			"IO Number": row[headers.indexOf("IO Number")].toString(),
			"IO BALANCE": row[headers.indexOf("IO BALANCE")].toString(),
			"Cost Center": row[headers.indexOf("Cost Center")].toString(),
			"Accrued?": row[headers.indexOf("Accrued?")].toString(),
			"CE Payment Status": row[headers.indexOf("CE Payment Status")].toString(),
			"Uploaded CE File": row[headers.indexOf("Uploaded CE File")].toString(),
			"SAP Ariba Reference Code 1": row[headers.indexOf("SAP Ariba Reference Code 1")].toString(),
			"SAP Ariba Reference Code 2": row[headers.indexOf("SAP Ariba Reference Code 2")].toString(),
			"SAP Ariba Reference Code 3": row[headers.indexOf("SAP Ariba Reference Code 3")].toString(),
			"SAP Ariba Reference Code 4": row[headers.indexOf("SAP Ariba Reference Code 4")].toString(),
			"SAP Ariba Reference Code 5": row[headers.indexOf("SAP Ariba Reference Code 5")].toString(),
		};

		return JSON.stringify(result);
	} catch (error) {
		Logger.log("Error in getRowDataByIONumberAndDetails: " + error.toString());
		return JSON.stringify({ error: error.toString() });
	}
}

// Automated Email for New Invoices added
function sendCENotificationEmail(fileName, userEmail, costCenter) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Auto Email");
	if (!sheet) {
		Logger.log("Email sheet not found.");
		return;
	}

	const headers = sheet
		.getRange(1, 1, 1, sheet.getLastColumn())
		.getValues()[0]
		.map((h) => h.trim());
	const emailColumnIndex = headers.indexOf("New Invoice Added") + 1;
	if (emailColumnIndex === 0) {
		Logger.log("New Invoice Added column not found.");
		return;
	}

	const emailAddresses = sheet
		.getRange(2, emailColumnIndex, sheet.getLastRow() - 1, 1)
		.getValues()
		.flat()
		.filter(String);
	if (emailAddresses.length === 0) {
		Logger.log("No emails found in the New Invoice Added column.");
		return;
	}

	// Get emails based on Cost Center
	const costCenterEmails = getEmailsByCostCenter(costCenter);
	const allEmails = [...new Set([...emailAddresses, ...costCenterEmails])]; // Combine and remove duplicates

	// Log the Cost Center and email recipients
	Logger.log(`Cost Center: ${costCenter}`);
	Logger.log(`Email recipients: ${allEmails.join(", ")}`);

	const subject = "New CE Invoice Added";
	const body = `
    <p>Dear Budget Team,</p>
    <p>A new Invoice has been added for File Name: <strong>${fileName}</strong>.</p>
    <p>To review the entry, please click the link below:</p>
    <p>
      <a href="https://docs.google.com/spreadsheets/d/1xgROcPRmkI6_dLfZmadbCI0Yv_0s_D7S23PIdNnkA40/edit?usp=sharing">
        View Tool Database
      </a>
    </p>
    <p>Best regards,<br>Your Budget Tool Team</p>
  `;

	MailApp.sendEmail({
		to: allEmails.join(","),
		subject: subject,
		htmlBody: body,
	});
}

// Helper function to send emails
function sendEmail(emailAddresses, subject, body) {
	if (emailAddresses.length > 0) {
		// Send email to all addresses
		MailApp.sendEmail({
			to: emailAddresses.join(","),
			subject: subject,
			htmlBody: body, // Use htmlBody for formatted content
		});
		console.log(`Notification email sent: ${subject}`);
	} else {
		console.error("No email addresses found.");
	}
}
