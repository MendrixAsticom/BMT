function checkToolDatabase() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName("Tool database");
	const logSheet = ss.getSheetByName("EmailLog") || ss.insertSheet("EmailLog");
	const logData = logSheet.getDataRange().getValues();

	const data = sheet.getDataRange().getValues(); // Get all data at once
	const today = new Date();
	const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");
	const emailList = {}; // Object to store CE numbers by email

	// Loop through the data starting from row 2 (to skip headers)
	for (let i = 1; i < data.length; i++) {
		const email = data[i][4]; // Column E (Email Address)
		const projectEndDate = data[i][6]; // Column G (CE End Date)
		const paymentStatus = data[i][25]; // Column Z (CE Payment Status)
		const costEstimateNo = data[i][14]; // Column O (Cost Estimate No.)

		// Ensure that the projectEndDate is a valid date and in the past
		if (projectEndDate && projectEndDate instanceof Date && projectEndDate < today && paymentStatus === "Not Yet Fully Paid") {
			const daysPastDue = calculateDaysPast(projectEndDate, today);

			// Check if this CE Number was already emailed today
			const alreadyEmailed = logData.some((row) => row[0] === email && row[1] === costEstimateNo && row[2] === todayFormatted);

			if (!alreadyEmailed) {
				// Group CE Numbers and other details by email
				if (!emailList[email]) {
					emailList[email] = []; // Create an array for this email if not already existing
				}

				// Push the CE details for this email
				emailList[email].push({
					ceNo: costEstimateNo,
					endDate: projectEndDate,
					daysPast: daysPastDue,
				});

				// Log the email being sent today to avoid resending
				logSheet.appendRow([email, costEstimateNo, todayFormatted]);
			}
		}
	}

	// Send a single email for each email address containing all CE Numbers
	for (let email in emailList) {
		const ceDetails = emailList[email];
		if (ceDetails.length > 0) {
			// If there are any CE Numbers to send
			sendEmailToProponent(email, ceDetails);
		}
	}
}

// Function to manually run the check
function manualRunCheckToolDatabase() {
	checkToolDatabase();
}

//Modify Email Template for Past Due Invoices
function sendEmailToProponent(email, ceDetails) {
	const subject = "Action Required: Past Due Invoices for Completed CES";

	const body = `
    <p style="font-size: 14px;">Hi,</p>
    <p style="font-size: 14px;">
      This is a reminder that the following CES, for which you are the proponent, have passed their Project End Date, and their invoices are pending submission:
    </p>
    <table style="border-collapse: collapse; font-size: 14px;">
      <thead>
        <tr>
          <th style="border: 1px solid black; padding: 5px;">CE Number</th>
          <th style="border: 1px solid black; padding: 5px;">CE End Date</th>
          <th style="border: 1px solid black; padding: 5px;">Days Past Due</th>
        </tr>
      </thead>
      <tbody>
        ${ceDetails
			.map(
				(ce) => `
        <tr>
          <td style="border: 1px solid black; padding: 5px;">${ce.ceNo}</td>
          <td style="border: 1px solid black; padding: 5px;">${Utilities.formatDate(ce.endDate, Session.getScriptTimeZone(), "MM/dd/yyyy")}</td>
          <td style="border: 1px solid black; padding: 5px;">${ce.daysPast}</td>
        </tr>`
			)
			.join("")}
      </tbody>
    </table>
    <p style="font-size: 14px;">
      Please send your corresponding invoices for these completed CES as soon as possible.
    </p>
    <p style="font-size: 14px;">
      Access this <a href="https://script.google.com/a/macros/globe.com.ph/s/AKfycbyHBZXqhD5Dzl5GQGoKfQEzd9rdTZP0-fUe3Hzk-xc/dev">link</a> to add the invoices for the corresponding CE Numbers.
    </p>
    <p style="font-size: 14px;">Thank you for your attention to this matter.</p>
  `;

	// Send the email using HTML body
	MailApp.sendEmail({
		to: email,
		subject: subject,
		htmlBody: body,
	});
}

// Helper function to calculate days past the end date
function calculateDaysPast(endDate, today) {
	const oneDay = 24 * 60 * 60 * 1000;
	return Math.floor((today - endDate) / oneDay);
}

// Function to manually trigger the check (for testing purposes)
function manualRunCheckToolDatabase() {
	checkToolDatabase(); // Call the main function
}
