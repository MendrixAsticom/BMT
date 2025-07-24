const testBulkInvoiceData = [
	{
		invoiceEmail: "mendrix.manlangit@gsupport.com.ph",
		invoiceVendorName: "Inv-LTE",
		invoiceNumber: "32423",
		invoiceAmount: "234,234",
		invoiceTimestamp: "7/23/2025",
		accrualRefDoc: "",
		invoiceDate: "7/3/2025",
		invoiceDueDate: "8/2/2025",
		uploadedFile: "https://drive.google.com/file/d/1SnkAgaUE7VawgJXRxSbRxGujLrh64j8M/view?usp=drivesdk",
	},
];

function testFunction() {
	Logger.log("test start");
	sendNotificationCEInvoiceEmail("example.txt", "mendrix.manlangit@gsupport.com.ph");
	Logger.log("test end");
}

function testFunctionCENumber() {
	Logger.log("test start");
	Logger.log(searchCENumber("nymber", "CMB10CBY2024", "SDA MANILA CORP/MARISKA M. BAUTISTA"));
	Logger.log("test end");
}

function testBulkInvoice() {
	Logger.log("test start");
	saveBulkInvoiceData("TEST_KALI_CE_021025_01_CINEMA_Corrected_1752858267738", "CMB10CBY2024", testBulkInvoiceData, 501);
	Logger.log("test end");
}

function testSearchVendor() {
	Logger.log("test start");
	Logger.log(checkVendorName("vendor"));
	Logger.log("test end");
}
