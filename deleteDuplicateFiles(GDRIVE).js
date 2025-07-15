function deleteDuplicateFilesInFolder() {
	// Replace with the folder ID you want to check for duplicates
	var folderId = "1T6RdopgpBGpr7dH0SEoe1BBavDDP7Yj-";
	var folder = DriveApp.getFolderById(folderId);
	var files = folder.getFiles();

	var fileMap = {}; // To keep track of filenames and file IDs

	while (files.hasNext()) {
		var file = files.next();
		var fileName = file.getName();
		var fileId = file.getId();

		// If the filename already exists in the map, it's a duplicate
		if (fileMap[fileName]) {
			// Delete the duplicate file
			DriveApp.getFileById(fileId).setTrashed(true);
			Logger.log("Deleted duplicate file: " + fileName);
		} else {
			// Store the first occurrence of each file
			fileMap[fileName] = fileId;
		}
	}

	Logger.log("Duplicate file removal completed.");
}
