function listPdfShareLinks() {
  var folderId = "1gkKH9LscoQXT_P4ZRYEj8jQ-Sx5ZqDPS";  
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(["File Name", "Shareable Link"]);

  while (files.hasNext()) {
    var file = files.next();

    if (file.getMimeType() === "application/pdf") {
      
      // Ensure file is shared
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      var name = file.getName();
      var link = file.getUrl();

      sheet.appendRow([name, link]);
    }
  }

  SpreadsheetApp.getUi().alert("Done! PDF links created.");
}
