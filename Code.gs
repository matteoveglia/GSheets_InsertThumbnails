function insertResizedImages(folderId, startRow, startCol) {
  PropertiesService.getScriptProperties().setProperty('shouldStop', 'false'); // Reset stop flag
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var fileList = [];

  // Collect files and extract the numerical part of the filenames
  while (files.hasNext()) {
    if (PropertiesService.getScriptProperties().getProperty('shouldStop') === 'true') {
      Logger.log('Script stopped by user during file collection');
      return;
    }
    var file = files.next();
    var fileName = file.getName();
    var numberPart = fileName.match(/\d+/); // Extract numbers from the filename
    if (numberPart) {
      fileList.push({ file: file, number: parseInt(numberPart[0], 10) });
    }
  }

  // Sort files based on the numerical part
  fileList.sort(function(a, b) {
    return a.number - b.number;
  });

  var totalFiles = fileList.length;
  var rowIndex = startRow; // Starting row
  var colIndex = startCol; // Starting column

  for (var i = 0; i < fileList.length; i++) {
    if (PropertiesService.getScriptProperties().getProperty('shouldStop') === 'true') {
      Logger.log('Script stopped by user during image insertion');
      return;
    }
    try {
      var fileObj = fileList[i];
      var resizedImage = ImgApp.doResize(fileObj.file.getId(), 800); // Adjust width as necessary
      var blob = resizedImage.blob;
      var imageData = "data:image/jpeg;base64," + Utilities.base64Encode(blob.getBytes());
      insertCellImage(sheet.getRange(rowIndex, colIndex), imageData, "Image " + (i + 1), "Image from frame " + fileObj.number); // Insert image into cell
      
      Logger.log('Inserted image ' + (i + 1) + ' of ' + totalFiles + ' at row ' + rowIndex);
      rowIndex++;
    } catch (e) {
      Logger.log('Error inserting image: ' + e.message);
    }
  }
}

function insertCellImage(range, imageData, altTitle = "", altDescription = "") {
  let image = SpreadsheetApp
                  .newCellImage()
                  .setSourceUrl(imageData)
                  .setAltTextTitle(altTitle)
                  .setAltTextDescription(altDescription)
                  .build();
  range.setValue(image);
}

function insertImagesFromUI(folderUrl, startRow, startCol) {
  var folderId = folderUrl.match(/[-\w]{25,}/);
  if (folderId) {
    insertResizedImages(folderId[0], parseInt(startRow), parseInt(startCol));
  } else {
    SpreadsheetApp.getUi().alert('Invalid folder URL');
  }
}

function stopImageInsertion() {
  PropertiesService.getScriptProperties().setProperty('shouldStop', 'true');
  Logger.log('Stop flag set');
}

function showUI() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Insert Thumbnails');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Insert Thumbnails from GDrive', 'showUI')
      .addToUi();
}