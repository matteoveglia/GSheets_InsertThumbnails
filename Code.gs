function initializeProcess(folderUrl, startRow, startCol) {
  PropertiesService.getScriptProperties().setProperty('shouldStop', 'false'); // Reset stop flag
  PropertiesService.getScriptProperties().setProperty('folderUrl', folderUrl);
  PropertiesService.getScriptProperties().setProperty('startRow', startRow);
  PropertiesService.getScriptProperties().setProperty('startCol', startCol);
  Logger.log('Initialization started: folderUrl=' + folderUrl + ', startRow=' + startRow + ', startCol=' + startCol);
  return true;
}

function collectFiles(folderUrl) {
  Logger.log('Collecting folder URL');
  var folderId = folderUrl.match(/[-\w]{25,}/)[0]; // Extract folder ID from URL
  var folder = DriveApp.getFolderById(folderId);
  Logger.log('Collecting files from folder');
  var files = folder.getFiles();
  Logger.log('Files collected');
  var fileList = [];

  while (files.hasNext()) {
    if (PropertiesService.getScriptProperties().getProperty('shouldStop') === 'true') {
      Logger.log('Script stopped by user during file collection');
      return;
    }
    var file = files.next();
    var fileName = file.getName();
    var numberPart = fileName.match(/\d+/); // Extract numbers from the filename
    if (numberPart) {
      fileList.push({ fileId: file.getId(), number: parseInt(numberPart[0], 10) });
    }
  }
  PropertiesService.getScriptProperties().setProperty('fileList', JSON.stringify(fileList));
  Logger.log('Collected ' + fileList.length + ' files');
  return true;
}

function sortFiles() {
  Logger.log('Sorting files');
  var fileList = JSON.parse(PropertiesService.getScriptProperties().getProperty('fileList'));
  fileList.sort(function(a, b) {
    return a.number - b.number;
  });
  PropertiesService.getScriptProperties().setProperty('fileList', JSON.stringify(fileList));
  Logger.log('Files sorted');
  return true;
}

function insertImagesFromUI() {
  Logger.log('UI Parser Started');
  var folderUrl = PropertiesService.getScriptProperties().getProperty('folderUrl');
  var startRow = parseInt(PropertiesService.getScriptProperties().getProperty('startRow'));
  var startCol = parseInt(PropertiesService.getScriptProperties().getProperty('startCol'));
  Logger.log('UI input received: folderUrl=' + folderUrl + ', startRow=' + startRow + ', startCol=' + startCol);
  
  var folderId = folderUrl.match(/[-\w]{25,}/)[0];
  insertResizedImages(folderId, startRow, startCol);
  Logger.log('UI Parser Completed');
  return true;
}

function insertResizedImages(folderId, startRow, startCol) {
  Logger.log('Image Inserter Started');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var fileList = JSON.parse(PropertiesService.getScriptProperties().getProperty('fileList'));
  var totalFiles = fileList.length;
  Logger.log('Total files to process: ' + totalFiles);

  var rowIndex = startRow; // Starting row
  var colIndex = startCol; // Starting column

  for (var i = 0; i < fileList.length; i++) {
    if (PropertiesService.getScriptProperties().getProperty('shouldStop') === 'true') {
      Logger.log('Script stopped by user during image insertion');
      return;
    }
    try {
      var fileObj = fileList[i];
      var file = DriveApp.getFileById(fileObj.fileId);
      Logger.log('Processing file ' + (i + 1) + ' of ' + totalFiles + ': ' + file.getName());
      var resizedImage = ImgApp.doResize(fileObj.fileId, 800); // Adjust width as necessary
      var blob = resizedImage.blob;
      var imageData = "data:image/jpeg;base64," + Utilities.base64Encode(blob.getBytes());
      insertCellImage(sheet.getRange(rowIndex, colIndex), imageData, "Image " + (i + 1), "Image from frame " + fileObj.number); // Insert image into cell

      Logger.log('Inserted image ' + (i + 1) + ' of ' + totalFiles + ' at row ' + rowIndex);
      rowIndex++;
    } catch (e) {
      Logger.log('Error inserting image: ' + e.message);
    }
  }
  Logger.log('Image insertion completed');
  resetForm(); // Reset form after completion
  return true;
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

function stopImageInsertion() {
  PropertiesService.getScriptProperties().setProperty('shouldStop', 'true');
  Logger.log('Stop flag set');
}

function showUI() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(650)
      .setHeight(505);
  SpreadsheetApp.getUi().showModalDialog(html, 'Insert Thumbnails');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Insert Thumbnails from GDrive', 'showUI')
      .addToUi();
}

function resetForm() {
  var ui = SpreadsheetApp.getUi();
  //ui.alert('Image insertion completed or stopped by user.');
}