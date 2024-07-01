function initializeProcess(folderUrl, startRow, startCol) {
  if (!folderUrl || !startRow || !startCol) {
    throw new Error('Invalid input: folderUrl, startRow, and startCol are required.');
  }
  PropertiesService.getScriptProperties().setProperty('shouldStop', 'false'); // Reset stop flag
  PropertiesService.getScriptProperties().setProperty('folderUrl', folderUrl);
  PropertiesService.getScriptProperties().setProperty('startRow', startRow);
  PropertiesService.getScriptProperties().setProperty('startCol', startCol);
  Logger.log('Initialization started: folderUrl=' + folderUrl + ', startRow=' + startRow + ', startCol=' + startCol);
  return true;
}

function collectFiles(folderUrl) {
  if (!folderUrl) {
    throw new Error('Invalid input: folderUrl is required.');
  }
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
    fileList.push({ fileId: file.getId(), name: fileName });
  }
  PropertiesService.getScriptProperties().setProperty('fileList', JSON.stringify(fileList));
  Logger.log('Collected ' + fileList.length + ' files');
  return true;
}

function sortFiles() {
  Logger.log('Sorting files');
  var fileList = JSON.parse(PropertiesService.getScriptProperties().getProperty('fileList'));
  fileList.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
  PropertiesService.getScriptProperties().setProperty('fileList', JSON.stringify(fileList));
  Logger.log('Files sorted');
  return true;
}

function previewSortedFiles() {
  Logger.log('Previewing sorted files');
  var fileList = JSON.parse(PropertiesService.getScriptProperties().getProperty('fileList'));
  return fileList.map(function(file) {
    return file.name;
  });
}

function insertImagesFromUI() {
  var startRow = parseInt(PropertiesService.getScriptProperties().getProperty('startRow'), 10);
  var startCol = parseInt(PropertiesService.getScriptProperties().getProperty('startCol'), 10);

  if (isNaN(startRow) || isNaN(startCol)) {
    throw new Error('Invalid input: startRow and startCol must be numbers.');
  }

  Logger.log('UI Parser Started');
  var folderUrl = PropertiesService.getScriptProperties().getProperty('folderUrl');
  var fileList = JSON.parse(PropertiesService.getScriptProperties().getProperty('fileList'));
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var rowIndex = startRow;
  var colIndex = startCol;
  var totalFiles = fileList.length;

  for (var i = 0; i < totalFiles; i++) {
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
      insertCellImage(sheet.getRange(rowIndex, colIndex), imageData, "Image " + (i + 1), "Image from frame " + fileObj.name); // Insert image into cell

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
  var html = HtmlService.createHtmlOutputFromFile('index')
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

function collectAndSortFiles(folderUrl) {
  collectFiles(folderUrl);
  sortFiles();
  return previewSortedFiles();
}