function initializeProcess(folderUrl, startRow, startCol) {
  if (!folderUrl || !startRow || !startCol) {
    throw new Error('Invalid input: folderUrl, startRow, and startCol are required.');
  }
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  
  PropertiesService.getScriptProperties().setProperties({
    'shouldStop': 'false',
    'folderUrl': folderUrl,
    'startRow': startRow,
    'startCol': startCol,
    'spreadsheetId': spreadsheetId,
    'sheetId': sheetId,
    'lastProcessedIndex': '-1'
  });
  
  Logger.log('Initialization started: folderUrl=' + folderUrl + ', startRow=' + startRow + ', startCol=' + startCol + ', spreadsheetId=' + spreadsheetId + ', sheetId=' + sheetId);
  return true;
}

function previewSortOrder() {
  // Reset necessary properties
  PropertiesService.getScriptProperties().setProperty('shouldStop', 'false');
  PropertiesService.getScriptProperties().setProperty('lastProcessedIndex', '-1');
  
  var folderUrl = PropertiesService.getScriptProperties().getProperty('folderUrl');
  if (!folderUrl) {
    throw new Error('Folder URL not set. Please initialize the process first.');
  }
  
  Logger.log('Previewing sort order for folder: ' + folderUrl);
  var fileList = collectFiles(folderUrl);
  sortFiles();
  return previewSortedFiles();
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
  return fileList;
}

function sortFiles() {
  Logger.log('Sorting files');
  var fileList = JSON.parse(PropertiesService.getScriptProperties().getProperty('fileList'));
  
  // Natural sort comparator
  function naturalCompare(a, b) {
    var ax = [], bx = [];
    a.name.replace(/(\d+)|(\D+)/g, function(_, $1, $2) { ax.push([$1 || Infinity, $2 || ""]) });
    b.name.replace(/(\d+)|(\D+)/g, function(_, $1, $2) { bx.push([$1 || Infinity, $2 || ""]) });
    
    while(ax.length && bx.length) {
      var an = ax.shift();
      var bn = bx.shift();
      var nn = (an[0] - bn[0]) || an[1].localeCompare(bn[1]);
      if(nn) return nn;
    }
    
    return ax.length - bx.length;
  }
  
  fileList.sort(naturalCompare);
  
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

function continueImageInsertion() {
  Logger.log('Continuing image insertion');
  var result = insertImagesFromUI();
  if (result === true) {
    // All images have been processed
    Logger.log('All images processed. Cleaning up.');
    cleanupAfterCompletion();
  } else if (result === false) {
    // More images to process, schedule next batch
    Logger.log('Scheduling next batch');
    ScriptApp.newTrigger('continueImageInsertion')
      .timeBased()
      .after(1000) // Wait 1 second before next execution
      .create();
  } else {
    // An error occurred
    Logger.log('Error occurred during image insertion: ' + result);
    cleanupAfterCompletion();
  }
}

function cleanupAfterCompletion() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'handleImageInsertion') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Delete all properties
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('Script complete, all properties deleted for new run.');
}

function handleImageInsertion() {
  try {
    var result = insertImagesFromUI();
    Logger.log("handleImageInsertion result: " + result);
    if (result === true) {
      // All images have been processed
      Logger.log('All images processed. Cleaning up.');
      cleanupAfterCompletion();
      return 'completed';
    } else if (result === false) {
      // More images to process, schedule next batch
      Logger.log('Scheduling next batch');
      ScriptApp.newTrigger('handleImageInsertion')
        .timeBased()
        .after(1000) // Wait 1 second before next execution
        .create();
      return 'continuing';
    } else if (result === 'stopped') {
      // Script was stopped by user
      Logger.log('Script stopped by user. Cleaning up.');
      cleanupAfterStop();
      return 'stopped';
    } else {
      // Unexpected result
      Logger.log('Unexpected result from insertImagesFromUI: ' + result);
      cleanupAfterCompletion();
      return 'error';
    }
  } catch (error) {
    Logger.log('Error in handleImageInsertion: ' + error.message);
    cleanupAfterCompletion();
    return 'error: ' + error.message;
  }
}

function insertImagesFromUI() {
  var startRow = parseInt(PropertiesService.getScriptProperties().getProperty('startRow'), 10);
  var startCol = parseInt(PropertiesService.getScriptProperties().getProperty('startCol'), 10);
  var batchSize = 75; // Process 75 images per execution
  var lastProcessedIndex = parseInt(PropertiesService.getScriptProperties().getProperty('lastProcessedIndex') || '-1', 10);
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('spreadsheetId');
  var sheetId = PropertiesService.getScriptProperties().getProperty('sheetId');

  if (isNaN(startRow) || isNaN(startCol)) {
    throw new Error('Invalid input: startRow and startCol must be numbers.');
  }

  Logger.log('UI Parser Started. Last processed index: ' + lastProcessedIndex);
  var folderUrl = PropertiesService.getScriptProperties().getProperty('folderUrl');
  var fileList = JSON.parse(PropertiesService.getScriptProperties().getProperty('fileList'));
  
  // Open the specific spreadsheet and sheet
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheets().filter(function(s) { return s.getSheetId() == sheetId; })[0];
  
  if (!sheet) {
    throw new Error('Sheet not found. spreadsheetId: ' + spreadsheetId + ', sheetId: ' + sheetId);
  }

  var rowIndex = startRow + lastProcessedIndex + 1;
  var colIndex = startCol;
  var totalFiles = fileList.length;
  var processedInThisExecution = 0;

  Logger.log('Starting batch processing from index ' + (lastProcessedIndex + 1) + ' at row ' + rowIndex);

  for (var i = lastProcessedIndex + 1; i < totalFiles && processedInThisExecution < batchSize; i++) {
    if (PropertiesService.getScriptProperties().getProperty('shouldStop') === 'true') {
      Logger.log('Script stopped by user during image insertion');
      return 'stopped';
    }
    try {
      var fileObj = fileList[i];
      var file = DriveApp.getFileById(fileObj.fileId);
      Logger.log('Processing file ' + (i + 1) + ' of ' + totalFiles + ': ' + file.getName());
      var resizedImage = resizeImage(fileObj.fileId, 800); // Resize to max width of 800px
      if (!resizedImage) {
        throw new Error('Failed to resize image');
      }
      var imageData = "data:image/jpeg;base64," + Utilities.base64Encode(resizedImage.getBytes());
      insertCellImage(sheet.getRange(rowIndex, colIndex), imageData, "Image " + (i + 1), "Image from frame " + fileObj.name);

      Logger.log('Inserted image ' + (i + 1) + ' of ' + totalFiles + ' at row ' + rowIndex);
      rowIndex++;
      processedInThisExecution++;
      PropertiesService.getScriptProperties().setProperty('lastProcessedIndex', i.toString());
    } catch (e) {
      Logger.log('Error processing image at index ' + i + ': ' + e.message);
      // Continue with the next image instead of stopping the entire process
    }
  }

  Logger.log('Batch completed. Processed ' + processedInThisExecution + ' images in this execution.');

  if (lastProcessedIndex + processedInThisExecution >= totalFiles - 1) {
    Logger.log('Image insertion completed');
    return true; // All images processed
  } else {
    Logger.log('More images to process. Last processed index: ' + PropertiesService.getScriptProperties().getProperty('lastProcessedIndex'));
    return false; // More images to process
  }
}

function insertCellImage(range, imageData, altTitle = "", altDescription = "") {
  try {
    let image = SpreadsheetApp
                    .newCellImage()
                    .setSourceUrl(imageData)
                    .setAltTextTitle(altTitle)
                    .setAltTextDescription(altDescription)
                    .build();
    range.setValue(image);
    Logger.log('Image inserted successfully');
  } catch (error) {
    Logger.log('Error inserting image: ' + error.message);
    throw error; // Rethrow the error to be caught in the calling function
  }
}

function getImageDimensions_(base64Image) {
  var decodedImage = Utilities.base64Decode(base64Image);
  var width = -1;
  var height = -1;

  for (var i = 0; i < decodedImage.length; i++) {
    if (decodedImage[i] == 0xFF && decodedImage[i+1] == 0xC0) {
      height = decodedImage[i+5] * 256 + decodedImage[i+6];
      width = decodedImage[i+7] * 256 + decodedImage[i+8];
      break;
    }
  }

  return {width: width, height: height};
}

function resizeImage(fileId, maxWidth) {
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var image = blob.getAs('image/jpeg');
  
  // Get image dimensions
  var imageData = Utilities.base64Encode(image.getBytes());
  var dimensions = getImageDimensions_(imageData);
  
  if (dimensions.width > maxWidth) {
    var newHeight = Math.round((maxWidth / dimensions.width) * dimensions.height);
    var resizedBlob = image.getAs('image/jpeg').setWidth(maxWidth).setHeight(newHeight);
    return resizedBlob;
  }
  
  return image;
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
  cleanupAfterStop();
  return true; // Indicate that stop process is complete
}

function showUI() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(650)
      .setHeight(505);
  SpreadsheetApp.getUi().showModalDialog(html, 'Insert Thumbnails');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
      .addItem('Insert Thumbnails from GDrive', 'showUI')
      .addToUi();
}

function cleanupAfterStop() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'handleImageInsertion') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Get all properties
  var properties = PropertiesService.getScriptProperties().getProperties();
  
  // Delete all properties except 'shouldStop'
  for (var key in properties) {
    if (key !== 'shouldStop') {
      PropertiesService.getScriptProperties().deleteProperty(key);
    }
  }
  
  // Ensure 'shouldStop' is set to 'true'
  PropertiesService.getScriptProperties().setProperty('shouldStop', 'true');
  
  Logger.log('All properties except shouldStop deleted, and shouldStop set to true');
}

function collectAndSortFiles(folderUrl) {
  try {
    collectFiles(folderUrl);
    sortFiles();
    return previewSortedFiles();
  } catch (error) {
    Logger.log('Error in collectAndSortFiles: ' + error.message);
    return ['Error: ' + error.message];
  }
}