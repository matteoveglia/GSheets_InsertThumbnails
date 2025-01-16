const BATCH_SIZE = 10;
const MAX_IMAGE_WIDTH = 800;
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const CHUNK_SIZE = 50;
const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
const START_TIME = Date.now();
const CACHE_EXPIRATION = 21600;
const CACHE = CacheService.getScriptCache();
const TRIGGER_DELAY = 500;

function isTimeUp() {
  return Date.now() - START_TIME > MAX_EXECUTION_TIME;
}

function getCachedImage(fileId) {
  return CACHE.get('img_' + fileId);
}

function setCachedImage(fileId, imageData) {
  CACHE.put('img_' + fileId, imageData, CACHE_EXPIRATION);
}

function initializeProcess(folderUrl, startRow, startCol) {
  if (!folderUrl || !startRow || !startCol) {
    throw new Error('Invalid input: folderUrl, startRow, and startCol are required.');
  }
  const props = {
    'shouldStop': 'false',
    'folderUrl': folderUrl,
    'startRow': startRow,
    'startCol': startCol,
    'spreadsheetId': SpreadsheetApp.getActiveSpreadsheet().getId(),
    'sheetId': SpreadsheetApp.getActiveSheet().getSheetId(),
    'lastProcessedIndex': '-1',
    'processingBatch': 'false'
  };
  
  SCRIPT_PROPERTIES.setProperties(props);
  Logger.log('Initialization complete with properties: ' + JSON.stringify(props));
  return true;
}

function previewSortOrder(folderUrl) { 
  SCRIPT_PROPERTIES.setProperty('shouldStop', 'false');
  SCRIPT_PROPERTIES.setProperty('lastProcessedIndex', '-1');
  
  if (!folderUrl) {
    throw new Error('Folder URL is required');
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
  
  Logger.log('Starting collectFiles with URL: ' + folderUrl);
  
  Logger.log('Current properties at start of collectFiles: ' + JSON.stringify(SCRIPT_PROPERTIES.getProperties()));

  try {
    const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
    if (!folderIdMatch) {
      Logger.log('Failed to extract folder ID from URL');
      throw new Error('Invalid folder URL format');
    }
    
    const folderId = folderIdMatch[0];
    Logger.log('Extracted folder ID: ' + folderId);
    
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
      Logger.log('Successfully accessed folder: ' + folder.getName());
    } catch (folderError) {
      Logger.log('Error accessing folder: ' + folderError.toString());
      throw new Error('Could not access the specified folder. Please check the URL and permissions.');
    }
    
    let files;
    try {
      files = folder.getFiles();
      Logger.log('Successfully got files iterator');
    } catch (filesError) {
      Logger.log('Error getting files: ' + filesError.toString());
      throw new Error('Could not retrieve files from the folder');
    }
    
    const fileList = [];
    let fileCount = 0;
    
    while (files.hasNext()) {
      if (isTimeUp()) {
        Logger.log('Approaching time limit during file collection');
        if (fileList.length > 0) {
          SCRIPT_PROPERTIES.setProperty('fileList', JSON.stringify(fileList));
          SCRIPT_PROPERTIES.setProperty('partialCollection', 'true');
          ScriptApp.newTrigger('continueFileCollection')
            .timeBased()
            .after(1000)
            .create();
        }
        return fileList;
      }

      if (SCRIPT_PROPERTIES.getProperty('shouldStop') === 'true') {
        Logger.log('Script stopped by user during file collection');
        return null;
      }
      
      const chunk = [];
      while (files.hasNext() && chunk.length < CHUNK_SIZE) {
        const file = files.next();
        chunk.push({
          fileId: file.getId(),
          name: file.getName()
        });
        fileCount++;
      }
      
      fileList.push(...chunk);
      
      if (fileCount % CHUNK_SIZE === 0) {
        SCRIPT_PROPERTIES.setProperty('fileList', JSON.stringify(fileList));
      }
      
      Utilities.sleep(20);
    }
    
    Logger.log('Total files collected: ' + fileCount);
    
    if (fileList.length === 0) {
      Logger.log('No files found in folder');
      throw new Error('No files found in the specified folder');
    }
    
    SCRIPT_PROPERTIES.setProperty('fileList', JSON.stringify(fileList));
    Logger.log('Successfully stored file list with ' + fileList.length + ' files');
    
    return fileList;
    
  } catch (error) {
    Logger.log('Final error in collectFiles: ' + error.toString());
    throw error;
  }
}

function continueFileCollection() {
  try {
    const props = SCRIPT_PROPERTIES.getProperties();
    const folderUrl = props.folderUrl;
    const existingFiles = JSON.parse(props.fileList || '[]');
    
    const newFiles = collectFiles(folderUrl);
    if (newFiles) {
      const combinedFiles = existingFiles.concat(newFiles);
      SCRIPT_PROPERTIES.setProperty('fileList', JSON.stringify(combinedFiles));
    }
    
    if (props.partialCollection !== 'true') {
      sortFiles();
    }
  } catch (error) {
    Logger.log('Error in continueFileCollection: ' + error.toString());
    cleanupAfterCompletion();
  }
}

function sortFiles() {
  Logger.log('Starting sortFiles function');
  
  try {
    const fileListStr = SCRIPT_PROPERTIES.getProperty('fileList');
    Logger.log('Retrieved fileList string: ' + (fileListStr ? 'non-null' : 'null'));
    
    if (!fileListStr) {
      throw new Error('No file list found in properties');
    }
    
    let fileList = JSON.parse(fileListStr);
    Logger.log('Parsed fileList, length: ' + fileList.length);
    
    if (!Array.isArray(fileList)) {
      throw new Error('File list is not an array');
    }
    
    fileList = fileList.filter(file => file && file.name);
    Logger.log('Filtered fileList, new length: ' + fileList.length);
    
    function naturalCompare(a, b) {
      if (!a || !b || !a.name || !b.name) {
        Logger.log('Invalid comparison objects:', a, b);
        return 0;
      }
      
      try {
        let ax = [], bx = [];
        
        a.name.replace(/(\d+)|(\D+)/g, function(_, $1, $2) { 
          ax.push([$1 || Infinity, $2 || ""]) 
        });
        
        b.name.replace(/(\d+)|(\D+)/g, function(_, $1, $2) { 
          bx.push([$1 || Infinity, $2 || ""]) 
        });
        
        while(ax.length && bx.length) {
          var an = ax.shift();
          var bn = bx.shift();
          var nn = (an[0] - bn[0]) || an[1].localeCompare(bn[1]);
          if(nn) return nn;
        }
        
        return ax.length - bx.length;
      } catch (e) {
        Logger.log('Error in comparison:', e);
        return 0;
      }
    }
    
    const chunkSize = 1000;
    for (let i = 0; i < fileList.length; i += chunkSize) {
      const chunk = fileList.slice(i, i + chunkSize);
      chunk.sort(naturalCompare);
      fileList.splice(i, chunk.length, ...chunk);
      
      if (i + chunkSize < fileList.length) {
        Utilities.sleep(50);
      }
    }
    
    Logger.log('Sort completed successfully');
    
    SCRIPT_PROPERTIES.setProperty('fileList', JSON.stringify(fileList));
    return true;
    
  } catch (error) {
    Logger.log('Error in sortFiles: ' + error.message);
    throw error;
  }
}

function previewSortedFiles() {
  Logger.log('Previewing sorted files');
  var fileList = JSON.parse(SCRIPT_PROPERTIES.getProperty('fileList'));
  return fileList.map(function(file) {
    return file.name;
  });
}

function cleanupAfterCompletion() {
  try {
    cleanupExistingTriggers();
    SCRIPT_PROPERTIES.deleteAllProperties();
    if (CACHE) {
      CACHE.removeAll(['img_*']);
    }
    Logger.log('Cleanup completed successfully');
  } catch (error) {
    Logger.log('Error during cleanup: ' + error.message);
  } finally {
    // Ensure these flags are cleared even if there's an error
    SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
    SCRIPT_PROPERTIES.setProperty('shouldStop', 'false');
  }
}

function createContinuationPoint() {
  if (SCRIPT_PROPERTIES.getProperty('shouldStop') === 'true') {
    Logger.log('Process stopped, skipping continuation');
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 10px;
            color: #333;
          }
          .message {
            font-size: 14px;
            line-height: 1.4;
          }
        </style>
      </head>
      <body>
        <div class="message">
          As there are more than 10 images batch mode has been activated. Processing will continue automatically.
        </div>
        <script>
          window.onload = function() {
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                console.error(error);
                google.script.host.close();
              })
              .handleImageInsertion();
          }
        </script>
      </body>
    </html>
  `)
  .setTitle('Batch Mode Activated')
  .setWidth(400)
  .setHeight(80);

  SpreadsheetApp.getActiveSpreadsheet().show(html);
}

function handleImageInsertion() {
  try {
    // Check stop flag first
    if (SCRIPT_PROPERTIES.getProperty('shouldStop') === 'true') {
      Logger.log('Process is stopped, skipping execution');
      return 'stopped';
    }

    const processingBatch = SCRIPT_PROPERTIES.getProperty('processingBatch');
    Logger.log('Current processing batch status: ' + processingBatch);
    
    if (processingBatch === 'true') {
      Logger.log('Another batch is currently processing, skipping this execution');
      return 'skipped';
    }

    SCRIPT_PROPERTIES.setProperty('processingBatch', 'true');
    Logger.log('Set processingBatch to true');

    cleanupExistingTriggers();
    
    const result = insertImagesFromUI();
    Logger.log('insertImagesFromUI result: ' + result);
    
    if (result === false) {
      SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
      Utilities.sleep(100); // Small delay before next batch
      createContinuationPoint();
      return 'continuing';
    }
    
    if (result === true) {
      cleanupAfterCompletion();
      return 'completed';
    }
    
    if (result === 'stopped') {
      cleanupAfterStop();
      return 'stopped';
    }

    SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
    return 'error';
  } catch (error) {
    Logger.log('Error in handleImageInsertion: ' + error.message);
    SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
    cleanupAfterCompletion();
    return 'error: ' + error.message;
  }
}

function insertImagesFromUI() {
  const props = SCRIPT_PROPERTIES.getProperties();
  const startRow = parseInt(props.startRow, 10);
  const startCol = parseInt(props.startCol, 10);
  const lastProcessedIndex = parseInt(props.lastProcessedIndex || '-1', 10);
  const spreadsheetId = props.spreadsheetId;
  const sheetId = props.sheetId;

  if (isNaN(startRow) || isNaN(startCol)) {
    throw new Error('Invalid input: startRow and startCol must be numbers.');
  }

  Logger.log('UI Parser Started. Last processed index: ' + lastProcessedIndex);
  var folderUrl = props.folderUrl;
  var fileList = JSON.parse(props.fileList);
  
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheets().filter(function(s) { return s.getSheetId() == sheetId; })[0];
  
  if (!sheet) {
    throw new Error('Sheet not found. spreadsheetId: ' + spreadsheetId + ', sheetId: ' + sheetId);
  }

  var rowIndex = startRow + lastProcessedIndex + 1;
  var totalFiles = fileList.length;
  var processedInThisExecution = 0;
  const maxRetries = 3;

  for (let i = lastProcessedIndex + 1; i < totalFiles && processedInThisExecution < BATCH_SIZE; i++) {
    // Check stop flag before each image
    if (SCRIPT_PROPERTIES.getProperty('shouldStop') === 'true') {
      Logger.log('Stop requested, halting batch immediately');
      return 'stopped';
    }

    if (isTimeUp()) {
      Logger.log('Approaching time limit, scheduling next batch');
      SCRIPT_PROPERTIES.setProperty('lastProcessedIndex', i.toString());
      return false;
    }

    let retryCount = 0;
    while (retryCount < maxRetries) {
      try {
        // Check stop flag before each attempt
        if (SCRIPT_PROPERTIES.getProperty('shouldStop') === 'true') {
          return 'stopped';
        }

        const fileObj = fileList[i];
        const file = DriveApp.getFileById(fileObj.fileId);
        Logger.log('Processing file ' + (i + 1) + ' of ' + totalFiles + ': ' + file.getName());
        
        const resizedImage = resizeImage(fileObj.fileId, MAX_IMAGE_WIDTH);
        if (!resizedImage) {
          throw new Error('Failed to resize image');
        }

        const imageBytes = resizedImage.getBytes();
        if (!imageBytes || imageBytes.length === 0) {
          throw new Error('Invalid image data');
        }

        const imageData = "data:image/jpeg;base64," + Utilities.base64Encode(imageBytes);
        const range = sheet.getRange(rowIndex, startCol);
        
        insertCellImage(range, imageData, "Image " + fileObj.name, "Image from frame " + fileObj.name);

        rowIndex++;
        processedInThisExecution++;
        SCRIPT_PROPERTIES.setProperty('lastProcessedIndex', i.toString());
        
        if (resizedImage.getBytes) {
          resizedImage.getBytes().length = 0;
        }
        
        if (processedInThisExecution % 5 === 0) {
          if (isTimeUp()) {
            SCRIPT_PROPERTIES.setProperty('lastProcessedIndex', i.toString());
            return false;
          }
        }
        
        if (processedInThisExecution % 5 === 0) {
          Utilities.sleep(50);
        }
        break;
      } catch (e) {
        Logger.log(`Error processing image at index ${i}, attempt ${retryCount + 1}: ${e.message}`);
        retryCount++;
        if (retryCount === maxRetries) {
          Logger.log(`Failed to process image at index ${i} after ${maxRetries} attempts`);
        } else {
          Utilities.sleep(1000);
        }
      }
    }
  }

  Logger.log('Batch completed. Processed ' + processedInThisExecution + ' images in this execution.');

  if (lastProcessedIndex + processedInThisExecution >= totalFiles - 1) {
    Logger.log('Image insertion completed');
    SCRIPT_PROPERTIES.setProperty('processingBatch', 'false'); // Ensure batch flag is cleared
    return true;
  } else {
    Logger.log('More images to process. Last processed index: ' + SCRIPT_PROPERTIES.getProperty('lastProcessedIndex'));
    return false;
  }
}

function insertCellImage(range, imageData, altTitle = "", altDescription = "") {
  try {
    if (!imageData || !imageData.startsWith('data:image')) {
      throw new Error('Invalid image data format');
    }

    let image = SpreadsheetApp
                    .newCellImage()
                    .setSourceUrl(imageData)
                    .setAltTextTitle(altTitle)
                    .setAltTextDescription(altDescription)
                    .build();
    
    range.setValue(image);
    Logger.log('Image inserted successfully at ' + range.getA1Notation());
  } catch (error) {
    Logger.log('Error inserting image: ' + error.message);
    throw error;
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
  
  var imageData = Utilities.base64Encode(image.getBytes());
  var dimensions = getImageDimensions_(imageData);
  
  if (dimensions.width > maxWidth) {
    var newHeight = Math.round((maxWidth / dimensions.width) * dimensions.height);
    var resizedBlob = image.getAs('image/jpeg').setWidth(maxWidth).setHeight(newHeight);
    return resizedBlob;
  }
  
  return image;
}

function stopImageInsertion() {
  SCRIPT_PROPERTIES.setProperty('shouldStop', 'true');
  SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
  Logger.log('Stop flag set and processing batch cleared');
  cleanupAfterStop();
  return true;
}

function showUI() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(650)
      .setHeight(540);
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
  
  var properties = SCRIPT_PROPERTIES.getProperties();
  
  for (var key in properties) {
    if (key !== 'shouldStop') {
      SCRIPT_PROPERTIES.deleteProperty(key);
    }
  }
  
  SCRIPT_PROPERTIES.setProperty('shouldStop', 'true');
  SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
  Logger.log('All properties except shouldStop deleted, and shouldStop set to true');
}

function collectAndSortFiles(folderUrl) {
  try {
    Logger.log('Starting collectAndSortFiles');
    
    SCRIPT_PROPERTIES.deleteProperty('fileList');
    
    const files = collectFiles(folderUrl);
    Logger.log('collectFiles returned: ' + (files ? files.length + ' files' : 'null'));
    
    if (!files || files.length === 0) {
      return ['No files were found in the specified folder'];
    }
    
    const sortResult = sortFiles();
    Logger.log('sortFiles returned: ' + sortResult);
    
    const preview = previewSortedFiles();
    Logger.log('Generated preview with ' + preview.length + ' items');
    
    return preview;
    
  } catch (error) {
    Logger.log('Error in collectAndSortFiles: ' + error.toString());
    return ['Error: ' + error.message];
  }
}

function cleanupExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'handleImageInsertion') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}