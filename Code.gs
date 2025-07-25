// ============================================================================
// CONSTANTS & CONFIGURATION
// ============================================================================
const BATCH_SIZE = 10;
const MAX_IMAGE_WIDTH = 800;
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const CHUNK_SIZE = 50;
const MAX_EXECUTION_TIME = 4.5 * 60 * 1000;
const START_TIME = Date.now();
const CACHE_EXPIRATION = 21600;
const CACHE = CacheService.getScriptCache();
const TRIGGER_DELAY = 500;

// Memory management constants
const MEMORY_CHECK_INTERVAL = 5; // Check memory every 5 images
const GARBAGE_COLLECTION_HINT_DELAY = 10; // ms delay for GC hint

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================
function isTimeUp() {
  return Date.now() - START_TIME > MAX_EXECUTION_TIME;
}

function getCachedImage(fileId) {
  return CACHE.get('img_' + fileId);
}

function setCachedImage(fileId, imageData) {
  CACHE.put('img_' + fileId, imageData, CACHE_EXPIRATION);
}

// ============================================================================
// MEMORY MANAGEMENT & LOGGING
// ============================================================================

/**
 * Properly dispose of image resources and force garbage collection hints
 * @param {Blob} blob - The image blob to dispose
 * @param {string} imageData - Base64 encoded image data to clear
 * @param {string} context - Context for logging
 */
function disposeImageResources(blob, imageData, context = 'unknown') {
  try {
    // Clear references
    blob = null;
    imageData = null;
    
    // Force garbage collection hint
    Utilities.sleep(GARBAGE_COLLECTION_HINT_DELAY);
    
    logWithContext('DEBUG', 'Image resources disposed', { context: context });
  } catch (error) {
    logWithContext('ERROR', 'Error disposing image resources', { 
      context: context, 
      error: error.message 
    });
  }
}

/**
 * Get approximate memory usage information
 * @returns {string} Memory usage description
 */
function getMemoryUsage() {
  try {
    // Approximate memory usage based on execution time and operations
    const executionTime = Date.now() - START_TIME;
    const cacheSize = CACHE.get('cache_size_estimate') || '0';
    return `Execution: ${executionTime}ms, Cache: ~${cacheSize}KB`;
  } catch (error) {
    return 'Memory info unavailable';
  }
}

/**
 * Enhanced logging with context, timestamps, and memory information
 * @param {string} level - Log level (DEBUG, INFO, WARN, ERROR)
 * @param {string} message - Log message
 * @param {Object} context - Additional context information
 */
function logWithContext(level, message, context = {}) {
  try {
    const timestamp = new Date().toISOString();
    const memoryUsage = getMemoryUsage();
    const contextStr = Object.keys(context).length > 0 ? JSON.stringify(context) : 'none';
    
    Logger.log(`[${timestamp}] [${level}] ${message} | Context: ${contextStr} | Memory: ${memoryUsage}`);
  } catch (error) {
    // Fallback to basic logging if enhanced logging fails
    Logger.log(`[${level}] ${message}`);
  }
}

/**
 * Perform aggressive memory cleanup during batch processing
 * @param {number} processedCount - Number of images processed so far
 */
function performMemoryCleanup(processedCount) {
  try {
    if (processedCount % MEMORY_CHECK_INTERVAL === 0) {
      // Clear cache entries older than current batch
      const cacheKeys = ['img_*'];
      cacheKeys.forEach(pattern => {
        try {
          CACHE.removeAll([pattern]);
        } catch (e) {
          logWithContext('WARN', 'Cache cleanup failed', { pattern: pattern, error: e.message });
        }
      });
      
      // Force garbage collection hint
      Utilities.sleep(GARBAGE_COLLECTION_HINT_DELAY * 2);
      
      logWithContext('INFO', 'Memory cleanup performed', { 
        processedCount: processedCount,
        interval: MEMORY_CHECK_INTERVAL 
      });
    }
  } catch (error) {
    logWithContext('ERROR', 'Memory cleanup failed', { error: error.message });
  }
}

// ============================================================================
// ERROR HANDLING & RETRY LOGIC
// ============================================================================

/**
 * Structured error class for consistent error handling
 */
class ScriptError {
  constructor(code, message, details = {}) {
    this.code = code;
    this.message = message;
    this.details = details;
    this.timestamp = new Date().toISOString();
  }
  
  toString() {
    return `[${this.code}] ${this.message} (${this.timestamp})`;
  }
}

// Error codes
const ERROR_CODES = {
  FOLDER_ACCESS: 'FOLDER_ACCESS_ERROR',
  FILE_ACCESS: 'FILE_ACCESS_ERROR',
  IMAGE_RESIZE: 'IMAGE_RESIZE_ERROR',
  SHEET_ACCESS: 'SHEET_ACCESS_ERROR',
  NETWORK: 'NETWORK_ERROR',
  MEMORY: 'MEMORY_ERROR',
  TIMEOUT: 'TIMEOUT_ERROR',
  VALIDATION: 'VALIDATION_ERROR'
};

// Retry configuration
const RETRY_CONFIG = {
  maxRetries: 3,
  baseDelay: 1000, // 1 second
  maxDelay: 30000, // 30 seconds
  backoffMultiplier: 2,
  retryableErrors: [
    ERROR_CODES.NETWORK,
    ERROR_CODES.FILE_ACCESS,
    ERROR_CODES.SHEET_ACCESS
  ]
};

/**
 * Calculate exponential backoff delay
 * @param {number} attempt - Current attempt number (0-based)
 * @returns {number} Delay in milliseconds
 */
function calculateBackoffDelay(attempt) {
  const delay = RETRY_CONFIG.baseDelay * Math.pow(RETRY_CONFIG.backoffMultiplier, attempt);
  return Math.min(delay, RETRY_CONFIG.maxDelay);
}

/**
 * Check if an error is retryable
 * @param {ScriptError|Error} error - The error to check
 * @returns {boolean} Whether the error is retryable
 */
function isRetryableError(error) {
  if (error instanceof ScriptError) {
    return RETRY_CONFIG.retryableErrors.includes(error.code);
  }
  
  // Check common retryable error patterns
  const retryablePatterns = [
    /timeout/i,
    /network/i,
    /connection/i,
    /service unavailable/i,
    /rate limit/i
  ];
  
  return retryablePatterns.some(pattern => pattern.test(error.message));
}

/**
 * Execute a function with retry logic and exponential backoff
 * @param {Function} fn - Function to execute
 * @param {string} operationName - Name of the operation for logging
 * @param {Object} context - Additional context for logging
 * @returns {*} Result of the function execution
 */
function executeWithRetry(fn, operationName, context = {}) {
  let lastError = null;
  
  for (let attempt = 0; attempt < RETRY_CONFIG.maxRetries; attempt++) {
    try {
      logWithContext('DEBUG', `Executing ${operationName}`, { 
        attempt: attempt + 1, 
        maxRetries: RETRY_CONFIG.maxRetries,
        ...context 
      });
      
      return fn();
      
    } catch (error) {
      lastError = error;
      
      logWithContext('WARN', `${operationName} failed`, { 
        attempt: attempt + 1, 
        maxRetries: RETRY_CONFIG.maxRetries,
        error: error.message,
        ...context 
      });
      
      // Don't retry on the last attempt or if error is not retryable
      if (attempt === RETRY_CONFIG.maxRetries - 1 || !isRetryableError(error)) {
        break;
      }
      
      // Calculate and apply backoff delay
      const delay = calculateBackoffDelay(attempt);
      logWithContext('INFO', `Retrying ${operationName} after delay`, { 
        delay: delay,
        nextAttempt: attempt + 2,
        ...context 
      });
      
      Utilities.sleep(delay);
    }
  }
  
  // All retries failed
  const finalError = lastError instanceof ScriptError ? lastError : 
    new ScriptError(ERROR_CODES.NETWORK, `${operationName} failed after ${RETRY_CONFIG.maxRetries} attempts: ${lastError.message}`, context);
  
  logWithContext('ERROR', `${operationName} failed permanently`, { 
    error: finalError.toString(),
    ...context 
  });
  
  throw finalError;
}

/**
 * Wrap common operations with error handling
 */
function safeFileAccess(fileId, operationName = 'file access') {
  return executeWithRetry(
    () => DriveApp.getFileById(fileId),
    operationName,
    { fileId: fileId }
  );
}

function safeFolderAccess(folderId, operationName = 'folder access') {
  return executeWithRetry(
    () => DriveApp.getFolderById(folderId),
    operationName,
    { folderId: folderId }
  );
}

function safeSheetAccess(spreadsheetId, sheetId, operationName = 'sheet access') {
  return executeWithRetry(
    () => {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheets().filter(s => s.getSheetId() == sheetId)[0];
      if (!sheet) {
        throw new ScriptError(ERROR_CODES.SHEET_ACCESS, 'Sheet not found', { spreadsheetId, sheetId });
      }
      return { spreadsheet, sheet };
    },
    operationName,
    { spreadsheetId: spreadsheetId, sheetId: sheetId }
  );
}

// ============================================================================
// INPUT VALIDATION FUNCTIONS
// ============================================================================

/**
 * Validate folder URL format
 * @param {string} folderUrl - The folder URL to validate
 * @returns {boolean} Whether the URL is valid
 */
function validateFolderUrl(folderUrl) {
  if (!folderUrl || typeof folderUrl !== 'string') {
    return false;
  }
  
  // Check for Google Drive folder URL patterns
  const patterns = [
    /^https:\/\/drive\.google\.com\/drive\/folders\/[-\w]{25,}/,
    /^https:\/\/drive\.google\.com\/drive\/u\/\d+\/folders\/[-\w]{25,}/,
    /^https:\/\/drive\.google\.com\/open\?id=[-\w]{25,}/
  ];
  
  return patterns.some(pattern => pattern.test(folderUrl));
}

/**
 * Validate row and column inputs
 * @param {*} row - Row value to validate
 * @param {*} col - Column value to validate
 * @returns {Object} Validation result with isValid and errors
 */
function validateRowCol(row, col) {
  const errors = [];
  
  // Convert to numbers
  const rowNum = parseInt(row);
  const colNum = parseInt(col);
  
  if (isNaN(rowNum) || rowNum < 1 || rowNum > 1000000) {
    errors.push('Row must be a number between 1 and 1,000,000');
  }
  
  if (isNaN(colNum) || colNum < 1 || colNum > 18278) {
    errors.push('Column must be a number between 1 and 18,278');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors,
    row: rowNum,
    col: colNum
  };
}

/**
 * Extract folder ID from various Google Drive URL formats
 * @param {string} url - The folder URL
 * @returns {string|null} The folder ID or null if not found
 */
// ============================================================================
// PERFORMANCE OPTIMIZATION & MONITORING
// ============================================================================

// Performance configuration
const PERFORMANCE_CONFIG = {
  // Adaptive batch sizing - reduced for better stop responsiveness
  minBatchSize: 1,
  maxBatchSize: 8,
  defaultBatchSize: 4,
  
  // Performance thresholds (milliseconds) - reduced for better responsiveness
  targetProcessingTime: 2500, // 2.5 seconds per batch
  maxProcessingTime: 3500,    // 3.5 seconds max
  
  // Cache configuration
  cacheExpirationHours: 24,
  maxCacheEntries: 100,
  
  // Performance monitoring
  enablePerformanceLogging: true,
  performanceLogInterval: 10 // Log every 10 images
};

/**
 * Performance monitoring class
 */
class PerformanceMonitor {
  constructor() {
    this.startTime = Date.now();
    this.imageProcessingTimes = [];
    this.memorySnapshots = [];
    this.batchMetrics = [];
  }
  
  startImageProcessing() {
    return Date.now();
  }
  
  endImageProcessing(startTime, imageSize = 0) {
    const processingTime = Date.now() - startTime;
    this.imageProcessingTimes.push({
      time: processingTime,
      size: imageSize,
      timestamp: Date.now()
    });
    return processingTime;
  }
  
  recordBatchMetrics(batchSize, batchTime, memoryUsage) {
    this.batchMetrics.push({
      batchSize: batchSize,
      batchTime: batchTime,
      memoryUsage: memoryUsage,
      timestamp: Date.now(),
      avgTimePerImage: batchTime / batchSize
    });
  }
  
  getAverageProcessingTime() {
    if (this.imageProcessingTimes.length === 0) return 0;
    const total = this.imageProcessingTimes.reduce((sum, item) => sum + item.time, 0);
    return total / this.imageProcessingTimes.length;
  }
  
  getRecommendedBatchSize() {
    if (this.batchMetrics.length < 2) {
      return PERFORMANCE_CONFIG.defaultBatchSize;
    }
    
    // Get recent batch metrics (last 3 batches)
    const recentMetrics = this.batchMetrics.slice(-3);
    const avgBatchTime = recentMetrics.reduce((sum, m) => sum + m.batchTime, 0) / recentMetrics.length;
    const avgBatchSize = recentMetrics.reduce((sum, m) => sum + m.batchSize, 0) / recentMetrics.length;
    
    // Calculate optimal batch size based on target processing time
    let recommendedSize = Math.round((PERFORMANCE_CONFIG.targetProcessingTime / avgBatchTime) * avgBatchSize);
    
    // Apply constraints
    recommendedSize = Math.max(PERFORMANCE_CONFIG.minBatchSize, recommendedSize);
    recommendedSize = Math.min(PERFORMANCE_CONFIG.maxBatchSize, recommendedSize);
    
    logWithContext('DEBUG', 'Batch size recommendation calculated', {
      avgBatchTime: avgBatchTime,
      avgBatchSize: avgBatchSize,
      recommendedSize: recommendedSize,
      targetTime: PERFORMANCE_CONFIG.targetProcessingTime
    });
    
    return recommendedSize;
  }
  
  logPerformanceReport() {
    if (!PERFORMANCE_CONFIG.enablePerformanceLogging) return;
    
    const totalTime = Date.now() - this.startTime;
    const avgImageTime = this.getAverageProcessingTime();
    const totalImages = this.imageProcessingTimes.length;
    
    logWithContext('INFO', 'Performance Report', {
      totalProcessingTime: totalTime,
      totalImagesProcessed: totalImages,
      averageTimePerImage: avgImageTime,
      totalBatches: this.batchMetrics.length,
      throughputImagesPerMinute: totalImages > 0 ? Math.round((totalImages / totalTime) * 60000) : 0
    });
  }
}

/**
 * Enhanced caching system with expiration and size limits
 */
class EnhancedCache {
  constructor() {
    this.cache = CacheService.getScriptCache();
  }
  
  /**
   * Generate cache key with metadata
   */
  generateCacheKey(fileId, width) {
    return `img_${fileId}_${width}_v2`;
  }
  
  /**
   * Store image in cache with metadata
   */
  put(fileId, width, imageData, fileSize = 0) {
    try {
      const key = this.generateCacheKey(fileId, width);
      
      // Check if image data is too large for cache (Google Apps Script limit is ~100KB per cache entry)
      const maxCacheSize = 90 * 1024; // 90KB to be safe
      if (imageData.length > maxCacheSize) {
        logWithContext('DEBUG', 'Image too large for cache, skipping', {
          fileId: fileId,
          dataSize: imageData.length,
          maxSize: maxCacheSize
        });
        return false;
      }
      
      const metadata = {
        data: imageData,
        timestamp: Date.now(),
        fileSize: fileSize,
        width: width
      };
      
      // Store with expiration
      const expirationSeconds = PERFORMANCE_CONFIG.cacheExpirationHours * 3600;
      this.cache.put(key, JSON.stringify(metadata), expirationSeconds);
      
      logWithContext('DEBUG', 'Image cached', {
        fileId: fileId,
        width: width,
        dataSize: imageData.length,
        fileSize: fileSize
      });
      
      return true;
    } catch (error) {
      logWithContext('WARN', 'Failed to cache image', {
        fileId: fileId,
        error: error.message
      });
      return false;
    }
  }
  
  /**
   * Retrieve image from cache
   */
  get(fileId, width) {
    try {
      const key = this.generateCacheKey(fileId, width);
      const cached = this.cache.get(key);
      
      if (!cached) {
        return null;
      }
      
      const metadata = JSON.parse(cached);
      
      // Check if cache entry is still valid
      const age = Date.now() - metadata.timestamp;
      const maxAge = PERFORMANCE_CONFIG.cacheExpirationHours * 3600 * 1000;
      
      if (age > maxAge) {
        this.cache.remove(key);
        logWithContext('DEBUG', 'Cache entry expired', {
          fileId: fileId,
          age: age,
          maxAge: maxAge
        });
        return null;
      }
      
      logWithContext('DEBUG', 'Cache hit', {
        fileId: fileId,
        width: width,
        age: age
      });
      
      return metadata.data;
    } catch (error) {
      logWithContext('WARN', 'Failed to retrieve from cache', {
        fileId: fileId,
        error: error.message
      });
      return null;
    }
  }
  
  /**
   * Clear old cache entries
   */
  cleanup() {
    try {
      // Note: Google Apps Script cache doesn't provide direct cleanup methods
      // Entries will expire automatically based on expiration time
      logWithContext('DEBUG', 'Cache cleanup requested', {});
    } catch (error) {
      logWithContext('WARN', 'Cache cleanup failed', {
        error: error.message
      });
    }
  }
  
  /**
   * Get cache hit rate (placeholder implementation)
   * Note: Google Apps Script cache doesn't provide hit rate metrics
   */
  getHitRate() {
    // Return a default value since GAS cache doesn't provide hit rate metrics
    return 0.0;
  }
}

/**
 * Adaptive batch processor
 */
class AdaptiveBatchProcessor {
  constructor() {
    this.performanceMonitor = new PerformanceMonitor();
    this.enhancedCache = new EnhancedCache();
    this.currentBatchSize = PERFORMANCE_CONFIG.defaultBatchSize;
  }
  
  /**
   * Process a batch of images with adaptive sizing
   */
  processBatch(fileList, startIndex, spreadsheetId, sheetId, startRow, startCol) {
    const batchStartTime = Date.now();
    let processedCount = 0;
    
    try {
      // Get current batch size recommendation
      this.currentBatchSize = this.performanceMonitor.getRecommendedBatchSize();
      
      logWithContext('INFO', 'Starting adaptive batch processing', {
        batchSize: this.currentBatchSize,
        startIndex: startIndex,
        totalFiles: fileList.length
      });
      
      const endIndex = Math.min(startIndex + this.currentBatchSize, fileList.length);
      
      for (let i = startIndex; i < endIndex; i++) {
        // Check for stop flag first for immediate responsiveness
        if (SCRIPT_PROPERTIES.getProperty('shouldStop') === 'true') {
          logWithContext('INFO', 'Stop requested during batch processing', {
            processedInBatch: processedCount,
            currentIndex: i
          });
          break;
        }
        
        if (isTimeUp()) {
          logWithContext('WARN', 'Time limit reached during batch processing', {
            processedInBatch: processedCount,
            currentIndex: i
          });
          break;
        }
        
        const imageStartTime = this.performanceMonitor.startImageProcessing();
        
        try {
          // Process single image with caching
          const result = this.processSingleImageWithCache(fileList[i], spreadsheetId, sheetId, startRow, startCol, i);
          
          // Check if processing was stopped
          if (result.stopped) {
            logWithContext('INFO', 'Processing stopped during image processing', {
              processedInBatch: processedCount,
              currentIndex: i
            });
            break;
          }
          
          if (result.success) {
            processedCount++;
            this.performanceMonitor.endImageProcessing(imageStartTime, result.imageSize);
            
            // Log progress periodically
            if (processedCount % PERFORMANCE_CONFIG.performanceLogInterval === 0) {
              this.performanceMonitor.logPerformanceReport();
            }
          }
        } catch (error) {
          logWithContext('ERROR', 'Failed to process image in batch', {
            index: i,
            fileId: fileList[i].fileId,
            error: error.message
          });
        }
      }
      
      // Record batch metrics
      const batchTime = Date.now() - batchStartTime;
      const memoryUsage = getMemoryUsage();
      this.performanceMonitor.recordBatchMetrics(processedCount, batchTime, memoryUsage);
      
      // Update performance metrics cache for UI
      updatePerformanceMetricsCache(
        this.currentBatchSize,
        this.performanceMonitor.getAverageProcessingTime(),
        memoryUsage,
        this.enhancedCache.getHitRate()
      );
      
      logWithContext('INFO', 'Batch processing completed', {
        processedCount: processedCount,
        batchTime: batchTime,
        nextBatchSize: this.performanceMonitor.getRecommendedBatchSize()
      });
      
      return {
        processedCount: processedCount,
        nextIndex: startIndex + processedCount,
        batchTime: batchTime,
        completed: startIndex + processedCount >= fileList.length
      };
      
    } catch (error) {
      logWithContext('ERROR', 'Batch processing failed', {
        startIndex: startIndex,
        processedCount: processedCount,
        error: error.message
      });
      throw error;
    }
  }
  
  /**
   * Process single image with caching support
   */
  processSingleImageWithCache(fileObj, spreadsheetId, sheetId, startRow, startCol, index) {
    try {
      // Check for stop flag immediately for better responsiveness
      if (SCRIPT_PROPERTIES.getProperty('shouldStop') === 'true') {
        logWithContext('DEBUG', 'Stop requested during single image processing', {
          fileId: fileObj.fileId,
          index: index
        });
        return { success: false, stopped: true };
      }
      
      // Check cache first
      const cachedImage = this.enhancedCache.get(fileObj.fileId, MAX_IMAGE_WIDTH);
      
      if (cachedImage) {
        // Use cached image with proper data URI format
        const imageData = `data:image/jpeg;base64,${cachedImage}`;
        const rowIndex = startRow + index;
        const range = SpreadsheetApp.openById(spreadsheetId)
          .getSheets().filter(s => s.getSheetId() == sheetId)[0]
          .getRange(rowIndex, startCol);
        
        insertCellImage(range, imageData, fileObj.fileName, '');
        
        logWithContext('DEBUG', 'Used cached image', {
          fileId: fileObj.fileId,
          index: index
        });
        
        return { success: true, imageSize: cachedImage.length, fromCache: true };
      }
      
      // Process image normally
      const resizedImage = resizeImage(fileObj.fileId, MAX_IMAGE_WIDTH);
      const base64Data = Utilities.base64Encode(resizedImage.getBytes());
      const imageData = `data:image/jpeg;base64,${base64Data}`;
      
      // Cache the processed image (cache base64 only to save space)
      this.enhancedCache.put(fileObj.fileId, MAX_IMAGE_WIDTH, base64Data, resizedImage.getBytes().length);
      
      // Insert into sheet
      const rowIndex = startRow + index;
      const { spreadsheet, sheet } = safeSheetAccess(spreadsheetId, sheetId, 'adaptive batch sheet access');
      const range = sheet.getRange(rowIndex, startCol);
      
      insertCellImage(range, imageData, fileObj.fileName, '');
      
      // Clean up resources
      disposeImageResources(resizedImage, 'processed image (adaptive batch)');
      
      return { success: true, imageSize: imageData.length, fromCache: false };
      
    } catch (error) {
      logWithContext('ERROR', 'Failed to process single image', {
        fileId: fileObj.fileId,
        index: index,
        error: error.message
      });
      return { success: false, error: error.message };
    }
  }
}

function extractFolderIdFromUrl(url) {
  if (!url) return null;
  
  // Pattern for folder ID extraction
  const patterns = [
    /\/folders\/([-\w]{25,})/,
    /[?&]id=([-\w]{25,})/,
    /^([-\w]{25,})$/ // Direct ID
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match) {
      return match[1];
    }
  }
  
  return null;
}

// ============================================================================
// INITIALIZATION FUNCTIONS
// ============================================================================

function initializeProcess(folderUrl, startRow, startCol) {
  // Input validation
  if (!folderUrl || !startRow || !startCol) {
    throw new ScriptError(ERROR_CODES.VALIDATION, 'Invalid input: folderUrl, startRow, and startCol are required.');
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
  logWithContext('INFO', 'Initialization complete', { properties: props });
  return true;
}

function insertImagesFromUI(folderUrl, startRow, startCol) {
  try {
    logWithContext('INFO', 'Starting image insertion from UI', { 
      folderUrl: folderUrl, 
      startRow: startRow, 
      startCol: startCol 
    });
    
    // Validate inputs
    if (!validateFolderUrl(folderUrl)) {
      throw new ScriptError(ERROR_CODES.VALIDATION, 'Invalid folder URL format. Please provide a valid Google Drive folder URL.');
    }
    
    const validation = validateRowCol(startRow, startCol);
    if (!validation.isValid) {
      throw new ScriptError(ERROR_CODES.VALIDATION, 'Invalid row/column values: ' + validation.errors.join(', '));
    }
    
    // Use validated values
    startRow = validation.row;
    startCol = validation.col;
    
    // Check if already processing
    const isProcessing = SCRIPT_PROPERTIES.getProperty('processingBatch');
    if (isProcessing === 'true') {
      throw new ScriptError(ERROR_CODES.VALIDATION, 'Image insertion is already in progress. Please wait for it to complete or stop it first.');
    }
    
    // Initialize process with error handling
    executeWithRetry(
      () => initializeProcess(folderUrl, startRow, startCol),
      'initialize process',
      { folderUrl: folderUrl, startRow: startRow, startCol: startCol }
    );
    
    // Get current sheet info with error handling
    const { spreadsheet, sheet } = executeWithRetry(
      () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const activeSheet = ss.getActiveSheet();
        return { 
          spreadsheet: ss, 
          sheet: activeSheet 
        };
      },
      'get active sheet',
      {}
    );
    
    // Collect and sort files with error handling first
    const preview = executeWithRetry(
      () => collectAndSortFiles(folderUrl),
      'collect and sort files',
      { folderUrl: folderUrl }
    );
    
    // Store sheet information and processing state
    SCRIPT_PROPERTIES.setProperties({
      'spreadsheetId': spreadsheet.getId(),
      'sheetId': sheet.getSheetId().toString(),
      'startRow': startRow.toString(),
      'startCol': startCol.toString(),
      'processingBatch': 'starting', // Use 'starting' state for UI detection
      'shouldStop': 'false',
      'lastProcessedIndex': '-1',
      'readyToProcess': 'true' // Flag to indicate processing should begin
    });
    
    logWithContext('INFO', 'Processing setup complete, UI can start polling', { 
      totalFiles: preview.length,
      startRow: startRow,
      startCol: startCol 
    });
    
    // Return immediately to allow UI polling to start
    // Processing will begin when getProcessingStatus detects readyToProcess flag
    return {
      success: true,
      message: `Setup complete. Processing ${preview.length} images will begin shortly.`,
      totalFiles: preview.length
    };
    
  } catch (error) {
    logWithContext('ERROR', 'Failed to start image insertion', { 
      folderUrl: folderUrl,
      startRow: startRow,
      startCol: startCol,
      error: error instanceof ScriptError ? error.toString() : error.message 
    });
    
    // Clean up on error
    try {
      SCRIPT_PROPERTIES.deleteProperty('processingBatch');
      cleanupExistingTriggers();
    } catch (cleanupError) {
      logWithContext('WARN', 'Cleanup after insertion start error failed', { 
        cleanupError: cleanupError.message 
      });
    }
    
    if (error instanceof ScriptError) {
      throw error;
    }
    
    throw new ScriptError(ERROR_CODES.VALIDATION, 'Failed to start image insertion: ' + error.message);
  }
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

// ============================================================================
// GOOGLE DRIVE OPERATIONS
// ============================================================================

function collectFiles(folderUrl) {
  if (!folderUrl) {
    throw new ScriptError(ERROR_CODES.VALIDATION, 'Invalid input: folderUrl is required.');
  }
  
  logWithContext('INFO', 'Starting file collection', { folderUrl: folderUrl });
  
  try {
    const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
    if (!folderIdMatch) {
      throw new ScriptError(ERROR_CODES.VALIDATION, 'Invalid folder URL format');
    }
    
    const folderId = folderIdMatch[0];
    logWithContext('DEBUG', 'Extracted folder ID', { folderId: folderId });
    
    const folder = safeFolderAccess(folderId, 'collect files folder access');
    logWithContext('INFO', 'Successfully accessed folder', { folderName: folder.getName() });
    
    const files = executeWithRetry(
      () => folder.getFiles(),
      'get files from folder',
      { folderId: folderId }
    );
    
    logWithContext('DEBUG', 'Successfully got files iterator', {});
    
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
      throw new ScriptError(ERROR_CODES.VALIDATION, 'No files found in the specified folder');
    }
    
    SCRIPT_PROPERTIES.setProperty('fileList', JSON.stringify(fileList));
    logWithContext('INFO', 'File collection completed', { totalFiles: fileList.length });
    
    return fileList;
    
  } catch (error) {
    if (error instanceof ScriptError) {
      throw error;
    }
    throw new ScriptError(ERROR_CODES.FILE_ACCESS, 'Failed to collect files: ' + error.message);
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

/**
 * Check if processing is complete
 * @returns {Object} Status object with completion info
 */
function getProcessingStatus() {
  try {
    const props = SCRIPT_PROPERTIES.getProperties();
    const processingState = props.processingBatch;
    const isProcessing = processingState === 'true' || processingState === 'starting';
    const shouldStop = props.shouldStop === 'true';
    const isComplete = props.isComplete === 'true';
    const lastProcessedIndex = parseInt(props.lastProcessedIndex || '-1');
    const fileListJson = props.fileList;
    const readyToProcess = props.readyToProcess === 'true';
    
    // If we're ready to process and UI is polling, start processing immediately
    if (readyToProcess && !isComplete && !shouldStop) {
      logWithContext('INFO', 'UI polling detected, starting image processing immediately', {});
      
      // Clear the readyToProcess flag to prevent multiple starts
      SCRIPT_PROPERTIES.deleteProperty('readyToProcess');
      
      // Start processing asynchronously using a minimal trigger
      try {
        const trigger = ScriptApp.newTrigger('handleImageInsertion')
          .timeBased()
          .after(1) // 1ms - minimal delay just to make it async
          .create();
        
        SCRIPT_PROPERTIES.setProperty('currentTriggerId', trigger.getUniqueId());
        
        logWithContext('INFO', 'Processing trigger created', {
          triggerId: trigger.getUniqueId()
        });
      } catch (triggerError) {
        logWithContext('ERROR', 'Failed to create processing trigger', {
          error: triggerError.message
        });
      }
    }
    
    let totalFiles = 0;
    
    if (fileListJson) {
      const fileList = JSON.parse(fileListJson);
      totalFiles = fileList.length;
      
      // If not explicitly marked complete, check if all files processed
      if (!isComplete && !isProcessing && (lastProcessedIndex + 1) >= totalFiles) {
        // Mark as complete if all files processed but flag not set
        SCRIPT_PROPERTIES.setProperty('isComplete', 'true');
        return {
          isProcessing: false,
          shouldStop: shouldStop,
          isComplete: true,
          processedCount: lastProcessedIndex + 1,
          totalFiles: totalFiles,
          progressPercentage: 100
        };
      }
    }
    
    return {
      isProcessing: isProcessing,
      shouldStop: shouldStop,
      isComplete: isComplete,
      processedCount: lastProcessedIndex + 1,
      totalFiles: totalFiles,
      progressPercentage: totalFiles > 0 ? Math.round(((lastProcessedIndex + 1) / totalFiles) * 100) : 0
    };
    
  } catch (error) {
    logWithContext('ERROR', 'Failed to get processing status', { error: error.message });
    return {
      isProcessing: false,
      shouldStop: false,
      isComplete: false,
      processedCount: 0,
      totalFiles: 0,
      progressPercentage: 0
    };
  }
}

function cleanupAfterCompletion() {
  try {
    logWithContext('INFO', 'Starting cleanup after completion', {});
    
    // Set completion flags - keep processingBatch true until the very end
    SCRIPT_PROPERTIES.setProperty('isComplete', 'true');
    SCRIPT_PROPERTIES.setProperty('shouldStop', 'false');
    
    cleanupExistingTriggers();
    
    // Clean up processing-related properties but keep completion status
    const propsToDelete = [
      'currentTriggerId',
      'batchSize',
      'currentBatchIndex',
      'processingStartTime'
    ];
    
    propsToDelete.forEach(prop => {
      try {
        SCRIPT_PROPERTIES.deleteProperty(prop);
      } catch (e) {
        // Ignore individual deletion errors
      }
    });
    
    if (CACHE) {
      try {
        CACHE.removeAll(['img_*']);
        logWithContext('DEBUG', 'Cache cleared successfully', {});
      } catch (cacheError) {
        logWithContext('WARN', 'Cache cleanup failed', { error: cacheError.message });
      }
    }
    
    // Force final garbage collection hint
    Utilities.sleep(GARBAGE_COLLECTION_HINT_DELAY * 3);
    
    // Only set processingBatch to false at the very end
    SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
    
    logWithContext('INFO', 'Cleanup completed successfully', {});
  } catch (error) {
    logWithContext('ERROR', 'Error during cleanup', { error: error.message });
  } finally {
    // Ensure processingBatch is set to false even if there's an error
    try {
      SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
    } catch (finalError) {
      logWithContext('ERROR', 'Failed to set processingBatch to false in finally block', { error: finalError.message });
    }
  }
}

/**
 * Clean up completion status after UI has detected it
 */
function cleanupCompletionStatus() {
  try {
    const propsToDelete = [
      'isComplete',
      'fileList',
      'lastProcessedIndex',
      'processingBatch',
      'shouldStop',
      'folderUrl',
      'sheetId',
      'spreadsheetId',
      'startCol',
      'startRow',
      'currentTriggerId',
      'batchSize',
      'currentBatchIndex',
      'processingStartTime',
      'partialCollection'
    ];
    
    propsToDelete.forEach(prop => {
      try {
        SCRIPT_PROPERTIES.deleteProperty(prop);
      } catch (e) {
        // Ignore individual deletion errors
      }
    });
    
    logWithContext('INFO', 'All completion status and process properties cleaned up', { 
      deletedProperties: propsToDelete.length 
    });
  } catch (error) {
    logWithContext('WARN', 'Failed to clean up completion status', { error: error.message });
  }
}

/**
 * Get current performance metrics for UI display
 * @returns {Object} Performance metrics object
 */
function getPerformanceMetrics() {
  try {
    const props = SCRIPT_PROPERTIES.getProperties();
    const lastProcessedIndex = parseInt(props.lastProcessedIndex || '-1');
    const fileListJson = props.fileList;
    
    if (!fileListJson) {
      return null;
    }
    
    const fileList = JSON.parse(fileListJson);
    const totalFiles = fileList.length;
    const processedCount = lastProcessedIndex + 1;
    
    // Get cached performance data if available
    const performanceData = CacheService.getScriptCache().get('performance_metrics');
    let metrics = {
      imagesProcessed: processedCount,
      totalImages: totalFiles
    };
    
    if (performanceData) {
      try {
        const data = JSON.parse(performanceData);
        metrics = { ...metrics, ...data };
      } catch (parseError) {
        logWithContext('WARN', 'Failed to parse performance data', { error: parseError.message });
      }
    }
    
    // Calculate estimated time remaining
    if (metrics.avgProcessingTime && processedCount > 0) {
      const remainingImages = totalFiles - processedCount;
      const estimatedMs = remainingImages * metrics.avgProcessingTime;
      metrics.estimatedTimeRemaining = estimatedMs;
    }
    
    // Calculate progress percentage
    metrics.progressPercentage = totalFiles > 0 ? Math.round((processedCount / totalFiles) * 100) : 0;
    
    return metrics;
    
  } catch (error) {
    logWithContext('ERROR', 'Failed to get performance metrics', { error: error.message });
    return null;
  }
}

/**
 * Get the total number of files collected for UI display
 * @returns {number|null} Total file count or null if not available
 */
function getTotalFileCount() {
  try {
    const fileListJson = SCRIPT_PROPERTIES.getProperty('fileList');
    
    if (!fileListJson) {
      return null;
    }
    
    const fileList = JSON.parse(fileListJson);
    return fileList.length;
    
  } catch (error) {
    logWithContext('ERROR', 'Failed to get total file count', { error: error.message });
    return null;
  }
}

/**
 * Update performance metrics in cache for UI consumption
 * @param {Object} metrics - Performance metrics to cache
 */
function updatePerformanceMetricsCache(metrics) {
  try {
    if (!metrics) return;
    
    const cacheData = {
      currentBatchSize: metrics.currentBatchSize,
      avgProcessingTime: metrics.avgProcessingTime,
      memoryUsage: metrics.memoryUsage,
      cacheHitRate: metrics.cacheHitRate,
      lastUpdated: Date.now()
    };
    
    CacheService.getScriptCache().put('performance_metrics', JSON.stringify(cacheData), 300); // 5 minutes
    
  } catch (error) {
    logWithContext('WARN', 'Failed to update performance metrics cache', { error: error.message });
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

// ============================================================================
// BATCH PROCESSING & TRIGGERS
// ============================================================================

function handleImageInsertion() {
  let batchProcessor = null;
  
  try {
    logWithContext('INFO', 'Starting image insertion batch processing', {});
    
    // Check if we should stop
    const shouldStop = SCRIPT_PROPERTIES.getProperty('shouldStop');
    if (shouldStop === 'true') {
      logWithContext('INFO', 'Stop requested, cleaning up', {});
      cleanupAfterStop();
      return;
    }
    
    // Check if we're still processing
    const isProcessing = SCRIPT_PROPERTIES.getProperty('processingBatch');
    if (isProcessing !== 'true' && isProcessing !== 'starting') {
      logWithContext('WARN', 'Processing batch flag not set, exiting', { processingBatch: isProcessing });
      return;
    }
    
    // If we're in 'starting' state, change to 'true' to indicate active processing
    if (isProcessing === 'starting') {
      SCRIPT_PROPERTIES.setProperty('processingBatch', 'true');
      logWithContext('INFO', 'Changed processing state from starting to active', {});
    }
    
    // Get processing parameters
    const props = SCRIPT_PROPERTIES.getProperties();
    const spreadsheetId = props.spreadsheetId;
    const sheetId = parseInt(props.sheetId);
    const startRow = parseInt(props.startRow);
    const startCol = parseInt(props.startCol);
    const lastProcessedIndex = parseInt(props.lastProcessedIndex || '-1');
    
    // Validate parameters (note: sheetId can be 0, which is valid)
    if (!spreadsheetId || isNaN(sheetId) || isNaN(startRow) || isNaN(startCol) || startRow < 1 || startCol < 1) {
      throw new ScriptError(ERROR_CODES.VALIDATION, 'Missing required processing parameters');
    }
    
    // Get file list
    const fileListJson = props.fileList;
    if (!fileListJson) {
      throw new ScriptError(ERROR_CODES.VALIDATION, 'No file list found in properties');
    }
    
    const fileList = JSON.parse(fileListJson);
    const nextIndex = lastProcessedIndex + 1;
    
    logWithContext('INFO', 'Processing parameters loaded', {
      totalFiles: fileList.length,
      nextIndex: nextIndex,
      remainingFiles: fileList.length - nextIndex
    });
    
    // Check if processing is complete
    if (nextIndex >= fileList.length) {
      logWithContext('INFO', 'All images processed, completing', {
        totalProcessed: fileList.length
      });
      cleanupAfterCompletion();
      return;
    }
    
    // Initialize adaptive batch processor
    batchProcessor = new AdaptiveBatchProcessor();
    
    // Process batch with adaptive sizing
    const batchResult = batchProcessor.processBatch(
      fileList, 
      nextIndex, 
      spreadsheetId, 
      sheetId, 
      startRow, 
      startCol
    );
    
    logWithContext('INFO', 'Batch processing result', {
      processedCount: batchResult.processedCount,
      batchTime: batchResult.batchTime,
      nextIndex: batchResult.nextIndex,
      completed: batchResult.completed
    });
    
    // Update progress
    SCRIPT_PROPERTIES.setProperty('lastProcessedIndex', (batchResult.nextIndex - 1).toString());
    
    // Perform memory cleanup
    performMemoryCleanup(batchResult.nextIndex);
    
    // Check if processing is complete
    if (batchResult.completed) {
      logWithContext('INFO', 'Image insertion completed successfully', {
        totalProcessed: fileList.length,
        totalBatches: batchProcessor.performanceMonitor.batchMetrics.length
      });
      
      // Generate final performance report
      batchProcessor.performanceMonitor.logPerformanceReport();
      
      cleanupAfterCompletion();
      return;
    }
    
    // Check if we should stop before scheduling next batch
    const shouldStopAgain = SCRIPT_PROPERTIES.getProperty('shouldStop');
    if (shouldStopAgain === 'true') {
      logWithContext('INFO', 'Stop requested after batch, cleaning up', {});
      cleanupAfterStop();
      return;
    }
    
    // Check time limit
    if (isTimeUp()) {
      logWithContext('INFO', 'Time limit reached, scheduling continuation', {
        processedInSession: batchResult.processedCount,
        nextIndex: batchResult.nextIndex
      });
      
      // Schedule next execution
      executeWithRetry(
        () => {
          const trigger = ScriptApp.newTrigger('handleImageInsertion')
            .timeBased()
            .after(1000) // 1 second delay for continuation
            .create();
          
          SCRIPT_PROPERTIES.setProperty('currentTriggerId', trigger.getUniqueId());
          
          logWithContext('INFO', 'Continuation trigger created', {
            triggerId: trigger.getUniqueId(),
            nextIndex: batchResult.nextIndex
          });
        },
        'create continuation trigger',
        { nextIndex: batchResult.nextIndex }
      );
      
      return;
    }
    
    // Continue processing in same execution if time allows
     logWithContext('DEBUG', 'Continuing processing in same execution', {
       nextIndex: batchResult.nextIndex,
       remainingTime: MAX_EXECUTION_TIME - (Date.now() - START_TIME)
     });
    
    // Small delay to prevent overwhelming the system
    Utilities.sleep(200);
    
    // Recursive call for next batch
    handleImageInsertion();
    
  } catch (error) {
    logWithContext('ERROR', 'Image insertion batch failed', {
      error: error instanceof ScriptError ? error.toString() : error.message
    });
    
    // Clean up on error
    try {
      SCRIPT_PROPERTIES.setProperty('shouldStop', 'true');
      cleanupAfterStop();
    } catch (cleanupError) {
      logWithContext('ERROR', 'Cleanup after error failed', {
        cleanupError: cleanupError.message
      });
    }
    
    throw error;
  } finally {
    // Final cleanup
    if (batchProcessor) {
      try {
        batchProcessor.enhancedCache.cleanup();
      } catch (cacheError) {
        logWithContext('WARN', 'Cache cleanup in finally block failed', {
          error: cacheError.message
        });
      }
    }
  }
}

// Legacy function removed - functionality moved to AdaptiveBatchProcessor

// ============================================================================
// IMAGE PROCESSING
// ============================================================================

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

  // Look for JPEG SOF (Start of Frame) markers
  // SOF0 (0xC0), SOF1 (0xC1), SOF2 (0xC2), SOF3 (0xC3), etc.
  for (var i = 0; i < decodedImage.length - 8; i++) {
    if (decodedImage[i] == 0xFF) {
      var marker = decodedImage[i+1];
      // Check for SOF markers (0xC0 to 0xCF, excluding 0xC4, 0xC8, 0xCC which are not SOF)
      if ((marker >= 0xC0 && marker <= 0xC3) || 
          (marker >= 0xC5 && marker <= 0xC7) || 
          (marker >= 0xC9 && marker <= 0xCB) || 
          (marker >= 0xCD && marker <= 0xCF)) {
        // SOF structure: FF Cn LL LL P YY YY XX XX
        // Where YY YY is height and XX XX is width
        height = (decodedImage[i+5] << 8) | decodedImage[i+6];
        width = (decodedImage[i+7] << 8) | decodedImage[i+8];
        break;
      }
    }
  }

  return {width: width, height: height};
}

function resizeImage(fileId, maxWidth) {
  let file = null;
  let blob = null;
  let image = null;
  let imageData = null;
  let resizedBlob = null;
  
  try {
    logWithContext('DEBUG', 'Starting image resize', { fileId: fileId, maxWidth: maxWidth });
    
    file = executeWithRetry(
      () => DriveApp.getFileById(fileId),
      'get file for resize',
      { fileId: fileId }
    );
    
    blob = executeWithRetry(
      () => file.getBlob(),
      'get image blob',
      { fileId: fileId }
    );
    
    image = executeWithRetry(
      () => blob.getAs('image/jpeg'),
      'convert to JPEG',
      { fileId: fileId }
    );
    
    imageData = Utilities.base64Encode(image.getBytes());
    const dimensions = getImageDimensions_(imageData);
    
    logWithContext('DEBUG', 'Image dimensions retrieved', { 
      width: dimensions.width, 
      height: dimensions.height,
      fileId: fileId 
    });
    
    if (dimensions.width > maxWidth) {
      const newHeight = Math.round((maxWidth / dimensions.width) * dimensions.height);
      
      resizedBlob = executeWithRetry(
        () => image.getAs('image/jpeg').setWidth(maxWidth).setHeight(newHeight),
        'resize image',
        { 
          fileId: fileId, 
          newWidth: maxWidth, 
          newHeight: newHeight 
        }
      );
      
      logWithContext('INFO', 'Image resized', { 
        originalWidth: dimensions.width,
        originalHeight: dimensions.height,
        newWidth: maxWidth,
        newHeight: newHeight,
        fileId: fileId
      });
      
      // Dispose of intermediate resources
      disposeImageResources(image, imageData, 'resize-intermediate');
      
      return resizedBlob;
    }
    
    // Dispose of resources when no resize needed
    disposeImageResources(blob, imageData, 'resize-no-change');
    
    return image;
    
  } catch (error) {
    // Cleanup on error
    try {
      disposeImageResources(blob, imageData, 'resize-error');
      disposeImageResources(image, null, 'resize-error-image');
      disposeImageResources(resizedBlob, null, 'resize-error-resized');
    } catch (cleanupError) {
      logWithContext('WARN', 'Cleanup after resize error failed', { error: cleanupError.message });
    }
    
    // Wrap in structured error if needed
    if (error instanceof ScriptError) {
      throw error;
    }
    
    throw new ScriptError(
      ERROR_CODES.IMAGE_RESIZE, 
      `Failed to resize image: ${error.message}`, 
      { 
        fileId: fileId,
        originalError: error.message 
      }
    );
  } finally {
    // Final cleanup
    file = null;
    blob = null;
    image = null;
    imageData = null;
    resizedBlob = null;
  }
}

// ============================================================================
// UI HANDLERS & MAIN FUNCTIONS
// ============================================================================

function stopImageInsertion() {
  SCRIPT_PROPERTIES.setProperty('shouldStop', 'true');
  SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
  Logger.log('Stop flag set and processing batch cleared');
  cleanupAfterStop();
  return true;
}

/**
 * Force reset the entire system - stops all processes and clears all state
 * This is used when the system gets stuck or needs a complete reset
 */
function forceResetSystem() {
  try {
    logWithContext('INFO', 'Starting force reset of system', {});
    
    // 1. Stop any running processes
    SCRIPT_PROPERTIES.setProperty('shouldStop', 'true');
    SCRIPT_PROPERTIES.setProperty('processingBatch', 'false');
    
    // 2. Clean up all triggers
    cleanupExistingTriggers();
    
    // 3. Clear all script properties
    SCRIPT_PROPERTIES.deleteAllProperties();
    
    // 4. Clear all cache
    try {
      CacheService.getScriptCache().removeAll(['img_*', 'performance_metrics']);
      logWithContext('DEBUG', 'Cache cleared during reset', {});
    } catch (cacheError) {
      logWithContext('WARN', 'Cache cleanup failed during reset', { error: cacheError.message });
    }
    
    // 5. Force garbage collection hint
    Utilities.sleep(GARBAGE_COLLECTION_HINT_DELAY * 3);
    
    logWithContext('INFO', 'Force reset completed successfully', {});
    return { success: true, message: 'System reset successfully. You can now start a new process.' };
    
  } catch (error) {
    logWithContext('ERROR', 'Force reset failed', { error: error.message });
    return { success: false, message: `Reset failed: ${error.message}` };
  }
}

function showUI() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(650)
      .setHeight(595);
  SpreadsheetApp.getUi().showModalDialog(html, 'Insert Thumbnails');
}

function onOpen() {
  // Initialize system on first launch
  initializeSystemOnFirstLaunch();
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
      .addItem('Insert Thumbnails from GDrive', 'showUI')
      .addToUi();
}

/**
 * Initialize system on first launch to ensure clean state
 * This runs every time the spreadsheet is opened to prevent stuck states
 */
function initializeSystemOnFirstLaunch() {
  try {
    // Check if there are any existing properties that might indicate a stuck state
    const properties = SCRIPT_PROPERTIES.getProperties();
    const hasStuckState = properties.processingBatch === 'true' || 
                         properties.shouldStop === 'true' ||
                         properties.currentTriggerId ||
                         properties.fileList;
    
    if (hasStuckState) {
      logWithContext('INFO', 'Detected potential stuck state on launch, performing cleanup', {
        processingBatch: properties.processingBatch,
        shouldStop: properties.shouldStop,
        hasTriggerId: !!properties.currentTriggerId,
        hasFileList: !!properties.fileList
      });
      
      // Clean up any stuck state
      cleanupExistingTriggers();
      SCRIPT_PROPERTIES.deleteAllProperties();
      
      // Clear cache
      try {
        CacheService.getScriptCache().removeAll(['img_*', 'performance_metrics']);
      } catch (cacheError) {
        logWithContext('WARN', 'Cache cleanup failed during initialization', { error: cacheError.message });
      }
      
      logWithContext('INFO', 'System initialization cleanup completed', {});
    } else {
      logWithContext('DEBUG', 'System state is clean on launch', {});
    }
    
  } catch (error) {
    logWithContext('ERROR', 'System initialization failed', { error: error.message });
  }
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
    logWithContext('INFO', 'Starting file collection and sorting process', { folderUrl: folderUrl });
    
    // Clean up any existing state
    cleanupExistingTriggers();
    SCRIPT_PROPERTIES.deleteProperty('fileList');
    
    // Collect files with error handling
    const files = executeWithRetry(
      () => collectFiles(folderUrl),
      'collect files',
      { folderUrl: folderUrl }
    );
    
    logWithContext('DEBUG', 'Files collected', { fileCount: files ? files.length : 0 });
    
    if (!files || files.length === 0) {
      logWithContext('WARN', 'No files found in folder', { folderUrl: folderUrl });
      return ['No files were found in the specified folder'];
    }
    
    // Sort files with error handling
    const sortResult = executeWithRetry(
      () => sortFiles(),
      'sort files',
      { totalFiles: files.length }
    );
    
    logWithContext('DEBUG', 'Files sorted', { sortResult: sortResult });
    
    // Generate preview with error handling
    const preview = executeWithRetry(
      () => previewSortedFiles(),
      'generate preview',
      { totalFiles: files.length }
    );
    
    logWithContext('INFO', 'File collection and sorting completed successfully', { 
      totalFiles: files.length,
      previewLength: preview.length 
    });
    
    return preview;
    
  } catch (error) {
    logWithContext('ERROR', 'File collection and sorting failed', { 
      folderUrl: folderUrl,
      error: error instanceof ScriptError ? error.toString() : error.message 
    });
    
    // Clean up on error
    try {
      cleanupExistingTriggers();
      SCRIPT_PROPERTIES.deleteProperty('fileList');
    } catch (cleanupError) {
      logWithContext('WARN', 'Cleanup after error failed', { 
        cleanupError: cleanupError.message 
      });
    }
    
    if (error instanceof ScriptError) {
      return [`Error: ${error.message}`];
    }
    
    return [`Error: Failed to collect and sort files - ${error.message}`];
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