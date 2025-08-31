// utils.gs - Utility functions for Google Apps Script

/**
 * Format date to consistent string format
 */
function formatDateString(date) {
  if (!date) return '';
  
  if (typeof date === 'string') {
    date = new Date(date);
  }
  
  if (!(date instanceof Date) || isNaN(date)) {
    return '';
  }
  
  return date.toISOString().split('T')[0]; // YYYY-MM-DD format
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Sanitize user input to prevent injection attacks
 */
function sanitizeInput(input) {
  if (typeof input !== 'string') {
    return input;
  }
  
  return input
    .replace(/[<>]/g, '') // Remove potential HTML tags
    .trim() // Remove leading/trailing whitespace
    .substring(0, 1000); // Limit length
}

/**
 * Validate media entry data (updated for pending entries)
 */
function validateEntryData(entryData) {
  const errors = [];
  
  // Required fields
  if (!entryData.title || entryData.title.trim() === '') {
    errors.push('Title is required');
  }
  
  if (!entryData.type || entryData.type.trim() === '') {
    errors.push('Type is required');
  }
  
  // Validate item types based on date combinations
  if (entryData.status === 'unknown-dates' && (entryData.startDate || entryData.finishDate)) {
    errors.push('Unknown-dates items should not have start or finish dates');
  }

  if (entryData.status === 'completed' && !entryData.startDate && !entryData.finishDate) {
    errors.push('Completed items need at least a finish date or both dates');
  }
  
  // Valid types
  const validTypes = ['videogame', 'film', 'series', 'book', 'paper'];
  if (entryData.type && !validTypes.includes(entryData.type.toLowerCase())) {
    errors.push('Invalid content type');
  }
  
  // Valid statuses - updated list
  const validStatuses = ['pending', 'in-progress', 'in-progress-no-dates', 'completed', 'completed-no-dates'];
  if (entryData.status && !validStatuses.includes(entryData.status)) {
    errors.push('Invalid status');
  }

  // Validate status logic
  if (entryData.status === 'completed-no-dates' && (!entryData.rating || entryData.rating === '0' || entryData.rating === 0)) {
    errors.push('Items marked as completed without dates must have a rating');
  }

  if (entryData.status === 'in-progress-no-dates' && entryData.rating && entryData.rating !== '0' && entryData.rating !== 0) {
    errors.push('Items in progress without dates should not have a rating');
  }
  
  // Date validation
  if (entryData.startDate) {
    const startDate = new Date(entryData.startDate);
    if (isNaN(startDate)) {
      errors.push('Invalid start date');
    }
    
    if (entryData.finishDate) {
      const finishDate = new Date(entryData.finishDate);
      if (isNaN(finishDate)) {
        errors.push('Invalid finish date');
      } else if (finishDate < startDate) {
        errors.push('Finish date cannot be before start date');
      }
    }
  }
  
  // Rating validation
  if (entryData.rating !== undefined && entryData.rating !== null && entryData.rating !== '') {
    const rating = parseInt(entryData.rating);
    if (isNaN(rating) || rating < 1 || rating > 10) {
      errors.push('Rating must be between 1 and 10');
    }
  }
  
  // Hype rating validation (for pending items)
  if (entryData.hypeRating !== undefined && entryData.hypeRating !== null && entryData.hypeRating !== '') {
    const hypeRating = parseInt(entryData.hypeRating);
    if (isNaN(hypeRating) || hypeRating < 1 || hypeRating > 10) {
      errors.push('Hype rating must be between 1 and 10');
    }
  }
  
  // Tags validation
  if (entryData.tags && typeof entryData.tags === 'string') {
    const tags = entryData.tags.split(',').map(tag => tag.trim());
    if (tags.length > 20) {
      errors.push('Maximum 20 tags allowed');
    }
    
    const invalidTags = tags.filter(tag => tag.length > 50);
    if (invalidTags.length > 0) {
      errors.push('Tags must be 50 characters or less');
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Validate pending entry data
 */
function validatePendingEntryData(entryData) {
  const errors = [];
  
  // Required fields for pending entries
  if (!entryData.title || entryData.title.trim() === '') {
    errors.push('Title is required');
  }
  
  if (!entryData.type || entryData.type.trim() === '') {
    errors.push('Type is required');
  }
  
  // Valid types
  const validTypes = ['videogame', 'film', 'series', 'book', 'paper'];
  if (entryData.type && !validTypes.includes(entryData.type.toLowerCase())) {
    errors.push('Invalid content type');
  }
  
  // Hype rating validation
  if (entryData.hypeRating !== undefined && entryData.hypeRating !== null && entryData.hypeRating !== '') {
    const hypeRating = parseInt(entryData.hypeRating);
    if (isNaN(hypeRating) || hypeRating < 1 || hypeRating > 10) {
      errors.push('Hype rating must be between 1 and 10');
    }
  }
  
  // Tags validation
  if (entryData.tags && typeof entryData.tags === 'string') {
    const tags = entryData.tags.split(',').map(tag => tag.trim());
    if (tags.length > 20) {
      errors.push('Maximum 20 tags allowed');
    }
    
    const invalidTags = tags.filter(tag => tag.length > 50);
    if (invalidTags.length > 0) {
      errors.push('Tags must be 50 characters or less');
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Create backup of current data
 */
function createBackup() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    // Create backup sheet with timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const backupSheetName = `Backup_${timestamp}`;
    
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const backupSheet = spreadsheet.insertSheet(backupSheetName);
    
    // Copy data to backup sheet
    if (data.length > 0) {
      backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      
      // Format header row
      backupSheet.getRange(1, 1, 1, data[0].length)
        .setBackground('#4285f4')
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    logUserAction('backup_created', { backupSheet: backupSheetName });
    
    return {
      success: true,
      backupSheet: backupSheetName,
      message: 'Backup created successfully'
    };
  } catch (error) {
    console.error('Error creating backup:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Clean up old backup sheets (keep only last 5)
 */
function cleanupBackups() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    // Find backup sheets
    const backupSheets = sheets.filter(sheet => 
      sheet.getName().startsWith('Backup_')
    );
    
    // Sort by creation date (newest first)
    backupSheets.sort((a, b) => {
      const dateA = a.getName().split('_')[1];
      const dateB = b.getName().split('_')[1];
      return dateB.localeCompare(dateA);
    });
    
    let deletedCount = 0;
    // Delete old backups (keep only 5 most recent)
    for (let i = 5; i < backupSheets.length; i++) {
      spreadsheet.deleteSheet(backupSheets[i]);
      deletedCount++;
    }
    
    logUserAction('backup_cleanup', { deletedCount: deletedCount });
    
    return {
      success: true,
      deleted: deletedCount,
      message: 'Backup cleanup completed'
    };
  } catch (error) {
    console.error('Error cleaning up backups:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Export data to CSV format (updated for new fields)
 */
function exportToCSV() {
  try {
    const entries = getAllEntries();
    
    if (entries.length === 0) {
      return {
        success: false,
        error: 'No data to export'
      };
    }
    
    // Define the order of headers for CSV export
    const orderedHeaders = [
      'id', 'title', 'type', 'startdate', 'finishdate', 
      'rating', 'notes', 'coverurl', 'createdat', 
      'status', 'tags', 'hyperating'
    ];
    
    let csvContent = orderedHeaders.join(',') + '\n';
    
    entries.forEach(entry => {
      const row = orderedHeaders.map(header => {
        let value = entry[header] || '';
        
        // Handle metadata specially
        if (header === 'metadata' && typeof entry[header] === 'object') {
          value = JSON.stringify(entry[header]);
        }
        
        // Handle commas and quotes in values
        if (typeof value === 'string') {
          value = value.replace(/"/g, '""'); // Escape quotes
          if (value.includes(',') || value.includes('"') || value.includes('\n')) {
            value = `"${value}"`; // Wrap in quotes
          }
        }
        
        return value;
      });
      
      csvContent += row.join(',') + '\n';
    });
    
    // Create blob
    const blob = Utilities.newBlob(csvContent, 'text/csv', 'multimedia_diary_export.csv');
    
    logUserAction('data_exported', { entryCount: entries.length });
    
    return {
      success: true,
      blob: blob,
      entryCount: entries.length,
      message: 'Data exported successfully'
    };
  } catch (error) {
    console.error('Error exporting data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Import data from CSV (updated for new fields)
 */
function importFromCSV(csvData) {
  try {
    const lines = csvData.split('\n');
    
    if (lines.length < 2) {
      throw new Error('CSV must contain at least header and one data row');
    }
    
    // Parse headers
    const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
    
    // Validate required headers
    const requiredHeaders = ['title', 'type'];
    const missingHeaders = requiredHeaders.filter(h => 
      !headers.some(header => header.toLowerCase().includes(h.toLowerCase()))
    );
    
    if (missingHeaders.length > 0) {
      throw new Error(`Missing required columns: ${missingHeaders.join(', ')}`);
    }
    
    const results = {
      success: 0,
      errors: [],
      total: lines.length - 1,
      pending: 0,
      active: 0
    };
    
    // Process data rows
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue; // Skip empty lines
      
      try {
        const values = parseCSVLine(line);
        
        if (values.length !== headers.length) {
          throw new Error(`Row ${i}: Column count mismatch`);
        }
        
        // Create entry object
        const entry = {};
        headers.forEach((header, index) => {
          entry[header.toLowerCase()] = values[index];
        });
        
        // Map common field variations
        entry.title = entry.title || entry.name;
        entry.type = entry.type || entry.category;
        entry.startdate = entry.startdate || entry.start_date || entry.datestarted;
        entry.finishdate = entry.finishdate || entry.finish_date || entry.datefinished;
        entry.rating = entry.rating || entry.score;
        entry.status = entry.status || (entry.startdate ? 'in-progress' : 'pending');
        
        // Determine if this should be a pending or active entry
        const isPending = entry.status === 'pending' || (!entry.startdate && !entry.finishdate);
        
        let result;
        if (isPending) {
          // Validate and add as pending entry
          const validation = validatePendingEntryData(entry);
          if (!validation.isValid) {
            throw new Error(validation.errors.join(', '));
          }
          
          result = addPendingEntry(entry);
          if (result.success) {
            results.pending++;
          }
        } else {
          // Validate and add as regular entry
          const validation = validateEntryData(entry);
          if (!validation.isValid) {
            throw new Error(validation.errors.join(', '));
          }
          
          result = addMediaEntry(entry);
          if (result.success) {
            results.active++;
          }
        }
        
        if (result.success) {
          results.success++;
        } else {
          throw new Error(result.error);
        }
        
      } catch (error) {
        results.errors.push(`Row ${i}: ${error.message}`);
      }
    }
    
    logUserAction('data_imported', results);
    
    return {
      success: true,
      results: results,
      message: `Import completed: ${results.success} successful (${results.active} active, ${results.pending} pending), ${results.errors.length} errors`
    };
    
  } catch (error) {
    console.error('Error importing CSV:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Parse a single CSV line handling quotes and commas
 */
function parseCSVLine(line) {
  const values = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        // Escaped quote
        current += '"';
        i++; // Skip next quote
      } else {
        // Toggle quote state
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // End of field
      values.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  // Add final field
  values.push(current.trim());
  
  return values;
}

/**
 * Generate random UUID for entry IDs
 */
function generateUUID() {
  return Utilities.getUuid();
}

/**
 * Log user action for audit trail
 */
function logUserAction(action, details = {}) {
  try {
    const logSheet = getOrCreateLogSheet();
    
    const logEntry = [
      new Date(),
      Session.getActiveUser().getEmail(),
      action,
      JSON.stringify(cleanSensitiveData(details)),
      Session.getScriptTimeZone()
    ];
    
    logSheet.appendRow(logEntry);
  } catch (error) {
    console.error('Error logging user action:', error);
  }
}

/**
 * Get or create audit log sheet
 */
function getOrCreateLogSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let logSheet = spreadsheet.getSheetByName('AuditLog');
  
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet('AuditLog');
    
    // Add headers
    const headers = ['Timestamp', 'User', 'Action', 'Details', 'Timezone'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    logSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold');
  }
  
  return logSheet;
}

/**
 * Get user preferences or create default ones (updated)
 */
function getUserPreferences() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const prefsSheet = getOrCreatePreferencesSheet();
    const data = prefsSheet.getDataRange().getValues();
    
    // Find user preferences
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        return JSON.parse(data[i][1] || '{}');
      }
    }
    
    // Return default preferences if not found
    return {
      theme: 'dark',
      defaultView: 'room',
      notifications: true,
      autoSave: true,
      showPendingInRoom: false,
      defaultHypeRating: 7,
      maxTagsPerEntry: 10
    };
  } catch (error) {
    console.error('Error getting user preferences:', error);
    return {};
  }
}

/**
 * Save user preferences
 */
function saveUserPreferences(preferences) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const prefsSheet = getOrCreatePreferencesSheet();
    const data = prefsSheet.getDataRange().getValues();
    
    let userRow = -1;
    
    // Find existing user row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        userRow = i + 1;
        break;
      }
    }
    
    const prefsJson = JSON.stringify(preferences);
    
    if (userRow > 0) {
      // Update existing row
      prefsSheet.getRange(userRow, 2).setValue(prefsJson);
    } else {
      // Add new row
      prefsSheet.appendRow([userEmail, prefsJson]);
    }
    
    logUserAction('preferences_saved');
    
    return {
      success: true,
      message: 'Preferences saved successfully'
    };
  } catch (error) {
    console.error('Error saving user preferences:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get or create preferences sheet
 */
function getOrCreatePreferencesSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let prefsSheet = spreadsheet.getSheetByName('UserPreferences');
  
  if (!prefsSheet) {
    prefsSheet = spreadsheet.insertSheet('UserPreferences');
    
    // Add headers
    const headers = ['UserEmail', 'Preferences'];
    prefsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    prefsSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold');
  }
  
  return prefsSheet;
}

/**
 * Performance monitoring - measure function execution time
 */
function measurePerformance(functionName, func) {
  const startTime = new Date().getTime();
  
  try {
    const result = func();
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    console.log(`Performance: ${functionName} took ${duration}ms`);
    
    // Log slow operations
    if (duration > 1000) {
      logUserAction('slow_operation', {
        function: functionName,
        duration: duration
      });
    }
    
    return result;
  } catch (error) {
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    console.error(`Performance: ${functionName} failed after ${duration}ms`, error);
    throw error;
  }
}

/**
 * Test function to verify cover update isolation
 */
function testCoverUpdateIsolation() {
  try {
    // 1. Create a pending entry
    const pendingEntry = {
      title: 'Test Book for Cover Update',
      type: 'book',
      status: 'pending'
    };
    
    const addResult = addPendingEntry(pendingEntry);
    if (!addResult.success) {
      throw new Error('Failed to create test entry: ' + addResult.error);
    }
    
    const entryId = addResult.id;
    
    // 2. Get the entry to verify initial state
    const entries = getAllEntries();
    const testEntry = entries.find(e => e.id === entryId);
    
    if (!testEntry) {
      throw new Error('Test entry not found after creation');
    }
    
    const originalStatus = testEntry.status;
    const originalRating = testEntry.rating;
    
    console.log(`Initial state - Status: ${originalStatus}, Rating: ${originalRating}`);
    
    // 3. Request new cover
    const coverResult = requestNewCover(entryId);
    if (!coverResult.success) {
      throw new Error('Cover update failed: ' + coverResult.error);
    }
    
    // 4. Verify the entry state after cover update
    const updatedEntries = getAllEntries();
    const updatedEntry = updatedEntries.find(e => e.id === entryId);
    
    if (!updatedEntry) {
      throw new Error('Updated entry not found');
    }
    
    // 5. Assert that only cover-related fields changed
    if (updatedEntry.status !== originalStatus) {
      throw new Error(`STATUS CHANGE BUG: Status changed from ${originalStatus} to ${updatedEntry.status}`);
    }
    
    if (updatedEntry.rating !== originalRating) {
      throw new Error(`RATING CHANGE BUG: Rating changed from ${originalRating} to ${updatedEntry.rating}`);
    }
    
    // 6. Clean up
    deleteEntry(entryId);
    
    return { 
      success: true, 
      message: 'Cover update isolation test passed - only cover fields were modified'
    };
    
  } catch (error) {
    console.error('Cover update isolation test failed:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Rate limiting helper
 */
function checkRateLimit(operation, maxPerHour = 100) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const now = new Date();
    const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
    
    const logSheet = getOrCreateLogSheet();
    const data = logSheet.getDataRange().getValues();
    
    // Count operations in last hour
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const logTime = new Date(data[i][0]);
      const logUser = data[i][1];
      const logAction = data[i][2];
      
      if (logUser === userEmail && 
          logAction === operation && 
          logTime > oneHourAgo) {
        count++;
      }
    }
    
    return {
      allowed: count < maxPerHour,
      current: count,
      limit: maxPerHour,
      resetTime: new Date(now.getTime() + 60 * 60 * 1000)
    };
  } catch (error) {
    console.error('Error checking rate limit:', error);
    return { allowed: true, current: 0, limit: maxPerHour };
  }
}

/**
 * Send email notification (optional feature)
 */
function sendNotificationEmail(to, subject, body) {
  try {
    // Check if user has notification preferences enabled
    const prefs = getUserPreferences();
    if (!prefs.notifications) {
      return { success: true, message: 'Notifications disabled by user' };
    }
    
    // Rate limit email notifications
    const rateLimit = checkRateLimit('email_notification', 10);
    if (!rateLimit.allowed) {
      return { 
        success: false, 
        error: 'Email rate limit exceeded',
        resetTime: rateLimit.resetTime
      };
    }
    
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      htmlBody: body
    });
    
    logUserAction('email_notification', { to: to, subject: subject });
    
    return { success: true, message: 'Email sent successfully' };
  } catch (error) {
    console.error('Error sending email:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Clean sensitive data for logging
 */
function cleanSensitiveData(data) {
  const cleaned = { ...data };
  
  // Remove sensitive fields
  const sensitiveFields = ['password', 'token', 'key', 'secret', 'email'];
  sensitiveFields.forEach(field => {
    if (cleaned[field]) {
      cleaned[field] = '[REDACTED]';
    }
  });
  
  // Truncate long strings
  Object.keys(cleaned).forEach(key => {
    if (typeof cleaned[key] === 'string' && cleaned[key].length > 200) {
      cleaned[key] = cleaned[key].substring(0, 200) + '... [TRUNCATED]';
    }
  });
  
  return cleaned;
}

/**
 * Validate and sanitize URL
 */
function validateURL(url) {
  if (!url || typeof url !== 'string') {
    return { isValid: false, error: 'URL is required and must be a string' };
  }
  
  try {
    const urlObj = new URL(url);
    
    // Only allow http and https protocols
    if (!['http:', 'https:'].includes(urlObj.protocol)) {
      return { isValid: false, error: 'Only HTTP and HTTPS URLs are allowed' };
    }
    
    // Block potentially dangerous domains
    const blockedDomains = ['localhost', '127.0.0.1', '0.0.0.0'];
    if (blockedDomains.some(domain => urlObj.hostname.includes(domain))) {
      return { isValid: false, error: 'Domain not allowed' };
    }
    
    return { isValid: true, cleanUrl: urlObj.toString() };
  } catch (error) {
    return { isValid: false, error: 'Invalid URL format' };
  }
}

/**
 * Get application configuration (updated)
 */
function getAppConfig() {
  return {
    version: '2.0.0',
    maxEntriesPerUser: 2000,
    maxPendingEntriesPerUser: 500,
    supportedTypes: ['videogame', 'film', 'series', 'book', 'paper'],
    supportedStatuses: ['pending', 'in-progress', 'completed', 'unknown-dates'],
    rateLimit: {
      addEntry: 50,
      addPendingEntry: 100,
      updateEntry: 100,
      deleteEntry: 20,
      startPendingEntry: 50
    },
    features: {
      metadata: true,
      backups: true,
      export: true,
      notifications: true,
      pendingList: true,
      tagging: true,
      hypeRating: true,
      search: true
    },
    validation: {
      maxTitleLength: 200,
      maxNotesLength: 2000,
      maxTagLength: 50,
      maxTagsPerEntry: 20
    }
  };
}

/**
 * Health check function (updated)
 */
function healthCheck() {
  const results = {
    timestamp: new Date().toISOString(),
    status: 'healthy',
    checks: {}
  };
  
  try {
    // Check sheet access
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const totalRows = sheet.getLastRow();
    const pendingEntries = getAllEntries().filter(e => e.status === 'pending').length;
    const activeEntries = getAllEntries().filter(e => e.status !== 'pending').length;
    
    results.checks.sheetAccess = { 
      status: 'pass', 
      message: 'Sheet accessible',
      totalRows: totalRows,
      pendingEntries: pendingEntries,
      activeEntries: activeEntries
    };
  } catch (error) {
    results.checks.sheetAccess = { 
      status: 'fail', 
      error: error.message 
    };
    results.status = 'unhealthy';
  }
  
  try {
    // Check script permissions
    const user = Session.getActiveUser().getEmail();
    results.checks.permissions = { 
      status: 'pass', 
      user: user 
    };
  } catch (error) {
    results.checks.permissions = { 
      status: 'fail', 
      error: error.message 
    };
    results.status = 'degraded';
  }
  
  try {
    // Check external API availability (optional)
    const testUrl = 'https://www.googleapis.com/books/v1/volumes?q=test';
    const response = UrlFetchApp.fetch(testUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    results.checks.externalAPI = {
      status: response.getResponseCode() === 200 ? 'pass' : 'fail',
      responseCode: response.getResponseCode()
    };
  } catch (error) {
    results.checks.externalAPI = { 
      status: 'fail', 
      error: error.message 
    };
  }
  
  // Check data integrity
  try {
    const entries = getAllEntries();
    const statusCounts = {
      pending: 0,
      'in-progress': 0,
      completed: 0,
      'unknown-dates': 0
    };
    
    entries.forEach(entry => {
      statusCounts[entry.status] = (statusCounts[entry.status] || 0) + 1;
    });
    
    results.checks.dataIntegrity = {
      status: 'pass',
      totalEntries: entries.length,
      statusBreakdown: statusCounts
    };
  } catch (error) {
    results.checks.dataIntegrity = {
      status: 'fail',
      error: error.message
    };
    if (results.status === 'healthy') results.status = 'degraded';
  }
  
  return results;
}

/**
 * Debug function to test all major functions (updated)
 */
function runDiagnostics() {
  const results = {
    timestamp: new Date().toISOString(),
    tests: []
  };
  
  // Test 1: Health Check
  try {
    const health = healthCheck();
    results.tests.push({
      name: 'Health Check',
      status: health.status === 'healthy' ? 'pass' : 'fail',
      details: health
    });
  } catch (error) {
    results.tests.push({
      name: 'Health Check',
      status: 'fail',
      error: error.message
    });
  }
  
  // Test 2: Add/Get/Delete Regular Entry
  try {
    const testEntry = {
      title: 'Diagnostic Test Entry',
      type: 'book',
      startDate: '2024-01-01',
      rating: 5,
      notes: 'This is a test entry for diagnostics'
    };
    
    const addResult = addMediaEntry(testEntry);
    if (!addResult.success) {
      throw new Error('Failed to add entry: ' + addResult.error);
    }
    
    const entries = getAllEntries();
    const foundEntry = entries.find(e => e.id === addResult.id);
    if (!foundEntry) {
      throw new Error('Added entry not found in list');
    }
    
    const deleteResult = deleteEntry(addResult.id);
    if (!deleteResult.success) {
      throw new Error('Failed to delete entry: ' + deleteResult.error);
    }
    
    results.tests.push({
      name: 'Regular Entry CRUD Operations',
      status: 'pass',
      details: 'Add, retrieve, and delete operations successful'
    });
  } catch (error) {
    results.tests.push({
      name: 'Regular Entry CRUD Operations',
      status: 'fail',
      error: error.message
    });
  }
  
  // Test 3: Add/Get/Start Pending Entry
  try {
    const testPendingEntry = {
      title: 'Diagnostic Test Pending Entry',
      type: 'film',
      tags: 'test, diagnostic',
      hypeRating: 8,
      notes: 'This is a test pending entry for diagnostics'
    };
    
    const addResult = addPendingEntry(testPendingEntry);
    if (!addResult.success) {
      throw new Error('Failed to add pending entry: ' + addResult.error);
    }
    
    const entries = getAllEntries();
    const foundEntry = entries.find(e => e.id === addResult.id);
    if (!foundEntry || foundEntry.status !== 'pending') {
      throw new Error('Added pending entry not found or has wrong status');
    }
    
    const startResult = startPendingEntry(addResult.id);
    if (!startResult.success) {
      throw new Error('Failed to start pending entry: ' + startResult.error);
    }
    
    // Verify it's now in-progress
    const updatedEntries = getAllEntries();
    const startedEntry = updatedEntries.find(e => e.id === addResult.id);
    if (!startedEntry || startedEntry.status !== 'in-progress') {
      throw new Error('Started entry not found or has wrong status');
    }
    
    const deleteResult = deleteEntry(addResult.id);
    if (!deleteResult.success) {
      throw new Error('Failed to delete started entry: ' + deleteResult.error);
    }
    
    results.tests.push({
      name: 'Pending Entry Operations',
      status: 'pass',
      details: 'Add pending, start, and delete operations successful'
    });
  } catch (error) {
    results.tests.push({
      name: 'Pending Entry Operations',
      status: 'fail',
      error: error.message
    });
  }
  
  // Test 4: Statistics
  try {
    const stats = getStatistics();
    results.tests.push({
      name: 'Statistics Generation',
      status: 'pass',
      details: stats
    });
  } catch (error) {
    results.tests.push({
      name: 'Statistics Generation',
      status: 'fail',
      error: error.message
    });
  }
  
  // Test 5: Backup Creation
  try {
    const backup = createBackup();
    if (backup.success) {
      // Clean up test backup
      const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
      const backupSheet = spreadsheet.getSheetByName(backup.backupSheet);
      if (backupSheet) {
        spreadsheet.deleteSheet(backupSheet);
      }
    }
    
    results.tests.push({
      name: 'Backup Creation',
      status: backup.success ? 'pass' : 'fail',
      details: backup.success ? 'Backup created and cleaned up' : backup.error
    });
  } catch (error) {
    results.tests.push({
      name: 'Backup Creation',
      status: 'fail',
      error: error.message
    });
  }
  
  // Test 6: Validation Functions
  try {
    // Test regular entry validation
    const validEntry = {
      title: 'Valid Entry',
      type: 'book',
      startDate: '2024-01-01',
      rating: 8
    };
    
    const validation1 = validateEntryData(validEntry);
    if (!validation1.isValid) {
      throw new Error('Valid entry failed validation: ' + validation1.errors.join(', '));
    }
    
    // Test pending entry validation
    const validPendingEntry = {
      title: 'Valid Pending Entry',
      type: 'film',
      tags: 'action, adventure',
      hypeRating: 9
    };
    
    const validation2 = validatePendingEntryData(validPendingEntry);
    if (!validation2.isValid) {
      throw new Error('Valid pending entry failed validation: ' + validation2.errors.join(', '));
    }
    
    // Test invalid entry
    const invalidEntry = {
      title: '',
      type: 'invalid_type',
      rating: 15
    };
    
    const validation3 = validateEntryData(invalidEntry);
    if (validation3.isValid) {
      throw new Error('Invalid entry passed validation when it should have failed');
    }
    
    results.tests.push({
      name: 'Validation Functions',
      status: 'pass',
      details: 'Entry and pending entry validation working correctly'
    });
  } catch (error) {
    results.tests.push({
      name: 'Validation Functions',
      status: 'fail',
      error: error.message
    });
  }
  
  // Test 7: User Preferences
  try {
    const originalPrefs = getUserPreferences();
    const testPrefs = {
      theme: 'light',
      defaultView: 'pending',
      notifications: false,
      defaultHypeRating: 5
    };
    
    const saveResult = saveUserPreferences(testPrefs);
    if (!saveResult.success) {
      throw new Error('Failed to save preferences: ' + saveResult.error);
    }
    
    const loadedPrefs = getUserPreferences();
    if (loadedPrefs.theme !== testPrefs.theme || loadedPrefs.defaultView !== testPrefs.defaultView) {
      throw new Error('Saved preferences do not match loaded preferences');
    }
    
    // Restore original preferences
    saveUserPreferences(originalPrefs);
    
    results.tests.push({
      name: 'User Preferences',
      status: 'pass',
      details: 'Save and load preferences working correctly'
    });
  } catch (error) {
    results.tests.push({
      name: 'User Preferences',
      status: 'fail',
      error: error.message
    });
  }
  
  return results;
}

/**
 * Initialize application with sample data for testing (updated with pending entries)
 */
function initializeWithSampleData() {
  try {
    // Initialize sheet first
    initializeSheet();
    
    const sampleActiveEntries = [
      {
        title: 'The Legend of Zelda: Breath of the Wild',
        type: 'videogame',
        startDate: '2024-01-15',
        finishDate: '2024-02-28',
        rating: 10,
        notes: 'Incredible open-world adventure game with amazing physics and exploration.'
      },
      {
        title: 'Inception',
        type: 'film',
        startDate: '2024-02-01',
        finishDate: '2024-02-01',
        rating: 9,
        notes: 'Mind-bending thriller about dreams within dreams.'
      },
      {
        title: 'Breaking Bad',
        type: 'series',
        startDate: '2024-01-01',
        rating: 10,
        notes: 'Currently watching this amazing crime drama series.'
      },
      {
        title: 'Sapiens: A Brief History of Humankind',
        type: 'book',
        startDate: '2024-02-10',
        finishDate: '2024-03-15',
        rating: 8,
        notes: 'Fascinating look at human history and development.'
      },
      {
        title: 'Attention Is All You Need',
        type: 'paper',
        startDate: '2024-03-01',
        rating: 9,
        notes: 'Groundbreaking paper on transformer architecture in deep learning.'
      },
      {
        title: 'The Witcher 3: Wild Hunt',
        type: 'videogame',
        // No dates - this will be 'unknown-dates' status
        rating: 9,
        notes: 'Amazing RPG with incredible storytelling and world-building.'
      }
    ];
    
    const samplePendingEntries = [
      {
        title: 'Elden Ring',
        type: 'videogame',
        tags: 'souls-like, open-world, fantasy',
        hypeRating: 10,
        notes: 'The latest From Software masterpiece everyone is talking about.'
      },
      {
        title: 'Dune: Part Two',
        type: 'film',
        tags: 'sci-fi, epic, denis-villeneuve',
        hypeRating: 9,
        notes: 'Really excited to see how they adapt the second half of the book.'
      },
      {
        title: 'The Bear',
        type: 'series',
        tags: 'comedy-drama, cooking, award-winning',
        hypeRating: 8,
        notes: 'Heard amazing things about this show about a Chicago restaurant.'
      },
      {
        title: 'Klara and the Sun',
        type: 'book',
        tags: 'kazuo-ishiguro, ai, literary-fiction',
        hypeRating: 7,
        notes: 'New Ishiguro novel about artificial intelligence and consciousness.'
      },
      {
        title: 'GPT-4 Technical Report',
        type: 'paper',
        tags: 'ai, llm, openai, technical',
        hypeRating: 9,
        notes: 'Need to read this important paper on the latest GPT model.'
      },
      {
        title: 'The Last of Us Part II',
        type: 'videogame',
        tags: 'action-adventure, post-apocalyptic, story-driven',
        hypeRating: 8,
        notes: 'Controversial but critically acclaimed sequel.'
      }
    ];
    
    const results = {
      activeSuccess: 0,
      pendingSuccess: 0,
      errors: [],
      total: sampleActiveEntries.length + samplePendingEntries.length
    };
    
    // Add active entries
    sampleActiveEntries.forEach((entry, index) => {
      try {
        const result = addMediaEntry(entry);
        if (result.success) {
          results.activeSuccess++;
        } else {
          results.errors.push(`Active Entry ${index + 1}: ${result.error}`);
        }
      } catch (error) {
        results.errors.push(`Active Entry ${index + 1}: ${error.message}`);
      }
    });
    
    // Add pending entries
    samplePendingEntries.forEach((entry, index) => {
      try {
        const result = addPendingEntry(entry);
        if (result.success) {
          results.pendingSuccess++;
        } else {
          results.errors.push(`Pending Entry ${index + 1}: ${result.error}`);
        }
      } catch (error) {
        results.errors.push(`Pending Entry ${index + 1}: ${error.message}`);
      }
    });
    
    logUserAction('sample_data_initialized', {
      activeEntries: results.activeSuccess,
      pendingEntries: results.pendingSuccess,
      totalErrors: results.errors.length
    });
    
    return {
      success: true,
      message: `Sample data initialized: ${results.activeSuccess} active entries, ${results.pendingSuccess} pending entries, ${results.errors.length} errors`,
      details: results
    };
  } catch (error) {
    console.error('Error initializing sample data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Clean up all sample/test data
 */
function cleanupSampleData() {
  try {
    const entries = getAllEntries();
    const sampleTitles = [
      'The Legend of Zelda: Breath of the Wild',
      'Inception', 
      'Breaking Bad',
      'Sapiens: A Brief History of Humankind',
      'Attention Is All You Need',
      'The Witcher 3: Wild Hunt',
      'Elden Ring',
      'Dune: Part Two',
      'The Bear',
      'Klara and the Sun',
      'GPT-4 Technical Report',
      'The Last of Us Part II',
      'Diagnostic Test Entry',
      'Diagnostic Test Pending Entry'
    ];
    
    let deletedCount = 0;
    const errors = [];
    
    entries.forEach(entry => {
      if (sampleTitles.includes(entry.title)) {
        try {
          const result = deleteEntry(entry.id);
          if (result.success) {
            deletedCount++;
          } else {
            errors.push(`Failed to delete "${entry.title}": ${result.error}`);
          }
        } catch (error) {
          errors.push(`Error deleting "${entry.title}": ${error.message}`);
        }
      }
    });
    
    logUserAction('sample_data_cleanup', {
      deletedCount: deletedCount,
      errorCount: errors.length
    });
    
    return {
      success: true,
      deletedCount: deletedCount,
      errors: errors,
      message: `Cleanup completed: ${deletedCount} sample entries deleted, ${errors.length} errors`
    };
  } catch (error) {
    console.error('Error cleaning up sample data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get comprehensive application statistics
 */
function getComprehensiveStats() {
  try {
    const entries = getAllEntries();
    
    const stats = {
      overview: {
        total: entries.length,
        pending: entries.filter(e => e.status === 'pending').length,
        inProgress: entries.filter(e => e.status === 'in-progress').length,
        completed: entries.filter(e => e.status === 'completed').length,
        unknownDates: entries.filter(e => e.status === 'unknown-dates').length
      },
      byType: {},
      byStatus: {},
      ratings: {
        average: 0,
        distribution: {},
        totalRated: 0
      },
      hypeRatings: {
        average: 0,
        distribution: {},
        totalRated: 0
      },
      tags: {
        mostUsed: [],
        total: 0
      },
      temporal: {
        entriesThisMonth: 0,
        entriesThisWeek: 0,
        completionsThisMonth: 0
      }
    };
    
    // Calculate type and status distributions
    entries.forEach(entry => {
      stats.byType[entry.type] = (stats.byType[entry.type] || 0) + 1;
      stats.byStatus[entry.status] = (stats.byStatus[entry.status] || 0) + 1;
    });
    
    // Calculate rating statistics
    const ratedEntries = entries.filter(e => e.rating && e.rating > 0);
    if (ratedEntries.length > 0) {
      const ratingsSum = ratedEntries.reduce((sum, entry) => sum + entry.rating, 0);
      stats.ratings.average = (ratingsSum / ratedEntries.length).toFixed(1);
      stats.ratings.totalRated = ratedEntries.length;
      
      // Rating distribution
      ratedEntries.forEach(entry => {
        stats.ratings.distribution[entry.rating] = (stats.ratings.distribution[entry.rating] || 0) + 1;
      });
    }
    
    // Calculate hype rating statistics
    const hypeRatedEntries = entries.filter(e => e.hyperating && e.hyperating > 0);
    if (hypeRatedEntries.length > 0) {
      const hypeRatingsSum = hypeRatedEntries.reduce((sum, entry) => sum + entry.hyperating, 0);
      stats.hypeRatings.average = (hypeRatingsSum / hypeRatedEntries.length).toFixed(1);
      stats.hypeRatings.totalRated = hypeRatedEntries.length;
      
      // Hype rating distribution
      hypeRatedEntries.forEach(entry => {
        stats.hypeRatings.distribution[entry.hyperating] = (stats.hypeRatings.distribution[entry.hyperating] || 0) + 1;
      });
    }
    
    // Calculate tag statistics
    const tagCounts = {};
    entries.forEach(entry => {
      if (entry.tags) {
        const tags = entry.tags.split(',').map(tag => tag.trim().toLowerCase());
        tags.forEach(tag => {
          if (tag) {
            tagCounts[tag] = (tagCounts[tag] || 0) + 1;
          }
        });
      }
    });
    
    stats.tags.total = Object.keys(tagCounts).length;
    stats.tags.mostUsed = Object.entries(tagCounts)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10)
      .map(([tag, count]) => ({ tag, count }));
    
    // Calculate temporal statistics
    const now = new Date();
    const thisMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const thisWeek = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    
    entries.forEach(entry => {
      const createdDate = new Date(entry.createdat);
      if (createdDate >= thisMonth) {
        stats.temporal.entriesThisMonth++;
      }
      if (createdDate >= thisWeek) {
        stats.temporal.entriesThisWeek++;
      }
      
      if (entry.finishdate) {
        const finishDate = new Date(entry.finishdate);
        if (finishDate >= thisMonth) {
          stats.temporal.completionsThisMonth++;
        }
      }
    });
    
    return stats;
  } catch (error) {
    console.error('Error getting comprehensive stats:', error);
    return null;
  }
}

/**
 * Batch refresh metadata for all entries or specific entries
 * @param {Array<string>} [entryIds] - Optional array of entry IDs to refresh. If not provided, refreshes all entries.
 * @returns {Object} - Result object with success count, error count, and details
 */
function batchRefreshMetadata(entryIds) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { 
        success: true, 
        refreshed: 0, 
        errors: 0, 
        details: [],
        message: 'No entries to refresh' 
      };
    }
    
    const headers = data[0];
    const results = {
      refreshed: 0,
      errors: 0,
      details: []
    };
    
    // Determine which entries to process
    const entriesToProcess = entryIds ? entryIds : null;
    
    for (let i = 1; i < data.length; i++) {
      const entryId = data[i][0];
      
      // Skip if entryIds is provided and this entry is not in the list
      if (entriesToProcess && !entriesToProcess.includes(entryId)) {
        continue;
      }
      
      try {
        // Get entry data
        const entry = {};
        headers.forEach((header, index) => {
          entry[header.toLowerCase()] = data[i][index];
        });
        
        // Skip if no title or type
        if (!entry.title || !entry.type) {
          results.details.push({
            entryId: entryId,
            title: entry.title || 'Unknown',
            error: 'Missing title or type'
          });
          results.errors++;
          continue;
        }
        
        // Fetch fresh metadata
        const newMetadata = fetchMetadata(entry.title, entry.type);
        
        // Prepare update data
        const updateData = {
          coverurl: newMetadata.coverURL,
          metadata: JSON.stringify(newMetadata)
        };
        
        // Update the entry
        const updateResult = updateEntry(entryId, updateData);
        
        if (updateResult.success) {
          results.refreshed++;
          results.details.push({
            entryId: entryId,
            title: entry.title,
            success: true,
            message: 'Metadata refreshed successfully'
          });
        } else {
          results.errors++;
          results.details.push({
            entryId: entryId,
            title: entry.title,
            error: updateResult.error || 'Update failed'
          });
        }
        
      } catch (error) {
        results.errors++;
        results.details.push({
          entryId: entryId,
          title: data[i][headers.indexOf('Title')] || 'Unknown',
          error: error.message
        });
      }
    }
    
    // Log the batch refresh operation
    logUserAction('batch_metadata_refresh', {
      totalProcessed: results.refreshed + results.errors,
      refreshed: results.refreshed,
      errors: results.errors,
      specificEntries: !!entryIds
    });
    
    return {
      success: true,
      refreshed: results.refreshed,
      errors: results.errors,
      details: results.details,
      message: `Batch refresh completed: ${results.refreshed} refreshed, ${results.errors} errors`
    };
    
  } catch (error) {
    console.error('Error in batch refresh metadata:', error);
    return {
      success: false,
      refreshed: 0,
      errors: 1,
      details: [{
        error: error.message
      }],
      message: `Batch refresh failed: ${error.message}`
    };
  }
}

/**
 * Get entries that need metadata refresh (e.g., old metadata or missing covers)
 * @returns {Array<string>} - Array of entry IDs that need metadata refresh
 */
function getEntriesNeedingRefresh() {
  try {
    const entries = getAllEntries();
    const needingRefresh = [];
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    
    entries.forEach(entry => {
      let needsRefresh = false;
      
      // Check if metadata is missing or invalid
      if (!entry.metadata || entry.metadata === '{}') {
        needsRefresh = true;
      }
      
      // Check if cover URL is missing
      if (!entry.coverurl || entry.coverurl === '') {
        needsRefresh = true;
      }
      
      // Check if metadata is old (older than 30 days)
      if (entry.metadata && entry.metadata !== '{}') {
        try {
          const metadata = typeof entry.metadata === 'string' ? JSON.parse(entry.metadata) : entry.metadata;
          if (metadata.fetchedAt) {
            const fetchedDate = new Date(metadata.fetchedAt);
            if (fetchedDate < thirtyDaysAgo) {
              needsRefresh = true;
            }
          }
        } catch (e) {
          needsRefresh = true;
        }
      }
      
      if (needsRefresh) {
        needingRefresh.push(entry.id);
      }
    });
    
    return needingRefresh;
  } catch (error) {
    console.error('Error getting entries needing refresh:', error);
    return [];
  }
}

/**
 * Get books with low-quality covers that need improvement
 * @returns {Array<Object>} - Array of book entries with poor quality covers
 */
function getBooksWithLowQualityCovers() {
  try {
    const entries = getAllEntries();
    const lowQualityBooks = [];
    
    entries.forEach(entry => {
      if (entry.type === 'book' && entry.coverurl) {
        const isLowQuality = isLowQualityCover(entry.coverurl);
        if (isLowQuality) {
          lowQualityBooks.push({
            id: entry.id,
            title: entry.title,
            currentCover: entry.coverurl,
            author: entry.author || '',
            reason: getLowQualityReason(entry.coverurl)
          });
        }
      }
    });
    
    return lowQualityBooks;
  } catch (error) {
    console.error('Error getting books with low quality covers:', error);
    return [];
  }
}

/**
 * Check if a cover URL is likely to be low quality
 * @param {string} coverUrl - The cover URL to check
 * @returns {boolean} - True if the cover is likely low quality
 */
function isLowQualityCover(coverUrl) {
  if (!coverUrl || coverUrl === '') {
    return true; // No cover is considered low quality
  }
  
  // Check for known low-quality indicators
  const lowQualityIndicators = [
    // Only specific low-quality Google Books patterns
    'books.google.com/books/content', // Only the lowest quality Google pattern
    'zoom=1', // Very low zoom level
    'w=90',   // Very small width (under 100px)
    // Placeholder patterns
    'via.placeholder.com',
    'placehold.co',
    'dummyimage.com'
  ];
  
  return lowQualityIndicators.some(indicator => 
    coverUrl.toLowerCase().includes(indicator.toLowerCase())
  );
}

/**
 * Get the reason why a cover is considered low quality
 * @param {string} coverUrl - The cover URL to analyze
 * @returns {string} - Reason for low quality
 */
function getLowQualityReason(coverUrl) {
  if (!coverUrl || coverUrl === '') {
    return 'No cover available';
  }
  
  const lowerUrl = coverUrl.toLowerCase();
  
  if (lowerUrl.includes('zoom=1')) {
    return 'Very low resolution zoom level';
  }
  
  if (lowerUrl.includes('w=90')) {
    return 'Very small image dimensions (under 100px)';
  }
  
  if (lowerUrl.includes('via.placeholder.com') || lowerUrl.includes('placehold.co')) {
    return 'Generic placeholder image';
  }
  
  if (lowerUrl.includes('books.google.com/books/content')) {
    return 'Very low quality Google Books image';
  }
  
  return 'Unknown quality issue';
}

/**
 * Refresh book covers specifically for books with low quality covers
 * @returns {Object} - Result object with refresh statistics
 */
function refreshBookCovers() {
  try {
    const lowQualityBooks = getBooksWithLowQualityCovers();
    
    if (lowQualityBooks.length === 0) {
      return {
        success: true,
        refreshed: 0,
        errors: 0,
        message: 'No books with low quality covers found'
      };
    }
    
    const bookIds = lowQualityBooks.map(book => book.id);
    const result = batchRefreshMetadata(bookIds);
    
    // Add book-specific information to the result
    result.booksProcessed = lowQualityBooks.length;
    result.bookDetails = lowQualityBooks;
    
    logUserAction('book_covers_refreshed', {
      booksProcessed: lowQualityBooks.length,
      refreshed: result.refreshed,
      errors: result.errors
    });
    
    return result;
  } catch (error) {
    console.error('Error refreshing book covers:', error);
    return {
      success: false,
      refreshed: 0,
      errors: 1,
      message: `Book cover refresh failed: ${error.message}`
    };
  }
}

/**
 * Manually set a custom cover URL for a book entry
 * @param {string} entryId - The ID of the book entry
 * @param {string} coverUrl - The new cover URL
 * @returns {Object} - Result object
 */
function setCustomBookCover(entryId, coverUrl) {
  try {
    // Validate the cover URL
    if (!coverUrl || coverUrl.trim() === '') {
      return { success: false, error: 'Cover URL is required' };
    }
    
    // Basic URL validation
    if (!coverUrl.startsWith('http://') && !coverUrl.startsWith('https://')) {
      return { success: false, error: 'Cover URL must start with http:// or https://' };
    }
    
    // Test if the URL is accessible
    const testResponse = UrlFetchApp.fetch(coverUrl, {
      method: 'HEAD',
      muteHttpExceptions: true
    });
    
    if (testResponse.getResponseCode() !== 200) {
      return { success: false, error: 'Cover URL is not accessible' };
    }
    
    // Check if it's actually an image
    const contentType = testResponse.getHeaders()['Content-Type'] || '';
    if (!contentType.startsWith('image/')) {
      return { success: false, error: 'URL does not point to an image' };
    }
    
    // Update the entry with the new cover URL
    const updateData = {
      coverurl: coverUrl
    };
    
    const updateResult = updateEntry(entryId, updateData);
    
    if (updateResult.success) {
      logUserAction('custom_book_cover_set', { entryId: entryId });
      return {
        success: true,
        message: 'Custom cover set successfully',
        coverUrl: coverUrl
      };
    } else {
      return updateResult;
    }
    
  } catch (error) {
    console.error('Error setting custom book cover:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
