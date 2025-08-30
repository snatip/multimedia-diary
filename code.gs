// Code.gs - Main Google Apps Script backend functions

// Configuration - Replace with your actual Google Sheets ID
const SHEET_ID = '1TExyILJ0mnE0Yb2rSEaZDxkNFbNh5my_DMF5niDd5zE';
const SHEET_NAME = 'MediaEntries';

/**
 * Main entry point for the web app
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('My Multimedia Diary')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl("https://em-content.zobj.net/source/apple/419/ledger_1f4d2.png");
}

/**
 * Include HTML files for templating
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialize the Google Sheet with proper headers
 */
function initializeSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Check if both 'ID' and 'Author' headers exist
    if (headers.indexOf('ID') === -1 || headers.indexOf('Author') === -1) {
      const newHeaders = [
        'ID', 'Title', 'Type', 'Author', 'StartDate', 'FinishDate', 
        'Rating', 'Notes', 'CoverURL', 'Metadata', 'CreatedAt',
        'Status', 'Tags', 'HypeRating'
      ];
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      
      sheet.getRange(1, 1, 1, newHeaders.length)
        .setBackground('#4285f4')
        .setFontColor('white')
        .setFontWeight('bold');
    }
    
    return { success: true, message: 'Sheet initialized successfully' };
  } catch (error) {
    console.error('Error initializing sheet:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Add a new media entry to the sheet (Refactored for maintainability)
 */
/**
 * Add a new media entry to the sheet (Refactored for maintainability)
 */
function addMediaEntry(entryData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Determine item type and set appropriate status
    let itemStatus;
    if (!entryData.startDate && !entryData.finishDate) {
      // Type 1: No remembered dates
      if (entryData.rating && entryData.rating !== '' && entryData.rating !== 'N/A') {
        // Type 1a: No dates but has rating = finished without recorded dates
        itemStatus = 'completed-no-dates';
      } else {
        // Type 1b: No dates, no rating = in progress without recorded dates
        itemStatus = 'in-progress-no-dates';
      }
    } else if (!entryData.startDate && entryData.finishDate) {
      // Type 2: Unknown start, known finish
      itemStatus = 'completed';
    } else if (entryData.startDate && !entryData.finishDate) {
      // Type 3: In progress with known start
      itemStatus = 'in-progress';  
    } else if (entryData.startDate && entryData.finishDate) {
      // Type 4: Full date info
      itemStatus = 'completed';
    } else {
      itemStatus = 'in-progress-no-dates'; // fallback
    }

    if (!entryData.title || !entryData.type) {
      throw new Error('Missing required fields: title or type');
    }
    
    const metadata = fetchMetadata(entryData.title, entryData.type);
    
    // --- THIS IS THE FIX ---
    // Handle the '0' rating from the form to save it as 'N/A'
    let ratingToSave = entryData.rating || '';
    if (entryData.rating === 0 || entryData.rating === '0') {
      ratingToSave = 'N/A';
    }

    const entryMap = {
      id: Utilities.getUuid(),
      createdat: new Date(),
      title: entryData.title,
      type: entryData.type,
      author: entryData.author || '',
      startdate: entryData.startDate ? new Date(entryData.startDate) : '',
      finishdate: entryData.finishDate ? new Date(entryData.finishDate) : '',
      rating: ratingToSave,
      notes: entryData.notes || '',
      coverurl: metadata.coverURL || '',
      metadata: JSON.stringify(metadata),
      status: itemStatus
    };

    const rowData = headers.map(header => entryMap[header.toLowerCase()] || '');
    sheet.appendRow(rowData);
    
    return { 
      success: true, 
      id: entryMap.id,
      message: 'Entry added successfully',
      metadata: metadata
    };
  } catch (error) {
    console.error('Error adding entry:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Add a new pending entry to the sheet (Refactored for maintainability)
 */
function addPendingEntry(entryData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (!entryData.title || !entryData.type) {
      throw new Error('Missing required fields: title or type');
    }
    
    const metadata = fetchMetadata(entryData.title, entryData.type);

    const entryMap = {
      id: Utilities.getUuid(),
      createdat: new Date(),
      title: entryData.title,
      type: entryData.type,
      notes: entryData.notes || '',
      coverurl: metadata.coverURL || '',
      metadata: JSON.stringify(metadata),
      status: 'pending',
      tags: entryData.tags || '',
      hyperating: entryData.hypeRating || ''
      // Author is intentionally omitted for pending items unless specified
    };
    
    const rowData = headers.map(header => entryMap[header.toLowerCase()] || '');
    sheet.appendRow(rowData);
    
    return { 
      success: true, 
      id: entryMap.id,
      message: 'Pending entry added successfully',
      metadata: metadata
    };
  } catch (error) {
    console.error('Error adding pending entry:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Move pending entry to active (started)
 */
function startPendingEntry(entryId) {
  try {
    const updateData = {
      startDate: new Date().toISOString(),
      status: 'in-progress'
    };
    
    return updateEntry(entryId, updateData);
  } catch (error) {
    console.error('Error starting pending entry:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get all media entries from the sheet
 */
// In your Code.gs file
function getAllEntries() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return []; // No data or only headers
    }
    
    const headers = data[0].map(header => header.toString().toLowerCase());
    const entries = [];
    
    for (let i = 1; i < data.length; i++) {
      const entry = {};
      headers.forEach((header, index) => {
        let value = data[i][index];
        
        if ((header.includes('date') || header === 'createdat') && value) {
          if (value instanceof Date) {
            value = value.toISOString();
          } else {
            value = new Date(value).toISOString();
          }
        }
        
        if (header === 'rating' || header === 'hyperating') {
          if (value === 'N/A') {
            value = 'N/A'; // Preserve the "Not Rated" state as a string
          } else {
            value = parseInt(value) || null; // Parse numbers, otherwise null
          }
        }

        if (header === 'metadata' && value) {
          try {
            value = JSON.parse(value);
          } catch (e) {
            value = {};
          }
        }
        
        entry[header] = value;
      });
      
      // Set status based on data with improved logic for item types
      if (!entry.status) {
        if (!entry.startdate && !entry.finishdate) {
          // Type 1: No dates - check if has rating to determine if finished
          if (entry.rating && entry.rating !== 'N/A' && entry.rating !== '' && typeof entry.rating === 'number' && entry.rating > 0) {
            entry.status = 'completed-no-dates';
          } else {
            entry.status = 'in-progress-no-dates';
          }
        } else if (!entry.startdate && entry.finishdate) {
          // Type 2: Unknown start, known finish
          entry.status = 'completed';
        } else if (entry.startdate && !entry.finishdate) {
          // Type 3: In progress (has start, no finish)
          entry.status = 'in-progress';
        } else if (entry.startdate && entry.finishdate) {
          // Type 4: Full date info
          entry.status = 'completed';
        } else {
          entry.status = 'in-progress-no-dates'; // fallback
        }
      }
      
      entries.push(entry);
    }
    
    // Sort by creation date (newest first). This still works because
    // JavaScript can parse ISO strings back into dates for comparison.
    entries.sort((a, b) => new Date(b.createdat) - new Date(a.createdat));
    
    return entries;
  } catch (error) {
    console.error('Error in getAllEntries:', error);
    return [];
  }
}

/**
 * Mark an entry as finished with current date
 */
function markEntryAsFinished(entryId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === entryId) {
        // Set finish date to current date
        sheet.getRange(i + 1, 5).setValue(new Date());
        return { 
          success: true, 
          finishDate: new Date().toISOString().split('T')[0],
          message: 'Entry marked as finished successfully'
        };
      }
    }
    
    throw new Error('Entry not found');
  } catch (error) {
    console.error('Error marking entry as finished:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Update an existing entry
 */
function updateEntry(entryId, updateData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === entryId) {
        // Update specified fields
        Object.keys(updateData).forEach(field => {
          const headerIndex = headers.findIndex(h => 
            h.toLowerCase() === field.toLowerCase()
          );
          
          if (headerIndex !== -1) {
            let value = updateData[field];
            
            // Handle date fields
            if (field.toLowerCase().includes('date') && value) {
              value = new Date(value);
            }
            
            // Handle empty date fields
            if (field.toLowerCase().includes('date') && !value) {
              value = '';
            }

            // Handle rating - convert 0 to 'N/A' for display, but keep 0 for logic
            if (field.toLowerCase() === 'rating') {
              if (value === 0 || value === '0') {
                value = 'N/A';
              }
            }
            
            sheet.getRange(i + 1, headerIndex + 1).setValue(value);
          }
        });
        
        // Log the update for debugging
        console.log(`Updated entry ${entryId} with status: ${updateData.status}`);
        
        return { success: true, message: 'Entry updated successfully' };
      }
    }
    
    throw new Error('Entry not found');
  } catch (error) {
    console.error('Error updating entry:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Delete an entry
 */
function deleteEntry(entryId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === entryId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Entry deleted successfully' };
      }
    }
    
    throw new Error('Entry not found');
  } catch (error) {
    console.error('Error deleting entry:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get statistics for the dashboard
 */
function getStatistics() {
  try {
    const entries = getAllEntries();
    
    // Filter for only entries with a valid, numeric rating
    const ratedEntries = entries.filter(e => typeof e.rating === 'number' && e.rating > 0);
    
    const stats = {
      total: entries.length,
      inProgress: entries.filter(e => e.status === 'in-progress').length,
      completed: entries.filter(e => e.status === 'completed').length,
      byType: {},
      averageRating: 0,
      recentActivity: entries.slice(0, 5)
    };
    
    entries.forEach(entry => {
      stats.byType[entry.type] = (stats.byType[entry.type] || 0) + 1;
    });
    
    // Calculate average rating based only on rated entries
    if (ratedEntries.length > 0) {
      const ratingsSum = ratedEntries.reduce((sum, entry) => sum + entry.rating, 0);
      stats.averageRating = (ratingsSum / ratedEntries.length).toFixed(1);
    }
    
    return stats;
  } catch (error) {
    console.error('Error getting statistics:', error);
    // Return a default error state
    return { total: 0, inProgress: 0, completed: 0, byType: {}, averageRating: 0, recentActivity: [] };
  }
}

/**
 * Test function to verify setup
 */
function testSetup() {
  try {
    // Test sheet access
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const testEntry = {
      title: 'Test Entry',
      type: 'book',
      startDate: '2024-01-01',
      rating: 8,
      notes: 'This is a test entry'
    };
    
    const result = addMediaEntry(testEntry);
    
    if (result.success) {
      // Clean up test entry
      deleteEntry(result.id);
      return { success: true, message: 'Setup test completed successfully' };
    } else {
      return result;
    }
  } catch (error) {
    console.error('Setup test failed:', error);
    return { success: false, error: error.message };
  }
}