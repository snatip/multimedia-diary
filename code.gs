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
    
    // Determine item type and set appropriate status based on new logic
    let itemStatus;
    const ratingValue = entryData.rating;
    const isFinished = entryData.isFinished; // This will come from the checkbox
    
    if (!entryData.startDate && !entryData.finishDate) {
        // Type 1: No remembered dates
        if (isFinished || (ratingValue && ratingValue !== '' && ratingValue !== 'N/A' && ratingValue !== '0' && ratingValue !== 0)) {
            // Either explicitly marked as finished OR has a rating = completed without dates
            itemStatus = 'completed-no-dates';
        } else {
            // Not marked as finished and no rating = in progress without known start date
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
      author: entryData.author || '', // CRITICAL FIX: Include author field for pending entries
      notes: entryData.notes || '',
      coverurl: metadata.coverURL || '',
      metadata: JSON.stringify(metadata),
      status: 'pending',
      tags: entryData.tags || '',
      hyperating: entryData.hypeRating || ''
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
 * Get all media entries from the sheet - DEBUG VERSION
 */
// In your Code.gs file
function getAllEntries() {
  try {
    console.log(`=== GET ALL ENTRIES DEBUG START ===`);
    
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log(`No data found - only headers or empty sheet`);
      console.log(`=== GET ALL ENTRIES DEBUG END ===`);
      return []; // No data or only headers
    }
    
    const headers = data[0].map(header => header.toString().toLowerCase());
    console.log(`Headers found:`, headers);
    console.log(`Total rows: ${data.length}`);
    
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
      // Only recalculate if status is missing or empty, preserve existing status
      console.log(`ENTRY ${i} - Status check:`);
      console.log(`  Current status: '${entry.status}'`);
      console.log(`  Has start date: ${!!entry.startdate}`);
      console.log(`  Has finish date: ${!!entry.finishdate}`);
      console.log(`  Rating: '${entry.rating}' (type: ${typeof entry.rating})`);
      
      if (!entry.status || entry.status === '') {
        console.log(`  Status is empty - calculating new status...`);
        if (!entry.startdate && !entry.finishdate) {
          // Type 1: No dates - check if has rating to determine if finished
          if (entry.rating && entry.rating !== 'N/A' && entry.rating !== '' && typeof entry.rating === 'number' && entry.rating > 0) {
            entry.status = 'completed-no-dates';
            console.log(`  → Set to 'completed-no-dates' (has rating > 0)`);
          } else {
            entry.status = 'in-progress-no-dates';
            console.log(`  → Set to 'in-progress-no-dates' (no rating or rating <= 0)`);
          }
        } else if (!entry.startdate && entry.finishdate) {
          // Type 2: Unknown start, known finish
          entry.status = 'completed';
          console.log(`  → Set to 'completed' (has finish date)`);
        } else if (entry.startdate && !entry.finishdate) {
          // Type 3: In progress (has start, no finish)
          entry.status = 'in-progress';
          console.log(`  → Set to 'in-progress' (has start date, no finish)`);
        } else if (entry.startdate && entry.finishdate) {
          // Type 4: Full date info
          entry.status = 'completed';
          console.log(`  → Set to 'completed' (has both dates)`);
        } else {
          entry.status = 'in-progress-no-dates'; // fallback
          console.log(`  → Set to 'in-progress-no-dates' (fallback)`);
        }
      } else {
        console.log(`  Status preserved: '${entry.status}'`);
      }
      // Important: If status already exists (including 'pending'), preserve it!
      
      console.log(`  Final status: '${entry.status}'`);
      console.log(`  Final rating: '${entry.rating}'`);
      
      entries.push(entry);
    }
    
    // Sort by creation date (newest first). This still works because
    // JavaScript can parse ISO strings back into dates for comparison.
    entries.sort((a, b) => new Date(b.createdat) - new Date(a.createdat));
    
    console.log(`GET ALL ENTRIES - Final entries array:`);
    entries.forEach((entry, index) => {
      console.log(`  Entry ${index}: ID=${entry.id}, Status='${entry.status}', Rating='${entry.rating}'`);
    });
    console.log(`=== GET ALL ENTRIES DEBUG END ===`);
    
    return entries;
  } catch (error) {
    console.error('Error in getAllEntries:', error);
    console.log(`=== GET ALL ENTRIES DEBUG END (ERROR) ===`);
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
 * Update an existing entry - DEBUG VERSION
 */
function updateEntry(entryId, updateData) {
  try {
    console.log(`=== UPDATE ENTRY DEBUG START ===`);
    console.log(`Entry ID: ${entryId}`);
    console.log(`Update Data Received:`, updateData);
    
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === entryId) {
        // SECURITY FIX: Log original state before any changes
        const originalEntry = {};
        headers.forEach((header, index) => {
          originalEntry[header.toLowerCase()] = data[i][index];
        });
        
        console.log(`ORIGINAL ENTRY STATE:`, originalEntry);
        console.log(`Original status: '${originalEntry.status}'`);
        console.log(`Original rating: '${originalEntry.rating}'`);
        
        console.log(`FIELDS TO UPDATE:`, Object.keys(updateData));
        
        // Update specified fields only
        Object.keys(updateData).forEach(field => {
          const headerIndex = headers.findIndex(h => 
            h.toLowerCase() === field.toLowerCase()
          );
          
          if (headerIndex !== -1) {
            let value = updateData[field];
            const originalValue = data[i][headerIndex];
            
            console.log(`UPDATING FIELD: ${field}`);
            console.log(`  Header index: ${headerIndex}`);
            console.log(`  Original value: '${originalValue}'`);
            console.log(`  New value: '${value}'`);
            
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
            
            console.log(`  Final value to be saved: '${value}'`);
            sheet.getRange(i + 1, headerIndex + 1).setValue(value);
          } else {
            console.log(`FIELD NOT FOUND: ${field} - skipping`);
          }
        });
        
        // Verify the actual state after update
        const updatedData = sheet.getDataRange().getValues();
        const updatedEntry = {};
        for (let j = 1; j < updatedData.length; j++) {
          if (updatedData[j][0] === entryId) {
            headers.forEach((header, index) => {
              updatedEntry[header.toLowerCase()] = updatedData[j][index];
            });
            break;
          }
        }
        
        console.log(`UPDATED ENTRY STATE:`, updatedEntry);
        console.log(`Updated status: '${updatedEntry.status}'`);
        console.log(`Updated rating: '${updatedEntry.rating}'`);
        
        // Check for unauthorized changes
        if (updatedEntry.status !== originalEntry.status) {
          console.error(`🚨 UPDATE ENTRY: UNAUTHORIZED STATUS CHANGE DETECTED!`);
          console.error(`Changed from: '${originalEntry.status}' to: '${updatedEntry.status}'`);
        }
        
        if (updatedEntry.rating !== originalEntry.rating) {
          console.error(`🚨 UPDATE ENTRY: UNAUTHORIZED RATING CHANGE DETECTED!`);
          console.error(`Changed from: '${originalEntry.rating}' to: '${updatedEntry.rating}'`);
        }
        
        console.log(`=== UPDATE ENTRY DEBUG END ===`);
        return { success: true, message: 'Entry updated successfully' };
      }
    }
    
    throw new Error('Entry not found');
  } catch (error) {
    console.error('Error updating entry:', error);
    console.log(`=== UPDATE ENTRY DEBUG END (ERROR) ===`);
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

/**
 * Request a new cover for an existing entry - DEBUG VERSION
 */
function requestNewCover(entryId) {
  try {
    console.log(`=== COVER UPDATE DEBUG START ===`);
    console.log(`Entry ID: ${entryId}`);
    
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === entryId) {
        // Get entry details and preserve ALL original data
        const entry = {};
        headers.forEach((header, index) => {
          entry[header.toLowerCase()] = data[i][index];
        });
        
        console.log(`ORIGINAL ENTRY STATE:`, entry);
        console.log(`Original status: '${entry.status}'`);
        console.log(`Original rating: '${entry.rating}'`);
        
        // Fetch new metadata with alternative sources
        const newMetadata = fetchAlternativeMetadata(entry.title, entry.type);
        
        // CRITICAL FIX: Only update cover-related fields, preserve everything else
        const updateData = {
          coverurl: newMetadata.coverURL,
          metadata: JSON.stringify(newMetadata)
        };
        
        console.log(`UPDATE DATA BEING SENT:`, updateData);
        console.log(`updateData.status: ${updateData.status}`);
        console.log(`updateData.rating: ${updateData.rating}`);
        
        const result = updateEntry(entryId, updateData);
        
        if (result.success) {
          // Verify the actual state after update
          const updatedData = sheet.getDataRange().getValues();
          const updatedEntry = {};
          for (let j = 1; j < updatedData.length; j++) {
            if (updatedData[j][0] === entryId) {
              headers.forEach((header, index) => {
                updatedEntry[header.toLowerCase()] = updatedData[j][index];
              });
              break;
            }
          }
          
          console.log(`UPDATED ENTRY STATE:`, updatedEntry);
          console.log(`Updated status: '${updatedEntry.status}'`);
          console.log(`Updated rating: '${updatedEntry.rating}'`);
          
          // Check for unauthorized changes
          if (updatedEntry.status !== entry.status) {
            console.error(`🚨 UNAUTHORIZED STATUS CHANGE DETECTED!`);
            console.error(`Changed from: '${entry.status}' to: '${updatedEntry.status}'`);
          }
          
          if (updatedEntry.rating !== entry.rating) {
            console.error(`🚨 UNAUTHORIZED RATING CHANGE DETECTED!`);
            console.error(`Changed from: '${entry.rating}' to: '${updatedEntry.rating}'`);
          }
          
          console.log(`=== COVER UPDATE DEBUG END ===`);
          return { 
            success: true, 
            coverURL: newMetadata.coverURL,
            message: 'New cover requested successfully'
          };
        } else {
          throw new Error(result.error);
        }
      }
    }
    
    throw new Error('Entry not found');
  } catch (error) {
    console.error('Error requesting new cover:', error);
    console.log(`=== COVER UPDATE DEBUG END (ERROR) ===`);
    return { success: false, error: error.message };
  }
}

/**
 * Fallback to placeholder image for an entry
 */
function fallbackToPlaceholder(entryId) {
  try {
    console.log(`=== FALLBACK TO PLACEHOLDER DEBUG START ===`);
    console.log(`Entry ID: ${entryId}`);
    
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === entryId) {
        // Get entry details and preserve original status
        const entry = {};
        headers.forEach((header, index) => {
          entry[header.toLowerCase()] = data[i][index];
        });
        
        console.log(`ORIGINAL ENTRY STATE:`, entry);
        console.log(`Original status: '${entry.status}'`);
        console.log(`Original rating: '${entry.rating}'`);
        
        // Generate placeholder URL
        const placeholderUrl = generateQualityPlaceholder(entry.title, entry.type);
        
        // Use updateEntry to preserve all existing data including status
        const updateData = {
          coverurl: placeholderUrl
        };
        
        console.log(`UPDATE DATA BEING SENT:`, updateData);
        console.log(`updateData.status: ${updateData.status}`);
        console.log(`updateData.rating: ${updateData.rating}`);
        
        const result = updateEntry(entryId, updateData);
        
        if (result.success) {
          // Verify the actual state after update
          const updatedData = sheet.getDataRange().getValues();
          const updatedEntry = {};
          for (let j = 1; j < updatedData.length; j++) {
            if (updatedData[j][0] === entryId) {
              headers.forEach((header, index) => {
                updatedEntry[header.toLowerCase()] = updatedData[j][index];
              });
              break;
            }
          }
          
          console.log(`UPDATED ENTRY STATE:`, updatedEntry);
          console.log(`Updated status: '${updatedEntry.status}'`);
          console.log(`Updated rating: '${updatedEntry.rating}'`);
          
          // Check for unauthorized changes
          if (updatedEntry.status !== entry.status) {
            console.error(`🚨 FALLBACK: UNAUTHORIZED STATUS CHANGE DETECTED!`);
            console.error(`Changed from: '${entry.status}' to: '${updatedEntry.status}'`);
          }
          
          if (updatedEntry.rating !== entry.rating) {
            console.error(`🚨 FALLBACK: UNAUTHORIZED RATING CHANGE DETECTED!`);
            console.error(`Changed from: '${entry.rating}' to: '${updatedEntry.rating}'`);
          }
          
          console.log(`=== FALLBACK TO PLACEHOLDER DEBUG END ===`);
          return { 
            success: true, 
            coverURL: placeholderUrl,
            message: 'Fallback to placeholder completed successfully'
          };
        } else {
          throw new Error(result.error);
        }
      }
    }
    
    throw new Error('Entry not found');
  } catch (error) {
    console.error('Error falling back to placeholder:', error);
    console.log(`=== FALLBACK TO PLACEHOLDER DEBUG END (ERROR) ===`);
    return { success: false, error: error.message };
  }
}
