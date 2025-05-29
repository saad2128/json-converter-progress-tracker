/**
 * Combined JSON Converter & Progress Tracker
 * Features: JSON conversion + comprehensive tracking with daily/weekly stats + JSON Export
 */

// ========================================
// PROJECT CONFIGURATION FUNCTIONS
// ========================================

function setProjectStartDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Project Start Date',
    'Enter the project start date (MM/DD/YYYY or YYYY-MM-DD format):\n\nExample: 01/15/2025 or 2025-01-15',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const dateInput = response.getResponseText().trim();
  if (!dateInput) {
    ui.alert('Error', 'Please enter a valid date.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const startDate = new Date(dateInput);
    if (isNaN(startDate.getTime())) {
      throw new Error('Invalid date format');
    }
    
    // Store the start date in script properties
    PropertiesService.getScriptProperties().setProperty('PROJECT_START_DATE', startDate.toISOString());
    
    const confirmMessage = `✅ Project start date set successfully!

📅 Start Date: ${startDate.toLocaleDateString()}
🕒 Days since start: ${Math.ceil((new Date() - startDate) / (1000 * 60 * 60 * 24))}

The daily average will now be calculated as:
Total Completed Tasks ÷ Days Since Project Start`;
    
    ui.alert('Success', confirmMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Invalid date format. Please use MM/DD/YYYY or YYYY-MM-DD format.\n\nError: ${error.message}`, ui.ButtonSet.OK);
  }
}

function getProjectStartDate() {
  const startDateString = PropertiesService.getScriptProperties().getProperty('PROJECT_START_DATE');
  if (startDateString) {
    return new Date(startDateString);
  }
  return null;
}

function viewProjectSettings() {
  const startDate = getProjectStartDate();
  const ui = SpreadsheetApp.getUi();
  
  if (!startDate) {
    ui.alert(
      'Project Settings', 
      '❌ No project start date configured.\n\nPlease set a project start date first using "Set Project Start Date" option.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const today = new Date();
  const daysSinceStart = Math.ceil((today - startDate) / (1000 * 60 * 60 * 24));
  const weeksSinceStart = Math.ceil(daysSinceStart / 7);
  
  const message = `📊 PROJECT SETTINGS

📅 Start Date: ${startDate.toLocaleDateString()}
📅 Current Date: ${today.toLocaleDateString()}

⏱️ PROJECT DURATION:
• Days since start: ${daysSinceStart}
• Weeks since start: ${weeksSinceStart}

📈 CALCULATION METHOD:
• Daily Average = Total Completed Tasks ÷ ${daysSinceStart} days
• This gives the true average daily completion rate since project inception

💡 TIP: The longer the project runs, the more accurate the daily average becomes.`;
  
  ui.alert('Project Settings', message, ui.ButtonSet.OK);
}

function resetProjectSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Project Settings',
    'Are you sure you want to reset the project start date?\n\nThis will affect daily average calculations.',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty('PROJECT_START_DATE');
    ui.alert('Success', 'Project start date has been reset. Please set a new start date.', ui.ButtonSet.OK);
  }
}

function combineActiveRowToJSON() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const activeRow = activeRange.getRow();
  
  const dataRange = sheet.getDataRange();
  const numCols = dataRange.getNumColumns();
  const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  const rowData = sheet.getRange(activeRow, 1, 1, numCols).getValues()[0];
  
  const completeJsonColIndex = headers.indexOf('complete_json');
  if (completeJsonColIndex === -1) {
    SpreadsheetApp.getUi().alert('Error', 'complete_json column not found!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const jsonObject = {};
  for (let i = 2; i < headers.length; i++) {
    if (i === completeJsonColIndex) continue;
    const header = headers[i] || `Column_${i + 1}`;
    jsonObject[header] = rowData[i];
  }
  
  const jsonString = JSON.stringify(jsonObject);
  sheet.getRange(activeRow, completeJsonColIndex + 1).setValue(jsonString);
  
  // Update tracker after JSON conversion
  updateTracker();
  
  console.log(`JSON created for row ${activeRow}:`, jsonString);
  SpreadsheetApp.getUi().alert('Success', `JSON data added and tracker updated for row ${activeRow}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function combineSelectedRowsToJSON() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  
  const dataRange = sheet.getDataRange();
  const numCols = dataRange.getNumColumns();
  const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  
  const completeJsonColIndex = headers.indexOf('complete_json');
  if (completeJsonColIndex === -1) {
    SpreadsheetApp.getUi().alert('Error', 'complete_json column not found!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
    const currentRow = startRow + rowOffset;
    if (currentRow === 1) continue;
    
    const rowData = sheet.getRange(currentRow, 1, 1, numCols).getValues()[0];
    const jsonObject = {};
    for (let i = 2; i < headers.length; i++) {
      if (i === completeJsonColIndex) continue;
      const header = headers[i] || `Column_${i + 1}`;
      jsonObject[header] = rowData[i];
    }
    
    const jsonString = JSON.stringify(jsonObject);
    sheet.getRange(currentRow, completeJsonColIndex + 1).setValue(jsonString);
  }
  
  // Update tracker after JSON conversion
  updateTracker();
  
  SpreadsheetApp.getUi().alert('Success', `JSON data added and tracker updated for ${numRows} row(s)`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function combineAllRowsToJSON() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const numRows = dataRange.getNumRows();
  const numCols = dataRange.getNumColumns();
  
  if (numRows <= 1) {
    SpreadsheetApp.getUi().alert('Info', 'No data rows to process', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  const completeJsonColIndex = headers.indexOf('complete_json');
  if (completeJsonColIndex === -1) {
    SpreadsheetApp.getUi().alert('Error', 'complete_json column not found!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  for (let currentRow = 2; currentRow <= numRows; currentRow++) {
    const rowData = sheet.getRange(currentRow, 1, 1, numCols).getValues()[0];
    const jsonObject = {};
    for (let i = 2; i < headers.length; i++) {
      if (i === completeJsonColIndex) continue;
      const header = headers[i] || `Column_${i + 1}`;
      jsonObject[header] = rowData[i];
    }
    
    const jsonString = JSON.stringify(jsonObject);
    sheet.getRange(currentRow, completeJsonColIndex + 1).setValue(jsonString);
  }
  
  // Update tracker after JSON conversion
  updateTracker();
  
  SpreadsheetApp.getUi().alert('Success', `JSON data added and tracker updated for ${numRows - 1} row(s)`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// JSON EXPORT FUNCTIONS (Direct Download)
// ========================================

function exportActiveRowAsJSON() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const activeRow = activeRange.getRow();
  
  if (activeRow === 1) {
    SpreadsheetApp.getUi().alert('Error', 'Cannot export header row. Please select a data row.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const jsonData = convertRowToJSON(sheet, activeRow);
  if (!jsonData) return;
  
  const fileName = `Row_${activeRow}_Export_${new Date().toISOString().split('T')[0]}.json`;
  createDownloadableJSONFile([jsonData], fileName);
}

function exportSelectedRowsAsJSON() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  
  const jsonArray = [];
  
  for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
    const currentRow = startRow + rowOffset;
    if (currentRow === 1) continue; // Skip header row
    
    const jsonData = convertRowToJSON(sheet, currentRow);
    if (jsonData) {
      jsonArray.push(jsonData);
    }
  }
  
  if (jsonArray.length === 0) {
    SpreadsheetApp.getUi().alert('Info', 'No valid data rows found to export', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const fileName = `Rows_${startRow}-${startRow + numRows - 1}_Export_${new Date().toISOString().split('T')[0]}.json`;
  createDownloadableJSONFile(jsonArray, fileName);
}

function exportAllRowsAsJSON() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const numRows = dataRange.getNumRows();
  
  if (numRows <= 1) {
    SpreadsheetApp.getUi().alert('Info', 'No data rows to export', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const jsonArray = [];
  
  for (let currentRow = 2; currentRow <= numRows; currentRow++) {
    const jsonData = convertRowToJSON(sheet, currentRow);
    if (jsonData) {
      jsonArray.push(jsonData);
    }
  }
  
  if (jsonArray.length === 0) {
    SpreadsheetApp.getUi().alert('Info', 'No valid data found to export', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const fileName = `${SpreadsheetApp.getActiveSpreadsheet().getName()}_All_Rows_Export_${new Date().toISOString().split('T')[0]}.json`;
  createDownloadableJSONFile(jsonArray, fileName);
}

function convertRowToJSON(sheet, rowNumber) {
  const dataRange = sheet.getDataRange();
  const numCols = dataRange.getNumColumns();
  const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  const rowData = sheet.getRange(rowNumber, 1, 1, numCols).getValues()[0];
  
  const completeJsonColIndex = headers.indexOf('complete_json');
  const jsonObject = {};
  
  // Add row identifier
  jsonObject._row_number = rowNumber;
  jsonObject._exported_at = new Date().toISOString();
  
  for (let i = 2; i < headers.length; i++) {
    if (i === completeJsonColIndex) continue;
    const header = headers[i] || `Column_${i + 1}`;
    jsonObject[header] = rowData[i];
  }
  
  return jsonObject;
}

function createDownloadableJSONFile(jsonArray, fileName) {
  try {
    // Create JSON string with proper formatting
    const jsonString = JSON.stringify(jsonArray, null, 2);
    
    // Create file in Google Drive
    const blob = Utilities.newBlob(jsonString, 'application/json', fileName);
    const file = DriveApp.createFile(blob);
    
    // Set file sharing to allow download
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Get download URL
    const fileId = file.getId();
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
    const viewUrl = file.getUrl();
    
    // Create success message with download instructions
    const message = `✅ JSON file created successfully!

📁 File Name: ${fileName}
📊 Records: ${jsonArray.length}
📅 Created: ${new Date().toLocaleString()}

🔗 DOWNLOAD OPTIONS:

1. DIRECT DOWNLOAD:
   Click this link to download immediately:
   ${downloadUrl}

2. VIEW IN DRIVE:
   ${viewUrl}

💡 TIP: Right-click the direct download link and select "Save link as..." if the file opens in browser instead of downloading.`;
    
    SpreadsheetApp.getUi().alert('JSON Export Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', `Failed to create JSON file: ${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createJSONFile() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Create Downloadable JSON File',
    'Choose what to export:\n\n1 - Active Row Only\n2 - Selected Rows\n3 - All Data Rows\n\nEnter your choice (1, 2, or 3):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const choice = response.getResponseText().trim();
  
  switch (choice) {
    case '1':
      exportActiveRowAsJSON();
      break;
    case '2':
      exportSelectedRowsAsJSON();
      break;
    case '3':
      exportAllRowsAsJSON();
      break;
    default:
      ui.alert('Invalid Choice', 'Please enter 1, 2, or 3.', ui.ButtonSet.OK);
  }
}

function quickExportJSON() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Quick JSON Export',
    'Export all data rows as downloadable JSON file?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    exportAllRowsAsJSON();
  }
}

function exportWithCustomName() {
  const ui = SpreadsheetApp.getUi();
  
  // Get custom file name
  const nameResponse = ui.prompt(
    'Custom Export Name',
    'Enter a custom name for your JSON file (without .json extension):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const customName = nameResponse.getResponseText().trim();
  if (!customName) {
    ui.alert('Error', 'Please enter a valid file name.', ui.ButtonSet.OK);
    return;
  }
  
  // Get export scope
  const scopeResponse = ui.prompt(
    'Export Scope',
    'What to export:\n1 - Active Row\n2 - Selected Rows\n3 - All Rows\n\nEnter choice:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (scopeResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const choice = scopeResponse.getResponseText().trim();
  const sheet = SpreadsheetApp.getActiveSheet();
  let jsonArray = [];
  let fileName = `${customName}_${new Date().toISOString().split('T')[0]}.json`;
  
  switch (choice) {
    case '1':
      const activeRow = sheet.getActiveRange().getRow();
      if (activeRow === 1) {
        ui.alert('Error', 'Cannot export header row.', ui.ButtonSet.OK);
        return;
      }
      const jsonData = convertRowToJSON(sheet, activeRow);
      if (jsonData) jsonArray = [jsonData];
      break;
      
    case '2':
      const activeRange = sheet.getActiveRange();
      const startRow = activeRange.getRow();
      const numRows = activeRange.getNumRows();
      
      for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
        const currentRow = startRow + rowOffset;
        if (currentRow === 1) continue;
        
        const rowData = convertRowToJSON(sheet, currentRow);
        if (rowData) {
          jsonArray.push(rowData);
        }
      }
      break;
      
    case '3':
      const dataRange = sheet.getDataRange();
      const numDataRows = dataRange.getNumRows();
      
      for (let currentRow = 2; currentRow <= numDataRows; currentRow++) {
        const rowData = convertRowToJSON(sheet, currentRow);
        if (rowData) {
          jsonArray.push(rowData);
        }
      }
      break;
      
    default:
      ui.alert('Invalid Choice', 'Please enter 1, 2, or 3.', ui.ButtonSet.OK);
      return;
  }
  
  if (jsonArray.length === 0) {
    ui.alert('No Data', 'No valid data found to export.', ui.ButtonSet.OK);
    return;
  }
  
  createDownloadableJSONFile(jsonArray, fileName);
}

// ========================================
// PROGRESS TRACKER FUNCTIONS
// ========================================

function createTrackerTab() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if tracker tab already exists
  let trackerSheet = spreadsheet.getSheetByName('Tracker');
  if (trackerSheet) {
    const response = SpreadsheetApp.getUi().alert(
      'Tracker Exists', 
      'Tracker tab already exists! Do you want to recreate it?', 
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response === SpreadsheetApp.getUi().Button.YES) {
      spreadsheet.deleteSheet(trackerSheet);
    } else {
      return;
    }
  }
  
  // Create new tracker sheet
  trackerSheet = spreadsheet.insertSheet('Tracker');
  
  // Set up main tracker headers
  const mainHeaders = [
    'Worker ID',
    'Total Tasks',
    'Completed Tasks',
    'In Progress Tasks', 
    'Pending Tasks',
    'Failed Tasks',
    'Completion Rate (%)',
    'Success Rate (%)',
    'Daily Avg',
    'Weekly Total',
    'Last Updated'
  ];
  
  trackerSheet.getRange(1, 1, 1, mainHeaders.length).setValues([mainHeaders]);
  
  // Format main headers
  const headerRange = trackerSheet.getRange(1, 1, 1, mainHeaders.length);
  headerRange.setBackground('#1f4e79');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  
  // Set column widths for main section
  trackerSheet.setColumnWidth(1, 120); // Worker ID
  trackerSheet.setColumnWidth(2, 100); // Total Tasks
  trackerSheet.setColumnWidth(3, 120); // Completed
  trackerSheet.setColumnWidth(4, 120); // In Progress
  trackerSheet.setColumnWidth(5, 100); // Pending
  trackerSheet.setColumnWidth(6, 100); // Failed
  trackerSheet.setColumnWidth(7, 140); // Completion Rate
  trackerSheet.setColumnWidth(8, 120); // Success Rate
  trackerSheet.setColumnWidth(9, 100); // Daily Avg
  trackerSheet.setColumnWidth(10, 110); // Weekly Total
  trackerSheet.setColumnWidth(11, 150); // Last Updated
  
  // Add summary section
  createSummarySection(trackerSheet);
  
  // Add daily/weekly stats section
  createDailyWeeklyStatsSection(trackerSheet);
  
  // Freeze header row
  trackerSheet.setFrozenRows(1);
  
  // Initial population
  updateTracker();
  
  SpreadsheetApp.getUi().alert('Success', 'Tracker tab created with daily/weekly stats!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function createSummarySection(trackerSheet) {
  // Summary section starting from column M (13)
  const summaryHeaders = [
    ['SUMMARY STATISTICS'],
    ['Total Developers:'],
    ['Total Tasks:'],
    ['Overall Completion Rate:'],
    ['Top Performer:'],
    ['Needs Attention:'],
    ['Today\'s Completions:'],
    ['This Week\'s Completions:']
  ];
  
  trackerSheet.getRange(1, 13, summaryHeaders.length, 1).setValues(summaryHeaders);
  
  // Format summary section
  const summaryHeaderRange = trackerSheet.getRange(1, 13, 1, 2);
  summaryHeaderRange.setBackground('#d5a6bd');
  summaryHeaderRange.setFontWeight('bold');
  summaryHeaderRange.setHorizontalAlignment('center');
  
  const summaryLabelsRange = trackerSheet.getRange(2, 13, 7, 1);
  summaryLabelsRange.setFontWeight('bold');
  
  // Set column widths for summary
  trackerSheet.setColumnWidth(13, 180);
  trackerSheet.setColumnWidth(14, 150);
}

function createDailyWeeklyStatsSection(trackerSheet) {
  // Daily/Weekly stats section starting from column P (16)
  const statsHeaders = [
    ['DAILY PERFORMANCE'],
    ['Worker ID', 'Today', 'Yesterday', '2 Days Ago', '3 Days Ago', '4 Days Ago', '5 Days Ago', '6 Days Ago']
  ];
  
  trackerSheet.getRange(1, 16, 1, 1).setValues([['DAILY PERFORMANCE']]);
  trackerSheet.getRange(2, 16, 1, 8).setValues(statsHeaders[1]);
  
  // Weekly stats section starting 2 rows below daily stats
  trackerSheet.getRange(1, 25, 1, 1).setValues([['WEEKLY PERFORMANCE']]);
  trackerSheet.getRange(2, 25, 1, 5).setValues([['Worker ID', 'This Week', 'Last Week', '2 Weeks Ago', '3 Weeks Ago']]);
  
  // Format daily/weekly headers
  const dailyHeaderRange = trackerSheet.getRange(1, 16, 1, 8);
  dailyHeaderRange.setBackground('#b6d7a8');
  dailyHeaderRange.setFontWeight('bold');
  dailyHeaderRange.setHorizontalAlignment('center');
  
  const weeklyHeaderRange = trackerSheet.getRange(1, 25, 1, 5);
  weeklyHeaderRange.setBackground('#ffd966');
  weeklyHeaderRange.setFontWeight('bold');
  weeklyHeaderRange.setHorizontalAlignment('center');
  
  const dailySubHeaders = trackerSheet.getRange(2, 16, 1, 8);
  dailySubHeaders.setFontWeight('bold');
  dailySubHeaders.setBackground('#d9ead3');
  
  const weeklySubHeaders = trackerSheet.getRange(2, 25, 1, 5);
  weeklySubHeaders.setFontWeight('bold');
  weeklySubHeaders.setBackground('#fff2cc');
  
  // Set column widths for stats sections
  for (let i = 16; i <= 23; i++) {
    trackerSheet.setColumnWidth(i, 80);
  }
  for (let i = 25; i <= 29; i++) {
    trackerSheet.setColumnWidth(i, 90);
  }
}

function updateTracker() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = getMainDataSheet(spreadsheet);
  let trackerSheet = spreadsheet.getSheetByName('Tracker');
  
  // Create tracker if it doesn't exist
  if (!trackerSheet) {
    createTrackerTab();
    return;
  }
  
  // Get main sheet data
  const mainDataRange = mainSheet.getDataRange();
  const mainData = mainDataRange.getValues();
  
  if (mainData.length <= 1) {
    SpreadsheetApp.getUi().alert('Info', 'No data found to process', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const headers = mainData[0];
  const workerIdIndex = headers.indexOf('worker_id');
  const statusIndex = headers.indexOf('status');
  
  if (workerIdIndex === -1 || statusIndex === -1) {
    SpreadsheetApp.getUi().alert('Error', 'Required columns (worker_id, status) not found in main sheet!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Analyze data by worker
  const workerStats = analyzeWorkerData(mainData, workerIdIndex, statusIndex, headers);
  
  // Clear existing tracker data (except headers)
  clearTrackerData(trackerSheet);
  
  // Populate main tracker data
  populateMainTrackerData(trackerSheet, workerStats);
  
  // Update summary statistics
  updateSummaryStatistics(trackerSheet, workerStats);
  
  // Update daily/weekly stats
  updateDailyWeeklyStats(trackerSheet, mainData, workerIdIndex, statusIndex, headers);
  
  // Apply conditional formatting
  applyConditionalFormatting(trackerSheet, Object.keys(workerStats).length);
}

function getMainDataSheet(spreadsheet) {
  // Try to find the main data sheet (first sheet or one with worker_id column)
  const sheets = spreadsheet.getSheets();
  
  for (let sheet of sheets) {
    if (sheet.getName() === 'Tracker') continue;
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.includes('worker_id') && headers.includes('status')) {
      return sheet;
    }
  }
  
  // Default to first sheet if no specific sheet found
  return sheets[0];
}

function analyzeWorkerData(mainData, workerIdIndex, statusIndex, headers) {
  const workerStats = {};
  const today = new Date();
  today.setHours(23, 59, 59, 999); // End of today for comparison
  const startOfToday = new Date(today);
  startOfToday.setHours(0, 0, 0, 0); // Start of today
  const sevenDaysAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  // Try to find a timestamp column for daily/weekly calculations
  const timestampIndex = findTimestampColumn(headers);
  
  for (let i = 1; i < mainData.length; i++) {
    const workerId = mainData[i][workerIdIndex];
    const status = mainData[i][statusIndex];
    
    if (!workerId) continue; // Skip empty worker IDs
    
    if (!workerStats[workerId]) {
      workerStats[workerId] = {
        total: 0,
        completed: 0,
        inProgress: 0,
        pending: 0,
        failed: 0,
        dailyCompletions: 0,
        weeklyCompletions: 0,
        totalCompletedTasks: 0 // Track all completed tasks regardless of date
      };
    }
    
    workerStats[workerId].total++;
    
    // Get timestamp for daily/weekly calculations
    let taskDate = null;
    if (timestampIndex !== -1 && mainData[i][timestampIndex]) {
      try {
        taskDate = new Date(mainData[i][timestampIndex]);
        // Validate the date
        if (isNaN(taskDate.getTime())) {
          taskDate = null;
        }
      } catch (e) {
        taskDate = null;
      }
    }
    
    // Categorize by status (case insensitive)
    const statusLower = status ? status.toString().toLowerCase().trim() : '';
    
    if (statusLower.includes('complete') || statusLower.includes('done') || 
        statusLower.includes('finished') || statusLower.includes('success')) {
      workerStats[workerId].completed++;
      workerStats[workerId].totalCompletedTasks++;
      
      // Count daily/weekly completions if we have a valid date
      if (taskDate && !isNaN(taskDate.getTime())) {
        // Daily completions (today only)
        if (taskDate >= startOfToday && taskDate <= today) {
          workerStats[workerId].dailyCompletions++;
        }
        
        // Weekly completions (last 7 days including today)
        if (taskDate >= sevenDaysAgo && taskDate <= today) {
          workerStats[workerId].weeklyCompletions++;
        }
      } else {
        // If no timestamp available, assume recent completion for weekly count
        // This ensures some data shows even without timestamps
        workerStats[workerId].weeklyCompletions++;
      }
    } else if (statusLower.includes('progress') || statusLower.includes('working') || 
               statusLower.includes('active') || statusLower.includes('ongoing')) {
      workerStats[workerId].inProgress++;
    } else if (statusLower.includes('fail') || statusLower.includes('error') || 
               statusLower.includes('reject') || statusLower.includes('cancel')) {
      workerStats[workerId].failed++;
    } else {
      workerStats[workerId].pending++;
    }
  }
  
  return workerStats;
}

function findTimestampColumn(headers) {
  // Look for common timestamp column names
  const timestampNames = [
    'timestamp', 'date', 'created_at', 'updated_at', 'completed_at', 
    'last_updated', 'time', 'datetime', 'created', 'modified',
    'completion_date', 'task_date', 'submitted_at', 'finished_at'
  ];
  
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] ? headers[i].toString().toLowerCase().trim() : '';
    if (timestampNames.some(name => header.includes(name))) {
      return i;
    }
  }
  
  return -1; // No timestamp column found
}

function isSameDay(date1, date2) {
  if (!date1 || !date2) return false;
  return date1.getDate() === date2.getDate() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getFullYear() === date2.getFullYear();
}

// Add a debug function to test timestamp detection
function debugTimestampColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timestampIndex = findTimestampColumn(headers);
  
  let message = `Timestamp Column Analysis:\n\n`;
  message += `Headers found: ${headers.join(', ')}\n\n`;
  
  if (timestampIndex !== -1) {
    message += `✅ Timestamp column detected: "${headers[timestampIndex]}" (Column ${timestampIndex + 1})\n\n`;
    
    // Check sample data
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() > 1) {
      const sampleData = sheet.getRange(2, timestampIndex + 1, Math.min(3, dataRange.getNumRows() - 1), 1).getValues();
      message += `Sample timestamp values:\n`;
      sampleData.forEach((row, index) => {
        const value = row[0];
        const date = value ? new Date(value) : null;
        message += `Row ${index + 2}: "${value}" → ${date && !isNaN(date.getTime()) ? date.toLocaleString() : 'Invalid Date'}\n`;
      });
    }
  } else {
    message += `❌ No timestamp column found.\n\n`;
    message += `To fix this, add a column with one of these names:\n`;
    message += `timestamp, date, created_at, updated_at, completed_at, last_updated, time, datetime, created, modified, completion_date, task_date, submitted_at, finished_at`;
  }
  
  SpreadsheetApp.getUi().alert('Timestamp Debug Info', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function clearTrackerData(trackerSheet) {
  const trackerDataRange = trackerSheet.getDataRange();
  if (trackerDataRange.getNumRows() > 2) {
    // Clear main tracker data (starting from row 3)
    trackerSheet.getRange(3, 1, trackerDataRange.getNumRows() - 2, 11).clearContent();
    
    // Clear daily stats (starting from row 3, columns P-W)
    trackerSheet.getRange(3, 16, trackerDataRange.getNumRows() - 2, 8).clearContent();
    
    // Clear weekly stats (starting from row 3, columns Y-AC)
    trackerSheet.getRange(3, 25, trackerDataRange.getNumRows() - 2, 5).clearContent();
  }
}

function populateMainTrackerData(trackerSheet, workerStats) {
  const trackerData = [];
  const currentTime = new Date().toLocaleString();
  const projectStartDate = getProjectStartDate();
  
  // Calculate days since project start for daily average calculation
  let daysSinceStart = 1; // Default to 1 to avoid division by zero
  if (projectStartDate) {
    const today = new Date();
    daysSinceStart = Math.max(1, Math.ceil((today - projectStartDate) / (1000 * 60 * 60 * 24)));
  }
  
  // Sort workers by completion rate (descending)
  const sortedWorkers = Object.keys(workerStats).sort((a, b) => {
    const aRate = workerStats[a].total > 0 ? (workerStats[a].completed / workerStats[a].total) * 100 : 0;
    const bRate = workerStats[b].total > 0 ? (workerStats[b].completed / workerStats[b].total) * 100 : 0;
    return bRate - aRate;
  });
  
  sortedWorkers.forEach(workerId => {
    const stats = workerStats[workerId];
    const completionRate = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;
    const successRate = stats.total > 0 ? Math.round(((stats.completed) / (stats.total - stats.pending)) * 100) : 0;
    
    // Calculate TRUE daily average: (Completed + Pending) ÷ Days since project start
    // This shows actual daily task handling capacity
    let dailyAvg = 0;
    if (projectStartDate) {
      const productiveTasks = stats.completed + stats.pending; // Tasks that show actual work
      if (productiveTasks > 0) {
        dailyAvg = Math.round((productiveTasks / daysSinceStart) * 100) / 100; // Round to 2 decimal places
      }
    } else if (!projectStartDate && stats.weeklyCompletions > 0) {
      // Fallback to old method if no project start date is set
      dailyAvg = Math.round((stats.weeklyCompletions / 7) * 10) / 10;
    }
    
    // Weekly total should show actual weekly completions
    const weeklyTotal = stats.weeklyCompletions || 0;
    
    trackerData.push([
      workerId,
      stats.total,
      stats.completed,
      stats.inProgress,
      stats.pending,
      stats.failed,
      completionRate,
      isNaN(successRate) ? 0 : successRate,
      dailyAvg,
      weeklyTotal,
      currentTime
    ]);
  });
  
  if (trackerData.length > 0) {
    trackerSheet.getRange(3, 1, trackerData.length, 11).setValues(trackerData);
  }
  
  // Add project info to tracker sheet
  addProjectInfoToTracker(trackerSheet, projectStartDate, daysSinceStart);
}

function addProjectInfoToTracker(trackerSheet, projectStartDate, daysSinceStart) {
  // Add project information starting from column AC (29)
  const projectInfoHeaders = [
    ['PROJECT INFO'],
    ['Start Date:'],
    ['Days Running:'],
    ['Calculation Method:'],
    ['Daily Avg Formula:'],
    ['Includes:']
  ];
  
  const projectInfoValues = [
    [''], // Empty for header
    [projectStartDate ? projectStartDate.toLocaleDateString() : 'Not Set'],
    [projectStartDate ? daysSinceStart : 'N/A'],
    [projectStartDate ? '(Completed + Pending) ÷ Days Since Start' : 'Weekly Completions ÷ 7'],
    [projectStartDate ? `(Completed + Pending) ÷ ${daysSinceStart}` : 'Weekly ÷ 7'],
    [projectStartDate ? 'Completed & Pending Tasks Only' : 'Completed Tasks Only']
  ];
  
  trackerSheet.getRange(1, 29, projectInfoHeaders.length, 1).setValues(projectInfoHeaders);
  trackerSheet.getRange(1, 30, projectInfoValues.length, 1).setValues(projectInfoValues);
  
  // Format project info section
  const projectHeaderRange = trackerSheet.getRange(1, 29, 1, 2);
  projectHeaderRange.setBackground('#ff9900');
  projectHeaderRange.setFontColor('white');
  projectHeaderRange.setFontWeight('bold');
  projectHeaderRange.setHorizontalAlignment('center');
  
  const projectLabelsRange = trackerSheet.getRange(2, 29, 5, 1);
  projectLabelsRange.setFontWeight('bold');
  projectLabelsRange.setBackground('#fff2cc');
  
  // Set column widths
  trackerSheet.setColumnWidth(29, 150);
  trackerSheet.setColumnWidth(30, 200);
}

function updateSummaryStatistics(trackerSheet, workerStats) {
  const workers = Object.keys(workerStats);
  const totalDevelopers = workers.length;
  const totalTasks = workers.reduce((sum, worker) => sum + workerStats[worker].total, 0);
  const totalCompleted = workers.reduce((sum, worker) => sum + workerStats[worker].completed, 0);
  const totalDailyCompletions = workers.reduce((sum, worker) => sum + workerStats[worker].dailyCompletions, 0);
  const totalWeeklyCompletions = workers.reduce((sum, worker) => sum + workerStats[worker].weeklyCompletions, 0);
  const overallCompletionRate = totalTasks > 0 ? Math.round((totalCompleted / totalTasks) * 100) : 0;
  
  // Find top performer and needs attention
  let topPerformer = '';
  let needsAttention = '';
  let highestRate = -1;
  let lowestRate = 101;
  
  workers.forEach(worker => {
    const stats = workerStats[worker];
    const completionRate = stats.total > 0 ? (stats.completed / stats.total) * 100 : 0;
    
    if (completionRate > highestRate) {
      highestRate = completionRate;
      topPerformer = worker;
    }
    
    if (completionRate < lowestRate) {
      lowestRate = completionRate;
      needsAttention = worker;
    }
  });
  
  // Update summary values - each value needs to be wrapped in an array for setValues()
  const summaryValues = [
    [totalDevelopers],
    [totalTasks],
    [`${overallCompletionRate}%`],
    [`${topPerformer} (${Math.round(highestRate)}%)`],
    [`${needsAttention} (${Math.round(lowestRate)}%)`],
    [totalDailyCompletions],
    [totalWeeklyCompletions]
  ];
  
  trackerSheet.getRange(2, 14, 7, 1).setValues(summaryValues);
}

function updateDailyWeeklyStats(trackerSheet, mainData, workerIdIndex, statusIndex, headers) {
  const timestampIndex = findTimestampColumn(headers);
  if (timestampIndex === -1) {
    // If no timestamp column, show placeholder
    trackerSheet.getRange(3, 16, 1, 1).setValues([['No timestamp data available']]);
    trackerSheet.getRange(3, 25, 1, 1).setValues([['No timestamp data available']]);
    return;
  }
  
  const dailyStats = calculateDailyStats(mainData, workerIdIndex, statusIndex, timestampIndex);
  const weeklyStats = calculateWeeklyStats(mainData, workerIdIndex, statusIndex, timestampIndex);
  
  // Populate daily stats
  const dailyData = [];
  Object.keys(dailyStats).forEach(workerId => {
    const stats = dailyStats[workerId];
    dailyData.push([
      workerId,
      stats.today || 0,
      stats.yesterday || 0,
      stats.day2 || 0,
      stats.day3 || 0,
      stats.day4 || 0,
      stats.day5 || 0,
      stats.day6 || 0
    ]);
  });
  
  if (dailyData.length > 0) {
    trackerSheet.getRange(3, 16, dailyData.length, 8).setValues(dailyData);
  }
  
  // Populate weekly stats
  const weeklyData = [];
  Object.keys(weeklyStats).forEach(workerId => {
    const stats = weeklyStats[workerId];
    weeklyData.push([
      workerId,
      stats.thisWeek || 0,
      stats.lastWeek || 0,
      stats.week2 || 0,
      stats.week3 || 0
    ]);
  });
  
  if (weeklyData.length > 0) {
    trackerSheet.getRange(3, 25, weeklyData.length, 5).setValues(weeklyData);
  }
}

function calculateDailyStats(mainData, workerIdIndex, statusIndex, timestampIndex) {
  const dailyStats = {};
  const today = new Date();
  
  for (let i = 1; i < mainData.length; i++) {
    const workerId = mainData[i][workerIdIndex];
    const status = mainData[i][statusIndex];
    const timestamp = mainData[i][timestampIndex];
    
    if (!workerId || !timestamp) continue;
    
    const taskDate = new Date(timestamp);
    if (isNaN(taskDate.getTime())) continue;
    
    const statusLower = status ? status.toString().toLowerCase().trim() : '';
    const isCompleted = statusLower.includes('complete') || statusLower.includes('done') || 
                       statusLower.includes('finished') || statusLower.includes('success');
    
    if (!isCompleted) continue;
    
    if (!dailyStats[workerId]) {
      dailyStats[workerId] = {
        today: 0, yesterday: 0, day2: 0, day3: 0, day4: 0, day5: 0, day6: 0
      };
    }
    
    // Calculate days difference
    const daysDiff = Math.floor((today - taskDate) / (1000 * 60 * 60 * 24));
    
    if (daysDiff === 0) dailyStats[workerId].today++;
    else if (daysDiff === 1) dailyStats[workerId].yesterday++;
    else if (daysDiff === 2) dailyStats[workerId].day2++;
    else if (daysDiff === 3) dailyStats[workerId].day3++;
    else if (daysDiff === 4) dailyStats[workerId].day4++;
    else if (daysDiff === 5) dailyStats[workerId].day5++;
    else if (daysDiff === 6) dailyStats[workerId].day6++;
  }
  
  return dailyStats;
}

function calculateWeeklyStats(mainData, workerIdIndex, statusIndex, timestampIndex) {
  const weeklyStats = {};
  const today = new Date();
  
  for (let i = 1; i < mainData.length; i++) {
    const workerId = mainData[i][workerIdIndex];
    const status = mainData[i][statusIndex];
    const timestamp = mainData[i][timestampIndex];
    
    if (!workerId || !timestamp) continue;
    
    const taskDate = new Date(timestamp);
    if (isNaN(taskDate.getTime())) continue;
    
    const statusLower = status ? status.toString().toLowerCase().trim() : '';
    const isCompleted = statusLower.includes('complete') || statusLower.includes('done') || 
                       statusLower.includes('finished') || statusLower.includes('success');
    
    if (!isCompleted) continue;
    
    if (!weeklyStats[workerId]) {
      weeklyStats[workerId] = {
        thisWeek: 0, lastWeek: 0, week2: 0, week3: 0
      };
    }
    
    // Calculate weeks difference
    const weeksDiff = Math.floor((today - taskDate) / (1000 * 60 * 60 * 24 * 7));
    
    if (weeksDiff === 0) weeklyStats[workerId].thisWeek++;
    else if (weeksDiff === 1) weeklyStats[workerId].lastWeek++;
    else if (weeksDiff === 2) weeklyStats[workerId].week2++;
    else if (weeksDiff === 3) weeklyStats[workerId].week3++;
  }
  
  return weeklyStats;
}

function applyConditionalFormatting(trackerSheet, numRows) {
  if (numRows === 0) return;
  
  // Clear existing conditional formatting
  trackerSheet.clearConditionalFormatRules();
  
  const rules = [];
  
  // Completion Rate formatting (Column G)
  const completionRateRange = trackerSheet.getRange(3, 7, numRows, 1);
  
  // Green for high completion (80%+)
  const highCompletionRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(80)
    .setBackground('#d9ead3')
    .setFontColor('#274e13')
    .setRanges([completionRateRange])
    .build();
  
  // Yellow for medium completion (50-79%)
  const mediumCompletionRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(50, 79)
    .setBackground('#fff2cc')
    .setFontColor('#7f6000')
    .setRanges([completionRateRange])
    .build();
  
  // Red for low completion (<50%)
  const lowCompletionRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(50)
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([completionRateRange])
    .build();
  
  // Success Rate formatting (Column H)
  const successRateRange = trackerSheet.getRange(3, 8, numRows, 1);
  
  const highSuccessRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(90)
    .setBackground('#b6d7a8')
    .setRanges([successRateRange])
    .build();
  
  const lowSuccessRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(70)
    .setBackground('#ea9999')
    .setRanges([successRateRange])
    .build();
  
  // Daily performance formatting
  const dailyRange = trackerSheet.getRange(3, 17, numRows, 7);
  const highDailyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(5)
    .setBackground('#b6d7a8')
    .setRanges([dailyRange])
    .build();
  
  rules.push(highCompletionRule, mediumCompletionRule, lowCompletionRule, 
             highSuccessRule, lowSuccessRule, highDailyRule);
  
  trackerSheet.setConditionalFormatRules(rules);
}

function refreshTracker() {
  updateTracker();
  SpreadsheetApp.getUi().alert('Success', 'Tracker refreshed with latest daily/weekly data!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function deleteTracker() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = spreadsheet.getSheetByName('Tracker');
  
  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('Info', 'No Tracker tab found to delete.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const response = SpreadsheetApp.getUi().alert(
    'Confirm Deletion', 
    'Are you sure you want to delete the Tracker tab?', 
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response === SpreadsheetApp.getUi().Button.YES) {
    spreadsheet.deleteSheet(trackerSheet);
    SpreadsheetApp.getUi().alert('Success', 'Tracker tab deleted successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function exportTrackerData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = spreadsheet.getSheetByName('Tracker');
  
  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Tracker tab not found!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Create a new spreadsheet with tracker data
  const newSpreadsheet = SpreadsheetApp.create(`Tracker Export - ${new Date().toDateString()}`);
  const newSheet = newSpreadsheet.getActiveSheet();
  
  // Copy data
  const trackerData = trackerSheet.getDataRange().getValues();
  newSheet.getRange(1, 1, trackerData.length, trackerData[0].length).setValues(trackerData);
  
  // Copy formatting for headers
  const headerRange = newSheet.getRange(1, 1, 2, trackerData[0].length);
  headerRange.setBackground('#1f4e79');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  const url = newSpreadsheet.getUrl();
  SpreadsheetApp.getUi().alert('Export Complete', `Tracker data with daily/weekly stats exported to: ${url}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// ANALYTICS & REPORTING FUNCTIONS
// ========================================

function generatePerformanceReport() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = spreadsheet.getSheetByName('Tracker');
  
  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Tracker tab not found! Please create tracker first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Create performance report sheet
  let reportSheet = spreadsheet.getSheetByName('Performance Report');
  if (reportSheet) {
    spreadsheet.deleteSheet(reportSheet);
  }
  
  reportSheet = spreadsheet.insertSheet('Performance Report');
  
  // Generate comprehensive report
  const reportData = generateReportData();
  populatePerformanceReport(reportSheet, reportData);
  
  SpreadsheetApp.getUi().alert('Success', 'Performance report generated successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateReportData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = getMainDataSheet(spreadsheet);
  const mainData = mainSheet.getDataRange().getValues();
  
  const headers = mainData[0];
  const workerIdIndex = headers.indexOf('worker_id');
  const statusIndex = headers.indexOf('status');
  const timestampIndex = findTimestampColumn(headers);
  
  const reportData = {
    totalTasks: mainData.length - 1,
    totalWorkers: new Set(mainData.slice(1).map(row => row[workerIdIndex]).filter(id => id)).size,
    completionTrend: calculateCompletionTrend(mainData, statusIndex, timestampIndex),
    topPerformers: getTopPerformers(mainData, workerIdIndex, statusIndex),
    productivityInsights: generateProductivityInsights(mainData, workerIdIndex, statusIndex, timestampIndex)
  };
  
  return reportData;
}

function calculateCompletionTrend(mainData, statusIndex, timestampIndex) {
  if (timestampIndex === -1) return 'No timestamp data available';
  
  const last30Days = {};
  const today = new Date();
  
  // Initialize last 30 days
  for (let i = 0; i < 30; i++) {
    const date = new Date(today.getTime() - i * 24 * 60 * 60 * 1000);
    const dateKey = date.toISOString().split('T')[0];
    last30Days[dateKey] = 0;
  }
  
  // Count completions by date
  for (let i = 1; i < mainData.length; i++) {
    const status = mainData[i][statusIndex];
    const timestamp = mainData[i][timestampIndex];
    
    if (!timestamp) continue;
    
    const statusLower = status ? status.toString().toLowerCase().trim() : '';
    const isCompleted = statusLower.includes('complete') || statusLower.includes('done') || 
                       statusLower.includes('finished') || statusLower.includes('success');
    
    if (isCompleted) {
      const taskDate = new Date(timestamp);
      if (!isNaN(taskDate.getTime())) {
        const dateKey = taskDate.toISOString().split('T')[0];
        if (last30Days.hasOwnProperty(dateKey)) {
          last30Days[dateKey]++;
        }
      }
    }
  }
  
  return last30Days;
}

function getTopPerformers(mainData, workerIdIndex, statusIndex) {
  const workerStats = {};
  
  for (let i = 1; i < mainData.length; i++) {
    const workerId = mainData[i][workerIdIndex];
    const status = mainData[i][statusIndex];
    
    if (!workerId) continue;
    
    if (!workerStats[workerId]) {
      workerStats[workerId] = { total: 0, completed: 0 };
    }
    
    workerStats[workerId].total++;
    
    const statusLower = status ? status.toString().toLowerCase().trim() : '';
    if (statusLower.includes('complete') || statusLower.includes('done') || 
        statusLower.includes('finished') || statusLower.includes('success')) {
      workerStats[workerId].completed++;
    }
  }
  
  // Sort by completion rate
  return Object.keys(workerStats)
    .map(workerId => ({
      workerId,
      completionRate: workerStats[workerId].total > 0 ? 
        Math.round((workerStats[workerId].completed / workerStats[workerId].total) * 100) : 0,
      completed: workerStats[workerId].completed,
      total: workerStats[workerId].total
    }))
    .sort((a, b) => b.completionRate - a.completionRate)
    .slice(0, 5);
}

function generateProductivityInsights(mainData, workerIdIndex, statusIndex, timestampIndex) {
  const insights = [];
  
  // Calculate average tasks per worker
  const workerTaskCount = {};
  for (let i = 1; i < mainData.length; i++) {
    const workerId = mainData[i][workerIdIndex];
    if (workerId) {
      workerTaskCount[workerId] = (workerTaskCount[workerId] || 0) + 1;
    }
  }
  
  const taskCounts = Object.values(workerTaskCount);
  const avgTasksPerWorker = taskCounts.length > 0 ? 
    Math.round(taskCounts.reduce((sum, count) => sum + count, 0) / taskCounts.length) : 0;
  
  insights.push(`Average tasks per worker: ${avgTasksPerWorker}`);
  
  // Calculate completion rate trends
  const totalTasks = mainData.length - 1;
  const completedTasks = mainData.slice(1).filter(row => {
    const status = row[statusIndex];
    const statusLower = status ? status.toString().toLowerCase().trim() : '';
    return statusLower.includes('complete') || statusLower.includes('done') || 
           statusLower.includes('finished') || statusLower.includes('success');
  }).length;
  
  const overallCompletionRate = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;
  insights.push(`Overall completion rate: ${overallCompletionRate}%`);
  
  // Performance recommendations
  if (overallCompletionRate < 60) {
    insights.push('⚠️ Recommendation: Overall completion rate is below 60%. Consider reviewing task assignment and support processes.');
  } else if (overallCompletionRate > 85) {
    insights.push('✅ Excellent: High completion rate indicates strong team performance!');
  }
  
  return insights;
}

function populatePerformanceReport(reportSheet, reportData) {
  // Set up report structure - ensure all rows have exactly 5 columns
  const reportContent = [
    ['PERFORMANCE REPORT', '', '', new Date().toLocaleDateString(), ''],
    ['', '', '', '', ''],
    ['OVERVIEW', '', '', '', ''],
    ['Total Tasks:', reportData.totalTasks, '', '', ''],
    ['Total Workers:', reportData.totalWorkers, '', '', ''],
    ['', '', '', '', ''],
    ['TOP PERFORMERS', '', '', '', ''],
    ['Rank', 'Worker ID', 'Completion Rate', 'Completed Tasks', 'Total Tasks']
  ];
  
  // Add top performers - ensure exactly 5 columns
  reportData.topPerformers.forEach((performer, index) => {
    reportContent.push([
      index + 1,
      performer.workerId,
      `${performer.completionRate}%`,
      performer.completed,
      performer.total
    ]);
  });
  
  // Add spacing and insights section
  reportContent.push(['', '', '', '', '']);
  reportContent.push(['PRODUCTIVITY INSIGHTS', '', '', '', '']);
  
  // Add insights - ensure exactly 5 columns
  reportData.productivityInsights.forEach(insight => {
    reportContent.push([insight, '', '', '', '']);
  });
  
  // Populate the sheet
  if (reportContent.length > 0) {
    reportSheet.getRange(1, 1, reportContent.length, 5).setValues(reportContent);
  }
  
  // Format the report
  const titleRange = reportSheet.getRange(1, 1, 1, 4);
  titleRange.setBackground('#1f4e79');
  titleRange.setFontColor('white');
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(14);
  
  // Find and format section headers
  const overviewRowIndex = reportContent.findIndex(row => row[0] === 'OVERVIEW') + 1;
  const topPerformersRowIndex = reportContent.findIndex(row => row[0] === 'TOP PERFORMERS') + 1;
  const insightsRowIndex = reportContent.findIndex(row => row[0] === 'PRODUCTIVITY INSIGHTS') + 1;
  
  [overviewRowIndex, topPerformersRowIndex, insightsRowIndex].forEach(row => {
    if (row > 0 && row <= reportContent.length) {
      const range = reportSheet.getRange(row, 1, 1, 5);
      range.setBackground('#d5a6bd');
      range.setFontWeight('bold');
    }
  });
  
  // Auto-resize columns
  reportSheet.autoResizeColumns(1, 5);
}

// ========================================
// AUTOMATION & SCHEDULING FUNCTIONS
// ========================================

function setupDailyTrackerUpdate() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'automaticTrackerUpdate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new daily trigger at 9 AM
  ScriptApp.newTrigger('automaticTrackerUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  SpreadsheetApp.getUi().alert('Success', 'Daily automatic tracker update scheduled for 9:00 AM', SpreadsheetApp.getUi().ButtonSet.OK);
}

function automaticTrackerUpdate() {
  updateTracker();
  
  // Send email notification to stakeholders (optional)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = spreadsheet.getSheetByName('Tracker');
  
  if (trackerSheet) {
    const summaryData = trackerSheet.getRange(2, 14, 7, 1).getValues();
    const emailBody = `
Daily Tracker Update - ${new Date().toLocaleDateString()}

Summary Statistics:
- Total Developers: ${summaryData[0][0]}
- Total Tasks: ${summaryData[1][0]}
- Overall Completion Rate: ${summaryData[2][0]}
- Today's Completions: ${summaryData[5][0]}
- This Week's Completions: ${summaryData[6][0]}

View full tracker: ${spreadsheet.getUrl()}
    `;
    
    // Note: Add email addresses of stakeholders who should receive daily updates
    // MailApp.sendEmail('stakeholder@company.com', 'Daily Progress Tracker Update', emailBody);
  }
}

function removeDailyUpdate() {
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'automaticTrackerUpdate') {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    }
  });
  
  SpreadsheetApp.getUi().alert('Success', `Removed ${removedCount} automatic update trigger(s)`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ========================================
// MENU FUNCTIONS
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 JSON Converter & Tracker')
    .addSubMenu(ui.createMenu('📝 JSON Conversion (Add to Sheet)')
      .addItem('Convert Active Row to JSON', 'combineActiveRowToJSON')
      .addItem('Convert Selected Rows to JSON', 'combineSelectedRowsToJSON')
      .addItem('Convert All Rows to JSON', 'combineAllRowsToJSON'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💾 JSON Export (Download Files)')
      .addItem('🔥 Quick Export All Rows', 'quickExportJSON')
      .addSeparator()
      .addItem('Export Active Row as JSON', 'exportActiveRowAsJSON')
      .addItem('Export Selected Rows as JSON', 'exportSelectedRowsAsJSON')
      .addItem('Export All Rows as JSON', 'exportAllRowsAsJSON')
      .addSeparator()
      .addItem('📝 Create JSON File (Interactive)', 'createJSONFile')
      .addItem('🏷️ Export with Custom Name', 'exportWithCustomName'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Progress Tracker')
      .addItem('Create Tracker Tab', 'createTrackerTab')
      .addItem('Refresh Tracker', 'refreshTracker')
      .addItem('Update Tracker', 'updateTracker')
      .addSeparator()
      .addItem('Delete Tracker', 'deleteTracker'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Project Settings')
      .addItem('📅 Set Project Start Date', 'setProjectStartDate')
      .addItem('👁️ View Project Settings', 'viewProjectSettings')
      .addItem('🔄 Reset Project Settings', 'resetProjectSettings'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📈 Analytics & Reports')
      .addItem('Generate Performance Report', 'generatePerformanceReport')
      .addItem('Export Tracker Data', 'exportTrackerData'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Setup Daily Auto-Update', 'setupDailyTrackerUpdate')
      .addItem('Remove Auto-Update', 'removeDailyUpdate'))
    .addSeparator()
    .addToUi();
}
