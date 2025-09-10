

/**mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  |    RR |  GET DATA FUNCTIONS
  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm*/


/** SAVE STAR FOLDER NAMES AND IDS TO SHEET mmmmmmmmmmmmmmmmmmmmmm  */

function startWriteFoldersToSheet() {
  listStarDataSubfoldersInFolder3Levels();
  replaceStarWithSTARinColB();
  return;
}


function listStarDataSubfoldersInFolder3Levels() {
  const processedStarReportsFolderId = getFolderIdProp('processedStarReports');
  varSheetName = "starFolderL3";
  varFolderId = processedStarReportsFolderId;
  listSubfoldersInFolder3Levels(varFolderId, varSheetName);
  return;
}


/**mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  |    RR |  MAIN FUNCTIONS
  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm*/


/**
 * Lists all subfolders within a specified parent folder and logs them in a Google Sheet.
 *
 * @function listFoldersInFolder
 */

function listFoldersInFolder(varFolderId, varSheetName) {
  try {
    var parentFolder = DriveApp.getFolderById(varFolderId);
    var subfolders = parentFolder.getFolders();

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(varSheetName);

    // If the sheet doesn't exist, create it
    if (!sheet) {
      sheet = spreadsheet.insertSheet(varSheetName);
      Logger.log("Created new sheet: " + varSheetName);
    }

    // Clear previous content starting from A2
    sheet.getRange('A2:D').clearContent();

    // Write headers
    sheet.getRange("C2").setValue("FOLDER NAME");
    sheet.getRange("D2").setValue("FOLDER LINK");
    sheet.getRange("E2").setValue("FOLDER ID");

    var row = 3; // Start writing data from row 3

    // Iterate through the subfolders
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      var folderName = folder.getName();
      var folderUrl = folder.getUrl();
      var folderId = folder.getId();

      // Write the folder details to the sheet
      sheet.getRange(row, 3).setValue(folderName);
      sheet.getRange(row, 4).setValue(folderUrl);
      sheet.getRange(row, 5).setValue(folderId);

      row++;
    }

    // Auto-resize columns
    sheet.autoResizeColumns(3, 3);

    Logger.log("Successfully listed folders in " + varSheetName + " sheet.");

  } catch (error) {
    Logger.log("Error in listFoldersInFolder: " + error.message);
  }
  return;
}

/**mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  |    RR |  LIST ALL SUB FOLDERS WITHIN A FOLDER
  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm*/

/**
 * Lists all subfolders within a given Google Drive folder and logs them in a specified sheet.
 *
 * @function listSubfoldersInFolder
 * @param {string} varFolderId - The ID of the parent folder.
 * @param {string} varSheetName - The name of the sheet where data will be stored.
 */

function listSubfoldersInFolder(varFolderId, varSheetName) {
  try {
    // Validate folder ID
    if (!varFolderId) {
      throw new Error("Folder ID is missing.");
    }

    // Get the parent folder by ID
    var parentFolder = DriveApp.getFolderById(varFolderId);
    var subfolders = parentFolder.getFolders();

    // Get the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(varSheetName);

    // If the sheet doesn't exist, create it
    if (!sheet) {
      sheet = spreadsheet.insertSheet(varSheetName);
      Logger.log("Created new sheet: " + varSheetName);
    }

    // Clear existing content while keeping headers
    sheet.clear();

    // Write headers
    sheet.getRange(1, 1).setValue("Folder Name");
    sheet.getRange(1, 2).setValue("Folder URL");

    var row = 2; // Start writing data from the second row

    // Iterate through the subfolders
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      var folderName = folder.getName();
      var folderUrl = folder.getUrl();

      // Write the folder name and URL to the sheet
      sheet.getRange(row, 1).setValue(folderName);
      sheet.getRange(row, 2).setValue(folderUrl);

      row++;
    }

    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, 2);

    SpreadsheetApp.flush();
    Logger.log("Successfully listed subfolders in sheet: " + varSheetName);

  } catch (error) {
    Logger.log("Error in listSubfoldersInFolder: " + error.message);
  }
  return;
}

/**mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  |    RR |  LIST SUBFOLDERS, 3 LEVELS
  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm*/

/**
 * Lists subfolders up to 3 levels deep within a given Google Drive folder and logs them in a specified sheet.
 *
 * @function listSubfoldersInFolder3Levels
 * @param {string} varFolderId - The ID of the parent folder.
 * @param {string} varSheetName - The name of the sheet where data will be stored.
 */

function listSubfoldersInFolder3Levels(varFolderId, varSheetName) {
  try {
    // Validate folder ID
    if (!varFolderId) {
      throw new Error("Folder ID is missing.");
    }

    // Get the parent folder by ID
    var parentFolder = DriveApp.getFolderById(varFolderId);

    // Get the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(varSheetName);

    // If the sheet doesn't exist, create it
    if (!sheet) {
      sheet = spreadsheet.insertSheet(varSheetName);
      Logger.log("Created new sheet: " + varSheetName);
    } else {
      sheet.clear(); // Clear existing content while keeping headers
    }

    // Write headers
    sheet.getRange(1, 1).setValue("Level");
    sheet.getRange(1, 2).setValue("Folder Name");
    sheet.getRange(1, 3).setValue("Folder URL");

    // Initialize row counter starting from row 2
    var row = 2;

    // Call the recursive function for the top-level folder with level 1
    row = listSubfoldersRecursive(parentFolder, 1, row, sheet);

    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, 3);

    SpreadsheetApp.flush();
    Logger.log("Successfully listed subfolders up to 3 levels in sheet: " + varSheetName);

  } catch (error) {
    Logger.log("Error in listSubfoldersInFolder3Levels: " + error.message);
  }
  return;
}

/**mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  |    RR |  RECURSIVELY LIST SUBFOLDERS, 3 LEVELS
  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm*/


/**
 * Recursively lists subfolders up to 3 levels deep and writes them to the sheet.
 *
 * @function listSubfoldersRecursive
 * @param {Folder} folder - The current Google Drive folder being processed.
 * @param {number} level - The current depth level (1 to 3).
 * @param {number} row - The current row number in the sheet.
 * @param {Sheet} sheet - The Google Sheet to write the data into.
 * @returns {number} The updated row number after adding subfolder data.
 */

function listSubfoldersRecursive(folder, level, row, sheet) {
  try {
    // Get all subfolders in the current folder
    var subfolders = folder.getFolders();

    // Iterate through each subfolder
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      var folderName = subfolder.getName();
      var folderUrl = subfolder.getUrl();

      // Write the level, folder name, and URL to the sheet
      sheet.getRange(row, 1).setValue("Level " + level);
      sheet.getRange(row, 2).setValue(folderName);
      sheet.getRange(row, 3).setValue(folderUrl);

      row++;

      // If the current level is less than 3, call the function recursively
      if (level < 3) {
        row = listSubfoldersRecursive(subfolder, level + 1, row, sheet);
      }
    }

  } catch (error) {
    Logger.log("Error in listSubfoldersRecursive: " + error.message);
  }

  // Return the updated row number
  return row;
}

/**mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  
  |    RR |  CHANGE STAR TO UPPERCASE
  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm*/

/**
 * Replaces all occurrences of 'Star' with 'STAR' in Column B
 * of the sheet named 'starFolderL3'.
 *
 * @returns {void}
 */
function replaceStarWithSTARinColB() {
  try {
    // Get target sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('starFolderL3');
    if (!sheet) {
      throw new Error("Sheet 'starFolderL3' not found.");
    }

    // Get number of rows with data in Column B
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return; // No data

    // Get values from Column B (2nd column)
    var range = sheet.getRange(1, 2, lastRow, 1);
    var values = range.getValues();

    // Loop through Column B and replace 'Star' with 'STAR'
    for (var r = 0; r < values.length; r++) {
      if (typeof values[r][0] === 'string' && values[r][0].includes('Star')) {
        values[r][0] = values[r][0].replace(/Star/g, 'STAR'); // case-sensitive
      }
    }

    // Write the updated values back
    range.setValues(values);
    Logger.log("Replacement complete in Column B: 'Star' → 'STAR'.");

  } catch (err) {
    Logger.log("Error: " + err.message);
    SpreadsheetApp.getUi().alert("Error: " + err.message);
  }
  return;
}

