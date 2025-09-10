

/**
 * @fileoverview STAR report mover using `tempRenamingComplete` as the ONLY source.
 * For each file in that folder, extract (REPORTKEY, YEARKEY) from the filename,
 * look them up in the active sheet (default 'starFolders'), then move the file to
 * the resolved destination Folder ID (Col F). If a same-named file exists there,
 * move to the `duplicateReports` folder instead. Includes rich console logging
 * for the extracted keys and the computed route map.
 *
 * Sheet schema (0-based indexes):
 *  A: LEVEL (ignored)
 *  B: REPORTKEY (e.g., WeeklySTAR_BWLLIH)
 *  C: YEARKEY (e.g., 2009)
 *  D: FOLDER NAME (ignored)
 *  E: FOLDER URL (ignored)
 *  F: FILEID (Drive Folder ID to move into)
 */

/** Column indexes (0-based) in the lookup sheet. */
const COL_REPORT_KEY = 1; // B
const COL_YEAR = 2;       // C
const COL_FOLDER_ID = 5;  // F

/**
 * Move all STAR reports from the `tempRenamingComplete` source.
 * Processes both Weekly and Monthly files; key is extracted directly from the filename.
 *
 * @param {string} [sheetName='starFolders'] - Lookup sheet name in the active spreadsheet.
 * @param {boolean} [includeSubfolders=true] - If true, scan nested subfolders under the source.
 * @returns {void}
 */
function moveStarReportsFromTemp(sheetName = 'starFolders', includeSubfolders = true) {
  const sourceFolderId = getFolderIdProp('tempRenamingComplete');
  const duplicateFolderId = getFolderIdProp('duplicateReports');

  if (!sourceFolderId) throw new Error('Missing property: tempRenamingComplete folder ID.');
  if (!duplicateFolderId) throw new Error('Missing property: duplicateReports folder ID.');

  Logger.log(`SOURCE: tempRenamingComplete → ${sourceFolderId}`);
  Logger.log(`DUPLICATES: duplicateReports → ${duplicateFolderId}`);
  Logger.log(`SHEET: ${sheetName}`);

  const src = getFolderSafe_(sourceFolderId);
  const dup = getFolderSafe_(duplicateFolderId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    Logger.log('Lookup sheet has no data rows. Nothing to do.');
    return;
  }

  const routeMap = buildRouteMap_(data, /*logSample=*/true);
  Logger.log(`Route map entries: ${routeMap.size}`);

  let scanned = 0, moved = 0, parked = 0, unmatched = 0, keyErrors = 0, noRoute = 0;

  const it = listFilesInFolder_(src, includeSubfolders);
  while (it.hasNext()) {
    const file = it.next();
    const fileName = file.getName();
    scanned++;

    // Extract from filename: REPORTKEY up to the first hyphen, YEAR from -YYYYMMDD
    const reportKey = extractReportKeyFromFilename_(fileName);
    const yearKey = extractYearFromFilename_(fileName);

    Logger.log(`KEYS: file="${fileName}" | reportKey="${reportKey}" | yearKey="${yearKey}" | normKey="${reportKey ? normalizeKey_(reportKey) : '(n/a)'}"`);

    if (!reportKey || !yearKey) {
      keyErrors++;
      unmatched++;
      Logger.log(`UNMATCHED: key extraction failed → ${fileName}`);
      continue;
    }

    const mapKey = `${normalizeKey_(reportKey)}|${String(yearKey).trim()}`;
    const targetFolderId = routeMap.get(mapKey);

    if (!targetFolderId) {
      noRoute++;
      unmatched++;
      Logger.log(`UNMATCHED: no route for ${fileName} → ${mapKey}`);
      continue;
    }

    // Confirm destination folder exists
    let dest;
    try {
      dest = getFolderSafe_(targetFolderId);
    } catch (e) {
      unmatched++;
      Logger.log(`UNMATCHED: destination folder not accessible for ${fileName} → ${targetFolderId} (${e})`);
      continue;
    }

    // Duplicate check at destination
    if (fileExistsInFolder(fileName, dest)) {
      try {
        safeMoveFile_(file, dup, 'duplicate');
        parked++;
      } catch (err) {
        Logger.log(`ERROR: duplicate move failed → ${fileName} :: ${String(err)}`);
      }
      continue;
    }

    // Move to destination
    try {
      safeMoveFile_(file, dest, 'move');
      moved++;
    } catch (err) {
      // Park in duplicates to avoid blocking
      try {
        safeMoveFile_(file, dup, 'fallback');
        Logger.log(`WARN: main move failed; parked → ${fileName} :: ${String(err)}`);
        parked++;
      } catch (err2) {
        Logger.log(`ERROR: park in duplicates failed → ${fileName} :: ${String(err)} | alt: ${String(err2)}`);
      }
    }
  }

  // Summary
  Logger.log(`Scanned: ${scanned} | Moved: ${moved} | Duplicates parked: ${parked}`);
  Logger.log(`Unmatched: ${unmatched} (keyErrors=${keyErrors}, noRoute=${noRoute})`);
}

/**
 * Build a lookup map from sheet data and optionally log a small sample.
 * Key format: "<REPORTKEY>|<YEAR>" (REPORTKEY normalized: uppercased, no spaces)
 *
 * @param {any[][]} data - Sheet values (including header row).
 * @param {boolean} [logSample=false] - If true, print up to 10 sample keys.
 * @returns {Map<string,string>} Map of key → destination folderId.
 */
function buildRouteMap_(data, logSample = false) {
  const map = new Map();
  const sample = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r] || [];
    const reportKey = normalizeKey_(String(row[COL_REPORT_KEY] || ''));
    const year = String(row[COL_YEAR] || '').trim();
    const folderId = String(row[COL_FOLDER_ID] || '').trim();
    if (!reportKey || !year || !folderId) continue;
    const k = `${reportKey}|${year}`;
    map.set(k, folderId);
    if (logSample && sample.length < 10) sample.push(k);
  }
  if (logSample && sample.length) Logger.log('Route map sample →\n' + sample.map(s => `• ${s}`).join('\n'));
  return map;
}

/**
 * Extract REPORTKEY from filename as the substring before the first hyphen, e.g.:
 *  "WeeklySTAR_BWLLIH-77609-20090104-USD-E.xls" → "WeeklySTAR_BWLLIH"
 *  "MonthlySTAR_SANLAB-9402-20141201-USD-E.xls" → "MonthlySTAR_SANLAB"
 *
 * @param {string} fileName - Drive file name.
 * @returns {string|null} Report key or null if not found.
 */
function extractReportKeyFromFilename_(fileName) {
  const m = String(fileName).match(/^([^\s-]+_[^\s-]+)-/);
  return m ? m[1] : null;
}

/**
 * Extract YEARKEY (YYYY) from the first 8-digit date block following a hyphen, e.g.:
 *  "...-20090104-..." → "2009"
 *
 * @param {string} fileName - Drive file name.
 * @returns {string|null} Four-digit year or null if not found.
 */
function extractYearFromFilename_(fileName) {
  const m = String(fileName).match(/-(\d{8})/);
  return m ? m[1].slice(0, 4) : null;
}

/**
 * Normalize a key for matching: remove spaces and uppercase.
 *
 * @param {string} s - Raw key.
 * @returns {string} Normalized key.
 */
function normalizeKey_(s) {
  return String(s || '').replace(/\s+/g, '').toUpperCase().trim();
}

/**
 * Check if a file with the same name already exists in a folder.
 *
 * @param {string} fileName - Name to look up.
 * @param {GoogleAppsScript.Drive.Folder} folder - Target folder.
 * @returns {boolean} True if at least one match exists.
 */
function fileExistsInFolder(fileName, folder) {
  return folder.getFilesByName(fileName).hasNext();
}

/**
 * Move a file, with Advanced Drive API fallback for Shared Drive edge cases.
 *
 * @param {GoogleAppsScript.Drive.File} file - File to move.
 * @param {GoogleAppsScript.Drive.Folder} destFolder - Destination folder.
 * @param {string} label - Context label for logging (e.g., 'move', 'duplicate', 'fallback').
 * @returns {void}
 */
function safeMoveFile_(file, destFolder, label) {
  try {
    file.moveTo(destFolder);
    Logger.log(`${label}: moved → ${destFolder.getId()} | ${file.getName()}`);
    return;
  } catch (e) {
    Logger.log(`${label}: moveTo failed (${e}). Attempting Advanced Drive move...`);
  }
  try {
    const fileId = file.getId();
    const parents = file.getParents();
    const oldParentId = parents.hasNext() ? parents.next().getId() : '';
    const newParentId = destFolder.getId();
    if (!oldParentId) throw new Error('Cannot resolve old parent.');
    if (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.update) {
      Drive.Files.update({}, fileId, { addParents: newParentId, removeParents: oldParentId, supportsAllDrives: true });
      Logger.log(`${label}: Advanced move (parents updated) → ${newParentId} | ${file.getName()}`);
    } else {
      throw new Error('Advanced Drive Service not available. Enable it under Services.');
    }
  } catch (err) {
    throw new Error(`safeMoveFile_: Advanced move failed → ${String(err)}`);
  }
}

/**
 * Get a Drive folder by ID with friendlier errors.
 *
 * @param {string} folderId - Drive folder ID.
 * @returns {GoogleAppsScript.Drive.Folder} Folder instance.
 */
function getFolderSafe_(folderId) {
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`Folder not found or inaccessible: ${folderId} (${e})`);
  }
}

/**
 * Enumerate files in a folder; optionally include all nested subfolders.
 *
 * @param {GoogleAppsScript.Drive.Folder} root - Starting folder.
 * @param {boolean} [includeSubfolders=true] - Include nested folders.
 * @returns {{hasNext: function(): boolean, next: function(): GoogleAppsScript.Drive.File}} Iterator-like object.
 */
function listFilesInFolder_(root, includeSubfolders = true) {
  if (!includeSubfolders) return root.getFiles();
  const files = [];
  const queue = [root];
  while (queue.length) {
    const f = queue.shift();
    const fi = f.getFiles();
    while (fi.hasNext()) files.push(fi.next());
    const sub = f.getFolders();
    while (sub.hasNext()) queue.push(sub.next());
  }
  return {
    _idx: 0,
    hasNext() { return this._idx < files.length; },
    next() { return files[this._idx++]; }
  };
}

/**
 * Read a Script Property (used for folder IDs).
 *
 * @param {string} key - Property name.
 * @returns {string} Value (or empty string if not set).
 */
function getFolderIdProp(key) {
  const props = PropertiesService.getScriptProperties();
  return String(props.getProperty(key) || '');
}

/**
 * Dry run: audit routes without moving anything. Logs file → key → target.
 *
 * @param {string} [sheetName='starFolders'] - Lookup sheet name.
 * @param {boolean} [includeSubfolders=true] - Scan recursively.
 * @returns {void}
 */
function dryRun_fromTemp(sheetName = 'starFolders', includeSubfolders = true) {
  const sourceFolderId = getFolderIdProp('tempRenamingComplete');
  if (!sourceFolderId) throw new Error('Missing property: tempRenamingComplete folder ID.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const routeMap = buildRouteMap_(sheet.getDataRange().getValues());

  const src = getFolderSafe_(sourceFolderId);
  const it = listFilesInFolder_(src, includeSubfolders);

  let scanned = 0;
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    scanned++;
    const reportKey = extractReportKeyFromFilename_(name);
    const yearKey = extractYearFromFilename_(name);
    const mapKey = `${normalizeKey_(String(reportKey || ''))}|${String(yearKey || '')}`;
    const target = routeMap.get(mapKey) || '(no route)';
    Logger.log(`${name} → ${mapKey} → ${target}`);
  }
  Logger.log(`Dry run scanned: ${scanned}`);
}



/** 

function processStarFilesInFolder(varSourceFolderId, varSheetName, varAltFolderId, varKeyword) {
  const folder = DriveApp.getFolderById(varSourceFolderId);
  const files = folder.getFiles();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(varSheetName);

  const data = sheet.getDataRange().getValues(); // Retrieve all data from the sheet
  const unmatchedFilesLog = []; // Log for files without matches

  Logger.log(`Source Folder ID: ${varSourceFolderId}`);
  if (!varSourceFolderId) {
    Logger.log("Error: Folder ID is missing or undefined.");
    return;
  }

  Logger.log(`Alt Folder ID: ${varAltFolderId}`);
  if (!varAltFolderId) {
    Logger.log("Error: Alt Folder ID is missing or undefined.");
    return;
  }

  // Loop through each file in the folder
  while (files.hasNext()) {
    const file = files.next();
    const starFileName = file.getName();

    // Extract yearKey and reportKey from the filename
    const starYearKey = extractStarYearKey(starFileName);
    const starReportKey = extractStarReportKey(starFileName, varKeyword);

    Logger.log(`Processing file: ${starFileName}, Extracted YearKey: ${starYearKey}, Extracted ReportKey: ${starReportKey}`);

    if (starYearKey && starReportKey) {
      let matchFound = false;

      // Search for matching yearKey and reportKey in the spreadsheet data
      for (let i = 1; i < data.length; i++) { // Skip the header row
        const sheetReportKey = data[i][1]?.toString().replace(/\s+/g, ''); // Trim spaces from COL B
        const sheetYearKey = data[i][2]?.toString().trim(); // Trim spaces from COL C
        const targetFolderId = data[i][5]; // COL F

        Logger.log(`Checking against sheet data - ReportKey: ${sheetReportKey}, YearKey: ${sheetYearKey}`);

        if (sheetReportKey === starReportKey && sheetYearKey === starYearKey) {
          // Check if the file already exists in the target folder
          const targetFolder = DriveApp.getFolderById(targetFolderId);
          if (fileExistsInFolder(starFileName, targetFolder)) {
            Logger.log(`Duplicate found for file: ${starFileName} in folder with ID: ${targetFolderId}`);
            const altFolder = DriveApp.getFolderById(varAltFolderId);
            file.moveTo(altFolder); // Move to alternate folder
            Logger.log(`Moved file: ${starFileName} to alternate folder with ID: ${varAltFolderId}`);
          } else {
            file.moveTo(targetFolder); // Move to target folder
            Logger.log(`Moved file: ${starFileName} to folder with ID: ${targetFolderId}`);
          }
          matchFound = true;
          break;
        }
      }

      if (!matchFound) {
        Logger.log(`No match found for file: ${starFileName}`);
        unmatchedFilesLog.push(starFileName); // Log file name if no match was found
      }
    } else {
      Logger.log(`Failed to extract keys for file: ${starFileName}`);
      unmatchedFilesLog.push(starFileName); // Log file name if extraction fails
    }
  }

  // Log unmatched files
  if (unmatchedFilesLog.length > 0) {
    Logger.log('Files with no matches: ' + unmatchedFilesLog.join(', '));
  } else {
    Logger.log('All files matched and moved successfully.');
  }
}

// Function to check if a file exists in a folder
function fileExistsInFolder(fileName, folder) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext(); // Return true if at least one file with the same name exists
}



function extractStarYearKey(fileName) {
  // Match the yyyyMMdd pattern within the filename
  const match = fileName.match(/-(\d{4}\d{2}\d{2})/);
  if (match) {
    const year = match[1].substring(0, 4); // Extract only the year
    return year;
  }
  Logger.log(`Invalid date format in filename: ${fileName}`);
  return null;
}

function extractStarReportKey(fileName, varKeyword) {
  if (!fileName) {
    Logger.log("File name is missing or undefined!");
    return null;
  }

  // Construct regex to capture the correct report key format
  const regex = new RegExp(`^(${varKeyword}_[A-Za-z]+)`, 'i');
  const match = fileName.match(regex);

  Logger.log(`Regex: ${regex}`);
  Logger.log(`Match: ${match ? match[1] : "No match"}`);
  return match ? match[1] : null;
}




function testExtractStarReportKey(varKeyword) {
  const fileName = 'MonthlySTAR_AEXHER-77609-20241100-USD-E.xlsx';
  const reportKey = extractStarReportKey(fileName, varKeyword);
  Logger.log(`Extracted Report Key: ${reportKey}`);
}

*/





