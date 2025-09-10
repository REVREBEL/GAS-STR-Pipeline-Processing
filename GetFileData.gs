/**
 * @fileoverview Loop Excel STAR reports in a folder, extract key cells from
 * "Response" or "Table of Contents" tabs, then hand off to an existing rename function.
 *
 * Uses a temporary on-the-fly conversion to Google Sheets (Advanced Drive Service)
 * to avoid clutter: the converted file is permanently deleted in a finally block.
 *
 * @version Drive API v2 (Advanced Service: Drive)
 * @version Spreadsheet service (Apps Script built-in)
 * @requires Enable Advanced Google Service: Drive (Drive API v2)
 * @requires Enable Google Cloud API: Google Drive API
 *
 * Notes:
 * - This script expects a global variable `unresolvedReports` already defined in your project.
 *   It may be either a Folder object or a Drive Folder ID string.
 * - The rename handoff is `renameStarReportFromCaptured_(file, captured)`. Implement this
 *   to call your established rename function. (If you share that function’s signature, we can
 *   wire it directly.)
 */

/** --------------------------------------------------------------------------
 *  CONFIG
 *  -------------------------------------------------------------------------- */

/**
 * If true, the temporary converted Spreadsheet will be permanently deleted via Drive.Files.remove().
 * If false, it will be sent to Trash instead (DriveApp.setTrashed(true)).
 * Keeping this true avoids clutter, assuming you’re comfortable with hard deletes.
 * @type {boolean}
 */
const PERMANENTLY_DELETE_TEMP = true;

/**
 * MIME types we treat as Excel spreadsheets.
 * @type {Set<string>}
 */
const EXCEL_MIME_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
  'application/vnd.ms-excel',                                         // .xls
  'application/vnd.ms-excel.sheet.macroEnabled.12'                    // .xlsm
]);

/** --------------------------------------------------------------------------
 *  ENTRY POINT
 *  -------------------------------------------------------------------------- */

/**
 * Entry point: Processes all Excel reports in the `unresolvedReports` folder.
 *
 * Behavior:
 *  - For each Excel file:
 *    - Convert temporarily → Google Sheet (Advanced Drive Service).
 *    - Read values from "Response" (preferred) or "Table of Contents" tabs.
 *    - If neither tab is present → throw an error with context.
 *    - Handoff to `renameStarReportFromCaptured_()` with the captured values.
 *
 * Logging:
 *  - Logs key steps and errors per file; continues to next file on error.
 */
function processUnresolvedReportsFolder() {
  const folder = resolveFolderFromVar_(typeof unresolvedReports !== 'undefined' ? unresolvedReports : null, 'unresolvedReports');
  if (!folder) {
    throw new Error('Folder `unresolvedReports` not found or not provided.');
  }

  const files = folder.getFiles();
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  console.log('RR | START | Processing Excel files in folder: %s (%s)', folder.getName(), folder.getId());

  while (files.hasNext()) {
    const file = files.next();
    try {
      if (!EXCEL_MIME_TYPES.has(file.getMimeType())) {
        skippedCount++;
        console.log('RR | SKIP | Not an Excel file: "%s" (%s)', file.getName(), file.getMimeType());
        continue;
      }

      const captured = extractFromExcelFile_(file);
      console.log('RR | OK   | Captured from "%s": %s', file.getName(), JSON.stringify(captured));

      // Handoff: Implement this function in your project to call your existing renamer.
      if (typeof renameStarReportFromCaptured_ === 'function') {
        renameStarReportFromCaptured_(file, captured);
      } else {
        console.warn('RR | WARN | renameStarReportFromCaptured_ is not defined. Skipping rename for "%s".', file.getName());
      }

      processedCount++;
    } catch (err) {
      errorCount++;
      console.error('RR | ERROR | "%s": %s', file.getName(), err && err.message ? err.message : err);
    }
  }

  console.log('RR | DONE | Processed: %s | Skipped: %s | Errors: %s', processedCount, skippedCount, errorCount);
}

/** --------------------------------------------------------------------------
 *  CORE LOGIC
 *  -------------------------------------------------------------------------- */

/**
 * Convert an Excel file to a temporary Google Sheet, extract the required cells,
 * then permanently delete (or trash) the temp file with a finally cleanup.
 *
 * @param {GoogleAppsScript.Drive.File} file - The Excel Drive file.
 * @returns {{
 *   source: { fileId: string, fileName: string },
 *   tabUsed: 'Response'|'Table of Contents',
 *   response?: { c2?: string, c3?: string, c4?: string },
 *   toc?: {
 *     layout: 'newer'|'older',
 *     b10?: string, b12?: string, e12?: string, m12?: string,
 *     b13?: string, h13?: string, l13?: string
 *   }
 * }}
 * @throws {Error} If neither "Response" nor "Table of Contents" tabs exist or if reading fails.
 */
function extractFromExcelFile_(file) {
  const fileId = file.getId();
  const fileName = file.getName();

  let tempSheetId = null;
  try {
    tempSheetId = convertExcelToTempSheet_(file);
    const ss = SpreadsheetApp.openById(tempSheetId);

    // Try "Response" first
    const responseSheet = findSheetByNames_(ss, ['Response']);
    if (responseSheet) {
      const c2 = getDisplayValueSafe_(responseSheet, 'C2');
      const c3 = getDisplayValueSafe_(responseSheet, 'C3');
      const c4 = getDisplayValueSafe_(responseSheet, 'C4');

      return {
        source: { fileId, fileName },
        tabUsed: 'Response',
        response: { c2, c3, c4 }
      };
    }

    // Fallback: "Table of Contents"
    const tocSheet = findSheetByNames_(ss, ['Table of Contents']);
    if (tocSheet) {
      // Newer layout first
      const b10n = getDisplayValueSafe_(tocSheet, 'B10');
      const b12n = getDisplayValueSafe_(tocSheet, 'B12');
      const e12n = getDisplayValueSafe_(tocSheet, 'E12');
      const m12n = getDisplayValueSafe_(tocSheet, 'M12');

      if (hasAnyValue_([b10n, b12n, e12n, m12n])) {
        return {
          source: { fileId, fileName },
          tabUsed: 'Table of Contents',
          toc: {
            layout: 'newer',
            b10: b10n,  // Hotel Name
            b12: b12n,  // Date Range Start - End
            e12: e12n,  // STR ID
            m12: m12n   // Date Created
          }
        };
      }

      // Older layout fallback
      const b10o = getDisplayValueSafe_(tocSheet, 'B10'); // Hotel Name
      const b13o = getDisplayValueSafe_(tocSheet, 'B13'); // Date Range or Month
      const h13o = getDisplayValueSafe_(tocSheet, 'H13'); // STR ID
      const l13o = getDisplayValueSafe_(tocSheet, 'L13'); // Date Created

      if (hasAnyValue_([b10o, b13o, h13o, l13o])) {
        return {
          source: { fileId, fileName },
          tabUsed: 'Table of Contents',
          toc: {
            layout: 'older',
            b10: b10o,
            b13: b13o,
            h13: h13o,
            l13: l13o
          }
        };
      }

      throw new Error('Found "Table of Contents" sheet but expected cells were empty.');
    }

    // Neither sheet exists → hard fail per your requirement
    throw new Error('Neither "Response" nor "Table of Contents" sheets were found.');
  } finally {
    // Cleanup temp conversion (no clutter)
    if (tempSheetId) {
      try {
        if (PERMANENTLY_DELETE_TEMP) {
          Drive.Files.remove(tempSheetId); // Hard delete via Advanced Drive
        } else {
          DriveApp.getFileById(tempSheetId).setTrashed(true);
        }
      } catch (cleanupErr) {
        console.warn('RR | WARN | Failed to cleanup temp sheet %s: %s', tempSheetId, cleanupErr && cleanupErr.message ? cleanupErr.message : cleanupErr);
      }
    }
  }
}

/** --------------------------------------------------------------------------
 *  HELPERS: Drive/Spreadsheet
 *  -------------------------------------------------------------------------- */

/**
 * Resolve a folder from an input that may be a Folder object or a Folder ID string.
 *
 * @param {GoogleAppsScript.Drive.Folder|string|null} input - Folder object or folder ID.
 * @param {string} varName - For error messages (e.g., 'unresolvedReports').
 * @returns {GoogleAppsScript.Drive.Folder|null} Folder or null if not found.
 */
function resolveFolderFromVar_(input, varName) {
  try {
    if (!input) return null;
    if (typeof input === 'string') {
      return DriveApp.getFolderById(input);
    }
    if (input && typeof input.getId === 'function') {
      return /** @type {GoogleAppsScript.Drive.Folder} */ (input);
    }
    console.warn('RR | WARN | %s was provided but not a string ID or Folder object.', varName);
    return null;
  } catch (err) {
    console.error('RR | ERROR | Could not resolve folder for %s: %s', varName, err && err.message ? err.message : err);
    return null;
  }
}

/**
 * Convert an Excel file to a temporary Google Sheet using the Advanced Drive Service.
 *
 * @param {GoogleAppsScript.Drive.File} excelFile - The Excel file to convert.
 * @returns {string} The new Spreadsheet file ID.
 * @throws {Error} If conversion fails.
 */
function convertExcelToTempSheet_(excelFile) {
  try {
    const blob = excelFile.getBlob();
    const inserted = Drive.Files.insert(
      { mimeType: MimeType.GOOGLE_SHEETS, title: 'TEMP_IMPORT_' + new Date().toISOString() },
      blob
    );
    if (!inserted || !inserted.id) {
      throw new Error('Drive.Files.insert returned no ID.');
    }
    return inserted.id;
  } catch (err) {
    throw new Error('Failed to convert Excel to temp Sheet: ' + (err && err.message ? err.message : err));
  }
}

/**
 * Find the first sheet that matches any of the given names (case-insensitive).
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The Spreadsheet.
 * @param {string[]} names - Candidate sheet names.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The matching sheet or null.
 */
function findSheetByNames_(ss, names) {
  const all = ss.getSheets();
  const set = new Set(names.map(n => String(n).trim().toLowerCase()));
  for (let i = 0; i < all.length; i++) {
    const s = all[i];
    const nm = String(s.getName() || '').trim().toLowerCase();
    if (set.has(nm)) return s;
  }
  return null;
}

/**
 * Get a cell’s display value safely (trimmed). Returns '' for missing/blank.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet.
 * @param {string} a1 - A1 notation (e.g., "C2").
 * @returns {string} Trimmed display value or ''.
 */
function getDisplayValueSafe_(sheet, a1) {
  try {
    return String(sheet.getRange(a1).getDisplayValue() || '').trim();
  } catch (_e) {
    return '';
  }
}

/**
 * Return true if any non-empty string exists in the list.
 *
 * @param {string[]} values - Candidate values.
 * @returns {boolean} True if any trimmed value is non-empty.
 */
function hasAnyValue_(values) {
  return values.some(v => String(v || '').trim().length > 0);
}

/** --------------------------------------------------------------------------
 *  RENAME HANDOFF (hook up to your existing script)
 *  -------------------------------------------------------------------------- */

/**
 * Handoff to your established rename workflow.
 * Implement this to call your existing rename function(s).
 *
 * @param {GoogleAppsScript.Drive.File} file - The original Excel file to rename.
 * @param {Object} captured - The captured data blob from extractFromExcelFile_().
 * @returns {void}
 *
 * @example
 * // Example wire-up:
 * // function renameStarReportFromCaptured_(file, captured) {
 * //   const newName = buildNameFromCaptured_(captured); // your existing logic
 * //   file.setName(newName);
 * //   // Optionally set appProperties or Labels here.
 * // }
 */
function renameStarReportFromCaptured_(file, captured) {
  // Placeholder: integrate with your existing renamer.
  // If you share the signature of your rename function, I can call it directly here.
  console.log('RR | INFO | (Demo) Would rename "%s" using captured data from tab: %s', file.getName(), captured.tabUsed);
}
