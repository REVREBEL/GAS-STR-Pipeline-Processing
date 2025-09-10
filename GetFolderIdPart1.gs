/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  READ VALUES FROM SHEET
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/***************************************************************************************************************
 * STEP 1: Read folder IDs from named ranges and save to Script Properties.
 * Shows summary in both Logger, toast, and RevRebelGlobalLibrary.showUserMessage().
 ***************************************************************************************************************/

const NAMED_RANGE_KEYS = [
  'startDataPipeline',
  'tempRenamingComplete',
  'processedStarReports',
  'duplicateReports',
  'unresolvedReports',
];


/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  CAPTURE VALUES + SAVE TO SCRIPT PROPERTIES
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/


/** Run this to capture values from named ranges into Script Properties */
function saveFolderIdsFromNamedRanges() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  const saved = {};
  const issues = [];

  NAMED_RANGE_KEYS.forEach((key) => {
    const rng = ss.getRangeByName(key);
    if (!rng) {
      issues.push(`Missing named range: ${key}`);
      return;
    }

    const raw = String(rng.getDisplayValue() || rng.getValue() || '').trim();
    if (!raw) {
      issues.push(`Please enter the Folder ID for: ${key}`);
      return;
    }

    const id = extractDriveId_(raw);
    if (!id) {
      issues.push(`Folder ID Format is Invaild for: "${key}": ${raw}`);
      return;
    }

    if (!folderExists_(id)) {
      issues.push(`Folder not found or no access for "${key}" (ID: ${id})`);
      return;
    }

    saved[key] = id;
  });

  if (Object.keys(saved).length) {
    props.setProperties(saved, true); // overwrite
  }

  const summary =
    `Results: ${Object.keys(saved).length}/${NAMED_RANGE_KEYS.length} Folder IDs Saved to Settings.` +
    (issues.length ? `\n\nReview the following:\n\n- ${issues.join('\n- ')}` : '');

  // Show via Logger
  Logger.log(summary);

  // Show via Sheet toast (optional)
  try {
    SpreadsheetApp.getActive().toast(summary, 'Folder ID Capture', 7);
  } catch (e) {
    // Not bound to sheet; ignore
  }

  // Show via RevRebelGlobalLibrary UI
  try {
    RevRebelGlobalMessagesLibrary.showUserMessage("Load Pipeline Settings", summary);
  } catch (e) {
    Logger.log(`UI message skipped: ${e}`);
  }
}




/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  HELPER FUNCTIONS
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/


/** Helper: get a folder ID by key (later use) */
function getFolderIdProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || null;
}

/** Helper: get a Drive Folder by key (later use) */
function getFolderByKey(key) {
  const id = getFolderIdProp(key);
  if (!id) throw new Error(`No folder ID stored for key "${key}".`);
  return DriveApp.getFolderById(id);
}

/** Parse a Drive folder ID from raw ID or full URL */
function extractDriveId_(input) {
  const s = input.trim();
  if (/^[A-Za-z0-9_-]{20,}$/.test(s)) return s;
  const m =
    s.match(/\/folders\/([A-Za-z0-9_-]+)/) ||
    s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  return m ? m[1] : null;
}

/** Validate folder exists */
function folderExists_(id) {
  try {
    const f = DriveApp.getFolderById(id);
    void f.getName();
    return true;
  } catch (e) {
    return false;
  }
}

