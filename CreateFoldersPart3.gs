/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  CREATE THE FOLDER STRUCTURE BY PROPERTY
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/***************************************************************************************************************
 * STEP 3: Build Processed STAR Folder Structure
 * - Parent: Script Property "processedStarReports"
 * - Source scan root: Script Property "startDataPipeline"
 * - Creates (no moves yet):
 *   L1: {propertyCode} [[{full property name}]]
 *   L2: {propertyCode} {reportName}
 *   L3: {propertyCode} {yyyy} {reportName}
 ***************************************************************************************************************/

/** Toggle: preview-only (no folder creation) */

const FOLDERS_DRY_RUN = false;


/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  BUILD INDEXES FOR PROPERTY DATA
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Build fast indexes for property lookup by PROPERTYCODE and STARID.
 * Assumes loadPropertyData_() returns an object with a .rows array of row objects
 * containing PROPERTYCODE, PROPERTYNAME, STARID (string/number ok).
 *
 * @param {Object} lookup Raw lookup from loadPropertyData_().
 * @param {Object[]} lookup.rows Array of row objects from your sheet.
 * @returns {{ byCode: Map<string,Object>, byStarId: Map<string,Object> }}
 */

function indexPropertyData_(lookup) {
  var byCode = new Map();
  var byStarId = new Map();

  var rows = Array.isArray(lookup && lookup.rows) ? lookup.rows : [];
  rows.forEach(function (r) {
    var code = String(r && r.PROPERTYCODE || '').toUpperCase().trim();
    var star = String(r && r.STARID || '').trim();
    if (code) byCode.set(code, r);
    if (star) byStarId.set(star, r);
  });

  return { byCode: byCode, byStarId: byStarId };
}

/**
 * Parse a STAR filename for canonical tokens.
 * Handles forms like:
 *  - MonthlySTAR_{CODE}-{STARID}-{YYYYMMDD}-USD-E.ext
 *  - WeeklySTAR_{CODE}-{STARID}-{YYYYMMDD}-USD-E.ext
 * Falls back to looser pattern CODE-STARID-YYYYMMDD when prefix is missing.
 *
 * @param {string} name The Drive file name.
 * @returns {{ code: (string|null), starId: (string|null), yyyymmdd: (string|null), reportName: (string|null) }}
 */

function parseCanonicalStarName_(name) {
  var m = String(name || '').match(/^(?:MonthlySTAR|WeeklySTAR)_([A-Z0-9]+)-(\d+)-(\d{8})\b/i);
  if (!m) {
    // Try a looser fallback: pick the first token-like CODE-STARID-YYYYMMDD anywhere
    m = String(name || '').match(/([A-Z0-9]{3,10})-(\d+)-(\d{8})/i);
  }
  if (!m) {
    return { code: null, starId: null, yyyymmdd: null, reportName: null };
  }

  var code = m[1] ? String(m[1]).toUpperCase() : null;
  var starId = m[2] ? String(m[2]) : null;
  var yyyymmdd = m[3] || null;
  var reportName = yyyymmdd ? (yyyymmdd.endsWith('00') ? 'MonthlySTAR' : 'WeeklySTAR') : null;

  return { code: code, starId: starId, yyyymmdd: yyyymmdd, reportName: reportName };
}

/**
 * Resolve a property row by trying CODE → STARID → fuzzy Name fallback.
 * Returns both the matched row and a reason if not found.
 *
 * @param {string} original Full original file name (for fuzzy name attempts).
 * @param {{code:(string|null), starId:(string|null)}} tokens Canonical parse tokens.
 * @param {{byCode:Map<string,Object>, byStarId:Map<string,Object>}} idx Indexes from indexPropertyData_().
 * @returns {{ row: (Object|null), reason: (string|null) }}
 */

function resolvePropertyRow_(original, tokens, idx) {
  // Try by CODE
  if (tokens.code && idx.byCode.has(tokens.code)) {
    return { row: idx.byCode.get(tokens.code), reason: null };
  }

  // Try by STARID
  if (tokens.starId && idx.byStarId.has(tokens.starId)) {
    return { row: idx.byStarId.get(tokens.starId), reason: null };
  }

  // Optional: fuzzy name search (very light-weight)
  var nameHit = null;
  var nameNeedle = String(original || '')
    .replace(/[_\-]/g, ' ')
    .replace(/\.(xlsx|xls|csv)$/i, '')
    .toUpperCase();

  idx.byCode.forEach(function (r) {
    var nm = String(r && r.PROPERTYNAME || '').toUpperCase();
    if (!nameHit && nm && nameNeedle.indexOf(nm) !== -1) {
      nameHit = r;
    }
  });

  if (nameHit) return { row: nameHit, reason: null };

  var reasonParts = [];
  if (tokens.code) reasonParts.push('CODE=' + tokens.code + ' not in lookup');
  if (tokens.starId) reasonParts.push('STARID=' + tokens.starId + ' not in lookup');
  if (!tokens.code && !tokens.starId) reasonParts.push('No CODE/STARID parsed');

  return { row: null, reason: reasonParts.join('; ') || 'No match' };
}

/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  GET PARENT FOLDER + SCAN REPORTS
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Patched entry for Step 3. Creates folders with better matching & diagnostics.
 *
 * @returns {void}
 */

function buildProcessedStarFolderTree() {
  var startFolder = getFolderByKey('tempRenamingComplete');
  var processedRoot = getFolderByKey('processedStarReports');
  var dupFolderId = getFolderIdProp('duplicateReport'); // optional, only for messages

  /**
   * Validate that required Drive folders exist and return them.
   * Throws readable errors if something is missing.
   *
   * @returns {{ startFolder: GoogleAppsScript.Drive.Folder, processedRoot: GoogleAppsScript.Drive.Folder }}
   */
  function validateEnvironment_() {
    var startFolder = getFolderByKey('tempRenamingComplete');
    var processedRoot = getFolderByKey('processedStarReports');

    if (!startFolder) {
      throw new Error('Missing folder for key "tempRenamingComplete". Set the Script Property or fix getFolderByKey().');
    }
    if (!processedRoot) {
      throw new Error('Missing folder for key "processedStarReports". Set the Script Property or fix getFolderByKey().');
    }
    return { startFolder: startFolder, processedRoot: processedRoot };
  }


  /**
   * Validate that the lookup has usable keys and log quick stats.
   *
   * @param {{ rows: Object[] }} lookup Normalized lookup from loadPropertyData_().
   * @returns {void}
   */
  function assertLookupShape_(lookup) {
    var rows = Array.isArray(lookup && lookup.rows) ? lookup.rows : [];
    if (!rows.length) throw new Error('Property lookup is empty. Check sheet name, permissions, or header row.');

    var missing = 0;
    var codeSet = new Set();
    var starSet = new Set();

    rows.forEach(function (r, i) {
      var hasName = !!(r && r.PROPERTYNAME);
      var hasCode = !!(r && r.PROPERTYCODE);
      var hasStar = !!(r && r.STARID);
      if (!hasName || !hasCode) missing++;
      if (hasCode) codeSet.add(String(r.PROPERTYCODE).toUpperCase());
      if (hasStar) starSet.add(String(r.STARID));
    });

    Logger.log('[Lookup] rows=%s uniqueCodes=%s uniqueStarIds=%s missingRequired=%s', rows.length, codeSet.size, starSet.size, missing);
  }


  var lookup = loadPropertyData_();
  var idx = indexPropertyData_(lookup);

  var stats = {
    scanned: 0,
    resolved: 0,
    createdL1: 0,
    createdL2: 0,
    createdL3: 0,
    skipped: 0,
    unresolved: 0,
    errors: 0,
    notes: []
  };

  var cache = {
    l1: new Map(),
    l2: new Map(),
    l3: new Map()
  };

  walkFolder_(startFolder, 0, 3, function (file, parentChain) {
    stats.scanned++;
    try {
      var original = file.getName();

      // Parse canonical tokens from the filename
      var canon = parseCanonicalStarName_(original);

      // If date wasn’t parsed by the above, try existing helper as a fallback
      var dateInfo = canon.yyyymmdd ? { yyyymmdd: canon.yyyymmdd } : (parseDateWithContext_(original, parentChain) || {});
      var yyyymmdd = dateInfo && dateInfo.yyyymmdd || null;
      var yyyy = yyyymmdd ? yyyymmdd.slice(0, 4) : null;
      var reportName = canon.reportName || (yyyymmdd ? (yyyymmdd.endsWith('00') ? 'MonthlySTAR' : 'WeeklySTAR') : null);

      // Resolve property row via CODE → STARID → fuzzy name
      var resolved = resolvePropertyRow_(original, { code: canon.code, starId: canon.starId }, idx);
      var row = resolved.row;

      var propertyCode = row && row.PROPERTYCODE || (canon.code || null);
      var propertyName = row && row.PROPERTYNAME || null;

      if (!propertyCode || !propertyName || !reportName || !yyyy) {
        stats.unresolved++;
        var why = [];
        if (!propertyCode) why.push('no propertyCode');
        if (!propertyName) why.push('no propertyName');
        if (!reportName)  why.push('no reportName');
        if (!yyyy)        why.push('no yyyy');

        var extra = resolved && resolved.reason ? (' | ' + resolved.reason) : '';
        stats.notes.push('Unresolved for "' + original + '"' + extra + (dupFolderId ? ' (would move later)' : ''));
        return;
      }

      stats.resolved++;

      // Extract original extension (everything after the last dot)
      var ext = original.includes('.') ? original.substring(original.lastIndexOf('.')) : '';

      // Build the resolved name with extension
      var resolvedName = [
        reportName,
        propertyCode + '-' + (row && row.STARID ? row.STARID : 'NOID'),
        yyyymmdd,
        'USD-E'
      ].join('_') + ext;

      Logger.log('File: "%s" → "%s"', original, resolvedName);

      // L1: {code} [[{full name}]]
      var l1Name = sanitizeForFolder_(propertyCode + ' ' + propertyName);
      var l1Key = String(propertyCode).toUpperCase();
      var l1Key = String(propertyCode).toUpperCase();
      var l1 = cache.l1.get(l1Key);
      if (!l1) {
        l1 = ensureChildFolderByName_(processedRoot, l1Name, function () { stats.createdL1++; });
        cache.l1.set(l1Key, l1);
      }



      // L2: {code} {reportName}
      var l2Name = sanitizeForFolder_(propertyCode + ' ' + reportName);
      var l2Key = String(propertyCode).toUpperCase() + '|' + reportName;
      var l2 = cache.l2.get(l2Key);
      if (!l2) {
        l2 = ensureChildFolderByName_(l1, l2Name, function () { stats.createdL2++; });
        cache.l2.set(l2Key, l2);
      }

      // L3: {code} {yyyy} {reportName}
      var l3Name = sanitizeForFolder_(propertyCode + ' ' + yyyy + ' ' + reportName);
      var l3Key = String(propertyCode).toUpperCase() + '|' + yyyy + '|' + reportName;
      if (!cache.l3.get(l3Key)) {
        var l3 = ensureChildFolderByName_(l2, l3Name, function () { stats.createdL3++; });
        cache.l3.set(l3Key, l3);
      }
    } catch (e) {
      stats.errors++;
      stats.notes.push('ERR: ' + file.getName() + ' → ' + (e && e.message ? e.message : e));
    }
  });

  var summary = [
    'Processed Folder Builder (v2)',
    'Scanned files: ' + stats.scanned,
    'Resolved (property/date): ' + stats.resolved,
    'Created L1: ' + stats.createdL1,
    'Created L2: ' + stats.createdL2,
    'Created L3: ' + stats.createdL3,
    'Unresolved: ' + stats.unresolved,
    'Errors: ' + stats.errors,
    stats.notes.length ? ('\nNotes:\n- ' + stats.notes.slice(0, 40).join('\n- ')) : ''
  ].join('\n');

  Logger.log(summary);
  try { SpreadsheetApp.getActive().toast(summary, 'Folder Builder', 8); } catch (_) {}

  try {
    RevRebelGlobalMessagesLibrary.showUserMessage('Folder Builder', summary);
  } catch (e) {
    Logger.log('UI message skipped: ' + e);
  }
}

/**
 * Quick probe to test parsing and resolution logic on sample names.
 *
 * @returns {void}
 */
function _probeStep3_() {
  var samples = [
    'MonthlySTAR_SFOLAU-19650-20130400-USD-E.xlsx',
    'MonthlySTAR_IAHHDK-16329-20130800-USD-E.xlsx',
    'WeeklySTAR_SANLAB-9402-20141201-USD-E.xlsx'
  ];
  var idx = indexPropertyData_(loadPropertyData_());
  samples.forEach(function (s) {
    var p = parseCanonicalStarName_(s);
    var r = resolvePropertyRow_(s, { code: p.code, starId: p.starId }, idx);
    Logger.log(JSON.stringify({ s: s, parsed: p, match: !!r.row, reason: r.reason }, null, 2));
  });
}


/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  LOAD PROPERTY DATA
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Load property rows from the sheet and normalize header keys.
 * Accepts headers like: PROPERTY NAME, 'PROPERTY NAME', Property-Name, etc.
 *
 * @returns {{ rows: Object[] }}
 */
function loadPropertyData_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PropertyData');
  var data = sheet.getDataRange().getValues();
  if (!data.length) return { rows: [] };

  var headers = data.shift().map(function (h) {
    return String(h)
      .toUpperCase()
      .replace(/[\"'`]/g, '')        // strip quotes/backticks
      .replace(/[^A-Z0-9]+/g, ' ')     // squash punctuation to spaces
      .replace(/\s+/g, '')            // remove spaces -> canonical key
      .trim();
  });

  var rows = data.map(function (row) {
    var out = {};
    headers.forEach(function (h, i) {
      if (h === 'PROPERTYNAME') out.PROPERTYNAME = row[i];
      if (h === 'PROPERTYCODE') out.PROPERTYCODE = row[i];
      if (h === 'STRID' || h === 'STARID') out.STARID = row[i];
    });
    return out;
  });

  return { rows: rows };
}


/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  HELPER FUNCTIONS
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/* ----------------------- Helpers specific to Step 3 ----------------------- */


/**
 * Ensure a child folder by name with DRY-RUN support and creation counter hook.
 *
 * @param {GoogleAppsScript.Drive.Folder} parent Parent folder.
 * @param {string} name Desired child folder name.
 * @param {Function} onCreate Callback invoked when a new folder would be created.
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function ensureChildFolderByName_(parent, name, onCreate) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();

  if (FOLDERS_DRY_RUN) {
    onCreate && onCreate();
    // Create a harmless stub that only implements getName();
    return /** @type {any} */ ({ getName: function () { return name; } });
  }

  var f = parent.createFolder(name);
  onCreate && onCreate();
  return f;
}


/** Remove characters Drive won’t like in folder names */

function sanitizeForFolder_(s) {
  return String(s).replace(/[\/\\:*?"<>|]/g, '-').trim();
}
