/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  RECURSIVE RENAMER LOGIC (STAR Reports)
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/***************************************************************************************************************
 * STEP 2: Recursive Renamer for Reports
 * - Start folder: Script Property "startDataPipeline"
 * - Complete (resolved): Script Property "tempRenamingComplete" → resolved files are moved here
 * - Fallback (unresolved): Script Property "unresolvedReports" → unresolved files are moved here
 * - Lookup: Sheet "PropertyData" (columns: PROPERTYNAME, STARID, PROPERTYCODE, GEOID, CITY, STATE)
 * - Depth: current + 3 levels
 * - Final base: {Monthly|Weekly}STAR_{GEOID}{PROPERTYCODE}-{STARID}-{yyyyMMdd}-USD-E
 * - Extensions: original file extension is preserved (e.g., .xlsx, .xls, .csv)
 * - Shared drives: moves use the Advanced Drive Service (v2) with supportsAllDrives=true
 ***************************************************************************************************************/

/** Toggle to preview without changing Drive */

const DRY_RUN = false;

/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  ENTRY
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Entry point: rename everything starting at the folder stored in Script Property "startDataPipeline".
 * Moves unresolved files into Script Property "unresolvedReports" and resolved files into "tempRenamingComplete".
 *
 * @returns {void}
 */

function renameAllStarReports() {
  const startFolder = getFolderByKey('startDataPipeline');
  const unresolvedFolderId = getFolderIdProp('unresolvedReports');
  const unresolvedFolder = unresolvedFolderId ? DriveApp.getFolderById(unresolvedFolderId) : null;
  const completeFolderId = getFolderIdProp('tempRenamingComplete');
  const completeFolder = completeFolderId ? DriveApp.getFolderById(completeFolderId) : null;

  let lookup;
  try {
    lookup = loadPropertyDataV2_(); // maps and rows (namespaced to avoid collisions)
    assertLookup_(lookup);
    Logger.log('[PropertyData] rows=%s', lookup && lookup.rows);
  } catch (e) {
    const fatal = `FATAL: PropertyData load failed — ${e && e.message ? e.message : e}`;
    Logger.log(fatal);
    try { SpreadsheetApp.getActive().toast(fatal, 'Report Renamer', 8); } catch (_) { /* ignore */ }
    return; // abort run; avoid spamming per-file errors
  }
  /** @type {{ scanned:number, renamed:number, skipped:number, unresolved:number, moved:number, completedMoved:number, errors:number, messages:string[] }} */
  const stats = { scanned: 0, renamed: 0, skipped: 0, unresolved: 0, moved: 0, completedMoved: 0, errors: 0, messages: [] };

  // Walk 3 levels deep
  walkFolder_(startFolder, 0, 3, (file, parentChain) => {
    stats.scanned++;
    try {
      const res = resolveAndRename_(file, parentChain, lookup, unresolvedFolder, completeFolder);
      if (res.action === 'RENAMED') stats.renamed++;
      else if (res.action === 'SKIPPED') stats.skipped++;
      else if (res.action === 'MOVED') { stats.unresolved++; stats.moved++; }
      else if (res.action === 'UNRESOLVED') stats.unresolved++;
      if (res.movedToComplete) stats.completedMoved++;
      if (res.message) stats.messages.push(res.message);
    } catch (e) {
      stats.errors++;
      stats.messages.push(`ERR ${file.getName()}: ${e && e.message ? e.message : e}`);
    }
  });

  const summary = [
    `Report Renamer`,
    `Scanned: ${stats.scanned}`,
    `Renamed: ${stats.renamed}`,
    `Moved to Complete: ${stats.completedMoved}`,
    `Skipped: ${stats.skipped}`,
    `Unresolved: ${stats.unresolved}`,
    `Moved to Unresolved Folder: ${stats.moved}`,
    `Errors: ${stats.errors}`,
    (stats.messages.length ? `\nNotes:\n- ${stats.messages.slice(0, 15).join('\n- ')}` : '')
  ].join('\n');

  Logger.log(summary);
  try { SpreadsheetApp.getActive().toast(summary, 'Report Renamer', 8); } catch (_) { /* ignore */ }
  try { RevRebelGlobalMessagesLibrary.showUserMessage('Report Renamer', summary); } catch (e) { Logger.log(`UI message skipped: ${e}`); }
}

/**
 * Resolve core values for a given Drive file and either rename/move it or mark/move as unresolved.
 *
 * @param {GoogleAppsScript.Drive.File} file Drive file to evaluate.
 * @param {GoogleAppsScript.Drive.Folder[]} parentChain Array of nearest parents (direct parent first).
 * @param {{ byCode: Map<string, any>, byName: Map<string, any> }} lookup Lookup object from loadPropertyData_().
 * @param {GoogleAppsScript.Drive.Folder|null} unresolvedFolder Destination for unresolved files (optional).
 * @param {GoogleAppsScript.Drive.Folder|null} completeFolder Destination for successfully resolved files (optional).
 * @returns {{ action: 'RENAMED'|'SKIPPED'|'MOVED'|'UNRESOLVED', movedToComplete?: boolean, message?: string }} Action result.
 */

function resolveAndRename_(file, parentChain, lookup, unresolvedFolder, completeFolder) {
  const original = file.getName();
  const ext = getFileExtension_(original); // keep original extension (e.g., .xlsx)
  const parsed = parseNameTokens_(original);

  // NEW: detect special report type (Pulse / Bandwidth / RPM)
  const special = detectSpecialReport_(original);
  
  // Ensure lookup maps exist to avoid undefined.get errors
  assertLookup_(lookup);

  // If already standard, still move to Complete
  const alreadyStandard = isAlreadyStandard_(original);
  if (alreadyStandard && !needsNormalization_(original)) {
    if (DRY_RUN) {
      return {
        action: 'SKIPPED',
        movedToComplete: !!completeFolder,
        message: `DRY_RUN: already standard → move to _TempRenamingComplete: "${original}"`
      };
    }
    if (completeFolder) {
      try {
        moveFileToFolder_(file, completeFolder);
        Logger.log('FILE: already standard → moved to Complete: "%s"', original);
        return { action: 'SKIPPED', movedToComplete: true, message: `Already standard → moved to _TempRenamingComplete: "${original}"` };
      } catch (e) {
        return { action: 'SKIPPED', message: `Already standard; move failed: ${e && e.message ? e.message : e}` };
      }
    }
    return { action: 'SKIPPED', message: `Already standard: ${original}` };
  }

  // Try to identify property via PROPERTYCODE first, else PROPERTYNAME
  const propMatch =
    findByPropertyCode_(parsed.codesInName, lookup) ||
    findByPropertyName_(original, lookup);

  // Try to harvest date from filename; if partial (MMYY), use parent folders to resolve year
  const dateInfo = parseDateWithContext_(original, parentChain);

  // Build final parts
  const geoId = propMatch?.row?.GEOID || null;
  const propCode = propMatch?.row?.PROPERTYCODE || propMatch?.code || null;
  const starId = propMatch?.row?.STARID || null;
  const yyyymmdd = dateInfo?.yyyymmdd || null;

  // Decide Monthly vs Weekly
  // If special report, we do NOT use Monthly/Weekly; otherwise compute Monthly/Weekly
  const versionLabel = (!special && yyyymmdd)
    ? (yyyymmdd.endsWith('00') ? 'Monthly' : 'Weekly')
    : null;


 // Essentials:
  // - Always need geoId, propCode, starId, yyyymmdd
  // - Need versionLabel ONLY for non-special reports
  const needsVersion = !special;
  if (!geoId || !propCode || !starId || !yyyymmdd || (needsVersion && !versionLabel)) {
    const why = `Unresolved -> ${[
      !geoId ? 'GEOID' : null,
      !propCode ? 'PROPERTYCODE' : null,
      !starId ? 'STARID' : null,
      !yyyymmdd ? 'DATE' : null,
      needsVersion && !versionLabel ? 'VERSION' : null
    ].filter(Boolean).join(', ')}`;

    if (unresolvedFolder) {
      if (DRY_RUN) return { action: 'MOVED', message: `${why}; would move "${original}" to Unresolved Reports.` };
      try {
        moveFileToFolder_(file, unresolvedFolder);
        Logger.log('FILE: unresolved (%s) → moved to Unresolved: "%s"', why, original);
        return { action: 'MOVED', message: `${why}; moved "${original}" to Unresolved Reports.` };
      } catch (e) {
        return { action: 'UNRESOLVED', message: `${why}; move failed: ${e && e.message ? e.message : e}` };
      }
    }
    return { action: 'UNRESOLVED', message: `${why}; no unresolved reports folder configured for "${original}".` };
  }

  // Build final name
  const prefix = special ? special.prefix : `${versionLabel}STAR_`;
  const finalBase = `${prefix}${geoId}${propCode}-${starId}-${yyyymmdd}-USD-E`;
  const finalName = `${finalBase}${ext}`;

  if (special) {
    Logger.log(`TYPE: ${special.key} → using prefix "${special.prefix}" for "${original}"`);
  } else {
    Logger.log(`TYPE: ${versionLabel}STAR → using prefix "${versionLabel}STAR_" for "${original}"`);
  }


  const nameChanged = original !== finalName;

  if (DRY_RUN) {
    const moveNote = completeFolder ? ' + move → _TempRenamingComplete' : '';
    return { action: 'SKIPPED', message: `DRY_RUN: "${original}" → "${finalName}"${moveNote}` };
  }

  // Apply rename if needed, then move to complete
  try {
    if (nameChanged) {
      file.setName(finalName);
      Logger.log('FILE: "%s" → "%s"', original, finalName);
    } else {
      Logger.log('FILE: no rename needed (computed name equals original): "%s"', original);
    }
  } catch (e) {
    return { action: 'SKIPPED', message: `Rename failed for "${original}": ${e && e.message ? e.message : e}` };
  }

  let movedToComplete = false;
  if (completeFolder) {
    try {
      moveFileToFolder_(file, completeFolder);
      Logger.log('FILE: moved to Complete folder: "%s"', finalName);
      movedToComplete = true;
    } catch (e) {
      return { action: nameChanged ? 'RENAMED' : 'SKIPPED', message: `"${original}" → "${finalName}"; move failed: ${e && e.message ? e.message : e}` };
    }
  }

  return { action: nameChanged ? 'RENAMED' : 'SKIPPED', movedToComplete, message: `"${original}" → "${finalName}"${movedToComplete ? '; moved to _TempRenamingComplete' : ''}` };
}


/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  LOOKUP + PARSING
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Build a header index map using logical keys + optional aliases.
 * Matches are case-insensitive and insensitive to spaces/punctuation.
 *
 * @param {string[]} headerRow The header row values (e.g., getValues()[0]).
 * @param {string[]} requiredKeys Logical keys you’ll use in code, e.g., ['SORT','PROPERTYNAME',...].
 * @param {Object<string, string[]>} [aliasMap] Map of logical key → array of acceptable header labels as seen in the sheet.
 * @returns {Object<string, number>} Map of logical key → 0-based column index.
 * @throws {Error} If any required key can’t be matched, throws with a helpful message listing seen headers.
 */

function idxMap_(headerRow, requiredKeys, aliasMap) {
  const norm = (s) => String(s || '')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9 ]+/g, '')
    .replace(/\s+/g, '');

  /** @type {Record<string, number>} */
  const byNorm = {};
  headerRow.forEach((h, i) => {
    const k = norm(h);
    if (!(k in byNorm)) byNorm[k] = i;
  });

  /** @type {Record<string, number>} */
  const out = {};

  /** @type {string[]} */
  const missing = [];

  requiredKeys.forEach((key) => {
    const candidates = [key, ...(aliasMap?.[key] || [])].map(norm);
    const found = candidates.find((c) => c in byNorm);
    if (found) out[key] = byNorm[found]; else missing.push(key);
  });

  if (missing.length) {
    const seen = headerRow.map((h, i) => `${i + 1}:${h}`).join(', ');
    throw new Error(`Column(s) not found: ${missing.join(', ')}. Seen headers → [${seen}]`);
  }
  return out;
}

/**
 * Load PropertyData into structures for fast matching.
 * Expects a sheet named "PropertyData". Supports optional SEARCH KEYWORDS column.
 *
 * Columns resolved (required): SORT, PROPERTYNAME, STARID, PROPERTYCODE, GEOID, CITY, STATE
 * Optional (if present): SEARCHKEYWORDS (aliases/synonyms; accepts comma/semicolon/pipe separated terms)
 *
 * @returns {{ byCode: Map<string, any>, byName: Map<string, any>, rows: number }}
 *   byCode: Map of PROPERTYCODE (uppercased) → row object
 *   byName: Map of normalized name/alias → row object
 *   rows:   Count of data rows processed
 * @throws {Error} If the sheet is missing or required columns cannot be resolved.
 */

function loadPropertyData_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('PropertyData');
  if (!sh) throw new Error('Missing sheet "PropertyData".');

  // Read header and data rows
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h).trim());
  const rows = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), header.length).getValues();

  // Try to include SEARCHKEYWORDS if present; otherwise fall back without it
  let H;
  try {
    H = idxMap_(
      header,
      ['SORT', 'PROPERTYNAME', 'STARID', 'PROPERTYCODE', 'GEOID', 'CITY', 'STATE', 'SEARCHKEYWORDS'],
      {
        PROPERTYNAME: ['PROPERTY NAME'],
        STARID: ['STAR ID', 'STR ID', 'STRID', 'STR'],
        PROPERTYCODE: ['PROPERTY CODE'],
        GEOID: ['GEO ID'],
        SEARCHKEYWORDS: ['SEARCH KEYWORDS', 'KEYWORDS']
      }
    );
  } catch (_) {
    H = idxMap_(
      header,
      ['SORT', 'PROPERTYNAME', 'STARID', 'PROPERTYCODE', 'GEOID', 'CITY', 'STATE'],
      {
        PROPERTYNAME: ['PROPERTY NAME'],
        STARID: ['STAR ID', 'STR ID', 'STRID', 'STR'],
        PROPERTYCODE: ['PROPERTY CODE'],
        GEOID: ['GEO ID']
      }
    );
    H.SEARCHKEYWORDS = -1; // optional, not found
  }

  /** @type {Map<string, any>} */
  const byCode = new Map(); // PROPERTYCODE → rowObj

  /** @type {Map<string, any>} */
  const byName = new Map(); // normalized PROPERTYNAME/alias → rowObj

  // Pre-compute first-word frequencies to avoid adding generic/ambiguous aliases (e.g., "resort")
  const firstWordCounts = (function () {
    const counts = Object.create(null);
    rows.forEach((r) => {
      const name = String(r[H.PROPERTYNAME] || '').trim();
      const fw = (name || '').replace(/[^A-Za-z]+/g, ' ').trim().split(' ').filter(Boolean)[0] || '';
      const k = normalizeName_(fw);
      if (k) counts[k] = (counts[k] || 0) + 1;
    });
    return counts;
  })();

  rows.forEach((r) => {
    const rowObj = {
      PROPERTYNAME: String(r[H.PROPERTYNAME] || '').trim(),
      STARID: String(r[H.STARID] || '').trim(),
      PROPERTYCODE: String(r[H.PROPERTYCODE] || '').trim(),
      GEOID: String(r[H.GEOID] || '').trim(),
      CITY: String(r[H.CITY] || '').trim(),
      STATE: String(r[H.STATE] || '').trim(),
      SEARCHKEYWORDS: (typeof H.SEARCHKEYWORDS === 'number' && H.SEARCHKEYWORDS >= 0 ? String(r[H.SEARCHKEYWORDS] || '').trim() : '')
    };

    // Index by PROPERTYCODE
    if (rowObj.PROPERTYCODE) byCode.set(rowObj.PROPERTYCODE.toUpperCase(), rowObj);

    // Index by full PROPERTYNAME
    if (rowObj.PROPERTYNAME) byName.set(normalizeName_(rowObj.PROPERTYNAME), rowObj);

    // Index by SEARCH KEYWORDS (aliases/synonyms)
    if (H.SEARCHKEYWORDS >= 0) {
      const raw = String(r[H.SEARCHKEYWORDS] || '');
      raw.split(/[;,|]/).map((s) => s.trim()).filter(Boolean).forEach((kw) => {
        const k = normalizeName_(kw);
        if (k) byName.set(k, rowObj);
      });
    }

    // Lightweight implicit alias: only add the **first word** if it is UNIQUE across the sheet
    // and not a generic hotel word. Prevents collisions like "Resort at Squaw Creek" winning for any “... Resort ...”.
    const genericFirst = new Set(['resort', 'hotel', 'the', 'inn', 'spa']);
    const firstWord = (rowObj.PROPERTYNAME || '').replace(/[^A-Za-z]+/g, ' ').trim().split(' ').filter(Boolean)[0] || '';
    const fwNorm = normalizeName_(firstWord);
    if (fwNorm && fwNorm.length >= 5 && firstWordCounts[fwNorm] === 1 && !genericFirst.has(fwNorm)) {
      byName.set(fwNorm, rowObj);
    }
  });

  return { byCode, byName, rows: rows.length };
}

/**
 * V2 loader with the same behavior but a unique name to avoid collisions with
 * any older loadPropertyData_ definitions that may exist in other project files.
 *
 * @returns {{ byCode: Map<string, any>, byName: Map<string, any>, rows: number }}
 */

function loadPropertyDataV2_() {
  // Self-contained loader to avoid collisions with any other loadPropertyData_ definitions
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('PropertyData');
  if (!sh) throw new Error('Missing sheet "PropertyData".');

  // Read header and data rows
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map((h) => String(h).trim());
  Logger.log('[PropertyData] header cols=%s', header.length);
  const rows = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), header.length).getValues();

  // Resolve columns; SEARCHKEYWORDS is optional
  let H;
  try {
    H = idxMap_(
      header,
      ['SORT', 'PROPERTYNAME', 'STARID', 'PROPERTYCODE', 'GEOID', 'CITY', 'STATE', 'SEARCHKEYWORDS'],
      {
        PROPERTYNAME: ['PROPERTY NAME'],
        STARID: ['STAR ID', 'STR ID', 'STRID', 'STR'],
        PROPERTYCODE: ['PROPERTY CODE'],
        GEOID: ['GEO ID'],
        SEARCHKEYWORDS: ['SEARCH KEYWORDS', 'KEYWORDS']
      }
    );
  } catch (_) {
    H = idxMap_(
      header,
      ['SORT', 'PROPERTYNAME', 'STARID', 'PROPERTYCODE', 'GEOID', 'CITY', 'STATE'],
      {
        PROPERTYNAME: ['PROPERTY NAME'],
        STARID: ['STAR ID', 'STR ID', 'STRID', 'STR'],
        PROPERTYCODE: ['PROPERTY CODE'],
        GEOID: ['GEO ID']
      }
    );
    H.SEARCHKEYWORDS = -1; // optional, not found
  }

  /** @type {Map<string, any>} */
  const byCode = new Map();

  /** @type {Map<string, any>} */
  const byName = new Map();

  // Pre-compute first-word frequencies to avoid adding generic/ambiguous aliases (e.g., "resort")
  const firstWordCounts = (function () {
    const counts = Object.create(null);
    rows.forEach((r) => {
      const name = String(r[H.PROPERTYNAME] || '').trim();
      const fw = (name || '').replace(/[^A-Za-z]+/g, ' ').trim().split(' ').filter(Boolean)[0] || '';
      const k = normalizeName_(fw);
      if (k) counts[k] = (counts[k] || 0) + 1;
    });
    return counts;
  })();

  rows.forEach((r) => {
    const rowObj = {
      PROPERTYNAME: String(r[H.PROPERTYNAME] || '').trim(),
      STARID: String(r[H.STARID] || '').trim(),
      PROPERTYCODE: String(r[H.PROPERTYCODE] || '').trim(),
      GEOID: String(r[H.GEOID] || '').trim(),
      CITY: String(r[H.CITY] || '').trim(),
      STATE: String(r[H.STATE] || '').trim(),
      SEARCHKEYWORDS: (typeof H.SEARCHKEYWORDS === 'number' && H.SEARCHKEYWORDS >= 0 ? String(r[H.SEARCHKEYWORDS] || '').trim() : '')
    };

    if (rowObj.PROPERTYCODE) byCode.set(rowObj.PROPERTYCODE.toUpperCase(), rowObj);
    if (rowObj.PROPERTYNAME) byName.set(normalizeName_(rowObj.PROPERTYNAME), rowObj);

    if (H.SEARCHKEYWORDS >= 0) {
      const raw = String(r[H.SEARCHKEYWORDS] || '');
      raw.split(/[;,|]/).map((s) => s.trim()).filter(Boolean).forEach((kw) => {
        const k = normalizeName_(kw);
        if (k) byName.set(k, rowObj);
      });
    }

    const fwRaw = (rowObj.PROPERTYNAME || '').replace(/[^A-Za-z]+/g, ' ').trim();
    const parts = fwRaw ? fwRaw.split(' ') : [];
    const firstWord = parts.length ? parts[0] : '';
    const fwNorm = normalizeName_(firstWord);
    if (fwNorm && fwNorm.length >= 5 && firstWordCounts[fwNorm] === 1 && !new Set(['resort', 'hotel', 'the', 'inn', 'spa']).has(fwNorm)) {
      byName.set(fwNorm, rowObj);
    }
  });

  Logger.log('[PropertyData] indexed byCode=%s, byName=%s, rows=%s', byCode.size, byName.size, rows.length);
  return { byCode, byName, rows: rows.length };
}

/**
 * Ensure lookup maps exist; throw a clear error if not.
 *
 * @param {{byName: any, byCode: any}} lookup Lookup container.
 * @returns {void}
 */

function assertLookup_(lookup) {
  var ok = lookup && (lookup.byName instanceof Map) && (lookup.byCode instanceof Map);
  if (!ok) {
    var byNameType = lookup && lookup.byName ? Object.prototype.toString.call(lookup.byName) : 'undefined';
    var byCodeType = lookup && lookup.byCode ? Object.prototype.toString.call(lookup.byCode) : 'undefined';
    throw new Error('Lookup not initialized (byName=' + byNameType + ', byCode=' + byCodeType + '). Check PropertyData headers and loadPropertyData_().');
  }
}

/**
 * Verify Advanced Drive is available; otherwise instruct how to enable it.
 *
 * @returns {void}
 */
function requireAdvancedDrive_() {
  var ok = (typeof Drive === 'object') && Drive && Drive.Files &&
    (typeof Drive.Files.get === 'function') && (typeof Drive.Files.update === 'function');
  if (!ok) {
    throw new Error('Advanced Drive API not available. Turn on Services → Advanced Google services → Drive API, and enable Drive API in Google Cloud console.');
  }
}

/**
 * Extract alpha tokens from a filename for partial property-name matching.
 *
 * - Returns distinct tokens (letters only) ≥ 3 characters.
 * - Adds **only one pass** of bigram fusions from the original core tokens
 *   (e.g., "wild"+"dunes" → "wilddunes").
 * - Filters common noise words: star/str/usd/xls/xlsx/csv/pdf/weekly/monthly/report/final/draft/copy.
 * - Hard-caps the output size to prevent pathological growth.
 *
 * @param {string} name Filename to tokenize.
 * @returns {string[]} Distinct, normalized alpha tokens and single-pass bigram fusions.
 */

function extractAlphaTokens_(name) {
  // Strip extension to avoid it joining into a token.
  const base = String(name || '').replace(/\.[^.]*$/, '');

  // Pull alphabetic words and normalize with your existing normalizer.
  const words = (base.match(/[A-Za-z]+/g) || [])
    .map((w) => normalizeName_(w))
    .filter(Boolean);

  // Common noise tokens to ignore.
  const stop = new Set([
    'star', 'str', 'usd', 'xls', 'xlsx', 'csv', 'pdf',
    'weekly', 'monthly', 'report', 'final', 'draft', 'copy',
    'hotel', 'resort', 'lodge', 'inn', 'spa', 'villas', 'villa', 'suites', 'suite', 'beach', 'mountain', 'waterfront', 'estates', 'condos', 'golf', 'house', 'park', 'place', 'plaza', 'center', 'centre', 'village', 'hall'
  ]);

  // Core tokens (letters only, length ≥ 3, not in stoplist).
  const core = words.filter((w) => w.length >= 3 && !stop.has(w));

  // IMPORTANT: compute bigrams from a **snapshot** of core tokens (do not mutate while iterating).
  const bigrams = [];
  for (let i = 0; i < core.length - 1; i++) {
    const fused = core[i] + core[i + 1];
    if (fused.length >= 5) bigrams.push(fused);
  }

  // Merge & de-duplicate and apply a sanity cap.
  const out = [];
  const seen = new Set();
  for (const t of core.concat(bigrams)) {
    if (!seen.has(t)) {
      seen.add(t);
      out.push(t);
      if (out.length >= 64) break; // cap to keep matching predictable & safe
    }
  }
  return out;
}

/**
 * Helper: split letters-only tokens (normalized) and filter generic words.
 *
 * @param {string} text Input text.
 * @param {Set<string>} [extraStop] Optional extra stop words.
 * @returns {string[]} Tokens ≥ 3 chars, lowercased, diacritic-stripped.
 */

function tokenizeWords_(text, extraStop) {
  const base = String(text || '');
  const words = (base.match(/[A-Za-z]+/g) || []).map((w) => normalizeName_(w)).filter(Boolean);
  const stop = new Set([
    'the', 'and', 'of', 'at', 'by', 'for', 'to', 'in', 'on', 'a', 'an',
    'star', 'str', 'usd', 'xls', 'xlsx', 'csv', 'pdf', 'weekly', 'monthly', 'report', 'final', 'draft', 'copy',
    // lodging generics ↓ (very low discriminative power)
    'hotel', 'resort', 'lodge', 'inn', 'spa', 'villas', 'villa', 'suites', 'suite', 'beach', 'mountain', 'waterfront', 'estates', 'condos', 'golf', 'house', 'park', 'place', 'plaza', 'center', 'centre', 'village', 'hall', 'house'
  ]);
  if (extraStop) { extraStop.forEach((w) => stop.add(w)); }
  return words.filter((w) => w.length >= 3 && !stop.has(w));
}

/**
 * Helper: return unique row objects from lookup.
 *
 * @param {{byCode: Map<string, any>, byName: Map<string, any>}} lookup Lookup maps.
 * @returns {any[]} Unique row list.
 */

function uniqueRows_(lookup) {
  const uniq = new Set();
  const rows = [];
  if (lookup && lookup.byCode instanceof Map) {
    for (const r of lookup.byCode.values()) if (r && !uniq.has(r)) { uniq.add(r); rows.push(r); }
  }
  if (lookup && lookup.byName instanceof Map) {
    for (const r of lookup.byName.values()) if (r && !uniq.has(r)) { uniq.add(r); rows.push(r); }
  }
  return rows;
}

/**
 * Build derived tokens for a property row (name tokens, bigrams, city/state tokens, aliases).
 *
 * @param {any} row Row object from PropertyData.
 * @returns {{ normName:string, nameTokens:string[], nameBigrams:string[], cityTok:string|null, stateTok:string|null, aliasList:string[] }}
 */

function propDerivedTokens_(row) {
  const normName = normalizeName_(row.PROPERTYNAME || '');
  const nameTokens = tokenizeWords_(row.PROPERTYNAME || '');
  const nameBigrams = [];
  for (let i = 0; i < nameTokens.length - 1; i++) {
    const fused = nameTokens[i] + nameTokens[i + 1];
    if (fused.length >= 5) nameBigrams.push(fused);
  }
  const cityTok = row.CITY ? (tokenizeWords_(row.CITY)[0] || null) : null;
  const stateTok = row.STATE ? (tokenizeWords_(row.STATE)[0] || null) : null;
  const aliasList = (row.SEARCHKEYWORDS || '')
    .split(/[;,|]/)
    .map((s) => s.trim())
    .filter(Boolean)
    .map((s) => normalizeName_(s));
  return { normName, nameTokens, nameBigrams, cityTok, stateTok, aliasList };
}

/**
 * Score how well a property row matches a filename.
 * Higher is better. Keep weights simple and interpretable.
 *
 * @param {string} fileText Cleaned, normalized filename.
 * @param {Set<string>} fileTokens Tokens (and bigrams) from the filename.
 * @param {any} row PropertyData row.
 * @returns {number} Score.
 */

function computeMatchScore_(fileText, fileTokens, row) {
  const d = propDerivedTokens_(row);
  let score = 0;

  // Strong signals
  if (d.normName && fileText.includes(d.normName)) score += 200; // full name appears
  for (const a of d.aliasList) { if (a && fileText.includes(a)) { score += 140; break; } }

  // Name bigrams present (e.g., wilddunes)
  for (const bg of d.nameBigrams) { if (fileText.includes(bg)) score += 40; }

  // Token overlap (cap contribution)
  let overlap = 0;
  for (const t of d.nameTokens) { if (fileTokens.has(t)) overlap++; }
  if (overlap >= 1) score += 20 * Math.min(overlap, 3); // up to +60
  if (overlap >= 2) score += 10; // small bonus for multi-token hit

  // City / state evidence
  if (d.cityTok && fileTokens.has(d.cityTok)) score += 60;
  if (d.stateTok && fileTokens.has(d.stateTok)) score += 25;

  // Brand+City fused (first significant token + city) if present
  if (d.cityTok && d.nameTokens.length) {
    const fused = d.nameTokens[0] + d.cityTok;
    if (fileText.includes(fused)) score += 80;
  }
  return score;
}

/**
 * Try to match by PROPERTYNAME using the simpler, earlier heuristic (kept because it worked well for you):
 *  1) Exact normalized match (full name or any alias in SEARCH KEYWORDS)
 *  2) Filename contains the full normalized property name → prefer the **longest**
 *  3) Otherwise, use filename tokens/bigrams; pick the row whose normalized name contains the **longest token** (≥4)
 *     – ties are treated as ambiguous → return null (so we don't mislabel)
 *
 * @param {string} filename The file name to search within.
 * @param {{ byName: Map<string, any> }} lookup Lookup with normalized name/alias keys.
 * @returns {{ row: any } | null}
 */

function findByPropertyName_(filename, lookup) {
  if (!lookup || !(lookup.byName instanceof Map)) return null;

  const clean = normalizeName_(filename);
  const tokens = extractAlphaTokens_(filename); // safe, single-pass bigrams

  // 1) Exact match on any key we indexed (full name or alias)
  const exact = lookup.byName.get(clean);
  if (exact) return { row: exact };

  // 2) Filename contains the full property name → choose the **longest** match
  let bestByContain = null; // { normName, row }
  for (const [normName, row] of lookup.byName.entries()) {
    if (clean.includes(normName)) {
      if (!bestByContain || normName.length > bestByContain.normName.length) {
        bestByContain = { normName, row };
      }
    }
  }
  if (bestByContain) return { row: bestByContain.row };

  // 3) Use filename tokens/bigrams (≥4 chars). Pick row with **longest** token contained in its name
  let bestToken = null; // { tokenLen, row }
  for (const [normName, row] of lookup.byName.entries()) {
    let localMax = 0;
    for (const t of tokens) {
      if (t.length >= 4 && normName.includes(t)) {
        if (t.length > localMax) localMax = t.length;
      }
    }
    if (localMax > 0) {
      if (!bestToken || localMax > bestToken.tokenLen) {
        bestToken = { tokenLen: localMax, row };
      } else if (bestToken && localMax === bestToken.tokenLen) {
        // tie → ambiguous
        bestToken = { tokenLen: localMax, row: null };
      }
    }
  }

  // Final sanity check for token-based guess: filename must contain the full normalized
  // property name OR one of its explicit aliases; otherwise treat as unresolved.
  if (bestToken && bestToken.row) {
    const rn = normalizeName_(bestToken.row.PROPERTYNAME || '');
    const aliases = String(bestToken.row.SEARCHKEYWORDS || '')
      .split(/[;,|]/)
      .map((s) => normalizeName_(String(s).trim()))
      .filter(Boolean);
    const ok = (rn && clean.includes(rn)) || aliases.some((a) => a && clean.includes(a));
    if (!ok) return null;
    return { row: bestToken.row };
  }

  return null;
}


/**
 * Try to find by PROPERTYCODE present in a set of name tokens.
 *
 * @param {string[]} codesInName Candidate property codes parsed from a file name/string.
 * @param {{ byCode: Map<string, any> }} lookup Lookup object returned from loadPropertyData_().
 * @returns {{ code: string, row: any } | null}
 */
function findByPropertyCode_(codesInName, lookup) {
  // Guard against missing lookup maps
  if (!lookup || !(lookup.byCode instanceof Map)) { return null; }
  for (const code of codesInName) {
    const row = lookup.byCode.get(String(code || '').toUpperCase());
    if (row) return { code, row };
  }
  return null;
}


/**
 * Normalize a name for robust matching (diacritic-safe, spacing/punctuation tolerant).
 *
 * @param {any} s Input string.
 * @returns {string}
 */

function normalizeName_(s) {
  return String(s || '')
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\u00A0/g, ' ')
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[‘’‚‛′ʹ`]/g, "'")
    .replace(/[^a-z0-9]+/g, '')
    .trim();
}

/**
 * Extract simple tokens from a name and collect candidate property codes (2–6 letters).
 *
 * @param {string} name The filename or text to tokenize.
 * @returns {{ tokens: string[], codesInName: string[] }}
 */

function parseNameTokens_(name) {
  const tokens = String(name).split(/[^A-Za-z0-9]+/).filter(Boolean);
  const codesInName = tokens.filter((t) => /^[A-Za-z]{2,6}$/.test(t));
  return { tokens, codesInName };
}

/**
 * Detect whether a name already conforms to the standard format (allows optional extension at the end).
 *
 * @param {string} name A filename to test.
 * @returns {boolean}
 */
function isAlreadyStandard_(name) {
  // Matches:
  // - MonthlySTAR_...  or WeeklySTAR_...
  // - PulseSTAR_...    or BandwidthSTAR_... or RPM_...
  return /((?:Monthly|Weekly)STAR_|(?:PulseSTAR_|BandwidthSTAR_|RPM_))[A-Z]{3,6}[A-Z]{2,6}-\d{4,6}-\d{8}-USD-E(\.[A-Za-z0-9]+)?$/i
    .test(name);
}


/**
 * Placeholder: decide whether a file that "matches shape" still needs normalization.
 * Currently no-op.
 *
 * @param {string} _name The filename.
 * @returns {boolean}
 */

function needsNormalization_(_name) { // eslint-disable-line no-unused-vars
  return false;
}

/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  DATE HANDLING
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Parse a date from a filename; when only MMYY is present, infer the year from parent folders.
 * Returns the synthetic monthly day as 00 for monthly files.
 *
 * @param {string} name The filename.
 * @param {GoogleAppsScript.Drive.Folder[]} parentChain Array of parent folders (nearest first).
 * @returns {{ yyyy:number, MM:number, dd:number, yyyymmdd:string } | null}
 */
function parseDateWithContext_(name, parentChain) {
  const base = name.replace(/\.[^.]*$/, '');

  // 1) Full YYYYMMDD
  let m = base.match(/(^|[^\d])(20\d{2})([01]\d)([0-3]\d)([^\d]|$)/);
  if (m) {
    const yyyy = +m[2], MM = +m[3], dd = +m[4];
    if (isValidYMD_(yyyy, MM, dd)) {
      return { yyyy, MM, dd, yyyymmdd: `${m[2]}${m[3]}${m[4]}` };
    }
  }

  // 2) YYYY[-_.]?MM[-_.]?DD
  m = base.match(/(^|[^\d])(20\d{2})[-_.]?([01]?\d)[-_.]?([0-3]?\d)([^\d]|$)/);
  if (m) {
    const yyyy = +m[2], MM = +m[3], dd = +m[4];
    if (isValidYMD_(yyyy, MM, dd)) {
      return { yyyy, MM, dd, yyyymmdd: `${m[2]}${String(MM).padStart(2, '0')}${String(dd).padStart(2, '0')}` };
    }
  }

  // 3) MMYYYY → monthly
  m = base.match(/(^|[^\d])([01]\d)(20\d{2})([^\d]|$)/);
  if (m) {
    const MM = +m[2], yyyy = +m[3];
    if (isValidYMD_(yyyy, MM, 1)) {
      return { yyyy, MM, dd: 0, yyyymmdd: `${yyyy}${String(MM).padStart(2, '0')}00` };
    }
  }

  // 4) MMYY → infer year
  m = base.match(/(^|[^\d])([01]\d)(\d{2})([^\d]|$)/);
  if (m) {
    const MM = +m[2], yy = +m[3];
    let yyyy = inferYearFromParents_(parentChain) || (2000 + yy);
    if (isValidYMD_(yyyy, MM, 1)) {
      return { yyyy, MM, dd: 0, yyyymmdd: `${yyyy}${String(MM).padStart(2, '0')}00` };
    }
  }

  // 5) Named month + year → monthly
  m = base.match(/\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b[^\d]{0,3}(20\d{2}|\d{2})/i);
  if (m) {
    const MM = monthToNumber_(m[1]);
    let yyyy = +m[2];
    if (yyyy < 100) yyyy += 2000;
    return { yyyy, MM, dd: 0, yyyymmdd: `${yyyy}${String(MM).padStart(2, '0')}00` };
  }

  return null;
}

/**
 * Attempt to infer a 4-digit year from nearest parent folder names (up to two levels up).
 *
 * @param {GoogleAppsScript.Drive.Folder[]} parentChain Parent folders (nearest first).
 * @returns {number|null}
 */
function inferYearFromParents_(parentChain) {
  for (let i = 0; i < Math.min(2, parentChain.length); i++) {
    const name = parentChain[i].getName();
    const m = name.match(/(20\d{2})/);
    if (m) return +m[1];
  }
  return null;
}

/**
 * Convert a month string to 1–12.
 *
 * @param {string} monStr Month token (Jan, January, etc.).
 * @returns {number|null}
 */
function monthToNumber_(monStr) {
  const m = monStr.slice(0, 3).toLowerCase();
  const map = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  return map[m] || null;
}

/**
 * Validate a Y-M-D tuple (YYYY, 1–12, 0–31). Uses Date roll to confirm.
 *
 * @param {number} y Year.
 * @param {number} m Month 1–12.
 * @param {number} d Day (0–31; 0 allowed for monthly synthesis).
 * @returns {boolean}
 */
function isValidYMD_(y, m, d) {
  if (!y || !m || d == null) return false;
  if (d === 0) return m >= 1 && m <= 12 && y >= 2000;
  const dt = new Date(y, m - 1, Math.max(1, d));
  return dt.getFullYear() === y && (dt.getMonth() + 1) === m && dt.getDate() === Math.max(1, d);
}

/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  DRIVE HELPERS
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Move a file to a destination folder (Shared‑Drive safe; ensures exactly one parent at all times).
 * Requires Advanced Drive Service (Drive API v2) enabled.
 *
 * @param {GoogleAppsScript.Drive.File} file File to move.
 * @param {GoogleAppsScript.Drive.Folder} destFolder Destination folder.
 * @returns {void}
 */

function moveFileToFolder_(file, destFolder) {
  // Require Advanced Drive to be available (Shared‑Drive safe moves)
  requireAdvancedDrive_();
  var fileId = file.getId();
  var destId = destFolder.getId();

  // Always use Advanced Drive for Shared‑Drive correctness
  var getOpts = { supportsAllDrives: true, supportsTeamDrives: true, fields: 'parents' };
  var meta = Drive.Files.get(fileId, getOpts);
  var parents = (meta.parents || []).map(function (p) { return p.id; });

  // Already the sole parent → nothing to do.
  if (parents.length === 1 && parents[0] === destId) return;

  // If destination already a parent, remove all others (single update, no add)
  if (parents.indexOf(destId) !== -1) {
    var removeOthers = parents.filter(function (id) { return id !== destId; }).join(',');
    if (removeOthers) {
      Drive.Files.update({}, fileId, null, {
        supportsAllDrives: true,
        supportsTeamDrives: true,
        removeParents: removeOthers
      });
    }
    return;
  }

  // Destination not a parent. Shared drives require exactly one parent at a time.
  // Step 1: If multiple parents, reduce to one (keep the first).
  if (parents.length > 1) {
    var keeper = parents[0];
    var removeExtra = parents.filter(function (id) { return id !== keeper; }).join(',');
    Drive.Files.update({}, fileId, null, {
      supportsAllDrives: true,
      supportsTeamDrives: true,
      removeParents: removeExtra
    });
    parents = [keeper];
  }

  // Step 2: If zero parents (rare), just add dest.
  if (parents.length === 0) {
    Drive.Files.update({}, fileId, null, {
      supportsAllDrives: true,
      supportsTeamDrives: true,
      addParents: destId
    });
    return;
  }

  // Step 3: Swap the sole parent for destination in one atomic update.
  var soleParent = parents[0];
  Drive.Files.update({}, fileId, null, {
    supportsAllDrives: true,
    supportsTeamDrives: true,
    addParents: destId,
    removeParents: soleParent
  });
}



/**
 * Walk a folder recursively to a given depth and invoke a callback for each file.
 *
 * @param {GoogleAppsScript.Drive.Folder} folder Starting folder.
 * @param {number} depth Current recursion depth (start with 0).
 * @param {number} maxDepth Maximum depth to recurse.
 * @param {(file: GoogleAppsScript.Drive.File, parentChain: GoogleAppsScript.Drive.Folder[]) => void} fn Callback for files.
 * @returns {void}
 */

function walkFolder_(folder, depth, maxDepth, fn) {
  // Collect parent chain for date inference (this folder + up to 2 parents)
  var parentChain = [];
  var p = folder;
  while (p) {
    parentChain.push(p);
    var it = p.getParents();
    p = it.hasNext() ? it.next() : null;
    if (parentChain.length >= 3) break;
  }

  // Files in this folder
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    fn(f, parentChain);
  }

  // Recurse subfolders
  if (depth < maxDepth) {
    var subs = folder.getFolders();
    while (subs.hasNext()) {
      var sub = subs.next();
      walkFolder_(sub, depth + 1, maxDepth, fn);
    }
  }
}

/**  mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
|    RR |  SCRIPT PROPERTIES HELPERS
|    mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm
*/

/**
 * Extract the extension from a filename (preserves the leading dot). If no extension, returns ''.
 *
 * @param {string} name The filename to inspect.
 * @returns {string} The extension including the dot (e.g., ".xlsx") or an empty string.
 */

function getFileExtension_(name) {
  var m = String(name).match(/(\.[^./\\]+)$/);
  return m ? m[1] : '';
}

/**
 * Get a stored Drive Folder ID from Script Properties by key or null.
 *
 * @param {string} key Script property key, e.g., 'startDataPipeline' or 'unresolvedReports'.
 * @returns {string|null}
 */

function getFolderIdProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || null;
}


/**
 * Resolve and return a Drive Folder by a Script Property key.
 *
 * @param {string} key Script property key that stores a folder ID.
 * @returns {GoogleAppsScript.Drive.Folder}
 * @throws {Error} If the key is not set in Script Properties.
 */
function getFolderByKey(key) {
  var id = getFolderIdProp(key);
  if (!id) throw new Error('No folder ID stored for key "' + key + '".');
  return DriveApp.getFolderById(id);
}

// REPLACE your existing SPECIAL_REPORTS + detectSpecialReport_ with this:

const SPECIAL_REPORTS = [
  // Allow "PULSE" or "PulseSTAR" anywhere (start or after a non-alnum), then require a boundary.
  { key: 'PULSE',     regex: /(?:^|[^A-Za-z0-9])PULSE(?:STAR)?(?=$|[^A-Za-z0-9])/i,     prefix: 'PulseSTAR_' },
  { key: 'BANDWIDTH', regex: /(?:^|[^A-Za-z0-9])BANDWIDTH(?:STAR)?(?=$|[^A-Za-z0-9])/i, prefix: 'BandwidthSTAR_' },
  // RPM typically appears as "RPM" or "RPM_"; boundary still fine
  { key: 'RPM',       regex: /(?:^|[^A-Za-z0-9])RPM(?=$|[^A-Za-z0-9])/i,                 prefix: 'RPM_' }
];

/**
 * If filename indicates a special report, return { key, prefix }, else null.
 * @param {string} name
 * @returns {{ key: 'PULSE'|'BANDWIDTH'|'RPM', prefix: string } | null}
 */
function detectSpecialReport_(name) {
  const s = String(name || '');
  for (const spec of SPECIAL_REPORTS) {
    if (spec.regex.test(s)) return { key: spec.key, prefix: spec.prefix };
  }
  return null;
}



