// =============================================================================
// CONTACT LOG PROCESSOR — Google Apps Script
// Investment Bank Deal Flow Automation
// =============================================================================
// ARCHITECTURE
//   Raw_Logs  (manual input)  →  OpenAI API  →  Clean_Data  →  Looker Studio
//
// SECURITY MODEL
//   - API key stored exclusively in PropertiesService (never in source code)
//   - Only unprocessed rows (col E = empty) are sent to the API
//   - One LLM call per row — zero bulk data exposure
// =============================================================================


// ---------------------------------------------------------------------------
// CONSTANTS
// ---------------------------------------------------------------------------

const SHEETS = {
  RAW:   'Raw_Logs',
  CLEAN: 'Clean_Data',
};

// Column indices (0-based) in Raw_Logs
const RAW_COL = {
  TIMESTAMP:     0,
  BANKER:        1,
  INVESTOR:      2,
  NOTES:         3,
  STATUS:        4,   // col E — written back by this script
};

// Column indices (0-based) in Clean_Data
const CLEAN_COL = {
  TIMESTAMP:         0,
  BANKER:            1,
  INVESTOR:          2,
  DEAL_STATUS:       3,
  DROP_REASON:       4,
  NEXT_STEPS:        5,
  NEXT_ACTION_DATE:  6,
};

const VALID_STATUSES = [
  'Approché',
  'Teaser Envoyé',
  'NDA Signé',
  'Management Pres',
  'En Due Diligence',
  'Offre Soumise',
  'Pass/Dropped',
];

const VALID_DROP_REASONS = [
  'Valuation',
  'Sector Mismatch',
  'Fund Lifecycle',
  'Conflict of Interest',
  'Other',
];

const STATUS = {
  PROCESSED: 'Traité',
  ERROR:     'Erreur',
};

const OPENAI_API_URL   = 'https://api.openai.com/v1/chat/completions';
const OPENAI_MODEL     = 'gpt-4o';
const MAX_OUTPUT_TOKENS = 512;
const ROW_DELAY_MS     = 600;   // Politeness delay between API calls (ms)


// ---------------------------------------------------------------------------
// MAIN ENTRY POINT
// ---------------------------------------------------------------------------

/**
 * processContactLogs()
 *
 * Reads every row in Raw_Logs where column E is empty, calls Claude to
 * structure the free-text notes, writes the result to Clean_Data, and
 * marks the source row as "Traité" (or "Erreur" on failure).
 *
 * Attach this function to a time-based trigger (see createTimeTrigger()).
 */
function processContactLogs() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet   = _getSheet(ss, SHEETS.RAW);
  const cleanSheet = _getSheet(ss, SHEETS.CLEAN);

  if (!rawSheet || !cleanSheet) return; // logged inside _getSheet

  const allRows = rawSheet.getDataRange().getValues();
  if (allRows.length <= 1) {
    Logger.log('[INFO] No data rows found in Raw_Logs.');
    return;
  }

  let processed = 0;
  let errors    = 0;

  for (let i = 1; i < allRows.length; i++) {
    const row    = allRows[i];
    const status = row[RAW_COL.STATUS];

    // Skip already-handled rows
    if (status !== '' && status !== null && status !== undefined) continue;

    const notes = (row[RAW_COL.NOTES] || '').toString().trim();
    if (!notes) {
      Logger.log(`[SKIP] Row ${i + 1}: empty notes.`);
      continue;
    }

    const banker   = (row[RAW_COL.BANKER]   || '').toString().trim();
    const investor = (row[RAW_COL.INVESTOR] || '').toString().trim();
    const timestamp = row[RAW_COL.TIMESTAMP];

    try {
      const structured = _callOpenAIAPI(banker, investor, notes);

      _appendToCleanData(cleanSheet, {
        timestamp,
        banker,
        investor,
        dealStatus:     structured.deal_status,
        dropReason:     structured.drop_reason     || '',
        nextSteps:      structured.next_steps      || '',
        nextActionDate: structured.next_action_date || '',
      });

      // Mark as processed — single-cell write is lightweight
      rawSheet.getRange(i + 1, RAW_COL.STATUS + 1).setValue(STATUS.PROCESSED);
      processed++;
      Logger.log(`[OK] Row ${i + 1}: ${investor} → ${structured.deal_status}`);

    } catch (err) {
      rawSheet.getRange(i + 1, RAW_COL.STATUS + 1).setValue(STATUS.ERROR);
      errors++;
      Logger.log(`[ERROR] Row ${i + 1}: ${err.message}`);
    }

    Utilities.sleep(ROW_DELAY_MS);
  }

  Logger.log(`[DONE] Processed: ${processed} | Errors: ${errors}`);
}


// ---------------------------------------------------------------------------
// OPENAI API LAYER
// ---------------------------------------------------------------------------

/**
 * _callOpenAIAPI(banker, investor, notes)
 *
 * Sends a single row's data to OpenAI and returns a validated JS object.
 * Only the minimum required context is transmitted — never the full sheet.
 * Uses response_format: json_object to guarantee parseable output.
 *
 * @param  {string} banker   - Banker name
 * @param  {string} investor - Investor/fund name
 * @param  {string} notes    - Free-text call notes
 * @returns {Object}          Validated structured object
 * @throws {Error}            On HTTP error, malformed JSON, or invalid values
 */
function _callOpenAIAPI(banker, investor, notes) {
  const apiKey = _getApiKey();

  const systemPrompt = `\
You are a precision data-extraction engine for an investment bank's M&A deal pipeline.
Your ONLY task: read banker call notes and return a single, raw JSON object.

OUTPUT RULES — ABSOLUTE:
1. Return ONLY the JSON object. Zero text before or after it.
2. No markdown, no code fences, no explanations, no apologies.
3. The JSON must have EXACTLY these four keys:
   - "deal_status"       : string — exactly one of the allowed values listed below
   - "drop_reason"       : string or null
   - "next_steps"        : string or null
   - "next_action_date"  : string (YYYY-MM-DD) or null

ALLOWED VALUES:
deal_status  → "Approché" | "Teaser Envoyé" | "NDA Signé" | "Management Pres" |
               "En Due Diligence" | "Offre Soumise" | "Pass/Dropped"
drop_reason  → "Valuation" | "Sector Mismatch" | "Fund Lifecycle" |
               "Conflict of Interest" | "Other"
               Use null when deal_status ≠ "Pass/Dropped".
next_action_date → Extract from notes if present. Format as YYYY-MM-DD.
                   Use null if no date is mentioned.

INFERENCE RULES:
- If status is ambiguous, choose the most recent confirmed stage (conservative).
- If the investor declined or passed, set deal_status = "Pass/Dropped".
- Never fabricate information not present in the notes.
- For relative dates (e.g. "next Tuesday"), resolve to an absolute date only
  if you can be certain; otherwise use null.

EXAMPLE OUTPUT:
{"deal_status":"NDA Signé","drop_reason":null,"next_steps":"Send the Information Memorandum draft for legal review before circulating to investor.","next_action_date":"2025-07-15"}`;

  const userContent = `Banker: ${banker}\nInvestor / Fund: ${investor}\nCall Notes:\n${notes}`;

  const payload = {
    model:           OPENAI_MODEL,
    max_tokens:      MAX_OUTPUT_TOKENS,
    response_format: { type: 'json_object' },  // Enforces JSON-only output
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userContent  },
    ],
  };

  const options = {
    method:          'post',
    contentType:     'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response     = UrlFetchApp.fetch(OPENAI_API_URL, options);
  const httpCode     = response.getResponseCode();
  const responseText = response.getContentText();

  if (httpCode !== 200) {
    throw new Error(`HTTP ${httpCode} from OpenAI API: ${responseText.substring(0, 300)}`);
  }

  const apiBody = JSON.parse(responseText);

  if (!apiBody.choices || !apiBody.choices[0] || !apiBody.choices[0].message || !apiBody.choices[0].message.content) {
    throw new Error(`Unexpected API response structure: ${responseText.substring(0, 300)}`);
  }

  return _validateAndParse(apiBody.choices[0].message.content.trim());
}


// ---------------------------------------------------------------------------
// RESPONSE VALIDATION
// ---------------------------------------------------------------------------

/**
 * _validateAndParse(rawText)
 *
 * Parses the LLM text output, validates field values, and returns a clean object.
 * Throws descriptive errors so the caller can mark the row "Erreur".
 *
 * @param  {string} rawText - Raw text from the LLM (expected to be pure JSON)
 * @returns {Object}         Validated structured data
 */
function _validateAndParse(rawText) {
  let data;
  try {
    // Strip accidental markdown fences if the model misbehaved
    const cleaned = rawText.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
    data = JSON.parse(cleaned);
  } catch (_) {
    throw new Error(`LLM returned non-JSON content: "${rawText.substring(0, 200)}"`);
  }

  // Validate deal_status
  if (!VALID_STATUSES.includes(data.deal_status)) {
    throw new Error(`Invalid deal_status: "${data.deal_status}"`);
  }

  // Validate drop_reason
  if (data.deal_status === 'Pass/Dropped') {
    if (data.drop_reason === null || data.drop_reason === undefined || data.drop_reason === '') {
      data.drop_reason = 'Other'; // Safe default
    } else if (!VALID_DROP_REASONS.includes(data.drop_reason)) {
      Logger.log(`[WARN] Unknown drop_reason "${data.drop_reason}" — defaulting to "Other".`);
      data.drop_reason = 'Other';
    }
  } else {
    data.drop_reason = null; // Force null for non-dropped deals
  }

  // Validate date format (loose check: YYYY-MM-DD or null)
  if (data.next_action_date !== null && data.next_action_date !== undefined && data.next_action_date !== '') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data.next_action_date)) {
      Logger.log(`[WARN] Invalid date format "${data.next_action_date}" — setting to null.`);
      data.next_action_date = null;
    }
  } else {
    data.next_action_date = null;
  }

  return data;
}


// ---------------------------------------------------------------------------
// SHEET OPERATIONS
// ---------------------------------------------------------------------------

/**
 * _appendToCleanData(sheet, data)
 *
 * Appends one structured row to the Clean_Data sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} data
 */
function _appendToCleanData(sheet, data) {
  sheet.appendRow([
    data.timestamp,
    data.banker,
    data.investor,
    data.dealStatus,
    data.dropReason     || '',
    data.nextSteps      || '',
    data.nextActionDate || '',
  ]);
}

/**
 * _getSheet(ss, name)
 *
 * Retrieves a sheet by name with a clear error log if missing.
 *
 * @param  {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param  {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function _getSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    Logger.log(`[ERROR] Sheet "${name}" not found. Run initializeSheets() first.`);
  }
  return sheet;
}


// ---------------------------------------------------------------------------
// SECURE CONFIGURATION
// ---------------------------------------------------------------------------

/**
 * _getApiKey()
 *
 * Retrieves the OpenAI API key from Script Properties.
 * Throws if not set to avoid silent failures.
 *
 * @returns {string}
 */
function _getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) {
    throw new Error(
      'OPENAI_API_KEY not set. Open Extensions → Apps Script → Project Settings → Script Properties and add it.'
    );
  }
  return key;
}

/**
 * setApiKey()
 *
 * ONE-TIME SETUP: stores your OpenAI API key securely.
 *
 * HOW TO USE:
 *   1. Replace 'YOUR_KEY_HERE' with your real key (starts with "sk-...").
 *   2. Run this function ONCE from the Apps Script editor.
 *   3. Immediately remove or blank out the key string from this source code.
 *   4. The key is now stored in Script Properties — never in the code.
 */
function setApiKey() {
  const key = 'YOUR_KEY_HERE'; // ← replace, run once, then clear this line
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', key);
  Logger.log('[SETUP] API key stored in Script Properties.');
}


// ---------------------------------------------------------------------------
// ONE-TIME SETUP UTILITIES
// ---------------------------------------------------------------------------

/**
 * initializeSheets()
 *
 * Creates Raw_Logs and Clean_Data sheets with correct headers if they
 * do not already exist. Safe to re-run — does not overwrite existing data.
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Raw_Logs ──────────────────────────────────────────────────────────────
  let rawSheet = ss.getSheetByName(SHEETS.RAW);
  if (!rawSheet) {
    rawSheet = ss.insertSheet(SHEETS.RAW);
    Logger.log('[SETUP] Created sheet: Raw_Logs');
  }
  if (rawSheet.getLastRow() === 0) {
    rawSheet.getRange(1, 1, 1, 5).setValues([[
      'Timestamp',
      'Nom du Banquier',
      "Nom de l'Investisseur",
      'Notes du call',
      'Statut de traitement',
    ]]);
    rawSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#D9EAD3');
    rawSheet.setColumnWidth(4, 450); // Notes column — needs room
    Logger.log('[SETUP] Raw_Logs headers written.');
  }

  // ── Clean_Data ────────────────────────────────────────────────────────────
  let cleanSheet = ss.getSheetByName(SHEETS.CLEAN);
  if (!cleanSheet) {
    cleanSheet = ss.insertSheet(SHEETS.CLEAN);
    Logger.log('[SETUP] Created sheet: Clean_Data');
  }
  if (cleanSheet.getLastRow() === 0) {
    cleanSheet.getRange(1, 1, 1, 7).setValues([[
      'Timestamp',
      'Nom du Banquier',
      "Nom de l'Investisseur",
      'Statut du Deal',
      'Drop Reason',
      'Next Steps',
      'Date Prochaine Action',
    ]]);
    cleanSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#CFE2F3');
    cleanSheet.setColumnWidth(6, 350); // Next Steps
    cleanSheet.setColumnWidth(5, 220); // Drop Reason
    Logger.log('[SETUP] Clean_Data headers written.');
  }

  Logger.log('[SETUP] initializeSheets() complete.');
}

/**
 * createTimeTrigger()
 *
 * Installs a time-based trigger so processContactLogs() runs automatically
 * every hour. Safe to re-run — removes duplicate triggers first.
 *
 * Run ONCE after deploying the script.
 */
function createTimeTrigger() {
  // Remove stale triggers for this function
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processContactLogs')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('processContactLogs')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('[SETUP] Time trigger created: processContactLogs → every 1 hour.');
}

/**
 * retryErrorRows()
 *
 * Convenience function: resets all "Erreur" rows back to empty so they
 * will be picked up on the next processContactLogs() run.
 * Use after fixing a configuration issue (e.g., wrong API key).
 */
function retryErrorRows() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = _getSheet(ss, SHEETS.RAW);
  if (!rawSheet) return;

  const data  = rawSheet.getDataRange().getValues();
  let cleared = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][RAW_COL.STATUS] === STATUS.ERROR) {
      rawSheet.getRange(i + 1, RAW_COL.STATUS + 1).clearContent();
      cleared++;
    }
  }

  Logger.log(`[RETRY] Cleared ${cleared} error row(s) for reprocessing.`);
}
