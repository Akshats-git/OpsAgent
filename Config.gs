// Config.gs — Configuration management with CacheService-backed reads.
// All org settings live in the Config sheet so non-technical admins can
// tweak behaviour without touching code.

const SHEET_NAMES = {
  AUDIT_LOG:    'AuditLog',
  HUMAN_REVIEW: 'HumanReview',
  MEMORY:       'Memory',
  REGISTRATIONS:'Registrations',
  CONFIG:       'Config',
  TEMPLATES:    'Templates',
};

const GMAIL_LABEL        = 'OpsAgent/Processed';
const TRIGGER_FUNCTION   = 'processEmails';
const TRIGGER_INTERVAL_M = 5;

const CONFIG_CACHE_KEY = 'opsagent_config_v1';
const CONFIG_CACHE_TTL = 300; // seconds — matches trigger interval

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Returns a plain object of every key→value pair from the Config sheet.
 * Result is cached for 5 minutes so we don't hit the sheet on every email.
 */
function getConfig() {
  const cache = CacheService.getScriptCache();
  const hit   = cache.get(CONFIG_CACHE_KEY);
  if (hit) return JSON.parse(hit);

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) throw new Error('Config sheet missing — run setup() first.');

  const rows   = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]).trim();
    const val = rows[i][1];
    if (key) config[key] = val !== undefined ? String(val) : '';
  }

  cache.put(CONFIG_CACHE_KEY, JSON.stringify(config), CONFIG_CACHE_TTL);
  return config;
}

function getConfigValue(key, fallback) {
  if (fallback === undefined) fallback = '';
  const val = getConfig()[key];
  return (val !== undefined && val !== '') ? val : fallback;
}

/** Call after saving new values to the Config sheet to force a fresh read. */
function clearConfigCache() {
  CacheService.getScriptCache().remove(CONFIG_CACHE_KEY);
}

/**
 * Returns the Gemini API key from Script Properties.
 * Stored there (not in the sheet) so it never appears in spreadsheet cells.
 */
function getGeminiApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error(
      'GEMINI_API_KEY not set.\n' +
      'Go to Extensions → Apps Script → Project Settings → Script Properties and add it.'
    );
  }
  return key;
}

/**
 * Returns the OpenAI API key from Script Properties.
 */
function getOpenAIApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) {
    throw new Error(
      'OPENAI_API_KEY not set.\n' +
      'Go to Extensions → Apps Script → Project Settings → Script Properties and add it.'
    );
  }
  return key;
}

/** Returns the current AI provider — 'gemini' or 'openai'. */
function getAIProvider() {
  return getConfigValue('AI_PROVIDER', 'gemini').toLowerCase();
}
