// Dashboard.gs — Serves the web-app dashboard and exposes server-side
// data functions called via google.script.run from the frontend.

// ---------------------------------------------------------------------------
// Web-app entry point
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('OpsAgent Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------------------
// Data functions (called from frontend via google.script.run)
// ---------------------------------------------------------------------------

/** Returns all stats needed for the top-bar cards and charts. */
function getDashboardStats() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet  = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  const reviewSheet = ss.getSheetByName(SHEET_NAMES.HUMAN_REVIEW);
  const memSheet    = ss.getSheetByName(SHEET_NAMES.MEMORY);

  const auditRows  = _sheetRows(auditSheet,  13);
  const reviewRows = _sheetRows(reviewSheet, 14);

  const valid      = auditRows.filter(r => r[5] && r[5] !== 'ERROR');
  const autoRows   = valid.filter(r => r[8] === 'auto');
  const pendingHR  = reviewRows.filter(r => r[12] === 'Pending');
  const errorRows  = auditRows.filter(r => r[5] === 'ERROR');

  const confidences = valid.map(r => parseFloat(r[6])).filter(c => !isNaN(c) && c > 0);
  const avgConf     = confidences.length
    ? (confidences.reduce((a, b) => a + b, 0) / confidences.length * 100).toFixed(1)
    : 0;

  const catCount = {};
  valid.forEach(r => {
    const cat = r[5];
    catCount[cat] = (catCount[cat] || 0) + 1;
  });

  const memCount = memSheet && memSheet.getLastRow() > 1 ? memSheet.getLastRow() - 1 : 0;

  return {
    totalProcessed:   valid.length,
    autoReplied:      autoRows.length,
    humanReview:      pendingHR.length,
    errors:           errorRows.length,
    memoryCorrections:memCount,
    avgConfidence:    Number(avgConf),
    autoRate:         valid.length ? ((autoRows.length / valid.length) * 100).toFixed(1) : '0',
    categoryDistribution: catCount,
    orgName:          getConfigValue('ORG_NAME', 'My Organization'),
  };
}

/** Returns the N most recent audit log entries (newest first). */
function getRecentAuditLog(limit) {
  if (!limit) limit = 25;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  const rows  = _sheetRows(sheet, 14);
  if (!rows.length) return [];

  return rows.slice(-limit).reverse().map(r => ({
    timestamp:    r[0] ? new Date(r[0]).toISOString() : '',
    sender:       r[3],
    subject:      r[4],
    category:     r[5],
    confidence:   parseFloat(r[6]) || 0,
    actionsTaken: r[7],
    status:       r[8],
    sensitiveFlag:r[10],
    error:        r[12],
    reasoning:    r[13] || '',
  }));
}

/** Returns all pending Human Review items. */
function getHumanReviewQueue() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HUMAN_REVIEW);
  const rows  = _sheetRows(sheet, 15);
  if (!rows.length) return [];

  return rows
    .map((r, i) => ({ _row: i + 2, data: r }))
    .filter(({ data: r }) => r[12] === 'Pending')
    .map(({ data: r }) => ({
      emailId:         String(r[1]),
      timestamp:       r[0] ? new Date(r[0]).toISOString() : '',
      sender:          r[3],
      subject:         r[4],
      agentCategory:   r[5],
      confidence:      parseFloat(r[6]) || 0,
      draftReply:      r[7],
      escalationReason:r[8],
      reasoning:       r[14] || '',
    }));
}

/**
 * Buckets the AuditLog by day for the last N days so the dashboard can draw a
 * trend chart. Returns { labels:[], auto:[], review:[], info:[] } aligned by index.
 */
function getActivityTrend(days) {
  if (!days) days = 7;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const rows  = _sheetRows(ss.getSheetByName(SHEET_NAMES.AUDIT_LOG), 14);

  // Build an ordered list of day-keys for the window (oldest → newest)
  const tz     = Session.getScriptTimeZone();
  const labels = [], keys = [], buckets = {};
  for (let i = days - 1; i >= 0; i--) {
    const d   = new Date();
    d.setDate(d.getDate() - i);
    const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    keys.push(key);
    labels.push(Utilities.formatDate(d, tz, 'MMM d'));
    buckets[key] = { auto: 0, review: 0, info: 0 };
  }

  rows.forEach(r => {
    if (!r[0] || r[5] === 'ERROR') return;
    const key = Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd');
    if (!buckets[key]) return;
    const status = String(r[8]);
    if (status === 'human_review')      buckets[key].review++;
    else if (status === 'request_info') buckets[key].info++;
    else                                buckets[key].auto++;
  });

  return {
    labels: labels,
    auto:   keys.map(k => buckets[k].auto),
    review: keys.map(k => buckets[k].review),
    info:   keys.map(k => buckets[k].info),
  };
}

/** Manually triggers a processing run — called by the "Run Now" button. */
function triggerManualRun() {
  try {
    processEmails();
    return { success: true, message: 'Processing run completed.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/** Returns recent Memory corrections for the dashboard Memory tab. */
function getMemoryEntries(limit) {
  if (!limit) limit = 20;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMORY);
  const rows  = _sheetRows(sheet, 8);
  if (!rows.length) return [];

  return rows.slice(-limit).reverse().map(r => ({
    timestamp:         r[0] ? new Date(r[0]).toISOString() : '',
    subject:           r[2],
    originalCategory:  r[4],
    correctedCategory: r[5],
    addedBy:           r[6],
  }));
}

// ---------------------------------------------------------------------------
// Shared helper
// ---------------------------------------------------------------------------

function _sheetRows(sheet, numCols) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const n    = sheet.getLastRow() - 1;
  // Clamp to the sheet's real width so requesting a not-yet-migrated column
  // (e.g. Reasoning) never throws an out-of-bounds error.
  const cols = Math.min(numCols, sheet.getLastColumn());
  return sheet.getRange(2, 1, n, cols).getValues();
}
