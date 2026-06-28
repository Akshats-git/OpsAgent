// DigestReport.gs — Generates a weekly operations digest as a Google Doc,
// stored in a Drive folder. Runs on a weekly trigger and on demand from the
// dashboard. Adds Google Docs + Google Drive to the tech stack and satisfies
// the "generate summaries / generate reports" example actions.

// ---------------------------------------------------------------------------
// Public entry points
// ---------------------------------------------------------------------------

/** Weekly-trigger target. Skips silently if disabled in Config. */
function generateWeeklyDigest() {
  if (getConfigValue('DIGEST_ENABLED', 'true') !== 'true') {
    Logger.log('Digest disabled in Config — skipping.');
    return { success: false, message: 'Digest disabled in Config.' };
  }
  return _buildDigest(7);
}

/** Dashboard "Generate Digest" button target. Always runs. */
function generateDigestNow() {
  return _buildDigest(7);
}

// ---------------------------------------------------------------------------
// Core
// ---------------------------------------------------------------------------

function _buildDigest(days) {
  try {
    const orgName = getConfigValue('ORG_NAME', 'Our Organization');
    const stats   = _collectStats(days);

    const now   = new Date();
    const start = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
    const range = `${_fmtDate(start)} – ${_fmtDate(now)}`;
    const title = `${orgName} — Weekly Operations Digest (${_fmtDate(now)})`;

    // ── Build the Google Doc ────────────────────────────────────────────────
    const doc  = DocumentApp.create(title);
    const body = doc.getBody();

    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(`Reporting period: ${range}`)
        .setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

    // Executive summary
    body.appendParagraph('Executive Summary').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(
      `Over the last ${days} days, the agent processed ${stats.total} email(s). ` +
      `${stats.auto} were handled automatically (${stats.autoRate}% automation rate), ` +
      `${stats.review} were escalated for human review, and ` +
      `${stats.infoRequested} had missing details and were sent a clarification request. ` +
      `Average classification confidence was ${stats.avgConf}%.`
    );

    // Key metrics table
    body.appendParagraph('Key Metrics').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendTable([
      ['Metric', 'Value'],
      ['Total processed',        String(stats.total)],
      ['Auto-handled',           String(stats.auto)],
      ['Sent to human review',   String(stats.review)],
      ['Info requests sent',     String(stats.infoRequested)],
      ['Complaints received',    String(stats.complaints)],
      ['Quarantined (errors)',   String(stats.quarantined)],
      ['New memory corrections', String(stats.memoryAdded)],
      ['Avg. confidence',        stats.avgConf + '%'],
      ['Automation rate',        stats.autoRate + '%'],
    ]);

    // Category breakdown
    body.appendParagraph('Category Breakdown').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const catTable = [['Category', 'Count']];
    Object.keys(stats.byCategory).sort((a, b) => stats.byCategory[b] - stats.byCategory[a])
      .forEach(cat => catTable.push([cat, String(stats.byCategory[cat])]));
    if (catTable.length === 1) catTable.push(['(none)', '0']);
    body.appendTable(catTable);

    // Attention items
    body.appendParagraph('Items Needing Attention').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (stats.attention.length === 0) {
      body.appendParagraph('Nothing outstanding — all clear. 🎉');
    } else {
      stats.attention.slice(0, 15).forEach(item =>
        body.appendListItem(`${item.category.toUpperCase()} — "${item.subject}" from ${item.sender}`)
            .setGlyphType(DocumentApp.GlyphType.BULLET)
      );
    }

    // Learning trend
    body.appendParagraph('Agent Learning').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(
      stats.memoryAdded > 0
        ? `${stats.memoryAdded} correction(s) were added to the agent's memory this period. ` +
          `These now serve as few-shot examples, improving future classification accuracy.`
        : `No new corrections this period — the agent's classifications were accepted as-is.`
    );

    body.appendParagraph('')
        .appendText('Generated automatically by OpsAgent.')
        .setItalic(true).setForegroundColor('#888888');

    doc.saveAndClose();

    // ── File it in the reports folder ───────────────────────────────────────
    const folder = _getReportsFolder();
    DriveApp.getFileById(doc.getId()).moveTo(folder);
    const url = doc.getUrl();

    Logger.log('Digest created: ' + url);

    // ── Notify the team via Chat (if configured) ────────────────────────────
    _notifyDigest(orgName, range, url, stats);

    return { success: true, url: url, message: 'Weekly digest generated.' };
  } catch (e) {
    Logger.log('Digest failed: ' + e.message);
    return { success: false, message: 'Digest failed: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// Stats collection
// ---------------------------------------------------------------------------

function _collectStats(days) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const cutoff    = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  const auditRows = _readRecent(ss.getSheetByName(SHEET_NAMES.AUDIT_LOG), cutoff);
  const memRows   = _readRecent(ss.getSheetByName(SHEET_NAMES.MEMORY), cutoff);

  const stats = {
    total: 0, auto: 0, review: 0, infoRequested: 0, complaints: 0,
    quarantined: 0, memoryAdded: memRows.length, byCategory: {},
    attention: [], avgConf: 0, autoRate: 0,
  };

  let confSum = 0, confCount = 0;
  auditRows.forEach(r => {
    const category = String(r[5]);
    const status   = String(r[8]);

    if (category === 'ERROR') {
      if (status.indexOf('quarantin') !== -1) stats.quarantined++;
      return;
    }

    stats.total++;
    stats.byCategory[category] = (stats.byCategory[category] || 0) + 1;

    if (status === 'auto')         stats.auto++;
    if (status === 'human_review') stats.review++;
    if (status === 'request_info') stats.infoRequested++;
    if (category === 'complaint')  stats.complaints++;

    const conf = parseFloat(r[6]);
    if (!isNaN(conf) && conf > 0) { confSum += conf; confCount++; }

    // Flag complaints and review items as attention-worthy
    if (category === 'complaint' || status === 'human_review') {
      stats.attention.push({ category: category, subject: String(r[4]), sender: String(r[3]) });
    }
  });

  stats.avgConf  = confCount ? Math.round((confSum / confCount) * 100) : 0;
  stats.autoRate = stats.total ? Math.round((stats.auto / stats.total) * 100) : 0;
  return stats;
}

/** Reads rows from a sheet whose first column (timestamp) is on/after cutoff. */
function _readRecent(sheet, cutoff) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const cols = sheet.getLastColumn();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues();
  return rows.filter(r => r[0] && new Date(r[0]) >= cutoff);
}

// ---------------------------------------------------------------------------
// Drive + notifications
// ---------------------------------------------------------------------------

function _getReportsFolder() {
  const name = getConfigValue('REPORTS_FOLDER_NAME', 'OpsAgent Reports');
  const it   = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function _notifyDigest(orgName, range, url, stats) {
  const webhookUrl = getConfigValue('CHAT_WEBHOOK_URL', '');
  if (!webhookUrl || getConfigValue('CHAT_ALERTS_ENABLED', 'true') !== 'true') return;

  const payload = {
    cards: [{
      header: { title: '📊 Weekly Digest Ready', subtitle: range },
      sections: [{
        widgets: [
          { keyValue: { topLabel: 'Processed',     content: String(stats.total) } },
          { keyValue: { topLabel: 'Automation',    content: stats.autoRate + '%' } },
          { keyValue: { topLabel: 'Needs review',  content: String(stats.review) } },
          { buttons: [{ textButton: { text: 'OPEN REPORT', onClick: { openLink: { url: url } } } }] },
        ],
      }],
    }],
  };
  try {
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
  } catch (e) { Logger.log('Digest Chat notify failed: ' + e.message); }
}

function _fmtDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy');
}
