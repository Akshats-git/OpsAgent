// EmailProcessor.gs — Main orchestration loop.
// Called by the 5-minute time trigger (and by the dashboard's "Run Now" button).
// Implements batch capping, execution-time guard, self-reply detection, and
// idempotency via the "OpsAgent/Processed" Gmail label.

const CHECKPOINT_KEY = 'opsagent_last_run_ts';

// ---------------------------------------------------------------------------
// Entry point — called by time trigger and manually from the dashboard
// ---------------------------------------------------------------------------

function processEmails() {
  const runStart  = Date.now();
  const batchSize = parseInt(getConfigValue('BATCH_SIZE', '10'), 10);

  Logger.log('=== OpsAgent run started ===');

  // ── Fetch unread threads not yet processed ────────────────────────────────
  // The Gmail label is the primary idempotency guard.
  // "-from:me" prevents processing our own auto-replies.
  const query = `is:unread -label:${GMAIL_LABEL} -from:me`;
  let threads;
  try {
    threads = GmailApp.search(query, 0, batchSize);
  } catch (e) {
    Logger.log('Gmail search failed: ' + e.message);
    return;
  }

  if (threads.length === 0) {
    Logger.log('No new threads. Done.');
    return;
  }

  Logger.log(`Found ${threads.length} thread(s) to process.`);

  // Pre-load shared context once per run to avoid redundant sheet reads
  const memoryExamples = getMemoryExamples(null);
  const templates      = getTemplates();
  const orgEmail       = Session.getActiveUser().getEmail().toLowerCase();

  let processed = 0;
  let errors    = 0;

  for (const thread of threads) {
    // ── Execution-time safety guard (4.5 min / 6 min limit) ─────────────────
    if (Date.now() - runStart > 4.5 * 60 * 1000) {
      Logger.log(
        `Approaching 6-min execution limit. ` +
        `Processed ${processed}/${threads.length}. Will resume next run.`
      );
      break;
    }

    try {
      const messages = thread.getMessages();
      const message  = messages[messages.length - 1]; // latest message in thread

      // ── Self-reply guard ──────────────────────────────────────────────────
      const fromField = message.getFrom().toLowerCase();
      if (fromField.includes(orgEmail)) {
        _applyLabelSilently(thread);
        continue;
      }

      const emailData = {
        messageId:  message.getId(),
        threadId:   thread.getId(),
        sender:     message.getFrom(),
        senderName: message.getFrom().replace(/<[^>]+>/, '').trim(),
        subject:    message.getSubject() || '(No Subject)',
        body:       message.getPlainBody() || message.getBody() || '',
        date:       message.getDate(),
      };

      _processOneEmail(emailData, memoryExamples, templates);
      processed++;

      // Brief pause between Gemini calls to stay within quota
      Utilities.sleep(600);

    } catch (e) {
      Logger.log(`Error on thread ${thread.getId()}: ${e.message}`);
      errors++;
      _logProcessingError(thread, e);
    }
  }

  PropertiesService.getScriptProperties().setProperty(CHECKPOINT_KEY, String(Date.now()));
  Logger.log(`=== Run complete — processed: ${processed}, errors: ${errors} ===`);
}

// ---------------------------------------------------------------------------
// Single-email pipeline
// ---------------------------------------------------------------------------

function _processOneEmail(emailData, memoryExamples, templates) {
  Logger.log(`→ "${emailData.subject}" from ${emailData.sender}`);

  // Step 1 — Gemini: classify + extract + draft reply
  const geminiResult = classifyAndExtract(emailData, memoryExamples, templates);
  Logger.log(
    `  Classified: ${geminiResult.category} ` +
    `(${(geminiResult.confidence * 100).toFixed(0)}% confidence)`
  );

  // Step 2 — Policy engine: decide auto vs human review
  const policyDecision = evaluatePolicy(emailData, geminiResult);
  Logger.log(`  Decision: ${policyDecision.action}`);

  // Step 3 — Resolve required actions for this category
  const requiredActions = getRequiredActions(geminiResult.category);

  // Step 4 — Execute
  const results = executeActions(emailData, geminiResult, policyDecision, requiredActions);

  if (results.errors.length > 0) {
    Logger.log(`  Completed with errors: ${results.errors.join(' | ')}`);
  } else {
    Logger.log(`  Done.`);
  }
}

// ---------------------------------------------------------------------------
// Error handling
// ---------------------------------------------------------------------------

function _logProcessingError(thread, error) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
    const msg   = thread.getMessages()[thread.getMessages().length - 1];

    sheet.appendRow([
      new Date(),
      msg.getId(),
      thread.getId(),
      msg.getFrom(),
      msg.getSubject() || '(No Subject)',
      'ERROR',
      0,
      'none',
      'error',
      '',
      '',
      '',
      error.message,
    ]);
  } catch (logErr) {
    Logger.log('Could not log error to AuditLog: ' + logErr.message);
  }

  // Always label the thread to prevent infinite retry on the same broken email
  _applyLabelSilently(thread);
}

function _applyLabelSilently(thread) {
  try {
    const label = GmailApp.getUserLabelByName(GMAIL_LABEL);
    if (label) label.addToThread(thread);
  } catch (e) { /* silent — label missing is non-critical */ }
}
