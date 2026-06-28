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
      _clearRetryCount(emailData.messageId);   // success → reset any retry counter
      processed++;

      // Brief pause between Gemini calls to stay within quota
      Utilities.sleep(600);

    } catch (e) {
      Logger.log(`Error on thread ${thread.getId()}: ${e.message}`);
      errors++;
      _handleProcessingError(thread, e);
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

const MAX_RETRIES = 3;
const QUARANTINE_LABEL = 'OpsAgent/Quarantine';

/**
 * Decides what to do with a failed email based on the failure type:
 *  - Transient (rate limit, timeout, 5xx): leave UNLABELED so the next run
 *    retries it. After MAX_RETRIES attempts, quarantine it.
 *  - Permanent (bad JSON, 400/401/404 config errors): quarantine immediately —
 *    retrying won't help and we mustn't block the batch forever.
 */
function _handleProcessingError(thread, error) {
  const msg     = thread.getMessages()[thread.getMessages().length - 1];
  const emailId = msg.getId();
  const transient = _isTransientError(error);

  if (transient) {
    const attempts = _incrementRetryCount(emailId);
    if (attempts < MAX_RETRIES) {
      // Leave the thread unlabeled — it stays in the queue and retries next run.
      Logger.log(`  Transient error (attempt ${attempts}/${MAX_RETRIES}) — will retry next run.`);
      _logErrorRow(thread, msg, error, `transient_retry_${attempts}`);
      return;
    }
    Logger.log(`  Transient error exhausted ${MAX_RETRIES} retries — quarantining.`);
    _logErrorRow(thread, msg, error, 'quarantined_after_retries');
  } else {
    Logger.log('  Permanent error — quarantining immediately.');
    _logErrorRow(thread, msg, error, 'quarantined_permanent');
  }

  // Quarantine: mark processed (so it leaves the active queue) + tag for the
  // operator to find and re-process manually once the root cause is fixed.
  _clearRetryCount(emailId);
  _applyLabelSilently(thread);
  _quarantineThread(thread);
}

/** Heuristic: is this error worth retrying? */
function _isTransientError(error) {
  const m = String(error && error.message || error).toLowerCase();
  return (
    m.indexOf(' 429') !== -1 || m.indexOf(' 500') !== -1 || m.indexOf(' 502') !== -1 ||
    m.indexOf(' 503') !== -1 || m.indexOf(' 504') !== -1 ||
    m.indexOf('rate limit') !== -1 || m.indexOf('quota') !== -1 ||
    m.indexOf('timeout') !== -1   || m.indexOf('timed out') !== -1 ||
    m.indexOf('unavailable') !== -1 || m.indexOf('overloaded') !== -1 ||
    m.indexOf('address unavailable') !== -1 || m.indexOf('dns') !== -1
  );
}

function _logErrorRow(thread, msg, error, statusTag) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.AUDIT_LOG);
    sheet.appendRow([
      new Date(), msg.getId(), thread.getId(), msg.getFrom(),
      msg.getSubject() || '(No Subject)', 'ERROR', 0, 'none', statusTag,
      '', '', '', String(error && error.message || error),
    ]);
  } catch (logErr) {
    Logger.log('Could not log error to AuditLog: ' + logErr.message);
  }
}

// ── Retry counter (per message ID, stored in Script Properties) ─────────────

function _retryKey(emailId) { return 'retry_' + emailId; }

function _incrementRetryCount(emailId) {
  const props = PropertiesService.getScriptProperties();
  const n = parseInt(props.getProperty(_retryKey(emailId)) || '0', 10) + 1;
  props.setProperty(_retryKey(emailId), String(n));
  return n;
}

function _clearRetryCount(emailId) {
  PropertiesService.getScriptProperties().deleteProperty(_retryKey(emailId));
}

// ── Labels ──────────────────────────────────────────────────────────────────

function _applyLabelSilently(thread) {
  try {
    const label = GmailApp.getUserLabelByName(GMAIL_LABEL);
    if (label) label.addToThread(thread);
  } catch (e) { /* silent — label missing is non-critical */ }
}

function _quarantineThread(thread) {
  try {
    const label = GmailApp.getUserLabelByName(QUARANTINE_LABEL) || GmailApp.createLabel(QUARANTINE_LABEL);
    label.addToThread(thread);
  } catch (e) { /* silent */ }
}

// ---------------------------------------------------------------------------
// Dev / demo helpers — run these manually from the editor
// ---------------------------------------------------------------------------

/**
 * Removes the OpsAgent/Processed and Quarantine labels from ALL threads and
 * clears every retry counter, so the agent will re-scan the whole inbox.
 * Handy when a test email got labeled during an earlier failed run and now
 * "Run Now does nothing". Does NOT delete any sheet data.
 */
function devResetProcessedLabels() {
  let cleared = 0;
  [GMAIL_LABEL, QUARANTINE_LABEL].forEach(name => {
    const label = GmailApp.getUserLabelByName(name);
    if (!label) return;
    let threads, start = 0;
    do {
      threads = label.getThreads(start, 100);
      threads.forEach(t => { label.removeFromThread(t); cleared++; });
      start += 100;
    } while (threads.length === 100);
  });

  // Wipe retry counters
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties())
        .filter(k => k.indexOf('retry_') === 0)
        .forEach(k => props.deleteProperty(k));

  Logger.log(`devResetProcessedLabels: cleared labels from ${cleared} thread(s) and reset retry counters.`);
  return cleared;
}
