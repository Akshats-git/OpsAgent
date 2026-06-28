// ActionRouter.gs — Executes every action the policy engine has decided on.
// Handles Gmail replies, Sheets writes, Calendar events, Google Forms links,
// Google Chat webhooks, and audit/human-review logging.

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Runs all required actions for a classified email.
 *
 * @param {Object} emailData       - Raw email data
 * @param {Object} geminiResult    - Output of classifyAndExtract()
 * @param {Object} policyDecision  - Output of evaluatePolicy()
 * @param {Array}  requiredActions - Output of getRequiredActions()
 * @returns {Object} results       - Flags for what happened and any errors
 */
function executeActions(emailData, geminiResult, policyDecision, requiredActions) {
  const results = {
    reply_sent:       false,
    calendar_event_id: null,
    form_sent:        false,
    chat_alert_sent:  false,
    audit_logged:     false,
    human_review_logged: false,
    info_requested:   false,
    errors:           [],
  };

  const autoReplyEnabled = getConfigValue('AUTO_REPLY_ENABLED', 'true') === 'true';
  const orgName          = getConfigValue('ORG_NAME', 'Our Organization');
  const isAutoAction     = policyDecision.action === 'auto';
  const isRequestInfo    = policyDecision.action === 'request_info';

  // ── 1. Send reply (only when auto-actioning and reply is enabled) ─────────
  if (requiredActions.includes('reply') && autoReplyEnabled && isAutoAction) {
    try {
      _sendReply(emailData, geminiResult.draft_reply, orgName);
      results.reply_sent = true;
    } catch (e) {
      results.errors.push('reply: ' + e.message);
    }
  }

  // ── 1b. Request more info (confident category, but details missing) ───────
  if (isRequestInfo && autoReplyEnabled) {
    try {
      _sendInfoRequest(emailData, policyDecision.missingFields, geminiResult, orgName);
      results.info_requested = true;
    } catch (e) {
      results.errors.push('info_request: ' + e.message);
    }
  }

  // ── 2. Update Registrations sheet + Calendar event ───────────────────────
  if (requiredActions.includes('registrations_sheet') && isAutoAction) {
    try {
      const calId = _updateRegistrations(
        emailData,
        geminiResult.extracted_fields,
        requiredActions.includes('calendar_event')
      );
      results.calendar_event_id = calId;
    } catch (e) {
      results.errors.push('registrations: ' + e.message);
    }
  }

  // ── 3. Send Google Forms link ─────────────────────────────────────────────
  if (requiredActions.includes('send_form') && isAutoAction) {
    try {
      _sendFormLink(emailData, geminiResult.category, orgName);
      results.form_sent = true;
    } catch (e) {
      results.errors.push('form_link: ' + e.message);
    }
  }

  // ── 4. Google Chat alert ──────────────────────────────────────────────────
  if (requiredActions.includes('chat_alert') || policyDecision.action === 'human_review') {
    try {
      _sendChatAlert(emailData, geminiResult, policyDecision);
      results.chat_alert_sent = true;
    } catch (e) {
      results.errors.push('chat_alert: ' + e.message);
    }
  }

  // ── 5. Log to Human Review sheet (before audit log so audit can reference it) ──
  if (policyDecision.action === 'human_review') {
    try {
      _logHumanReview(emailData, geminiResult, policyDecision);
      results.human_review_logged = true;
    } catch (e) {
      results.errors.push('human_review_log: ' + e.message);
    }
  }

  // ── 6. Audit log — always ─────────────────────────────────────────────────
  try {
    _logAudit(emailData, geminiResult, policyDecision, results, requiredActions);
    results.audit_logged = true;
  } catch (e) {
    results.errors.push('audit_log: ' + e.message);
  }

  // ── 7. Gmail labelling ────────────────────────────────────────────────────
  // For request_info we do NOT mark Processed — we mark the thread read and tag
  // it AwaitingInfo. When the sender replies with the details, the thread goes
  // unread again, re-enters the queue, and gets completed on the next run.
  try {
    if (isRequestInfo) {
      _markAwaitingInfo(emailData.threadId);
    } else {
      _applyProcessedLabel(emailData.threadId);
    }
  } catch (e) {
    results.errors.push('label: ' + e.message);
  }

  return results;
}

// ---------------------------------------------------------------------------
// Gmail
// ---------------------------------------------------------------------------

function _sendReply(emailData, draftReply, orgName) {
  const thread = GmailApp.getThreadById(emailData.threadId);
  if (!thread) throw new Error('Thread not found: ' + emailData.threadId);
  const body = `${draftReply}\n\n—\n${orgName} Team\n*(Automated response — powered by AI)*`;
  thread.reply(body);
}

function _sendFormLink(emailData, category, orgName) {
  const urlKey = category === 'volunteer_application' ? 'FORMS_VOLUNTEER_URL' : 'FORMS_SPEAKER_URL';
  const formUrl = getConfigValue(urlKey, '');
  if (!formUrl) return; // Form URL not configured — silently skip

  const label = category === 'volunteer_application' ? 'volunteer application form' : 'speaker application form';
  const body  =
    `Hi,\n\nThank you for your interest! To help us learn more about you, ` +
    `please complete our ${label}:\n\n${formUrl}\n\n` +
    `We'll review your submission and get back to you shortly.\n\n` +
    `—\n${orgName} Team`;

  const thread = GmailApp.getThreadById(emailData.threadId);
  thread.reply(body);
}

function _applyProcessedLabel(threadId) {
  const label  = GmailApp.getUserLabelByName(GMAIL_LABEL);
  const thread = GmailApp.getThreadById(threadId);
  if (label && thread) label.addToThread(thread);
  // If this thread was previously awaiting info, clear that tag now.
  const awaiting = GmailApp.getUserLabelByName(AWAITING_INFO_LABEL);
  if (awaiting && thread) {
    try { awaiting.removeFromThread(thread); } catch (e) { /* not present */ }
  }
}

/**
 * Sends a friendly clarification email listing the specific details we still
 * need, then marks the thread read + tags it AwaitingInfo so it re-enters the
 * queue only when the sender replies.
 */
function _sendInfoRequest(emailData, missingFields, geminiResult, orgName) {
  const thread = GmailApp.getThreadById(emailData.threadId);
  if (!thread) throw new Error('Thread not found: ' + emailData.threadId);

  const fields = (missingFields && missingFields.length)
    ? missingFields : ['a few more details to proceed'];
  const bullets = fields.map(f => `  • ${f}`).join('\n');

  const body =
    `Thanks so much for reaching out — we'd love to help with this!\n\n` +
    `Before we can complete your request, could you share:\n\n${bullets}\n\n` +
    `Just reply to this email with those details and we'll take it from there.\n\n` +
    `—\n${orgName} Team\n*(Automated response — powered by AI)*`;

  thread.reply(body);
  thread.markRead();
}

function _markAwaitingInfo(threadId) {
  const thread = GmailApp.getThreadById(threadId);
  if (!thread) return;
  const label = GmailApp.getUserLabelByName(AWAITING_INFO_LABEL) ||
                GmailApp.createLabel(AWAITING_INFO_LABEL);
  label.addToThread(thread);
  thread.markRead(); // belt-and-suspenders so it won't re-match is:unread
}

// ---------------------------------------------------------------------------
// Google Sheets — Registrations
// ---------------------------------------------------------------------------

function _updateRegistrations(emailData, fields, createEvent) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.REGISTRATIONS);
  let calEventId = '';

  if (createEvent && fields.event_date) {
    try { calEventId = _createCalendarEvent(fields); }
    catch (e) { Logger.log('Calendar event failed: ' + e.message); }
  }

  sheet.appendRow([
    new Date(),
    emailData.messageId,
    fields.sender_name  || emailData.senderName,
    emailData.sender,
    fields.event_name   || 'Unknown Event',
    fields.event_date   || 'Not specified',
    fields.participants_count || '1',
    fields.key_ask      || '',
    calEventId,
    'Confirmed',
  ]);

  return calEventId;
}

// ---------------------------------------------------------------------------
// Google Calendar
// ---------------------------------------------------------------------------

function _createCalendarEvent(fields) {
  const calId    = getConfigValue('CALENDAR_ID', 'primary');
  const calendar = CalendarApp.getCalendarById(calId) || CalendarApp.getDefaultCalendar();

  let start = new Date(fields.event_date);
  if (isNaN(start.getTime())) {
    // Fallback: 7 days from now at 10:00
    start = new Date();
    start.setDate(start.getDate() + 7);
    start.setHours(10, 0, 0, 0);
  }

  const end = new Date(start.getTime() + 2 * 60 * 60 * 1000); // +2 hours
  const event = calendar.createEvent(
    fields.event_name || 'Registered Event',
    start,
    end,
    { description: `Auto-created by OpsAgent\nRegistered attendees: ${fields.participants_count || '1'}` }
  );
  return event.getId();
}

// ---------------------------------------------------------------------------
// Google Chat
// ---------------------------------------------------------------------------

function _sendChatAlert(emailData, geminiResult, policyDecision) {
  const webhookUrl = getConfigValue('CHAT_WEBHOOK_URL', '');
  if (!webhookUrl) return;

  const isEscalation  = policyDecision.action === 'human_review';
  const confidencePct = (geminiResult.confidence * 100).toFixed(0) + '%';
  const reasonText    = policyDecision.reasons.length > 0
    ? policyDecision.reasons.join('; ')
    : 'Business opportunity — heads up';

  // Google Chat card format (v1 cards API — works with simple webhook)
  const payload = {
    cards: [{
      header: {
        title:    isEscalation ? '🚨 Human Review Required' : `📬 ${_titleCase(geminiResult.category)} Alert`,
        subtitle: emailData.subject,
        imageUrl: 'https://www.gstatic.com/images/branding/product/1x/gmail_2020q4_48dp.png',
      },
      sections: [{
        widgets: [
          { keyValue: { topLabel: 'From',       content: emailData.sender } },
          { keyValue: { topLabel: 'Category',   content: geminiResult.category } },
          { keyValue: { topLabel: 'Confidence', content: confidencePct } },
          { keyValue: { topLabel: isEscalation ? 'Escalation reason' : 'Reason', content: reasonText } },
        ],
      }],
    }],
  };

  UrlFetchApp.fetch(webhookUrl, {
    method:      'post',
    contentType: 'application/json',
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

// ---------------------------------------------------------------------------
// Google Sheets — Audit Log
// ---------------------------------------------------------------------------

function _logAudit(emailData, geminiResult, policyDecision, results, requiredActions) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);

  const taken = [];
  if (results.reply_sent)           taken.push('auto_reply');
  if (results.info_requested)       taken.push('requested_info');
  if (results.calendar_event_id)    taken.push('calendar_event');
  if (results.form_sent)            taken.push('form_link');
  if (results.chat_alert_sent)      taken.push('chat_alert');
  if (results.human_review_logged)  taken.push('human_review_queue');

  sheet.appendRow([
    new Date(),
    emailData.messageId,
    emailData.threadId,
    emailData.sender,
    emailData.subject,
    geminiResult.category,
    geminiResult.confidence,
    taken.join(', ') || 'none',
    policyDecision.action,
    geminiResult.draft_reply.substring(0, 500),
    policyDecision.sensitiveFlags.join(', '),
    JSON.stringify(geminiResult.extracted_fields).substring(0, 500),
    results.errors.join('; '),
    geminiResult.reasoning || '',   // Reasoning (col 14) — transparency
  ]);
}

// ---------------------------------------------------------------------------
// Google Sheets — Human Review queue
// ---------------------------------------------------------------------------

function _logHumanReview(emailData, geminiResult, policyDecision) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HUMAN_REVIEW);

  sheet.appendRow([
    new Date(),
    emailData.messageId,
    emailData.threadId,
    emailData.sender,
    emailData.subject,
    geminiResult.category,
    geminiResult.confidence,
    geminiResult.draft_reply,
    policyDecision.reasons.join('; '),
    '',          // HumanOverrideCategory — filled by reviewer
    '',          // ReviewedBy
    '',          // ReviewTimestamp
    'Pending',   // Status
    '',          // Notes
    geminiResult.reasoning || '',   // AgentReasoning (col 15) — why it chose this
  ]);
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function _titleCase(str) {
  return str.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}
