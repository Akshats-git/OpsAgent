// Setup.gs — One-time initialisation. Run setup() from the Apps Script editor
// before using the agent. Safe to re-run; it skips sheets that already exist.

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('OpsAgent: Starting setup…');

  _setupSheets(ss);
  _setupGmailLabel();
  _setupTimeTrigger();
  _setupOnEditTrigger();
  _setupDigestTrigger();

  Logger.log('OpsAgent: Setup complete.');
  _safeAlert(
    'Setup complete!\n\n' +
    'One remaining step:\n' +
    'Extensions → Apps Script → Project Settings → Script Properties\n' +
    'Add a property named  GEMINI_API_KEY  with your Gemini API key.'
  );
}

/**
 * Shows a UI alert when one is available, otherwise logs. getUi() throws when
 * a function is run from the editor/trigger context (no bound UI), so anything
 * runnable both ways must go through this.
 */
function _safeAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log('[alert] ' + msg);
  }
}

/** Removes the time-driven trigger to pause the agent without deleting data. */
function pauseAgent() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === TRIGGER_FUNCTION) ScriptApp.deleteTrigger(t);
  });
  Logger.log('Agent paused — time trigger removed.');
}

/** Re-installs the time trigger after a pause. */
function resumeAgent() {
  _setupTimeTrigger();
  Logger.log('Agent resumed.');
}

// ---------------------------------------------------------------------------
// Sheet scaffolding
// ---------------------------------------------------------------------------

function _setupSheets(ss) {
  _makeSheet(ss, SHEET_NAMES.CONFIG, [
    ['Key', 'Value', 'Description'],
    ['ORG_NAME',             'My Organization',     'Name shown in email signatures and alerts'],
    ['CONFIDENCE_THRESHOLD', '0.75',                'Min confidence (0–1) for auto-action. Below → Human Review'],
    ['BATCH_SIZE',           '10',                  'Max emails processed per trigger run'],
    ['AI_PROVIDER',          'gemini',              'AI provider: "gemini" or "openai"'],
    ['GEMINI_MODEL',         'gemini-2.0-flash',    'Gemini model ID'],
    ['OPENAI_MODEL',         'gpt-4o-mini',         'OpenAI model ID (used when AI_PROVIDER = openai)'],
    ['CHAT_WEBHOOK_URL',     '',                    'Google Chat space webhook URL for escalation alerts'],
    ['CALENDAR_ID',          'primary',             'Google Calendar ID for event creation'],
    ['FORMS_VOLUNTEER_URL',  '',                    'Google Form URL sent to volunteer applicants'],
    ['FORMS_SPEAKER_URL',    '',                    'Google Form URL sent to speaker applicants'],
    ['SENSITIVE_KEYWORDS',   'legal,harassment,urgent,refund,lawsuit,discrimination,threat', 'Comma-separated — always route to Human Review'],
    ['MAX_MEMORY_EXAMPLES',  '5',                   'Max few-shot corrections injected per Gemini call'],
    ['AUTO_REPLY_ENABLED',   'true',                'false = dry-run mode (no emails sent)'],
    ['CALENDAR_ENABLED',     'true',                'false = skip Calendar event creation'],
    ['CHAT_ALERTS_ENABLED',  'true',                'false = skip Google Chat alerts'],
    ['DIGEST_ENABLED',       'true',                'false = skip the weekly Google Docs digest'],
    ['REPORTS_FOLDER_NAME',  'OpsAgent Reports',    'Google Drive folder where weekly digests are saved'],
  ]);

  _makeSheet(ss, SHEET_NAMES.AUDIT_LOG, [[
    'Timestamp','EmailID','ThreadID','Sender','Subject',
    'Category','Confidence','ActionsTaken','Status',
    'DraftReply','SensitiveFlags','ExtractedFields','Error','Reasoning',
  ]]);

  _makeSheet(ss, SHEET_NAMES.HUMAN_REVIEW, [[
    'Timestamp','EmailID','ThreadID','Sender','Subject',
    'AgentCategory','Confidence','DraftReply','EscalationReason',
    'HumanOverrideCategory','ReviewedBy','ReviewTimestamp','Status','Notes','AgentReasoning',
  ]]);

  _makeSheet(ss, SHEET_NAMES.MEMORY, [[
    'Timestamp','EmailID','Subject','EmailSnippet',
    'OriginalCategory','CorrectedCategory','AddedBy','Notes',
  ]]);

  _makeSheet(ss, SHEET_NAMES.REGISTRATIONS, [[
    'Timestamp','EmailID','Name','SenderEmail','EventName',
    'EventDate','Participants','AdditionalInfo','CalendarEventID','ReplyStatus',
  ]]);

  _makeSheet(ss, SHEET_NAMES.TEMPLATES, [
    ['Category','Tone','KeyPoints','SampleOpening','SampleClosing','IncludeFormLink','FormLinkKey'],
    ['sponsorship',         'Professional and enthusiastic',
     'Mention audience size, sponsorship tiers, 3-day SLA',
     'Thank you for your interest in sponsoring [ORG_NAME]!',
     'Our team will follow up within 3 business days.',
     'FALSE', ''],
    ['event_registration',  'Warm and welcoming',
     'Confirm registration, provide next steps, mention calendar invite',
     'Your registration is confirmed — we are thrilled to have you!',
     'See you there! Reach out anytime if you have questions.',
     'FALSE', ''],
    ['speaker_application', 'Respectful and encouraging',
     'Acknowledge application, 5-day review timeline, request abstract via form',
     'Thank you for your interest in speaking at an [ORG_NAME] event!',
     'We review applications on a rolling basis and will respond within 5 days.',
     'TRUE', 'FORMS_SPEAKER_URL'],
    ['volunteer_application','Energetic and welcoming',
     'Thank for enthusiasm, send volunteer form, mention onboarding process',
     'We love your enthusiasm to volunteer with [ORG_NAME]!',
     'Fill out the form and our team lead will be in touch shortly.',
     'TRUE', 'FORMS_VOLUNTEER_URL'],
    ['complaint',           'Empathetic and reassuring',
     'Acknowledge concern, DO NOT resolve, assure human follow-up in 24 h',
     'We sincerely apologise for the experience you had.',
     'A member of our team will personally follow up with you within 24 hours.',
     'FALSE', ''],
    ['partnership',         'Professional and collaborative',
     'Express interest, mention org mission, 5-day response SLA',
     'Thank you for reaching out about a potential partnership with [ORG_NAME]!',
     'We will review your proposal and respond within 5 business days.',
     'FALSE', ''],
    ['feedback',            'Appreciative and humble',
     'Thank sincerely, mention it will be shared with the team',
     'Thank you so much for taking the time to share your feedback!',
     'Your input genuinely helps us improve — we really appreciate it.',
     'FALSE', ''],
    ['community_query',     'Helpful and welcoming',
     'Answer the question, point to relevant resources or links',
     'Thanks for reaching out to the [ORG_NAME] community!',
     'Hope that helps — feel free to ask if anything is unclear.',
     'FALSE', ''],
    ['escalation',          'Serious and reassuring',
     'Acknowledge urgency, DO NOT resolve, assure immediate human escalation',
     'Thank you for flagging this — we are treating it with priority.',
     'This has been escalated to our team and someone will respond very shortly.',
     'FALSE', ''],
    ['internal_ops',        'Brief and collegial',
     'Acknowledge, confirm it will be routed to the relevant organisers',
     'Thanks — noted and routing this to the team.',
     'Will follow up internally and circle back.',
     'FALSE', ''],
    ['general_support',     'Helpful and friendly',
     'Answer directly and concisely, offer follow-up if needed',
     'Thank you for reaching out to [ORG_NAME]!',
     'Don\'t hesitate to get in touch if you need anything else.',
     'FALSE', ''],
  ]);

  _makeSheet(ss, SHEET_NAMES.ROUTING, _routingDefaults());

  Logger.log('Sheets ready.');
}

/** Default category → team routing table. Shared by setup and migration. */
function _routingDefaults() {
  return [
    ['Category','Team','NotifyEmail'],
    ['sponsorship',          'Sponsorship',  ''],
    ['partnership',          'Partnerships', ''],
    ['event_registration',   'Events',       ''],
    ['speaker_application',  'Events',       ''],
    ['volunteer_application','Volunteers',   ''],
    ['complaint',            'Support',      ''],
    ['escalation',           'Leadership',   ''],
    ['feedback',             'Support',      ''],
    ['community_query',      'Support',      ''],
    ['internal_ops',         'Operations',   ''],
    ['general_support',      'Support',      ''],
  ];
}

/**
 * Creates a sheet with styled headers and default data rows.
 * Skips creation if the sheet already exists to avoid overwriting live data.
 */
function _makeSheet(ss, name, rows) {
  if (ss.getSheetByName(name)) {
    Logger.log('Sheet already exists, skipping: ' + name);
    return;
  }

  const sheet = ss.insertSheet(name);

  // Style header row
  const headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setValues([rows[0]]);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  sheet.setFrozenRows(1);

  // Write default data rows (if any)
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, rows[0].length).setValues(rows.slice(1));
  }

  // Auto-resize first column for readability
  sheet.autoResizeColumn(1);
  Logger.log('Created sheet: ' + name);
}

// ---------------------------------------------------------------------------
// Gmail label
// ---------------------------------------------------------------------------

function _setupGmailLabel() {
  [GMAIL_LABEL, AWAITING_INFO_LABEL, 'OpsAgent/Quarantine'].forEach(name => {
    const existing = GmailApp.getUserLabels().some(l => l.getName() === name);
    if (!existing) {
      GmailApp.createLabel(name);
      Logger.log('Created Gmail label: ' + name);
    } else {
      Logger.log('Gmail label already exists: ' + name);
    }
  });
}

// ---------------------------------------------------------------------------
// Triggers
// ---------------------------------------------------------------------------

function _setupTimeTrigger() {
  // Remove duplicates first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === TRIGGER_FUNCTION) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(TRIGGER_FUNCTION)
    .timeBased()
    .everyMinutes(TRIGGER_INTERVAL_M)
    .create();
  Logger.log('Time trigger set: ' + TRIGGER_FUNCTION + ' every ' + TRIGGER_INTERVAL_M + ' min.');
}

function _setupOnEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  Logger.log('onEdit trigger installed for memory loop.');
}

function _setupDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === DIGEST_FUNCTION) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(DIGEST_FUNCTION)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log('Weekly digest trigger set: ' + DIGEST_FUNCTION + ' Mondays 8am.');
}

// ---------------------------------------------------------------------------
// Migration helper — run once after pulling Phase B if you already ran setup()
// before. Adds new Config rows, new sheet columns, the new Gmail labels, and
// the weekly digest trigger WITHOUT recreating or wiping any existing sheet.
// ---------------------------------------------------------------------------

function devUpgradeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. New Config rows (only added if the key is absent)
  _ensureConfigRows(ss, [
    ['AI_PROVIDER',         'gemini',            'AI provider: "gemini" or "openai"'],
    ['OPENAI_MODEL',        'gpt-4o-mini',       'OpenAI model ID (used when AI_PROVIDER = openai)'],
    ['DIGEST_ENABLED',      'true',              'false = skip the weekly Google Docs digest'],
    ['REPORTS_FOLDER_NAME', 'OpsAgent Reports',  'Google Drive folder where weekly digests are saved'],
  ]);

  // 2. New trailing columns
  _ensureColumn(ss, SHEET_NAMES.AUDIT_LOG,    'Reasoning');
  _ensureColumn(ss, SHEET_NAMES.HUMAN_REVIEW, 'AgentReasoning');

  // 3. Routing sheet (Phase C) — created only if absent
  if (!ss.getSheetByName(SHEET_NAMES.ROUTING)) {
    _makeSheet(ss, SHEET_NAMES.ROUTING, _routingDefaults());
  }

  // 4. New-category template rows (Phase C) — appended if absent
  _ensureTemplateRows(ss);

  // 5. New labels + digest trigger
  _setupGmailLabel();
  _setupDigestTrigger();

  clearConfigCache();
  Logger.log('devUpgradeSheets: migration complete.');
  _safeAlert('Upgrade complete — Config rows, columns, labels, and the weekly digest trigger are now in place.');
}

/** Appends any Config keys that aren't already present. */
function _ensureConfigRows(ss, rows) {
  const sheet    = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const existing = sheet.getDataRange().getValues().map(r => String(r[0]).trim());
  rows.forEach(row => {
    if (existing.indexOf(row[0]) === -1) {
      sheet.appendRow(row);
      Logger.log('Config: added ' + row[0]);
    }
  });
}

/** Appends template rows for the Phase C categories if they're missing. */
function _ensureTemplateRows(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.TEMPLATES);
  if (!sheet) return;
  const existing = sheet.getDataRange().getValues().map(r => String(r[0]).trim());
  const newRows = [
    ['community_query', 'Helpful and welcoming',
     'Answer the question, point to relevant resources or links',
     'Thanks for reaching out to the [ORG_NAME] community!',
     'Hope that helps — feel free to ask if anything is unclear.', 'FALSE', ''],
    ['escalation', 'Serious and reassuring',
     'Acknowledge urgency, DO NOT resolve, assure immediate human escalation',
     'Thank you for flagging this — we are treating it with priority.',
     'This has been escalated to our team and someone will respond very shortly.', 'FALSE', ''],
    ['internal_ops', 'Brief and collegial',
     'Acknowledge, confirm it will be routed to the relevant organisers',
     'Thanks — noted and routing this to the team.',
     'Will follow up internally and circle back.', 'FALSE', ''],
  ];
  newRows.forEach(row => {
    if (existing.indexOf(row[0]) === -1) {
      sheet.appendRow(row);
      Logger.log('Templates: added ' + row[0]);
    }
  });
}

/** Adds a trailing header column to a sheet if that header isn't present. */
function _ensureColumn(ss, sheetName, header) {
  const sheet  = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  if (headers.indexOf(header) !== -1) return;
  const col = sheet.getLastColumn() + 1;
  const cell = sheet.getRange(1, col);
  cell.setValue(header).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  Logger.log(`${sheetName}: added column "${header}" at ${col}`);
}
