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

  Logger.log('OpsAgent: Setup complete.');
  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    'One remaining step:\n' +
    'Extensions → Apps Script → Project Settings → Script Properties\n' +
    'Add a property named  GEMINI_API_KEY  with your Gemini API key.'
  );
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
  ]);

  _makeSheet(ss, SHEET_NAMES.AUDIT_LOG, [[
    'Timestamp','EmailID','ThreadID','Sender','Subject',
    'Category','Confidence','ActionsTaken','Status',
    'DraftReply','SensitiveFlags','ExtractedFields','Error',
  ]]);

  _makeSheet(ss, SHEET_NAMES.HUMAN_REVIEW, [[
    'Timestamp','EmailID','ThreadID','Sender','Subject',
    'AgentCategory','Confidence','DraftReply','EscalationReason',
    'HumanOverrideCategory','ReviewedBy','ReviewTimestamp','Status','Notes',
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
    ['general_support',     'Helpful and friendly',
     'Answer directly and concisely, offer follow-up if needed',
     'Thank you for reaching out to [ORG_NAME]!',
     'Don\'t hesitate to get in touch if you need anything else.',
     'FALSE', ''],
  ]);

  Logger.log('Sheets ready.');
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
  const existing = GmailApp.getUserLabels().some(l => l.getName() === GMAIL_LABEL);
  if (!existing) {
    GmailApp.createLabel(GMAIL_LABEL);
    Logger.log('Created Gmail label: ' + GMAIL_LABEL);
  } else {
    Logger.log('Gmail label already exists: ' + GMAIL_LABEL);
  }
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
