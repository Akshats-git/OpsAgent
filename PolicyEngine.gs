// PolicyEngine.gs — Applies organisation-defined rules to decide whether an
// email should be handled automatically or routed to the Human Review queue.
// All thresholds and keyword lists are read from the Config sheet so admins
// can adjust them without touching code.

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Evaluates a classified email against org policy.
 *
 * @param {Object} emailData    - Raw email fields (subject, body, sender, …)
 * @param {Object} geminiResult - Output of classifyAndExtract()
 * @returns {Object} decision   - { action: 'auto'|'human_review', reasons: [], sensitiveFlags: [] }
 */
function evaluatePolicy(emailData, geminiResult) {
  const threshold        = parseFloat(getConfigValue('CONFIDENCE_THRESHOLD', '0.75'));
  const sensitiveKeywords = _loadSensitiveKeywords();

  const decision = {
    action:         'auto',
    reasons:        [],
    sensitiveFlags: Array.isArray(geminiResult.sensitive_flags) ? [...geminiResult.sensitive_flags] : [],
  };

  // ── Rule 1: Complaints & escalations ALWAYS go to human review ──────────
  if (geminiResult.category === 'complaint' || geminiResult.category === 'escalation') {
    decision.action = 'human_review';
    decision.reasons.push(`Category is "${geminiResult.category}" — always requires human review`);
    return decision; // No further checks needed
  }

  // ── Rule 2: Sensitive keywords detected in subject or body ───────────────
  const emailText      = `${emailData.subject} ${emailData.body}`.toLowerCase();
  const hitKeywords    = sensitiveKeywords.filter(kw => emailText.includes(kw));
  if (hitKeywords.length > 0) {
    decision.action = 'human_review';
    decision.reasons.push(`Sensitive keyword(s) found: ${hitKeywords.join(', ')}`);
    // Add to sensitiveFlags if not already present
    hitKeywords.forEach(kw => {
      if (!decision.sensitiveFlags.includes(kw)) decision.sensitiveFlags.push(kw);
    });
    return decision;
  }

  // ── Rule 3: Gemini itself flagged sensitive content ──────────────────────
  if (geminiResult.sensitive_flags && geminiResult.sensitive_flags.length > 0) {
    const geminiHit = geminiResult.sensitive_flags.some(flag =>
      sensitiveKeywords.some(kw => flag.toLowerCase().includes(kw))
    );
    if (geminiHit) {
      decision.action = 'human_review';
      decision.reasons.push(
        `Gemini flagged sensitive content: ${geminiResult.sensitive_flags.join(', ')}`
      );
      return decision;
    }
  }

  // ── Rule 4: Confidence below threshold ──────────────────────────────────
  if (geminiResult.confidence < threshold) {
    decision.action = 'human_review';
    decision.reasons.push(
      `Confidence ${(geminiResult.confidence * 100).toFixed(0)}% < ` +
      `threshold ${(threshold * 100).toFixed(0)}%`
    );
    return decision;
  }

  // ── Rule 5: Required fields missing → request more info ──────────────────
  // The agent is confident about the category but the email lacks details it
  // needs to actually complete the action (e.g. a registration with no date).
  const missing = getMissingFields(geminiResult.category, geminiResult.extracted_fields);
  if (missing.length > 0) {
    decision.action       = 'request_info';
    decision.missingFields = missing;
    decision.reasons.push(`Missing required detail(s): ${missing.join(', ')}`);
    return decision;
  }

  // All rules passed — proceed automatically
  return decision;
}

/**
 * Required fields per category. If any are blank in the extracted data, the
 * agent asks the sender to provide them instead of completing the action.
 * Only categories listed here are gated; everything else proceeds freely.
 */
const REQUIRED_FIELDS = {
  event_registration: [
    { key: 'event_name', label: 'which event you want to register for' },
    { key: 'event_date', label: 'the event date you are interested in' },
  ],
};

/** Returns an array of human-readable labels for fields that are missing. */
function getMissingFields(category, fields) {
  const required = REQUIRED_FIELDS[category];
  if (!required) return [];
  fields = fields || {};
  return required
    .filter(f => {
      const v = fields[f.key];
      return !v || String(v).trim() === '' ||
             String(v).trim().toLowerCase() === 'not specified' ||
             String(v).trim().toLowerCase() === 'unknown';
    })
    .map(f => f.label);
}

/**
 * Returns the set of action types needed for a given category.
 * ActionRouter uses this list to decide what to execute.
 */
function getRequiredActions(category) {
  // Every email always gets a reply and an audit-log entry
  const actions = new Set(['reply', 'audit_log']);

  switch (category) {
    case 'event_registration':
      actions.add('registrations_sheet');
      if (getConfigValue('CALENDAR_ENABLED', 'true') === 'true') actions.add('calendar_event');
      break;

    case 'volunteer_application':
      actions.add('send_form');
      break;

    case 'speaker_application':
      actions.add('send_form');
      break;

    case 'complaint':
    case 'escalation':
      actions.add('chat_alert');
      break;

    case 'internal_ops':
      // Internal logistics — always notify the team channel
      actions.add('chat_alert');
      break;

    case 'sponsorship':
    case 'partnership':
      // Notify the team of inbound business opportunities
      if (getConfigValue('CHAT_ALERTS_ENABLED', 'true') === 'true') {
        actions.add('chat_alert');
      }
      break;
  }

  return [...actions];
}

// ---------------------------------------------------------------------------
// Team routing — maps each category to an owning team. Read from the Routing
// sheet so admins can re-assign teams without touching code.
// ---------------------------------------------------------------------------

/**
 * Returns { team, notifyEmail } for a category, or null if no routing exists.
 */
function getRoute(category) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ROUTING);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === category) {
      return { team: String(rows[i][1]).trim(), notifyEmail: String(rows[i][2]).trim() };
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

function _loadSensitiveKeywords() {
  const raw = getConfigValue(
    'SENSITIVE_KEYWORDS',
    'legal,harassment,urgent,refund,lawsuit,discrimination,threat'
  );
  return raw.split(',').map(k => k.trim().toLowerCase()).filter(Boolean);
}
