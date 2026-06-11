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

  // ── Rule 1: Complaints ALWAYS go to human review ────────────────────────
  if (geminiResult.category === 'complaint') {
    decision.action = 'human_review';
    decision.reasons.push('Category is "complaint" — always requires human review');
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

  // All rules passed — proceed automatically
  return decision;
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
// Private helpers
// ---------------------------------------------------------------------------

function _loadSensitiveKeywords() {
  const raw = getConfigValue(
    'SENSITIVE_KEYWORDS',
    'legal,harassment,urgent,refund,lawsuit,discrimination,threat'
  );
  return raw.split(',').map(k => k.trim().toLowerCase()).filter(Boolean);
}
