// ReviewActions.gs — Resolves items in the Human Review queue.
// These functions are the "second half" of the human-in-the-loop: a reviewer
// approves (optionally editing the draft and correcting the category) and the
// agent actually sends the reply. Callable from the dashboard button AND
// directly from the editor as a fallback.

// ---------------------------------------------------------------------------
// Approve & Send
// ---------------------------------------------------------------------------

/**
 * Sends the (optionally edited) reply for a Human Review item and closes it out.
 * If the reviewer changed the category, the correction is also recorded to Memory.
 *
 * @param {string} emailId       - Message ID identifying the HumanReview row
 * @param {string} editedDraft   - Final reply body (may be edited by the reviewer)
 * @param {string} finalCategory - Category the reviewer settled on
 * @returns {Object} { success, message }
 */
function approveAndSend(emailId, editedDraft, finalCategory) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheet  = ss.getSheetByName(SHEET_NAMES.HUMAN_REVIEW);
    const found  = _findReviewRow(sheet, emailId);
    if (!found) return { success: false, message: 'Review item not found (it may already be resolved).' };

    const { row, data } = found;
    if (String(data[HR_COL.STATUS - 1]) !== 'Pending') {
      return { success: false, message: 'This item was already resolved.' };
    }

    const threadId      = String(data[HR_COL.THREAD_ID - 1]);
    const subject       = String(data[HR_COL.SUBJECT - 1]);
    const agentCategory = String(data[HR_COL.AGENT_CATEGORY - 1]);
    const reviewer      = Session.getActiveUser().getEmail() || 'dashboard-user';
    const orgName       = getConfigValue('ORG_NAME', 'Our Organization');
    const draft         = (editedDraft && editedDraft.trim()) ? editedDraft : String(data[HR_COL.DRAFT_REPLY - 1]);
    const category      = (finalCategory && finalCategory.trim()) ? finalCategory.trim() : agentCategory;

    // ── Send the reply (reuses ActionRouter's sender) ───────────────────────
    _sendReply({ threadId: threadId }, draft, orgName);

    // ── Close out the review row ────────────────────────────────────────────
    const now = new Date();
    sheet.getRange(row, HR_COL.OVERRIDE).setValue(category);
    sheet.getRange(row, HR_COL.REVIEWED_BY).setValue(reviewer);
    sheet.getRange(row, HR_COL.REVIEW_TS).setValue(now);
    sheet.getRange(row, HR_COL.STATUS).setValue('Approved & Sent');

    // ── Record correction to Memory if the category was changed ─────────────
    recordMemoryCorrection(emailId, subject, agentCategory, category, reviewer);

    // ── Audit trail ─────────────────────────────────────────────────────────
    _logReviewResolution(emailId, threadId, subject, category, reviewer, 'human_approved_sent');

    return { success: true, message: 'Reply sent and review closed.' };
  } catch (e) {
    return { success: false, message: 'Send failed: ' + e.message };
  }
}

/**
 * Dismisses a Human Review item without sending a reply (e.g. spam, duplicate,
 * or handled out-of-band). Records the correction if the category was changed.
 *
 * @param {string} emailId       - Message ID identifying the HumanReview row
 * @param {string} reason        - Optional dismissal note
 * @returns {Object} { success, message }
 */
function dismissReview(emailId, reason) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.HUMAN_REVIEW);
    const found = _findReviewRow(sheet, emailId);
    if (!found) return { success: false, message: 'Review item not found.' };

    const { row, data } = found;
    const reviewer = Session.getActiveUser().getEmail() || 'dashboard-user';

    sheet.getRange(row, HR_COL.REVIEWED_BY).setValue(reviewer);
    sheet.getRange(row, HR_COL.REVIEW_TS).setValue(new Date());
    sheet.getRange(row, HR_COL.STATUS).setValue('Dismissed');
    if (reason) sheet.getRange(row, HR_COL.NOTES).setValue(reason);

    _logReviewResolution(
      emailId, String(data[HR_COL.THREAD_ID - 1]), String(data[HR_COL.SUBJECT - 1]),
      String(data[HR_COL.AGENT_CATEGORY - 1]), reviewer, 'human_dismissed'
    );

    return { success: true, message: 'Review dismissed.' };
  } catch (e) {
    return { success: false, message: 'Dismiss failed: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Locates a HumanReview row by its EmailID column. Returns {row, data} or null. */
function _findReviewRow(sheet, emailId) {
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][HR_COL.EMAIL_ID - 1]) === String(emailId)) {
      return { row: i + 2, data: values[i] };
    }
  }
  return null;
}

/** Appends a row to the AuditLog recording a human resolution action. */
function _logReviewResolution(emailId, threadId, subject, category, reviewer, actionTag) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  sheet.appendRow([
    new Date(), emailId, threadId, reviewer, subject,
    category, 1, actionTag, 'human_review_resolved',
    '', '', '', '',
  ]);
}
