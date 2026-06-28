// MemoryLoop.gs — The feedback flywheel.
// When a human reviewer fills in the "HumanOverrideCategory" column in the
// HumanReview sheet, this onEdit trigger fires, writes the correction to the
// Memory sheet, and updates the row status — so every future Gemini call
// benefits from the correction via few-shot injection.

// Column indices in HumanReview (1-based, matching _logHumanReview row order)
const HR_COL = {
  TIMESTAMP:       1,
  EMAIL_ID:        2,
  THREAD_ID:       3,
  SENDER:          4,
  SUBJECT:         5,
  AGENT_CATEGORY:  6,
  CONFIDENCE:      7,
  DRAFT_REPLY:     8,
  ESCALATION:      9,
  OVERRIDE:        10,   // ← Human fills this column
  REVIEWED_BY:     11,
  REVIEW_TS:       12,
  STATUS:          13,
  NOTES:           14,
};

// ---------------------------------------------------------------------------
// Trigger handler — installed as an installable onEdit trigger by Setup.gs
// ---------------------------------------------------------------------------

function onEditTrigger(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.HUMAN_REVIEW) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Only react to edits in the HumanOverrideCategory column, skip header
  if (col !== HR_COL.OVERRIDE || row <= 1) return;

  const override = String(e.range.getValue()).trim().toLowerCase();
  if (!override) return; // Cell was cleared — ignore

  // Validate the override value is a known category
  if (!CATEGORIES.includes(override)) {
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEET_NAMES.HUMAN_REVIEW)
      .getRange(row, HR_COL.NOTES)
      .setValue(`Invalid category "${override}". Valid: ${CATEGORIES.join(', ')}`);
    return;
  }

  const reviewSheet = sheet;
  const rowData     = reviewSheet.getRange(row, 1, 1, 14).getValues()[0];

  const emailId       = String(rowData[HR_COL.EMAIL_ID - 1]);
  const subject       = String(rowData[HR_COL.SUBJECT - 1]);
  const agentCategory = String(rowData[HR_COL.AGENT_CATEGORY - 1]);
  const reviewer      = Session.getActiveUser().getEmail();
  const now           = new Date();

  // ── Mark the HumanReview row as reviewed ────────────────────────────────
  const statusValue = (agentCategory === override) ? 'Confirmed Correct' : 'Corrected';
  reviewSheet.getRange(row, HR_COL.REVIEWED_BY).setValue(reviewer);
  reviewSheet.getRange(row, HR_COL.REVIEW_TS).setValue(now);
  reviewSheet.getRange(row, HR_COL.STATUS).setValue(statusValue);

  // ── If it's just a confirmation (agent was right), no memory update needed
  if (agentCategory === override) {
    Logger.log(`Memory: "${subject}" confirmed correct as ${agentCategory}.`);
    return;
  }

  // ── Write correction to Memory sheet ─────────────────────────────────────
  recordMemoryCorrection(emailId, subject, agentCategory, override, reviewer);
}

/**
 * Appends a single correction to the Memory sheet so it becomes a future
 * few-shot example. Shared by the onEdit trigger and the dashboard approve flow.
 * No-op if the original and corrected categories are the same.
 */
function recordMemoryCorrection(emailId, subject, originalCategory, correctedCategory, reviewer) {
  if (originalCategory === correctedCategory) return;

  let emailSnippet = '';
  try {
    const message = GmailApp.getMessageById(emailId);
    emailSnippet  = (message.getPlainBody() || '').substring(0, 350);
  } catch (err) {
    emailSnippet = '(email body unavailable)';
    Logger.log('Memory: could not fetch email body: ' + err.message);
  }

  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const memorySheet = ss.getSheetByName(SHEET_NAMES.MEMORY);

  memorySheet.appendRow([
    new Date(),
    emailId,
    subject,
    emailSnippet,
    originalCategory,   // OriginalCategory (what the AI said)
    correctedCategory,  // CorrectedCategory (what the human said)
    reviewer,
    '',                 // Notes — reviewer can fill optionally
  ]);

  // Bust the config cache so the next run picks up the fresh memory count
  clearConfigCache();

  Logger.log(
    `Memory updated: "${subject}" corrected from [${originalCategory}] → [${correctedCategory}] by ${reviewer}`
  );
}
