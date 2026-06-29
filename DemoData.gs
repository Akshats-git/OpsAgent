// DemoData.gs — Populates the sheets with realistic sample data so the
// dashboard looks alive for screenshots, judging, and demos WITHOUT waiting
// for real email traffic. Run seedDemoData() from the editor; clearDemoData()
// wipes every data row (keeping headers) when you're done.
//
// NOTE: For a *live* pipeline demo, send real emails from a DIFFERENT account
// (the agent skips mail from its own address). This helper is purely for
// pre-populating the dashboards/logs.

function seedDemoData() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = Date.now();
  const day = 24 * 60 * 60 * 1000;
  const ago = (d, h) => new Date(now - d * day - (h || 0) * 60 * 60 * 1000);

  // ── AuditLog ────────────────────────────────────────────────────────────
  const audit = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  const A = [
    [ago(6,2), 'Sarah from TechCorp', 'We would love to sponsor your hackathon', 'sponsorship',     0.94, 'auto_reply, chat_alert',        'auto',         'Mentions a sponsorship budget and brand visibility.'],
    [ago(6,1), 'raj.patel@gmail.com', 'Register me for the AI workshop on July 12', 'event_registration', 0.91, 'auto_reply, calendar_event', 'auto',     'Explicit registration request with a clear event and date.'],
    [ago(5,3), 'anonymous@proton.me', 'The event was poorly organised and rude staff', 'complaint',  0.88, 'chat_alert, human_review_queue','human_review', 'Strong negative sentiment about staff conduct.'],
    [ago(5,0), 'lena@designhub.io',  'Interested in a long-term partnership',        'partnership',  0.86, 'auto_reply, chat_alert',        'auto',         'Proposes ongoing collaboration, not a one-off ask.'],
    [ago(4,5), 'dr.mehta@univ.edu',  'I would like to speak at your conference',     'speaker_application', 0.90, 'auto_reply, form_link', 'auto',         'Offers to present; requesting a speaking slot.'],
    [ago(4,2), 'volunteer.jay@gmail.com', 'How can I help out as a volunteer?',      'volunteer_application', 0.89, 'auto_reply, form_link','auto',       'Expresses willingness to volunteer.'],
    [ago(3,6), 'kid@school.com',     'Can I register? (no date given)',              'event_registration', 0.79, 'requested_info',          'request_info', 'Registration intent but the event/date is missing.'],
    [ago(3,1), 'maria@ngo.org',      'Loved the workshop, thank you!',               'feedback',     0.95, 'auto_reply',                    'auto',         'Positive feedback expressing gratitude.'],
    [ago(2,4), 'press@dailynews.com','URGENT: legal notice regarding your event',    'escalation',   0.92, 'chat_alert, human_review_queue','human_review', 'Contains legal/urgent language requiring escalation.'],
    [ago(2,0), 'student@campus.edu', 'What time does the venue open?',               'community_query', 0.93, 'auto_reply',                 'auto',         'General logistical question from a community member.'],
    [ago(1,3), 'organiser@team.com', 'Reminder: AV setup at 9am tomorrow',           'internal_ops', 0.90, 'chat_alert',                    'auto',         'Internal logistics message between organisers.'],
    [ago(1,1), 'curious@gmail.com',  'Is the event free to attend?',                 'community_query', 0.94, 'auto_reply',                 'auto',         'Simple eligibility/cost question.'],
    [ago(0,4), 'bigco@enterprise.com','Sponsorship deck attached — lets talk',       'sponsorship',  0.88, 'auto_reply, chat_alert',        'auto',         'Inbound sponsorship with materials attached.'],
    [ago(0,2), 'angry@customer.com', 'I demand a refund immediately',                'complaint',    0.85, 'chat_alert, human_review_queue','human_review', 'Refund demand — sensitive keyword triggered review.'],
    [ago(0,1), 'newbie@gmail.com',   'Thanks for the quick help earlier!',           'feedback',     0.96, 'auto_reply',                    'auto',         'Appreciative note following prior support.'],
  ];
  A.forEach(r => audit.appendRow([
    r[0], 'demo-' + Utilities.getUuid().substring(0, 8), 'thread-demo', r[1], r[2],
    r[3], r[4], r[5], r[6], 'Sample drafted reply for "' + r[2] + '"…',
    (r[3] === 'complaint' || r[3] === 'escalation') ? 'refund/legal' : '',
    '{"sender_name":"' + r[1] + '"}', '', r[7],
  ]));

  // ── HumanReview (pending) ───────────────────────────────────────────────
  const hr = ss.getSheetByName(SHEET_NAMES.HUMAN_REVIEW);
  [
    [ago(2,4), 'press@dailynews.com', 'URGENT: legal notice regarding your event', 'escalation', 0.92,
     'Thank you for flagging this — we are treating it with priority and have escalated it to our leadership team.',
     'Category is "escalation" — always requires human review', 'Contains legal/urgent language requiring escalation.'],
    [ago(0,2), 'angry@customer.com', 'I demand a refund immediately', 'complaint', 0.85,
     'We sincerely apologise for the experience. A team member will personally follow up within 24 hours.',
     'Sensitive keyword(s) found: refund', 'Refund demand — sensitive keyword triggered review.'],
  ].forEach(r => hr.appendRow([
    r[0], 'demo-' + Utilities.getUuid().substring(0, 8), 'thread-demo', r[1], r[2],
    r[3], r[4], r[5], r[6], '', '', '', 'Pending', '', r[7],
  ]));

  // ── Memory (corrections) ────────────────────────────────────────────────
  const mem = ss.getSheetByName(SHEET_NAMES.MEMORY);
  [
    [ago(5,0), 'Collab on a workshop series', 'We run workshops and want to co-host with you…', 'sponsorship', 'partnership'],
    [ago(3,0), 'Quick question about parking', 'Where can attendees park near the venue?',       'general_support', 'community_query'],
    [ago(1,0), 'Following up on AV logistics', 'Internal note: confirm mic count for main hall',  'general_support', 'internal_ops'],
  ].forEach(r => mem.appendRow([
    r[0], 'demo-' + Utilities.getUuid().substring(0, 8), r[1], r[2], r[3], r[4], 'demo.reviewer@org.com', '',
  ]));

  // ── Registrations ───────────────────────────────────────────────────────
  const reg = ss.getSheetByName(SHEET_NAMES.REGISTRATIONS);
  [
    [ago(6,1), 'Raj Patel',   'raj.patel@gmail.com', 'AI Workshop',     'July 12, 2026', '1'],
    [ago(4,0), 'Priya Sharma','priya@gmail.com',     'Hackathon 2026',  'July 20, 2026', '3'],
    [ago(2,0), 'Sam Lee',     'sam.lee@gmail.com',   'Design Sprint',   'July 15, 2026', '2'],
  ].forEach(r => reg.appendRow([
    r[0], 'demo-' + Utilities.getUuid().substring(0, 8), r[1], r[2], r[3], r[4], r[5], '', 'demo-cal-id', 'Confirmed',
  ]));

  clearConfigCache();
  Logger.log('seedDemoData: inserted ' + A.length + ' audit rows + review/memory/registration samples.');
  _safeAlert('Demo data seeded! Open the dashboard to see it populated.');
}

/** Removes every data row from the activity sheets (keeps headers + Config/Templates/Routing). */
function clearDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_NAMES.AUDIT_LOG, SHEET_NAMES.HUMAN_REVIEW, SHEET_NAMES.MEMORY, SHEET_NAMES.REGISTRATIONS]
    .forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
    });
  clearConfigCache();
  Logger.log('clearDemoData: activity sheets cleared (headers kept).');
  _safeAlert('Activity sheets cleared.');
}
