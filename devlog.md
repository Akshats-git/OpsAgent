# OpsAgent — Dev Log

A running account of how this thing got built, what broke, what worked, and why we made the calls we did.

---

## Session 1 — Architecture, Design Review, and Full First Implementation

**Date:** 2026-06-11

### Where we started

The initial idea was solid: Gmail + Gemini + Apps Script + Sheets + Calendar + Chat, all wired together into a system that can process a student org's email inbox autonomously. The design had a lot of good instincts — Config sheet for non-technical admins, a memory loop for few-shot learning, human-in-the-loop for uncertain cases. But before writing a single line of code, we did a design review pass and caught a handful of things that would have caused demo failures.

### Problems found in the design — and how we fixed them

**Problem 1: Idempotency — the sneaky killer**

The original plan was to query `is:unread` emails every 5 minutes. The problem: if two trigger runs overlap (Apps Script doesn't prevent this), or if an action partially fails halfway through, the same email gets processed twice. You end up double-replying to someone who registered for an event. Not great.

The fix was simple once you see it: apply a Gmail label (`OpsAgent/Processed`) as the very first thing when an email enters the pipeline, and exclude that label from the search query. The label is the lock. Even if the trigger runs twice at the same time, the second run sees the label and skips the thread.

**Problem 2: Apps Script's 6-minute execution limit**

Apps Script kills any function that runs longer than 6 minutes. If the inbox has 30 unread emails and each Gemini call takes 2 seconds (plus sheet reads, Gmail API calls, etc.), you're easily at 4–5 minutes for a batch of 10. If you try to process 30, you get terminated mid-run with no cleanup.

The fix was two-layered: first, a configurable `BATCH_SIZE` in the Config sheet (default 10) so the admin can tune it. Second, a runtime guard in `EmailProcessor.gs` that checks elapsed time every iteration — if we're past 4.5 minutes, we log how many threads remain and bail cleanly. The next trigger run picks up where we left off because of the label system.

**Problem 3: Memory prompt bloat**

The original memory loop design would inject *all* past corrections into every Gemini prompt. That's fine at 5 corrections. At 200 corrections, you're hitting context window limits and your prompt tokens cost skyrockets. More importantly, irrelevant corrections (a complaint being corrected to partnership) add noise when you're classifying a registration email.

The fix: `getMemoryExamples()` in `GeminiService.gs` pulls the most recent N corrections, prioritises ones where the *corrected* category matches what we're currently classifying, then fills remaining slots with others. The cap (default 5) lives in the Config sheet. The few-shot block is prefixed with clear explanatory text so Gemini understands these are *mistakes to learn from*, not positive examples to mimic.

**Problem 4: Self-reply loops**

The agent sends emails from the org's Gmail account. Those sent messages can appear in the inbox search results (especially if someone replies to the agent's auto-reply and cc's the same address). If the agent processes its own sent email, you get infinite loops.

The fix: in `EmailProcessor.gs`, we grab the sender's email from the latest message in each thread and compare it against `Session.getActiveUser().getEmail()`. If it's a match, we apply the processed label and continue — no Gemini call, no reply, no log entry.

### What we built

The codebase is split across 8 files. Here's what each one does and why we structured it that way:

**`appsscript.json`** — The manifest. Declares all OAuth scopes (Gmail modify, Sheets, Calendar, Drive, external requests). The `webapp` section sets `executeAs: USER_DEPLOYING` so the dashboard runs as the org admin, and `access: ANYONE_ANONYMOUS` so team members don't need a Google login to view stats.

**`Config.gs`** — One place to rule all configuration. Reads the Config sheet and caches the result in `CacheService` for 5 minutes (matching the trigger interval). This means the agent isn't doing a sheet read on every single email. `getGeminiApiKey()` intentionally reads from `PropertiesService` instead of the sheet — API keys should never appear in spreadsheet cells.

**`Setup.gs`** — The one-time initialisation function. Creates all 6 sheet tabs with styled headers and default data, creates the `OpsAgent/Processed` Gmail label, installs the 5-minute time trigger, and installs the onEdit trigger for the memory loop. Safe to re-run: it skips sheets that already exist, so you can't accidentally wipe live data. Also exposes `pauseAgent()` and `resumeAgent()` for operators who want to temporarily stop the agent without touching triggers manually.

**`GeminiService.gs`** — Everything Gemini. The prompt is built in `_buildPrompt()`, which assembles: the category list, per-category response guidelines, template tone/key-point guidance, the few-shot correction block (if any corrections exist), and the email itself. We set `temperature: 0.15` — low enough for consistent classification, high enough that the replies don't sound robotic. We also set `responseMimeType: 'application/json'` to tell Gemini to return JSON natively, which is more reliable than just asking in the prompt. Safety settings are all set to `BLOCK_NONE` because complaint emails naturally contain words that would otherwise get blocked.

**`PolicyEngine.gs`** — Four rules, applied in order:
1. Complaints always escalate — no exceptions, no matter the confidence score.
2. Sensitive keywords in subject or body → human review. The keyword list lives in the Config sheet.
3. If Gemini itself flagged sensitive content → human review (belt-and-suspenders).
4. Confidence below threshold → human review.

If all four pass, the email gets automated. `getRequiredActions()` then resolves which specific actions to take based on category (e.g., event_registration also needs a Registrations sheet update and possibly a Calendar event).

**`ActionRouter.gs`** — Executes the actions. Every email gets a reply and an audit log entry. Specific categories get additional actions: registrations update the Registrations sheet, volunteer/speaker apps get a form link reply, complaint/sponsorship/partnership emails send a Google Chat webhook. The Chat payload uses the Card format (structured messages with headers and key-value widgets) rather than plain text, so it renders nicely in the Chat UI. Calendar event creation parses the extracted `event_date` field, falls back to 7 days from now at 10:00 if parsing fails, and creates a 2-hour block.

**`EmailProcessor.gs`** — The main loop. Fetches threads matching the Gmail query, checks the execution timer each iteration, detects self-replies, assembles the email data object, and calls the 4-step pipeline: Gemini → Policy → RequiredActions → Execute. There's a 600ms sleep between Gemini calls to stay within the free-tier rate limits (15 RPM for Gemini 1.5 Flash).

**`MemoryLoop.gs`** — The feedback flywheel. An installable onEdit trigger watches the HumanReview sheet. When a reviewer fills in column J (HumanOverrideCategory), the trigger fires, validates the value, fetches the original email body (truncated to 350 chars for the Memory sheet), and appends a correction row. It also updates the HumanReview row status to either "Corrected" or "Confirmed Correct" depending on whether the override matches the agent's original classification. The next time `getMemoryExamples()` runs, this correction is in the pool.

**`Dashboard.gs`** — The web-app backend. `doGet()` serves the HTML file. The other functions (`getDashboardStats`, `getRecentAuditLog`, `getHumanReviewQueue`, `getMemoryEntries`, `triggerManualRun`) are called from the frontend via `google.script.run`. We read only the exact columns we need from each sheet to keep response times fast.

**`dashboard.html`** — The frontend. Built with Tailwind CSS (CDN), Alpine.js (reactive state, no build step), and Chart.js (donut chart for category distribution). The design is dark-themed (slate-950 background) with gradient stat cards, animated confidence bars, color-coded category badges, and 30-second auto-refresh. The "Run Now" button calls `triggerManualRun()` server-side and shows a toast on completion. The Human Review tab shows pending items with an expandable draft reply section and clear instructions for correcting via the sheet. The Memory tab shows corrections with the original (struck-through) and corrected category side by side.

### What we added beyond the original design

- **Templates sheet** — per-category tone and key-point guidance injected into every Gemini prompt. Admins can edit tone without touching the prompt code.
- **`pauseAgent()` / `resumeAgent()`** convenience functions — one-click pause from the Apps Script editor.
- **Dashboard "Run Now" button** — critical for demos where you don't want to wait up to 5 minutes for the trigger to fire.
- **30-second auto-refresh** with a live/paused toggle in the dashboard header.
- **Error emails are labeled** even when logging fails, preventing infinite retry on a broken email.
- **Rate limiting sleep (600ms)** between Gemini calls to avoid hitting the free-tier rate limit.

### What the demo sequence looks like

1. Send a registration email to the inbox → within one trigger cycle (or via "Run Now"): agent auto-replies with confirmation, adds row to Registrations sheet, creates Calendar event.
2. Send an ambiguous email → goes to HumanReview sheet with status "Pending", Chat alert fires.
3. Fill in column J of the HumanReview row with the correct category → Memory sheet gets a new row instantly.
4. Send a similar ambiguous email → agent now gets it right (or gets significantly higher confidence) because the correction was injected as a few-shot example.

### Things to watch out for during deployment

- The `GEMINI_API_KEY` must be set in Script Properties *before* running `setup()`. The setup function reminds you, but if you skip it, every trigger run will throw and log errors.
- The Chat webhook URL is optional — if left blank, Chat alerts are silently skipped.
- `AUTO_REPLY_ENABLED = false` in the Config sheet puts the agent in dry-run mode (classifies and logs, but sends no emails). Useful for the first few days of testing.
- The `FORMS_VOLUNTEER_URL` and `FORMS_SPEAKER_URL` fields are also optional — if blank, the form-link step is silently skipped.
- The web app needs to be deployed as "Execute as: Me" and "Access: Anyone" from the Apps Script → Deploy menu. The URL it gives you is your dashboard URL.

---

## Session 2 — Deployment Reality Check + Phase A (Closing the Loops)

**Date:** 2026-06-11 (later)

### The deployment shakedown

Got the system into a live Apps Script project bound to a Sheet and immediately hit the usual real-world papercuts. Logging them here because anyone redeploying will hit the same ones:

- **`myFunction was deleted`** — the default `Code.gs` stub Apps Script auto-creates was set as the active function. Deleted it; run `setup` from the dropdown instead.
- **`Cannot read properties of null (reading 'getSheetByName')`** — I'd created a *standalone* script first. `SpreadsheetApp.getActiveSpreadsheet()` returns null there. Fix: create the script from *inside* a Sheet (Extensions → Apps Script) so it's container-bound.
- **`Gemini API 404: models/gemini-1.5-flash is not found`** — the model needs a version suffix or a current name. Switched the default to `gemini-2.0-flash`.
- **OpenAI escape hatch** — for local testing the user wanted to use an OpenAI key instead of Gemini. Added an `AI_PROVIDER` switch in the Config sheet (`gemini` | `openai`) and a parallel `_callOpenAI()` path using the chat-completions endpoint with `response_format: json_object`. Same prompt, same normalised output shape — the rest of the pipeline doesn't know or care which provider answered.

### The hard look in the mirror

After the first deploy we went back to the original problem statement and graded ourselves honestly against the six evaluation criteria. Three real holes surfaced, and one of them was embarrassing:

**Hole 1 — the human review loop went nowhere.** This was the big one. A low-confidence or complaint email got flagged, its draft stored in the HumanReview sheet, a Chat alert fired... and then *nothing*. A human could correct the category (and Memory would update), but the customer never actually got a reply. We'd built the "raise your hand" half of human-in-the-loop and completely forgotten the "lower your hand and act" half. For a criterion literally named *Human-AI Collaboration*, that's a hole a judge finds in one question: "so what happens after the human reviews it?"

**Hole 2 — transient failures ate emails.** The old `_logProcessingError` slapped the `Processed` label on a thread for *any* exception. So a Gemini rate-limit (429) or a momentary timeout would permanently bury an email that would've succeeded thirty seconds later. A reliability bug hiding inside the error handler.

**Hole 3 — the reports we promised didn't exist.** (Deferred to Phase B, but noted.)

### What Phase A actually changed

**A1 — Closed the human loop.** New file `ReviewActions.gs` with `approveAndSend(emailId, editedDraft, finalCategory)`: it finds the HumanReview row, sends the (possibly edited) reply to the original thread, stamps the row as `Approved & Sent` with reviewer + timestamp, and — if the reviewer changed the category — records the correction to Memory. There's also `dismissReview()` for spam/duplicates. I deliberately wrote these as clean standalone server functions, *then* wired the dashboard button on top. The user preferred a dashboard button (fair — deploying is genuinely one click), but `google.script.run` calls can silently no-op, which is exactly the "Run Now does nothing" symptom we were already fighting. Keeping the logic standalone means it's also callable straight from the editor as a demo-day fallback.

To avoid two code paths writing Memory differently, I pulled the Memory-append logic out of the onEdit trigger into a shared `recordMemoryCorrection()` that both the sheet-edit path and the dashboard-approve path call. Single source of truth.

The dashboard's Human Review tab is now interactive: an editable draft textarea, a category dropdown (defaulting to what the agent guessed), and **Approve & Send** / **Dismiss** buttons with per-item busy state. I also taught the 30-second auto-refresh to *not* clobber a card while the reviewer is mid-edit on it — small thing, but losing a half-written reply to a background refresh would be infuriating.

**A2 — Made failures recoverable.** Rewrote the error handler to classify failures. `_isTransientError()` sniffs the message for 429/5xx/rate-limit/quota/timeout/overloaded signatures. Transient errors leave the thread *unlabeled* so the next 5-minute run retries it, backed by a per-message retry counter in Script Properties. After `MAX_RETRIES` (3) it gives up and quarantines. Permanent errors (bad JSON, 400/401/404 config problems) quarantine immediately — no point hammering a misconfigured key 3 times per email. Quarantined threads get an `OpsAgent/Quarantine` label so the operator can find them, fix the root cause, and re-process. On success we clear the retry counter so a flaky email that eventually works doesn't carry stale state.

**A3 — A reset button for demos and stuck inboxes.** Added `devResetProcessedLabels()`: strips the Processed and Quarantine labels off every thread and wipes all retry counters, so the agent re-scans the whole inbox from scratch. This directly fixes the "I sent a test email and Run Now does nothing" situation — which is almost always because the email got labeled during an earlier failed run and is now invisible to the `-label:OpsAgent/Processed` query. Doesn't touch any sheet data, so it's safe to run anytime.

### Where we stand now

The end-to-end story is finally *complete* rather than just *plausible*: an uncertain email gets flagged → a human opens the dashboard, tweaks the draft and category, hits Approve & Send → the reply goes out AND the correction trains the agent. And the pipeline no longer quietly loses mail when an API has a bad minute.

Next up (Phase B), pending the user's call: Google Docs weekly digests in a Drive folder, a "request more info" action for incomplete emails, and surfacing the agent's reasoning for transparency.

---

## Session 3 — Phase B (Reports, Smarter Workflows, Transparency)

**Date:** 2026-06-11 (later still)

With the loops closed in Phase A, Phase B was about going from "correct" to "complete" — filling the gaps against the example actions in the brief and making the agent's thinking visible. Three things shipped.

### B1 — Weekly digest as a Google Doc (new file: `DigestReport.gs`)

This closes two gaps at once: the brief's "generate summaries / generate reports" example actions, and the two Google technologies we'd promised but never wired up — **Docs and Drive**.

`generateWeeklyDigest()` reads the last 7 days of the AuditLog, aggregates it (total processed, automation rate, review/info-request counts, complaints, quarantines, new memory corrections, average confidence, per-category breakdown, and a list of attention-worthy items), then builds a properly formatted Google Doc with a title, executive summary, two tables, and bulleted attention list. The doc is filed into an "OpsAgent Reports" Drive folder (auto-created if missing) and, if a Chat webhook is configured, the team gets a card with an **OPEN REPORT** button linking straight to it.

It runs two ways: a weekly time trigger (Mondays 8am) and a **Digest** button I added to the dashboard header for on-demand generation during demos — clicking it opens the freshly generated doc in a new tab.

One detail I got right early: the stats reader (`_readRecent`) reads the sheet's *actual* column width rather than a hardcoded count, so it doesn't break when columns get added later. That same defensiveness saved me in B3.

### B2 — "Request more info" for incomplete emails

This was the missing example action that makes the agent feel genuinely smart. Previously, if someone emailed "I'd like to register for your event!" with no event name or date, the agent would confidently send a *confirmation* for an event it knew nothing about and write a garbage row to the Registrations sheet. Now it recognises the gap and asks.

Implementation: I added a third disposition to the policy engine alongside `auto` and `human_review` — **`request_info`**. There's a `REQUIRED_FIELDS` map (currently event_registration needs both an event name and a date) and a `getMissingFields()` check that runs *after* all the confidence/sensitivity rules pass. If the agent is confident about the category but the must-have fields are blank (or "not specified" / "unknown"), it flips to `request_info`, and the action router sends a friendly templated email listing exactly what's missing — no second Gemini call needed.

The tricky part was the re-entry loop, which is the same idempotency tension from Phase A wearing a different hat. If I marked the thread `Processed`, the sender's reply-with-details would never get picked up. If I left it untouched, the thread stays unread and the agent asks for the same info every 5 minutes — an infinite nag loop. The fix: on a `request_info`, the agent sends the clarification, **marks the thread read**, and tags it `OpsAgent/AwaitingInfo` instead of `Processed`. The thread now sits quietly outside the `is:unread` query. The moment the sender replies, the thread goes unread again, re-enters the queue, gets reprocessed with the new details (the reply quotes the original for context), and — now complete — finally gets the `Processed` label, which also strips the `AwaitingInfo` tag on the way out. Clean state machine, no nagging.

### B3 — Reasoning transparency

The agent was already producing a `reasoning` field on every classification (e.g. "Mentions a budget and brand visibility, which indicates a sponsorship inquiry rather than a partnership") — and we were throwing it straight in the bin. For a criterion literally called *Agent Reasoning*, surfacing that was free points.

Now the reasoning is persisted to a new trailing **Reasoning** column in the AuditLog and an **AgentReasoning** column in HumanReview, and shown in the UI: a small italic "↳ reasoning" subline under each audit row's subject, and a dedicated blue "Agent's reasoning" panel on every Human Review card — so a reviewer sees *why* the agent made the call it's asking them to check, not just the verdict.

### The migration trap (and how I defused it)

Adding columns to sheets is easy on a fresh install — `setup()` just creates them. The problem: the user had **already run `setup()`**, so their AuditLog had 13 columns and HumanReview had 14. Two failure modes loomed:

1. New `appendRow` calls writing a 14th/15th value into a sheet whose header row was still the old width — cosmetically the data lands but the header cell is blank.
2. The dashboard's `_sheetRows()` doing `getRange(2, 1, n, 14)` on a 13-column sheet → hard "out of bounds" crash.

I fixed both. `_sheetRows()` now clamps its column count to `sheet.getLastColumn()`, so asking for a column that doesn't exist yet returns `undefined` (which maps to an empty string) instead of throwing. And I wrote a one-shot **`devUpgradeSheets()`** migration that idempotently adds the new Config rows, the two new header columns, the new Gmail labels (`AwaitingInfo`, `Quarantine`), and the weekly digest trigger — all without recreating or wiping a single existing sheet. Run it once after pulling Phase B and an already-live instance is fully upgraded.

Also added the `documents` OAuth scope to the manifest, since `DocumentApp` needs it and we maintain an explicit scope list.

### Where we stand

The agent now covers the full sweep of the brief's example actions: auto-reply, request more information, update Sheets, generate summaries/reports, escalate to humans, create workflow records, and route alerts. Every decision is logged *with its reasoning*, uncertain cases round-trip cleanly through a human, incomplete cases round-trip cleanly through the sender, and the whole week rolls up into a shareable Doc. Three new Google technologies joined the stack this session (Docs, Drive, and a real use for the AwaitingInfo workflow state).

Deployment note for an existing instance: paste the updated files + the new `DigestReport.gs`, run **`devUpgradeSheets()`** once, then redeploy the web app as a new version.

---

*More entries to follow as the system is deployed and iterated.*
