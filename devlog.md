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

*More entries to follow as the system is deployed and iterated.*
