// GeminiService.gs — All interactions with AI APIs (Gemini + OpenAI).
// Handles classification, field extraction, confidence scoring, reply drafting,
// and few-shot injection from the Memory sheet.

const GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta/models';
const OPENAI_API_BASE = 'https://api.openai.com/v1/chat/completions';

const CATEGORIES = [
  'sponsorship',
  'event_registration',
  'speaker_application',
  'volunteer_application',
  'complaint',
  'partnership',
  'feedback',
  'general_support',
];

// ---------------------------------------------------------------------------
// Main public function
// ---------------------------------------------------------------------------

/**
 * Sends an email to Gemini for classification + extraction + reply drafting.
 * Returns a structured object matching the response schema below.
 *
 * @param {Object} emailData      - { sender, senderName, subject, body, … }
 * @param {Array}  memoryExamples - Correction history rows from the Memory sheet
 * @param {Array}  templates      - Template rows from the Templates sheet
 * @returns {Object} geminiResult - { category, confidence, reasoning,
 *                                    extracted_fields, draft_reply, sensitive_flags }
 */
function classifyAndExtract(emailData, memoryExamples, templates) {
  const provider = getAIProvider();
  const orgName  = getConfigValue('ORG_NAME', 'Our Organization');
  const prompt   = _buildPrompt(emailData, memoryExamples, templates, orgName);

  let raw;
  if (provider === 'openai') {
    raw = _callOpenAI(prompt);
  } else {
    raw = _callGemini(prompt);
  }

  let result;
  try {
    result = JSON.parse(raw);
  } catch (e) {
    throw new Error(`AI response was not valid JSON:\n${raw.substring(0, 500)}`);
  }

  // Defensive normalisation
  result.category        = CATEGORIES.includes(result.category) ? result.category : 'general_support';
  result.confidence      = Math.max(0, Math.min(1, parseFloat(result.confidence) || 0));
  result.sensitive_flags = Array.isArray(result.sensitive_flags) ? result.sensitive_flags : [];
  result.extracted_fields = result.extracted_fields || {};
  result.draft_reply      = result.draft_reply || '';
  result.reasoning        = result.reasoning   || '';

  return result;
}

// ---------------------------------------------------------------------------
// Provider: Gemini
// ---------------------------------------------------------------------------

function _callGemini(prompt) {
  const model  = getConfigValue('GEMINI_MODEL', 'gemini-2.0-flash');
  const apiKey = getGeminiApiKey();
  const url    = `${GEMINI_API_BASE}/${model}:generateContent?key=${apiKey}`;

  const response = UrlFetchApp.fetch(url, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(_buildGeminiBody(prompt)),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code !== 200) throw new Error(`Gemini API ${code}: ${text.substring(0, 400)}`);

  const json = JSON.parse(text);
  if (!json.candidates || json.candidates.length === 0) {
    const reason = (json.promptFeedback && json.promptFeedback.blockReason) || 'unknown';
    throw new Error(`Gemini returned no candidates. Block reason: ${reason}`);
  }
  return json.candidates[0].content.parts[0].text;
}

// ---------------------------------------------------------------------------
// Provider: OpenAI
// ---------------------------------------------------------------------------

function _callOpenAI(prompt) {
  const model  = getConfigValue('OPENAI_MODEL', 'gpt-4o-mini');
  const apiKey = getOpenAIApiKey();

  const body = {
    model: model,
    messages: [
      { role: 'system', content: 'You are a JSON-only email classification agent. Respond with ONLY valid JSON — no markdown fences, no extra text.' },
      { role: 'user',   content: prompt },
    ],
    temperature: 0.15,
    max_tokens:  2048,
    response_format: { type: 'json_object' },
  };

  const response = UrlFetchApp.fetch(OPENAI_API_BASE, {
    method:             'post',
    contentType:        'application/json',
    headers:            { 'Authorization': 'Bearer ' + apiKey },
    payload:            JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code !== 200) throw new Error(`OpenAI API ${code}: ${text.substring(0, 400)}`);

  const json = JSON.parse(text);
  if (!json.choices || json.choices.length === 0) {
    throw new Error('OpenAI returned no choices');
  }
  return json.choices[0].message.content;
}

// ---------------------------------------------------------------------------
// Memory helpers (used by EmailProcessor)
// ---------------------------------------------------------------------------

/**
 * Retrieves up to MAX_MEMORY_EXAMPLES corrections from the Memory sheet.
 * Prioritises examples that match the hint category, then fills with others.
 * Pass null for categoryHint to get a general mix.
 */
function getMemoryExamples(categoryHint) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMORY);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const max  = parseInt(getConfigValue('MAX_MEMORY_EXAMPLES', '5'), 10);
  const data = sheet.getDataRange().getValues().slice(1); // skip header

  const examples = data
    .filter(r => r[0] && r[4] && r[5])           // must have timestamp, original, corrected
    .slice(-max * 4)                               // take tail (most recent) only
    .map(r => ({
      subject:           String(r[2]),
      snippet:           String(r[3]).substring(0, 300),
      originalCategory:  String(r[4]),
      correctedCategory: String(r[5]),
    }));

  if (!categoryHint) return examples.slice(-max);

  const relevant = examples.filter(e => e.correctedCategory === categoryHint);
  const others   = examples.filter(e => e.correctedCategory !== categoryHint);
  return [...relevant, ...others].slice(0, max);
}

/**
 * Reads all template rows from the Templates sheet.
 */
function getTemplates() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.TEMPLATES);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getDataRange().getValues().slice(1).map(r => ({
    category:     String(r[0]),
    tone:         String(r[1]),
    keyPoints:    String(r[2]),
    opening:      String(r[3]),
    closing:      String(r[4]),
    includeForm:  String(r[5]).toLowerCase() === 'true',
    formKey:      String(r[6]),
  })).filter(t => t.category);
}

// ---------------------------------------------------------------------------
// Prompt construction
// ---------------------------------------------------------------------------

function _buildPrompt(emailData, memoryExamples, templates, orgName) {
  return `You are the intelligent email operations agent for "${orgName}", a student organisation.
Your job: classify incoming emails, extract structured data, and draft personalised replies.

## Allowed categories (use EXACTLY one of these strings)
${CATEGORIES.map((c, i) => `${i + 1}. ${c}`).join('\n')}

## Per-category response guidelines
- sponsorship         : Acknowledge interest; mention 3-day follow-up SLA and review process.
- event_registration  : Confirm registration warmly; mention that a calendar invite was created.
- speaker_application : Thank them; mention 5-day rolling review; say a form link follows.
- volunteer_application: Welcome enthusiasm; say a form link follows for details.
- complaint           : Express empathy; DO NOT attempt to resolve; promise human follow-up within 24 h.
- partnership         : Express collaborative interest; 5-day response SLA.
- feedback            : Thank sincerely; say feedback will be shared with the team.
- general_support     : Answer helpfully and concisely.

## Templates (use as tone/style guidance, not verbatim)
${_buildTemplatesBlock(templates)}

${_buildFewShotBlock(memoryExamples)}

## Email to classify
From   : ${emailData.sender}
Subject: ${emailData.subject}
Body:
${emailData.body.substring(0, 3000)}

## Your task
1. Pick the single best category.
2. Set confidence 0.0–1.0 (be honest — underconfidence is safer than overconfidence).
3. Extract every structured field you can find in the email.
4. Write a personalised, warm, human-sounding reply (2–4 paragraphs). Do not sound templated.
   Sign off as the "${orgName} Team".
5. List any words/phrases that suggest sensitivity (legal risk, harassment, urgency, refund demands, etc.).
6. Explain your classification reasoning briefly.

Respond with ONLY valid JSON — no markdown fences, no extra text.

Required JSON shape:
{
  "category": "<one of the 8 categories>",
  "confidence": <0.0–1.0>,
  "reasoning": "<1–2 sentences>",
  "extracted_fields": {
    "sender_name": "",
    "organization": "",
    "event_name": "",
    "event_date": "",
    "participants_count": "",
    "topic": "",
    "urgency": "",
    "key_ask": ""
  },
  "draft_reply": "<full reply body, no salutation line — start from the first content sentence>",
  "sensitive_flags": ["<word or phrase>"]
}`;
}

function _buildTemplatesBlock(templates) {
  if (!templates || templates.length === 0) return '(none configured)';
  return templates.map(t =>
    `• ${t.category}: Tone = ${t.tone}. Key points = ${t.keyPoints}.`
  ).join('\n');
}

function _buildFewShotBlock(examples) {
  if (!examples || examples.length === 0) return '';
  let block = '## Correction history — learn from these past mistakes\n';
  block += 'These are emails the agent misclassified. A human corrected them.\n\n';
  examples.forEach((ex, i) => {
    block += `### Past example ${i + 1}\n`;
    block += `Subject : "${ex.subject}"\n`;
    block += `Snippet : "${ex.snippet}"\n`;
    block += `Agent classified as : ${ex.originalCategory}\n`;
    block += `CORRECT category    : **${ex.correctedCategory}**\n\n`;
  });
  return block;
}

// ---------------------------------------------------------------------------
// Request body builder
// ---------------------------------------------------------------------------

function _buildGeminiBody(prompt) {
  return {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature:     0.15,   // Low temperature → consistent classification
      topK:            10,
      topP:            0.9,
      maxOutputTokens: 2048,
      // Ask Gemini to return JSON — more reliable than asking in the prompt alone
      responseMimeType: 'application/json',
    },
    safetySettings: [
      { category: 'HARM_CATEGORY_HARASSMENT',       threshold: 'BLOCK_NONE' },
      { category: 'HARM_CATEGORY_HATE_SPEECH',      threshold: 'BLOCK_NONE' },
      { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT',threshold: 'BLOCK_NONE' },
      { category: 'HARM_CATEGORY_DANGEROUS_CONTENT',threshold: 'BLOCK_NONE' },
    ],
  };
}
