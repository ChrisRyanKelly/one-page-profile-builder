/***** CONFIG (edit these) *****/
const TEMPLATE_DOC_ID = '1xvRxNEhjDrUFhsKehL1Uy7Z0yov0003HmG44jqCICwg';
const OUTPUT_DOC_FOLDER_ID = '1UF6VuW6GivJl3Kv3HoTdis44-TU3qF0u';
const GEMINI_API_PROP_KEY = 'GEMINI_API_KEY';

const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta';
const MODEL_CANDIDATES = [
  'models/gemini-2.5-flash',
  'models/gemini-2.0-flash',
  'models/gemini-2.5-pro'
];

const REDACT_NAMES = true;          // Mask StudentName inside prompts
const MAX_OUTPUT_TOKENS = 256;      // Keep small for reliability
const CANONICALISE_SUPPORT = false; // Preserve personalised support from the model
/***** END CONFIG *****/

const EXPECTED_KEYS = [
  'StudentName','YearGroup','AboutMe','IWishMyTeacherKnew','HowToSupportMe','ProfileDate'
];

const TARGET_SUPPORT = [
  'Give brief written and verbal instructions; highlight the key steps/success criteria.',
  'Check in discreetly during independent work (“Are you on track?”).',
  'Offer a quick private check before I start a task.',
  'Let me use a low-key help signal (e.g., small desk card) instead of a raised hand.',
  'Break tasks into smaller chunks with mini-deadlines.',
  'Use visuals/models to show what good work looks like.'
];

/** === One-time helper === **/
function createOnFormSubmitTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(ScriptApp.deleteTrigger);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  Logger.log('Installable trigger created for onFormSubmit.');
}

/** === Main: on form submission, make a Doc with AI-summarised content === **/
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // serialize bursts a bit; keep it short
  try {
    if (!e) throw new Error('Missing event object. Use an installable spreadsheet trigger.');
    const tz = Session.getScriptTimeZone();
    const profileDate = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // Read header + row
    let sheet, row, headers, values;
    if (e.range) {
      sheet = e.range.getSheet();
      row = e.range.getRow();
      const lastCol = sheet.getLastColumn();
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      values  = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    } else if (e.namedValues) {
      sheet = SpreadsheetApp.getActiveSheet();
      row = sheet.getActiveRange().getRow();
      headers = Object.keys(e.namedValues);
      values  = headers.map(h => (Array.isArray(e.namedValues[h]) ? e.namedValues[h][0] : e.namedValues[h]));
    } else {
      throw new Error('Event object has neither range nor namedValues.');
    }

    // Map data by normalised keys
    const toKey = s => String(s).replace(/[^\p{L}\p{N}_]/gu, '').trim();
    const escRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const clean = v => {
      const t = String(v ?? '').replace(/<[^>]+>/g, '').trim();
      return t === '' ? '—' : t;
    };
    const data = {};
    const seen = new Map();
    headers.forEach((h, i) => {
      const k = toKey(h);
      if (seen.has(k) && clean(values[i]) !== '—' && clean(values[seen.get(k)]) !== '—') {
        console.warn(`Header collision on key "${k}" from "${h}" and "${headers[seen.get(k)]}".`);
      }
      seen.set(k, i);
      data[k] = values[i] == null ? '' : String(values[i]);
    });
    data['ProfileDate'] = profileDate;

    // Pull original fields
    const yearGroup   = (data['YearGroup']   || 'Secondary').trim();
    const studentName = (data['StudentName'] || 'Student').trim();
    const rawInputs = {
      aboutMe:            (data['AboutMe'] || '').trim(),
      iWishMyTeacherKnew: (data['IWishMyTeacherKnew'] || '').trim(),
      howToSupportMe:     (data['HowToSupportMe'] || '').trim(),
    };

    // Try AI summary (one structured call); fall back to simple tidy
    let supportLines = [];
    try {
      const summary  = generateProfileSummary_(rawInputs, yearGroup, studentName);
      const polished = polishToSLTStyle_(summary);

      data['AboutMe']            = polished.aboutMe || data['AboutMe'] || '—';
      data['IWishMyTeacherKnew'] = polished.iWishMyTeacherKnew || data['IWishMyTeacherKnew'] || '—';
      supportLines               = Array.isArray(polished.howToSupportMe) ? polished.howToSupportMe : [];
      data['HowToSupportMe']     = composeSupportParagraph_(supportLines);

    } catch (aiErr) {
      console.warn('AI summary failed; using local tidy fallback:', aiErr);
      const tidy = s => String(s || '')
        .replace(/\bweekends\b/gi, 'the weekend')
        .replace(/\bfavorite\b/gi, 'favourite')
        .replace(/\bI am\b/g, 'I’m')
        .replace(/\bI do not\b/g, 'I don’t')
        .replace(/\s+/g, ' ')
        .trim();

      data['AboutMe'] = rawInputs.aboutMe
        ? tidy(rawInputs.aboutMe)
        : 'I enjoy my lessons and work best when I know exactly what I’m meant to do. I’ve got a good group of friends and I try my best in class.';
      data['IWishMyTeacherKnew'] = rawInputs.iWishMyTeacherKnew
        ? tidy(rawInputs.iWishMyTeacherKnew) + (/[.!?]$/.test(rawInputs.iWishMyTeacherKnew) ? '' : '.')
        : 'Sometimes I’m not sure what we’re supposed to be doing and I feel nervous about asking for help. I find it easier when instructions are clear and I can check privately that I’m on the right track.';
      supportLines = TARGET_SUPPORT.slice();
      data['HowToSupportMe'] = composeSupportParagraph_(supportLines);
    }

    // Create Doc from template and replace placeholders
    const template = DriveApp.getFileById(TEMPLATE_DOC_ID);
    const parentFolder = OUTPUT_DOC_FOLDER_ID
      ? DriveApp.getFolderById(OUTPUT_DOC_FOLDER_ID)
      : (template.getParents().hasNext() ? template.getParents().next() : DriveApp.getRootFolder());

    const safeStudentName = (data['StudentName'] || 'Unknown').trim().replace(/\s+/g, ' ');
    const baseName = `${(data['YearGroup'] || 'Unknown').trim()}_${safeStudentName}_OnePageProfile_${profileDate}`
      .replace(/[^\p{L}\p{N}\- ]/gu, '_').replace(/\s+/g, '_');

    const docCopy = template.makeCopy(baseName, parentFolder);
    const doc = DocumentApp.openById(docCopy.getId());
    const body = doc.getBody();

    Object.keys(data).forEach(k => {
      body.replaceText(`\\{\\{${escRe(k)}\\}\\}`, clean(data[k]));
    });
    EXPECTED_KEYS.forEach(k => body.replaceText(`\\{\\{${escRe(k)}\\}\\}`, '—')); // ensure no {{Missing}}
    doc.saveAndClose();

    Logger.log('Created Doc: %s', doc.getUrl()); // staff can QA via Drive or Sheet link if you store it

  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** === AI: single structured call with light resilience === **/
function generateProfileSummary_(raw, yearGroup, studentName) {
  const maskName = (s) => {
    if (!REDACT_NAMES || !studentName) return String(s || '');
    const esc = studentName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return String(s || '').replace(new RegExp(`\\b${esc}\\b`, 'gi'), 'STUDENT_NAME');
  };
  const maskedName = REDACT_NAMES && studentName ? 'STUDENT_NAME' : (studentName || 'the student');

  const systemInstruction = {
    role: "system",
    parts: [{ text:
`You assist a UK school ALN/SEND coordinator. Use clear, professional UK English.
Strictly output JSON that matches the schema—no extra prose.
Tone: warm, concise, student-centred. No diagnoses or medical claims.
Field styles:
• aboutMe: first person ("I ..."), ~70–120 words; use UK spellings and "at the weekend".
• iWishMyTeacherKnew: first person, ~40–90 words; emphasise clarity and discreet support.
• howToSupportMe: 4–5 items, each a plain sentence (no numbering, no quotes).` }]
  };

  const responseSchema = {
    type: "OBJECT",
    properties: {
      aboutMe: { type: "STRING" },
      iWishMyTeacherKnew: { type: "STRING" },
      howToSupportMe: { type: "ARRAY", items: { type: "STRING" } }
    },
    required: ["aboutMe","iWishMyTeacherKnew","howToSupportMe"],
    propertyOrdering: ["aboutMe","iWishMyTeacherKnew","howToSupportMe"]
  };

  const styleExample = `
About Me:
I live with my mum and dad and I enjoy playing football at the weekend. In school, my favourite lesson is English. I’ve got a good group of friends and I work well with them. I can be quiet at first, but I try my best and I like knowing exactly what I’m meant to do.

I Wish My Teacher Knew:
Sometimes I’m not sure what we’re supposed to be doing and I feel nervous about asking for help. I find it easier when instructions are clear and I can check privately that I’m on the right track.

How to Support Me items (examples):
Give brief written and verbal instructions; highlight the key steps/success criteria.
Check in discreetly during independent work (“Are you on track?”).
Let me use a low-key help signal (e.g., small desk card) instead of a raised hand.
Use visuals/models to show what good work looks like.
`;

  const prompt = [
    `Student name: ${maskedName}`,
    `Year group: ${yearGroup || 'Secondary'}`,
    `STYLE EXAMPLE (mirror tone/structure; do not copy text):`,
    styleExample.trim(),
    `RAW INPUTS:`,
    `About Me: """${maskName(raw.aboutMe)}"""`,
    `I Wish My Teacher Knew: """${maskName(raw.iWishMyTeacherKnew)}"""`,
    `How To Support Me: """${maskName(raw.howToSupportMe)}"""`,
    `Return ONLY JSON that matches the schema.`
  ].join('\n\n');

  const body = {
    systemInstruction,
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.18,
      topK: 40,
      topP: 0.9,
      maxOutputTokens: MAX_OUTPUT_TOKENS,
      responseMimeType: "application/json",
      responseSchema
    }
  };

  let lastErr;
  for (const model of MODEL_CANDIDATES) {
    try {
      const url = `${GEMINI_BASE}/${model}:generateContent?key=${encodeURIComponent(getApiKey_())}`;
      const resp = withRetry_(() => UrlFetchApp.fetch(url, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      }));
      const code = resp.getResponseCode();
      const rawText = resp.getContentText();
      if (code !== 200) throw new Error(`HTTP ${code}: ${rawText.slice(0, 400)}`);

      const data   = JSON.parse(rawText);
      const cand   = data && data.candidates && data.candidates[0];
      const parts  = cand && cand.content && cand.content.parts || [];
      const jsonStr = parts.map(p => p && p.text).filter(Boolean).join('\n').trim();
      if (!jsonStr) throw new Error('Empty JSON payload from model');

      const parsed = JSON.parse(jsonStr);

      // Unmask name
      if (REDACT_NAMES && studentName) {
        const unmask = s => String(s || '').replace(/\bSTUDENT_NAME\b/g, studentName);
        parsed.aboutMe            = unmask(parsed.aboutMe);
        parsed.iWishMyTeacherKnew = unmask(parsed.iWishMyTeacherKnew);
        parsed.howToSupportMe     = Array.isArray(parsed.howToSupportMe) ? parsed.howToSupportMe.map(unmask) : [];
      }

      return parsed;

    } catch (err) {
      lastErr = err;
      Utilities.sleep(600 + Math.floor(Math.random() * 800)); // brief cool-off before next model
    }
  }
  throw lastErr || new Error('All models failed to return JSON.');
}

/** === Light UK polish + house-style supports (optional) === **/
function polishToSLTStyle_(summary) {
  const s = {...summary};

  if (s.aboutMe) {
    let t = s.aboutMe
      .replace(/\bweekends\b/gi, 'the weekend')
      .replace(/\bfavorite\b/gi, 'favourite')
      .replace(/\bI have\b/gi, 'I’ve');
    if (/I(’|'|`)ve got a good group of friends\b/i.test(t) && !/work well with them/i.test(t)) {
      t = t.replace(/I(’|'|`)ve got a good group of friends\b/i, 'I’ve got a good group of friends and I work well with them');
    }
    s.aboutMe = tidySentences_(t);
  }

  if (s.iWishMyTeacherKnew) {
    let t = s.iWishMyTeacherKnew
      .replace(/\bI am\b/g, 'I’m')
      .replace(/\bI do not\b/g, 'I don’t')
      .replace(/\bsupposed to be doing\b/gi, 'what we’re supposed to be doing')
      .replace(/\bafraid to ask for help\b/gi, 'nervous about asking for help')
      .trim();
    s.iWishMyTeacherKnew = tidySentences_(dedupeSentences_(t));
  }

  // Keep model items; tidy, dedupe, cap to 5.
  const modelItems = Array.isArray(s.howToSupportMe) ? s.howToSupportMe.map(tidySentence_) : [];
  s.howToSupportMe = dedupeKeepOrder_(modelItems).slice(0, 5);

  return s;
}

function composeSupportParagraph_(items) {
  const list = dedupeKeepOrder_((items || []).map(tidySentence_)).slice(0, 5);
  return list.length ? tidySentences_(list.join(' ')) : '—';
}

/** === Minimal resilience: retry only when it helps === **/
function withRetry_(fn, opts = {}) {
  const {
    retries = 4,
    minDelayMs = 600,
    maxDelayMs = 6000,
    factor = 2,
    jitter = true
  } = opts;

  let attempt = 0, lastErr;
  while (attempt <= retries) {
    try {
      return fn();
    } catch (err) {
      lastErr = err;
      const msg = String(err);
      const transient =
        /HTTP (429|500|502|503|504)/.test(msg) ||
        /timed out|Request failed|Service unavailable|reset by peer/i.test(msg);

      if (!transient || attempt === retries) throw err;

      const base = Math.min(minDelayMs * Math.pow(factor, attempt), maxDelayMs);
      const delay = jitter ? base / 2 + Math.random() * (base / 2) : base;
      Utilities.sleep(delay);
      attempt++;
    }
  }
  throw lastErr;
}

/** === Utilities === **/
function tidySentence_(s) {
  let t = String(s || '').trim().replace(/\s+/g, ' ');
  if (!t) return '';
  if (!/[.!?]["’"]?$/.test(t)) t += '.';
  return t;
}
function tidySentences_(s) {
  return String(s || '')
    .replace(/\s+/g, ' ')
    .replace(/\s,/, ',')
    .trim()
    .replace(/ +([,.!?;:])/g, '$1');
}
function dedupeSentences_(s) {
  const parts = String(s || '').split(/(?<=[.!?])\s+/);
  const seen = new Set();
  const out = [];
  for (const p of parts) {
    const key = p.trim().toLowerCase();
    if (key && !seen.has(key)) {
      seen.add(key);
      out.push(p.trim());
    }
  }
  return out.join(' ');
}
function dedupeKeepOrder_(arr) {
  const seen = new Set(), out = [];
  for (const x of (arr || [])) {
    const t = String(x || '').trim(); if (!t) continue;
    const k = t.toLowerCase(); if (!seen.has(k)) { seen.add(k); out.push(t); }
  }
  return out;
}
function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty(GEMINI_API_PROP_KEY);
  if (!key) throw new Error(`Missing ${GEMINI_API_PROP_KEY} in Script Properties.`);
  return key;
}

/** === Dev helper: run on selected row === **/
function runOnActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (row === 1) throw new Error('Select a data row, not the header row.');
  const fake = { range: sheet.getRange(row, 1, 1, sheet.getLastColumn()) };
  onFormSubmit(fake);
}