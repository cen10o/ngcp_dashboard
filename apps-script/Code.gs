// ============================================================
//  NGCP Dashboard — Google Apps Script Web App
//
//  DEPLOY STEPS:
//  1. Open your Google Sheet → Extensions → Apps Script
//  2. Delete any existing code and paste this entire file
//  3. Fill in CONIFG below with your real tab names & column indices
//  4. Set your secret token (ONE-TIME SETUP):
//       Apps Script editor → Project Settings (gear icon)
//       → Script properties → Add property
//       Name:  DASHBOARD_TOKEN
//       Value: any strong random string you choose
//              (e.g. generate one at: https://1password.com/password-generator/)
//  5. Click Deploy → New deployment → Web app
//       Execute as: Me
//       Who has access: Anyone   ← token is the auth layer, not Google login
//  6. Copy the Web app URL → paste into index.html as SCRIPT_URL
//  7. Share the token value with anyone who needs dashboard access
//
//  PRIVACY NOTE:
//  This script NEVER returns individual responses. Only aggregated numbers
//  (averages, counts) leave the spreadsheet. Names, emails, and IDs stay
//  in the sheet regardless of what columns exist.
// ============================================================

// ── CONFIGURATION ──────────────────────────────────────────
// Column indices are 0-based. Column A = 0, B = 1, C = 2 …
// Google Forms puts the timestamp in column 0 (A).
//
// !! IMPORTANT: Only include columns that contain Likert ratings
//    or open-ended text answers. NEVER include columns that hold
//    names, email addresses, student IDs, or any other PII.
//    The aggregation functions only ever return numbers and keyword
//    counts — but setting wrong columns here would cause raw PII
//    to flow into open-ended text matching, even if it isn't sent.

const CONFIG = {

  // Each survey tab that feeds the "Engagement by event" chart.
  surveys: [
    {
      tabName:          'Orientation Feedback',
      eventLabel:       'Orientation',
      likertColumns:    [4],           // col 6 was empty — add correct column if there is a second rating question
      openEndedColumns: [9, 14],
    },
    {
      tabName:          'Authenticity & Professionalism',
      eventLabel:       'Authenticity & Professionalism',
      likertColumns:    [6, 8, 10],    // col 14 contained open-ended text, moved below
      openEndedColumns: [14, 15, 16, 17, 18],
    },
    {
      tabName:          'Scavenger Hunt',
      eventLabel:       'Scavenger Hunt',
      likertColumns:    [6, 8, 10],    // col 14 contained open-ended text, moved below
      openEndedColumns: [14, 15, 16, 17, 18],
    },
    {
      tabName:          'Informational Interviews',
      eventLabel:       'Informational Interviews',
      likertColumns:    [6, 8, 10],    // col 14 contained open-ended text, moved below
      openEndedColumns: [14, 15, 16, 17, 18],
    },
    {
      tabName:          'Summer Intern Fun Bus', // !! TAB NOT FOUND — update to exact tab name from your sheet
      eventLabel:       'Fun Bus',
      likertColumns:    [1, 2, 3],
      openEndedColumns: [10],
    },
    {
      tabName:          'Building brand (Semester)',        // !! TAB NOT FOUND — update to exact tab name from your sheet
      eventLabel:       'Building Your Brand',
      likertColumns:    [6, 8, 10],
      openEndedColumns: [14, 15, 16, 17, 18],
    },
  ],

  // "Impact: before vs. after" grouped bar chart.
  impact: {
    preTab:      'Summer Internship Survey (Pre)',
    postTab:     'Summer Internship Survey (Post)',
    questions:   ['Career clarity', 'Job skills', 'Confidence', 'Networking'],
    preColumns:  [20, 22, 23, 24],
    postColumns: [8, 10, 11, 12],
  },

  // "Mentor experience" grouped bar chart.
  mentor: {
    startTab:     'Mentor Survey',
    endTab:       'Mentor End',
    questions:    ['Would mentor again', 'Felt confident', 'Recieved adequate training', 'Recieved adequate resources & support'],
    startColumns: [4, 5, 6],
    endColumns:   [9, 15, 13, 10],
  },

  // Keyword themes for open-ended responses.
  // Counts how many responses contain at least one keyword per theme.
  themes: [
    { label: 'Hands-on experience', keywords: ['hands-on', 'hands on', 'practical', 'real-world', 'experience', 'applied'] },
    { label: 'Mentor support',       keywords: ['mentor', 'guidance', 'support', 'helped me', 'advisor'] },
    { label: 'Career clarity',       keywords: ['career', 'future', 'clarity', 'direction', 'goal', 'path'] },
    { label: 'Professional skills',  keywords: ['professional', 'workplace', 'communication', 'teamwork', 'collaboration'] },
    { label: 'Networking',           keywords: ['network', 'connect', 'relationship', 'people', 'peers', 'colleagues'] },
    { label: 'Technical skills',     keywords: ['technical', 'technology', 'software', 'coding', 'data', 'tool'] },
  ],
};

// ── AUTHENTICATION ─────────────────────────────────────────

function doGet(e) {
  const token    = e && e.parameter ? e.parameter.token : null;
  const expected = PropertiesService.getScriptProperties().getProperty('DASHBOARD_TOKEN');

  // Reject if the token property was never configured.
  if (!expected) {
    return jsonResponse({ error: 'Server misconfiguration: DASHBOARD_TOKEN not set in Script Properties.' });
  }

  // Reject if the caller didn't supply the right token.
  if (token !== expected) {
    return jsonResponse({ error: 'Unauthorized' });
  }

  return jsonResponse(aggregateData());
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LIKERT TEXT → NUMBER MAP ───────────────────────────────
// Handles surveys where Google Forms stores scale responses as
// text labels instead of numbers.
const LIKERT_MAP = {
  // Agreement scale
  'strongly agree':             5,
  'agree':                      4,
  'neutral':                    3,
  'neither agree nor disagree': 3,
  'disagree':                   2,
  'strongly disagree':          1,
  // Satisfaction scale (used by Summer Intern Fun Bus)
  'very satisfied':             5,
  'satisfied':                  4,
  'dissatisfied':               2,
  'very dissatisfied':          1,
};

function parseLikert(val) {
  if (val === '' || val === null || val === undefined) return NaN;
  const n = parseFloat(val);
  if (!isNaN(n)) return n; // already a number — use as-is
  const key = String(val).trim().toLowerCase();
  return key in LIKERT_MAP ? LIKERT_MAP[key] : NaN;
}

// ── AGGREGATION ────────────────────────────────────────────
// All functions below return only numbers and labels — no raw
// row data, no names, no emails, no IDs ever leave this script.

function aggregateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const engagement = buildEngagement(ss);
  const impact     = buildImpact(ss);
  const mentor     = buildMentor(ss);
  const themes     = buildThemes(ss);
  const stats      = buildStats(ss, engagement, impact);

  return { stats, engagement, impact, mentor, themes, lastUpdated: new Date().toISOString() };
}

function avgLikertColumn(rows, col) {
  const vals = rows.map(r => parseLikert(r[col])).filter(v => !isNaN(v) && v >= 1 && v <= 5);
  if (!vals.length) return 0;
  return round1(vals.reduce((a, b) => a + b, 0) / vals.length);
}

function buildEngagement(ss) {
  const labels = [], values = [];
  for (const survey of CONFIG.surveys) {
    const sheet = ss.getSheetByName(survey.tabName);
    if (!sheet) continue;
    const rows = getRows(sheet);
    const allVals = survey.likertColumns.flatMap(col =>
      rows.map(r => parseLikert(r[col])).filter(v => !isNaN(v) && v >= 1 && v <= 5)
    );
    labels.push(survey.eventLabel);
    values.push(allVals.length ? round1(allVals.reduce((a, b) => a + b, 0) / allVals.length) : 0);
  }
  return { labels, values };
}

function buildImpact(ss) {
  const { preTab, postTab, questions, preColumns, postColumns } = CONFIG.impact;
  const before = [], after = [];
  const preSheet  = ss.getSheetByName(preTab);
  const postSheet = ss.getSheetByName(postTab);
  for (let i = 0; i < questions.length; i++) {
    before.push(preSheet  ? avgLikertColumn(getRows(preSheet),  preColumns[i])  : 0);
    after.push(postSheet  ? avgLikertColumn(getRows(postSheet), postColumns[i]) : 0);
  }
  return { labels: questions, before, after };
}

function buildMentor(ss) {
  const { startTab, endTab, questions, startColumns, endColumns } = CONFIG.mentor;
  const start = [], end = [];
  const startSheet = ss.getSheetByName(startTab);
  const endSheet   = ss.getSheetByName(endTab);
  for (let i = 0; i < questions.length; i++) {
    start.push(startSheet ? avgLikertColumn(getRows(startSheet), startColumns[i]) : 0);
    end.push(endSheet     ? avgLikertColumn(getRows(endSheet),   endColumns[i])   : 0);
  }
  return { labels: questions, start, end };
}

function buildThemes(ss) {
  // Collects raw open-ended text locally, then discards it.
  // Only the keyword match counts are returned.
  const allText = [];
  for (const survey of CONFIG.surveys) {
    const sheet = ss.getSheetByName(survey.tabName);
    if (!sheet) continue;
    const rows = getRows(sheet);
    for (const row of rows) {
      for (const col of survey.openEndedColumns) {
        const cell = row[col];
        if (cell) allText.push(String(cell).toLowerCase());
      }
    }
  }
  const labels = [], counts = [];
  for (const theme of CONFIG.themes) {
    labels.push(theme.label);
    counts.push(allText.filter(text => theme.keywords.some(kw => text.includes(kw))).length);
  }
  return { labels, counts };
}

function buildStats(ss, engagement, impact) {
  const totalResponses = CONFIG.surveys.reduce((sum, survey) => {
    const sheet = ss.getSheetByName(survey.tabName);
    return sum + (sheet ? Math.max(0, sheet.getLastRow() - 1) : 0);
  }, 0);

  const avgRating = engagement.values.length
    ? round1(engagement.values.reduce((a, b) => a + b, 0) / engagement.values.length)
    : 0;

  const confidenceIdx = 2; // index of "Confidence" in impact.questions
  const confidenceGrowth = round1(
    (impact.after[confidenceIdx] || 0) - (impact.before[confidenceIdx] || 0)
  );

  const preSheet = ss.getSheetByName(CONFIG.impact.preTab);
  const internsEnrolled = preSheet ? Math.max(0, preSheet.getLastRow() - 1) : 0;

  return [
    { lbl: 'Interns enrolled',  num: String(internsEnrolled) },
    { lbl: 'Survey responses',  num: String(totalResponses) },
    { lbl: 'Avg. event rating', num: String(avgRating) },
    { lbl: 'Confidence growth', num: (confidenceGrowth >= 0 ? '+' : '') + confidenceGrowth },
  ];
}

// ── HELPERS ────────────────────────────────────────────────

function getRows(sheet) {
  return sheet.getDataRange().getValues().slice(1); // skip header
}

function round1(n) {
  return Math.round(n * 10) / 10;
}

// ── LOCAL TEST ─────────────────────────────────────────────
// Run this in the Apps Script editor (not via HTTP) to preview
// the JSON without needing a token. Check Execution Log for output.
function testAggregation() {
  Logger.log(JSON.stringify(aggregateData(), null, 2));
}

// Run this to diagnose low engagement averages.
// It logs the first 3 raw values from each Likert column so you can
// see exactly what is stored in those cells (numbers vs. text vs. empty).
function debugEngagement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const survey of CONFIG.surveys) {
    const sheet = ss.getSheetByName(survey.tabName);
    if (!sheet) { Logger.log(survey.tabName + ': TAB NOT FOUND'); continue; }
    const rows = getRows(sheet);
    Logger.log('── ' + survey.tabName + ' (' + rows.length + ' responses) ──');
    for (const col of survey.likertColumns) {
      const samples = rows.slice(0, 3).map(r => JSON.stringify(r[col]));
      const parsed  = rows.map(r => parseLikert(r[col])).filter(v => !isNaN(v) && v >= 1 && v <= 5);
      Logger.log('  col ' + col + ': samples=' + samples.join(', ') + '  →  ' + parsed.length + ' valid values, avg=' + (parsed.length ? round1(parsed.reduce((a,b)=>a+b,0)/parsed.length) : 'N/A'));
    }
  }
}
