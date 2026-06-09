// ============================================================
//  NGCP Dashboard — Google Apps Script Web App
//  HOW TO DEPLOY:
//  1. Open your Google Sheet → Extensions → Apps Script
//  2. Delete any existing code and paste this entire file
//  3. Fill in CONFIG below with your real tab names & column indices
//  4. Click Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Copy the Web app URL and paste it into index.html as SCRIPT_URL
// ============================================================

// ── CONFIGURATION ──────────────────────────────────────────
// Column indices are 0-based. Column A = 0, B = 1, C = 2 …
// Google Forms always puts the timestamp in column 0 (A).
// Your first question is usually column 1 (B), and so on.

const CONFIG = {

  // Each survey that feeds the "Engagement by event" bar chart.
  // Add or remove objects to match your actual tabs.
  surveys: [
    {
      tabName:          'Orientation',       // exact name of the Sheet tab
      eventLabel:       'Orientation',       // label shown on the chart
      likertColumns:    [1, 2, 3, 4],        // columns that contain 1–5 ratings
      openEndedColumns: [5, 6],              // columns that contain free-text answers
    },
    {
      tabName:          'Fun Bus',
      eventLabel:       'Fun Bus',
      likertColumns:    [1, 2, 3],
      openEndedColumns: [4],
    },
    {
      tabName:          'Friday Learning',
      eventLabel:       'Friday Learning',
      likertColumns:    [1, 2, 3, 4],
      openEndedColumns: [5],
    },
    {
      tabName:          'Pre-Internship',
      eventLabel:       'Pre-internship',
      likertColumns:    [1, 2, 3, 4],
      openEndedColumns: [5, 6],
    },
    {
      tabName:          'Post-Internship',
      eventLabel:       'Post-internship',
      likertColumns:    [1, 2, 3, 4],
      openEndedColumns: [5, 6],
    },
  ],

  // "Impact: before vs. after" grouped bar chart.
  // Pair the same question from two different tabs.
  impact: {
    preTab:      'Pre-Internship',   // tab with pre-program responses
    postTab:     'Post-Internship',  // tab with post-program responses
    // One label per question, in the same order as the columns below.
    questions:   ['Career clarity', 'Job skills', 'Confidence', 'Networking'],
    preColumns:  [1, 2, 3, 4],      // Likert columns in the pre tab
    postColumns: [1, 2, 3, 4],      // Likert columns in the post tab (same questions)
  },

  // "Mentor experience" grouped bar chart.
  mentor: {
    startTab:     'Mentor Start',    // tab filled out at program start
    endTab:       'Mentor End',      // tab filled out at program end
    questions:    ['Felt prepared', 'Felt supported', 'Would mentor again'],
    startColumns: [1, 2, 3],
    endColumns:   [1, 2, 3],
  },

  // Keyword themes for open-ended responses.
  // The script counts how many responses contain at least one keyword per theme.
  // Add keywords that match how your interns actually write.
  themes: [
    { label: 'Hands-on experience', keywords: ['hands-on', 'hands on', 'practical', 'real-world', 'experience', 'applied'] },
    { label: 'Mentor support',       keywords: ['mentor', 'guidance', 'support', 'helped me', 'advisor'] },
    { label: 'Career clarity',       keywords: ['career', 'future', 'clarity', 'direction', 'goal', 'path'] },
    { label: 'Professional skills',  keywords: ['professional', 'workplace', 'communication', 'teamwork', 'collaboration'] },
    { label: 'Networking',           keywords: ['network', 'connect', 'relationship', 'people', 'peers', 'colleagues'] },
    { label: 'Technical skills',     keywords: ['technical', 'technology', 'software', 'coding', 'data', 'tool'] },
  ],
};

// ── WEB APP ENTRY POINT ────────────────────────────────────

function doGet() {
  const result = aggregateData();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── AGGREGATION ────────────────────────────────────────────

function aggregateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const engagement = buildEngagement(ss);
  const impact     = buildImpact(ss);
  const mentor     = buildMentor(ss);
  const themes     = buildThemes(ss);
  const stats      = buildStats(ss, engagement, impact);

  return { stats, engagement, impact, mentor, themes, lastUpdated: new Date().toISOString() };
}

// Average all Likert columns across every row in a tab.
function avgLikertColumn(rows, col) {
  const vals = rows.map(r => parseFloat(r[col])).filter(v => !isNaN(v) && v >= 1 && v <= 5);
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
      rows.map(r => parseFloat(r[col])).filter(v => !isNaN(v) && v >= 1 && v <= 5)
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
  // Collect all open-ended text from every survey tab
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
  // Total responses = sum of all data rows across survey tabs
  const totalResponses = CONFIG.surveys.reduce((sum, survey) => {
    const sheet = ss.getSheetByName(survey.tabName);
    return sum + (sheet ? Math.max(0, sheet.getLastRow() - 1) : 0);
  }, 0);

  const avgRating = engagement.values.length
    ? round1(engagement.values.reduce((a, b) => a + b, 0) / engagement.values.length)
    : 0;

  // Confidence column is index 2 in the impact questions array (0-based)
  const confidenceIdx = 2;
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

// Returns all data rows (skips the header row at index 0)
function getRows(sheet) {
  const values = sheet.getDataRange().getValues();
  return values.slice(1);
}

function round1(n) {
  return Math.round(n * 10) / 10;
}

// ── LOCAL TEST ─────────────────────────────────────────────
// Run this function from the Apps Script editor to preview
// the JSON that will be sent to your dashboard.
function testAggregation() {
  Logger.log(JSON.stringify(aggregateData(), null, 2));
}
