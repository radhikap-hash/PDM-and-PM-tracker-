/**
 * PDM & PM tracker → GitHub Pages publisher
 * -----------------------------------------
 * Bind this script to the "PDM & PM" Google Sheet.
 * It reads two tabs (Products, Engineering), builds data.json in the
 * exact shape dashboard.html expects, and commits/updates it in the
 * GitHub repo via the Contents API.
 *
 * ONE-TIME SETUP
 * 1. Extensions > Apps Script, paste this file in as Code.gs.
 * 2. Project Settings > Script Properties, add:
 *      GITHUB_TOKEN   = a fine-grained PAT with "Contents: read & write"
 *                       scoped to this one repo
 *      GITHUB_OWNER   = radhikap-hash
 *      GITHUB_REPO    = PDM-and-PM-tracker-
 *      GITHUB_BRANCH  = main
 *      GITHUB_PATH    = data.json
 * 3. Run setup() once from the Apps Script editor to create the menu
 *    and the two tabs (if they don't exist yet) with header rows.
 * 4. Reload the Sheet — a "Dashboard" menu appears with "Publish now".
 * 5. Optional: run enableAutoPublish() once to also publish automatically
 *    ~2 minutes after any edit (debounced, so a burst of edits produces
 *    one commit, not one per keystroke).
 *
 * SHEET LAYOUT
 * Tab "Products" (one row per product variant):
 *   Category | Product Name | ID | PDM | PM/TL | Stage (1-11) |
 *   Days In Stage | Days Total | Current Doc | Blocker | Tentative Date
 *
 * Tab "Engineering" (one row per weekly update, keyed by Product ID —
 * latest row per ID wins):
 *   Product ID | Period | Owners | Platform | Approach |
 *   Completed (one bullet per line, or separated by " | ") |
 *   Next Steps (same format) | Risk | Status | ETA
 */

const PRODUCTS_SHEET = 'Products';
const ENGINEERING_SHEET = 'Engineering';

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, PRODUCTS_SHEET,
    ['Category','Product Name','ID','PDM','PM/TL','Stage (1-11)','Days In Stage','Days Total','Current Doc','Blocker','Tentative Date']);
  ensureSheet_(ss, ENGINEERING_SHEET,
    ['Product ID','Period','Owners','Platform','Approach','Completed','Next Steps','Risk','Status','ETA']);
  onOpen();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Publish now', 'publishToGithub')
    .addItem('Enable auto-publish (~2 min after edits)', 'enableAutoPublish')
    .addItem('Disable auto-publish', 'disableAutoPublish')
    .addToUi();
}

/** Reads both tabs and returns the JS object matching dashboard.html's schema. */
function buildPayload_() {
  const ss = SpreadsheetApp.getActive();

  const prodSheet = ss.getSheetByName(PRODUCTS_SHEET);
  const prodRows = prodSheet.getDataRange().getValues().slice(1).filter(r => r[2]); // needs an ID
  const products = prodRows.map(r => ({
    cat: str_(r[0]), name: str_(r[1]), id: str_(r[2]),
    pdm: str_(r[3]) || null, pm: str_(r[4]) || null,
    stage: Number(r[5]) || 1,
    daysStage: Number(r[6]) || 0, daysTotal: Number(r[7]) || 0,
    doc: str_(r[8]) || 'TBD', blocker: str_(r[9]) || null,
    tentative: str_(r[10]) || null
  }));

  const engSheet = ss.getSheetByName(ENGINEERING_SHEET);
  const engineering = {};
  if (engSheet) {
    const engRows = engSheet.getDataRange().getValues().slice(1).filter(r => r[0]);
    engRows.forEach(r => {
      engineering[str_(r[0])] = {
        period: str_(r[1]), owners: str_(r[2]) || 'Not specified',
        platform: str_(r[3]) || 'Not specified', approach: str_(r[4]) || 'Not specified',
        completed: splitList_(r[5]), next: splitList_(r[6]),
        risk: str_(r[7]) || null, status: str_(r[8]) || 'In Progress',
        eta: str_(r[9]) || null
      };
      // later rows for the same Product ID overwrite earlier ones, so the
      // sheet's row order naturally lets you keep a history and only the
      // last row per ID is published.
    });
  }

  return {
    generatedAt: new Date().toISOString(),
    products: products,
    engineering: engineering
  };
}

function str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function splitList_(v) {
  const s = str_(v);
  if (!s) return [];
  return s.split(/\n|\s\|\s/).map(x => x.trim()).filter(Boolean);
}

/** Commits data.json to GitHub via the Contents API (create or update). */
function publishToGithub() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER');
  const repo = props.getProperty('GITHUB_REPO');
  const branch = props.getProperty('GITHUB_BRANCH') || 'main';
  const path = props.getProperty('GITHUB_PATH') || 'data.json';

  if (!token || !owner || !repo) {
    throw new Error('Set GITHUB_TOKEN, GITHUB_OWNER and GITHUB_REPO in Script Properties first.');
  }

  const payload = buildPayload_();
  const content = Utilities.base64Encode(JSON.stringify(payload, null, 2), Utilities.Charset.UTF_8);
  const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/vnd.github+json'
  };

  // Look up the current file SHA (required by GitHub to update an existing file).
  let sha = null;
  const getResp = UrlFetchApp.fetch(`${apiUrl}?ref=${branch}`, { headers, muteHttpExceptions: true });
  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  }

  const body = {
    message: `Publish dashboard data — ${new Date().toISOString()}`,
    content: content,
    branch: branch
  };
  if (sha) body.sha = sha;

  const putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers,
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = putResp.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error(`GitHub publish failed (${code}): ${putResp.getContentText()}`);
  }
  Logger.log('Published data.json — %s', new Date());
}

/**
 * Debounced auto-publish: any edit (re)schedules a single publish ~2 min later,
 * so a burst of edits produces one GitHub commit instead of one per keystroke.
 * This is the handler for the INSTALLABLE onEdit trigger created by
 * enableAutoPublish() below — a simple onEdit(e) trigger can't create other
 * triggers, so it must be installed this way rather than just naming a
 * function "onEdit".
 */
function onEditDebounced_(e) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'publishToGithub') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('publishToGithub')
    .timeBased()
    .after(2 * 60 * 1000)
    .create();
}

function enableAutoPublish() {
  disableAutoPublish(); // clear any existing installable trigger first
  ScriptApp.newTrigger('onEditDebounced_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('Auto-publish enabled — edits will publish to GitHub ~2 minutes later.');
}

function disableAutoPublish() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'publishToGithub' || fn === 'onEditDebounced_') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert('Auto-publish disabled. Use "Publish now" manually.');
}
