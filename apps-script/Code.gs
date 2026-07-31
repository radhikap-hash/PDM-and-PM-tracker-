/**
 * PDM & PM tracker → GitHub Pages publisher
 * -----------------------------------------
 * Bind this script to the "PDM & PM" Google Sheet.
 * Reads the "Project details" tab (the master product/status table) and the
 * "Engineering" tab (weekly updates), builds data.json in the shape
 * dashboard.html expects, and commits/updates it in the GitHub repo via the
 * Contents API.
 *
 * Column detection is header-driven (matches by column NAME, not position),
 * so reordering a column in the sheet won't break the sync.
 *
 * ONE-TIME SETUP
 * 1. Extensions > Apps Script, paste this file in as Code.gs (replacing
 *    whatever was there before — do not hand-edit the GITHUB_TOKEN /
 *    GITHUB_OWNER / GITHUB_REPO lines in the code itself, they read from
 *    Script Properties, see step 2).
 * 2. Project Settings > Script Properties, add:
 *      GITHUB_TOKEN   = a fine-grained PAT with "Contents: read & write"
 *                       scoped to this one repo
 *      GITHUB_OWNER   = radhikap-hash
 *      GITHUB_REPO    = PDM-and-PM-tracker-
 *      GITHUB_BRANCH  = main
 *      GITHUB_PATH    = data.json
 * 3. Run setup() once from the Apps Script editor toolbar (approve the
 *    permissions prompt when asked).
 * 4. Reload the Sheet — a "Dashboard" menu appears with "Publish now".
 *
 * SHEET LAYOUT THIS EXPECTS (matches your actual "PDM & PM" sheet)
 *
 * Tab "Project details" — one row per product variant:
 *   Category | Product / Solution Name | Project ID | Type | Stage / Status |
 *   Product Manager | Project Manager | Start Date | End Date
 *   - Category only needs to be filled on the first row of each group —
 *     blank cells below it inherit the category above (like your sheet does).
 *   - Project ID can be a compound string like "ZMD-DLD-AI (CAM-001)" — the
 *     short code in parentheses is pulled out and used as the product's id
 *     (must match the ids used as keys in the Engineering tab and in
 *     dashboard.html's CAM_ENGINEERING). If there's no parenthetical, the
 *     whole cell is used as-is (e.g. "NVR-001").
 *   - Stage / Status is free text (e.g. "Procurement & Development",
 *     "Ordered for learning"). It's matched against the sheet's own
 *     "Stages" legend (see below); unmatched text is kept verbatim as
 *     statusLabel and shown in the dashboard, defaulting to stage 1 so it
 *     still has *some* position on the progress bar.
 *
 * "Stages" legend — anywhere below the product table in the SAME tab:
 *   a row with "Stages" in column A, then rows "Stage 1" / "Stage 2" / ...
 *   in column A with the stage name in column B. This becomes the
 *   dashboard's canonical stage list — edit it here, not in the HTML.
 *
 * Tab "Engineering" — one row per weekly update (latest row per Product ID
 * wins), columns detected by header name, expected headers close to:
 *   Product ID | Period | Owners | Platform | Approach | Completed |
 *   Next Steps | Risk | Status | ETA
 *   - Completed / Next Steps: one bullet per line in the cell (Alt+Enter),
 *     or separated by " | ".
 */

const PROJECT_SHEET_NAMES = ['Project details', 'Products'];
const ENGINEERING_SHEET_NAMES = ['Engineering'];

function setup() {
  onOpen();
  SpreadsheetApp.getUi().alert('Menu installed. Fill in Script Properties (see the comment at the top of Code.gs), then use Dashboard \u25b8 Publish now.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Publish now', 'publishToGithub')
    .addItem('Enable auto-publish (~2 min after edits)', 'enableAutoPublish')
    .addItem('Disable auto-publish', 'disableAutoPublish')
    .addToUi();
}

function findSheet_(ss, names) {
  for (const n of names) {
    const sh = ss.getSheetByName(n);
    if (sh) return sh;
  }
  return null;
}

function str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function splitList_(v) {
  const s = str_(v);
  if (!s) return [];
  return s.split(/\n|\s\|\s/).map(x => x.trim()).filter(Boolean);
}

function findCol_(headerRow, candidates) {
  const norm = headerRow.map(h => str_(h).toLowerCase());
  for (const cand of candidates) {
    const idx = norm.findIndex(h => h.includes(cand));
    if (idx !== -1) return idx;
  }
  return -1;
}

function readStageLegend_(values) {
  let legendRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (str_(values[i][0]).toLowerCase() === 'stages') { legendRow = i; break; }
  }
  const stages = [];
  const nameToNum = {};
  if (legendRow === -1) return { stages, nameToNum };
  for (let i = legendRow + 1; i < values.length; i++) {
    const label = str_(values[i][0]);
    const m = label.match(/^stage\s*(\d+)$/i);
    if (!m) { if (label) break; else continue; }
    const num = parseInt(m[1], 10);
    const name = str_(values[i][1]);
    if (!name) continue;
    stages[num - 1] = name;
    nameToNum[name.toLowerCase()] = num;
  }
  return { stages: stages.filter(Boolean), nameToNum };
}

function matchStage_(text, nameToNum) {
  const t = str_(text).toLowerCase();
  if (!t) return { stage: 1, matched: false };
  if (nameToNum[t] !== undefined) return { stage: nameToNum[t], matched: true };
  for (const name in nameToNum) {
    if (t.includes(name) || name.includes(t)) return { stage: nameToNum[name], matched: true };
  }
  return { stage: 1, matched: false };
}

function extractId_(raw) {
  const s = str_(raw);
  const m = s.match(/\(([^)]+)\)\s*$/);
  return m ? m[1].trim() : s;
}

function buildPayload_() {
  const ss = SpreadsheetApp.getActive();

  const projSheet = findSheet_(ss, PROJECT_SHEET_NAMES);
  if (!projSheet) throw new Error('No "Project details" (or "Products") sheet found.');
  const values = projSheet.getDataRange().getValues();

  let headerRowIdx = -1;
  for (let i = 0; i < values.length; i++) {
    if (str_(values[i][0]).toLowerCase() === 'category') { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) throw new Error('Could not find the header row (looking for "Category" in column A).');
  const header = values[headerRowIdx];

  const colCat   = findCol_(header, ['category']);
  const colName  = findCol_(header, ['product / solution name', 'solution name', 'product name', 'name']);
  const colId    = findCol_(header, ['project id', 'id']);
  const colStage = findCol_(header, ['stage / status', 'stage', 'status']);
  const colPdm   = findCol_(header, ['product manager', 'pdm']);
  const colPm    = findCol_(header, ['project manager', 'pm']);

  const { stages, nameToNum } = readStageLegend_(values);

  const products = [];
  let lastCat = '';
  for (let i = headerRowIdx + 1; i < values.length; i++) {
    const row = values[i];
    const idRaw = colId !== -1 ? row[colId] : '';
    const nameRaw = colName !== -1 ? row[colName] : '';
    if (!str_(idRaw) && !str_(nameRaw)) {
      if (str_(row[0]).toLowerCase() === 'stages') break;
      const nextBlank = (i + 1 < values.length) && !str_(values[i+1][colId !== -1 ? colId : 0]);
      if (nextBlank) break;
      continue;
    }
    const catCell = colCat !== -1 ? str_(row[colCat]) : '';
    if (catCell) lastCat = catCell;

    const statusText = colStage !== -1 ? str_(row[colStage]) : '';
    const { stage, matched } = matchStage_(statusText, nameToNum);

    products.push({
      cat: lastCat || 'Uncategorized',
      name: str_(nameRaw) || str_(idRaw),
      id: extractId_(idRaw),
      pdm: colPdm !== -1 ? (str_(row[colPdm]) || null) : null,
      pm: colPm !== -1 ? (str_(row[colPm]) || null) : null,
      stage: stage,
      daysStage: 0,
      daysTotal: 0,
      doc: 'TBD',
      blocker: null,
      tentative: null,
      statusLabel: statusText || null
    });
  }

  const engSheet = findSheet_(ss, ENGINEERING_SHEET_NAMES);
  const engineering = {};
  if (engSheet) {
    const evals = engSheet.getDataRange().getValues();
    if (evals.length > 1) {
      const ehead = evals[0];
      const eColId = findCol_(ehead, ['product id', 'id']);
      const eColPeriod = findCol_(ehead, ['period']);
      const eColOwners = findCol_(ehead, ['owner']);
      const eColPlatform = findCol_(ehead, ['platform']);
      const eColApproach = findCol_(ehead, ['approach']);
      const eColCompleted = findCol_(ehead, ['completed']);
      const eColNext = findCol_(ehead, ['next']);
      const eColRisk = findCol_(ehead, ['risk']);
      const eColStatus = findCol_(ehead, ['status']);
      const eColEta = findCol_(ehead, ['eta']);

      for (let i = 1; i < evals.length; i++) {
        const row = evals[i];
        const id = eColId !== -1 ? str_(row[eColId]) : '';
        if (!id) continue;
        engineering[id] = {
          period: eColPeriod !== -1 ? str_(row[eColPeriod]) : '',
          owners: (eColOwners !== -1 && str_(row[eColOwners])) || 'Not specified',
          platform: (eColPlatform !== -1 && str_(row[eColPlatform])) || 'Not specified',
          approach: (eColApproach !== -1 && str_(row[eColApproach])) || 'Not specified',
          completed: eColCompleted !== -1 ? splitList_(row[eColCompleted]) : [],
          next: eColNext !== -1 ? splitList_(row[eColNext]) : [],
          risk: (eColRisk !== -1 && str_(row[eColRisk])) || null,
          status: (eColStatus !== -1 && str_(row[eColStatus])) || 'In Progress',
          eta: (eColEta !== -1 && str_(row[eColEta])) || null
        };
      }
    }
  }

  return {
    generatedAt: new Date().toISOString(),
    stages: stages,
    products: products,
    engineering: engineering
  };
}

function publishToGithub() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER');
  const repo = props.getProperty('GITHUB_REPO');
  const branch = props.getProperty('GITHUB_BRANCH') || 'main';
  const path = props.getProperty('GITHUB_PATH') || 'data.json';

  if (!token || !owner || !repo) {
    throw new Error('Set GITHUB_TOKEN, GITHUB_OWNER and GITHUB_REPO in Script Properties first (Project Settings \u25b8 Script Properties \u2014 not in the code).');
  }

  const payload = buildPayload_();
  const content = Utilities.base64Encode(JSON.stringify(payload, null, 2), Utilities.Charset.UTF_8);
  const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/vnd.github+json'
  };

  let sha = null;
  const getResp = UrlFetchApp.fetch(`${apiUrl}?ref=${branch}`, { headers, muteHttpExceptions: true });
  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  }

  const body = {
    message: `Publish dashboard data \u2014 ${new Date().toISOString()}`,
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
  Logger.log('Published data.json \u2014 %s (%s products)', new Date(), payload.products.length);
}

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
  disableAutoPublish();
  ScriptApp.newTrigger('onEditDebounced_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('Auto-publish enabled \u2014 edits will publish to GitHub ~2 minutes later.');
}

function disableAutoPublish() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'publishToGithub' || fn === 'onEditDebounced_') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert('Auto-publish disabled. Use "Publish now" manually.');
}

/** Handy for debugging from the Apps Script editor: logs the payload without publishing. */
function testBuildPayload() {
  Logger.log(JSON.stringify(buildPayload_(), null, 2));
}
