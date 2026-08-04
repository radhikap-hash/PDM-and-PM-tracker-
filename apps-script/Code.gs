/**
 * PDM & PM tracker → GitHub Pages publisher
 * -----------------------------------------
 * Bind to the "PDM & PM" Google Sheet (Extensions > Apps Script).
 *
 * READS:
 *  - "Project details" — master list of variants: Category, Name, Project ID,
 *    Stage/Status (free text), PDM, PM, Start/End Date. This decides each
 *    variant's overall category and CURRENT stage number.
 *  - "Stages/documents" — the canonical Stage | Activity | Document | Owner
 *    template (11 stages, one row per document). Used to build the dashboard's
 *    stage list and per-stage document checklist, so the sheet — not the
 *    HTML — is the single source of truth for both.
 *  - The five per-owner tabs (Sensor-Durga, Drone-Bharath, Praddumna- Camera,
 *    Delivery bot-Swapnil, Server-Tarun) — each repeats the SAME Stage|
 *    Activity|Document|Owner template in columns A-D, then has one 8-column
 *    block per variant to the right:
 *      Planned Start | Due Date | Completion date | Status | Key milestones |
 *      Blocker | Immediate next step | Documents
 *    row-aligned to the template, with the variant's "Name (CODE)" as a
 *    merged header directly above the column labels. One row = one specific
 *    document deliverable for that variant.
 *
 * BUILDS data.json with:
 *   stages       — array of stage names (from Stages/documents)
 *   stageDocs    — {stageNum: [documentName, ...]} (from Stages/documents)
 *   products     — one entry per variant (from Project details)
 *   engineering  — engineering[id][stageNum] = rolled-up {plannedStart,
 *                  dueDate, completedDate, status, milestone, blocker,
 *                  nextStep} — aggregated across that stage's document rows
 *   documents    — documents[id] = [{stage, document, plannedStart, dueDate,
 *                  completedDate, status, link}, ...] — one entry per real
 *                  document row, for the per-document checklist detail
 *
 * ONE-TIME SETUP
 * 1. Paste this whole file in as Code.gs (Extensions > Apps Script).
 * 2. Project Settings > Script Properties > Add script property, then SAVE:
 *      GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, GITHUB_BRANCH (see bottom).
 * 3. Run setup() once from the editor toolbar, approve the permissions prompt.
 * 4. Reload the Sheet — a "Dashboard" menu appears.
 * 5. Dashboard ▸ Preview data (no publish) FIRST — this is the fast way to
 *    see if the owner-tab parsing lined up with your actual columns before
 *    anything gets published. Read the warnings carefully.
 * 6. Dashboard ▸ Publish now.
 *
 * IMPORTANT CAVEAT: the owner-tab block parser below was written from two
 * screenshots (Sensor-Durga tab), not by reading your other four tabs
 * directly. If Drone-Bharath / Praddumna- Camera / Delivery bot-Swapnil /
 * Server-Tarun use a different column layout, Preview data's warnings will
 * say so — report them back and the parser gets adjusted, rather than
 * silently publishing wrong data.
 */

const PROJECT_SHEET_NAMES = ['Project details', 'Products'];
const TEMPLATE_SHEET_NAMES = ['Stages/documents', 'Stages / documents'];
const OWNER_SHEET_NAMES = ['Sensor-Durga', 'Drone-Bharath', 'Praddumna- Camera', 'Delivery bot-Swapnil', 'Server-Tarun'];

const BLOCK_FIELD_CANDIDATES = {
  plannedStart: ['planned start'],
  dueDate: ['due date'],
  completedDate: ['completion date', 'completed date'],
  status: ['status'],
  milestone: ['key milestone', 'milestone'],
  blocker: ['blocker'],
  nextStep: ['immediate next step', 'next step'],
  document: ['document']
};

function setup() {
  onOpen();
  SpreadsheetApp.getUi().alert('Menu installed. Set Script Properties (see comment block at top of Code.gs), then Dashboard \u25b8 Preview data, then Dashboard \u25b8 Publish now.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Preview data (no publish)', 'previewData')
    .addItem('Publish now', 'publishToGithubUI')
    .addItem('Enable auto-publish (~2 min after edits)', 'enableAutoPublish')
    .addItem('Disable auto-publish', 'disableAutoPublish')
    .addToUi();
}

function str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function dateStr_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Etc/UTC', 'd-MMM-yyyy');
  }
  return str_(v) || null;
}
function findSheet_(ss, names) {
  for (const n of names) { const sh = ss.getSheetByName(n); if (sh) return sh; }
  return null;
}
function findCol_(headerRow, candidates) {
  const norm = headerRow.map(h => str_(h).toLowerCase());
  for (const cand of candidates) {
    const idx = norm.findIndex(h => h.includes(cand));
    if (idx !== -1) return idx;
  }
  return -1;
}
function normDoc_(s) { return str_(s).toLowerCase(); }
function normName_(s) { return str_(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

/**
 * Parses the Stage|Activity|Document|Owner template starting in column A of
 * the given sheet. Returns {templateRows: [{rowIndex(0-based sheet row),
 * stage, activity, document}], headerRowIdx}.
 * Stage and Activity are forward-filled down blank cells (matches how the
 * sheet groups multiple document rows under one stage/activity).
 */
function parseTemplate_(values) {
  let headerRowIdx = -1;
  for (let i = 0; i < values.length; i++) {
    if (str_(values[i][0]).toLowerCase() === 'stage') { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) return { templateRows: [], headerRowIdx: -1 };

  const templateRows = [];
  let lastStage = null, lastActivity = '';
  let blankStreak = 0;
  for (let i = headerRowIdx + 1; i < values.length; i++) {
    const row = values[i];
    const stageCell = str_(row[0]);
    const activityCell = str_(row[1]);
    const docCell = str_(row[2]);
    const ownerCell = str_(row[3]);

    if (!stageCell && !activityCell && !docCell && !ownerCell) {
      blankStreak++;
      if (blankStreak >= 3) break; // end of template
      continue;
    }
    blankStreak = 0;

    const m = stageCell.match(/^stage\s*(\d+)$/i);
    if (m) lastStage = parseInt(m[1], 10);
    if (activityCell) lastActivity = activityCell;

    if (!docCell) continue; // a row with a stage/activity label but no document isn't a checklist item
    templateRows.push({ rowIndex: i, stage: lastStage, activity: lastActivity, document: docCell });
  }
  return { templateRows, headerRowIdx };
}

/** Reads Stages/documents to build the canonical stages[] list and stageDocs{} checklist. */
// Hardcoded fallback — the definitive 10-stage list you confirmed directly
// (not guessed from a screenshot this time). Used whenever the "Stages/documents"
// tab can't be found under any of the names above, so stage-name matching
// (e.g. "Procurement & Development" -> stage 9) still works instead of
// silently defaulting everything to stage 1.
const HARDCODED_STAGES = [
  'Requirement Analysis','Proposal','Commercial Finalization','Statement of Work (SOW)','Project Kickoff',
  'Project Planning','Detailed Design','Critical Design Review (CDR)',
  'Procurement & Development','Integration, Verification & Validation'
];
const HARDCODED_STAGE_DOCS = {
  1: ["MRD","PRD","PRD Matrix"],
  2: ["Proposal Presentation (PPT)","NDA","Quotation (Price Breakup, Commercial Terms, Payment Schedule, Delivery Schedule)"],
  3: ["Final SOW (Scope Matrix, Deliverables List, Exclusions List, Assumptions, Dependencies, Acceptance Criteria)"],
  4: ["Kickoff Presentation","Meeting Minutes (MoM)","RACI Matrix"],
  5: ["Work Breakdown Structure (WBS)","PDM Document","Hardware Architecture","Software Architecture","Block Diagram","Data Flow Diagram (DFD)","Interface Control Document (ICD)","Milestone Plan","Resource Plan","Budget Tracker","Procurement Plan","Risk Register"],
  6: ["Schematics","PCB Layout","Bill of Materials (BOM)","PCB Stack-up","Simulation Reports","Software Design Document (SDD)","API Specification","Database Design","UI Mockups","CAD Models","Mechanical Drawings","Thermal Analysis Report"],
  7: ["CDR Package","Design Review Presentation","Simulation Results","Review Minutes","Risk Closure Report","Manufacturing Readiness Report","Design Freeze Approval"],
  8: ["Approved Vendor List (AVL)","RFQ","Purchase Request (PR)","Purchase Order (PO)","Delivery Tracker","Incoming Inspection Report","PCB Assembly Report","Bring-up Report","Debug Logs","Source Code Repository","Code Review Report","Unit Test Report","Dataset","AI Training Report","Sprint Board","Bug Tracker","Version Control Log"],
  9: ["Verification Plan","Validation Plan","Test Cases","Test Reports","Defect Log","Regression Test Report","Requirements Traceability Matrix (RTM)","Customer Acceptance Test (CAT/UAT) Report","Final Validation Report"],
  10: []
};

function readStagesAndDocs_(ss, warnings) {
  const sh = findSheet_(ss, TEMPLATE_SHEET_NAMES);
  if (!sh) {
    const actual = ss.getSheets().map(s => s.getName()).join(', ');
    warnings.push(`No tab named "Stages/documents" found (tried: ${TEMPLATE_SHEET_NAMES.join(', ')}) — using a hardcoded fallback stage list instead so status matching still works. Actual tabs in this spreadsheet: ${actual}. If one of those is meant to be the Stages/documents tab, tell me its exact name and I'll add it.`);
    return { stages: HARDCODED_STAGES.slice(), stageDocs: HARDCODED_STAGE_DOCS };
  }
  const { templateRows } = parseTemplate_(sh.getDataRange().getValues());
  if (!templateRows.length) {
    warnings.push(`"${sh.getName()}" tab was found but no rows parsed from it (expected a header row with "Stage" in column A, then Stage/Activity/Document/Owner rows below) — using the hardcoded fallback stage list instead.`);
    return { stages: HARDCODED_STAGES.slice(), stageDocs: HARDCODED_STAGE_DOCS };
  }
  const stages = [];
  const stageDocs = {};
  templateRows.forEach(r => {
    if (!r.stage) return;
    if (!stages[r.stage - 1] && r.activity) stages[r.stage - 1] = r.activity;
    if (!stageDocs[r.stage]) stageDocs[r.stage] = [];
    stageDocs[r.stage].push(r.document);
  });
  // fill any stage-name gaps (a stage with docs but no activity text recorded on its first row)
  for (let i = 0; i < stages.length; i++) if (!stages[i]) stages[i] = `Stage ${i+1}`;
  return { stages: stages.filter(Boolean), stageDocs };
}

/** Extracts "SOS-1" from "SOS Band (SOS-1)"; returns {name, code}. */
function parseVariantHeader_(raw) {
  const s = str_(raw);
  const m = s.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
  if (m) return { name: m[1].trim(), code: m[2].trim() };
  return { name: s, code: s };
}

/**
 * Parses one owner tab: its own Stage|Activity|Document|Owner template (cols
 * A-D) for row alignment, then every 8-column variant block to the right.
 * Returns an array of {name, code, entries: [{stage, document, plannedStart,
 * dueDate, completedDate, status, milestone, blocker, nextStep, link}, ...]}
 * — one per variant block found. name/code come straight off the block's
 * own "Name (CODE)" header; resolving that to a Project details id happens
 * in buildPayload_ (by name first, since owner-tab codes like SEN-1/BAN-1/
 * SOS-1 don't match Project details' SEN-001..SEN-004 numbering).
 */
function parseOwnerTab_(sheet, warnings) {
  const values = sheet.getDataRange().getValues();
  const { templateRows, headerRowIdx } = parseTemplate_(values);
  if (!templateRows.length) {
    warnings.push(`Tab "${sheet.getName()}": couldn't find the Stage/Activity/Document/Owner template in columns A-D (expected a header row with "Stage" in column A).`);
    return [];
  }
  const variantHeaderRow = values[headerRowIdx - 1] || [];
  const subHeaderRow = values[headerRowIdx] || [];

  // Find each variant block's starting column: any non-empty cell in the row
  // ABOVE the Stage/Activity/Document/Owner header, at column E (index 4) or later.
  const blockStarts = [];
  for (let c = 4; c < variantHeaderRow.length; c++) {
    if (str_(variantHeaderRow[c])) blockStarts.push(c);
  }
  if (!blockStarts.length) {
    warnings.push(`Tab "${sheet.getName()}": no variant blocks found (expected a merged "Name (CODE)" header in row ${headerRowIdx} above the Stage/Activity/Document/Owner row, starting at column E).`);
    return [];
  }

  const blocks = [];
  blockStarts.forEach((startCol, bi) => {
    const endCol = (bi + 1 < blockStarts.length) ? blockStarts[bi + 1] : Math.min(startCol + 10, subHeaderRow.length);
    const { name, code } = parseVariantHeader_(variantHeaderRow[startCol]);

    // locate each field within [startCol, endCol) by header text
    const localHeaders = subHeaderRow.slice(startCol, endCol);
    const fieldCol = {};
    Object.keys(BLOCK_FIELD_CANDIDATES).forEach(field => {
      const idx = findCol_(localHeaders, BLOCK_FIELD_CANDIDATES[field]);
      if (idx !== -1) fieldCol[field] = startCol + idx;
    });
    if (Object.keys(fieldCol).length < 4) {
      warnings.push(`Tab "${sheet.getName()}", block "${str_(variantHeaderRow[startCol])}": only recognized ${Object.keys(fieldCol).length}/8 expected columns (Planned Start/Due Date/Completion date/Status/Key milestones/Blocker/Immediate next step/Documents) — check for renamed headers.`);
    }

    const entries = [];
    templateRows.forEach(t => {
      const row = values[t.rowIndex];
      const get = (f) => fieldCol[f] !== undefined ? row[fieldCol[f]] : '';
      const plannedStart = dateStr_(get('plannedStart'));
      const dueDate = dateStr_(get('dueDate'));
      const completedDate = dateStr_(get('completedDate'));
      const status = str_(get('status')) || null;
      const milestone = str_(get('milestone')) || null;
      const blocker = str_(get('blocker')) || null;
      const nextStep = str_(get('nextStep')) || null;
      const docCell = fieldCol.document !== undefined ? row[fieldCol.document] : '';
      const document = str_(docCell) || null;
      let link = null;
      try {
        if (fieldCol.document !== undefined) {
          const rich = sheet.getRange(t.rowIndex + 1, fieldCol.document + 1).getRichTextValue();
          if (rich) link = rich.getLinkUrl();
        }
      } catch (e) { /* no rich text / no link — fine */ }

      if (!plannedStart && !dueDate && !completedDate && !status && !milestone && !blocker && !nextStep && !document) return; // nothing filled in for this doc row

      entries.push({
        stage: t.stage, document: t.document,
        plannedStart, dueDate, completedDate, status, milestone, blocker, nextStep, document_text: document, link
      });
    });

    if (entries.length) blocks.push({ name, code, entries });
  });
  return blocks;
}

/** Rolls up a variant's per-document rows into one summary record per stage. */
function rollupByStage_(entries) {
  const byStage = {};
  entries.forEach(e => {
    if (!e.stage) return;
    if (!byStage[e.stage]) byStage[e.stage] = [];
    byStage[e.stage].push(e);
  });
  const out = {};
  Object.keys(byStage).forEach(stageNum => {
    const rows = byStage[stageNum];
    const dates = arr => rows.map(r=>r[arr]).filter(Boolean);
    const allCompleted = rows.length > 0 && rows.every(r => /complet/i.test(r.status||'') || r.completedDate);
    const inProgressRow = rows.find(r => r.status && !/complet/i.test(r.status)) || rows.find(r => !r.completedDate) || rows[0];
    out[stageNum] = {
      plannedStart: dates('plannedStart').sort()[0] || null,
      dueDate: dates('dueDate').sort().slice(-1)[0] || null,
      completedDate: allCompleted ? (dates('completedDate').sort().slice(-1)[0] || null) : null,
      status: (inProgressRow && inProgressRow.status) || (allCompleted ? 'Completed' : null),
      milestone: Array.from(new Set(rows.map(r=>r.milestone).filter(Boolean))).join('; ') || null,
      blocker: Array.from(new Set(rows.map(r=>r.blocker).filter(Boolean))).join('; ') || null,
      nextStep: (inProgressRow && inProgressRow.nextStep) || Array.from(new Set(rows.map(r=>r.nextStep).filter(Boolean))).join('; ') || null
    };
  });
  return out;
}

function buildPayload_() {
  const ss = SpreadsheetApp.getActive();
  const warnings = [];

  // --- Project details: master variant list ---
  const projSheet = findSheet_(ss, PROJECT_SHEET_NAMES);
  if (!projSheet) {
    const actual = ss.getSheets().map(s => s.getName()).join(', ');
    throw new Error(`No sheet named "Project details" or "Products" found. Tabs in this spreadsheet: ${actual}`);
  }
  const pdValues = projSheet.getDataRange().getValues();
  let pdHeaderIdx = -1;
  for (let i = 0; i < pdValues.length; i++) {
    if (str_(pdValues[i][0]).toLowerCase() === 'category') { pdHeaderIdx = i; break; }
  }
  if (pdHeaderIdx === -1) throw new Error(`Could not find the header row in "${projSheet.getName()}" (looking for "Category" in column A).`);
  const pdHeader = pdValues[pdHeaderIdx];
  const colCat = findCol_(pdHeader, ['category']);
  const colName = findCol_(pdHeader, ['product / solution name', 'solution name', 'product name', 'name']);
  const colId = findCol_(pdHeader, ['project id', 'id']);
  const colStage = findCol_(pdHeader, ['stage / status', 'stage', 'status']);
  const colPdm = findCol_(pdHeader, ['product manager', 'pdm']);
  const colPm = findCol_(pdHeader, ['program manager', 'project manager', 'pm']);
  const colStart = findCol_(pdHeader, ['start date']);
  const colEnd = findCol_(pdHeader, ['end date']);

  const { stages, stageDocs } = readStagesAndDocs_(ss, warnings);
  const stageNameToNum = {};
  stages.forEach((name, idx) => { stageNameToNum[name.toLowerCase()] = idx + 1; });

  function matchStage(text) {
    const t = str_(text).toLowerCase();
    if (!t) return { stage: 1, matched: false };
    if (stageNameToNum[t] !== undefined) return { stage: stageNameToNum[t], matched: true };
    for (const name in stageNameToNum) {
      if (t.includes(name) || name.includes(t)) return { stage: stageNameToNum[name], matched: true };
    }
    return { stage: 1, matched: false };
  }
  function extractId(raw) {
    const s = str_(raw);
    const m = s.match(/\(([^)]+)\)\s*$/);
    return m ? m[1].trim() : s;
  }

  const products = [];
  let lastCat = '';
  for (let i = pdHeaderIdx + 1; i < pdValues.length; i++) {
    const row = pdValues[i];
    const idRaw = colId !== -1 ? row[colId] : '';
    const nameRaw = colName !== -1 ? row[colName] : '';
    if (!str_(idRaw) && !str_(nameRaw)) {
      if (str_(row[0]).toLowerCase() === 'stages') break;
      const nextBlank = (i + 1 < pdValues.length) && !str_(pdValues[i+1][colId !== -1 ? colId : 0]);
      if (nextBlank) break;
      continue;
    }
    const catCell = colCat !== -1 ? str_(row[colCat]) : '';
    if (catCell) lastCat = catCell;
    const statusText = colStage !== -1 ? str_(row[colStage]) : '';
    const sm = matchStage(statusText);
    if (statusText && !sm.matched) warnings.push(`Project details row ${i+1}: status "${statusText}" didn't match any stage name — defaulted to stage 1.`);
    const id = extractId(idRaw);
    if (!id) warnings.push(`Project details row ${i+1} ("${str_(nameRaw)}"): no Project ID.`);
    products.push({
      cat: lastCat || 'Uncategorized', name: str_(nameRaw) || id, id: id,
      pdm: colPdm !== -1 ? (str_(row[colPdm]) || null) : null,
      pm: colPm !== -1 ? (str_(row[colPm]) || null) : null,
      stage: sm.stage,
      plannedStart: colStart !== -1 ? dateStr_(row[colStart]) : null,
      plannedEnd: colEnd !== -1 ? dateStr_(row[colEnd]) : null,
      statusLabel: statusText || null
    });
  }
  if (!products.length) warnings.push('No product rows were read from Project details at all.');
  const knownIds = products.map(p => p.id);
  const nameToId = {};
  products.forEach(p => { nameToId[normName_(p.name)] = p.id; });

  // --- Owner tabs: per-variant, per-document tracking ---
  const documents = {};
  const engineering = {};
  let anyOwnerTabFound = false;
  OWNER_SHEET_NAMES.forEach(tabName => {
    const sh = ss.getSheetByName(tabName);
    if (!sh) { warnings.push(`Owner tab "${tabName}" not found — skipped. (Tabs present: ${ss.getSheets().map(s=>s.getName()).join(', ')})`); return; }
    anyOwnerTabFound = true;
    const blocks = parseOwnerTab_(sh, warnings);
    blocks.forEach(block => {
      // Owner-tab codes (SEN-1, BAN-1, SOS-1, ...) are a different numbering
      // system than Project details' Project IDs (SEN-001..SEN-004) — match
      // by the block's NAME against Project details' product name first.
      let id = nameToId[normName_(block.name)];
      if (!id && knownIds.indexOf(block.code) !== -1) id = block.code; // exact code match, if it happens to line up
      if (!id) {
        warnings.push(`Tab "${tabName}": block "${block.name} (${block.code})" didn't match any Project details product by name or code — its data won't show up on the dashboard until the name matches exactly, or the code is fixed.`);
        return;
      }
      if (!documents[id]) documents[id] = [];
      documents[id] = documents[id].concat(block.entries.map(e => ({
        stage: e.stage, document: e.document, plannedStart: e.plannedStart, dueDate: e.dueDate,
        completedDate: e.completedDate, status: e.status, link: e.link
      })));
      const rollup = rollupByStage_(block.entries);
      if (!engineering[id]) engineering[id] = {};
      Object.keys(rollup).forEach(stageNum => { engineering[id][stageNum] = rollup[stageNum]; });
    });
  });
  if (!anyOwnerTabFound) warnings.push('None of the expected owner tabs were found — no milestones/blockers/next steps/documents will show for any variant.');

  return {
    payload: {
      generatedAt: new Date().toISOString(),
      stages: stages,
      stageDocs: stageDocs,
      products: products,
      engineering: engineering,
      documents: documents
    },
    warnings: warnings
  };
}

function previewData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const { payload, warnings } = buildPayload_();
    const catCounts = {};
    payload.products.forEach(p => { catCounts[p.cat] = (catCounts[p.cat]||0) + 1; });
    const catSummary = Object.keys(catCounts).map(c => `${c}: ${catCounts[c]}`).join('\n');
    const docCount = Object.keys(payload.documents).reduce((n,id)=>n+payload.documents[id].length,0);
    const withoutDocs = payload.products.filter(p => !payload.documents[p.id]).map(p => p.id);
    let msg = `Read ${payload.products.length} products across ${Object.keys(catCounts).length} categories:\n${catSummary}\n\n`
      + `${payload.stages.length} stages, ${Object.keys(payload.stageDocs).length} stages with documents defined.\n`
      + `${docCount} document rows read across ${Object.keys(payload.documents).length} variants.`;
    if (withoutDocs.length) msg += `\n\nVariants with NO document rows at all: ${withoutDocs.join(', ')}`;
    if (warnings.length) msg += `\n\n\u26a0 ${warnings.length} warning(s):\n- ` + warnings.slice(0,10).join('\n- ');
    Logger.log(JSON.stringify({ payload, warnings }, null, 2));
    ui.alert('Preview (nothing published)', msg, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log('Preview failed: ' + e.stack);
    ui.alert('Preview failed', e.message, ui.ButtonSet.OK);
  }
}

function publishToGithub() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER');
  const repo = props.getProperty('GITHUB_REPO');
  const branch = props.getProperty('GITHUB_BRANCH') || 'main';
  const path = props.getProperty('GITHUB_PATH') || 'data.json';
  if (!token || !owner || !repo) {
    throw new Error('GITHUB_TOKEN, GITHUB_OWNER and/or GITHUB_REPO are missing from Script Properties.');
  }
  const { payload, warnings } = buildPayload_();
  const content = Utilities.base64Encode(JSON.stringify(payload, null, 2), Utilities.Charset.UTF_8);
  const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  const headers = { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json' };
  let sha = null;
  const getResp = UrlFetchApp.fetch(`${apiUrl}?ref=${branch}`, { headers, muteHttpExceptions: true });
  if (getResp.getResponseCode() === 200) sha = JSON.parse(getResp.getContentText()).sha;
  else if (getResp.getResponseCode() !== 404) throw new Error(`GitHub GET failed (${getResp.getResponseCode()}): ${getResp.getContentText()}`);
  const body = { message: `Publish dashboard data \u2014 ${new Date().toISOString()}`, content, branch };
  if (sha) body.sha = sha;
  const putResp = UrlFetchApp.fetch(apiUrl, { method: 'put', headers, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
  const code = putResp.getResponseCode();
  if (code !== 200 && code !== 201) throw new Error(`GitHub PUT failed (${code}): ${putResp.getContentText()}`);
  Logger.log('Published \u2014 %s products, %s doc rows, %s warnings', payload.products.length, Object.keys(payload.documents).reduce((n,id)=>n+payload.documents[id].length,0), warnings.length);
  return { payload, warnings };
}

function publishToGithubUI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const { payload, warnings } = publishToGithub();
    const docCount = Object.keys(payload.documents).reduce((n,id)=>n+payload.documents[id].length,0);
    let msg = `Published ${payload.products.length} products and ${docCount} document rows to GitHub.`;
    if (warnings.length) msg += `\n\n\u26a0 ${warnings.length} warning(s) \u2014 see Executions log for the full list:\n- ` + warnings.slice(0,5).join('\n- ');
    ui.alert('Publish succeeded', msg, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log('Publish failed: ' + e.stack);
    ui.alert('Publish failed', e.message, ui.ButtonSet.OK);
  }
}

function onEditDebounced_(e) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'publishToGithub') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('publishToGithub').timeBased().after(2 * 60 * 1000).create();
}
function enableAutoPublish() {
  disableAutoPublish();
  ScriptApp.newTrigger('onEditDebounced_').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getUi().alert('Auto-publish enabled \u2014 edits will publish to GitHub ~2 minutes later.');
}
function disableAutoPublish() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'publishToGithub' || fn === 'onEditDebounced_') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert('Auto-publish disabled. Use "Publish now" manually.');
}
function testBuildPayload() { Logger.log(JSON.stringify(buildPayload_(), null, 2)); }
