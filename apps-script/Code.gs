/**
 * PDM & PM Tracker → GitHub Pages Publisher
 * ==========================================
 *
 * Google Sheet → parsed project data → data.json → GitHub Pages
 *
 * IMPORTANT:
 * 1. Keep the GitHub token in Apps Script > Project Settings > Script Properties.
 * 2. Run setup() once.
 * 3. Run Dashboard > Preview data first.
 * 4. Then Dashboard > Publish now.
 *
 * Script Properties:
 *   GITHUB_TOKEN
 *   GITHUB_OWNER
 *   GITHUB_REPO
 *   GITHUB_BRANCH
 *   GITHUB_PATH (optional; defaults to data.json)
 */

const PROJECT_SHEET_NAMES = ['Project details', 'Products'];
const TEMPLATE_SHEET_NAMES = ['Stages/documents', 'Stages / documents'];
const OWNER_SHEET_NAMES = [
  'Sensor-Durga',
  'Drone-Bharath',
  'Praddumna- Camera',
  'Delivery bot-Swapnil',
  'Server-Tarun'
];

const SHEET_DASHBOARD_TAB = 'Live Dashboard';
const GITHUB_API_VERSION = '2022-11-28';

const HARDCODED_STAGES = [
  'Requirement Analysis',
  'Proposal',
  'Commercial Finalization',
  'Statement of Work (SOW)',
  'Project Kickoff',
  'Project Planning',
  'Detailed Design',
  'Critical Design Review (CDR)',
  'Procurement & Development',
  'Integration, Verification & Validation'
];

const HARDCODED_STAGE_DOCS = {
  1: ['MRD', 'PRD', 'PRD Matrix'],
  2: [
    'Proposal Presentation (PPT)',
    'NDA',
    'Quotation (Price Breakup, Commercial Terms, Payment Schedule, Delivery Schedule)'
  ],
  3: [
    'Final SOW (Scope Matrix, Deliverables List, Exclusions List, Assumptions, Dependencies, Acceptance Criteria)'
  ],
  4: ['Kickoff Presentation', 'Meeting Minutes (MoM)', 'RACI Matrix'],
  5: [
    'Work Breakdown Structure (WBS)',
    'PDM Document',
    'Hardware Architecture',
    'Software Architecture',
    'Block Diagram',
    'Data Flow Diagram (DFD)',
    'Interface Control Document (ICD)',
    'Milestone Plan',
    'Resource Plan',
    'Budget Tracker',
    'Procurement Plan',
    'Risk Register'
  ],
  6: [
    'Schematics',
    'PCB Layout',
    'Bill of Materials (BOM)',
    'PCB Stack-up',
    'Simulation Reports',
    'Software Design Document (SDD)',
    'API Specification',
    'Database Design',
    'UI Mockups',
    'CAD Models',
    'Mechanical Drawings',
    'Thermal Analysis Report'
  ],
  7: [
    'CDR Package',
    'Design Review Presentation',
    'Simulation Results',
    'Review Minutes',
    'Risk Closure Report',
    'Manufacturing Readiness Report',
    'Design Freeze Approval'
  ],
  8: [
    'Approved Vendor List (AVL)',
    'RFQ',
    'Purchase Request (PR)',
    'Purchase Order (PO)',
    'Delivery Tracker',
    'Incoming Inspection Report',
    'PCB Assembly Report',
    'Bring-up Report',
    'Debug Logs',
    'Source Code Repository',
    'Code Review Report',
    'Unit Test Report',
    'Dataset',
    'AI Training Report',
    'Sprint Board',
    'Bug Tracker',
    'Version Control Log'
  ],
  9: [
    'Verification Plan',
    'Validation Plan',
    'Test Cases',
    'Test Reports',
    'Defect Log',
    'Regression Test Report',
    'Requirements Traceability Matrix (RTM)',
    'Customer Acceptance Test (CAT/UAT) Report',
    'Final Validation Report'
  ],
  10: []
};

const BLOCK_FIELD_CANDIDATES = {
  plannedStart: ['planned start'],
  dueDate: ['due date'],
  completedDate: ['completion date', 'completed date'],
  status: ['status'],
  milestone: ['key milestones', 'key milestone', 'milestones', 'milestone'],
  blocker: ['blocker'],
  nextStep: ['immediate next step', 'next action', 'next step'],
  document: ['documents', 'document']
};


/* ============================================================
   MENU / SETUP
   ============================================================ */

function setup() {
  onOpen();
  SpreadsheetApp.getUi().alert(
    'Setup complete',
    'Set GitHub Script Properties, then use Dashboard → Preview data. ' +
    'After checking the warnings, use Dashboard → Publish now.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard')
    .addItem('Preview data (no publish)', 'previewData')
    .addItem('Publish now', 'publishToGithubUI')
    .addItem('Rebuild Live Dashboard only', 'rebuildSheetDashboardUI')
    .addSeparator()
    .addItem('Test GitHub connection', 'testGithubConnection')
    .addItem('Enable auto-publish (~2 min)', 'enableAutoPublish')
    .addItem('Disable auto-publish', 'disableAutoPublish')
    .addToUi();
}


/* ============================================================
   HELPERS
   ============================================================ */

function str_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

function norm_(v) {
  return str_(v).toLowerCase().replace(/\s+/g, ' ').trim();
}

function normName_(v) {
  return norm_(v).replace(/[^a-z0-9]+/g, ' ').trim();
}

function dateStr_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(
      v,
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'd-MMM-yyyy'
    );
  }
  return str_(v) || null;
}

function findSheet_(ss, names) {
  for (const name of names) {
    const exact = ss.getSheetByName(name);
    if (exact) return exact;
  }

  const wanted = names.map(n => n.trim().toLowerCase());

  for (const sh of ss.getSheets()) {
    if (wanted.indexOf(sh.getName().trim().toLowerCase()) !== -1) {
      return sh;
    }
  }

  return null;
}

function findCol_(headerRow, candidates) {
  const headers = headerRow.map(h => norm_(h));

  for (const candidate of candidates) {
    const idx = headers.findIndex(h => h.includes(norm_(candidate)));
    if (idx !== -1) return idx;
  }

  return -1;
}

function uniqueStrings_(values) {
  return Array.from(new Set(
    values.map(str_).filter(Boolean)
  ));
}

function isCompleted_(entry) {
  return !!entry.completedDate ||
    /completed|complete|closed|approved/i.test(str_(entry.status));
}


/* ============================================================
   STAGE / DOCUMENT TEMPLATE
   ============================================================ */

function hasBlockSubHeaders_(row) {
  if (!row) return false;

  const slice = row.slice(4);

  return Object.keys(BLOCK_FIELD_CANDIDATES).some(field =>
    findCol_(slice, BLOCK_FIELD_CANDIDATES[field]) !== -1
  );
}

function parseTemplate_(values) {
  let headerRowIdx = -1;

  for (let i = 0; i < values.length; i++) {
    const first = norm_(values[i][0]);
    if (first === 'stage' || first === 'process') {
      headerRowIdx = i;
      break;
    }
  }

  if (headerRowIdx === -1) {
    return {
      templateRows: [],
      headerRowIdx: -1,
      subHeaderRowIdx: -1,
      variantHeaderRowIdx: -1
    };
  }

  let subHeaderRowIdx;
  let variantHeaderRowIdx;
  let dataStartRow;

  if (hasBlockSubHeaders_(values[headerRowIdx])) {
    subHeaderRowIdx = headerRowIdx;
    variantHeaderRowIdx = headerRowIdx - 1;
    dataStartRow = headerRowIdx + 1;
  } else if (
    headerRowIdx + 1 < values.length &&
    hasBlockSubHeaders_(values[headerRowIdx + 1])
  ) {
    subHeaderRowIdx = headerRowIdx + 1;
    variantHeaderRowIdx = headerRowIdx;
    dataStartRow = headerRowIdx + 2;
  } else {
    subHeaderRowIdx = headerRowIdx;
    variantHeaderRowIdx = headerRowIdx - 1;
    dataStartRow = headerRowIdx + 1;
  }

  const templateRows = [];
  let lastStage = null;
  let lastActivity = '';
  let blankStreak = 0;

  for (let i = dataStartRow; i < values.length; i++) {
    const row = values[i];

    const stageCell = str_(row[0]);
    const activityCell = str_(row[1]);
    const docCell = str_(row[2]);
    const ownerCell = str_(row[3]);

    if (!stageCell && !activityCell && !docCell && !ownerCell) {
      blankStreak++;
      if (blankStreak >= 3) break;
      continue;
    }

    blankStreak = 0;

    const match = stageCell.match(/^stage\s*(\d+)$/i);
    if (match) lastStage = parseInt(match[1], 10);

    if (activityCell) lastActivity = activityCell;

    if (!docCell) continue;

    templateRows.push({
      rowIndex: i,
      stage: lastStage,
      activity: lastActivity,
      document: docCell
    });
  }

  return {
    templateRows,
    headerRowIdx,
    subHeaderRowIdx,
    variantHeaderRowIdx
  };
}

function readStagesAndDocs_(ss, warnings) {
  const sh = findSheet_(ss, TEMPLATE_SHEET_NAMES);

  if (!sh) {
    warnings.push(
      'No Stages/documents tab found. Using the 10-stage fallback.'
    );
    return {
      stages: HARDCODED_STAGES.slice(),
      stageDocs: HARDCODED_STAGE_DOCS
    };
  }

  const values = sh.getDataRange().getValues();
  const parsed = parseTemplate_(values);

  if (!parsed.templateRows.length) {
    warnings.push(
      '"' + sh.getName() +
      '" was found but no Stage/Document rows were parsed. ' +
      'Using the 10-stage fallback.'
    );

    return {
      stages: HARDCODED_STAGES.slice(),
      stageDocs: HARDCODED_STAGE_DOCS
    };
  }

  const stages = [];
  const stageDocs = {};

  parsed.templateRows.forEach(r => {
    if (!r.stage) return;

    if (!stages[r.stage - 1] && r.activity) {
      stages[r.stage - 1] = r.activity;
    }

    if (!stageDocs[r.stage]) stageDocs[r.stage] = [];
    stageDocs[r.stage].push(r.document);
  });

  if (!stages.filter(Boolean).length) {
    warnings.push(
      '"' + sh.getName() +
      '" contains rows but no recognizable Stage N labels. ' +
      'Using the 10-stage fallback.'
    );

    return {
      stages: HARDCODED_STAGES.slice(),
      stageDocs: HARDCODED_STAGE_DOCS
    };
  }

  for (let i = 0; i < HARDCODED_STAGES.length; i++) {
    if (!stages[i]) stages[i] = HARDCODED_STAGES[i];
  }

  return {
    stages: stages.slice(0, 10),
    stageDocs
  };
}


/* ============================================================
   VARIANT HEADER PARSING
   ============================================================ */

/**
 * Handles:
 *   "SOS Band (SOS-1)"
 *   "Dual Lens IP Camera (ZMD-DLD-AI (CAM-001))"
 *
 * The OLD parser stopped at the first ")" and therefore failed on
 * nested parentheses. This parser uses the LAST "(...)" pair.
 */
function parseVariantHeader_(raw) {
  const s = str_(raw);

  const lastOpen = s.lastIndexOf('(');
  const lastClose = s.lastIndexOf(')');

  if (
    lastOpen > 0 &&
    lastClose > lastOpen
  ) {
    return {
      name: s.substring(0, lastOpen).trim(),
      code: s.substring(lastOpen + 1, lastClose).trim()
    };
  }

  return {
    name: s,
    code: s
  };
}


/* ============================================================
   OWNER TAB PARSING
   ============================================================ */

function parseOwnerTab_(sheet, warnings) {
  const values = sheet.getDataRange().getValues();
  const parsed = parseTemplate_(values);

  if (!parsed.templateRows.length) {
    warnings.push(
      'Tab "' + sheet.getName() +
      '": could not find Stage/Activity/Document/Owner template.'
    );
    return [];
  }

  const variantHeaderRow =
    values[parsed.variantHeaderRowIdx] || [];

  const subHeaderRow =
    values[parsed.subHeaderRowIdx] || [];

  const blockStarts = [];

  for (let c = 4; c < variantHeaderRow.length; c++) {
    if (str_(variantHeaderRow[c])) {
      blockStarts.push(c);
    }
  }

  if (!blockStarts.length) {
    warnings.push(
      'Tab "' + sheet.getName() +
      '": no variant blocks found from column E onward.'
    );
    return [];
  }

  const blocks = [];

  blockStarts.forEach((startCol, blockIndex) => {
    const endCol =
      blockIndex + 1 < blockStarts.length
        ? blockStarts[blockIndex + 1]
        : Math.min(startCol + 10, subHeaderRow.length);

    const variant =
      parseVariantHeader_(variantHeaderRow[startCol]);

    const localHeaders =
      subHeaderRow.slice(startCol, endCol);

    const fieldCol = {};

    Object.keys(BLOCK_FIELD_CANDIDATES).forEach(field => {
      const localIndex =
        findCol_(
          localHeaders,
          BLOCK_FIELD_CANDIDATES[field]
        );

      if (localIndex !== -1) {
        fieldCol[field] =
          startCol + localIndex;
      }
    });

    const recognized =
      Object.keys(fieldCol).length;

    if (recognized < 4) {
      warnings.push(
        'Tab "' + sheet.getName() +
        '", block "' + variantHeaderRow[startCol] +
        '": only ' + recognized +
        '/8 tracker columns were recognized.'
      );
    }

    const entries = [];

    parsed.templateRows.forEach(t => {
      const row = values[t.rowIndex];

      const get = field =>
        fieldCol[field] !== undefined
          ? row[fieldCol[field]]
          : '';

      const plannedStart = dateStr_(get('plannedStart'));
      const dueDate = dateStr_(get('dueDate'));
      const completedDate = dateStr_(get('completedDate'));
      const status = str_(get('status')) || null;
      const milestone = str_(get('milestone')) || null;
      const blocker = str_(get('blocker')) || null;
      const nextStep = str_(get('nextStep')) || null;
      const document = str_(get('document')) || null;

      let link = null;

      try {
        if (fieldCol.document !== undefined) {
          const rich =
            sheet
              .getRange(
                t.rowIndex + 1,
                fieldCol.document + 1
              )
              .getRichTextValue();

          if (rich) link = rich.getLinkUrl();
        }
      } catch (e) {}

      if (
        !plannedStart &&
        !dueDate &&
        !completedDate &&
        !status &&
        !milestone &&
        !blocker &&
        !nextStep &&
        !document
      ) {
        return;
      }

      entries.push({
        rawStage: t.stage,
        activity: t.activity,
        document,
        plannedStart,
        dueDate,
        completedDate,
        status,
        milestone,
        blocker,
        nextStep,
        link
      });
    });

    if (entries.length) {
      blocks.push({
        name: variant.name,
        code: variant.code,
        entries
      });
    }
  });

  return blocks;
}


/* ============================================================
   STAGE NORMALIZATION
   ============================================================ */

/**
 * Some owner tabs currently contain a wrong stage number beside a
 * document group. For example, the data currently published has
 * verification documents under stage 10 even though the canonical
 * 10-stage model puts verification under stage 9.
 *
 * The document name is therefore used as the strongest signal.
 */
function resolveEntryStage_(entry, stages, stageDocs) {
  const raw = str_(entry.document);

  if (!raw) {
    return entry.rawStage || 1;
  }

  const text = norm_(raw);
  let bestStage = null;
  let bestScore = 0;

  Object.keys(stageDocs).forEach(stageKey => {
    const stageNum = Number(stageKey);

    (stageDocs[stageNum] || []).forEach(docName => {
      const docNorm = norm_(docName);

      if (!docNorm) return;

      let score = 0;

      if (text === docNorm) {
        score = 100;
      } else if (text.includes(docNorm)) {
        score = 80 + Math.min(docNorm.length / 100, 0.9);
      } else if (docNorm.includes(text) && text.length >= 8) {
        score = 60 + Math.min(text.length / 100, 0.9);
      }

      if (score > bestScore) {
        bestScore = score;
        bestStage = stageNum;
      }
    });
  });

  return bestStage || entry.rawStage || 1;
}


/* ============================================================
   ENGINEERING ROLLUP
   ============================================================ */

function rollupByStage_(entries, stages, stageDocs) {
  const byStage = {};

  entries.forEach(entry => {
    const stage =
      resolveEntryStage_(
        entry,
        stages,
        stageDocs
      );

    if (!byStage[stage]) {
      byStage[stage] = [];
    }

    byStage[stage].push({
      ...entry,
      stage
    });
  });

  const out = {};

  Object.keys(byStage).forEach(stageNum => {
    const rows = byStage[stageNum];

    const completedCount =
      rows.filter(isCompleted_).length;

    const completionRatio =
      rows.length
        ? completedCount / rows.length
        : 0;

    const incompleteRows =
      rows.filter(r => !isCompleted_(r));

    const activeRow =
      incompleteRows[0] ||
      rows[rows.length - 1];

    out[stageNum] = {
      stage: Number(stageNum),
      plannedStart:
        rows.map(r => r.plannedStart).filter(Boolean).sort()[0] || null,

      dueDate:
        rows.map(r => r.dueDate).filter(Boolean).sort().slice(-1)[0] || null,

      completedDate:
        completionRatio === 1
          ? rows.map(r => r.completedDate).filter(Boolean).sort().slice(-1)[0] || null
          : null,

      status:
        (activeRow && activeRow.status) ||
        (completionRatio === 1 ? 'Completed' : null),

      milestone:
        uniqueStrings_(
          rows.map(r => r.milestone)
        ).join('; ') || null,

      blocker:
        uniqueStrings_(
          rows.map(r => r.blocker)
        ).join('; ') || null,

      nextStep:
        uniqueStrings_(
          rows.map(r => r.nextStep)
        ).join('; ') || null,

      completedCount,
      totalCount: rows.length,
      completionRatio
    };
  });

  return out;
}


/* ============================================================
   PRODUCT SUMMARY
   ============================================================ */

function calculateProductProgress_(engineering, fallbackStage, totalStages) {
  const stageNumbers =
    Object.keys(engineering)
      .map(Number)
      .filter(n => n >= 1 && n <= totalStages);

  if (!stageNumbers.length) {
    return {
      currentStage: fallbackStage || 1,
      progressPct: 0,
      completedStages: 0
    };
  }

  let weightedProgress = 0;
  let completedStages = 0;

  stageNumbers.forEach(stageNum => {
    const e = engineering[stageNum];

    weightedProgress +=
      Number(e.completionRatio || 0);

    if (Number(e.completionRatio || 0) >= 1) {
      completedStages++;
    }
  });

  const progressPct =
    Math.round(
      (weightedProgress / totalStages) * 100
    );

  /*
   * Current stage:
   * 1. Prefer the highest stage that has actual incomplete work.
   * 2. If all recorded stages are complete, current = next stage.
   * 3. If there is no engineering data, use Project Details stage.
   */
  const activeStages =
    stageNumbers.filter(
      n => Number(
        engineering[n].completionRatio || 0
      ) < 1
    );

  let currentStage;

  if (activeStages.length) {
    currentStage =
      Math.max.apply(null, activeStages);
  } else {
    const highest =
      Math.max.apply(null, stageNumbers);

    currentStage =
      Math.min(
        highest + 1,
        totalStages
      );
  }

  /*
   * If Project Details has an explicitly higher stage and there is
   * no engineering contradiction, retain it.
   */
  if (
    fallbackStage &&
    !activeStages.length &&
    fallbackStage > currentStage
  ) {
    currentStage = fallbackStage;
  }

  return {
    currentStage,
    progressPct: Math.min(
      100,
      Math.max(0, progressPct)
    ),
    completedStages
  };
}


/* ============================================================
   BUILD PAYLOAD
   ============================================================ */

function buildPayload_() {
  const ss = SpreadsheetApp.getActive();
  const warnings = [];

  const projSheet =
    findSheet_(
      ss,
      PROJECT_SHEET_NAMES
    );

  if (!projSheet) {
    throw new Error(
      'Project details / Products sheet not found.'
    );
  }

  const stageData =
    readStagesAndDocs_(
      ss,
      warnings
    );

  const stages =
    stageData.stages.length
      ? stageData.stages
      : HARDCODED_STAGES;

  const stageDocs =
    stageData.stageDocs;

  const totalStages =
    stages.length || 10;

  /* ----------------------------------------------------------
     PROJECT DETAILS
     ---------------------------------------------------------- */

  const values =
    projSheet
      .getDataRange()
      .getValues();

  let headerIndex = -1;

  for (let i = 0; i < values.length; i++) {
    if (norm_(values[i][0]) === 'category') {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) {
    throw new Error(
      'Could not find Category header in Project details.'
    );
  }

  const header = values[headerIndex];

  const colCat =
    findCol_(header, ['category']);

  const colName =
    findCol_(
      header,
      [
        'product / solution name',
        'solution name',
        'product name',
        'name'
      ]
    );

  const colId =
    findCol_(
      header,
      ['project id', 'id']
    );

  const colStage =
    findCol_(
      header,
      ['stage / status', 'stage', 'status']
    );

  const colPdm =
    findCol_(
      header,
      ['product manager', 'pdm']
    );

  const colPm =
    findCol_(
      header,
      ['program manager', 'project manager', 'pm']
    );

  const colStart =
    findCol_(
      header,
      ['start date']
    );

  const colEnd =
    findCol_(
      header,
      ['end date']
    );

  const stageNameToNum = {};

  stages.forEach((name, index) => {
    stageNameToNum[norm_(name)] =
      index + 1;
  });

  function matchStage(text) {
    const t = norm_(text);

    if (!t) {
      return {
        stage: 1,
        matched: false
      };
    }

    if (stageNameToNum[t]) {
      return {
        stage: stageNameToNum[t],
        matched: true
      };
    }

    for (const name in stageNameToNum) {
      if (
        t.includes(name) ||
        name.includes(t)
      ) {
        return {
          stage: stageNameToNum[name],
          matched: true
        };
      }
    }

    return {
      stage: 1,
      matched: false
    };
  }

  function extractId(raw) {
    const s = str_(raw);

    const lastOpen = s.lastIndexOf('(');
    const lastClose = s.lastIndexOf(')');

    if (
      lastOpen !== -1 &&
      lastClose > lastOpen
    ) {
      return s.substring(
        lastOpen + 1,
        lastClose
      ).trim();
    }

    return s;
  }

  const products = [];
  let lastCategory = '';

  for (
    let i = headerIndex + 1;
    i < values.length;
    i++
  ) {
    const row = values[i];

    const idRaw =
      colId !== -1
        ? row[colId]
        : '';

    const nameRaw =
      colName !== -1
        ? row[colName]
        : '';

    if (!str_(idRaw) && !str_(nameRaw)) {
      if (norm_(row[0]) === 'stages') break;
      continue;
    }

    const category =
      colCat !== -1
        ? str_(row[colCat])
        : '';

    if (category) {
      lastCategory = category;
    }

    const statusText =
      colStage !== -1
        ? str_(row[colStage])
        : '';

    const stageMatch =
      matchStage(statusText);

    const id =
      extractId(idRaw);

    products.push({
      cat:
        lastCategory ||
        'Uncategorized',

      name:
        str_(nameRaw) ||
        id,

      id,

      pdm:
        colPdm !== -1
          ? str_(row[colPdm]) || null
          : null,

      pm:
        colPm !== -1
          ? str_(row[colPm]) || null
          : null,

      /*
       * This is retained as the fallback only.
       * Actual current stage is calculated from engineering data below.
       */
      stage:
        stageMatch.stage,

      plannedStart:
        colStart !== -1
          ? dateStr_(row[colStart])
          : null,

      plannedEnd:
        colEnd !== -1
          ? dateStr_(row[colEnd])
          : null,

      statusLabel:
        statusText || null
    });
  }

  if (!products.length) {
    warnings.push(
      'No products were found in Project details.'
    );
  }

  /* ----------------------------------------------------------
     MATCH MAPS
     ---------------------------------------------------------- */

  const knownIds =
    products.map(p => p.id);

  const nameToId = {};

  products.forEach(p => {
    nameToId[normName_(p.name)] =
      p.id;
  });

  const prefixToIds = {};

  products.forEach(p => {
    const match =
      str_(p.id).match(
        /^([A-Za-z]+)/
      );

    if (!match) return;

    const prefix =
      match[1].toUpperCase();

    if (!prefixToIds[prefix]) {
      prefixToIds[prefix] = [];
    }

    prefixToIds[prefix].push(
      p.id
    );
  });

  /* ----------------------------------------------------------
     OWNER TABS
     ---------------------------------------------------------- */

  const documents = {};
  const engineering = {};

  OWNER_SHEET_NAMES.forEach(tabName => {
    const sh =
      findSheet_(
        ss,
        [tabName]
      );

    if (!sh) {
      warnings.push(
        'Owner tab "' +
        tabName +
        '" was not found.'
      );
      return;
    }

    const blocks =
      parseOwnerTab_(
        sh,
        warnings
      );

    blocks.forEach(block => {
      let id =
        nameToId[
          normName_(block.name)
        ];

      /*
       * Exact code match first.
       */
      if (
        !id &&
        knownIds.indexOf(
          block.code
        ) !== -1
      ) {
        id = block.code;
      }

      /*
       * Prefix fallback.
       */
      if (!id) {
        const prefixMatch =
          block.code.match(
            /^([A-Za-z]+)/
          );

        const prefix =
          prefixMatch
            ? prefixMatch[1].toUpperCase()
            : '';

        const candidates =
          prefixToIds[prefix] || [];

        if (candidates.length === 1) {
          id = candidates[0];

          warnings.push(
            'Tab "' +
            tabName +
            '": "' +
            block.name +
            ' (' +
            block.code +
            ')" matched ' +
            id +
            ' by code prefix.'
          );
        } else if (
          candidates.length > 1
        ) {
          warnings.push(
            'Tab "' +
            tabName +
            '": "' +
            block.name +
            ' (' +
            block.code +
            ')" has ambiguous prefix match: ' +
            candidates.join(', ')
          );
        }
      }

      if (!id) {
        warnings.push(
          'Tab "' +
          tabName +
          '": "' +
          block.name +
          ' (' +
          block.code +
          ')" could not be matched to Project details.'
        );
        return;
      }

      if (!documents[id]) {
        documents[id] = [];
      }

      if (!engineering[id]) {
        engineering[id] = {};
      }

      block.entries.forEach(entry => {
        const stage =
          resolveEntryStage_(
            entry,
            stages,
            stageDocs
          );

        documents[id].push({
          stage,
          document:
            entry.document || null,
          plannedStart:
            entry.plannedStart || null,
          dueDate:
            entry.dueDate || null,
          completedDate:
            entry.completedDate || null,
          status:
            entry.status || null,
          link:
            entry.link || null
        });
      });

      const normalizedEntries =
        block.entries.map(entry => ({
          ...entry,
          stage:
            resolveEntryStage_(
              entry,
              stages,
              stageDocs
            )
        }));

      const rollup =
        rollupByStage_(
          normalizedEntries,
          stages,
          stageDocs
        );

      Object.keys(rollup).forEach(stageNum => {
        const newRecord =
          rollup[stageNum];

        if (!engineering[id][stageNum]) {
          engineering[id][stageNum] =
            newRecord;
          return;
        }

        /*
         * If multiple owner tabs contain the same product,
         * merge the information instead of overwriting it.
         */
        const old =
          engineering[id][stageNum];

        engineering[id][stageNum] = {
          stage:
            Number(stageNum),

          plannedStart:
            old.plannedStart ||
            newRecord.plannedStart ||
            null,

          dueDate:
            newRecord.dueDate ||
            old.dueDate ||
            null,

          completedDate:
            newRecord.completedDate ||
            old.completedDate ||
            null,

          status:
            newRecord.status ||
            old.status ||
            null,

          milestone:
            uniqueStrings_([
              old.milestone,
              newRecord.milestone
            ]).join('; ') || null,

          blocker:
            uniqueStrings_([
              old.blocker,
              newRecord.blocker
            ]).join('; ') || null,

          nextStep:
            uniqueStrings_([
              old.nextStep,
              newRecord.nextStep
            ]).join('; ') || null,

          completedCount:
            Number(old.completedCount || 0) +
            Number(newRecord.completedCount || 0),

          totalCount:
            Number(old.totalCount || 0) +
            Number(newRecord.totalCount || 0),

          completionRatio:
            (
              Number(old.totalCount || 0) +
              Number(newRecord.totalCount || 0)
            )
              ? (
                  Number(old.completedCount || 0) +
                  Number(newRecord.completedCount || 0)
                ) /
                (
                  Number(old.totalCount || 0) +
                  Number(newRecord.totalCount || 0)
                )
              : 0
        };
      });
    });
  });

  /* ----------------------------------------------------------
     CALCULATE ACTUAL CURRENT STAGE + PROGRESS
     ---------------------------------------------------------- */

  const productById = {};

  products.forEach(p => {
    productById[p.id] = p;
  });

  products.forEach(product => {
    const eng =
      engineering[product.id] || {};

    const summary =
      calculateProductProgress_(
        eng,
        product.stage,
        totalStages
      );

    product.currentStage =
      summary.currentStage;

    product.progressPct =
      summary.progressPct;

    product.completedStages =
      summary.completedStages;

    const current =
      eng[summary.currentStage] || {};

    product.currentStageName =
      stages[
        summary.currentStage - 1
      ] ||
      'Stage ' +
      summary.currentStage;

    product.milestones =
      uniqueStrings_(
        Object.keys(eng)
          .map(n => eng[n].milestone)
      );

    product.blockers =
      uniqueStrings_(
        Object.keys(eng)
          .map(n => eng[n].blocker)
      );

    product.nextActions =
      uniqueStrings_(
        Object.keys(eng)
          .map(n => eng[n].nextStep)
      );

    product.currentMilestones =
      uniqueStrings_([
        current.milestone
      ]);

    product.currentBlockers =
      uniqueStrings_([
        current.blocker
      ]);

    product.currentNextActions =
      uniqueStrings_([
        current.nextStep
      ]);
  });

  /* ----------------------------------------------------------
     PAYLOAD
     ---------------------------------------------------------- */

  return {
    payload: {
      generatedAt:
        new Date().toISOString(),

      stages,

      stageDocs,

      products,

      engineering,

      documents
    },

    warnings
  };
}


/* ============================================================
   LIVE DASHBOARD SHEET
   ============================================================ */

function writeSheetDashboard_(payload) {
  const ss =
    SpreadsheetApp.getActive();

  let sh =
    ss.getSheetByName(
      SHEET_DASHBOARD_TAB
    );

  if (!sh) {
    sh =
      ss.insertSheet(
        SHEET_DASHBOARD_TAB
      );
  }

  sh.clear();

  const headers = [
    'Category',
    'Variant',
    'Project ID',
    'PDM',
    'PM',
    'Current Stage',
    'Progress %',
    'Completed Stages',
    'Stage Name',
    'Planned Start',
    'Due Date',
    'Completed Date',
    'Status',
    'Key Milestones',
    'Blocker',
    'Immediate Next Action'
  ];

  const rows = [];

  payload.products.forEach(p => {
    const eng =
      payload.engineering[p.id] || {};

    const stageNumbers =
      Object.keys(eng)
        .map(Number)
        .sort((a, b) => a - b);

    /*
     * Always show all stages with data, plus current stage.
     */
    const stageSet =
      new Set(stageNumbers);

    stageSet.add(
      p.currentStage
    );

    Array.from(stageSet)
      .sort((a, b) => a - b)
      .forEach(stageNum => {
        const e =
          eng[stageNum] || {};

        rows.push([
          p.cat,
          p.name,
          p.id,
          p.pdm || '',
          p.pm || '',
          p.currentStage +
            ' / ' +
            payload.stages.length,
          p.progressPct + '%',
          p.completedStages,
          payload.stages[
            stageNum - 1
          ] || 'Stage ' + stageNum,
          e.plannedStart || '',
          e.dueDate || '',
          e.completedDate || '',
          e.status || '',
          e.milestone || '',
          e.blocker || '',
          e.nextStep || ''
        ]);
      });
  });

  sh
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#EEF3FA');

  if (rows.length) {
    sh
      .getRange(
        2,
        1,
        rows.length,
        headers.length
      )
      .setValues(rows);

    /*
     * Highlight blockers.
     */
    rows.forEach((row, index) => {
      if (row[14]) {
        sh
          .getRange(
            index + 2,
            1,
            1,
            headers.length
          )
          .setBackground('#FCE8E8');
      }
    });
  }

  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  sh.autoResizeColumns(
    1,
    headers.length
  );

  return rows.length;
}


/* ============================================================
   PREVIEW
   ============================================================ */

function previewData() {
  const ui =
    SpreadsheetApp.getUi();

  try {
    const result =
      buildPayload_();

    const payload =
      result.payload;

    const warnings =
      result.warnings;

    const docCount =
      Object.keys(
        payload.documents
      ).reduce(
        (n, id) =>
          n +
          payload.documents[id].length,
        0
      );

    const summary =
      payload.products
        .map(p =>
          p.id +
          ': Stage ' +
          p.currentStage +
          ' — ' +
          p.progressPct +
          '%'
        )
        .slice(0, 20)
        .join('\n');

    let message =
      'Products: ' +
      payload.products.length +
      '\n' +
      'Stages: ' +
      payload.stages.length +
      '\n' +
      'Document rows: ' +
      docCount +
      '\n\n' +
      'Calculated progress:\n' +
      summary;

    if (warnings.length) {
      message +=
        '\n\n⚠ ' +
        warnings.length +
        ' warning(s):\n- ' +
        warnings
          .slice(0, 12)
          .join('\n- ');
    }

    Logger.log(
      JSON.stringify(
        result,
        null,
        2
      )
    );

    ui.alert(
      'Preview — nothing published',
      message,
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log(
      'Preview failed: ' +
      e.stack
    );

    ui.alert(
      'Preview failed',
      e.message,
      ui.ButtonSet.OK
    );
  }
}


/* ============================================================
   GITHUB
   ============================================================ */

function getGithubConfig_() {
  const props =
    PropertiesService
      .getScriptProperties();

  const token =
    str_(
      props.getProperty(
        'GITHUB_TOKEN'
      )
    );

  const owner =
    str_(
      props.getProperty(
        'GITHUB_OWNER'
      )
    );

  const repo =
    str_(
      props.getProperty(
        'GITHUB_REPO'
      )
    );

  const branch =
    str_(
      props.getProperty(
        'GITHUB_BRANCH'
      )
    ) || 'main';

  const path =
    str_(
      props.getProperty(
        'GITHUB_PATH'
      )
    ) || 'data.json';

  if (!token || !owner || !repo) {
    throw new Error(
      'Missing GitHub Script Properties: GITHUB_TOKEN, GITHUB_OWNER and/or GITHUB_REPO.'
    );
  }

  return {
    token,
    owner,
    repo,
    branch,
    path
  };
}

function githubHeaders_(token) {
  return {
    Authorization:
      'Bearer ' + token,

    Accept:
      'application/vnd.github+json',

    'X-GitHub-Api-Version':
      GITHUB_API_VERSION
  };
}

function encodeGithubPath_(path) {
  return str_(path)
    .split('/')
    .map(part => encodeURIComponent(part))
    .join('/');
}

function publishToGithub() {
  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      'Another publish is already running. Please wait and try again.'
    );
  }

  try {
    const config =
      getGithubConfig_();

    const result =
      buildPayload_();

    const payload =
      result.payload;

    const warnings =
      result.warnings;

    /*
     * Update the Google Sheet dashboard too.
     */
    writeSheetDashboard_(
      payload
    );

    const json =
      JSON.stringify(
        payload,
        null,
        2
      );

    const content =
      Utilities.base64Encode(
        json,
        Utilities.Charset.UTF_8
      );

    const apiUrl =
      'https://api.github.com/repos/' +
      encodeURIComponent(config.owner) +
      '/' +
      encodeURIComponent(config.repo) +
      '/contents/' +
      encodeGithubPath_(config.path);

    const headers =
      githubHeaders_(
        config.token
      );

    let sha = null;

    const getResp =
      UrlFetchApp.fetch(
        apiUrl +
        '?ref=' +
        encodeURIComponent(
          config.branch
        ),
        {
          method: 'get',
          headers,
          muteHttpExceptions: true
        }
      );

    const getCode =
      getResp.getResponseCode();

    if (getCode === 200) {
      const existing =
        JSON.parse(
          getResp.getContentText()
        );

      sha =
        existing.sha || null;

    } else if (getCode !== 404) {
      throw new Error(
        'GitHub GET failed (' +
        getCode +
        '): ' +
        getResp.getContentText()
      );
    }

    const body = {
      message:
        'Publish dashboard data — ' +
        new Date().toISOString(),

      content,

      branch:
        config.branch
    };

    if (sha) {
      body.sha = sha;
    }

    const putResp =
      UrlFetchApp.fetch(
        apiUrl,
        {
          method: 'put',
          headers,
          contentType:
            'application/json',
          payload:
            JSON.stringify(body),
          muteHttpExceptions: true
        }
      );

    const putCode =
      putResp.getResponseCode();

    if (
      putCode !== 200 &&
      putCode !== 201
    ) {
      throw new Error(
        'GitHub PUT failed (' +
        putCode +
        '): ' +
        putResp.getContentText()
      );
    }

    Logger.log(
      'Published %s products, %s document rows, %s warnings.',
      payload.products.length,
      Object.keys(payload.documents)
        .reduce(
          (n, id) =>
            n +
            payload.documents[id].length,
          0
        ),
      warnings.length
    );

    return result;

  } finally {
    lock.releaseLock();
  }
}

function publishToGithubUI() {
  const ui =
    SpreadsheetApp.getUi();

  try {
    const result =
      publishToGithub();

    const payload =
      result.payload;

    const warnings =
      result.warnings;

    const docCount =
      Object.keys(
        payload.documents
      ).reduce(
        (n, id) =>
          n +
          payload.documents[id].length,
        0
      );

    let msg =
      'Published successfully.\n\n' +
      'Products: ' +
      payload.products.length +
      '\n' +
      'Document rows: ' +
      docCount +
      '\n' +
      'Live Dashboard updated.';

    if (warnings.length) {
      msg +=
        '\n\n⚠ ' +
        warnings.length +
        ' warning(s):\n- ' +
        warnings
          .slice(0, 8)
          .join('\n- ') +
        '\n\nSee Apps Script → Executions for the full log.';
    }

    ui.alert(
      'Publish succeeded',
      msg,
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log(
      'Publish failed: ' +
      e.stack
    );

    ui.alert(
      'Publish failed',
      e.message,
      ui.ButtonSet.OK
    );
  }
}


/* ============================================================
   LIVE SHEET ONLY
   ============================================================ */

function rebuildSheetDashboardUI() {
  const ui =
    SpreadsheetApp.getUi();

  try {
    const result =
      buildPayload_();

    const count =
      writeSheetDashboard_(
        result.payload
      );

    let msg =
      'Wrote ' +
      count +
      ' rows to "' +
      SHEET_DASHBOARD_TAB +
      '".\n\nGitHub was not changed.';

    if (result.warnings.length) {
      msg +=
        '\n\n⚠ ' +
        result.warnings.length +
        ' warning(s):\n- ' +
        result.warnings
          .slice(0, 8)
          .join('\n- ');
    }

    ui.alert(
      'Live Dashboard rebuilt',
      msg,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert(
      'Rebuild failed',
      e.message,
      ui.ButtonSet.OK
    );
  }
}


/* ============================================================
   AUTO PUBLISH
   ============================================================ */

function onEditDebounced_() {
  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      if (
        trigger.getHandlerFunction() ===
        'publishToGithub'
      ) {
        ScriptApp.deleteTrigger(
          trigger
        );
      }
    });

  ScriptApp
    .newTrigger(
      'publishToGithub'
    )
    .timeBased()
    .after(
      2 * 60 * 1000
    )
    .create();
}

function enableAutoPublish() {
  disableAutoPublish(false);

  ScriptApp
    .newTrigger(
      'onEditDebounced_'
    )
    .forSpreadsheet(
      SpreadsheetApp.getActive()
    )
    .onEdit()
    .create();

  SpreadsheetApp
    .getUi()
    .alert(
      'Auto-publish enabled',
      'Edits will publish approximately 2 minutes after the last edit.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
}

function disableAutoPublish(showAlert) {
  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      const fn =
        trigger.getHandlerFunction();

      if (
        fn === 'publishToGithub' ||
        fn === 'onEditDebounced_'
      ) {
        ScriptApp.deleteTrigger(
          trigger
        );
      }
    });

  if (showAlert !== false) {
    SpreadsheetApp
      .getUi()
      .alert(
        'Auto-publish disabled',
        'Use Dashboard → Publish now when you want to publish manually.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
  }
}


/* ============================================================
   TEST
   ============================================================ */

function testBuildPayload() {
  const result =
    buildPayload_();

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}

function testGithubConnection() {
  const config =
    getGithubConfig_();

  const url =
    'https://api.github.com/repos/' +
    encodeURIComponent(config.owner) +
    '/' +
    encodeURIComponent(config.repo);

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'get',
        headers:
          githubHeaders_(
            config.token
          ),
        muteHttpExceptions: true
      }
    );

  const code =
    response.getResponseCode();

  if (code !== 200) {
    throw new Error(
      'GitHub connection failed (' +
      code +
      '): ' +
      response.getContentText()
    );
  }

  const data =
    JSON.parse(
      response.getContentText()
    );

  SpreadsheetApp
    .getUi()
    .alert(
      'GitHub connection successful',
      'Repository: ' +
      data.full_name +
      '\nConfigured branch: ' +
      config.branch,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
}
