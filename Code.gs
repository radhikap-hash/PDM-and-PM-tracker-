const CONFIG = {
  SPREADSHEET_ID: '1xVZsxiIdWAASMGL6Bj7H3A-uyGg0r3PyhZAv6tVyMlU',
  SHEET_NAME: 'Tracker',
  TIMEZONE: 'Asia/Kolkata',
  TOTAL_STAGES: 10
};

const STAGE_DEFINITIONS = [
  {number:1,name:'Market Research'},
  {number:2,name:'Proposal'},
  {number:3,name:'Commercial Finalization'},
  {number:4,name:'Statement of Work (SOW)'},
  {number:5,name:'Project Kickoff'},
  {number:6,name:'Project Planning'},
  {number:7,name:'Detailed Design'},
  {number:8,name:'Critical Design Review (CDR)'},
  {number:9,name:'Procurement & Development'},
  {number:10,name:'Integration, Verification & Validation'}
];

const REQUIRED_HEADERS = [
  'Project ID','Stage #','Stage Name','Owner','Planned Start','Due Date',
  'Completed Date','Status','Key Milestones','Blocker',
  'Immediate Next Step / Action Item','Document name','Document Link',
  'Document Status','Remarks'
];

function doGet() {
  try {
    return ContentService.createTextOutput(JSON.stringify(buildDashboardData_()))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success:false,
      error:String(err && err.message || err),
      generatedAt:new Date().toISOString(),
      totalProjects:0,
      totalStages:CONFIG.TOTAL_STAGES,
      projects:[], products:[], engineering:{}, documents:{}, stageDocs:{},
      stages:STAGE_DEFINITIONS.map(s=>s.name),
      stageDefinitions:STAGE_DEFINITIONS
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function buildDashboardData_() {
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet=ss.getSheetByName(CONFIG.SHEET_NAME);
  if(!sheet) throw new Error('Sheet not found: '+CONFIG.SHEET_NAME);
  const values=sheet.getDataRange().getValues();
  if(!values || values.length<1) return emptyPayload_();

  const headers=values[0].map(v=>String(v||'').trim());
  validateHeaders_(headers);
  const index={}; headers.forEach((h,i)=>index[h]=i);
  const projects={};

  for(let r=1;r<values.length;r++){
    const row=values[r];
    const projectId=clean_(row[index['Project ID']]);
    if(!projectId) continue;
    const stageNumber=parseStageNumber_(row[index['Stage #']]);
    if(stageNumber<1 || stageNumber>CONFIG.TOTAL_STAGES) continue;

    if(!projects[projectId]) projects[projectId]={projectId:projectId,stageMap:{}};
    if(!projects[projectId].stageMap[stageNumber]) projects[projectId].stageMap[stageNumber]=createEmptyStage_(stageNumber);
    const stage=projects[projectId].stageMap[stageNumber];

    stage.stageName=clean_(row[index['Stage Name']]) || stage.stageName;
    stage.owner=mergeText_(stage.owner,row[index['Owner']]);
    stage.plannedStart=formatDate_(row[index['Planned Start']]) || stage.plannedStart;
    stage.dueDate=formatDate_(row[index['Due Date']]) || stage.dueDate;
    stage.completedDate=formatDate_(row[index['Completed Date']]) || stage.completedDate;
    stage.status=strongerStatus_(stage.status,normalizeStatus_(row[index['Status']]));
    stage.keyMilestones=mergeText_(stage.keyMilestones,row[index['Key Milestones']]);
    stage.blocker=mergeText_(stage.blocker,row[index['Blocker']]);
    stage.immediateNextStep=mergeText_(stage.immediateNextStep,row[index['Immediate Next Step / Action Item']]);
    stage.remarks=mergeText_(stage.remarks,row[index['Remarks']]);

    parseDocuments_(row[index['Document name']],row[index['Document Link']],row[index['Document Status']],row[index['Remarks']]).forEach(doc=>{
      if(!stage.documents.some(x=>x.name===doc.name && x.link===doc.link)) stage.documents.push(doc);
    });
  }

  const projectList=Object.keys(projects).sort().map(id=>finalizeProject_(projects[id]));
  return toDashboardPayload_(projectList);
}

function finalizeProject_(project) {
  const stages=[];
  for(let i=1;i<=CONFIG.TOTAL_STAGES;i++) stages.push(project.stageMap[i] || createEmptyStage_(i));
  stages.forEach(s=>s.status=normalizeStatus_(s.status));

  const completedStages=stages.filter(s=>s.status==='Completed');
  const blocked=stages.find(s=>s.status==='Blocked');
  const inProgress=stages.find(s=>s.status==='In Progress');
  let currentStage=1;

  if(blocked) currentStage=blocked.stageNumber;
  else if(inProgress) currentStage=inProgress.stageNumber;
  else {
    const next=stages.find(s=>s.status==='Not Started');
    if(next) currentStage=next.stageNumber;
    else if(completedStages.length) currentStage=completedStages[completedStages.length-1].stageNumber;
  }

  const progress=Math.round(completedStages.length/CONFIG.TOTAL_STAGES*100);
  return {
    projectId:project.projectId,
    projectName:project.projectId,
    category:inferCategory_(project.projectId),
    currentStage:currentStage,
    currentStageName:stages[currentStage-1].stageName,
    progress:progress,
    progressPct:progress,
    completedStages:completedStages.length,
    totalStages:CONFIG.TOTAL_STAGES,
    stages:stages
  };
}

function toDashboardPayload_(projectList) {
  const engineering={};
  const documents={};
  const products=[];

  projectList.forEach(p=>{
    const eng={};
    const docs=[];

    p.stages.forEach(s=>{
      eng[s.stageNumber]={
        stageNumber:s.stageNumber,stageName:s.stageName,owner:s.owner,
        plannedStart:s.plannedStart,dueDate:s.dueDate,completedDate:s.completedDate,
        status:s.status,milestone:s.keyMilestones,blocker:s.blocker,
        nextStep:s.immediateNextStep,remarks:s.remarks,documents:s.documents
      };

      s.documents.forEach(doc=>docs.push({
        productId:p.projectId,projectId:p.projectId,
        stage:s.stageNumber,stageNumber:s.stageNumber,
        document:doc.name,documentName:doc.name,name:doc.name,
        link:doc.link,url:doc.link,documentLink:doc.link,
        status:doc.status,remarks:doc.remarks
      }));
    });

    engineering[p.projectId]=eng;
    documents[p.projectId]=docs;

    const current=p.stages[p.currentStage-1] || createEmptyStage_(p.currentStage);
    products.push({
      id:p.projectId,projectId:p.projectId,name:p.projectName,
      projectName:p.projectName,productName:p.projectName,
      cat:p.category,category:p.category,
      pdm:current.owner || '',pm:'',
      stage:p.currentStage,currentStage:p.currentStage,
      currentStageName:p.currentStageName,
      progress:p.progress,progressPct:p.progressPct,
      completedStages:p.completedStages,totalStages:p.totalStages,
      status:current.status,blocker:current.blocker || '',
      plannedStart:current.plannedStart || '',dueDate:current.dueDate || '',
      completedDate:current.completedDate || ''
    });
  });

  return {
    success:true,
    generatedAt:new Date().toISOString(),
    source:{spreadsheetId:CONFIG.SPREADSHEET_ID,sheet:CONFIG.SHEET_NAME},
    totalProjects:projectList.length,
    totalStages:CONFIG.TOTAL_STAGES,
    projects:projectList,
    products:products,
    engineering:engineering,
    documents:documents,
    stageDocs:documents,
    stages:STAGE_DEFINITIONS.map(s=>s.name),
    stageDefinitions:STAGE_DEFINITIONS
  };
}

function inferCategory_(id) {
  const x=String(id||'').toUpperCase();
  if(x.indexOf('CAM')===0) return 'Camera';
  if(x.indexOf('SEN')===0) return 'Sensor';
  if(x.indexOf('EDG')===0 || x.indexOf('EDGE')===0) return 'Edge';
  if(x.indexOf('DRN')===0 || x.indexOf('DRONE')===0) return 'Drone';
  if(x.indexOf('VDB')===0 || x.indexOf('VDEL')===0) return 'VDelibot';
  if(x.indexOf('NVR')===0) return 'NVR';
  return 'Project';
}

function createEmptyStage_(n) {
  const d=STAGE_DEFINITIONS.find(x=>x.number===n);
  return {stageNumber:n,stageName:d?d.name:'Stage '+n,owner:'',plannedStart:'',dueDate:'',completedDate:'',status:'Not Started',keyMilestones:'',blocker:'',immediateNextStep:'',remarks:'',documents:[]};
}

function parseDocuments_(names,links,statuses,remarks) {
  const a=splitCell_(names),b=splitCell_(links),c=splitCell_(statuses),d=splitCell_(remarks);
  const max=Math.max(a.length,b.length,c.length,d.length),out=[];
  for(let i=0;i<max;i++){
    const name=a[i]||'',link=b[i]||'';
    if(!name&&!link) continue;
    out.push({name:name,link:link,status:normalizeDocumentStatus_(c[i]||''),remarks:d[i]||''});
  }
  return out;
}

function splitCell_(v) { const t=clean_(v); return t?t.split(/\r?\n/).map(x=>x.trim()).filter(Boolean):[]; }
function normalizeStatus_(v) {
  const x=clean_(v).toLowerCase();
  if(['completed','complete','closed','done','approved'].includes(x)) return 'Completed';
  if(['in progress','in-progress','inprogress','ongoing','started','active'].includes(x)) return 'In Progress';
  if(['blocked','block'].includes(x)) return 'Blocked';
  if(['on hold','on-hold','hold'].includes(x)) return 'On Hold';
  return 'Not Started';
}
function strongerStatus_(oldStatus,newStatus) {
  const rank={'Not Started':1,'On Hold':2,'In Progress':3,'Blocked':4,'Completed':5};
  return (rank[newStatus]||1)>(rank[oldStatus]||1)?newStatus:oldStatus;
}
function normalizeDocumentStatus_(v) { return clean_(v)||'Not Uploaded'; }
function formatDate_(v) {
  if(!v) return '';
  if(Object.prototype.toString.call(v)==='[object Date]') {
    if(isNaN(v.getTime())) return '';
    return Utilities.formatDate(v,CONFIG.TIMEZONE,'dd/MM/yyyy');
  }
  return clean_(v);
}
function parseStageNumber_(v) { if(typeof v==='number') return Math.floor(v); const m=String(v||'').match(/\d+/); return m?Number(m[0]):0; }
function clean_(v) { return v===null||v===undefined?'':String(v).trim(); }
function mergeText_(oldValue,newValue) {
  const incoming=clean_(newValue);
  if(!incoming) return oldValue||'';
  if(!oldValue) return incoming;
  const p=oldValue.split(/\s*;\s*/).filter(Boolean);
  if(!p.includes(incoming)) p.push(incoming);
  return p.join('; ');
}
function validateHeaders_(headers) {
  const missing=REQUIRED_HEADERS.filter(h=>headers.indexOf(h)===-1);
  if(missing.length) throw new Error('Tracker is missing these headers: '+missing.join(', '));
}
function emptyPayload_() {
  return {success:true,generatedAt:new Date().toISOString(),source:{spreadsheetId:CONFIG.SPREADSHEET_ID,sheet:CONFIG.SHEET_NAME},totalProjects:0,totalStages:CONFIG.TOTAL_STAGES,projects:[],products:[],engineering:{},documents:{},stageDocs:{},stages:STAGE_DEFINITIONS.map(s=>s.name),stageDefinitions:STAGE_DEFINITIONS};
}
function testDashboard() {
  const d=buildDashboardData_();
  Logger.log('Projects: '+d.projects.length);
  d.projects.forEach(p=>Logger.log(p.projectId+' | Stage '+p.currentStage+' | '+p.currentStageName+' | '+p.progress+'%'));
}
function getDataJson() { return JSON.stringify(buildDashboardData_(),null,2); }
