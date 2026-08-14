from pathlib import Path

cp = Path('apps-script/Code.gs')
code = cp.read_text(encoding='utf-8')

if 'const DOCUMENT_SOURCE_SHEET_NAMES' not in code:
    code = code.replace("const OWNER_SHEET_NAMES = [", "const DOCUMENT_SOURCE_SHEET_NAMES = ['Documents'];\n\nconst OWNER_SHEET_NAMES = [", 1)

if 'function readDocumentsSheet_' not in code:
    fn = r'''
/* ============================================================
   DOCUMENTS SHEET — AUTHORITATIVE DOCUMENT SOURCE
   ============================================================ */
function readDocumentsSheet_(ss, stages, products, warnings) {
  const sheet = findSheet_(ss, DOCUMENT_SOURCE_SHEET_NAMES);
  if (!sheet) return { found: false, documents: {} };
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { found: true, documents: {} };
  let headerIdx = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r];
    const hasDoc = findCol_(row, ['document name','document names','document','file name']) !== -1;
    const hasProject = findCol_(row, ['project id','project','product id','variant id','id']) !== -1;
    const hasStage = findCol_(row, ['stage #','stage no','stage number','stage name','stage']) !== -1;
    if (hasDoc && (hasProject || hasStage)) { headerIdx = r; break; }
  }
  if (headerIdx === -1) return { found: true, documents: {} };
  const header = values[headerIdx];
  const colId = findCol_(header, ['project id','product id','variant id','project','id']);
  const colName = findCol_(header, ['product / solution name','solution name','product name','variant name']);
  const colStageNum = findCol_(header, ['stage #','stage no','stage number']);
  const colStage = findCol_(header, ['stage name','stage','activity']);
  const colDoc = findCol_(header, ['document name','document names','file name','document']);
  const colLink = findCol_(header, ['document link','file link','drive link','url','link']);
  const colStatus = findCol_(header, ['status','document status']);
  const colStart = findCol_(header, ['planned start','start date']);
  const colDue = findCol_(header, ['due date','deadline']);
  const colCompleted = findCol_(header, ['completed date','completion date']);
  if (colDoc === -1) return { found: true, documents: {} };

  const knownIds = products.map(p => p.id);
  const nameToId = {};
  products.forEach(p => { nameToId[normName_(p.name)] = p.id; });
  const stageNameToNum = {};
  stages.forEach((s,i) => { stageNameToNum[norm_(s)] = i + 1; });

  function resolveStage_(rawNum, rawName) {
    const n = parseInt(str_(rawNum),10);
    if (!isNaN(n) && n >= 1 && n <= stages.length) return n;
    const text = norm_(rawName);
    if (!text) return null;
    if (stageNameToNum[text]) return stageNameToNum[text];
    const m = text.match(/stage\s*(\d+)/i);
    if (m) return parseInt(m[1],10);
    for (const s in stageNameToNum) if (text.includes(s) || s.includes(text)) return stageNameToNum[s];
    return null;
  }
  function resolveId_(rawId, rawName) {
    const id = str_(rawId);
    if (id && knownIds.indexOf(id) !== -1) return id;
    const name = normName_(rawName);
    if (name && nameToId[name]) return nameToId[name];
    return knownIds.find(x => norm_(x) === norm_(id)) || null;
  }
  function richLink_(r,c) {
    if (c === -1) return null;
    try { const rich = sheet.getRange(r+1,c+1).getRichTextValue(); return rich ? rich.getLinkUrl() : null; } catch(e) { return null; }
  }

  const documents = {};
  for (let r = headerIdx + 1; r < values.length; r++) {
    const row = values[r];
    const document = str_(row[colDoc]);
    if (!document) continue;
    const id = resolveId_(colId === -1 ? '' : row[colId], colName === -1 ? '' : row[colName]);
    if (!id) continue;
    const stage = resolveStage_(colStageNum === -1 ? '' : row[colStageNum], colStage === -1 ? '' : row[colStage]);
    if (!stage) continue;
    let link = colLink === -1 ? '' : str_(row[colLink]);
    if (!link) link = richLink_(r,colLink);
    if (!documents[id]) documents[id] = [];
    documents[id].push({
      stage: stage,
      document: document,
      plannedStart: colStart === -1 ? null : dateStr_(row[colStart]),
      dueDate: colDue === -1 ? null : dateStr_(row[colDue]),
      completedDate: colCompleted === -1 ? null : dateStr_(row[colCompleted]),
      status: colStatus === -1 ? null : (str_(row[colStatus]) || null),
      link: link || null
    });
  }
  return { found: true, documents: documents };
}
'''
    code = code.replace("\nfunction buildPayload_() {", "\n" + fn + "\nfunction buildPayload_() {", 1)

marker = "\n  if (!anyOwnerTabFound) warnings.push('None of the expected owner tabs were found — no milestones/blockers/next steps/documents will show for any variant.');"
if 'const documentSource = readDocumentsSheet_' not in code and marker in code:
    merge = """

  // Documents sheet is authoritative for document names, status and links.
  const documentSource = readDocumentsSheet_(ss, stages, products, warnings);
  if (documentSource.found) {
    Object.keys(documents).forEach(id => delete documents[id]);
    Object.keys(documentSource.documents).forEach(id => { documents[id] = documentSource.documents[id]; });
  }
"""
    code = code.replace(marker, merge + marker, 1)
cp.write_text(code, encoding='utf-8')

# Dashboard patch: use every DATA.documents row for the expanded stage and add Open/Download buttons.
dp = Path('dashboard.html')
dash = dp.read_text(encoding='utf-8')
if 'DOCUMENTS SHEET LIVE SOURCE PATCH' not in dash:
    css = '''<style>\n/* DOCUMENTS SHEET LIVE SOURCE PATCH */\n.doc-row .doc-actions{display:flex;align-items:center;gap:8px;flex-shrink:0}.doc-row .doc-actions a{display:inline-block;padding:4px 8px;border:1px solid var(--line);border-radius:5px;background:var(--panel);font-size:10.5px;font-weight:600;text-decoration:none}.doc-row .doc-actions a.download{color:var(--teal)}.doc-row .doc-actions a.open{color:var(--sub)}\n</style>'''
    js = '''<script>\n/* DOCUMENTS SHEET LIVE SOURCE PATCH */\n(function(){\nfunction esc(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;')}\nfunction docs(pid,n){return ((DATA.documents&&DATA.documents[pid])||[]).filter(x=>Number(x.stage)===Number(n)&&(x.document||x.documentName||x.name))}\nfunction dl(u){if(!u)return'';var s=String(u),m=s.match(/drive\\.google\\.com\\/file\\/d\\/([^/]+)/i)||s.match(/[?&]id=([^&]+)/i),id=m?m[1]:null;return id?'https://drive.google.com/uc?export=download&id='+encodeURIComponent(id):s}\nfunction refresh(){if(!window.selectedProduct||!window.expandedStage)return;var a=docs(window.selectedProduct,window.expandedStage);if(!a.length)return;var p=document.getElementById('detail');if(!p)return;var row=Array.from(p.querySelectorAll('.stage-row')).find(function(r){var n=r.querySelector('.stage-num');return n&&Number(n.textContent.trim())===Number(window.expandedStage)});if(!row)return;var inner=row.nextElementSibling&&row.nextElementSibling.querySelector('.stage-detail-inner');if(!inner)return;var sec=inner.querySelector('.doc-section');if(!sec){sec=document.createElement('div');sec.className='doc-section';inner.appendChild(sec)}sec.innerHTML='<div class="k">Documents</div>'+a.map(function(d){var n=d.document||d.documentName||d.name||'Document',s=d.status||'Available',u=d.link||d.url||d.documentLink||'',c=/complet/i.test(s)?'ready':/progress|track|revision/i.test(s)?'progress':'notstarted',act=u?'<div class="doc-actions"><a class="open" href="'+esc(u)+'" target="_blank" onclick="event.stopPropagation()">Open ↗</a><a class="download" href="'+esc(dl(u))+'" target="_blank" download onclick="event.stopPropagation()">Download ↓</a></div>':'';return '<div class="doc-row"><span>'+esc(n)+'</span><div style="display:flex;align-items:center;gap:8px"><span class="doc-status '+c+'">'+esc(s)+'</span>'+act+'</div></div>'}).join('')}\nfunction install(){if(window.__documentsPatchInstalled||typeof window.renderDetail!=='function')return;var old=window.renderDetail;window.renderDetail=function(){old();setTimeout(refresh,0)};window.__documentsPatchInstalled=true;setTimeout(refresh,0)}\nvar t=setInterval(function(){install();if(window.__documentsPatchInstalled)clearInterval(t)},100)\n})();\n</script>'''
    dash=dash.replace('</head>',css+'\n</head>',1).replace('</body>',js+'\n</body>',1)
dp.write_text(dash,encoding='utf-8')
