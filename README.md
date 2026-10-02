<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Document Revision Compare</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --bg:#12181f;
    --panel:#1a222c;
    --panel-2:#212b37;
    --line:#2c3947;
    --ink:#dfe6ee;
    --ink-dim:#8b98a8;
    --accent:#3ea6a0;
    --accent-2:#e8a33d;
    --add-bg:#173a2c;
    --add-ink:#7fe0a8;
    --del-bg:#3a1e1e;
    --del-ink:#f19b9b;
    --mono: "JetBrains Mono","Roboto Mono",Consolas,monospace;
    --sans: "Inter","Segoe UI",Arial,sans-serif;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    background:
      linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0) 200px),
      var(--bg);
    color:var(--ink);
    font-family:var(--sans);
    -webkit-font-smoothing:antialiased;
  }
  .wrap{max-width:1180px;margin:0 auto;padding:28px 20px 60px;}

  header.top{
    display:flex;align-items:baseline;justify-content:space-between;
    border-bottom:1px solid var(--line);
    padding-bottom:16px;margin-bottom:18px;flex-wrap:wrap;gap:8px;
  }
  header.top h1{
    font-size:22px;margin:0;font-weight:650;letter-spacing:.2px;
  }
  header.top h1 span{color:var(--accent);}
  header.top p{margin:0;color:var(--ink-dim);font-size:13px;}

  .modes{display:flex;gap:8px;margin-bottom:18px;flex-wrap:wrap;}
  .mode-btn{
    background:var(--panel);border:1px solid var(--line);color:var(--ink-dim);
    font-size:13px;font-weight:600;padding:9px 16px;border-radius:8px;cursor:pointer;
    display:flex;align-items:center;gap:7px;transition:border-color .15s,color .15s,background .15s;
  }
  .mode-btn .dot{width:7px;height:7px;border-radius:50%;background:var(--ink-dim);}
  .mode-btn:hover{border-color:var(--accent);color:var(--ink);}
  .mode-btn.active{background:rgba(62,166,160,.12);border-color:var(--accent);color:var(--ink);}
  .mode-btn.active .dot{background:var(--accent);}

  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
  @media(max-width:760px){.grid2{grid-template-columns:1fr;}}

  .drop{
    background:var(--panel);
    border:1.5px dashed var(--line);
    border-radius:10px;
    padding:20px;
    text-align:center;
    cursor:pointer;
    transition:border-color .15s, background .15s;
    position:relative;
  }
  .drop:hover, .drop.drag{border-color:var(--accent);background:var(--panel-2);}
  .drop input{display:none;}
  .drop .tag{
    display:inline-block;font-size:11px;color:var(--accent);
    background:rgba(62,166,160,.12);border:1px solid rgba(62,166,160,.35);
    border-radius:5px;padding:2px 7px;margin-bottom:10px;
  }
  .drop .tag.b{color:var(--accent-2);background:rgba(232,163,61,.12);border-color:rgba(232,163,61,.35);}
  .drop .fname{font-size:14px;font-weight:600;margin-top:4px;word-break:break-all;}
  .drop .meta{font-size:12px;color:var(--ink-dim);margin-top:3px;}
  .drop .hint{font-size:12.5px;color:var(--ink-dim);}

  .controls{
    display:flex;align-items:center;justify-content:space-between;
    gap:14px;margin:20px 0 4px;flex-wrap:wrap;
  }
  .opts{display:flex;gap:16px;flex-wrap:wrap;}
  .opt{display:flex;align-items:center;gap:6px;font-size:13px;color:var(--ink-dim);}
  .opt input{accent-color:var(--accent);}

  button.run{
    background:var(--accent);color:#0b1512;border:none;
    font-weight:700;font-size:14px;padding:11px 22px;border-radius:8px;
    cursor:pointer;transition:filter .15s, transform .05s;
  }
  button.run:hover{filter:brightness(1.08);}
  button.run:active{transform:translateY(1px);}
  button.run:disabled{background:#3a4451;color:#75828f;cursor:not-allowed;}

  .status{font-size:13px;color:var(--accent-2);min-height:18px;margin-top:10px;}

  .summary{
    display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:22px 0 18px;
  }
  @media(max-width:700px){.summary{grid-template-columns:repeat(2,1fr);}}
  .card{
    background:var(--panel);border:1px solid var(--line);border-radius:10px;
    padding:14px 16px;
  }
  .card .n{font-size:26px;font-weight:700;font-family:var(--mono);}
  .card .l{font-size:12px;color:var(--ink-dim);margin-top:2px;}
  .card.add .n{color:var(--add-ink);}
  .card.del .n{color:var(--del-ink);}

  .toolbar{
    display:flex;justify-content:space-between;align-items:center;
    margin:6px 0 10px;flex-wrap:wrap;gap:10px;
  }
  .filters{display:flex;gap:6px;flex-wrap:wrap;}
  .filters button{
    background:transparent;border:1px solid var(--line);color:var(--ink-dim);
    font-size:12.5px;padding:6px 12px;border-radius:20px;cursor:pointer;
  }
  .filters button.active{background:var(--accent);border-color:var(--accent);color:#0b1512;font-weight:600;}
  .export{
    background:transparent;border:1px solid var(--line);color:var(--ink);
    font-size:12.5px;padding:6px 12px;border-radius:6px;cursor:pointer;
  }
  .export:hover{border-color:var(--accent);}

  table.diff{
    width:100%;border-collapse:collapse;font-family:var(--mono);font-size:12.5px;
    background:var(--panel);border:1px solid var(--line);border-radius:10px;overflow:hidden;
  }
  table.diff thead th{
    background:var(--panel-2);color:var(--ink-dim);text-align:left;
    padding:9px 10px;font-size:11px;font-weight:600;border-bottom:1px solid var(--line);
  }
  table.diff td{padding:7px 10px;vertical-align:top;border-bottom:1px solid rgba(255,255,255,0.03);}
  table.diff td.ln{color:var(--ink-dim);width:44px;}
  table.diff td.pg{color:var(--ink-dim);width:90px;white-space:nowrap;}
  table.diff td.status{width:78px;font-weight:600;}
  table.diff td.txt{white-space:pre-wrap;word-break:break-word;width:44%;}
  tr.eq td.status{color:var(--ink-dim);}
  tr.del td{background:var(--del-bg);}
  tr.del td.status{color:var(--del-ink);}
  tr.add td{background:var(--add-bg);}
  tr.add td.status{color:var(--add-ink);}
  .wdel{background:rgba(241,155,155,.28);border-radius:2px;}
  .wadd{background:rgba(127,224,168,.28);border-radius:2px;}
  tr.hidden{display:none;}

  .empty{color:var(--ink-dim);font-size:13px;text-align:center;padding:40px 0;}
  footer{color:var(--ink-dim);font-size:11.5px;text-align:center;margin-top:26px;}
</style>
</head>
<body>
<div class="wrap">

  <header class="top">
    <div>
      <h1>Document Revision <span>Compare</span></h1>
      <p>Runs entirely in this browser tab — files are never uploaded anywhere.</p>
    </div>
  </header>

  <div class="modes">
    <button class="mode-btn active" data-mode="pdf"><span class="dot"></span>PDF comparison</button>
    <button class="mode-btn" data-mode="docx"><span class="dot"></span>Word file comparison</button>
    <button class="mode-btn" data-mode="xlsx"><span class="dot"></span>Excel sheet comparison</button>
  </div>

  <div class="grid2">
    <div class="drop" id="drop1">
      <span class="tag">FILE 1 &middot; OLDER</span>
      <input type="file" id="file1" accept=".pdf,application/pdf">
      <div class="fname" id="name1">Click or drop the older PDF here</div>
      <div class="meta" id="meta1"></div>
      <div class="hint" id="hint1">.pdf</div>
    </div>
    <div class="drop" id="drop2">
      <span class="tag b">FILE 2 &middot; NEWER</span>
      <input type="file" id="file2" accept=".pdf,application/pdf">
      <div class="fname" id="name2">Click or drop the newer PDF here</div>
      <div class="meta" id="meta2"></div>
      <div class="hint" id="hint2">.pdf</div>
    </div>
  </div>

  <div class="controls">
    <div class="opts">
      <label class="opt"><input type="checkbox" id="optCase"> Ignore case</label>
      <label class="opt"><input type="checkbox" id="optSpace" checked> Ignore extra whitespace</label>
      <label class="opt"><input type="checkbox" id="optWord" checked> <span id="optWordLabel">Highlight word-level changes</span></label>
      <label class="opt xlsx-only" id="optFormulaWrap" style="display:none;"><input type="checkbox" id="optFormula"> Compare formulas (not just displayed values)</label>
    </div>
    <button class="run" id="runBtn" disabled>Compare documents</button>
  </div>
  <div class="status" id="status"></div>

  <div id="results" style="display:none;">
    <div class="summary">
      <div class="card add"><div class="n" id="cAdd">0</div><div class="l">added</div></div>
      <div class="card del"><div class="n" id="cDel">0</div><div class="l">removed</div></div>
      <div class="card"><div class="n" id="cSame">0</div><div class="l">unchanged</div></div>
      <div class="card"><div class="n" id="cPages">0 / 0</div><div class="l" id="cPagesLabel">pages (File 1 / File 2)</div></div>
    </div>

    <div class="toolbar">
      <div class="filters">
        <button class="active" data-f="all">All</button>
        <button data-f="changed">Changed only</button>
        <button data-f="add">Added</button>
        <button data-f="del">Removed</button>
      </div>
      <button class="export" id="exportBtn">Export diff as CSV</button>
    </div>

    <table class="diff">
      <thead>
        <tr>
          <th class="ln">#</th>
          <th class="status">Status</th>
          <th class="pg">Location</th>
          <th class="txt">File 1</th>
          <th class="txt">File 2</th>
        </tr>
      </thead>
      <tbody id="diffBody"></tbody>
    </table>
  </div>

  <footer id="footerNote">Text is extracted in reading order per line; very complex multi-column layouts may interleave. No data leaves this page.</footer>
</div>

<script>
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const MODE_NOUN  = { pdf: 'PDF', docx: 'Word file', xlsx: 'Excel / spreadsheet file' };
const MODE_EXT   = { pdf: ['pdf'], docx: null, xlsx: null };   // Word/Excel: any file, detected by content
const MODE_ACCEPT= {
  pdf:  '.pdf,application/pdf',
  docx: '.doc,.docx,.docm,.dot,.dotx,.dotm,.rtf,.odt,.ott,.txt,.htm,.html,*/*',
  xlsx: '.xls,.xlsx,.xlsm,.xlsb,.xlt,.xltx,.xltm,.csv,.tsv,.txt,.ods,.ots,.fods,.xml,.htm,.html,*/*'
};
const MODE_HINT  = {
  pdf:  '.pdf',
  docx: '.doc .docx .docm .dotx .rtf .odt .txt .html and more',
  xlsx: '.xls .xlsx .xlsm .xlsb .csv .tsv .ods .txt and more'
};
const MODE_UNIT_LABEL = { pdf: 'pages', docx: 'paragraphs', xlsx: 'sheets' };
const MODE_FOOTER = {
  pdf:  'Text is extracted in reading order per line; very complex multi-column layouts may interleave. No data leaves this page.',
  docx: 'Compares paragraph text (including table cell text) in document order for .doc, .docx, .rtf, .odt, .html, .txt etc. Formatting/images are not compared. No data leaves this page.',
  xlsx: 'Compares cell values sheet by sheet and row by row for .xls, .xlsx, .xlsm, .xlsb, .csv, .ods etc. By default compares displayed values; tick "Compare formulas" to also flag cells whose formula changed even if the result looks the same. No data leaves this page.'
};

const state = { mode: 'pdf', doc1: null, doc2: null };

function setMode(m){
  state.mode = m;
  state.doc1 = null; state.doc2 = null;
  document.querySelectorAll('.mode-btn').forEach(b => b.classList.toggle('active', b.dataset.mode === m));

  document.getElementById('file1').accept = MODE_ACCEPT[m];
  document.getElementById('file2').accept = MODE_ACCEPT[m];
  document.getElementById('name1').textContent = 'Click or drop the older ' + MODE_NOUN[m] + ' here';
  document.getElementById('name2').textContent = 'Click or drop the newer ' + MODE_NOUN[m] + ' here';
  document.getElementById('meta1').textContent = '';
  document.getElementById('meta2').textContent = '';
  document.getElementById('hint1').textContent = MODE_HINT[m];
  document.getElementById('hint2').textContent = MODE_HINT[m];
  document.getElementById('optWordLabel').textContent = (m === 'xlsx') ? 'Highlight changed cells' : 'Highlight word-level changes';
  document.getElementById('optFormulaWrap').style.display = (m === 'xlsx') ? 'flex' : 'none';
  document.getElementById('cPagesLabel').textContent = MODE_UNIT_LABEL[m] + ' (File 1 / File 2)';
  document.getElementById('footerNote').textContent = MODE_FOOTER[m];

  document.getElementById('runBtn').disabled = true;
  document.getElementById('results').style.display = 'none';
  document.getElementById('status').textContent = '';
}
document.querySelectorAll('.mode-btn').forEach(btn=>{
  btn.addEventListener('click', ()=> setMode(btn.dataset.mode));
});

function setupDrop(dropId, fileId, nameId, metaId, key){
  const drop = document.getElementById(dropId);
  const input = document.getElementById(fileId);
  drop.addEventListener('click', ()=> input.click());
  drop.addEventListener('dragover', e=>{ e.preventDefault(); drop.classList.add('drag'); });
  drop.addEventListener('dragleave', ()=> drop.classList.remove('drag'));
  drop.addEventListener('drop', e=>{
    e.preventDefault(); drop.classList.remove('drag');
    if(e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0], nameId, metaId, key);
  });
  input.addEventListener('change', e=>{
    if(e.target.files.length) handleFile(e.target.files[0], nameId, metaId, key);
  });
}

async function handleFile(file, nameId, metaId, key){
  const ext = file.name.split('.').pop().toLowerCase();
  const allowed = MODE_EXT[state.mode];
  if(allowed && !allowed.includes(ext)){
    alert('Please choose a .' + allowed.join(' / .') + ' file for ' + MODE_NOUN[state.mode] + ' comparison.');
    return;
  }
  document.getElementById(nameId).textContent = file.name;
  document.getElementById(metaId).textContent = 'Reading...';

  try{
    let extracted;
    if(state.mode === 'pdf'){
      const buf = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
      extracted = await extractPdfGroups(pdf);
    } else if(state.mode === 'docx'){
      extracted = await extractWordGroups(file);
    } else {
      extracted = await extractXlsxGroups(file);
    }
    state[key] = { file, extracted };
    let metaText = extracted.count + ' ' + extracted.unitLabel + ' found';
    if(state.mode === 'xlsx') metaText += extracted.hasFormulas ? ' (formulas detected)' : ' (no formulas detected)';
    document.getElementById(metaId).textContent = metaText;
  } catch(err){
    document.getElementById(metaId).textContent = 'Could not read this file';
    alert('Error reading "' + file.name + '":\n' + err.message);
    return;
  }
  document.getElementById('runBtn').disabled = !(state.doc1 && state.doc2);
}

setupDrop('drop1','file1','name1','meta1','doc1');
setupDrop('drop2','file2','name2','meta2','doc2');

// ---------- PDF extraction (layout-aware) ----------
async function extractPdfGroups(pdf){
  const groups = [];
  for(let p = 1; p <= pdf.numPages; p++){
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const items = content.items.map(it => ({
      str: it.str, x: it.transform[4], y: it.transform[5], w: it.width
    })).filter(it => it.str.trim() !== '');

    items.sort((a,b)=> b.y - a.y || a.x - b.x);
    const rowsRaw = [];
    const TOL = 2.5;
    for(const it of items){
      let row = rowsRaw.find(r => Math.abs(r.y - it.y) <= TOL);
      if(!row){ row = { y: it.y, items: [] }; rowsRaw.push(row); }
      row.items.push(it);
    }
    rowsRaw.sort((a,b)=> b.y - a.y);

    const rows = [];
    for(const row of rowsRaw){
      row.items.sort((a,b)=> a.x - b.x);
      let line = '', lastEnd = null;
      for(const it of row.items){
        if(lastEnd === null){ line = it.str; }
        else{
          const gap = it.x - lastEnd;
          line += (gap > 8 ? '   ' : (line.endsWith(' ') ? '' : ' ')) + it.str;
        }
        lastEnd = it.x + it.w;
      }
      const text = line.replace(/\s+/g,' ').trim();
      if(text) rows.push({ label: 'p.'+p, text, cells: null });
    }
    const sig = rows.map(r=>r.text).join('\n');
    groups.push({ sig, rows });
  }
  return { groups, count: pdf.numPages, unitLabel: 'pages' };
}

// ---------- Word-family text extraction (docx/docm/dotx, doc, rtf, odt, html, txt) ----------
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const ODF_TEXT_NS = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';

function u8StartsWith(u8, s){
  for(let i = 0; i < s.length; i++){ if(u8[i] !== s.charCodeAt(i)) return false; }
  return true;
}
function decodeBytes(u8, label){
  try { return new TextDecoder(label).decode(u8); }
  catch(e){ let s = ''; for(let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]); return s; }
}
function textToParas(text){
  return text.split(/[\r\n\u2028\u2029]+/).map(t => t.replace(/\s+/g,' ').trim()).filter(Boolean);
}

// --- .docx / .docm / .dotx / .dotm ---
async function docxZipToParas(zip){
  const xml = await zip.file('word/document.xml').async('string');
  const dom = new DOMParser().parseFromString(xml, 'application/xml');
  const ps = dom.getElementsByTagNameNS(W_NS, 'p');
  const out = [];
  for(let i = 0; i < ps.length; i++){
    let text = '';
    const els = ps[i].getElementsByTagName('*');
    for(let j = 0; j < els.length; j++){
      const ln = els[j].localName;
      if(ln === 't') text += els[j].textContent;
      else if((ln === 'tab' || ln === 'br' || ln === 'cr') && els[j].parentNode && els[j].parentNode.localName === 'r') text += ' ';
    }
    text = text.replace(/\s+/g,' ').trim();
    if(text) out.push(text);
  }
  return out;
}

// --- .odt / .ott ---
function odfNodeText(node){
  let s = '';
  for(let c = node.firstChild; c; c = c.nextSibling){
    if(c.nodeType === 3) s += c.nodeValue;
    else if(c.nodeType === 1){
      const ln = c.localName;
      if(ln === 's' || ln === 'tab' || ln === 'line-break') s += ' ';
      else s += odfNodeText(c);
    }
  }
  return s;
}
async function odtZipToParas(zip){
  const xml = await zip.file('content.xml').async('string');
  const dom = new DOMParser().parseFromString(xml, 'application/xml');
  const els = dom.getElementsByTagName('*');
  const out = [];
  for(let i = 0; i < els.length; i++){
    const e = els[i];
    if(e.namespaceURI === ODF_TEXT_NS && (e.localName === 'p' || e.localName === 'h')){
      const t = odfNodeText(e).replace(/\s+/g,' ').trim();
      if(t) out.push(t);
    }
  }
  return out;
}

// --- legacy binary .doc / .dot (Word 97-2003) via CFB + piece table ---
function docBinaryToParas(u8){
  if(!(typeof XLSX !== 'undefined' && XLSX.CFB)) throw new Error('Compound-file reader not loaded.');
  const cfb = XLSX.CFB.read(u8, { type: 'array' });
  const wdE = XLSX.CFB.find(cfb, '/WordDocument');
  if(!wdE) throw new Error('This is an OLE file but not a Word document (no WordDocument stream). If it is an Excel file, use the Excel comparison tab.');
  const wd = Uint8Array.from(wdE.content);
  if(wd.length < 0x1AA) throw new Error('Unsupported/old Word format (Word 6/95). Please re-save as .docx or .rtf.');
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  if(dv.getUint16(0, true) !== 0xA5EC) throw new Error('Unsupported/old Word format (Word 6/95). Please re-save as .docx or .rtf.');
  const flags = dv.getUint16(0x0A, true);
  if(flags & 0x0100) throw new Error('This .doc is password-protected. Remove the password and try again.');
  const tE = XLSX.CFB.find(cfb, (flags & 0x0200) ? '/1Table' : '/0Table');
  if(!tE) throw new Error('Word table stream not found (corrupt .doc?).');
  const tbl = Uint8Array.from(tE.content);
  const tv = new DataView(tbl.buffer, tbl.byteOffset, tbl.byteLength);

  const ccpText = dv.getInt32(0x4C, true);
  const fcClx = dv.getUint32(0x1A2, true);
  const lcbClx = dv.getUint32(0x1A6, true);
  let pos = fcClx;
  const end = fcClx + lcbClx;
  while(pos < end && tbl[pos] === 1){ pos += 3 + tv.getUint16(pos + 1, true); }   // skip Prc entries
  if(tbl[pos] !== 2) throw new Error('Could not locate the text piece table in this .doc.');
  const lcb = tv.getUint32(pos + 1, true);
  pos += 5;
  const n = Math.floor((lcb - 4) / 12);
  const cpBase = pos, pcdBase = pos + 4 * (n + 1);

  let text = '';
  for(let i = 0; i < n; i++){
    const cpStart = tv.getUint32(cpBase + 4 * i, true);
    let cpEnd = tv.getUint32(cpBase + 4 * (i + 1), true);
    if(cpEnd > ccpText) cpEnd = ccpText;
    const len = cpEnd - cpStart;
    if(len <= 0) break;
    const fcRaw = tv.getUint32(pcdBase + 8 * i + 2, true);
    const compressed = (fcRaw & 0x40000000) !== 0;
    const fc = fcRaw & 0x3FFFFFFF;
    if(compressed) text += decodeBytes(wd.subarray(fc / 2, fc / 2 + len), 'windows-1252');
    else text += decodeBytes(wd.subarray(fc, fc + len * 2), 'utf-16le');
  }
  text = text
    .replace(/\x13[^\x14\x15]*\x15/g, '')     // fields with no result
    .replace(/\x13[^\x14\x15]*\x14/g, '')     // field instruction, keep result
    .replace(/[\x01\x08\x13\x14\x15\x1F]/g, '')
    .replace(/\x1E/g, '-')
    .replace(/[\x07\x0C]/g, '\n')              // cell/row end, page break
    .replace(/\x0B/g, ' ');
  return textToParas(text);
}

// --- .rtf ---
const RTF_SKIP = new Set(['fonttbl','colortbl','stylesheet','info','pict','header','footer','headerl','headerr','headerf',
  'footerl','footerr','footerf','object','themedata','colorschememapping','latentstyles','datastore','listtable',
  'listoverridetable','rsidtbl','generator','xmlnstbl','fldinst','private','template','revtbl','filetbl','pgptbl','mmathPr']);
const RTF_SYMBOLS = { bullet:'\u2022', endash:'\u2013', emdash:'\u2014', lquote:'\u2018', rquote:'\u2019',
  ldblquote:'\u201C', rdblquote:'\u201D', emspace:' ', enspace:' ', qmspace:' ' };

function rtfToParas(str){
  let out = '', i = 0;
  const n = str.length, stack = [], bytes = [];
  let skip = false, uc = 1, ucSkip = 0;
  const flush = () => { if(bytes.length){ out += decodeBytes(Uint8Array.from(bytes), 'windows-1252'); bytes.length = 0; } };
  const isAlpha = c => c !== undefined && /[a-zA-Z]/.test(c);
  const isDigit = c => c !== undefined && /[0-9]/.test(c);
  while(i < n){
    const ch = str[i];
    if(ch === '{'){ flush(); stack.push({ skip, uc }); i++; }
    else if(ch === '}'){ flush(); const s = stack.pop(); if(s){ skip = s.skip; uc = s.uc; } i++; }
    else if(ch === '\\'){
      i++;
      const c = str[i];
      if(c === undefined) break;
      if(c === '\\' || c === '{' || c === '}'){
        i++;
        if(!skip){ if(ucSkip > 0) ucSkip--; else { flush(); out += c; } }
      } else if(c === "'"){
        const hex = str.substr(i + 1, 2); i += 3;
        if(!skip){ if(ucSkip > 0) ucSkip--; else bytes.push(parseInt(hex, 16)); }
      } else if(c === '*'){ skip = true; i++; }
      else if(c === '~'){ i++; if(!skip){ flush(); out += ' '; } }
      else if(c === '_'){ i++; if(!skip){ flush(); out += '-'; } }
      else if(isAlpha(c)){
        let j = i; while(isAlpha(str[j])) j++;
        const word = str.slice(i, j);
        let k = j, param = null;
        if(str[k] === '-' || isDigit(str[k])){
          let m = k; if(str[m] === '-') m++;
          while(isDigit(str[m])) m++;
          param = parseInt(str.slice(k, m), 10); k = m;
        }
        if(str[k] === ' ') k++;
        i = k;
        flush();
        if(RTF_SKIP.has(word)){ skip = true; continue; }
        if(skip) continue;
        if(word === 'par' || word === 'line' || word === 'sect' || word === 'page' || word === 'cell' || word === 'row') out += '\n';
        else if(word === 'tab') out += ' ';
        else if(word === 'uc'){ uc = param === null ? 1 : param; }
        else if(word === 'u'){ let cp = param; if(cp < 0) cp += 65536; out += String.fromCharCode(cp); ucSkip = uc; }
        else if(RTF_SYMBOLS[word]) out += RTF_SYMBOLS[word];
      } else { i++; }
    } else if(ch === '\r' || ch === '\n'){ i++; }
    else {
      i++;
      if(skip) continue;
      if(ucSkip > 0){ ucSkip--; continue; }
      flush(); out += ch;
    }
  }
  flush();
  return textToParas(out);
}

// --- html-based ".doc" and plain html ---
function htmlToParas(str){
  if(/quoted-printable/i.test(str.slice(0, 4000))){
    const s = str.replace(/=\r?\n/g, '');
    try { str = decodeURIComponent(s.replace(/=([0-9A-Fa-f]{2})/g, '%$1')); } catch(e){}
    const h = str.search(/<html/i); if(h > 0) str = str.slice(h);
  }
  const doc = new DOMParser().parseFromString(str, 'text/html');
  doc.querySelectorAll('script,style,head').forEach(e => e.remove());
  const sel = 'p,h1,h2,h3,h4,h5,h6,li,td,th,pre,div,blockquote,caption';
  const out = [];
  doc.body.querySelectorAll(sel).forEach(el => {
    if(el.querySelector(sel)) return;
    const t = el.textContent.replace(/\s+/g,' ').trim();
    if(t) out.push(t);
  });
  if(!out.length){ const t = doc.body.textContent.replace(/\s+/g,' ').trim(); if(t) out.push(t); }
  return out;
}

// --- dispatcher: sniff content (not extension) ---
async function wordFileToParas(file){
  const u8 = new Uint8Array(await file.arrayBuffer());
  if(u8StartsWith(u8, '%PDF')) throw new Error('This is a PDF. Switch to the "PDF comparison" tab.');
  if(u8[0] === 0x50 && u8[1] === 0x4B){                      // ZIP container
    const zip = await JSZip.loadAsync(u8);
    if(zip.file('word/document.xml')) return docxZipToParas(zip);
    if(zip.file('content.xml') && !zip.file('xl/workbook.xml')) return odtZipToParas(zip);
    if(zip.file('xl/workbook.xml')) throw new Error('This is an Excel file. Use the "Excel sheet comparison" tab.');
    throw new Error('Unrecognised ZIP-based document (no Word/OpenDocument text found).');
  }
  if(u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0) return docBinaryToParas(u8);
  const head = decodeBytes(u8.subarray(0, 4096), 'windows-1252').replace(/^\uFEFF/, '').replace(/^\s+/, '');
  if(/^\{\\rtf/.test(head)) return rtfToParas(decodeBytes(u8, 'windows-1252'));
  if(/<html|<!doctype html|<body/i.test(head) || /^MIME-Version/i.test(head)) return htmlToParas(decodeBytes(u8, 'utf-8'));
  return textToParas(decodeBytes(u8, 'utf-8'));
}

async function extractWordGroups(file){
  const paras = await wordFileToParas(file);
  if(!paras.length) throw new Error('No readable text found in this file.');
  const rows = paras.map((text, i) => ({ label: '¶' + (i+1), text, cells: null }));
  return { groups: [{ sig: paras.join('\n'), rows }], count: rows.length, unitLabel: 'paragraphs' };
}

// ---------- Excel / spreadsheet extraction (xls, xlsx, xlsm, xlsb, csv, ods, ...) ----------
async function extractXlsxGroups(file){
  const u8 = new Uint8Array(await file.arrayBuffer());
  if(u8StartsWith(u8, '%PDF')) throw new Error('This is a PDF. Switch to the "PDF comparison" tab.');
  if(u8[0] === 0x50 && u8[1] === 0x4B){
    const zip = await JSZip.loadAsync(u8);
    if(zip.file('word/document.xml')) throw new Error('This is a Word file. Use the "Word file comparison" tab.');
  } else if(u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0){
    const cfb = XLSX.CFB.read(u8, { type: 'array' });
    if(XLSX.CFB.find(cfb, '/WordDocument')) throw new Error('This is a Word file. Use the "Word file comparison" tab.');
  }
  let wb;
  try { wb = XLSX.read(u8, { type: 'array', cellFormula: true }); }
  catch(err){ throw new Error('Could not open as a spreadsheet (password-protected or unsupported?): ' + err.message); }

  let anyFormulas = false;
  const groups = [];
  for(const sheetName of wb.SheetNames){
    const ws = wb.Sheets[sheetName];
    const rows = [];
    if(ws['!ref']){
      const range = XLSX.utils.decode_range(ws['!ref']);
      for(let r = range.s.r; r <= range.e.r; r++){
        const vals = [], formulas = [];
        let nonEmpty = false;
        for(let c = range.s.c; c <= range.e.c; c++){
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          let v = '';
          if(cell){
            if(cell.t === 'e') v = cell.w || '#ERR';
            else if(cell.w !== undefined) v = String(cell.w);
            else if(cell.v !== undefined) v = String(cell.v);
          }
          if(v !== '') nonEmpty = true;
          vals.push(v.trim());
          const f = cell && cell.f ? '=' + cell.f : '';
          if(f) anyFormulas = true;
          formulas.push(f);
        }
        if(nonEmpty){
          rows.push({ label: sheetName + '!R' + (r+1), vals, formulas });
        }
      }
    }
    groups.push({ rows, sheetName });
  }
  if(!groups.length) throw new Error('No sheets found in this file.');
  return { groups, count: wb.SheetNames.length, unitLabel: 'sheets', hasFormulas: anyFormulas };
}

// builds the comparable {text, cells} for each row, given the current "include formulas" choice
function materializeXlsxGroups(extracted, includeFormulas){
  return extracted.groups.map(g => {
    const rows = g.rows.map(r => {
      const cells = includeFormulas
        ? r.vals.map((v, i) => r.formulas[i] ? (v + '  {' + r.formulas[i] + '}') : v)
        : r.vals.slice();
      return { label: r.label, text: cells.join(' | '), cells };
    });
    return { sig: rows.map(r => r.text).join('\n'), rows };
  });
}

function normalize(s, ignoreCase, ignoreSpace){
  let t = s;
  if(ignoreSpace) t = t.replace(/\s+/g,' ').trim();
  if(ignoreCase) t = t.toLowerCase();
  return t;
}

// ---------- generic LCS diff over an array of comparable keys ----------
function lcsDiff(keysA, keysB){
  const n = keysA.length, m = keysB.length;
  const dp = new Int32Array((n+1)*(m+1));
  const idx = (i,j)=> i*(m+1)+j;
  for(let i=n-1;i>=0;i--){
    for(let j=m-1;j>=0;j--){
      if(keysA[i] === keysB[j]) dp[idx(i,j)] = dp[idx(i+1,j+1)] + 1;
      else dp[idx(i,j)] = Math.max(dp[idx(i+1,j)], dp[idx(i,j+1)]);
    }
  }
  const ops = [];
  let i=0,j=0;
  while(i<n && j<m){
    if(keysA[i]===keysB[j]){ ops.push({t:'eq', a:i, b:j}); i++; j++; }
    else if(dp[idx(i+1,j)] >= dp[idx(i,j+1)]){ ops.push({t:'del', a:i, b:-1}); i++; }
    else { ops.push({t:'ins', a:-1, b:j}); j++; }
  }
  while(i<n){ ops.push({t:'del', a:i, b:-1}); i++; }
  while(j<m){ ops.push({t:'ins', a:-1, b:j}); j++; }
  return ops;
}

// word-level diff for a single changed text pair -> returns [htmlA, htmlB]
function wordDiffHtml(lineA, lineB){
  const wa = lineA.split(/(\s+)/);
  const wb = lineB.split(/(\s+)/);
  const ops = lcsDiff(wa, wb);
  let htmlA = '', htmlB = '';
  for(const op of ops){
    if(op.t === 'eq'){
      const v = esc(wa[op.a]);
      htmlA += v; htmlB += v;
    } else if(op.t === 'del'){
      htmlA += '<span class="wdel">'+esc(wa[op.a])+'</span>';
    } else {
      htmlB += '<span class="wadd">'+esc(wb[op.b])+'</span>';
    }
  }
  return [htmlA, htmlB];
}

// cell-level diff for a changed spreadsheet row -> returns [htmlA, htmlB]
function cellDiffHtml(cellsA, cellsB){
  const ops = lcsDiff(cellsA, cellsB);
  const partsA = [], partsB = [];
  for(const op of ops){
    if(op.t === 'eq'){
      const v = esc(cellsA[op.a]);
      partsA.push(v); partsB.push(v);
    } else if(op.t === 'del'){
      partsA.push('<span class="wdel">'+esc(cellsA[op.a])+'</span>');
    } else {
      partsB.push('<span class="wadd">'+esc(cellsB[op.b])+'</span>');
    }
  }
  return [partsA.join(' | '), partsB.join(' | ')];
}

function esc(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ---------- generic group/row diff (shared by all 3 modes) ----------
function computeDiffRows(groupsA, groupsB, ignoreCase, ignoreSpace){
  const sigA = groupsA.map(g => normalize(g.sig, ignoreCase, ignoreSpace));
  const sigB = groupsB.map(g => normalize(g.sig, ignoreCase, ignoreSpace));
  const groupOps = lcsDiff(sigA, sigB);

  const finalRows = [];
  let block = null;
  const flushBlock = () => {
    if(!block) return;
    const subA = block.aIdx.flatMap(i => groupsA[i].rows);
    const subB = block.bIdx.flatMap(i => groupsB[i].rows);
    const keysA = subA.map(r => normalize(r.text, ignoreCase, ignoreSpace));
    const keysB = subB.map(r => normalize(r.text, ignoreCase, ignoreSpace));
    const ops = lcsDiff(keysA, keysB);
    for(let k=0;k<ops.length;k++){
      const op = ops[k];
      if(op.t === 'eq'){
        finalRows.push({ status:'eq', a: subA[op.a], b: subB[op.b] });
      } else if(op.t === 'del'){
        const next = ops[k+1];
        if(next && next.t === 'ins'){
          finalRows.push({ status:'chg', a: subA[op.a], b: subB[next.b] });
          k++;
        } else {
          finalRows.push({ status:'del', a: subA[op.a], b:null });
        }
      } else {
        finalRows.push({ status:'add', a:null, b: subB[op.b] });
      }
    }
    block = null;
  };

  for(const op of groupOps){
    if(op.t === 'eq'){
      flushBlock();
      const ga = groupsA[op.a].rows, gb = groupsB[op.b].rows;
      const len = Math.max(ga.length, gb.length);
      for(let i=0;i<len;i++){
        finalRows.push({ status:'eq', a: ga[i]||null, b: gb[i]||null });
      }
    } else {
      if(!block) block = { aIdx: [], bIdx: [] };
      if(op.t === 'del') block.aIdx.push(op.a); else block.bIdx.push(op.b);
    }
  }
  flushBlock();
  return finalRows;
}

// ---------- main compare ----------
document.getElementById('runBtn').addEventListener('click', runCompare);

async function runCompare(){
  const statusEl = document.getElementById('status');
  const runBtn = document.getElementById('runBtn');
  runBtn.disabled = true;
  statusEl.textContent = 'Comparing...';
  await new Promise(r=>setTimeout(r,10));

  const ignoreCase = document.getElementById('optCase').checked;
  const ignoreSpace = document.getElementById('optSpace').checked;
  const highlight = document.getElementById('optWord').checked;

  const includeFormulas = state.mode === 'xlsx' && document.getElementById('optFormula').checked;
  const groupsA = state.mode === 'xlsx' ? materializeXlsxGroups(state.doc1.extracted, includeFormulas) : state.doc1.extracted.groups;
  const groupsB = state.mode === 'xlsx' ? materializeXlsxGroups(state.doc2.extracted, includeFormulas) : state.doc2.extracted.groups;
  const finalRows = computeDiffRows(groupsA, groupsB, ignoreCase, ignoreSpace);

  renderResults(finalRows, state.doc1.extracted, state.doc2.extracted, highlight);
  statusEl.textContent = 'Done.';
  runBtn.disabled = false;
}

function renderResults(rows, extractedA, extractedB, highlight){
  document.getElementById('results').style.display = 'block';
  const body = document.getElementById('diffBody');
  body.innerHTML = '';

  let add=0, del=0, eq=0;
  const frag = document.createDocumentFragment();
  rows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.className = r.status === 'chg' ? 'del' : r.status;
    tr.dataset.status = r.status;

    const tdN = document.createElement('td'); tdN.className='ln'; tdN.textContent = idx+1;
    const tdS = document.createElement('td'); tdS.className='status';
    const tdP = document.createElement('td'); tdP.className='pg';
    const tdA = document.createElement('td'); tdA.className='txt';
    const tdB = document.createElement('td'); tdB.className='txt';

    if(r.status === 'eq'){
      tdS.textContent = 'Same'; eq++;
      tdP.textContent = (r.a? r.a.label:'') + (r.b && r.a && r.b.label!==r.a.label ? ' / '+r.b.label : (!r.a && r.b ? r.b.label:''));
      tdA.textContent = r.a ? r.a.text : '';
      tdB.textContent = r.b ? r.b.text : '';
    } else if(r.status === 'del'){
      tdS.textContent = 'Removed'; del++;
      tdP.textContent = r.a ? r.a.label : '';
      tdA.textContent = r.a ? r.a.text : '';
      tdB.textContent = '';
    } else if(r.status === 'add'){
      tdS.textContent = 'Added'; add++;
      tdP.textContent = r.b ? r.b.label : '';
      tdA.textContent = '';
      tdB.textContent = r.b ? r.b.text : '';
    } else if(r.status === 'chg'){
      tdS.textContent = 'Changed'; del++; add++;
      tdP.textContent = (r.a?r.a.label:'') + (r.b?' / '+r.b.label:'');
      let ha, hb;
      if(highlight && r.a.cells && r.b.cells){
        [ha, hb] = cellDiffHtml(r.a.cells, r.b.cells);
      } else if(highlight){
        [ha, hb] = wordDiffHtml(r.a.text, r.b.text);
      } else {
        ha = esc(r.a.text); hb = esc(r.b.text);
      }
      tdA.innerHTML = ha;
      tdB.innerHTML = hb;
    }

    tr.appendChild(tdN); tr.appendChild(tdS); tr.appendChild(tdP); tr.appendChild(tdA); tr.appendChild(tdB);
    frag.appendChild(tr);
  });
  body.appendChild(frag);

  document.getElementById('cAdd').textContent = add;
  document.getElementById('cDel').textContent = del;
  document.getElementById('cSame').textContent = eq;
  document.getElementById('cPages').textContent = extractedA.count + ' / ' + extractedB.count;

  applyFilter(document.querySelector('.filters button.active')?.dataset.f || 'all');
}

document.querySelectorAll('.filters button').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.filters button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    applyFilter(btn.dataset.f);
  });
});

function applyFilter(f){
  document.querySelectorAll('#diffBody tr').forEach(tr=>{
    const s = tr.dataset.status;
    let show = true;
    if(f==='changed') show = s!=='eq';
    else if(f==='add') show = (s==='add' || s==='chg');
    else if(f==='del') show = (s==='del' || s==='chg');
    tr.classList.toggle('hidden', !show);
  });
}

document.getElementById('exportBtn').addEventListener('click', ()=>{
  const rows = [['#','Status','Location','File 1','File 2']];
  document.querySelectorAll('#diffBody tr').forEach(tr=>{
    if(tr.classList.contains('hidden')) return;
    const cells = tr.querySelectorAll('td');
    rows.push([...cells].map(c => '"' + c.textContent.replace(/"/g,'""') + '"'));
  });
  const csv = rows.map(r=>r.join(',')).join('\n');
  const blob = new Blob([csv], {type:'text/csv'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'document_diff.csv';
  a.click();
});
</script>
</body>
</html>
