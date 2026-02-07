const repDropZone = document.getElementById("repDropZone");
const repFileInput = document.getElementById("repFileInput");
const repFileList = document.getElementById("repFileList");
const repFileGrid = document.getElementById("repFileGrid");
const repFileCount = document.getElementById("repFileCount");
const repClearFiles = document.getElementById("repClearFiles");
const repPairs = document.getElementById("repPairs");
const repAddPair = document.getElementById("repAddPair");
const repScanBtn = document.getElementById("repScanBtn");
const repExportCSV = document.getElementById("repExportCSV");
const repExportXLSX = document.getElementById("repExportXLSX");
const repResultsTable = document.getElementById("repResultsTable").querySelector("tbody");
const repProgress = document.getElementById("repProgress");
const repHead = document.querySelector('#repResultsTable thead');
const repSelectFile = document.getElementById('repSelectFile');
const repSelectType = document.getElementById('repSelectType');
const repSelectLayer = document.getElementById('repSelectLayer');
const repSelectRule = document.getElementById('repSelectRule');
const repFilterOriginal = document.getElementById('repFilterOriginalH');
const repFilterUpdated = document.getElementById('repFilterUpdatedH');

let repResults = [];
const repFilesMap = new Map();
let repSortKey = null;
let repSortDir = 'asc';

function ensureOnePair() {
  if (!repPairs.querySelector(".pair-row")) addPairRow();
}

function addPairRow() {
  const row = document.createElement("div");
  row.className = "pair-row";
  row.style.display = "flex";
  row.style.gap = "8px";
  row.style.marginBottom = "6px";
  const findInput = document.createElement("input");
  findInput.type = "text";
  findInput.placeholder = "查找关键字";
  const replaceInput = document.createElement("input");
  replaceInput.type = "text";
  replaceInput.placeholder = "替换为";
  const removeBtn = document.createElement("button");
  const isFirst = !repPairs.querySelector('.pair-row');
  removeBtn.textContent = isFirst ? "➕" : "✖";
  removeBtn.onclick = () => {
    if (removeBtn.textContent === "➕") {
      addPairRow();
    } else {
      row.remove(); ensureOnePair();
    }
  };
  row.appendChild(findInput);
  row.appendChild(replaceInput);
  row.appendChild(removeBtn);
  repPairs.appendChild(row);
}

function getPairs() {
  const rows = Array.from(repPairs.querySelectorAll(".pair-row"));
  return rows
    .map(r => {
      const inputs = r.querySelectorAll("input");
      return { find: inputs[0].value.trim(), to: inputs[1].value };
    })
    .filter(p => p.find);
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function displayRepFileList(files) {
  if (!repFileList) return;
  repFileList.style.display = "block";
  repFileGrid.innerHTML = "";
  repFileCount.textContent = files.length;
  files.forEach(f=>{ repFilesMap.set(f.name, f); });
  Array.from(files).forEach((file, index) => {
    const item = document.createElement("div");
    item.className = "file-item";
    item.innerHTML = `<div class="file-icon">📄</div><div class="file-info"><div class="file-name" title="${file.name}">${file.name}</div><div class="file-details"><span>大小: ${formatFileSize(file.size)}</span><span>类型: ${file.type || 'DXF文件'}</span><span>修改时间: ${file.lastModified ? new Date(file.lastModified).toLocaleString() : '未知'}</span></div></div><button class="remove-btn" title="移除文件">×</button>`;
    item.querySelector(".remove-btn").onclick = () => removeRepFile(index);
    repFileGrid.appendChild(item);
  });
}

function removeRepFile(index) {
  const files = Array.from(repFileInput.files);
  const newFiles = files.filter((_, i) => i !== index);
  const dt = new DataTransfer();
  newFiles.forEach(f => dt.items.add(f));
  repFileInput.files = dt.files;
  if (newFiles.length) {
    displayRepFileList(newFiles);
    const totalSize = newFiles.reduce((s, f) => s + f.size, 0);
    repDropZone.innerHTML = `✅ 已选择 ${newFiles.length} 个文件 (${formatFileSize(totalSize)})`;
  } else {
    repFileList.style.display = "none";
    repDropZone.innerHTML = "📁 将 DXF 文件拖拽到此处或点击选择文件";
  }
}

repDropZone.addEventListener("click", () => repFileInput.click());
repDropZone.addEventListener("dragover", e => { e.preventDefault(); repDropZone.classList.add("dragover"); });
repDropZone.addEventListener("dragleave", () => repDropZone.classList.remove("dragover"));
repDropZone.addEventListener("drop", async e => {
  e.preventDefault();
  repDropZone.classList.remove("dragover");
  const items = e.dataTransfer.items ? Array.from(e.dataTransfer.items) : [];
  if (items.length) {
    const files = await collectFilesFromItems(items);
    const dxfFiles = files.filter(f => f.name.toLowerCase().endsWith('.dxf'));
    if (!dxfFiles.length) { showAlert("⚠️ 未发现 DXF 文件", "warning"); return; }
    const dt = new DataTransfer();
    dxfFiles.forEach(f => dt.items.add(f));
    repFileInput.files = dt.files;
    const totalSize = dxfFiles.reduce((s, f) => s + f.size, 0);
    repDropZone.innerHTML = `✅ 已选择 ${dxfFiles.length} 个文件 (${formatFileSize(totalSize)})`;
    displayRepFileList(dxfFiles);
    showAlert(`✅ 成功添加 ${dxfFiles.length} 个 DXF 文件`, "success");
    return;
  }
  const files = Array.from(e.dataTransfer.files).filter(f => f.name.toLowerCase().endsWith(".dxf"));
  if (!files.length) { showAlert("⚠️ 请拖拽 DXF 文件！", "warning"); return; }
  const dt = new DataTransfer();
  files.forEach(f => dt.items.add(f));
  repFileInput.files = dt.files;
  const totalSize = files.reduce((s, f) => s + f.size, 0);
  repDropZone.innerHTML = `✅ 已选择 ${files.length} 个文件 (${formatFileSize(totalSize)})`;
  displayRepFileList(files);
  showAlert(`✅ 成功添加 ${files.length} 个 DXF 文件`, "success");
});

async function collectFilesFromItems(items){
  const entries = items.map(it => it.webkitGetAsEntry && it.webkitGetAsEntry()).filter(Boolean);
  const out = [];
  const walk = async (entry) => {
    if (entry.isFile) {
      await new Promise(resolve => entry.file(f => { out.push(f); resolve(); }));
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      await new Promise(resolve => reader.readEntries(async ents => { for(const e of ents){ await walk(e); } resolve(); }));
    }
  };
  for(const e of entries) await walk(e);
  return out;
}

repFileInput.addEventListener("change", e => {
  const files = Array.from(e.target.files);
  if (files.length) {
    const totalSize = files.reduce((s, f) => s + f.size, 0);
    repDropZone.innerHTML = `✅ 已选择 ${files.length} 个文件 (${formatFileSize(totalSize)})`;
    displayRepFileList(files);
    files.forEach(f=>{ repFilesMap.set(f.name, f); });
  } else {
    repFileList.style.display = "none";
  }
});

if (repClearFiles) {
  repClearFiles.addEventListener("click", () => {
    repFileInput.value = "";
    repFileList.style.display = "none";
    repDropZone.innerHTML = "📁 将 DXF 文件拖拽到此处或点击选择文件";
  });
}

ensureOnePair();

function setupRepControls() {
  const onInput = debounce(() => renderRepResults(), 200);
  [repFilterOriginal, repFilterUpdated].forEach(el => { if (el) el.addEventListener('input', onInput); });
  [repSelectFile, repSelectType, repSelectLayer, repSelectRule].forEach(el => { if (el) el.addEventListener('change', () => renderRepResults()); });
  // 双击表头清空筛选
  if (repHead) repHead.addEventListener('dblclick', () => {
    [repFilterOriginal, repFilterUpdated].forEach(el => { if (el) el.value = ''; });
    [repSelectFile, repSelectType, repSelectLayer, repSelectRule].forEach(el => { if (el) el.value = ''; });
    renderRepResults();
  });
  if (repHead) {
    repHead.querySelectorAll('th').forEach(th => {
      th.style.cursor = 'pointer';
      th.addEventListener('click', () => {
        const key = th.dataset.key;
        if (!key) return;
        if (repSortKey === key) {
          repSortDir = repSortDir === 'asc' ? 'desc' : 'asc';
        } else {
          repSortKey = key;
          repSortDir = 'asc';
        }
        renderRepResults();
      });
    });
  }
}

repScanBtn.addEventListener("click", async () => {
  const files = Array.from(repFileInput.files).filter(f => f.name.toLowerCase().endsWith('.dxf'));
  if (!files.length) { showAlert("⚠️ 请先选择 DXF 文件！", "warning"); return; }
  const pairs = getPairs();
  if (!pairs.length) { showAlert("⚠️ 请至少添加一组替换规则！", "warning"); return; }
  repResults = [];
  repResultsTable.innerHTML = "";
  repExportCSV.disabled = true;
  repExportXLSX.disabled = true;
  repScanBtn.disabled = true;
  repProgress.textContent = `准备处理 ${files.length} 个文件...`;
  const LARGE_SIZE = 6 * 1024 * 1024;
  const normalFiles = files.filter(f => (f.size || 0) <= LARGE_SIZE);
  const largeFiles = files.filter(f => (f.size || 0) > LARGE_SIZE);
  const parsed = await parseFilesWithWorkers(normalFiles);
  for (let i = 0; i < parsed.length; i++) {
    const p = parsed[i];
    repProgress.textContent = `处理中 (${i+1}/${parsed.length + largeFiles.length})：${p.file.name}`;
    if (p.error) { showAlert(`❌ 无法解析：${p.file.name}`, "error"); continue; }
    for (const e of p.entities) {
      let updated = e.text;
      let applied = [];
      for (const pr of pairs) {
        const re = new RegExp(escapeRegExp(pr.find), "gi");
        if (re.test(updated)) { updated = updated.replace(re, pr.to); applied.push(`${pr.find}→${pr.to}`); }
      }
      if (updated !== e.text) {
        repResults.push({ 文件名: p.file.name, 对象类型: e.type, 图层: e.layer || "", 原内容: e.text, 替换后: updated, 使用规则: applied.join("；"), __skip: false });
      }
    }
  }
  for (let j = 0; j < largeFiles.length; j++) {
    const f = largeFiles[j];
    repProgress.textContent = `处理中 (${parsed.length + j + 1}/${parsed.length + largeFiles.length})：${f.name}`;
    const text = await f.text();
    for (const pr of pairs) {
      const re = new RegExp(escapeRegExp(pr.find), "gi");
      let m;
      let idx = 0;
      while((m = text.slice(idx).match(re))){
        const pos = idx + m.index;
        const before = text.substr(pos, pr.find.length);
        const after = pr.to;
        const contextStart = Math.max(0, pos - 20);
        const contextEnd = Math.min(text.length, pos + pr.find.length + 20);
        const originalSnippet = text.slice(contextStart, contextEnd);
        const updatedSnippet = originalSnippet.replace(new RegExp(escapeRegExp(pr.find), 'i'), pr.to);
        repResults.push({ 文件名: f.name, 对象类型: '文本', 图层: '-', 原内容: originalSnippet, 替换后: updatedSnippet, 使用规则: `${pr.find}→${pr.to}`, __skip: false });
        idx = pos + pr.find.length;
      }
    }
  }
  renderRepResults();
  repScanBtn.disabled = false;
  document.getElementById('repConfirmBtn').disabled = repResults.length === 0;
});

let repPageIndex = 0; let repPageSize = 10;
const repPageSizeSel = document.getElementById('repPageSize');
const repPrevPage = document.getElementById('repPrevPage');
const repNextPage = document.getElementById('repNextPage');
const repPageInfo = document.getElementById('repPageInfo');
if (repPageSizeSel) repPageSizeSel.addEventListener('change', () => { repPageSize = parseInt(repPageSizeSel.value,10)||10; renderRepResults(); });
if (repPrevPage) repPrevPage.addEventListener('click', () => { if (repPageIndex>0){ repPageIndex--; renderRepResults(); } });
if (repNextPage) repNextPage.addEventListener('click', () => { const pages = Math.ceil(filteredRep().length/repPageSize); if (repPageIndex < pages-1){ repPageIndex++; renderRepResults(); } });

function filteredRep(){
  populateRepSelects();
  const f = {
    file: repSelectFile ? repSelectFile.value : '',
    type: repSelectType ? repSelectType.value : '',
    layer: repSelectLayer ? repSelectLayer.value : '',
    rule: repSelectRule ? repSelectRule.value : '',
    original: repFilterOriginal ? repFilterOriginal.value.trim() : '',
    updated: repFilterUpdated ? repFilterUpdated.value.trim() : ''
  };
  const inc = (s,q) => !q || (String(s||'').toLowerCase().includes(q.toLowerCase()));
  const eq = (s,v) => !v || String(s||'') === v;
  let filtered = repResults.filter(r =>
    eq(r.文件名,f.file) && eq(r.对象类型,f.type) && eq(r.图层||'-',f.layer) && inc(r.使用规则,f.rule) && inc(r.原内容,f.original) && inc(r.替换后,f.updated)
  );
  if (repSortKey) {
    const k = repSortKey; const dir = repSortDir === 'asc' ? 1 : -1;
    filtered.sort((a,b)=>{ const av=String(a[k]||'').toLowerCase(); const bv=String(b[k]||'').toLowerCase(); if(av<bv) return -1*dir; if(av>bv) return 1*dir; return 0; });
  }
  return filtered;
}

function renderRepResults() {
  repResultsTable.innerHTML = "";
  let filtered = filteredRep();
  if (!filtered.length) {
    repResultsTable.innerHTML = `<tr><td colspan="6" style="text-align:center; padding: 20px;">未发现需要替换的内容</td></tr>`;
    repProgress.textContent = "完成，未发现变化";
    return;
  }
  const pages = Math.ceil(filtered.length/repPageSize) || 1;
  if (repPageIndex >= pages) repPageIndex = pages-1;
  const start = repPageIndex*repPageSize;
  const pageRows = filtered.slice(start, start+repPageSize);
  if (repPageInfo) repPageInfo.textContent = `第 ${repPageIndex+1} / ${pages} 页，共 ${filtered.length} 项`;
  for (const r of pageRows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${r.文件名}</td><td>${r.对象类型}</td><td>${r.图层 || '-'}</td><td>${r.原内容}</td><td>${r.替换后}</td><td>${r.使用规则}</td><td><button class="toggle-btn">${r.__skip ? '恢复' : '撤销替换'}</button></td>`;
    tr.querySelector('.toggle-btn').addEventListener('click', () => { r.__skip = !r.__skip; renderRepResults(); });
    repResultsTable.appendChild(tr);
  }
  repExportCSV.disabled = false;
  repExportXLSX.disabled = false;
  repProgress.textContent = `完成，共 ${filtered.length} 项替换`;
}

repExportCSV.addEventListener("click", () => {
  let csv = "文件名,对象类型,图层,原内容,替换后,使用规则\n";
  repResults.forEach(r => { csv += `${r.文件名},${r.对象类型},${r.图层},"${r.原内容.replace(/"/g,'""')}","${r.替换后.replace(/"/g,'""')}",${r.使用规则}\n`; });
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "替换结果.csv";
  a.click();
  URL.revokeObjectURL(url);
});

repExportXLSX.addEventListener("click", async () => {
  if (typeof XLSX === "undefined") await loadSheetJS();
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(repResults);
  // 为"文件名"列添加相对超链接，工作簿与DXF在同一目录时可点击打开
  for (let i = 0; i < repResults.length; i++) {
    const addr = XLSX.utils.encode_cell({ r: i + 1, c: 0 });
    if (!ws[addr]) ws[addr] = { t: 's', v: repResults[i].文件名 };
    ws[addr].l = { Target: repResults[i].文件名 };
  }
  XLSX.utils.book_append_sheet(wb, ws, "替换结果");
  XLSX.writeFile(wb, "替换结果.xlsx");
});

window.addEventListener('load', setupRepControls);

function populateRepSelects() {
  const uniq = (arr) => Array.from(new Set(arr.filter(x => x !== undefined))).sort((a,b)=>String(a).localeCompare(String(b)));
  const files = uniq(repResults.map(r => r.文件名));
  const types = uniq(repResults.map(r => r.对象类型));
  const layers = uniq(repResults.map(r => r.图层 || '-'));
  const rules = uniq(repResults.map(r => r.使用规则));
  const fill = (sel, list) => {
    if (!sel) return;
    const prev = sel.value;
    sel.innerHTML = '<option value="">全部</option>' + list.map(v => `<option value="${v}">${v}</option>`).join('');
    if (list.includes(prev)) sel.value = prev;
  };
  fill(repSelectFile, files);
  fill(repSelectType, types);
  fill(repSelectLayer, layers);
  fill(repSelectRule, rules);
}

document.getElementById('repConfirmBtn').addEventListener('click', async () => {
  const files = Array.from(repFileInput.files);
  if(!files.length){ showAlert('⚠️ 没有文件', 'warning'); return; }
  const overwrite = document.getElementById('repOverwrite')?.checked;
  const byFile = new Map();
  repResults.filter(r=>!r.__skip).forEach(r => {
    const key = r.文件名;
    const arr = byFile.get(key) || [];
    arr.push({ before: r.原内容, after: r.替换后 });
    byFile.set(key, arr);
  });
  let changed = 0;
  if (overwrite && !window.showDirectoryPicker) {
    showAlert('ℹ️ 当前打开方式不支持直接保存到目录，已自动改为下载新文件', 'info');
  }
  if (overwrite && window.showDirectoryPicker) {
    try{
      const dir = await window.showDirectoryPicker();
      for(const f of files){
        const rules = byFile.get(f.name);
        if(!rules || !rules.length) continue;
        let text = await f.text();
        const counts = new Map();
        rules.forEach(r => { const k = JSON.stringify(r); counts.set(k, (counts.get(k)||0)+1); });
        for(const [k, n] of counts.entries()){ const { before, after } = JSON.parse(k); text = replaceLimited(text, before, after, n); }
        const fh = await dir.getFileHandle(f.name, { create: true });
        const ws = await fh.createWritable();
        await ws.write(text);
        await ws.close();
        changed++;
      }
      showAlert(`✅ 已写入 ${changed} 个文件到所选目录`, 'success');
      return;
    }catch(e){
      showAlert(`❌ 写入失败：${e}`, 'error');
    }
  }
  for(const f of files){
    const rules = byFile.get(f.name);
    if(!rules || !rules.length) continue;
    let text = await f.text();
    const counts = new Map();
    rules.forEach(r => { const k = JSON.stringify(r); counts.set(k, (counts.get(k)||0)+1); });
    for(const [k, n] of counts.entries()){ const { before, after } = JSON.parse(k); text = replaceLimited(text, before, after, n); }
    const blob = new Blob([text], { type: 'application/dxf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = f.name;
    a.click(); URL.revokeObjectURL(url);
    changed++;
  }
  showAlert(`✅ 已生成 ${changed} 个替换后的文件`, 'success');
});

function replaceLimited(text, before, after, count){
  if (!before) return text;
  let remaining = count;
  let idx = 0;
  const lowerBefore = before; // exact match sequence
  while(remaining>0){
    const pos = text.indexOf(lowerBefore, idx);
    if (pos === -1) break;
    text = text.slice(0, pos) + after + text.slice(pos + lowerBefore.length);
    idx = pos + after.length;
    remaining--;
  }
  return text;
}

async function parseFilesWithWorkers(files){
  let workerTestOk = true;
  try {
    const t = new Worker('dxf_worker.js'); t.terminate();
  } catch (e) {
    workerTestOk = false;
  }
  if (!workerTestOk) {
    const resultsArr = [];
    if (typeof showAlert === 'function') {
      showAlert('ℹ️ 当前以本地文件方式打开，浏览器禁用多线程解析，已自动切换为单线程处理', 'info');
    }
    for (const file of files) {
      try {
        const text = await file.text();
        const parser = new DxfParser();
        let dxf;
        try {
          dxf = parser.parseSync(text);
        } catch(err) {
          resultsArr.push({ file, error: String(err) });
          continue;
        }
        const entities = (dxf.entities || []).map(ent => {
          let content = '';
          if (ent.type === 'TEXT' || ent.type === 'MTEXT' || ent.type === 'ATTRIB') content = ent.text || '';
          else if (ent.type === 'INSERT') content = ent.name || '';
          return { type: ent.type, layer: ent.layer || '', text: content };
        }).filter(e => e.text);
        resultsArr.push({ file, entities });
      } catch(err) {
        resultsArr.push({ file, error: String(err) });
      }
    }
    return resultsArr;
  }
  const workers = [];
  const maxWorkers = Math.min((navigator.hardwareConcurrency || 4), Math.max(1, files.length));
  for(let i=0;i<maxWorkers;i++) workers.push(new Worker('dxf_worker.js'));
  let nextId = 1;
  const tasks = files.map(file => ({ file, id: nextId++ }));
  const resultsArr = [];
  const queue = tasks.slice();
  const runOnWorker = (worker) => new Promise(resolve => {
    const pump = () => {
      const task = queue.shift();
      if(!task){ resolve(); return; }
      fileToText(task.file).then(text => {
        const onMsg = (ev) => {
          const data = ev.data || {};
          if(data.id !== task.id) return;
          worker.removeEventListener('message', onMsg);
          if(data.ok){ resultsArr.push({ file: task.file, entities: data.entities }); }
          else { resultsArr.push({ file: task.file, error: data.error }); }
          pump();
        };
        worker.addEventListener('message', onMsg);
        worker.postMessage({ id: task.id, op: 'parse', text });
      }).catch(err => { resultsArr.push({ file: task.file, error: String(err) }); pump(); });
    };
    pump();
  });
  await Promise.all(workers.map(w => runOnWorker(w)));
  workers.forEach(w => w.terminate());
  return resultsArr;
}

function fileToText(file){ return file.text(); }
async function repOpenFileInCAD(filename){
  const f = repFilesMap.get(filename);
  if(!f){ showAlert(`❌ 未找到文件：${filename}`, 'error'); return; }
  try{
    // 复用 app.js 的检测逻辑（若未加载则直接走下载）
    let available = false;
    try{
      const res = await fetch('http://localhost:8765/open', { method: 'OPTIONS' });
      available = res.ok;
    }catch(e){ available = false; }
    if(!available){
      const url = URL.createObjectURL(f);
      const a = document.createElement('a');
      a.href = url;
      a.download = f.name;
      a.click();
      URL.revokeObjectURL(url);
      showAlert('ℹ️ 未检测到打开服务，已为你下载文件，请手动用AutoCAD打开', 'info');
      return;
    }
    const buf = await f.arrayBuffer();
    const b64 = btoa(String.fromCharCode(...new Uint8Array(buf)));
    const res = await fetch('http://localhost:8765/open', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: f.name, content: b64 })
    });
    if(res.ok) showAlert(`✅ 已请求在本机打开：${f.name}`, 'success');
    else showAlert(`❌ 打开失败：${f.name}`, 'error');
  }catch(err){ showAlert(`❌ 打开失败：${err}`, 'error'); }
}
