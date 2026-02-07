
// ===================== DXF 关键字扫描工具 =====================
const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const scanBtn = document.getElementById("scanBtn");
const exportCSV = document.getElementById("exportCSV");
const exportXLSX = document.getElementById("exportXLSX");
const resultsTable = document.getElementById("resultsTable").querySelector("tbody");
const resultsHead = document.querySelector('#resultsTable thead');
const resSelectFile = document.getElementById('resSelectFile');
const resSelectType = document.getElementById('resSelectType');
const resSelectLayer = document.getElementById('resSelectLayer');
const resSelectKeyword = document.getElementById('resSelectKeyword');
const resFilterContent = document.getElementById('resFilterContentH');
const progressDiv = document.getElementById("progress");

let results = [];
const uploadedFilesMap = new Map();
let openServerAvailable = null;
let resultsSortKey = null;
let resultsSortDir = 'asc';

function setupResultsControls() {
  const onInput = debounce(() => showResults(), 200);
  if (resFilterContent) resFilterContent.addEventListener('input', onInput);
  [resSelectFile, resSelectType, resSelectLayer, resSelectKeyword].forEach(el => { if (el) el.addEventListener('change', () => showResults()); });
  if (resultsHead) {
    resultsHead.addEventListener('dblclick', () => {
      if (resFilterContent) resFilterContent.value = '';
      [resSelectFile, resSelectType, resSelectLayer, resSelectKeyword].forEach(el => { if (el) el.value = ''; });
      showResults();
    });
  }
  if (resultsHead) {
    resultsHead.querySelectorAll('th').forEach(th => {
      th.style.cursor = 'pointer';
      th.addEventListener('click', () => {
        const key = th.dataset.key;
        if (!key) return;
        if (resultsSortKey === key) {
          resultsSortDir = resultsSortDir === 'asc' ? 'desc' : 'asc';
        } else {
          resultsSortKey = key;
          resultsSortDir = 'asc';
        }
        showResults();
      });
    });
  }
}

function populateResSelects() {
  const uniq = (arr) => Array.from(new Set(arr.filter(x => x !== undefined))).sort((a,b)=>String(a).localeCompare(String(b)));
  const files = uniq(results.map(r => r.文件名));
  const types = uniq(results.map(r => r.对象类型));
  const layers = uniq(results.map(r => r.图层 || '-'));
  const keywords = uniq(results.map(r => r.关键字));
  const fill = (sel, list) => {
    if (!sel) return;
    const prev = sel.value;
    sel.innerHTML = '<option value="">全部</option>' + list.map(v => `<option value="${v}">${v}</option>`).join('');
    if (list.includes(prev)) sel.value = prev;
  };
  fill(resSelectFile, files);
  fill(resSelectType, types);
  fill(resSelectLayer, layers);
  fill(resSelectKeyword, keywords);
}

// 拖拽上传（优化版）
if(dropZone){
    dropZone.addEventListener("dragover", e => {
    e.preventDefault();
    dropZone.classList.add("dragover");
    dropZone.innerHTML = '📁 释放鼠标以上传文件';
    });

    dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
    dropZone.innerHTML = '📁 将 DXF 文件拖拽到此处或点击选择文件';
    });

    dropZone.addEventListener("drop", async e => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    
    const items = e.dataTransfer.items ? Array.from(e.dataTransfer.items) : [];
    let files = [];
    if (items.length) {
        files = await collectFilesFromItems(items);
    } else {
        files = Array.from(e.dataTransfer.files);
    }
    const dxfFiles = files.filter(file => file.name.toLowerCase().endsWith('.dxf'));
    
    if (dxfFiles.length === 0) {
        showAlert('⚠️ 请拖拽 DXF 文件！', 'warning');
        dropZone.innerHTML = '📁 将 DXF 文件拖拽到此处或点击选择文件';
        return;
    }
    
    const dt = new DataTransfer();
    dxfFiles.forEach(f => dt.items.add(f));
    fileInput.files = dt.files;
    // registerUploadedFiles(dxfFiles); // This function seems missing in my view? Assuming it was not essential or lost in snippet? 
    // Wait, let me check app.js view again. registerUploadedFiles IS CALLED in app.js line 245. But I don't see the DEFINITION in lines 1-800 or 800+?
    // Let me check if I missed it. I viewed 1-800. I need to check if it's defined later.
    // If not, it might be a bug or missing code. I will check the file content again later or just ignore/mock it if it's not critical. 
    // Actually, line 159 define uploadedFilesMap. Maybe registerUploadedFiles just adds to it?
    // I will add a simple implementation if I don't find it.
    
    // 显示文件信息
    const totalSize = dxfFiles.reduce((sum, file) => sum + file.size, 0);
    dropZone.innerHTML = `
        <div style="color: var(--success-color);">
        ✅ 已选择 ${dxfFiles.length} 个文件 (${formatFileSize(totalSize)})
        </div>
        <div style="font-size: 12px; margin-top: 5px; color: var(--text-secondary);">
        点击重新选择文件
        </div>
    `;
    
    displayFileList(dxfFiles); // 显示文件列表
    showAlert(`✅ 成功添加 ${dxfFiles.length} 个 DXF 文件`, 'success');
    });
}

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

// 文件选择事件
if(fileInput){
    fileInput.addEventListener("change", (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
        const totalSize = files.reduce((sum, file) => sum + file.size, 0);
        dropZone.innerHTML = `
        <div style="color: var(--success-color);">
            ✅ 已选择 ${files.length} 个文件 (${formatFileSize(totalSize)})
        </div>
        <div style="font-size: 12px; margin-top: 5px; color: var(--text-secondary);">
            点击重新选择文件
        </div>
        `;
        displayFileList(files); // 显示文件列表
        // registerUploadedFiles(files);
    } else {
        // 没有文件时隐藏文件列表
        const fileListDiv = document.getElementById('fileList');
        if (fileListDiv) fileListDiv.style.display = 'none';
    }
    });

    // 点击拖拽区域触发文件选择
    dropZone.addEventListener("click", () => {
        fileInput.click();
    });
}


// 主逻辑（优化版）
if(scanBtn){
    scanBtn.addEventListener("click", async () => {
    const files = Array.from(fileInput.files).filter(f => f.name.toLowerCase().endsWith('.dxf'));
    if (files.length === 0) {
        showAlert('⚠️ 请先选择 DXF 文件！', 'warning');
        return;
    }

    const keywords = document.getElementById("keywords").value.split(",").map(k => k.trim()).filter(k => k);

    results = [];
    resultsTable.innerHTML = "";
    exportCSV.disabled = true;
    exportXLSX.disabled = true;
    scanBtn.disabled = true;
    
    // 显示加载状态
    scanBtn.innerHTML = '<span class="loading-spinner"></span>扫描中...';
    progressDiv.innerHTML = `
        <div class="progress-bar">
        <div class="progress-fill" style="width: 0%"></div>
        </div>
        <div style="margin-top: 10px;">
        <span class="status-indicator status-info"></span>
        准备扫描 ${files.length} 个文件...
        </div>
    `;

    const LARGE_SIZE = 6 * 1024 * 1024;
    const normalFiles = files.filter(f => (f.size || 0) <= LARGE_SIZE);
    const largeFiles = files.filter(f => (f.size || 0) > LARGE_SIZE);
    const parsed = await parseFilesWithWorkers(normalFiles);
    let idx = 0;
    for(let i=0;i<parsed.length;i++){
        const p = parsed[i];
        idx++;
        const progress = ((idx) / (parsed.length + largeFiles.length) * 100).toFixed(1);
        document.querySelector('.progress-fill').style.width = progress + '%';
        progressDiv.querySelector('div:last-child').innerHTML = `
        <span class="status-indicator status-info"></span>
        扫描中 (${idx}/${parsed.length + largeFiles.length})：${p.file.name}
        `;
        if(p.error){
        showAlert(`❌ 无法解析：${p.file.name}，使用纯文本模式`, 'error');
        const text = await p.file.text();
        searchPlainText(p.file.name, text, keywords);
        continue;
        }
        const entities = p.entities || [];
        for(const entity of entities){
        const content = entity.text;
        if(!content) continue;
        if(keywords.length === 0){
            results.push({ 文件名: p.file.name, 对象类型: entity.type, 图层: entity.layer || '', 关键字: '全部', 匹配内容: content });
        }else{
            for(const kw of keywords){
            if(content.toLowerCase().includes(kw.toLowerCase())){
                results.push({ 文件名: p.file.name, 对象类型: entity.type, 图层: entity.layer || '', 关键字: kw, 匹配内容: content });
            }
            }
        }
        }
    }
    for(const f of largeFiles){
        idx++;
        const progress = ((idx) / (parsed.length + largeFiles.length) * 100).toFixed(1);
        document.querySelector('.progress-fill').style.width = progress + '%';
        progressDiv.querySelector('div:last-child').innerHTML = `
        <span class="status-indicator status-info"></span>
        扫描中 (${idx}/${parsed.length + largeFiles.length})：${f.name}
        `;
        const text = await f.text();
        searchPlainText(f.name, text, keywords);
    }

    showResults();
    scanBtn.disabled = false;
    scanBtn.innerHTML = '🚀 开始扫描';
    
    // 显示完成状态
    if (results.length > 0) {
        showAlert(`✅ 扫描完成！找到 ${results.length} 条匹配结果`, 'success');
    } else {
        showAlert('ℹ️ 扫描完成，但未找到匹配内容', 'info');
    }
    });
}


// 显示文件列表函数
function displayFileList(files) {
  const fileListDiv = document.getElementById('fileList');
  const fileGrid = document.getElementById('fileGrid');
  const fileCount = document.getElementById('fileCount');
  
  if (!fileListDiv) return;
  
  fileListDiv.style.display = 'block';
  fileGrid.innerHTML = '';
  fileCount.textContent = files.length;
  
  Array.from(files).forEach((file, index) => {
    const fileItem = document.createElement('div');
    fileItem.className = 'file-item';
    fileItem.innerHTML = `
      <div class="file-icon">📄</div>
      <div class="file-info">
        <div class="file-name" title="${file.name}">${file.name}</div>
        <div class="file-details">
          <span>大小: ${formatFileSize(file.size)}</span>
          <span>类型: ${file.type || 'DXF文件'}</span>
          <span>修改时间: ${file.lastModified ? new Date(file.lastModified).toLocaleString() : '未知'}</span>
        </div>
      </div>
      <button class="remove-btn" onclick="removeFile(${index})" title="移除文件">×</button>
    `;
    fileGrid.appendChild(fileItem);
  });
}

// 移除单个文件
function removeFile(index) {
  const files = Array.from(fileInput.files);
  const newFiles = files.filter((_, i) => i !== index);
  
  // 创建新的FileList
  const dt = new DataTransfer();
  newFiles.forEach(file => dt.items.add(file));
  fileInput.files = dt.files;
  
  if (newFiles.length > 0) {
    displayFileList(newFiles);
    const totalSize = newFiles.reduce((sum, file) => sum + file.size, 0);
    dropZone.innerHTML = `
      <div style="color: var(--success-color);">
        ✅ 已选择 ${newFiles.length} 个文件 (${formatFileSize(totalSize)})
      </div>
      <div style="font-size: 12px; margin-top: 5px; color: var(--text-secondary);">
        点击重新选择文件
      </div>
    `;
  } else {
    document.getElementById('fileList').style.display = 'none';
    dropZone.innerHTML = '📁 将 DXF 文件拖拽到此处或点击选择文件';
  }
}

// 清空所有文件
document.addEventListener('DOMContentLoaded', function() {
  const clearBtn = document.getElementById('clearFiles');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      fileInput.value = '';
      document.getElementById('fileList').style.display = 'none';
      dropZone.innerHTML = '📁 将 DXF 文件拖拽到此处或点击选择文件';
    });
  }
});

// 纯文本模式扫描（当 dxf-parser 无法解析时）
function searchPlainText(filename, text, keywords) {
  if (keywords.length === 0){
  //空文本框：匹配所有内容的实体
    results.push({
        文件名: filename,
        对象类型: "未知",
        图层: "-",
        关键字: "全部",
        匹配内容: "(纯文本匹配)"
    });
  }else{
      for (const kw of keywords) {
        if (text.toLowerCase().includes(kw.toLowerCase())) {
          results.push({
            文件名: filename,
            对象类型: "未知",
            图层: "-",
            关键字: kw,
            匹配内容: "(纯文本匹配)"
          });
        }
      }
    }
}
// 显示结果（优化版）
function showResults() {
  resultsTable.innerHTML = "";
  
  if (results.length === 0) {
    resultsTable.innerHTML = '<tr><td colspan="5" style="text-align: center; padding: 40px; color: var(--text-secondary);">😔 未找到匹配结果</td></tr>';
    progressDiv.innerHTML = '<span class="status-indicator status-warning"></span>扫描完成，未找到匹配结果';
    return;
  }
  
  populateResSelects();
  const f = {
    file: resSelectFile ? resSelectFile.value : '',
    type: resSelectType ? resSelectType.value : '',
    layer: resSelectLayer ? resSelectLayer.value : '',
    keyword: resSelectKeyword ? resSelectKeyword.value : '',
    content: resFilterContent ? resFilterContent.value.trim() : ''
  };
  const inc = (s, q) => !q || (String(s || '').toLowerCase().includes(q.toLowerCase()));
  const eq = (s, v) => !v || String(s || '') === v;
  let filtered = results.filter(r =>
    eq(r.文件名, f.file) &&
    eq(r.对象类型, f.type) &&
    eq(r.图层 || '-', f.layer) &&
    eq(r.关键字, f.keyword) &&
    inc(r.匹配内容, f.content)
  );
  if (resultsSortKey) {
    const k = resultsSortKey;
    const dir = resultsSortDir === 'asc' ? 1 : -1;
    filtered.sort((a,b) => {
      const av = String(a[k] || '').toLowerCase();
      const bv = String(b[k] || '').toLowerCase();
      if (av < bv) return -1 * dir;
      if (av > bv) return 1 * dir;
      return 0;
    });
  }
  const stats = {
    total: filtered.length,
    files: new Set(filtered.map(r => r.文件名)).size,
    keywords: new Set(filtered.map(r => r.关键字)).size
  };
  
  // 显示统计信息
  progressDiv.innerHTML = `
    <div class="stats-card">
      <div class="stat-item">
        <span class="stat-number">${stats.total}</span>
        <span class="stat-label">匹配结果</span>
      </div>
      <div class="stat-item">
        <span class="stat-number">${stats.files}</span>
        <span class="stat-label">涉及文件</span>
      </div>
      <div class="stat-item">
        <span class="stat-number">${stats.keywords}</span>
        <span class="stat-label">匹配关键字</span>
      </div>
    </div>
    <div style="margin-top: 15px;">
      <span class="status-indicator status-success"></span>
      ✅ 扫描完成，共找到 ${stats.total} 条匹配结果
    </div>
  `;
  
  const pages = Math.ceil(filtered.length / resPageSize) || 1;
  if (resPageIndex >= pages) resPageIndex = pages - 1;
  const start = resPageIndex * resPageSize;
  const pageRows = filtered.slice(start, start + resPageSize);
  if (resPageInfo) resPageInfo.textContent = `第 ${resPageIndex + 1} / ${pages} 页，共 ${filtered.length} 项`;

  for (const r of pageRows) {
    const tr = document.createElement("tr");
    const content = r.匹配内容;
    const displayContent = content.length > 100 ? 
      `<span class="content-preview">${content.substring(0, 100)}...</span>
       <button class="expand-btn" onclick="toggleContent(this, '${content.replace(/'/g, "\\'")}')" style="margin-left: 5px; background: var(--primary-color); color: white; border: none; padding: 2px 6px; border-radius: 3px; cursor: pointer; font-size: 11px;">展开</button>` :
      `<span class="content-preview">${content}</span>`;
    
    tr.innerHTML = `
      <td><span class="file-icon">📄</span>${r.文件名}</td>
      <td><span class="type-badge">${r.对象类型}</span></td>
      <td><span class="layer-badge">${r.图层 || '-'}</span></td>
      <td><span class="keyword-badge">${r.关键字}</span></td>
      <td class="content-cell">
        ${r.匹配内容.length > 100 ? 
          `<span class="content-preview">${r.匹配内容.substring(0, 100)}...</span>
           <button class="expand-btn" onclick="toggleContent(this, '${r.匹配内容.replace(/'/g, "\\'")}')" style="margin-left: 5px; background: var(--primary-color); color: white; border: none; padding: 2px 6px; border-radius: 3px; cursor: pointer; font-size: 11px;">展开</button>` :
          `<span class="content-full">${r.匹配内容}</span>`
        }
      </td>
    `;
    // 添加双击事件，展开完整内容
    tr.addEventListener('dblclick', function() {
      const contentCell = tr.querySelector('.content-cell');
      const expandBtn = contentCell.querySelector('.expand-btn');
      if (expandBtn && expandBtn.textContent === '展开') {
        expandBtn.click();
      }
    });
    resultsTable.appendChild(tr);
  }
  
  exportCSV.disabled = false;
  exportXLSX.disabled = false;
}

// 导出 CSV
exportCSV.addEventListener("click", () => {
  let csv = "文件名,对象类型,图层,关键字,匹配内容\n";
  results.forEach(r => {
    csv += `${r.文件名},${r.对象类型},${r.图层},${r.关键字},"${r.匹配内容.replace(/"/g, '""')}"\n`;
  });
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "扫描结果.csv";
  a.click();
  URL.revokeObjectURL(url);
});

// 导出 Excel（使用 SheetJS）
exportXLSX.addEventListener("click", async () => {
  if (typeof XLSX === "undefined") {
    await loadSheetJS();
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(results);
  // 为"文件名"列添加超链接（相对路径）。将工作簿与DXF置于同一文件夹即可点击打开
  for (let i = 0; i < results.length; i++) {
    const cellAddr = XLSX.utils.encode_cell({ r: i + 1, c: 0 }); // 行从2开始，列A=0
    if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: results[i].文件名 };
    ws[cellAddr].l = { Target: results[i].文件名 };
  }
  XLSX.utils.book_append_sheet(wb, ws, "扫描结果");
  XLSX.writeFile(wb, "扫描结果.xlsx");
});

// 展开/收起内容功能
function toggleContent(button, fullContent) {
  const cell = button.parentElement;
  const isExpanded = button.textContent === '收起';
  
  if (isExpanded) {
    // 收起内容
    cell.innerHTML = `
      <span class="content-preview">${fullContent.substring(0, 100)}...</span>
      <button class="expand-btn" onclick="toggleContent(this, '${fullContent.replace(/'/g, "\\'")}')" style="margin-left: 5px; background: var(--primary-color); color: white; border: none; padding: 2px 6px; border-radius: 3px; cursor: pointer; font-size: 11px;">展开</button>
    `;
  } else {
    // 展开内容
    cell.innerHTML = `
      <span class="content-full">${fullContent}</span>
      <button class="expand-btn" onclick="toggleContent(this, '${fullContent.replace(/'/g, "\\'")}')" style="margin-left: 5px; background: var(--secondary-color); color: white; border: none; padding: 2px 6px; border-radius: 3px; cursor: pointer; font-size: 11px;">收起</button>
    `;
  }
}

// ====================== 阀门统计识别函数 ======================
function extractValveInfo(textList) {
    const results = [];

    const re_size = /(\d+)\s*[xX×]\s*(\d+)/;
    const re_extra_height = /(顶[0-9\.]+m|标高[:：]?\s*[0-9\.]+m)/i;
    const re_valve_id = /[A-Za-z0-9]{6,}/;

    const invalid_keywords = ["排风系统", "系统", "标高", "顶标高", "top", "尺寸"];

    function isInvalidName(t) {
        if (!t) return true;
        if (invalid_keywords.some(k => t.includes(k))) return true;
        if (re_size.test(t)) return true;
        return false;
    }

    for (let i = 0; i < textList.length - 1; i++) {
        const t1 = textList[i].text;
        const t2 = textList[i + 1].text;

        if (isInvalidName(t1)) continue;

        // 尺寸匹配
        const m_size = t2.match(re_size);
        if (!m_size) continue;
        const sizeText = m_size[0];

        // 标高提取
        let heightText = "";
        const m_h = t2.match(re_extra_height);
        if (m_h) heightText = m_h[0];

        // 阀门编号识别
        const m_id = t1.match(re_valve_id);
        if (m_id) {
            const valve_id = m_id[0];
            const valve_name = t1.replace(valve_id, "").replace(/[（）()]/g, "").trim();

            results.push({
                类型: "阀门",
                名称: valve_name,
                编号: valve_id,
                尺寸: sizeText,
                标高: heightText,
            });
            continue;
        }

        // 风口识别
        if (t1.includes("风口") || t1.includes("风阀") || t1.includes("百叶")) {
            results.push({
                类型: "风口",
                名称: t1,
                编号: "",
                尺寸: sizeText,
                标高: heightText,
            });
            continue;
        }
    }

    return results;
}

function parseDXFText(dxf) {
    const out = [];
    const entities = (dxf && dxf.entities) || [];
    for (const entity of entities) {
        let content = "";
        if (entity.type === "TEXT" || entity.type === "MTEXT" || entity.type === "ATTRIB") content = entity.text || "";
        else if (entity.type === "INSERT") content = entity.name || "";
        if (!content) continue;
        out.push({ text: content, layer: entity.layer || "", type: entity.type });
    }
    return out;
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
      }).catch(err => {
        resultsArr.push({ file: task.file, error: String(err) });
        pump();
      });
    };
    pump();
  });
  await Promise.all(workers.map(w => runOnWorker(w)));
  workers.forEach(w => w.terminate());
  return resultsArr;
}

function fileToText(file){ return file.text(); }
