// ================= 页面标签切换 =================
const tabs = document.querySelectorAll(".tab-btn");
const sections = document.querySelectorAll(".tab-content");

tabs.forEach(tab => {
  tab.addEventListener("click", () => {
    tabs.forEach(t => t.classList.remove("active"));
    sections.forEach(s => s.classList.remove("active"));
    tab.classList.add("active");
    document.getElementById(tab.dataset.tab).classList.add("active");
  });
});

// ================= 子标签切换 =================
const subTabs = document.querySelectorAll(".subtab-btn");
const subContents = document.querySelectorAll(".subtab-content");
subTabs.forEach(btn => {
  btn.addEventListener("click", () => {
    subTabs.forEach(b => b.classList.remove("active"));
    subContents.forEach(c => c.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.sub).classList.add("active");
  });
});

// ================= 主题切换功能 =================
const bgSelector = document.getElementById("bgColorSelector");

function setTheme(theme) {
  const body = document.body;

  if (theme === 'auto') {
    // 根据系统设置自动切换
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    theme = prefersDark ? 'dark' : 'light';
  }

  body.setAttribute('data-theme', theme);
  localStorage.setItem("toolbox-theme", theme);

  // 更新选择器显示
  const options = bgSelector.querySelectorAll('option');
  options.forEach(opt => opt.selected = opt.value === (theme === 'auto' ? 'auto' : theme));
}

// 监听主题切换
bgSelector.addEventListener("change", e => setTheme(e.target.value));

// 监听系统主题变化
window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', e => {
  if (bgSelector.value === 'auto') {
    setTheme('auto');
  }
});

// 页面加载时恢复主题设置
window.addEventListener("load", () => {
  const saved = localStorage.getItem("toolbox-theme") || 'light';
  setTheme(saved);
  bgSelector.value = saved;
});

// 添加平滑滚动
function smoothScroll() {
  document.documentElement.style.scrollBehavior = 'smooth';
}

// 页面加载完成后启用平滑滚动
window.addEventListener('load', smoothScroll);

// DXF Result Controls might need to be called if it exists (it's in dxf_search.js now)
// We can check if it exists or let dxf_search.js attach itself
// Since we are loading dxf_search.js separately, we should let it attach its own listeners.
// But we need to make sure setupResultsControls is called.
// I'll add a check here or better, add it in dxf_search.js 'load' event.
// In dxf_search.js, I didn't add the listener. I will trust my copy which didn't include the 'load' listener at the bottom.
// Wait, I did not copy the 'window.addEventListener' at the bottom of dxf_search. JS.
// I should add it there. But for now I will fix this in a subsequent edit or assume it's fine.
// Actually, in the code I wrote for dxf_search.js, I did NOT include `window.addEventListener('load', setupResultsControls);`.
// I should add it to app.js to call it if it exists.
window.addEventListener('load', () => {
  if (typeof setupResultsControls === 'function') setupResultsControls();
});


// ================= 工具函数 =================

// 显示提示信息
function showAlert(message, type = 'info') {
  const alertDiv = document.createElement('div');
  alertDiv.className = `alert alert-${type}`;
  alertDiv.innerHTML = `
    <span class="status-indicator status-${type}"></span>
    ${message}
    <button onclick="this.parentElement.remove()" style="margin-left: auto; background: none; border: none; font-size: 18px; cursor: pointer;">×</button>
  `;

  // 插入到主内容区域顶部
  const main = document.querySelector('main');
  main.insertBefore(alertDiv, main.firstChild);

  // 5秒后自动消失
  setTimeout(() => {
    if (alertDiv.parentElement) {
      alertDiv.remove();
    }
  }, 5000);
}

// 防抖函数
function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

// 文件大小格式化
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 检测是否为移动设备
function isMobile() {
  return window.innerWidth <= 768;
}

// 键盘快捷键支持
document.addEventListener('keydown', (e) => {
  // Ctrl/Cmd + Enter 开始扫描
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
    // Check if we are on DXF tab? Or just try to click scanBtn if it exists
    const scanBtn = document.getElementById("scanBtn");
    if (scanBtn && !scanBtn.disabled && scanBtn.offsetParent !== null) { // offsetParent check visibility
      e.preventDefault();
      scanBtn.click();
    }
  }

  // Ctrl/Cmd + S 导出结果
  if ((e.ctrlKey || e.metaKey) && e.key === 's') {
    const exportCSV = document.getElementById("exportCSV");
    if (exportCSV && !exportCSV.disabled && exportCSV.offsetParent !== null) {
      e.preventDefault();
      exportCSV.click();
    }
  }
});

function loadSheetJS() {
  return new Promise(resolve => {
    if (typeof XLSX !== "undefined") { resolve(); return; }
    const script = document.createElement("script");
    script.src = "./xlsx.full.min.js";
    script.onload = resolve;
    document.body.appendChild(script);
  });
}
