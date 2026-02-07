# 办公自动化工具箱 (Office Toolbox JS)

这是一个办公自动化工具箱项目，包含：

- 纯前端网页工具箱（本仓库根目录）。
- WPF 桌面版 `DWG办公工具箱`（Windows 10+，MSIX 打包）。
- WPF Win7 兼容版 `DWG办公工具箱_win7兼容`（.NET Framework 4.8，EXE 安装包）。

## 🛠️ 技术方案

-   **核心架构**: 原生 HTML5 / CSS3 / JavaScript (ES6+)。
-   **模块化设计**:
    -   `index.html`: UI 骨架与路由。
    -   `app.js`: 核心应用壳（主题、Tab 切换、通用工具）。
    -   `dxf_search.js`: DXF 搜索与提取逻辑。
    -   `replace.js`: DXF 文本替换逻辑。
    -   `excel_merge.js`: Excel 各类合并模式实现。
-   **依赖库**:
    -   `dxf-parser.min.js`: 用于解析 DXF 文件结构。
    -   `xlsx.full.min.js` (SheetJS): 用于强大的 Excel 读取与写入（支持样式）。
-   **样式**: 原生 CSS Variable 实现深色/浅色主题切换，响应式布局。

## ✨ 功能特性

### 1. 📐 CAD (DXF) 工具集
-   **关键字查找**: 批量扫描文件，查找特定文本，导出定位列表。
-   **信息统计**: 智能识别阀门、风口等设备的编号、尺寸、标高信息。
-   **文本替换**: 批量替换图纸中的文字内容，支持正则，支持预览与导出。

### 2. 📊 Excel 工具集
提供 7 种强大的合并/提取模式，**完美保留原文件格式（字体、颜色、边框、合并单元格）**：
1.  **多工作簿 -> 一个工作簿**: 自动重命名 Sheet，生成带跳转链接的索引目录。
2.  **多工作簿 -> 一个工作表**: 纵向堆叠或横向拼接。支持**跳过表头**与**去除表尾**（如制表人签名行）。
3.  **工作簿内部合并**: 将一个文件内的所有 Sheet 汇总到一张表。
4.  **同名表合并**: 自动提取多个文件中名字相同的 Sheet 进行合并。
5.  **指定位置提取**: 提取所有文件的 A1, B2 等特定单元格生成台账。
6.  **同文件名汇总**: 专门处理不同文件夹下的同名文件。
7.  **动态字段合并**: 按表头名称自动对齐数据列。

### 3. 📝 Word 工具集
-   *(开发中) 格式设置与模板填充工具*

## 🚀 使用说明

### 在线/本地运行
1.  下载本项目所有文件到本地目录。
2.  直接使用 Chrome / Edge 浏览器打开 `index.html` 即可使用。
3.  (推荐) 使用 VS Code 的 "Live Server" 插件运行，体验更佳。

### WPF 桌面版（Windows 10+）
- 目录：`DWG办公工具箱`
- MSIX：`DWG办公工具箱\msix\DwgOfficeToolbox.msix`
- 证书：`DWG办公工具箱\msix\cert\DwgOfficeToolbox.pfx`（需导入“受信任的人”）

### WPF Win7 兼容版（.NET Framework 4.8）
- 目录：`DWG办公工具箱_win7兼容`
- 构建脚本：`DWG办公工具箱_win7兼容\build-exe-win7.ps1`
- EXE 安装包：`DWG办公工具箱_win7兼容\dist\setup\DwgOfficeToolbox_Win7_Setup.exe`
- Win7 机器需预装 **.NET Framework 4.8**

### 开发与扩展
-   **添加新工具**: 在 `index.html` 添加 Tab 按钮及 Content 区域，新建 `.js` 文件并在底部引入。
-   **测试**: 运行 `tests.html` 可进行 Excel 核心逻辑的集成测试（无需上传文件，使用内存模拟）。

## 📅 后续开发计划
1.  **Word 工具开发**: 实现文档内容的批量替换与格式统一。
2.  **性能优化**: 针对超大 Excel 文件（>50MB）引入 Web Worker 处理。
3.  **UI 升级**: 增加更多可视化图表展示统计结果。
