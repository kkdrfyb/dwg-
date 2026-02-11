

WPF 桌面端 DWG 批量文字查找（ODA CLI + DXF 解析）更新版计划
摘要

输出目录改为：输入目录下新建 output 文件夹，并保留转换结果。
DXF 解析规则新增：块内文字解析、属性文字解析。
文件解析阈值调整为 50MB，超过则走纯文本回退。
仍为 WPF .NET 8 桌面端，无前端网页。
1. 范围与目标
支持 DWG 文件与文件夹混选，文件夹递归。
使用 ODA CLI 转 DXF2018，静默模式、递归、转换前审计修复。
按关键词扫描 DXF 文本，结果表格展示并支持筛选、导出 CSV/XLSX。
默认输出目录为输入目录下的 output 文件夹。
2. 关键流程与数据流
用户选择文件与文件夹。
归并输入清单，去重，排除已被文件夹覆盖的单文件。
为每个输入根生成输出目录 <输入目录>\output。
调用 ODA CLI 转换生成 DXF2018。
扫描 DXF 并生成匹配结果。
UI 展示与导出。
转换结果默认保留。
3. ODA CLI 调用规范
可执行文件默认路径：ODAFileConverter.exe。
输出目录规则：
输入为文件夹：输出为 该文件夹\output。
输入为单文件：输出为 该文件所在目录\output。
命令模板：
文件夹：
ODAFileConverter.exe "<inputFolder>" "<outputFolder>" DWG DXF2018 -quiet -audit -recurse
单文件：
ODAFileConverter.exe "<inputFile>" "<outputFolder>" DWG DXF2018 -quiet -audit
输出目录若不存在则创建。
默认允许覆盖同名 DXF，界面提示“可能覆盖旧结果”。
4. DXF 文本解析规则（更新）
解析目标实体

TEXT、MTEXT、ATTRIB、INSERT。
增加解析：块内文字、属性文字。
块内文字解析

解析 BLOCKS 中定义的文字实体。
对每个 INSERT，取其对应 BLOCK 的 TEXT/MTEXT 内容，作为该插入实例的可搜索文本。
属性文字解析

对每个 INSERT，读取其 ATTRIB 值作为可搜索文本。
若仅存在 ATTDEF 而无 ATTRIB，使用 ATTDEF 的默认值作为文本（避免完全丢失属性内容）。
INSERT 兼容规则

保留“块名作为文本内容”的旧逻辑以兼容原 JS 行为。
文本规范化

将 \P 替换为换行或空格。
搜索时忽略大小写。
解析库选择

采用 netDxf（MIT）加载 DXF，便于可靠解析 BLOCK、INSERT、ATTRIB。
解析失败时回退到“纯文本扫描”。
5. 文件解析阈值
LargeFileThresholdMB = 50。
大于 50MB 的 DXF 不进入结构化解析，直接做纯文本扫描。
6. 性能与并发
ODA 转换阶段顺序执行，避免外部进程并发过多。
DXF 解析与扫描阶段并发处理，默认 min(逻辑核心数, 6)。
UI 刷新节流，避免 1000 文件导致界面卡顿。
7. UI 设计（WPF 单窗口）
输入区：选择文件、选择文件夹两个按钮，混合列表展示。
关键词输入框（逗号分隔）。
控制区：开始、取消、导出 CSV、导出 Excel。
进度区：当前文件、累计进度、日志列表。
结果区：DataGrid + 筛选控件（文件名/对象类型/图层/关键字/内容包含）。
8. 导出格式
CSV：UTF-8 BOM。
Excel：ClosedXML 生成 .xlsx。
9. 配置与接口（对外变化）
配置文件：settings.json

OdaExePath
OutputFolderName 默认 output
LargeFileThresholdMB 默认 50
MaxParseConcurrency 默认 6
KeepConvertedDxf 默认 true
ClearOutputBeforeConvert 默认 false
对外行为变化

转换输出默认进入输入目录下 output 文件夹。
解析增加块内文字与属性文字。
10. 测试用例与场景
解析 TEXT/MTEXT/ATTRIB 正确输出。
INSERT 能读取块内文字。
INSERT 能读取 ATTRIB 属性值。
ATTDEF 无 ATTRIB 时能回退使用默认值。
解析失败与大文件时纯文本回退。
输出目录创建与覆盖策略验证。
CSV/XLSX 导出内容字段正确。
11. 明确假设与默认值
ODA CLI 可随应用分发且许可允许安装包打包。
DXF 默认使用 ASCII。
默认保留输出目录中的 DXF 文件。
不包含阀门/风口统计与文字替换功能。
