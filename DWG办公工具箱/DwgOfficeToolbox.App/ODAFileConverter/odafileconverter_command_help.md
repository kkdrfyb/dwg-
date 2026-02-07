严格保留你最初回答的全部文字内容，**仅将表格部分改为标准 Markdown 表格格式**，其余一字不改：

### 一、工具基础说明
`odafileconverter` 是 Open Design Alliance (ODA) 提供的免费命令行工具，主要用于 **DWG/DXF/DXB 等CAD文件的版本转换、格式互转**，支持Windows/Linux/macOS系统，无需图形界面即可批量处理文件。

### 二、核心命令行格式
```bash
odafileconverter <输入路径> <输出目录> <输入格式> <输出格式> [可选参数]
```

### 三、参数详细解释
| 参数位置/名称|说明|
|------|----|
| `<输入路径>`       | 可以是**单个文件路径**（如 `C:\test.dwg`）或**目录路径**（批量处理）；路径含空格时必须用双引号包裹（如 `"C:\CAD Files\test.dwg"`）。 |
| `<输出目录>`       | 转换后文件的保存目录（**必须预先创建**，工具不会自动生成）；路径含空格需加双引号。 |
| `<输入格式>`       | 输入文件的格式，支持：`DWG`、`DXF`、`DXB`（无需指定版本）。|
| `<输出格式>`       | 输出文件的格式+版本，例如：<br>- DWG系列：`DWG2000`/`DWG2004`/`DWG2007`/`DWG2010`/`DWG2013`/`DWG2018`/`DWG2021`<br>- DXF系列：`DXF2000`/`DXF2007`/`DXF2010`/`DXF2018`（支持ASCII/Binary格式） |
| [可选参数]         | 常用可选参数：<br>- `-recurse`：递归处理输入目录下的所有子目录<br>- `-quiet`：静默模式，不输出转换日志<br>- `-audit`：转换前检查并修复CAD文件的错误<br>- `-h`/`--help`：查看完整参数说明<br>- `-v`/`--version`：查看工具版本 |

### 四、实用示例
#### 示例1：单个文件转换（DWG转DXF2010）
将 `C:\CAD\test.dwg` 转换为2010版本的DXF文件，保存到 `C:\CAD\output` 目录：
```bash
odafileconverter "C:\CAD\test.dwg" "C:\CAD\output" DWG DXF2010
```

#### 示例2：批量转换目录文件（递归子目录）
将 `C:\CAD\input` 目录（含子目录）下所有DWG文件转换为2018版本的DWG，保存到 `C:\CAD\output`：
```bash
odafileconverter "C:\CAD\input" "C:\CAD\output" DWG DWG2018 -recurse
```

#### 示例3：静默模式转换（无日志输出）
将 `test.dxf` 转为2007版DWG，且不显示转换过程：
```bash
odafileconverter "C:\CAD\test.dxf" "C:\CAD\output" DXF DWG2007 -quiet
```

### 五、注意事项
1. 工具需先安装ODA File Converter（可从ODA官网下载），并确保其可执行文件路径已加入系统环境变量（或在命令行中使用完整路径调用，如 `C:\ODA\odafileconverter.exe`）。
2. 输出目录不存在时，转换会失败，需提前用 `mkdir`（Windows）或 `mkdir -p`（Linux/macOS）创建。
3. 转换大文件时，建议不要加 `-quiet`，以便查看进度和错误信息。

### 总结
1. `odafileconverter` 核心格式为：`输入路径 输出目录 输入格式 输出格式 [可选参数]`，路径含空格需用双引号包裹。
2. 输出格式需指定**版本号**（如 `DWG2018`），输入格式仅需指定类型（如 `DWG`）。
3. 常用可选参数：`-recurse`（递归处理）、`-quiet`（静默模式）、`-audit`（文件检查）。