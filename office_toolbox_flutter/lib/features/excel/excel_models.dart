enum ExcelMode {
  mergeWorkbooks,
  mergeToSheet,
  internalMerge,
  sameNameSheet,
  samePosition,
  sameFilename,
  mergeDynamic,
}

enum ExcelDirection { vertical, horizontal }

enum InternalMergeMode { newSheetFirst, firstSheet }

enum SameNameMode { rename, skip }

class ExcelInputFile {
  ExcelInputFile({required this.name, required this.path, required this.size});

  final String name;
  final String path;
  final int size;

  Map<String, dynamic> toMap() => {
        'name': name,
        'path': path,
        'size': size,
      };

  factory ExcelInputFile.fromMap(Map<String, dynamic> map) {
    return ExcelInputFile(
      name: map['name'] as String,
      path: map['path'] as String,
      size: map['size'] as int? ?? 0,
    );
  }
}

class ExcelJob {
  ExcelJob({
    required this.mode,
    required this.files,
    this.headerRows = 1,
    this.footerRows = 0,
    this.direction = ExcelDirection.vertical,
    this.internalMode = InternalMergeMode.newSheetFirst,
    this.sameNameMode = SameNameMode.rename,
    this.cellRange = '',
    this.preview = false,
  });

  final ExcelMode mode;
  final List<ExcelInputFile> files;
  final int headerRows;
  final int footerRows;
  final ExcelDirection direction;
  final InternalMergeMode internalMode;
  final SameNameMode sameNameMode;
  final String cellRange;
  final bool preview;

  Map<String, dynamic> toMap() => {
        'mode': mode.name,
        'files': files.map((file) => file.toMap()).toList(),
        'headerRows': headerRows,
        'footerRows': footerRows,
        'direction': direction.name,
        'internalMode': internalMode.name,
        'sameNameMode': sameNameMode.name,
        'cellRange': cellRange,
        'preview': preview,
      };

  factory ExcelJob.fromMap(Map<String, dynamic> map) {
    return ExcelJob(
      mode: ExcelMode.values.byName(map['mode'] as String),
      files: (map['files'] as List)
          .map((item) => ExcelInputFile.fromMap(Map<String, dynamic>.from(item as Map)))
          .toList(),
      headerRows: map['headerRows'] as int? ?? 1,
      footerRows: map['footerRows'] as int? ?? 0,
      direction: ExcelDirection.values.byName(map['direction'] as String),
      internalMode: InternalMergeMode.values.byName(map['internalMode'] as String),
      sameNameMode: SameNameMode.values.byName(map['sameNameMode'] as String),
      cellRange: map['cellRange'] as String? ?? '',
      preview: map['preview'] as bool? ?? false,
    );
  }
}

class ExcelOutput {
  ExcelOutput({required this.filename, required this.bytes});

  final String filename;
  final List<int> bytes;

  Map<String, dynamic> toMap() => {
        'filename': filename,
        'bytes': bytes,
      };

  factory ExcelOutput.fromMap(Map<String, dynamic> map) {
    return ExcelOutput(
      filename: map['filename'] as String,
      bytes: (map['bytes'] as List).cast<int>(),
    );
  }
}

class ExcelPreview {
  ExcelPreview({required this.sheetName, required this.rows});

  final String sheetName;
  final List<List<String>> rows;

  Map<String, dynamic> toMap() => {
        'sheetName': sheetName,
        'rows': rows,
      };

  factory ExcelPreview.fromMap(Map<String, dynamic> map) {
    return ExcelPreview(
      sheetName: map['sheetName'] as String,
      rows: (map['rows'] as List)
          .map((row) => (row as List).map((cell) => cell.toString()).toList())
          .toList(),
    );
  }
}

class ExcelJobResult {
  ExcelJobResult({
    required this.outputs,
    required this.sheetNames,
    this.preview,
    this.canceled = false,
  });

  final List<ExcelOutput> outputs;
  final List<String> sheetNames;
  final ExcelPreview? preview;
  final bool canceled;

  Map<String, dynamic> toMap() => {
        'outputs': outputs.map((o) => o.toMap()).toList(),
        'sheetNames': sheetNames,
        'preview': preview?.toMap(),
        'canceled': canceled,
      };

  factory ExcelJobResult.fromMap(Map<String, dynamic> map) {
    return ExcelJobResult(
      outputs: (map['outputs'] as List)
          .map((item) => ExcelOutput.fromMap(Map<String, dynamic>.from(item as Map)))
          .toList(),
      sheetNames: (map['sheetNames'] as List?)?.map((e) => e.toString()).toList() ?? [],
      preview: map['preview'] == null
          ? null
          : ExcelPreview.fromMap(Map<String, dynamic>.from(map['preview'] as Map)),
      canceled: map['canceled'] as bool? ?? false,
    );
  }
}

extension ExcelModeLabel on ExcelMode {
  String get label {
    switch (this) {
      case ExcelMode.mergeWorkbooks:
        return '多工作簿 -> 工作簿';
      case ExcelMode.mergeToSheet:
        return '多工作簿 -> 工作表';
      case ExcelMode.internalMerge:
        return '工作簿内部汇总';
      case ExcelMode.sameNameSheet:
        return '同名 Sheet 汇总';
      case ExcelMode.samePosition:
        return '指定位置/单元格提取';
      case ExcelMode.sameFilename:
        return '同名文件汇总';
      case ExcelMode.mergeDynamic:
        return '合并动态字段';
    }
  }
}
