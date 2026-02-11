import 'dart:async';
import 'dart:convert';
import 'dart:io';
import 'dart:isolate';
import 'dart:math';
import 'package:csv/csv.dart';
import 'package:excel/excel.dart';

import '../../core/task/task_exceptions.dart';
import 'excel_models.dart';

class ExcelProgress {
  ExcelProgress({required this.completed, required this.total, required this.message});

  final int completed;
  final int total;
  final String message;
}

class ExcelIsolateRunner {
  Isolate? _isolate;
  SendPort? _sendPort;
  StreamSubscription? _subscription;
  final Completer<ExcelJobResult> _resultCompleter = Completer<ExcelJobResult>();

  Future<ExcelJobResult> run(
    ExcelJob job, {
    required void Function(ExcelProgress progress) onProgress,
  }) async {
    final receivePort = ReceivePort();
    _isolate = await Isolate.spawn(_excelIsolateEntry, receivePort.sendPort);

    _subscription = receivePort.listen((message) {
      if (message is Map) {
        final type = message['type'];
        if (type == 'ready') {
          _sendPort = message['sendPort'] as SendPort;
          _sendPort?.send({'type': 'start', 'job': job.toMap()});
        } else if (type == 'progress') {
          onProgress(
            ExcelProgress(
              completed: message['completed'] as int? ?? 0,
              total: message['total'] as int? ?? 0,
              message: message['message'] as String? ?? '',
            ),
          );
        } else if (type == 'result') {
          if (!_resultCompleter.isCompleted) {
            final result = ExcelJobResult.fromMap(Map<String, dynamic>.from(message['result'] as Map));
            _resultCompleter.complete(result);
          }
          _dispose();
        } else if (type == 'error') {
          if (!_resultCompleter.isCompleted) {
            _resultCompleter.completeError(Exception(message['message']));
          }
          _dispose();
        }
      }
    });

    return _resultCompleter.future;
  }

  void cancel() {
    _sendPort?.send({'type': 'cancel'});
  }

  void _dispose() {
    _subscription?.cancel();
    _subscription = null;
    _isolate?.kill(priority: Isolate.immediate);
    _isolate = null;
  }
}

void _excelIsolateEntry(SendPort mainSendPort) {
  final port = ReceivePort();
  mainSendPort.send({'type': 'ready', 'sendPort': port.sendPort});
  var canceled = false;

  port.listen((message) {
    if (message is Map) {
      final type = message['type'];
      if (type == 'cancel') {
        canceled = true;
        return;
      }
      if (type == 'start') {
        final jobMap = Map<String, dynamic>.from(message['job'] as Map);
        try {
          final job = ExcelJob.fromMap(jobMap);
          final result = _processJob(
            job,
            onProgress: (completed, total, msg) {
              mainSendPort.send({
                'type': 'progress',
                'completed': completed,
                'total': total,
                'message': msg,
              });
            },
            isCanceled: () => canceled,
          );
          mainSendPort.send({'type': 'result', 'result': result.toMap()});
        } on TaskCanceled {
          final result = ExcelJobResult(outputs: [], sheetNames: [], canceled: true);
          mainSendPort.send({'type': 'result', 'result': result.toMap()});
        } catch (error, stack) {
          mainSendPort.send({'type': 'error', 'message': '$error\n$stack'});
        }
      }
    }
  });
}

ExcelJobResult _processJob(
  ExcelJob job, {
  required void Function(int completed, int total, String message) onProgress,
  required bool Function() isCanceled,
}) {
  final outputs = <ExcelOutput>[];
  ExcelPreview? preview;
  final sheetNames = <String>[];

  void checkCanceled() {
    if (isCanceled()) throw TaskCanceled();
  }

  void addOutput(Excel excel, String filename) {
    if (job.preview) return;
    outputs.add(ExcelOutput(filename: filename, bytes: _encodeExcel(excel)));
  }

  Excel readExcel(ExcelInputFile input) {
    final ext = input.name.toLowerCase();
    if (ext.endsWith('.csv')) {
      return _readCsv(File(input.path));
    }
    if (ext.endsWith('.xlsx')) {
      final bytes = File(input.path).readAsBytesSync();
      return Excel.decodeBytes(bytes);
    }
    throw UnsupportedError('不支持的文件格式: ${input.name}');
  }

  Excel createWorkbookWithSheet(String sheetName) {
    final excel = Excel.createExcel();
    final defaultName = excel.getDefaultSheet() ?? 'Sheet1';
    if (defaultName != sheetName) {
      excel.rename(defaultName, sheetName);
    }
    return excel;
  }

  Excel? lastWorkbook;

  switch (job.mode) {
    case ExcelMode.mergeWorkbooks:
      {
        final excel = createWorkbookWithSheet('索引');
        final indexSheet = excel['索引'];
        final usedNames = <String>{'索引'};
        final indexRows = <List<String>>[
          ['工作表-列表', '源文件与工作表'],
        ];
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          for (final sheetName in wb.sheets.keys) {
            final sourceSheet = wb.sheets[sheetName]!;
            final cleanName = _cleanFileName(file.name);
            var newName = _uniqueSheetName('$cleanName-$sheetName', usedNames);
            usedNames.add(newName);
            final targetSheet = excel[newName];
            _copySheet(sourceSheet, targetSheet, checkCanceled: checkCanceled);
            indexRows.add([newName, '${file.name} - $sheetName']);
          }
        }
        _writeRows(indexSheet, indexRows, linkStyle: _linkStyle());
        lastWorkbook = excel;
        addOutput(excel, '合并结果_多工作簿.xlsx');
        break;
      }
    case ExcelMode.mergeToSheet:
      {
        if (job.direction == ExcelDirection.horizontal) {
          final excel = createWorkbookWithSheet('结果');
          final sheet = excel['结果'];
          sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value =
              TextCellValue('横向合并暂不支持复杂格式');
          lastWorkbook = excel;
          addOutput(excel, '横向合并结果.xlsx');
          break;
        }

        final excel = createWorkbookWithSheet('汇总结果');
        final target = excel['汇总结果'];
        var currentRow = 0;
        var isFirst = true;
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          for (final sheetName in wb.sheets.keys) {
            checkCanceled();
            final sheet = wb.sheets[sheetName]!;
            currentRow = _appendSheetData(
              target,
              sheet,
              startRow: currentRow,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
              includeHeader: isFirst,
              sourceName: '${file.name}-$sheetName',
              copyColumns: isFirst,
              checkCanceled: checkCanceled,
            );
            isFirst = false;
          }
        }
        lastWorkbook = excel;
        addOutput(excel, '多簿汇总.xlsx');
        break;
      }
    case ExcelMode.internalMerge:
      {
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '处理: ${file.name}');
          final wb = readExcel(file);
          final out = createWorkbookWithSheet('全工作簿汇总');
          final summary = out['全工作簿汇总'];
          var currRow = 0;
          var sheetIndex = 0;
          for (final sheetName in wb.sheets.keys) {
            final sheet = wb.sheets[sheetName]!;
            currRow = _appendSheetData(
              summary,
              sheet,
              startRow: currRow,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
              includeHeader: sheetIndex == 0,
              sourceName: sheetName,
              copyColumns: sheetIndex == 0,
              checkCanceled: checkCanceled,
            );
            sheetIndex++;
          }

          final usedNames = <String>{'全工作簿汇总'};
          var sheetCursor = 0;
          for (final sheetName in wb.sheets.keys) {
            if (job.internalMode == InternalMergeMode.firstSheet && sheetCursor == 0) {
              sheetCursor++;
              continue;
            }
            final uniqueName = _uniqueSheetName(sheetName, usedNames);
            usedNames.add(uniqueName);
            final targetSheet = out[uniqueName];
            _copySheet(wb.sheets[sheetName]!, targetSheet, checkCanceled: checkCanceled);
            sheetCursor++;
          }

          final filename = '汇总_${_cleanFileName(file.name)}.xlsx';
          addOutput(out, filename);
          lastWorkbook = out;
        }
        break;
      }
    case ExcelMode.sameNameSheet:
      {
        final excel = Excel.createExcel();
        final tempSheet = excel.getDefaultSheet() ?? 'Sheet1';
        final used = <String>{};
        final sheetMap = <String, List<_SheetRef>>{};
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          for (final name in wb.sheets.keys) {
            sheetMap.putIfAbsent(name, () => []).add(_SheetRef(file.name, wb.sheets[name]!));
          }
        }
        var added = false;
        for (final entry in sheetMap.entries) {
          checkCanceled();
          final sheetName = _uniqueSheetName(entry.key, used);
          used.add(sheetName);
          final target = excel[sheetName];
          var currRow = 0;
          for (var idx = 0; idx < entry.value.length; idx++) {
            final item = entry.value[idx];
            currRow = _appendSheetData(
              target,
              item.sheet,
              startRow: currRow,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
              includeHeader: idx == 0,
              sourceName: item.fileName,
              copyColumns: idx == 0,
              checkCanceled: checkCanceled,
            );
          }
          added = true;
        }
        if (added && tempSheet.isNotEmpty && excel.sheets.containsKey(tempSheet)) {
          excel.delete(tempSheet);
        }
        lastWorkbook = excel;
        addOutput(excel, '同名表汇总.xlsx');
        break;
      }
    case ExcelMode.samePosition:
      {
        final cells = _parseCellList(job.cellRange);
        if (cells.isEmpty) {
          final excel = createWorkbookWithSheet('全部合并');
          final target = excel['全部合并'];
          var currentRow = 0;
          var isFirst = true;
          for (var i = 0; i < job.files.length; i++) {
            checkCanceled();
            final file = job.files[i];
            onProgress(i + 1, job.files.length, '读取: ${file.name}');
            final wb = readExcel(file);
            final sheet = wb.sheets.values.first;
            currentRow = _appendSheetData(
              target,
              sheet,
              startRow: currentRow,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
              includeHeader: isFirst,
              sourceName: file.name,
              copyColumns: isFirst,
              checkCanceled: checkCanceled,
            );
            isFirst = false;
          }
          lastWorkbook = excel;
          addOutput(excel, '提取结果.xlsx');
        } else {
          final excel = createWorkbookWithSheet('提取结果');
          final sheet = excel['提取结果'];
          final header = ['文件名', ...cells];
          _writeRow(sheet, 0, header);
          for (var i = 0; i < job.files.length; i++) {
            checkCanceled();
            final file = job.files[i];
            onProgress(i + 1, job.files.length, '读取: ${file.name}');
            final wb = readExcel(file);
            final firstSheet = wb.sheets.values.first;
            final row = <String>[file.name];
            for (final cell in cells) {
              final idx = CellIndex.indexByString(cell);
              final data = firstSheet.cell(idx);
              row.add(data.value?.toString() ?? '');
            }
            _writeRow(sheet, i + 1, row);
          }
          lastWorkbook = excel;
          addOutput(excel, '提取结果.xlsx');
        }
        break;
      }
    case ExcelMode.sameFilename:
      {
        final excel = createWorkbookWithSheet('索引');
        final indexSheet = excel['索引'];
        final usedNames = <String>{'索引'};
        final indexRows = <List<String>>[
          ['列表', '来源'],
        ];
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          for (final sheetName in wb.sheets.keys) {
            var finalName = sheetName;
            if (usedNames.contains(finalName)) {
              if (job.sameNameMode == SameNameMode.skip) {
                continue;
              }
              finalName = _uniqueSheetName('${_cleanFileName(file.name)}-$sheetName', usedNames);
            }
            usedNames.add(finalName);
            final targetSheet = excel[finalName];
            _copySheet(wb.sheets[sheetName]!, targetSheet, checkCanceled: checkCanceled);
            indexRows.add([finalName, file.name]);
          }
        }
        _writeRows(indexSheet, indexRows, linkStyle: _linkStyle());
        lastWorkbook = excel;
        addOutput(excel, '同名文件汇总.xlsx');
        break;
      }
    case ExcelMode.mergeDynamic:
      {
        final excel = createWorkbookWithSheet('动态合并');
        final sheet = excel['动态合并'];
        final headerOrder = <String>[];
        final headerSet = <String>{};
        final headerIndex = <String, int>{};
        var currentRow = 0;
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          final firstSheet = wb.sheets.values.first;
          if (firstSheet.maxRows == 0) continue;
          final headerRow = firstSheet.row(0);
          final headers = <String>[];
          for (var c = 0; c < firstSheet.maxColumns; c++) {
            final value = headerRow.isNotEmpty && c < headerRow.length ? headerRow[c]?.value : null;
            final text = value?.toString().trim() ?? '';
            headers.add(text);
            if (text.isNotEmpty && !headerSet.contains(text)) {
              headerSet.add(text);
              headerOrder.add(text);
              headerIndex[text] = headerOrder.length - 1;
            }
          }
          if (currentRow == 0) {
            _writeRow(sheet, 0, headerOrder);
            currentRow = 1;
          }
          for (var r = 1; r < firstSheet.maxRows; r++) {
            checkCanceled();
            final row = firstSheet.row(r);
            if (row.every((cell) => cell == null)) continue;
            final values = List<String>.filled(headerOrder.length, '');
            for (var c = 0; c < headers.length; c++) {
              final header = headers[c];
              if (header.isEmpty) continue;
              final index = headerIndex[header];
              if (index == null) continue;
              final cell = c < row.length ? row[c] : null;
              values[index] = cell?.value?.toString() ?? '';
            }
            _writeRow(sheet, currentRow, values);
            currentRow++;
          }
        }
        lastWorkbook = excel;
        addOutput(excel, '动态字段合并.xlsx');
        break;
      }
  }

  if (lastWorkbook != null) {
    sheetNames.addAll(lastWorkbook.sheets.keys);
    if (job.preview) {
      preview = _buildPreview(lastWorkbook);
      outputs.clear();
    }
  }

  return ExcelJobResult(outputs: outputs, sheetNames: sheetNames, preview: preview);
}

ExcelPreview _buildPreview(Excel workbook) {
  final firstSheet = workbook.sheets.values.first;
  final rows = <List<String>>[];
  final maxRows = min(50, firstSheet.maxRows);
  for (var r = 0; r < maxRows; r++) {
    final row = firstSheet.row(r);
    if (row.isEmpty) {
      rows.add([]);
      continue;
    }
    final values = <String>[];
    for (final cell in row) {
      values.add(cell?.value?.toString() ?? '');
    }
    rows.add(values);
  }
  return ExcelPreview(sheetName: firstSheet.sheetName, rows: rows);
}

Excel _readCsv(File file) {
  final bytes = file.readAsBytesSync();
  final content = _decodeText(bytes);
  final rows = const CsvDecoder().convert(content);
  final excel = Excel.createExcel();
  final sheet = excel[excel.getDefaultSheet() ?? 'Sheet1'];
  for (var r = 0; r < rows.length; r++) {
    final row = rows[r];
    for (var c = 0; c < row.length; c++) {
      final value = row[c];
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r)).value =
          TextCellValue(value?.toString() ?? '');
    }
  }
  return excel;
}

String _decodeText(List<int> bytes) {
  try {
    return utf8.decode(bytes);
  } catch (_) {
    return latin1.decode(bytes);
  }
}

int _appendSheetData(
  Sheet target,
  Sheet source, {
  required int startRow,
  required int headerRows,
  required int footerRows,
  required bool includeHeader,
  required bool copyColumns,
  required void Function() checkCanceled,
  String? sourceName,
}) {
  if (source.maxRows == 0 || source.maxColumns == 0) return startRow;
  final sourceStart = includeHeader ? 0 : headerRows;
  var sourceEnd = source.maxRows - 1 - footerRows;
  if (sourceEnd < sourceStart) return startRow;
  final sourceMaxCol = max(0, source.maxColumns - 1);

  if (copyColumns) {
    source.getColumnWidths.forEach((col, width) {
      target.setColumnWidth(col, width);
    });
  }

  for (var r = sourceStart; r <= sourceEnd; r++) {
    checkCanceled();
    final row = source.row(r);
    final newRow = startRow + (r - sourceStart);

    if (source.getRowHeights.containsKey(r)) {
      target.setRowHeight(newRow, source.getRowHeights[r]!);
    }

    for (var c = 0; c <= sourceMaxCol; c++) {
      final data = c < row.length ? row[c] : null;
      if (data == null) continue;
      final cell = target.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: newRow));
      cell.value = data.value;
      if (data.cellStyle != null) {
        cell.cellStyle = data.cellStyle;
      }
    }

    if (sourceName != null) {
      final linkCell = target.cell(CellIndex.indexByColumnRow(columnIndex: sourceMaxCol + 1, rowIndex: newRow));
      linkCell.value = TextCellValue(sourceName);
      linkCell.cellStyle = _linkStyle();
    }
  }

  if (source.spannedItems.isNotEmpty) {
    for (final span in source.spannedItems) {
      final range = _parseSpan(span);
      if (range == null) continue;
      if (range.start.rowIndex < sourceStart || range.end.rowIndex > sourceEnd) continue;
      final targetStart = CellIndex.indexByColumnRow(
        columnIndex: range.start.columnIndex,
        rowIndex: startRow + (range.start.rowIndex - sourceStart),
      );
      final targetEnd = CellIndex.indexByColumnRow(
        columnIndex: range.end.columnIndex,
        rowIndex: startRow + (range.end.rowIndex - sourceStart),
      );
      target.merge(targetStart, targetEnd);
    }
  }

  return startRow + (sourceEnd - sourceStart + 1);
}

void _copySheet(
  Sheet source,
  Sheet target, {
  required void Function() checkCanceled,
}) {
  if (source.maxRows == 0 || source.maxColumns == 0) return;

  source.getColumnWidths.forEach((col, width) {
    target.setColumnWidth(col, width);
  });
  source.getRowHeights.forEach((row, height) {
    target.setRowHeight(row, height);
  });

  for (var r = 0; r < source.maxRows; r++) {
    checkCanceled();
    final row = source.row(r);
    for (var c = 0; c < source.maxColumns; c++) {
      final data = c < row.length ? row[c] : null;
      if (data == null) continue;
      final cell = target.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r));
      cell.value = data.value;
      if (data.cellStyle != null) {
        cell.cellStyle = data.cellStyle;
      }
    }
  }

  if (source.spannedItems.isNotEmpty) {
    for (final span in source.spannedItems) {
      final range = _parseSpan(span);
      if (range == null) continue;
      target.merge(range.start, range.end);
    }
  }
}

void _writeRows(
  Sheet sheet,
  List<List<String>> rows, {
  CellStyle? linkStyle,
}) {
  for (var r = 0; r < rows.length; r++) {
    _writeRow(sheet, r, rows[r]);
    if (r > 0 && linkStyle != null) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r));
      cell.cellStyle = linkStyle;
      final second = sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r));
      second.cellStyle = linkStyle;
    }
  }
}

void _writeRow(Sheet sheet, int rowIndex, List<String> values) {
  for (var c = 0; c < values.length; c++) {
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: rowIndex)).value =
        TextCellValue(values[c]);
  }
}

List<int> _encodeExcel(Excel excel) {
  final bytes = excel.encode();
  if (bytes == null) return <int>[];
  return bytes;
}

CellStyle _linkStyle() {
  return CellStyle(
    fontColorHex: '0000FF'.excelColor,
    underline: Underline.Single,
  );
}

String _cleanFileName(String name) {
  return name.replaceAll(RegExp(r'\.[^/.]+$'), '');
}

String _uniqueSheetName(String base, Set<String> used) {
  var sanitized = _sanitizeSheetName(base);
  var name = sanitized;
  var counter = 1;
  while (used.contains(name)) {
    name = _sanitizeSheetName('${sanitized}_copy$counter');
    counter++;
  }
  return name;
}

String _sanitizeSheetName(String name) {
  var sanitized = name.replaceAll(RegExp(r'[\\/\?\*\[\]:]'), '_');
  if (sanitized.length > 31) {
    sanitized = sanitized.substring(0, 31);
  }
  if (sanitized.isEmpty) {
    return 'Sheet';
  }
  return sanitized;
}

List<String> _parseCellList(String input) {
  if (input.trim().isEmpty) return [];
  return input
      .split(RegExp(r'[,，]'))
      .map((c) => c.trim().toUpperCase())
      .where((c) => c.isNotEmpty)
      .toList();
}

_MergeRange? _parseSpan(String span) {
  final parts = span.split(':');
  try {
    final start = CellIndex.indexByString(parts[0]);
    final end = parts.length > 1 ? CellIndex.indexByString(parts[1]) : start;
    return _MergeRange(start, end);
  } catch (_) {
    return null;
  }
}

class _MergeRange {
  _MergeRange(this.start, this.end);

  final CellIndex start;
  final CellIndex end;
}

class _SheetRef {
  _SheetRef(this.fileName, this.sheet);

  final String fileName;
  final Sheet sheet;
}
