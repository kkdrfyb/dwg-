import 'dart:async';
import 'dart:collection';
import 'dart:convert';
import 'dart:io';
import 'dart:isolate';
import 'dart:math';
import 'package:csv/csv.dart';
import 'package:excel/excel.dart';

import '../../core/task/task_exceptions.dart';
import 'excel_models.dart';

class ExcelProgress {
  ExcelProgress({
    required this.completed,
    required this.total,
    required this.message,
  });

  final int completed;
  final int total;
  final String message;
}

class ExcelIsolateRunner {
  Isolate? _isolate;
  SendPort? _sendPort;
  StreamSubscription? _subscription;
  final Completer<ExcelJobResult> _resultCompleter =
      Completer<ExcelJobResult>();

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
            final result = ExcelJobResult.fromMap(
              Map<String, dynamic>.from(message['result'] as Map),
            );
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
          final result = ExcelJobResult(
            outputs: [],
            sheetNames: [],
            canceled: true,
          );
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
          sheet
              .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0))
              .value = TextCellValue(
            '横向合并暂不支持复杂格式',
          );
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
    case ExcelMode.mergeToSheetSummary:
      {
        final allRows = <Map<String, String>>[];
        final headers = <String>{};
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '汇总: ${file.name}');
          final wb = readExcel(file);
          for (final sheetName in wb.sheets.keys) {
            final source = wb.sheets[sheetName]!;
            final parsed = _sheetToRecords(
              source,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
            );
            if (parsed == null) {
              continue;
            }
            headers.addAll(parsed.headers);
            allRows.addAll(parsed.rows);
          }
        }

        final aggregated = _aggregateRecords(
          allRows,
          explicitHeaderOrder: headers.toList(),
        );
        final excel = createWorkbookWithSheet('Result');
        final sheet = excel['Result'];
        _writeRows(sheet, aggregated);
        lastWorkbook = excel;
        addOutput(excel, '汇总结果（多簿汇总到一簿）.xlsx');
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
            if (job.internalMode == InternalMergeMode.firstSheet &&
                sheetCursor == 0) {
              sheetCursor++;
              continue;
            }
            final uniqueName = _uniqueSheetName(sheetName, usedNames);
            usedNames.add(uniqueName);
            final targetSheet = out[uniqueName];
            _copySheet(
              wb.sheets[sheetName]!,
              targetSheet,
              checkCanceled: checkCanceled,
            );
            sheetCursor++;
          }

          final filename = '汇总_${_cleanFileName(file.name)}.xlsx';
          addOutput(out, filename);
          lastWorkbook = out;
        }
        break;
      }
    case ExcelMode.reorderColumns:
      {
        final preferredOrder = _parseFieldOrder(job.fieldOrder);
        final aliasMap = _parseAliasRules(job.aliasRules);
        final usedFiles = <String>{};
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '调整列序: ${file.name}');
          final source = readExcel(file);
          final out = _reorderWorkbookColumns(
            source,
            preferredOrder: preferredOrder,
            aliasMap: aliasMap,
            checkCanceled: checkCanceled,
          );
          final filename = _uniqueFileName(
            _sanitizeFileName('${_cleanFileName(file.name)}.xlsx'),
            usedFiles,
          );
          addOutput(out, filename);
          lastWorkbook = out;
        }
        break;
      }
    case ExcelMode.splitWorkbook:
      {
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          final wb = readExcel(file);
          onProgress(
            i + 1,
            job.files.length,
            '拆分: ${file.name} (${wb.sheets.length} 个工作表)',
          );
          for (final sheetName in wb.sheets.keys) {
            checkCanceled();
            final outSheetName = _sanitizeSheetName(sheetName);
            final out = createWorkbookWithSheet(outSheetName);
            final target = out[outSheetName];
            _copySheet(
              wb.sheets[sheetName]!,
              target,
              checkCanceled: checkCanceled,
            );
            final filename = _sanitizeFileName(
              '${_cleanFileName(file.name)}-$sheetName.xlsx',
            );
            addOutput(out, filename);
            lastWorkbook = out;
          }
        }
        break;
      }
    case ExcelMode.splitWorksheet:
      {
        final allRows = <Map<String, String>>[];
        final headerOrder = <String>[];
        final headerSeen = <String>{};
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          for (final source in wb.sheets.values) {
            final parsed = _sheetToRecords(
              source,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
            );
            if (parsed == null) {
              continue;
            }
            for (final header in parsed.headers) {
              if (headerSeen.add(header)) {
                headerOrder.add(header);
              }
            }
            allRows.addAll(parsed.rows);
          }
        }
        if (allRows.isEmpty) {
          final empty = createWorkbookWithSheet('Result');
          empty['Result']
              .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0))
              .value = TextCellValue(
            '未找到可拆分数据',
          );
          addOutput(empty, '拆分结果.xlsx');
          lastWorkbook = empty;
          break;
        }

        final splitHeader = _resolveSplitHeader(headerOrder, job.splitKey);
        final grouped = <String, List<Map<String, String>>>{};
        for (final row in allRows) {
          final key = (row[splitHeader] ?? '').trim();
          final groupKey = key.isEmpty ? '空值' : key;
          grouped.putIfAbsent(groupKey, () => <Map<String, String>>[]).add(row);
        }

        if (job.splitWorksheetOutputMode ==
                SplitWorksheetOutputMode.oneWorkbook ||
            job.splitWorksheetOutputMode == SplitWorksheetOutputMode.both) {
          final out = createWorkbookWithSheet('Result(总表)');
          _writeRows(out['Result(总表)'], _recordsToTable(allRows, headerOrder));
          final usedNames = <String>{'Result(总表)'};
          for (final entry in grouped.entries) {
            final sheetName = _uniqueSheetName(entry.key, usedNames);
            usedNames.add(sheetName);
            _writeRows(
              out[sheetName],
              _recordsToTable(entry.value, headerOrder),
            );
          }
          addOutput(out, '拆分结果.xlsx');
          lastWorkbook = out;
        }

        if (job.splitWorksheetOutputMode ==
                SplitWorksheetOutputMode.separateFiles ||
            job.splitWorksheetOutputMode == SplitWorksheetOutputMode.both) {
          final usedFiles = <String>{};
          for (final entry in grouped.entries) {
            checkCanceled();
            final sheetName = _sanitizeSheetName(entry.key);
            final out = createWorkbookWithSheet(sheetName);
            _writeRows(
              out[sheetName],
              _recordsToTable(entry.value, headerOrder),
            );
            final base = _sanitizeFileName('${entry.key}.xlsx');
            final filename = _uniqueFileName(base, usedFiles);
            addOutput(out, filename);
            lastWorkbook = out;
          }
        }
        break;
      }
    case ExcelMode.regroupSameSheetToWorkbook:
      {
        final sheetMap = <String, List<_SheetRef>>{};
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          for (final sheetName in wb.sheets.keys) {
            sheetMap
                .putIfAbsent(sheetName, () => <_SheetRef>[])
                .add(_SheetRef(file.name, wb.sheets[sheetName]!));
          }
        }

        final usedFiles = <String>{};
        var done = 0;
        for (final entry in sheetMap.entries) {
          checkCanceled();
          done++;
          onProgress(done, max(1, sheetMap.length), '重组: ${entry.key}');
          final excel = Excel.createExcel();
          final defaultSheet = excel.getDefaultSheet() ?? 'Sheet1';
          final usedNames = <String>{};
          for (final item in entry.value) {
            final sheetName = _uniqueSheetName(
              _cleanFileName(item.fileName),
              usedNames,
            );
            usedNames.add(sheetName);
            final target = excel[sheetName];
            _copySheet(item.sheet, target, checkCanceled: checkCanceled);
          }
          if (excel.sheets.containsKey(defaultSheet) &&
              !usedNames.contains(defaultSheet)) {
            excel.delete(defaultSheet);
          }

          final base = _sanitizeFileName('${entry.key}.xlsx');
          final filename = _uniqueFileName(base, usedFiles);
          addOutput(excel, filename);
          lastWorkbook = excel;
        }
        break;
      }
    case ExcelMode.internalSummary:
      {
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '汇总: ${file.name}');
          final wb = readExcel(file);

          final allRows = <Map<String, String>>[];
          final headers = <String>{};
          for (final sheetName in wb.sheets.keys) {
            final source = wb.sheets[sheetName]!;
            final parsed = _sheetToRecords(
              source,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
            );
            if (parsed == null) {
              continue;
            }
            headers.addAll(parsed.headers);
            allRows.addAll(parsed.rows);
          }

          final summaryRows = _aggregateRecords(
            allRows,
            explicitHeaderOrder: headers.toList(),
          );
          final summaryName = _uniqueSheetName('汇总表', wb.sheets.keys.toSet());
          final summarySheet = wb[summaryName];
          _writeRows(summarySheet, summaryRows);
          addOutput(wb, '${_cleanFileName(file.name)}.xlsx');
          lastWorkbook = wb;
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
            sheetMap
                .putIfAbsent(name, () => [])
                .add(_SheetRef(file.name, wb.sheets[name]!));
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
        if (added &&
            tempSheet.isNotEmpty &&
            excel.sheets.containsKey(tempSheet)) {
          excel.delete(tempSheet);
        }
        lastWorkbook = excel;
        addOutput(excel, '同名表汇总.xlsx');
        break;
      }
    case ExcelMode.sameNameSheetSummary:
      {
        final excel = createWorkbookWithSheet('Result');
        final tempName = excel.getDefaultSheet() ?? 'Result';
        var wroteAny = false;
        final sheetMap = <String, List<Sheet>>{};
        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '汇总: ${file.name}');
          final wb = readExcel(file);
          for (final name in wb.sheets.keys) {
            sheetMap.putIfAbsent(name, () => <Sheet>[]).add(wb.sheets[name]!);
          }
        }

        final usedNames = <String>{};
        for (final entry in sheetMap.entries) {
          final allRows = <Map<String, String>>[];
          final headers = <String>{};
          for (final sheet in entry.value) {
            final parsed = _sheetToRecords(
              sheet,
              headerRows: job.headerRows,
              footerRows: job.footerRows,
            );
            if (parsed == null) {
              continue;
            }
            headers.addAll(parsed.headers);
            allRows.addAll(parsed.rows);
          }
          if (allRows.isEmpty) {
            continue;
          }
          final rows = _aggregateRecords(
            allRows,
            explicitHeaderOrder: headers.toList(),
          );
          final outName = _uniqueSheetName(entry.key, usedNames);
          usedNames.add(outName);
          final target = excel[outName];
          _writeRows(target, rows);
          wroteAny = true;
        }
        if (wroteAny && excel.sheets.containsKey(tempName)) {
          excel.delete(tempName);
        }
        lastWorkbook = excel;
        addOutput(excel, '汇总结果（同名表汇总到一表）.xlsx');
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
    case ExcelMode.samePositionSummary:
      {
        final excel = createWorkbookWithSheet('Result');
        final sheet = excel['Result'];
        final cells = _parseCellList(job.cellRange);
        if (cells.isNotEmpty) {
          final rows = <List<String>>[
            ['单元格', '汇总值', '明细'],
          ];
          for (final cellRef in cells) {
            final values = <String>[];
            for (final file in job.files) {
              checkCanceled();
              final wb = readExcel(file);
              for (final source in wb.sheets.values) {
                final index = CellIndex.indexByString(cellRef);
                final value = source.cell(index).value?.toString().trim() ?? '';
                if (value.isNotEmpty) {
                  values.add(value);
                }
              }
            }
            final merged = _aggregatePositionValues(values);
            rows.add([cellRef, merged.summary, merged.detail]);
          }
          _writeRows(sheet, rows);
        } else {
          final orderedKeys = <String>[];
          final valueMap = <String, List<String>>{};
          for (final file in job.files) {
            checkCanceled();
            final wb = readExcel(file);
            for (final source in wb.sheets.values) {
              for (var r = 0; r < source.maxRows; r++) {
                final key =
                    source
                        .cell(
                          CellIndex.indexByColumnRow(
                            columnIndex: 0,
                            rowIndex: r,
                          ),
                        )
                        .value
                        ?.toString()
                        .trim() ??
                    '';
                if (key.isEmpty) {
                  continue;
                }
                final value =
                    source
                        .cell(
                          CellIndex.indexByColumnRow(
                            columnIndex: 1,
                            rowIndex: r,
                          ),
                        )
                        .value
                        ?.toString()
                        .trim() ??
                    '';
                if (!valueMap.containsKey(key)) {
                  valueMap[key] = <String>[];
                  orderedKeys.add(key);
                }
                if (value.isNotEmpty) {
                  valueMap[key]!.add(value);
                }
              }
            }
          }

          final rows = <List<String>>[];
          for (final key in orderedKeys) {
            final merged = _aggregatePositionValues(
              valueMap[key] ?? const <String>[],
            );
            rows.add([key, merged.detail]);
          }
          _writeRows(sheet, rows);
        }
        lastWorkbook = excel;
        addOutput(excel, '汇总结果（同位置汇总到一表）.xlsx');
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
              finalName = _uniqueSheetName(
                '${_cleanFileName(file.name)}-$sheetName',
                usedNames,
              );
            }
            usedNames.add(finalName);
            final targetSheet = excel[finalName];
            _copySheet(
              wb.sheets[sheetName]!,
              targetSheet,
              checkCanceled: checkCanceled,
            );
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
        final aliasMap = _parseAliasRules(job.aliasRules);
        final preferredOrder = _parseFieldOrder(job.fieldOrder);
        final discoveredHeaders = <String>[];
        final headerSet = <String>{};
        final records = <Map<String, String>>[];

        for (var i = 0; i < job.files.length; i++) {
          checkCanceled();
          final file = job.files[i];
          onProgress(i + 1, job.files.length, '读取: ${file.name}');
          final wb = readExcel(file);
          for (final entry in wb.sheets.entries) {
            final sourceSheet = entry.value;
            if (sourceSheet.maxRows == 0 || sourceSheet.maxColumns == 0) {
              continue;
            }
            final headerRow = sourceSheet.row(0);
            final headers = <String>[];
            for (var c = 0; c < sourceSheet.maxColumns; c++) {
              final value = headerRow.isNotEmpty && c < headerRow.length
                  ? headerRow[c]?.value
                  : null;
              final text = _normalizeHeaderName(
                value?.toString() ?? '',
                aliasMap,
              );
              headers.add(text);
              if (text.isNotEmpty && headerSet.add(text)) {
                discoveredHeaders.add(text);
              }
            }

            for (var r = 1; r < sourceSheet.maxRows; r++) {
              checkCanceled();
              final rowCells = sourceSheet.row(r);
              if (rowCells.every((cell) => cell == null)) {
                continue;
              }
              final row = <String, String>{};
              for (var c = 0; c < headers.length; c++) {
                final header = headers[c];
                if (header.isEmpty) {
                  continue;
                }
                final cell = c < rowCells.length ? rowCells[c] : null;
                final text = cell?.value?.toString().trim() ?? '';
                row[header] = text;
              }
              if (row.values.every((value) => value.isEmpty)) {
                continue;
              }
              row['文件名'] = file.name;
              row['来源表'] = entry.key;
              if (headerSet.add('文件名')) {
                discoveredHeaders.add('文件名');
              }
              if (headerSet.add('来源表')) {
                discoveredHeaders.add('来源表');
              }
              records.add(row);
            }
          }
        }

        final headerOrder = _mergePreferredHeaderOrder(
          preferredOrder: preferredOrder,
          discoveredHeaders: discoveredHeaders,
        );
        final excel = createWorkbookWithSheet('动态合并');
        final sheet = excel['动态合并'];
        _writeRows(sheet, _recordsToTable(records, headerOrder));
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

  return ExcelJobResult(
    outputs: outputs,
    sheetNames: sheetNames,
    preview: preview,
  );
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
      sheet
          .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r))
          .value = TextCellValue(
        value?.toString() ?? '',
      );
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

List<String> _parseFieldOrder(String input) {
  final parts = input
      .split(RegExp(r'[,，;\n\r]+'))
      .map((item) => item.trim())
      .where((item) => item.isNotEmpty);
  final result = <String>[];
  final seen = <String>{};
  for (final item in parts) {
    if (seen.add(item)) {
      result.add(item);
    }
  }
  return result;
}

Map<String, String> _parseAliasRules(String input) {
  final rules = <String, String>{};
  final parts = input
      .split(RegExp(r'[\n\r,，;；]+'))
      .map((item) => item.trim())
      .where((item) => item.isNotEmpty);
  for (final item in parts) {
    String? left;
    String? right;
    if (item.contains('=>')) {
      final pair = item.split('=>');
      if (pair.length >= 2) {
        left = pair.first.trim();
        right = pair.sublist(1).join('=>').trim();
      }
    } else if (item.contains('->')) {
      final pair = item.split('->');
      if (pair.length >= 2) {
        left = pair.first.trim();
        right = pair.sublist(1).join('->').trim();
      }
    } else if (item.contains('=')) {
      final pair = item.split('=');
      if (pair.length >= 2) {
        left = pair.first.trim();
        right = pair.sublist(1).join('=').trim();
      }
    } else if (item.contains('：')) {
      final pair = item.split('：');
      if (pair.length >= 2) {
        left = pair.first.trim();
        right = pair.sublist(1).join('：').trim();
      }
    } else if (item.contains(':')) {
      final pair = item.split(':');
      if (pair.length >= 2) {
        left = pair.first.trim();
        right = pair.sublist(1).join(':').trim();
      }
    }
    if (left == null || right == null || left.isEmpty || right.isEmpty) {
      continue;
    }
    rules[left] = right;
    rules[left.toLowerCase()] = right;
  }
  return rules;
}

String _normalizeHeaderName(String input, Map<String, String> aliasMap) {
  final raw = input.trim().replaceAll(RegExp(r'\s+'), ' ');
  if (raw.isEmpty) {
    return '';
  }
  return aliasMap[raw] ?? aliasMap[raw.toLowerCase()] ?? raw;
}

List<String> _mergePreferredHeaderOrder({
  required List<String> preferredOrder,
  required List<String> discoveredHeaders,
}) {
  final output = <String>[];
  final seen = <String>{};
  for (final item in preferredOrder) {
    final name = item.trim();
    if (name.isNotEmpty && seen.add(name)) {
      output.add(name);
    }
  }
  for (final item in discoveredHeaders) {
    final name = item.trim();
    if (name.isNotEmpty && seen.add(name)) {
      output.add(name);
    }
  }
  return output;
}

Excel _reorderWorkbookColumns(
  Excel source, {
  required List<String> preferredOrder,
  required Map<String, String> aliasMap,
  required void Function() checkCanceled,
}) {
  final out = Excel.createExcel();
  final defaultSheet = out.getDefaultSheet() ?? 'Sheet1';
  final usedNames = <String>{};
  for (final entry in source.sheets.entries) {
    checkCanceled();
    final outSheetName = _uniqueSheetName(entry.key, usedNames);
    usedNames.add(outSheetName);
    final target = out[outSheetName];
    _reorderSingleSheet(
      source: entry.value,
      target: target,
      preferredOrder: preferredOrder,
      aliasMap: aliasMap,
      checkCanceled: checkCanceled,
    );
  }
  if (out.sheets.containsKey(defaultSheet) &&
      !usedNames.contains(defaultSheet)) {
    out.delete(defaultSheet);
  }
  return out;
}

void _reorderSingleSheet({
  required Sheet source,
  required Sheet target,
  required List<String> preferredOrder,
  required Map<String, String> aliasMap,
  required void Function() checkCanceled,
}) {
  if (source.maxRows == 0 || source.maxColumns == 0) {
    return;
  }

  final headerRow = _detectHeaderRowIndex(
    source: source,
    preferredOrder: preferredOrder,
    aliasMap: aliasMap,
  );

  final sourceHeaders = <String>[];
  for (var c = 0; c < source.maxColumns; c++) {
    final value =
        source
            .cell(
              CellIndex.indexByColumnRow(columnIndex: c, rowIndex: headerRow),
            )
            .value
            ?.toString() ??
        '';
    sourceHeaders.add(_normalizeHeaderName(value, aliasMap));
  }

  final normalizedPreferred = preferredOrder
      .map((item) => _normalizeHeaderName(item, aliasMap))
      .where((item) => item.isNotEmpty)
      .toList();

  final targetHeaders = <String>[];
  final mapping = <int?>[];
  final usedCols = <int>{};

  if (normalizedPreferred.isNotEmpty) {
    for (final header in normalizedPreferred) {
      var matchedCol = -1;
      for (var c = 0; c < sourceHeaders.length; c++) {
        if (usedCols.contains(c)) {
          continue;
        }
        if (sourceHeaders[c] == header) {
          matchedCol = c;
          break;
        }
      }
      targetHeaders.add(header);
      if (matchedCol >= 0) {
        usedCols.add(matchedCol);
        mapping.add(matchedCol);
      } else {
        mapping.add(null);
      }
    }
  } else {
    for (var c = 0; c < sourceHeaders.length; c++) {
      final header = sourceHeaders[c].isEmpty ? '列${c + 1}' : sourceHeaders[c];
      targetHeaders.add(header);
      mapping.add(c);
      usedCols.add(c);
    }
  }

  for (var c = 0; c < sourceHeaders.length; c++) {
    if (usedCols.contains(c)) {
      continue;
    }
    final header = sourceHeaders[c].isEmpty ? '列${c + 1}' : sourceHeaders[c];
    targetHeaders.add(header);
    mapping.add(c);
  }

  for (var r = 0; r < headerRow; r++) {
    checkCanceled();
    final row = source.row(r);
    for (var c = 0; c < source.maxColumns; c++) {
      final data = c < row.length ? row[c] : null;
      if (data == null) {
        continue;
      }
      final cell = target.cell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
      );
      cell.value = data.value;
      if (data.cellStyle != null) {
        cell.cellStyle = data.cellStyle;
      }
    }
  }

  for (var c = 0; c < targetHeaders.length; c++) {
    target
        .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: headerRow))
        .value = TextCellValue(
      targetHeaders[c],
    );
  }

  for (var r = headerRow + 1; r < source.maxRows; r++) {
    checkCanceled();
    final row = source.row(r);
    for (var c = 0; c < mapping.length; c++) {
      final sourceCol = mapping[c];
      if (sourceCol == null) {
        continue;
      }
      final data = sourceCol < row.length ? row[sourceCol] : null;
      if (data == null) {
        continue;
      }
      final cell = target.cell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
      );
      cell.value = data.value;
      if (data.cellStyle != null) {
        cell.cellStyle = data.cellStyle;
      }
    }
  }

  source.getRowHeights.forEach((row, height) {
    target.setRowHeight(row, height);
  });
  for (var outCol = 0; outCol < mapping.length; outCol++) {
    final sourceCol = mapping[outCol];
    if (sourceCol == null) {
      continue;
    }
    final width = source.getColumnWidths[sourceCol];
    if (width != null) {
      target.setColumnWidth(outCol, width);
    }
  }
}

int _detectHeaderRowIndex({
  required Sheet source,
  required List<String> preferredOrder,
  required Map<String, String> aliasMap,
}) {
  final preferred = preferredOrder
      .map((item) => _normalizeHeaderName(item, aliasMap))
      .where((item) => item.isNotEmpty)
      .toSet();
  final maxScan = min(source.maxRows, 20);
  var bestRow = 0;
  var bestScore = -1;
  for (var r = 0; r < maxScan; r++) {
    var nonEmpty = 0;
    var matched = 0;
    for (var c = 0; c < source.maxColumns; c++) {
      final value =
          source
              .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r))
              .value
              ?.toString() ??
          '';
      final header = _normalizeHeaderName(value, aliasMap);
      if (header.isEmpty) {
        continue;
      }
      nonEmpty++;
      if (preferred.contains(header)) {
        matched++;
      }
    }
    final score = preferred.isEmpty ? nonEmpty : matched * 100 + nonEmpty;
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow;
}

_SheetRecords? _sheetToRecords(
  Sheet source, {
  required int headerRows,
  required int footerRows,
}) {
  if (source.maxRows == 0 || source.maxColumns == 0) {
    return null;
  }

  var headerRow = headerRows;
  if (headerRow < 0) {
    headerRow = 0;
  }
  if (headerRow >= source.maxRows) {
    headerRow = source.maxRows - 1;
  }

  String readCell(int row, int col) {
    return source
            .cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row))
            .value
            ?.toString()
            .trim() ??
        '';
  }

  bool rowHasAnyValue(int row) {
    for (var c = 0; c < source.maxColumns; c++) {
      if (readCell(row, c).isNotEmpty) {
        return true;
      }
    }
    return false;
  }

  if (!rowHasAnyValue(headerRow)) {
    for (var r = headerRow; r < source.maxRows; r++) {
      if (rowHasAnyValue(r)) {
        headerRow = r;
        break;
      }
    }
  }

  final endRow = source.maxRows - 1 - footerRows;
  if (endRow <= headerRow) {
    return null;
  }

  final headers = <String>[];
  final usedHeaders = <String, int>{};
  for (var c = 0; c < source.maxColumns; c++) {
    var header = readCell(headerRow, c);
    if (header.isEmpty) {
      header = '列${c + 1}';
    }
    final count = usedHeaders[header] ?? 0;
    usedHeaders[header] = count + 1;
    if (count > 0) {
      header = '${header}_${count + 1}';
    }
    headers.add(header);
  }

  final rows = <Map<String, String>>[];
  for (var r = headerRow + 1; r <= endRow; r++) {
    final row = <String, String>{};
    var hasValue = false;
    for (var c = 0; c < headers.length; c++) {
      final value = readCell(r, c);
      if (value.isNotEmpty) {
        hasValue = true;
      }
      row[headers[c]] = value;
    }
    if (hasValue) {
      rows.add(row);
    }
  }

  if (rows.isEmpty) {
    return null;
  }
  return _SheetRecords(headers: headers, rows: rows);
}

List<List<String>> _aggregateRecords(
  List<Map<String, String>> rows, {
  List<String>? explicitHeaderOrder,
}) {
  if (rows.isEmpty) {
    return const <List<String>>[];
  }

  final headerOrder = <String>[];
  final seen = <String>{};
  if (explicitHeaderOrder != null) {
    for (final header in explicitHeaderOrder) {
      if (header.trim().isEmpty || seen.contains(header)) {
        continue;
      }
      headerOrder.add(header);
      seen.add(header);
    }
  }
  for (final row in rows) {
    for (final key in row.keys) {
      if (key.trim().isEmpty || seen.contains(key)) {
        continue;
      }
      headerOrder.add(key);
      seen.add(key);
    }
  }

  final valuesByHeader = <String, List<String>>{};
  for (final header in headerOrder) {
    valuesByHeader[header] = rows
        .map((row) => row[header]?.trim() ?? '')
        .where((value) => value.isNotEmpty)
        .toList();
  }

  final numericHeaders = <String>{};
  final keyHeaders = <String>{};
  for (final header in headerOrder) {
    final values = valuesByHeader[header]!;
    if (values.isNotEmpty &&
        values.every((value) => _tryParseNumber(value) != null)) {
      numericHeaders.add(header);
    }
    final lower = header.toLowerCase();
    final isKey = RegExp(r'(名称|名字|品名|编号|编码|id|code|型号|item)').hasMatch(lower);
    if (isKey) {
      keyHeaders.add(header);
    }
  }

  if (keyHeaders.isEmpty) {
    final candidate = headerOrder.firstWhere(
      (header) => !numericHeaders.contains(header),
      orElse: () => headerOrder.first,
    );
    keyHeaders.add(candidate);
  }

  final listHeaders = <String>[];
  final sumHeaders = <String>[];
  for (final header in headerOrder) {
    if (keyHeaders.contains(header)) {
      continue;
    }
    final isMonthLike = header.contains('月');
    if (numericHeaders.contains(header) && !isMonthLike) {
      sumHeaders.add(header);
    } else {
      listHeaders.add(header);
    }
  }

  final groups = <String, _AggregateGroup>{};
  for (final row in rows) {
    final key = keyHeaders
        .map((header) => row[header]?.trim() ?? '')
        .join('|#|');
    final group = groups.putIfAbsent(
      key,
      () => _AggregateGroup(
        keyValues: {for (final h in keyHeaders) h: row[h]?.trim() ?? ''},
      ),
    );

    for (final header in listHeaders) {
      final value = row[header]?.trim() ?? '';
      if (value.isNotEmpty) {
        group.listValues
            .putIfAbsent(header, LinkedHashSet<String>.new)
            .add(value);
      }
    }
    for (final header in sumHeaders) {
      final value = row[header]?.trim() ?? '';
      final number = _tryParseNumber(value);
      if (number != null) {
        group.sumValues[header] = (group.sumValues[header] ?? 0) + number;
        final detail = group.sumDetails.putIfAbsent(
          header,
          () => <String, int>{},
        );
        detail[value] = (detail[value] ?? 0) + 1;
      }
    }
  }

  final outHeader = <String>[
    ...keyHeaders,
    ...listHeaders,
    ...sumHeaders,
    ...sumHeaders.map((header) => '$header明细'),
  ];

  final table = <List<String>>[outHeader];
  for (final group in groups.values) {
    final row = <String>[];
    for (final header in keyHeaders) {
      row.add(group.keyValues[header] ?? '');
    }
    for (final header in listHeaders) {
      final list = group.listValues[header];
      row.add(list == null ? '' : list.join(','));
    }
    for (final header in sumHeaders) {
      row.add(_formatNumber(group.sumValues[header] ?? 0));
    }
    for (final header in sumHeaders) {
      final detail = group.sumDetails[header];
      if (detail == null || detail.isEmpty) {
        row.add('');
      } else {
        row.add(
          detail.entries
              .map((entry) => '${entry.key}*${entry.value}')
              .join(','),
        );
      }
    }
    table.add(row);
  }

  return table;
}

_PositionAggregate _aggregatePositionValues(List<String> values) {
  final clean = values
      .map((value) => value.trim())
      .where((value) => value.isNotEmpty)
      .toList();
  if (clean.isEmpty) {
    return const _PositionAggregate(summary: '', detail: '');
  }
  final numbers = clean.map(_tryParseNumber).toList();
  final allNumeric = numbers.every((value) => value != null);
  if (allNumeric) {
    final sum = numbers.fold<double>(0, (prev, value) => prev + (value ?? 0));
    return _PositionAggregate(
      summary: _formatNumber(sum),
      detail: clean.join('+'),
    );
  }
  final merged = LinkedHashSet<String>.of(clean).join(',');
  return _PositionAggregate(summary: merged, detail: merged);
}

double? _tryParseNumber(String text) {
  final clean = text.replaceAll(',', '').trim();
  if (!RegExp(r'^-?\d+(\.\d+)?$').hasMatch(clean)) {
    return null;
  }
  return double.tryParse(clean);
}

String _formatNumber(double value) {
  if (value == value.roundToDouble()) {
    return value.toInt().toString();
  }
  return value
      .toStringAsFixed(4)
      .replaceFirst(RegExp(r'0+$'), '')
      .replaceFirst(RegExp(r'\.$'), '');
}

List<List<String>> _recordsToTable(
  List<Map<String, String>> rows,
  List<String> headerOrder,
) {
  final table = <List<String>>[headerOrder];
  for (final row in rows) {
    table.add(
      headerOrder
          .map((header) => row[header]?.toString().trim() ?? '')
          .toList(),
    );
  }
  return table;
}

String _resolveSplitHeader(List<String> headers, String splitKey) {
  if (headers.isEmpty) {
    return '';
  }
  final key = splitKey.trim();
  if (key.isEmpty) {
    return headers.first;
  }

  for (final header in headers) {
    if (header.toLowerCase() == key.toLowerCase()) {
      return header;
    }
  }

  final colIndex = _columnLabelToIndex(key);
  if (colIndex != null && colIndex >= 0 && colIndex < headers.length) {
    return headers[colIndex];
  }

  final numericIndex = int.tryParse(key);
  if (numericIndex != null &&
      numericIndex >= 1 &&
      numericIndex <= headers.length) {
    return headers[numericIndex - 1];
  }

  return headers.first;
}

int? _columnLabelToIndex(String input) {
  final text = input.trim().toUpperCase();
  if (!RegExp(r'^[A-Z]+$').hasMatch(text)) {
    return null;
  }
  var result = 0;
  for (var i = 0; i < text.length; i++) {
    result = result * 26 + (text.codeUnitAt(i) - 'A'.codeUnitAt(0) + 1);
  }
  return result - 1;
}

String _sanitizeFileName(String input) {
  final name = input.replaceAll(RegExp(r'[\\/:*?"<>|]'), '_').trim();
  if (name.isEmpty) {
    return 'output.xlsx';
  }
  return name;
}

String _uniqueFileName(String baseName, Set<String> used) {
  var name = _sanitizeFileName(baseName);
  if (used.add(name)) {
    return name;
  }
  final dot = name.lastIndexOf('.');
  final stem = dot > 0 ? name.substring(0, dot) : name;
  final ext = dot > 0 ? name.substring(dot) : '';
  var i = 2;
  while (true) {
    final candidate = _sanitizeFileName('${stem}_$i$ext');
    if (used.add(candidate)) {
      return candidate;
    }
    i++;
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
      final cell = target.cell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: newRow),
      );
      cell.value = data.value;
      if (data.cellStyle != null) {
        cell.cellStyle = data.cellStyle;
      }
    }

    if (sourceName != null) {
      final linkCell = target.cell(
        CellIndex.indexByColumnRow(
          columnIndex: sourceMaxCol + 1,
          rowIndex: newRow,
        ),
      );
      linkCell.value = TextCellValue(sourceName);
      linkCell.cellStyle = _linkStyle();
    }
  }

  if (source.spannedItems.isNotEmpty) {
    for (final span in source.spannedItems) {
      final range = _parseSpan(span);
      if (range == null) continue;
      if (range.start.rowIndex < sourceStart ||
          range.end.rowIndex > sourceEnd) {
        continue;
      }
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
      final cell = target.cell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
      );
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

void _writeRows(Sheet sheet, List<List<String>> rows, {CellStyle? linkStyle}) {
  for (var r = 0; r < rows.length; r++) {
    _writeRow(sheet, r, rows[r]);
    if (r > 0 && linkStyle != null) {
      final cell = sheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
      );
      cell.cellStyle = linkStyle;
      final second = sheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r),
      );
      second.cellStyle = linkStyle;
    }
  }
}

void _writeRow(Sheet sheet, int rowIndex, List<String> values) {
  for (var c = 0; c < values.length; c++) {
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: rowIndex))
        .value = TextCellValue(
      values[c],
    );
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

class _SheetRecords {
  _SheetRecords({required this.headers, required this.rows});

  final List<String> headers;
  final List<Map<String, String>> rows;
}

class _AggregateGroup {
  _AggregateGroup({required this.keyValues});

  final Map<String, String> keyValues;
  final Map<String, LinkedHashSet<String>> listValues =
      <String, LinkedHashSet<String>>{};
  final Map<String, double> sumValues = <String, double>{};
  final Map<String, Map<String, int>> sumDetails = <String, Map<String, int>>{};
}

class _PositionAggregate {
  const _PositionAggregate({required this.summary, required this.detail});

  final String summary;
  final String detail;
}

class _SheetRef {
  _SheetRef(this.fileName, this.sheet);

  final String fileName;
  final Sheet sheet;
}
