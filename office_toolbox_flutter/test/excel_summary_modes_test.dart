import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:office_toolbox_flutter/features/excel/excel_isolate.dart';
import 'package:office_toolbox_flutter/features/excel/excel_models.dart';

void main() {
  test(
    'merge to workbook summary matches sample aggregation columns',
    () async {
      final tempDir = await Directory.systemTemp.createTemp(
        'excel_summary_case_',
      );

      final fileA = await _createStoreWorkbook(
        tempDir,
        filename: '门店A.xlsx',
        month1Rows: const <List<dynamic>>[
          <dynamic>['苹果', '山东', 'A', 1, 1164],
          <dynamic>['西瓜', '江苏', 'B', 1, 750],
          <dynamic>['苹果', '河北', 'B', 1, 890],
          <dynamic>['圣女果', '福建', 'A', 1, 590],
          <dynamic>['芒果', '广东', 'C', 1, 680],
        ],
        month2Rows: const <List<dynamic>>[
          <dynamic>['苹果', '山东', 'A', 2, 990],
          <dynamic>['西瓜', '江苏', 'B', 2, 862],
          <dynamic>['苹果', '河北', 'B', 2, 1140],
          <dynamic>['圣女果', '福建', 'A', 2, 600],
          <dynamic>['芒果', '广东', 'C', 2, 702],
        ],
      );

      final fileB = await _createStoreWorkbook(
        tempDir,
        filename: '门店B.xlsx',
        month1Rows: const <List<dynamic>>[
          <dynamic>['葡萄', '山东', 'A', 1, 2164],
          <dynamic>['哈密瓜', '江苏', 'B', 1, 1750],
          <dynamic>['苹果', '河北', 'B', 1, 1090],
          <dynamic>['圣女果', '福建', 'A', 1, 990],
          <dynamic>['火龙果', '广东', 'C', 1, 866],
        ],
        month2Rows: const <List<dynamic>>[
          <dynamic>['葡萄', '山东', 'A', 2, 1964],
          <dynamic>['哈密瓜', '江苏', 'B', 2, 1850],
          <dynamic>['苹果', '河北', 'B', 2, 890],
          <dynamic>['圣女果', '福建', 'A', 2, 1020],
          <dynamic>['火龙果', '广东', 'C', 2, 960],
        ],
      );

      final result = await ExcelIsolateRunner().run(
        ExcelJob(
          mode: ExcelMode.mergeToSheetSummary,
          files: <ExcelInputFile>[fileA, fileB],
          headerRows: 2,
          footerRows: 1,
        ),
        onProgress: (_) {},
      );

      expect(result.outputs.length, 1);
      final out = Excel.decodeBytes(result.outputs.first.bytes);
      final sheet = out['Result'];

      expect(_value(sheet, 0, 0), '水果名');
      expect(_value(sheet, 0, 1), '产地');
      expect(_value(sheet, 0, 2), '品级');
      expect(_value(sheet, 0, 3), '月份');
      expect(_value(sheet, 0, 4), '销量/元');
      expect(_value(sheet, 0, 5), '销量/元');

      final appleRow = _findRowByFirstCell(sheet, '苹果');
      expect(appleRow, isNot(-1));
      expect(_value(sheet, appleRow, 1), '山东,河北');
      expect(_value(sheet, appleRow, 2), 'A,B');
      expect(_value(sheet, appleRow, 3), '1,2');
      expect(_value(sheet, appleRow, 4), '6164');
      expect(_value(sheet, appleRow, 5), '1164*1,890*2,990*1,1140*1,1090*1');

      await tempDir.delete(recursive: true);
    },
  );

  test(
    'internal summary keeps compact columns and puts summary sheet first',
    () async {
      final tempDir = await Directory.systemTemp.createTemp(
        'excel_internal_case_',
      );

      final fileA = await _createStoreWorkbook(
        tempDir,
        filename: '门店A.xlsx',
        month1Rows: const <List<dynamic>>[
          <dynamic>['苹果', '山东', 'A', 1, 1164],
          <dynamic>['西瓜', '江苏', 'B', 1, 750],
          <dynamic>['苹果', '河北', 'B', 1, 890],
          <dynamic>['圣女果', '福建', 'A', 1, 590],
          <dynamic>['芒果', '广东', 'C', 1, 680],
        ],
        month2Rows: const <List<dynamic>>[
          <dynamic>['苹果', '山东', 'A', 2, 990],
          <dynamic>['西瓜', '江苏', 'B', 2, 862],
          <dynamic>['苹果', '河北', 'B', 2, 1140],
          <dynamic>['圣女果', '福建', 'A', 2, 600],
          <dynamic>['芒果', '广东', 'C', 2, 702],
        ],
      );

      final result = await ExcelIsolateRunner().run(
        ExcelJob(
          mode: ExcelMode.internalSummary,
          files: <ExcelInputFile>[fileA],
          headerRows: 2,
          footerRows: 1,
        ),
        onProgress: (_) {},
      );

      expect(result.outputs.length, 1);
      final out = Excel.decodeBytes(result.outputs.first.bytes);
      expect(out.sheets.keys.first, '汇总表');
      expect(out.sheets.containsKey('1月份'), isTrue);
      expect(out.sheets.containsKey('2月份'), isTrue);

      final summary = out['汇总表'];
      expect(_value(summary, 0, 0), '水果名');
      expect(_value(summary, 0, 1), '品级');
      expect(_value(summary, 0, 2), '销量/元');
      expect(_value(summary, 0, 3), '销量/元');

      final appleRow = _findRowByFirstCell(summary, '苹果');
      expect(appleRow, isNot(-1));
      expect(_value(summary, appleRow, 1), 'A,B');
      expect(_value(summary, appleRow, 2), '4184');
      expect(_value(summary, appleRow, 3), '1164*1,890*1,990*1,1140*1');

      await tempDir.delete(recursive: true);
    },
  );
}

Future<ExcelInputFile> _createStoreWorkbook(
  Directory tempDir, {
  required String filename,
  required List<List<dynamic>> month1Rows,
  required List<List<dynamic>> month2Rows,
}) async {
  final excel = Excel.createExcel();
  final defaultSheet = excel.getDefaultSheet() ?? 'Sheet1';
  if (defaultSheet != '1月份') {
    excel.rename(defaultSheet, '1月份');
  }
  final month2 = excel['2月份'];
  _fillStoreSheet(
    excel['1月份'],
    title: '连锁店：${filename.replaceAll('.xlsx', '')}',
    dataRows: month1Rows,
    footer: '制表人：张伟',
  );
  _fillStoreSheet(
    month2,
    title: '连锁店：${filename.replaceAll('.xlsx', '')}',
    dataRows: month2Rows,
    footer: '制表人：李四',
  );

  final path = '${tempDir.path}${Platform.pathSeparator}$filename';
  await File(path).writeAsBytes(excel.encode()!, flush: true);
  return ExcelInputFile(
    name: filename,
    path: path,
    size: File(path).lengthSync(),
  );
}

void _fillStoreSheet(
  Sheet sheet, {
  required String title,
  required List<List<dynamic>> dataRows,
  required String footer,
}) {
  final rows = <List<dynamic>>[
    <dynamic>[title, '', '', '', ''],
    const <dynamic>['水果名', '产地', '品级', '月份', '销量/元'],
    ...dataRows,
    <dynamic>[footer, '', '', '日期：2016-3-1', ''],
  ];

  for (var r = 0; r < rows.length; r++) {
    final row = rows[r];
    for (var c = 0; c < row.length; c++) {
      final cell = sheet.cell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
      );
      final value = row[c];
      if (value is int) {
        cell.value = IntCellValue(value);
      } else if (value is double) {
        cell.value = DoubleCellValue(value);
      } else {
        cell.value = TextCellValue(value.toString());
      }
    }
  }
}

String _value(Sheet sheet, int row, int col) {
  return sheet
          .cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row))
          .value
          ?.toString() ??
      '';
}

int _findRowByFirstCell(Sheet sheet, String value) {
  for (var r = 1; r < sheet.maxRows; r++) {
    if (_value(sheet, r, 0) == value) {
      return r;
    }
  }
  return -1;
}
