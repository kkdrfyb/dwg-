import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:office_toolbox_flutter/features/excel/excel_isolate.dart';
import 'package:office_toolbox_flutter/features/excel/excel_models.dart';

void main() {
  test(
    'split worksheet keeps template style/merge and writes numeric cells',
    () async {
      final tempDir = await Directory.systemTemp.createTemp(
        'excel_split_case_',
      );

      Future<ExcelInputFile> createSource(
        String filename,
        List<List<String>> rows,
      ) async {
        final excel = Excel.createExcel();
        final defaultSheet = excel.getDefaultSheet() ?? 'Sheet1';
        if (defaultSheet != '1月份') {
          excel.rename(defaultSheet, '1月份');
        }
        final sheet = excel['1月份'];
        final headerStyle = CellStyle(
          backgroundColorHex: ExcelColor.yellow100,
          leftBorder: Border(borderStyle: BorderStyle.Thin),
          rightBorder: Border(borderStyle: BorderStyle.Thin),
          topBorder: Border(borderStyle: BorderStyle.Thin),
          bottomBorder: Border(borderStyle: BorderStyle.Thin),
          bold: true,
        );

        for (var r = 0; r < rows.length; r++) {
          final row = rows[r];
          for (var c = 0; c < row.length; c++) {
            final cell = sheet.cell(
              CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
            );
            cell.value = TextCellValue(row[c]);
            if (r == 1) {
              cell.cellStyle = headerStyle;
            }
          }
        }

        sheet.merge(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
          CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),
        );

        final path = '${tempDir.path}${Platform.pathSeparator}$filename';
        await File(path).writeAsBytes(excel.encode()!, flush: true);
        return ExcelInputFile(
          name: filename,
          path: path,
          size: File(path).lengthSync(),
        );
      }

      final fileA = await createSource('门店A.xlsx', <List<String>>[
        <String>['连锁店：门店A', '', ''],
        <String>['水果名', '产地', '销量/元'],
        <String>['苹果', '山东', '1164'],
        <String>['西瓜', '江苏', '750'],
        <String>['制表人：张伟', '', ''],
      ]);
      final fileB = await createSource('门店B.xlsx', <List<String>>[
        <String>['连锁店：门店B', '', ''],
        <String>['水果名', '产地', '销量/元'],
        <String>['苹果', '河北', '1090'],
        <String>['芒果', '广东', '680'],
        <String>['制表人：李四', '', ''],
      ]);

      final runner = ExcelIsolateRunner();
      final job = ExcelJob(
        mode: ExcelMode.splitWorksheet,
        files: <ExcelInputFile>[fileA, fileB],
        headerRows: 2,
        footerRows: 1,
        splitKey: '水果名',
        splitWorksheetOutputMode: SplitWorksheetOutputMode.oneWorkbook,
      );

      final result = await runner.run(job, onProgress: (_) {});
      expect(result.outputs.length, 1);

      final outExcel = Excel.decodeBytes(result.outputs.first.bytes);
      final totalSheet = outExcel['Result(总表)'];
      final appleSheet = outExcel['苹果'];

      String value(Sheet sheet, int row, int col) {
        return sheet
                .cell(
                  CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
                )
                .value
                ?.toString() ??
            '';
      }

      expect(totalSheet.spannedItems, contains('A1:C1'));
      expect(value(totalSheet, 0, 0), '连锁店：门店A');
      expect(value(totalSheet, 1, 0), '水果名');
      expect(value(totalSheet, totalSheet.maxRows - 1, 0), '制表人：张伟');

      final headerCell = totalSheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
      );
      expect(headerCell.cellStyle?.backgroundColor, ExcelColor.yellow100);
      expect(headerCell.cellStyle?.bottomBorder.borderStyle, BorderStyle.Thin);

      expect(value(appleSheet, 0, 0), '连锁店：门店A');
      expect(value(appleSheet, 1, 0), '水果名');
      expect(value(appleSheet, 2, 0), '苹果');
      expect(value(appleSheet, 3, 0), '苹果');
      expect(value(appleSheet, appleSheet.maxRows - 1, 0), '制表人：张伟');

      final salesCell = appleSheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 2),
      );
      expect(salesCell.value, isA<IntCellValue>());
      expect((salesCell.value as IntCellValue).value, 1164);

      await tempDir.delete(recursive: true);
    },
  );

  test(
    'same position summary writes numeric values and grid borders',
    () async {
      final tempDir = await Directory.systemTemp.createTemp('excel_type_case_');

      final excel = Excel.createExcel();
      final sheetName = excel.getDefaultSheet() ?? 'Sheet1';
      final sheet = excel[sheetName];
      sheet
          .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0))
          .value = TextCellValue(
        '名称',
      );
      sheet
          .cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 0))
          .value = TextCellValue(
        '100.5',
      );

      final inputPath = '${tempDir.path}${Platform.pathSeparator}source.xlsx';
      await File(inputPath).writeAsBytes(excel.encode()!, flush: true);

      final job = ExcelJob(
        mode: ExcelMode.samePositionSummary,
        files: <ExcelInputFile>[
          ExcelInputFile(
            name: 'source.xlsx',
            path: inputPath,
            size: File(inputPath).lengthSync(),
          ),
        ],
        cellRange: 'B1',
      );

      final result = await ExcelIsolateRunner().run(job, onProgress: (_) {});
      expect(result.outputs.length, 1);

      final outExcel = Excel.decodeBytes(result.outputs.first.bytes);
      final resultSheet = outExcel['Result'];

      final summaryCell = resultSheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
      );
      expect(summaryCell.value, isA<DoubleCellValue>());
      expect(
        (summaryCell.value as DoubleCellValue).value,
        closeTo(100.5, 0.0001),
      );

      final headerCell = resultSheet.cell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
      );
      expect(headerCell.cellStyle?.leftBorder.borderStyle, BorderStyle.Thin);
      expect(headerCell.cellStyle?.rightBorder.borderStyle, BorderStyle.Thin);

      await tempDir.delete(recursive: true);
    },
  );
}
