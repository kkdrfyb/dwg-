import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';

Future<String?> pickSavePath({
  required String dialogTitle,
  required String suggestedName,
  required List<String> allowedExtensions,
}) {
  return FilePicker.platform.saveFile(
    dialogTitle: dialogTitle,
    fileName: suggestedName,
    allowedExtensions: allowedExtensions,
    type: FileType.custom,
  );
}

Future<void> exportCsv({
  required String dialogTitle,
  required String suggestedName,
  required List<String> headers,
  required List<List<String>> rows,
}) async {
  final path = await pickSavePath(
    dialogTitle: dialogTitle,
    suggestedName: suggestedName,
    allowedExtensions: const ['csv'],
  );
  if (path == null) return;

  final buffer = StringBuffer();
  buffer.write('\uFEFF');
  buffer.writeln(headers.map(_escapeCsv).join(','));
  for (final row in rows) {
    buffer.writeln(row.map(_escapeCsv).join(','));
  }

  final file = File(path);
  await file.writeAsString(buffer.toString(), encoding: utf8, flush: true);
}

Future<void> exportXlsx({
  required String dialogTitle,
  required String suggestedName,
  required String sheetName,
  required List<String> headers,
  required List<List<String>> rows,
}) async {
  final path = await pickSavePath(
    dialogTitle: dialogTitle,
    suggestedName: suggestedName,
    allowedExtensions: const ['xlsx'],
  );
  if (path == null) return;

  final excel = Excel.createExcel();
  final sheet = excel[sheetName];
  sheet.appendRow(_toCellRow(headers));
  for (final row in rows) {
    sheet.appendRow(_toCellRow(row));
  }
  if (excel.sheets.length > 1 && excel.sheets.keys.first != sheetName) {
    excel.delete(excel.sheets.keys.first);
  }

  final bytes = excel.encode();
  if (bytes == null) return;
  final file = File(path);
  await file.writeAsBytes(bytes, flush: true);
}

Future<String?> saveBytes({
  required String dialogTitle,
  required String suggestedName,
  required List<String> allowedExtensions,
  required List<int> bytes,
}) async {
  final path = await pickSavePath(
    dialogTitle: dialogTitle,
    suggestedName: suggestedName,
    allowedExtensions: allowedExtensions,
  );
  if (path == null) return null;
  final file = File(path);
  await file.writeAsBytes(bytes, flush: true);
  return path;
}

String _escapeCsv(String value) {
  final needsQuote = value.contains(',') || value.contains('"') || value.contains('\n') || value.contains('\r');
  final escaped = value.replaceAll('"', '""');
  return needsQuote ? '"$escaped"' : escaped;
}

List<CellValue> _toCellRow(List<String> values) {
  return values.map(TextCellValue.new).toList();
}
