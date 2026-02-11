import 'dart:io';

import 'dxf_parser.dart';

Future<String> readDxfFileAsync(String path) async {
  final bytes = await File(path).readAsBytes();
  return decodeDxfText(bytes);
}

String replaceLimited(String text, String before, String after, int count) {
  if (before.isEmpty || count <= 0) return text;
  var remaining = count;
  var idx = 0;
  while (remaining > 0) {
    final pos = text.indexOf(before, idx);
    if (pos == -1) break;
    text = text.substring(0, pos) + after + text.substring(pos + before.length);
    idx = pos + after.length;
    remaining--;
  }
  return text;
}
