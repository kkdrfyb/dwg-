import 'dart:typed_data';

import 'package:flutter_test/flutter_test.dart';
import 'package:gbk_codec/gbk_codec.dart';
import 'package:office_toolbox_flutter/features/dxf/dxf_parser.dart';

void main() {
  test('decode GBK fallback', () {
    final bytes = Uint8List.fromList(gbk_bytes.encode('测试DXF'));
    final decoded = decodeDxfText(bytes);
    expect(decoded, '测试DXF');
  });
}
