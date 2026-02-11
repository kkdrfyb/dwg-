import 'package:flutter_test/flutter_test.dart';
import 'package:office_toolbox_flutter/features/dxf/dxf_parser.dart';

void main() {
  test('parse basic DXF text entities', () {
    const sample = '0\nSECTION\n2\nENTITIES\n0\nTEXT\n8\nLayer1\n1\nHello\n0\nMTEXT\n8\nLayer2\n1\nLine1\\PLine2\n0\nENDSEC\n0\nEOF\n';

    final entities = parseDxfEntities(sample);

    expect(entities.length, 2);
    expect(entities[0].text, 'Hello');
    expect(entities[0].layer, 'Layer1');
    expect(entities[1].text, 'Line1\nLine2');
    expect(entities[1].layer, 'Layer2');
  });
}
