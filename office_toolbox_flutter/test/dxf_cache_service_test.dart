import 'dart:io';

import 'package:flutter_test/flutter_test.dart';
import 'package:office_toolbox_flutter/core/logging/log_service.dart';
import 'package:office_toolbox_flutter/features/dxf/dxf_cache_service.dart';

void main() {
  late DxfCacheService service;

  setUp(() {
    service = DxfCacheService(log: LogService.forTesting());
  });

  test('build ODA command should keep quoted single argument line', () {
    final command = service.buildOdaProcessCommandForTest(
      r'D:\office_toolbox\ODAFileConverter\ODAFileConverter.exe',
      <String>[
        r'F:\测试目录 A\01 目录.dwg',
        r'F:\测试目录 A\output',
        'ACAD2010',
        'DXF',
        '0',
        '0',
        '01 目录.dwg',
      ],
      expectedOutputPath: r'F:\测试目录 A\output\01 目录.dxf',
      timeoutSeconds: 90,
    );

    expect(command, contains('-ArgumentList'));
    expect(command, contains(r'"F:\测试目录 A\01 目录.dwg"'));
    expect(command, contains(r'"F:\测试目录 A\output"'));
    expect(command, contains('-WindowStyle Hidden -PassThru'));
    expect(command, isNot(contains('-ArgumentList @(')));
  });

  test('legacy version mapping should be stable', () {
    expect(service.toLegacyVersionForTest('DXF2010'), 'ACAD2010');
    expect(service.toLegacyVersionForTest('DWG2018'), 'ACAD2018');
    expect(service.toLegacyVersionForTest('ACAD2013'), 'ACAD2013');
    expect(service.toLegacyVersionForTest('UNKNOWN'), 'ACAD2010');
  });

  test('waitForFileReady returns for a stable file', () async {
    final tempDir = await Directory.systemTemp.createTemp('dxf_ready_test_');
    final file = File('${tempDir.path}${Platform.pathSeparator}sample.dxf');
    file.writeAsStringSync('0\nEOF\n');

    await service
        .waitForFileReadyForTest(file.path)
        .timeout(const Duration(seconds: 5));

    expect(file.existsSync(), isTrue);
    tempDir.deleteSync(recursive: true);
  });

  test('waitForFileReady ignores missing files', () async {
    final tempDir = await Directory.systemTemp.createTemp('dxf_ready_test_');
    final missingPath =
        '${tempDir.path}${Platform.pathSeparator}missing_file.dxf';

    await service.waitForFileReadyForTest(missingPath);
    expect(File(missingPath).existsSync(), isFalse);

    tempDir.deleteSync(recursive: true);
  });
}
