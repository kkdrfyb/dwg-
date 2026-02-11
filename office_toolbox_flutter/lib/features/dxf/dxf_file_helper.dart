import 'dart:io';

import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:path/path.dart' as p;

Future<List<PlatformFile>> collectDxfFilesFromPaths(List<String> paths) async {
  if (paths.isEmpty) return [];
  final results = await compute(_scanCadPaths, paths);
  return results
      .map(
        (item) => PlatformFile(
          name: item['name'] as String,
          path: item['path'] as String,
          size: item['size'] as int,
        ),
      )
      .toList();
}

List<Map<String, Object>> _scanCadPaths(List<String> paths) {
  final results = <Map<String, Object>>[];
  final seen = <String>{};

  bool isCadFile(String path) {
    final lower = path.toLowerCase();
    return lower.endsWith('.dwg') || lower.endsWith('.dxf');
  }

  void addFile(File file) {
    final normalized = p.normalize(file.path);
    final lower = normalized.toLowerCase();
    final outputMarker = '${p.separator}output${p.separator}'.toLowerCase();
    if (lower.contains(outputMarker)) return;
    if (!isCadFile(normalized)) return;
    if (!seen.add(normalized)) return;
    final stat = file.statSync();
    results.add({
      'path': normalized,
      'name': p.basename(normalized),
      'size': stat.size,
    });
  }

  for (final input in paths) {
    final normalized = p.normalize(input);
    if (normalized.isEmpty) continue;
    final type = FileSystemEntity.typeSync(normalized, followLinks: false);
    if (type == FileSystemEntityType.directory) {
      final dir = Directory(normalized);
      if (!dir.existsSync()) continue;
      for (final entity in dir.listSync(recursive: true, followLinks: false)) {
        if (entity is File) {
          addFile(entity);
        }
      }
      continue;
    }
    if (type == FileSystemEntityType.file) {
      addFile(File(normalized));
    }
  }

  results.sort((a, b) => (a['path'] as String).compareTo(b['path'] as String));
  return results;
}
