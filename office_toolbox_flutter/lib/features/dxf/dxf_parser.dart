import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';

import 'package:gbk_codec/gbk_codec.dart';

import 'dxf_models.dart';

const int dxfLargeFileThreshold = 6 * 1024 * 1024;

String decodeDxfText(Uint8List bytes) {
  try {
    return utf8.decode(bytes);
  } catch (_) {}
  try {
    return gbk_bytes.decode(bytes);
  } catch (_) {}
  try {
    return latin1.decode(bytes);
  } catch (_) {}
  return utf8.decode(bytes, allowMalformed: true);
}

String readDxfFile(String path) {
  final bytes = File(path).readAsBytesSync();
  return decodeDxfText(bytes);
}

List<DxfEntityText> parseDxfEntities(String content) {
  final lines = content.split(RegExp(r'\r?\n'));
  final entities = <DxfEntityText>[];

  bool inEntities = false;
  bool expectSectionName = false;
  String? currentType;
  String currentLayer = '';
  final currentText = StringBuffer();

  void finalize() {
    if (currentType == null) return;
    final text = _normalizeText(currentText.toString());
    if (text.trim().isNotEmpty) {
      entities.add(
        DxfEntityText(
          type: currentType!,
          layer: currentLayer,
          text: text,
        ),
      );
    }
    currentType = null;
    currentLayer = '';
    currentText.clear();
  }

  for (var i = 0; i + 1 < lines.length; i += 2) {
    final code = int.tryParse(lines[i].trim());
    final value = lines[i + 1];
    if (code == null) continue;

    if (code == 0) {
      final marker = value.trim();
      if (marker == 'SECTION') {
        expectSectionName = true;
        continue;
      }
      if (marker == 'ENDSEC') {
        finalize();
        inEntities = false;
        expectSectionName = false;
        continue;
      }
      if (!inEntities) {
        continue;
      }
      finalize();
      currentType = marker;
      currentLayer = '';
      currentText.clear();
      continue;
    }

    if (expectSectionName && code == 2) {
      final sectionName = value.trim();
      inEntities = sectionName == 'ENTITIES';
      expectSectionName = false;
      continue;
    }

    if (!inEntities || currentType == null) continue;

    if (code == 8) {
      currentLayer = value.trim();
      continue;
    }

    if ((code == 1 || code == 3) && _isTextEntity(currentType!)) {
      if (currentText.isNotEmpty) {
        currentText.write('\n');
      }
      currentText.write(value);
      continue;
    }

    if (code == 2 && currentType == 'INSERT') {
      currentText.write(value.trim());
      continue;
    }
  }

  finalize();
  return entities;
}

bool _isTextEntity(String type) {
  return type == 'TEXT' || type == 'MTEXT' || type == 'ATTRIB';
}

String _normalizeText(String text) {
  return text.replaceAll('\\P', '\n').replaceAll('\\~', ' ');
}
