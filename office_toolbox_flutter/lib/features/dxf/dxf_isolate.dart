import 'dart:math';

import 'dxf_models.dart';
import 'dxf_parser.dart';

Map<String, dynamic> scanDxfFileForKeywords(Map<String, dynamic> request) {
  final path = request['path'] as String;
  final name = request['name'] as String;
  final keywords = (request['keywords'] as List).cast<String>();
  final lowerKeywords = keywords.map((k) => k.toLowerCase()).toList();
  final size = request['size'] as int? ?? 0;
  final forcePlain = request['forcePlain'] == true;
  final forceParse = request['forceParse'] == true;
  final largeFile = (size > dxfLargeFileThreshold || forcePlain) && !forceParse;

  try {
    final text = readDxfFile(path);
    if (largeFile) {
      return {
        'ok': true,
        'plainText': true,
        'results': _plainTextMatches(name, text, keywords),
      };
    }
    final entities = parseDxfEntities(text);
    final results = <Map<String, String>>[];
    if (keywords.isEmpty) {
      for (final entity in entities) {
        results.add({
          'fileName': name,
          'objectType': entity.type,
          'layer': entity.layer,
          'keyword': '全部',
          'content': entity.text,
        });
      }
    } else {
      for (final entity in entities) {
        final content = entity.text;
        if (content.isEmpty) continue;
        final lowerContent = content.toLowerCase();
        for (var i = 0; i < lowerKeywords.length; i++) {
          if (lowerContent.contains(lowerKeywords[i])) {
            results.add({
              'fileName': name,
              'objectType': entity.type,
              'layer': entity.layer,
              'keyword': keywords[i],
              'content': content,
            });
          }
        }
      }
    }
    return {'ok': true, 'plainText': false, 'results': results};
  } catch (error) {
    try {
      final text = readDxfFile(path);
      return {
        'ok': false,
        'plainText': true,
        'error': error.toString(),
        'results': _plainTextMatches(name, text, keywords),
      };
    } catch (inner) {
      return {'ok': false, 'error': inner.toString(), 'results': <Map<String, String>>[]};
    }
  }
}

Map<String, dynamic> extractValveInfoFromFile(Map<String, dynamic> request) {
  final path = request['path'] as String;
  final name = request['name'] as String;

  try {
    final text = readDxfFile(path);
    final entities = parseDxfEntities(text);
    final results = _extractValveInfo(entities)
        .map((item) => {
              'fileName': name,
              'kind': item['kind'] ?? '',
              'name': item['name'] ?? '',
              'code': item['code'] ?? '',
              'size': item['size'] ?? '',
              'height': item['height'] ?? '',
            })
        .toList();
    return {'ok': true, 'results': results};
  } catch (error) {
    return {'ok': false, 'error': error.toString(), 'results': <Map<String, String>>[]};
  }
}

Map<String, dynamic> scanDxfFileForReplace(Map<String, dynamic> request) {
  final path = request['path'] as String;
  final name = request['name'] as String;
  final size = request['size'] as int? ?? 0;
  final forcePlain = request['forcePlain'] == true;
  final pairs = (request['pairs'] as List)
      .map((item) => Map<String, String>.from(item as Map))
      .where((pair) => pair['find']?.trim().isNotEmpty == true)
      .toList();
  final rules = pairs
      .map(
        (pair) => _ReplaceRule(
          pair['find'] ?? '',
          pair['replace'] ?? '',
        ),
      )
      .where((rule) => rule.find.isNotEmpty)
      .toList();

  if (rules.isEmpty) {
    return {'ok': true, 'results': <Map<String, String>>[]};
  }

  try {
    final text = readDxfFile(path);
    final largeFile = size > dxfLargeFileThreshold || forcePlain;
    if (largeFile) {
      return {
        'ok': true,
        'plainText': true,
        'results': _plainTextReplaceMatches(name, path, text, pairs),
      };
    }

    final entities = parseDxfEntities(text);
    final results = <Map<String, String>>[];
    for (final entity in entities) {
      final original = entity.text;
      if (original.isEmpty) continue;
      var updated = original;
      final applied = <String>[];
      for (final rule in rules) {
        if (rule.regex.hasMatch(updated)) {
          updated = updated.replaceAll(rule.regex, rule.replaceWith);
          applied.add('${rule.find}→${rule.replaceWith}');
        }
      }
      if (updated != original) {
        results.add({
          'fileName': name,
          'filePath': path,
          'objectType': entity.type,
          'layer': entity.layer,
          'originalText': original,
          'updatedText': updated,
          'rule': applied.join('；'),
        });
      }
    }
    return {'ok': true, 'plainText': false, 'results': results};
  } catch (error) {
    try {
      final text = readDxfFile(path);
      return {
        'ok': false,
        'plainText': true,
        'error': error.toString(),
        'results': _plainTextReplaceMatches(name, path, text, pairs),
      };
    } catch (inner) {
      return {'ok': false, 'error': inner.toString(), 'results': <Map<String, String>>[]};
    }
  }
}

List<Map<String, String>> _plainTextMatches(
  String fileName,
  String text,
  List<String> keywords,
) {
  final results = <Map<String, String>>[];
  if (keywords.isEmpty) {
    results.add({
      'fileName': fileName,
      'objectType': '未知',
      'layer': '-',
      'keyword': '全部',
      'content': '(纯文本匹配)',
    });
    return results;
  }
  for (final keyword in keywords) {
    if (keyword.trim().isEmpty) continue;
    final re = RegExp(RegExp.escape(keyword), caseSensitive: false);
    for (final match in re.allMatches(text)) {
      final pos = match.start;
      final contextStart = max(0, pos - 20);
      final contextEnd = min(text.length, pos + keyword.length + 20);
      final snippet = text.substring(contextStart, contextEnd);
      results.add({
        'fileName': fileName,
        'objectType': '文本',
        'layer': '-',
        'keyword': keyword,
        'content': snippet,
      });
    }
  }
  return results;
}

List<Map<String, String>> _plainTextReplaceMatches(
  String fileName,
  String filePath,
  String text,
  List<Map<String, String>> pairs,
) {
  final results = <Map<String, String>>[];
  for (final pair in pairs) {
    final find = pair['find'] ?? '';
    final replaceWith = pair['replace'] ?? '';
    if (find.isEmpty) continue;
    final re = RegExp(RegExp.escape(find), caseSensitive: false);
    for (final match in re.allMatches(text)) {
      final pos = match.start;
      final contextStart = max(0, pos - 20);
      final contextEnd = min(text.length, pos + find.length + 20);
      final originalSnippet = text.substring(contextStart, contextEnd);
      final updatedSnippet = originalSnippet.replaceFirst(re, replaceWith);
      results.add({
        'fileName': fileName,
        'filePath': filePath,
        'objectType': '文本',
        'layer': '-',
        'originalText': originalSnippet,
        'updatedText': updatedSnippet,
        'rule': '$find→$replaceWith',
      });
    }
  }
  return results;
}

class _ReplaceRule {
  _ReplaceRule(this.find, this.replaceWith)
      : regex = RegExp(RegExp.escape(find), caseSensitive: false);

  final String find;
  final String replaceWith;
  final RegExp regex;
}

List<Map<String, String>> _extractValveInfo(List<DxfEntityText> entities) {
  final results = <Map<String, String>>[];

  final sizeRe = RegExp(r'(\d+)\s*[xX×]\s*(\d+)');
  final heightRe = RegExp(r'(顶[0-9\.]+m|标高[:：]?\s*[0-9\.]+m)', caseSensitive: false);
  final idRe = RegExp(r'[A-Za-z0-9]{6,}');
  const invalidKeywords = ['排风系统', '系统', '标高', '顶标高', 'top', '尺寸'];

  bool isInvalid(String text) {
    if (text.trim().isEmpty) return true;
    if (invalidKeywords.any((k) => text.contains(k))) return true;
    if (sizeRe.hasMatch(text)) return true;
    return false;
  }

  for (var i = 0; i + 1 < entities.length; i++) {
    final t1 = entities[i].text;
    final t2 = entities[i + 1].text;

    if (isInvalid(t1)) continue;

    final sizeMatch = sizeRe.firstMatch(t2);
    if (sizeMatch == null) continue;
    final sizeText = sizeMatch.group(0) ?? '';

    var heightText = '';
    final heightMatch = heightRe.firstMatch(t2);
    if (heightMatch != null) {
      heightText = heightMatch.group(0) ?? '';
    }

    final idMatch = idRe.firstMatch(t1);
    if (idMatch != null) {
      final valveId = idMatch.group(0) ?? '';
      final valveName = t1
          .replaceAll(valveId, '')
          .replaceAll('（', '')
          .replaceAll('）', '')
          .replaceAll('(', '')
          .replaceAll(')', '')
          .trim();
      results.add({
        'kind': '阀门',
        'name': valveName,
        'code': valveId,
        'size': sizeText,
        'height': heightText,
      });
      continue;
    }

    if (t1.contains('风口') || t1.contains('风阀') || t1.contains('百叶')) {
      results.add({
        'kind': '风口',
        'name': t1,
        'code': '',
        'size': sizeText,
        'height': heightText,
      });
    }
  }

  return results;
}
