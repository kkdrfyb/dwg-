import 'dart:io';

import 'package:crypto/crypto.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';
import 'package:sqlite3/sqlite3.dart' as sqlite;

import '../../core/logging/log_service.dart';
import '../../core/task/task_exceptions.dart';
import '../../core/task/task_service.dart';
import 'dxf_isolate.dart';
import 'dxf_models.dart';

class DxfPreparedFile {
  const DxfPreparedFile({
    required this.sourcePath,
    required this.sourceName,
    required this.dxfPath,
    required this.dxfSize,
    required this.needsIndex,
  });

  final String sourcePath;
  final String sourceName;
  final String dxfPath;
  final int dxfSize;
  final bool needsIndex;
}

class DxfCacheService {
  DxfCacheService({required LogService log}) : _log = log;

  final LogService _log;
  sqlite.Database? _db;
  String? _dbPath;

  String? get activeDbPath => _dbPath;

  Future<sqlite.Database> _openDbForSources(List<PlatformFile> sources) async {
    final resolvedPath = await _resolveDbPathForSources(sources);
    if (_db != null && _dbPath == resolvedPath) {
      return _db!;
    }
    _db?.dispose();
    final parentDir = Directory(p.dirname(resolvedPath));
    if (!parentDir.existsSync()) {
      parentDir.createSync(recursive: true);
    }
    final db = sqlite.sqlite3.open(resolvedPath);
    _initSchema(db);
    _db = db;
    _dbPath = resolvedPath;
    return db;
  }

  Future<sqlite.Database> _openCurrentDb() async {
    if (_db != null) {
      return _db!;
    }
    return _openDbForSources(const <PlatformFile>[]);
  }

  Future<String> _resolveDbPathForSources(List<PlatformFile> sources) async {
    final sourcePaths = sources
        .map((file) => file.path)
        .whereType<String>()
        .toList();
    if (sourcePaths.isEmpty) {
      final dir = await getApplicationSupportDirectory();
      return p.join(dir.path, 'office_toolbox_dxf.db');
    }

    final parentDirs =
        sourcePaths.map((path) => p.dirname(path)).toSet().toList()
          ..sort((a, b) => a.compareTo(b));
    final commonRoot = _findCommonPath(parentDirs);
    if (commonRoot != null) {
      final outputDir = p.join(commonRoot, 'output');
      final outputFolder = Directory(outputDir);
      if (!outputFolder.existsSync()) {
        outputFolder.createSync(recursive: true);
      }
      await _hideOutputDir(outputDir);
      return p.join(outputDir, 'dxf_index.sqlite');
    }

    final dir = await getApplicationSupportDirectory();
    final key = md5
        .convert(
          parentDirs.map((item) => item.toLowerCase()).join('|').codeUnits,
        )
        .toString()
        .substring(0, 12);
    return p.join(dir.path, 'office_toolbox_dxf_$key.db');
  }

  String? _findCommonPath(List<String> paths) {
    if (paths.isEmpty) {
      return null;
    }
    if (paths.length == 1) {
      return paths.first;
    }

    final parts = paths
        .map((path) => p.normalize(path).split(RegExp(r'[\\/]')))
        .toList();
    final minLength = parts
        .map((item) => item.length)
        .reduce((a, b) => a < b ? a : b);
    final common = <String>[];
    for (var i = 0; i < minLength; i++) {
      final token = parts.first[i].toLowerCase();
      final same = parts.every((item) => item[i].toLowerCase() == token);
      if (!same) {
        break;
      }
      common.add(parts.first[i]);
    }

    if (common.length <= 1) {
      return null;
    }
    return p.joinAll(common);
  }

  void _initSchema(sqlite.Database db) {
    db.execute('''
CREATE TABLE IF NOT EXISTS file_meta (
  source_path TEXT PRIMARY KEY,
  source_name TEXT,
  source_size INTEGER,
  source_mtime INTEGER,
  source_md5 TEXT,
  dxf_path TEXT,
  dxf_size INTEGER,
  dxf_mtime INTEGER,
  dxf_md5 TEXT,
  indexed_at INTEGER
);
''');
    db.execute('''
CREATE TABLE IF NOT EXISTS dxf_text (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  source_path TEXT,
  file_name TEXT,
  object_type TEXT,
  layer TEXT,
  content TEXT,
  content_lower TEXT
);
''');
    db.execute(
      'CREATE INDEX IF NOT EXISTS idx_dxf_text_source ON dxf_text(source_path);',
    );
    db.execute(
      'CREATE INDEX IF NOT EXISTS idx_dxf_text_content ON dxf_text(content_lower);',
    );
  }

  Future<void> close() async {
    _db?.dispose();
    _db = null;
    _dbPath = null;
  }

  @visibleForTesting
  String buildOdaProcessCommandForTest(
    String exe,
    List<String> args, {
    String? expectedOutputPath,
    int timeoutSeconds = 90,
  }) {
    return _buildOdaProcessCommand(
      exe,
      args,
      expectedOutputPath,
      timeoutSeconds,
    );
  }

  @visibleForTesting
  String toLegacyVersionForTest(String outputVersion) {
    return _toLegacyVersion(outputVersion);
  }

  @visibleForTesting
  Future<void> waitForFileReadyForTest(String path) {
    return _waitForFileReady(path);
  }

  Future<List<DxfPreparedFile>> prepareSources(
    List<PlatformFile> sources,
    TaskContext context, {
    void Function(double progress, String message)? onProgress,
    bool pruneMissing = false,
  }) async {
    final validSources = sources.where((file) => file.path != null).toList();
    final db = await _openDbForSources(validSources);
    if (pruneMissing) {
      final sourcePaths = validSources.map((file) => file.path!).toList();
      await _pruneMissingSources(db, sourcePaths);
      await _syncOutputFolders(validSources);
    }

    final prepared = <DxfPreparedFile>[];
    for (var i = 0; i < validSources.length; i++) {
      if (context.isCanceled()) throw TaskCanceled();
      final file = validSources[i];
      final progress = validSources.isEmpty ? 1.0 : i / validSources.length;
      onProgress?.call(progress, '检查文件: ${file.name}');
      context.updateProgress(progress, message: '检查文件: ${file.name}');

      final sourcePath = file.path!;
      final sourceName = file.name;
      final sourceFile = File(sourcePath);
      if (!sourceFile.existsSync()) {
        await _log.warn('源文件不存在: $sourcePath', context: 'dxf');
        continue;
      }

      final ext = p.extension(sourcePath).toLowerCase();
      final sourceStat = sourceFile.statSync();
      final sourceSize = sourceStat.size;
      final sourceMtime = sourceStat.modified.millisecondsSinceEpoch;

      final meta = _getMeta(db, sourcePath);
      final outputInfo = await _ensureDxf(
        db,
        sourcePath: sourcePath,
        sourceName: sourceName,
        ext: ext,
        sourceSize: sourceSize,
        sourceMtime: sourceMtime,
        meta: meta,
        context: context,
        baseProgress: progress,
      );

      if (outputInfo == null) {
        continue;
      }

      prepared.add(outputInfo);
    }
    onProgress?.call(1, '文件准备完成');
    context.updateProgress(1, message: '文件准备完成');
    return prepared;
  }

  Future<void> ensureIndex(
    List<PlatformFile> sources,
    TaskContext context, {
    void Function(double progress, String message)? onProgress,
  }) async {
    final prepared = await prepareSources(
      sources,
      context,
      onProgress: onProgress,
      pruneMissing: true,
    );
    await _indexPrepared(prepared, context, onProgress: onProgress);
  }

  Future<List<DxfSearchResult>> queryKeywords(List<String> keywords) async {
    final db = await _openCurrentDb();
    if (keywords.isEmpty) {
      final rows = db.select(
        'SELECT file_name, object_type, layer, content FROM dxf_text',
      );
      return rows
          .map(
            (row) => DxfSearchResult(
              fileName: row['file_name'] as String? ?? '',
              objectType: row['object_type'] as String? ?? '',
              layer: row['layer'] as String? ?? '',
              keyword: '全部',
              content: row['content'] as String? ?? '',
            ),
          )
          .toList();
    }

    final lowerKeywords = keywords
        .map((k) => k.toLowerCase())
        .where((k) => k.isNotEmpty)
        .toList();
    if (lowerKeywords.isEmpty) return [];

    final clauses = List.filled(
      lowerKeywords.length,
      'content_lower LIKE ?',
    ).join(' OR ');
    final params = lowerKeywords.map((k) => '%$k%').toList();
    final rows = db.select(
      'SELECT file_name, object_type, layer, content, content_lower FROM dxf_text WHERE $clauses',
      params,
    );

    final results = <DxfSearchResult>[];
    for (final row in rows) {
      final content = row['content'] as String? ?? '';
      final lower = row['content_lower'] as String? ?? content.toLowerCase();
      for (var i = 0; i < lowerKeywords.length; i++) {
        if (lower.contains(lowerKeywords[i])) {
          results.add(
            DxfSearchResult(
              fileName: row['file_name'] as String? ?? '',
              objectType: row['object_type'] as String? ?? '',
              layer: row['layer'] as String? ?? '',
              keyword: keywords[i],
              content: content,
            ),
          );
        }
      }
    }
    return results;
  }

  Future<void> _pruneMissingSources(
    sqlite.Database db,
    List<String> sourcePaths,
  ) async {
    List<sqlite.Row> removedRows;
    if (sourcePaths.isEmpty) {
      removedRows = db.select('SELECT source_path, dxf_path FROM file_meta');
      db.execute('DELETE FROM dxf_text');
      db.execute('DELETE FROM file_meta');
    } else {
      final placeholders = List.filled(sourcePaths.length, '?').join(',');
      removedRows = db.select(
        'SELECT source_path, dxf_path FROM file_meta WHERE source_path NOT IN ($placeholders)',
        sourcePaths,
      );
      db.execute(
        'DELETE FROM dxf_text WHERE source_path NOT IN ($placeholders)',
        sourcePaths,
      );
      db.execute(
        'DELETE FROM file_meta WHERE source_path NOT IN ($placeholders)',
        sourcePaths,
      );
    }

    for (final row in removedRows) {
      final sourcePath = row['source_path'] as String? ?? '';
      final dxfPath = row['dxf_path'] as String? ?? '';
      if (!_isGeneratedOutputDxf(sourcePath, dxfPath)) {
        continue;
      }
      final file = File(dxfPath);
      if (!file.existsSync()) {
        continue;
      }
      try {
        file.deleteSync();
        await _log.info('删除失效 DXF 缓存: $dxfPath', context: 'dxf');
      } catch (error) {
        await _log.warn(
          '删除失效 DXF 缓存失败: $dxfPath',
          context: 'dxf',
          error: error,
        );
      }
    }
  }

  bool _isGeneratedOutputDxf(String sourcePath, String dxfPath) {
    if (sourcePath.isEmpty || dxfPath.isEmpty) {
      return false;
    }
    if (sourcePath == dxfPath) {
      return false;
    }
    if (p.extension(sourcePath).toLowerCase() != '.dwg') {
      return false;
    }
    if (p.extension(dxfPath).toLowerCase() != '.dxf') {
      return false;
    }
    return p.basename(p.dirname(dxfPath)).toLowerCase() == 'output';
  }

  Future<void> _syncOutputFolders(List<PlatformFile> sources) async {
    final expectedBySourceDir = <String, Set<String>>{};
    for (final source in sources) {
      final sourcePath = source.path;
      if (sourcePath == null) {
        continue;
      }
      if (p.extension(sourcePath).toLowerCase() != '.dwg') {
        continue;
      }
      final sourceDir = p.dirname(sourcePath);
      expectedBySourceDir
          .putIfAbsent(sourceDir, () => <String>{})
          .add(p.basenameWithoutExtension(sourcePath).toLowerCase());
    }

    for (final entry in expectedBySourceDir.entries) {
      final outputDir = Directory(p.join(entry.key, 'output'));
      if (!outputDir.existsSync()) {
        continue;
      }
      for (final entity in outputDir.listSync()) {
        if (entity is! File) {
          continue;
        }
        final dxfPath = entity.path;
        if (p.extension(dxfPath).toLowerCase() != '.dxf') {
          continue;
        }
        final baseName = p.basenameWithoutExtension(dxfPath).toLowerCase();
        if (entry.value.contains(baseName)) {
          continue;
        }
        try {
          entity.deleteSync();
          await _log.info('清理冗余 DXF 缓存: $dxfPath', context: 'dxf');
        } catch (error) {
          await _log.warn(
            '清理冗余 DXF 缓存失败: $dxfPath',
            context: 'dxf',
            error: error,
          );
        }
      }
    }
  }

  Map<String, Object?>? _getMeta(sqlite.Database db, String sourcePath) {
    final rows = db.select(
      'SELECT source_path, source_name, source_size, source_mtime, source_md5, '
      'dxf_path, dxf_size, dxf_mtime, dxf_md5, indexed_at '
      'FROM file_meta WHERE source_path = ?',
      [sourcePath],
    );
    if (rows.isEmpty) return null;
    return rows.first;
  }

  Future<DxfPreparedFile?> _ensureDxf(
    sqlite.Database db, {
    required String sourcePath,
    required String sourceName,
    required String ext,
    required int sourceSize,
    required int sourceMtime,
    required Map<String, Object?>? meta,
    required TaskContext context,
    required double baseProgress,
  }) async {
    final isDxf = ext == '.dxf';
    final isDwg = ext == '.dwg';
    if (context.isCanceled()) throw TaskCanceled();
    if (!isDxf && !isDwg) {
      await _log.warn('不支持的格式: $sourcePath', context: 'dxf');
      return null;
    }

    final metaSourceMd5 = meta?['source_md5'] as String?;
    var sourceMd5 = metaSourceMd5;
    final metaSize = meta?['source_size'] as int?;
    final metaMtime = meta?['source_mtime'] as int?;
    final metaIndexedAt = meta?['indexed_at'] as int?;

    var needsHash =
        meta == null || metaSize != sourceSize || metaMtime != sourceMtime;
    if (needsHash) {
      sourceMd5 = await _computeFileMd5(sourcePath);
    }

    String dxfPath;
    if (isDxf) {
      dxfPath = sourcePath;
    } else {
      final outputDir = p.join(p.dirname(sourcePath), 'output');
      final outputFolder = Directory(outputDir);
      if (!outputFolder.existsSync()) {
        outputFolder.createSync(recursive: true);
      }
      await _hideOutputDir(outputDir);
      dxfPath = p.join(
        outputDir,
        '${p.basenameWithoutExtension(sourcePath)}.dxf',
      );
    }

    var dxfFile = File(dxfPath);
    if (!dxfFile.existsSync() && isDwg) {
      final alt = _findDxfByBasename(
        p.dirname(dxfPath),
        p.basenameWithoutExtension(sourcePath),
      );
      if (alt != null) {
        dxfPath = alt;
        dxfFile = File(dxfPath);
      }
    }

    var needsConvert = false;
    if (isDwg) {
      final metaDxfPath = meta?['dxf_path'] as String?;
      if (meta != null &&
          metaSourceMd5 != null &&
          metaSourceMd5 == sourceMd5 &&
          metaDxfPath != null) {
        final metaDxfFile = File(metaDxfPath);
        if (metaDxfFile.existsSync()) {
          dxfPath = metaDxfPath;
          dxfFile = metaDxfFile;
        } else if (!dxfFile.existsSync()) {
          needsConvert = true;
        }
      } else {
        if (!dxfFile.existsSync()) {
          needsConvert = true;
        } else if (meta == null) {
          needsConvert = false;
        } else if (metaSourceMd5 != null && metaSourceMd5 != sourceMd5) {
          needsConvert = true;
        }
      }

      if (needsConvert) {
        context.updateProgress(baseProgress, message: '转换 DWG: $sourceName');
        await _log.info('开始转换 DWG -> DXF: $sourcePath', context: 'dxf');
        await _convertDwgToDxf(sourcePath, p.dirname(dxfPath));
        await _log.info('完成转换 DWG -> DXF: $sourcePath', context: 'dxf');
        dxfFile = File(dxfPath);
        if (!dxfFile.existsSync()) {
          final alt = _findDxfByBasename(
            p.dirname(dxfPath),
            p.basenameWithoutExtension(sourcePath),
          );
          if (alt != null) {
            dxfPath = alt;
            dxfFile = File(dxfPath);
          }
        }
        if (!dxfFile.existsSync()) {
          await _log.error('DWG 转 DXF 失败: $sourcePath', context: 'dxf');
          return null;
        }
      }
    }

    if (!dxfFile.existsSync()) {
      await _log.error('DXF 文件不存在: $dxfPath', context: 'dxf');
      return null;
    }
    await _waitForFileReady(dxfPath);

    final dxfStat = dxfFile.statSync();
    final dxfSize = dxfStat.size;
    final dxfMtime = dxfStat.modified.millisecondsSinceEpoch;
    var dxfMd5 = meta?['dxf_md5'] as String?;
    final metaDxfSize = meta?['dxf_size'] as int?;
    final metaDxfMtime = meta?['dxf_mtime'] as int?;
    final metaDxfMd5 = meta?['dxf_md5'] as String?;
    if (meta == null || metaDxfSize != dxfSize || metaDxfMtime != dxfMtime) {
      await _waitForFileReady(dxfPath);
      dxfMd5 = await _computeFileMd5(dxfPath);
    }

    final sourceContentChanged =
        meta == null ||
        metaSourceMd5 == null ||
        sourceMd5 == null ||
        metaSourceMd5 != sourceMd5;
    final dxfContentChanged =
        meta == null ||
        metaDxfMd5 == null ||
        dxfMd5 == null ||
        metaDxfMd5 != dxfMd5;
    final needsIndex =
        meta == null ||
        metaIndexedAt == null ||
        sourceContentChanged ||
        dxfContentChanged ||
        needsConvert;

    _upsertMeta(
      db,
      sourcePath: sourcePath,
      sourceName: sourceName,
      sourceSize: sourceSize,
      sourceMtime: sourceMtime,
      sourceMd5: sourceMd5,
      dxfPath: dxfPath,
      dxfSize: dxfSize,
      dxfMtime: dxfMtime,
      dxfMd5: dxfMd5,
      indexedAt: metaIndexedAt,
    );

    return DxfPreparedFile(
      sourcePath: sourcePath,
      sourceName: sourceName,
      dxfPath: dxfPath,
      dxfSize: dxfSize,
      needsIndex: needsIndex,
    );
  }

  void _upsertMeta(
    sqlite.Database db, {
    required String sourcePath,
    required String sourceName,
    required int sourceSize,
    required int sourceMtime,
    required String? sourceMd5,
    required String dxfPath,
    required int dxfSize,
    required int dxfMtime,
    required String? dxfMd5,
    required int? indexedAt,
  }) {
    db.execute(
      'INSERT INTO file_meta '
      '(source_path, source_name, source_size, source_mtime, source_md5, '
      'dxf_path, dxf_size, dxf_mtime, dxf_md5, indexed_at) '
      'VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) '
      'ON CONFLICT(source_path) DO UPDATE SET '
      'source_name=excluded.source_name, '
      'source_size=excluded.source_size, '
      'source_mtime=excluded.source_mtime, '
      'source_md5=excluded.source_md5, '
      'dxf_path=excluded.dxf_path, '
      'dxf_size=excluded.dxf_size, '
      'dxf_mtime=excluded.dxf_mtime, '
      'dxf_md5=excluded.dxf_md5, '
      'indexed_at=excluded.indexed_at',
      [
        sourcePath,
        sourceName,
        sourceSize,
        sourceMtime,
        sourceMd5,
        dxfPath,
        dxfSize,
        dxfMtime,
        dxfMd5,
        indexedAt,
      ],
    );
  }

  Future<void> _indexPrepared(
    List<DxfPreparedFile> prepared,
    TaskContext context, {
    void Function(double progress, String message)? onProgress,
  }) async {
    if (prepared.isEmpty) return;
    final db = await _openCurrentDb();
    final needsIndex = prepared.where((file) => file.needsIndex).toList();
    if (needsIndex.isEmpty) return;

    for (var i = 0; i < needsIndex.length; i++) {
      if (context.isCanceled()) throw TaskCanceled();
      final file = needsIndex[i];
      final progress = needsIndex.isEmpty ? 1.0 : i / needsIndex.length;
      onProgress?.call(progress, '解析索引: ${file.sourceName}');
      context.updateProgress(progress, message: '解析索引: ${file.sourceName}');

      final response = await compute(scanDxfFileForKeywords, {
        'path': file.dxfPath,
        'name': file.sourceName,
        'keywords': const <String>[],
        'size': file.dxfSize,
        'forceParse': true,
      });

      if (response['ok'] != true && response['error'] != null) {
        await _log.warn(
          '解析失败: ${file.sourceName}',
          context: 'dxf',
          error: response['error'],
        );
      }

      final items = (response['results'] as List)
          .map(
            (item) => {
              'fileName': item['fileName'] as String? ?? file.sourceName,
              'objectType': item['objectType'] as String? ?? '',
              'layer': item['layer'] as String? ?? '',
              'content': item['content'] as String? ?? '',
            },
          )
          .toList();

      db.execute('BEGIN');
      try {
        db.execute('DELETE FROM dxf_text WHERE source_path = ?', [
          file.sourcePath,
        ]);
        final stmt = db.prepare(
          'INSERT INTO dxf_text '
          '(source_path, file_name, object_type, layer, content, content_lower) '
          'VALUES (?, ?, ?, ?, ?, ?)',
        );
        for (final item in items) {
          final content = item['content'] ?? '';
          stmt.execute([
            file.sourcePath,
            item['fileName'] ?? '',
            item['objectType'] ?? '',
            item['layer'] ?? '',
            content,
            content.toLowerCase(),
          ]);
        }
        stmt.dispose();

        final now = DateTime.now().millisecondsSinceEpoch;
        db.execute(
          'UPDATE file_meta SET indexed_at = ? WHERE source_path = ?',
          [now, file.sourcePath],
        );
        db.execute('COMMIT');
      } catch (error) {
        db.execute('ROLLBACK');
        await _log.error(
          '写入索引失败: ${file.sourceName}',
          context: 'dxf',
          error: error,
        );
      }
    }
    onProgress?.call(1, '索引更新完成');
    context.updateProgress(1, message: '索引更新完成');
  }

  Future<String> _computeFileMd5(String path) async {
    return compute(_md5Worker, path);
  }

  Future<String> convertDxfToDwg({
    required String dxfPath,
    required String outputDwgPath,
  }) async {
    final sourceFile = File(dxfPath);
    if (!sourceFile.existsSync()) {
      throw Exception('DXF 文件不存在: $dxfPath');
    }

    final normalizedOutput = p.normalize(outputDwgPath);
    final outputDir = p.dirname(normalizedOutput);
    final outputFolder = Directory(outputDir);
    if (!outputFolder.existsSync()) {
      outputFolder.createSync(recursive: true);
    }

    final generatedPath = p.join(
      outputDir,
      '${p.basenameWithoutExtension(dxfPath)}.dwg',
    );
    await _convertWithOda(
      inputPath: dxfPath,
      outputDir: outputDir,
      inputType: 'DXF',
      outputType: 'DWG',
      outputVersion: 'DWG2010',
    );

    final generatedFile = File(generatedPath);
    if (!generatedFile.existsSync()) {
      throw Exception('DXF 转 DWG 后未找到输出文件: $generatedPath');
    }

    final generatedNorm = p.normalize(generatedPath).toLowerCase();
    final targetNorm = normalizedOutput.toLowerCase();
    if (generatedNorm == targetNorm) {
      return normalizedOutput;
    }

    final targetFile = File(normalizedOutput);
    if (targetFile.existsSync()) {
      targetFile.deleteSync();
    }

    try {
      generatedFile.renameSync(normalizedOutput);
    } catch (_) {
      generatedFile.copySync(normalizedOutput);
      generatedFile.deleteSync();
    }
    return normalizedOutput;
  }

  Future<void> _convertDwgToDxf(String dwgPath, String outputDir) async {
    await _convertWithOda(
      inputPath: dwgPath,
      outputDir: outputDir,
      inputType: 'DWG',
      outputType: 'DXF',
      outputVersion: 'DXF2010',
    );
  }

  Future<void> _convertWithOda({
    required String inputPath,
    required String outputDir,
    required String inputType,
    required String outputType,
    required String outputVersion,
  }) async {
    final exe = _resolveOdaExecutable();
    if (exe == null) {
      throw Exception('未找到 ODAFileConverter，请检查 ODAFileConverter 文件夹');
    }
    final outputFolder = Directory(outputDir);
    if (!outputFolder.existsSync()) {
      outputFolder.createSync(recursive: true);
    }
    await _hideOutputDir(outputDir);

    final normalizedOutputType = outputType.toUpperCase();
    final normalizedOutputVersion = outputVersion.toUpperCase();
    final legacyVersion = _toLegacyVersion(normalizedOutputVersion);
    final expectedExt = normalizedOutputType.toLowerCase();
    final expectedBase = p.basenameWithoutExtension(inputPath);
    final expectedPath = p.join(outputDir, '$expectedBase.$expectedExt');

    final attempts = <({String label, List<String> args})>[
      (
        label: 'legacy',
        args: [
          p.dirname(inputPath),
          outputDir,
          legacyVersion,
          normalizedOutputType,
          '0',
          '0',
          p.basename(inputPath),
        ],
      ),
    ];

    final odaTimeout = _odaTimeoutForInput(inputPath);
    final errors = <String>[];
    for (final attempt in attempts) {
      final result = await _runOda(
        exe,
        attempt.args,
        expectedOutputPath: expectedPath,
        timeout: odaTimeout,
      );
      if (result.exitCode != 0) {
        errors.add(
          '${attempt.label}: ${_formatOdaError(result)} | args=${attempt.args.join(' | ')}',
        );
        continue;
      }

      final expectedFile = File(expectedPath);
      if (expectedFile.existsSync()) {
        return;
      }
      final altPath = _findOutputByBasename(
        outputDir,
        expectedBase,
        expectedExt,
      );
      if (altPath != null) {
        return;
      }

      errors.add(
        '${attempt.label}: exit=0 但未找到输出文件 $expectedPath | args=${attempt.args.join(' | ')}',
      );
    }

    await _log.error(
      'ODA 转换失败: $inputPath',
      context: 'dxf',
      error: '可执行文件: $exe; ${errors.join(' || ')}',
    );
    throw Exception('ODA 转换失败');
  }

  String _toLegacyVersion(String outputVersion) {
    final upper = outputVersion.toUpperCase();
    if (upper.startsWith('ACAD')) {
      return upper;
    }
    final match = RegExp(r'(DWG|DXF)(\d{4})').firstMatch(upper);
    if (match != null) {
      return 'ACAD${match.group(2)}';
    }
    return 'ACAD2010';
  }

  String _formatOdaError(ProcessResult result) {
    final stderr = '${result.stderr}'.trim();
    if (stderr.isNotEmpty) {
      return 'exit=${result.exitCode}, stderr=$stderr';
    }
    final stdout = '${result.stdout}'.trim();
    if (stdout.isNotEmpty) {
      return 'exit=${result.exitCode}, stdout=$stdout';
    }
    return 'exit=${result.exitCode}';
  }

  Duration _odaTimeoutForInput(String inputPath) {
    try {
      final bytes = File(inputPath).lengthSync();
      final mb = bytes / (1024 * 1024);
      final seconds = (60 + mb * 3).round().clamp(90, 900);
      return Duration(seconds: seconds);
    } catch (_) {
      return const Duration(minutes: 3);
    }
  }

  Future<ProcessResult> _runOda(
    String exe,
    List<String> args, {
    String? expectedOutputPath,
    Duration timeout = const Duration(minutes: 3),
  }) async {
    if (!Platform.isWindows) {
      return Process.run(exe, args, runInShell: false);
    }
    final command = _buildOdaProcessCommand(
      exe,
      args,
      expectedOutputPath,
      timeout.inSeconds,
    );
    return Process.run('powershell', [
      '-NoProfile',
      '-NonInteractive',
      '-Command',
      command,
    ], runInShell: false);
  }

  String _buildOdaProcessCommand(
    String exe,
    List<String> args,
    String? expectedOutputPath,
    int timeoutSeconds,
  ) {
    String escapePsSingleQuoted(String value) => value.replaceAll("'", "''");
    String quoteWinArg(String value) {
      final escaped = value.replaceAll('"', r'\"');
      return '"$escaped"';
    }

    final argLine = args.map(quoteWinArg).join(' ');
    final filePath = "'${escapePsSingleQuoted(exe)}'";
    final quotedArgLine = "'${escapePsSingleQuoted(argLine)}'";
    final timeout = timeoutSeconds < 5 ? 5 : timeoutSeconds;
    final timeoutOutputCheck =
        expectedOutputPath == null
        ? ''
        : '''
if (Test-Path -LiteralPath '${escapePsSingleQuoted(expectedOutputPath)}') {
  Write-Error "ODA process timeout (output exists but process is still running)"
  exit 124
}
''';
    return '''
\$p = Start-Process -FilePath $filePath -ArgumentList $quotedArgLine -WindowStyle Hidden -PassThru
\$deadline = (Get-Date).AddSeconds($timeout)
while ((Get-Date) -lt \$deadline) {
  \$p.Refresh()
  if (\$p.HasExited) {
    exit \$p.ExitCode
  }
  Start-Sleep -Milliseconds 300
}
if (-not \$p.HasExited) {
  Stop-Process -Id \$p.Id -Force -ErrorAction SilentlyContinue
}
$timeoutOutputCheck
Write-Error "ODA process timeout"
exit 124
''';
  }

  String? _resolveOdaExecutable() {
    final exeName = Platform.isWindows
        ? 'ODAFileConverter.exe'
        : 'odafileconverter';
    final envPath = Platform.environment['ODA_FILE_CONVERTER'];
    if (envPath != null && envPath.isNotEmpty) {
      final envFile = File(envPath);
      if (envFile.existsSync()) return envPath;
      final envDir = Directory(envPath);
      if (envDir.existsSync()) {
        final candidate = p.join(envDir.path, exeName);
        if (File(candidate).existsSync()) return candidate;
      }
    }

    final startDirs = <String>{
      File(Platform.resolvedExecutable).parent.path,
      Directory.current.path,
    };

    for (final startDir in startDirs) {
      final resolved = _searchOdaInAncestors(startDir, exeName);
      if (resolved != null) return resolved;
    }

    final pathEnv = Platform.environment['PATH'];
    if (pathEnv != null && pathEnv.isNotEmpty) {
      for (final dir in pathEnv.split(';')) {
        final trimmed = dir.trim();
        if (trimmed.isEmpty) continue;
        final candidate = p.join(trimmed, exeName);
        if (File(candidate).existsSync()) return candidate;
      }
    }

    return null;
  }

  String? _searchOdaInAncestors(String startDir, String exeName) {
    var current = Directory(startDir);
    while (true) {
      final direct = p.join(current.path, exeName);
      if (File(direct).existsSync()) return direct;
      final bundled = p.join(current.path, 'ODAFileConverter', exeName);
      if (File(bundled).existsSync()) return bundled;

      final parent = current.parent;
      if (parent.path == current.path) break;
      current = parent;
    }
    return null;
  }

  Future<void> _hideOutputDir(String outputDir) async {
    if (!Platform.isWindows) return;
    try {
      await Process.run('attrib', ['+h', outputDir], runInShell: true);
    } catch (_) {
      // ignore
    }
  }

  Future<void> _waitForFileReady(String path) async {
    final file = File(path);
    if (!file.existsSync()) {
      return;
    }

    final deadline = DateTime.now().add(const Duration(seconds: 12));
    int? lastSize;
    while (true) {
      if (!file.existsSync()) {
        return;
      }
      try {
        final raf = file.openSync(mode: FileMode.read);
        final size = raf.lengthSync();
        raf.closeSync();
        if (size > 0 && lastSize != null && size == lastSize) {
          return;
        }
        lastSize = size;
      } on FileSystemException {
        // DXF may still be locked by converter; retry briefly.
      }

      if (DateTime.now().isAfter(deadline)) {
        return;
      }
      await Future.delayed(const Duration(milliseconds: 250));
    }
  }

  String? _findDxfByBasename(String dir, String baseName) {
    return _findOutputByBasename(dir, baseName, 'dxf');
  }

  String? _findOutputByBasename(String dir, String baseName, String ext) {
    final folder = Directory(dir);
    if (!folder.existsSync()) return null;
    final lowerBase = baseName.toLowerCase();
    final expectedExt = '.${ext.toLowerCase()}';
    for (final entity in folder.listSync()) {
      if (entity is! File) continue;
      final name = p.basenameWithoutExtension(entity.path).toLowerCase();
      if (name == lowerBase &&
          p.extension(entity.path).toLowerCase() == expectedExt) {
        return entity.path;
      }
    }
    return null;
  }
}

String _md5Worker(String path) {
  final bytes = File(path).readAsBytesSync();
  return md5.convert(bytes).toString();
}
