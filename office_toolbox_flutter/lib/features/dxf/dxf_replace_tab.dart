import 'dart:io';
import 'dart:math';

import 'package:desktop_drop/desktop_drop.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:path/path.dart' as p;
import 'package:provider/provider.dart';

import '../../core/export/export_utils.dart';
import '../../core/logging/log_service.dart';
import '../../core/task/task_exceptions.dart';
import '../../core/task/task_limiter.dart';
import '../../core/task/task_service.dart';
import '../../widgets/section_card.dart';
import 'dxf_cache_service.dart';
import 'dxf_file_helper.dart';
import 'dxf_isolate.dart';
import 'dxf_models.dart';
import 'dxf_utils.dart';

class _ReplacePayload {
  const _ReplacePayload({required this.fileName, required this.response});

  final String fileName;
  final Map<String, dynamic> response;
}

class _ApplySummary {
  const _ApplySummary({
    required this.written,
    required this.skipped,
    required this.failed,
  });

  final int written;
  final int skipped;
  final int failed;
}

class DxfReplaceTab extends StatefulWidget {
  const DxfReplaceTab({super.key});

  @override
  State<DxfReplaceTab> createState() => _DxfReplaceTabState();
}

class _DxfReplaceTabState extends State<DxfReplaceTab> {
  final ScrollController _fileScrollController = ScrollController();
  final List<PlatformFile> _files = [];
  final List<_PairRow> _pairRows = [_PairRow()];
  final List<DxfReplaceResult> _results = [];
  final Map<String, DxfPreparedFile> _preparedByDxfPath = {};
  final List<DxfReplacePair> _scanPairs = [];

  bool _isRunning = false;
  bool _isDragging = false;
  double _progress = 0;
  String _status = '';
  String? _activeTaskId;

  bool _overwrite = false;
  String? _outputDir;

  String _filterFile = '';
  String _filterType = '';
  String _filterLayer = '';
  String _filterRule = '';
  String _filterOriginal = '';
  String _filterUpdated = '';

  int _rowsPerPage = PaginatedDataTable.defaultRowsPerPage;

  @override
  void dispose() {
    for (final row in _pairRows) {
      row.dispose();
    }
    _fileScrollController.dispose();
    super.dispose();
  }

  Future<void> _pickFiles() async {
    final result = await FilePicker.platform.pickFiles(
      allowMultiple: true,
      type: FileType.custom,
      allowedExtensions: const ['dwg', 'dxf'],
    );
    if (result == null) return;

    final files = result.files.where((file) => file.path != null).toList();
    if (files.isEmpty) return;
    _addFiles(files);
  }

  Future<void> _pickFolder() async {
    final path = await FilePicker.platform.getDirectoryPath();
    if (path == null) return;

    final files = await collectDxfFilesFromPaths([path]);
    if (!mounted) return;
    if (files.isEmpty) {
      _showSnack('文件夹内未找到 DWG/DXF 文件');
      return;
    }
    _addFiles(files);
  }

  Future<void> _handleDrop(List<String> paths) async {
    if (paths.isEmpty) return;
    final files = await collectDxfFilesFromPaths(paths);
    if (!mounted) return;
    if (files.isEmpty) {
      _showSnack('未识别到 DWG/DXF 文件');
      return;
    }
    _addFiles(files);
  }

  void _addFiles(List<PlatformFile> files) {
    final existing = _files.map((f) => f.path).whereType<String>().toSet();
    final added = <PlatformFile>[];
    for (final file in files) {
      final path = file.path;
      if (path == null) continue;
      if (existing.add(path)) {
        added.add(file);
      }
    }
    if (added.isEmpty) {
      _showSnack('没有新增文件');
      return;
    }
    setState(() {
      _files.addAll(added);
    });
  }

  void _removeFile(PlatformFile file) {
    setState(() => _files.remove(file));
  }

  void _clearFiles() {
    setState(() {
      _files.clear();
      _preparedByDxfPath.clear();
      _scanPairs.clear();
    });
  }

  void _addPair() {
    setState(() => _pairRows.add(_PairRow()));
  }

  void _removePair(int index) {
    if (index < 0 || index >= _pairRows.length) return;
    setState(() {
      _pairRows[index].dispose();
      _pairRows.removeAt(index);
      if (_pairRows.isEmpty) {
        _pairRows.add(_PairRow());
      }
    });
  }

  List<DxfReplacePair> _collectPairs() {
    return _pairRows
        .map(
          (row) => DxfReplacePair(
            find: row.findController.text.trim(),
            replaceWith: row.replaceController.text,
          ),
        )
        .where((pair) => pair.isValid)
        .toList();
  }

  Future<void> _startScan() async {
    if (_files.isEmpty) {
      _showSnack('请先选择 DWG/DXF 文件');
      return;
    }

    final pairs = _collectPairs();
    if (pairs.isEmpty) {
      _showSnack('请至少输入一组替换规则');
      return;
    }
    _scanPairs
      ..clear()
      ..addAll(
        pairs.map(
          (pair) =>
              DxfReplacePair(find: pair.find, replaceWith: pair.replaceWith),
        ),
      );

    final taskService = context.read<TaskService>();
    final log = context.read<LogService>();
    final cache = context.read<DxfCacheService>();
    final handle = taskService.startTask('DXF 文字替换模拟');

    setState(() {
      _isRunning = true;
      _progress = 0;
      _status = '准备处理 ${_files.length} 个文件...';
      _results.clear();
      _preparedByDxfPath.clear();
      _activeTaskId = handle.id;
    });

    var preparedFiles = <DxfPreparedFile>[];
    final payloads = await taskService.runTask<List<_ReplacePayload>>(handle, (
      context,
    ) async {
      final prepared = await cache.prepareSources(
        _files,
        context,
        onProgress: (progress, message) {
          if (!mounted) return;
          setState(() {
            _progress = progress;
            _status = message;
          });
        },
      );
      preparedFiles = prepared;

      if (prepared.isEmpty) return <_ReplacePayload>[];

      final tasks = prepared
          .map(
            (file) => () async {
              try {
                final response = await compute(scanDxfFileForReplace, {
                  'path': file.dxfPath,
                  'name': file.sourceName,
                  'size': file.dxfSize,
                  'pairs': pairs
                      .map(
                        (pair) => {
                          'find': pair.find,
                          'replace': pair.replaceWith,
                        },
                      )
                      .toList(),
                });
                return _ReplacePayload(
                  fileName: file.sourceName,
                  response: response,
                );
              } catch (error) {
                return _ReplacePayload(
                  fileName: file.sourceName,
                  response: {
                    'ok': false,
                    'error': error.toString(),
                    'results': <Map<String, String>>[],
                  },
                );
              }
            },
          )
          .toList();

      final limiter = TaskLimiter.auto();
      return limiter.run<_ReplacePayload>(
        tasks,
        isCanceled: context.isCanceled,
        onItemCompleted: (index, completed, total) {
          final progress = completed / max(1, total);
          context.updateProgress(progress, message: '处理中 ($completed/$total)');
          if (mounted) {
            setState(() {
              _progress = progress;
              _status =
                  '处理中 ($completed/$total): ${prepared[index].sourceName}';
            });
          }
        },
      );
    });

    if (!mounted) return;

    if (payloads == null) {
      setState(() {
        _isRunning = false;
        _status = '已取消';
      });
      return;
    }

    for (final prepared in preparedFiles) {
      _preparedByDxfPath[_normalizePathKey(prepared.dxfPath)] = prepared;
    }

    final results = <DxfReplaceResult>[];
    for (final payload in payloads) {
      final response = payload.response;
      if (response['ok'] != true && response['error'] != null) {
        await log.warn(
          '解析失败: ${payload.fileName}',
          context: 'dxf',
          error: response['error'],
        );
      }

      final items = (response['results'] as List)
          .map(
            (item) => DxfReplaceResult(
              fileName: item['fileName'] as String,
              filePath: item['filePath'] as String,
              objectType: item['objectType'] as String,
              layer: item['layer'] as String,
              originalText: item['originalText'] as String,
              updatedText: item['updatedText'] as String,
              rule: item['rule'] as String,
            ),
          )
          .toList();
      results.addAll(items);
    }

    setState(() {
      _results
        ..clear()
        ..addAll(results);
      _isRunning = false;
      _progress = 1;
      _status = results.isEmpty ? '未发现可替换内容' : '完成，共 ${results.length} 条';
    });
  }

  void _cancel() {
    final taskId = _activeTaskId;
    if (taskId == null) return;
    context.read<TaskService>().cancelTask(taskId);
  }

  Future<void> _pickOutputDir() async {
    final path = await FilePicker.platform.getDirectoryPath();
    if (path == null) return;
    setState(() => _outputDir = path);
  }

  Future<void> _applyReplace() async {
    final activeResults = _results.where((r) => !r.skip).toList();
    if (activeResults.isEmpty) {
      _showSnack('没有可执行的替换项');
      return;
    }

    if (!_overwrite && _outputDir != null) {
      final outputDir = Directory(_outputDir!);
      if (!outputDir.existsSync()) {
        outputDir.createSync(recursive: true);
      }
    }

    final taskService = context.read<TaskService>();
    final log = context.read<LogService>();
    final cache = context.read<DxfCacheService>();
    final handle = taskService.startTask('DXF 确认替换');

    setState(() {
      _isRunning = true;
      _progress = 0;
      _status = '准备写入文件...';
      _activeTaskId = handle.id;
    });

    final byFile = <String, List<DxfReplaceResult>>{};
    final allByFile = <String, List<DxfReplaceResult>>{};
    for (final result in activeResults) {
      byFile.putIfAbsent(result.filePath, () => []).add(result);
    }
    for (final result in _results) {
      allByFile.putIfAbsent(result.filePath, () => []).add(result);
    }

    final summary = await taskService.runTask<_ApplySummary>(handle, (
      context,
    ) async {
      final targets = byFile.keys.toList();
      final usedOutputPaths = <String>{};
      var written = 0;
      var skipped = 0;
      var failed = 0;

      for (var index = 0; index < targets.length; index++) {
        if (context.isCanceled()) throw TaskCanceled();
        final dxfPath = targets[index];

        final fileResults = byFile[dxfPath] ?? [];
        final progress = index / max(1, targets.length);
        context.updateProgress(progress, message: '写入: ${p.basename(dxfPath)}');
        if (mounted) {
          setState(() {
            _progress = progress;
            _status =
                '写入中 (${index + 1}/${targets.length}): ${p.basename(dxfPath)}';
          });
        }

        try {
          final sourceInfo = _preparedByDxfPath[_normalizePathKey(dxfPath)];
          final sourcePath = sourceInfo?.sourcePath ?? dxfPath;
          final sourceExt = p.extension(sourcePath).toLowerCase();
          final isDwgSource = sourceExt == '.dwg';

          final dxfFile = File(dxfPath);
          if (!dxfFile.existsSync()) {
            failed++;
            await log.warn('DXF 文件不存在，跳过写入: $dxfPath', context: 'dxf');
            continue;
          }

          final originalText = await readDxfFileAsync(dxfPath);
          var updatedText = originalText;
          final totalRows = allByFile[dxfPath]?.length ?? fileResults.length;
          final useFastRulePath =
              _scanPairs.isNotEmpty && fileResults.length == totalRows;

          if (useFastRulePath) {
            for (final pair in _scanPairs) {
              if (pair.find.isEmpty) {
                continue;
              }
              final pattern = RegExp(
                RegExp.escape(pair.find),
                caseSensitive: false,
              );
              updatedText = updatedText.replaceAllMapped(
                pattern,
                (_) => pair.replaceWith,
              );
            }
          } else {
            final counts = <String, int>{};
            for (final item in fileResults) {
              final key = '${item.originalText}:::${item.updatedText}';
              counts[key] = (counts[key] ?? 0) + 1;
            }
            for (final entry in counts.entries) {
              final parts = entry.key.split(':::');
              if (parts.length != 2) continue;
              updatedText = replaceLimited(
                updatedText,
                parts[0],
                parts[1],
                entry.value,
              );
            }
          }

          if (updatedText == originalText) {
            skipped++;
            await log.warn('未检测到实际文本变化，跳过: $sourcePath', context: 'dxf');
            continue;
          }

          if (isDwgSource) {
            context.updateProgress(
              progress,
              message: '转换 DWG: ${p.basename(sourcePath)}',
            );
            if (mounted) {
              setState(() {
                _status =
                    '转换 DWG (${index + 1}/${targets.length}): ${p.basename(sourcePath)}';
              });
            }
            if (_overwrite) {
              await dxfFile.writeAsString(updatedText, flush: true);
              await cache.convertDxfToDwg(
                dxfPath: dxfPath,
                outputDwgPath: sourcePath,
              );
              await log.info('已回写 DWG: $sourcePath', context: 'dxf');
            } else {
              final outputDwgPath = _resolveOutputPath(
                sourcePath,
                preferredExtension: '.dwg',
                usedPaths: usedOutputPaths,
              );
              final outputParent = Directory(p.dirname(outputDwgPath));
              if (!outputParent.existsSync()) {
                outputParent.createSync(recursive: true);
              }

              final tempDir = await Directory.systemTemp.createTemp(
                'office_toolbox_replace_',
              );
              final tempDxfPath = p.join(
                tempDir.path,
                '${p.basenameWithoutExtension(sourcePath)}_replace_tmp.dxf',
              );
              try {
                await File(tempDxfPath).writeAsString(updatedText, flush: true);
                await cache.convertDxfToDwg(
                  dxfPath: tempDxfPath,
                  outputDwgPath: outputDwgPath,
                );
              } finally {
                if (tempDir.existsSync()) {
                  tempDir.deleteSync(recursive: true);
                }
              }
              usedOutputPaths.add(_normalizePathKey(outputDwgPath));
              await log.info('已输出 DWG: $outputDwgPath', context: 'dxf');
            }
          } else {
            final outputPath = _resolveOutputPath(
              sourcePath,
              preferredExtension: '.dxf',
              usedPaths: usedOutputPaths,
            );
            final outputFile = File(outputPath);
            final parent = outputFile.parent;
            if (!parent.existsSync()) {
              parent.createSync(recursive: true);
            }
            await outputFile.writeAsString(updatedText, flush: true);
            usedOutputPaths.add(_normalizePathKey(outputPath));
            await log.info('已写入 DXF: $outputPath', context: 'dxf');
          }
          written++;
        } catch (error) {
          failed++;
          await log.error(
            '写入失败: ${p.basename(dxfPath)}',
            context: 'dxf',
            error: error,
          );
        }
      }
      context.updateProgress(1, message: '完成');
      return _ApplySummary(written: written, skipped: skipped, failed: failed);
    });

    if (!mounted) return;

    if (summary == null) {
      setState(() {
        _isRunning = false;
        _status = '替换未完成，请查看日志';
      });
      return;
    }

    setState(() {
      _isRunning = false;
      _progress = 1;
      final tail = summary.failed > 0 || summary.skipped > 0
          ? '（成功 ${summary.written}，跳过 ${summary.skipped}，失败 ${summary.failed}）'
          : '（成功 ${summary.written}）';
      if (_overwrite) {
        _status = '替换完成，已覆盖写入 DWG/DXF $tail';
      } else if (_outputDir != null) {
        _status = '替换完成，已输出到 ${_outputDir} $tail';
      } else {
        _status = '替换完成，已输出到源目录 $tail';
      }
    });
  }

  String _resolveOutputPath(
    String inputPath, {
    String? preferredExtension,
    Set<String>? usedPaths,
  }) {
    if (_overwrite) return inputPath;

    final outputDir = _outputDir ?? p.dirname(inputPath);
    final base = p.basenameWithoutExtension(inputPath);
    final ext =
        preferredExtension ??
        (p.extension(inputPath).isEmpty ? '.dxf' : p.extension(inputPath));

    String buildPath(int? suffix) {
      final tail = suffix == null ? '_replaced' : '_replaced_${suffix}';
      return p.join(outputDir, '${base}${tail}${ext}');
    }

    var candidate = buildPath(null);
    if (usedPaths == null) {
      return candidate;
    }

    var key = _normalizePathKey(candidate);
    if (!usedPaths.contains(key) && !File(candidate).existsSync()) {
      return candidate;
    }

    var seq = 2;
    while (true) {
      candidate = buildPath(seq);
      key = _normalizePathKey(candidate);
      if (!usedPaths.contains(key) && !File(candidate).existsSync()) {
        return candidate;
      }
      seq++;
    }
  }

  String _normalizePathKey(String inputPath) {
    return p.normalize(inputPath).toLowerCase();
  }

  String _outputModeLabel() {
    if (_overwrite) {
      return '覆盖原文件';
    }
    if (_outputDir == null) {
      return '输出到源目录';
    }
    return '输出目录: ${_outputDir}';
  }

  Future<void> _exportCsv() async {
    if (_results.isEmpty) return;
    await exportCsv(
      dialogTitle: '导出替换结果 CSV',
      suggestedName: '替换结果.csv',
      headers: const ['文件名', '对象类型', '图层', '原内容', '替换后', '使用规则'],
      rows: _results
          .map(
            (r) => [
              r.fileName,
              r.objectType,
              r.layer,
              r.originalText,
              r.updatedText,
              r.rule,
            ],
          )
          .toList(),
    );
  }

  Future<void> _exportXlsx() async {
    if (_results.isEmpty) return;
    await exportXlsx(
      dialogTitle: '导出替换结果 Excel',
      suggestedName: '替换结果.xlsx',
      sheetName: '替换结果',
      headers: const ['文件名', '对象类型', '图层', '原内容', '替换后', '使用规则'],
      rows: _results
          .map(
            (r) => [
              r.fileName,
              r.objectType,
              r.layer,
              r.originalText,
              r.updatedText,
              r.rule,
            ],
          )
          .toList(),
    );
  }

  void _showSnack(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(
      context,
    ).showSnackBar(SnackBar(content: Text(message)));
  }

  List<DxfReplaceResult> get _filteredResults {
    return _results.where((item) {
      if (_filterFile.isNotEmpty && item.fileName != _filterFile) return false;
      if (_filterType.isNotEmpty && item.objectType != _filterType)
        return false;
      if (_filterLayer.isNotEmpty && item.layer != _filterLayer) return false;
      if (_filterRule.isNotEmpty && !item.rule.contains(_filterRule))
        return false;
      if (_filterOriginal.isNotEmpty &&
          !item.originalText.toLowerCase().contains(
            _filterOriginal.toLowerCase(),
          )) {
        return false;
      }
      if (_filterUpdated.isNotEmpty &&
          !item.updatedText.toLowerCase().contains(
            _filterUpdated.toLowerCase(),
          )) {
        return false;
      }
      return true;
    }).toList();
  }

  List<String> _unique(List<String> values) {
    final set = values.where((v) => v.isNotEmpty).toSet().toList();
    set.sort((a, b) => a.compareTo(b));
    return ['全部', ...set];
  }

  @override
  Widget build(BuildContext context) {
    final filtered = _filteredResults;
    final files = _unique(_results.map((e) => e.fileName).toList());
    final types = _unique(_results.map((e) => e.objectType).toList());
    final layers = _unique(_results.map((e) => e.layer).toList());
    final theme = Theme.of(context);

    return LayoutBuilder(
      builder: (context, constraints) {
        final tableHeight = max(320.0, constraints.maxHeight - 360.0);
        return ListView(
          children: [
            SectionCard(
              title: 'DWG 文字内容替换',
              subtitle: '自动转 DXF，支持批量替换、预览与导出。',
              trailing: Row(
                mainAxisSize: MainAxisSize.min,
                children: [
                  FilledButton.icon(
                    onPressed: _isRunning ? null : _startScan,
                    icon: const Icon(Icons.find_replace),
                    label: const Text('执行模拟'),
                  ),
                  const SizedBox(width: 8),
                  OutlinedButton.icon(
                    onPressed: _isRunning ? _cancel : null,
                    icon: const Icon(Icons.cancel),
                    label: const Text('取消'),
                  ),
                ],
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  DropTarget(
                    onDragEntered: (_) => setState(() => _isDragging = true),
                    onDragExited: (_) => setState(() => _isDragging = false),
                    onDragDone: (detail) async {
                      setState(() => _isDragging = false);
                      await _handleDrop(
                        detail.files.map((file) => file.path).toList(),
                      );
                    },
                    child: Container(
                      width: double.infinity,
                      padding: const EdgeInsets.all(12),
                      decoration: BoxDecoration(
                        color: _isDragging
                            ? theme.colorScheme.primary.withOpacity(0.08)
                            : theme.colorScheme.surfaceVariant.withOpacity(0.2),
                        borderRadius: BorderRadius.circular(12),
                        border: Border.all(
                          color: _isDragging
                              ? theme.colorScheme.primary
                              : theme.dividerColor,
                        ),
                      ),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Wrap(
                            spacing: 8,
                            runSpacing: 8,
                            crossAxisAlignment: WrapCrossAlignment.center,
                            children: [
                              FilledButton.icon(
                                onPressed: _pickFiles,
                                icon: const Icon(Icons.upload_file),
                                label: const Text('选择 DWG 文件'),
                              ),
                              OutlinedButton.icon(
                                onPressed: _pickFolder,
                                icon: const Icon(Icons.folder_open),
                                label: const Text('选择文件夹'),
                              ),
                              Text(
                                '支持拖拽文件或文件夹到此区域',
                                style: theme.textTheme.bodySmall,
                              ),
                            ],
                          ),
                        ],
                      ),
                    ),
                  ),
                  const SizedBox(height: 12),
                  Container(
                    padding: const EdgeInsets.all(12),
                    decoration: BoxDecoration(
                      color: theme.colorScheme.surface,
                      borderRadius: BorderRadius.circular(12),
                      border: Border.all(color: theme.dividerColor),
                    ),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Row(
                          children: [
                            Text(
                              '已上传文件 (${_files.length})',
                              style: theme.textTheme.titleSmall?.copyWith(
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                            const Spacer(),
                            TextButton.icon(
                              onPressed: _files.isEmpty ? null : _clearFiles,
                              icon: const Icon(Icons.delete_outline),
                              label: const Text('清空'),
                            ),
                          ],
                        ),
                        const SizedBox(height: 8),
                        SizedBox(
                          height: 180,
                          child: _files.isEmpty
                              ? Center(
                                  child: Text(
                                    '暂无文件',
                                    style: theme.textTheme.bodySmall,
                                  ),
                                )
                              : Scrollbar(
                                  controller: _fileScrollController,
                                  thumbVisibility: true,
                                  child: ListView.separated(
                                    controller: _fileScrollController,
                                    itemCount: _files.length,
                                    separatorBuilder: (_, __) =>
                                        const Divider(height: 1),
                                    itemBuilder: (context, index) {
                                      final file = _files[index];
                                      final path = file.path ?? '';
                                      return ListTile(
                                        dense: true,
                                        visualDensity: VisualDensity.compact,
                                        leading: const Icon(
                                          Icons.description_outlined,
                                          size: 20,
                                        ),
                                        title: Text(
                                          file.name,
                                          maxLines: 1,
                                          overflow: TextOverflow.ellipsis,
                                        ),
                                        subtitle: path.isEmpty
                                            ? null
                                            : Text(
                                                path,
                                                maxLines: 1,
                                                overflow: TextOverflow.ellipsis,
                                              ),
                                        trailing: IconButton(
                                          tooltip: '移除',
                                          icon: const Icon(Icons.close),
                                          onPressed: () => _removeFile(file),
                                        ),
                                      );
                                    },
                                  ),
                                ),
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(height: 16),
                  Text('替换规则', style: Theme.of(context).textTheme.titleMedium),
                  const SizedBox(height: 8),
                  Column(
                    children: _pairRows.asMap().entries.map((entry) {
                      final index = entry.key;
                      final row = entry.value;
                      return Padding(
                        padding: const EdgeInsets.only(bottom: 8),
                        child: Row(
                          children: [
                            Expanded(
                              child: TextField(
                                controller: row.findController,
                                decoration: const InputDecoration(
                                  labelText: '查找内容',
                                ),
                              ),
                            ),
                            const SizedBox(width: 8),
                            Expanded(
                              child: TextField(
                                controller: row.replaceController,
                                decoration: const InputDecoration(
                                  labelText: '替换为',
                                ),
                              ),
                            ),
                            const SizedBox(width: 8),
                            IconButton(
                              onPressed: () =>
                                  index == 0 ? _addPair() : _removePair(index),
                              icon: Icon(
                                index == 0
                                    ? Icons.add_circle
                                    : Icons.remove_circle,
                              ),
                            ),
                          ],
                        ),
                      );
                    }).toList(),
                  ),
                  const SizedBox(height: 8),
                  Wrap(
                    spacing: 12,
                    runSpacing: 8,
                    crossAxisAlignment: WrapCrossAlignment.center,
                    children: [
                      FilledButton.icon(
                        onPressed: _results.isEmpty || _isRunning
                            ? null
                            : _applyReplace,
                        icon: const Icon(Icons.check_circle),
                        label: const Text('确认替换'),
                      ),
                      OutlinedButton.icon(
                        onPressed: _results.isEmpty ? null : _exportCsv,
                        icon: const Icon(Icons.table_view),
                        label: const Text('导出 CSV'),
                      ),
                      OutlinedButton.icon(
                        onPressed: _results.isEmpty ? null : _exportXlsx,
                        icon: const Icon(Icons.grid_on),
                        label: const Text('导出 Excel'),
                      ),
                      Row(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          Checkbox(
                            value: _overwrite,
                            onChanged: (value) =>
                                setState(() => _overwrite = value ?? false),
                          ),
                          const Text('覆盖原文件'),
                        ],
                      ),
                      OutlinedButton.icon(
                        onPressed: _overwrite ? null : _pickOutputDir,
                        icon: const Icon(Icons.folder_open),
                        label: Text(_outputDir == null ? '选择输出目录' : '输出目录已选择'),
                      ),
                    ],
                  ),
                  const SizedBox(height: 6),
                  Text(_outputModeLabel(), style: theme.textTheme.bodySmall),
                  const SizedBox(height: 12),
                  if (_isRunning) ...[
                    LinearProgressIndicator(value: _progress),
                    const SizedBox(height: 8),
                  ],
                  if (_status.isNotEmpty) Text(_status),
                ],
              ),
            ),
            const SizedBox(height: 12),
            Card(
              child: Padding(
                padding: const EdgeInsets.all(12),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Wrap(
                      spacing: 12,
                      runSpacing: 12,
                      children: [
                        _buildDropdown('文件', files, _filterFile, (value) {
                          setState(
                            () =>
                                _filterFile = value == '全部' ? '' : value ?? '',
                          );
                        }),
                        _buildDropdown('类型', types, _filterType, (value) {
                          setState(
                            () =>
                                _filterType = value == '全部' ? '' : value ?? '',
                          );
                        }),
                        _buildDropdown('图层', layers, _filterLayer, (value) {
                          setState(
                            () =>
                                _filterLayer = value == '全部' ? '' : value ?? '',
                          );
                        }),
                        SizedBox(
                          width: 220,
                          child: TextField(
                            decoration: const InputDecoration(
                              labelText: '原内容过滤',
                            ),
                            onChanged: (value) =>
                                setState(() => _filterOriginal = value.trim()),
                          ),
                        ),
                        SizedBox(
                          width: 220,
                          child: TextField(
                            decoration: const InputDecoration(
                              labelText: '替换后过滤',
                            ),
                            onChanged: (value) =>
                                setState(() => _filterUpdated = value.trim()),
                          ),
                        ),
                        SizedBox(
                          width: 220,
                          child: TextField(
                            decoration: const InputDecoration(
                              labelText: '规则过滤',
                            ),
                            onChanged: (value) =>
                                setState(() => _filterRule = value.trim()),
                          ),
                        ),
                        OutlinedButton.icon(
                          onPressed: () {
                            setState(() {
                              _filterFile = '';
                              _filterType = '';
                              _filterLayer = '';
                              _filterRule = '';
                              _filterOriginal = '';
                              _filterUpdated = '';
                            });
                          },
                          icon: const Icon(Icons.filter_alt_off),
                          label: const Text('清除筛选'),
                        ),
                      ],
                    ),
                    const SizedBox(height: 12),
                    SizedBox(
                      height: tableHeight,
                      child: SingleChildScrollView(
                        scrollDirection: Axis.horizontal,
                        child: PaginatedDataTable(
                          header: Text('替换结果 (${filtered.length})'),
                          rowsPerPage: _rowsPerPage,
                          onRowsPerPageChanged: (value) {
                            if (value == null) return;
                            setState(() => _rowsPerPage = value);
                          },
                          columns: const [
                            DataColumn(label: Text('文件名')),
                            DataColumn(label: Text('对象类型')),
                            DataColumn(label: Text('图层')),
                            DataColumn(label: Text('原内容')),
                            DataColumn(label: Text('替换后')),
                            DataColumn(label: Text('使用规则')),
                            DataColumn(label: Text('操作')),
                          ],
                          source: _ReplaceDataSource(
                            filtered,
                            onToggle: (item) {
                              setState(() {
                                item.skip = !item.skip;
                              });
                            },
                          ),
                        ),
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ],
        );
      },
    );
  }

  Widget _buildDropdown(
    String label,
    List<String> items,
    String value,
    ValueChanged<String?> onChanged,
  ) {
    return SizedBox(
      width: 160,
      child: InputDecorator(
        decoration: InputDecoration(labelText: label),
        child: DropdownButtonHideUnderline(
          child: DropdownButton<String>(
            value: value.isEmpty ? '全部' : value,
            isDense: true,
            items: items
                .map((item) => DropdownMenuItem(value: item, child: Text(item)))
                .toList(),
            onChanged: onChanged,
          ),
        ),
      ),
    );
  }
}

class _ReplaceDataSource extends DataTableSource {
  _ReplaceDataSource(this.rows, {required this.onToggle});

  final List<DxfReplaceResult> rows;
  final void Function(DxfReplaceResult) onToggle;

  @override
  DataRow? getRow(int index) {
    if (index >= rows.length) return null;
    final row = rows[index];
    return DataRow.byIndex(
      index: index,
      cells: [
        DataCell(Text(row.fileName)),
        DataCell(Text(row.objectType)),
        DataCell(Text(row.layer.isEmpty ? '-' : row.layer)),
        DataCell(Text(row.originalText)),
        DataCell(Text(row.updatedText)),
        DataCell(Text(row.rule)),
        DataCell(
          TextButton(
            onPressed: () => onToggle(row),
            child: Text(row.skip ? '恢复' : '撤销'),
          ),
        ),
      ],
    );
  }

  @override
  bool get isRowCountApproximate => false;

  @override
  int get rowCount => rows.length;

  @override
  int get selectedRowCount => 0;
}

class _PairRow {
  _PairRow()
    : findController = TextEditingController(),
      replaceController = TextEditingController();

  final TextEditingController findController;
  final TextEditingController replaceController;

  void dispose() {
    findController.dispose();
    replaceController.dispose();
  }
}
