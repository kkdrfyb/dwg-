import 'dart:math';

import 'package:desktop_drop/desktop_drop.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/export/export_utils.dart';
import '../../core/logging/log_service.dart';
import '../../core/task/task_limiter.dart';
import '../../core/task/task_service.dart';
import '../../widgets/section_card.dart';
import 'dxf_cache_service.dart';
import 'dxf_file_helper.dart';
import 'dxf_isolate.dart';
import 'dxf_models.dart';

class _ValvePayload {
  const _ValvePayload({required this.fileName, required this.response});

  final String fileName;
  final Map<String, dynamic> response;
}

class DxfValveTab extends StatefulWidget {
  const DxfValveTab({super.key});

  @override
  State<DxfValveTab> createState() => _DxfValveTabState();
}

class _DxfValveTabState extends State<DxfValveTab> {
  final ScrollController _fileScrollController = ScrollController();
  final List<PlatformFile> _files = [];
  final List<DxfValveResult> _results = [];

  bool _isRunning = false;
  bool _isDragging = false;
  double _progress = 0;
  String _status = '';
  String? _activeTaskId;

  int _rowsPerPage = PaginatedDataTable.defaultRowsPerPage;

  @override
  void dispose() {
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
    setState(() {
      _files.remove(file);
    });
  }

  void _clearFiles() {
    setState(() => _files.clear());
  }

  Future<void> _startScan() async {
    if (_files.isEmpty) {
      _showSnack('请先选择 DWG/DXF 文件');
      return;
    }

    final taskService = context.read<TaskService>();
    final log = context.read<LogService>();
    final cache = context.read<DxfCacheService>();
    final handle = taskService.startTask('DXF 阀门风口统计');

    setState(() {
      _isRunning = true;
      _progress = 0;
      _status = '准备处理 ${_files.length} 个文件...';
      _results.clear();
      _activeTaskId = handle.id;
    });

    final payloads = await taskService.runTask<List<_ValvePayload>>(handle, (context) async {
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

      if (prepared.isEmpty) {
        return <_ValvePayload>[];
      }

      final tasks = prepared
          .map(
            (file) => () async {
              try {
                final response = await compute(extractValveInfoFromFile, {
                  'path': file.dxfPath,
                  'name': file.sourceName,
                });
                return _ValvePayload(fileName: file.sourceName, response: response);
              } catch (error) {
                return _ValvePayload(
                  fileName: file.sourceName,
                  response: {'ok': false, 'error': error.toString(), 'results': <Map<String, String>>[]},
                );
              }
            },
          )
          .toList();

      final limiter = TaskLimiter.auto();
      return limiter.run<_ValvePayload>(
        tasks,
        isCanceled: context.isCanceled,
        onItemCompleted: (index, completed, total) {
          final progress = completed / max(1, total);
          context.updateProgress(progress, message: '解析中 ($completed/$total)');
          if (mounted) {
            setState(() {
              _progress = progress;
              _status = '解析中 ($completed/$total): ${prepared[index].sourceName}';
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

    final results = <DxfValveResult>[];
    for (final payload in payloads) {
      final response = payload.response;
      if (response['ok'] != true && response['error'] != null) {
        await log.warn('解析失败: ${payload.fileName}', context: 'dxf', error: response['error']);
      }

      final items = (response['results'] as List)
          .map(
            (item) => DxfValveResult(
              fileName: item['fileName'] as String,
              kind: item['kind'] as String,
              name: item['name'] as String,
              code: item['code'] as String,
              size: item['size'] as String,
              height: item['height'] as String,
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
      _status = results.isEmpty ? '未识别到阀门/风口信息' : '完成，共 ${results.length} 条';
    });
  }

  void _cancel() {
    final taskId = _activeTaskId;
    if (taskId == null) return;
    context.read<TaskService>().cancelTask(taskId);
  }

  Future<void> _exportCsv() async {
    if (_results.isEmpty) return;
    await exportCsv(
      dialogTitle: '导出阀门风口统计 CSV',
      suggestedName: '阀门风口统计.csv',
      headers: const ['文件名', '类型', '名称', '编号', '尺寸', '标高'],
      rows: _results
          .map((r) => [r.fileName, r.kind, r.name, r.code, r.size, r.height])
          .toList(),
    );
  }

  Future<void> _exportXlsx() async {
    if (_results.isEmpty) return;
    await exportXlsx(
      dialogTitle: '导出阀门风口统计 Excel',
      suggestedName: '阀门风口统计.xlsx',
      sheetName: '阀门风口统计',
      headers: const ['文件名', '类型', '名称', '编号', '尺寸', '标高'],
      rows: _results
          .map((r) => [r.fileName, r.kind, r.name, r.code, r.size, r.height])
          .toList(),
    );
  }

  void _showSnack(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text(message)),
    );
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    return LayoutBuilder(
      builder: (context, constraints) {
        final tableHeight = max(320.0, constraints.maxHeight - 280.0);
        return ListView(
          children: [
            SectionCard(
              title: '阀门/风口信息统计',
              subtitle: '自动转 DXF，识别阀门、风口的编号、尺寸与标高。',
              trailing: Row(
                mainAxisSize: MainAxisSize.min,
                children: [
                  FilledButton.icon(
                    onPressed: _isRunning ? null : _startScan,
                    icon: const Icon(Icons.play_arrow),
                    label: const Text('开始提取'),
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
                      await _handleDrop(detail.files.map((file) => file.path).toList());
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
                              Text('支持拖拽文件或文件夹到此区域',
                                  style: theme.textTheme.bodySmall),
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
                              style: theme.textTheme.titleSmall?.copyWith(fontWeight: FontWeight.w600),
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
                                  child: Text('暂无文件', style: theme.textTheme.bodySmall),
                                )
                              : Scrollbar(
                                  controller: _fileScrollController,
                                  thumbVisibility: true,
                                  child: ListView.separated(
                                    controller: _fileScrollController,
                                    itemCount: _files.length,
                                    separatorBuilder: (_, __) => const Divider(height: 1),
                                    itemBuilder: (context, index) {
                                      final file = _files[index];
                                      final path = file.path ?? '';
                                      return ListTile(
                                        dense: true,
                                        visualDensity: VisualDensity.compact,
                                        leading: const Icon(Icons.description_outlined, size: 20),
                                        title: Text(file.name, maxLines: 1, overflow: TextOverflow.ellipsis),
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
                  const SizedBox(height: 12),
                  if (_isRunning) ...[
                    LinearProgressIndicator(value: _progress),
                    const SizedBox(height: 8),
                  ],
                  if (_status.isNotEmpty) Text(_status),
                  const SizedBox(height: 12),
                  Wrap(
                    spacing: 8,
                    children: [
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
                    ],
                  ),
                ],
              ),
            ),
            const SizedBox(height: 12),
            Card(
              child: Padding(
                padding: const EdgeInsets.all(12),
                child: SizedBox(
                  height: tableHeight,
                  child: SingleChildScrollView(
                    scrollDirection: Axis.horizontal,
                    child: PaginatedDataTable(
                      header: Text('识别结果 (${_results.length})'),
                      rowsPerPage: _rowsPerPage,
                      onRowsPerPageChanged: (value) {
                        if (value == null) return;
                        setState(() => _rowsPerPage = value);
                      },
                      columns: const [
                        DataColumn(label: Text('文件名')),
                        DataColumn(label: Text('类型')),
                        DataColumn(label: Text('名称')),
                        DataColumn(label: Text('编号')),
                        DataColumn(label: Text('尺寸')),
                        DataColumn(label: Text('标高')),
                      ],
                      source: _ValveDataSource(_results),
                    ),
                  ),
                ),
              ),
            ),
          ],
        );
      },
    );
  }
}

class _ValveDataSource extends DataTableSource {
  _ValveDataSource(this.rows);

  final List<DxfValveResult> rows;

  @override
  DataRow? getRow(int index) {
    if (index >= rows.length) return null;
    final row = rows[index];
    return DataRow.byIndex(
      index: index,
      cells: [
        DataCell(Text(row.fileName)),
        DataCell(Text(row.kind)),
        DataCell(Text(row.name)),
        DataCell(Text(row.code)),
        DataCell(Text(row.size)),
        DataCell(Text(row.height)),
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
