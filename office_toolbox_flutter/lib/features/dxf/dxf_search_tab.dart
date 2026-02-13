import 'dart:math';

import 'package:desktop_drop/desktop_drop.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/export/export_utils.dart';
import '../../core/task/task_service.dart';
import '../../core/task/task_exceptions.dart';
import '../../widgets/section_card.dart';
import 'dxf_cache_service.dart';
import 'dxf_file_helper.dart';
import 'dxf_models.dart';

class DxfSearchTab extends StatefulWidget {
  const DxfSearchTab({super.key});

  @override
  State<DxfSearchTab> createState() => _DxfSearchTabState();
}

class _DxfSearchTabState extends State<DxfSearchTab> {
  final TextEditingController _keywordsController = TextEditingController();
  final ScrollController _fileScrollController = ScrollController();
  final ScrollController _resultScrollController = ScrollController();
  final List<PlatformFile> _files = [];
  final List<DxfSearchResult> _results = [];

  bool _isRunning = false;
  bool _isDragging = false;
  bool _pendingIndex = false;
  double _progress = 0;
  String _status = '';
  String? _activeTaskId;

  String _filterFile = '';
  String _filterType = '';
  String _filterLayer = '';
  String _filterKeyword = '';
  String _filterContent = '';

  @override
  void dispose() {
    _keywordsController.dispose();
    _fileScrollController.dispose();
    _resultScrollController.dispose();
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
    _scheduleIndex();
  }

  void _removeFile(PlatformFile file) {
    setState(() {
      _files.remove(file);
    });
    _scheduleIndex();
  }

  void _clearFiles() {
    setState(() {
      _files.clear();
      _results.clear();
      _clearFilters();
    });
    _scheduleIndex();
  }

  void _clearFilters() {
    _filterFile = '';
    _filterType = '';
    _filterLayer = '';
    _filterKeyword = '';
    _filterContent = '';
  }

  void _scheduleIndex() {
    if (_isRunning) {
      _pendingIndex = true;
      return;
    }
    _runIndex();
  }

  Future<void> _runIndex() async {
    final taskService = context.read<TaskService>();
    final cache = context.read<DxfCacheService>();
    final handle = taskService.startTask('DXF 索引更新');

    setState(() {
      _isRunning = true;
      _progress = 0;
      _status = '准备更新索引...';
      _activeTaskId = handle.id;
    });

    final ok = await taskService.runTask<bool>(handle, (taskContext) async {
      await cache.ensureIndex(
        _files,
        taskContext,
        onProgress: (progress, message) {
          if (!mounted) return;
          setState(() {
            _progress = progress;
            _status = message;
          });
        },
      );
      return true;
    });

    if (!mounted) return;

    if (ok == null) {
      setState(() {
        _isRunning = false;
        _status = '已取消';
      });
    } else {
      setState(() {
        _isRunning = false;
        _progress = 1;
        _status = '索引更新完成';
      });
    }

    if (_pendingIndex) {
      _pendingIndex = false;
      _runIndex();
    }
  }

  Future<void> _startScan() async {
    if (_files.isEmpty) {
      _showSnack('请先选择 DWG/DXF 文件');
      return;
    }

    final keywords = _keywordsController.text
        .split(',')
        .map((k) => k.trim())
        .where((k) => k.isNotEmpty)
        .toList();

    final taskService = context.read<TaskService>();
    final cache = context.read<DxfCacheService>();
    final handle = taskService.startTask('DXF 关键字查询');

    setState(() {
      _isRunning = true;
      _progress = 0;
      _status = '准备查询...';
      _results.clear();
      _clearFilters();
      _activeTaskId = handle.id;
    });

    final results = await taskService.runTask<List<DxfSearchResult>>(handle, (
      taskContext,
    ) async {
      await cache.ensureIndex(
        _files,
        taskContext,
        onProgress: (progress, message) {
          if (!mounted) return;
          setState(() {
            _progress = progress;
            _status = message;
          });
        },
      );
      if (taskContext.isCanceled()) {
        throw TaskCanceled();
      }
      final data = await cache.queryKeywords(keywords);
      return data;
    });

    if (!mounted) return;

    if (results == null) {
      setState(() {
        _isRunning = false;
        _status = '已取消';
      });
      return;
    }

    setState(() {
      _results
        ..clear()
        ..addAll(results);
      _isRunning = false;
      _progress = 1;
      _status = results.isEmpty
          ? '查询完成，未找到匹配内容'
          : '查询完成，找到 ${results.length} 条结果';
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
      dialogTitle: '导出扫描结果 CSV',
      suggestedName: '扫描结果.csv',
      headers: const ['文件名', '对象类型', '图层', '关键字', '匹配内容'],
      rows: _results
          .map((r) => [r.fileName, r.objectType, r.layer, r.keyword, r.content])
          .toList(),
    );
  }

  Future<void> _exportXlsx() async {
    if (_results.isEmpty) return;
    await exportXlsx(
      dialogTitle: '导出扫描结果 Excel',
      suggestedName: '扫描结果.xlsx',
      sheetName: '扫描结果',
      headers: const ['文件名', '对象类型', '图层', '关键字', '匹配内容'],
      rows: _results
          .map((r) => [r.fileName, r.objectType, r.layer, r.keyword, r.content])
          .toList(),
    );
  }

  void _showSnack(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(
      context,
    ).showSnackBar(SnackBar(content: Text(message)));
  }

  List<DxfSearchResult> get _filteredResults {
    return _results.where((item) {
      if (_filterFile.isNotEmpty && item.fileName != _filterFile) return false;
      if (_filterType.isNotEmpty && item.objectType != _filterType)
        return false;
      if (_filterLayer.isNotEmpty && item.layer != _filterLayer) return false;
      if (_filterKeyword.isNotEmpty && item.keyword != _filterKeyword)
        return false;
      if (_filterContent.isNotEmpty &&
          !item.content.toLowerCase().contains(_filterContent.toLowerCase())) {
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
    final keywords = _unique(_results.map((e) => e.keyword).toList());
    final theme = Theme.of(context);
    final dbPath = context.read<DxfCacheService>().activeDbPath;

    return LayoutBuilder(
      builder: (context, constraints) {
        final tableHeight = max(320.0, constraints.maxHeight - 320.0);
        return ListView(
          children: [
            SectionCard(
              title: 'DWG 关键字查找',
              subtitle: '自动转 DXF 并建立索引，支持批量查询与导出。',
              trailing: Row(
                mainAxisSize: MainAxisSize.min,
                children: [
                  FilledButton.icon(
                    onPressed: _isRunning ? null : _startScan,
                    icon: const Icon(Icons.search),
                    label: const Text('开始扫描'),
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
                  TextField(
                    controller: _keywordsController,
                    decoration: const InputDecoration(
                      labelText: '关键字（逗号分隔）',
                      hintText: '例如：预埋件,HILTI',
                    ),
                  ),
                  const SizedBox(height: 12),
                  if (_isRunning) ...[
                    LinearProgressIndicator(value: _progress),
                    const SizedBox(height: 8),
                  ],
                  if (_status.isNotEmpty) Text(_status),
                  if (dbPath != null && dbPath.isNotEmpty) ...[
                    const SizedBox(height: 4),
                    SelectableText(
                      '索引数据库: $dbPath',
                      style: theme.textTheme.bodySmall,
                    ),
                  ],
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
                        _buildDropdown('关键字', keywords, _filterKeyword, (
                          value,
                        ) {
                          setState(
                            () => _filterKeyword = value == '全部'
                                ? ''
                                : value ?? '',
                          );
                        }),
                        SizedBox(
                          width: 220,
                          child: TextField(
                            decoration: const InputDecoration(
                              labelText: '匹配内容过滤',
                            ),
                            onChanged: (value) =>
                                setState(() => _filterContent = value.trim()),
                          ),
                        ),
                        OutlinedButton.icon(
                          onPressed: () {
                            setState(_clearFilters);
                          },
                          icon: const Icon(Icons.filter_alt_off),
                          label: const Text('清除筛选'),
                        ),
                      ],
                    ),
                    const SizedBox(height: 12),
                    SizedBox(
                      height: tableHeight,
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Padding(
                            padding: const EdgeInsets.only(bottom: 8),
                            child: Text(
                              '匹配结果 (${filtered.length}/${_results.length})',
                            ),
                          ),
                          Expanded(
                            child: Scrollbar(
                              controller: _resultScrollController,
                              thumbVisibility: true,
                              child: SingleChildScrollView(
                                controller: _resultScrollController,
                                child: SingleChildScrollView(
                                  scrollDirection: Axis.horizontal,
                                  child: DataTable(
                                    headingRowHeight: 42,
                                    dataRowMinHeight: 44,
                                    dataRowMaxHeight: 64,
                                    columns: const [
                                      DataColumn(label: Text('文件名')),
                                      DataColumn(label: Text('对象类型')),
                                      DataColumn(label: Text('图层')),
                                      DataColumn(label: Text('关键字')),
                                      DataColumn(label: Text('匹配内容')),
                                    ],
                                    rows: filtered
                                        .map(
                                          (row) => DataRow(
                                            cells: [
                                              DataCell(Text(row.fileName)),
                                              DataCell(Text(row.objectType)),
                                              DataCell(
                                                Text(
                                                  row.layer.isEmpty
                                                      ? '-'
                                                      : row.layer,
                                                ),
                                              ),
                                              DataCell(Text(row.keyword)),
                                              DataCell(Text(row.content)),
                                            ],
                                          ),
                                        )
                                        .toList(),
                                  ),
                                ),
                              ),
                            ),
                          ),
                        ],
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
