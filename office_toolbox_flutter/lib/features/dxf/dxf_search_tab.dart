import 'dart:math';

import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/export/export_utils.dart';
import '../../core/logging/log_service.dart';
import '../../core/task/task_limiter.dart';
import '../../core/task/task_service.dart';
import '../../widgets/section_card.dart';
import 'dxf_isolate.dart';
import 'dxf_models.dart';

class _SearchPayload {
  const _SearchPayload({required this.file, required this.response});

  final PlatformFile file;
  final Map<String, dynamic> response;
}

class DxfSearchTab extends StatefulWidget {
  const DxfSearchTab({super.key});

  @override
  State<DxfSearchTab> createState() => _DxfSearchTabState();
}

class _DxfSearchTabState extends State<DxfSearchTab> {
  final TextEditingController _keywordsController = TextEditingController();
  final List<PlatformFile> _files = [];
  final List<DxfSearchResult> _results = [];

  bool _isRunning = false;
  double _progress = 0;
  String _status = '';
  String? _activeTaskId;

  String _filterFile = '';
  String _filterType = '';
  String _filterLayer = '';
  String _filterKeyword = '';
  String _filterContent = '';

  int _rowsPerPage = PaginatedDataTable.defaultRowsPerPage;

  @override
  void dispose() {
    _keywordsController.dispose();
    super.dispose();
  }

  Future<void> _pickFiles() async {
    final result = await FilePicker.platform.pickFiles(
      allowMultiple: true,
      type: FileType.custom,
      allowedExtensions: const ['dxf'],
    );
    if (result == null) return;

    setState(() {
      _files
        ..clear()
        ..addAll(result.files.where((file) => file.path != null));
    });
  }

  void _removeFile(PlatformFile file) {
    setState(() {
      _files.remove(file);
    });
  }

  void _clearFiles() {
    setState(() {
      _files.clear();
    });
  }

  Future<void> _startScan() async {
    if (_files.isEmpty) {
      _showSnack('请先选择 DXF 文件');
      return;
    }

    final keywords = _keywordsController.text
        .split(',')
        .map((k) => k.trim())
        .where((k) => k.isNotEmpty)
        .toList();

    final taskService = context.read<TaskService>();
    final log = context.read<LogService>();
    final handle = taskService.startTask('DXF 关键字扫描');

    setState(() {
      _isRunning = true;
      _progress = 0;
      _status = '准备扫描 ${_files.length} 个文件...';
      _results.clear();
      _activeTaskId = handle.id;
    });

    final tasks = _files
        .map(
          (file) => () async {
            try {
              final response = await compute(scanDxfFileForKeywords, {
                'path': file.path,
                'name': file.name,
                'keywords': keywords,
                'size': file.size,
              });
              return _SearchPayload(file: file, response: response);
            } catch (error) {
              return _SearchPayload(
                file: file,
                response: {'ok': false, 'error': error.toString(), 'results': <Map<String, String>>[]},
              );
            }
          },
        )
        .toList();

    final payloads = await taskService.runTask<List<_SearchPayload>>(handle, (context) async {
      final limiter = TaskLimiter.auto();
      return limiter.run<_SearchPayload>(
        tasks,
        isCanceled: context.isCanceled,
        onItemCompleted: (index, completed, total) {
          final progress = completed / max(1, total);
          context.updateProgress(progress, message: '扫描中 ($completed/$total)');
          if (mounted) {
            setState(() {
              _progress = progress;
              _status = '扫描中 ($completed/$total): ${_files[index].name}';
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

    final results = <DxfSearchResult>[];
    for (final payload in payloads) {
      final response = payload.response;
      final ok = response['ok'] == true;
      if (!ok && response['error'] != null) {
        await log.warn('解析失败: ${payload.file.name}', context: 'dxf', error: response['error']);
      }
      final items = (response['results'] as List)
          .map((item) => DxfSearchResult(
                fileName: item['fileName'] as String,
                objectType: item['objectType'] as String,
                layer: item['layer'] as String,
                keyword: item['keyword'] as String,
                content: item['content'] as String,
              ))
          .toList();
      results.addAll(items);
    }

    setState(() {
      _results
        ..clear()
        ..addAll(results);
      _isRunning = false;
      _progress = 1;
      _status = results.isEmpty
          ? '扫描完成，未找到匹配内容'
          : '扫描完成，找到 ${results.length} 条结果';
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
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text(message)),
    );
  }

  List<DxfSearchResult> get _filteredResults {
    return _results.where((item) {
      if (_filterFile.isNotEmpty && item.fileName != _filterFile) return false;
      if (_filterType.isNotEmpty && item.objectType != _filterType) return false;
      if (_filterLayer.isNotEmpty && item.layer != _filterLayer) return false;
      if (_filterKeyword.isNotEmpty && item.keyword != _filterKeyword) return false;
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

    return LayoutBuilder(
      builder: (context, constraints) {
        final tableHeight = max(320.0, constraints.maxHeight - 320.0);
        return ListView(
          children: [
            SectionCard(
              title: 'DXF 关键字查找',
              subtitle: '支持批量扫描、过滤与导出。',
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
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    crossAxisAlignment: WrapCrossAlignment.center,
                    children: [
                      FilledButton.icon(
                        onPressed: _pickFiles,
                        icon: const Icon(Icons.upload_file),
                        label: const Text('选择 DXF 文件'),
                      ),
                      OutlinedButton.icon(
                        onPressed: _files.isEmpty ? null : _clearFiles,
                        icon: const Icon(Icons.delete_outline),
                        label: const Text('清空列表'),
                      ),
                      Text('已选择 ${_files.length} 个文件'),
                    ],
                  ),
                  if (_files.isNotEmpty) ...[
                    const SizedBox(height: 12),
                    Wrap(
                      spacing: 8,
                      runSpacing: 8,
                      children: _files
                          .map(
                            (file) => Chip(
                              label: Text(file.name),
                              onDeleted: () => _removeFile(file),
                            ),
                          )
                          .toList(),
                    ),
                  ],
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
                          setState(() => _filterFile = value == '全部' ? '' : value ?? '');
                        }),
                        _buildDropdown('类型', types, _filterType, (value) {
                          setState(() => _filterType = value == '全部' ? '' : value ?? '');
                        }),
                        _buildDropdown('图层', layers, _filterLayer, (value) {
                          setState(() => _filterLayer = value == '全部' ? '' : value ?? '');
                        }),
                        _buildDropdown('关键字', keywords, _filterKeyword, (value) {
                          setState(() => _filterKeyword = value == '全部' ? '' : value ?? '');
                        }),
                        SizedBox(
                          width: 220,
                          child: TextField(
                            decoration: const InputDecoration(
                              labelText: '匹配内容过滤',
                            ),
                            onChanged: (value) => setState(() => _filterContent = value.trim()),
                          ),
                        ),
                        OutlinedButton.icon(
                          onPressed: () {
                            setState(() {
                              _filterFile = '';
                              _filterType = '';
                              _filterLayer = '';
                              _filterKeyword = '';
                              _filterContent = '';
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
                          header: Text('匹配结果 (${filtered.length})'),
                          rowsPerPage: _rowsPerPage,
                          onRowsPerPageChanged: (value) {
                            if (value == null) return;
                            setState(() => _rowsPerPage = value);
                          },
                          columns: const [
                            DataColumn(label: Text('文件名')),
                            DataColumn(label: Text('对象类型')),
                            DataColumn(label: Text('图层')),
                            DataColumn(label: Text('关键字')),
                            DataColumn(label: Text('匹配内容')),
                          ],
                          source: _SearchDataSource(filtered),
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
                .map(
                  (item) => DropdownMenuItem(
                    value: item,
                    child: Text(item),
                  ),
                )
                .toList(),
            onChanged: onChanged,
          ),
        ),
      ),
    );
  }
}

class _SearchDataSource extends DataTableSource {
  _SearchDataSource(this.rows);

  final List<DxfSearchResult> rows;

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
        DataCell(Text(row.keyword)),
        DataCell(Text(row.content)),
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
