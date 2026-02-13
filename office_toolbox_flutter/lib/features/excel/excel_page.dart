import 'dart:io';
import 'dart:math';

import 'package:excel/excel.dart' hide Border;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/export/export_utils.dart';
import '../../core/logging/log_service.dart';
import '../../core/task/task_service.dart';
import '../../widgets/section_card.dart';
import 'excel_isolate.dart';
import 'excel_models.dart';

class _ExcelModeGuide {
  const _ExcelModeGuide({required this.description, required this.demo});

  final String description;
  final String demo;
}

const Map<ExcelMode, _ExcelModeGuide> _excelModeGuides = {
  ExcelMode.mergeWorkbooks: _ExcelModeGuide(
    description: '将多个文件的所有工作表复制到一个新工作簿，并生成索引表。',
    demo: '示例：门店A.xlsx + 门店B.xlsx -> 输出工作表：门店A-1月份、门店B-1月份...',
  ),
  ExcelMode.mergeToSheet: _ExcelModeGuide(
    description: '将多个工作簿的数据提取到一个工作表，支持表头/表尾控制。',
    demo: '示例：两份月报纵向叠加为一张“汇总结果”表。',
  ),
  ExcelMode.internalMerge: _ExcelModeGuide(
    description: '对每个工作簿内部多工作表执行提取合并。',
    demo: '示例：门店A.xlsx 的 1月份+2月份 -> 全工作簿汇总。',
  ),
  ExcelMode.reorderColumns: _ExcelModeGuide(
    description: '按目标字段顺序重排列，并支持字段别名映射。',
    demo: '示例：字段顺序=品名,价格,数量；别名“商品名=品名”。',
  ),
  ExcelMode.splitWorkbook: _ExcelModeGuide(
    description: '按工作表拆分工作簿，每个工作表输出一个文件。',
    demo: '示例：门店A.xlsx -> 门店A-1月份.xlsx、门店A-2月份.xlsx。',
  ),
  ExcelMode.splitWorksheet: _ExcelModeGuide(
    description: '按某列字段值拆分明细，可输出为一个文件多表或多个文件。',
    demo: '示例：按“水果名”拆分 -> 苹果表、西瓜表。',
  ),
  ExcelMode.regroupSameSheetToWorkbook: _ExcelModeGuide(
    description: '将多个文件的同名工作表重组为单独工作簿。',
    demo: '示例：A/B 文件都有“1月份” -> 输出 1月份.xlsx（内含门店A、门店B表）。',
  ),
  ExcelMode.mergeToSheetSummary: _ExcelModeGuide(
    description: '跨文件跨表做字段聚合汇总（自动识别数值列求和）。',
    demo: '示例：按“品名”汇总，得到销量总和与明细。',
  ),
  ExcelMode.internalSummary: _ExcelModeGuide(
    description: '对单个工作簿内部多表做聚合汇总，生成“汇总表”。',
    demo: '示例：门店A.xlsx 的多月明细按品名汇总到一表。',
  ),
  ExcelMode.sameNameSheet: _ExcelModeGuide(
    description: '跨文件按同名工作表执行提取合并。',
    demo: '示例：所有文件的“1月份”提取到结果中的“1月份”表。',
  ),
  ExcelMode.sameNameSheetSummary: _ExcelModeGuide(
    description: '跨文件按同名工作表执行聚合汇总。',
    demo: '示例：所有“1月份”表汇总后输出“1月份”结果表。',
  ),
  ExcelMode.samePosition: _ExcelModeGuide(
    description: '按指定单元格位置提取值；为空则提取首表内容。',
    demo: '示例：提取 A1,H1,H2 -> 形成清单表。',
  ),
  ExcelMode.samePositionSummary: _ExcelModeGuide(
    description: '按指定位置做汇总（数值求和，文本去重拼接）。',
    demo: '示例：H1 汇总总营业额，A1 汇总店名列表。',
  ),
  ExcelMode.sameFilename: _ExcelModeGuide(
    description: '将同名文件场景汇总到一个工作簿，可重命名或跳过重复表名。',
    demo: '示例：不同目录同名文件统一汇总并保留索引。',
  ),
  ExcelMode.mergeDynamic: _ExcelModeGuide(
    description: '动态字段合并，自动对齐列；支持别名映射和目标字段顺序。',
    demo: '示例：人民币价格/美元价格合并到同一结果表。',
  ),
};

class ExcelPage extends StatefulWidget {
  const ExcelPage({super.key});

  @override
  State<ExcelPage> createState() => _ExcelPageState();
}

class _ExcelPageState extends State<ExcelPage> {
  final List<PlatformFile> _files = [];
  final TextEditingController _headerController = TextEditingController(
    text: '1',
  );
  final TextEditingController _footerController = TextEditingController(
    text: '0',
  );
  final TextEditingController _cellRangeController = TextEditingController();
  final TextEditingController _splitKeyController = TextEditingController(
    text: 'A',
  );
  final TextEditingController _fieldOrderController = TextEditingController();
  final TextEditingController _aliasRulesController = TextEditingController();
  ExcelMode _mode = ExcelMode.mergeWorkbooks;
  ExcelDirection _direction = ExcelDirection.vertical;
  InternalMergeMode _internalMode = InternalMergeMode.newSheetFirst;
  SameNameMode _sameNameMode = SameNameMode.rename;
  SplitWorksheetOutputMode _splitWorksheetOutputMode =
      SplitWorksheetOutputMode.oneWorkbook;
  int _headerRows = 1;
  int _footerRows = 0;
  String _cellRange = '';
  String _splitKey = 'A';
  String _fieldOrder = '';
  String _aliasRules = '';

  bool _isRunning = false;
  double _progress = 0;
  String _status = '';
  String? _activeTaskId;
  ExcelJobResult? _lastResult;
  ExcelIsolateRunner? _runner;

  @override
  void dispose() {
    _headerController.dispose();
    _footerController.dispose();
    _cellRangeController.dispose();
    _splitKeyController.dispose();
    _fieldOrderController.dispose();
    _aliasRulesController.dispose();
    super.dispose();
  }

  Future<void> _pickFiles() async {
    final result = await FilePicker.platform.pickFiles(
      allowMultiple: true,
      type: FileType.custom,
      allowedExtensions: const ['xlsx', 'xls', 'csv'],
    );
    if (result == null) return;

    setState(() {
      _files
        ..clear()
        ..addAll(result.files.where((file) => file.path != null));
    });
  }

  void _removeFile(PlatformFile file) {
    setState(() => _files.remove(file));
  }

  void _clearFiles() {
    setState(() => _files.clear());
  }

  void _cancel() {
    final taskId = _activeTaskId;
    if (taskId == null) return;
    context.read<TaskService>().cancelTask(taskId);
    _runner?.cancel();
  }

  bool get _needsHeaderFooter {
    return _mode == ExcelMode.mergeToSheet ||
        _mode == ExcelMode.mergeToSheetSummary ||
        _mode == ExcelMode.internalMerge ||
        _mode == ExcelMode.internalSummary ||
        _mode == ExcelMode.sameNameSheet ||
        _mode == ExcelMode.sameNameSheetSummary ||
        _mode == ExcelMode.splitWorksheet ||
        (_mode == ExcelMode.samePosition && _cellRange.trim().isEmpty);
  }

  Future<void> _runJob({required bool preview}) async {
    if (_files.isEmpty) {
      _showSnack('请先选择 Excel 文件');
      return;
    }
    if (_files.any((file) => file.name.toLowerCase().endsWith('.xls'))) {
      _showSnack('暂不支持 .xls，请先转换为 .xlsx');
      return;
    }

    final job = ExcelJob(
      mode: _mode,
      files: _files
          .map(
            (file) => ExcelInputFile(
              name: file.name,
              path: file.path!,
              size: file.size,
            ),
          )
          .toList(),
      headerRows: _headerRows,
      footerRows: _footerRows,
      direction: _direction,
      internalMode: _internalMode,
      sameNameMode: _sameNameMode,
      cellRange: _cellRange,
      splitKey: _splitKey,
      fieldOrder: _fieldOrder,
      aliasRules: _aliasRules,
      splitWorksheetOutputMode: _splitWorksheetOutputMode,
      preview: preview,
    );

    final taskService = context.read<TaskService>();
    final log = context.read<LogService>();
    final handle = taskService.startTask(preview ? 'Excel 预览' : 'Excel 处理');
    final runner = ExcelIsolateRunner();
    _runner = runner;

    setState(() {
      _isRunning = true;
      _progress = 0;
      _status = preview ? '预览处理中...' : '处理中...';
      _activeTaskId = handle.id;
      _lastResult = null;
    });

    final result = await taskService.runTask<ExcelJobResult>(handle, (
      context,
    ) async {
      final res = await runner.run(
        job,
        onProgress: (progress) {
          final total = max(1, progress.total);
          final value = progress.completed / total;
          context.updateProgress(value, message: progress.message);
          if (mounted) {
            setState(() {
              _progress = value;
              _status = progress.message;
            });
          }
        },
      );
      return res;
    });

    if (!mounted) return;

    if (result == null) {
      setState(() {
        _isRunning = false;
        _status = '处理失败';
      });
      await log.error('Excel 处理失败', context: 'excel');
      return;
    }
    if (result.canceled) {
      setState(() {
        _isRunning = false;
        _status = '已取消';
      });
      await log.warn('Excel 操作已取消', context: 'excel');
      return;
    }

    setState(() {
      _isRunning = false;
      _progress = 1;
      _status = '完成';
      _lastResult = result;
    });

    if (preview) {
      if (result.preview == null) {
        _showSnack('预览不可用');
        return;
      }
      _showPreview(result.preview!);
      return;
    }

    await _saveOutputs(result);
  }

  Future<void> _saveOutputs(ExcelJobResult result) async {
    if (result.outputs.isEmpty) {
      _showSnack('没有可保存的输出');
      return;
    }

    if (result.outputs.length == 1) {
      final output = result.outputs.first;
      await saveBytes(
        dialogTitle: '保存 Excel 文件',
        suggestedName: output.filename,
        allowedExtensions: const ['xlsx'],
        bytes: output.bytes,
      );
      return;
    }

    final dir = await FilePicker.platform.getDirectoryPath();
    if (dir == null) return;
    for (final output in result.outputs) {
      final path = '$dir${Platform.pathSeparator}${output.filename}';
      await File(path).writeAsBytes(output.bytes, flush: true);
    }
    _showSnack('已保存 ${result.outputs.length} 个文件');
  }

  Future<void> _exportCsv() async {
    final result = _lastResult;
    if (result == null || result.outputs.isEmpty) return;
    if (result.outputs.length > 1) {
      _showSnack('多工作簿结果暂不支持直接导出 CSV');
      return;
    }

    final bytes = result.outputs.first.bytes;
    final excel = Excel.decodeBytes(bytes);
    if (excel.sheets.isEmpty) {
      _showSnack('无可导出的工作表');
      return;
    }
    final sheet = excel.sheets.values.first;
    final rows = <List<String>>[];
    for (var r = 0; r < sheet.maxRows; r++) {
      final row = sheet.row(r);
      rows.add(row.map((cell) => cell?.value?.toString() ?? '').toList());
    }

    await exportCsv(
      dialogTitle: '导出 CSV',
      suggestedName: '合并结果.csv',
      headers: rows.isNotEmpty ? rows.first : const [],
      rows: rows.length > 1 ? rows.sublist(1) : const [],
    );
  }

  void _showPreview(ExcelPreview preview) {
    showDialog<void>(
      context: context,
      builder: (context) {
        final rows = preview.rows;
        final maxCols = rows.fold<int>(0, (prev, row) => max(prev, row.length));
        return AlertDialog(
          title: Text('预览：${preview.sheetName} (前 ${rows.length} 行)'),
          content: SizedBox(
            width: min(MediaQuery.of(context).size.width * 0.8, 900),
            height: min(MediaQuery.of(context).size.height * 0.7, 520),
            child: rows.isEmpty
                ? const Center(child: Text('空表'))
                : SingleChildScrollView(
                    child: Table(
                      defaultVerticalAlignment:
                          TableCellVerticalAlignment.middle,
                      border: TableBorder.all(color: Colors.black12),
                      children: rows
                          .map(
                            (row) => TableRow(
                              children: List.generate(
                                maxCols,
                                (index) => Padding(
                                  padding: const EdgeInsets.all(6),
                                  child: Text(
                                    index < row.length ? row[index] : '',
                                  ),
                                ),
                              ),
                            ),
                          )
                          .toList(),
                    ),
                  ),
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(context).pop(),
              child: const Text('关闭'),
            ),
          ],
        );
      },
    );
  }

  void _showSnack(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(
      context,
    ).showSnackBar(SnackBar(content: Text(message)));
  }

  @override
  Widget build(BuildContext context) {
    return ListView(
      children: [
        SectionCard(
          title: 'Excel 工具集',
          subtitle: '合并/筛选/导出，支持进度与取消。',
          trailing: Row(
            mainAxisSize: MainAxisSize.min,
            children: [
              FilledButton.icon(
                onPressed: _isRunning ? null : () => _runJob(preview: true),
                icon: const Icon(Icons.preview),
                label: const Text('预览'),
              ),
              const SizedBox(width: 8),
              FilledButton.icon(
                onPressed: _isRunning ? null : () => _runJob(preview: false),
                icon: const Icon(Icons.play_arrow),
                label: const Text('开始处理'),
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
                    label: const Text('选择文件'),
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
              Text('功能模式', style: Theme.of(context).textTheme.titleMedium),
              const SizedBox(height: 8),
              Wrap(
                spacing: 8,
                runSpacing: 8,
                children: ExcelMode.values
                    .map(
                      (mode) => ChoiceChip(
                        label: Text(mode.label),
                        selected: _mode == mode,
                        onSelected: (selected) {
                          if (!selected) return;
                          setState(() => _mode = mode);
                        },
                      ),
                    )
                    .toList(),
              ),
              const SizedBox(height: 16),
              _buildOptions(),
              const SizedBox(height: 8),
              _buildModeGuide(),
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
                    onPressed: (_lastResult?.outputs.isNotEmpty ?? false)
                        ? _exportCsv
                        : null,
                    icon: const Icon(Icons.table_view),
                    label: const Text('导出 CSV'),
                  ),
                ],
              ),
            ],
          ),
        ),
      ],
    );
  }

  Widget _buildOptions() {
    final widgets = <Widget>[];

    if (_mode == ExcelMode.mergeToSheet) {
      widgets.add(
        _buildEnumDropdown<ExcelDirection>(
          label: '排列方向',
          value: _direction,
          items: const [
            DropdownMenuItem(
              value: ExcelDirection.vertical,
              child: Text('竖向叠加'),
            ),
            DropdownMenuItem(
              value: ExcelDirection.horizontal,
              child: Text('横向并列'),
            ),
          ],
          onChanged: (value) =>
              setState(() => _direction = value ?? ExcelDirection.vertical),
        ),
      );
    }

    if (_mode == ExcelMode.internalMerge ||
        _mode == ExcelMode.internalSummary) {
      widgets.add(
        _buildEnumDropdown<InternalMergeMode>(
          label: '汇总方式',
          value: _internalMode,
          items: const [
            DropdownMenuItem(
              value: InternalMergeMode.newSheetFirst,
              child: Text('新建汇总表(最前)'),
            ),
            DropdownMenuItem(
              value: InternalMergeMode.firstSheet,
              child: Text('合并到第1个工作表'),
            ),
          ],
          onChanged: (value) => setState(
            () => _internalMode = value ?? InternalMergeMode.newSheetFirst,
          ),
        ),
      );
    }

    if (_mode == ExcelMode.sameFilename) {
      widgets.add(
        _buildEnumDropdown<SameNameMode>(
          label: '同名工作表处理',
          value: _sameNameMode,
          items: const [
            DropdownMenuItem(value: SameNameMode.rename, child: Text('自动重命名')),
            DropdownMenuItem(value: SameNameMode.skip, child: Text('跳过重复')),
          ],
          onChanged: (value) =>
              setState(() => _sameNameMode = value ?? SameNameMode.rename),
        ),
      );
    }

    if (_mode == ExcelMode.samePosition ||
        _mode == ExcelMode.samePositionSummary) {
      widgets.add(
        TextField(
          controller: _cellRangeController,
          decoration: const InputDecoration(
            labelText: '指定单元格 (如 A1,B2)',
            hintText: '为空则合并首个工作表',
          ),
          onChanged: (value) => setState(() => _cellRange = value),
        ),
      );
    }

    if (_mode == ExcelMode.splitWorksheet) {
      widgets.add(
        TextField(
          controller: _splitKeyController,
          decoration: const InputDecoration(
            labelText: '拆分字段 (列名或列字母)',
            hintText: '如: 水果名 或 A',
          ),
          onChanged: (value) => setState(() => _splitKey = value.trim()),
        ),
      );
      widgets.add(
        _buildEnumDropdown<SplitWorksheetOutputMode>(
          label: '输出方式',
          value: _splitWorksheetOutputMode,
          items: const [
            DropdownMenuItem(
              value: SplitWorksheetOutputMode.oneWorkbook,
              child: Text('输出到一个文件'),
            ),
            DropdownMenuItem(
              value: SplitWorksheetOutputMode.separateFiles,
              child: Text('输出到多个文件'),
            ),
            DropdownMenuItem(
              value: SplitWorksheetOutputMode.both,
              child: Text('两种都输出'),
            ),
          ],
          onChanged: (value) => setState(
            () => _splitWorksheetOutputMode =
                value ?? SplitWorksheetOutputMode.oneWorkbook,
          ),
        ),
      );
    }

    if (_mode == ExcelMode.reorderColumns || _mode == ExcelMode.mergeDynamic) {
      widgets.add(
        TextField(
          controller: _fieldOrderController,
          decoration: const InputDecoration(
            labelText: '目标字段顺序 (逗号分隔)',
            hintText: '如: 品名,价格,数量,厂家,备注',
          ),
          onChanged: (value) => setState(() => _fieldOrder = value.trim()),
        ),
      );
      widgets.add(
        TextField(
          controller: _aliasRulesController,
          minLines: 2,
          maxLines: 4,
          decoration: const InputDecoration(
            labelText: '字段别名映射 (可选)',
            hintText: '每行一条，格式: 原字段=目标字段',
          ),
          onChanged: (value) => setState(() => _aliasRules = value.trim()),
        ),
      );
    }

    if (_needsHeaderFooter) {
      widgets.add(
        Row(
          children: [
            Expanded(
              child: TextField(
                controller: _headerController,
                decoration: const InputDecoration(labelText: '保留表头行数'),
                keyboardType: TextInputType.number,
                onChanged: (value) =>
                    setState(() => _headerRows = int.tryParse(value) ?? 0),
              ),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: TextField(
                controller: _footerController,
                decoration: const InputDecoration(labelText: '去除表尾行数'),
                keyboardType: TextInputType.number,
                onChanged: (value) =>
                    setState(() => _footerRows = int.tryParse(value) ?? 0),
              ),
            ),
          ],
        ),
      );
    }

    if (widgets.isEmpty) {
      return const SizedBox.shrink();
    }

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: widgets
          .map(
            (widget) => Padding(
              padding: const EdgeInsets.only(bottom: 12),
              child: widget,
            ),
          )
          .toList(),
    );
  }

  Widget _buildModeGuide() {
    final guide = _excelModeGuides[_mode];
    if (guide == null) {
      return const SizedBox.shrink();
    }
    final theme = Theme.of(context);
    return Container(
      width: double.infinity,
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(
        color: theme.colorScheme.surfaceContainerHighest.withValues(alpha: 0.18),
        border: Border.all(color: theme.dividerColor),
        borderRadius: BorderRadius.circular(10),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            '功能说明',
            style: theme.textTheme.titleSmall?.copyWith(
              fontWeight: FontWeight.w700,
            ),
          ),
          const SizedBox(height: 6),
          Text(guide.description),
          const SizedBox(height: 8),
          Text(
            '简单演示',
            style: theme.textTheme.titleSmall?.copyWith(
              fontWeight: FontWeight.w700,
            ),
          ),
          const SizedBox(height: 6),
          Text(guide.demo),
        ],
      ),
    );
  }

  Widget _buildEnumDropdown<T>({
    required String label,
    required T value,
    required List<DropdownMenuItem<T>> items,
    required ValueChanged<T?> onChanged,
  }) {
    return InputDecorator(
      decoration: InputDecoration(labelText: label),
      child: DropdownButtonHideUnderline(
        child: DropdownButton<T>(
          value: value,
          isDense: true,
          items: items,
          onChanged: onChanged,
        ),
      ),
    );
  }
}
