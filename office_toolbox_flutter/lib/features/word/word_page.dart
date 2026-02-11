import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../../core/task/task_service.dart';
import '../../widgets/section_card.dart';

class WordPage extends StatelessWidget {
  const WordPage({super.key});

  @override
  Widget build(BuildContext context) {
    final taskService = context.read<TaskService>();

    return SingleChildScrollView(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          SectionCard(
            title: 'Word 工具集',
            subtitle: '内容提取、差异对比、批量替换。',
            trailing: FilledButton.icon(
              onPressed: () => taskService.runDemoTask(title: 'Word 扫描演示'),
              icon: const Icon(Icons.play_arrow),
              label: const Text('演示任务'),
            ),
            child: Wrap(
              spacing: 8,
              runSpacing: 8,
              children: const [
                Chip(label: Text('内容提取')),
                Chip(label: Text('差异对比')),
                Chip(label: Text('批量替换')),
                Chip(label: Text('CSV/XLSX 导出')),
              ],
            ),
          ),
          const SizedBox(height: 16),
          SectionCard(
            title: '开发说明',
            subtitle: 'Word 模块将在后续版本中实现。',
            child: Text(
              '优先目标：模板填充、段落/表格批量替换与差异对比。',
              style: Theme.of(context).textTheme.bodyMedium,
            ),
          ),
        ],
      ),
    );
  }
}
