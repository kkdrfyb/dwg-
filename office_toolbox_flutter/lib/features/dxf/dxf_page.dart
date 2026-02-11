import 'package:flutter/material.dart';

import '../../widgets/section_card.dart';
import 'dxf_replace_tab.dart';
import 'dxf_search_tab.dart';
import 'dxf_valve_tab.dart';

class DxfPage extends StatelessWidget {
  const DxfPage({super.key});

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        const SectionCard(
          title: 'CAD (DXF) 工具集',
          subtitle: '关键字查找、阀门统计、批量替换。',
          child: Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              Chip(label: Text('关键字查找')),
              Chip(label: Text('阀门风口统计')),
              Chip(label: Text('文字替换')),
              Chip(label: Text('CSV/XLSX 导出')),
              Chip(label: Text('后台任务与取消')),
            ],
          ),
        ),
        const SizedBox(height: 12),
        Expanded(
          child: DefaultTabController(
            length: 3,
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: const [
                TabBar(
                  tabs: [
                    Tab(text: '关键字查找'),
                    Tab(text: '阀门统计'),
                    Tab(text: '文字替换'),
                  ],
                ),
                SizedBox(height: 12),
                Expanded(
                  child: TabBarView(
                    children: [
                      DxfSearchTab(),
                      DxfValveTab(),
                      DxfReplaceTab(),
                    ],
                  ),
                ),
              ],
            ),
          ),
        ),
      ],
    );
  }
}
