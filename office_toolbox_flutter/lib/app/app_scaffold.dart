import 'package:flutter/material.dart';

import '../features/dxf/dxf_page.dart';
import '../features/excel/excel_page.dart';
import '../features/logs/logs_page.dart';
import '../features/word/word_page.dart';
import '../widgets/task_list.dart';

class AppScaffold extends StatefulWidget {
  const AppScaffold({super.key});

  @override
  State<AppScaffold> createState() => _AppScaffoldState();
}

class _AppScaffoldState extends State<AppScaffold> {
  int _index = 0;

  final _destinations = const <NavigationDestination>[
    NavigationDestination(icon: Icon(Icons.table_view), label: 'Excel'),
    NavigationDestination(icon: Icon(Icons.architecture), label: 'CAD'),
    NavigationDestination(icon: Icon(Icons.description), label: 'Word'),
    NavigationDestination(icon: Icon(Icons.receipt_long), label: 'Logs'),
  ];

  Widget _buildPage(int index) {
    switch (index) {
      case 0:
        return const ExcelPage();
      case 1:
        return const DxfPage();
      case 2:
        return const WordPage();
      case 3:
        return const LogsPage();
      default:
        return const ExcelPage();
    }
  }

  void _showTasksSheet(BuildContext context) {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      showDragHandle: true,
      builder: (_) => const TaskListPanel(inSheet: true),
    );
  }

  @override
  Widget build(BuildContext context) {
    return LayoutBuilder(
      builder: (context, constraints) {
        final isWide = constraints.maxWidth >= 960;
        final showTaskPanel = constraints.maxWidth >= 1200;

        return Scaffold(
          appBar: AppBar(
            title: const Text('Office Toolbox'),
            actions: [
              if (!isWide)
                IconButton(
                  tooltip: 'Tasks',
                  icon: const Icon(Icons.track_changes),
                  onPressed: () => _showTasksSheet(context),
                ),
            ],
          ),
          body: Row(
            children: [
              if (isWide)
                NavigationRail(
                  selectedIndex: _index,
                  onDestinationSelected: (value) => setState(() => _index = value),
                  labelType: NavigationRailLabelType.all,
                  destinations: _destinations
                      .map(
                        (destination) => NavigationRailDestination(
                          icon: destination.icon,
                          label: Text(destination.label),
                        ),
                      )
                      .toList(),
                ),
              Expanded(
                child: Container(
                  decoration: const BoxDecoration(
                    gradient: LinearGradient(
                      colors: [Color(0xFFF4F7FA), Color(0xFFE9F0F5)],
                      begin: Alignment.topLeft,
                      end: Alignment.bottomRight,
                    ),
                  ),
                  child: SafeArea(
                    child: Padding(
                      padding: const EdgeInsets.all(24),
                      child: _buildPage(_index),
                    ),
                  ),
                ),
              ),
              if (showTaskPanel)
                const SizedBox(
                  width: 340,
                  child: TaskListPanel(),
                ),
            ],
          ),
          bottomNavigationBar: isWide
              ? null
              : NavigationBar(
                  selectedIndex: _index,
                  onDestinationSelected: (value) => setState(() => _index = value),
                  destinations: _destinations,
                ),
          floatingActionButton: isWide
              ? null
              : FloatingActionButton.extended(
                  onPressed: () => _showTasksSheet(context),
                  icon: const Icon(Icons.track_changes),
                  label: const Text('Tasks'),
                ),
        );
      },
    );
  }
}
