import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../core/logging/log_service.dart';
import '../core/task/task_controller.dart';
import '../core/task/task_service.dart';
import '../features/dxf/dxf_cache_service.dart';
import 'app_scaffold.dart';
import 'theme.dart';

class OfficeToolboxApp extends StatelessWidget {
  const OfficeToolboxApp({super.key, required this.logService});

  final LogService logService;

  @override
  Widget build(BuildContext context) {
    return MultiProvider(
      providers: [
        ChangeNotifierProvider<LogService>.value(value: logService),
        ChangeNotifierProvider(create: (_) => TaskController()),
        ProxyProvider2<LogService, TaskController, TaskService>(
          update: (context, log, tasks, previous) => TaskService(log: log, tasks: tasks),
        ),
        Provider<DxfCacheService>(
          create: (context) => DxfCacheService(log: context.read<LogService>()),
        ),
      ],
      child: MaterialApp(
        title: 'Office Toolbox',
        theme: buildAppTheme(),
        home: const AppScaffold(),
      ),
    );
  }
}
