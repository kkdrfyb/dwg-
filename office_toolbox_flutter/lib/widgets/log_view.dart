import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:provider/provider.dart';

import '../core/logging/log_service.dart';

class LogView extends StatelessWidget {
  const LogView({super.key});

  @override
  Widget build(BuildContext context) {
    final formatter = DateFormat('HH:mm:ss');

    return Consumer<LogService>(
      builder: (context, logService, child) {
        final entries = logService.entries;
        if (entries.isEmpty) {
          return const Center(child: Text('暂无日志记录'));
        }

        return ListView.separated(
          itemCount: entries.length,
          separatorBuilder: (context, index) => const Divider(height: 1),
          itemBuilder: (context, index) {
            final entry = entries[index];
            final contextLabel = entry.context == null ? '' : ' · ${entry.context}';
            final subtitle = '${entry.level.label}$contextLabel';
            return ListTile(
              leading: Text(formatter.format(entry.timestamp)),
              title: Text(entry.message),
              subtitle: Text(subtitle),
              trailing: entry.error == null ? null : const Icon(Icons.error_outline),
            );
          },
        );
      },
    );
  }
}
