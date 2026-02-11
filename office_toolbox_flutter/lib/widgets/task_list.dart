import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../core/task/task_controller.dart';
import '../core/task/task_models.dart';
import '../core/task/task_service.dart';

class TaskListPanel extends StatelessWidget {
  const TaskListPanel({super.key, this.inSheet = false});

  final bool inSheet;

  @override
  Widget build(BuildContext context) {
    final taskService = context.read<TaskService>();
    final height = inSheet ? MediaQuery.of(context).size.height * 0.7 : null;

    return SizedBox(
      height: height,
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Padding(
            padding: const EdgeInsets.fromLTRB(16, 16, 16, 8),
            child: Row(
              children: [
                Text(
                  '任务中心',
                  style: Theme.of(context).textTheme.titleMedium,
                ),
                const SizedBox(width: 8),
                Consumer<TaskController>(
                  builder: (context, tasks, child) {
                    return Text(
                      '(${tasks.tasks.length})',
                      style: Theme.of(context).textTheme.bodyMedium,
                    );
                  },
                ),
              ],
            ),
          ),
          const Divider(height: 1),
          Expanded(
            child: Consumer<TaskController>(
              builder: (context, tasks, child) {
                if (tasks.tasks.isEmpty) {
                  return const Center(child: Text('暂无任务'));
                }

                return ListView.separated(
                  padding: const EdgeInsets.all(12),
                  itemCount: tasks.tasks.length,
                  separatorBuilder: (context, index) => const SizedBox(height: 8),
                  itemBuilder: (context, index) {
                    final task = tasks.tasks[index];
                    return Card(
                      child: Padding(
                        padding: const EdgeInsets.all(12),
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            Row(
                              children: [
                                Expanded(
                                  child: Text(
                                    task.title,
                                    style: Theme.of(context).textTheme.titleSmall,
                                  ),
                                ),
                                Text(task.status.label),
                              ],
                            ),
                            const SizedBox(height: 8),
                            if (task.status == TaskStatus.running)
                              LinearProgressIndicator(value: task.progress)
                            else
                              LinearProgressIndicator(value: task.progress, minHeight: 4),
                            if (task.message != null) ...[
                              const SizedBox(height: 8),
                              Text(task.message!),
                            ],
                            if (task.status == TaskStatus.running)
                              Align(
                                alignment: Alignment.centerRight,
                                child: TextButton.icon(
                                  onPressed: () => taskService.cancelTask(task.id),
                                  icon: const Icon(Icons.cancel),
                                  label: const Text('取消'),
                                ),
                              ),
                          ],
                        ),
                      ),
                    );
                  },
                );
              },
            ),
          ),
        ],
      ),
    );
  }
}
