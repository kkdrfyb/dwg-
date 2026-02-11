import 'package:flutter/foundation.dart';
import 'package:uuid/uuid.dart';

import 'task_models.dart';

class TaskController extends ChangeNotifier {
  final List<TaskItem> _tasks = [];
  final Uuid _uuid = const Uuid();

  List<TaskItem> get tasks => List.unmodifiable(_tasks);

  TaskItem create(String title) {
    final item = TaskItem(
      id: _uuid.v4(),
      title: title,
      status: TaskStatus.queued,
      progress: 0,
      createdAt: DateTime.now(),
    );
    _tasks.insert(0, item);
    notifyListeners();
    return item;
  }

  void start(String id) {
    _update(id, (item) => item.copyWith(status: TaskStatus.running, startedAt: DateTime.now()));
  }

  void updateProgress(String id, double progress, {String? message}) {
    _update(id, (item) => item.copyWith(progress: progress, message: message));
  }

  void complete(String id, {String? message}) {
    _update(
      id,
      (item) => item.copyWith(
        status: TaskStatus.completed,
        progress: 1,
        endedAt: DateTime.now(),
        message: message,
      ),
    );
  }

  void fail(String id, String error, {String? message}) {
    _update(
      id,
      (item) => item.copyWith(
        status: TaskStatus.failed,
        endedAt: DateTime.now(),
        message: message,
        error: error,
      ),
    );
  }

  void cancel(String id, {String? message}) {
    _update(
      id,
      (item) => item.copyWith(
        status: TaskStatus.canceled,
        endedAt: DateTime.now(),
        message: message,
      ),
    );
  }

  void _update(String id, TaskItem Function(TaskItem) update) {
    final index = _tasks.indexWhere((task) => task.id == id);
    if (index == -1) return;
    _tasks[index] = update(_tasks[index]);
    notifyListeners();
  }
}
