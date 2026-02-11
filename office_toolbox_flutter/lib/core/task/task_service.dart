import 'dart:async';

import '../logging/log_service.dart';
import 'task_controller.dart';
import 'task_exceptions.dart';

class TaskContext {
  TaskContext({
    required this.isCanceled,
    required this.updateProgress,
  });

  final bool Function() isCanceled;
  final void Function(double progress, {String? message}) updateProgress;
}

class TaskHandle {
  TaskHandle(this.id, this._token);

  final String id;
  final CancelToken _token;

  bool get isCanceled => _token.isCanceled;

  void cancel() => _token.cancel();
}

class TaskService {
  TaskService({required this.log, required this.tasks});

  final LogService log;
  final TaskController tasks;
  final Map<String, CancelToken> _tokens = {};

  TaskHandle startTask(String title) {
    final task = tasks.create(title);
    tasks.start(task.id);
    final token = _register(task.id);
    return TaskHandle(task.id, token);
  }

  Future<T?> runTask<T>(
    TaskHandle handle,
    Future<T> Function(TaskContext context) action, {
    String? successMessage,
  }) async {
    final context = TaskContext(
      isCanceled: () => handle.isCanceled,
      updateProgress: (progress, {message}) {
        tasks.updateProgress(handle.id, progress, message: message);
      },
    );

    await log.info('Task started: ${handle.id}', context: 'task');

    try {
      final result = await action(context);
      if (handle.isCanceled) {
        tasks.cancel(handle.id, message: 'Canceled by user');
        await log.warn('Task canceled: ${handle.id}', context: 'task');
        return null;
      }
      tasks.complete(handle.id, message: successMessage ?? 'Completed');
      await log.info('Task completed: ${handle.id}', context: 'task');
      return result;
    } on TaskCanceled catch (error) {
      tasks.cancel(handle.id, message: error.message ?? 'Canceled');
      await log.warn('Task canceled: ${handle.id}', context: 'task');
      return null;
    } catch (error) {
      tasks.fail(handle.id, error.toString(), message: 'Failed');
      await log.error('Task failed: ${handle.id}', context: 'task', error: error);
      return null;
    } finally {
      _tokens.remove(handle.id);
    }
  }

  Future<void> runDemoTask({required String title, int steps = 16}) async {
    final handle = startTask(title);
    await runTask<void>(handle, (context) async {
      for (var step = 0; step <= steps; step++) {
        if (context.isCanceled()) {
          throw TaskCanceled();
        }
        final progress = step / steps;
        context.updateProgress(progress, message: 'Step $step/$steps');
        await Future<void>.delayed(const Duration(milliseconds: 180));
      }
    });
  }

  void cancelTask(String id) {
    final token = _tokens[id];
    if (token == null) return;
    token.cancel();
  }

  CancelToken _register(String id) {
    final token = CancelToken();
    _tokens[id] = token;
    return token;
  }
}

class CancelToken {
  bool _canceled = false;

  bool get isCanceled => _canceled;

  void cancel() {
    _canceled = true;
  }
}
