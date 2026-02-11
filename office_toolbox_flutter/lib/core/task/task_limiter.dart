import 'dart:io';
import 'dart:math';

import 'task_exceptions.dart';

class TaskLimiter {
  TaskLimiter({required this.maxConcurrent}) : assert(maxConcurrent > 0);

  factory TaskLimiter.auto({int cap = 6}) {
    final cores = Platform.numberOfProcessors;
    final computed = max(2, min(cores, cap));
    return TaskLimiter(maxConcurrent: computed);
  }

  final int maxConcurrent;

  Future<List<T>> run<T>(
    List<Future<T> Function()> tasks, {
    bool Function()? isCanceled,
    void Function(int completed, int total)? onProgress,
    void Function(int index, int completed, int total)? onItemCompleted,
  }) async {
    if (tasks.isEmpty) return <T>[];

    final results = List<T?>.filled(tasks.length, null);
    var cursor = 0;
    var completed = 0;
    var canceled = false;

    Future<void> runNext() async {
      if (canceled) return;
      final index = cursor++;
      if (index >= tasks.length) return;
      if (isCanceled?.call() == true) {
        canceled = true;
        return;
      }

      final result = await tasks[index]();
      results[index] = result;
      completed++;
      onItemCompleted?.call(index, completed, tasks.length);
      onProgress?.call(completed, tasks.length);

      if (!canceled) {
        await runNext();
      }
    }

    final slots = min(maxConcurrent, tasks.length);
    await Future.wait(List.generate(slots, (_) => runNext()));

    if (canceled || isCanceled?.call() == true) {
      throw TaskCanceled();
    }

    return results.cast<T>();
  }
}
