class TaskCanceled implements Exception {
  TaskCanceled([this.message]);

  final String? message;

  @override
  String toString() => message ?? 'Task was canceled';
}
