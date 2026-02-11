enum TaskStatus { queued, running, completed, failed, canceled }

extension TaskStatusX on TaskStatus {
  String get label {
    switch (this) {
      case TaskStatus.queued:
        return 'Queued';
      case TaskStatus.running:
        return 'Running';
      case TaskStatus.completed:
        return 'Completed';
      case TaskStatus.failed:
        return 'Failed';
      case TaskStatus.canceled:
        return 'Canceled';
    }
  }
}

class TaskItem {
  TaskItem({
    required this.id,
    required this.title,
    required this.status,
    required this.progress,
    required this.createdAt,
    this.startedAt,
    this.endedAt,
    this.message,
    this.error,
  });

  final String id;
  final String title;
  final TaskStatus status;
  final double progress;
  final DateTime createdAt;
  final DateTime? startedAt;
  final DateTime? endedAt;
  final String? message;
  final String? error;

  TaskItem copyWith({
    TaskStatus? status,
    double? progress,
    DateTime? startedAt,
    DateTime? endedAt,
    String? message,
    String? error,
  }) {
    return TaskItem(
      id: id,
      title: title,
      status: status ?? this.status,
      progress: progress ?? this.progress,
      createdAt: createdAt,
      startedAt: startedAt ?? this.startedAt,
      endedAt: endedAt ?? this.endedAt,
      message: message ?? this.message,
      error: error ?? this.error,
    );
  }
}
