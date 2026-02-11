import 'dart:io';

import 'package:flutter/foundation.dart';
import 'package:intl/intl.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';

enum LogLevel { debug, info, warn, error }

extension LogLevelX on LogLevel {
  String get label {
    switch (this) {
      case LogLevel.debug:
        return 'DEBUG';
      case LogLevel.info:
        return 'INFO';
      case LogLevel.warn:
        return 'WARN';
      case LogLevel.error:
        return 'ERROR';
    }
  }
}

class LogEntry {
  LogEntry({
    required this.timestamp,
    required this.level,
    required this.message,
    this.context,
    this.error,
  });

  final DateTime timestamp;
  final LogLevel level;
  final String message;
  final String? context;
  final Object? error;

  String toLine(DateFormat formatter) {
    final time = formatter.format(timestamp);
    final ctx = context == null ? '' : ' [$context]';
    final err = error == null ? '' : ' | $error';
    return '$time ${level.label}$ctx: $message$err\n';
  }
}

class LogService extends ChangeNotifier {
  LogService._(this._logFile, this._formatter);

  final File _logFile;
  final DateFormat _formatter;
  final List<LogEntry> _entries = [];
  Future<void> _writeQueue = Future.value();

  static Future<LogService> create() async {
    final baseDir = await getApplicationSupportDirectory();
    final logDir = Directory(p.join(baseDir.path, 'office_toolbox', 'logs'));
    await logDir.create(recursive: true);
    final file = File(p.join(logDir.path, 'office_toolbox.log'));
    return LogService._(file, DateFormat('yyyy-MM-dd HH:mm:ss'));
  }

  @visibleForTesting
  static LogService forTesting({String? path}) {
    final filePath = path ?? p.join(Directory.systemTemp.path, 'office_toolbox_test.log');
    return LogService._(File(filePath), DateFormat('yyyy-MM-dd HH:mm:ss'));
  }

  String get logFilePath => _logFile.path;
  List<LogEntry> get entries => List.unmodifiable(_entries);

  Future<void> debug(String message, {String? context, Object? error}) {
    return _log(LogLevel.debug, message, context: context, error: error);
  }

  Future<void> info(String message, {String? context, Object? error}) {
    return _log(LogLevel.info, message, context: context, error: error);
  }

  Future<void> warn(String message, {String? context, Object? error}) {
    return _log(LogLevel.warn, message, context: context, error: error);
  }

  Future<void> error(String message, {String? context, Object? error}) {
    return _log(LogLevel.error, message, context: context, error: error);
  }

  Future<void> _log(
    LogLevel level,
    String message, {
    String? context,
    Object? error,
  }) {
    final entry = LogEntry(
      timestamp: DateTime.now(),
      level: level,
      message: message,
      context: context,
      error: error,
    );
    _entries.add(entry);
    if (_entries.length > 500) {
      _entries.removeRange(0, _entries.length - 500);
    }
    notifyListeners();

    final line = entry.toLine(_formatter);
    _writeQueue = _writeQueue.then((_) async {
      try {
        await _logFile.writeAsString(line, mode: FileMode.append, flush: true);
      } catch (_) {}
    });
    return _writeQueue;
  }

  void clearMemory() {
    _entries.clear();
    notifyListeners();
  }
}
