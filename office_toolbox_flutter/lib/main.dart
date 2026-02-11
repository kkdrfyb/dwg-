import 'package:flutter/material.dart';

import 'app/app.dart';
import 'core/logging/log_service.dart';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  final logService = await LogService.create();
  runApp(OfficeToolboxApp(logService: logService));
}
