import 'package:flutter_test/flutter_test.dart';

import 'package:office_toolbox_flutter/app/app.dart';
import 'package:office_toolbox_flutter/core/logging/log_service.dart';

void main() {
  testWidgets('App boots with shell UI', (WidgetTester tester) async {
    final logService = LogService.forTesting();
    await tester.pumpWidget(OfficeToolboxApp(logService: logService));

    expect(find.text('Office Toolbox'), findsOneWidget);
    expect(find.text('Excel 工具集'), findsOneWidget);
  });
}
