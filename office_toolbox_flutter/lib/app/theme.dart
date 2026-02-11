import 'package:flutter/material.dart';

ThemeData buildAppTheme() {
  const primary = Color(0xFF1E4D5A);
  const secondary = Color(0xFFF29F05);
  const surface = Color(0xFFF9FBFC);
  const background = Color(0xFFF4F7FA);

  final colorScheme = ColorScheme.fromSeed(
    seedColor: primary,
    brightness: Brightness.light,
  ).copyWith(
    primary: primary,
    secondary: secondary,
    surface: surface,
  );

  final baseTextTheme = ThemeData.light().textTheme;
  final textTheme = baseTextTheme.copyWith(
    displayLarge: baseTextTheme.displayLarge?.copyWith(fontWeight: FontWeight.w800),
    displayMedium: baseTextTheme.displayMedium?.copyWith(fontWeight: FontWeight.w800),
    displaySmall: baseTextTheme.displaySmall?.copyWith(fontWeight: FontWeight.w700),
    headlineLarge: baseTextTheme.headlineLarge?.copyWith(fontWeight: FontWeight.w700),
    headlineMedium: baseTextTheme.headlineMedium?.copyWith(fontWeight: FontWeight.w700),
    headlineSmall: baseTextTheme.headlineSmall?.copyWith(fontWeight: FontWeight.w700),
    titleLarge: baseTextTheme.titleLarge?.copyWith(fontWeight: FontWeight.w700),
    titleMedium: baseTextTheme.titleMedium?.copyWith(fontWeight: FontWeight.w600),
    titleSmall: baseTextTheme.titleSmall?.copyWith(fontWeight: FontWeight.w600),
    bodyLarge: baseTextTheme.bodyLarge?.copyWith(fontWeight: FontWeight.w600),
    bodyMedium: baseTextTheme.bodyMedium?.copyWith(fontWeight: FontWeight.w500),
    bodySmall: baseTextTheme.bodySmall?.copyWith(fontWeight: FontWeight.w500),
    labelLarge: baseTextTheme.labelLarge?.copyWith(fontWeight: FontWeight.w600),
    labelMedium: baseTextTheme.labelMedium?.copyWith(fontWeight: FontWeight.w600),
    labelSmall: baseTextTheme.labelSmall?.copyWith(fontWeight: FontWeight.w600),
  ).apply(
    bodyColor: Colors.black87,
    displayColor: Colors.black87,
  );

  return ThemeData(
    useMaterial3: true,
    colorScheme: colorScheme,
    scaffoldBackgroundColor: background,
    fontFamily: 'Microsoft YaHei UI',
    fontFamilyFallback: const ['Segoe UI', 'Microsoft YaHei', 'Arial', 'sans-serif'],
    textTheme: textTheme,
    appBarTheme: const AppBarTheme(
      backgroundColor: primary,
      foregroundColor: Colors.white,
      centerTitle: false,
      elevation: 0,
    ),
    cardTheme: CardThemeData(
      color: surface,
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: BorderSide(color: Colors.black12),
      ),
    ),
    inputDecorationTheme: const InputDecorationTheme(
      border: OutlineInputBorder(),
    ),
  );
}
