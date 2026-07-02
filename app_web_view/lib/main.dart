import 'package:flutter/material.dart';

import 'screens/splash_screen.dart';
import 'theme/app_theme.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const AITranslatorApp());
}

class AITranslatorApp extends StatelessWidget {
  const AITranslatorApp({super.key, this.initialToken});

  /// Giữ tương thích test cũ — token thật lấy từ SharedPreferences trong SplashScreen.
  final String? initialToken;

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'AI Translator',
      debugShowCheckedModeBanner: false,
      theme: AppTheme.dark(),
      home: const SplashScreen(),
    );
  }
}
