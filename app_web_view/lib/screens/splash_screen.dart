import 'package:flutter/material.dart';

import '../services/auth_service.dart';
import '../theme/app_theme.dart';
import 'webview_screen.dart';

class SplashScreen extends StatefulWidget {
  const SplashScreen({super.key});

  @override
  State<SplashScreen> createState() => _SplashScreenState();
}

class _SplashScreenState extends State<SplashScreen> {
  final _auth = AuthService();

  @override
  void initState() {
    super.initState();
    _bootstrap();
  }

  Future<void> _bootstrap() async {
    await Future<void>.delayed(const Duration(milliseconds: 800));
    if (!mounted) return;

    // Thử khôi phục session — nếu có thì inject token vào WebView
    try {
      await _auth.restoreSession();
    } catch (_) {
      // Không có session → vẫn mở WebView, website tự xử lý
    }

    if (!mounted) return;

    // Luôn mở WebView trang chủ — y chang mở web trên trình duyệt
    Navigator.of(context).pushReplacement(
      MaterialPageRoute(
        builder: (_) => WebViewScreen(
          authService: _auth,
          profile: _auth.profile,
          token: _auth.token ?? '',
          initialPath: '/',
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppColors.bg,
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Image.asset('assets/logo.png', width: 100, height: 100),
            const SizedBox(height: 32),
            const CircularProgressIndicator(color: AppColors.teal),
            const SizedBox(height: 16),
            const Text(
              'AI Translator',
              style: TextStyle(
                color: AppColors.teal,
                fontSize: 18,
                fontWeight: FontWeight.w600,
                letterSpacing: 1.2,
              ),
            ),
          ],
        ),
      ),
    );
  }
}
