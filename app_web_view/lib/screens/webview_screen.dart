import 'dart:convert';

import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_inappwebview/flutter_inappwebview.dart';
import 'package:url_launcher/url_launcher.dart';

import '../config/app_config.dart';
import '../models/user_profile.dart';
import '../services/auth_service.dart';
import '../services/session_storage.dart' as app_storage;
import '../theme/app_theme.dart';
import 'profile_screen.dart';

class WebViewScreen extends StatefulWidget {
  const WebViewScreen({
    super.key,
    required this.authService,
    required this.profile,
    required this.token,
    this.initialPath = '/',
    this.guestMode = false,
  });

  final AuthService authService;
  final UserProfile? profile;
  final String token;
  final String initialPath;
  final bool guestMode;

  @override
  State<WebViewScreen> createState() => _WebViewScreenState();
}

class _WebViewScreenState extends State<WebViewScreen> {
  InAppWebViewController? _controller;
  UserProfile? _profile;
  double _progress = 0;
  /// Token hiện tại — được xóa khi user logout
  String _activeToken = '';

  static const _chromeUA =
      'Mozilla/5.0 (Linux; Android 14; Pixel 8) '
      'AppleWebKit/537.36 (KHTML, like Gecko) '
      'Chrome/126.0.0.0 Mobile Safari/537.36';

  @override
  void initState() {
    super.initState();
    _profile = widget.profile;
    _activeToken = widget.token;
    SystemChrome.setSystemUIOverlayStyle(const SystemUiOverlayStyle(
      statusBarColor: Colors.transparent,
    ));
  }

  String get _startUrl {
    final path = widget.initialPath.startsWith('/')
        ? widget.initialPath
        : '/${widget.initialPath}';
    return '${AppConfig.baseUrl}$path';
  }

  bool get _hasSession => _activeToken.isNotEmpty;

  /// Build source JS inject token — dùng _activeToken (có thể thay đổi khi logout)
  String _buildTokenScript() {
    if (!_hasSession) return '';
    final user = _profile;
    final userJson = user != null ? jsonEncode(user.toStorageJson()) : '{}';
    final tokenEsc = _activeToken.replaceAll('\\', '\\\\').replaceAll("'", "\\'");
    final userEsc = userJson.replaceAll('\\', '\\\\').replaceAll("'", "\\'");
    return '''
      (function() {
        try {
          localStorage.setItem('token', '$tokenEsc');
          localStorage.setItem('user', '$userEsc');
        } catch(e) {}
      })();
    ''';
  }

  /// Thêm script inject token vào WebView (dùng addUserScript — có thể xóa khi logout)
  Future<void> _addTokenScript() async {
    if (!_hasSession || _controller == null) return;
    await _controller!.addUserScript(
      userScript: UserScript(
        source: _buildTokenScript(),
        injectionTime: UserScriptInjectionTime.AT_DOCUMENT_START,
        forMainFrameOnly: true,
      ),
    );
  }

  /// Xóa token script khi logout — ngăn re-inject token cũ
  Future<void> _clearTokenScript() async {
    await _controller?.removeAllUserScripts();
  }

  /// Xử lý Google OAuth: load URL trực tiếp trong WebView (không mở Chrome Custom Tab)
  /// → shouldOverrideUrlLoading sẽ bắt aitranslator://oauth?token=... và hoàn tất login
  Future<void> _handleGoogleAuth() async {
    try {
      final authUrl = await widget.authService.getGoogleAuthUrl();
      if (!mounted) return;
      // Load Google auth page thẳng vào WebView — Chrome user agent đã được set
      await _controller?.loadUrl(
        urlRequest: URLRequest(
          url: WebUri(authUrl),
          headers: AppConfig.defaultHeaders,
        ),
      );
    } catch (e) {
      if (!mounted) return;
      final msg = e.toString().replaceAll('\\', '\\\\').replaceAll("'", "\\'");
      await _controller?.evaluateJavascript(source: '''
        try {
          if (typeof showAuthMessage === 'function') {
            showAuthMessage('$msg', 'error');
          }
        } catch(_) {}
      ''');
    }
  }

  Future<void> _openProfile() async {
    if (_profile == null) return;
    final updated = await Navigator.of(context).push<UserProfile>(
      MaterialPageRoute(
        builder: (_) => ProfileScreen(
          authService: widget.authService,
          profile: _profile!,
        ),
      ),
    );
    if (updated != null) {
      setState(() => _profile = updated);
      final token = widget.authService.token ?? '';
      final userJson = jsonEncode(updated.toStorageJson());
      final tokenEsc = token.replaceAll('\\', '\\\\').replaceAll("'", "\\'");
      final userEsc = userJson.replaceAll('\\', '\\\\').replaceAll("'", "\\'");
      await _controller?.evaluateJavascript(source: '''
        try {
          localStorage.setItem('token', '$tokenEsc');
          localStorage.setItem('user', '$userEsc');
        } catch(e) {}
      ''');
      await _controller?.reload();
    }
  }

  @override
  Widget build(BuildContext context) {
    return PopScope(
      canPop: false,
      onPopInvokedWithResult: (didPop, _) async {
        if (didPop) return;
        if (_controller != null && await _controller!.canGoBack()) {
          await _controller!.goBack();
        }
      },
      child: Scaffold(
        backgroundColor: AppColors.bg,
        body: Stack(
          children: [
            InAppWebView(
              initialUrlRequest: URLRequest(
                url: WebUri(_startUrl),
                headers: AppConfig.defaultHeaders,
              ),
              initialSettings: InAppWebViewSettings(
                javaScriptEnabled: true,
                domStorageEnabled: true,
                databaseEnabled: true,
                cacheEnabled: true,
                allowFileAccess: true,
                allowContentAccess: true,
                allowFileAccessFromFileURLs: true,
                thirdPartyCookiesEnabled: true,
                mediaPlaybackRequiresUserGesture: false,
                allowsInlineMediaPlayback: true,
                useHybridComposition: true,
                supportZoom: false,
                builtInZoomControls: false,
                displayZoomControls: false,
                useOnDownloadStart: true,
                // User agent chuẩn Chrome → bypass in-app browser detection
                userAgent: _chromeUA,
              ),
              onWebViewCreated: (controller) {
                _controller = controller;

                // Inject token trước khi page JS chạy (có thể xóa khi logout)
                _addTokenScript();

                // Google auth
                controller.addJavaScriptHandler(
                  handlerName: 'flutterGoogleAuth',
                  callback: (_) { _handleGoogleAuth(); return null; },
                );

                // Download — mở Chrome ngoài
                controller.addJavaScriptHandler(
                  handlerName: 'flutterDownload',
                  callback: (args) async {
                    final url = args.isNotEmpty ? args[0].toString() : '';
                    if (url.isEmpty) return null;
                    try {
                      final uri = Uri.parse(url);
                      if (await canLaunchUrl(uri)) {
                        await launchUrl(uri, mode: LaunchMode.externalApplication);
                      }
                    } catch (e) { debugPrint('Download error: $e'); }
                    return null;
                  },
                );

                // Logout — xóa token script + session
                controller.addJavaScriptHandler(
                  handlerName: 'flutterLogout',
                  callback: (_) async {
                    await _clearTokenScript();          // xóa user scripts
                    await app_storage.SessionStorage.clear();
                    widget.authService.apiSetToken('');
                    if (mounted) setState(() {
                      _activeToken = '';
                      _profile = null;
                    });
                    return null;
                  },
                );

                controller.addJavaScriptHandler(
                  handlerName: 'flutterReady',
                  callback: (_) => null,
                );
              },
              onProgressChanged: (_, p) {
                setState(() => _progress = p / 100);
              },
              onLoadStop: (controller, url) async {
                await controller.evaluateJavascript(source: r'''
                  (function() {
                    window.__FLUTTER_APP__ = true;
                    document.documentElement.classList.add('flutter-app');

                    // Override Google auth
                    if (typeof isInAppBrowser !== 'undefined') {
                      isInAppBrowser = function() { return false; };
                    }
                    if (typeof signInWithGoogle !== 'undefined') {
                      signInWithGoogle = function() {
                        window.flutter_inappwebview.callHandler('flutterGoogleAuth');
                      };
                    }

                    // Intercept localStorage.removeItem('token') → báo Flutter logout
                    if (!window.__flLogoutPatched) {
                      window.__flLogoutPatched = true;
                      var _origRemove = Storage.prototype.removeItem;
                      Storage.prototype.removeItem = function(key) {
                        _origRemove.call(this, key);
                        if (key === 'token' && this === localStorage) {
                          try {
                            window.flutter_inappwebview.callHandler('flutterLogout');
                          } catch(e) {}
                        }
                      };
                    }

                    // Làm sáng nền mobile nav dropdown (homepage)
                    if (!document.getElementById('__fl_nav_css')) {
                      var s = document.createElement('style');
                      s.id = '__fl_nav_css';
                      s.textContent = `
                        .nav-links.mobile-menu {
                          background: rgba(8, 18, 48, 0.97) !important;
                          backdrop-filter: blur(16px) !important;
                          border-bottom: 1px solid rgba(0, 210, 210, 0.25) !important;
                          box-shadow: 0 8px 32px rgba(0,0,0,0.4) !important;
                        }
                        .nav-links.mobile-menu a {
                          color: #e0eaff !important;
                          padding: 10px 0 !important;
                          border-bottom: 1px solid rgba(255,255,255,0.07) !important;
                        }
                        .nav-links.mobile-menu a:hover { color: #00d2d2 !important; }
                      `;
                      document.head.appendChild(s);
                    }
                  })();
                ''');
              },
              onDownloadStartRequest: (controller, request) async {
                final url = request.url.toString();
                if (url.isEmpty) return;
                try {
                  final uri = Uri.parse(url);
                  if (await canLaunchUrl(uri)) {
                    await launchUrl(uri, mode: LaunchMode.externalApplication);
                  }
                } catch (e) {
                  debugPrint('Download start error: $e');
                }
              },
              // Bắt scheme aitranslator:// từ Google OAuth redirect
              shouldOverrideUrlLoading: (controller, action) async {
                final url = action.request.url?.toString() ?? '';
                if (url.startsWith('${AppConfig.oauthScheme}://')) {
                  final uri = Uri.parse(url);
                  final token = uri.queryParameters['token'];
                  final error = uri.queryParameters['error'];
                  if (token != null && token.isNotEmpty) {
                    // Lưu session và inject vào WebView
                    try {
                      final profile =
                          await widget.authService.finishGoogleLogin(token);
                      setState(() => _profile = profile);
                      final userJson = jsonEncode(profile.toStorageJson());
                      final tokenEsc = token
                          .replaceAll('\\', '\\\\')
                          .replaceAll("'", "\\'");
                      final userEsc = userJson
                          .replaceAll('\\', '\\\\')
                          .replaceAll("'", "\\'");
                      await controller.evaluateJavascript(source: '''
                        try {
                          localStorage.setItem('token', '$tokenEsc');
                          localStorage.setItem('user', '$userEsc');
                          window.location.href = '${AppConfig.baseUrl}/dashboard';
                        } catch(e) {}
                      ''');
                    } catch (_) {}
                  } else if (error != null) {
                    await controller.evaluateJavascript(source: '''
                      try {
                        if (typeof showAuthMessage === 'function') {
                          showAuthMessage('Google OAuth: $error', 'error');
                        }
                      } catch(_) {}
                    ''');
                  }
                  return NavigationActionPolicy.CANCEL;
                }
                return NavigationActionPolicy.ALLOW;
              },
            ),

            // Progress bar
            if (_progress < 1)
              Positioned(
                top: 0, left: 0, right: 0,
                child: LinearProgressIndicator(
                  value: _progress,
                  backgroundColor: Colors.transparent,
                  color: AppColors.teal,
                  minHeight: 2,
                ),
              ),

            // Nút reload nhỏ ở góc dưới phải (không che navbar website)
            Positioned(
              bottom: 20,
              right: 14,
              child: Container(
                decoration: BoxDecoration(
                  shape: BoxShape.circle,
                  boxShadow: [
                    BoxShadow(
                      color: Colors.black.withValues(alpha: 0.45),
                      blurRadius: 8,
                      offset: const Offset(0, 2),
                    ),
                  ],
                ),
                child: FloatingActionButton(
                  mini: true,
                  backgroundColor: const Color(0xFF0F1E3D),
                  foregroundColor: Colors.white,
                  elevation: 4,
                  onPressed: () => _controller?.reload(),
                  tooltip: 'Tải lại',
                  child: const Icon(Icons.refresh, size: 20),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}

