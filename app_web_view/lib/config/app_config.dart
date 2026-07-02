import 'generated_base_url.dart';

/// Base URL của backend — truyền khi build:
/// flutter run --dart-define=BASE_URL=https://your-ngrok.ngrok-free.dev
class AppConfig {
  static const String baseUrl = String.fromEnvironment(
    'BASE_URL',
    defaultValue: kGeneratedBaseUrl,
  );

  static const String oauthScheme = 'aitranslator';

  static const Map<String, String> defaultHeaders = {
    'Accept': 'application/json',
    'ngrok-skip-browser-warning': 'true',
  };

  static String get apiAuth => '$baseUrl/api/auth';
  static String get dashboardUrl => '$baseUrl/dashboard';
  static String get authPageUrl => '$baseUrl/auth';
}
