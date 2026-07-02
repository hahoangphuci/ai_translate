import 'dart:io';

import 'package:flutter_web_auth_2/flutter_web_auth_2.dart';

import '../config/app_config.dart';
import '../models/user_profile.dart';
import 'api_client.dart';
import 'session_storage.dart';

class AuthService {
  AuthService({ApiClient? client}) : _api = client ?? ApiClient();

  final ApiClient _api;
  UserProfile? _cachedProfile;

  UserProfile? get profile => _cachedProfile;

  Future<UserProfile> loginWithEmail({
    required String email,
    required String password,
  }) async {
    final data = await _api.postJson('/login', {
      'email': email.trim(),
      'password': password,
    });
    final token = data['access_token']?.toString();
    if (token == null || token.isEmpty) {
      throw ApiException('Không nhận được token đăng nhập');
    }
    final userJson = data['user'];
    UserProfile? user;
    if (userJson is Map<String, dynamic>) {
      user = UserProfile.fromJson(userJson);
    }
    _api.token = token;
    user ??= await fetchProfile();
    await SessionStorage.saveSession(token: token, user: user);
    return user;
  }

  /// Lấy Google OAuth URL để load trong WebView (không mở Chrome Custom Tab)
  Future<String> getGoogleAuthUrl() async {
    final data = await _api.postJson('/google/authorize', {
      'callback_scheme': AppConfig.oauthScheme,
    });
    final authUrl = data['auth_url']?.toString();
    if (authUrl == null || authUrl.isEmpty) {
      throw ApiException('Không lấy được URL đăng nhập Google');
    }
    return authUrl;
  }

  /// Xử lý token sau khi Google OAuth callback (dùng sau khi WebView bắt aitranslator://)
  Future<UserProfile> finishGoogleLogin(String token) async {
    _api.token = token;
    final profile = await fetchProfile();
    await SessionStorage.saveSession(token: token, user: profile);
    return profile;
  }

  /// Legacy: dùng Chrome Custom Tab (giữ lại phòng khi cần)
  Future<UserProfile> loginWithGoogle() async {
    final authUrl = await getGoogleAuthUrl();
    final result = await FlutterWebAuth2.authenticate(
      url: authUrl,
      callbackUrlScheme: AppConfig.oauthScheme,
    );
    final uri = Uri.parse(result);
    final error = uri.queryParameters['error'];
    if (error != null && error.isNotEmpty) {
      throw ApiException('Google OAuth: $error');
    }
    final token = uri.queryParameters['token'];
    if (token == null || token.isEmpty) {
      throw ApiException('Không nhận được token từ Google');
    }
    return finishGoogleLogin(token);
  }

  Future<UserProfile> fetchProfile() async {
    final data = await _api.getJson('/profile');
    final profile = UserProfile.fromJson(data);
    _cachedProfile = profile;
    await SessionStorage.saveUser(profile);
    return profile;
  }

  Future<UserProfile> updateProfile({String? name}) async {
    final body = <String, dynamic>{};
    if (name != null) body['name'] = name;
    final data = await _api.patchJson('/profile', body);
    final profile = UserProfile.fromJson(data);
    await SessionStorage.saveUser(profile);
    return profile;
  }

  Future<UserProfile> uploadAvatar(File file) async {
    final data = await _api.uploadAvatar(file);
    final userJson = data['user'];
    if (userJson is Map<String, dynamic>) {
      final profile = UserProfile.fromJson(userJson);
      await SessionStorage.saveUser(profile);
      return profile;
    }
    return fetchProfile();
  }

  Future<UserProfile?> restoreSession() async {
    final token = await SessionStorage.getToken();
    if (token == null || token.isEmpty) return null;
    _api.token = token;
    try {
      return await fetchProfile();
    } catch (_) {
      await logout();
      return null;
    }
  }

  Future<void> logout() async {
    _api.token = null;
    _cachedProfile = null;
    await SessionStorage.clear();
  }

  String? get token => _api.token;

  /// Dùng để set token từ bên ngoài (vd: sau khi nhận từ OAuth redirect)
  void apiSetToken(String token) => _api.token = token;
}
