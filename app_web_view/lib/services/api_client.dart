import 'dart:convert';
import 'dart:io';

import 'package:http/http.dart' as http;

import '../config/app_config.dart';

class ApiException implements Exception {
  ApiException(this.message, {this.statusCode});

  final String message;
  final int? statusCode;

  @override
  String toString() => message;
}

class ApiClient {
  ApiClient({this.token});

  String? token;

  Map<String, String> _headers({bool jsonBody = false}) {
    final h = Map<String, String>.from(AppConfig.defaultHeaders);
    if (jsonBody) h['Content-Type'] = 'application/json';
    if (token != null && token!.isNotEmpty) {
      h['Authorization'] = 'Bearer $token';
    }
    return h;
  }

  Future<Map<String, dynamic>> getJson(String path) async {
    final res = await http.get(
      Uri.parse('${AppConfig.apiAuth}$path'),
      headers: _headers(),
    );
    return _decode(res);
  }

  Future<Map<String, dynamic>> postJson(
    String path,
    Map<String, dynamic> body,
  ) async {
    final res = await http.post(
      Uri.parse('${AppConfig.apiAuth}$path'),
      headers: _headers(jsonBody: true),
      body: jsonEncode(body),
    );
    return _decode(res);
  }

  Future<Map<String, dynamic>> patchJson(
    String path,
    Map<String, dynamic> body,
  ) async {
    final res = await http.patch(
      Uri.parse('${AppConfig.apiAuth}$path'),
      headers: _headers(jsonBody: true),
      body: jsonEncode(body),
    );
    return _decode(res);
  }

  Future<Map<String, dynamic>> uploadAvatar(File file) async {
    final req = http.MultipartRequest(
      'POST',
      Uri.parse('${AppConfig.apiAuth}/profile/avatar'),
    );
    req.headers.addAll(_headers());
    req.files.add(await http.MultipartFile.fromPath('file', file.path));
    final streamed = await req.send();
    final res = await http.Response.fromStream(streamed);
    return _decode(res);
  }

  Map<String, dynamic> _decode(http.Response res) {
    Map<String, dynamic> data = {};
    try {
      final parsed = jsonDecode(res.body);
      if (parsed is Map<String, dynamic>) data = parsed;
    } catch (_) {}

    if (res.statusCode >= 200 && res.statusCode < 300) {
      return data;
    }

    final msg = data['message']?.toString() ??
        data['error']?.toString() ??
        'HTTP ${res.statusCode}';
    throw ApiException(msg, statusCode: res.statusCode);
  }
}
