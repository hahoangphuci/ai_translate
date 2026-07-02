import 'dart:io';

import 'package:flutter/material.dart';
import 'package:image_picker/image_picker.dart';

import '../models/user_profile.dart';
import '../services/auth_service.dart';
import '../theme/app_theme.dart';
import '../widgets/app_logo.dart';
import 'login_screen.dart';

class ProfileScreen extends StatefulWidget {
  const ProfileScreen({
    super.key,
    required this.authService,
    required this.profile,
  });

  final AuthService authService;
  final UserProfile profile;

  @override
  State<ProfileScreen> createState() => _ProfileScreenState();
}

class _ProfileScreenState extends State<ProfileScreen> {
  late TextEditingController _nameCtrl;
  late UserProfile _profile;
  bool _saving = false;
  bool _uploading = false;

  @override
  void initState() {
    super.initState();
    _profile = widget.profile;
    _nameCtrl = TextEditingController(text: _profile.displayName);
  }

  @override
  void dispose() {
    _nameCtrl.dispose();
    super.dispose();
  }

  Future<void> _pickAvatar() async {
    final picker = ImagePicker();
    final picked = await picker.pickImage(
      source: ImageSource.gallery,
      maxWidth: 1024,
      maxHeight: 1024,
      imageQuality: 85,
    );
    if (picked == null) return;

    setState(() => _uploading = true);
    try {
      final updated = await widget.authService.uploadAvatar(File(picked.path));
      setState(() => _profile = updated);
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Đã cập nhật ảnh đại diện')),
        );
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text(e.toString().replaceFirst('ApiException: ', ''))),
        );
      }
    } finally {
      if (mounted) setState(() => _uploading = false);
    }
  }

  Future<void> _saveName() async {
    final name = _nameCtrl.text.trim();
    if (name.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Tên không được để trống')),
      );
      return;
    }
    setState(() => _saving = true);
    try {
      final updated = await widget.authService.updateProfile(name: name);
      setState(() => _profile = updated);
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Đã lưu hồ sơ')),
        );
        Navigator.of(context).pop(updated);
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text(e.toString().replaceFirst('ApiException: ', ''))),
        );
      }
    } finally {
      if (mounted) setState(() => _saving = false);
    }
  }

  Future<void> _logout() async {
    await widget.authService.logout();
    if (!mounted) return;
    Navigator.of(context).pushAndRemoveUntil(
      MaterialPageRoute(builder: (_) => LoginScreen(authService: widget.authService)),
      (_) => false,
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Hồ sơ')),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(24),
        child: Column(
          children: [
            Stack(
              alignment: Alignment.center,
              children: [
                UserAvatar(
                  profile: _profile,
                  radius: 56,
                  showEditBadge: true,
                  onTap: _uploading ? null : _pickAvatar,
                ),
                if (_uploading)
                  const Positioned(
                    child: CircularProgressIndicator(color: AppColors.teal),
                  ),
              ],
            ),
            const SizedBox(height: 8),
            TextButton(
              onPressed: _uploading ? null : _pickAvatar,
              child: const Text('Đổi ảnh đại diện'),
            ),
            const SizedBox(height: 24),
            TextFormField(
              controller: _nameCtrl,
              decoration: const InputDecoration(
                labelText: 'Họ và tên',
                prefixIcon: Icon(Icons.person_outline),
              ),
            ),
            const SizedBox(height: 16),
            TextFormField(
              initialValue: _profile.email,
              readOnly: true,
              decoration: const InputDecoration(
                labelText: 'Email',
                prefixIcon: Icon(Icons.email_outlined),
              ),
            ),
            const SizedBox(height: 8),
            Align(
              alignment: Alignment.centerLeft,
              child: Text(
                'Gói: ${_profile.plan.toUpperCase()} · Token: ${_profile.tokenBalance}',
                style: const TextStyle(color: AppColors.textMuted, fontSize: 13),
              ),
            ),
            const SizedBox(height: 28),
            ElevatedButton(
              onPressed: _saving ? null : _saveName,
              child: _saving
                  ? const SizedBox(
                      width: 22,
                      height: 22,
                      child: CircularProgressIndicator(strokeWidth: 2),
                    )
                  : const Text('Lưu thay đổi'),
            ),
            const SizedBox(height: 12),
            OutlinedButton.icon(
              onPressed: _logout,
              icon: const Icon(Icons.logout, color: AppColors.error),
              label: const Text('Đăng xuất', style: TextStyle(color: AppColors.error)),
            ),
          ],
        ),
      ),
    );
  }
}
