import 'package:flutter/material.dart';

import '../models/user_profile.dart';
import '../theme/app_theme.dart';

class AppLogo extends StatelessWidget {
  const AppLogo({super.key, this.size = 72, this.showTitle = true});

  final double size;
  final bool showTitle;

  @override
  Widget build(BuildContext context) {
    return Column(
      mainAxisSize: MainAxisSize.min,
      children: [
        Container(
          width: size,
          height: size,
          decoration: BoxDecoration(
            borderRadius: BorderRadius.circular(size * 0.22),
            boxShadow: [
              BoxShadow(
                color: AppColors.teal.withValues(alpha: 0.25),
                blurRadius: 24,
                spreadRadius: 2,
              ),
            ],
          ),
          clipBehavior: Clip.antiAlias,
          child: Image.asset('assets/logo.png', fit: BoxFit.cover),
        ),
        if (showTitle) ...[
          const SizedBox(height: 16),
          const Text(
            'AI Translator',
            style: TextStyle(
              fontSize: 24,
              fontWeight: FontWeight.w700,
              color: AppColors.textPrimary,
            ),
          ),
          const SizedBox(height: 4),
          const Text(
            'Dịch thuật AI giữ định dạng',
            style: TextStyle(color: AppColors.textMuted, fontSize: 14),
          ),
        ],
      ],
    );
  }
}

class UserAvatar extends StatelessWidget {
  const UserAvatar({
    super.key,
    required this.profile,
    this.radius = 48,
    this.onTap,
    this.showEditBadge = false,
  });

  final UserProfile profile;
  final double radius;
  final VoidCallback? onTap;
  final bool showEditBadge;

  @override
  Widget build(BuildContext context) {
    final url = profile.avatarUrl;
    final initials = profile.initials;

    Widget avatar;
    if (url != null && url.isNotEmpty) {
      avatar = CircleAvatar(
        radius: radius,
        backgroundColor: AppColors.card,
        backgroundImage: NetworkImage(url),
        onBackgroundImageError: (_, __) {},
      );
    } else {
      avatar = CircleAvatar(
        radius: radius,
        backgroundColor: AppColors.teal.withValues(alpha: 0.15),
        foregroundColor: AppColors.teal,
        child: Text(
          initials,
          style: TextStyle(fontSize: radius * 0.55, fontWeight: FontWeight.w700),
        ),
      );
    }

    if (!showEditBadge) {
      return GestureDetector(onTap: onTap, child: avatar);
    }

    return GestureDetector(
      onTap: onTap,
      child: Stack(
        children: [
          avatar,
          Positioned(
            right: 0,
            bottom: 0,
            child: Container(
              padding: const EdgeInsets.all(6),
              decoration: BoxDecoration(
                color: AppColors.teal,
                shape: BoxShape.circle,
                border: Border.all(color: AppColors.bg, width: 2),
              ),
              child: Icon(Icons.camera_alt, size: radius * 0.35, color: AppColors.bg),
            ),
          ),
        ],
      ),
    );
  }
}
