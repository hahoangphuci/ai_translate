class UserProfile {
  const UserProfile({
    required this.id,
    required this.email,
    this.name,
    this.avatarUrl,
    this.plan = 'free',
    this.role = 'user',
    this.tokenBalance = 0,
  });

  final int id;
  final String email;
  final String? name;
  final String? avatarUrl;
  final String plan;
  final String role;
  final int tokenBalance;

  factory UserProfile.fromJson(Map<String, dynamic> json) {
    return UserProfile(
      id: json['id'] as int? ?? 0,
      email: (json['email'] ?? '').toString(),
      name: json['name']?.toString(),
      avatarUrl: json['avatar_url']?.toString(),
      plan: (json['plan'] ?? 'free').toString(),
      role: (json['role'] ?? 'user').toString(),
      tokenBalance: int.tryParse('${json['token_balance']}') ?? 0,
    );
  }

  Map<String, dynamic> toStorageJson() => {
        'id': id,
        'email': email,
        'name': name,
        'avatar_url': avatarUrl,
        'plan': plan,
        'role': role,
        'token_balance': tokenBalance,
      };

  String get displayName =>
      (name != null && name!.trim().isNotEmpty) ? name!.trim() : email;

  String get initials {
    final s = displayName.trim();
    if (s.isEmpty) return 'U';
    final parts = s.split(RegExp(r'\s+'));
    if (parts.length >= 2) {
      return '${parts.first[0]}${parts[1][0]}'.toUpperCase();
    }
    return s[0].toUpperCase();
  }

  UserProfile copyWith({
    String? name,
    String? avatarUrl,
    int? tokenBalance,
  }) {
    return UserProfile(
      id: id,
      email: email,
      name: name ?? this.name,
      avatarUrl: avatarUrl ?? this.avatarUrl,
      plan: plan,
      role: role,
      tokenBalance: tokenBalance ?? this.tokenBalance,
    );
  }
}
