# Build AI Translator APK (release)
# Yêu cầu: Flutter SDK, Android SDK, JDK 17 (tự tải vào tools/jdk-17 nếu chưa có)
param(
    [string]$BaseUrl = "https://legacy-unpaid-sternum.ngrok-free.dev"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$AppDir = Join-Path $Root "app_web_view"
$JdkDir = Join-Path $Root "tools\jdk-17"
$ApkOut = Join-Path $AppDir "build\app\outputs\flutter-apk\app-release.apk"
$DownloadDir = Join-Path $Root "api_base\utils\download"

# Đọc version từ pubspec.yaml
$pubspec = Get-Content (Join-Path $AppDir "pubspec.yaml") -Raw
if ($pubspec -match 'version:\s*([\d.]+)\+(\d+)') {
    $versionName = $Matches[1]
    $versionCode = $Matches[2]
} else {
    $versionName = "1.0.0"
    $versionCode = "1"
}

# Cập nhật generated_base_url.dart
$genFile = Join-Path $AppDir "lib\config\generated_base_url.dart"
@"
// AUTO-GENERATED — cập nhật bằng scripts/build-apk.ps1 hoặc sửa thủ công.
const String kGeneratedBaseUrl = '$BaseUrl';
"@ | Set-Content -Path $genFile -Encoding UTF8

# JDK 17
if (-not (Test-Path "$JdkDir\bin\java.exe")) {
    Write-Host "Downloading Microsoft JDK 17..."
    $toolsDir = Join-Path $Root "tools"
    New-Item -ItemType Directory -Force -Path $toolsDir | Out-Null
    $zip = Join-Path $toolsDir "jdk17.zip"
    Invoke-WebRequest -Uri "https://aka.ms/download-jdk/microsoft-jdk-17.0.15-windows-x64.zip" -OutFile $zip -UseBasicParsing
    Expand-Archive -Path $zip -DestinationPath $toolsDir -Force
    $extracted = Get-ChildItem $toolsDir -Directory | Where-Object { $_.Name -like "jdk-17*" } | Select-Object -First 1
    if ($extracted -and $extracted.FullName -ne $JdkDir) {
        if (Test-Path $JdkDir) { Remove-Item $JdkDir -Recurse -Force }
        Rename-Item $extracted.FullName "jdk-17"
    }
    Remove-Item $zip -Force -ErrorAction SilentlyContinue
}

$env:JAVA_HOME = $JdkDir
$env:PATH = "$JdkDir\bin;$env:PATH"

Push-Location $AppDir
try {
    flutter pub get
    flutter build apk --release --dart-define=BASE_URL=$BaseUrl
} finally {
    Pop-Location
}

if (-not (Test-Path $ApkOut)) {
    throw "Build failed — APK not found at $ApkOut"
}

$destName = "AI_Translator_v$versionName.apk"
New-Item -ItemType Directory -Force -Path $DownloadDir | Out-Null
Copy-Item $ApkOut (Join-Path $DownloadDir $destName) -Force

Write-Host ""
Write-Host "APK built: $ApkOut"
Write-Host "Copied to: $DownloadDir\$destName"
Write-Host "Version: $versionName+$versionCode | Base URL: $BaseUrl"
