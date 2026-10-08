#!/bin/sh
# Builds "XLSX Editor.app" into ./publish. Self-contained: runs on Macs without .NET installed.
#   ./build-mac.sh            build only
#   ./build-mac.sh --install  build and copy to /Applications
set -e
cd "$(dirname "$0")"

DOTNET=$(command -v dotnet || echo "$HOME/.dotnet/dotnet")
RID=${RID:-osx-$([ "$(uname -m)" = arm64 ] && echo arm64 || echo x64)} # override: RID=osx-x64 ./build-mac.sh
NAME="XLSX Editor"
APP="publish/$NAME.app"
VERSION=1.0.0

rm -rf publish
"$DOTNET" publish src/XlsxEditor.App -c Release -r "$RID" --self-contained -o publish/bin

mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
cp -R publish/bin/. "$APP/Contents/MacOS/"
cp src/XlsxEditor.App/Assets/AppIcon.icns "$APP/Contents/Resources/"
cat > "$APP/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key><string>$NAME</string>
  <key>CFBundleDisplayName</key><string>$NAME</string>
  <key>CFBundleIdentifier</key><string>com.jamezarkk.xlsxeditor</string>
  <key>CFBundleExecutable</key><string>XlsxEditor.App</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>$VERSION</string>
  <key>CFBundleVersion</key><string>$VERSION</string>
  <key>NSHighResolutionCapable</key><true/>
  <key>CFBundleIconFile</key><string>AppIcon</string>
  <!-- Lets Finder offer this app for .xlsx (Open With, Get Info → Change All). -->
  <key>CFBundleDocumentTypes</key>
  <array>
    <dict>
      <key>CFBundleTypeName</key><string>Excel Workbook</string>
      <key>CFBundleTypeRole</key><string>Editor</string>
      <key>LSHandlerRank</key><string>Default</string>
      <key>LSItemContentTypes</key><array><string>org.openxmlformats.spreadsheetml.sheet</string></array>
    </dict>
  </array>
</dict>
</plist>
EOF

# Ad-hoc signature: enough to run on this Mac (Apple Silicon refuses unsigned code).
# ponytail: not notarized; other Macs will need right-click → Open the first time.
codesign --force --deep --sign - "$APP"
echo "Built: $APP"

if [ "$1" = "--install" ]; then
  rm -rf "/Applications/$NAME.app"
  cp -R "$APP" /Applications/
  echo "Installed: /Applications/$NAME.app"
fi
