#!/usr/bin/env bash
set -euo pipefail
APP_NAME="SMS Commercial Document Generator"
APP_PATH="/Applications/${APP_NAME}.app"
USER_APP_PATH="$HOME/Applications/${APP_NAME}.app"
DESKTOP_APP_PATH="$HOME/Desktop/${APP_NAME}.app"

printf '\nUninstalling %s...\n' "$APP_NAME"

printf 'Closing running application...\n'
osascript -e "quit app \"${APP_NAME}\"" >/dev/null 2>&1 || true
sleep 2
pkill -f "${APP_NAME}" >/dev/null 2>&1 || true
sleep 1

printf 'Removing app bundles...\n'
rm -rf "$APP_PATH" "$USER_APP_PATH" "$DESKTOP_APP_PATH"

printf 'Removing user data, caches, logs, and preferences...\n'
rm -rf "$HOME/Library/Application Support/${APP_NAME}"
rm -rf "$HOME/Library/Caches/${APP_NAME}"
rm -rf "$HOME/Library/Logs/${APP_NAME}"
rm -f "$HOME/Library/Preferences/com.smscontrols.commercialdocumentgenerator.plist"
rm -f "$HOME/Library/Preferences/com.smscontrols.documentgenerator.plist"

printf 'Refreshing Launchpad/Dock...\n'
defaults write com.apple.dock ResetLaunchPad -bool true >/dev/null 2>&1 || true
killall Dock >/dev/null 2>&1 || true

printf 'Verifying removal...\n'
FOUND="$(find /Applications "$HOME/Applications" "$HOME/Desktop" -maxdepth 2 -name "${APP_NAME}.app" -print 2>/dev/null || true)"
if [ -n "$FOUND" ]; then
  printf 'WARNING: Some app copies still exist:\n%s\n' "$FOUND"
  exit 1
fi

printf 'Uninstalled %s successfully.\n\n' "$APP_NAME"
