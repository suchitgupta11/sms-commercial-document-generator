# Uninstall — macOS

Run from the project/release folder:

```bash
chmod +x Uninstall.command
./Uninstall.command
```

The uninstaller:

- closes the running app,
- removes `/Applications/SMS Commercial Document Generator.app`,
- removes duplicate copies from `~/Applications` and Desktop,
- removes app support/cache/log/preference files,
- refreshes Launchpad.

Verify removal:

```bash
find /Applications "$HOME/Applications" "$HOME/Desktop" -maxdepth 2 -name "SMS Commercial Document Generator.app" -print
```

If nothing prints, the app bundle has been removed.
