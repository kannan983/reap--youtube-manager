## Download

**`REAP_YouTube_Manager.exe`** (below, under Assets) - Windows 10/11, 64-bit. One file, nothing to install.

1. Download the exe and put it anywhere (e.g. Desktop).
2. Double-click it. If Windows says "Windows protected your PC": **More info** > **Run anyway** (the exe is not code-signed).
3. Open the **Help** tab in the app, or read the [User guide](https://github.com/kannan983/reap--youtube-manager#user-guide).

## What's in v2.0.0

- Rewritten in C# (.NET 10 + Avalonia): one app for all 12 language channels instead of 12 separate exes.
- One auth folder per language, remembered; picking a language loads its folder and signs in with the saved token.
- Header always shows the language, the auth folder and the real YouTube channel name.
- Language mismatch warning before generating an Excel sheet.
- Fix: Excel save failed when the playlist description had a line break (`[Errno 22] Invalid argument`).
- Fix: saved-token expiry check error.
- Light / Dark / System theme, Help tab, live Log tab, logs folder path with Copy path / Open folder, About page.
- Reuses `token.json` from the old Python app.

MIT licensed. Source code and developer guide: https://github.com/kannan983/reap--youtube-manager
