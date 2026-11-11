# REAP YouTube Manager

Desktop tool for REAP YouTube channels: fix chapter titles and descriptions in bulk, and export a playlist to an Excel sheet.
One app for all 12 channel languages. Open source.

## Download and run (Windows 10/11, 64-bit)

1. Download `REAP_YouTube_Manager.exe` from the **Releases** page of this repo.
2. Double-click it. Nothing to install (it includes everything it needs).
3. If Windows SmartScreen says "Windows protected your PC": click **More info** > **Run anyway**.

## How to use

1. **Channel Language** (tab 1): pick the language. The orange badge at the top, the window title and the status bar always show it.
   It is picked automatically when the auth folder name contains the language (e.g. `reap_Tamil`).
2. **Auth Folder**: browse to the folder for that channel. It must contain `oauth_client.json` and `api.txt`.
   Keep these files private, never upload them.
3. **Sign in**
   - First time: **Authorize (Browser)** and log in with the channel's Google account.
   - After that: **Authorize with Token** (or tick "Sign in automatically" and it happens at start).
   - A `token.json` from the old Python app also works.
4. **2. Rename & Describe**: load a playlist. Bold titles are not in `Chapter N - Topic` format; check/edit the
   Corrected and Description cells, then **Apply Changes**.
5. **3. Excel Export**: load a playlist, check the preview, **Generate Excel Sheet**.
   The file is saved in `<auth folder>\Excel sheet\`. You get a warning if the playlist's language code
   (e.g. `_ta` in the playlist title) does not match the selected language.
6. **Theme**: System / Light / Dark, top right.

### Saved login stops working every 7 days?

The Google Cloud OAuth app is in "Testing" mode. Fix once in Google Cloud Console:
**APIs & Services > OAuth consent screen > Publish app > In production**.

## Logs (for reporting problems)

- **Log** tab: live log, "Show debug details" for more, "Copy all".
- Log files: `%LOCALAPPDATA%\REAP_YouTube_Manager\logs` (the full path is shown in the Log and About tabs,
  with **Copy path** and **Open logs folder**). Files older than 30 days are deleted automatically.
- Settings: `%LOCALAPPDATA%\REAP_YouTube_Manager\settings.json`.

## Build from source

Requires the .NET 10 SDK. Works on Windows and Linux.

```
cd ReapManager
dotnet run                                           # run on this machine
dotnet run -- --selftest                             # logic self-test
dotnet publish -c Release -r win-x64 -o ../publish   # single Windows exe
```

The original Python version (`REAP <Language>/gen3_youtube_gem.py`) is kept for reference.
