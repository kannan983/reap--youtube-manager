# REAP Multi-Language YouTube Channel Manager - Executables

## 🚀 Quick Start

Download the exe for your language from the **Releases** page
(built automatically by GitHub Actions, see `.github/workflows/build.yml`).

### How to Run:
1. Download `REAP_<Language>_Manager.exe` from Releases
2. Double-click it to run
3. In tab "1. Authentication", browse to your auth folder containing `oauth_client.json` and `api.txt`
   (these are NOT in this repo — keep them private)

### Available Languages:
- **English**: `REAP_English_Manager.exe`
- **Malayalam**: `REAP_Malayalam_Manager.exe`
- **Hindi**: `REAP_Hindi_Manager.exe`
- **Telugu**: `REAP_Telugu_Manager.exe`
- **Tamil**: `REAP_Tamil_Manager.exe`
- **Kannada**: `REAP_Kannada_Manager.exe`
- **Bengali**: `REAP_Bengali_Manager.exe`
- **Gujarati**: `REAP_Gujarati_Manager.exe`
- **Marathi**: `REAP_Marathi_Manager.exe`
- **Punjabi**: `REAP_Punjabi_Manager.exe`
- **Assamees**: `REAP_Assamees_Manager.exe`
- **Odia**: `REAP_Odia_Manager.exe`

## 🛡️ Windows Security Notice

If Windows Defender flags the executables as suspicious:

1. **Temporary Solution**: Click "More info" → "Run anyway"
2. **Permanent Solution**: Add the folder where you saved the exe to Windows Defender exclusions
3. **Alternative**: Run the source Python file directly with `python gen3_youtube_gem.py`

## 🔧 Technical Details

- **Built with**: PyInstaller --onefile
- **Self-contained**: Each executable includes all dependencies
- **No installation required**: Just run the .exe file
- **Size**: Typically 20-50 MB per executable

## 🐛 Troubleshooting

1. **Auth buttons disabled**: the selected auth folder must contain `oauth_client.json` and `api.txt`
2. **Logs**: written to `logs/` inside the selected auth folder
3. **Excel output**: saved to `Excel sheet/` inside the selected auth folder

## 🔨 Build Locally (Windows)

```
pip install -r "REAP English/requirements.txt" pyinstaller
pyinstaller --onefile --windowed --name REAP_English_Manager "REAP English/gen3_youtube_gem.py"
```

To publish a release with all 12 exes: `git tag v1.0.1 && git push origin v1.0.1`
