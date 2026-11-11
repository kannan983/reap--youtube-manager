using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using ClosedXML.Excel;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.YouTube.v3;
using Google.Apis.YouTube.v3.Data;
using Ellipse = Avalonia.Controls.Shapes.Ellipse;

namespace ReapManager;

public record FolderEntry(string Language, string Path)
{
    public override string ToString() => $"{Language}  -  {Path}";
}

public record PlaylistInfo(string Id, string Title, string Description, long VideoCount);

public class RenameRow
{
    public string Original { get; init; } = "";
    public string Corrected { get; set; } = "";
    public string Description { get; set; } = "";
    public string VideoId { get; init; } = "";
    public bool Needs { get; init; }
}

public class MainWindow : Window
{
    const string OAuthClientFile = "oauth_client.json", ApiKeyFile = "api.txt", PyTokenFile = "token.json";
    const string StoreTokenFile = "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
    static readonly string[] Scopes = [YouTubeService.Scope.YoutubeForceSsl];
    static readonly string[] LanguageNames = Core.Languages.Select(l => l.Name).ToArray();
    static readonly IBrush Accent = new SolidColorBrush(Color.Parse("#1F4E79"));
    static readonly IBrush Orange = new SolidColorBrush(Color.Parse("#E67E22"));
    static readonly IBrush WarnBrush = new SolidColorBrush(Color.Parse("#D35400"));
    static readonly FontFamily Mono = new("Cascadia Mono, Consolas, DejaVu Sans Mono, Liberation Mono, monospace");

    readonly Settings _settings = Settings.Current;
    string? _authFolder, _apiKey;
    CancellationTokenSource? _authCts;
    YouTubeService? _ytOAuth, _ytKey;
    readonly Dictionary<string, PlaylistInfo> _playlists = new();
    List<ExcelRow>? _excelRows;
    List<RenameRow> _renameRows = new();
    bool _signedIn, _switching;
    string? _channelTitle;

    // header + status
    readonly TextBlock _langBadgeText = new() { Foreground = Brushes.White, FontWeight = FontWeight.Bold };
    readonly TextBlock _headerChannel = new() { Foreground = Brushes.White, FontSize = 13, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 4, 0, 0) };
    readonly ComboBox _themeCombo = new() { ItemsSource = new[] { "System", "Light", "Dark" }, MinWidth = 110 };
    readonly Ellipse _led = new() { Width = 14, Height = 14, Stroke = Brushes.DimGray, StrokeThickness = 1, Margin = new Thickness(0, 0, 8, 0) };
    readonly TextBlock _statusText = new() { Text = "Welcome!", VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
    readonly TextBlock _statusLang = new() { VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeight.Bold, Margin = new Thickness(12, 0, 0, 0) };
    // tab 1
    readonly ComboBox _langCombo = new() { ItemsSource = LanguageNames, MinWidth = 160 };
    readonly ComboBox _folderCombo = new() { HorizontalAlignment = HorizontalAlignment.Stretch, PlaceholderText = "No folder saved for this language yet - click Browse..." };
    readonly Button _browseBtn = new() { Content = "Browse..." };
    readonly Button _tokenBtn = new() { Content = "Authorize with Token", IsEnabled = false };
    readonly Button _browserBtn = new() { Content = "Authorize (Browser)", IsEnabled = false };
    readonly Button _abortBtn = new() { Content = "Sign out / Reset", IsEnabled = false };
    readonly CheckBox _autoSignIn = new() { Content = "Sign in automatically with the saved token (at start and when switching language)" };
    readonly TextBlock _authStatus = new() { Text = "Status: Awaiting folder.", FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap };
    readonly TextBlock _tokenInfo = new() { TextWrapping = TextWrapping.Wrap, Opacity = 0.8 };
    // tab 2
    readonly ComboBox _plRename = new() { IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Stretch };
    readonly Button _loadVideosBtn = new() { Content = "Load Videos", IsEnabled = false };
    readonly TextBlock _videoStatus = new() { Text = "Status: Select playlist & load.", Margin = new Thickness(0, 8), TextWrapping = TextWrapping.Wrap };
    readonly DataGrid _videoGrid = new() { CanUserResizeColumns = true, GridLinesVisibility = DataGridGridLinesVisibility.All };
    readonly TextBlock _renameCount = new() { Text = "0 videos to rename.", FontWeight = FontWeight.Bold, Foreground = WarnBrush, VerticalAlignment = VerticalAlignment.Center };
    readonly Button _applyBtn = new() { Content = "Apply Changes", IsEnabled = false };
    readonly ProgressBar _progress = new() { IsVisible = false, Width = 200, Minimum = 0, Maximum = 100 };
    // tab 3
    readonly ComboBox _plExcel = new() { IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Stretch };
    readonly Button _loadExcelBtn = new() { Content = "Load Video Info", IsEnabled = false };
    readonly TextBlock _excelStatus = new() { Text = "Status: Select playlist & load for preview.", Margin = new Thickness(0, 8), TextWrapping = TextWrapping.Wrap };
    readonly DataGrid _excelGrid = new() { IsReadOnly = true, CanUserResizeColumns = true, GridLinesVisibility = DataGridGridLinesVisibility.All };
    readonly Button _genExcelBtn = new() { Content = "Generate Excel Sheet", IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 8, 0, 0) };
    // tab 4: log
    readonly TextBox _logBox = new() { IsReadOnly = true, AcceptsReturn = true, TextWrapping = TextWrapping.NoWrap, FontFamily = Mono, FontSize = 12 };
    readonly CheckBox _verbose = new() { Content = "Show debug details" };
    readonly StringBuilder _logText = new();
    // shell
    readonly TabItem _tab2 = new() { Header = "2. Rename & Describe" };
    readonly TabItem _tab3 = new() { Header = "3. Excel Export" };

    public MainWindow()
    {
        MinWidth = 760; MinHeight = 520; Width = 1100; Height = 760; FontSize = 13;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        _langCombo.SelectedItem = LanguageNames.Contains(_settings.Language) ? _settings.Language : "English";
        _themeCombo.SelectedItem = _settings.Theme;
        _autoSignIn.IsChecked = _settings.AutoSignIn;
        _verbose.IsChecked = _settings.VerboseLog;
        UpdateLanguageDisplay();

        _langCombo.SelectionChanged += async (_, _) =>
        {
            if (_switching || _langCombo.SelectedItem is not string lang || lang == _settings.Language) return;
            _settings.Language = lang; _settings.Save();
            Log.Info($"Channel language set to {lang}.");
            await SwitchToLanguage(lang, autoSignIn: true);
        };
        _folderCombo.SelectionChanged += (_, _) =>
        {
            // picking another language's folder from the history = switch to that language
            if (!_switching && _folderCombo.SelectedItem is FolderEntry f && f.Language != _settings.Language)
                _langCombo.SelectedItem = f.Language;
        };
        _themeCombo.SelectionChanged += (_, _) =>
        {
            _settings.Theme = (string)_themeCombo.SelectedItem!; _settings.Save();
            Application.Current!.RequestedThemeVariant = App.ThemeFromName(_settings.Theme);
            Log.Debug($"Theme set to {_settings.Theme}.");
        };
        _autoSignIn.IsCheckedChanged += (_, _) => { _settings.AutoSignIn = _autoSignIn.IsChecked == true; _settings.Save(); };
        _verbose.IsCheckedChanged += (_, _) => { _settings.VerboseLog = _verbose.IsChecked == true; _settings.Save(); RebuildLogView(); };

        _browseBtn.Click += async (_, _) => await BrowseFolder();
        _tokenBtn.Click += async (_, _) => await StartAuth(browser: false);
        _browserBtn.Click += async (_, _) => await StartAuth(browser: true);
        _abortBtn.Click += (_, _) => { Log.Info("Sign out / reset clicked."); _authCts?.Cancel(); SetAuthState(false, "Signed out."); };
        _loadVideosBtn.Click += async (_, _) => await LoadRenameVideos();
        _applyBtn.Click += async (_, _) => await ApplyChanges();
        _loadExcelBtn.Click += async (_, _) => await LoadExcelPreview();
        _genExcelBtn.Click += async (_, _) => await GenerateExcel();

        BuildGrids();
        _tab2.Content = BuildRenameTab();
        _tab3.Content = BuildExcelTab();
        Content = new DockPanel
        {
            Children =
            {
                Dock(BuildHeader(), Avalonia.Controls.Dock.Top),
                Dock(BuildStatusBar(), Avalonia.Controls.Dock.Bottom),
                new TabControl
                {
                    Margin = new Thickness(8, 4, 8, 0),
                    Items =
                    {
                        new TabItem { Header = "1. Authentication", Content = BuildAuthTab() },
                        _tab2, _tab3,
                        new TabItem { Header = "Help", Content = BuildHelpTab() },
                        new TabItem { Header = "Log", Content = BuildLogTab() },
                        new TabItem { Header = "About", Content = BuildAboutTab() },
                    }
                }
            }
        };

        Log.Written += OnLogWritten;
        RebuildLogView();
        SetAuthState(false);
        Opened += async (_, _) => { FitToScreen(); await RestoreLastSession(); };
    }

    // ================= layout =================
    Control BuildHeader()
    {
        var badge = new Border { Background = Orange, CornerRadius = new CornerRadius(12), Padding = new Thickness(12, 4), VerticalAlignment = VerticalAlignment.Center, Child = _langBadgeText };
        var themeLabel = new TextBlock { Text = "Theme:", Foreground = Brushes.White, VerticalAlignment = VerticalAlignment.Center };
        var right = Row(badge, themeLabel, _themeCombo);
        right.VerticalAlignment = VerticalAlignment.Center;
        var left = new StackPanel
        {
            Children =
            {
                new TextBlock { Text = AppInfo.Name, FontSize = 20, FontWeight = FontWeight.Bold, Foreground = Brushes.White },
                new TextBlock { Text = "Rename videos and export playlists to Excel for REAP YouTube channels", FontSize = 12, Foreground = new SolidColorBrush(Color.Parse("#D0E4F5")), TextWrapping = TextWrapping.Wrap },
                _headerChannel,
            }
        };
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), ColumnSpacing = 12 };
        Grid.SetColumn(right, 1);
        g.Children.Add(left); g.Children.Add(right);
        return new Border { Background = Accent, Padding = new Thickness(16, 10), Child = g };
    }

    Control BuildStatusBar()
    {
        var p = new DockPanel { Margin = new Thickness(10, 6) };
        p.Children.Add(Dock(_led, Avalonia.Controls.Dock.Left));
        p.Children.Add(Dock(_statusLang, Avalonia.Controls.Dock.Right));
        p.Children.Add(_statusText);
        return new Border { BorderBrush = Brushes.Gray, BorderThickness = new Thickness(0, 1, 0, 0), Child = p };
    }

    Control BuildAuthTab() => new ScrollViewer
    {
        Content = new StackPanel
        {
            Margin = new Thickness(12), Spacing = 10,
            Children =
            {
                Wrap(new TextBlock { Text = "Channel Language:", VerticalAlignment = VerticalAlignment.Center }, _langCombo,
                     new TextBlock { Text = "Each language has its own auth folder and YouTube channel.", Opacity = 0.7, VerticalAlignment = VerticalAlignment.Center }),
                Grid3("Auth Folder:", _folderCombo, _browseBtn),
                new TextBlock { Text = "The list remembers every folder you have used, with its language. Picking one switches to that language.", Opacity = 0.7, TextWrapping = TextWrapping.Wrap },
                Wrap(_tokenBtn, _browserBtn, _abortBtn),
                _autoSignIn,
                _authStatus,
                _tokenInfo,
                Panel("New here?", "Open the Help tab for the full step-by-step guide (first-time setup for all 12 languages, everyday use, and fixing problems)."),
            }
        }
    };

    Control BuildRenameTab() => new DockPanel
    {
        Margin = new Thickness(12),
        Children =
        {
            Dock(Grid3("Select Playlist:", _plRename, _loadVideosBtn), Avalonia.Controls.Dock.Top),
            Dock(_videoStatus, Avalonia.Controls.Dock.Top),
            Dock(Grid3(null, _renameCount, Row(_applyBtn, _progress)), Avalonia.Controls.Dock.Bottom),
            _videoGrid
        }
    };

    Control BuildExcelTab() => new DockPanel
    {
        Margin = new Thickness(12),
        Children =
        {
            Dock(Grid3("Select Playlist:", _plExcel, _loadExcelBtn), Avalonia.Controls.Dock.Top),
            Dock(_excelStatus, Avalonia.Controls.Dock.Top),
            Dock(_genExcelBtn, Avalonia.Controls.Dock.Bottom),
            _excelGrid
        }
    };

    static Control BuildHelpTab()
    {
        var p = new StackPanel { Margin = new Thickness(12), Spacing = 12, MaxWidth = 900, HorizontalAlignment = HorizontalAlignment.Left };
        p.Children.Add(new TextBlock { Text = "How to use REAP YouTube Manager", FontSize = 20, FontWeight = FontWeight.Bold });
        foreach (var (title, body) in HelpText.Sections) p.Children.Add(Panel(title, body));
        return new ScrollViewer { Content = p };
    }

    Control BuildLogTab()
    {
        var copyAll = new Button { Content = "Copy all" };
        var clear = new Button { Content = "Clear view" };
        copyAll.Click += async (_, _) => { await Clipboard!.SetTextAsync(_logBox.Text ?? ""); Status("Log copied to clipboard."); };
        clear.Click += (_, _) => { _logText.Clear(); _logBox.Text = ""; };
        var tools = Wrap(_verbose, copyAll, clear);
        tools.Margin = new Thickness(0, 0, 0, 8);
        return new DockPanel
        {
            Margin = new Thickness(12),
            Children =
            {
                Dock(LogsPathRow(), Avalonia.Controls.Dock.Top),
                Dock(tools, Avalonia.Controls.Dock.Top),
                _logBox
            }
        };
    }

    Control BuildAboutTab()
    {
        var gh = new Button { Content = "Open on GitHub" };
        gh.Click += async (_, _) => await Launcher.LaunchUriAsync(new Uri(AppInfo.RepoUrl));
        var repo = new TextBox { Text = AppInfo.RepoUrl, IsReadOnly = true };
        return new ScrollViewer
        {
            Content = new StackPanel
            {
                Margin = new Thickness(12), Spacing = 10,
                Children =
                {
                    new TextBlock { Text = AppInfo.Name, FontSize = 22, FontWeight = FontWeight.Bold },
                    new TextBlock { Text = $"Version {AppInfo.Version}", Opacity = 0.8 },
                    new TextBlock { TextWrapping = TextWrapping.Wrap, Text =
                        "A desktop tool for REAP YouTube channels: fix chapter titles and descriptions in bulk, " +
                        "and export a playlist to an Excel sheet (course code, chapter, URL, title, description, order, language)." },
                    new TextBlock { Text = "Open source", FontWeight = FontWeight.Bold, Margin = new Thickness(0, 8, 0, 0) },
                    new TextBlock { TextWrapping = TextWrapping.Wrap, Text = "The source code is on GitHub. Bug reports and improvements are welcome." },
                    Grid3(null, repo, gh),
                    new TextBlock { Text = "Logs (send these when reporting a problem)", FontWeight = FontWeight.Bold, Margin = new Thickness(0, 8, 0, 0) },
                    LogsPathRow(),
                    new TextBlock { Text = "Settings file", FontWeight = FontWeight.Bold, Margin = new Thickness(0, 8, 0, 0) },
                    new TextBox { Text = Path.Combine(AppInfo.DataDir, "settings.json"), IsReadOnly = true, FontFamily = Mono },
                    new TextBlock { Text = "System", FontWeight = FontWeight.Bold, Margin = new Thickness(0, 8, 0, 0) },
                    new SelectableTextBlock { Text = $"{RuntimeInformation.OSDescription} ({RuntimeInformation.OSArchitecture})\n.NET {Environment.Version}", Opacity = 0.8 },
                }
            }
        };
    }

    /// Full logs folder path in a selectable box + Copy path + Open folder.
    Control LogsPathRow()
    {
        var box = new TextBox { Text = Log.Dir, IsReadOnly = true, FontFamily = Mono };
        var copy = new Button { Content = "Copy path" };
        var open = new Button { Content = "Open logs folder" };
        copy.Click += async (_, _) => { await Clipboard!.SetTextAsync(Log.Dir); Status("Logs folder path copied."); };
        open.Click += async (_, _) =>
        {
            if (!await Launcher.LaunchDirectoryInfoAsync(new DirectoryInfo(Log.Dir)))
                await Message("Logs folder", $"Could not open the folder automatically. Paste this path into File Explorer:\n{Log.Dir}");
        };
        var g = Grid3("Logs folder:", box, Row(copy, open));
        g.Margin = new Thickness(0, 0, 0, 8);
        return g;
    }

    void BuildGrids()
    {
        var star = new DataGridLength(1, DataGridLengthUnitType.Star);
        _videoGrid.Columns.Add(new DataGridTemplateColumn
        {
            Header = "Original", Width = star, IsReadOnly = true,
            CellTemplate = new FuncDataTemplate<RenameRow>((r, _) => new TextBlock
            {
                Text = r?.Original, Margin = new Thickness(6, 0), VerticalAlignment = VerticalAlignment.Center,
                FontWeight = r?.Needs == true ? FontWeight.Bold : FontWeight.Normal
            })
        });
        _videoGrid.Columns.Add(new DataGridTextColumn { Header = "Corrected (editable)", Binding = new Binding(nameof(RenameRow.Corrected)), Width = star });
        _videoGrid.Columns.Add(new DataGridTextColumn { Header = "Description (editable)", Binding = new Binding(nameof(RenameRow.Description)), Width = star });

        string[] props = [nameof(ExcelRow.CourseCode), nameof(ExcelRow.ChapterName), nameof(ExcelRow.YouTubeUrl), nameof(ExcelRow.VideoTitle),
                          nameof(ExcelRow.VideoDescription), nameof(ExcelRow.OrderNo), nameof(ExcelRow.LanguageCode)];
        for (int i = 0; i < props.Length; i++)
            _excelGrid.Columns.Add(new DataGridTextColumn { Header = Core.ExcelHeaders[i], Binding = new Binding(props[i]) });
    }

    static T Dock<T>(T c, Dock d) where T : Control { DockPanel.SetDock(c, d); return c; }
    static StackPanel Row(params Control[] cs)
    {
        var p = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        p.Children.AddRange(cs);
        return p;
    }
    static WrapPanel Wrap(params Control[] cs)
    {
        var p = new WrapPanel { ItemSpacing = 8, LineSpacing = 8 };
        p.Children.AddRange(cs);
        return p;
    }
    static Grid Grid3(string? label, Control middle, Control right)
    {
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"), ColumnSpacing = 8 };
        if (label != null) g.Children.Add(new TextBlock { Text = label, VerticalAlignment = VerticalAlignment.Center });
        Grid.SetColumn(middle, 1); Grid.SetColumn(right, 2);
        g.Children.Add(middle); g.Children.Add(right);
        return g;
    }
    static Border Panel(string title, string body) => new()
    {
        BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4), Padding = new Thickness(10),
        Child = new StackPanel
        {
            Spacing = 4,
            Children = { new TextBlock { Text = title, FontWeight = FontWeight.Bold }, new SelectableTextBlock { Text = body, TextWrapping = TextWrapping.Wrap } }
        }
    };

    /// Size the window to the laptop/monitor it opens on (85% of the work area, capped).
    void FitToScreen()
    {
        var screen = Screens.ScreenFromWindow(this) ?? Screens.Primary;
        if (screen == null) return;
        var wa = screen.WorkingArea; var s = screen.Scaling;
        Width = Math.Clamp(wa.Width / s * 0.85, MinWidth, 1300);
        Height = Math.Clamp(wa.Height / s * 0.85, MinHeight, 900);
        Position = new PixelPoint(wa.X + (int)((wa.Width - Width * s) / 2), wa.Y + (int)((wa.Height - Height * s) / 2));
        Log.Debug($"Screen {wa.Width}x{wa.Height} @ {s:0.##}x, window {Width:0}x{Height:0}.");
    }

    void UpdateLanguageDisplay()
    {
        var lang = _settings.Language;
        var folder = _authFolder == null ? "not set" : System.IO.Path.GetFileName(_authFolder.TrimEnd('/', '\\'));
        var channel = _channelTitle ?? "not signed in";
        Title = $"{AppInfo.Name} - {lang} channel";
        _langBadgeText.Text = $"Language: {lang} ({Core.LanguageCode(lang)})";
        _headerChannel.Text = $"Auth folder: {folder}     YouTube channel: {channel}";
        _statusLang.Text = $"{lang}  |  {channel}";
    }

    void Status(string s) => _statusText.Text = s;

    async Task<bool> Message(string title, string text, bool yesNo = false)
    {
        bool result = false;
        var dlg = new Window
        {
            Title = title, Width = 480, SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner, CanResize = false
        };
        var ok = new Button { Content = yesNo ? "Yes" : "OK", MinWidth = 70, IsDefault = true };
        var no = new Button { Content = "No", MinWidth = 70, IsVisible = yesNo, IsCancel = true };
        ok.Click += (_, _) => { result = true; dlg.Close(); };
        no.Click += (_, _) => dlg.Close();
        var buttons = Row(ok, no);
        buttons.HorizontalAlignment = HorizontalAlignment.Right;
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(16), Spacing = 14,
            Children = { new SelectableTextBlock { Text = text, TextWrapping = TextWrapping.Wrap }, buttons }
        };
        await dlg.ShowDialog(this);
        return result;
    }

    // ================= log view =================
    void OnLogWritten(LogLevel level, string line) => Dispatcher.UIThread.Post(() =>
    {
        if (level == LogLevel.Debug && _verbose.IsChecked != true) return;
        _logText.AppendLine(line);
        if (_logText.Length > 400_000) RebuildLogView(); // keep memory bounded
        else _logBox.Text = _logText.ToString();
        _logBox.CaretIndex = _logBox.Text?.Length ?? 0;
    });

    void RebuildLogView()
    {
        bool verbose = _verbose.IsChecked == true;
        _logText.Clear();
        foreach (var (level, line) in Log.Snapshot())
            if (verbose || level != LogLevel.Debug) _logText.AppendLine(line);
        _logBox.Text = _logText.ToString();
        _logBox.CaretIndex = _logBox.Text.Length;
    }

    // ================= tab 1: auth =================
    Task RestoreLastSession() => SwitchToLanguage(_settings.Language, autoSignIn: true);

    /// Load the saved auth folder for a language (signing out of the previous channel first).
    async Task SwitchToLanguage(string lang, bool autoSignIn)
    {
        _authCts?.Cancel();
        if (_signedIn) SetAuthState(false, $"Switched to {lang}.");
        _authFolder = _settings.AuthFolders.GetValueOrDefault(lang);
        if (_authFolder != null && !Directory.Exists(_authFolder))
        {
            Log.Warn($"Saved {lang} auth folder no longer exists: {_authFolder}");
            _authFolder = null;
        }
        Log.Info(_authFolder == null ? $"No auth folder saved for {lang} yet." : $"{lang} auth folder: {_authFolder}");
        RefreshFolderList();
        UpdateAuthButtons();
        UpdateLanguageDisplay();
        if (autoSignIn && _settings.AutoSignIn && _tokenBtn.IsEnabled) await StartAuth(browser: false);
    }

    void RefreshFolderList()
    {
        _switching = true;
        var items = Core.Languages.Where(l => _settings.AuthFolders.ContainsKey(l.Name))
                                  .Select(l => new FolderEntry(l.Name, _settings.AuthFolders[l.Name])).ToList();
        _folderCombo.ItemsSource = items;
        _folderCombo.SelectedItem = items.FirstOrDefault(i => i.Language == _settings.Language);
        _langCombo.SelectedItem = _settings.Language;
        _switching = false;
    }

    async Task BrowseFolder()
    {
        var picked = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = $"Select the {_settings.Language} auth folder" });
        var path = picked.FirstOrDefault()?.TryGetLocalPath();
        if (path == null) return;

        // The folder name decides the language when it contains one (e.g. reap_Tamil); otherwise the selected language.
        var lang = Core.DetectLanguage(path) ?? _settings.Language;
        if (lang != _settings.Language)
            Log.Info($"Folder name says {lang}, switching language from {_settings.Language} to {lang}.");
        foreach (var other in _settings.AuthFolders.Where(kv => kv.Value == path && kv.Key != lang).Select(kv => kv.Key).ToList())
        {
            Log.Info($"Folder was saved for {other}; moving it to {lang}.");
            _settings.AuthFolders.Remove(other);
        }
        _settings.AuthFolders[lang] = path;
        _settings.Language = lang;
        _settings.Save();
        Log.Info($"Saved {lang} auth folder: {path}");
        await SwitchToLanguage(lang, autoSignIn: false);
    }

    void UpdateAuthButtons()
    {
        bool oauthOk = false, apiOk = false, pyToken = false, storeToken = false;
        _apiKey = null;
        if (_authFolder != null)
        {
            oauthOk = File.Exists(Path.Combine(_authFolder, OAuthClientFile));
            pyToken = File.Exists(Path.Combine(_authFolder, PyTokenFile));
            storeToken = File.Exists(Path.Combine(_authFolder, StoreTokenFile));
            try
            {
                var m = Regex.Match(File.ReadAllText(Path.Combine(_authFolder, ApiKeyFile)), @"""API Key""\s*:\s*""(.*?)""");
                if (m.Success) { _apiKey = m.Groups[1].Value; apiOk = true; }
            }
            catch (Exception e) { Log.Debug($"api.txt not readable: {e.Message}"); }
            Log.Debug($"Auth folder check: oauth_client.json={oauthOk}, api.txt key={apiOk}, token.json={pyToken}, saved token={storeToken}");
        }
        bool filesOk = oauthOk && apiOk;
        _browserBtn.IsEnabled = filesOk && !_signedIn;
        _tokenBtn.IsEnabled = filesOk && (pyToken || storeToken) && !_signedIn;
        _abortBtn.IsEnabled = _authFolder != null;
        if (!_signedIn)
            _authStatus.Text = _authFolder == null ? "Status: Awaiting folder."
                : !oauthOk ? "Status: Error! oauth_client.json is missing in this folder."
                : !apiOk ? "Status: Error! api.txt is missing or has no \"API Key\"."
                : "Status: Ready. Use Token or Browser.";
        _tokenInfo.Text = _authFolder == null ? ""
            : storeToken ? "Saved login: found (from this app)."
            : pyToken ? "Saved login: found (token.json from the old Python app)."
            : "Saved login: none yet. Use 'Authorize (Browser)' once.";
    }

    /// Seeds the Google token store from the old Python app's token.json, so existing logins keep working.
    TokenResponse? ImportPythonToken()
    {
        var path = Path.Combine(_authFolder!, PyTokenFile);
        if (!File.Exists(path)) return null;
        using var doc = JsonDocument.Parse(File.ReadAllText(path));
        if (!doc.RootElement.TryGetProperty("refresh_token", out var rt)) return null;
        Log.Info("Using refresh token from token.json (old Python app).");
        return new TokenResponse { RefreshToken = rt.GetString() };
    }

    async Task StartAuth(bool browser)
    {
        if (_authFolder == null) return;
        var mode = browser ? "browser" : "token";
        _browserBtn.IsEnabled = _tokenBtn.IsEnabled = false; _abortBtn.IsEnabled = true;
        _authStatus.Text = $"Status: Starting '{mode}' auth..."; Status("Auth in progress...");
        Log.Info($"Starting auth (mode: {mode}).");
        _authCts?.Dispose();
        _authCts = new CancellationTokenSource();
        try
        {
            var secrets = GoogleClientSecrets.FromFile(Path.Combine(_authFolder, OAuthClientFile)).Secrets;
            Log.Debug($"OAuth client id: {secrets.ClientId}");
            var store = new FileDataStore(_authFolder, fullPath: true);
            UserCredential cred;
            string msg;
            if (browser)
            {
                Log.Info("Opening browser for Google sign-in...");
                cred = await GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, Scopes, "user", _authCts.Token, store);
                msg = "Auth successful.";
            }
            else
            {
                var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer { ClientSecrets = secrets, Scopes = Scopes, DataStore = store });
                var token = await flow.LoadTokenAsync("user", _authCts.Token) ?? ImportPythonToken();
                if (token?.RefreshToken == null) { await AuthFailed("Token invalid/missing. Use 'Authorize (Browser)'."); return; }
                cred = new UserCredential(flow, "user", token);
                if (!await cred.RefreshTokenAsync(_authCts.Token)) { await AuthFailed("Token refresh failed. Use 'Authorize (Browser)'."); return; }
                msg = "Token loaded.";
            }
            Log.Debug($"Access token valid until {cred.Token.IssuedUtc.AddSeconds(cred.Token.ExpiresInSeconds ?? 0):u} (refreshed automatically).");
            DisposeServices();
            _ytOAuth = new YouTubeService(new BaseClientService.Initializer { HttpClientInitializer = cred, ApplicationName = AppInfo.Name });
            _ytKey = new YouTubeService(new BaseClientService.Initializer { ApiKey = _apiKey, ApplicationName = AppInfo.Name });
            _channelTitle = await FetchChannelTitle();
            _authStatus.Text = $"Status: {msg} Signed in to YouTube channel: {_channelTitle ?? "(unknown)"}";
            SetAuthState(true, msg);
            await LoadPlaylists();
        }
        catch (OperationCanceledException) { Log.Info("Auth aborted."); SetAuthState(false, "Aborted."); }
        catch (TokenResponseException e) when (e.Error?.Error == "invalid_grant")
        {
            Log.Error($"Saved token rejected by Google: {e.Error.ErrorDescription}");
            await AuthFailed("The saved login expired or was revoked.\n\nClick 'Authorize (Browser)' to sign in again.\n\n" +
                             "If this happens about every 7 days, the Google Cloud OAuth app is in 'Testing' mode: " +
                             "publish it to 'In production' (APIs & Services > OAuth consent screen).");
        }
        catch (Exception e) { Log.Error($"Auth error: {e}"); await AuthFailed($"Auth failed: {e.Message}"); }
    }

    async Task AuthFailed(string msg)
    {
        var first = msg.Split('\n')[0];
        SetAuthState(false, first);
        _authStatus.Text = $"Status: {first}";
        await Message("Sign-in failed", msg);
    }

    async Task<string?> FetchChannelTitle()
    {
        try
        {
            var req = _ytOAuth!.Channels.List("snippet");
            req.Mine = true;
            var title = (await req.ExecuteAsync()).Items?.FirstOrDefault()?.Snippet.Title;
            Log.Info($"Signed-in YouTube channel: {title ?? "(none found)"}");
            return title;
        }
        catch (Exception e) { Log.Warn($"Could not read channel name: {e.Message}"); return null; }
    }

    void SetAuthState(bool ok, string message = "")
    {
        _signedIn = ok;
        if (!ok) _channelTitle = null;
        UpdateLanguageDisplay();
        _led.Fill = new SolidColorBrush(ok ? Color.Parse("#14C814") : Color.Parse("#D01414"));
        _tab2.IsEnabled = _tab3.IsEnabled = ok;
        if (ok)
        {
            _browserBtn.IsEnabled = _tokenBtn.IsEnabled = false; _abortBtn.IsEnabled = true;
            Status($"Signed in ({message})"); Log.Info($"Authenticated! ({message})");
            return;
        }
        DisposeServices();
        _playlists.Clear(); _renameRows = new(); _excelRows = null;
        _plRename.ItemsSource = _plExcel.ItemsSource = null;
        _videoGrid.ItemsSource = _excelGrid.ItemsSource = null;
        Status(message == "" ? "Not signed in." : $"Not signed in. ({message})");
        UpdateAuthButtons();
    }

    void DisposeServices()
    {
        _ytOAuth?.Dispose(); _ytKey?.Dispose();
        _ytOAuth = _ytKey = null;
    }

    protected override void OnClosed(EventArgs e)
    {
        Log.Written -= OnLogWritten;
        _authCts?.Cancel(); _authCts?.Dispose();
        DisposeServices();
        Log.Info("Main window closed.");
        base.OnClosed(e);
    }

    // ================= playlists / videos =================
    record PlaylistItemView(string Id, string Text) { public override string ToString() => Text; }

    async Task LoadPlaylists()
    {
        Status("Loading playlists..."); Log.Info("Fetching playlists...");
        _plRename.IsEnabled = _plExcel.IsEnabled = _loadVideosBtn.IsEnabled = _loadExcelBtn.IsEnabled = false;
        try
        {
            _playlists.Clear();
            string? page = null;
            do
            {
                var req = _ytOAuth!.Playlists.List("snippet,contentDetails");
                req.Mine = true; req.MaxResults = 50; req.PageToken = page;
                var r = await req.ExecuteAsync();
                Log.Debug($"Playlists page: {r.Items.Count} items, next={r.NextPageToken ?? "-"}");
                foreach (var i in r.Items)
                    _playlists[i.Id] = new PlaylistInfo(i.Id, i.Snippet.Title, i.Snippet.Description ?? "", i.ContentDetails.ItemCount ?? 0);
                page = r.NextPageToken;
            } while (page != null);

            var items = _playlists.Values.Select(p =>
            {
                var d = p.Description.Length > 30 ? p.Description[..30] + "..." : p.Description;
                return new PlaylistItemView(p.Id, $"{p.Title} - {d.ReplaceLineEndings(" ")} - {p.VideoCount} videos");
            }).ToList();
            _plRename.ItemsSource = items; _plExcel.ItemsSource = items;
            Log.Info($"Loaded {items.Count} playlists.");
            if (items.Count == 0) { await Message("Info", "No playlists found on this channel."); return; }
            _plRename.SelectedIndex = _plExcel.SelectedIndex = 0;
            _plRename.IsEnabled = _plExcel.IsEnabled = _loadVideosBtn.IsEnabled = _loadExcelBtn.IsEnabled = true;
            Status($"Signed in. {items.Count} playlists loaded.");
        }
        catch (Exception e) { Log.Error($"Error fetching playlists: {e}"); Status($"Error loading playlists: {e.Message}"); }
    }

    async Task<List<VideoInfo>> FetchVideos(string playlistId)
    {
        Log.Info($"Fetching videos for {playlistId}...");
        var videos = new List<VideoInfo>();
        string? page = null;
        do
        {
            var req = _ytKey!.PlaylistItems.List("snippet");
            req.PlaylistId = playlistId; req.MaxResults = 50; req.PageToken = page;
            var r = await req.ExecuteAsync();
            Log.Debug($"Videos page: {r.Items.Count} items, next={r.NextPageToken ?? "-"}");
            foreach (var i in r.Items)
                videos.Add(new VideoInfo(i.Snippet.ResourceId.VideoId, i.Snippet.Title, i.Snippet.Description ?? ""));
            page = r.NextPageToken;
        } while (page != null);
        Log.Info($"Fetched {videos.Count} videos.");
        return videos;
    }

    // ================= tab 2: rename =================
    async Task LoadRenameVideos()
    {
        if (_plRename.SelectedItem is not PlaylistItemView pl) return;
        Status("Loading videos..."); _videoGrid.ItemsSource = null;
        _loadVideosBtn.IsEnabled = _applyBtn.IsEnabled = false;
        try
        {
            var videos = await FetchVideos(pl.Id);
            _renameRows = videos.Select(v =>
            {
                var (corrected, desc, needs) = Core.CorrectTitle(v.Title);
                if (needs) Log.Debug($"Needs rename: '{v.Title}' -> '{corrected}'");
                return new RenameRow { Original = v.Title, Corrected = corrected, Description = desc, VideoId = v.VideoId, Needs = needs };
            }).ToList();
            _videoGrid.ItemsSource = _renameRows;
            int count = _renameRows.Count(r => r.Needs);
            _videoStatus.Text = $"Loaded {videos.Count} videos. Bold titles will be renamed; you can edit the Corrected and Description cells.";
            _renameCount.Text = $"{count} videos to rename.";
            _applyBtn.IsEnabled = count > 0;
            Status("Videos loaded.");
        }
        catch (Exception e) { Log.Error($"Error fetching videos: {e}"); _videoStatus.Text = $"Status: No videos / error: {e.Message}"; }
        finally { _loadVideosBtn.IsEnabled = true; }
    }

    async Task ApplyChanges()
    {
        var todo = _renameRows.Where(r => r.Needs).ToList();
        if (todo.Count == 0) { await Message("Info", "No videos marked."); return; }
        if (!await Message("Confirm", $"Update {todo.Count} videos on the {_settings.Language} channel \"{_channelTitle}\"?", yesNo: true)) return;

        _applyBtn.IsEnabled = false; _progress.Value = 0; _progress.IsVisible = true;
        Log.Info($"Starting update: {todo.Count} videos.");
        int ok = 0, fail = 0;
        for (int i = 0; i < todo.Count; i++)
        {
            var row = todo[i];
            try
            {
                var get = _ytOAuth!.Videos.List("snippet"); get.Id = row.VideoId;
                var res = await get.ExecuteAsync();
                if (res.Items.Count == 0) { Log.Error($"Video not found {row.VideoId}."); fail++; }
                else
                {
                    var body = new Video
                    {
                        Id = row.VideoId,
                        Snippet = new VideoSnippet { Title = row.Corrected, Description = row.Description, CategoryId = res.Items[0].Snippet.CategoryId }
                    };
                    await _ytOAuth.Videos.Update(body, "snippet").ExecuteAsync();
                    ok++; Log.Info($"Updated {row.VideoId}: '{row.Corrected}'");
                }
            }
            catch (Exception e) { Log.Error($"Error updating {row.VideoId}: {e.Message}"); fail++; }
            _progress.Value = (i + 1) * 100.0 / todo.Count;
            Status($"Updating... {(int)_progress.Value}%");
        }
        _progress.IsVisible = false;
        var msg = $"Update complete. OK: {ok}, Failed: {fail}.";
        Log.Info(msg); Status(msg);
        await Message("Update Complete", fail > 0 ? msg + "\n\nSee the Log tab for the failed videos." : msg);
        await LoadRenameVideos();
    }

    // ================= tab 3: excel =================
    async Task LoadExcelPreview()
    {
        if (_plExcel.SelectedItem is not PlaylistItemView pl) return;
        _excelStatus.Text = $"Status: Fetching videos for {pl.Id}..."; _excelStatus.ClearValue(TextBlock.ForegroundProperty);
        _excelGrid.ItemsSource = null; _excelRows = null;
        _loadExcelBtn.IsEnabled = _genExcelBtn.IsEnabled = false;
        try
        {
            var videos = await FetchVideos(pl.Id);
            var title = _playlists[pl.Id].Title;
            Log.Info($"Processing {videos.Count} videos for Excel (Playlist Title: {title}).");
            _excelRows = Core.BuildExcelRows(title, videos);
            _excelGrid.ItemsSource = _excelRows;
            _genExcelBtn.IsEnabled = _excelRows.Count > 0;
            var warn = LanguageMismatch();
            _excelStatus.Text = $"Status: {_excelRows.Count} videos processed. Ready to generate Excel." + (warn == null ? "" : "\n" + warn);
            if (warn != null) { _excelStatus.Foreground = WarnBrush; Log.Warn(warn); }
        }
        catch (Exception e) { Log.Error($"Error fetching videos: {e}"); _excelStatus.Text = $"Status: Failed - {e.Message}"; }
        finally { _loadExcelBtn.IsEnabled = true; }
    }

    string? LanguageMismatch()
    {
        var got = _excelRows?.FirstOrDefault()?.LanguageCode;
        var want = Core.LanguageCode(_settings.Language);
        return got == null || got == want ? null
            : $"Warning: this playlist's language code is '{got}', but the selected channel language is {_settings.Language} ({want}).";
    }

    async Task GenerateExcel()
    {
        if (_excelRows == null || _authFolder == null || _plExcel.SelectedItem is not PlaylistItemView pl)
        {
            await Message("Error", "No data loaded or auth folder not set."); return;
        }
        var warn = LanguageMismatch();
        if (warn != null && !await Message("Language mismatch", warn + "\n\nGenerate the Excel file anyway?", yesNo: true)) return;

        var p = _playlists[pl.Id];
        var desc = string.IsNullOrEmpty(p.Description) ? "No_Description" : p.Description;
        var fileName = $"{Core.SanitizeFilename(desc)}_{Core.CourseCode(p.Title)}.xlsx";
        var path = Path.Combine(_authFolder, "Excel sheet", fileName);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            Log.Info($"Generating Excel file automatically: {path}");
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                for (int c = 0; c < Core.ExcelHeaders.Length; c++)
                {
                    var cell = ws.Cell(1, c + 1);
                    cell.Value = Core.ExcelHeaders[c];
                    cell.Style.Font.Bold = true;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                for (int r = 0; r < _excelRows.Count; r++)
                {
                    var cells = _excelRows[r].Cells;
                    for (int c = 0; c < cells.Length; c++) ws.Cell(r + 2, c + 1).Value = cells[c];
                }
                wb.SaveAs(path);
            }
            Log.Info("Excel file generation successful.");
            Status($"Excel saved: {fileName}");
            await Message("Success", $"Excel file saved successfully to:\n{path}");
        }
        catch (Exception e)
        {
            Log.Error($"Excel save failed: {e}");
            await Message("Error", $"Failed to save Excel file: {e.Message}\n\nSee the Log tab for details.");
        }
    }
}
