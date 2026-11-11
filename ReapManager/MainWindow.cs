using System.Text.Json;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Ellipse = Avalonia.Controls.Shapes.Ellipse;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ClosedXML.Excel;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.YouTube.v3;
using Google.Apis.YouTube.v3.Data;

namespace ReapManager;

public record PlaylistInfo(string Id, string Title, string Description, long VideoCount);

public class RenameRow
{
    public string Original { get; init; } = "";
    public string Corrected { get; set; } = "";
    public string Description { get; set; } = "";
    public string VideoId { get; init; } = "";
    public bool Needs { get; init; }
}

static class Log
{
    static StreamWriter? _w;
    public static void Init(string dir)
    {
        Directory.CreateDirectory(dir);
        _w?.Dispose();
        _w = new StreamWriter(Path.Combine(dir, DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log")) { AutoFlush = true };
        Info("Logging initialized.");
    }
    static void Write(string level, string msg)
    {
        var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss,fff} - {level} - {msg}";
        Console.WriteLine(line);
        _w?.WriteLine(line);
    }
    public static void Info(string m) => Write("INFO", m);
    public static void Error(string m) => Write("ERROR", m);
}

public class MainWindow : Window
{
    const string OAuthClientFile = "oauth_client.json", ApiKeyFile = "api.txt", PyTokenFile = "token.json";
    const string StoreTokenFile = "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user";
    static readonly string[] Scopes = [YouTubeService.Scope.YoutubeForceSsl];
    static readonly string[] Languages =
        ["English", "Hindi", "Tamil", "Telugu", "Kannada", "Malayalam", "Bengali", "Gujarati", "Marathi", "Punjabi", "Assamese", "Odia"];

    string? _authFolder, _apiKey;
    CancellationTokenSource? _authCts;
    YouTubeService? _ytOAuth, _ytKey;
    readonly Dictionary<string, PlaylistInfo> _playlists = new();
    List<ExcelRow>? _excelRows;
    List<RenameRow> _renameRows = new();

    // tab 1
    readonly ComboBox _langCombo = new() { ItemsSource = Languages, SelectedIndex = 0, MinWidth = 160 };
    readonly TextBox _folderBox = new() { IsReadOnly = true, HorizontalAlignment = HorizontalAlignment.Stretch };
    readonly Button _browseBtn = new() { Content = "Browse..." };
    readonly Button _tokenBtn = new() { Content = "Authorize with Token", IsEnabled = false };
    readonly Button _browserBtn = new() { Content = "Authorize (Browser)", IsEnabled = false };
    readonly Button _abortBtn = new() { Content = "Abort / Reset", IsEnabled = false };
    readonly TextBlock _authStatus = new() { Text = "Status: Awaiting folder.", FontWeight = FontWeight.Bold };
    // tab 2
    readonly ComboBox _plRename = new() { IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Stretch };
    readonly Button _loadVideosBtn = new() { Content = "Load Videos", IsEnabled = false };
    readonly TextBlock _videoStatus = new() { Text = "Status: Select playlist & load." };
    readonly DataGrid _videoGrid = new() { CanUserResizeColumns = true, GridLinesVisibility = DataGridGridLinesVisibility.All };
    readonly TextBlock _renameCount = new() { Text = "0 videos to rename.", FontWeight = FontWeight.Bold, Foreground = new SolidColorBrush(Color.Parse("#d35400")) };
    readonly Button _applyBtn = new() { Content = "Apply Changes", IsEnabled = false };
    readonly ProgressBar _progress = new() { IsVisible = false, Width = 200, Minimum = 0, Maximum = 100 };
    // tab 3
    readonly ComboBox _plExcel = new() { IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Stretch };
    readonly Button _loadExcelBtn = new() { Content = "Load Video Info", IsEnabled = false };
    readonly TextBlock _excelStatus = new() { Text = "Status: Select playlist & load for preview." };
    readonly DataGrid _excelGrid = new() { IsReadOnly = true, CanUserResizeColumns = true, GridLinesVisibility = DataGridGridLinesVisibility.All };
    readonly Button _genExcelBtn = new() { Content = "Generate Excel Sheet", IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Right };
    // shell
    readonly TabItem _tab2 = new() { Header = "2. Rename & Describe" };
    readonly TabItem _tab3 = new() { Header = "3. Excel Export" };
    readonly Ellipse _led = new() { Width = 14, Height = 14, Stroke = Brushes.DimGray, StrokeThickness = 1 };
    readonly TextBlock _statusText = new() { Text = "Welcome!", VerticalAlignment = VerticalAlignment.Center };

    public MainWindow()
    {
        Width = 900; Height = 700; FontSize = 13;
        UpdateTitle();
        _langCombo.SelectionChanged += (_, _) => UpdateTitle();

        _browseBtn.Click += async (_, _) => await BrowseFolder();
        _tokenBtn.Click += async (_, _) => await StartAuth(browser: false);
        _browserBtn.Click += async (_, _) => await StartAuth(browser: true);
        _abortBtn.Click += (_, _) => { _authCts?.Cancel(); SetAuthState(false, "Reset."); };
        _loadVideosBtn.Click += async (_, _) => await LoadRenameVideos();
        _applyBtn.Click += async (_, _) => await ApplyChanges();
        _loadExcelBtn.Click += async (_, _) => await LoadExcelPreview();
        _genExcelBtn.Click += async (_, _) => await GenerateExcel();

        _videoGrid.Columns.Add(new DataGridTemplateColumn
        {
            Header = "Original", Width = new DataGridLength(1, DataGridLengthUnitType.Star), IsReadOnly = true,
            CellTemplate = new FuncDataTemplate<RenameRow>((r, _) => new TextBlock
            {
                Text = r?.Original, Margin = new Thickness(6, 0), VerticalAlignment = VerticalAlignment.Center,
                FontWeight = r?.Needs == true ? FontWeight.Bold : FontWeight.Normal
            })
        });
        _videoGrid.Columns.Add(new DataGridTextColumn { Header = "Corrected", Binding = new Binding(nameof(RenameRow.Corrected)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        _videoGrid.Columns.Add(new DataGridTextColumn { Header = "Description", Binding = new Binding(nameof(RenameRow.Description)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });

        string[] props = [nameof(ExcelRow.CourseCode), nameof(ExcelRow.ChapterName), nameof(ExcelRow.YouTubeUrl), nameof(ExcelRow.VideoTitle),
                          nameof(ExcelRow.VideoDescription), nameof(ExcelRow.OrderNo), nameof(ExcelRow.LanguageCode)];
        for (int i = 0; i < props.Length; i++)
            _excelGrid.Columns.Add(new DataGridTextColumn { Header = Core.ExcelHeaders[i], Binding = new Binding(props[i]) });

        var tab1 = new StackPanel
        {
            Margin = new Thickness(12), Spacing = 10,
            Children =
            {
                Row(new TextBlock { Text = "Channel Language:", VerticalAlignment = VerticalAlignment.Center }, _langCombo),
                Grid3("Auth Folder:", _folderBox, _browseBtn),
                Row(_tokenBtn, _browserBtn, _abortBtn),
                _authStatus,
                new Border
                {
                    BorderBrush = Brushes.LightGray, BorderThickness = new Thickness(1), Padding = new Thickness(10),
                    Child = new TextBlock
                    {
                        TextWrapping = TextWrapping.Wrap,
                        Text = "Instructions:\n1. Pick the channel language.\n2. Select the auth folder (must contain oauth_client.json and api.txt).\n" +
                               "3. Use 'Token' (saved login) or 'Browser' (first time).\n4. Status / LED will update."
                    }
                }
            }
        };
        _tab2.Content = new DockPanel
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
        _tab3.Content = new DockPanel
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
        _videoStatus.Margin = _excelStatus.Margin = new Thickness(0, 8);
        _genExcelBtn.Margin = new Thickness(0, 8, 0, 0);

        var statusBar = new DockPanel
        {
            Margin = new Thickness(8, 4), LastChildFill = true,
            Children = { Dock(_led, Avalonia.Controls.Dock.Left), _statusText }
        };
        _led.Margin = new Thickness(0, 0, 8, 0);
        Content = new DockPanel
        {
            Children =
            {
                Dock(statusBar, Avalonia.Controls.Dock.Bottom),
                new TabControl { Items = { new TabItem { Header = "1. Authentication", Content = tab1 }, _tab2, _tab3 } }
            }
        };
        SetAuthState(false);
    }

    // ---------- layout helpers ----------
    static T Dock<T>(T c, Dock d) where T : Control { DockPanel.SetDock(c, d); return c; }
    static StackPanel Row(params Control[] cs)
    {
        var p = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
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

    void UpdateTitle() => Title = $"Reap {_langCombo.SelectedItem} YouTube Channel Manager";
    void Status(string s) => _statusText.Text = s;

    async Task<bool> Message(string title, string text, bool yesNo = false)
    {
        bool result = false;
        var dlg = new Window
        {
            Title = title, Width = 460, SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner, CanResize = false
        };
        var ok = new Button { Content = yesNo ? "Yes" : "OK", MinWidth = 70 };
        var no = new Button { Content = "No", MinWidth = 70, IsVisible = yesNo };
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

    // ---------- tab 1: auth ----------
    async Task BrowseFolder()
    {
        var picked = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = "Select Auth Folder" });
        var path = picked.FirstOrDefault()?.TryGetLocalPath();
        if (path == null) return;
        _authFolder = path;
        _folderBox.Text = path;
        Log.Init(Path.Combine(path, "logs"));
        UpdateAuthButtons();
    }

    void UpdateAuthButtons()
    {
        bool oauthOk = false, apiOk = false, tokenExists = false;
        _apiKey = null;
        if (_authFolder != null)
        {
            oauthOk = File.Exists(Path.Combine(_authFolder, OAuthClientFile));
            tokenExists = File.Exists(Path.Combine(_authFolder, PyTokenFile)) || File.Exists(Path.Combine(_authFolder, StoreTokenFile));
            try
            {
                var m = Regex.Match(File.ReadAllText(Path.Combine(_authFolder, ApiKeyFile)), @"""API Key""\s*:\s*""(.*?)""");
                if (m.Success) { _apiKey = m.Groups[1].Value; apiOk = true; }
            }
            catch { /* missing api.txt => apiOk stays false */ }
        }
        bool filesOk = oauthOk && apiOk;
        _browserBtn.IsEnabled = filesOk;
        _tokenBtn.IsEnabled = filesOk && tokenExists;
        _abortBtn.IsEnabled = _authFolder != null;
        _authStatus.Text = _authFolder == null ? "Status: Awaiting folder."
            : !filesOk ? "Status: Error! oauth_client.json/api.txt missing."
            : "Status: Ready. Use Token or Browser.";
    }

    /// Seeds the Google token store from the old Python app's token.json, so existing logins keep working.
    TokenResponse? ImportPythonToken()
    {
        var path = Path.Combine(_authFolder!, PyTokenFile);
        if (!File.Exists(path)) return null;
        using var doc = JsonDocument.Parse(File.ReadAllText(path));
        return doc.RootElement.TryGetProperty("refresh_token", out var rt) ? new TokenResponse { RefreshToken = rt.GetString() } : null;
    }

    async Task StartAuth(bool browser)
    {
        if (_authFolder == null) return;
        _browserBtn.IsEnabled = _tokenBtn.IsEnabled = false; _abortBtn.IsEnabled = true;
        _authStatus.Text = $"Status: Starting '{(browser ? "browser" : "token")}' auth..."; Status("Auth in progress...");
        Log.Info($"Starting auth (mode: {(browser ? "browser" : "token")}).");
        _authCts = new CancellationTokenSource();
        try
        {
            var secrets = GoogleClientSecrets.FromFile(Path.Combine(_authFolder, OAuthClientFile)).Secrets;
            var store = new FileDataStore(_authFolder, fullPath: true);
            UserCredential? cred;
            string msg;
            if (browser)
            {
                cred = await GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, Scopes, "user", _authCts.Token, store);
                msg = "Auth successful.";
            }
            else
            {
                var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer { ClientSecrets = secrets, Scopes = Scopes, DataStore = store });
                var token = await flow.LoadTokenAsync("user", _authCts.Token) ?? ImportPythonToken();
                if (token?.RefreshToken == null) { await AuthFailed("Token invalid/missing."); return; }
                cred = new UserCredential(flow, "user", token);
                if (!await cred.RefreshTokenAsync(_authCts.Token)) { await AuthFailed("Token refresh failed."); return; }
                msg = "Token loaded.";
            }
            var init = new BaseClientService.Initializer { HttpClientInitializer = cred, ApplicationName = "REAP YouTube Manager" };
            _ytOAuth = new YouTubeService(init);
            _ytKey = new YouTubeService(new BaseClientService.Initializer { ApiKey = _apiKey, ApplicationName = "REAP YouTube Manager" });
            _authStatus.Text = $"Status: {msg}";
            SetAuthState(true, msg);
            await LoadPlaylists();
        }
        catch (OperationCanceledException) { SetAuthState(false, "Aborted."); }
        catch (Exception e) { Log.Error($"Auth error: {e}"); await AuthFailed($"Auth failed: {e.Message}"); }
    }

    async Task AuthFailed(string msg)
    {
        _authStatus.Text = $"Status: {msg}";
        SetAuthState(false, msg);
        await Message("Auth Failed", msg);
    }

    void SetAuthState(bool ok, string message = "")
    {
        _led.Fill = new SolidColorBrush(ok ? Color.Parse("#14c814") : Color.Parse("#d01414"));
        _tab2.IsEnabled = _tab3.IsEnabled = ok;
        if (ok)
        {
            _browserBtn.IsEnabled = _tokenBtn.IsEnabled = false; _abortBtn.IsEnabled = true;
            Status($"Authenticated! ({message})"); Log.Info($"Authenticated! ({message})");
            return;
        }
        _ytOAuth = _ytKey = null;
        _plRename.ItemsSource = _plExcel.ItemsSource = null;
        _videoGrid.ItemsSource = _excelGrid.ItemsSource = null;
        Status(message == "" ? "Not Authenticated." : $"Not Authenticated. ({message})");
        UpdateAuthButtons();
    }

    // ---------- playlists / videos ----------
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
                foreach (var i in r.Items)
                    _playlists[i.Id] = new PlaylistInfo(i.Id, i.Snippet.Title, i.Snippet.Description ?? "", i.ContentDetails.ItemCount ?? 0);
                page = r.NextPageToken;
            } while (page != null);

            var items = _playlists.Values.Select(p =>
            {
                var d = p.Description.Length > 30 ? p.Description[..30] + "..." : p.Description;
                return new PlaylistItemView(p.Id, $"{p.Title} - {d} - {p.VideoCount} videos");
            }).ToList();
            _plRename.ItemsSource = items; _plExcel.ItemsSource = items;
            if (items.Count == 0) { await Message("Info", "No playlists found/error."); return; }
            _plRename.SelectedIndex = _plExcel.SelectedIndex = 0;
            _plRename.IsEnabled = _plExcel.IsEnabled = _loadVideosBtn.IsEnabled = _loadExcelBtn.IsEnabled = true;
            Status("Playlists loaded.");
        }
        catch (Exception e) { Log.Error($"Error fetching playlists: {e}"); Status($"Error: {e.Message}"); }
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
            foreach (var i in r.Items)
                videos.Add(new VideoInfo(i.Snippet.ResourceId.VideoId, i.Snippet.Title, i.Snippet.Description ?? ""));
            page = r.NextPageToken;
        } while (page != null);
        return videos;
    }

    // ---------- tab 2: rename ----------
    async Task LoadRenameVideos()
    {
        if (_plRename.SelectedItem is not PlaylistItemView pl) return;
        Status("Loading rename videos..."); _videoGrid.ItemsSource = null;
        _loadVideosBtn.IsEnabled = _applyBtn.IsEnabled = false;
        try
        {
            var videos = await FetchVideos(pl.Id);
            _renameRows = videos.Select(v =>
            {
                var (corrected, desc, needs) = Core.CorrectTitle(v.Title);
                return new RenameRow { Original = v.Title, Corrected = corrected, Description = desc, VideoId = v.VideoId, Needs = needs };
            }).ToList();
            _videoGrid.ItemsSource = _renameRows;
            int count = _renameRows.Count(r => r.Needs);
            _videoStatus.Text = $"Loaded {videos.Count} videos."; _renameCount.Text = $"{count} to rename.";
            _applyBtn.IsEnabled = count > 0;
            Status("Videos fetched for processing.");
        }
        catch (Exception e) { Log.Error($"Error fetching videos: {e}"); _videoStatus.Text = $"Status: No videos / error: {e.Message}"; }
        finally { _loadVideosBtn.IsEnabled = true; }
    }

    async Task ApplyChanges()
    {
        var todo = _renameRows.Where(r => r.Needs).ToList();
        if (todo.Count == 0) { await Message("Info", "No videos marked."); return; }
        if (!await Message("Confirm", $"Update {todo.Count} videos?", yesNo: true)) return;

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
                    ok++; Log.Info($"Updated {row.VideoId}");
                }
            }
            catch (Exception e) { Log.Error($"Error updating {row.VideoId}: {e.Message}"); fail++; }
            _progress.Value = (i + 1) * 100.0 / todo.Count;
            Status($"Updating... {(int)_progress.Value}%");
        }
        _progress.IsVisible = false;
        var msg = $"Update complete. OK: {ok}, Failed: {fail}.";
        Status(msg);
        await Message("Update Complete", msg);
        await LoadRenameVideos();
    }

    // ---------- tab 3: excel ----------
    async Task LoadExcelPreview()
    {
        if (_plExcel.SelectedItem is not PlaylistItemView pl) return;
        _excelStatus.Text = $"Status: Fetching videos for {pl.Id}...";
        _excelGrid.ItemsSource = null; _excelRows = null;
        _loadExcelBtn.IsEnabled = _genExcelBtn.IsEnabled = false;
        try
        {
            var videos = await FetchVideos(pl.Id);
            var title = _playlists[pl.Id].Title;
            Log.Info($"Processing {videos.Count} videos for Excel (Playlist Title: {title}).");
            _excelRows = Core.BuildExcelRows(title, videos);
            _excelGrid.ItemsSource = _excelRows;
            _excelStatus.Text = $"Status: {_excelRows.Count} videos processed. Ready to generate Excel.";
            _genExcelBtn.IsEnabled = _excelRows.Count > 0;
        }
        catch (Exception e) { Log.Error($"Error fetching videos: {e}"); _excelStatus.Text = $"Status: Failed - {e.Message}"; }
        finally { _loadExcelBtn.IsEnabled = true; }
    }

    async Task GenerateExcel()
    {
        if (_excelRows == null || _authFolder == null || _plExcel.SelectedItem is not PlaylistItemView pl)
        {
            await Message("Error", "No data loaded or auth folder not set."); return;
        }
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
            await Message("Success", $"Excel file saved successfully to:\n{path}");
        }
        catch (Exception e)
        {
            Log.Error($"Excel save failed: {e}");
            await Message("Error", $"Failed to save Excel file: {e.Message}");
        }
    }
}
