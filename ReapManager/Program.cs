using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml.Styling;
using Avalonia.Styling;
using Avalonia.Themes.Fluent;
using Avalonia.Threading;

namespace ReapManager;

class App : Application
{
    public static ThemeVariant ThemeFromName(string name) =>
        name == "Light" ? ThemeVariant.Light : name == "Dark" ? ThemeVariant.Dark : ThemeVariant.Default;

    public override void Initialize()
    {
        Styles.Add(new FluentTheme());
        Styles.Add(new StyleInclude(new Uri("avares://REAP_YouTube_Manager"))
        {
            Source = new Uri("avares://Avalonia.Controls.DataGrid/Themes/Fluent.xaml")
        });
        RequestedThemeVariant = ThemeFromName(Settings.Current.Theme);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        // Log UI-thread exceptions instead of crashing the app.
        Dispatcher.UIThread.UnhandledException += (_, e) => { Log.Error($"Unhandled UI error: {e.Exception}"); e.Handled = true; };
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            desktop.MainWindow = new MainWindow();
        base.OnFrameworkInitializationCompleted();
    }
}

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        if (args.Contains("--selftest")) return Core.SelfTest();

        Log.Init();
        Log.Info($"{AppInfo.Name} v{AppInfo.Version} starting.");
        Log.Info($"OS: {RuntimeInformation.OSDescription} ({RuntimeInformation.OSArchitecture}), .NET {Environment.Version}");
        Log.Info($"Log file: {Log.FilePath}");
        AppDomain.CurrentDomain.UnhandledException += (_, e) => Log.Error($"FATAL unhandled error: {e.ExceptionObject}");
        TaskScheduler.UnobservedTaskException += (_, e) => { Log.Error($"Background task error: {e.Exception}"); e.SetObserved(); };
        Settings.Load();

        try
        {
            return AppBuilder.Configure<App>().UsePlatformDetect().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception e) { Log.Error($"FATAL: {e}"); throw; }
        finally { Log.Info("Application closed."); Log.Close(); }
    }
}
