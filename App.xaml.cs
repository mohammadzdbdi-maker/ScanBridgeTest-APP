using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.Media;
using MessageBox = System.Windows.MessageBox;

namespace ScanBridgeTest;

public partial class App : System.Windows.Application
{
    private const string AppDisplayName = "Scanbridge";
    private ScanBridgeService? _service;
    private MainWindow? _mainWindow;
    private NotifyIcon? _notifyIcon;
    private ToolStripMenuItem? _openItem;
    private ToolStripMenuItem? _exitItem;
    private int _lastConnectedClients = 0;
    private LocalizationManager _localization = LocalizationManager.Instance;

    private static void LogStartupTrace(string message)
    {
        try
        {
            string path = System.IO.Path.Combine(AppContext.BaseDirectory, "startup-trace.log");
            System.IO.File.AppendAllText(path, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
        }
        catch { }
    }

    private static void LogStartupError(Exception ex, string section)
    {
        try
        {
            string path = System.IO.Path.Combine(AppContext.BaseDirectory, "startup-error.log");
            string text = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {section}\n{ex}\n\n";
            System.IO.File.AppendAllText(path, text);
        }
        catch { }
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        LogStartupTrace("OnStartup entered. Args: " + string.Join(" ", e.Args));
        base.OnStartup(e);
        ShutdownMode = ShutdownMode.OnExplicitShutdown;
        LogStartupTrace("base.OnStartup completed.");

        this.DispatcherUnhandledException += (s, args) =>
        {
            var lang = _localization.CurrentLanguage;
            string title = lang switch
            {
                AppLanguage.English => "Runtime Error",
                _ => "خطا در اجرا"
            };

            string detailsLabel = lang switch
            {
                AppLanguage.English => "Details",
                _ => "جزئیات"
            };

            string message = lang switch
            {
                AppLanguage.English => $"Application error:\n{args.Exception.Message}\n\n{detailsLabel}:\n{args.Exception.InnerException?.Message}",
                _ => $"خطای برنامه:\n{args.Exception.Message}\n\n{detailsLabel}:\n{args.Exception.InnerException?.Message}"
            };

            LogStartupError(args.Exception, "DispatcherUnhandledException");
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };

        try
        {
            LogStartupTrace("Creating ScanBridgeService...");
            _service = new ScanBridgeService();
            LogStartupTrace("Starting ScanBridgeService...");
            _service.Start();
            LogStartupTrace("ScanBridgeService started.");

            LogStartupTrace("Creating MainWindow...");
            _mainWindow = new MainWindow(_service);
            LogStartupTrace("MainWindow created.");

            SetupTrayIcon();
            LogStartupTrace("Tray icon created.");

            // گوش دادن به تغییر زبان برنامه
            _localization.LanguageChanged += (_, _) =>
            {
                Dispatcher.BeginInvoke(UpdateTrayLanguage);
            };

            bool silentStartup = e.Args.Any(a => a.Equals("--startup", StringComparison.OrdinalIgnoreCase));
            LogStartupTrace("silentStartup=" + silentStartup);
            if (!silentStartup)
            {
                ForceShowMainWindow();
                Dispatcher.BeginInvoke(new Action(ForceShowMainWindow), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            }

            _service.ConnectionStatusChanged += (_, args) =>
            {
                _lastConnectedClients = args.ConnectedClients;
                UpdateTrayTooltip();
            };
        }
        catch (Exception ex)
        {
            var lang = _localization.CurrentLanguage;
            string title = lang switch
            {
                AppLanguage.English => "Startup Error",
                _ => "خطای شروع"
            };

            string message = lang switch
            {
                AppLanguage.English => $"Error while starting the application:\n{ex.Message}",
                _ => $"خطا هنگام راه اندازی برنامه:\n{ex.Message}"
            };

            LogStartupError(ex, "OnStartup");
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ForceShowMainWindow()
    {
        if (_mainWindow is null)
            return;

        try
        {
            LogStartupTrace("ForceShowMainWindow called.");
            _mainWindow.ShowInTaskbar = true;
            _mainWindow.WindowState = WindowState.Normal;
            _mainWindow.Show();
            _mainWindow.Activate();
            _mainWindow.Topmost = true;
            _mainWindow.Topmost = false;
            _mainWindow.Focus();
            LogStartupTrace("MainWindow show/activate completed. IsVisible=" + _mainWindow.IsVisible);
        }
        catch (Exception ex)
        {
            LogStartupError(ex, "ForceShowMainWindow");
        }
    }

    private Icon LoadTrayIcon()
    {
        try
        {
            string iconPath = System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "app-icon.ico");
            if (File.Exists(iconPath))
                return new Icon(iconPath);
        }
        catch { }

        try
        {
            var resourceInfo = GetResourceStream(new Uri("pack://application:,,,/Assets/app-icon.ico", UriKind.Absolute));
            if (resourceInfo?.Stream != null)
                return new Icon(resourceInfo.Stream);
        }
        catch { }

        try
        {
            string? exePath = Environment.ProcessPath;
            if (!string.IsNullOrWhiteSpace(exePath) && File.Exists(exePath))
            {
                Icon? extracted = Icon.ExtractAssociatedIcon(exePath);
                if (extracted != null)
                    return extracted;
            }
        }
        catch { }

        return SystemIcons.Application;
    }

    private void SetupTrayIcon()
    {
        Icon appIcon = LoadTrayIcon();

        _notifyIcon = new NotifyIcon
        {
            Icon = appIcon,
            Visible = true
        };

        var contextMenu = new ContextMenuStrip();

        _openItem = new ToolStripMenuItem();
        _openItem.Click += (_, _) => RestoreWindow();

        _exitItem = new ToolStripMenuItem();
        _exitItem.Click += (_, _) => ExitApplication();

        contextMenu.Items.Add(_openItem);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(_exitItem);
        _notifyIcon.ContextMenuStrip = contextMenu;

        _notifyIcon.DoubleClick += (_, _) => RestoreWindow();

        // اعمال متن‌های اولیه بر اساس زبان جاری
        UpdateTrayLanguage();
    }

    private void UpdateTrayLanguage()
    {
        if (_notifyIcon is null) return;

        var lang = _localization.CurrentLanguage;

        // به‌روزرسانی متن دکمه باز کردن
        if (_openItem != null)
        {
            _openItem.Text = lang switch
            {
                AppLanguage.English => "Open",
                _ => "ورود"
            };
        }

        // به‌روزرسانی متن دکمه خروج
        if (_exitItem != null)
        {
            _exitItem.Text = lang switch
            {
                AppLanguage.English => "Exit",
                _ => "خروج"
            };
        }

        // به‌روزرسانی تول‌تیپ روی آیکون
        UpdateTrayTooltip();
    }

    private void UpdateTrayTooltip()
    {
        if (_notifyIcon is null) return;

        var lang = _localization.CurrentLanguage;

        string tooltip = _lastConnectedClients switch
        {
            0 => lang switch
            {
                AppLanguage.English => $"{AppDisplayName} — Waiting for connection",
                _ => $"{AppDisplayName} — منتظر اتصال"
            },
            1 => lang switch
            {
                AppLanguage.English => $"{AppDisplayName} — 1 device connected",
                _ => $"{AppDisplayName} — ۱ دستگاه متصل"
            },
            _ => lang switch
            {
                AppLanguage.English => $"{AppDisplayName} — {_lastConnectedClients} devices connected",
                _ => $"{AppDisplayName} — {_lastConnectedClients} دستگاه متصل"
            }
        };

        _notifyIcon.Text = tooltip.Length > 63 ? tooltip.Substring(0, 63) : tooltip;
    }

    private void RestoreWindow()
    {
        ForceShowMainWindow();
    }

    private void ExitApplication()
    {
        _notifyIcon?.Dispose();
        _service?.Dispose();
        System.Windows.Application.Current.Shutdown();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _notifyIcon?.Dispose();
        _service?.Dispose();
        base.OnExit(e);
    }

    public void ShowUnexpectedDisconnectAlert()
    {
        if (_notifyIcon is null) return;
        SystemSounds.Hand.Play();

        var lang = _localization.CurrentLanguage;

        _notifyIcon.BalloonTipTitle = lang switch
        {
            AppLanguage.English => "Disconnected",
            _ => "قطع اتصال"
        };

        _notifyIcon.BalloonTipText = lang switch
        {
            AppLanguage.English => "Connection with phone lost",
            _ => "اتصال با گوشی قطع شد"
        };

        _notifyIcon.ShowBalloonTip(3000);
    }
}
