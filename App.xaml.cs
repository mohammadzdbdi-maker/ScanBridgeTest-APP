using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Drawing;
using System.Media;
using MessageBox = System.Windows.MessageBox;

namespace ScanBridgeTest;

public partial class App : System.Windows.Application
{
    // تنظیم DPI awareness از طریق کد به‌جای app.manifest: مانیفست exe اگه یه ذره غلط باشه
    // کل اجرای برنامه رو با "CreateProcess failed; code 14001 (side-by-side configuration
    // incorrect)" می‌ترکونه (این یه بار واقعاً اتفاق افتاد). این روش با یه فراخوانی API استاندارد
    // ویندوز (از Windows 10 1703 به بعد، که این برنامه با WebView2 حداقلش همینه) هیچ ریسکی
    // برای exe نداره؛ اگه هم روی ویندوز خیلی قدیمی نبود، فقط false برمی‌گردونه و رد می‌شیم.
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

    private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);

    // اگه سیستم به‌اندازه‌ی کافی جدید نبود که از تابع بالا (Windows 10 1703+) پشتیبانی کنه،
    // این دو تا رده‌ی پایین‌تر رو هم امتحان می‌کنیم تا حداقل روی ویندوز 8.1 به بعد
    // (SetProcessDpiAwareness) یا حتی ویستا به بعد (SetProcessDPIAware) یه DPI awareness
    // پایه فعال بشه، نه اینکه کاملاً بی‌خیال بشیم و اپ روی سیستم‌های قدیمی‌تر بیت‌مپ کشیده بشه.
    [DllImport("shcore.dll", SetLastError = true)]
    private static extern int SetProcessDpiAwareness(int value);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetProcessDPIAware();

    private static void ApplyBestAvailableDpiAwareness()
    {
        try { if (SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)) return; } catch { }
        try { if (SetProcessDpiAwareness(2 /* PROCESS_PER_MONITOR_DPI_AWARE */) == 0) return; } catch { }
        try { SetProcessDPIAware(); } catch { }
    }

    private const string AppDisplayName = "Scanbridge";
    private ScanBridgeService? _service;
    private MainWindow? _mainWindow;
    private NotifyIcon? _notifyIcon;
    private ToolStripMenuItem? _openItem;
    private ToolStripMenuItem? _exitItem;
    private int _lastConnectedClients = 0;
    private LocalizationManager _localization = LocalizationManager.Instance;

    // مسیر نصب (AppContext.BaseDirectory) گاهی write-protected است (مثلاً برنامه داخل Program
    // Files و کاربر بدون دسترسی مدیر نصب کرده). قبلاً در این حالت هر سه فایل لاگ
    // (startup-trace.log/startup-error.log) بی‌صدا هیچ‌وقت نوشته نمی‌شدند و کل مکانیزم تشخیص
    // خطا (از جمله گزارش تشخیصی که کاربر برای پشتیبانی می‌فرستد) بدون هیچ رد یا هشداری از کار
    // می‌افتاد (باگ گزارش ممیزی). حالا اگر نوشتن در مسیر نصب شکست بخورد، یک‌بار هم پوشه‌ی
    // AppData کاربر (که تقریباً همیشه قابل‌نوشتن است) امتحان می‌شود.
    private static void AppendAppLogLine(string fileName, string text)
    {
        try
        {
            string path = System.IO.Path.Combine(AppContext.BaseDirectory, fileName);
            System.IO.File.AppendAllText(path, text);
            return;
        }
        catch { }

        try
        {
            string fallbackDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Scanbridge", "logs");
            System.IO.Directory.CreateDirectory(fallbackDir);
            System.IO.File.AppendAllText(System.IO.Path.Combine(fallbackDir, fileName), text);
        }
        catch { }
    }

    private static void LogStartupTrace(string message)
    {
        AppendAppLogLine("startup-trace.log", $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
    }

    private static void LogStartupError(Exception ex, string section)
    {
        AppendAppLogLine("startup-error.log", $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {section}\n{ex}\n\n");
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        // باید همین اول تابع و قبل از ساخته‌شدن هر پنجره‌ای صدا زده بشه تا WPF از همون اول
        // با آگاهی درست از DPI هر مانیتور رندر کنه (وگرنه روی صفحه‌های اسکیل‌شده — مثلاً
        // ۱۳۶۶×۷۶۸ با اسکیل ۱۲۵٪/۱۵۰٪ — ویندوز کل پنجره رو بیت‌مپ می‌کشه و کادرها/دکمه‌ها از
        // لبه‌ی صفحه بیرون می‌زنن).
        ApplyBestAvailableDpiAwareness();

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

        // تا اینجا فقط خطاهای thread رابط کاربری (DispatcherUnhandledException) لاگ می‌شدند. دو
        // دسته‌ی مهم دیگر از خطا اصلاً هیچ‌جا ثبت نمی‌شدند و در نتیجه در گزارش تشخیصی هم دیده
        // نمی‌شدند - یعنی اگر یکی از این‌ها باعث مشکل کاربر می‌شد، فایل گزارش تشخیصی که برای
        // پشتیبانی می‌فرستاد کاملاً خالی از سرنخ بود:
        // ۱) استثناهای thread پس‌زمینه (مثلاً داخل Task.Run بدون await/catch) - این‌ها کل برنامه را
        //    از پا در می‌آورند، بدون هیچ اثری در لاگ.
        // ۲) استثناهای Task هایی که «آتش‌وفراموش» صدا زده شده‌اند (الگوی رایج در این پروژه:
        //    _ = SomeAsyncMethodAsync(); بدون await) - در دات‌نت این‌ها به‌طور پیش‌فرض برنامه را
        //    نمی‌ترکانند، اما بی‌صدا هم قورت داده می‌شدند و هیچ اثری باقی نمی‌ماند.
        AppDomain.CurrentDomain.UnhandledException += (s, args) =>
        {
            if (args.ExceptionObject is Exception bgEx)
                LogStartupError(bgEx, "AppDomain.UnhandledException" + (args.IsTerminating ? " (Terminating)" : ""));
        };

        System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (s, args) =>
        {
            LogStartupError(args.Exception, "TaskScheduler.UnobservedTaskException");
            args.SetObserved();
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

            // ShutdownMode بالای این متد روی OnExplicitShutdown تنظیم شده (چون برنامه معمولاً
            // فقط با آیکون تری، بدون هیچ پنجره‌ی قابل‌مشاهده، اجرا می‌شود). یعنی اگر همین‌جا
            // ساخت سرویس/پنجره‌ی اصلی/آیکون تری شکست بخورد، WPF خودش تصمیم به بستن برنامه
            // نمی‌گیرد - فرآیند بدون هیچ پنجره یا آیکون تری، کاملاً نامرئی، تا ابد زنده می‌ماند
            // (فقط از Task Manager قابل بستن است؛ باگ ۱۸ گزارش ممیزی). چون تا اینجا رسیدن یعنی
            // راه‌اندازی قطعاً شکست خورده، باید صریحاً بسته شود.
            Shutdown(1);
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
            // قبلاً اینجا WindowState.Normal بود که همیشه WindowState="Maximized" تعریف‌شده در
            // MainWindow.xaml را بی‌اثر می‌کرد - یعنی برنامه هیچ‌وقت واقعاً تمام‌صفحه باز نمی‌شد
            // (نه بار اول، نه دفعات بعد)، حتی وقتی از تری هم دوباره باز می‌شد. حالا با Maximized
            // هماهنگ با تنظیم پیش‌فرض XAML، برنامه همیشه تمام‌صفحه باز می‌شود.
            _mainWindow.WindowState = WindowState.Maximized;
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
