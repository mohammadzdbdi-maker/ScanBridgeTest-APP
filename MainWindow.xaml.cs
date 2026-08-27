using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using ClosedXML.Excel;

namespace ScanBridgeTest;

public partial class MainWindow : Window
{
    private enum ExportTarget
    {
        Excel,
        Pdf
    }

    private enum FormulaRegistrationMode
    {
        Unknown,
        NoPrescription,
        PrescriptionBased
    }

    private const string StartupRegistryPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupRegistryValueName = "ScanBridge";
    private readonly ScanBridgeService _service;
    private LocalizationManager _localization = LocalizationManager.Instance;
    private TtTeckSettings _ttTeckSettings = new();
    private ScanbridgeLicense _activeLicense = ScanbridgeLicense.Missing();
    private string _activeLicenseCode = string.Empty;
    private const string LegacyLicenseCodePrefix = "SCB1";
    private const string ServerLicenseCodePrefix = "SCB2";
    private const string LicenseApiBaseUrl = "https://license.scanbridge.ir/api";
    private const string LicenseHmacSecret = "Scanbridge-Local-License-v1--replace-with-server-rsa-before-public-release";
    private const string LicenseServerPublicKeyPem = @"-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAvynv+ffiDEixbyWorvwx
Z+r1CWdbr0dr8nzIIZwian+L9Z/HHl1gwzMA7/fuPMnNB9qYxmTa8uvwMHLjEaRe
5CxEGMeDk6pEW+4yKT3uDiro5XQhOM8x40pvifpErLpkrcHfVgSSZliqiUvSSugd
RrfEatgisWHZoBVanmCP5v1dbhHtTtmyZqxgREQBrM9ITaRBVkY7pcLzEBoLS+iZ
fLlHgRPTb7P8JVLHkbfw5xNKE5w/9z5gn0Cd12Gb0OFDnM8wsoeIGl2hBTO6sq2E
LH5zxyzvoNQOunr1BN7MamctPMenbZS8Wil3m6I9XulFoX6Xu8KNMEe7QMaCIDJJ
nQIDAQAB
-----END PUBLIC KEY-----";
    private readonly HttpClient _licenseHttpClient = new();
    private readonly System.Windows.Threading.DispatcherTimer _licenseHeartbeatTimer = new() { Interval = TimeSpan.FromHours(6) };
    private DateTime _lastLicenseOnlineValidationUtc = DateTime.MinValue;

    // یادآور بله: یک ربات مشترک برای کل برنامه (نه هر داروخانه یک ربات جدا).
    // TODO: توکن و یوزرنیم واقعی ربات بله را اینجا جایگزین کنید (یوزرنیم بدون @).
    private const string SharedBaleBotToken = "1891930342:sQTq-cG_mwKSesa7QBaUKkqmXqqEOyuKckM";
    private const string SharedBaleBotUsername = "ReminderPharmacybot";
    private bool _isInitializingUi = true;
    private bool _isExportingOnlyTtTeck = false;
    private bool _isTtTeckHistoryFilterActive = false;
    private bool _isFormulaHistoryFilterActive = false;
    private bool _isTtacPanelFormulaOnly = false;
    private string _ttacPanelSearchText = string.Empty;
    private string _receiveStatusSearchText = string.Empty;
    private bool _autoOpenInfantFormulaRegistration = true;
    private readonly HashSet<string> _autoOpenedFormulaRegistrationKeys = new();
    private readonly HashSet<string> _specialPrescriptionGenericCodes = new(StringComparer.OrdinalIgnoreCase);
    private bool _specialPrescriptionGenericCodesLoaded = false;
    private readonly HashSet<string> _noPrescriptionFormulaProductIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _prescriptionFormulaProductIds = new(StringComparer.OrdinalIgnoreCase);
    // شناسه فرآورده -> کد گروه عکس. دو دیکشنری جدا برای نسخه‌دار/بدون‌نسخه چون کد گروه فقط
    // داخل همان فایل یکتاست (مثلاً کد ۱۱ در Rx یک محصول است و در No-Rx محصول کاملاً دیگری).
    private readonly Dictionary<string, string> _prescriptionFormulaPhotoCodes = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _noPrescriptionFormulaPhotoCodes = new(StringComparer.OrdinalIgnoreCase);
    private bool _formulaProductIdsLoaded = false;
    private bool _isDateRangeInitializing = false;
    private bool _isAutoAdvancingTtacField = false;
    private bool _focusTtacMobileAfterStyledMessageClose = false;
    private bool _switchTtacToPrescriptionAfterStyledMessageClose = false;
    private DateTime? _historyFilterFrom = null;
    private DateTime? _historyFilterTo = null;
    private readonly PersianCalendar _persianCalendar = new();
    private ExportTarget _pendingExportTarget = ExportTarget.Excel;
    private string _productDetailsCurrentBarcode = string.Empty;
    private string _historyLoadedPharmacyKey = "default";
    // قبلاً Dictionary معمولی بود - وقتی دو گوشی تقریباً هم‌زمان اسکن می‌کردند، استعلام تی‌تک هر
    // کدام (بعد از await روی HTTP) روی یک ترد pool مجزا ادامه پیدا می‌کرد و همزمان روی همین
    // دیکشنری می‌نوشت؛ Dictionary معمولی برای نوشتن هم‌زمان از چند ترد thread-safe نیست (حتی با
    // کلیدهای متفاوت) و می‌تواند خراب شود یا exception بدهد (باگ ۱۲ گزارش ممیزی).
    // ConcurrentDictionary خواندن/نوشتن هم‌زمان را بدون قفل دستی ایمن می‌کند.
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, DrugInfo> _ttTeckDetailsByBarcode = new();
    // فقط برای سریال‌سازی نوشتن فایل کش تی‌تک (SaveTtTeckDetailsCache) - خودِ دیکشنری با
    // ConcurrentDictionary ایمن شده، ولی File.WriteAllText اگر از دو ترد هم‌زمان صدا زده شود
    // می‌تواند فایل را خراب کند یا با خطای اشغال‌بودن فایل شکست بخورد.
    private readonly object _ttTeckDetailsCacheFileLock = new();
    private readonly Dictionary<string, string> _deviceAliases = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<ConnectedDeviceInfo> _lastConnectedDevices = new();
    private bool _usbInternetTipShown;
    private Services.AdbUsbBridge? _adbBridge;
    private string _editingDeviceOriginalName = string.Empty;
    private TtTeckHistoryRow? _pendingRetryTtTeckRow;
    private TtTeckHistoryRow? _pendingRegistrationTtTeckRow;
    private string _lastTtTeckWebViewUrl = "https://newstatisticsreports.ttac.ir/pharmacyDashboard";
    private Microsoft.Web.WebView2.Core.CoreWebView2Environment? _ttTeckWebViewEnvironment;
    private bool _ttacBrowserSlowWarningShown = false;
    private DateTime _ttacBrowserOpenedAtUtc = DateTime.MinValue;
    private Microsoft.Web.WebView2.Wpf.WebView2? _ttTeckWebView;
    private Microsoft.Web.WebView2.Wpf.WebView2 TtTeckWebView
    {
        get
        {
            if (_ttTeckWebView == null)
            {
                _ttTeckWebView = new Microsoft.Web.WebView2.Wpf.WebView2();
                TtTeckWebViewHost.Children.Clear();
                TtTeckWebViewHost.Children.Add(_ttTeckWebView);
            }

            return _ttTeckWebView;
        }
    }
    private long? _ttacCurrentPrescriptionId;
    private string _ttacCurrentCaptchaId = string.Empty;
    private string _ttacCurrentNationalId = string.Empty;
    private string _ttacCurrentBirthDate = string.Empty;
    private string _ttacCurrentPatientFullName = string.Empty;
    private bool _ttacCurrentIsElectronic;
    private string? _ttacAccessTokenOverride;
    private string _ttacPharmacyDisplayName = string.Empty;
    private DateTime _ttacAccessTokenExpiresAtUtc = DateTime.MinValue;
    // برای تشخیص لحظه‌ی دقیق «تمام شدن» توکن تی‌تک (نه فقط وضعیت لحظه‌ای) - هر ثانیه با تیک تایمر
    // شمارش معکوس مقایسه می‌شود تا فقط دقیقاً همان لحظه‌ی گذار از معتبر به نامعتبر یک بار اطلاع‌رسانی شود.
    private bool _ttacTokenWasValidLastCheck;
    private readonly System.Windows.Threading.DispatcherTimer _ttacTokenCountdownTimer = new() { Interval = TimeSpan.FromSeconds(1) };
    private readonly HttpClient _ttacHttpClient = new();
    private readonly List<TtacRegistrationHistoryEntry> _ttacRegistrationHistory = new();
    private string _ttacRegistrationHistoryLoadedPharmacyKey = "default";
    private Func<Task>? _pendingTtacRetryAction;
    private string? _pendingTtacRetryLabel;
    // کلیدهای ثبت‌قلم‌هایی که یک‌بار درخواستشان به تی‌تک ارسال شده - برای جلوگیری از ارسال
    // خودکار و بی‌صدای دوباره‌ی همان قلم بعد از قطع/وصل نشست (مشکل ۳ گزارش باگ).
    private readonly HashSet<string> _ttacSubmittedItemKeys = new(StringComparer.OrdinalIgnoreCase);
    private bool _isTtacTokenMonitorRunning;
    // وقتی کاربر روی دکمه‌ی «ورود به داروخانه ...» در صفحه‌ی اصلی کلیک می‌کند، این فیلد
    // نام‌کاربری همان داروخانه را نگه می‌دارد تا مرورگر داخلی دقیقاً همان حساب را (نه آخرین
    // حساب استفاده‌شده) خودکار پر کند. بعد از پر کردن صفحه‌ی ورود یا موفق شدن ورود پاک می‌شود.
    private string? _pendingTtacAutofillUsername;
    // وقتی ورود با یک داروخانه‌ی خاص در مرورگر داخلی شکست می‌خورد، این فیلد نام‌کاربری همان
    // داروخانه را نگه می‌دارد تا دکمه‌ی «تلاش مجدد» دقیقاً با همان حساب دوباره تلاش کند.
    private string? _ttacRetryUsername;
    // وقتی ورود از پنجره‌ی انتخاب داروخانه شروع شده باشد true است؛ بعد از موفق شدن ورود، بنر
    // سبز «ورود موفق شد» داخل همان پنجره نشان داده می‌شود.
    private bool _ttacQuickLoginInProgress;
    // جلوی اعمال دوباره‌ی توکن/بستن پنجره وقتی هم اسکریپت، هم ناوبری و هم مانیتور توکن را می‌بینند.
    private bool _ttacLoginSuccessHandled;
    // وقتی کاربر داروخانه‌ی دیگری را می‌زند، توکن قبلی را تا دیدن صفحه‌ی ورود idp نادیده می‌گیریم.
    private bool _ttacWaitingForFreshLogin;
    private bool _ttacSawIdpLoginPage;
    private readonly List<string> _queuedReceiveStatusBarcodes = new();
    private readonly HashSet<string> _receiveStatusKnownBarcodes = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _cargoDeliveryKnownBarcodes = new(StringComparer.OrdinalIgnoreCase);
    private string _cargoDeliveryLoadedPharmacyKey = string.Empty;
    private bool _isCargoDeliveryBulkRunning = false;
    // --- ویژگی هشدار تاریخ انقضای نزدیک (کالاهای تحویل بار) ---
    private static readonly HttpClient _baleHttpClient = new();
    private ExpiryAlertSettings _expiryAlertSettings = new();
    // آیا کاربر همین الان یک فیلتر بازه‌ی دستی روی لیست «تاریخ نزدیک» اعمال کرده؟ این کاملاً جدا
    // از پین‌شدن اقلام است - اقلامی که به آستانه‌ی تنظیمات رسیده‌اند همیشه (چه این فیلتر فعال باشد
    // چه نباشد) بالای لیست پین می‌شوند؛ این فیلتر فقط اقلام «اضافه‌ی» دیگر را کنترل می‌کند.
    private bool _expiryFilterActive;
    private readonly System.Windows.Threading.DispatcherTimer _expiryAlertCheckTimer = new() { Interval = TimeSpan.FromHours(6) };

    // بررسی بروزرسانی برنامه: یک فایل JSON ساده روی سایت ({"version","message","url"}) - همان
    // الگویی که اپ اندروید هم برای «پیام‌ها»/بروزرسانی استفاده می‌کند (UPDATE_CHECK_URL در
    // MainActivity.kt). هیچ سرور اختصاصی لازم نیست؛ بعد از هر انتشار نسخه‌ی جدید کافی است همین
    // یک فایل روی سایت بازنویسی شود. نگاه کنید به CheckForAppUpdateAsync.
    private readonly HttpClient _updateCheckHttpClient = new();
    private readonly System.Windows.Threading.DispatcherTimer _appUpdateCheckTimer = new() { Interval = TimeSpan.FromHours(24) };
    private const string AppUpdateCheckUrl = "https://scanbridge.ir/app/update-desktop.json";
    private bool _openReceiveStatusAfterStyledMessageClose = false;
    private string _receiveStatusBarcodeAfterStyledMessageClose = string.Empty;
    private string _receiveStatusLoadedPharmacyKey = string.Empty;
    private TtacRepeatFormulaContext? _lastFormulaRepeatContext;
    private bool _isMonthlyArchivePromptOpen = false;
    private bool _ttac5173FlowAfterStyledMessageClose = false;
    private string _ttac5173BarcodeAfterStyledMessageClose = string.Empty;
    private TtacRepeatFormulaContext? _ttac5173ReturnContext;
    private bool _returnToRegistrationAfterStyledMessageClose = false;
    private string _returnRegistrationBarcodeAfterStyledMessageClose = string.Empty;
    private TtacRepeatFormulaContext? _returnRegistrationContextAfterStyledMessageClose;
    private bool _styledMessageLogoutAction = false;
    private string _styledMessageLinkUrl = string.Empty;

    private sealed class TtacSessionExpiredException : Exception
    {
        public TtacSessionExpiredException(string message) : base(message) { }
    }

    private sealed class TtacRepeatFormulaContext
    {
        public string Amount { get; set; } = "1";
        public string NationalId { get; set; } = string.Empty;
        public string BirthDay { get; set; } = string.Empty;
        public string BirthMonth { get; set; } = string.Empty;
        public string BirthYear { get; set; } = string.Empty;
        public string Mobile { get; set; } = string.Empty;
        public string MedicalCouncil { get; set; } = string.Empty;
        public bool IsElectronic { get; set; }
    }

    public ObservableCollection<ScanRecord> HistoryItems { get; } = new();
    public ObservableCollection<HistoryDisplayRow> HistoryViewItems { get; } = new();
    public ObservableCollection<TtTeckHistoryRow> TtTeckHistoryItems { get; } = new();
    public ObservableCollection<ProductDetailField> ProductDetailsExtraFields { get; } = new();
    public ObservableCollection<ProductDetailField> OperationDetailsLeftFields { get; } = new();
    public ObservableCollection<ProductDetailField> OperationDetailsRightFields { get; } = new();
    public ObservableCollection<TtacRegistrationLogRow> TtacRegistrationLogItems { get; } = new();
    public ObservableCollection<ReceiveStatusRow> ReceiveStatusItems { get; } = new();
    public ObservableCollection<CargoDeliveryRow> CargoDeliveryItems { get; } = new();
    public ObservableCollection<ExpiryWatchDisplayRow> ExpiryWatchDisplayItems { get; } = new();
    public ObservableCollection<DeviceRowDisplayViewModel> DeviceRows { get; } = new();
    public ObservableCollection<AppMessage> Messages { get; } = new();

    public string MessageLinkButtonText => _localization.GetString("OpenLink2");

    public string DeleteButtonText => _localization.GetString("Delete");

    public string RetryLookupButtonText => _localization.GetString("RetryTtTeckConfirm");

    public string RegisterInTtTeckButtonText => _localization.GetString("Register");

    public string CopyButtonText => GetHistoryCopyButtonText();

    public string ReceiveStatusButtonText => _localization.GetString("ReceiveStatus");

    public MainWindow(ScanBridgeService service)
    {
        TraceMainWindowStartup("MainWindow constructor entered");
        InitializeComponent();
        _isInitializingUi = false;
        TraceMainWindowStartup("InitializeComponent completed");
        DataContext = this;
        TraceMainWindowStartup("DataContext set");
        _service = service;
        _ttacHttpClient.Timeout = TimeSpan.FromSeconds(35);
        _licenseHttpClient.Timeout = TimeSpan.FromSeconds(20);
        _licenseHeartbeatTimer.Tick += async (_, _) => await ValidateActiveLicenseOnlineBestEffortAsync();
        EnsureBundledDataFilesAvailable();
        TraceMainWindowStartup("Service assigned");

        TraceMainWindowStartup("LoadTtTeckSettings start");
        LoadTtTeckSettings();
        TraceMainWindowStartup("LoadTtTeckSettings done");

        TraceMainWindowStartup("LoadDeviceAliases start");
        LoadDeviceAliases();
        TraceMainWindowStartup("LoadDeviceAliases done");

        TraceMainWindowStartup("LoadTtTeckDetailsCache start");
        LoadTtTeckDetailsCache();
        TraceMainWindowStartup("LoadTtTeckDetailsCache done");

        TraceMainWindowStartup("LoadTtacRegistrationHistory start");
        LoadTtacRegistrationHistory();
        TraceMainWindowStartup("LoadTtacRegistrationHistory done");

        TraceMainWindowStartup("InitializeDateRangeFilterControls start");
        InitializeDateRangeFilterControls();
        TraceMainWindowStartup("InitializeDateRangeFilterControls done");

        TraceMainWindowStartup("InitializeTtTeckRegistrationControls start");
        InitializeTtTeckRegistrationControls();
        TraceMainWindowStartup("InitializeTtTeckRegistrationControls done");

        TraceMainWindowStartup("InitializeExpiryFilterControls start");
        InitializeExpiryFilterControls();
        TraceMainWindowStartup("InitializeExpiryFilterControls done");

        TraceMainWindowStartup("InitializeHighUsageBarcodeFeature start");
        InitializeHighUsageBarcodeFeature();
        TraceMainWindowStartup("InitializeHighUsageBarcodeFeature done");

        TraceMainWindowStartup("UpdateLanguageUI start");
        UpdateLanguageUI();
        TraceMainWindowStartup("UpdateLanguageUI done");
        _ttacTokenCountdownTimer.Tick += (_, _) => TtacTokenCountdownTick();
        _ttacTokenCountdownTimer.Start();
        
        _localization.LanguageChanged += (_, _) => 
        {
            Dispatcher.BeginInvoke(() => UpdateLanguageUI());
        };
        TraceMainWindowStartup("LanguageChanged handler attached");

        TraceMainWindowStartup("UpdateSystemInfo start");
        UpdateSystemInfo();
        TraceMainWindowStartup("UpdateSystemInfo done");

        DevicesList.ItemsSource = DeviceRows;
        MessagesList.ItemsSource = Messages;
        HistoryListBox.ItemsSource = HistoryViewItems;
        TtTeckHistoryListBox.ItemsSource = TtTeckHistoryItems;
        TtacPanelListBox.ItemsSource = TtTeckHistoryItems;
        ProductDetailsExtraFieldsList.ItemsSource = ProductDetailsExtraFields;
        OperationDetailsLeftList.ItemsSource = OperationDetailsLeftFields;
        OperationDetailsRightList.ItemsSource = OperationDetailsRightFields;
        TtTeckRegistrationLogsList.ItemsSource = TtacRegistrationLogItems;
        ReceiveStatusListBox.ItemsSource = ReceiveStatusItems;
        System.Windows.Data.CollectionViewSource.GetDefaultView(ReceiveStatusItems).Filter = item => item is ReceiveStatusRow row && DoesReceiveStatusRowMatchFilter(row);
        CargoDeliveryListBox.ItemsSource = CargoDeliveryItems;
        ExpiryWatchListBox.ItemsSource = ExpiryWatchDisplayItems;
        LoadAndApplyLicense();
        TraceMainWindowStartup("ItemsSource assignments done");

        Loaded += (_, _) =>
        {
            TraceMainWindowStartup("MainWindow Loaded entered");
            LoadHistoryFromCsv();
            TraceMainWindowStartup("LoadHistoryFromCsv done");
            RefreshQrCode();
            TraceMainWindowStartup("RefreshQrCode done");
            UpdateConnectionStatus(_service.ConnectionState, _service.ConnectedClients);
            RefreshTtacQuickLoginButtons();
            SetStartupRegistryEntry(false);
            StartupOnWindowsBootCheckBox.Visibility = Visibility.Collapsed;
            LoadMessages();
            if (!IsLicenseValid())
                Dispatcher.BeginInvoke(new Action(ShowLicenseOverlayStrict), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            else
            {
                StartLicenseHeartbeatTimer();
                _ = ValidateActiveLicenseOnlineBestEffortAsync();
            }

            // WebView2 اولین بار کمی زمان می‌برد. چند ثانیه بعد از بالا آمدن برنامه،
            // مرورگر داخلی تی‌تک را در پس‌زمینه آماده می‌کنیم تا هنگام ورود سریع‌تر باز شود.
            _ = Dispatcher.BeginInvoke(new Action(async () =>
            {
                try
                {
                    await Task.Delay(2500);
                    await EnsureTtTeckWebViewAsync();
                }
                catch { }
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

            // هشدار تاریخ انقضای نزدیک: تنظیمات را بارگذاری کن، چند ثانیه بعد از بالا آمدن
            // برنامه یک بار بررسی کن، و بعد هر ۶ ساعت یک‌بار دوباره بررسی کن (چون ممکن است
            // برنامه روزها باز بماند).
            LoadExpiryAlertSettings();
            PopulateExpiryFilterCombos();
            RefreshExpiryWatchDisplayList();
            _expiryAlertCheckTimer.Tick += async (_, _) => await CheckExpiryAlertsAsync();
            _expiryAlertCheckTimer.Start();
            _ = Dispatcher.BeginInvoke(new Action(async () =>
            {
                try
                {
                    await Task.Delay(8000);
                    await CheckExpiryAlertsAsync();
                }
                catch { }
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

            // بررسی بروزرسانی: یک بار چند ثانیه بعد از بالا آمدن برنامه، و بعد هر ۲۴ ساعت یک‌بار
            // دیگر (چون ممکن است برنامه روزها بدون بسته شدن باز بماند - داروخانه‌ها معمولاً برنامه
            // را در طول روز کاری باز نگه می‌دارند).
            _appUpdateCheckTimer.Tick += async (_, _) => await CheckForAppUpdateAsync(manualTrigger: false);
            _appUpdateCheckTimer.Start();
            _ = Dispatcher.BeginInvoke(new Action(async () =>
            {
                try
                {
                    await Task.Delay(5000);
                    await CheckForAppUpdateAsync(manualTrigger: false);
                }
                catch { }
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

            TraceMainWindowStartup("MainWindow Loaded completed");
        };

        _service.ScanReceived += async (_, args) =>
        {
            // نکته مهم: این رویداد به‌صورت synchronous از ترد پس‌زمینه‌ی پردازش صف در
            // ScanBridgeService فراخوانی می‌شود (قبل از اولین await). اگر هر خطایی از این
            // بلوک بیرون بزند، به آن ترد پس‌زمینه برمی‌گردد و می‌تواند کل پردازش اسکن‌های
            // بعدی را ساکت متوقف کند. به همین دلیل کل بدنه در try/catch پوشانده شده است.
            try
            {
                string incomingBarcode = args.Barcode ?? "";
                string originalDeviceName = args.DeviceName ?? "";

                // اگر به هر دلیل جای بارکد و نام دستگاه از سمت سرویس/CSV عوض شد، همین‌جا اصلاح شود.
                if (!IsLikelyScannedCode(incomingBarcode) && IsLikelyScannedCode(originalDeviceName))
                {
                    (incomingBarcode, originalDeviceName) = (originalDeviceName, incomingBarcode);
                }

                incomingBarcode = CleanBarcodeForExternalUse(incomingBarcode);

                // اگر QR اتصال خود برنامه اشتباهاً به عنوان اسکن برگشت، وارد تاریخچه و تی‌تک نشود.
                if (IsScanBridgePairPayload(incomingBarcode))
                    return;

                if (!HasLicenseModule("barcodeBridge"))
                {
                    Dispatcher.BeginInvoke(new Action(ShowLicenseOverlayStrict));
                    return;
                }

                // اگر پنجره‌ی استعلام قیمت باز است، اسکن داخل همان کادر فعال می‌نشیند
                // (نام / بارکد / ژنریک). تایپ کیبورد جداگانه قطع است تا متن دو بار نیاید.
                bool priceLookupActive = Dispatcher.CheckAccess()
                    ? PriceLookupOverlay.Visibility == Visibility.Visible
                    : Dispatcher.Invoke(() => PriceLookupOverlay.Visibility == Visibility.Visible);
                if (priceLookupActive)
                {
                    if (Dispatcher.CheckAccess())
                        ApplyScannedValueToPriceLookup(incomingBarcode);
                    else
                        Dispatcher.Invoke(() => ApplyScannedValueToPriceLookup(incomingBarcode));
                    return;
                }

                bool cargoModeActive = await Dispatcher.InvokeAsync(() => CargoDeliveryOverlay.Visibility == Visibility.Visible);
                if (cargoModeActive)
                {
                    await AddCargoDeliveryBarcodeAsync(incomingBarcode, showErrors: true);
                    return;
                }

                // اگر بانک «بارکد پرمصرف» در حالت دریافت فعال باشد (کاربر یک زیرگروه را برای دریافت
                // انتخاب کرده)، این بارکد به‌جای مسیر عادی تاریخچه/تی‌تک، به همان زیرگروه اضافه شود.
                bool highUsageCaptureHandled = await Dispatcher.InvokeAsync(() => TryCaptureHighUsageBarcode(incomingBarcode));
                if (highUsageCaptureHandled)
                    return;

                var timestampLocal = ToLocalTimestamp(args.TimestampUtc);
                var record = new ScanRecord(timestampLocal, incomingBarcode, GetDeviceDisplayName(originalDeviceName));

                var barcodeType = BarcodeDetector.DetectBarcodeType(incomingBarcode);
                record.Source = barcodeType;
                bool ttacAllowedByLicense = HasLicenseModule("ttac");
                bool shouldLookupTtTeck = ttacAllowedByLicense && _ttTeckSettings.IsEnabled && IsTtTeckLookupCandidate(incomingBarcode, barcodeType);

                if (shouldLookupTtTeck)
                    record.DrugName = GetTtTeckLookupPendingText();
                else if (ttacAllowedByLicense && _ttTeckSettings.IsEnabled)
                    record.DrugName = GetTtTeckLookupSkippedText();
                else
                    record.DrugName = "جستجو غیرفعال";

                // اول همه بارکدها ثبت شوند و پیام سبز دریافت بارکد نمایش داده شود؛ استعلام تی‌تک جداگانه انجام می‌شود.
                AddHistoryRecord(record);
                ShowScanToast(record, false, string.Empty);

                if (shouldLookupTtTeck)
                {
                    await LookupTtTeckForRecordAsync(record, false);
                    QueueReceiveStatusBarcode(record.Barcode);
                    if (HasValidTtacToken())
                        _ = AddReceiveStatusBarcodeAsync(record.Barcode, showErrors: false);
                }
            }
            catch (Exception ex)
            {
                LogBackgroundHandlerError(ex, "ScanReceived handler");
            }
        };

        _service.ConnectionStatusChanged += (_, args) => UpdateConnectionStatus(args.State, args.ConnectedClients);

        // با تغییر IP شبکه، QR اتصال و متن آی‌پی خودکار به‌روز می‌شوند (دکمه‌ی «ساخت مجدد بارکد» حذف شده).
        _service.LanIpChanged += (_, _) =>
        {
            Dispatcher.BeginInvoke(new Action(OnLanIpChanged), System.Windows.Threading.DispatcherPriority.Background);
        };
        _service.ConnectedDevicesChanged += (_, args) => UpdateDeviceRows(args.Devices);

        // پل ADB: اتصال صفر-ضربه‌ی گوشی با کابل (بدون Tethering) — adb reverse روی پورت 5050
        _adbBridge = new Services.AdbUsbBridge();
        _adbBridge.StatusTip += status =>
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                switch (status)
                {
                    case "UNAUTHORIZED":
                        ShowStyledMessage(
                            "تأیید روی گوشی لازم است",
                            "گوشی با کابل وصل شده ولی «USB Debugging» هنوز تأیید نشده. روی صفحه‌ی گوشی پیام «Allow USB debugging?» را با زدن تیک Always allow و OK تأیید کنید.",
                            false);
                        break;
                    case "REVERSE_ON":
                        ShowStyledMessage(
                            "اتصال سریع کابل فعال شد ⚡",
                            "گوشی تأیید شد. از این به بعد اپ گوشی با «اتصال با کابل USB» بدون هیچ تنظیمی، فقط با وصل کردن کابل متصل می‌شود.",
                            false);
                        break;
                }
            }));
        };
        _adbBridge.Start();

        _service.UnexpectedDisconnection += (_, _) =>
        {
            Dispatcher.BeginInvoke(() =>
            {
                ((App)System.Windows.Application.Current).ShowUnexpectedDisconnectAlert();
            });
        };

        // وقتی یک سیستم دیگر روی همین شبکه‌ی محلی با همان لایسنس، تنظیمات TtTeck/آستانه‌های تاریخ
        // نزدیک/وضعیت فعال بودن بله را عوض کند، همان‌جا با WebSocket برای این سیستم هم می‌فرستد -
        // اینجا اعمالش می‌کنیم.
        _service.PeerDesktopSettingsReceived += (_, args) => ApplyPeerDesktopSettings(args.PayloadJson, args.VersionUtcMs);

        // ویژگی «ورود اطلاعات از راه دور» (فرم ثبت شیرخشک روی گوشی) - نگاه کنید به
        // MainWindow.RemoteFormulaEntry.cs
        _service.RemoteEntryValueReceived += (_, args) => HandleRemoteEntryValueFromPhone(args.Barcode, args.StepId, args.Value);
        _service.RemoteEntrySubmitReceived += (_, args) => HandleRemoteEntrySubmitFromPhone(args.Barcode);
        _service.RemoteEntryBackReceived += (_, args) => HandleRemoteEntryBackFromPhone(args.Barcode);
        _service.RemoteEntryRepeatArmReceived += (_, _) => HandleRemoteEntryRepeatArmFromPhone();
        _service.PriceLookupPhoneMessageReceived += (_, args) => HandlePriceLookupPhoneMessage(args.Type, args.Json);

        // همگام‌سازی بانک بارکد پرمصرف بین چند سیستم هم‌شبکه با همان لایسنس - نگاه کنید به
        // MainWindow.HighUsageBarcode.cs (ApplyHighUsageBarcodeOperation/ApplyHighUsageBarcodeSnapshot).
        _service.HighUsageBarcodeOperationReceived += (_, args) => ApplyHighUsageBarcodeOperation(args.OpJson, args.FromComputerId);
        _service.HighUsageBarcodeSnapshotReceived += (_, args) => ApplyHighUsageBarcodeSnapshot(args.PayloadJson, args.VersionUtcMs, args.FromComputerId);

        TraceMainWindowStartup("MainWindow constructor completed");
    }

    private void EnsureBundledDataFilesAvailable()
    {
        ExtractEmbeddedTextFileIfMissing("special-prescription-generics.txt");
        ExtractEmbeddedTextFileIfMissing("No-Rx-Formula.txt");
        ExtractEmbeddedTextFileIfMissing("Rx-Formula.txt");
    }

    private void ExtractEmbeddedTextFileIfMissing(string fileName)
    {
        try
        {
            string targetPath = Path.Combine(AppContext.BaseDirectory, fileName);
            if (File.Exists(targetPath) && new FileInfo(targetPath).Length > 0)
                return;

            var assembly = Assembly.GetExecutingAssembly();
            string? resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase)
                                     || name.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));
            if (string.IsNullOrWhiteSpace(resourceName))
                return;

            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
                return;

            Directory.CreateDirectory(Path.GetDirectoryName(targetPath) ?? AppContext.BaseDirectory);
            using var fileStream = File.Create(targetPath);
            stream.CopyTo(fileStream);
        }
        catch { }
    }

    // مسیر نصب (AppContext.BaseDirectory) گاهی write-protected است (مثلاً برنامه داخل Program
    // Files و کاربر بدون دسترسی مدیر نصب کرده). قبلاً در این حالت این دو فایل لاگ بی‌صدا هیچ‌وقت
    // نوشته نمی‌شدند - یعنی خود گزارش تشخیصی هم که کاربر برای پشتیبانی می‌فرستد، همیشه خالی
    // می‌ماند، بدون هیچ هشداری که چرا (باگ گزارش ممیزی). حالا اگر نوشتن در مسیر نصب شکست بخورد،
    // یک‌بار هم پوشه‌ی AppData کاربر (تقریباً همیشه قابل‌نوشتن) امتحان می‌شود؛ GenerateDiagnosticsReport
    // هم همین مسیر جایگزین را چک می‌کند (نگاه کنید به GetAppLogFilePathForReading).
    private static void AppendAppLogLine(string fileName, string text)
    {
        try
        {
            string path = Path.Combine(AppContext.BaseDirectory, fileName);
            File.AppendAllText(path, text);
            return;
        }
        catch { }

        try
        {
            string fallbackDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Scanbridge", "logs");
            Directory.CreateDirectory(fallbackDir);
            File.AppendAllText(Path.Combine(fallbackDir, fileName), text);
        }
        catch { }
    }

    // برای GenerateDiagnosticsReport: مسیر واقعی یک فایل لاگ را برمی‌گرداند - اول مسیر نصب، اگر
    // آنجا وجود نداشت (مثلاً چون write-protected بوده و AppendAppLogLine مجبور شده fallback را
    // استفاده کند)، مسیر AppData را چک می‌کند.
    private static string? GetAppLogFilePathForReading(string fileName)
    {
        string primaryPath = Path.Combine(AppContext.BaseDirectory, fileName);
        if (File.Exists(primaryPath))
            return primaryPath;

        try
        {
            string fallbackPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Scanbridge", "logs", fileName);
            if (File.Exists(fallbackPath))
                return fallbackPath;
        }
        catch { }

        return null;
    }

    private static void TraceMainWindowStartup(string message)
    {
        AppendAppLogLine("startup-trace.log", $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
    }

    // برای گزارش خطاهایی که در هندلرهای پس‌زمینه/رویدادها رخ می‌دهند (مثل ScanReceived)
    // و نباید کل برنامه یا صف پردازش اسکن را متوقف کنند، ولی باید جایی ثبت شوند تا
    // در صورت نیاز قابل بررسی باشند.
    private static void LogBackgroundHandlerError(Exception ex, string section)
    {
        AppendAppLogLine("startup-error.log", $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {section}\n{ex}\n\n");
    }


    private void MainWindow_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Escape)
            return;

        e.Handled = true;

        if (StyledMessageOverlay.Visibility == Visibility.Visible)
        {
            CloseStyledMessage();
            return;
        }
        if (OperationDetailsOverlay.Visibility == Visibility.Visible)
        {
            CloseOperationDetails();
            return;
        }
        if (RepeatFormulaBarcodeOverlay.Visibility == Visibility.Visible)
        {
            CloseRepeatFormulaBarcodeDialog();
            return;
        }
        if (TtacLoginOverlay.Visibility == Visibility.Visible)
        {
            CloseTtacLoginOverlay();
            return;
        }
        if (TtTeckWebViewOverlay.Visibility == Visibility.Visible)
        {
            TtTeckWebViewOverlay.Visibility = Visibility.Collapsed;
            MainContent.Effect = null;
            return;
        }
        if (TtTeckRegistrationOverlay.Visibility == Visibility.Visible)
        {
            CloseTtTeckRegistrationOverlay();
            return;
        }
        if (RetryTtTeckOverlay.Visibility == Visibility.Visible)
        {
            CloseRetryTtTeckOverlay();
            return;
        }
        if (CargoDeliveryOverlay.Visibility == Visibility.Visible)
        {
            CloseCargoDeliveryPanel();
            return;
        }
        if (ReceiveStatusOverlay.Visibility == Visibility.Visible)
        {
            CloseReceiveStatusPanel();
            return;
        }
        if (TtacPanelOverlay.Visibility == Visibility.Visible)
        {
            CloseTtacPanel();
            return;
        }
        if (DateRangeFilterOverlay.Visibility == Visibility.Visible)
        {
            CloseDateRangeFilter();
            return;
        }
        if (ProductDetailsOverlay.Visibility == Visibility.Visible)
        {
            CloseProductDetails();
            return;
        }
        if (ExportTypeSelectOverlay.Visibility == Visibility.Visible)
        {
            CloseExportTypeSelection();
            return;
        }
        if (ConfirmDeleteHistoryOverlay.Visibility == Visibility.Visible)
        {
            CloseConfirmDeleteOverlay();
            return;
        }
        if (HistoryOverlay.Visibility == Visibility.Visible)
        {
            CloseHistoryOverlay();
            return;
        }
        if (MessagesOverlay.Visibility == Visibility.Visible)
        {
            CloseMessagesOverlay();
            return;
        }
        if (LanguageOverlay.Visibility == Visibility.Visible)
        {
            CloseLanguageOverlay();
            return;
        }
        if (DeviceAliasOverlay.Visibility == Visibility.Visible)
        {
            CloseDeviceAliasEditor();
            return;
        }
        if (FindName("LicenseOverlay") is FrameworkElement licenseOverlay && licenseOverlay.Visibility == Visibility.Visible)
        {
            CloseLicenseOverlaySafe();
            return;
        }
        if (SupportOverlay.Visibility == Visibility.Visible)
        {
            CloseSupportOverlay();
            return;
        }
        if (TtTeckSettingsOverlay.Visibility == Visibility.Visible)
        {
            CloseTtTeckSettings();
            return;
        }
    }

    // ---------- Language Switching ----------

    private void LanguageButton_Click(object sender, RoutedEventArgs e)
    {
        LanguageOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void LanguageOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseLanguageOverlay();
    }

    private void LanguageCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseLanguageOverlay()
    {
        LanguageOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void PersianOptionButton_Click(object sender, RoutedEventArgs e)
    {
        _localization.CurrentLanguage = AppLanguage.Persian;
        _localization.SaveSettings(_localization.CurrentLanguage);
        CloseLanguageOverlay();
    }

    private void EnglishOptionButton_Click(object sender, RoutedEventArgs e)
    {
        _localization.CurrentLanguage = AppLanguage.English;
        _localization.SaveSettings(_localization.CurrentLanguage);
        CloseLanguageOverlay();
    }

    private void UpdateLanguageUI()
    {
        this.FlowDirection = _localization.CurrentLanguage == AppLanguage.English
            ? System.Windows.FlowDirection.LeftToRight
            : System.Windows.FlowDirection.RightToLeft;

        AppTitleText.Text = "Scanbridge";
        MessagesButtonText.Text = _localization.GetString("Messages");
        BarcodeTitle.Text = _localization.GetString("BarcodeTitle");
        ConnectedDevicesTitle.Text = _localization.GetString("ConnectedDevices");
        NoDevicesText.Text = _localization.GetString("NoDevices");
        SystemInfoTitle.Text = _localization.GetString("SystemInfo");
        StartupOnWindowsBootCheckBox.Content = _localization.GetString("AutoStartup");
        UserPanelTitle.Text = _localization.GetString("UserPanel");
        HistoryButton.Content = _localization.GetString("History");
        TtacPanelButton.Content = _localization.GetString("TtTeckPanel");
        CargoDeliveryPanelButton.Content = _localization.GetString("CargoDelivery");
        ReceiveStatusPanelButton.Content = _localization.GetString("ReceiveStatus");
        PrintQrButton.Content = _localization.GetString("Print");
        SupportButton.Content = _localization.GetString("Support");
        SupportMessageText.Text = _localization.GetString("SupportMessageText");
        if (SupportSiteButton != null) SupportSiteButton.Content = _localization.GetString("Website");
        if (SupportOverlayTitleText != null) SupportOverlayTitleText.Text = _localization.GetString("SupportOverlayTitle");
        if (SupportOverlayMessageText != null) SupportOverlayMessageText.Text = _localization.GetString("SupportMessageText");
        if (SupportOverlayWhatsAppButton != null) SupportOverlayWhatsAppButton.Content = _localization.GetString("WhatsAppContact");
        if (SupportOverlaySiteButton != null) SupportOverlaySiteButton.Content = _localization.GetString("Website");
        if (SupportOverlayCloseButton != null) SupportOverlayCloseButton.Content = _localization.GetString("Close");
        RemoteSupportButton.Content = _localization.GetString("RemoteSupport");
        MessagesHeaderText.Text = _localization.GetString("Messages");
        NoMessagesText.Text = _localization.GetString("NoMessages");
        LanguageSelectTitle.Text = _localization.GetString("LanguageSelectTitle");
        LanguageSelectSubtitle.Text = _localization.GetString("SelectLanguage");
        SetLanguageOptionTexts(_localization.GetString("PersianOption"), _localization.GetString("EnglishOption"));
        HistoryTitle.Text = _localization.GetString("HistoryTitle");
        HistoryClearButton.Content = _localization.GetString("ClearHistory");
        SettingsButtonText.Text = _localization.GetString("Settings");
        TtTeckSettingsTitle.Text = _localization.GetString("Settings");
        TtTeckEnableText.Text = _localization.GetString("TtTeckEnableText");
        ExportTypeTitle.Text = _localization.GetString("SelectExportType");
        ExportTypeDescription.Text = _localization.GetString("ExportTypeDescription");
        ExportAllTitle.Text = _localization.GetString("ExportAllTitle");
        ExportAllDesc.Text = _localization.GetString("ExportAllDesc");
        ExportTtTeckTitle.Text = _localization.GetString("ExportTtTeckTitle");
        ExportTtTeckDesc.Text = _localization.GetString("ExportTtTeckDesc");
        if (LanguageCardTitle != null) LanguageCardTitle.Text = _localization.GetString("LanguageCardTitle");
        ChangeLanguageButton.Content = _localization.GetString("ChangeLanguage");
        UpdateSystemInfo();

        UpdateAdditionalLocalizedTexts();
        UpdateStaticLanguageTexts();
        UpdateHistoryCopyButtonsText();
        MessagesList?.Items.Refresh();
        HistoryListBox?.Items.Refresh();
        TtTeckHistoryListBox?.Items.Refresh();
        try
        {
            ApplyHistoryFilters();
        }
        catch { }
    }

    private void SetLanguageOptionTexts(string persianOptionText, string englishOptionText)
    {
        SetNamedTextIfExists("PersianLanguageOptionText", persianOptionText);
        SetNamedTextIfExists("EnglishLanguageOptionText", englishOptionText);
        UpdateStaticLanguageTexts();
    }

    private void SetNamedTextIfExists(string name, string text)
    {
        if (FindName(name) is TextBlock textBlock)
            textBlock.Text = text;
    }

    private void SetNamedFlowDirectionIfExists(string name, System.Windows.FlowDirection flowDirection)
    {
        if (FindName(name) is FrameworkElement element)
            element.FlowDirection = flowDirection;
    }

    private void SetTtacRegistrationProductHelpText(string text)
    {
        SetNamedTextIfExists("TtTeckRegistrationProductHelpText", text);
        ReplaceStaticTextBlockText("برای ثبت سریع، از Enter استفاده کنید. فقط برای بازنشانی کپچا به موس نیاز دارید.", text);
        ReplaceStaticTextBlockText("For fast registration use Enter. Use the mouse only to reset the captcha.", text);
    }

    private void UpdateStaticLanguageTexts()
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        string persianOption = _localization.GetString("PersianOption");
        string englishOption = _localization.GetString("EnglishOption");

        ReplaceStaticTextBlockText("فارسی / Persian", persianOption);
        ReplaceStaticTextBlockText("فارسی", persianOption);
        ReplaceStaticTextBlockText("Persian", persianOption);
        ReplaceStaticTextBlockText("English / انگلیسی", englishOption);
        ReplaceStaticTextBlockText("انگلیسی", englishOption);
        ReplaceStaticTextBlockText("English", englishOption);
        ReplaceStaticTextBlockText("Select Language", _localization.GetString("SelectLanguage"));
    }

    private void ReplaceStaticTextBlockText(string oldText, string newText)
    {
        if (string.IsNullOrEmpty(oldText))
            return;

        try
        {
            foreach (var textBlock in FindVisualChildren<TextBlock>(this))
            {
                if (string.Equals(textBlock.Text, oldText, StringComparison.Ordinal))
                    textBlock.Text = newText;
            }
        }
        catch { }
    }

    private void UpdateAdditionalLocalizedTexts()
    {
        Title = "Scanbridge";
        AppTitleText.Text = "Scanbridge";

        ExportExcelButton.Content = "📊 Excel";
        ConfirmDeleteTitle.Text = _localization.GetString("ConfirmDeleteTitle");
        ConfirmDeleteMessage.Text = _localization.GetString("ConfirmDeleteMessage");
        ConfirmDeleteWarning.Text = _localization.GetString("ConfirmDeleteWarning");
        ConfirmDeleteYesButton.Content = _localization.GetString("ConfirmDeleteYes");
        ConfirmDeleteNoButton.Content = _localization.GetString("ConfirmDeleteNo");
        SuccessTitle.Text = _localization.GetString("ExportSuccessTitle");
        SuccessMessage.Text = _localization.GetString("ExportSuccessMessage");
        ExportTypeOkButton.Content = _localization.GetString("OkConfirm");
        ExportTypeCancelButton.Content = _localization.GetString("CancelButton");
        TtTeckSaveButton.Content = _localization.GetString("Save");
        TtTeckCancelButton.Content = _localization.GetString("CancelButton");
        StyledMessageOkButton.Content = _localization.GetString("OK");

        UpdateHistoryFilterLocalizedTexts();
        UpdateProductDetailsLocalizedTexts();
        UpdateTtTeckRegistrationLocalizedTexts();
        UpdateTtacLoginLocalizedTexts();
        UpdateExpiryAlertSettingsLocalizedTexts();
        UpdateTtacSavedLoginsLocalizedTexts();
        UpdateCargoDeliveryLocalizedTexts();
        RefreshExpiryWatchDisplayList();

        UpdateTtTeckFilterButtonText();
    }

    private void UpdateExpiryAlertSettingsLocalizedTexts()
    {
        if (ExpiryAlertSectionTitle == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;

        ExpiryAlertSectionTitle.Text = _localization.GetString("NearExpiryAlert");
        ExpiryAlertSectionDescription.Text = _localization.GetString("TtTeckPlusOnlyWhenAnItemRegisteredInCargoDeliveryHasANearExpirationDateItIsFlaggedInTheSamePanel");
        ExpiryThresholdMonthsLabel.Text = _localization.GetString("HowManyMonthsBeforeExpiryToAlert");
        RepeatReminderDaysLabel.Text = _localization.GetString("RepeatTheReminderEveryHowManyDays");
        BaleSectionTitle.Text = _localization.GetString("BaleReminder");
        if (BaleActivateButton != null && (_baleActivationPollTimer == null))
            BaleActivateButton.Content = _localization.GetString("ActivateBaleReminder");
        if (BaleTestMessageButton != null)
            BaleTestMessageButton.Content = _localization.GetString("SendTestMessage");
        BaleHintText.Text = _localization.GetString("TappingThisButtonOpensBaleJustTapStartThereAndNearExpiryAlertsWillAlsoBeSentToYouOnBaleFromThenOn");

        UpdateBaleConnectionStatusText();

        if (ExpiryFilterLabel != null)
        {
            ExpiryFilterLabel.Text = _localization.GetString("DisplayRange");
            ExpiryFilterFromLabel.Text = _localization.GetString("From");
            ExpiryFilterToLabel.Text = _localization.GetString("To");
            ExpiryFilterApplyButton.Content = _localization.GetString("ApplyFilter");
            if (ExpiryFilterClearButton != null)
                ExpiryFilterClearButton.Content = _localization.GetString("ClearFilter2");
            ExpiryWatchExportExcelButton.Content = _localization.GetString("ExportExcel");
            PopulateExpiryFilterCombos();
        }
    }

    private string GetHistoryCopyButtonText()
    {
        return _localization.GetString("Copy");
    }

    private string GetHistoryCopiedButtonText()
    {
        return _localization.GetString("Copied");
    }

    private void UpdateHistoryFilterLocalizedTexts()
    {
        if (HistorySearchLabel == null)
            return;

        HistorySearchLabel.Text = _localization.GetString("SearchHistory");
        HistorySearchTextBox.ToolTip = _localization.GetString("SearchHistoryTooltip");
        HistoryDateFilterLabel.Text = _localization.GetString("DateAndTimeFilter");
        DateRangeTitle.Text = _localization.GetString("SelectDateRange");
        DateRangeFromTitle.Text = _localization.GetString("FromDateTime");
        DateRangeToTitle.Text = _localization.GetString("ToDateTime");
        SetDateRangePartLabels(
            _localization.GetString("Year"),
            _localization.GetString("Month"),
            _localization.GetString("Day"),
            _localization.GetString("Hour"),
            _localization.GetString("Minute"));
        DateRangeApplyButton.Content = _localization.GetString("ApplyFilter");
        DateRangeClearButton.Content = _localization.GetString("ClearFilter");
        DateRangeCancelButton.Content = _localization.GetString("CancelButton");
        ExportPdfButton.Content = "📄 PDF";
        RetryTtTeckTitle.Text = _localization.GetString("RetryTtTeckTitle");
        RetryTtTeckReasonLabel.Text = _localization.GetString("RetryTtTeckReason");
        RetryTtTeckConfirmButton.Content = _localization.GetString("RetryTtTeckConfirm");
        RetryTtTeckCancelButton.Content = _localization.GetString("CancelButton");

        UpdateHistoryDateRangeButtonText();
    }

    private void SetDateRangePartLabels(string year, string month, string day, string hour, string minute)
    {
        DateRangeFromYearLabel.Text = year;
        DateRangeFromMonthLabel.Text = month;
        DateRangeFromDayLabel.Text = day;
        DateRangeFromHourLabel.Text = hour;
        DateRangeFromMinuteLabel.Text = minute;

        DateRangeToYearLabel.Text = year;
        DateRangeToMonthLabel.Text = month;
        DateRangeToDayLabel.Text = day;
        DateRangeToHourLabel.Text = hour;
        DateRangeToMinuteLabel.Text = minute;
    }

    private void UpdateProductDetailsLocalizedTexts()
    {
        if (ProductDetailsTitle == null)
            return;

        ProductDetailsTitle.Text = _localization.GetString("ProductDetailsTitle");
        ProductDetailsPersianNameLabel.Text = _localization.GetString("PersianNameLabel");
        ProductDetailsEnglishNameLabel.Text = _localization.GetString("EnglishNameLabel");
        ProductDetailsBarcodeLabel.Text = _localization.GetString("BarcodeLabel");
        ProductDetailsDateTimeLabel.Text = _localization.GetString("DateTimeLabel");
        ProductDetailsDeviceLabel.Text = _localization.GetString("DeviceLabel");
        ProductDetailsStatusLabel.Text = _localization.GetString("StatusLabel");
        ProductDetailsCopyBarcodeButton.Content = _localization.GetString("CopyBarcode");
        ProductDetailsCloseButton.Content = _localization.GetString("Close");
        ProductDetailsExtraTitle.Text = _localization.GetString("AllTtTeckInfo");
        DeviceAliasTitle.Text = _localization.GetString("EditDeviceName");
        DeviceAliasOriginalLabel.Text = _localization.GetString("CurrentDeviceName");
        DeviceAliasNewNameLabel.Text = _localization.GetString("CustomName");
        DeviceAliasSaveButton.Content = _localization.GetString("Save");
        DeviceAliasResetButton.Content = _localization.GetString("RemoveCustomName");
        DeviceAliasCancelButton.Content = _localization.GetString("CancelButton");
        DeviceReconnectTitle.Text = _localization.GetString("DeviceReconnectTitle");
        DeviceReconnectDescription.Text = _localization.GetString("DeviceReconnectDesc");
        AllowDeviceReconnectButton.Content = _localization.GetString("AllowPhonesReconnect");
    }

    private void UpdateSystemInfo()
    {
        ComputerNameText.Text = _localization.GetFormattedString("SystemInfoComputer", _service.ComputerName);
        LanIpText.Text = _localization.GetFormattedString("SystemInfoIp", _service.LanIp);
        PortText.Text = _localization.GetFormattedString("SystemInfoPort", ScanBridgeService.Port);
        if (AllowDeviceReconnectMainButton != null)
        {
            AllowDeviceReconnectMainButton.Content = _localization.GetString("AllowPhonesReconnect");
            AllowDeviceReconnectMainButton.Visibility = Visibility.Collapsed;
        }

        UpdateTtacConnectionStatusUI();
    }

    private string GetTtacTokenRemainingText(bool english)
    {
        if (string.IsNullOrWhiteSpace(_ttacAccessTokenOverride) || DateTime.UtcNow >= _ttacAccessTokenExpiresAtUtc)
            return string.Empty;

        TimeSpan remaining = _ttacAccessTokenExpiresAtUtc - DateTime.UtcNow;
        if (remaining.TotalSeconds < 0)
            remaining = TimeSpan.Zero;

        string time = remaining.TotalHours >= 1
            ? $"{(int)remaining.TotalHours:00}:{remaining.Minutes:00}:{remaining.Seconds:00}"
            : $"{remaining.Minutes:00}:{remaining.Seconds:00}";

        return _localization.GetFormattedString("TokenRemaining", time);
    }

    private bool HasValidTtacToken()
    {
        return !string.IsNullOrWhiteSpace(_ttacAccessTokenOverride) && DateTime.UtcNow < _ttacAccessTokenExpiresAtUtc;
    }

    // با هر تیک تایمر شمارش معکوس (هر ثانیه) صدا زده می‌شود. اگر دقیقاً همین بین این تیک و تیک قبلی
    // توکن تی‌تک از حالت معتبر به نامعتبر رفته باشد (یعنی واقعاً «تمام شده»، نه این‌که از اول متصل نبوده)
    // هم روی سیستم هم روی گوشیِ متصل پیغام «توکن تمام شد» نشان داده می‌شود.
    private void TtacTokenCountdownTick()
    {
        UpdateTtacTokenValidityTracking(HasValidTtacToken());
        UpdateTtacConnectionStatusUI();
    }

    // suppressExpiryNotification: وقتی خودِ کاربر آگاهانه قطع اتصال می‌کند یا داروخانه را عوض می‌کند
    // (که پیغام مخصوص به خودش را دارد)، این گذار نباید دوباره پیغام «توکن تمام شد» را نشان دهد.
    private void UpdateTtacTokenValidityTracking(bool isValidNow, bool suppressExpiryNotification = false)
    {
        if (_ttacTokenWasValidLastCheck && !isValidNow && !suppressExpiryNotification && HasLicenseModule("ttac"))
        {
            NotifyTtacTokenExpired();
        }
        _ttacTokenWasValidLastCheck = isValidNow;
    }

    private void NotifyTtacTokenExpired()
    {
        try
        {
            string title = _localization.GetString("TtacTokenExpiredTitle");
            string message = _localization.GetString("TtacTokenExpiredMessage");
            ShowStyledMessage(title, message, true);
            _service?.BroadcastAlert(title, message, false);
        }
        catch { }
    }

    private void UpdateTtacConnectionStatusUI()
    {
        if (FindName("TtacConnectionStatusPanel") is not Border panel ||
            FindName("TtacConnectionStatusIcon") is not TextBlock icon ||
            FindName("TtacConnectionStatusTitle") is not TextBlock title ||
            FindName("TtacConnectionStatusText") is not TextBlock statusText ||
            FindName("TtacDisconnectButton") is not System.Windows.Controls.Button disconnectButton)
        {
            return;
        }

        var openSiteButton = FindName("TtacOpenSiteButton") as System.Windows.Controls.Button;

        bool licenseAllowsTtac = HasLicenseModule("ttac");
        panel.Visibility = licenseAllowsTtac ? Visibility.Visible : Visibility.Collapsed;
        if (!licenseAllowsTtac)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        bool connected = HasValidTtacToken();

        title.Text = _localization.GetString("TTACConnectionStatus");
        disconnectButton.ToolTip = _localization.GetString("DisconnectTTAC");
        if (openSiteButton != null)
        {
            openSiteButton.Content = _localization.GetString("OpenTTACSite");
            openSiteButton.ToolTip = _localization.GetString("OpenTTACWebsiteInTheInternalBrowser");
            openSiteButton.Visibility = Visibility.Visible;
        }

        if (connected)
        {
            panel.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xEC, 0xFD, 0xF5));
            panel.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81));
            icon.Text = "●";
            icon.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81));
            string displayName = string.IsNullOrWhiteSpace(_ttacPharmacyDisplayName)
                ? TryExtractTtacDisplayNameFromToken(_ttacAccessTokenOverride)
                : _ttacPharmacyDisplayName;
            _ttacPharmacyDisplayName = displayName ?? string.Empty;
            string remainingText = GetTtacTokenRemainingText(english);
            string baseStatus = !string.IsNullOrWhiteSpace(displayName)
                ? (_localization.GetFormattedString("ConnectedWithName", displayName))
                : (_localization.GetString("Connected"));
            statusText.Text = string.IsNullOrWhiteSpace(remainingText) ? baseStatus : $"{baseStatus} | {remainingText}";
            statusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x06, 0x5F, 0x46));
            disconnectButton.Visibility = Visibility.Visible;
        }
        else
        {
            panel.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF3, 0xF4, 0xF6));
            panel.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xE5, 0xE7, 0xEB));
            icon.Text = "●";
            icon.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF));
            statusText.Text = _localization.GetString("NotConnected");
            statusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x6B, 0x72, 0x80));
            disconnectButton.Visibility = Visibility.Collapsed;
        }
    }

    private void TtacOpenSiteButton_Click(object sender, RoutedEventArgs e)
    {
        // همه‌ی مسیرهای ورود به تی‌تک یکسان: اگر حسابی ذخیره شده، پنجره‌ی انتخاب داروخانه باز
        // می‌شود و اگر هیچ حسابی نیست، فرم ذخیره‌ی حساب در تنظیمات بالا می‌آید.
        ShowTtacLoginOverlay();
    }

    private async void TtacDisconnectButton_Click(object sender, RoutedEventArgs e)
    {
        await DisconnectTtacSessionAsync(showMessage: true);
    }

    private async Task DisconnectTtacSessionAsync(bool showMessage)
    {
        SaveCurrentReceiveStatusItemsForCurrentPharmacy();
        SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
        SaveCurrentHistoryForCurrentPharmacy();
        SaveTtacRegistrationHistory();
        _ttacAccessTokenOverride = null;
        _ttacPharmacyDisplayName = string.Empty;
        _ttacAccessTokenExpiresAtUtc = DateTime.MinValue;
        UpdateTtacTokenValidityTracking(false, suppressExpiryNotification: true);
        _pendingTtacRetryAction = null;
        _pendingTtacRetryLabel = null;
        _pendingTtacAutofillUsername = null;
        _ttacRetryUsername = null;
        _ttacQuickLoginInProgress = false;
        ReceiveStatusItems.Clear();
        _receiveStatusKnownBarcodes.Clear();
        _queuedReceiveStatusBarcodes.Clear();
        _receiveStatusLoadedPharmacyKey = string.Empty;
        _cargoDeliveryLoadedPharmacyKey = string.Empty;
        CargoDeliveryItems.Clear();
        _cargoDeliveryKnownBarcodes.Clear();
        _ttacRegistrationHistoryLoadedPharmacyKey = "default";
        _ttacRegistrationHistory.Clear();
        _historyLoadedPharmacyKey = "default";
        // قبلاً اینجا فقط HistoryItems.Clear() + ApplyHistoryFilters() صدا زده می‌شد - یعنی
        // تاریخچه‌ی واقعیِ روی دیسکِ بخش «default» هیچ‌وقت دوباره بارگذاری نمی‌شد و پنل تاریخچه
        // تا ری‌استارت بعدی برنامه خالی می‌ماند، حتی اگر «default» قبلاً اسکن‌هایی داشت
        // (باگ گزارش ممیزی، «موارد کوچک‌تر»). LoadHistoryFromCsv خودش هم پاک‌سازی هم بارگذاری
        // واقعی از فایل CSV همین کلید را انجام می‌دهد.
        LoadHistoryFromCsv();

        try
        {
            if (_ttTeckWebView?.CoreWebView2 != null)
            {
                try
                {
                    await _ttTeckWebView.CoreWebView2.ExecuteScriptAsync("try{localStorage.clear();sessionStorage.clear();}catch(e){}");
                }
                catch { }

                try
                {
                    _ttTeckWebView.CoreWebView2.CookieManager.DeleteAllCookies();
                }
                catch { }

                try
                {
                    _ttTeckWebView.CoreWebView2.Navigate("about:blank");
                }
                catch { }
            }
        }
        catch { }

        UpdateTtacConnectionStatusUI();

        // با خروج از تی‌تک، دیگر مشخص نیست کدام داروخانه فعال است (_ttacPharmacyDisplayName چند
        // خط بالاتر خالی شد)؛ اگر اینجا لیست «تاریخ نزدیک» را رفرش نکنیم، نشان قرمز روی دکمه‌ی
        // «تاریخ نزدیک» همچنان تعداد داروخانه‌ی قبلی (که دیگر فعال نیست) را نشان می‌دهد - چون آن
        // نشان فقط با ورود به تی‌تک یا باز کردن دستی همان پنل به‌روز می‌شد، نه با خروج.
        RefreshExpiryWatchDisplayList();

        if (showMessage)
        {
            ShowStyledMessage(
                _localization.GetString("TTACDisconnectedSuccess"),
                _localization.GetString("TTACTokenAndInternalBrowserSessionWereCleared"),
                isError: true,
                customIcon: "✗");
        }
    }

    private void UpdateHistoryCopyButtonsText()
    {
        string copyButtonText = GetHistoryCopyButtonText();
        UpdateListBoxCopyButtonsText(HistoryListBox, copyButtonText);
        UpdateListBoxCopyButtonsText(TtTeckHistoryListBox, copyButtonText);
        UpdateListBoxCopyButtonsText(TtacPanelListBox, copyButtonText);
    }

    private void UpdateListBoxCopyButtonsText(System.Windows.Controls.ListBox listBox, string copyButtonText)
    {
        foreach (var item in listBox.Items)
        {
            if (listBox.ItemContainerGenerator.ContainerFromItem(item) is ListBoxItem container)
            {
                foreach (var button in FindVisualChildren<System.Windows.Controls.Button>(container))
                {
                    if (button.Tag is string)
                        button.Content = copyButtonText;
                }
            }
        }
    }

    private static IEnumerable<T> FindVisualChildren<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            DependencyObject child = VisualTreeHelper.GetChild(parent, i);
            if (child is T result)
                yield return result;

            foreach (var childOfChild in FindVisualChildren<T>(child))
                yield return childOfChild;
        }
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            DependencyObject child = VisualTreeHelper.GetChild(parent, i);
            if (child is T result)
                return result;
            T? childOfChild = FindVisualChild<T>(child);
            if (childOfChild != null)
                return childOfChild;
        }
        return null;
    }

    private static string CleanBarcodeForExternalUse(string? barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return string.Empty;

        string cleaned = new string(barcode
                .Trim()
                .Where(c => !char.IsControl(c)
                            && c != '\u200E'
                            && c != '\u200F'
                            && c != '\u202A'
                            && c != '\u202B'
                            && c != '\u202C'
                            && c != '\u202D'
                            && c != '\u202E')
                .ToArray())
            .Trim();

        // \u0628\u0631\u062E\u06CC \u0627\u0633\u06A9\u0646\u0631\u0647\u0627 \u067E\u06CC\u0634 \u0627\u0632 \u0628\u0627\u0631\u06A9\u062F \u0648\u0627\u0642\u0639\u06CC \u067E\u06CC\u0634\u0648\u0646\u062F\u0647\u0627\u06CC\u06CC \u0645\u062B\u0644 ]C1 \u06CC\u0627 ]d2 \u0645\u06CC\u200C\u0641\u0631\u0633\u062A\u0646\u062F (\u0647\u0645\u0627\u0646 \u067E\u06CC\u0634\u0648\u0646\u062F\u0647\u0627\u06CC\u06CC
        // \u06A9\u0647 DrugLookupService.ExtractUIDFromBarcode \u0647\u0645 \u062D\u0630\u0641 \u0645\u06CC\u200C\u06A9\u0646\u062F). \u0627\u06AF\u0631 \u0627\u06CC\u0646 \u067E\u06CC\u0634\u0648\u0646\u062F\u0647\u0627 \u0627\u06CC\u0646\u062C\u0627 \u062D\u0630\u0641
        // \u0646\u0634\u0648\u0646\u062F\u060C BarcodeDetector.DetectBarcodeType \u0628\u0627\u0631\u06A9\u062F \u0645\u0639\u062A\u0628\u0631 \u062A\u06CC\u200C\u062A\u06A9 \u0631\u0627 Unknown/QRCode \u062A\u0634\u062E\u06CC\u0635
        // \u0645\u06CC\u200C\u062F\u0647\u062F \u0686\u0648\u0646 \u062F\u06CC\u06AF\u0631 \u0628\u0627 "01" \u0634\u0631\u0648\u0639 \u0646\u0645\u06CC\u200C\u0634\u0648\u062F\u060C \u0648 \u062F\u0631 \u0641\u06CC\u0644\u062A\u0631 \u00AB\u0641\u0642\u0637 \u0627\u0642\u0644\u0627\u0645 \u062A\u06CC\u200C\u062A\u06A9\u00BB \u062C\u0627 \u0645\u06CC\u200C\u0627\u0641\u062A\u062F.
        cleaned = cleaned
            .Replace("]d2", "", StringComparison.OrdinalIgnoreCase)
            .Replace("]C1", "", StringComparison.OrdinalIgnoreCase)
            .Replace("]e0", "", StringComparison.OrdinalIgnoreCase)
            .Trim();

        return cleaned;
    }

    private static bool IsScanBridgePairPayload(string? barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return false;

        string trimmed = barcode.TrimStart();
        return barcode.IndexOf("SCANBRIDGE_PAIR", StringComparison.OrdinalIgnoreCase) >= 0
               || trimmed.StartsWith("{", StringComparison.Ordinal);
    }

    private static bool IsTtTeckLookupCandidate(string? barcode, BarcodeSource detectedSource)
    {
        if (detectedSource == BarcodeSource.TtTeck)
            return true;

        if (string.IsNullOrWhiteSpace(barcode))
            return false;

        string digitsOnly = new string(barcode.Where(char.IsDigit).ToArray());

        // UID خالص بیست‌رقمی
        if (digitsOnly.Length == 20)
            return true;

        // ساختار رایج تی‌تک: AI 01 + GTIN چهارده‌رقمی + AI 21 + UID بیست‌رقمی
        int idx01 = digitsOnly.IndexOf("01", StringComparison.Ordinal);
        while (idx01 >= 0)
        {
            if (digitsOnly.Length >= idx01 + 38 && digitsOnly.Substring(idx01 + 16, 2) == "21")
                return true;

            idx01 = digitsOnly.IndexOf("01", idx01 + 1, StringComparison.Ordinal);
        }

        // بعضی اسکنرها بخش قبل از 21 را متفاوت می‌فرستند؛ اما بعد از 21 باید حداقل UID بیست‌رقمی وجود داشته باشد.
        int idx21 = digitsOnly.IndexOf("21", StringComparison.Ordinal);
        while (idx21 >= 0)
        {
            if (digitsOnly.Length >= idx21 + 22)
                return true;

            idx21 = digitsOnly.IndexOf("21", idx21 + 1, StringComparison.Ordinal);
        }

        // سازگاری با روش قبلی/فال‌بک DrugLookupService.ExtractUIDFromBarcode: ۱۸ کاراکتر بعد
        // از "01" (بدون نیاز به وجود صریح "21")، UID بیست‌رقمی شروع می‌شود. اگر این تابع این
        // حالت را گیت نکند، بارکدی که فقط از این مسیر قابل تشخیص است، بدون هیچ اطلاعی «بارکد
        // تی‌تک نیست» برچسب می‌خورد.
        idx01 = digitsOnly.IndexOf("01", StringComparison.Ordinal);
        if (idx01 != -1 && digitsOnly.Length >= idx01 + 38)
            return true;

        // آخرین راه‌حل DrugLookupService: اگر فقط یک توالی ۲۰ رقمی در کل بارکد وجود داشته باشد،
        // همان به‌عنوان UID در نظر گرفته می‌شود.
        var twentyDigitCandidates = System.Text.RegularExpressions.Regex.Matches(digitsOnly, @"\d{20}")
            .Cast<System.Text.RegularExpressions.Match>()
            .Select(m => m.Value)
            .Distinct()
            .ToList();
        if (twentyDigitCandidates.Count == 1)
            return true;

        return false;
    }

    private string GetTtTeckLookupSkippedText()
    {
        return _localization.GetString("NotATtTeckBarcode");
    }

    // ---------- Cargo Delivery Panel ----------

    private async void CargoDeliveryPanelButton_Click(object sender, RoutedEventArgs e)
    {
        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () =>
            {
                OpenCargoDeliveryPanelNow();
                await Task.CompletedTask;
            };
            _pendingTtacRetryLabel = _localization.GetString("PendingOpenCargoDelivery");
            ShowTtacLoginOverlay();
            return;
        }

        OpenCargoDeliveryPanelNow();
    }

    private void OpenCargoDeliveryPanelNow()
    {
        UpdateCargoDeliveryLocalizedTexts();
        CargoDeliveryOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        System.Windows.Input.Keyboard.ClearFocus();
    }

    private async void ExpiryWatchPanelButton_Click(object sender, RoutedEventArgs e)
    {
        // مثل پنل «تحویل بار»، ورود به تی‌تک برای دیدن «تاریخ نزدیک» هم الزامی است - چون خودِ
        // اقلام این لیست هم فقط از همان مسیر (اسکن‌های تحویل بار) ثبت می‌شوند.
        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () =>
            {
                OpenExpiryWatchPanel();
                await Task.CompletedTask;
            };
            _pendingTtacRetryLabel = _localization.GetString("PendingOpenExpiryWatch");
            ShowTtacLoginOverlay();
            return;
        }

        OpenExpiryWatchPanel();
    }

    private void OpenExpiryWatchPanel()
    {
        UpdateExpiryAlertSettingsLocalizedTexts();
        // هر بار که پنل باز می‌شود، فیلتر دستی خاموش می‌شود تا طبق طراحی اولیه، پیش‌فرض فقط
        // اقلام پین‌شده (رسیده به آستانه‌ی تنظیمات) دیده شوند - نه آخرین فیلتری که قبلاً اعمال شده.
        _expiryFilterActive = false;
        ResetExpiryFilterCombosToDefault();
        RefreshExpiryWatchDisplayList();
        ExpiryWatchOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        System.Windows.Input.Keyboard.ClearFocus();
    }

    private void ExpiryWatchOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseExpiryWatchPanel();
    }

    private void ExpiryWatchCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseExpiryWatchPanelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseExpiryWatchPanel();
    }

    private void CloseExpiryWatchPanel()
    {
        ExpiryWatchOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void ExpiryFilterApplyButton_Click(object sender, RoutedEventArgs e)
    {
        _expiryFilterActive = true;
        RefreshExpiryWatchDisplayList();
    }

    private void ExpiryFilterClearButton_Click(object sender, RoutedEventArgs e)
    {
        _expiryFilterActive = false;
        ResetExpiryFilterCombosToDefault();
        RefreshExpiryWatchDisplayList();
    }

    private void ResetExpiryFilterCombosToDefault()
    {
        if (ExpiryFilterFromCombo == null || ExpiryFilterToCombo == null)
            return;
        SelectComboItemByTag(ExpiryFilterFromCombo, 0);
        SelectComboItemByTag(ExpiryFilterToCombo, Math.Max(1, _expiryAlertSettings.ThresholdMonths));
    }

    // بازه‌های آماده برای فیلتر «تاریخ نزدیک» - بر حسب ماه از امروز.
    private static readonly int[] ExpiryFilterFromMonthOptions = { 0, 1, 2, 3, 6, 9, 12 };
    private static readonly int[] ExpiryFilterToMonthOptions = { 1, 2, 3, 6, 9, 12, 18, 24 };
    private const int ExpiryFilterUnlimitedMonths = 999;

    private void InitializeExpiryFilterControls()
    {
        PopulateExpiryFilterCombos();
    }

    private string GetExpiryFilterMonthLabel(int months)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        if (months >= ExpiryFilterUnlimitedMonths)
            return _localization.GetString("Unlimited");
        if (months == 0)
            return _localization.GetString("Today");
        if (months % 12 == 0)
        {
            int years = months / 12;
            if (english)
                return years == 1 ? "1 year" : $"{years} years";
            return years == 1 ? "1 سال" : $"{years} سال";
        }
        return _localization.GetFormattedString("MonthsCount", months);
    }

    private void PopulateExpiryFilterCombos()
    {
        if (ExpiryFilterFromCombo == null || ExpiryFilterToCombo == null)
            return;

        int? previousFrom = GetSelectedComboTag(ExpiryFilterFromCombo, int.MinValue) is int pf && pf != int.MinValue ? pf : null;
        int? previousTo = GetSelectedComboTag(ExpiryFilterToCombo, int.MinValue) is int pt && pt != int.MinValue ? pt : null;

        ExpiryFilterFromCombo.Items.Clear();
        foreach (int m in ExpiryFilterFromMonthOptions)
            ExpiryFilterFromCombo.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = GetExpiryFilterMonthLabel(m), Tag = m });

        ExpiryFilterToCombo.Items.Clear();
        foreach (int m in ExpiryFilterToMonthOptions)
            ExpiryFilterToCombo.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = GetExpiryFilterMonthLabel(m), Tag = m });
        ExpiryFilterToCombo.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = GetExpiryFilterMonthLabel(ExpiryFilterUnlimitedMonths), Tag = ExpiryFilterUnlimitedMonths });

        int fromToSelect = previousFrom ?? 0;
        int toToSelect = previousTo ?? Math.Max(1, _expiryAlertSettings.ThresholdMonths);

        SelectComboItemByTag(ExpiryFilterFromCombo, fromToSelect);
        SelectComboItemByTag(ExpiryFilterToCombo, toToSelect);
    }

    private static void SelectComboItemByTag(System.Windows.Controls.ComboBox combo, int tag)
    {
        foreach (var obj in combo.Items)
        {
            if (obj is System.Windows.Controls.ComboBoxItem item && item.Tag is int t && t == tag)
            {
                combo.SelectedItem = item;
                return;
            }
        }
        if (combo.Items.Count > 0)
            combo.SelectedIndex = 0;
    }

    private static int GetSelectedComboTag(System.Windows.Controls.ComboBox? combo, int fallback)
    {
        if (combo?.SelectedItem is System.Windows.Controls.ComboBoxItem item && item.Tag is int tag)
            return tag;
        return fallback;
    }

    private void RefreshExpiryWatchDisplayList()
    {
        try
        {
            bool english = _localization.CurrentLanguage == AppLanguage.English;
            string pharmacyKey = GetReceiveStatusStorageKey();
            var store = LoadExpiryWatchStore();
            store.TryGetValue(pharmacyKey, out var items);
            items ??= new List<ExpiryWatchItem>();

            var watchingAll = items.Where(x => x.Status == ExpiryWatchStatus.Watching).ToList();

            DateTime today = DateTime.Now.Date;
            DateTime alertThresholdDate = today.AddMonths(Math.Max(1, _expiryAlertSettings.ThresholdMonths));
            // «due» یعنی هم به آستانه‌ی هشدار رسیده هم هنوز جواب داده نشده (فروخته شد/حواسم هست
            // زده نشده) - همان اقلامی که نشان قرمز روی دکمه‌ی «تاریخ نزدیک» را فعال نگه می‌دارند.
            // عمداً به NextAlertDueUtc گره نخورده، چون آن فقط زمان‌بندی چرخه‌ی خودکار بعدی است و
            // با هر بار فایر شدن هشدار (چه کاربر پاسخ بدهد چه ندهد) جلو کشیده می‌شود.
            int dueCount = watchingAll.Count(x => x.ExpirationDate <= alertThresholdDate && x.NeedsResponse);

            // اقلام «پین» - این‌ها همیشه بالای لیست نشان داده می‌شوند، کاملاً مستقل از فیلتر بازه‌ی
            // دستی کاربر: همان اقلامی که به آستانه‌ی هشدار تنظیمات رسیده‌اند و هنوز پاسخ داده
            // نشده‌اند. نمایش پیش‌فرض (بدون فیلتر) دقیقاً همین لیست است - نه خالی، نه همه‌چیز.
            var pinnedItems = watchingAll
                .Where(x => x.ExpirationDate <= alertThresholdDate && x.NeedsResponse)
                .OrderBy(x => x.ExpirationDate)
                .ToList();

            // اقلام «اضافه» - فقط وقتی کاربر صریحاً فیلتر را با دکمه‌ی «اعمال فیلتر» فعال کرده
            // باشد محاسبه می‌شوند و زیر اقلام پین‌شده می‌آیند؛ اقلامی که از قبل پین شده‌اند دوباره
            // در این بخش تکرار نمی‌شوند.
            var otherItems = new List<ExpiryWatchItem>();
            if (_expiryFilterActive)
            {
                int fromMonths = GetSelectedComboTag(ExpiryFilterFromCombo, 0);
                int toMonthsRaw = GetSelectedComboTag(ExpiryFilterToCombo, Math.Max(1, _expiryAlertSettings.ThresholdMonths));
                DateTime fromDate = today.AddMonths(fromMonths);
                DateTime? toDate = toMonthsRaw >= ExpiryFilterUnlimitedMonths ? (DateTime?)null : today.AddMonths(toMonthsRaw);

                var pinnedBarcodes = new HashSet<string>(pinnedItems.Select(x => x.Barcode), StringComparer.OrdinalIgnoreCase);
                otherItems = watchingAll
                    .Where(x => !pinnedBarcodes.Contains(x.Barcode))
                    .Where(x => x.ExpirationDate >= fromDate && (toDate == null || x.ExpirationDate <= toDate.Value))
                    .OrderBy(x => x.ExpirationDate)
                    .ToList();
            }

            var filtered = pinnedItems.Concat(otherItems).ToList();

            ExpiryWatchDisplayItems.Clear();
            int index = 1;
            foreach (var item in filtered)
            {
                string name = string.IsNullOrWhiteSpace(item.ProductName)
                    ? (string.IsNullOrWhiteSpace(item.ProductEnName) ? item.Barcode : item.ProductEnName)
                    : item.ProductName;

                bool isPinned = item.ExpirationDate <= alertThresholdDate && item.NeedsResponse;
                if (isPinned)
                    name = "📌 " + name;

                int daysLeft = (int)Math.Ceiling((item.ExpirationDate - today).TotalDays);
                string statusText = daysLeft < 0
                    ? (_localization.GetString("ExpirationDateHasPassed"))
                    : (_localization.GetFormattedString("DaysLeft", daysLeft));
                var statusBrush = daysLeft <= 30
                    ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xB9, 0x1C, 0x1C))
                    : new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x92, 0x40, 0x0E));

                string batchCode = string.IsNullOrWhiteSpace(item.BatchCode) ? "-" : item.BatchCode;
                string detailText = _localization.GetFormattedString("LotExpiry", batchCode, item.PersianExpirationText);

                var cardBackground = isPinned
                    ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFE, 0xF2, 0xF2))
                    : new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFB, 0xEB));
                var cardBorder = isPinned
                    ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFC, 0xA5, 0xA5))
                    : new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFD, 0xE6, 0x8A));

                ExpiryWatchDisplayItems.Add(new ExpiryWatchDisplayRow
                {
                    RowNumber = index++,
                    Barcode = item.Barcode,
                    ProductName = name,
                    BatchCode = batchCode,
                    ExpirationText = item.PersianExpirationText,
                    DetailText = detailText,
                    StatusText = statusText,
                    StatusBrush = statusBrush,
                    IsPinned = isPinned,
                    SoldButtonText = _localization.GetString("Sold"),
                    AckButtonText = _localization.GetString("GotIt"),
                    CardBackground = cardBackground,
                    CardBorder = cardBorder
                });
            }

            bool ttacPlusAllowed = IsLicenseValid() && _activeLicense.Plan == "TtacPlus";

            if (ExpiryWatchPanelButtonHost != null)
                ExpiryWatchPanelButtonHost.Visibility = ttacPlusAllowed ? Visibility.Visible : Visibility.Collapsed;

            if (ExpiryWatchPanelButton != null)
            {
                string baseLabel = _localization.GetString("UpcomingExpiry");
                ExpiryWatchPanelButton.Content = dueCount > 0 ? $"{baseLabel} ({dueCount})" : baseLabel;
            }

            if (CargoDeliveryAlertBadge != null)
                CargoDeliveryAlertBadge.Visibility = (ttacPlusAllowed && dueCount > 0) ? Visibility.Visible : Visibility.Collapsed;

            if (ExpiryWatchTitle != null)
                ExpiryWatchTitle.Text = _localization.GetString("UpcomingExpiry");
        }
        catch { }
    }

    private void ExpiryWatchExportExcelButton_Click(object sender, RoutedEventArgs e)
    {
        var rows = ExpiryWatchDisplayItems.ToList();
        if (rows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        var saveFileDialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"Scanbridge_ExpiryWatch_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = _localization.GetString("SaveNearExpiryReport")
        };
        if (saveFileDialog.ShowDialog() != true)
            return;

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("ExpiryWatch");
            string[] headers = _localization.GetStringArray("ExpiryWatchHeaders");
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(245, 127, 23);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            int r = 2;
            foreach (var row in rows)
            {
                worksheet.Cell(r, 1).Value = row.RowNumber;
                worksheet.Cell(r, 2).Value = row.ProductName;
                worksheet.Cell(r, 3).Value = row.Barcode;
                worksheet.Cell(r, 4).Value = row.BatchCode;
                worksheet.Cell(r, 5).Value = row.ExpirationText;
                worksheet.Cell(r, 6).Value = row.StatusText;
                r++;
            }
            worksheet.Columns().AdjustToContents();
            worksheet.RangeUsed()?.CreateTable();
            workbook.SaveAs(saveFileDialog.FileName);
            ShowStyledMessage(_localization.GetString("ExportSuccessful"), saveFileDialog.FileName);
        }
        catch (Exception ex)
        {
            ShowStyledMessage(_localization.GetString("ExportFailed"), ex.Message, true);
        }
    }

    private void ExpiryWatchSoldButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button btn && btn.Tag is string barcode && !string.IsNullOrWhiteSpace(barcode))
            MarkExpiryItemSold(barcode);
    }

    private void ExpiryWatchAckButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button btn && btn.Tag is string barcode && !string.IsNullOrWhiteSpace(barcode))
            AcknowledgeExpiryItem(barcode);
    }

    private void UpdateCargoDeliveryLocalizedTexts()
    {
        if (CargoDeliveryTitle == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        CargoDeliveryTitle.Text = _localization.GetString("CargoDelivery");
        CargoDeliveryManualLabel.Text = _localization.GetString("ManualUIDBarcode");
        CargoDeliveryManualTextBox.ToolTip = _localization.GetString("InCargoModeScansAreAddedOnlyHere");
        CargoDeliveryManualAddButton.Content = _localization.GetString("AddManually");
        CargoDeliverySelectAllCheckBox.Content = _localization.GetString("SelectAll");
        CargoDeliveryAddSelectedButton.Content = _localization.GetString("AddSelected");
        CargoDeliveryAddAllButton.Content = _localization.GetString("AddAll");
        CargoDeliveryClearAllButton.Content = _localization.GetString("ClearAll");
        CargoDeliveryExportExcelButton.Content = "Excel";
        CargoDeliveryExportPdfButton.Content = "PDF";
        if (string.IsNullOrWhiteSpace(CargoDeliveryBulkStatusText.Text))
            CargoDeliveryBulkStatusText.Text = string.Empty;

        // اقلامی که از قبل در لیست تحویل بار هستند خودشان زبان جدید را نمی‌گیرند چون این کلاس
        // اعلان تغییر (INotifyPropertyChanged) ندارد؛ برای همین پرچم زبان‌شان را دستی به‌روز
        // می‌کنیم و با Refresh لیست را دوباره رسم می‌کنیم.
        bool anyLanguageChanged = false;
        foreach (var row in CargoDeliveryItems)
        {
            if (row.English != english)
            {
                row.English = english;
                anyLanguageChanged = true;
            }
        }
        if (anyLanguageChanged)
            CargoDeliveryListBox?.Items.Refresh();
    }

    private void RefreshCargoDeliveryRowNumbers()
    {
        int index = 1;
        foreach (var row in CargoDeliveryItems)
            row.RowNumber = index++;
        CargoDeliveryListBox?.Items.Refresh();
    }

    private void RefreshReceiveStatusRowNumbers()
    {
        int index = 1;
        foreach (var row in ReceiveStatusItems)
            row.RowNumber = index++;
        ReceiveStatusListBox?.Items.Refresh();
    }

    private void CargoDeliveryOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseCargoDeliveryPanel();
    }

    private void CargoDeliveryCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseCargoDeliveryPanelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseCargoDeliveryPanel();
    }

    private void CloseCargoDeliveryPanel()
    {
        CargoDeliveryOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void CargoDeliveryManualTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;

        e.Handled = true;
        _ = AddCargoDeliveryBarcodeAsync(CargoDeliveryManualTextBox.Text, showErrors: true);
    }

    private void CargoDeliveryManualAddButton_Click(object sender, RoutedEventArgs e)
    {
        _ = AddCargoDeliveryBarcodeAsync(CargoDeliveryManualTextBox.Text, showErrors: true);
    }

    private async Task AddCargoDeliveryBarcodeAsync(string? rawBarcode, bool showErrors)
    {
        string barcode = CleanBarcodeForExternalUse(rawBarcode);
        if (string.IsNullOrWhiteSpace(barcode))
            return;

        if (!Dispatcher.CheckAccess())
        {
            var uiTask = await Dispatcher.InvokeAsync(() => AddCargoDeliveryBarcodeAsync(barcode, showErrors));
            await uiTask;
            return;
        }

        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            if (showErrors)
            {
                _pendingTtacRetryAction = async () => await AddCargoDeliveryBarcodeAsync(barcode, showErrors: true);
                _pendingTtacRetryLabel = _localization.GetString("PendingAddCargoBarcode");
                ShowTtacLoginOverlay();
            }
            return;
        }

        if (_cargoDeliveryKnownBarcodes.Contains(barcode))
            return;

        _cargoDeliveryKnownBarcodes.Add(barcode);
        CargoDeliveryManualAddButton.IsEnabled = false;
        string oldContent = CargoDeliveryManualAddButton.Content?.ToString() ?? string.Empty;
        CargoDeliveryManualAddButton.Content = _localization.GetString("Checking");

        try
        {
            var receiveRow = await BuildReceiveStatusRowAsync(barcode);
            if (receiveRow == null)
                return;

            var cargoRow = CargoDeliveryRow.FromReceiveStatus(receiveRow, _localization.CurrentLanguage == AppLanguage.English);
            CargoDeliveryItems.Insert(0, cargoRow);
            RefreshCargoDeliveryRowNumbers();
            SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
            CargoDeliveryManualTextBox.Text = string.Empty;

            // اگر این قلم تاریخ انقضای قابل‌فهمی داشت، برای پایش «تاریخ نزدیک» ثبت می‌شود.
            // این ثبت مستقل از لیست تحویل بار است (که ممکن است در آرشیو ماهانه پاک شود)
            // تا هشدار انقضا تا ۶ ماه بعد هم زنده بماند.
            TryRegisterExpiryWatchItem(cargoRow);
            RefreshExpiryWatchDisplayList();
            _ = CheckExpiryAlertsAsync();
        }
        catch (Exception ex)
        {
            _cargoDeliveryKnownBarcodes.Remove(barcode);
            if (showErrors)
                ShowStyledMessage(_localization.GetString("CargoDelivery"), ex.Message, true);
        }
        finally
        {
            CargoDeliveryManualAddButton.IsEnabled = true;
            CargoDeliveryManualAddButton.Content = oldContent;
        }
    }

    private async Task ConfirmCargoDeliveryRowAsync(CargoDeliveryRow row)
    {
        if (row == null || !row.IsActionEnabled || row.ReceiveId <= 0)
            return;

        try
        {
            row.IsActionEnabled = false;
            row.StatusText = _localization.GetString("AddingToPharmacy");
            CargoDeliveryListBox.Items.Refresh();

            using var doc = await SendTtacJsonAsync(HttpMethod.Post, "https://statisticsreports.ttac.ir/declaration/ConfirmReceive", new { Id = row.ReceiveId });
            string? message = doc == null ? null : ReadJsonString(doc.RootElement, "Message");
            row.MarkAdded(_localization.CurrentLanguage == AppLanguage.English, string.IsNullOrWhiteSpace(message) || message == "null" ? null : message);
            CargoDeliveryListBox.Items.Refresh();
            SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
        }
        catch (Exception ex)
        {
            // فقط وقتی توکن واقعاً معتبر نیست، این خطا «انقضای نشست» محسوب می‌شود؛ وگرنه ممکن
            // است پیام خطا فقط عدد ۴۰۱/۴۰۳ یا Unauthorized را داخل خودش داشته باشد (مثلاً یک
            // پیام خطای عادی سامانه) و نباید پنجره‌ی ورود را با هشدار «نشست منقضی شده» باز کند.
            if (IsTtacSessionExpiredException(ex) && !HasValidTtacToken())
            {
                row.IsActionEnabled = true;
                throw;
            }

            row.IsActionEnabled = true;
            row.StatusText = ex.Message;
            row.StatusBrush = System.Windows.Media.Brushes.Firebrick;
            CargoDeliveryListBox.Items.Refresh();
            SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
        }
    }

    private async void CargoDeliveryConfirmButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is CargoDeliveryRow row)
        {
            button.IsEnabled = false;
            try
            {
                await ConfirmCargoDeliveryRowAsync(row);
            }
            catch (Exception ex)
            {
                HandleTtacOperationException(ex,
                    _localization.GetString("CargoDeliveryFailed"),
                    async () => await ConfirmCargoDeliveryRowAsync(row),
                    pendingLabel: _localization.GetString("PendingConfirmCargoRow"));
            }
            finally
            {
                button.IsEnabled = row.IsActionEnabled;
            }
        }
    }

    private async void CargoDeliveryAddSelectedButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = CargoDeliveryItems.Where(x => x.IsSelected && x.IsActionEnabled).ToList();
        await ProcessCargoDeliveryRowsAsync(selected);
    }

    private async void CargoDeliveryAddAllButton_Click(object sender, RoutedEventArgs e)
    {
        var rows = CargoDeliveryItems.Where(x => x.IsActionEnabled).ToList();
        await ProcessCargoDeliveryRowsAsync(rows);
    }

    private void SetCargoDeliveryBulkControlsEnabled(bool enabled)
    {
        CargoDeliveryAddSelectedButton.IsEnabled = enabled;
        CargoDeliveryAddAllButton.IsEnabled = enabled;
        CargoDeliveryClearAllButton.IsEnabled = enabled;
        CargoDeliveryManualAddButton.IsEnabled = enabled;
        CargoDeliveryExportExcelButton.IsEnabled = enabled;
        CargoDeliveryExportPdfButton.IsEnabled = enabled;
        CargoDeliverySelectAllCheckBox.IsEnabled = enabled;
    }

    private void SetCargoDeliveryBulkStatus(string text)
    {
        CargoDeliveryBulkStatusText.Text = text;
    }

    private async Task ProcessCargoDeliveryRowsAsync(List<CargoDeliveryRow> rows)
    {
        if (_isCargoDeliveryBulkRunning)
            return;

        rows = rows.Where(x => x.IsActionEnabled).ToList();
        if (rows.Count == 0)
        {
            ShowStyledMessage(
                _localization.GetString("CargoDelivery"),
                _localization.GetString("NoItemIsReadyToAdd"),
                true);
            return;
        }

        _isCargoDeliveryBulkRunning = true;
        SetCargoDeliveryBulkControlsEnabled(false);
        int success = 0;
        int failed = 0;

        try
        {
            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                if (!row.IsActionEnabled)
                    continue;

                SetCargoDeliveryBulkStatus(_localization.GetFormattedString("AddingCount", i + 1, rows.Count));

                try
                {
                    await ConfirmCargoDeliveryRowAsync(row);
                    if (!row.IsActionEnabled && row.StatusBrush == System.Windows.Media.Brushes.Green)
                        success++;
                    else
                        failed++;
                }
                catch (Exception ex) when (IsTtacSessionExpiredException(ex) && !HasValidTtacToken())
                {
                    var remaining = rows.Skip(i).Where(x => x.IsActionEnabled).ToList();
                    _pendingTtacRetryAction = async () => await ProcessCargoDeliveryRowsAsync(remaining);
                    _pendingTtacRetryLabel = _localization.GetString("PendingContinueCargoDelivery");
                    ShowTtacLoginOverlay(sessionExpired: true);
                    TtacLoginStatusText.Text = _localization.GetString("TTACSessionExpiredLoginAgainCargoDeliveryWillContinueAutomatically");
                    return;
                }
                catch
                {
                    failed++;
                }
            }

            SetCargoDeliveryBulkStatus(string.Empty);
            ShowStyledMessage(
                _localization.GetString("CargoDeliveryCompleted"),
                _localization.GetFormattedString("SuccessFailed", success, failed));
        }
        finally
        {
            _isCargoDeliveryBulkRunning = false;
            SetCargoDeliveryBulkControlsEnabled(true);
            if (CargoDeliveryOverlay.Visibility == Visibility.Visible && !TtacLoginOverlay.IsVisible)
            {
                if (!CargoDeliveryBulkStatusText.Text.Contains("منقضی", StringComparison.OrdinalIgnoreCase))
                    SetCargoDeliveryBulkStatus(string.Empty);
            }
        }
    }

    private void CargoDeliverySelectAllCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        bool isChecked = CargoDeliverySelectAllCheckBox.IsChecked == true;
        foreach (var row in CargoDeliveryItems)
            row.IsSelected = isChecked;
        CargoDeliveryListBox.Items.Refresh();
    }

    private void DeleteCargoDeliveryButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is CargoDeliveryRow row)
        {
            CargoDeliveryItems.Remove(row);
            if (!string.IsNullOrWhiteSpace(row.Barcode))
                _cargoDeliveryKnownBarcodes.Remove(row.Barcode);
            RefreshCargoDeliveryRowNumbers();
            SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
        }
    }

    private void CargoDeliveryClearAllButton_Click(object sender, RoutedEventArgs e)
    {
        CargoDeliveryItems.Clear();
        _cargoDeliveryKnownBarcodes.Clear();
        CargoDeliverySelectAllCheckBox.IsChecked = false;
        RefreshCargoDeliveryRowNumbers();
        SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
    }

    private void CargoDeliveryListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (CargoDeliveryListBox.SelectedItem is CargoDeliveryRow row)
            ShowCargoDeliveryRowDetails(row);
    }

    private void ShowCargoDeliveryRowDetails(CargoDeliveryRow row)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        var left = new List<ProductDetailField>
        {
            DetailField(_localization.GetString("Product"), row.ProductName),
            DetailField(_localization.GetString("EnglishProduct"), row.ProductEnName),
            DetailField(_localization.GetString("BarcodeUID"), row.Barcode),
            DetailField("IRC", row.Irc),
            DetailField("UID", row.UID),
            DetailField("GTIN", row.GTIN),
            DetailField(_localization.GetString("LotBatch"), row.LotNumber)
        };

        var right = new List<ProductDetailField>
        {
            DetailField(_localization.GetString("Expiration"), row.Expiration),
            DetailField(_localization.GetString("GenericCode"), row.GenericCode),
            DetailField(_localization.GetString("GenericName"), row.GenericName),
            DetailField(_localization.GetString("SenderDistributor"), row.SenderName),
            DetailField(_localization.GetString("Quantity"), row.Quantity),
            DetailField(_localization.GetString("SentDate"), row.SentDatePersian),
            DetailField(_localization.GetString("Status"), row.StatusText),
            DetailField(_localization.GetString("ReceiveItemID"), row.ReceiveId > 0 ? row.ReceiveId.ToString(CultureInfo.InvariantCulture) : "-")
        };

        ShowOperationDetails(_localization.GetString("CargoItemDetails"), left, right);
    }

    private ProductDetailField DetailField(string label, string? value)
    {
        return new ProductDetailField
        {
            Label = label,
            Value = string.IsNullOrWhiteSpace(value) ? "-" : value
        };
    }

    private void ShowOperationDetails(string title, IEnumerable<ProductDetailField> leftFields, IEnumerable<ProductDetailField> rightFields)
    {
        OperationDetailsTitle.Text = title;
        OperationDetailsOkButton.Content = _localization.GetString("Close");
        OperationDetailsLeftFields.Clear();
        OperationDetailsRightFields.Clear();
        foreach (var field in leftFields)
            OperationDetailsLeftFields.Add(field);
        foreach (var field in rightFields)
            OperationDetailsRightFields.Add(field);
        System.Windows.Controls.Panel.SetZIndex(OperationDetailsOverlay, 280);
        OperationDetailsOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void CloseOperationDetails()
    {
        OperationDetailsOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = CargoDeliveryOverlay.Visibility == Visibility.Visible
            || ReceiveStatusOverlay.Visibility == Visibility.Visible
            || TtacPanelOverlay.Visibility == Visibility.Visible
            || HistoryOverlay.Visibility == Visibility.Visible
            ? new System.Windows.Media.Effects.BlurEffect { Radius = 18 }
            : null;
    }

    private void OperationDetailsOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseOperationDetails();
    }

    private void OperationDetailsCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void OperationDetailsCloseButton_Click(object sender, RoutedEventArgs e)
    {
        CloseOperationDetails();
    }

    private void ReceiveStatusListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (ReceiveStatusListBox.SelectedItem is ReceiveStatusRow row)
            ShowReceiveStatusRowDetails(row);
    }

    private void ShowReceiveStatusRowDetails(ReceiveStatusRow row)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        var left = new List<ProductDetailField>
        {
            DetailField(_localization.GetString("Product"), row.ProductName),
            DetailField(_localization.GetString("EnglishProduct"), row.ProductEnName),
            DetailField(_localization.GetString("BarcodeUID"), row.Barcode),
            DetailField("IRC", row.Irc),
            DetailField("UID", row.UID),
            DetailField("GTIN", row.GTIN),
            DetailField(_localization.GetString("LotBatch"), row.LotNumber)
        };
        var right = new List<ProductDetailField>
        {
            DetailField(_localization.GetString("Expiration"), row.Expiration),
            DetailField(_localization.GetString("GenericCode"), row.GenericCode),
            DetailField(_localization.GetString("GenericName"), row.GenericName),
            DetailField(_localization.GetString("SenderDistributor"), row.SenderName),
            DetailField(_localization.GetString("Quantity"), row.Quantity),
            DetailField(_localization.GetString("SentDate"), row.SentDatePersian),
            DetailField(_localization.GetString("Status"), row.StatusText),
            DetailField(_localization.GetString("ReceiveItemID"), row.ReceiveId > 0 ? row.ReceiveId.ToString(CultureInfo.InvariantCulture) : "-")
        };
        ShowOperationDetails(_localization.GetString("ReceiveStatusDetails"), left, right);
    }

    private string[] GetCargoDeliveryExportHeaders()
    {
        return _localization.GetStringArray("CargoDeliveryExportHeaders");
    }

    private void CargoDeliveryExportExcelButton_Click(object sender, RoutedEventArgs e)
    {
        var rows = CargoDeliveryItems.ToList();
        if (rows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        var saveFileDialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"Scanbridge_CargoDelivery_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = _localization.GetString("SaveCargoDeliveryReport")
        };
        if (saveFileDialog.ShowDialog() != true)
            return;

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("CargoDelivery");
            string[] headers = GetCargoDeliveryExportHeaders();
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(16, 185, 129);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            int r = 2;
            int index = 1;
            foreach (var row in rows)
            {
                worksheet.Cell(r, 1).Value = index++;
                worksheet.Cell(r, 2).Value = row.IsSelected ? "✓" : "";
                worksheet.Cell(r, 3).Value = row.ProductName;
                worksheet.Cell(r, 4).Value = row.ProductEnName;
                worksheet.Cell(r, 5).Value = row.Barcode;
                worksheet.Cell(r, 6).Value = row.Irc;
                worksheet.Cell(r, 7).Value = row.LotNumber;
                worksheet.Cell(r, 8).Value = row.Quantity;
                worksheet.Cell(r, 9).Value = row.SenderName;
                worksheet.Cell(r, 10).Value = row.StatusText;
                r++;
            }
            worksheet.Columns().AdjustToContents();
            worksheet.RangeUsed()?.CreateTable();
            workbook.SaveAs(saveFileDialog.FileName);
            ShowStyledMessage(_localization.GetString("ExportSuccessful"), saveFileDialog.FileName);
        }
        catch (Exception ex)
        {
            ShowStyledMessage(_localization.GetString("ExportFailed"), ex.Message, true);
        }
    }

    private void CargoDeliveryExportPdfButton_Click(object sender, RoutedEventArgs e)
    {
        var rows = CargoDeliveryItems.ToList();
        if (rows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }
        var printDialog = new System.Windows.Controls.PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var document = new FlowDocument
        {
            FlowDirection = _localization.CurrentLanguage == AppLanguage.English ? System.Windows.FlowDirection.LeftToRight : System.Windows.FlowDirection.RightToLeft,
            FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
            FontSize = 9,
            PageWidth = printDialog.PrintableAreaWidth,
            PageHeight = printDialog.PrintableAreaHeight,
            PagePadding = new Thickness(28),
            ColumnWidth = printDialog.PrintableAreaWidth
        };
        document.Blocks.Add(new Paragraph(new Run(_localization.GetString("ScanbridgeCargoDelivery")))
        {
            FontSize = 22,
            FontWeight = FontWeights.Bold,
            TextAlignment = TextAlignment.Center,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x06, 0x5F, 0x46)),
            Margin = new Thickness(0, 0, 0, 14)
        });
        var table = new Table { CellSpacing = 0 };
        document.Blocks.Add(table);
        string[] headers = GetCargoDeliveryExportHeaders();
        foreach (var _ in headers) table.Columns.Add(new TableColumn());
        var group = new TableRowGroup();
        table.RowGroups.Add(group);
        var headerRow = new TableRow();
        group.Rows.Add(headerRow);
        foreach (string h in headers) headerRow.Cells.Add(CreatePdfCell(h, true));
        int idx = 1;
        foreach (var row in rows)
        {
            var tr = new TableRow();
            group.Rows.Add(tr);
            tr.Cells.Add(CreatePdfCell((idx++).ToString(), false));
            tr.Cells.Add(CreatePdfCell(row.IsSelected ? "✓" : "", false));
            tr.Cells.Add(CreatePdfCell(row.ProductName, false));
            tr.Cells.Add(CreatePdfCell(row.ProductEnName, false));
            tr.Cells.Add(CreatePdfCell(row.Barcode, false));
            tr.Cells.Add(CreatePdfCell(row.Irc, false));
            tr.Cells.Add(CreatePdfCell(row.LotNumber, false));
            tr.Cells.Add(CreatePdfCell(row.Quantity, false));
            tr.Cells.Add(CreatePdfCell(row.SenderName, false));
            tr.Cells.Add(CreatePdfCell(row.StatusText, false));
        }
        printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, "Scanbridge Cargo Delivery");
        ShowStyledMessage(_localization.GetString("PDFReport"), GetLocalizedPdfSuccessPathText());
    }

    // ---------- Receive Status Panel ----------

    private void QueueReceiveStatusBarcode(string? barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return;

        if (_receiveStatusKnownBarcodes.Contains(barcode))
            return;

        if (!_queuedReceiveStatusBarcodes.Any(x => x.Equals(barcode, StringComparison.OrdinalIgnoreCase)))
            _queuedReceiveStatusBarcodes.Add(barcode);
    }

    private async Task ProcessQueuedReceiveStatusBarcodesAsync()
    {
        if (!HasValidTtacToken())
            return;

        var items = _queuedReceiveStatusBarcodes.ToList();
        foreach (string barcode in items)
        {
            await AddReceiveStatusBarcodeAsync(barcode, showErrors: false);
            _queuedReceiveStatusBarcodes.RemoveAll(x => x.Equals(barcode, StringComparison.OrdinalIgnoreCase));
        }
    }

    private async void ReceiveStatusPanelButton_Click(object sender, RoutedEventArgs e)
    {
        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () =>
            {
                OpenReceiveStatusPanelNow();
                await Task.CompletedTask;
            };
            _pendingTtacRetryLabel = _localization.GetString("PendingOpenReceiveStatus");
            ShowTtacLoginOverlay();
            return;
        }

        OpenReceiveStatusPanelNow();
    }

    private void OpenReceiveStatusPanelNow()
    {
        UpdateReceiveStatusLocalizedTexts();
        ReceiveStatusOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        ReceiveStatusManualTextBox.Focus();
        _ = ProcessQueuedReceiveStatusBarcodesAsync();
    }

    private void UpdateReceiveStatusLocalizedTexts()
    {
        if (ReceiveStatusTitle == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        ReceiveStatusTitle.Text = _localization.GetString("ReceiveStatus");
        ReceiveStatusManualLabel.Text = _localization.GetString("ManualUIDBarcode");
        ReceiveStatusManualTextBox.ToolTip = _localization.GetString("EnterUIDOrFullBarcodeToReceiveConfirm");
        ReceiveStatusManualAddButton.Content = _localization.GetString("AddManually");
        ReceiveStatusSearchLabel.Text = _localization.GetString("Filter");
        ReceiveStatusSearchTextBox.ToolTip = _localization.GetString("SearchByProductBarcodeIRCLotDistributorDateOrStatus");
        ReceiveStatusClearSearchButton.Content = _localization.GetString("Clear");
        ReceiveStatusExportExcelButton.Content = "Excel";
        ReceiveStatusExportPdfButton.Content = "PDF";
    }

    private void ReceiveStatusOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseReceiveStatusPanel();
    }

    private void ReceiveStatusCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseReceiveStatusPanelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseReceiveStatusPanel();
    }

    private void CloseReceiveStatusPanel()
    {
        ReceiveStatusOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void ReceiveStatusManualTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;

        e.Handled = true;
        _ = AddManualReceiveStatusAsync();
    }

    private void ReceiveStatusManualAddButton_Click(object sender, RoutedEventArgs e)
    {
        _ = AddManualReceiveStatusAsync();
    }

    private async Task AddManualReceiveStatusAsync()
    {
        string barcode = CleanBarcodeForExternalUse(ReceiveStatusManualTextBox.Text);
        if (string.IsNullOrWhiteSpace(barcode))
        {
            ShowStyledMessage(_localization.GetString("ManualEntry"),
                _localization.GetString("EnterTheUIDOrFullBarcodeFirst"), true);
            return;
        }

        await AddReceiveStatusBarcodeAsync(barcode, showErrors: true);
    }

    private async Task AddReceiveStatusBarcodeAsync(string barcode, bool showErrors)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return;

        if (!Dispatcher.CheckAccess())
        {
            var uiTask = await Dispatcher.InvokeAsync(() => AddReceiveStatusBarcodeAsync(barcode, showErrors));
            await uiTask;
            return;
        }

        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            QueueReceiveStatusBarcode(barcode);
            if (showErrors)
            {
                _pendingTtacRetryAction = async () => await AddReceiveStatusBarcodeAsync(barcode, showErrors: true);
                _pendingTtacRetryLabel = _localization.GetString("PendingAddReceiveStatusBarcode");
                ShowTtacLoginOverlay();
            }
            return;
        }

        if (_receiveStatusKnownBarcodes.Contains(barcode))
        {
            if (showErrors)
                ShowStyledMessage(_localization.GetString("ReceiveStatus"),
                    _localization.GetString("ThisBarcodeIsAlreadyInTheReceiveStatusList"), true);
            return;
        }

        ReceiveStatusManualAddButton.IsEnabled = false;
        string oldContent = ReceiveStatusManualAddButton.Content?.ToString() ?? string.Empty;
        ReceiveStatusManualAddButton.Content = _localization.GetString("Checking");

        try
        {
            var row = await BuildReceiveStatusRowAsync(barcode);
            if (row == null)
                return;

            ReceiveStatusItems.Insert(0, row);
            RefreshReceiveStatusRowNumbers();
            _receiveStatusKnownBarcodes.Add(barcode);
            // این متد هم از مسیر «آنی هنگام اسکن» (وقتی توکن معتبر است) صدا زده می‌شود، هم از
            // صف ProcessQueuedReceiveStatusBarcodesAsync. قبلاً فقط آن حلقه‌ی صف، آیتم را از
            // _queuedReceiveStatusBarcodes حذف می‌کرد - یعنی مسیر آنی هرچقدر هم موفق می‌شد، صف
            // خودش خالی نمی‌شد و در طول یک روز پرمصرف بدون باز شدن پنل، فقط بزرگ‌تر می‌شد؛ وقتی
            // پنل بالاخره باز می‌شد، همه‌ی آن بارکدهای از قبل موفق دوباره (بی‌فایده) پردازش
            // می‌شدند (باگ گزارش ممیزی). حذف همین‌جا تضمین می‌کند موفقیت از هر مسیری صف را خالی کند.
            _queuedReceiveStatusBarcodes.RemoveAll(x => x.Equals(barcode, StringComparison.OrdinalIgnoreCase));
            SaveCurrentReceiveStatusItemsForCurrentPharmacy();
            ReceiveStatusManualTextBox.Text = string.Empty;
        }
        catch (Exception ex)
        {
            if (IsJsonNullElementException(ex))
            {
                if (!_receiveStatusKnownBarcodes.Contains(barcode))
                {
                    ReceiveStatusItems.Insert(0, CreateReceiveStatusNotFoundRow(
                        barcode,
                        _localization.GetString("TTACReturnedAnEmptyResultForThisReceiveStatusRequest")));
                    RefreshReceiveStatusRowNumbers();
                    _receiveStatusKnownBarcodes.Add(barcode);
                    _queuedReceiveStatusBarcodes.RemoveAll(x => x.Equals(barcode, StringComparison.OrdinalIgnoreCase));
                }
                return;
            }

            if (showErrors)
            {
                HandleTtacOperationException(ex,
                    _localization.GetString("ReceiveStatusFailed"),
                    async () => await AddReceiveStatusBarcodeAsync(barcode, showErrors: true),
                    pendingLabel: _localization.GetString("PendingAddReceiveStatusBarcode"));
            }
        }
        finally
        {
            ReceiveStatusManualAddButton.IsEnabled = true;
            ReceiveStatusManualAddButton.Content = oldContent;
            ReceiveStatusManualTextBox.Focus();
        }
    }

    private (string Gtin, string Uid, string LotNumber) ExtractTtacBarcodePartsForReceiveStatus(string barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return (string.Empty, string.Empty, string.Empty);

        string cleaned = new string(barcode.Trim().Where(c => !char.IsControl(c) && !char.IsWhiteSpace(c)).ToArray());
        cleaned = cleaned
            .Replace("]d2", "", StringComparison.OrdinalIgnoreCase)
            .Replace("]C1", "", StringComparison.OrdinalIgnoreCase);

        var match = Regex.Match(cleaned, @"01(?<gtin>\d{14})21(?<uid>\d{20})17\d{6}10(?<lot>[A-Za-z0-9\-_/]+)");
        if (match.Success)
            return (match.Groups["gtin"].Value, match.Groups["uid"].Value, match.Groups["lot"].Value);

        string digitsOnly = new string(cleaned.Where(char.IsDigit).ToArray());
        string uid = string.Empty;
        var uidMatch = Regex.Match(digitsOnly, @"21(?<uid>\d{20})");
        if (uidMatch.Success)
            uid = uidMatch.Groups["uid"].Value;
        else if (digitsOnly.Length == 20)
            uid = digitsOnly;

        string gtin = string.Empty;
        var gtinMatch = Regex.Match(digitsOnly, @"01(?<gtin>\d{14})");
        if (gtinMatch.Success)
            gtin = gtinMatch.Groups["gtin"].Value;

        return (gtin, uid, string.Empty);
    }

    private (string PersianName, string EnglishName) GetReceiveStatusFallbackProductNames(string barcode, string productName = "", string productEnName = "")
    {
        string persian = productName ?? string.Empty;
        string english = productEnName ?? string.Empty;

        if (_ttTeckDetailsByBarcode.TryGetValue(barcode, out var cachedInfo))
        {
            if (string.IsNullOrWhiteSpace(persian))
                persian = cachedInfo.PersianName ?? string.Empty;
            if (string.IsNullOrWhiteSpace(english))
                english = cachedInfo.EnglishName ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(persian) || string.IsNullOrWhiteSpace(english))
        {
            var historyRecord = HistoryItems.FirstOrDefault(x => x.Barcode.Equals(barcode, StringComparison.OrdinalIgnoreCase));
            if (historyRecord != null)
            {
                var split = SplitTtTeckProductNames(historyRecord.DrugName);
                if (string.IsNullOrWhiteSpace(persian))
                    persian = split.PersianName;
                if (string.IsNullOrWhiteSpace(english))
                    english = split.EnglishName;
            }
        }

        return (persian, english);
    }

    private (string Irc, string BatchCode, string GenericCode, string GenericName, string Expiration, string PersianName, string EnglishName) GetCachedReceiveStatusInfo(string barcode)
    {
        string irc = string.Empty;
        string batchCode = string.Empty;
        string genericCode = string.Empty;
        string genericName = string.Empty;
        string expiration = string.Empty;
        string persianName = string.Empty;
        string englishName = string.Empty;

        if (!_ttTeckDetailsByBarcode.TryGetValue(barcode, out var info) || info == null)
            return (irc, batchCode, genericCode, genericName, expiration, persianName, englishName);

        irc = info.IRC ?? string.Empty;
        persianName = info.PersianName ?? string.Empty;
        englishName = info.EnglishName ?? string.Empty;

        if (info.ExtraFields == null)
            return (irc, batchCode, genericCode, genericName, expiration, persianName, englishName);

        foreach (var field in info.ExtraFields)
        {
            string key = NormalizeProductIdentifierLabel(field.Key);
            string value = field.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(value))
                continue;

            if (string.IsNullOrWhiteSpace(irc) && (key == "irc" || key.Contains("شناسهفرآورده") || key.Contains("شناسهفراورده")))
                irc = new string(ToEnglishDigits(value).Where(char.IsDigit).ToArray());
            else if (string.IsNullOrWhiteSpace(batchCode) && (key.Contains("batch") || key.Contains("lot") || key.Contains("سریساخت") || key.Contains("شمارهبچ") || key.Contains("شمارهلات")))
                batchCode = ToEnglishDigits(value.Trim());
            else if (string.IsNullOrWhiteSpace(genericCode) && ((key.Contains("generic") && key.Contains("code")) || (key.Contains("ژنریک") && key.Contains("کد"))))
                genericCode = new string(ToEnglishDigits(value).Where(char.IsDigit).ToArray());
            else if (string.IsNullOrWhiteSpace(genericName) && (key.Contains("genericname") || key.Contains("نامژنریک")))
                genericName = value.Trim();
            else if (string.IsNullOrWhiteSpace(expiration) && (key.Contains("expiration") || key.Contains("expire") || key.Contains("expiry") || key.Contains("تاریخانقضا")))
                expiration = value.Trim();
        }

        return (irc, batchCode, genericCode, genericName, expiration, persianName, englishName);
    }

    private async Task<long?> TryGetProductIdByIrcAsync(string irc)
    {
        if (string.IsNullOrWhiteSpace(irc))
            return null;

        try
        {
            using var productDoc = await SendTtacJsonAsync(HttpMethod.Get, $"https://statisticsreports.ttac.ir/product/GetDetailProductByIrc?irc={Uri.EscapeDataString(irc)}");
            return productDoc == null ? null : FindLongRecursive(productDoc.RootElement, "Id");
        }
        catch
        {
            return null;
        }
    }

    private ReceiveStatusRow CreateReceiveStatusNotFoundRow(
        string barcode,
        string statusText,
        string productName = "",
        string productEnName = "",
        string irc = "",
        string uid = "",
        string gtin = "",
        string genericCode = "",
        string genericName = "",
        string lotNumber = "",
        string expiration = "")
    {
        var fallbackNames = GetReceiveStatusFallbackProductNames(barcode, productName, productEnName);
        string finalProductName = string.IsNullOrWhiteSpace(fallbackNames.PersianName) ? barcode : fallbackNames.PersianName;
        string finalProductEnName = fallbackNames.EnglishName;

        return new ReceiveStatusRow
        {
            ReceiveId = 0,
            IsConfirmable = false,
            Barcode = barcode,
            Irc = irc,
            UID = uid,
            GTIN = gtin,
            ProductName = finalProductName,
            ProductEnName = finalProductEnName,
            GenericCode = genericCode,
            GenericName = genericName,
            LotNumber = lotNumber,
            Expiration = expiration,
            SenderName = "-",
            Quantity = "-",
            SentDatePersian = "-",
            StatusText = statusText,
            StatusBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17))
        };
    }

    private async Task<ReceiveStatusRow?> BuildReceiveStatusRowAsync(string barcode)
    {
        var barcodeParts = ExtractTtacBarcodePartsForReceiveStatus(barcode);
        var fallbackNames = GetReceiveStatusFallbackProductNames(barcode);
        var cachedInfo = GetCachedReceiveStatusInfo(barcode);
        if (string.IsNullOrWhiteSpace(barcodeParts.LotNumber) && !string.IsNullOrWhiteSpace(cachedInfo.BatchCode))
            barcodeParts.LotNumber = cachedInfo.BatchCode;
        if (string.IsNullOrWhiteSpace(fallbackNames.PersianName) && !string.IsNullOrWhiteSpace(cachedInfo.PersianName))
            fallbackNames.PersianName = cachedInfo.PersianName;
        if (string.IsNullOrWhiteSpace(fallbackNames.EnglishName) && !string.IsNullOrWhiteSpace(cachedInfo.EnglishName))
            fallbackNames.EnglishName = cachedInfo.EnglishName;

        using var catalogDoc = await SendTtacJsonAsync(HttpMethod.Get, $"https://statisticsreports.ttac.ir/product/InstanceCatalog?uid={Uri.EscapeDataString(barcode)}");
        if (catalogDoc == null || !catalogDoc.RootElement.TryGetProperty("Result", out var catalog) || catalog.ValueKind != JsonValueKind.Object)
        {
            string knownIrc = cachedInfo.Irc;
            string knownBatch = !string.IsNullOrWhiteSpace(barcodeParts.LotNumber) ? barcodeParts.LotNumber : cachedInfo.BatchCode;
            long? knownProductId = await TryGetProductIdByIrcAsync(knownIrc);
            var alreadyConfirmedWithoutCatalog = await TryBuildAlreadyConfirmedReceiveStatusRowFallbackAsync(
                barcode,
                knownIrc,
                knownBatch,
                knownProductId?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                barcodeParts.Uid,
                barcodeParts.Gtin,
                fallbackNames.PersianName,
                fallbackNames.EnglishName,
                cachedInfo.GenericCode,
                cachedInfo.GenericName,
                cachedInfo.Expiration);
            if (alreadyConfirmedWithoutCatalog != null)
                return alreadyConfirmedWithoutCatalog;

            return CreateReceiveStatusNotFoundRow(
                barcode,
                _localization.GetString("ProductWasNotFoundInTTACCatalog"),
                fallbackNames.PersianName,
                fallbackNames.EnglishName,
                knownIrc,
                barcodeParts.Uid,
                barcodeParts.Gtin,
                cachedInfo.GenericCode,
                cachedInfo.GenericName,
                knownBatch,
                cachedInfo.Expiration);
        }

        string irc = ReadJsonString(catalog, "Irc") ?? string.Empty;
        string uid = ReadJsonString(catalog, "UID") ?? string.Empty;
        string gtin = ReadJsonString(catalog, "GTIN") ?? string.Empty;
        string batchCode = ReadJsonString(catalog, "BatchCode") ?? string.Empty;
        string persianName = ReadJsonString(catalog, "PersianName") ?? string.Empty;
        string englishName = ReadJsonString(catalog, "EnglishName") ?? string.Empty;
        string genericCodeText = ReadJsonString(catalog, "GenericCode") ?? string.Empty;
        string genericName = ReadJsonString(catalog, "GenericName") ?? string.Empty;
        string expiration = ReadJsonString(catalog, "Expiration") ?? string.Empty;

        if (string.IsNullOrWhiteSpace(uid)) uid = barcodeParts.Uid;
        if (string.IsNullOrWhiteSpace(gtin)) gtin = barcodeParts.Gtin;
        if (string.IsNullOrWhiteSpace(batchCode)) batchCode = !string.IsNullOrWhiteSpace(barcodeParts.LotNumber) ? barcodeParts.LotNumber : cachedInfo.BatchCode;
        if (string.IsNullOrWhiteSpace(persianName)) persianName = !string.IsNullOrWhiteSpace(fallbackNames.PersianName) ? fallbackNames.PersianName : cachedInfo.PersianName;
        if (string.IsNullOrWhiteSpace(englishName)) englishName = !string.IsNullOrWhiteSpace(fallbackNames.EnglishName) ? fallbackNames.EnglishName : cachedInfo.EnglishName;
        if (string.IsNullOrWhiteSpace(irc)) irc = cachedInfo.Irc;
        if (string.IsNullOrWhiteSpace(genericCodeText)) genericCodeText = cachedInfo.GenericCode;
        if (string.IsNullOrWhiteSpace(genericName)) genericName = cachedInfo.GenericName;
        if (string.IsNullOrWhiteSpace(expiration)) expiration = cachedInfo.Expiration;

        if (string.IsNullOrWhiteSpace(irc) || string.IsNullOrWhiteSpace(batchCode) || string.IsNullOrWhiteSpace(genericCodeText))
        {
            var alreadyConfirmedIncompleteCatalog = await TryBuildAlreadyConfirmedReceiveStatusRowFallbackAsync(
                barcode,
                irc,
                batchCode,
                string.Empty,
                uid,
                gtin,
                persianName,
                englishName,
                genericCodeText,
                genericName,
                expiration);
            if (alreadyConfirmedIncompleteCatalog != null)
                return alreadyConfirmedIncompleteCatalog;

            return CreateReceiveStatusNotFoundRow(barcode, _localization.GetString("CatalogResponseIsIncomplete"), persianName, englishName, irc, uid, gtin, genericCodeText, genericName, batchCode, expiration);
        }

        using var productDoc = await SendTtacJsonAsync(HttpMethod.Get, $"https://statisticsreports.ttac.ir/product/GetDetailProductByIrc?irc={Uri.EscapeDataString(irc)}");
        long? productId = productDoc == null ? null : FindLongRecursive(productDoc.RootElement, "Id");
        if (!productId.HasValue)
            throw new InvalidOperationException(_localization.GetString("ProductIDWasNotFound"));

        long? genericId = null;
        string normalizedGenericCode = ToEnglishDigits(genericCodeText).Trim();
        if (!string.IsNullOrWhiteSpace(normalizedGenericCode) && normalizedGenericCode != "0")
        {
            using var genericDoc = await SendTtacJsonAsync(HttpMethod.Get, $"https://statisticsreports.ttac.ir/Generic/GetByCode?code={Uri.EscapeDataString(normalizedGenericCode)}");
            genericId = genericDoc == null ? null : FindLongRecursive(genericDoc.RootElement, "Id");
        }

        string genericIdQuery = genericId.HasValue ? $"&genericId={genericId.Value}" : string.Empty;
        string listUrl = $"https://statisticsreports.ttac.ir/declaration/GetReceiveListToConfirm?searchExp=&productId={productId.Value}{genericIdQuery}&lotNumber={Uri.EscapeDataString(batchCode)}&PageSize=50&PageNumber=1";
        using var listDoc = await SendTtacJsonAsync(HttpMethod.Get, listUrl);
        if (listDoc == null || !listDoc.RootElement.TryGetProperty("Result", out var list) || list.ValueKind != JsonValueKind.Array || list.GetArrayLength() == 0)
        {
            var alreadyConfirmedRow = await TryBuildAlreadyConfirmedReceiveStatusRowAsync(
                barcode,
                productId.Value,
                batchCode,
                irc,
                uid,
                gtin,
                persianName,
                englishName,
                genericCodeText,
                genericName,
                expiration);
            if (alreadyConfirmedRow != null)
                return alreadyConfirmedRow;

            return CreateReceiveStatusNotFoundRow(
                barcode,
                _localization.GetString("NoReceivableItemWasFoundForThisProductBatch"),
                persianName,
                englishName,
                irc,
                uid,
                gtin,
                genericCodeText,
                genericName,
                batchCode,
                expiration);
        }

        // ابتدا سعی می‌شود دقیقاً همان ردیفی از لیست پیدا شود که UID آن با UID بسته‌ی
        // اسکن‌شده یکی است — نه صرفاً اولین ردیف با همان محصول/سری‌ساخت. اگر تی‌تک UID را در
        // پاسخ این endpoint برنمی‌گرداند (یا با بسته‌ی ما مطابقتی پیدا نشد)، برای حفظ رفتار قبلی
        // به‌عنوان آخرین راه‌حل اولین ردیف انتخاب می‌شود.
        JsonElement item = default;
        bool matchedByUid = false;
        if (!string.IsNullOrWhiteSpace(uid))
        {
            foreach (var candidate in list.EnumerateArray())
            {
                string? candidateUid = FindStringRecursive(candidate, "UID", "Uid", "uid", "PackageUID", "InstanceUID");
                if (!string.IsNullOrWhiteSpace(candidateUid) && string.Equals(candidateUid.Trim(), uid.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    item = candidate;
                    matchedByUid = true;
                    break;
                }
            }
        }

        if (!matchedByUid)
            item = list.EnumerateArray().First();

        long? receiveId = FindLongRecursive(item, "Id");
        if (!receiveId.HasValue)
            throw new InvalidOperationException(_localization.GetString("ReceiveItemIDWasNotFound"));

        string productName = ReadJsonString(item, "ProductName") ?? persianName;
        string productEnName = ReadJsonString(item, "ProductEnName") ?? englishName;
        string lotNumber = ReadJsonString(item, "LotNumber") ?? batchCode;
        string senderName = ReadJsonString(item, "SenderName") ?? string.Empty;
        string quantity = ReadJsonString(item, "Quantity") ?? string.Empty;
        string sentDatePersian = ReadJsonString(item, "SentDatePersian") ?? string.Empty;
        string confirmStatus = ReadJsonString(item, "ConfirmStatusDesc") ?? string.Empty;

        return new ReceiveStatusRow
        {
            ReceiveId = receiveId.Value,
            Barcode = barcode,
            Irc = irc,
            UID = uid,
            GTIN = gtin,
            ProductName = productName,
            ProductEnName = productEnName,
            GenericCode = genericCodeText,
            GenericName = genericName,
            LotNumber = lotNumber,
            Expiration = expiration,
            SenderName = senderName,
            Quantity = quantity,
            SentDatePersian = sentDatePersian,
            IsConfirmable = true,
            StatusText = string.IsNullOrWhiteSpace(confirmStatus) ? (_localization.GetString("ReadyToConfirm")) : confirmStatus,
            StatusBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0))
        };
    }

    private async Task<ReceiveStatusRow?> TryBuildAlreadyConfirmedReceiveStatusRowFallbackAsync(
        string barcode,
        string irc,
        string batchCode,
        string productIdText,
        string uid,
        string gtin,
        string persianName,
        string englishName,
        string genericCodeText,
        string genericName,
        string expiration)
    {
        var urls = new List<(string Url, string MatchMode)>();
        if (!string.IsNullOrWhiteSpace(productIdText) && !string.IsNullOrWhiteSpace(batchCode))
            urls.Add(($"https://statisticsreports.ttac.ir/declaration/confirmedList?searchExp=&productId={Uri.EscapeDataString(productIdText)}&lotNumber={Uri.EscapeDataString(batchCode)}&PageSize=50&PageNumber=1", "product-lot"));
        if (!string.IsNullOrWhiteSpace(batchCode))
            urls.Add(($"https://statisticsreports.ttac.ir/declaration/confirmedList?searchExp=&lotNumber={Uri.EscapeDataString(batchCode)}&PageSize=50&PageNumber=1", "lot-only"));
        if (!string.IsNullOrWhiteSpace(uid))
            urls.Add(($"https://statisticsreports.ttac.ir/declaration/confirmedList?searchExp={Uri.EscapeDataString(uid)}&PageSize=50&PageNumber=1", "search"));
        if (!string.IsNullOrWhiteSpace(gtin))
            urls.Add(($"https://statisticsreports.ttac.ir/declaration/confirmedList?searchExp={Uri.EscapeDataString(gtin)}&PageSize=50&PageNumber=1", "search"));
        if (!string.IsNullOrWhiteSpace(barcode))
            urls.Add(($"https://statisticsreports.ttac.ir/declaration/confirmedList?searchExp={Uri.EscapeDataString(barcode)}&PageSize=50&PageNumber=1", "search"));
        if (!string.IsNullOrWhiteSpace(persianName))
            urls.Add(($"https://statisticsreports.ttac.ir/declaration/confirmedList?searchExp={Uri.EscapeDataString(persianName)}&PageSize=50&PageNumber=1", "search"));

        foreach (var request in urls.DistinctBy(x => x.Url))
        {
            try
            {
                using var confirmedDoc = await SendTtacJsonAsync(HttpMethod.Get, request.Url);
                var row = BuildAlreadyConfirmedReceiveStatusRowFromDocument(
                    confirmedDoc,
                    barcode,
                    irc,
                    batchCode,
                    uid,
                    gtin,
                    persianName,
                    englishName,
                    genericCodeText,
                    genericName,
                    expiration,
                    request.MatchMode);
                if (row != null)
                    return row;
            }
            catch { }
        }

        return null;
    }

    private ReceiveStatusRow? BuildAlreadyConfirmedReceiveStatusRowFromDocument(
        JsonDocument? confirmedDoc,
        string barcode,
        string irc,
        string batchCode,
        string uid,
        string gtin,
        string persianName,
        string englishName,
        string genericCodeText,
        string genericName,
        string expiration,
        string matchMode)
    {
        if (confirmedDoc == null ||
            !confirmedDoc.RootElement.TryGetProperty("Result", out var result) ||
            result.ValueKind != JsonValueKind.Array ||
            result.GetArrayLength() == 0)
        {
            return null;
        }

        var candidates = result.EnumerateArray().ToList();
        JsonElement item = default;
        bool exactIrcLot = false;
        bool productLotMatch = false;
        bool lotOnlyMatch = false;

        if (!string.IsNullOrWhiteSpace(irc) && !string.IsNullOrWhiteSpace(batchCode))
        {
            item = candidates.FirstOrDefault(x =>
                string.Equals(ReadJsonString(x, "Irc"), irc, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(ReadJsonString(x, "LotNumber"), batchCode, StringComparison.OrdinalIgnoreCase));
            exactIrcLot = item.ValueKind != JsonValueKind.Undefined && item.ValueKind != JsonValueKind.Null;
        }

        if (!exactIrcLot && string.Equals(matchMode, "product-lot", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(batchCode))
        {
            item = candidates.FirstOrDefault(x => string.Equals(ReadJsonString(x, "LotNumber"), batchCode, StringComparison.OrdinalIgnoreCase));
            productLotMatch = item.ValueKind != JsonValueKind.Undefined && item.ValueKind != JsonValueKind.Null;
        }

        if (!exactIrcLot && !productLotMatch && string.Equals(matchMode, "lot-only", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(batchCode))
        {
            item = candidates.FirstOrDefault(x => string.Equals(ReadJsonString(x, "LotNumber"), batchCode, StringComparison.OrdinalIgnoreCase));
            lotOnlyMatch = item.ValueKind != JsonValueKind.Undefined && item.ValueKind != JsonValueKind.Null;
        }

        if (!exactIrcLot && !productLotMatch && !lotOnlyMatch && string.Equals(matchMode, "search", StringComparison.OrdinalIgnoreCase))
        {
            item = candidates.FirstOrDefault(x =>
                (!string.IsNullOrWhiteSpace(irc) && string.Equals(ReadJsonString(x, "Irc"), irc, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrWhiteSpace(batchCode) && string.Equals(ReadJsonString(x, "LotNumber"), batchCode, StringComparison.OrdinalIgnoreCase)));
        }

        if (item.ValueKind == JsonValueKind.Undefined || item.ValueKind == JsonValueKind.Null)
            return null;

        bool strongMatch = exactIrcLot || productLotMatch;
        bool probableMatch = !strongMatch;

        string productName = ReadJsonString(item, "ProductName") ?? persianName;
        string productEnName = ReadJsonString(item, "ProductEnName") ?? englishName;
        string lotNumber = ReadJsonString(item, "LotNumber") ?? batchCode;
        string senderName = ReadJsonString(item, "SenderName") ?? string.Empty;
        string quantity = ReadJsonString(item, "ConfirmedQuantity") ?? ReadJsonString(item, "Quantity") ?? string.Empty;
        string sentDatePersian = ReadJsonString(item, "SentDatePersian") ?? string.Empty;
        string confirmedDescription = ReadJsonString(item, "ConfirmDescription") ?? string.Empty;
        string rowGenericCode = ReadJsonString(item, "GenericCode") ?? genericCodeText;
        string rowGenericName = ReadJsonString(item, "GenericName") ?? genericName;
        string rowIrc = ReadJsonString(item, "Irc") ?? irc;

        string baseStatus = strongMatch
            ? (_localization.GetString("AlreadyReceivedConfirmed"))
            : (_localization.GetString("ProbablyAlreadyReceivedConfirmedNeedsReview"));

        return new ReceiveStatusRow
        {
            ReceiveId = 0,
            IsConfirmable = false,
            Barcode = barcode,
            Irc = rowIrc,
            UID = uid,
            GTIN = gtin,
            ProductName = productName,
            ProductEnName = productEnName,
            GenericCode = rowGenericCode,
            GenericName = rowGenericName,
            LotNumber = lotNumber,
            Expiration = expiration,
            SenderName = senderName,
            Quantity = quantity,
            SentDatePersian = sentDatePersian,
            StatusText = string.IsNullOrWhiteSpace(confirmedDescription) ? baseStatus : $"{baseStatus} - {confirmedDescription}",
            StatusBrush = strongMatch ? System.Windows.Media.Brushes.Green : new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17))
        };
    }

    private async Task<ReceiveStatusRow?> TryBuildAlreadyConfirmedReceiveStatusRowAsync(
        string barcode,
        long productId,
        string batchCode,
        string irc,
        string uid,
        string gtin,
        string persianName,
        string englishName,
        string genericCodeText,
        string genericName,
        string expiration)
    {
        try
        {
            string url = $"https://statisticsreports.ttac.ir/declaration/confirmedList?searchExp=&productId={productId}&lotNumber={Uri.EscapeDataString(batchCode)}&PageSize=50&PageNumber=1";
            using var confirmedDoc = await SendTtacJsonAsync(HttpMethod.Get, url);
            return BuildAlreadyConfirmedReceiveStatusRowFromDocument(
                confirmedDoc,
                barcode,
                irc,
                batchCode,
                uid,
                gtin,
                persianName,
                englishName,
                genericCodeText,
                genericName,
                expiration,
                "product-lot");
        }
        catch
        {
            return null;
        }
    }

    private async void ConfirmReceiveStatusButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button button || button.Tag is not ReceiveStatusRow row)
            return;

        if (!row.IsConfirmable || row.ReceiveId <= 0)
        {
            ShowStyledMessage(
                _localization.GetString("ReceiveStatus"),
                string.IsNullOrWhiteSpace(row.StatusText)
                    ? (_localization.GetString("ThisItemIsNotConfirmable"))
                    : row.StatusText,
                true);
            return;
        }

        button.IsEnabled = false;
        try
        {
            using var doc = await SendTtacJsonAsync(HttpMethod.Post, "https://statisticsreports.ttac.ir/declaration/ConfirmReceive", new { Id = row.ReceiveId });
            string? message = doc == null ? null : ReadJsonString(doc.RootElement, "Message");
            row.StatusText = string.IsNullOrWhiteSpace(message) || message == "null"
                ? (_localization.GetString("ConfirmedSuccessfully"))
                : message;
            row.IsConfirmable = false;
            row.StatusBrush = System.Windows.Media.Brushes.Green;
            ReceiveStatusListBox.Items.Refresh();
            SaveCurrentReceiveStatusItemsForCurrentPharmacy();
            ShowStyledMessage(_localization.GetString("ReceiveStatus"), row.StatusText);
        }
        catch (Exception ex)
        {
            if (IsJsonNullElementException(ex))
            {
                row.StatusText = _localization.GetString("TTACReturnedAnEmptyResponseAfterConfirmationPleaseRefreshCheckThePortal");
                row.StatusBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17));
                row.IsConfirmable = false;
                ReceiveStatusListBox.Items.Refresh();
                SaveCurrentReceiveStatusItemsForCurrentPharmacy();
                return;
            }

            row.StatusText = ex.Message;
            row.StatusBrush = System.Windows.Media.Brushes.Firebrick;
            ReceiveStatusListBox.Items.Refresh();
            SaveCurrentReceiveStatusItemsForCurrentPharmacy();
            HandleTtacOperationException(ex,
                _localization.GetString("ReceiveStatusFailed"),
                async () =>
                {
                    ConfirmReceiveStatusButton_Click(button, new RoutedEventArgs());
                    await Task.CompletedTask;
                },
                pendingLabel: _localization.GetString("PendingConfirmReceiveStatus"));
        }
        finally
        {
            button.IsEnabled = true;
        }
    }

    private void ReceiveStatusCopyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is string text && !string.IsNullOrWhiteSpace(text))
            System.Windows.Clipboard.SetText(CleanBarcodeForExternalUse(text));
    }

    private void DeleteReceiveStatusButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.Button button && button.Tag is ReceiveStatusRow row)
        {
            ReceiveStatusItems.Remove(row);
            if (!string.IsNullOrWhiteSpace(row.Barcode))
                _receiveStatusKnownBarcodes.Remove(row.Barcode);
            RefreshReceiveStatusRowNumbers();
            SaveCurrentReceiveStatusItemsForCurrentPharmacy();
        }
    }

    private void ReceiveStatusExportExcelButton_Click(object sender, RoutedEventArgs e)
    {
        ExportReceiveStatusToExcel();
    }

    private void ReceiveStatusExportPdfButton_Click(object sender, RoutedEventArgs e)
    {
        PrintReceiveStatusPdfReport();
    }

    private string[] GetReceiveStatusExportHeaders()
    {
        return _localization.GetStringArray("ReceiveStatusExportHeaders");
    }

    private void ExportReceiveStatusToExcel()
    {
        var rows = ReceiveStatusItems.ToList();
        if (rows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        var saveFileDialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"Scanbridge_ReceiveStatus_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = _localization.GetString("SaveReceiveStatusReport")
        };

        if (saveFileDialog.ShowDialog() != true)
            return;

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("ReceiveStatus");
            string[] headers = GetReceiveStatusExportHeaders();
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(8, 145, 178);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int r = 2;
            int index = 1;
            foreach (var row in rows)
            {
                worksheet.Cell(r, 1).Value = index++;
                worksheet.Cell(r, 2).Value = row.ProductName;
                worksheet.Cell(r, 3).Value = row.ProductEnName;
                worksheet.Cell(r, 4).Value = row.Barcode;
                worksheet.Cell(r, 5).Value = row.Irc;
                worksheet.Cell(r, 6).Value = row.GenericCode;
                worksheet.Cell(r, 7).Value = row.LotNumber;
                worksheet.Cell(r, 8).Value = row.Quantity;
                worksheet.Cell(r, 9).Value = row.SenderName;
                worksheet.Cell(r, 10).Value = row.SentDatePersian;
                worksheet.Cell(r, 11).Value = row.StatusText;
                r++;
            }
            worksheet.Columns().AdjustToContents();
            worksheet.RangeUsed()?.CreateTable();
            workbook.SaveAs(saveFileDialog.FileName);
            ShowStyledMessage(_localization.GetString("ExportSuccessful"), saveFileDialog.FileName);
        }
        catch (Exception ex)
        {
            ShowStyledMessage(_localization.GetString("ExportFailed"), ex.Message, true);
        }
    }

    private void PrintReceiveStatusPdfReport()
    {
        var rows = ReceiveStatusItems.ToList();
        if (rows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        var printDialog = new System.Windows.Controls.PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var document = CreateReceiveStatusReportDocument(rows);
        document.PageWidth = printDialog.PrintableAreaWidth;
        document.PageHeight = printDialog.PrintableAreaHeight;
        document.PagePadding = new Thickness(28);
        document.ColumnWidth = printDialog.PrintableAreaWidth;
        printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, "Scanbridge Receive Status Report");
        ShowStyledMessage(_localization.GetString("PDFReport"), GetLocalizedPdfSuccessPathText());
    }

    private FlowDocument CreateReceiveStatusReportDocument(List<ReceiveStatusRow> rows)
    {
        var document = new FlowDocument
        {
            FlowDirection = _localization.CurrentLanguage == AppLanguage.English ? System.Windows.FlowDirection.LeftToRight : System.Windows.FlowDirection.RightToLeft,
            FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
            FontSize = 9
        };
        document.Blocks.Add(new Paragraph(new Run(_localization.GetString("ScanbridgeReceiveStatusReport")))
        {
            FontSize = 22,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x1A, 0x23, 0x7E)),
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 14)
        });

        var table = new Table { CellSpacing = 0 };
        document.Blocks.Add(table);
        string[] headers = GetReceiveStatusExportHeaders();
        foreach (string _ in headers)
            table.Columns.Add(new TableColumn());
        var group = new TableRowGroup();
        table.RowGroups.Add(group);
        var headerRow = new TableRow();
        group.Rows.Add(headerRow);
        foreach (string header in headers)
            headerRow.Cells.Add(CreatePdfCell(header, true));

        int index = 1;
        foreach (var row in rows)
        {
            var tr = new TableRow();
            group.Rows.Add(tr);
            tr.Cells.Add(CreatePdfCell(index++.ToString(), false));
            tr.Cells.Add(CreatePdfCell(row.ProductName, false));
            tr.Cells.Add(CreatePdfCell(row.ProductEnName, false));
            tr.Cells.Add(CreatePdfCell(row.Barcode, false));
            tr.Cells.Add(CreatePdfCell(row.Irc, false));
            tr.Cells.Add(CreatePdfCell(row.GenericCode, false));
            tr.Cells.Add(CreatePdfCell(row.LotNumber, false));
            tr.Cells.Add(CreatePdfCell(row.Quantity, false));
            tr.Cells.Add(CreatePdfCell(row.SenderName, false));
            tr.Cells.Add(CreatePdfCell(row.SentDatePersian, false));
            tr.Cells.Add(CreatePdfCell(row.StatusText, false));
        }

        return document;
    }

    // ---------- Dedicated Ttac Panel ----------

    private async void TtacPanelButton_Click(object sender, RoutedEventArgs e)
    {
        string? token = await GetTtacAccessTokenAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () =>
            {
                OpenTtacPanelNow();
                await Task.CompletedTask;
            };
            _pendingTtacRetryLabel = _localization.GetString("PendingOpenTtacPanel");
            ShowTtacLoginOverlay();
            return;
        }

        OpenTtacPanelNow();
    }

    private void OpenTtacPanelNow()
    {
        _isTtacPanelFormulaOnly = false;
        RefreshTtTeckHistoryItems();
        UpdateTtacPanelButtons();
        TtacPanelOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void TtacPanelAllButton_Click(object sender, RoutedEventArgs e)
    {
        _isTtacPanelFormulaOnly = false;
        RefreshTtTeckHistoryItems();
        UpdateTtacPanelButtons();
    }

    private void TtacPanelFormulaButton_Click(object sender, RoutedEventArgs e)
    {
        _isTtacPanelFormulaOnly = true;
        RefreshTtTeckHistoryItems();
        UpdateTtacPanelButtons();
    }

    private void TtacPanelOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseTtacPanel();
    }

    private void TtacPanelCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseTtacPanelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseTtacPanel();
    }

    private void CloseTtacPanel()
    {
        TtacPanelOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void UpdateTtacPanelButtons()
    {
        if (TtacPanelTitle == null)
            return;

        TtacPanelTitle.Text = _localization.GetString("TtTeckPanel2");

        TtacPanelAllButton.Content = _localization.GetString("AllTtTeckItems");

        TtacPanelFormulaButton.Content = _localization.GetString("Formula");

        if (TtacPanelReceiveStatusButton != null)
            TtacPanelReceiveStatusButton.Content = _localization.GetString("ReceiveStatus");
        if (TtacPanelExportExcelButton != null)
            TtacPanelExportExcelButton.Content = _localization.GetString("Excel");
        if (TtacPanelExportPdfButton != null)
            TtacPanelExportPdfButton.Content = _localization.GetString("PDF");

        if (TtacManualBarcodeLabel != null)
        {
            bool english = _localization.CurrentLanguage == AppLanguage.English;
            TtacManualBarcodeLabel.Text = _localization.GetString("ManualUIDBarcode");
            TtacManualBarcodeTextBox.ToolTip = _localization.GetString("IfScanningIsNotPossibleEnterTheUIDOrFullBarcodeHere");
            TtacManualAddButton.Content = _localization.GetString("AddManually");
            TtacPanelSearchLabel.Text = _localization.GetString("Filter");
            TtacPanelSearchTextBox.ToolTip = _localization.GetString("SearchByProductBarcodeNationalIDMobilePatientNameDateOrTime");
            TtacPanelClearSearchButton.Content = _localization.GetString("Clear");
        }

        TtacPanelAllButton.Background = new SolidColorBrush(_isTtacPanelFormulaOnly
            ? System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF)
            : System.Windows.Media.Color.FromRgb(0x7C, 0x3A, 0xED));

        TtacPanelFormulaButton.Background = new SolidColorBrush(_isTtacPanelFormulaOnly
            ? System.Windows.Media.Color.FromRgb(0xBE, 0x18, 0x5D)
            : System.Windows.Media.Color.FromRgb(0xEC, 0x48, 0x99));
    }

    private void TtacPanelSearchTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        _ttacPanelSearchText = TtacPanelSearchTextBox?.Text?.Trim() ?? string.Empty;
        RefreshTtTeckHistoryItems();
    }

    private void TtacPanelClearSearchButton_Click(object sender, RoutedEventArgs e)
    {
        TtacPanelSearchTextBox.Text = string.Empty;
        _ttacPanelSearchText = string.Empty;
        RefreshTtTeckHistoryItems();
    }

    private void ReceiveStatusSearchTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        _receiveStatusSearchText = ReceiveStatusSearchTextBox?.Text?.Trim() ?? string.Empty;
        System.Windows.Data.CollectionViewSource.GetDefaultView(ReceiveStatusItems).Refresh();
    }

    private void ReceiveStatusClearSearchButton_Click(object sender, RoutedEventArgs e)
    {
        ReceiveStatusSearchTextBox.Text = string.Empty;
        _receiveStatusSearchText = string.Empty;
        System.Windows.Data.CollectionViewSource.GetDefaultView(ReceiveStatusItems).Refresh();
    }

    private void TtacPanelExportExcelButton_Click(object sender, RoutedEventArgs e)
    {
        ExportTtacPanelItemsToExcel();
    }

    private void TtacPanelExportPdfButton_Click(object sender, RoutedEventArgs e)
    {
        PrintTtacPanelPdfReport();
    }

    private string GetTtacPanelExportTitle()
    {
        if (_isTtacPanelFormulaOnly)
        {
            return _localization.GetString("ScanbridgeFormulaItemsReport");
        }

        return _localization.GetString("ScanbridgeTtTeckItemsReport");
    }

    private string[] GetTtacPanelExportHeaders()
    {
        return _localization.GetStringArray("TtacPanelExportHeaders");
    }

    private List<TtTeckHistoryRow> GetCurrentTtacPanelExportRows()
    {
        return TtTeckHistoryItems
            .OrderByDescending(r => r.TimestampLocal)
            .ToList();
    }

    private void ExportTtacPanelItemsToExcel()
    {
        var rows = GetCurrentTtacPanelExportRows();
        if (rows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        var saveFileDialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"{(_isTtacPanelFormulaOnly ? "Scanbridge_Formula" : "Scanbridge_TtTeck")}_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = _localization.GetString("SaveReportToExcel")
        };

        if (saveFileDialog.ShowDialog() != true)
            return;

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(_isTtacPanelFormulaOnly ? "Formula" : "TtTeck");
            string[] headers = GetTtacPanelExportHeaders();

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(124, 58, 237);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int rowNumber = 2;
            int index = 1;
            foreach (var row in rows)
            {
                worksheet.Cell(rowNumber, 1).Value = index;
                worksheet.Cell(rowNumber, 2).Value = row.PersianDateText;
                worksheet.Cell(rowNumber, 3).Value = row.TimeText;
                worksheet.Cell(rowNumber, 4).Value = row.DeviceName;
                worksheet.Cell(rowNumber, 5).Value = row.Barcode;
                worksheet.Cell(rowNumber, 6).Value = row.PersianProductName;
                worksheet.Cell(rowNumber, 7).Value = row.EnglishProductName;
                worksheet.Cell(rowNumber, 8).Value = row.RegistrationButtonText;

                for (int col = 1; col <= headers.Length; col++)
                    worksheet.Cell(rowNumber, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                rowNumber++;
                index++;
            }

            worksheet.Column(1).Width = 8;
            worksheet.Column(2).Width = 15;
            worksheet.Column(3).Width = 12;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 48;
            worksheet.Column(6).Width = 42;
            worksheet.Column(7).Width = 42;
            worksheet.Column(8).Width = 18;
            worksheet.RangeUsed()?.CreateTable();
            workbook.SaveAs(saveFileDialog.FileName);

            ShowStyledMessage(
                _localization.GetString("ExportSuccessful"),
                (_localization.GetString("ExcelReportWasSavedSuccessfullyN")) + saveFileDialog.FileName);
        }
        catch (Exception ex)
        {
            ShowStyledMessage(_localization.GetString("ExportFailed"), ex.Message, true);
        }
    }

    private void PrintTtacPanelPdfReport()
    {
        var rows = GetCurrentTtacPanelExportRows();
        if (rows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        var printDialog = new System.Windows.Controls.PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var document = CreateTtacPanelReportDocument(rows);
        document.PageWidth = printDialog.PrintableAreaWidth;
        document.PageHeight = printDialog.PrintableAreaHeight;
        document.PagePadding = new Thickness(32);
        document.ColumnWidth = printDialog.PrintableAreaWidth;

        printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, GetTtacPanelExportTitle());
        ShowStyledMessage(
            _localization.GetString("PDFReport"),
            GetLocalizedPdfSuccessPathText());
    }

    private FlowDocument CreateTtacPanelReportDocument(List<TtTeckHistoryRow> rows)
    {
        var flowDirection = _localization.CurrentLanguage == AppLanguage.English
            ? System.Windows.FlowDirection.LeftToRight
            : System.Windows.FlowDirection.RightToLeft;

        var document = new FlowDocument
        {
            FlowDirection = flowDirection,
            FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
            FontSize = 10
        };

        document.Blocks.Add(new Paragraph(new Run(GetTtacPanelExportTitle()))
        {
            FontSize = 22,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x1A, 0x23, 0x7E)),
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 14)
        });

        document.Blocks.Add(new Paragraph(new Run(DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)))
        {
            FontSize = 12,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x4B, 0x55, 0x63)),
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 18)
        });

        var table = new Table { CellSpacing = 0 };
        document.Blocks.Add(table);
        string[] headers = GetTtacPanelExportHeaders();
        foreach (string _ in headers)
            table.Columns.Add(new TableColumn());

        var group = new TableRowGroup();
        table.RowGroups.Add(group);
        var headerRow = new TableRow();
        group.Rows.Add(headerRow);
        foreach (string header in headers)
            headerRow.Cells.Add(CreatePdfCell(header, true));

        int index = 1;
        foreach (var row in rows)
        {
            var tableRow = new TableRow();
            group.Rows.Add(tableRow);
            tableRow.Cells.Add(CreatePdfCell(index.ToString(), false));
            tableRow.Cells.Add(CreatePdfCell(row.PersianDateText, false));
            tableRow.Cells.Add(CreatePdfCell(row.TimeText, false));
            tableRow.Cells.Add(CreatePdfCell(row.DeviceName, false));
            tableRow.Cells.Add(CreatePdfCell(row.Barcode, false));
            tableRow.Cells.Add(CreatePdfCell(row.PersianProductName, false));
            tableRow.Cells.Add(CreatePdfCell(row.EnglishProductName, false));
            tableRow.Cells.Add(CreatePdfCell(row.RegistrationButtonText, false));
            index++;
        }

        return document;
    }

    private void TtacManualBarcodeTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;

        e.Handled = true;
        _ = AddManualTtacBarcodeAsync();
    }

    private void TtacManualAddButton_Click(object sender, RoutedEventArgs e)
    {
        _ = AddManualTtacBarcodeAsync();
    }

    private async Task AddManualTtacBarcodeAsync()
    {
        string barcode = CleanBarcodeForExternalUse(TtacManualBarcodeTextBox.Text);
        if (string.IsNullOrWhiteSpace(barcode))
        {
            ShowStyledMessage(
                _localization.GetString("ManualEntry"),
                _localization.GetString("EnterTheUIDOrFullBarcodeFirst"),
                true);
            return;
        }

        string digitsOnly = new string(ToEnglishDigits(barcode).Where(char.IsDigit).ToArray());
        if (digitsOnly.Length < 20 && !IsTtTeckLookupCandidate(barcode, BarcodeDetector.DetectBarcodeType(barcode)))
        {
            ShowStyledMessage(
                _localization.GetString("InvalidUID"),
                _localization.GetString("TheEnteredValueDoesNotLookLikeATTACUIDBarcode"),
                true);
            return;
        }

        TtacManualAddButton.IsEnabled = false;
        string originalButtonText = TtacManualAddButton.Content?.ToString() ?? string.Empty;
        TtacManualAddButton.Content = _localization.GetString("Adding");

        try
        {
            var record = new ScanRecord(DateTime.Now, barcode, _localization.GetString("ManualEntry"))
            {
                Source = BarcodeDetector.DetectBarcodeType(barcode),
                DrugName = GetTtTeckLookupPendingText()
            };

            AddHistoryRecord(record);
            SaveHistoryItemsToCsv();
            TtacManualBarcodeTextBox.Text = string.Empty;
            await LookupTtTeckForRecordAsync(record, false);
            SaveHistoryItemsToCsv();
            RefreshTtTeckHistoryItems();
            UpdateTtacPanelButtons();
        }
        catch (Exception ex)
        {
            ShowStyledMessage(GetLocalizedLookupFailedTitle(), ex.Message, true);
        }
        finally
        {
            TtacManualAddButton.IsEnabled = true;
            TtacManualAddButton.Content = originalButtonText;
            TtacManualBarcodeTextBox.Focus();
        }
    }

    // ---------- History Modal ----------

    private void HistoryButton_Click(object sender, RoutedEventArgs e)
    {
        _isTtTeckHistoryFilterActive = false;
        _isFormulaHistoryFilterActive = false;
        ApplyHistoryFilterMode();
        SuccessMessageGrid.Visibility = Visibility.Collapsed;
        HistoryOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void HistoryOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseHistoryOverlay();
    }

    private void HistoryCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseHistoryOverlay_Click(object sender, RoutedEventArgs e)
    {
        CloseHistoryOverlay();
    }

    private void CloseHistoryOverlay()
    {
        HistoryOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void HistoryClearButton_Click(object sender, RoutedEventArgs e)
    {
        ConfirmDeleteHistoryOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void ConfirmDeleteOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseConfirmDeleteOverlay();
    }

    private void ConfirmDeleteCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseConfirmDeleteOverlay()
    {
        ConfirmDeleteHistoryOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void ConfirmDeleteYesButton_Click(object sender, RoutedEventArgs e)
    {
        _service.ClearHistoryFile();
        HistoryItems.Clear();
        HistoryViewItems.Clear();
        TtTeckHistoryItems.Clear();
        CloseConfirmDeleteOverlay();
        CloseHistoryOverlay();
    }

    private void ConfirmDeleteNoButton_Click(object sender, RoutedEventArgs e)
    {
        CloseConfirmDeleteOverlay();
    }

    private void InitializeDateRangeFilterControls()
    {
        _isDateRangeInitializing = true;

        int currentYear = _persianCalendar.GetYear(DateTime.Now);
        FillIntCombo(FromYearComboBox, currentYear - 5, currentYear + 1);
        FillIntCombo(ToYearComboBox, currentYear - 5, currentYear + 1);
        FillIntCombo(FromMonthComboBox, 1, 12);
        FillIntCombo(ToMonthComboBox, 1, 12);
        FillIntCombo(FromHourComboBox, 0, 23);
        FillIntCombo(ToHourComboBox, 0, 23);
        FillIntCombo(FromMinuteComboBox, 0, 59);
        FillIntCombo(ToMinuteComboBox, 0, 59);

        SetDateRangeComboSelection(DateTime.Now.Date, DateTime.Now.Date.AddHours(23).AddMinutes(59));
        _isDateRangeInitializing = false;
    }

    private static void FillIntCombo(System.Windows.Controls.ComboBox comboBox, int from, int to)
    {
        comboBox.Items.Clear();
        for (int i = from; i <= to; i++)
            comboBox.Items.Add(i);
    }

    private void SetDateRangeComboSelection(DateTime from, DateTime to)
    {
        SetPersianDateComboSelection(FromYearComboBox, FromMonthComboBox, FromDayComboBox, from);
        SelectComboValue(FromHourComboBox, from.Hour);
        SelectComboValue(FromMinuteComboBox, from.Minute);

        SetPersianDateComboSelection(ToYearComboBox, ToMonthComboBox, ToDayComboBox, to);
        SelectComboValue(ToHourComboBox, to.Hour);
        SelectComboValue(ToMinuteComboBox, to.Minute);
    }

    private void SetPersianDateComboSelection(System.Windows.Controls.ComboBox yearCombo, System.Windows.Controls.ComboBox monthCombo, System.Windows.Controls.ComboBox dayCombo, DateTime dateTime)
    {
        int year = _persianCalendar.GetYear(dateTime);
        int month = _persianCalendar.GetMonth(dateTime);
        int day = _persianCalendar.GetDayOfMonth(dateTime);

        SelectComboValue(yearCombo, year);
        SelectComboValue(monthCombo, month);
        UpdateDayCombo(yearCombo, monthCombo, dayCombo, day);
    }

    private static void SelectComboValue(System.Windows.Controls.ComboBox comboBox, int value)
    {
        if (comboBox.Items.Contains(value))
            comboBox.SelectedItem = value;
        else if (comboBox.Items.Count > 0)
            comboBox.SelectedIndex = 0;
    }

    private void DateRangeDatePart_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_isDateRangeInitializing)
            return;

        UpdateDayCombo(FromYearComboBox, FromMonthComboBox, FromDayComboBox);
        UpdateDayCombo(ToYearComboBox, ToMonthComboBox, ToDayComboBox);
    }

    private void UpdateDayCombo(System.Windows.Controls.ComboBox yearCombo, System.Windows.Controls.ComboBox monthCombo, System.Windows.Controls.ComboBox dayCombo, int? preferredDay = null)
    {
        if (yearCombo.SelectedItem is not int year || monthCombo.SelectedItem is not int month)
            return;

        int selectedDay = preferredDay ?? (dayCombo.SelectedItem is int day ? day : 1);
        int daysInMonth = GetPersianMonthDays(year, month);

        dayCombo.Items.Clear();
        for (int i = 1; i <= daysInMonth; i++)
            dayCombo.Items.Add(i);

        SelectComboValue(dayCombo, Math.Min(selectedDay, daysInMonth));
    }

    private int GetPersianMonthDays(int year, int month)
    {
        if (month <= 6)
            return 31;
        if (month <= 11)
            return 30;
        return _persianCalendar.IsLeapYear(year) ? 30 : 29;
    }

    private void HistoryDateRangeButton_Click(object sender, RoutedEventArgs e)
    {
        OpenDateRangeFilter();
    }

    private void OpenDateRangeFilter()
    {
        DateTime from = _historyFilterFrom ?? DateTime.Now.Date;
        DateTime to = _historyFilterTo ?? DateTime.Now.Date.AddHours(23).AddMinutes(59);
        SetDateRangeComboSelection(from, to);

        DateRangeFilterOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void DateRangeFilterOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseDateRangeFilter();
    }

    private void DateRangeFilterCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void DateRangeCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseDateRangeFilter();
    }

    private void DateRangeApplyButton_Click(object sender, RoutedEventArgs e)
    {
        DateTime from = ReadDateRangeComboValue(true);
        DateTime to = ReadDateRangeComboValue(false);

        if (from > to)
        {
            ShowStyledMessage(GetLocalizedInvalidDateRangeTitle(), GetLocalizedInvalidDateRangeMessage(), true);
            return;
        }

        _historyFilterFrom = from;
        _historyFilterTo = to;
        CloseDateRangeFilter();
        UpdateHistoryDateRangeButtonText();
        ApplyHistoryFilters();
    }

    private void DateRangeClearButton_Click(object sender, RoutedEventArgs e)
    {
        _historyFilterFrom = null;
        _historyFilterTo = null;
        CloseDateRangeFilter();
        UpdateHistoryDateRangeButtonText();
        ApplyHistoryFilters();
    }

    private void CloseDateRangeFilter()
    {
        DateRangeFilterOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private DateTime ReadDateRangeComboValue(bool isFrom)
    {
        var yearCombo = isFrom ? FromYearComboBox : ToYearComboBox;
        var monthCombo = isFrom ? FromMonthComboBox : ToMonthComboBox;
        var dayCombo = isFrom ? FromDayComboBox : ToDayComboBox;
        var hourCombo = isFrom ? FromHourComboBox : ToHourComboBox;
        var minuteCombo = isFrom ? FromMinuteComboBox : ToMinuteComboBox;

        int year = GetComboInt(yearCombo, _persianCalendar.GetYear(DateTime.Now));
        int month = GetComboInt(monthCombo, _persianCalendar.GetMonth(DateTime.Now));
        int day = GetComboInt(dayCombo, _persianCalendar.GetDayOfMonth(DateTime.Now));
        int hour = GetComboInt(hourCombo, isFrom ? 0 : 23);
        int minute = GetComboInt(minuteCombo, isFrom ? 0 : 59);

        return _persianCalendar.ToDateTime(year, month, day, hour, minute, 0, 0);
    }

    private static int GetComboInt(System.Windows.Controls.ComboBox comboBox, int fallback)
    {
        return comboBox.SelectedItem is int value ? value : fallback;
    }

    private void UpdateHistoryDateRangeButtonText()
    {
        if (HistoryDateRangeButton == null)
            return;

        if (_historyFilterFrom.HasValue && _historyFilterTo.HasValue)
        {
            HistoryDateRangeButton.Content = _localization.GetFormattedString("DateRangeFromTo", FormatPersianDateTime(_historyFilterFrom.Value), FormatPersianDateTime(_historyFilterTo.Value));
            return;
        }

        HistoryDateRangeButton.Content = _localization.GetString("SelectDateRange");
    }

    private string FormatPersianDateTime(DateTime dateTime)
    {
        return $"{_persianCalendar.GetYear(dateTime):0000}/{_persianCalendar.GetMonth(dateTime):00}/{_persianCalendar.GetDayOfMonth(dateTime):00} {dateTime:HH:mm}";
    }

    private string GetLocalizedInvalidDateRangeTitle()
    {
        return _localization.GetString("InvalidRange");
    }

    private string GetLocalizedInvalidDateRangeMessage()
    {
        return _localization.GetString("TheStartDateTimeMustBeBeforeTheEndDateTime");
    }

    private void HistorySearchTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        ApplyHistoryFilters();
    }

    private void ApplyHistoryFilters()
    {
        if (HistoryViewItems == null)
            return;

        HistoryViewItems.Clear();
        int rowNumber = 1;
        foreach (var item in GetFilteredHistoryRecords())
        {
            HistoryViewItems.Add(new HistoryDisplayRow(rowNumber++, item, GetDeviceDisplayName(item.DeviceName)));
        }

        RefreshTtTeckHistoryItems();
        ApplyHistoryFilterMode();
    }

    private IEnumerable<ScanRecord> GetFilteredHistoryRecords()
    {
        return HistoryItems
            .Where(DoesRecordMatchDateFilter)
            .Where(DoesRecordMatchSearch)
            .OrderByDescending(r => r.TimestampLocal);
    }

    private bool DoesRecordMatchDateFilter(ScanRecord item)
    {
        if (_historyFilterFrom.HasValue && item.TimestampLocal < _historyFilterFrom.Value)
            return false;

        if (_historyFilterTo.HasValue && item.TimestampLocal > _historyFilterTo.Value)
            return false;

        return true;
    }

    private bool DoesRecordMatchSearch(ScanRecord item)
    {
        string search = HistorySearchTextBox?.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(search))
            return true;

        return ContainsText(item.Barcode, search)
               || ContainsText(item.DeviceName, search)
               || ContainsText(item.DrugName, search)
               || ContainsText(item.TimeText, search)
               || ContainsText(item.PersianDateText, search);
    }

    private static bool ContainsText(string? value, string search)
    {
        return !string.IsNullOrWhiteSpace(value)
               && value.IndexOf(search, StringComparison.CurrentCultureIgnoreCase) >= 0;
    }

    // ---------- TtTeck Registration Phase 1 ----------

    private void InitializeTtTeckRegistrationControls()
    {
        _isDateRangeInitializing = true;
        try
        {
            TtTeckBirthDayTextBox.Text = string.Empty;
            TtTeckBirthMonthTextBox.Text = string.Empty;
            TtTeckBirthYearTextBox.Text = string.Empty;
            TtTeckRegistrationBirthDateTextBox.Text = string.Empty;
            TtTeckRegistrationTypeComboBox.SelectedIndex = 0;
            UpdateTtTeckRegistrationTypeButtons();
        }
        finally
        {
            _isDateRangeInitializing = false;
        }
    }

    private void TtTeckBirthDatePart_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_isDateRangeInitializing)
            return;

        if (TtTeckBirthYearComboBox == null || TtTeckBirthMonthComboBox == null || TtTeckBirthDayComboBox == null)
            return;

        UpdateDayCombo(TtTeckBirthYearComboBox, TtTeckBirthMonthComboBox, TtTeckBirthDayComboBox);

        if (TtTeckRegistrationBirthDateTextBox != null &&
            TtTeckBirthYearComboBox.SelectedItem is int year &&
            TtTeckBirthMonthComboBox.SelectedItem is int month &&
            TtTeckBirthDayComboBox.SelectedItem is int day)
        {
            TtTeckRegistrationBirthDateTextBox.Text = $"{year:0000}/{month:00}/{day:00}";
        }
    }

    private void OpenTtTeckRegistrationButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn || btn.Tag is not TtTeckHistoryRow row)
            return;

        FormulaRegistrationMode formulaMode = GetFormulaRegistrationModeForRow(row);
        bool forceElectronicPrescription = formulaMode == FormulaRegistrationMode.PrescriptionBased || ShouldOpenAsPrescriptionBasedByGenericCode(row);
        bool forceNonePrescription = formulaMode == FormulaRegistrationMode.NoPrescription;
        _ = OpenTtTeckRegistrationForRowAfterLoginAsync(row, forceNonePrescription, forceElectronicPrescription);
    }

    private async Task OpenTtTeckRegistrationForRowAfterLoginAsync(TtTeckHistoryRow row, bool forceNonePrescription, bool forceElectronicPrescription = false)
    {
        if (row.IsRegistered)
        {
            ShowTtacRegistrationHistoryForBarcode(row.Barcode);
            return;
        }

        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () =>
            {
                OpenTtTeckRegistrationForRow(row, forceNonePrescription, forceElectronicPrescription);
                await Task.CompletedTask;
            };
            _pendingTtacRetryLabel = _localization.GetString("PendingOpenRegistrationForm");
            ShowTtacLoginOverlay();
            return;
        }

        OpenTtTeckRegistrationForRow(row, forceNonePrescription, forceElectronicPrescription);
    }

    private void OpenTtTeckRegistrationForRow(TtTeckHistoryRow row, bool forceNonePrescription, bool forceElectronicPrescription = false)
    {
        if (row.IsRegistered)
        {
            ShowTtacRegistrationHistoryForBarcode(row.Barcode);
            return;
        }

        // اگر فرم ثبت همین لحظه برای همین بارکد باز است (مثلاً کاربر وسط پر کردن فیلدهاست) و
        // دوباره اسکن/باز شدن برای همان قلم درخواست شده، فرم پاک نشود - فقط دوباره روی صفحه
        // بیاید و فوکوس بگیرد. این از پاک شدن فیلدهای نیمه‌تکمیل (کد ملی، تاریخ تولد، موبایل،
        // کپچا) با یک اسکن تکراری یا trigger دوبل اسکنر جلوگیری می‌کند.
        if (TtTeckRegistrationOverlay.Visibility == Visibility.Visible
            && _pendingRegistrationTtTeckRow != null
            && string.Equals(_pendingRegistrationTtTeckRow.Barcode, row.Barcode, StringComparison.OrdinalIgnoreCase))
        {
            MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
            FocusAndSelect(TtTeckRegistrationNationalIdTextBox);
            return;
        }

        _pendingRegistrationTtTeckRow = row;
        _ttacCurrentPrescriptionId = null;
        _ttacCurrentCaptchaId = string.Empty;
        _ttacCurrentNationalId = string.Empty;
        _ttacCurrentBirthDate = string.Empty;
        _ttacCurrentPatientFullName = string.Empty;
        TtTeckRegistrationCaptchaTextBox.Text = string.Empty;
        TtTeckRegistrationCaptchaImage.Source = null;
        TtTeckRegistrationResultText.Text = string.Empty;
        TtacRegistrationLogItems.Clear();
        TtTeckRegistrationProductText.Text = string.IsNullOrWhiteSpace(row.ProductDisplayName) ? row.Barcode : row.ProductDisplayName;
        TtTeckRegistrationBarcodeText.Text = row.Barcode;
        UpdateTtTeckRegistrationProductPhoto(row.Barcode);
        TtTeckRegistrationAmountTextBox.Text = "1";
        TtTeckRegistrationNationalIdTextBox.Text = string.Empty;
        TtTeckRegistrationMobileTextBox.Text = string.Empty;
        TtTeckRegistrationMedicalCouncilTextBox.Text = string.Empty;
        TtTeckBirthDayTextBox.Text = string.Empty;
        TtTeckBirthMonthTextBox.Text = string.Empty;
        TtTeckBirthYearTextBox.Text = string.Empty;
        TtTeckRegistrationBirthDateTextBox.Text = string.Empty;

        if (forceElectronicPrescription)
            TtTeckRegistrationTypeComboBox.SelectedIndex = 1;
        else
            TtTeckRegistrationTypeComboBox.SelectedIndex = 0;

        UpdateTtTeckRegistrationTypeButtons();
        ValidateTtacRegistrationFields();
        UpdateTtacRegistrationStageButtons();
        TtTeckRegistrationOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        TtTeckRegistrationNationalIdTextBox.Focus();
        TtTeckRegistrationNationalIdTextBox.SelectAll();

        // به محض باز شدن فرم، کپچای جدید گرفته می‌شود؛ فوکوس روی فیلد اول می‌ماند تا کاربر با Enter جلو برود.
        Dispatcher.BeginInvoke(new Action(async () =>
        {
            await LoadTtacCaptchaAsync(false);
            FocusAndSelect(TtTeckRegistrationNationalIdTextBox);
        }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

        // اگر این یک قلم شیرخشک است و گوشی‌ای وصل است، همین فرم روی گوشی هم مرحله‌به‌مرحله شروع
        // می‌شود (نگاه کنید به MainWindow.RemoteFormulaEntry.cs).
        StartRemoteFormulaEntryIfPossible(row);
    }

    private FormulaRegistrationMode GetFormulaRegistrationModeForRow(TtTeckHistoryRow row)
    {
        FormulaRegistrationMode mode = GetFormulaRegistrationModeForBarcode(row.Barcode);
        if (mode != FormulaRegistrationMode.Unknown)
            return mode;

        return IsFormulaTextFallback(row.PersianProductName, row.EnglishProductName, row.ProductDisplayName, row.StatusText)
            ? FormulaRegistrationMode.NoPrescription
            : FormulaRegistrationMode.Unknown;
    }

    private FormulaRegistrationMode GetFormulaRegistrationModeForRecord(ScanRecord record)
    {
        FormulaRegistrationMode mode = GetFormulaRegistrationModeForBarcode(record.Barcode);
        if (mode != FormulaRegistrationMode.Unknown)
            return mode;

        return IsFormulaTextFallback(record.DrugName)
            ? FormulaRegistrationMode.NoPrescription
            : FormulaRegistrationMode.Unknown;
    }

    private FormulaRegistrationMode GetFormulaRegistrationModeForBarcode(string barcode)
    {
        EnsureFormulaProductIdsLoaded();
        string? productId = GetProductIdentifierForBarcode(barcode);
        if (string.IsNullOrWhiteSpace(productId))
            return FormulaRegistrationMode.Unknown;

        if (_prescriptionFormulaProductIds.Contains(productId))
            return FormulaRegistrationMode.PrescriptionBased;

        if (_noPrescriptionFormulaProductIds.Contains(productId))
            return FormulaRegistrationMode.NoPrescription;

        return FormulaRegistrationMode.Unknown;
    }

    private bool IsFormulaTextFallback(params string?[] values)
    {
        return values.Any(IsInfantFormulaText);
    }

    private void EnsureFormulaProductIdsLoaded()
    {
        if (_formulaProductIdsLoaded)
            return;

        _formulaProductIdsLoaded = true;
        _noPrescriptionFormulaProductIds.Clear();
        _prescriptionFormulaProductIds.Clear();
        _noPrescriptionFormulaPhotoCodes.Clear();
        _prescriptionFormulaPhotoCodes.Clear();

        LoadFormulaProductIdFile(_noPrescriptionFormulaProductIds, _noPrescriptionFormulaPhotoCodes, "No-Rx-Formula.txt");
        LoadFormulaProductIdFile(_prescriptionFormulaProductIds, _prescriptionFormulaPhotoCodes, "Rx-Formula.txt");
    }

    private void LoadFormulaProductIdFile(HashSet<string> idTarget, Dictionary<string, string> photoCodeTarget, string fileName)
    {
        foreach (string path in GetFormulaFileCandidates(fileName))
        {
            try
            {
                if (!File.Exists(path))
                    continue;

                foreach (string rawLine in File.ReadLines(path))
                {
                    string? productId = ExtractProductIdentifierFromFormulaLine(rawLine);
                    if (string.IsNullOrWhiteSpace(productId))
                        continue;

                    idTarget.Add(productId);

                    string? photoCode = ExtractPhotoGroupCodeFromFormulaLine(rawLine);
                    if (!string.IsNullOrWhiteSpace(photoCode))
                        photoCodeTarget[productId] = photoCode;
                }

                if (idTarget.Count > 0)
                    break;
            }
            catch { }
        }
    }

    private IEnumerable<string> GetFormulaFileCandidates(string fileName)
    {
        yield return Path.Combine(AppContext.BaseDirectory, fileName);
        yield return Path.Combine(AppContext.BaseDirectory, "Config", fileName);
        yield return Path.Combine(Environment.CurrentDirectory, fileName);
        yield return Path.Combine(Environment.CurrentDirectory, "Config", fileName);
    }

    private static string? ExtractProductIdentifierFromFormulaLine(string? line)
    {
        if (string.IsNullOrWhiteSpace(line))
            return null;

        string value = ToEnglishDigits(line.Trim());
        if (value.Contains("شناسه", StringComparison.OrdinalIgnoreCase) && value.Contains("فرآورده", StringComparison.OrdinalIgnoreCase))
            return null;

        string[] tabParts = value.Split('\t', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (tabParts.Length >= 2)
        {
            // ستون دوم همیشه «شناسه فرآورده» است - عمداً آخرین ستون گرفته نمی‌شود چون ممکن است
            // یک ستون سوم اختیاری «کد گروه عکس» هم بعد از آن آمده باشد.
            string? id = NormalizeProductIdentifier(tabParts[1]);
            if (!string.IsNullOrWhiteSpace(id))
                return id;
        }

        var endMatch = Regex.Match(value, @"(?<!\d)(\d{10,20})\s*$");
        return endMatch.Success ? NormalizeProductIdentifier(endMatch.Groups[1].Value) : null;
    }

    /// <summary>
    /// ستون سوم اختیاری «کد گروه عکس» را از یک خط Rx-Formula.txt / No-Rx-Formula.txt می‌خواند.
    /// اگر این ستون وجود نداشته باشد یا خالی باشد، null برمی‌گرداند (یعنی برای این قلم عکسی نیست).
    /// </summary>
    private static string? ExtractPhotoGroupCodeFromFormulaLine(string? line)
    {
        if (string.IsNullOrWhiteSpace(line))
            return null;

        string value = ToEnglishDigits(line.Trim());
        if (value.Contains("شناسه", StringComparison.OrdinalIgnoreCase) && value.Contains("فرآورده", StringComparison.OrdinalIgnoreCase))
            return null;

        string[] tabParts = value.Split('\t', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (tabParts.Length < 3)
            return null;

        string code = tabParts[2].Trim();
        return string.IsNullOrWhiteSpace(code) ? null : code;
    }

    private string? GetProductIdentifierForBarcode(string barcode)
    {
        if (!_ttTeckDetailsByBarcode.TryGetValue(barcode, out var info))
            return null;

        string? fromIrc = NormalizeProductIdentifier(info.IRC);
        if (!string.IsNullOrWhiteSpace(fromIrc))
            return fromIrc;

        if (info.ExtraFields == null)
            return null;

        foreach (var field in info.ExtraFields)
        {
            string normalizedKey = NormalizeProductIdentifierLabel(field.Key);
            if (normalizedKey == "irc" || normalizedKey.Contains("شناسهفرآورده") || normalizedKey.Contains("شناسهفراورده"))
            {
                string? value = NormalizeProductIdentifier(field.Value);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
        }

        return null;
    }

    /// <summary>
    /// اگر بارکد داده‌شده متعلق به یک قلم شیرخشک باشد که برایش عکس ثبت شده، مسیر فایل عکس
    /// (JPG با پس‌زمینه‌ی سفید یکدست، همانی که پیش‌تر آماده شد) را برمی‌گرداند؛ در غیر این صورت
    /// null (که یعنی چیزی نمایش داده نشود - نه یک آیکن جایگزین).
    /// </summary>
    private string? GetFormulaPhotoPathForBarcode(string barcode)
    {
        EnsureFormulaProductIdsLoaded();
        string? productId = GetProductIdentifierForBarcode(barcode);
        if (string.IsNullOrWhiteSpace(productId))
            return null;

        if (_prescriptionFormulaPhotoCodes.TryGetValue(productId, out var rxCode))
        {
            string? rxPath = FindFormulaPhotoFile("Rx", rxCode);
            if (rxPath != null)
                return rxPath;
        }

        if (_noPrescriptionFormulaPhotoCodes.TryGetValue(productId, out var noRxCode))
        {
            string? noRxPath = FindFormulaPhotoFile("NoRx", noRxCode);
            if (noRxPath != null)
                return noRxPath;
        }

        return null;
    }

    private static string? FindFormulaPhotoFile(string subFolder, string code)
    {
        if (string.IsNullOrWhiteSpace(code))
            return null;

        foreach (string baseDir in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            try
            {
                string path = Path.Combine(baseDir, "FormulaPhotos", subFolder, $"{code}.jpg");
                if (File.Exists(path))
                    return path;
            }
            catch { }
        }

        return null;
    }

    /// <summary>
    /// عکس شیرخشک قلم انتخاب‌شده در پنل ثبت تی‌تک را نشان می‌دهد؛ اگر عکسی برای این قلم نباشد،
    /// کادر عکس کاملاً مخفی می‌شود (طبق تصمیم: بدون آیکن/جای‌خالی جایگزین).
    /// </summary>
    private void UpdateTtTeckRegistrationProductPhoto(string barcode)
    {
        string? path = GetFormulaPhotoPathForBarcode(barcode);
        if (string.IsNullOrWhiteSpace(path))
        {
            TtTeckRegistrationProductPhoto.Source = null;
            TtTeckRegistrationProductPhotoBorder.Visibility = Visibility.Collapsed;
            return;
        }

        try
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.UriSource = new Uri(path, UriKind.Absolute);
            bitmap.EndInit();
            bitmap.Freeze();
            TtTeckRegistrationProductPhoto.Source = bitmap;
            TtTeckRegistrationProductPhotoBorder.Visibility = Visibility.Visible;
        }
        catch
        {
            TtTeckRegistrationProductPhoto.Source = null;
            TtTeckRegistrationProductPhotoBorder.Visibility = Visibility.Collapsed;
        }
    }

    /// <summary>
    /// عکس شیرخشک را داخل یک Image/Border مشخص بارگذاری می‌کند؛ برای استفاده مشترک در
    /// دیالوگ «جزئیات محصول» و دیالوگ «تاریخچه ثبت این محصول». اگر عکسی نباشد، کادر
    /// کاملاً مخفی می‌شود (بدون آیکن/جای‌خالی جایگزین، مطابق تصمیم قبلی).
    /// </summary>
    private static void SetOverlayProductPhoto(System.Windows.Controls.Image image, Border border, string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            image.Source = null;
            border.Visibility = Visibility.Collapsed;
            return;
        }

        try
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.UriSource = new Uri(path, UriKind.Absolute);
            bitmap.EndInit();
            bitmap.Freeze();
            image.Source = bitmap;
            border.Visibility = Visibility.Visible;
        }
        catch
        {
            image.Source = null;
            border.Visibility = Visibility.Collapsed;
        }
    }

    private static string? NormalizeProductIdentifier(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        string digits = new string(ToEnglishDigits(value).Where(char.IsDigit).ToArray());
        return string.IsNullOrWhiteSpace(digits) ? null : digits;
    }

    private static string NormalizeProductIdentifierLabel(string? label)
    {
        return (label ?? string.Empty)
            .Replace(" ", "")
            .Replace("_", "")
            .Replace("-", "")
            .Replace("ي", "ی")
            .Replace("ك", "ک")
            .Trim()
            .ToLowerInvariant();
    }

    private bool ShouldOpenAsPrescriptionBasedByGenericCode(TtTeckHistoryRow row)
    {
        EnsureSpecialPrescriptionGenericCodesLoaded();
        if (_specialPrescriptionGenericCodes.Count == 0)
            return false;

        string? genericCode = GetGenericCodeForTtTeckRow(row);
        if (string.IsNullOrWhiteSpace(genericCode))
            return false;

        return _specialPrescriptionGenericCodes.Contains(genericCode);
    }

    private void EnsureSpecialPrescriptionGenericCodesLoaded()
    {
        if (_specialPrescriptionGenericCodesLoaded)
            return;

        _specialPrescriptionGenericCodesLoaded = true;
        _specialPrescriptionGenericCodes.Clear();

        foreach (string path in GetSpecialPrescriptionGenericCodeFileCandidates())
        {
            try
            {
                if (!File.Exists(path))
                    continue;

                foreach (string rawLine in File.ReadLines(path))
                {
                    string? code = ExtractGenericCodeFromSpecialPrescriptionLine(rawLine);
                    if (!string.IsNullOrWhiteSpace(code))
                        _specialPrescriptionGenericCodes.Add(code);
                }

                if (_specialPrescriptionGenericCodes.Count > 0)
                    break;
            }
            catch { }
        }
    }

    private IEnumerable<string> GetSpecialPrescriptionGenericCodeFileCandidates()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "special-prescription-generics.txt");
        yield return Path.Combine(AppContext.BaseDirectory, "Config", "special-prescription-generics.txt");
        yield return Path.Combine(Environment.CurrentDirectory, "special-prescription-generics.txt");
        yield return Path.Combine(Environment.CurrentDirectory, "Config", "special-prescription-generics.txt");
    }

    private static string? ExtractGenericCodeFromSpecialPrescriptionLine(string? line)
    {
        if (string.IsNullOrWhiteSpace(line))
            return null;

        string value = ToEnglishDigits(line.Trim());
        if (value.Contains("کد", StringComparison.OrdinalIgnoreCase) && value.Contains("ژنریک", StringComparison.OrdinalIgnoreCase))
            return null;
        if (value.Contains("GenericCode", StringComparison.OrdinalIgnoreCase) || value.Contains("Generic Code", StringComparison.OrdinalIgnoreCase))
            return null;

        string[] tabParts = value.Split('\t', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (tabParts.Length >= 2)
        {
            string lastPart = tabParts[^1];
            var codeMatch = Regex.Match(lastPart, @"\d+");
            if (codeMatch.Success)
                return codeMatch.Value.TrimStart('0').Length == 0 ? "0" : codeMatch.Value.TrimStart('0');
        }

        var endMatch = Regex.Match(value, @"(?<!\d)(\d{1,10})\s*$");
        if (endMatch.Success)
        {
            string code = endMatch.Groups[1].Value;
            return code.TrimStart('0').Length == 0 ? "0" : code.TrimStart('0');
        }

        return null;
    }

    private string? GetGenericCodeForTtTeckRow(TtTeckHistoryRow row)
    {
        if (!_ttTeckDetailsByBarcode.TryGetValue(row.Barcode, out var info) || info.ExtraFields == null)
            return null;

        foreach (var field in info.ExtraFields)
        {
            string label = NormalizeGenericCodeLabel(field.Key);
            if (label == "genericcode" || (label.Contains("generic") && label.Contains("code")) || (label.Contains("ژنریک") && label.Contains("کد")))
            {
                string value = ToEnglishDigits(field.Value ?? string.Empty);
                var match = Regex.Match(value, @"\d+");
                if (match.Success)
                {
                    string code = match.Value.TrimStart('0');
                    return code.Length == 0 ? "0" : code;
                }
            }
        }

        return null;
    }

    private static string NormalizeGenericCodeLabel(string? label)
    {
        return (label ?? string.Empty)
            .Replace(" ", "")
            .Replace("_", "")
            .Replace("-", "")
            .Replace("ي", "ی")
            .Replace("ك", "ک")
            .Trim()
            .ToLowerInvariant();
    }

    private string GetLocalizedProductDisplayName(string? persianName, string? englishName)
    {
        if (_localization.CurrentLanguage == AppLanguage.English)
        {
            return !string.IsNullOrWhiteSpace(englishName) ? englishName : (persianName ?? string.Empty);
        }

        return !string.IsNullOrWhiteSpace(persianName) ? persianName : (englishName ?? string.Empty);
    }

    private TtTeckHistoryRow CreateTtTeckHistoryRowFromRecord(ScanRecord item)
    {
        var (persianProductName, englishProductName) = SplitTtTeckProductNames(item.DrugName);
        bool hasCachedDetails = _ttTeckDetailsByBarcode.TryGetValue(item.Barcode, out var cachedDetails);
        if (hasCachedDetails)
        {
            persianProductName = cachedDetails?.PersianName ?? persianProductName;
            englishProductName = cachedDetails?.EnglishName ?? englishProductName;
        }

        bool isPending = !hasCachedDetails && IsTtTeckLookupPending(item.DrugName);
        bool isFailed = !hasCachedDetails && !isPending && IsTtTeckLookupFailed(item.DrugName);

        if (string.IsNullOrWhiteSpace(persianProductName) || isFailed || isPending)
        {
            if (isPending)
            {
                persianProductName = GetTtTeckLookupPendingText();
                if (_localization.CurrentLanguage == AppLanguage.English && string.IsNullOrWhiteSpace(englishProductName))
                    englishProductName = GetTtTeckLookupPendingText();
            }
            else
            {
                persianProductName = $"بارکد: {item.Barcode}";
                if (_localization.CurrentLanguage == AppLanguage.English && string.IsNullOrWhiteSpace(englishProductName))
                    englishProductName = $"Barcode: {item.Barcode}";
            }
        }

        bool isRegisteredInTtac = IsBarcodeTtacRegistered(item.Barcode);

        return new TtTeckHistoryRow
        {
            RowNumber = TtTeckHistoryItems.Count + 1,
            IsRegistered = isRegisteredInTtac,
            RegistrationButtonText = GetRegistrationButtonText(isRegisteredInTtac),
            RegistrationButtonBackground = GetRegistrationButtonBrush(isRegisteredInTtac),
            ProductDisplayName = GetLocalizedProductDisplayName(persianProductName, englishProductName),
            PersianProductName = persianProductName,
            EnglishProductName = englishProductName,
            Barcode = item.Barcode,
            TimeText = item.TimeText,
            PersianDateText = item.PersianDateText,
            TimestampLocal = item.TimestampLocal,
            DeviceName = GetDeviceDisplayName(item.DeviceName),
            StatusText = cachedDetails?.Message ?? item.DrugName,
            RetryReason = string.IsNullOrWhiteSpace(item.DrugName) ? GetLocalizedUnknownLookupReason() : item.DrugName,
            RetryButtonVisibility = isFailed ? Visibility.Visible : Visibility.Collapsed,
            IsInfantFormula = IsInfantFormulaRecord(item)
        };
    }

    private bool IsInfantFormulaRecord(ScanRecord item)
    {
        if (GetFormulaRegistrationModeForRecord(item) != FormulaRegistrationMode.Unknown)
            return true;

        string text = item.DrugName ?? string.Empty;
        if (_ttTeckDetailsByBarcode.TryGetValue(item.Barcode, out var info))
        {
            text += " " + info.PersianName + " " + info.EnglishName + " " + info.Message;
            if (info.ExtraFields != null)
                text += " " + string.Join(" ", info.ExtraFields.Values);
        }

        return IsInfantFormulaText(text);
    }

    private static bool IsInfantFormulaText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        string value = text.Trim();
        string normalized = value
            .Replace("ي", "ی")
            .Replace("ك", "ک")
            .Replace("‌", "")
            .Replace(" ", "")
            .ToLowerInvariant();

        if (normalized.Contains("شیرخشک") || normalized.Contains("شيرخشک") || normalized.Contains("شيرخشك"))
            return true;

        string lower = value.ToLowerInvariant();
        return lower.Contains("infant formula")
               || lower.Contains("baby formula")
               || lower.Contains("milk powder")
               || lower.Contains("formula milk")
               || lower.Contains("follow-on formula")
               || lower.Contains("follow on formula");
    }

    private void TryAutoOpenInfantFormulaRegistration(ScanRecord record)
    {
        if (!_autoOpenInfantFormulaRegistration)
            return;

        FormulaRegistrationMode formulaMode = GetFormulaRegistrationModeForRecord(record);
        if (formulaMode == FormulaRegistrationMode.Unknown)
            return;

        // اگر کاربر روی دیالوگ موفقیتِ گوشی دکمه‌ی «ثبت مجدد» را زده بود، همین اسکن (اولین اسکنِ
        // بعد از آن) همان درخواستی است که منتظرش بودیم - این پرچم یک‌بارمصرف همین‌جا مصرف می‌شود،
        // چه بارکد قبلاً ثبت شده باشد (شاخه‌ی زیر) چه فرم واقعاً باز شود (پایین‌تر).
        bool repeatArmed = _remoteEntryRepeatArmed;
        _remoteEntryRepeatArmed = false;

        // اگر این بارکد قبلاً در همین داروخانه ثبت شده، به‌جای فرم ثبت، تاریخچه‌ی ثبت همان محصول
        // را خودکار نشان بده تا کاربر نتواند دوباره ثبت کند.
        if (IsBarcodeTtacRegistered(record.Barcode))
        {
            Dispatcher.BeginInvoke(new Action(() => ShowTtacRegistrationHistoryForBarcode(record.Barcode)));

            // چون در این حالت هیچ‌وقت فرم ثبت باز نمی‌شود، ویزارد «ورود از راه دور» هم هیچ‌وقت
            // روی گوشی شروع نمی‌شود - پس اگر گوشی وصل است، به‌جایش فقط عکس محصول و یک پیام کوتاه
            // «قبلاً ثبت شده» روی گوشی نشان می‌دهیم (همان مکانیزم BroadcastAlert موجود).
            _service?.BroadcastAlert(
                _localization.GetString("AlreadyRegisteredFormulaTitleForPhone"),
                _localization.GetString("ThisProductWasAlreadyRegisteredForThisPharmacyForPhone"),
                true,
                GetFormulaPhotoPathForBarcode(record.Barcode));
            return;
        }

        // کلید بر اساس «بارکد + داروخانه‌ی فعلی» است (نه فقط بارکد) - هم به این دلیل که هر اسکن
        // دوباره‌ی همان جعبه (مثلاً دبل‌تریگر اسکنر یا اسکن دوباره چند ثانیه بعد) نباید فرم را
        // دوباره باز کند، و هم برای این‌که با تغییر داروخانه (خروج از تی‌تک و ورود به داروخانه‌ی
        // دیگر) همین بارکد دوباره بتواند فرم را باز کند - چون در داروخانه‌ی جدید وضعیت ثبتش کاملاً
        // جداست. این کلید مستقل از تشخیص «تغییر داروخانه» در ApplyTtacAccessToken کار می‌کند، پس
        // حتی اگر آن تشخیص به هر دلیلی درست عمل نکند، رفتار درست همچنان تضمین می‌شود. این کلید
        // موقع بستن فرم ثبت (انصراف یا ثبت موفق) پاک می‌شود تا اسکن دوباره‌ی همان بارکد دوباره
        // فرم را باز کند.
        string key = GetReceiveStatusStorageKey() + "|" + record.Barcode;
        if (!_autoOpenedFormulaRegistrationKeys.Add(key))
            return;

        Dispatcher.BeginInvoke(new Action(async () =>
        {
            var row = CreateTtTeckHistoryRowFromRecord(record);
            bool forceElectronicPrescription = formulaMode == FormulaRegistrationMode.PrescriptionBased;
            bool forceNonePrescription = formulaMode == FormulaRegistrationMode.NoPrescription;
            await OpenTtTeckRegistrationForRowAfterLoginAsync(row, forceNonePrescription, forceElectronicPrescription);

            // اگر ورود تی‌تک لازم بود (توکن معتبر نبود) و کاربر پاپ‌آپ ورود را بدون تکمیل بست،
            // OpenTtTeckRegistrationForRowAfterLoginAsync برمی‌گردد بدون اینکه واقعاً فرم ثبت را
            // باز کرده باشد. قبلاً کلید بالا (_autoOpenedFormulaRegistrationKeys.Add) در همین حالت
            // هم می‌ماند - یعنی اسکن دوباره‌ی همین بارکد تا آخر نشست دیگر هیچ‌وقت خودکار فرم را
            // باز نمی‌کرد، چون هیچ نقطه‌ای (نه CloseTtTeckRegistrationOverlay، چون اصلاً باز نشد)
            // آن را پاک نمی‌کرد (باگ گزارش ممیزی). اگر واقعاً باز نشده، همین‌جا پاک می‌شود تا اسکن
            // بعدی دوباره تلاش کند.
            bool actuallyOpened = TtTeckRegistrationOverlay.Visibility == Visibility.Visible
                && _pendingRegistrationTtTeckRow != null
                && string.Equals(_pendingRegistrationTtTeckRow.Barcode, record.Barcode, StringComparison.OrdinalIgnoreCase);
            if (!actuallyOpened)
            {
                _autoOpenedFormulaRegistrationKeys.Remove(key);
            }
            else if (repeatArmed && _lastFormulaRepeatContext != null)
            {
                // «ثبت مجدد» از روی گوشی: کادرها را با اطلاعات همان ثبتِ قبلی (کد ملی/تاریخ
                // تولد/موبایل/شماره نظام) پر کن - دقیقاً مثل مسیر مشابه دسکتاپیِ
                // OpenRepeatFormulaRegistrationForBarcodeAsync - و یک کپچای تازه بگیر. علامت‌زدن
                // _remoteEntryWaitingForCaptcha کافی است: NotifyRemoteEntryCaptchaLoaded (که
                // LoadTtacCaptchaAsync بعد از لود موفق کپچا صدا می‌زند) خودش مرحله‌ی کپچا را روی
                // گوشی نشان می‌دهد - دیگر لازم نیست کاربر مراحل کد ملی/تاریخ تولد/موبایل را که
                // خالی می‌شد یکی‌یکی رد کند.
                ApplyFormulaRepeatContextToOpenForm(_lastFormulaRepeatContext);
                if (IsRemoteFormulaEntryActiveFor(record.Barcode))
                    _remoteEntryWaitingForCaptcha = true;
                await LoadTtacCaptchaAsync(true);
                FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
            }
        }));
    }

    private void TtTeckRegistrationOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseTtTeckRegistrationOverlay();
    }

    private void TtTeckRegistrationCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void TtTeckRegistrationCloseButton_Click(object sender, RoutedEventArgs e)
    {
        CloseTtTeckRegistrationOverlay();
    }

    private void CloseTtTeckRegistrationOverlay()
    {
        // اگر فرم ثبتِ شیرخشک با اسکن خودکار باز شده بود و کاربر بدون ثبت بستش، کلید کشِ همان
        // بارکد را پاک کن تا اسکن دوباره‌ی همان بارکد دوباره فرم را باز کند. (بعد از ثبت موفق هم
        // پاک کردن بی‌ضرر است چون همان بارکد حالا «ثبت‌شده» است و به‌جای فرم، تاریخچه‌ی ثبت آن
        // نشان داده می‌شود.)
        if (_pendingRegistrationTtTeckRow != null)
        {
            string key = GetReceiveStatusStorageKey() + "|" + _pendingRegistrationTtTeckRow.Barcode;
            _autoOpenedFormulaRegistrationKeys.Remove(key);
        }

        // اگر «ورود اطلاعات از راه دور» برای همین فرم فعال بود و همین‌جا (نه بعد از ثبت
        // موفق/ناموفق که خودش با EndRemoteFormulaEntry(notifyPhone:false) این حالت را از قبل پاک
        // کرده) بسته می‌شود، یعنی کاربر خودش فرم را بست - به گوشی خبر بده تا ویزارد پاک شود.
        EndRemoteFormulaEntry();

        TtTeckRegistrationOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void UpdateTtacRegistrationStageButtons(bool isBusy = false)
    {
        if (_isInitializingUi ||
            TtTeckRegistrationCreatePrescriptionButton == null ||
            TtTeckRegistrationOpenWebButton == null ||
            TtTeckRegistrationGetCaptchaButton == null ||
            TtTeckRegistrationSubmitItemButton == null ||
            TtTeckRegistrationSaveDraftButton == null ||
            TtTeckRegistrationCancelButton == null)
        {
            return;
        }

        bool hasPrescription = _ttacCurrentPrescriptionId.HasValue;

        TtTeckRegistrationOpenWebButton.IsEnabled = !isBusy;
        bool fieldsOk = !hasPrescription && ValidateTtacRegistrationFields();
        TtTeckRegistrationGetCaptchaButton.IsEnabled = !isBusy && !hasPrescription;
        TtTeckRegistrationCreatePrescriptionButton.IsEnabled = !isBusy && !hasPrescription && fieldsOk;
        TtTeckRegistrationSubmitItemButton.IsEnabled = !isBusy && hasPrescription;
        TtTeckRegistrationSaveDraftButton.IsEnabled = !isBusy && hasPrescription;
        TtTeckRegistrationCancelButton.IsEnabled = !isBusy;
    }

    private void TtTeckRegistrationTypeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (TtTeckRegistrationMedicalCouncilPanel == null)
            return;

        bool isElectronic = TtTeckRegistrationTypeComboBox?.SelectedIndex == 1;
        TtTeckRegistrationMedicalCouncilPanel.Visibility = isElectronic ? Visibility.Visible : Visibility.Collapsed;
        UpdateTtTeckRegistrationTypeButtons();
    }

    private void TtTeckRegistrationNonePrescriptionButton_Click(object sender, RoutedEventArgs e)
    {
        bool changed = TtTeckRegistrationTypeComboBox.SelectedIndex != 0;
        TtTeckRegistrationTypeComboBox.SelectedIndex = 0;
        if (changed)
            ResetTtacRegistrationAfterTypeChange(focusMedicalCouncilAfterLoad: false);
        UpdateTtTeckRegistrationTypeButtons();
        // بعد از تغییر به فاقد نسخه، مستقیم روی کپچا برو.
        FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
    }

    private void TtTeckRegistrationElectronicPrescriptionButton_Click(object sender, RoutedEventArgs e)
    {
        bool changed = TtTeckRegistrationTypeComboBox.SelectedIndex != 1;
        TtTeckRegistrationTypeComboBox.SelectedIndex = 1;
        if (changed)
            ResetTtacRegistrationAfterTypeChange(focusMedicalCouncilAfterLoad: true);
        UpdateTtTeckRegistrationTypeButtons();
        // در نسخه‌محور بعد از تغییر نوع، مستقیم روی شماره نظام پزشکی برو.
        FocusAndSelect(TtTeckRegistrationMedicalCouncilTextBox);
    }

    private void ResetTtacRegistrationAfterTypeChange(bool focusMedicalCouncilAfterLoad)
    {
        // اگر کاربر بعد از خطای سامانه تشخیص داد نوع ثبت را اشتباه زده،
        // بدون بستن فرم، نسخه قبلی و کپچای مصرف‌شده پاک می‌شود تا بتواند مسیر جدید را شروع کند.
        _ttacCurrentPrescriptionId = null;
        _ttacCurrentCaptchaId = string.Empty;
        _ttacCurrentIsElectronic = TtTeckRegistrationTypeComboBox.SelectedIndex == 1;
        TtTeckRegistrationCaptchaTextBox.Text = string.Empty;
        TtTeckRegistrationCaptchaImage.Source = null;
        TtTeckRegistrationResultText.Text = _localization.GetString("RegistrationTypeChangedGettingANewCaptcha");
        UpdateTtacRegistrationStageButtons();

        // کپچای جدید خودکار گرفته می‌شود؛ بعد از آن فوکوس مطابق نوع ثبت تنظیم می‌شود.
        Dispatcher.BeginInvoke(new Action(async () =>
        {
            await LoadTtacCaptchaAsync(false);
            if (focusMedicalCouncilAfterLoad)
                FocusAndSelect(TtTeckRegistrationMedicalCouncilTextBox);
            else
                FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
        }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
    }

    private void UpdateTtTeckRegistrationTypeButtons()
    {
        if (_isInitializingUi ||
            TtTeckRegistrationNonePrescriptionButton == null ||
            TtTeckRegistrationElectronicPrescriptionButton == null ||
            TtTeckRegistrationTypeComboBox == null)
            return;

        bool isElectronic = TtTeckRegistrationTypeComboBox?.SelectedIndex == 1;
        TtTeckRegistrationNonePrescriptionButton.Background = new SolidColorBrush(isElectronic
            ? System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF)
            : System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81));
        TtTeckRegistrationElectronicPrescriptionButton.Background = new SolidColorBrush(isElectronic
            ? System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81)
            : System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF));
    }

    private void FocusAndSelect(System.Windows.Controls.Control control)
    {
        control.Focus();
        switch (control)
        {
            case System.Windows.Controls.TextBox textBox:
                textBox.SelectAll();
                break;
            case PasswordBox passwordBox:
                passwordBox.SelectAll();
                break;
        }
    }

    private void TtTeckBirthDateBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.TextBox textBox)
            FocusAndSelect(textBox);
    }

    private void TryAutoAdvanceTtacDateField(object sender)
    {
        if (_isAutoAdvancingTtacField)
            return;

        try
        {
            _isAutoAdvancingTtacField = true;

            if (sender == TtTeckBirthDayTextBox)
            {
                string value = ToEnglishDigits(TtTeckBirthDayTextBox.Text.Trim());
                if (int.TryParse(value, out int day) && day >= 1 && day <= 31 && (value.Length >= 2 || day >= 4))
                    FocusAndSelect(TtTeckBirthMonthTextBox);
            }
            else if (sender == TtTeckBirthMonthTextBox)
            {
                string value = ToEnglishDigits(TtTeckBirthMonthTextBox.Text.Trim());
                if (int.TryParse(value, out int month) && month >= 1 && month <= 12 && (value.Length >= 2 || month >= 2))
                    FocusAndSelect(TtTeckBirthYearTextBox);
            }
            else if (sender == TtTeckBirthYearTextBox)
            {
                string value = ToEnglishDigits(TtTeckBirthYearTextBox.Text.Trim());
                if (value.Length >= 4 && int.TryParse(value, out int year) && year >= 1200 && year <= 1500)
                {
                    if (TtTeckRegistrationTypeComboBox.SelectedIndex == 1)
                        FocusAndSelect(TtTeckRegistrationMedicalCouncilTextBox);
                    else
                        FocusAndSelect(TtTeckRegistrationMobileTextBox);
                }
            }
        }
        finally
        {
            _isAutoAdvancingTtacField = false;
        }
    }

    private void TtacRegistrationField_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (_isInitializingUi)
            return;

        ValidateTtacRegistrationFields();
        UpdateTtacRegistrationStageButtons();
        TryAutoAdvanceNationalIdField(sender);
        TryAutoAdvanceTtacDateField(sender);
        TryAutoAdvanceTtacMobileField(sender);
    }

    private void TryAutoAdvanceNationalIdField(object sender)
    {
        if (_isAutoAdvancingTtacField || sender != TtTeckRegistrationNationalIdTextBox)
            return;

        string nationalId = ToEnglishDigits(TtTeckRegistrationNationalIdTextBox.Text.Trim());
        if (nationalId.Length == 10 && nationalId.All(char.IsDigit))
        {
            try
            {
                _isAutoAdvancingTtacField = true;
                TtTeckRegistrationNationalIdTextBox.Text = nationalId;
                TtTeckRegistrationNationalIdTextBox.CaretIndex = TtTeckRegistrationNationalIdTextBox.Text.Length;
                FocusAndSelect(TtTeckBirthDayTextBox);
            }
            finally
            {
                _isAutoAdvancingTtacField = false;
            }
        }
    }

    private void TryAutoAdvanceTtacMobileField(object sender)
    {
        if (_isAutoAdvancingTtacField || sender != TtTeckRegistrationMobileTextBox)
            return;

        string mobile = ToEnglishDigits(TtTeckRegistrationMobileTextBox.Text.Trim());
        if (mobile.Length == 11 && mobile.All(char.IsDigit) && mobile.StartsWith("09"))
        {
            try
            {
                _isAutoAdvancingTtacField = true;
                TtTeckRegistrationMobileTextBox.Text = mobile;
                TtTeckRegistrationMobileTextBox.CaretIndex = TtTeckRegistrationMobileTextBox.Text.Length;
                FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
            }
            finally
            {
                _isAutoAdvancingTtacField = false;
            }
        }
    }

    private bool ValidateTtacRegistrationFields()
    {
        if (_isInitializingUi ||
            TtTeckRegistrationAmountTextBox == null ||
            TtTeckRegistrationNationalIdTextBox == null ||
            TtTeckBirthDayTextBox == null ||
            TtTeckBirthMonthTextBox == null ||
            TtTeckBirthYearTextBox == null ||
            TtTeckRegistrationCaptchaTextBox == null ||
            TtTeckRegistrationMedicalCouncilTextBox == null ||
            TtTeckRegistrationMobileTextBox == null ||
            TtTeckRegistrationTypeComboBox == null)
        {
            return false;
        }

        bool amountOk = int.TryParse(ToEnglishDigits(TtTeckRegistrationAmountTextBox.Text.Trim()), out int amount) && amount > 0;
        string nationalId = ToEnglishDigits(TtTeckRegistrationNationalIdTextBox.Text.Trim());
        bool nationalOk = nationalId.Length == 10 && nationalId.All(char.IsDigit);
        bool dayOk = int.TryParse(ToEnglishDigits(TtTeckBirthDayTextBox.Text.Trim()), out int day) && day is >= 1 and <= 31;
        bool monthOk = int.TryParse(ToEnglishDigits(TtTeckBirthMonthTextBox.Text.Trim()), out int month) && month is >= 1 and <= 12;
        bool yearOk = int.TryParse(ToEnglishDigits(TtTeckBirthYearTextBox.Text.Trim()), out int year) && year is >= 1200 and <= 1500;
        bool captchaOk = !string.IsNullOrWhiteSpace(TtTeckRegistrationCaptchaTextBox.Text) && !string.IsNullOrWhiteSpace(_ttacCurrentCaptchaId);
        bool isElectronic = TtTeckRegistrationTypeComboBox.SelectedIndex == 1;
        bool medicalOk = !isElectronic || !string.IsNullOrWhiteSpace(TtTeckRegistrationMedicalCouncilTextBox.Text.Trim());

        string mobile = ToEnglishDigits(TtTeckRegistrationMobileTextBox.Text.Trim());
        bool mobileOk = string.IsNullOrWhiteSpace(mobile) || (mobile.Length == 11 && mobile.All(char.IsDigit) && mobile.StartsWith("09"));

        SetValidationBorder(TtTeckRegistrationAmountTextBox, amountOk);
        SetValidationBorder(TtTeckRegistrationNationalIdTextBox, nationalOk);
        SetValidationBorder(TtTeckBirthDayTextBox, dayOk);
        SetValidationBorder(TtTeckBirthMonthTextBox, monthOk);
        SetValidationBorder(TtTeckBirthYearTextBox, yearOk);
        SetValidationBorder(TtTeckRegistrationCaptchaTextBox, captchaOk);
        SetValidationBorder(TtTeckRegistrationMedicalCouncilTextBox, medicalOk);
        SetValidationBorder(TtTeckRegistrationMobileTextBox, mobileOk, string.IsNullOrWhiteSpace(mobile));

        return amountOk && nationalOk && dayOk && monthOk && yearOk && captchaOk && medicalOk && mobileOk;
    }

    private void SetValidationBorder(System.Windows.Controls.TextBox textBox, bool isValid, bool isNeutral = false)
    {
        textBox.BorderThickness = new Thickness(2);
        textBox.BorderBrush = isNeutral
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xD1, 0xD5, 0xDB))
            : isValid
                ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81))
                : new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xEF, 0x44, 0x44));
    }

    private void TtTeckRegistrationAmountTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        if (string.IsNullOrWhiteSpace(TtTeckRegistrationAmountTextBox.Text))
            return;
        FocusAndSelect(TtTeckRegistrationNationalIdTextBox);
    }

    private void TtTeckRegistrationNationalIdTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        string nationalId = ToEnglishDigits(TtTeckRegistrationNationalIdTextBox.Text.Trim());
        if (nationalId.Length != 10 || !nationalId.All(char.IsDigit))
        {
            TtTeckRegistrationResultText.Text = _localization.GetString("NationalIDMustBe10Digits");
            return;
        }
        TtTeckRegistrationNationalIdTextBox.Text = nationalId;
        FocusAndSelect(TtTeckBirthDayTextBox);
    }

    private void TtTeckRegistrationMobileTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        string mobile = ToEnglishDigits(TtTeckRegistrationMobileTextBox.Text.Trim());
        if (!string.IsNullOrWhiteSpace(mobile) && (!mobile.All(char.IsDigit) || mobile.Length < 10))
        {
            TtTeckRegistrationResultText.Text = _localization.GetString("MobileNumberIsNotValid");
            return;
        }
        TtTeckRegistrationMobileTextBox.Text = mobile;
        FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
    }

    private void TtTeckBirthDayTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        if (!int.TryParse(ToEnglishDigits(TtTeckBirthDayTextBox.Text.Trim()), out int day) || day < 1 || day > 31)
            return;
        TtTeckBirthDayTextBox.Text = day.ToString();
        FocusAndSelect(TtTeckBirthMonthTextBox);
    }

    private void TtTeckBirthMonthTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        if (!int.TryParse(ToEnglishDigits(TtTeckBirthMonthTextBox.Text.Trim()), out int month) || month < 1 || month > 12)
            return;
        TtTeckBirthMonthTextBox.Text = month.ToString();
        FocusAndSelect(TtTeckBirthYearTextBox);
    }

    private void TtTeckBirthYearTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        if (!int.TryParse(ToEnglishDigits(TtTeckBirthYearTextBox.Text.Trim()), out int year) || year < 1200 || year > 1500)
            return;
        TtTeckBirthYearTextBox.Text = year.ToString();

        if (TtTeckRegistrationTypeComboBox.SelectedIndex == 1)
            FocusAndSelect(TtTeckRegistrationMedicalCouncilTextBox);
        else
            FocusAndSelect(TtTeckRegistrationMobileTextBox);
    }

    private void TtTeckRegistrationMedicalCouncilTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        if (string.IsNullOrWhiteSpace(TtTeckRegistrationMedicalCouncilTextBox.Text))
            return;
        FocusAndSelect(TtTeckRegistrationMobileTextBox);
    }

    private void TtTeckRegistrationCaptchaTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        if (string.IsNullOrWhiteSpace(TtTeckRegistrationCaptchaTextBox.Text))
            return;
        if (TtTeckRegistrationCreatePrescriptionButton.IsEnabled)
            TtTeckRegistrationCreatePrescriptionButton_Click(TtTeckRegistrationCreatePrescriptionButton, new RoutedEventArgs());
        else if (TtTeckRegistrationSubmitItemButton.IsEnabled)
            TtTeckRegistrationSubmitItemButton_Click(TtTeckRegistrationSubmitItemButton, new RoutedEventArgs());
    }

    private void TtTeckRegistrationSaveDraftButton_Click(object sender, RoutedEventArgs e)
    {
        _ttacCurrentPrescriptionId = null;
        _ttacCurrentCaptchaId = string.Empty;
        TtTeckRegistrationCaptchaTextBox.Text = string.Empty;
        TtTeckRegistrationCaptchaImage.Source = null;
        TtTeckRegistrationResultText.Text = _localization.GetString("ReadyForANewPrescription");
        UpdateTtacRegistrationStageButtons();
    }

    private string ExtractTtacSuccessSummary(JsonElement? root)
    {
        string defaultMessage = _localization.GetString("ItemRegistrationRequestWasSentSuccessfully");

        if (root == null)
            return defaultMessage;

        string? message = ReadJsonString(root.Value, "Message");
        if (!string.IsNullOrWhiteSpace(message) && message != "null")
            return message;

        string detailMessage = BuildTtacSuccessfulInquiryMessage(root.Value);
        return string.IsNullOrWhiteSpace(detailMessage) ? defaultMessage : detailMessage;
    }

    private string BuildTtacSuccessfulInquiryMessage(JsonElement root)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        var lines = new List<string>
        {
            _localization.GetString("TTACRegistrationInquiryWasSuccessful")
        };

        AddTtacResultLine(lines, root, _localization.GetString("PersianProductName"), "FaProductName", "faProductName", "PersianProductName", "persianProductName", "PersianName", "persianName");
        AddTtacResultLine(lines, root, _localization.GetString("EnglishProductName"), "EnProductName", "enProductName", "EnglishProductName", "englishProductName", "EnglishName", "englishName");
        AddTtacResultLine(lines, root, _localization.GetString("ProductName"), "ProductName", "productName", "ProductTitle", "productTitle", "Name", "name");
        AddTtacResultLine(lines, root, "UID", "UID", "uid", "Uid");
        AddTtacResultLine(lines, root, _localization.GetString("BatchCode"), "BatchCode", "batchCode", "BatchNumber", "batchNumber");
        AddTtacResultLine(lines, root, _localization.GetString("PackageCount"), "PackageCount", "packageCount");
        AddTtacResultLine(lines, root, _localization.GetString("PrescriptionItemID"), "PrescriptionItemId", "prescriptionItemId");
        AddTtacResultLine(lines, root, _localization.GetString("InquiredAmount"), "Amount", "amount", "InquiryAmount", "inquiryAmount", "Count", "count");

        string? priceMessage = FindJsonStringRecursive(root, "PriceMessage", "priceMessage");
        if (!string.IsNullOrWhiteSpace(priceMessage) && priceMessage != "null")
            lines.Add(priceMessage.Trim());

        AddTtacResultLine(lines, root, _localization.GetString("ProductPrice"), "Price", "price", "ProductPrice", "productPrice", "TotalPrice", "totalPrice", "ConsumerPrice", "consumerPrice", "SalePrice", "salePrice");
        AddTtacResultLine(lines, root, _localization.GetString("InsurancePayment"), "InsurancePayment", "insurancePayment", "InsuranceShare", "insuranceShare", "InsurerPayment", "insurerPayment", "OrganizationShare", "organizationShare");
        AddTtacResultLine(lines, root, _localization.GetString("PatientPayment"), "PatientPayment", "patientPayment", "PatientShare", "patientShare", "CustomerPayment", "customerPayment", "Payable", "payable", "PayAmount", "payAmount", "CashPayment", "cashPayment");
        AddTtacResultLine(lines, root, _localization.GetString("CurrencyDifference"), "CurrencyDifference", "currencyDifference", "CurrencyShare", "currencyShare", "SubsidyDifference", "subsidyDifference");

        AddAdditionalTtacFinancialLines(lines, root);

        return string.Join(Environment.NewLine, lines.Distinct());
    }

    private void AddTtacResultLine(List<string> lines, JsonElement root, string label, params string[] propertyNames)
    {
        string? value = FindJsonStringRecursive(root, propertyNames);
        if (string.IsNullOrWhiteSpace(value) || value == "null")
            return;

        bool money = IsTtacMoneyLabel(label);
        lines.Add($"{label}: {(money ? FormatTtacDisplayValue(value) : value.Trim())}");
    }

    private bool IsTtacMoneyLabel(string label)
    {
        return label.Contains("Price", StringComparison.OrdinalIgnoreCase)
               || label.Contains("Payment", StringComparison.OrdinalIgnoreCase)
               || label.Contains("Share", StringComparison.OrdinalIgnoreCase)
               || label.Contains("Difference", StringComparison.OrdinalIgnoreCase)
               || label.Contains("قیمت", StringComparison.OrdinalIgnoreCase)
               || label.Contains("پرداخت", StringComparison.OrdinalIgnoreCase)
               || label.Contains("بیمه", StringComparison.OrdinalIgnoreCase)
               || label.Contains("سهم", StringComparison.OrdinalIgnoreCase)
               || label.Contains("ارزی", StringComparison.OrdinalIgnoreCase)
               || label.Contains("مابه", StringComparison.OrdinalIgnoreCase);
    }

    private void AddAdditionalTtacFinancialLines(List<string> lines, JsonElement root)
    {
        var addedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in EnumerateJsonLeafValues(root))
        {
            string key = item.Key;
            string value = item.Value;
            if (string.IsNullOrWhiteSpace(value) || value == "null")
                continue;

            string normalized = key.Replace("_", "").Replace("-", "").ToLowerInvariant();
            if (normalized == "pricemessage" || normalized == "haspricemessage")
                continue;

            bool looksFinancial = normalized.Contains("price")
                                  || normalized.Contains("payment")
                                  || normalized.Contains("payable")
                                  || normalized.Contains("insurance")
                                  || normalized.Contains("share")
                                  || normalized.Contains("currency")
                                  || normalized.Contains("subsidy")
                                  || key.Contains("قیمت", StringComparison.OrdinalIgnoreCase)
                                  || key.Contains("پرداخت", StringComparison.OrdinalIgnoreCase)
                                  || key.Contains("بیمه", StringComparison.OrdinalIgnoreCase)
                                  || key.Contains("سهم", StringComparison.OrdinalIgnoreCase)
                                  || key.Contains("ارزی", StringComparison.OrdinalIgnoreCase);

            if (!looksFinancial || !addedKeys.Add(key))
                continue;

            string label = GetFriendlyTtacResultLabel(key);
            string line = $"{label}: {FormatTtacDisplayValue(value)}";
            if (!lines.Contains(line))
                lines.Add(line);
        }
    }

    private IEnumerable<KeyValuePair<string, string>> EnumerateJsonLeafValues(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                foreach (var child in EnumerateJsonLeafValues(property.Value))
                {
                    string key = string.IsNullOrWhiteSpace(child.Key) ? property.Name : child.Key;
                    yield return new KeyValuePair<string, string>(key, child.Value);
                }

                if (property.Value.ValueKind is JsonValueKind.String or JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False)
                    yield return new KeyValuePair<string, string>(property.Name, property.Value.ToString());
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                foreach (var child in EnumerateJsonLeafValues(item))
                    yield return child;
            }
        }
    }

    private string? FindJsonStringRecursive(JsonElement element, params string[] propertyNames)
    {
        var names = new HashSet<string>(propertyNames, StringComparer.OrdinalIgnoreCase);
        return FindJsonStringRecursive(element, names);
    }

    private string? FindJsonStringRecursive(JsonElement element, HashSet<string> propertyNames)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (propertyNames.Contains(property.Name))
                {
                    string value = property.Value.ValueKind == JsonValueKind.String ? property.Value.GetString() ?? string.Empty : property.Value.ToString();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }

                string? child = FindJsonStringRecursive(property.Value, propertyNames);
                if (!string.IsNullOrWhiteSpace(child))
                    return child;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                string? child = FindJsonStringRecursive(item, propertyNames);
                if (!string.IsNullOrWhiteSpace(child))
                    return child;
            }
        }

        return null;
    }

    private string GetFriendlyTtacResultLabel(string key)
    {
        string normalized = key.Replace("_", "").Replace("-", "").ToLowerInvariant();
        bool english = _localization.CurrentLanguage == AppLanguage.English;

        if (normalized.Contains("product") && normalized.Contains("price")) return _localization.GetString("ProductPrice");
        if (normalized.Contains("total") && normalized.Contains("price")) return _localization.GetString("TotalPrice");
        if (normalized.Contains("insurance")) return _localization.GetString("InsurancePayment");
        if (normalized.Contains("patient") || normalized.Contains("customer") || normalized.Contains("payable")) return _localization.GetString("PatientPayment");
        if (normalized.Contains("currency")) return _localization.GetString("CurrencyDifference");
        if (normalized.Contains("price")) return _localization.GetString("Price");
        if (normalized.Contains("payment")) return _localization.GetString("Payment");
        return key;
    }

    private string FormatTtacDisplayValue(string value)
    {
        value = value.Trim();
        if (decimal.TryParse(ToEnglishDigits(value).Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal number))
        {
            string formatted = string.Format(CultureInfo.InvariantCulture, "{0:N0}", number);
            return formatted + _localization.GetString("RialSuffix");
        }

        return value;
    }

    private void AddTtacRegistrationLog(bool success, string title, string message)
    {
        TtacRegistrationLogItems.Insert(0, new TtacRegistrationLogRow
        {
            Title = $"{DateTime.Now:HH:mm:ss} - {title}",
            Message = message,
            StatusBrush = success ? System.Windows.Media.Brushes.Green : System.Windows.Media.Brushes.Firebrick
        });
    }

    private string GetTtacRegistrationHistoryPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "ttac-registration-history.json");
    }

    private Dictionary<string, List<TtacRegistrationHistoryEntry>> LoadTtacRegistrationHistoryStore()
    {
        try
        {
            string path = GetTtacRegistrationHistoryPath();
            if (!File.Exists(path))
                return new Dictionary<string, List<TtacRegistrationHistoryEntry>>(StringComparer.OrdinalIgnoreCase);

            string json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json))
                return new Dictionary<string, List<TtacRegistrationHistoryEntry>>(StringComparer.OrdinalIgnoreCase);

            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind == JsonValueKind.Array)
            {
                var oldItems = JsonSerializer.Deserialize<List<TtacRegistrationHistoryEntry>>(json) ?? new List<TtacRegistrationHistoryEntry>();
                return new Dictionary<string, List<TtacRegistrationHistoryEntry>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = oldItems
                };
            }

            if (doc.RootElement.ValueKind == JsonValueKind.Object)
            {
                return JsonSerializer.Deserialize<Dictionary<string, List<TtacRegistrationHistoryEntry>>>(json)
                       ?? new Dictionary<string, List<TtacRegistrationHistoryEntry>>(StringComparer.OrdinalIgnoreCase);
            }
        }
        catch { }

        return new Dictionary<string, List<TtacRegistrationHistoryEntry>>(StringComparer.OrdinalIgnoreCase);
    }

    private void SaveTtacRegistrationHistoryStore(Dictionary<string, List<TtacRegistrationHistoryEntry>> store)
    {
        try
        {
            string json = JsonSerializer.Serialize(store, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(GetTtacRegistrationHistoryPath(), json);
        }
        catch { }
    }

    private void LoadTtacRegistrationHistory()
    {
        LoadTtacRegistrationHistoryForKey("default");
    }

    private void LoadTtacRegistrationHistoryForKey(string key)
    {
        key = string.IsNullOrWhiteSpace(key) ? "default" : key;
        var store = LoadTtacRegistrationHistoryStore();
        _ttacRegistrationHistory.Clear();
        if (store.TryGetValue(key, out var items))
            _ttacRegistrationHistory.AddRange(items);
        _ttacRegistrationHistoryLoadedPharmacyKey = key;
    }

    private void SaveTtacRegistrationHistory()
    {
        SaveTtacRegistrationHistoryForKey(string.IsNullOrWhiteSpace(_ttacRegistrationHistoryLoadedPharmacyKey) ? "default" : _ttacRegistrationHistoryLoadedPharmacyKey);
    }

    private void SaveTtacRegistrationHistoryForKey(string key)
    {
        key = string.IsNullOrWhiteSpace(key) ? "default" : key;
        var store = LoadTtacRegistrationHistoryStore();
        store[key] = _ttacRegistrationHistory.ToList();
        SaveTtacRegistrationHistoryStore(store);
    }

    private void LoadTtacRegistrationHistoryForPharmacy(string pharmacyName)
    {
        string key = GetReceiveStatusStorageKey(pharmacyName);
        if (string.Equals(_ttacRegistrationHistoryLoadedPharmacyKey, key, StringComparison.OrdinalIgnoreCase))
            return;

        SaveTtacRegistrationHistory();
        LoadTtacRegistrationHistoryForKey(key);
        RefreshTtTeckHistoryItems();
        TtacPanelListBox?.Items.Refresh();
    }

    private void AddPersistentTtacRegistrationHistory(bool success, string message)
    {
        if (_pendingRegistrationTtTeckRow == null)
            return;

        string nationalId = ToEnglishDigits(TtTeckRegistrationNationalIdTextBox.Text.Trim());
        string mobile = ToEnglishDigits(TtTeckRegistrationMobileTextBox.Text.Trim());
        string amount = string.IsNullOrWhiteSpace(TtTeckRegistrationAmountTextBox.Text) ? "1" : ToEnglishDigits(TtTeckRegistrationAmountTextBox.Text.Trim());
        bool isElectronic = TtTeckRegistrationTypeComboBox.SelectedIndex == 1;

        var entry = new TtacRegistrationHistoryEntry
        {
            RegisteredAt = DateTime.Now,
            Barcode = _pendingRegistrationTtTeckRow.Barcode,
            ProductName = _pendingRegistrationTtTeckRow.PersianProductName,
            RegistrationType = isElectronic ? (_localization.GetString("PrescriptionBased")) : (_localization.GetString("WithoutPrescription")),
            PrescriptionId = _ttacCurrentPrescriptionId,
            Amount = amount,
            NationalIdFull = nationalId,
            MobileFull = mobile,
            NationalIdMasked = MaskNationalId(nationalId),
            MobileMasked = MaskMobile(mobile),
            PatientFullName = _ttacCurrentPatientFullName,
            Success = success,
            Message = message
        };

        _ttacRegistrationHistory.Insert(0, entry);
        SaveTtacRegistrationHistory();
    }

    private static string MaskNationalId(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length < 4)
            return string.Empty;
        return "‎" + new string('*', Math.Max(0, value.Length - 4)) + value[^4..] + "‎";
    }

    private static string MaskMobile(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length < 7)
            return string.Empty;
        return "‎" + value[..4] + "***" + value[^4..] + "‎";
    }

    private bool IsBarcodeTtacRegistered(string barcode)
    {
        return _ttacRegistrationHistory.Any(x => x.Success && x.Barcode == barcode);
    }

    private string GetRegistrationButtonText(bool registered)
    {
        if (registered)
        {
            return _localization.GetString("Registered");
        }

        return RegisterInTtTeckButtonText;
    }

    private System.Windows.Media.Brush GetRegistrationButtonBrush(bool registered)
    {
        return registered
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81))
            : new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0));
    }

    // برچسب‌های پیام تی‌تک که مربوط به خود فرآورده هستند (نه بیمار/پرداخت این ثبت خاص)؛
    // برای جدا کردن ستون «اطلاعات فرآورده» از ستون «اطلاعات بیمار و پرداخت» در دیالوگ
    // تاریخچه استفاده می‌شود. AddTtacResultLine همین برچسب‌ها را هنگام ساخت پیام می‌نویسد.
    private static readonly string[] TtacFormulaMessageLabels =
    {
        "نام فارسی فرآورده", "نام انگلیسی فرآورده", "نام فرآورده", "سری ساخت", "تعداد بسته", "شناسه قلم نسخه",
        "Persian product name", "English product name", "Product name", "Batch code", "Package count", "Prescription item ID"
    };

    private static (string formulaLines, string otherLines) SplitTtacMessageForColumns(string? message)
    {
        var formula = new List<string>();
        var other = new List<string>();
        if (!string.IsNullOrWhiteSpace(message))
        {
            foreach (var rawLine in message.Replace("\r\n", "\n").Split('\n'))
            {
                string line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                bool isFormulaLine = TtacFormulaMessageLabels.Any(label => line.StartsWith(label + ":", StringComparison.Ordinal));
                (isFormulaLine ? formula : other).Add(line);
            }
        }

        return (string.Join(Environment.NewLine, formula), string.Join(Environment.NewLine, other));
    }

    private void ShowTtacRegistrationHistoryForBarcode(string barcode)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        var entries = _ttacRegistrationHistory
            .Where(x => x.Barcode == barcode)
            .OrderByDescending(x => x.RegisteredAt)
            .Take(8)
            .ToList();

        if (entries.Count == 0)
        {
            ShowStyledMessage(
                _localization.GetString("RegistrationHistory"),
                _localization.GetString("NoRegistrationHistoryWasFoundForThisProduct"),
                true);
            return;
        }

        string formulaInfoText = string.Empty;
        var messageBuilder = new StringBuilder();
        var patientBuilder = new StringBuilder();
        bool first = true;
        foreach (var entry in entries)
        {
            string registrationType = entry.RegistrationType;
            if (english)
            {
                registrationType = registrationType.Contains("نسخه", StringComparison.OrdinalIgnoreCase) && !registrationType.Contains("فاقد", StringComparison.OrdinalIgnoreCase)
                    ? "Prescription-based"
                    : "Without prescription";
            }

            var (formulaLines, otherLines) = SplitTtacMessageForColumns(entry.Message);
            if (string.IsNullOrWhiteSpace(formulaInfoText) && !string.IsNullOrWhiteSpace(formulaLines))
                formulaInfoText = formulaLines;

            if (!first)
            {
                messageBuilder.AppendLine(TtacHistoryDividerLine);
                patientBuilder.AppendLine(TtacHistoryDividerLine);
            }
            first = false;

            if (!string.IsNullOrWhiteSpace(otherLines))
                messageBuilder.AppendLine(otherLines);

            patientBuilder.AppendLine($"{entry.RegisteredAt:yyyy/MM/dd HH:mm:ss} - {registrationType}");
            if (entry.PrescriptionId.HasValue)
                patientBuilder.AppendLine($"{(_localization.GetString("PrescriptionID"))}: {entry.PrescriptionId.Value}");
            patientBuilder.AppendLine($"{(_localization.GetString("Amount"))}: {entry.Amount}");
            if (!string.IsNullOrWhiteSpace(entry.NationalIdMasked))
                patientBuilder.AppendLine($"{(_localization.GetString("NationalID"))}: {entry.NationalIdMasked}");
            if (!string.IsNullOrWhiteSpace(entry.PatientFullName))
                patientBuilder.AppendLine($"{(_localization.GetString("Patient"))}: {entry.PatientFullName}");
            if (!string.IsNullOrWhiteSpace(entry.MobileMasked))
                patientBuilder.AppendLine($"{(_localization.GetString("Mobile"))}: {entry.MobileMasked}");
            patientBuilder.AppendLine(_localization.GetFormattedString("ResultStatus", entry.Success ? _localization.GetString("Successful") : _localization.GetString("Failed")));
        }

        if (string.IsNullOrWhiteSpace(formulaInfoText))
            formulaInfoText = entries[0].ProductName;

        ShowTtacHistoryDetailOverlay(
            _localization.GetString("RegistrationHistoryForThisProduct"),
            formulaInfoText,
            messageBuilder.ToString(),
            patientBuilder.ToString(),
            GetFormulaPhotoPathForBarcode(barcode));
    }

    // خط جداکننده‌ی بین ثبت‌های مختلف یک محصول در دیالوگ «تاریخچه ثبت این محصول»؛ هم برای
    // ساخت متن استفاده می‌شود، هم بعداً برای شکستن دوباره‌ی متن رنگی قیمت به ازای هر ثبت.
    private const string TtacHistoryDividerLine = "――――――――――――";

    private void ShowTtacHistoryDetailOverlay(string title, string formulaInfo, string messageInfo, string patientInfo, string? photoPath)
    {
        TtacHistoryDetailTitle.Text = title;
        TtacHistoryDetailFormulaText.Text = string.IsNullOrWhiteSpace(formulaInfo) ? "-" : formulaInfo;
        TtacHistoryDetailSideText.Text = string.IsNullOrWhiteSpace(patientInfo) ? "-" : patientInfo;

        // قیمت‌ها با همون منطق رنگی قرمز/سبز که در دیالوگ «ثبت شد» استفاده می‌شود نمایش داده می‌شوند
        // (AppendTtacColorizedRuns)، چون این متن هم از همان entry.Message می‌آید.
        TtacHistoryDetailMessageText.Inlines.Clear();
        if (string.IsNullOrWhiteSpace(messageInfo))
        {
            TtacHistoryDetailMessageText.Inlines.Add(new Run("-"));
        }
        else
        {
            string[] chunks = messageInfo.Replace("\r\n", "\n").Split(new[] { TtacHistoryDividerLine }, StringSplitOptions.None);
            for (int i = 0; i < chunks.Length; i++)
            {
                if (i > 0)
                {
                    TtacHistoryDetailMessageText.Inlines.Add(new LineBreak());
                    TtacHistoryDetailMessageText.Inlines.Add(new Run(TtacHistoryDividerLine) { Foreground = System.Windows.Media.Brushes.Gray });
                    TtacHistoryDetailMessageText.Inlines.Add(new LineBreak());
                }

                AppendTtacColorizedRuns(TtacHistoryDetailMessageText.Inlines, chunks[i].Trim('\n', '\r', ' '));
            }
        }

        SetOverlayProductPhoto(TtacHistoryDetailPhoto, TtacHistoryDetailPhotoBorder, photoPath);

        System.Windows.Controls.Panel.SetZIndex(TtacHistoryDetailOverlay, 355);
        TtacHistoryDetailOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void TtacHistoryDetailOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseTtacHistoryDetailOverlay();
    }

    private void TtacHistoryDetailCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseTtacHistoryDetailButton_Click(object sender, RoutedEventArgs e)
    {
        CloseTtacHistoryDetailOverlay();
    }

    private void CloseTtacHistoryDetailOverlay()
    {
        TtacHistoryDetailOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private async void TtTeckRegistrationOpenWebButton_Click(object sender, RoutedEventArgs e)
    {
        string url = TtTeckRegistrationTypeComboBox?.SelectedIndex == 1
            ? "https://newstatisticsreports.ttac.ir/pharmacyDashboard/electronicPrescription"
            : "https://newstatisticsreports.ttac.ir/pharmacyDashboard/nonePrescription";

        await OpenTtTeckInternalBrowserAsync(url);
    }

    private void UpdateTtTeckRegistrationLocalizedTexts()
    {
        if (TtTeckRegistrationTitle == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        var flow = english ? System.Windows.FlowDirection.LeftToRight : System.Windows.FlowDirection.RightToLeft;

        SetNamedFlowDirectionIfExists("TtTeckRegistrationCaptchaColumn", flow);
        SetNamedFlowDirectionIfExists("TtTeckRegistrationInputsColumn", flow);
        SetNamedFlowDirectionIfExists("TtTeckRegistrationProductColumn", flow);
        TtTeckRegistrationProductText.FlowDirection = flow;
        TtTeckRegistrationBarcodeText.FlowDirection = System.Windows.FlowDirection.LeftToRight;
        TtTeckRegistrationResultText.FlowDirection = flow;
        if (_pendingRegistrationTtTeckRow != null)
        {
            string localizedProductName = GetLocalizedProductDisplayName(_pendingRegistrationTtTeckRow.PersianProductName, _pendingRegistrationTtTeckRow.EnglishProductName);
            TtTeckRegistrationProductText.Text = string.IsNullOrWhiteSpace(localizedProductName) ? _pendingRegistrationTtTeckRow.Barcode : localizedProductName;
        }
        TtTeckBirthDayTextBox.ToolTip = _localization.GetString("Day");
        TtTeckBirthMonthTextBox.ToolTip = _localization.GetString("Month");
        TtTeckBirthYearTextBox.ToolTip = _localization.GetString("Year");

        if (english)
        {
            TtTeckRegistrationTitle.Text = "Register in TtTeck";
            TtTeckRegistrationProductLabel.Text = "Selected product";
            SetTtacRegistrationProductHelpText("For fast registration use Enter. Use the mouse only to reset the captcha.");
            TtTeckRegistrationTypeLabel.Text = "Registration type";
            TtTeckRegistrationNonePrescriptionItem.Content = "Without prescription";
            TtTeckRegistrationNonePrescriptionButton.Content = "Without prescription";
            TtTeckRegistrationElectronicPrescriptionButton.Content = "Prescription-based";
            TtTeckRegistrationElectronicPrescriptionItem.Content = "Prescription-based";
            TtTeckRegistrationAmountLabel.Text = "Amount";
            TtTeckRegistrationNationalIdLabel.Text = "Patient national ID";
            TtTeckRegistrationMobileLabel.Text = "Patient / visitor mobile number (optional)";
            TtTeckRegistrationBirthDateLabel.Text = "Patient birth date (Persian calendar)";
            TtTeckRegistrationBirthDateHint.Text = "Day / Month / Year - example: 20 / 6 / 1379";
            TtTeckRegistrationMedicalCouncilLabel.Text = "Medical council number";
            TtTeckRegistrationMedicalCouncilHint.Text = "Required for prescription-based registration; it will be sent as MD_number.";
            TtTeckRegistrationPhaseNote.Text = "Enter the small text in the image.";
            TtTeckRegistrationGetCaptchaButton.Content = "Reset";
            TtTeckRegistrationCreatePrescriptionButton.Content = "Create prescription";
            TtTeckRegistrationSubmitItemButton.Content = "Submit item";
            TtTeckRegistrationSaveDraftButton.Content = "Finish / new prescription";
            TtTeckRegistrationOpenWebButton.Content = "Open TtTeck portal";
            TtTeckRegistrationOpenWebButton.Visibility = Visibility.Collapsed;
            TtTeckRegistrationCancelButton.Content = "Cancel";
            TtTeckWebViewTitle.Text = "TtTeck internal browser";
            TtTeckWebViewBackButton.Content = "Back";
            TtTeckWebViewRefreshButton.Content = "Refresh";
            TtTeckWebViewOpenExternalButton.Content = "Open in browser";
        }
        else
        {
            TtTeckRegistrationTitle.Text = "ثبت در سامانه تی‌تک";
            TtTeckRegistrationProductLabel.Text = "محصول انتخاب‌شده";
            SetTtacRegistrationProductHelpText("برای ثبت سریع، از Enter استفاده کنید. فقط برای بازنشانی کپچا به موس نیاز دارید.");
            TtTeckRegistrationTypeLabel.Text = "نوع ثبت";
            TtTeckRegistrationNonePrescriptionItem.Content = "فاقد نسخه";
            TtTeckRegistrationNonePrescriptionButton.Content = "فاقد نسخه";
            TtTeckRegistrationElectronicPrescriptionButton.Content = "نسخه‌محور";
            TtTeckRegistrationElectronicPrescriptionItem.Content = "نسخه‌محور";
            TtTeckRegistrationAmountLabel.Text = "تعداد";
            TtTeckRegistrationNationalIdLabel.Text = "کد ملی بیمار";
            TtTeckRegistrationMobileLabel.Text = "شماره همراه بیمار / مراجعه‌کننده (اختیاری)";
            TtTeckRegistrationBirthDateLabel.Text = "تاریخ تولد بیمار (شمسی)";
            TtTeckRegistrationBirthDateHint.Text = "روز / ماه / سال  - مثال: 20 / 6 / 1379";
            TtTeckRegistrationMedicalCouncilLabel.Text = "شماره نظام پزشکی";
            TtTeckRegistrationMedicalCouncilHint.Text = "برای نسخه‌محور الزامی است؛ برنامه آن را به صورت MD_شماره ارسال می‌کند.";
            TtTeckRegistrationPhaseNote.Text = "متن ریز تصویر را در کادر زیر وارد کنید.";
            TtTeckRegistrationGetCaptchaButton.Content = "بازنشانی";
            TtTeckRegistrationCreatePrescriptionButton.Content = "ایجاد نسخه";
            TtTeckRegistrationSubmitItemButton.Content = "ثبت قلم";
            TtTeckRegistrationSaveDraftButton.Content = "اتمام / نسخه جدید";
            TtTeckRegistrationOpenWebButton.Content = "باز کردن پنل تی‌تک";
            TtTeckRegistrationOpenWebButton.Visibility = Visibility.Collapsed;
            TtTeckRegistrationCancelButton.Content = "انصراف";
            TtTeckWebViewTitle.Text = "مرورگر داخلی تی‌تک";
            TtTeckWebViewBackButton.Content = "بازگشت";
            TtTeckWebViewRefreshButton.Content = "تازه‌سازی";
            TtTeckWebViewOpenExternalButton.Content = "باز کردن در مرورگر";
        }
    }

    private void UpdateTtacLoginLocalizedTexts()
    {
        if (TtacLoginTitle == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        TtacLoginTitle.Text = _localization.GetString("TTACLogin");
        TtacLoginDescription.Text = _localization.GetString("TheTTACSessionHasExpiredOrYouHavenTLoggedInYetPleaseLogInUsingTheInternalBrowser");

        // فرم ورود مستقیم (نام‌کاربری/رمز/دکمه سبز) به دلیل محافظت ضدربات سایت تی‌تک قابل
        // اعتماد نبود (حتی گرفتن صفحه‌ی ورود هم شکست می‌خورد) - طبق درخواست کاربر مخفی نگه
        // داشته می‌شود و فقط مسیر «مرورگر داخلی» (که موتور واقعی مرورگر دارد) فعال است.
        TtacLoginUsernameLabel.Visibility = Visibility.Collapsed;
        TtacLoginUsernameTextBox.Visibility = Visibility.Collapsed;
        TtacLoginUsernameHintText.Visibility = Visibility.Collapsed;
        TtacLoginPasswordLabel.Visibility = Visibility.Collapsed;
        TtacLoginPasswordBox.Visibility = Visibility.Collapsed;
        TtacLoginRememberCheckBox.Visibility = Visibility.Collapsed;
        TtacLoginSubmitButton.Visibility = Visibility.Collapsed;

        TtacLoginOpenBrowserButton.Content = _localization.GetString("InternalBrowser");
        TtacLoginOpenBrowserButton.Width = english ? 150 : 150;
        TtacLoginCancelButton.Content = _localization.GetString("CancelButton");
    }

    private async void TtTeckRegistrationGetCaptchaButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadTtacCaptchaAsync(true);
    }

    private bool IsTtacSessionExpiredException(Exception ex)
    {
        return ex is TtacSessionExpiredException
               || ex.Message.Contains("نشست تی‌تک", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("session expired", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("Forbidden", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("401", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("403", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsJsonNullElementException(Exception ex)
    {
        return ex is InvalidOperationException
               && ex.Message.Contains("requires an element of type", StringComparison.OrdinalIgnoreCase)
               && ex.Message.Contains("Null", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsReceiveStatusOperationTitle(string? title)
    {
        return !string.IsNullOrWhiteSpace(title)
               && (title.Contains("تعیین وضعیت", StringComparison.OrdinalIgnoreCase)
                   || title.Contains("Receive status", StringComparison.OrdinalIgnoreCase));
    }

    private void HandleTtacOperationException(Exception ex, string title, Func<Task> retryAction, Action<string, string>? onFailureMessageShown = null, string? pendingLabel = null)
    {
        // فقط وقتی توکن واقعاً معتبر نیست، «انقضای نشست» است؛ اگر توکن سر جایش باشد، پیام خطا
        // فقط اتفاقی ۴۰۱/۴۰۳ یا Unauthorized را داخل خودش دارد (مثل پیام‌های خطای عادی سامانه)
        // و باید همان پیام واقعی خطا نشان داده شود، نه پنجره‌ی ورود.
        if (IsTtacSessionExpiredException(ex) && !HasValidTtacToken())
        {
            _pendingTtacRetryAction = retryAction;
            _pendingTtacRetryLabel = pendingLabel;
            ShowTtacLoginOverlay(sessionExpired: true);
            TtacLoginStatusText.Text = _localization.GetString("YourTTACSessionExpiredLoginAgainThePreviousOperationWillContinueAutomatically");
            return;
        }

        if (IsJsonNullElementException(ex) && IsReceiveStatusOperationTitle(title))
        {
            AddTtacRegistrationLog(false, title, _localization.GetString("TTACReturnedAnEmptyResultForThisReceiveStatusRequest"));
            return;
        }

        AddTtacRegistrationLog(false, title, ex.Message);
        ShowStyledMessage(title, GetFriendlyTtacErrorMessage(ex.Message), true);
        // فقط دقیقاً همین‌جا (یعنی وقتی واقعاً یک پیام خطای نهایی به کاربر نشان داده شده - نه در
        // حالت نشست‌منقضی‌شده که خودکار دوباره تلاش می‌شود) به فراخواننده اطلاع داده می‌شود، تا
        // پیامی که (فقط برای شیرخشک) قرار است روی گوشی هم نشان داده شود، دقیقاً با چیزی که روی
        // دسکتاپ دیده می‌شود یکی باشد.
        onFailureMessageShown?.Invoke(title, GetFriendlyTtacErrorMessage(ex.Message));
    }

    // بعضی خطاهای تی‌تک پیام خامِ بلند و گیج‌کننده‌ای دارند (مثلاً خطای ۵۱۷۳ که لیست بلندی از
    // شماره سری‌ها را همراه با «سهمیه‌ی باقیمانده» برمی‌گرداند). برای کاربر، به‌جای آن پیام خام،
    // یک پیام کوتاه و واضح نشان بده؛ نسخه‌ی خام همچنان در لاگ ثبت می‌شود.
    private string GetFriendlyTtacErrorMessage(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
            return message;

        if (message.Contains("5173", StringComparison.OrdinalIgnoreCase)
            && (message.Contains("سهمیه", StringComparison.OrdinalIgnoreCase)
                || message.Contains("quota", StringComparison.OrdinalIgnoreCase)))
        {
            return _localization.GetString("TtacQuotaExhaustedFriendly");
        }

        return message;
    }

    // جلوگیری از اجرای هم‌زمانِ عملیات در انتظار (مثلاً وقتی هم NavigationCompleted توکن را پیدا
    // می‌کند و هم مانیتور) - عملیات فقط یک بار اجرا می‌شود.
    private bool _isTtacPendingRetryRunning;

    private async Task RunPendingTtacRetryIfAnyAsync()
    {
        if (_isTtacPendingRetryRunning)
            return;

        var retry = _pendingTtacRetryAction;
        _pendingTtacRetryAction = null;
        _pendingTtacRetryLabel = null;
        if (retry == null)
            return;

        _isTtacPendingRetryRunning = true;
        try
        {
            await Dispatcher.BeginInvoke(new Action(async () =>
            {
                try
                {
                    await retry();
                }
                catch (Exception ex)
                {
                    ShowStyledMessage(GetLocalizedLookupFailedTitle(), ex.Message, true);
                }
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }
        finally
        {
            _isTtacPendingRetryRunning = false;
        }
    }

    // ---------- TTAC Direct Login ----------
    //
    // پیش از این، هر بار که این پنجره باز می‌شد، فایل ذخیره‌شده بی‌قید‌وشرط پاک می‌شد (چند خط
    // پایین‌تر در همین متد) - یعنی حتی اگر کاربر تیک «به خاطر بسپار» را می‌زد، همان لحظه‌ی
    // باز شدن دوباره‌ی این پنجره (که به خاطر منقضی‌شدن نشست تی‌تک خیلی زیاد اتفاق می‌افتد)
    // ذخیره پاک می‌شد و رمز هیچ‌وقت واقعاً یادش نمی‌ماند. همچنین دکمه‌ی «ورود» اصلاً به ورود
    // مستقیم (LoginToTtacDirectAsync) وصل نبود و فقط مرورگر داخلی را باز می‌کرد. این دو مورد
    // اینجا اصلاح شده‌اند. ذخیره الان یک لیست از چند ورود (برای چند داروخانه/کد) است، نه یک
    // مورد تکی، تا اسم داروخانه بر اساس نام کاربری هم پیشنهاد داده شود.

    private sealed class TtacSavedLogin
    {
        public string Username { get; set; } = string.Empty;
        public string PasswordProtectedBase64 { get; set; } = string.Empty;
        public string PharmacyName { get; set; } = string.Empty;
        public DateTime LastUsedUtc { get; set; } = DateTime.UtcNow;

        // برای نمایش در لیست تنظیمات: اگر نام داروخانه دارد آن را نشان بده وگرنه خود کد.
        public string DisplayLabel => string.IsNullOrWhiteSpace(PharmacyName) ? Username : PharmacyName;
    }

    private const int MaxSavedTtacLogins = 12;

    private string GetTtacSavedLoginPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "ttac-login.dat");
    }

    /// <summary>
    /// مسیر فایلی که شناسه‌ی لایسنس فعال هنگام ذخیره‌ی حساب‌های ذخیره‌شده را نگه می‌دارد.
    /// اگر لایسنس عوض شود (نه تمدید)، حساب‌های ذخیره‌شده پاک می‌شوند.
    /// </summary>
    private string GetTtacSavedLoginLicenseIdPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "ttac-login-license.dat");
    }

    private string ReadSavedLoginLicenseId()
    {
        try
        {
            string path = GetTtacSavedLoginLicenseIdPath();
            if (File.Exists(path))
                return File.ReadAllText(path).Trim();
        }
        catch { }
        return string.Empty;
    }

    private void WriteSavedLoginLicenseId(string licenseId)
    {
        try
        {
            File.WriteAllText(GetTtacSavedLoginLicenseIdPath(), licenseId, Encoding.UTF8);
        }
        catch { }
    }

    /// <summary>
    /// اگر لایسنس عوض شده باشد (تمدید نشده)، حساب‌های ذخیره‌شده را پاک می‌کند.
    /// تمدید لایسنس یعنی LicenseId یکی باشد؛ تعویض لایسنس یعنی LicenseId فرق کند.
    /// </summary>
    private void ClearSavedLoginsIfLicenseChanged()
    {
        try
        {
            string currentLicenseId = _activeLicense?.LicenseId ?? string.Empty;
            string savedLicenseId = ReadSavedLoginLicenseId();

            // اگر هنوز لایسنسی ذخیره نشده (اولین اجرا یا بعد از پاک شدن)، کاری نکن.
            if (string.IsNullOrEmpty(savedLicenseId))
                return;

            // اگر لایسنس عوض شده (متفاوت از لایسنسی که هنگام ذخیره حساب‌ها فعال بوده)
            if (!string.Equals(currentLicenseId, savedLicenseId, StringComparison.OrdinalIgnoreCase))
            {
                // حساب‌های ذخیره‌شده را پاک کن
                SaveTtacLoginsList(new List<TtacSavedLogin>());
                RefreshTtacSavedLoginsList();
                RefreshTtacQuickLoginButtons();
            }
        }
        catch { }
    }

    private List<TtacSavedLogin> LoadSavedTtacLogins()
    {
        try
        {
            string path = GetTtacSavedLoginPath();
            if (!File.Exists(path))
                return new List<TtacSavedLogin>();

            string json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json))
                return new List<TtacSavedLogin>();

            // فرمت جدید: آرایه‌ای از چند ورود.
            try
            {
                var list = JsonSerializer.Deserialize<List<TtacSavedLogin>>(json);
                if (list != null)
                    return list;
            }
            catch { }

            // سازگاری با فایل‌های قدیمی: یک شیء تکی {"username":..,"password":..}.
            try
            {
                var single = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                if (single != null && single.TryGetValue("username", out var oldUsername) && single.TryGetValue("password", out var oldPassword))
                {
                    return new List<TtacSavedLogin>
                    {
                        new TtacSavedLogin { Username = oldUsername, PasswordProtectedBase64 = oldPassword }
                    };
                }
            }
            catch { }
        }
        catch { }

        return new List<TtacSavedLogin>();
    }

    private void SaveTtacLoginsList(List<TtacSavedLogin> logins)
    {
        try
        {
            File.WriteAllText(GetTtacSavedLoginPath(), JsonSerializer.Serialize(logins, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private void UpsertSavedTtacLogin(string username, string password, string? pharmacyName)
    {
        try
        {
            byte[] encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(password), null, DataProtectionScope.CurrentUser);
            string protectedBase64 = Convert.ToBase64String(encrypted);

            var logins = LoadSavedTtacLogins();
            var existing = logins.FirstOrDefault(x => string.Equals(x.Username, username, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.PasswordProtectedBase64 = protectedBase64;
                existing.LastUsedUtc = DateTime.UtcNow;
                if (!string.IsNullOrWhiteSpace(pharmacyName))
                    existing.PharmacyName = pharmacyName;
            }
            else
            {
                logins.Add(new TtacSavedLogin
                {
                    Username = username,
                    PasswordProtectedBase64 = protectedBase64,
                    PharmacyName = pharmacyName ?? string.Empty,
                    LastUsedUtc = DateTime.UtcNow
                });
            }

            // فقط جدیدترین‌ها نگه داشته می‌شوند تا فایل بی‌رویه بزرگ نشود.
            logins = logins.OrderByDescending(x => x.LastUsedUtc).Take(MaxSavedTtacLogins).ToList();
            SaveTtacLoginsList(logins);

            // شناسه‌ی لایسنس فعلی را ذخیره کن تا در اجراهای بعدی مشخص باشد
            // حساب‌ها متعلق به کدام لایسنس بوده‌اند.
            WriteSavedLoginLicenseId(_activeLicense?.LicenseId ?? string.Empty);
        }
        catch { }
    }

    private void RemoveSavedTtacLogin(string username)
    {
        try
        {
            var logins = LoadSavedTtacLogins();
            int removed = logins.RemoveAll(x => string.Equals(x.Username, username, StringComparison.OrdinalIgnoreCase));
            if (removed > 0)
                SaveTtacLoginsList(logins);
        }
        catch { }
    }

    // ---------- بخش «حساب‌های ذخیره‌شده تی‌تک» در تنظیمات ----------

    private void RefreshTtacSavedLoginsList()
    {
        if (TtacSavedLoginsItemsControl == null)
            return;

        var logins = LoadSavedTtacLogins().OrderByDescending(x => x.LastUsedUtc).ToList();
        TtacSavedLoginsItemsControl.ItemsSource = logins;

        if (TtacSavedLoginsEmptyText != null)
            TtacSavedLoginsEmptyText.Visibility = logins.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void UpdateTtacSavedLoginsLocalizedTexts()
    {
        if (TtacSavedLoginsSectionTitle == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        TtacSavedLoginsSectionTitle.Text = _localization.GetString("SavedTTACAccounts");
        TtacSavedLoginsSectionDescription.Text = _localization.GetString("SaveTheUsernamePasswordAndPharmacyNameALoginToButtonWillAppearOnTheMainScreenForEachSavedPharmacyAndItFillsTheCodeAndPasswordAutomaticallyOnTheTTACLoginPage");
        TtacSavedLoginUsernameLabel.Text = _localization.GetString("UsernamePharmacyCode");
        TtacSavedLoginPasswordLabel.Text = _localization.GetString("Password");
        TtacSavedLoginPharmacyNameLabel.Text = _localization.GetString("PharmacyName");
        TtacSaveSavedLoginButton.Content = _localization.GetString("Save");
        TtacSavedLoginsEmptyText.Text = _localization.GetString("NoSavedAccountsYet");
        TtacLoginWithoutSavingButton.Content = _localization.GetString("LoginWithoutSaving");
        TtacLoginWithoutSavingHint.Text = _localization.GetString("LoginWithoutSavingHint");
    }

    // «ورود بدون ذخیره حساب»: مرورگر داخلی را مستقیم باز می‌کند تا کاربری که نمی‌خواهد حسابش
    // را ذخیره کند هم بتواند بدون پر کردن فرم، وارد شود. هیچ حسابی ذخیره یا خودکار پر نمی‌شود.
    private async void TtacLoginWithoutSavingButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TtacLoginWithoutSavingButton.IsEnabled = false;
            await OpenTtTeckInternalBrowserAsync("https://newstatisticsreports.ttac.ir/pharmacyDashboard");
            _ = MonitorTtacConnectionAfterBrowserOpenAsync();
        }
        catch { }
        finally
        {
            TtacLoginWithoutSavingButton.IsEnabled = true;
        }
    }

    private bool _isPasswordVisible = false;

    // چشم = رمز دیده می‌شود؛ چشم با خط = رمز مخفی است.
    private void UpdateTtacSavedLoginPasswordToggleIcon(bool passwordVisible)
    {
        if (TtacSavedLoginPasswordSlash != null)
            TtacSavedLoginPasswordSlash.Visibility = passwordVisible ? Visibility.Collapsed : Visibility.Visible;
        if (TtacSavedLoginPasswordToggleIcon != null)
            TtacSavedLoginPasswordToggleIcon.Text = "👁";
        if (TtacSavedLoginPasswordToggle != null)
            TtacSavedLoginPasswordToggle.ToolTip = passwordVisible ? "مخفی کردن رمز عبور" : "نمایش رمز عبور";
    }

    private void TtacSavedLoginPasswordToggle_Click(object sender, RoutedEventArgs e)
    {
        _isPasswordVisible = !_isPasswordVisible;
        if (_isPasswordVisible)
        {
            TtacSavedLoginPasswordTextBox.Text = TtacSavedLoginPasswordBox.Password;
            TtacSavedLoginPasswordTextBox.Visibility = Visibility.Visible;
            TtacSavedLoginPasswordBox.Visibility = Visibility.Collapsed;
            UpdateTtacSavedLoginPasswordToggleIcon(true);
            TtacSavedLoginPasswordTextBox.Focus();
            TtacSavedLoginPasswordTextBox.CaretIndex = TtacSavedLoginPasswordTextBox.Text.Length;
        }
        else
        {
            TtacSavedLoginPasswordBox.Password = TtacSavedLoginPasswordTextBox.Text;
            TtacSavedLoginPasswordBox.Visibility = Visibility.Visible;
            TtacSavedLoginPasswordTextBox.Visibility = Visibility.Collapsed;
            UpdateTtacSavedLoginPasswordToggleIcon(false);
            TtacSavedLoginPasswordBox.Focus();
        }
    }

    private void TtacSaveSavedLoginButton_Click(object sender, RoutedEventArgs e)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        string username = TtacSavedLoginUsernameTextBox.Text?.Trim() ?? string.Empty;
        string password = _isPasswordVisible
            ? (TtacSavedLoginPasswordTextBox.Text ?? string.Empty)
            : (TtacSavedLoginPasswordBox.Password ?? string.Empty);
        string pharmacyName = TtacSavedLoginPharmacyNameTextBox.Text?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
        {
            TtacSavedLoginStatusText.Text = _localization.GetString("EnterTheUsernameCodeAndThePassword");
            TtacSavedLoginStatusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xDC, 0x26, 0x26));
            TtacSavedLoginStatusText.Visibility = Visibility.Visible;
            return;
        }

        UpsertSavedTtacLogin(username, password, string.IsNullOrWhiteSpace(pharmacyName) ? null : pharmacyName);

        TtacSavedLoginUsernameTextBox.Text = string.Empty;
        TtacSavedLoginPasswordBox.Password = string.Empty;
        TtacSavedLoginPasswordTextBox.Text = string.Empty;
        TtacSavedLoginPharmacyNameTextBox.Text = string.Empty;
        _isPasswordVisible = false;
        TtacSavedLoginPasswordBox.Visibility = Visibility.Visible;
        TtacSavedLoginPasswordTextBox.Visibility = Visibility.Collapsed;
        UpdateTtacSavedLoginPasswordToggleIcon(false);

        TtacSavedLoginStatusText.Text = _localization.GetString("SavedAQuickLoginButtonWasAddedToTheMainScreen");
        TtacSavedLoginStatusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x05, 0x96, 0x69));
        TtacSavedLoginStatusText.Visibility = Visibility.Visible;

        RefreshTtacSavedLoginsList();
        RefreshTtacQuickLoginButtons();
    }

    private void TtacSavedLoginEditButton_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not TtacSavedLogin login)
            return;

        // نام کاربری و نام داروخانه را در فرم پر کن تا کاربر بتواند ویرایش کند.
        // رمز عبور به‌صورت رمزنگاری‌شده ذخیره شده و قابل نمایش نیست؛ کاربر باید دوباره وارد کند.
        TtacSavedLoginUsernameTextBox.Text = login.Username;
        TtacSavedLoginPharmacyNameTextBox.Text = login.PharmacyName;
        TtacSavedLoginPasswordBox.Password = string.Empty;
        TtacSavedLoginPasswordTextBox.Text = string.Empty;
        // رمز عبور را به حالت مخفی برگردان
        _isPasswordVisible = false;
        TtacSavedLoginPasswordBox.Visibility = Visibility.Visible;
        TtacSavedLoginPasswordTextBox.Visibility = Visibility.Collapsed;
        UpdateTtacSavedLoginPasswordToggleIcon(false);
        TtacSavedLoginPasswordBox.Focus();

        TtacSavedLoginStatusText.Text = _localization.GetFormattedString("EditingLogin", login.DisplayLabel);
        TtacSavedLoginStatusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0));
        TtacSavedLoginStatusText.Visibility = Visibility.Visible;
    }

    private void TtacSavedLoginDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        if ((sender as FrameworkElement)?.DataContext is not TtacSavedLogin login)
            return;

        RemoveSavedTtacLogin(login.Username);
        TtacSavedLoginStatusText.Text = _localization.GetFormattedString("LoginDeleted", login.DisplayLabel);
        TtacSavedLoginStatusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x6B, 0x72, 0x80));
        TtacSavedLoginStatusText.Visibility = Visibility.Visible;

        RefreshTtacSavedLoginsList();
        RefreshTtacQuickLoginButtons();
    }

    // بعد از هر ورود موفق (چه با فرم مستقیم، چه با مرورگر داخلی) صدا زده می‌شود تا اگر این
    // نام‌کاربری قبلاً «به خاطر سپرده» شده، اسم داروخانه‌اش برای دفعه‌ی بعد به‌روز بماند.
    private void UpdateSavedLoginPharmacyName(string username, string pharmacyName)
    {
        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(pharmacyName))
            return;

        try
        {
            var logins = LoadSavedTtacLogins();
            var existing = logins.FirstOrDefault(x => string.Equals(x.Username, username, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
                return;

            existing.PharmacyName = pharmacyName;
            SaveTtacLoginsList(logins);
        }
        catch { }
    }

    // آخرین ورودِ ذخیره‌شده را (اگر باشد) داخل فرم پر می‌کند - برای ورود سریع دفعه‌ی بعد.
    private void PrefillTtacLoginFields()
    {
        try
        {
            var logins = LoadSavedTtacLogins();
            var latest = logins.OrderByDescending(x => x.LastUsedUtc).FirstOrDefault();
            if (latest == null)
                return;

            TtacLoginUsernameTextBox.Text = latest.Username;
            byte[] encrypted = Convert.FromBase64String(latest.PasswordProtectedBase64);
            byte[] plain = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
            TtacLoginPasswordBox.Password = Encoding.UTF8.GetString(plain);
            TtacLoginRememberCheckBox.IsChecked = true;
            UpdateTtacLoginUsernameHint();
        }
        catch { }
    }

    private void SaveTtacLoginIfRequested(string username, string password)
    {
        if (TtacLoginRememberCheckBox.IsChecked != true)
        {
            RemoveSavedTtacLogin(username);
            return;
        }

        UpsertSavedTtacLogin(username, password, null);
    }

    // وقتی کاربر داخل کادر «نام کاربری» تایپ می‌کند، اگر این کد قبلاً ذخیره شده، اسم
    // داروخانه‌اش را زیر کادر نشان می‌دهد؛ اگر دقیقاً با یک ورود ذخیره‌شده یکی باشد و کادر
    // رمز عبور هنوز خالی باشد، رمز را هم خودکار پر می‌کند (مثل autofill مرورگر).
    private void TtacLoginUsernameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateTtacLoginUsernameHint();
    }

    private void UpdateTtacLoginUsernameHint()
    {
        if (TtacLoginUsernameHintText == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        string typed = TtacLoginUsernameTextBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(typed))
        {
            TtacLoginUsernameHintText.Visibility = Visibility.Collapsed;
            return;
        }

        var logins = LoadSavedTtacLogins();
        var exact = logins.FirstOrDefault(x => string.Equals(x.Username, typed, StringComparison.OrdinalIgnoreCase));
        if (exact != null)
        {
            if (string.IsNullOrEmpty(TtacLoginPasswordBox.Password))
            {
                try
                {
                    byte[] encrypted = Convert.FromBase64String(exact.PasswordProtectedBase64);
                    byte[] plain = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                    TtacLoginPasswordBox.Password = Encoding.UTF8.GetString(plain);
                    TtacLoginRememberCheckBox.IsChecked = true;
                }
                catch { }
            }

            TtacLoginUsernameHintText.Text = !string.IsNullOrWhiteSpace(exact.PharmacyName)
                ? (_localization.GetFormattedString("CodeUsedFor", exact.PharmacyName))
                : (_localization.GetString("ThisCodeIsSaved"));
            TtacLoginUsernameHintText.Visibility = Visibility.Visible;
            return;
        }

        var prefixMatches = logins
            .Where(x => x.Username.StartsWith(typed, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (prefixMatches.Count > 0)
        {
            string names = string.Join("، ", prefixMatches
                .Select(x => string.IsNullOrWhiteSpace(x.PharmacyName) ? x.Username : x.PharmacyName)
                .Distinct());
            TtacLoginUsernameHintText.Text = _localization.GetFormattedString("SavedCodesMatching", names);
            TtacLoginUsernameHintText.Visibility = Visibility.Visible;
        }
        else
        {
            TtacLoginUsernameHintText.Visibility = Visibility.Collapsed;
        }
    }

    // ---------- پنجره‌ی انتخاب داروخانه برای ورود به تی‌تک ----------

    // با کلیک روی «ورود به سایت تی‌تک»، این پنجره‌ی کوچک (با بلور پشتش) باز می‌شود و به ازای هر
    // حساب ذخیره‌شده در تنظیمات یک دکمه‌ی «ورود به داروخانه ...» نشان می‌دهد. کلیک روی هرکدام،
    // مرورگر داخلی تی‌تک را با همان حساب (اجباری) باز می‌کند تا فقط دکمه‌ی «ورود به سیستم» سایت
    // را بزنی. اگر هیچ حسابی ذخیره نشده باشد، این پنجره اصلاً باز نمی‌شود و مرورگر مستقیم بالا
    // می‌آید (رفتار قبلی).
    private void RefreshTtacQuickLoginButtons()
    {
        if (TtacQuickLoginButtonsPanel == null)
            return;

        TtacQuickLoginButtonsPanel.Children.Clear();
        var logins = LoadSavedTtacLogins();
        if (logins.Count == 0)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        foreach (var login in logins.OrderByDescending(x => x.LastUsedUtc))
        {
            string label = !string.IsNullOrWhiteSpace(login.PharmacyName) ? login.PharmacyName : login.Username;
            // اگر نام ذخیره‌شده خودش با «داروخانه» شروع شده، دوباره جلوش «داروخانه» نگذاریم.
            string pharmacyPart = label;
            if (!english && !pharmacyPart.Contains("داروخانه", StringComparison.Ordinal))
                pharmacyPart = "داروخانه " + pharmacyPart;
            string content = _localization.GetFormattedString("LoginTo", label, pharmacyPart);

            var button = new System.Windows.Controls.Button
            {
                Content = content,
                Tag = login.Username,
                Height = 54,
                Margin = new Thickness(0, 6, 0, 6),
                FontSize = 16,
                Style = (Style)FindResource("RoundedButtonStyle"),
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0)),
                ToolTip = _localization.GetFormattedString("OpenTtacLoginPage", login.Username)
            };
            button.Click += TtacQuickLoginButton_Click;
            TtacQuickLoginButtonsPanel.Children.Add(button);
        }
    }

    private bool ShowTtacQuickLoginOverlay(bool sessionExpired = false)
    {
        RefreshTtacQuickLoginButtons();
        if (TtacQuickLoginButtonsPanel.Children.Count == 0)
            return false;

        TtacQuickLoginOverlayTitle.Text = _localization.GetString("TTACLogin");
        TtacQuickLoginOverlaySubtitle.Text = _localization.GetString("TtacQuickLoginSubtitle");
        TtacQuickLoginCancelButton.Content = _localization.GetString("CancelButton");
        if (TtacQuickLoginWarningBorder != null)
        {
            if (sessionExpired)
            {
                // اگر نشست وسط یک عملیات منقضی شده، نام همان عملیات (که بعد از ورود خودکار ادامه
                // پیدا می‌کند) را داخل هشدار نشان بده.
                string warningText;
                if (!string.IsNullOrWhiteSpace(_pendingTtacRetryLabel))
                    warningText = _localization.GetFormattedString("TtacSessionExpiredWithPending", _pendingTtacRetryLabel);
                else if (_pendingTtacRetryAction != null)
                    warningText = _localization.GetString("TtacSessionExpiredPendingGeneric");
                else
                    warningText = _localization.GetString("TtacSessionExpiredWarning");
                TtacQuickLoginWarningText.Text = warningText;
                TtacQuickLoginWarningBorder.Visibility = Visibility.Visible;
            }
            else
            {
                TtacQuickLoginWarningBorder.Visibility = Visibility.Collapsed;
            }
        }
        TtacQuickLoginOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        return true;
    }

    private void CloseTtacQuickLoginOverlay()
    {
        TtacQuickLoginOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = TtTeckSettingsOverlay.Visibility == Visibility.Visible
            || TtTeckWebViewOverlay.Visibility == Visibility.Visible
            || TtacPanelOverlay.Visibility == Visibility.Visible
            || CargoDeliveryOverlay.Visibility == Visibility.Visible
            || ReceiveStatusOverlay.Visibility == Visibility.Visible
            || TtTeckRegistrationOverlay.Visibility == Visibility.Visible
            || HistoryOverlay.Visibility == Visibility.Visible
            ? new System.Windows.Media.Effects.BlurEffect { Radius = 18 }
            : null;
    }

    private void TtacQuickLoginOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseTtacQuickLoginOverlay();
    }

    private void TtacQuickLoginCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void TtacQuickLoginCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseTtacQuickLoginOverlay();
    }

    private async void TtacQuickLoginButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var username = (sender as FrameworkElement)?.Tag as string;
            if (string.IsNullOrWhiteSpace(username))
                return;

            CloseTtacQuickLoginOverlay();

            _pendingTtacAutofillUsername = username;
            _ttacRetryUsername = username;
            _ttacQuickLoginInProgress = true;

            // اگر همین داروخانه از قبل وارد است، دوباره مرورگر را باز نکن.
            if (HasValidTtacToken() && TokenMatchesRequestedPharmacy(_ttacAccessTokenOverride, username))
            {
                _ttacQuickLoginInProgress = false;
                ShowTtacLoginSuccessBanner();
                return;
            }

            // توکن/کوکی داروخانه‌ی قبلی را پاک کن تا همان نشست فوری بسته نشود.
            if (HasValidTtacToken() || _ttTeckWebView?.CoreWebView2 != null)
                await ClearTtacWebViewSessionForSwitchAsync();

            await OpenTtTeckInternalBrowserAsync("https://newstatisticsreports.ttac.ir/pharmacyDashboard");
            _ = MonitorTtacConnectionAfterBrowserOpenAsync();
        }
        catch { }
    }

    // ---------- حساب‌های ذخیره‌شده‌ی تی‌تک (تنظیمات) ----------
    // این بخش کاملاً مستقل از فرم قدیمی «ورود مستقیم» (که مخفی مانده و دیگر استفاده نمی‌شود) است:
    // یک‌بار (نه هر بار که کاربر تایپ می‌کند) موقع کامل شدن ناوبری در مرورگر داخلی تی‌تک صدا زده
    // می‌شود - فقط وقتی هنوز توکن معتبری به‌دست نیامده (یعنی احتمالاً صفحه‌ی ورود است). برخلاف
    // تلاش قبلی (که به‌طور کامل حذف شد چون دکمه‌ی ورود صفحه را خراب می‌کرد)، این نسخه هیچ رویداد
    // submit/click صفحه را قبضه نمی‌کند و هیچ اطلاعاتی از صفحه به بیرون نمی‌فرستد - فقط یک‌بار
    // مقدار دو فیلد را می‌گذارد و یک شنونده‌ی ساده روی «تایپ در فیلد یوزرنیم» اضافه می‌کند تا اگر
    // کاربر بین چند حساب ذخیره‌شده جابه‌جا شد، رمز هم خودش را با آن هماهنگ کند.
    private async Task TryAutofillTtacWebViewLoginAsync(string? forceUsername = null)
    {
        try
        {
            var core = TtTeckWebView?.CoreWebView2;
            if (core == null)
                return;

            // اگر کاربر از صفحه‌ی اصلی روی دکمه‌ی یک داروخانه‌ی خاص کلیک کرده، همان حساب را
            // اجباری پر کن (حتی اگر فیلد یوزرنیم از قبل توسط خودِ مرورگر مقدار داشته باشد).
            if (string.IsNullOrWhiteSpace(forceUsername) && !string.IsNullOrWhiteSpace(_pendingTtacAutofillUsername))
                forceUsername = _pendingTtacAutofillUsername;

            var savedLogins = LoadSavedTtacLogins();
            if (savedLogins.Count == 0)
            {
                AppendTtacWebViewDebugLog("Autofill: no saved logins on disk.");
                return;
            }

            var credentials = new List<TtacAutofillCredential>();
            string bestGuessUsername = string.Empty;
            DateTime bestGuessTime = DateTime.MinValue;
            foreach (var login in savedLogins)
            {
                try
                {
                    byte[] encrypted = Convert.FromBase64String(login.PasswordProtectedBase64);
                    byte[] plain = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                    credentials.Add(new TtacAutofillCredential { U = login.Username, P = Encoding.UTF8.GetString(plain) });
                    if (login.LastUsedUtc > bestGuessTime)
                    {
                        bestGuessTime = login.LastUsedUtc;
                        bestGuessUsername = login.Username;
                    }
                }
                catch { }
            }

            if (credentials.Count == 0)
            {
                AppendTtacWebViewDebugLog($"Autofill: decryption of {savedLogins.Count} saved login(s) failed or empty.");
                return;
            }

            // اگر کاربر روی دکمه‌ی یک داروخانه‌ی خاص کلیک کرده، آن حساب را اجباری پر کن؛
            // فقط اگر واقعاً در میان حساب‌های ذخیره‌شده باشد.
            bool force = false;
            if (!string.IsNullOrWhiteSpace(forceUsername))
            {
                if (credentials.Any(c => string.Equals(c.U, forceUsername, StringComparison.OrdinalIgnoreCase)))
                {
                    bestGuessUsername = forceUsername;
                    force = true;
                }
                else
                {
                    forceUsername = null;
                }
            }

            // مهم: اسکریپت جاوااسکریپت کلیدهای u/p (حروف کوچک) را می‌خواند؛ پس حتماً با
            // PropertyNamingPolicy.CamelCase سریالایز می‌کنیم - وگرنه کلیدها U/P (حرف بزرگ)
            // می‌شوند و اسکریپت موقع مقایسه‌ی کد با خطای toLowerCase روی undefined از کار می‌افتد
            // (این دقیقاً همان باگی بود که باعث می‌شد رمز هیچ‌وقت پر نشود).
            string credentialsJson = JsonSerializer.Serialize(credentials, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            string bestGuessJson = JsonSerializer.Serialize(bestGuessUsername);
            string forceJson = JsonSerializer.Serialize(force);
            try
            {
                string result = await core.ExecuteScriptAsync(BuildTtacLoginAutofillScript(credentialsJson, bestGuessJson, forceJson));
                AppendTtacWebViewDebugLog($"Autofill script result => {result}");

                if (force && TtTeckWebView?.Source?.Host?.Equals("idp.ttac.ir", StringComparison.OrdinalIgnoreCase) == true)
                {
                    // روی خودِ صفحه‌ی ورود مصرف شد؛ دیگر لازم نیست اجباری بماند.
                    _pendingTtacAutofillUsername = null;
                }
            }
            catch (Exception ex)
            {
                AppendTtacWebViewDebugLog($"Autofill ExecuteScriptAsync error: {ex.Message}");
            }
        }
        catch { }
    }

    private sealed class TtacAutofillCredential
    {
        public string U { get; set; } = string.Empty;
        public string P { get; set; } = string.Empty;
    }

    private static string BuildTtacLoginAutofillScript(string credentialsJson, string bestGuessUsernameJson, string forceJson)
    {
        return @"
(function() {
  try {
    var __log = (window.__sbAutofillLog = window.__sbAutofillLog || []);
    function __push(msg) { try { __log.push(msg); } catch (e) {} }

    function setNativeValue(el, value) {
      try {
        var proto = Object.getPrototypeOf(el);
        var desc = Object.getOwnPropertyDescriptor(proto, 'value') || Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
        if (desc && desc.set) { desc.set.call(el, value); } else { el.value = value; }
      } catch (e1) {
        try { el.value = value; } catch (e2) { return false; }
      }
      try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch (e3) {}
      try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch (e4) {}
      return true;
    }

    // بررسی دقیق‌تر «دیده شدن» فیلد - فقط offsetParent کافی نیست، چون فیلدهای طعمه‌ی ضدربات
    // (honeypot) هم می‌توانند offsetParent داشته باشند ولی با opacity:0 یا اندازه‌ی صفر مخفی
    // باشند؛ چنین فیلدهایی هرگز نباید هدف پرکردن خودکار قرار بگیرند.
    function isReallyVisible(el) {
      if (el.offsetParent === null || el.disabled) return false;
      var rect = el.getBoundingClientRect();
      if (rect.width < 4 || rect.height < 4) return false;
      var style = window.getComputedStyle(el);
      if (style.visibility === 'hidden' || style.display === 'none' || parseFloat(style.opacity || '1') < 0.05) return false;
      return true;
    }

    var credentials = " + credentialsJson + @";
    var bestGuessUsername = " + bestGuessUsernameJson + @";
    // وقتی کاربر از صفحه‌ی اصلی یک داروخانه‌ی خاص را انتخاب کرده، این پرچم true است و باید
    // همان حساب (حتی اگر فیلد یوزرنیم از قبل مقداری داشته باشد) پر شود.
    var forceUsername = " + forceJson + @";
    if (!credentials || credentials.length === 0) return 'no-credentials';

    // در برابر هر دو شکل کلید (u/U یا p/P) مقاوم باش - اگه کلیدی نبود/تهی بود، آن ورودی را رد کن.
    credentials = credentials.map(function (c) {
      return { u: c && (c.u != null ? c.u : c.U), p: c && (c.p != null ? c.p : c.P) };
    }).filter(function (c) { return c && c.u && c.p; });
    if (credentials.length === 0) return 'no-credentials';

    var userField = null;
    var passField = null;
    var attachedUser = null;
    var attachedPass = null;
    var lastAutoFilledPassword = null;
    var passwordUserTouched = false;
    // به محض اینکه کاربر روی هر کدام از فیلدها کلیک کند یا تایپ کند، دیگر چیزی را
    // بازنویسی نمی‌کنیم تا بتواند بین حساب‌های ذخیره‌شده جابه‌جا شود.
    var userActive = false;

    function findMatch(typed) {
      if (!typed) return null;
      var t = typed.trim().toLowerCase();
      for (var k = 0; k < credentials.length; k++) {
        if (credentials[k].u.toLowerCase() === t) return credentials[k];
      }
      return null;
    }

    // فیلدها را نسبت به دکمه‌ی ورود (متن «ورود») پیدا می‌کنیم، نه صرفاً اولین فیلد پسورد کل صفحه
    // - چون فرم‌های دولتی با محافظت ضدربات گاهی فیلدهای طعمه‌ی دیگری هم جای دیگری از صفحه دارند.
    function locateFields() {
      // اول سراغ فیلدهای مشخص خودِ سایت برو (id/name استاندارد) - مطمئن‌تر از حدس زدن
      var byIdU = document.getElementById('username') || document.querySelector('input[name=username]');
      var byIdP = document.getElementById('password') || document.querySelector('input[name=password]');
      if (byIdU && byIdP && isReallyVisible(byIdU) && isReallyVisible(byIdP)) {
        return { u: byIdU, p: byIdP };
      }

      var inputs = Array.prototype.slice.call(document.querySelectorAll('input'));
      var visible = inputs.filter(isReallyVisible);

      var allClickable = Array.prototype.slice.call(document.querySelectorAll('button, input[type=submit], input[type=button], a'));
      var loginButton = null;
      for (var bi = 0; bi < allClickable.length; bi++) {
        var el = allClickable[bi];
        var text = (el.innerText || el.value || el.textContent || '').trim();
        if (text.indexOf('ورود') !== -1 && isReallyVisible(el)) { loginButton = el; break; }
      }

      var searchSpace = visible;
      if (loginButton) {
        var beforeButton = visible.filter(function (el) {
          return (el.compareDocumentPosition(loginButton) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0;
        });
        if (beforeButton.length > 0) searchSpace = beforeButton;
      }

      var pwdIdx = -1;
      for (var i = searchSpace.length - 1; i >= 0; i--) {
        if ((searchSpace[i].getAttribute('type') || '').toLowerCase() === 'password') { pwdIdx = i; break; }
      }
      if (pwdIdx === -1) return null;

      var p = searchSpace[pwdIdx];
      var u = null;
      for (var j = pwdIdx - 1; j >= 0; j--) {
        var t = (searchSpace[j].getAttribute('type') || 'text').toLowerCase();
        if (t === 'text' || t === 'email' || t === 'tel' || t === 'number' || t === 'search' || t === '') { u = searchSpace[j]; break; }
      }
      if (!u) return null;
      return { u: u, p: p };
    }

    function applyFill(match) {
      var okU = true;
      var okP = true;
      if ((userField.value || '') !== match.u) okU = setNativeValue(userField, match.u);
      okP = setNativeValue(passField, match.p);
      if (!okU) __push('setNativeValue(USER) failed');
      if (!okP) __push('setNativeValue(PASS) failed');
      lastAutoFilledPassword = match.p;
      passwordUserTouched = false;
      return okU && okP;
    }

    function userInputListener() {
      var match = findMatch(userField.value);
      if (match) {
        applyFill(match);
      } else if (lastAutoFilledPassword !== null && passField.value === lastAutoFilledPassword) {
        setNativeValue(passField, '');
        lastAutoFilledPassword = null;
      }
    }

    // شنونده‌ها را فقط یک‌بار به هر فیلد وصل می‌کنیم (اگر صفحه فیلد را عوض کرد، برای فیلد جدید
    // دوباره وصل می‌شود).
    function attachListeners() {
      if (attachedUser !== userField) {
        userField.setAttribute('data-scanbridge-autofill', '1');
        userField.addEventListener('input', userInputListener);
        userField.addEventListener('focus', function () { userActive = true; });
        attachedUser = userField;
      }
      if (attachedPass !== passField) {
        passField.addEventListener('input', function () {
          if (passField.value !== lastAutoFilledPassword) passwordUserTouched = true;
        });
        passField.addEventListener('focus', function () { userActive = true; });
        attachedPass = passField;
      }
    }

    // یک بار تلاش برای پیدا کردن فرم و پر کردن آن. اگر فرم هنوز رندر نشده (مثلاً Angular بعد از
    // لود مدل آن را نشان می‌دهد) یا فیلدها مخفی باشند، false برمی‌گرداند و فراخوان صبر می‌کند و
    // دوباره تلاش می‌کند.
    function tryFill() {
      try {
        if (passwordUserTouched || userActive) { __push('user-active, stop'); return true; }
        var loc = locateFields();
        if (!loc) {
          __push('no-visible-fields');
          return false;
        }
        userField = loc.u;
        passField = loc.p;
        attachListeners();
        __push('located u=' + (loc.u.id || loc.u.name) + ' p=' + (loc.p.id || loc.p.name) + ' uValLen=' + (userField.value || '').length + ' pValLen=' + (passField.value || '').length);

        var currentUser = (userField.value || '').trim();
        var match = null;
        if (forceUsername && bestGuessUsername) match = findMatch(bestGuessUsername);
        if (!match && currentUser) match = findMatch(currentUser);
        if (!match && !currentUser && bestGuessUsername) match = findMatch(bestGuessUsername);
        if (!match) {
          __push('no-match for user=' + currentUser);
          return false;
        }

        if (lastAutoFilledPassword !== match.p || (passField.value || '') !== match.p) {
          var ok = applyFill(match);
          __push('applyFill user=' + match.u + ' -> ' + (ok ? 'ok' : 'partial-fail') + ' pValLenNow=' + (passField.value || '').length);
        } else {
          __push('already-filled');
        }
        return true;
      } catch (e) {
        __push('tryFill-error: ' + (e && e.message ? e.message : String(e)));
        return false;
      }
    }

    tryFill();

    // اگر فرم هنوز ظاهر نشده یا صفحه مقدار فیلد پسورد را بعداً خالی کرد، تا ۶۰ ثانیه (هر ۵۰۰ms)
    // دوباره تلاش می‌کنیم. به محض اینکه کاربر خودش در فیلدها تایپ کند، دست نگه می‌داریم.
    var tries = 0;
    var timer = setInterval(function () {
      tries++;
      if (tries > 120) { clearInterval(timer); return; }
      tryFill();
    }, 500);

    // اگر صفحه فیلدها را حذف/دوباره بسازد، بلافاصله (بدون انتظار برای تیک بعدی تایمر) دوباره
    // تلاش می‌کنیم.
    try {
      if (window.MutationObserver) {
        var mo = new MutationObserver(function () { tryFill(); });
        mo.observe(document.body || document.documentElement, { childList: true, subtree: true });
      }
    } catch (e) {}

    return 'started';
  } catch (e) {
    return 'error:' + (e && e.message ? e.message : String(e));
  }
})();
";
    }

    private void ShowTtacLoginOverlay(bool sessionExpired = false)
    {
        // ۱) اگر حساب ذخیره‌شده‌ای هست، پنجره‌ی انتخاب داروخانه را نشان بده تا کاربر فقط
        //    داروخانه را انتخاب کند و کد/رمز آن خودکار پر شود. اگر نشست منقضی شده، یک
        //    هشدار هم داخل همین پنجره نمایش داده می‌شود.
        if (ShowTtacQuickLoginOverlay(sessionExpired))
            return;

        // ۲) اگر هیچ حسابی ذخیره نشده، به‌جای پیام قدیمی «مرورگر داخلی»، مستقیم فرم ذخیره‌ی
        //    حساب تی‌تک را داخل تنظیمات باز کن و به همان بخش اسکرول کن.
        OpenTtTeckSettings();
        if (TtacSavedLoginStatusText != null)
        {
            TtacSavedLoginStatusText.Text = _localization.GetString("SaveAccountToLoginFirst");
            TtacSavedLoginStatusText.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xDC, 0x26, 0x26));
            TtacSavedLoginStatusText.Visibility = Visibility.Visible;
        }
        if (TtacLoginWithoutSavingHint != null)
        {
            TtacLoginWithoutSavingHint.Text = _localization.GetString("LoginWithoutSavingHint");
        }
        Dispatcher.BeginInvoke(new Action(() =>
        {
            TtacSavedLoginsSection?.BringIntoView();
            TtacSavedLoginUsernameTextBox?.Focus();
        }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
    }

    private void CloseTtacLoginOverlay()
    {
        TtacLoginOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void TtacLoginOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseTtacLoginOverlay();
    }

    private void TtacLoginCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void TtacLoginCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseTtacLoginOverlay();
    }

    private async void TtacLoginOpenBrowserButton_Click(object sender, RoutedEventArgs e)
    {
        CloseTtacLoginOverlay();
        await OpenTtTeckInternalBrowserAsync("https://newstatisticsreports.ttac.ir/pharmacyDashboard");
        _ = MonitorTtacConnectionAfterBrowserOpenAsync();
    }

    private async void TtacLoginSubmitButton_Click(object sender, RoutedEventArgs e)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        string username = TtacLoginUsernameTextBox.Text?.Trim() ?? string.Empty;
        string password = TtacLoginPasswordBox.Password ?? string.Empty;

        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
        {
            TtacLoginStatusText.Text = _localization.GetString("EnterUsernameAndPasswordOrUseTheInternalBrowser");
            return;
        }

        TtacLoginSubmitButton.IsEnabled = false;
        TtacLoginOpenBrowserButton.IsEnabled = false;
        TtacLoginStatusText.Text = _localization.GetString("LoggingIn");

        try
        {
            TtacTokenResult tokenResult = await LoginToTtacDirectAsync(username, password);
            DateTime expiresAtUtc = DateTime.UtcNow.AddSeconds(Math.Max(60, tokenResult.ExpiresInSeconds - 60));

            ApplyTtacAccessToken(tokenResult.AccessToken, expiresAtUtc);
            SaveTtacLoginIfRequested(username, password);
            UpdateSavedLoginPharmacyName(username, _ttacPharmacyDisplayName);

            CloseTtacLoginOverlay();
            UpdateTtacConnectionStatusUI();

            if (_pendingTtacRetryAction != null)
                await RunPendingTtacRetryIfAnyAsync();
        }
        catch (Exception ex)
        {
            // ورود مستقیم شکست خورد؛ کاربر می‌تواند دوباره تلاش کند یا از مرورگر داخلی استفاده کند
            // (که همیشه کار می‌کند چون رندر واقعی سایت تی‌تک است).
            TtacLoginStatusText.Text = ex.Message;
        }
        finally
        {
            TtacLoginSubmitButton.IsEnabled = true;
            TtacLoginOpenBrowserButton.IsEnabled = true;
        }
    }

    private string? TryExtractTtacDisplayNameFromToken(string? jwt)
    {
        if (string.IsNullOrWhiteSpace(jwt))
            return null;

        try
        {
            string[] parts = jwt.Split('.');
            if (parts.Length < 2)
                return null;

            string payload = parts[1]
                .Replace('-', '+')
                .Replace('_', '/');
            payload = payload.PadRight(payload.Length + ((4 - payload.Length % 4) % 4), '=');
            string json = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            string[] preferredClaims =
            {
                "companyName",
                "pharmacyName",
                "organizationName",
                "orgName",
                "name",
                "given_name",
                "family_name",
                "preferred_username"
            };

            foreach (string claim in preferredClaims)
            {
                string? value = ReadJsonString(root, claim);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    if (claim == "given_name")
                    {
                        string? familyName = ReadJsonString(root, "family_name");
                        if (!string.IsNullOrWhiteSpace(familyName))
                            value = value + " " + familyName;
                    }

                    return value.Trim();
                }
            }
        }
        catch { }

        return null;
    }

    private static string? ReadJwtClaim(string? jwt, params string[] names)
    {
        if (string.IsNullOrWhiteSpace(jwt) || names == null || names.Length == 0)
            return null;
        try
        {
            string[] parts = jwt.Split('.');
            if (parts.Length < 2)
                return null;
            string payload = parts[1].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + ((4 - payload.Length % 4) % 4), '=');
            string json = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
            using var doc = JsonDocument.Parse(json);
            foreach (string name in names)
            {
                if (doc.RootElement.TryGetProperty(name, out var el))
                {
                    string? value = el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value.Trim();
                }
            }
        }
        catch { }
        return null;
    }

    private bool TokenMatchesRequestedPharmacy(string? token, string? username)
    {
        if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(username))
            return false;

        string requested = username.Trim();
        string? claimUser = ReadJwtClaim(token, "preferred_username", "unique_name", "username", "userName", "nameid", "sub", "name", "pharmacyCode", "gln");
        if (!string.IsNullOrWhiteSpace(claimUser) && string.Equals(claimUser, requested, StringComparison.OrdinalIgnoreCase))
            return true;

        try
        {
            var saved = LoadSavedTtacLogins()
                .FirstOrDefault(x => string.Equals(x.Username, requested, StringComparison.OrdinalIgnoreCase));
            string? display = TryExtractTtacDisplayNameFromToken(token);
            if (saved != null && !string.IsNullOrWhiteSpace(saved.PharmacyName) && !string.IsNullOrWhiteSpace(display))
            {
                if (display.Contains(saved.PharmacyName, StringComparison.OrdinalIgnoreCase)
                    || saved.PharmacyName.Contains(display, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch { }

        return false;
    }

    private bool ShouldIgnoreStaleTtacToken()
        => _ttacWaitingForFreshLogin && !_ttacSawIdpLoginPage;

    private void MarkTtacIdpLoginSeen(string? url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || string.IsNullOrWhiteSpace(uri.Host))
            return;
        if (uri.Host.EndsWith("idp.ttac.ir", StringComparison.OrdinalIgnoreCase))
            _ttacSawIdpLoginPage = true;
    }

    private async Task ClearTtacWebViewSessionForSwitchAsync()
    {
        try
        {
            if (HasValidTtacToken())
            {
                SaveCurrentReceiveStatusItemsForCurrentPharmacy();
                SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
                SaveCurrentHistoryForCurrentPharmacy();
                SaveTtacRegistrationHistory();
            }
        }
        catch { }

        _ttacAccessTokenOverride = null;
        _ttacAccessTokenExpiresAtUtc = DateTime.MinValue;
        UpdateTtacTokenValidityTracking(false, suppressExpiryNotification: true);
        _ttacLoginSuccessHandled = false;
        _ttacSawIdpLoginPage = false;
        _ttacWaitingForFreshLogin = true;
        UpdateTtacConnectionStatusUI();

        try
        {
            if (_ttTeckWebView?.CoreWebView2 != null)
            {
                try { await _ttTeckWebView.CoreWebView2.ExecuteScriptAsync("try{localStorage.clear();sessionStorage.clear();}catch(e){}"); } catch { }
                try { _ttTeckWebView.CoreWebView2.CookieManager.DeleteAllCookies(); } catch { }
                try { _ttTeckWebView.CoreWebView2.Navigate("about:blank"); } catch { }
                await Task.Delay(400);
            }
        }
        catch { }
    }

    private sealed class TtacTokenResult
    {
        public string AccessToken { get; set; } = string.Empty;
        public int ExpiresInSeconds { get; set; } = 7200;
    }

    // حداکثر تعداد ریدایرکتی که دنبال می‌کنیم؛ چون در برخی حالت‌ها (مثلاً صفحه‌ی تایید/consent
    // بین صفحه‌ی ورود و بازگشت نهایی) ممکن است بیش از چند مرحله طول بکشد، این عدد را بالاتر از
    // قبل (12) گذاشتیم.
    private const int TtacLoginMaxRedirects = 20;

    private async Task<TtacTokenResult> LoginToTtacDirectAsync(string username, string password)
    {
        var handler = new HttpClientHandler
        {
            AllowAutoRedirect = false,
            UseCookies = true,
            CookieContainer = new System.Net.CookieContainer()
        };

        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(45) };
        client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36");
        client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");

        string state = Guid.NewGuid().ToString("N");
        string nonce = Guid.NewGuid().ToString("N");
        string authorizeUrl = "https://idp.ttac.ir/identity/connect/authorize" +
                              "?client_id=Statistics" +
                              "&redirect_uri=https%3A%2F%2Fnewstatisticsreports.ttac.ir%2Fcallback" +
                              "&response_type=id_token%20token" +
                              "&scope=profile%20openid%20roles%20demo-website" +
                              $"&state={state}&nonce={nonce}";

        Uri currentUri = new Uri(authorizeUrl);
        HttpResponseMessage response = await client.GetAsync(currentUri);
        string html = string.Empty;

        for (int i = 0; i < TtacLoginMaxRedirects; i++)
        {
            if (IsRedirect(response.StatusCode) && response.Headers.Location != null)
            {
                Uri nextUri = response.Headers.Location.IsAbsoluteUri
                    ? response.Headers.Location
                    : new Uri(currentUri, response.Headers.Location);

                if (TryExtractTokenFromUri(nextUri, out var tokenResult))
                    return tokenResult;

                currentUri = nextUri;
                response = await client.GetAsync(currentUri);
                continue;
            }

            html = await response.Content.ReadAsStringAsync();
            break;
        }

        if (string.IsNullOrWhiteSpace(html))
            throw new InvalidOperationException(_localization.GetString("TTACLoginPageWasNotReceived"));

        // همه‌ی فیلدهای پنهان صفحه (نه فقط idsrv.xsrf) استخراج می‌شوند؛ چون بعضی وقت‌ها این
        // صفحه فیلدهای پنهان دیگری هم دارد (مثل ReturnUrl) که اگر ارسال نشوند سرور درخواست را رد
        // می‌کند و همین می‌توانست دلیل ورود ناموفق قبلی باشد.
        var form = ExtractAllHiddenInputs(html);
        if (!form.ContainsKey("idsrv.xsrf") || string.IsNullOrWhiteSpace(form["idsrv.xsrf"]))
            throw new InvalidOperationException(_localization.GetString("TheLoginPageSecurityTokenWasNotFoundPleaseUseInternalBrowserLogin"));

        form["username"] = username;
        form["password"] = password;

        var post = new HttpRequestMessage(HttpMethod.Post, currentUri)
        {
            Content = new FormUrlEncodedContent(form)
        };
        post.Headers.Referrer = currentUri;
        response = await client.SendAsync(post);

        string lastBody = string.Empty;
        for (int i = 0; i < TtacLoginMaxRedirects; i++)
        {
            if (IsRedirect(response.StatusCode) && response.Headers.Location != null)
            {
                Uri nextUri = response.Headers.Location.IsAbsoluteUri
                    ? response.Headers.Location
                    : new Uri(currentUri, response.Headers.Location);

                if (TryExtractTokenFromUri(nextUri, out var tokenResult))
                    return tokenResult;

                currentUri = nextUri;
                response = await client.GetAsync(currentUri);
                continue;
            }

            lastBody = await response.Content.ReadAsStringAsync();

            // برخی سرویس‌های OpenID Connect بین صفحه‌ی ورود و بازگشت نهایی یک صفحه‌ی «اجازه
            // دسترسی/consent» نشان می‌دهند که باید خودکار تایید شود تا توکن صادر شود. اگر چنین
            // فرمی دیدیم (فیلدهای ورود دیگر در صفحه نیستند ولی فرم دیگری با دکمه‌ی تایید هست)
            // آن را هم خودکار ارسال می‌کنیم؛ همین می‌توانست یکی از دلایل ورود ناموفق قبلی باشد.
            if (!lastBody.Contains("name=\"password\"", StringComparison.OrdinalIgnoreCase) &&
                !lastBody.Contains("name='password'", StringComparison.OrdinalIgnoreCase) &&
                TryBuildConsentForm(lastBody, out var consentForm))
            {
                var consentPost = new HttpRequestMessage(HttpMethod.Post, currentUri)
                {
                    Content = new FormUrlEncodedContent(consentForm)
                };
                consentPost.Headers.Referrer = currentUri;
                response = await client.SendAsync(consentPost);
                continue;
            }

            break;
        }

        string bodyLower = lastBody ?? string.Empty;
        bool looksLikeCredentialError =
            bodyLower.Contains("نام کاربری یا کلمه عبور", StringComparison.OrdinalIgnoreCase) ||
            bodyLower.Contains("نام کاربری یا رمز عبور", StringComparison.OrdinalIgnoreCase) ||
            bodyLower.Contains("invalid username or password", StringComparison.OrdinalIgnoreCase) ||
            bodyLower.Contains("validation-summary-errors", StringComparison.OrdinalIgnoreCase) ||
            bodyLower.Contains("field-validation-error", StringComparison.OrdinalIgnoreCase);

        if (looksLikeCredentialError)
            throw new InvalidOperationException(_localization.GetString("LoginFailedCheckUsernameOrPassword"));

        // به این نقطه که رسیدیم یعنی نه توکن گرفتیم و نه خطای مشخص اعتبارسنجی دیدیم - یعنی
        // سایت صفحه‌ی غیرمنتظره‌ای برگردانده (مثلاً کد دو مرحله‌ای، کپچا، یا تغییر ساختار سایت)
        // که این فرم مستقیم نمی‌تواند آن را تشخیص/تکمیل کند. در این حالت باید از مرورگر داخلی
        // استفاده شود.
        throw new InvalidOperationException(_localization.GetString("LoginTokenWasNotReceivedTheSiteMayRequireAnExtraStepThisFormCanTHandlePleaseUseInternalBrowserLogin"));
    }

    // همه‌ی فیلدهای <input type="hidden" name="..." value="..."> صفحه را (با هر ترتیب
    // اتریبیوت) استخراج می‌کند تا هنگام ارسال فرم چیزی از قلم نیفتد.
    private static Dictionary<string, string> ExtractAllHiddenInputs(string html)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var matches = Regex.Matches(html, "<input\\b[^>]*type=[\\\"']hidden[\\\"'][^>]*>", RegexOptions.IgnoreCase);
        foreach (Match m in matches)
        {
            string tag = m.Value;
            var nameMatch = Regex.Match(tag, "name=[\\\"']([^\\\"']*)[\\\"']", RegexOptions.IgnoreCase);
            if (!nameMatch.Success)
                continue;
            string name = System.Net.WebUtility.HtmlDecode(nameMatch.Groups[1].Value);
            var valueMatch = Regex.Match(tag, "value=[\\\"']([^\\\"']*)[\\\"']", RegexOptions.IgnoreCase);
            string value = valueMatch.Success ? System.Net.WebUtility.HtmlDecode(valueMatch.Groups[1].Value) : string.Empty;
            result[name] = value;
        }

        return result;
    }

    // اگر بعد از ورود، به‌جای توکن یک صفحه‌ی «اجازه دسترسی» (consent) دیدیم، فیلدهای پنهانش را
    // به همراه یک فیلد تاییدِ رایج (allow/AllowSelected و مشابه) برمی‌گرداند تا خودکار تایید شود.
    private static bool TryBuildConsentForm(string html, out Dictionary<string, string> form)
    {
        form = new Dictionary<string, string>();
        if (string.IsNullOrWhiteSpace(html) || !html.Contains("<form", StringComparison.OrdinalIgnoreCase))
            return false;

        bool looksLikeConsent =
            html.Contains("consent", StringComparison.OrdinalIgnoreCase) ||
            html.Contains("AllowSelected", StringComparison.OrdinalIgnoreCase) ||
            html.Contains("scopesconsented", StringComparison.OrdinalIgnoreCase) ||
            html.Contains("اجازه دسترسی", StringComparison.OrdinalIgnoreCase);
        if (!looksLikeConsent)
            return false;

        form = ExtractAllHiddenInputs(html);
        if (form.Count == 0)
            return false;

        if (!form.ContainsKey("AllowSelected"))
            form["AllowSelected"] = "true";
        if (!form.ContainsKey("button"))
            form["button"] = "yes";

        return true;
    }

    private static bool IsRedirect(System.Net.HttpStatusCode statusCode)
    {
        int code = (int)statusCode;
        return code is >= 300 and < 400;
    }

    private static string ExtractInputValue(string html, string name)
    {
        string pattern = $"<input[^>]*name=[\\\"']{Regex.Escape(name)}[\\\"'][^>]*value=[\\\"']([^\\\"']*)[\\\"'][^>]*>";
        var match = Regex.Match(html, pattern, RegexOptions.IgnoreCase);
        if (match.Success)
            return System.Net.WebUtility.HtmlDecode(match.Groups[1].Value);

        pattern = $"<input[^>]*value=[\\\"']([^\\\"']*)[\\\"'][^>]*name=[\\\"']{Regex.Escape(name)}[\\\"'][^>]*>";
        match = Regex.Match(html, pattern, RegexOptions.IgnoreCase);
        return match.Success ? System.Net.WebUtility.HtmlDecode(match.Groups[1].Value) : string.Empty;
    }

    private static bool TryExtractTokenFromUri(Uri uri, out TtacTokenResult result)
    {
        result = new TtacTokenResult();
        string fragment = uri.Fragment;
        if (string.IsNullOrWhiteSpace(fragment))
            return false;

        if (fragment.StartsWith("#"))
            fragment = fragment[1..];

        var values = fragment.Split('&', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Split('=', 2))
            .Where(parts => parts.Length == 2)
            .ToDictionary(parts => Uri.UnescapeDataString(parts[0]), parts => Uri.UnescapeDataString(parts[1]), StringComparer.OrdinalIgnoreCase);

        if (!values.TryGetValue("access_token", out var accessToken) || string.IsNullOrWhiteSpace(accessToken))
            return false;

        result.AccessToken = accessToken;
        if (values.TryGetValue("expires_in", out var expiresText) && int.TryParse(expiresText, out int expires))
            result.ExpiresInSeconds = expires;

        return true;
    }

    private string GetHistoryCsvPathForKey(string key)
    {
        key = string.IsNullOrWhiteSpace(key) ? "default" : key;
        if (string.Equals(key, "default", StringComparison.OrdinalIgnoreCase))
            return Path.Combine(AppContext.BaseDirectory, "scans.csv");
        return Path.Combine(AppContext.BaseDirectory, $"scans-{key}.csv");
    }

    private string GetActiveHistoryCsvPath()
    {
        return GetHistoryCsvPathForKey(_historyLoadedPharmacyKey);
    }

    private void SaveHistoryItemsForKey(string key)
    {
        try
        {
            string csvPath = GetHistoryCsvPathForKey(key);

            // این متد الان از جاهایی هم صدا زده می‌شود که ممکن است روی ترد پس‌زمینه (بعد از
            // اتمام استعلام تی‌تک) اجرا شوند، نه فقط ترد UI. چون HistoryItems یک
            // ObservableCollection متصل به لیست‌ویو است، پیمایش مستقیم آن از یک ترد پس‌زمینه
            // هم‌زمان با تغییرش روی ترد UI می‌تواند استثنا بدهد؛ برای همین ابتدا یک کپی امن
            // روی ترد UI گرفته می‌شود.
            List<ScanRecord> snapshot = Dispatcher.CheckAccess()
                ? HistoryItems.OrderBy(r => r.TimestampLocal).ToList()
                : Dispatcher.Invoke(() => HistoryItems.OrderBy(r => r.TimestampLocal).ToList());

            // ستون چهارم (drugName) اضافه شد تا نام دارو بعد از بستن و باز کردن مجدد برنامه از بین نرود.
            var lines = new List<string> { "timestamp_iso,deviceName,barcode,drugName" };
            foreach (var item in snapshot)
                lines.Add($"{item.TimestampLocal.ToUniversalTime():O},{EscapeCsvLocal(item.DeviceName)},{EscapeCsvLocal(item.Barcode)},{EscapeCsvLocal(item.DrugName)}");
            File.WriteAllLines(csvPath, lines, Encoding.UTF8);
        }
        catch { }
    }

    private void SaveCurrentHistoryForCurrentPharmacy()
    {
        SaveHistoryItemsForKey(string.IsNullOrWhiteSpace(_historyLoadedPharmacyKey) ? "default" : _historyLoadedPharmacyKey);
    }

    private void LoadHistoryItemsForPharmacy(string pharmacyName)
    {
        string key = GetReceiveStatusStorageKey(pharmacyName);
        if (string.Equals(_historyLoadedPharmacyKey, key, StringComparison.OrdinalIgnoreCase))
            return;

        SaveCurrentHistoryForCurrentPharmacy();
        _historyLoadedPharmacyKey = key;
        LoadHistoryFromCsv();
    }

    private string GetReceiveStatusHistoryPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "receive-status-history.json");
    }

    private string GetReceiveStatusStorageKey(string? pharmacyName = null)
    {
        string value = string.IsNullOrWhiteSpace(pharmacyName) ? _ttacPharmacyDisplayName : pharmacyName;
        if (string.IsNullOrWhiteSpace(value))
            return "default";

        foreach (char invalid in Path.GetInvalidFileNameChars())
            value = value.Replace(invalid, '_');

        return value.Trim();
    }

    private Dictionary<string, List<ReceiveStatusStorageRow>> LoadReceiveStatusStore()
    {
        try
        {
            string path = GetReceiveStatusHistoryPath();
            if (!File.Exists(path))
                return new Dictionary<string, List<ReceiveStatusStorageRow>>(StringComparer.OrdinalIgnoreCase);

            return JsonSerializer.Deserialize<Dictionary<string, List<ReceiveStatusStorageRow>>>(File.ReadAllText(path))
                   ?? new Dictionary<string, List<ReceiveStatusStorageRow>>(StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            return new Dictionary<string, List<ReceiveStatusStorageRow>>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private void SaveReceiveStatusStore(Dictionary<string, List<ReceiveStatusStorageRow>> store)
    {
        try
        {
            File.WriteAllText(GetReceiveStatusHistoryPath(), JsonSerializer.Serialize(store, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private void SaveCurrentReceiveStatusItemsForCurrentPharmacy()
    {
        string key = string.IsNullOrWhiteSpace(_receiveStatusLoadedPharmacyKey)
            ? GetReceiveStatusStorageKey()
            : _receiveStatusLoadedPharmacyKey;
        SaveReceiveStatusItemsForKey(key);
    }

    private void SaveReceiveStatusItemsForKey(string key)
    {
        if (string.IsNullOrWhiteSpace(key))
            return;

        var store = LoadReceiveStatusStore();
        store[key] = ReceiveStatusItems.Select(ReceiveStatusStorageRow.FromRow).ToList();
        SaveReceiveStatusStore(store);
    }

    private void LoadReceiveStatusItemsForPharmacy(string pharmacyName)
    {
        string key = GetReceiveStatusStorageKey(pharmacyName);
        if (string.Equals(_receiveStatusLoadedPharmacyKey, key, StringComparison.OrdinalIgnoreCase))
            return;

        var store = LoadReceiveStatusStore();
        ReceiveStatusItems.Clear();
        _receiveStatusKnownBarcodes.Clear();

        if (store.TryGetValue(key, out var rows))
        {
            foreach (var stored in rows)
            {
                var row = stored.ToRow();
                ReceiveStatusItems.Add(row);
                if (!string.IsNullOrWhiteSpace(row.Barcode))
                    _receiveStatusKnownBarcodes.Add(row.Barcode);
            }
        }

        _receiveStatusLoadedPharmacyKey = key;
        RefreshReceiveStatusRowNumbers();
    }

    private string GetCargoDeliveryHistoryPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "cargo-delivery-history.json");
    }

    private Dictionary<string, List<CargoDeliveryStorageRow>> LoadCargoDeliveryStore()
    {
        try
        {
            string path = GetCargoDeliveryHistoryPath();
            if (!File.Exists(path))
                return new Dictionary<string, List<CargoDeliveryStorageRow>>(StringComparer.OrdinalIgnoreCase);

            return JsonSerializer.Deserialize<Dictionary<string, List<CargoDeliveryStorageRow>>>(File.ReadAllText(path))
                   ?? new Dictionary<string, List<CargoDeliveryStorageRow>>(StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            return new Dictionary<string, List<CargoDeliveryStorageRow>>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private void SaveCargoDeliveryStore(Dictionary<string, List<CargoDeliveryStorageRow>> store)
    {
        try
        {
            File.WriteAllText(GetCargoDeliveryHistoryPath(), JsonSerializer.Serialize(store, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private void SaveCurrentCargoDeliveryItemsForCurrentPharmacy()
    {
        string key = string.IsNullOrWhiteSpace(_cargoDeliveryLoadedPharmacyKey)
            ? GetReceiveStatusStorageKey()
            : _cargoDeliveryLoadedPharmacyKey;
        SaveCargoDeliveryItemsForKey(key);
    }

    private void SaveCargoDeliveryItemsForKey(string key)
    {
        if (string.IsNullOrWhiteSpace(key))
            return;

        var store = LoadCargoDeliveryStore();
        store[key] = CargoDeliveryItems.Select(CargoDeliveryStorageRow.FromRow).ToList();
        SaveCargoDeliveryStore(store);
    }

    private void LoadCargoDeliveryItemsForPharmacy(string pharmacyName)
    {
        string key = GetReceiveStatusStorageKey(pharmacyName);
        if (string.Equals(_cargoDeliveryLoadedPharmacyKey, key, StringComparison.OrdinalIgnoreCase))
            return;

        var store = LoadCargoDeliveryStore();
        CargoDeliveryItems.Clear();
        _cargoDeliveryKnownBarcodes.Clear();

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        if (store.TryGetValue(key, out var rows))
        {
            foreach (var stored in rows)
            {
                var row = stored.ToRow(english);
                CargoDeliveryItems.Add(row);
                if (!string.IsNullOrWhiteSpace(row.Barcode))
                    _cargoDeliveryKnownBarcodes.Add(row.Barcode);
            }
        }

        _cargoDeliveryLoadedPharmacyKey = key;
        RefreshCargoDeliveryRowNumbers();
    }

    // =========================================================================================
    // هشدار تاریخ انقضای نزدیک (Expiry Alert) - ویژگی جدید
    // وقتی کالایی در «تحویل بار» ثبت می‌شود و تاریخ انقضای قابل‌فهمی دارد، برای پایش جداگانه
    // ثبت می‌شود (مستقل از لیست تحویل بار که ممکن است هر ماه آرشیو/پاک شود). هر بار که برنامه
    // بالا می‌آید یا هر ۶ ساعت، اگر تا ۶ ماه (قابل‌تنظیم) به تاریخ انقضا مانده باشد و قبلاً
    // «فروخته شده» علامت نخورده باشد، یک هشدار داخل برنامه نشان داده می‌شود و - در صورت تنظیم
    // بودن - پیام بله هم ارسال می‌شود. با «حواسم هست» هشدار بعدی ۱ ماه دیگر می‌آید؛ با
    // «فروخته شد» دیگر هیچ‌وقت هشدار نمی‌آید.
    // =========================================================================================

    private static bool TryParseTtacExpirationDate(string? raw, out DateTime expirationDate)
    {
        expirationDate = default;
        if (string.IsNullOrWhiteSpace(raw))
            return false;

        string text = ToEnglishDigits(raw.Trim());

        // تلاش اول: فرمت میلادی/ISO استاندارد (چیزی که تی‌تک معمولاً برای Expiration برمی‌گرداند)
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var gregorianDate)
            && gregorianDate.Year >= 1900 && gregorianDate.Year <= 2200)
        {
            expirationDate = gregorianDate.Date;
            return true;
        }

        // تلاش دوم: تاریخ شمسی به‌صورت yyyy/MM/dd یا yyyy-MM-dd (سال بین ۱۳۰۰ تا ۱۴۹۹)
        var match = Regex.Match(text, @"^(?<y>1[34]\d{2})[/\-.](?<m>\d{1,2})[/\-.](?<d>\d{1,2})");
        if (match.Success)
        {
            try
            {
                int y = int.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture);
                int m = int.Parse(match.Groups["m"].Value, CultureInfo.InvariantCulture);
                int d = int.Parse(match.Groups["d"].Value, CultureInfo.InvariantCulture);
                var persianCalendar = new PersianCalendar();
                expirationDate = persianCalendar.ToDateTime(y, m, d, 0, 0, 0, 0).Date;
                return true;
            }
            catch
            {
                return false;
            }
        }

        return false;
    }

    private string GetExpiryAlertSettingsPath() => Path.Combine(AppContext.BaseDirectory, "expiry-alert-settings.json");

    private void LoadExpiryAlertSettings()
    {
        try
        {
            string path = GetExpiryAlertSettingsPath();
            _expiryAlertSettings = File.Exists(path)
                ? JsonSerializer.Deserialize<ExpiryAlertSettings>(File.ReadAllText(path)) ?? new ExpiryAlertSettings()
                : new ExpiryAlertSettings();
        }
        catch
        {
            _expiryAlertSettings = new ExpiryAlertSettings();
        }
    }

    private void SaveExpiryAlertSettings()
    {
        try
        {
            File.WriteAllText(GetExpiryAlertSettingsPath(), JsonSerializer.Serialize(_expiryAlertSettings, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    // =========================================================================================
    // یادآور بله - فعال‌سازی با یک دکمه: کاربر روی «فعال‌سازی» می‌زند، بله باز می‌شود و کاربر آنجا
    // «شروع» را می‌زند؛ برنامه با یک کد یکتا (که به لینک شروع چسبانده شده) هر چند ثانیه یک‌بار از
    // ربات می‌پرسد که آیا کسی با این کد پیام داده یا نه (getUpdates)؛ وقتی پیدا شد، chat_id همان
    // پیام به‌عنوان چت‌آیدی این داروخانه ذخیره می‌شود.
    // =========================================================================================

    private string? _balePendingStartCode;
    private System.Windows.Threading.DispatcherTimer? _baleActivationPollTimer;
    private int _baleActivationPollAttempts;
    // وقتی به‌صورت برنامه‌ای IsChecked چک‌باکس را ست می‌کنیم (نه با کلیک کاربر)، این پرچم جلوی
    // اجرای دوباره‌ی BaleNotificationsEnabledCheckBox_Click را می‌گیرد.
    private bool _suppressBaleCheckboxEvent;

    private void UpdateBaleConnectionStatusText()
    {
        if (BaleConnectionStatusText == null)
            return;

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        bool configured = _expiryAlertSettings.IsBaleConfigured;
        bool enabled = _expiryAlertSettings.IsBaleNotificationsEnabled;

        BaleConnectionStatusText.Text = !configured
            ? (_localization.GetString("Inactive"))
            : enabled
                ? (_localization.GetString("ActiveAlertsAreAlsoSentOnBale"))
                : (_localization.GetString("PausedBaleAlertsAreTurnedOff"));

        if (BaleTestMessageButton != null)
            BaleTestMessageButton.Visibility = configured ? Visibility.Visible : Visibility.Collapsed;

        if (BaleNotificationsEnabledCheckBox != null)
        {
            BaleNotificationsEnabledCheckBox.Visibility = configured ? Visibility.Visible : Visibility.Collapsed;
            BaleNotificationsEnabledCheckBox.Content = _localization.GetString("SendBaleAlerts");
            _suppressBaleCheckboxEvent = true;
            BaleNotificationsEnabledCheckBox.IsChecked = enabled;
            _suppressBaleCheckboxEvent = false;
        }
    }

    private void BaleNotificationsEnabledCheckBox_Click(object sender, RoutedEventArgs e)
    {
        if (_suppressBaleCheckboxEvent)
            return;

        bool wasEnabled = _expiryAlertSettings.IsBaleNotificationsEnabled;
        bool nowEnabled = BaleNotificationsEnabledCheckBox.IsChecked ?? true;
        _expiryAlertSettings.IsBaleNotificationsEnabled = nowEnabled;
        SaveExpiryAlertSettings();
        UpdateBaleConnectionStatusText();

        // برای سیستم‌های دیگرِ هم‌شبکه که همان لایسنس را دارند هم پخش کن (نسخه‌ی جدید را خودش
        // دوباره ست/ذخیره می‌کند - بی‌ضرر است، فقط برای هماهنگی نسخه‌ها لازم است).
        PublishDesktopSettingsForSync();

        // فقط وقتی از غیرفعال به فعال برمی‌گردد (و نه بار اول فعال‌سازی) یک پیام تاییدیه در بله
        // فرستاده می‌شود تا کاربر مطمئن شود ربات دوباره فعال شده است.
        if (nowEnabled && !wasEnabled && _expiryAlertSettings.IsBaleConfigured && !string.IsNullOrWhiteSpace(SharedBaleBotToken))
        {
            _ = SendBaleMessageAsync(
                SharedBaleBotToken,
                _expiryAlertSettings.BaleChatId,
                _localization.GetString("BaleRemindersForScanbridgeWereTurnedBackOn"));
        }
    }

    private async void BaleTestMessageButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(SharedBaleBotToken) || string.IsNullOrWhiteSpace(_expiryAlertSettings.BaleChatId))
            return;

        BaleTestMessageButton.IsEnabled = false;
        bool ok = await SendBaleMessageAsync(
            SharedBaleBotToken,
            _expiryAlertSettings.BaleChatId,
            "این یک پیام آزمایشی از اسکن‌بریج است. اگر این پیام را در بله می‌بینید، یادآور تاریخ انقضا برای شما درست فعال شده است.");
        BaleTestMessageButton.IsEnabled = true;

        ShowStyledMessage(
            _localization.GetString("BaleTest"),
            ok
                ? (_localization.GetString("TestMessageSentCheckBale"))
                : (_localization.GetString("SendingFailedCheckYourInternetConnection")),
            !ok);
    }

    private void BaleActivateButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(SharedBaleBotToken) || string.IsNullOrWhiteSpace(SharedBaleBotUsername))
        {
            ShowStyledMessage(
                _localization.GetString("BaleReminder2"),
                _localization.GetString("BaleBotIsNotConfiguredInThisAppBuildYet"),
                true);
            return;
        }

        _balePendingStartCode = "sb" + Guid.NewGuid().ToString("N").Substring(0, 10);

        try
        {
            string url = $"https://ble.ir/{SharedBaleBotUsername}?start={_balePendingStartCode}";
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch { }

        if (BaleConnectionStatusText != null)
            BaleConnectionStatusText.Text = _localization.GetString("WaitingForConfirmationInBalePleaseWaitAFewSecondsAfterTappingStart");

        StartBaleActivationPolling();
    }

    private void StartBaleActivationPolling()
    {
        StopBaleActivationPolling();
        _baleActivationPollAttempts = 0;

        _baleActivationPollTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
        _baleActivationPollTimer.Tick += async (_, _) =>
        {
            _baleActivationPollAttempts++;
            if (_baleActivationPollAttempts > 30)
            {
                StopBaleActivationPolling();
                if (BaleConnectionStatusText != null)
                    BaleConnectionStatusText.Text = _localization.GetString("TimedOutTapActivateBaleReminderAgain");
                return;
            }

            string? chatId = await TryFindBaleChatIdAsync(_balePendingStartCode ?? "");
            if (!string.IsNullOrWhiteSpace(chatId))
            {
                StopBaleActivationPolling();
                _expiryAlertSettings.BaleChatId = chatId;
                SaveExpiryAlertSettings();
                UpdateBaleConnectionStatusText();
            }
        };
        _baleActivationPollTimer.Start();
    }

    private void StopBaleActivationPolling()
    {
        _baleActivationPollTimer?.Stop();
        _baleActivationPollTimer = null;
    }

    private static async Task<string?> TryFindBaleChatIdAsync(string startCode)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(startCode) || string.IsNullOrWhiteSpace(SharedBaleBotToken))
                return null;

            string url = $"https://tapi.bale.ai/bot{SharedBaleBotToken}/getUpdates?limit=50";
            using var response = await _baleHttpClient.GetAsync(url);
            if (!response.IsSuccessStatusCode)
                return null;

            string json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            if (!doc.RootElement.TryGetProperty("result", out var result) || result.ValueKind != JsonValueKind.Array)
                return null;

            foreach (var update in result.EnumerateArray())
            {
                if (!update.TryGetProperty("message", out var message))
                    continue;
                if (!message.TryGetProperty("text", out var textProp))
                    continue;

                string text = textProp.GetString() ?? "";
                if (!text.Contains(startCode, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!message.TryGetProperty("chat", out var chat) || !chat.TryGetProperty("id", out var idProp))
                    continue;

                return idProp.ValueKind == JsonValueKind.Number
                    ? idProp.GetInt64().ToString(CultureInfo.InvariantCulture)
                    : idProp.GetString();
            }
        }
        catch { }

        return null;
    }

    private string GetExpiryWatchStorePath() => Path.Combine(AppContext.BaseDirectory, "expiry-watch-items.json");

    private Dictionary<string, List<ExpiryWatchItem>> LoadExpiryWatchStore()
    {
        try
        {
            string path = GetExpiryWatchStorePath();
            if (!File.Exists(path))
                return new Dictionary<string, List<ExpiryWatchItem>>(StringComparer.OrdinalIgnoreCase);

            return JsonSerializer.Deserialize<Dictionary<string, List<ExpiryWatchItem>>>(File.ReadAllText(path))
                   ?? new Dictionary<string, List<ExpiryWatchItem>>(StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            return new Dictionary<string, List<ExpiryWatchItem>>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private void SaveExpiryWatchStore(Dictionary<string, List<ExpiryWatchItem>> store)
    {
        try
        {
            File.WriteAllText(GetExpiryWatchStorePath(), JsonSerializer.Serialize(store, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private void TryRegisterExpiryWatchItem(CargoDeliveryRow row)
    {
        try
        {
            if (row == null || string.IsNullOrWhiteSpace(row.Barcode))
                return;
            if (!TryParseTtacExpirationDate(row.Expiration, out var expirationDate))
                return; // این کالا تاریخ انقضای قابل‌فهمی نداشت؛ نادیده گرفته می‌شود

            string pharmacyKey = GetReceiveStatusStorageKey();
            var store = LoadExpiryWatchStore();
            if (!store.TryGetValue(pharmacyKey, out var items))
            {
                items = new List<ExpiryWatchItem>();
                store[pharmacyKey] = items;
            }

            var existing = items.FirstOrDefault(x => string.Equals(x.Barcode, row.Barcode, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.ProductName = row.ProductName;
                existing.ProductEnName = row.ProductEnName;
                existing.BatchCode = row.LotNumber;
                existing.Quantity = row.Quantity;
                existing.ExpirationRaw = row.Expiration;
                existing.ExpirationDate = expirationDate;
                // اگر قبلاً «فروخته شده» علامت خورده، دست‌نخورده می‌ماند - فرض بر این است که
                // همان بسته دوباره اسکن شده، نه یک محموله‌ی واقعاً جدید.
            }
            else
            {
                items.Add(new ExpiryWatchItem
                {
                    Barcode = row.Barcode,
                    ProductName = row.ProductName,
                    ProductEnName = row.ProductEnName,
                    BatchCode = row.LotNumber,
                    Quantity = row.Quantity,
                    ExpirationRaw = row.Expiration,
                    ExpirationDate = expirationDate,
                    Status = ExpiryWatchStatus.Watching,
                    NextAlertDueUtc = DateTime.UtcNow,
                    CreatedAtUtc = DateTime.UtcNow
                });
            }

            SaveExpiryWatchStore(store);
        }
        catch { }
    }

    private async Task CheckExpiryAlertsAsync()
    {
        try
        {
            string pharmacyKey = GetReceiveStatusStorageKey();
            var store = LoadExpiryWatchStore();
            if (!store.TryGetValue(pharmacyKey, out var items) || items.Count == 0)
                return;

            DateTime nowUtc = DateTime.UtcNow;
            DateTime thresholdDate = DateTime.Now.AddMonths(Math.Max(1, _expiryAlertSettings.ThresholdMonths));

            var due = items
                .Where(x => x.Status == ExpiryWatchStatus.Watching
                            && x.ExpirationDate <= thresholdDate
                            && x.NextAlertDueUtc <= nowUtc)
                .OrderBy(x => x.ExpirationDate)
                .ToList();

            if (due.Count == 0)
                return;

            // زمان هشدار بعدی همین الان جلو کشیده می‌شود (طبق تنظیم «دوباره هر چند روز یادآوری کند») -
            // چه کاربر روی هشدار کلیک کند چه نکند - تا اسپم نشود ولی دوباره یادآوری شود.
            int repeatDays = Math.Max(1, _expiryAlertSettings.RepeatReminderDays);
            foreach (var item in due)
            {
                item.LastAlertedUtc = nowUtc;
                item.NextAlertDueUtc = nowUtc.AddDays(repeatDays);
                // با فایر شدن هشدار، این قلم دوباره «منتظر پاسخ» می‌شود - این جدا از NextAlertDueUtc
                // است، پس حتی اگر کاربر پاپ‌آپ را ببندد و هیچ دکمه‌ای نزند، نشان قرمز روی دکمه‌ی
                // «تاریخ نزدیک» می‌ماند تا خودش با «فروخته شد»/«حواسم هست» پاکش کند.
                item.NeedsResponse = true;
            }
            SaveExpiryWatchStore(store);

            if (_expiryAlertSettings.IsBaleConfigured && _expiryAlertSettings.IsBaleNotificationsEnabled && !string.IsNullOrWhiteSpace(SharedBaleBotToken))
            {
                string baleText = BuildBaleExpiryAlertText(due);
                _ = SendBaleMessageAsync(SharedBaleBotToken, _expiryAlertSettings.BaleChatId, baleText);
            }

            await Dispatcher.InvokeAsync(() =>
            {
                ShowExpiryAlertPopup(due);
                RefreshExpiryWatchDisplayList();
            });
        }
        catch { }
    }

    private static string BuildBaleExpiryAlertText(List<ExpiryWatchItem> items)
    {
        var sb = new StringBuilder();
        sb.AppendLine("⏰ هشدار تاریخ انقضای نزدیک - Scanbridge");
        sb.AppendLine();
        DateTime today = DateTime.Now.Date;
        foreach (var item in items)
        {
            string name = string.IsNullOrWhiteSpace(item.ProductName) ? item.ProductEnName : item.ProductName;
            if (string.IsNullOrWhiteSpace(name)) name = item.Barcode;

            int daysLeft = (int)Math.Ceiling((item.ExpirationDate - today).TotalDays);
            string daysText = daysLeft < 0
                ? "تاریخ انقضا گذشته است"
                : daysLeft == 0
                    ? "امروز تاریخش تمام می‌شود"
                    : $"{daysLeft} روز دیگر تاریخش تمام می‌شود";

            sb.AppendLine($"⚠️ {name} ⚠️");
            sb.AppendLine($"سری ساخت: {(string.IsNullOrWhiteSpace(item.BatchCode) ? "-" : item.BatchCode)} | انقضا: {item.PersianExpirationText} | {daysText}");
            sb.AppendLine();
        }
        return sb.ToString().TrimEnd();
    }

    private static async Task<bool> SendBaleMessageAsync(string botToken, string chatId, string text)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(botToken) || string.IsNullOrWhiteSpace(chatId))
                return false;

            string url = $"https://tapi.bale.ai/bot{botToken}/sendMessage";
            var body = new { chat_id = chatId, text };
            using var content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
            using var response = await _baleHttpClient.PostAsync(url, content);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private void ShowExpiryAlertPopup(List<ExpiryWatchItem> due)
    {
        if (due == null || due.Count == 0)
            return;

        // این پاپ‌آپ ممکن است روی داشبورد اصلی باز شود یا روی هر پنل دیگری (تحویل بار، تنظیمات،
        // تاریخ نزدیک و ...) که کاربر همان لحظه باز داشته. برای اینکه همیشه پشتش بلور شود - نه
        // فقط داشبورد - کل ریشه‌ی محتوای پنجره (RootContentGrid، که هم داشبورد و هم همه‌ی
        // اورلی‌ها زیرش هستند) بلور می‌شود، نه فقط MainContent. فقط اگر از قبل بلور نبود این کار
        // را می‌کنیم و با بسته‌شدن پاپ‌آپ هم فقط در همان صورت برش می‌داریم.
        bool blurWasAlreadyActive = RootContentGrid.Effect != null;
        if (!blurWasAlreadyActive)
            RootContentGrid.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };

        // این پاپ‌آپ غیرمودال است (win.Show نه ShowDialog) و ممکن است کاربر مدتی آن را باز نگه
        // دارد؛ اگر قبل از پاسخ به آن (زدن «فروخته شد»/«باشه») داروخانه را عوض کند،
        // GetReceiveStatusStorageKey() در لحظه‌ی کلیک دیگر همان داروخانه‌ای را برنمی‌گرداند که
        // این هشدارها برایش ساخته شده - باید همین‌جا (قبل از باز شدن پنجره) ثبت شود (باگ ۱۵
        // گزارش ممیزی).
        string capturedPharmacyKeyForAlert = GetReceiveStatusStorageKey();

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        var win = new ExpiryAlertWindow(due, english) { Owner = this };
        win.ItemMarkedSold += barcode => MarkExpiryItemSold(barcode, capturedPharmacyKeyForAlert);
        win.ItemAcknowledged += barcode => AcknowledgeExpiryItem(barcode, capturedPharmacyKeyForAlert);
        win.Closed += (_, _) =>
        {
            if (!blurWasAlreadyActive)
                RootContentGrid.Effect = null;
        };
        win.Show();
    }

    // pharmacyKeyOverride: وقتی از دکمه‌ی داخل پنل «تاریخ نزدیک» زده می‌شود، null است و کلید
    // داروخانه‌ی جاری استفاده می‌شود (رفتار قبلی، درست است چون آن پنل همیشه داروخانه‌ی جاری را
    // نشان می‌دهد). وقتی از پاپ‌آپ غیرمودال زده می‌شود، ShowExpiryAlertPopup کلید داروخانه‌ای که
    // پاپ‌آپ برایش باز شده را صراحتاً می‌فرستد، تا اگر کاربر بین باز شدن پاپ‌آپ و کلیک روی دکمه
    // داروخانه را عوض کرده باشد، هنوز روی داروخانه‌ی درست اعمال شود.
    private void MarkExpiryItemSold(string barcode, string? pharmacyKeyOverride = null)
    {
        try
        {
            string pharmacyKey = pharmacyKeyOverride ?? GetReceiveStatusStorageKey();
            var store = LoadExpiryWatchStore();
            if (store.TryGetValue(pharmacyKey, out var items))
            {
                var item = items.FirstOrDefault(x => string.Equals(x.Barcode, barcode, StringComparison.OrdinalIgnoreCase));
                if (item != null)
                {
                    item.Status = ExpiryWatchStatus.Sold;
                    item.SoldAtUtc = DateTime.UtcNow;
                    item.NeedsResponse = false;
                    SaveExpiryWatchStore(store);
                }
            }
        }
        catch { }
        // چه از دکمه‌ی داخل پنل زده شده باشد چه از پاپ‌آپ هشدار، لیست و نوتیف روی «تاریخ نزدیک» باید
        // بلافاصله به‌روز شوند.
        RefreshExpiryWatchDisplayList();
    }

    private void AcknowledgeExpiryItem(string barcode, string? pharmacyKeyOverride = null)
    {
        try
        {
            string pharmacyKey = pharmacyKeyOverride ?? GetReceiveStatusStorageKey();
            var store = LoadExpiryWatchStore();
            if (store.TryGetValue(pharmacyKey, out var items))
            {
                var item = items.FirstOrDefault(x => string.Equals(x.Barcode, barcode, StringComparison.OrdinalIgnoreCase));
                if (item != null)
                {
                    item.NextAlertDueUtc = DateTime.UtcNow.AddDays(Math.Max(1, _expiryAlertSettings.RepeatReminderDays));
                    item.NeedsResponse = false;
                    SaveExpiryWatchStore(store);
                }
            }
        }
        catch { }
        RefreshExpiryWatchDisplayList();
    }


    private void ApplyTtacAccessToken(string token, DateTime expiresAtUtc)
    {
        string oldPharmacy = _ttacPharmacyDisplayName;
        string? oldToken = _ttacAccessTokenOverride;
        string oldStorageKey = _receiveStatusLoadedPharmacyKey;
        string newPharmacy = TryExtractTtacDisplayNameFromToken(token) ?? string.Empty;
        string newStorageKey = GetReceiveStatusStorageKey(newPharmacy);
        bool tokenChanged = !string.Equals(oldToken, token, StringComparison.Ordinal);
        bool pharmacyChanged = !string.IsNullOrWhiteSpace(oldStorageKey)
                              && !string.Equals(oldStorageKey, newStorageKey, StringComparison.OrdinalIgnoreCase);

        if (pharmacyChanged)
        {
            SaveReceiveStatusItemsForKey(oldStorageKey);
            SaveCargoDeliveryItemsForKey(string.IsNullOrWhiteSpace(_cargoDeliveryLoadedPharmacyKey) ? oldStorageKey : _cargoDeliveryLoadedPharmacyKey);
            SaveHistoryItemsForKey(string.IsNullOrWhiteSpace(_historyLoadedPharmacyKey) ? "default" : _historyLoadedPharmacyKey);
            SaveTtacRegistrationHistory();
        }

        _ttacAccessTokenOverride = token;
        _ttacPharmacyDisplayName = newPharmacy;
        _ttacAccessTokenExpiresAtUtc = expiresAtUtc;
        // ورود موفق انجام شد؛ دیگر نیازی به پر کردن اجباری یک حساب خاص نیست.
        _pendingTtacAutofillUsername = null;

        // وقتی کاربر وارد تی‌تک شده یعنی حالت داروخانه‌ای فعال است؛
        // اگر از نصب قبلی جستجوی تی‌تک خاموش مانده باشد، شیر خشک/ثبت خودکار کار نمی‌کند.
        if (!_ttTeckSettings.IsEnabled)
        {
            _ttTeckSettings.IsEnabled = true;
            if (TtTeckEnabledCheckBox != null)
                TtTeckEnabledCheckBox.IsChecked = true;
            SaveTtTeckSettings();
        }

        if (pharmacyChanged)
        {
            _queuedReceiveStatusBarcodes.Clear();
            // این کش فقط جلوی باز شدن مکرر فرم ثبت برای یک بارکد در همین داروخانه را می‌گیرد
            // (مثلاً اسکن دوباره‌ی همان جعبه)؛ اگر با تغییر داروخانه پاک نشود، بارکدی که قبلاً در
            // داروخانه‌ی قبلی باعث باز شدن فرم شده، در داروخانه‌ی جدید (که وضعیت ثبتش کاملاً جداست)
            // دیگر هرگز خودکار باز نمی‌شود - حتی اگر از تاریخچه‌ی گوشی دوباره ارسال شود.
            _autoOpenedFormulaRegistrationKeys.Clear();
        }

        if (pharmacyChanged || string.IsNullOrWhiteSpace(_historyLoadedPharmacyKey) || string.Equals(_historyLoadedPharmacyKey, "default", StringComparison.OrdinalIgnoreCase))
            LoadHistoryItemsForPharmacy(newPharmacy);

        if (pharmacyChanged || string.IsNullOrWhiteSpace(_receiveStatusLoadedPharmacyKey))
            LoadReceiveStatusItemsForPharmacy(newPharmacy);

        if (pharmacyChanged || string.IsNullOrWhiteSpace(_cargoDeliveryLoadedPharmacyKey))
            LoadCargoDeliveryItemsForPharmacy(newPharmacy);

        if (pharmacyChanged || string.IsNullOrWhiteSpace(_ttacRegistrationHistoryLoadedPharmacyKey) || string.Equals(_ttacRegistrationHistoryLoadedPharmacyKey, "default", StringComparison.OrdinalIgnoreCase))
            LoadTtacRegistrationHistoryForPharmacy(newPharmacy);

        // هشدار «تاریخ نزدیک» به داروخانه وابسته است - هر داروخانه فهرست پایش خودش را دارد
        // (GetReceiveStatusStorageKey از همین _ttacPharmacyDisplayName که چند خط بالاتر ست شد
        // کلید می‌سازد). قبلاً بررسی فقط با تایمر انجام می‌شد (۸ ثانیه بعد از باز شدن برنامه و
        // هر ۶ ساعت)؛ اگر آن لحظه هنوز وارد تی‌تک نشده بودیم (حالت معمول، چون ورود خودکار نیست)،
        // بررسی زیر کلید «default» انجام می‌شد که هیچ‌وقت اقلام واقعی این داروخانه را ندارد -
        // در نتیجه نه پاپ‌آپ هشدار خودکار می‌آمد و نه نشان روی دکمه‌ی «تاریخ نزدیک» به‌روز می‌شد،
        // مگر کاربر خودش یک‌بار وارد آن پنل می‌شد (که آن هم فقط نمایش را رفرش می‌کرد، نه یک
        // بررسی واقعی). حالا همین‌جا - دقیقاً لحظه‌ای که مشخص شد کاربر وارد کدام داروخانه شده -
        // هم نشان روی دکمه با اقلام همان داروخانه به‌روز می‌شود، هم یک بررسی کامل اجرا می‌شود که
        // در صورت لزوم خودش پاپ‌آپ هشدار را هم باز می‌کند.
        if (!string.IsNullOrWhiteSpace(newPharmacy))
        {
            RefreshExpiryWatchDisplayList();
            _ = CheckExpiryAlertsAsync();
        }

        UpdateTtacConnectionStatusUI();
        _ = CheckMonthlyArchiveReminderForCurrentPharmacyAsync();
    }

    private string GetArchiveStatePath()
    {
        return Path.Combine(AppContext.BaseDirectory, "archive-state.json");
    }

    private Dictionary<string, MonthlyArchiveState> LoadArchiveStateStore()
    {
        try
        {
            string path = GetArchiveStatePath();
            if (!File.Exists(path))
                return new Dictionary<string, MonthlyArchiveState>(StringComparer.OrdinalIgnoreCase);

            return JsonSerializer.Deserialize<Dictionary<string, MonthlyArchiveState>>(File.ReadAllText(path))
                   ?? new Dictionary<string, MonthlyArchiveState>(StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            return new Dictionary<string, MonthlyArchiveState>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private void SaveArchiveStateStore(Dictionary<string, MonthlyArchiveState> store)
    {
        try
        {
            File.WriteAllText(GetArchiveStatePath(), JsonSerializer.Serialize(store, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private string GetCurrentArchiveMonthKey()
    {
        return DateTime.Now.ToString("yyyy-MM", CultureInfo.InvariantCulture);
    }

    private string GetCurrentArchivePersianMonthKey()
    {
        DateTime now = DateTime.Now;
        return $"{_persianCalendar.GetYear(now):0000}-{_persianCalendar.GetMonth(now):00}";
    }

    private bool HasAnyCurrentPharmacyDataForArchive()
    {
        return HistoryItems.Count > 0
               || _ttacRegistrationHistory.Count > 0
               || ReceiveStatusItems.Count > 0
               || CargoDeliveryItems.Count > 0;
    }

    private async Task CheckMonthlyArchiveReminderForCurrentPharmacyAsync()
    {
        if (_isMonthlyArchivePromptOpen || string.IsNullOrWhiteSpace(_ttacPharmacyDisplayName))
            return;

        string pharmacyKey = GetReceiveStatusStorageKey(_ttacPharmacyDisplayName);
        string currentMonth = GetCurrentArchiveMonthKey();
        var store = LoadArchiveStateStore();

        if (!store.TryGetValue(pharmacyKey, out var state) || string.IsNullOrWhiteSpace(state.LastSeenMonth))
        {
            store[pharmacyKey] = new MonthlyArchiveState { LastSeenMonth = currentMonth };
            SaveArchiveStateStore(store);
            return;
        }

        if (string.Equals(state.LastSeenMonth, currentMonth, StringComparison.OrdinalIgnoreCase))
            return;

        string today = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        if (string.Equals(state.LastDismissedDate, today, StringComparison.OrdinalIgnoreCase))
            return;

        if (!HasAnyCurrentPharmacyDataForArchive())
        {
            state.LastSeenMonth = currentMonth;
            store[pharmacyKey] = state;
            SaveArchiveStateStore(store);
            return;
        }

        _isMonthlyArchivePromptOpen = true;
        try
        {
            string title = _localization.GetString("MonthlyArchive");
            string message = _localization.GetString("ANewMonthHasStartedDoYouWantScanbridgeToCreateAMultiSheetExcelArchiveForThePreviousDataAndStartThisMonthWithACleanWorkspaceRawDataWillAlsoBeKeptInArchive");

            var result = System.Windows.MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                bool archiveOk = await CreateMonthlyArchiveForCurrentPharmacyAsync(pharmacyKey, state.LastSeenMonth, currentMonth);
                if (archiveOk)
                {
                    state.LastArchivedMonth = state.LastSeenMonth;
                    state.LastSeenMonth = currentMonth;
                    state.LastDismissedDate = string.Empty;
                    store[pharmacyKey] = state;
                    SaveArchiveStateStore(store);
                    ShowStyledMessage(title, _localization.GetString("MonthlyArchiveWasCreatedSuccessfully"));
                }
                else
                {
                    // موفق نشد؛ داده‌ی فعلی دست‌نخورده باقی می‌ماند و LastSeenMonth تغییر
                    // نمی‌کند تا دوباره در فرصت بعدی به کاربر پیشنهاد آرشیو داده شود.
                    ShowStyledMessage(
                        _localization.GetString("ArchiveFailed"),
                        _localization.GetString("TheArchiveCouldNotBeFullyBackedUpAFileMayBeLockedOrTheDiskIsFullSoNothingWasDeletedPleaseTryAgain"),
                        true);
                }
            }
            else
            {
                state.LastDismissedDate = today;
                store[pharmacyKey] = state;
                SaveArchiveStateStore(store);
            }
        }
        catch (Exception ex)
        {
            ShowStyledMessage(_localization.GetString("ArchiveFailed"), ex.Message, true);
        }
        finally
        {
            _isMonthlyArchivePromptOpen = false;
        }
    }

    /// <summary>
    /// آرشیو ماهانه می‌سازد. اگر پشتیبان‌گیری خام (JSON/CSV) هرکدام از فایل‌ها شکست بخورد،
    /// false برمی‌گرداند و داده‌ی فعلی پاک نمی‌شود - قبلاً این متد صرف‌نظر از موفقیت
    /// CopyIfExists، بدون قید و شرط داده‌ی جاری را پاک می‌کرد (باگ ۵ گزارش).
    /// </summary>
    /// <summary>ماه قبل از monthKey (فرمت "yyyy-MM") را برمی‌گرداند؛ اگر پارس نشد خودِ monthKey را.</summary>
    private static string GetPredecessorMonthKey(string monthKey)
    {
        if (DateTime.TryParseExact(monthKey + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            return dt.AddMonths(-1).ToString("yyyy-MM", CultureInfo.InvariantCulture);
        return monthKey;
    }

    private async Task<bool> CreateMonthlyArchiveForCurrentPharmacyAsync(string pharmacyKey, string archiveMonth, string currentMonth)
    {
        await Task.Yield();
        string persianMonth = GetCurrentArchivePersianMonthKey();
        string safeArchiveMonth = string.IsNullOrWhiteSpace(archiveMonth) ? GetCurrentArchiveMonthKey() : archiveMonth;

        // اگر کاربر چندبار پشت‌سرهم «خیر» زده باشد، داده‌ی این آرشیو ممکن است چند ماه را با هم
        // پوشش بدهد (نه فقط safeArchiveMonth) - چون خودِ داده هرگز بین این ماه‌ها پاک نشده است.
        // قبلاً پوشه/فایل فقط با نام اولین ماه (safeArchiveMonth) ساخته می‌شد که گمراه‌کننده بود
        // (باگ ۸ گزارش)؛ الان اگر بازه بیش از یک ماه باشد، هر دو سر بازه در نام می‌آید.
        string coverageEndMonth = GetPredecessorMonthKey(currentMonth);
        string archiveLabel = string.Equals(safeArchiveMonth, coverageEndMonth, StringComparison.OrdinalIgnoreCase)
            ? safeArchiveMonth
            : $"{safeArchiveMonth}_to_{coverageEndMonth}";

        string archiveRoot = Path.Combine(AppContext.BaseDirectory, "Archive", pharmacyKey, archiveLabel);
        string rawDir = Path.Combine(archiveRoot, "Data");
        Directory.CreateDirectory(rawDir);

        string excelPath = Path.Combine(archiveRoot, $"Scanbridge_Archive_{pharmacyKey}_{persianMonth}_{archiveLabel}.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                AddArchiveScanHistorySheet(workbook);
                AddArchiveTtTeckSheet(workbook, false);
                AddArchiveTtTeckSheet(workbook, true);
                AddArchiveTtacRegistrationSheet(workbook);
                AddArchiveReceiveStatusSheet(workbook);
                AddArchiveCargoDeliverySheet(workbook);
                workbook.SaveAs(excelPath);
            }
        }
        catch
        {
            // اگر خود اکسل ترکیبی ساخته نشد، به هیچ وجه نباید داده‌ی جاری پاک شود.
            return false;
        }

        bool rawBackupOk = true;
        rawBackupOk &= CopyIfExists(GetActiveHistoryCsvPath(), Path.Combine(rawDir, Path.GetFileName(GetActiveHistoryCsvPath())));
        rawBackupOk &= CopyIfExists(GetTtacRegistrationHistoryPath(), Path.Combine(rawDir, "ttac-registration-history.json"));
        rawBackupOk &= CopyIfExists(GetReceiveStatusHistoryPath(), Path.Combine(rawDir, "receive-status-history.json"));
        rawBackupOk &= CopyIfExists(GetCargoDeliveryHistoryPath(), Path.Combine(rawDir, "cargo-delivery-history.json"));

        if (!rawBackupOk)
        {
            // خروجی اکسل ترکیبی ساخته شد ولی حداقل یکی از فایل‌های خام پشتیبان‌گیری نشد
            // (مثلاً به‌خاطر قفل‌بودن فایل توسط آنتی‌ویروس). داده‌ی جاری پاک نمی‌شود تا چیزی
            // از بین نرود؛ کاربر می‌تواند دوباره تلاش کند.
            return false;
        }

        HistoryItems.Clear();
        ApplyHistoryFilters();
        SaveHistoryItemsForKey(pharmacyKey);

        _ttacRegistrationHistory.Clear();
        SaveTtacRegistrationHistoryForKey(pharmacyKey);

        ReceiveStatusItems.Clear();
        _receiveStatusKnownBarcodes.Clear();
        SaveReceiveStatusItemsForKey(pharmacyKey);

        CargoDeliveryItems.Clear();
        _cargoDeliveryKnownBarcodes.Clear();
        SaveCargoDeliveryItemsForKey(pharmacyKey);

        return true;
    }

    private static bool CopyIfExists(string source, string destination)
    {
        try
        {
            if (File.Exists(source))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                File.Copy(source, destination, true);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private void AddArchiveScanHistorySheet(XLWorkbook workbook)
    {
        var ws = workbook.Worksheets.Add("Scan History");
        string[] headers = { "ردیف", "تاریخ", "ساعت", "دستگاه", "بارکد", "محصول / وضعیت" };
        WriteArchiveHeader(ws, headers);
        int r = 2, i = 1;
        foreach (var item in HistoryItems.OrderByDescending(x => x.TimestampLocal))
        {
            ws.Cell(r, 1).Value = i++;
            ws.Cell(r, 2).Value = item.PersianDateText;
            ws.Cell(r, 3).Value = item.TimeText;
            ws.Cell(r, 4).Value = item.DeviceName;
            ws.Cell(r, 5).Value = item.Barcode;
            ws.Cell(r, 6).Value = item.DrugName;
            r++;
        }
        ws.Columns().AdjustToContents();
    }

    private void AddArchiveTtTeckSheet(XLWorkbook workbook, bool formulaOnly)
    {
        var ws = workbook.Worksheets.Add(formulaOnly ? "Formula" : "TtTeck Items");
        string[] headers = { "ردیف", "تاریخ", "ساعت", "دستگاه", "بارکد / UID", "نام فارسی", "نام انگلیسی", "وضعیت ثبت" };
        WriteArchiveHeader(ws, headers);
        int r = 2, i = 1;
        foreach (var item in HistoryItems.OrderByDescending(x => x.TimestampLocal))
        {
            if (!IsTtTeckHistoryRecord(item))
                continue;
            if (formulaOnly && !IsInfantFormulaRecord(item))
                continue;
            var row = CreateTtTeckHistoryRowFromRecord(item);
            ws.Cell(r, 1).Value = i++;
            ws.Cell(r, 2).Value = row.PersianDateText;
            ws.Cell(r, 3).Value = row.TimeText;
            ws.Cell(r, 4).Value = row.DeviceName;
            ws.Cell(r, 5).Value = row.Barcode;
            ws.Cell(r, 6).Value = row.PersianProductName;
            ws.Cell(r, 7).Value = row.EnglishProductName;
            ws.Cell(r, 8).Value = row.RegistrationButtonText;
            r++;
        }
        ws.Columns().AdjustToContents();
    }

    private void AddArchiveTtacRegistrationSheet(XLWorkbook workbook)
    {
        var ws = workbook.Worksheets.Add("TTAC Registrations");
        string[] headers = { "تاریخ ثبت", "محصول", "نوع ثبت", "شناسه نسخه", "تعداد", "کد ملی", "موبایل", "بیمار", "نتیجه", "پیام" };
        WriteArchiveHeader(ws, headers);
        int r = 2;
        foreach (var item in _ttacRegistrationHistory.OrderByDescending(x => x.RegisteredAt))
        {
            ws.Cell(r, 1).Value = item.RegisteredAt.ToString("yyyy/MM/dd HH:mm:ss");
            ws.Cell(r, 2).Value = item.ProductName;
            ws.Cell(r, 3).Value = item.RegistrationType;
            ws.Cell(r, 4).Value = item.PrescriptionId?.ToString() ?? "";
            ws.Cell(r, 5).Value = item.Amount;
            ws.Cell(r, 6).Value = string.IsNullOrWhiteSpace(item.NationalIdFull) ? item.NationalIdMasked : item.NationalIdFull;
            ws.Cell(r, 7).Value = string.IsNullOrWhiteSpace(item.MobileFull) ? item.MobileMasked : item.MobileFull;
            ws.Cell(r, 8).Value = item.PatientFullName;
            ws.Cell(r, 9).Value = item.Success ? "موفق" : "ناموفق";
            ws.Cell(r, 10).Value = item.Message;
            r++;
        }
        ws.Columns().AdjustToContents();
    }

    private void AddArchiveReceiveStatusSheet(XLWorkbook workbook)
    {
        var ws = workbook.Worksheets.Add("Receive Status");
        string[] headers = { "محصول", "بارکد / UID", "IRC", "سری ساخت", "تعداد", "ارسال‌کننده", "تاریخ ارسال", "وضعیت" };
        WriteArchiveHeader(ws, headers);
        int r = 2;
        foreach (var row in ReceiveStatusItems)
        {
            ws.Cell(r, 1).Value = row.ProductName;
            ws.Cell(r, 2).Value = row.Barcode;
            ws.Cell(r, 3).Value = row.Irc;
            ws.Cell(r, 4).Value = row.LotNumber;
            ws.Cell(r, 5).Value = row.Quantity;
            ws.Cell(r, 6).Value = row.SenderName;
            ws.Cell(r, 7).Value = row.SentDatePersian;
            ws.Cell(r, 8).Value = row.StatusText;
            r++;
        }
        ws.Columns().AdjustToContents();
    }

    private void AddArchiveCargoDeliverySheet(XLWorkbook workbook)
    {
        var ws = workbook.Worksheets.Add("Cargo Delivery");
        string[] headers = { "محصول", "بارکد / UID", "IRC", "سری ساخت", "تعداد", "ارسال‌کننده", "وضعیت", "انتخاب‌شده" };
        WriteArchiveHeader(ws, headers);
        int r = 2;
        foreach (var row in CargoDeliveryItems)
        {
            ws.Cell(r, 1).Value = row.ProductName;
            ws.Cell(r, 2).Value = row.Barcode;
            ws.Cell(r, 3).Value = row.Irc;
            ws.Cell(r, 4).Value = row.LotNumber;
            ws.Cell(r, 5).Value = row.Quantity;
            ws.Cell(r, 6).Value = row.SenderName;
            ws.Cell(r, 7).Value = row.StatusText;
            ws.Cell(r, 8).Value = row.IsSelected ? "بله" : "خیر";
            r++;
        }
        ws.Columns().AdjustToContents();
    }

    private static void WriteArchiveHeader(IXLWorksheet ws, string[] headers)
    {
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(26, 35, 126);
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
    }

    private async Task<string?> TryReadTtacTokenFromWebViewStorageAsync()
    {
        if (_ttTeckWebView?.CoreWebView2 == null)
            return null;

        string script = string.Join("\n", new[]
        {
            "(function(){",
            "  function findToken(storage){",
            "    for(let i=0;i<storage.length;i++){",
            "      const k=storage.key(i);",
            "      const v=storage.getItem(k);",
            "      if(!v) continue;",
            "      if(k && k.toLowerCase().includes('access_token')) return v;",
            "      try{",
            "        const o=JSON.parse(v);",
            "        if(o && o.access_token) return o.access_token;",
            "        if(o && o.accessToken) return o.accessToken;",
            "      }catch(e){}",
            "      const m=v.match(/\\\"access_token\\\"\\s*:\\s*\\\"([^\\\"]+)\\\"/);",
            "      if(m) return m[1];",
            "    }",
            "    return null;",
            "  }",
            "  return findToken(window.localStorage) || findToken(window.sessionStorage) || null;",
            "})();",
        });

        string resultJson = await _ttTeckWebView.CoreWebView2.ExecuteScriptAsync(script);
        return JsonSerializer.Deserialize<string?>(resultJson);
    }

    private async Task<string?> GetTtacAccessTokenOnUiThreadAsync(bool openLoginIfMissing = true)
    {
        if (Dispatcher.CheckAccess())
            return await GetTtacAccessTokenAsync(openLoginIfMissing);

        var tokenTask = await Dispatcher.InvokeAsync(() => GetTtacAccessTokenAsync(openLoginIfMissing));
        return await tokenTask;
    }

    private async Task<string?> GetTtacAccessTokenAsync(bool openLoginIfMissing = true)
    {
        bool sessionExpired = false;
        try
        {
            if (!ShouldIgnoreStaleTtacToken())
            {
                string? webViewToken = await TryReadTtacTokenFromWebViewStorageAsync();
                if (!string.IsNullOrWhiteSpace(webViewToken))
                {
                    ApplyTtacAccessToken(webViewToken, DateTime.UtcNow.AddMinutes(90));
                    return _ttacAccessTokenOverride;
                }
            }
        }
        catch { }

        if (!string.IsNullOrWhiteSpace(_ttacAccessTokenOverride) && DateTime.UtcNow < _ttacAccessTokenExpiresAtUtc)
        {
            if (string.IsNullOrWhiteSpace(_ttacPharmacyDisplayName))
                _ttacPharmacyDisplayName = TryExtractTtacDisplayNameFromToken(_ttacAccessTokenOverride) ?? string.Empty;
            UpdateTtacConnectionStatusUI();
            return _ttacAccessTokenOverride;
        }

        if (!string.IsNullOrWhiteSpace(_ttacAccessTokenOverride) && DateTime.UtcNow >= _ttacAccessTokenExpiresAtUtc)
        {
            _ttacAccessTokenOverride = null;
            _ttacPharmacyDisplayName = string.Empty;
            _ttacAccessTokenExpiresAtUtc = DateTime.MinValue;
            UpdateTtacTokenValidityTracking(false);
            sessionExpired = true;
            UpdateTtacConnectionStatusUI();
        }

        try
        {
            if (TtTeckWebView.CoreWebView2 == null && openLoginIfMissing)
                await OpenTtTeckInternalBrowserAsync("https://newstatisticsreports.ttac.ir/pharmacyDashboard");

            if (TtTeckWebView.CoreWebView2 == null)
                return null;

            string script = string.Join("\n", new[]
            {
                "(function(){",
                "  function findToken(storage){",
                "    for(let i=0;i<storage.length;i++){",
                "      const k=storage.key(i);",
                "      const v=storage.getItem(k);",
                "      if(!v) continue;",
                "      if(k && k.toLowerCase().includes('access_token')) return v;",
                "      try{",
                "        const o=JSON.parse(v);",
                "        if(o && o.access_token) return o.access_token;",
                "        if(o && o.accessToken) return o.accessToken;",
                "      }catch(e){}",
                "      const m=v.match(/\\\"access_token\\\"\\\\s*:\\\\s*\\\"([^\\\"]+)\\\"/);",
                "      if(m) return m[1];",
                "    }",
                "    return null;",
                "  }",
                "  return findToken(window.localStorage) || findToken(window.sessionStorage) || null;",
                "})();",
            });

            if (!ShouldIgnoreStaleTtacToken())
            {
                string resultJson = await TtTeckWebView.CoreWebView2.ExecuteScriptAsync(script);
                string? token = JsonSerializer.Deserialize<string?>(resultJson);
                if (!string.IsNullOrWhiteSpace(token))
                {
                    ApplyTtacAccessToken(token, DateTime.UtcNow.AddMinutes(90));
                    return _ttacAccessTokenOverride;
                }
            }
        }
        catch { }

        if (openLoginIfMissing)
        {
            ShowTtacLoginOverlay(sessionExpired);
        }

        UpdateTtacConnectionStatusUI();
        return null;
    }

    private async Task<HttpRequestMessage?> CreateTtacRequestAsync(HttpMethod method, string url, object? body = null)
    {
        string? token = await GetTtacAccessTokenAsync();
        if (string.IsNullOrWhiteSpace(token))
            return null;

        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.TryAddWithoutValidation("Origin", "https://newstatisticsreports.ttac.ir");
        request.Headers.TryAddWithoutValidation("Referer", "https://newstatisticsreports.ttac.ir/");
        request.Headers.TryAddWithoutValidation("Accept", "application/json, text/plain, */*");

        if (body != null)
        {
            string json = JsonSerializer.Serialize(body);
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");
        }

        return request;
    }

    private string GetFriendlyTtacConnectionErrorMessage(Exception ex)
    {
        bool english = _localization.CurrentLanguage == AppLanguage.English;
        if (ex is TaskCanceledException || ex is TimeoutException)
        {
            return _localization.GetString("TTACIsRespondingSlowlyOrDidNotRespondInTimeThisIsUsuallyCausedByHighTTACTrafficSlowInternetVPNProxyDNSIssuesOrTemporaryTTACOutagePleaseWaitAMomentAndTryAgain");
        }

        if (ex is HttpRequestException || ex.Message.Contains("No such host", StringComparison.OrdinalIgnoreCase) || ex.Message.Contains("Name or service not known", StringComparison.OrdinalIgnoreCase))
        {
            return _localization.GetString("ConnectionToTTACFailedCheckInternetDNSVPNProxyOrTTACAvailability");
        }

        return ex.Message;
    }

    private async Task<JsonDocument?> SendTtacJsonAsync(HttpMethod method, string url, object? body = null)
    {
        try
        {
            var request = await CreateTtacRequestAsync(method, url, body);
            if (request == null)
                throw new TtacSessionExpiredException(_localization.GetString("TTACLoginIsRequired"));

            var response = await _ttacHttpClient.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized || response.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                throw new TtacSessionExpiredException(_localization.GetString("TtTeckSessionExpiredPleaseLoginAgain"));
            }

            if (!response.IsSuccessStatusCode)
                throw new InvalidOperationException(ExtractTtacErrorMessage(content, $"HTTP {(int)response.StatusCode}"));

            if (string.IsNullOrWhiteSpace(content))
                return null;

            var doc = JsonDocument.Parse(content);
            if (doc.RootElement.TryGetProperty("Success", out var successProp) && successProp.ValueKind == JsonValueKind.False)
            {
                string message = ExtractTtacErrorMessage(content, _localization.GetString("RequestFailed"));
                doc.Dispose();
                throw new InvalidOperationException(message);
            }

            return doc;
        }
        catch (TtacSessionExpiredException)
        {
            throw;
        }
        catch (TaskCanceledException ex)
        {
            throw new TimeoutException(GetFriendlyTtacConnectionErrorMessage(ex), ex);
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException(GetFriendlyTtacConnectionErrorMessage(ex), ex);
        }
    }

    private string ExtractTtacErrorMessage(string? json, string fallback)
    {
        if (string.IsNullOrWhiteSpace(json))
            return fallback;

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            string? message = ReadJsonString(root, "Message");
            string? statusCode = ReadJsonString(root, "StatusCode");

            if (!string.IsNullOrWhiteSpace(message))
            {
                if (!string.IsNullOrWhiteSpace(statusCode) && statusCode != "null")
                    return _localization.GetFormattedString("StatusCodeFormat", message, statusCode);

                return message;
            }
        }
        catch { }

        return fallback + (json.Length > 300 ? " " + json.Substring(0, 300) : " " + json);
    }

    private async Task LoadTtacCaptchaAsync(bool focusCaptchaAfterLoad = false)
    {
        UpdateTtacRegistrationStageButtons(true);
        try
        {
            using var doc = await SendTtacJsonAsync(HttpMethod.Get, "https://statisticsreports.ttac.ir/captcha/generate");
            if (doc == null)
                return;

            var root = doc.RootElement;
            if (!root.TryGetProperty("Result", out var result))
                result = root;

            string captchaId = ReadJsonString(result, "CaptchaId") ?? ReadJsonString(result, "captchaId") ?? string.Empty;
            string imageData = ReadJsonString(result, "ImageData") ?? ReadJsonString(result, "imageData") ?? ReadJsonString(result, "Image") ?? string.Empty;

            if (string.IsNullOrWhiteSpace(captchaId) || string.IsNullOrWhiteSpace(imageData))
                throw new InvalidOperationException("Captcha response was not recognized.");

            _ttacCurrentCaptchaId = captchaId;
            TtTeckRegistrationCaptchaImage.Source = LoadBitmapImage(Convert.FromBase64String(imageData));
            TtTeckRegistrationCaptchaTextBox.Text = string.Empty;
            TtTeckRegistrationResultText.Text = _localization.GetString("CaptchaReceivedEnterTheCodeAndCreatePrescription");
            ValidateTtacRegistrationFields();
            UpdateTtacRegistrationStageButtons();
            NotifyRemoteEntryCaptchaLoaded(imageData);
        }
        catch (Exception ex)
        {
            HandleTtacOperationException(ex, _localization.GetString("CaptchaFailed"), () => LoadTtacCaptchaAsync(focusCaptchaAfterLoad), pendingLabel: _localization.GetString("PendingLoadCaptcha"));
        }
        finally
        {
            UpdateTtacRegistrationStageButtons(false);
            if (focusCaptchaAfterLoad && TtTeckRegistrationCaptchaTextBox != null)
                FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
        }
    }

    private static string? ReadJsonString(JsonElement element, string name)
    {
        if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var prop))
            return prop.ValueKind == JsonValueKind.String ? prop.GetString() : prop.ToString();
        return null;
    }

    private static long? FindLongRecursive(JsonElement element, string propertyName)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var p in element.EnumerateObject())
            {
                if (p.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    if (p.Value.ValueKind == JsonValueKind.Number && p.Value.TryGetInt64(out long n)) return n;
                    if (long.TryParse(p.Value.ToString(), out n)) return n;
                }
                var child = FindLongRecursive(p.Value, propertyName);
                if (child.HasValue) return child;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var child = FindLongRecursive(item, propertyName);
                if (child.HasValue) return child;
            }
        }
        return null;
    }

    // جستجوی بازگشتی برای یک فیلد رشته‌ای (مثل UID) در یک شیء/آرایه JSON با نام‌های احتمالی مختلف.
    // برای تشخیص اینکه یک آیتم از لیست «قابل‌تأیید تی‌تک» دقیقاً مربوط به همان بسته‌ی فیزیکی
    // اسکن‌شده است یا نه (نه صرفاً اولین آیتم با همان محصول/سری‌ساخت).
    private static string? FindStringRecursive(JsonElement element, params string[] propertyNames)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var p in element.EnumerateObject())
            {
                foreach (var name in propertyNames)
                {
                    if (p.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && p.Value.ValueKind == JsonValueKind.String)
                    {
                        var value = p.Value.GetString();
                        if (!string.IsNullOrWhiteSpace(value))
                            return value;
                    }
                }

                var child = FindStringRecursive(p.Value, propertyNames);
                if (!string.IsNullOrWhiteSpace(child))
                    return child;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var child = FindStringRecursive(item, propertyNames);
                if (!string.IsNullOrWhiteSpace(child))
                    return child;
            }
        }

        return null;
    }

    private string GetTtacBirthDateText()
    {
        string dayText = ToEnglishDigits(TtTeckBirthDayTextBox.Text.Trim());
        string monthText = ToEnglishDigits(TtTeckBirthMonthTextBox.Text.Trim());
        string yearText = ToEnglishDigits(TtTeckBirthYearTextBox.Text.Trim());

        if (int.TryParse(dayText, out int day) &&
            int.TryParse(monthText, out int month) &&
            int.TryParse(yearText, out int year))
        {
            string normalized = $"{year:0000}/{month:00}/{day:00}";
            TtTeckRegistrationBirthDateTextBox.Text = normalized;
            return ConvertBirthDateToTtacApiDate(normalized);
        }

        string inputDate = NormalizeDateInput(TtTeckRegistrationBirthDateTextBox?.Text);
        if (!string.IsNullOrWhiteSpace(inputDate))
            return ConvertBirthDateToTtacApiDate(inputDate);

        return string.Empty;
    }

    private string ConvertBirthDateToTtacApiDate(string normalizedDate)
    {
        var parts = normalizedDate.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 3 ||
            !int.TryParse(parts[0], out int year) ||
            !int.TryParse(parts[1], out int month) ||
            !int.TryParse(parts[2], out int day))
        {
            return string.Empty;
        }

        // نکته مهم: سایت تاریخ را شمسی نمایش می‌دهد، اما در API مقدار میلادی ارسال می‌کند.
        // مثال مشاهده‌شده در HAR: 1379/06/20 در سایت => 2000/09/10 در درخواست API.
        if (year >= 1200 && year <= 1500)
        {
            try
            {
                DateTime gregorian = _persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
                return gregorian.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
            }
            catch (ArgumentOutOfRangeException)
            {
                // تاریخ شمسیِ نامعتبر (مثلاً 31 مهر یا 30 اسفند سال غیرکبیسه) —
                // به‌جای کرشِ فرایند ثبت، تاریخ ارسال نمی‌شود.
                return string.Empty;
            }
        }

        // اگر کاربر خودش تاریخ میلادی وارد کرد، همان فرمت استاندارد به API ارسال می‌شود.
        return $"{year:0000}/{month:00}/{day:00}";
    }

    private static string NormalizeDateInput(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return string.Empty;

        string value = ToEnglishDigits(input.Trim())
            .Replace('-', '/')
            .Replace('.', '/')
            .Replace('\\', '/');

        var parts = value.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 3)
            return string.Empty;

        if (!int.TryParse(parts[0], out int year) ||
            !int.TryParse(parts[1], out int month) ||
            !int.TryParse(parts[2], out int day))
        {
            return string.Empty;
        }

        if (month < 1 || month > 12 || day < 1 || day > 31)
            return string.Empty;

        // سال شمسی یا میلادی را قبول می‌کنیم.
        if (!((year >= 1200 && year <= 1500) || (year >= 1900 && year <= 2100)))
            return string.Empty;

        return $"{year:0000}/{month:00}/{day:00}";
    }

    private static string ToEnglishDigits(string input)
    {
        return input
            .Replace('۰', '0').Replace('۱', '1').Replace('۲', '2').Replace('۳', '3').Replace('۴', '4')
            .Replace('۵', '5').Replace('۶', '6').Replace('۷', '7').Replace('۸', '8').Replace('۹', '9')
            .Replace('٠', '0').Replace('١', '1').Replace('٢', '2').Replace('٣', '3').Replace('٤', '4')
            .Replace('٥', '5').Replace('٦', '6').Replace('٧', '7').Replace('٨', '8').Replace('٩', '9');
    }

    private string ExtractPatientFullNameFromTtacDeclareResponse(JsonElement root)
    {
        try
        {
            string first = FindJsonStringRecursive(root, "FirstName", "firstName") ?? string.Empty;
            string last = FindJsonStringRecursive(root, "LastName", "lastName") ?? string.Empty;
            string full = (first + " " + last).Trim();
            return full;
        }
        catch
        {
            return string.Empty;
        }
    }

    private async void TtTeckRegistrationCreatePrescriptionButton_Click(object sender, RoutedEventArgs e)
    {
        await CreateCurrentTtacPrescriptionAsync();
    }

    // بدنه‌ی همان دکمه‌ی «ایجاد نسخه»، جدا شده به یک متد قابل‌await تا هم از کلیک محلی هم از
    // جریان «ورود اطلاعات از راه دور» (وقتی دکمه‌ی نهایی روی گوشی زده می‌شود) با یک منطق واحد صدا
    // زده شود - نگاه کنید به MainWindow.RemoteFormulaEntry.cs.
    private async Task CreateCurrentTtacPrescriptionAsync()
    {
        UpdateTtacRegistrationStageButtons(true);
        try
        {
            string nationalId = TtTeckRegistrationNationalIdTextBox.Text.Trim();
            string captchaInput = TtTeckRegistrationCaptchaTextBox.Text.Trim();
            bool isElectronic = TtTeckRegistrationTypeComboBox.SelectedIndex == 1;

            if (string.IsNullOrWhiteSpace(nationalId) || string.IsNullOrWhiteSpace(captchaInput) || string.IsNullOrWhiteSpace(_ttacCurrentCaptchaId))
                throw new InvalidOperationException(_localization.GetString("NationalIDAndCaptchaAreRequired"));

            _ttacCurrentNationalId = nationalId;
            _ttacCurrentBirthDate = GetTtacBirthDateText();
            if (string.IsNullOrWhiteSpace(_ttacCurrentBirthDate))
                throw new InvalidOperationException(_localization.GetString("EnterBirthDateInTheCorrectFormatPersianExample13790620"));

            _ttacCurrentIsElectronic = isElectronic;
            TtTeckRegistrationResultText.Text = _localization.GetFormattedString("BirthDateInfo", NormalizeDateInput(TtTeckRegistrationBirthDateTextBox?.Text), _ttacCurrentBirthDate);

            object body;
            string url;
            if (isElectronic)
            {
                string medicalCode = TtTeckRegistrationMedicalCouncilTextBox.Text.Trim();
                if (string.IsNullOrWhiteSpace(medicalCode))
                    throw new InvalidOperationException(_localization.GetString("MedicalCouncilNumberIsRequired"));

                if (!medicalCode.StartsWith("MD_", StringComparison.OrdinalIgnoreCase))
                    medicalCode = "MD_" + medicalCode;

                url = "https://statisticsreports.ttac.ir/prescription/declare";
                body = new
                {
                    medicalCouncilCode = medicalCode,
                    insurance = 37,
                    captchaId = _ttacCurrentCaptchaId,
                    captchaInput,
                    type = 1,
                    nationalId,
                    birthDate = _ttacCurrentBirthDate
                };
            }
            else
            {
                url = "https://statisticsreports.ttac.ir/prescription/declareNonePrescription";
                body = new
                {
                    insurance = 37,
                    captchaId = _ttacCurrentCaptchaId,
                    captchaInput,
                    type = 1,
                    nationalId,
                    birthDate = _ttacCurrentBirthDate
                };
            }

            using var doc = await SendTtacJsonAsync(HttpMethod.Post, url, body);
            if (doc == null)
                return;

            bool success = doc.RootElement.TryGetProperty("Success", out var successProp) && successProp.ValueKind == JsonValueKind.True;
            if (!success)
                throw new InvalidOperationException(doc.RootElement.TryGetProperty("Message", out var msg) ? msg.ToString() : (_localization.GetString("CreatingPrescriptionFailed")));

            _ttacCurrentPrescriptionId = FindLongRecursive(doc.RootElement, "prescriptionId");
            _ttacCurrentPatientFullName = ExtractPatientFullNameFromTtacDeclareResponse(doc.RootElement);
            if (!_ttacCurrentPrescriptionId.HasValue)
                throw new InvalidOperationException(_localization.GetString("PrescriptionIDWasNotFoundInThePortalResponse"));

            TtTeckRegistrationResultText.Text = _localization.GetFormattedString("PrescriptionCreated", _ttacCurrentPrescriptionId.Value);
            AddTtacRegistrationLog(true, _localization.GetString("CreatePrescription"), _localization.GetFormattedString("PrescriptionIdFormat", _ttacCurrentPrescriptionId.Value));
            UpdateTtacRegistrationStageButtons(false);
            TtTeckRegistrationSubmitItemButton.Focus();
        }
        catch (Exception ex)
        {
            bool isFormulaRegistrationForAlert = IsCurrentRegistrationFormulaItem();
            HandleTtacOperationException(ex, _localization.GetString("CreatePrescriptionFailed"), async () =>
            {
                await CreateCurrentTtacPrescriptionAsync();
            }, pendingLabel: _localization.GetString("PendingCreatePrescription"),
            onFailureMessageShown: (shownTitle, shownMessage) =>
            {
                // فقط برای شیرخشک، و فقط وقتی واقعاً یک پیام خطای نهایی نشان داده شده (نه در حالت
                // نشست‌منقضی‌شده که خودش خودکار دوباره تلاش می‌کند). این خطا (مثلاً کپچای اشتباه یا
                // کد ملی نامعتبر) اگر ورود از راه دور فعال بود، روی گوشی هم نشان داده می‌شود - بدون
                // پایان‌دادن به کل جریان، تا کاربر بتواند با دکمه‌ی «قبلی» برگردد و فیلد را اصلاح کند.
                if (isFormulaRegistrationForAlert)
                {
                    ShowRemoteEntryErrorAndAllowRetry(shownTitle, shownMessage);
                }
            });
        }
        finally
        {
            UpdateTtacRegistrationStageButtons(false);
        }
    }

    private bool IsCurrentRegistrationFormulaItem()
    {
        return _pendingRegistrationTtTeckRow != null && GetFormulaRegistrationModeForRow(_pendingRegistrationTtTeckRow) != FormulaRegistrationMode.Unknown;
    }

    private TtacRepeatFormulaContext CreateCurrentFormulaRepeatContext()
    {
        return new TtacRepeatFormulaContext
        {
            Amount = string.IsNullOrWhiteSpace(TtTeckRegistrationAmountTextBox.Text) ? "1" : ToEnglishDigits(TtTeckRegistrationAmountTextBox.Text.Trim()),
            NationalId = ToEnglishDigits(TtTeckRegistrationNationalIdTextBox.Text.Trim()),
            BirthDay = ToEnglishDigits(TtTeckBirthDayTextBox.Text.Trim()),
            BirthMonth = ToEnglishDigits(TtTeckBirthMonthTextBox.Text.Trim()),
            BirthYear = ToEnglishDigits(TtTeckBirthYearTextBox.Text.Trim()),
            Mobile = ToEnglishDigits(TtTeckRegistrationMobileTextBox.Text.Trim()),
            MedicalCouncil = TtTeckRegistrationMedicalCouncilTextBox.Text.Trim(),
            IsElectronic = TtTeckRegistrationTypeComboBox.SelectedIndex == 1
        };
    }

    private void ApplyFormulaRepeatContextToOpenForm(TtacRepeatFormulaContext context)
    {
        TtTeckRegistrationAmountTextBox.Text = string.IsNullOrWhiteSpace(context.Amount) ? "1" : context.Amount;
        TtTeckRegistrationNationalIdTextBox.Text = context.NationalId;
        TtTeckBirthDayTextBox.Text = context.BirthDay;
        TtTeckBirthMonthTextBox.Text = context.BirthMonth;
        TtTeckBirthYearTextBox.Text = context.BirthYear;
        TtTeckRegistrationBirthDateTextBox.Text = $"{context.BirthYear}/{context.BirthMonth}/{context.BirthDay}";
        TtTeckRegistrationMobileTextBox.Text = context.Mobile;
        TtTeckRegistrationMedicalCouncilTextBox.Text = context.MedicalCouncil;
        UpdateTtTeckRegistrationTypeButtons();
        ValidateTtacRegistrationFields();
        UpdateTtacRegistrationStageButtons();
    }

    private async Task<string> CompletePrescriptionAndSendSmsIfNeededAsync()
    {
        if (!_ttacCurrentPrescriptionId.HasValue)
            return string.Empty;

        string mobile = ToEnglishDigits(TtTeckRegistrationMobileTextBox.Text.Trim());
        if (string.IsNullOrWhiteSpace(mobile))
            return string.Empty;

        var body = new
        {
            prescriptionId = _ttacCurrentPrescriptionId.Value,
            promptModel = mobile
        };

        using var doc = await SendTtacJsonAsync(HttpMethod.Post, "https://statisticsreports.ttac.ir/prescription/CompletePrescriptionAndSendSMS", body);
        string? message = doc == null ? null : ReadJsonString(doc.RootElement, "Message");
        if (!string.IsNullOrWhiteSpace(message) && message != "null")
            return message;

        return _localization.GetString("CompletionSMSWasSent");
    }

    private async void TtTeckRegistrationSubmitItemButton_Click(object sender, RoutedEventArgs e)
    {
        await SubmitCurrentTtacItemAsync(viaRemoteEntry: false);
    }

    // بدنه‌ی همان دکمه‌ی «ثبت قلم»، جدا شده به یک متد قابل‌await با یک پارامتر جدید
    // (viaRemoteEntry) تا هم از کلیک محلی هم از جریان «ورود اطلاعات از راه دور» (گوشی) با یک
    // منطق واحد صدا زده شود. تنها رفتار متفاوت در دو مسیر، سناریوی نادرِ «احتمال ثبت تکراری»
    // است (نگاه کنید پایین‌تر) - نگاه کنید به MainWindow.RemoteFormulaEntry.cs.
    private async Task SubmitCurrentTtacItemAsync(bool viaRemoteEntry)
    {
        UpdateTtacRegistrationStageButtons(true);
        // بیرون از try تعریف می‌شود تا در catch هم در دسترس باشد (برای پاک‌کردن کلید در
        // شکست‌های قطعی که درخواست به تی‌تک نرسیده).
        string? submissionKey = null;
        try
        {
            if (_pendingRegistrationTtTeckRow == null)
                throw new InvalidOperationException(_localization.GetString("NoBarcodeIsSelected"));

            if (!_ttacCurrentPrescriptionId.HasValue)
                throw new InvalidOperationException(_localization.GetString("CreateThePrescriptionFirst"));

            bool isFormulaRegistration = IsCurrentRegistrationFormulaItem();
            string mobile = ToEnglishDigits(TtTeckRegistrationMobileTextBox.Text.Trim());
            if (isFormulaRegistration && string.IsNullOrWhiteSpace(mobile))
                throw new InvalidOperationException(_localization.GetString("MobileNumberIsRequiredForFormulaRegistration"));
            if (!string.IsNullOrWhiteSpace(mobile) && (!mobile.All(char.IsDigit) || mobile.Length != 11 || !mobile.StartsWith("09")))
                throw new InvalidOperationException(_localization.GetString("MobileNumberIsNotValid"));

            string amount = string.IsNullOrWhiteSpace(TtTeckRegistrationAmountTextBox.Text) ? "1" : TtTeckRegistrationAmountTextBox.Text.Trim();
            string url = _ttacCurrentIsElectronic
                ? "https://statisticsreports.ttac.ir/prescription/checkUid"
                : "https://statisticsreports.ttac.ir/prescription/checknoneprescriptionuid";

            // کلید یکتا برای این تلاش ثبت. اگر همین ترکیب نسخه/قلم/تعداد قبلاً یک‌بار به تی‌تک
            // ارسال شده باشد (یعنی الان داریم بعد از قطع/وصل نشست دوباره اینجا رسیده‌ایم)، به‌جای
            // ارسال خودکار و بی‌صدا، از کاربر تأیید صریح گرفته می‌شود تا خطر ثبت دوباره‌ی همان
            // قلم روی سامانه تی‌تک از بین برود.
            submissionKey = $"{_ttacCurrentPrescriptionId.Value}|{_pendingRegistrationTtTeckRow.Barcode}|{amount}";
            if (_ttacSubmittedItemKeys.Contains(submissionKey))
            {
                if (viaRemoteEntry)
                {
                    // وقتی از گوشی ثبت می‌شود، کسی پای سیستم نیست که به این دیالوگ تایید بدهد؛ طبق
                    // تصمیم صریح کاربر، این سناریوی نادر خودکار ادامه پیدا نمی‌کند - پیام روی گوشی
                    // نشان داده می‌شود و کاربر باید خودش دوباره دکمه‌ی ثبت را بزند (یا با «قبلی»
                    // برگردد و چیزی را اصلاح کند)؛ جریان پایان نمی‌یابد.
                    AddTtacRegistrationLog(false,
                        _localization.GetString("SubmitItem"),
                        _localization.GetString("AutomaticResendSkippedByTheUserToAvoidADuplicateTTACRegistration"));
                    ShowRemoteEntryErrorAndAllowRetry(
                        _localization.GetString("PossibleDuplicateSubmission"),
                        _localization.GetString("ARegistrationRequestForThisExactItemAndPrescriptionMayHaveAlreadyReachedTTACOnceBeforeEGTheSessionExpiredRightAfterSendingSendingItAgainCouldRegisterItTwiceSendAgainAnyway"));
                    return;
                }

                var confirmResult = System.Windows.MessageBox.Show(
                    this,
                    _localization.GetString("ARegistrationRequestForThisExactItemAndPrescriptionMayHaveAlreadyReachedTTACOnceBeforeEGTheSessionExpiredRightAfterSendingSendingItAgainCouldRegisterItTwiceSendAgainAnyway"),
                    _localization.GetString("PossibleDuplicateSubmission"),
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Warning);

                if (confirmResult != System.Windows.MessageBoxResult.Yes)
                {
                    AddTtacRegistrationLog(false,
                        _localization.GetString("SubmitItem"),
                        _localization.GetString("AutomaticResendSkippedByTheUserToAvoidADuplicateTTACRegistration"));
                    return;
                }
            }

            _ttacSubmittedItemKeys.Add(submissionKey);

            var body = new
            {
                UID = _pendingRegistrationTtTeckRow.Barcode,
                Amount = amount,
                SellingMethod = false,
                PrescriptionId = _ttacCurrentPrescriptionId.Value,
                NationalId = _ttacCurrentNationalId,
                BirthDate = _ttacCurrentBirthDate,
                Type = 1
            };

            using var doc = await SendTtacJsonAsync(HttpMethod.Post, url, body);
            string successMessage = ExtractTtacSuccessSummary(doc?.RootElement);
            bool canRepeatFormulaRegistration = isFormulaRegistration;
            var repeatContext = canRepeatFormulaRegistration ? CreateCurrentFormulaRepeatContext() : null;
            // پیش از بستن پنل ثبت (که ممکن است قلم در حال ثبتِ بعدی را عوض کند) مسیر عکس همین
            // قلم شیرخشک را نگه می‌داریم تا در دیالوگ «ثبت شد» نشان داده شود.
            string? formulaPhotoPath = canRepeatFormulaRegistration && _pendingRegistrationTtTeckRow != null
                ? GetFormulaPhotoPathForBarcode(_pendingRegistrationTtTeckRow.Barcode)
                : null;
            string smsMessage = await CompletePrescriptionAndSendSmsIfNeededAsync();
            if (!string.IsNullOrWhiteSpace(smsMessage))
                successMessage += Environment.NewLine + smsMessage;
            TtTeckRegistrationResultText.Text = successMessage;
            AddTtacRegistrationLog(true, _localization.GetString("SubmitItem"), successMessage);
            AddPersistentTtacRegistrationHistory(true, successMessage);

            // اگر ورود از راه دور برای همین قلم فعال بود، همین‌جا (بدون فرستادن پیام لغو جداگانه
            // به گوشی) پایان می‌یابد - چون خودِ BroadcastAlert چند خط پایین‌تر، نتیجه‌ی موفق را با
            // همان دیالوگ آشنا روی گوشی نشان می‌دهد.
            if (canRepeatFormulaRegistration)
                EndRemoteFormulaEntry(notifyPhone: false);

            CloseTtTeckRegistrationOverlay();
            _ttacCurrentPrescriptionId = null;
            _ttacCurrentCaptchaId = string.Empty;
            ApplyHistoryFilters();

            if (canRepeatFormulaRegistration && repeatContext != null)
                _lastFormulaRepeatContext = repeatContext;
            string registeredTitle = _localization.GetString("Registered");
            ShowStyledMessage(
                registeredTitle,
                successMessage,
                false,
                canRepeatFormulaRegistration,
                photoPath: formulaPhotoPath);

            // فقط برای شیرخشک: همین پیام موفقیت که روی دسکتاپ دیده شد، عیناً روی گوشی هم به شکل
            // یک هشدار با دکمه‌ی «باشه» نشان داده می‌شود - همراه با همان عکس شیرخشکی که در دیالوگ
            // «ثبت شد» دسکتاپ دیده شد (اگر عکسی برای این قلم موجود باشد).
            if (canRepeatFormulaRegistration)
                _service?.BroadcastAlert(registeredTitle, successMessage, true, formulaPhotoPath, canRepeat: true);
        }
        catch (Exception ex)
        {
            // اگر این تلاش قطعاً به تی‌تک نرسیده (خطای عادی سامانه - نه انقضای نشست و نه
            // timeout که در آن‌ها ممکن است درخواست رسیده باشد)، کلیدش را از مجموعه بردار تا
            // دیالوگ «احتمال ثبت تکراری» برای تلاش بعدیِ همین قلم اشتباه ظاهر نشود.
            if (submissionKey != null
                && !IsTtacSessionExpiredException(ex)
                && ex is not TaskCanceledException
                && ex is not TimeoutException)
            {
                _ttacSubmittedItemKeys.Remove(submissionKey);
            }

            // چون formulaMode بالای try محاسبه می‌شود ولی این catch بیرون از دامنه‌ی آن متغیر است،
            // دوباره از همان تابع پرس‌وجو می‌شود - چون _pendingRegistrationTtTeckRow تا اینجا (چه
            // موفق چه ناموفق) دست‌نخورده مانده، نتیجه با همانی که در try محاسبه شد یکی است.
            bool isFormulaRegistrationForAlert = IsCurrentRegistrationFormulaItem();
            HandleTtacOperationException(ex, _localization.GetString("SubmitItemFailed"), async () =>
            {
                await SubmitCurrentTtacItemAsync(viaRemoteEntry);
            },
            pendingLabel: _localization.GetString("PendingSubmitItem"),
            onFailureMessageShown: (shownTitle, shownMessage) =>
            {
                // فقط برای شیرخشک، و فقط وقتی واقعاً یک پیام خطای نهایی نشان داده شده (نه در حالت
                // نشست‌منقضی‌شده که خودش خودکار دوباره تلاش می‌کند و هنوز نتیجه‌ی نهایی‌ای نیست). این
                // خطا اگر ورود از راه دور فعال بود، روی گوشی هم نشان داده می‌شود - بدون پایان‌دادن به
                // کل جریان، تا کاربر بتواند با دکمه‌ی «قبلی» برگردد و فیلد را اصلاح کند.
                if (isFormulaRegistrationForAlert)
                {
                    ShowRemoteEntryErrorAndAllowRetry(shownTitle, shownMessage);
                }
            });
        }
        finally
        {
            UpdateTtacRegistrationStageButtons(false);
        }
    }

    private string GetTtTeckLookupPendingText()
    {
        return _localization.GetString("WaitingForTtTeckLookup");
    }

    private bool IsTtTeckLookupPending(string? text)
    {
        return !string.IsNullOrWhiteSpace(text) && text.TrimStart().StartsWith("⏳", StringComparison.Ordinal);
    }

    private bool IsTtTeckLookupFailed(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return true;

        return text.StartsWith("❌")
               || text.StartsWith("خطا")
               || text.Contains("No such host", StringComparison.OrdinalIgnoreCase)
               || text.Contains("timeout", StringComparison.OrdinalIgnoreCase)
               || text.Contains("در دسترس نیست", StringComparison.OrdinalIgnoreCase)
               || text.Contains("یافت نشد", StringComparison.OrdinalIgnoreCase);
    }

    private async Task LookupTtTeckForRecordAsync(ScanRecord record, bool showResultMessage)
    {
        try
        {
            var lookupService = new DrugLookupService();
            string? ttacToken = await GetTtacAccessTokenOnUiThreadAsync(false);
            DrugInfo result = await lookupService.GetDrugNameAsync(record.Barcode, ttacToken);

            if (result.Success)
            {
                record.DrugName = $"{result.PersianName} | {result.EnglishName}";
                _ttTeckDetailsByBarcode[record.Barcode] = result;
                SaveTtTeckDetailsCache();
                TryAutoOpenInfantFormulaRegistration(record);

                if (showResultMessage)
                {
                    ShowStyledMessage(
                        _localization.GetString("LookupSuccessful"),
                        result.PersianName ?? result.EnglishName ?? result.Message);
                }
            }
            else
            {
                record.DrugName = NormalizeTtTeckFailureReason(result.Message);

                if (showResultMessage)
                    ShowStyledMessage(GetLocalizedLookupFailedTitle(), record.DrugName, true);
            }
        }
        catch (Exception ex)
        {
            record.DrugName = NormalizeTtTeckFailureReason($"خطا: {ex.Message}");
            if (showResultMessage)
                ShowStyledMessage(GetLocalizedLookupFailedTitle(), record.DrugName, true);
        }
        finally
        {
            Dispatcher.Invoke(ApplyHistoryFilters);
            // نام دارو همین‌جا هم ذخیره شود، نه فقط جاهایی که بعد از این متد صراحتاً
            // SaveHistoryItemsToCsv را صدا می‌زنند - وگرنه در مسیر اسکن معمولی (خط ۳۸۴)
            // نتیجه‌ی استعلام فقط در حافظه می‌ماند و با ری‌استارت برنامه از بین می‌رود.
            SaveHistoryItemsToCsv();
        }
    }

    private string GetLocalizedLookupFailedTitle()
    {
        return _localization.GetString("LookupFailed");
    }

    private ScanRecord? FindHistoryRecord(TtTeckHistoryRow row)
    {
        return HistoryItems.FirstOrDefault(item =>
            item.Barcode == row.Barcode &&
            item.TimestampLocal == row.TimestampLocal);
    }

    private void RetryTtTeckLookupButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn || btn.Tag is not TtTeckHistoryRow row)
            return;

        _pendingRetryTtTeckRow = row;
        RetryTtTeckBarcodeText.Text = row.Barcode;
        RetryTtTeckReasonText.Text = NormalizeTtTeckFailureReason(
            string.IsNullOrWhiteSpace(row.RetryReason)
                ? GetLocalizedUnknownLookupReason()
                : row.RetryReason);

        System.Windows.Controls.Panel.SetZIndex(RetryTtTeckOverlay, 340);
        RetryTtTeckOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private string GetLocalizedUnknownLookupReason()
    {
        return _localization.GetString("ThePreviousLookupResultIsNotAvailable");
    }


    private string NormalizeTtTeckFailureReason(string? reason)
    {
        if (string.IsNullOrWhiteSpace(reason))
            return GetLocalizedUnknownLookupReason();

        if (reason.Contains("No such host", StringComparison.OrdinalIgnoreCase) ||
            (reason.Contains("newapi.ttac.ir", StringComparison.OrdinalIgnoreCase) && reason.Contains("443", StringComparison.OrdinalIgnoreCase)))
        {
            return _localization.GetString("ConnectionToTtTeckFailedWindowsCannotResolveNewapiTtacIrCheckInternetDNSVPNOrProxySettings");
        }

        if (reason.Contains("timeout", StringComparison.OrdinalIgnoreCase) || reason.Contains("timed out", StringComparison.OrdinalIgnoreCase))
        {
            return _localization.GetString("TtTeckDidNotRespondInTimePleaseRetryAFewMinutesLater");
        }

        return reason;
    }

    private void RetryTtTeckOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseRetryTtTeckOverlay();
    }

    private void RetryTtTeckCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void RetryTtTeckCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseRetryTtTeckOverlay();
    }

    private async void RetryTtTeckConfirmButton_Click(object sender, RoutedEventArgs e)
    {
        if (_pendingRetryTtTeckRow == null)
            return;

        var record = FindHistoryRecord(_pendingRetryTtTeckRow);
        if (record == null)
        {
            CloseRetryTtTeckOverlay();
            return;
        }

        record.DrugName = GetTtTeckLookupPendingText();
        ApplyHistoryFilters();
        CloseRetryTtTeckOverlay();
        await LookupTtTeckForRecordAsync(record, true);
    }

    private void CloseRetryTtTeckOverlay()
    {
        RetryTtTeckOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void HistoryTtTeckFilterButton_Click(object sender, RoutedEventArgs e)
    {
        _isTtTeckHistoryFilterActive = !_isTtTeckHistoryFilterActive;
        if (_isTtTeckHistoryFilterActive)
            _isFormulaHistoryFilterActive = false;
        ApplyHistoryFilters();
    }

    private void HistoryFormulaFilterButton_Click(object sender, RoutedEventArgs e)
    {
        _isFormulaHistoryFilterActive = !_isFormulaHistoryFilterActive;
        if (_isFormulaHistoryFilterActive)
            _isTtTeckHistoryFilterActive = false;
        ApplyHistoryFilters();
    }

    private void ApplyHistoryFilterMode()
    {
        if (HistoryListBox == null || TtTeckHistoryListBox == null)
            return;

        // تاریخچه فقط تاریخچه همه بارکدهاست؛ بخش‌های تی‌تک در پنل جداگانه نمایش داده می‌شوند.
        TtTeckHistoryListBox.Visibility = Visibility.Collapsed;
        HistoryListBox.Visibility = Visibility.Visible;
        UpdateHistoryCopyButtonsText();
    }

    private void UpdateTtTeckFilterButtonText()
    {
        if (HistoryTtTeckFilterButton == null)
            return;

        HistoryTtTeckFilterButton.Content = _isTtTeckHistoryFilterActive ? _localization.GetString("ShowAll") : _localization.GetString("TtTeckFilter");

        HistoryTtTeckFilterButton.Background = new SolidColorBrush(_isTtTeckHistoryFilterActive
            ? System.Windows.Media.Color.FromRgb(0x5B, 0x21, 0xB6)
            : System.Windows.Media.Color.FromRgb(0x7C, 0x3A, 0xED));

        if (HistoryFormulaFilterButton != null)
        {
            HistoryFormulaFilterButton.Content = _isFormulaHistoryFilterActive ? _localization.GetString("ShowAll") : _localization.GetString("FormulaRegistration");
            HistoryFormulaFilterButton.Background = new SolidColorBrush(_isFormulaHistoryFilterActive
                ? System.Windows.Media.Color.FromRgb(0xBE, 0x18, 0x5D)
                : System.Windows.Media.Color.FromRgb(0xEC, 0x48, 0x99));
        }
    }

    private void RefreshTtTeckHistoryItems()
    {
        TtTeckHistoryItems.Clear();

        foreach (var item in GetFilteredHistoryRecords())
        {
            if (!IsTtTeckHistoryRecord(item))
                continue;

            if ((_isFormulaHistoryFilterActive || _isTtacPanelFormulaOnly) && !IsInfantFormulaRecord(item))
                continue;

            var row = CreateTtTeckHistoryRowFromRecord(item);
            if (!DoesTtacPanelRowMatchFilter(row))
                continue;
            TtTeckHistoryItems.Add(row);
        }
    }

    private bool DoesTtacPanelRowMatchFilter(TtTeckHistoryRow row)
    {
        string search = _ttacPanelSearchText;
        if (string.IsNullOrWhiteSpace(search))
            return true;

        var registrations = _ttacRegistrationHistory.Where(x => x.Barcode == row.Barcode).ToList();
        string normalizedSearch = ToEnglishDigits(search);
        return ContainsText(row.Barcode, search)
               || ContainsText(row.ProductDisplayName, search)
               || ContainsText(row.PersianProductName, search)
               || ContainsText(row.EnglishProductName, search)
               || ContainsText(row.DeviceName, search)
               || ContainsText(row.PersianDateText, search)
               || ContainsText(row.TimeText, search)
               || registrations.Any(x => ContainsText(x.NationalIdMasked, search)
                                       || ContainsText(x.MobileMasked, search)
                                       || ContainsText(x.NationalIdFull, normalizedSearch)
                                       || ContainsText(x.MobileFull, normalizedSearch)
                                       || ContainsText(x.PatientFullName, search)
                                       || ContainsText(x.Message, search)
                                       || ContainsText(x.RegistrationType, search));
    }

    private bool DoesReceiveStatusRowMatchFilter(ReceiveStatusRow row)
    {
        string search = _receiveStatusSearchText;
        if (string.IsNullOrWhiteSpace(search))
            return true;

        return ContainsText(row.Barcode, search)
               || ContainsText(row.Irc, search)
               || ContainsText(row.UID, search)
               || ContainsText(row.GTIN, search)
               || ContainsText(row.ProductName, search)
               || ContainsText(row.ProductEnName, search)
               || ContainsText(row.GenericCode, search)
               || ContainsText(row.GenericName, search)
               || ContainsText(row.LotNumber, search)
               || ContainsText(row.SenderName, search)
               || ContainsText(row.Quantity, search)
               || ContainsText(row.SentDatePersian, search)
               || ContainsText(row.StatusText, search)
               || ContainsText(row.DetailText, search);
    }

    private static bool IsTtTeckHistoryRecord(ScanRecord item)
    {
        if (item.Source == BarcodeSource.TtTeck)
            return true;

        try
        {
            return BarcodeDetector.DetectBarcodeType(item.Barcode) == BarcodeSource.TtTeck
                   || IsTtTeckLookupCandidate(item.Barcode, item.Source);
        }
        catch
        {
            return IsTtTeckLookupCandidate(item.Barcode, item.Source);
        }
    }

    private void DeleteHistoryItemButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn)
            return;

        ScanRecord? record = btn.Tag switch
        {
            ScanRecord directRecord => directRecord,
            HistoryDisplayRow displayRow => displayRow.Record,
            TtTeckHistoryRow ttRow => FindHistoryRecord(ttRow),
            _ => null
        };

        if (record == null)
            return;

        RemoveHistoryRecord(record);
    }

    private void RemoveHistoryRecord(ScanRecord record)
    {
        var target = HistoryItems.FirstOrDefault(item =>
            item.Barcode == record.Barcode &&
            item.TimestampLocal == record.TimestampLocal);

        if (target == null)
            return;

        HistoryItems.Remove(target);
        SaveHistoryItemsToCsv();
        ApplyHistoryFilters();
    }

    private void SaveHistoryItemsToCsv()
    {
        SaveCurrentHistoryForCurrentPharmacy();
    }

    private static string EscapeCsvLocal(string? value)
    {
        value ??= string.Empty;
        if (value.Contains('"') || value.Contains(',') || value.Contains('\n') || value.Contains('\r'))
            return '"' + value.Replace("\"", "\"\"") + '"';

        return value;
    }

    private async void HistoryCopyButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn)
            return;

        if (btn.Tag is not string barcode)
            return;

        System.Windows.Clipboard.SetText(CleanBarcodeForExternalUse(barcode));

        var originalColor = (btn.Background as SolidColorBrush)?.Color ?? System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17);
        btn.Content = GetHistoryCopiedButtonText();

        btn.Background = new SolidColorBrush(originalColor);

        var colorAnimation = new ColorAnimation
        {
            From = originalColor,
            To = System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF),
            Duration = new Duration(System.TimeSpan.FromMilliseconds(300)),
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseInOut }
        };

        btn.Background.BeginAnimation(SolidColorBrush.ColorProperty, colorAnimation);

        await Task.Delay(2000);

        var returnColorAnimation = new ColorAnimation
        {
            From = System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF),
            To = originalColor,
            Duration = new Duration(System.TimeSpan.FromMilliseconds(300)),
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseInOut }
        };

        btn.Background.BeginAnimation(SolidColorBrush.ColorProperty, returnColorAnimation);

        btn.Content = GetHistoryCopyButtonText();
    }

    // ---------- TtTeck Details Cache ----------

    private string GetTtTeckDetailsCachePath()
    {
        return Path.Combine(AppContext.BaseDirectory, "ttteck-details-cache.json");
    }

    private void LoadTtTeckDetailsCache()
    {
        try
        {
            string path = GetTtTeckDetailsCachePath();
            if (!File.Exists(path))
                return;

            string json = File.ReadAllText(path);
            var items = JsonSerializer.Deserialize<List<DrugInfo>>(json) ?? new List<DrugInfo>();
            _ttTeckDetailsByBarcode.Clear();

            foreach (var item in items)
            {
                if (!string.IsNullOrWhiteSpace(item.OriginalBarcode))
                    _ttTeckDetailsByBarcode[item.OriginalBarcode] = item;
            }
        }
        catch { }
    }

    private void SaveTtTeckDetailsCache()
    {
        try
        {
            string path = GetTtTeckDetailsCachePath();
            var items = _ttTeckDetailsByBarcode.Values
                .Where(i => !string.IsNullOrWhiteSpace(i.OriginalBarcode))
                .GroupBy(i => i.OriginalBarcode)
                .Select(g => g.First())
                .ToList();

            string json = JsonSerializer.Serialize(items, new JsonSerializerOptions { WriteIndented = true });

            // قفل فقط دور خودِ نوشتن فایل - اگر دو اسکن هم‌زمان (از دو گوشی) هرکدام بخواهند کش را
            // ذخیره کنند، بدون این قفل ممکن است هر دو هم‌زمان File.WriteAllText را صدا بزنند و
            // فایل خراب/نصفه‌نوشته شود یا یکی با خطای اشغال‌بودن فایل شکست بخورد.
            lock (_ttTeckDetailsCacheFileLock)
            {
                File.WriteAllText(path, json);
            }
        }
        catch { }
    }

    private void FillProductDetailsExtraFields(string barcode)
    {
        ProductDetailsExtraFields.Clear();

        if (!_ttTeckDetailsByBarcode.TryGetValue(barcode, out var info))
        {
            ProductDetailsExtraTitle.Visibility = Visibility.Collapsed;
            ProductDetailsExtraFieldsList.Visibility = Visibility.Collapsed;
            return;
        }

        var addedLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in info.ExtraFields)
        {
            if (string.IsNullOrWhiteSpace(field.Value))
                continue;

            if (!addedLabels.Add(field.Key))
                continue;

            ProductDetailsExtraFields.Add(new ProductDetailField
            {
                Label = field.Key,
                Value = field.Value
            });
        }

        ProductDetailsExtraTitle.Visibility = ProductDetailsExtraFields.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
        ProductDetailsExtraFieldsList.Visibility = ProductDetailsExtraFields.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
    }

    // ---------- Product Details ----------

    private void HistoryListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (HistoryListBox.SelectedItem is HistoryDisplayRow row)
            ShowProductDetails(row.Record);
    }

    private void TtTeckHistoryListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is System.Windows.Controls.ListBox listBox && listBox.SelectedItem is TtTeckHistoryRow row)
            ShowProductDetails(row);
    }

    private void ShowProductDetails(ScanRecord record)
    {
        UpdateProductDetailsLocalizedTexts();
        var (persianName, englishName) = SplitTtTeckProductNames(record.DrugName);

        if (_ttTeckDetailsByBarcode.TryGetValue(record.Barcode, out var cachedInfo))
        {
            persianName = cachedInfo.PersianName ?? persianName;
            englishName = cachedInfo.EnglishName ?? englishName;
        }

        if (string.IsNullOrWhiteSpace(persianName) && !string.IsNullOrWhiteSpace(record.DrugName))
            persianName = record.DrugName;

        _productDetailsCurrentBarcode = record.Barcode;
        ProductDetailsPersianNameText.Text = string.IsNullOrWhiteSpace(persianName) ? "-" : persianName;
        ProductDetailsEnglishNameText.Text = string.IsNullOrWhiteSpace(englishName) ? "-" : englishName;
        ProductDetailsBarcodeText.Text = record.Barcode;
        ProductDetailsDateTimeText.Text = $"{record.PersianDateText} - {record.TimeText}";
        ProductDetailsDeviceText.Text = string.IsNullOrWhiteSpace(record.DeviceName) ? "-" : GetDeviceDisplayName(record.DeviceName);
        ProductDetailsStatusText.Text = cachedInfo?.Message ?? GetRecordStatusText(record);
        ProductDetailsCopyBarcodeButton.Tag = record.Barcode;
        FillProductDetailsExtraFields(record.Barcode);
        SetOverlayProductPhoto(ProductDetailsPhoto, ProductDetailsPhotoBorder, GetFormulaPhotoPathForBarcode(record.Barcode));

        ShowProductDetailsOverlayOnTop();
        _ = RefreshProductDetailsFromTtacCatalogAsync(record.Barcode);
    }

    private void ShowProductDetails(TtTeckHistoryRow row)
    {
        UpdateProductDetailsLocalizedTexts();
        _productDetailsCurrentBarcode = row.Barcode;
        var rowPersianName = row.PersianProductName;
        var rowEnglishName = row.EnglishProductName;
        if (_ttTeckDetailsByBarcode.TryGetValue(row.Barcode, out var cachedRowInfo))
        {
            rowPersianName = cachedRowInfo.PersianName ?? rowPersianName;
            rowEnglishName = cachedRowInfo.EnglishName ?? rowEnglishName;
        }

        ProductDetailsPersianNameText.Text = string.IsNullOrWhiteSpace(rowPersianName) ? "-" : rowPersianName;
        ProductDetailsEnglishNameText.Text = string.IsNullOrWhiteSpace(rowEnglishName) ? "-" : rowEnglishName;
        ProductDetailsBarcodeText.Text = row.Barcode;
        ProductDetailsDateTimeText.Text = $"{row.PersianDateText} - {row.TimeText}";
        ProductDetailsDeviceText.Text = string.IsNullOrWhiteSpace(row.DeviceName) ? "-" : row.DeviceName;
        ProductDetailsStatusText.Text = cachedRowInfo?.Message ?? (string.IsNullOrWhiteSpace(row.StatusText) ? GetLocalizedTtTeckFoundText() : row.StatusText);
        ProductDetailsCopyBarcodeButton.Tag = row.Barcode;
        FillProductDetailsExtraFields(row.Barcode);
        SetOverlayProductPhoto(ProductDetailsPhoto, ProductDetailsPhotoBorder, GetFormulaPhotoPathForBarcode(row.Barcode));

        ShowProductDetailsOverlayOnTop();
        _ = RefreshProductDetailsFromTtacCatalogAsync(row.Barcode);
    }

    private async Task RefreshProductDetailsFromTtacCatalogAsync(string barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return;

        try
        {
            var lookup = new DrugLookupService();
            string? ttacToken = await GetTtacAccessTokenOnUiThreadAsync(false);
            var refreshed = await lookup.GetDrugNameAsync(barcode, ttacToken);
            if (refreshed == null || !refreshed.Success)
                return;

            if (string.IsNullOrWhiteSpace(refreshed.OriginalBarcode))
                refreshed.OriginalBarcode = barcode;

            _ttTeckDetailsByBarcode[barcode] = refreshed;
            SaveTtTeckDetailsCache();

            await Dispatcher.BeginInvoke(new Action(() =>
            {
                if (ProductDetailsOverlay.Visibility != Visibility.Visible || _productDetailsCurrentBarcode != barcode)
                    return;

                ProductDetailsPersianNameText.Text = string.IsNullOrWhiteSpace(refreshed.PersianName) ? ProductDetailsPersianNameText.Text : refreshed.PersianName;
                ProductDetailsEnglishNameText.Text = string.IsNullOrWhiteSpace(refreshed.EnglishName) ? ProductDetailsEnglishNameText.Text : refreshed.EnglishName;
                if (!string.IsNullOrWhiteSpace(refreshed.GTIN))
                    ProductDetailsStatusText.Text = refreshed.Message ?? ProductDetailsStatusText.Text;
                FillProductDetailsExtraFields(barcode);
                ApplyHistoryFilters();
            }));
        }
        catch { }
    }

    private void ShowProductDetailsOverlayOnTop()
    {
        System.Windows.Controls.Panel.SetZIndex(ProductDetailsOverlay, 260);
        ProductDetailsOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private string GetRecordStatusText(ScanRecord record)
    {
        if (IsTtTeckHistoryRecord(record))
            return string.IsNullOrWhiteSpace(record.DrugName) ? GetLocalizedTtTeckFoundText() : record.DrugName;

        return _localization.GetString("RegularBarcode");
    }

    private string GetLocalizedTtTeckFoundText()
    {
        return _localization.GetString("TtTeckProduct");
    }

    private void ProductDetailsOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseProductDetails();
    }

    private void ProductDetailsCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseProductDetailsButton_Click(object sender, RoutedEventArgs e)
    {
        CloseProductDetails();
    }

    private void CloseProductDetails()
    {
        ProductDetailsOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private async void ProductDetailsCopyBarcodeButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_productDetailsCurrentBarcode))
            return;

        System.Windows.Clipboard.SetText(CleanBarcodeForExternalUse(_productDetailsCurrentBarcode));
        var originalContent = ProductDetailsCopyBarcodeButton.Content;
        ProductDetailsCopyBarcodeButton.Content = GetHistoryCopiedButtonText();
        await Task.Delay(1500);
        ProductDetailsCopyBarcodeButton.Content = originalContent;
    }

    // ---------- Export Type Selection ----------

    private void ExportExcelButton_Click(object sender, RoutedEventArgs e)
    {
        _pendingExportTarget = ExportTarget.Excel;
        ShowExportTypeSelection();
    }

    private void ExportPdfButton_Click(object sender, RoutedEventArgs e)
    {
        _pendingExportTarget = ExportTarget.Pdf;
        ShowExportTypeSelection();
    }

    private void PrintHistoryPdfReport()
    {
        var reportRows = GetCurrentHistoryReportRows().ToList();
        if (reportRows.Count == 0)
        {
            ShowStyledMessage(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        var printDialog = new System.Windows.Controls.PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var document = CreateHistoryReportDocument(reportRows);
        document.PageWidth = printDialog.PrintableAreaWidth;
        document.PageHeight = printDialog.PrintableAreaHeight;
        document.PagePadding = new Thickness(32);
        document.ColumnWidth = printDialog.PrintableAreaWidth;

        printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, "Scanbridge History Report");
        ShowSuccessDialog(GetLocalizedPdfSuccessPathText(), true);
    }

    private IEnumerable<HistoryReportRow> GetCurrentHistoryReportRows()
    {
        var rows = GetFilteredHistoryRecords();

        if (_isExportingOnlyTtTeck)
            rows = rows.Where(IsTtTeckHistoryRecord);

        foreach (var item in rows)
        {
            var (persianName, englishName) = SplitTtTeckProductNames(item.DrugName);
            yield return new HistoryReportRow
            {
                Date = item.PersianDateText,
                Time = item.TimeText,
                Barcode = item.Barcode,
                DeviceName = GetDeviceDisplayName(item.DeviceName),
                ProductName = _isExportingOnlyTtTeck
                    ? (!string.IsNullOrWhiteSpace(persianName) ? persianName : item.DrugName)
                    : item.DrugName,
                EnglishProductName = englishName
            };
        }
    }

    private FlowDocument CreateHistoryReportDocument(List<HistoryReportRow> rows)
    {
        var flowDirection = _localization.CurrentLanguage == AppLanguage.English
            ? System.Windows.FlowDirection.LeftToRight
            : System.Windows.FlowDirection.RightToLeft;

        var document = new FlowDocument
        {
            FlowDirection = flowDirection,
            FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
            FontSize = 11
        };

        var title = GetLocalizedPdfReportTitle();
        document.Blocks.Add(new Paragraph(new Run(title))
        {
            FontSize = 22,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x1A, 0x23, 0x7E)),
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 14)
        });

        string filterText = _isExportingOnlyTtTeck ? GetLocalizedTtTeckFoundText() : GetLocalizedAllScansText();
        document.Blocks.Add(new Paragraph(new Run($"{filterText} - {DateTime.Now:yyyy-MM-dd HH:mm}"))
        {
            FontSize = 12,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x4B, 0x55, 0x63)),
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 18)
        });

        var table = new Table { CellSpacing = 0 };
        document.Blocks.Add(table);

        string[] headers = _isExportingOnlyTtTeck
            ? GetLocalizedPdfTtTeckHeaders()
            : GetLocalizedPdfAllHeaders();

        for (int i = 0; i < headers.Length; i++)
            table.Columns.Add(new TableColumn());

        var rowGroup = new TableRowGroup();
        table.RowGroups.Add(rowGroup);

        var headerRow = new TableRow();
        rowGroup.Rows.Add(headerRow);
        foreach (string header in headers)
        {
            headerRow.Cells.Add(CreatePdfCell(header, true));
        }

        int counter = 1;
        foreach (var row in rows)
        {
            var tableRow = new TableRow();
            rowGroup.Rows.Add(tableRow);

            if (_isExportingOnlyTtTeck)
            {
                tableRow.Cells.Add(CreatePdfCell(counter.ToString(), false));
                tableRow.Cells.Add(CreatePdfCell(row.Date, false));
                tableRow.Cells.Add(CreatePdfCell(row.Time, false));
                tableRow.Cells.Add(CreatePdfCell(row.DeviceName, false));
                tableRow.Cells.Add(CreatePdfCell(row.Barcode, false));
                tableRow.Cells.Add(CreatePdfCell(row.ProductName, false));
                tableRow.Cells.Add(CreatePdfCell(row.EnglishProductName, false));
            }
            else
            {
                tableRow.Cells.Add(CreatePdfCell(counter.ToString(), false));
                tableRow.Cells.Add(CreatePdfCell(row.Date, false));
                tableRow.Cells.Add(CreatePdfCell(row.Time, false));
                tableRow.Cells.Add(CreatePdfCell(row.DeviceName, false));
                tableRow.Cells.Add(CreatePdfCell(row.Barcode, false));
                tableRow.Cells.Add(CreatePdfCell(row.ProductName, false));
            }

            counter++;
        }

        return document;
    }

    private TableCell CreatePdfCell(string? text, bool isHeader)
    {
        var paragraph = new Paragraph(new Run(string.IsNullOrWhiteSpace(text) ? "-" : text))
        {
            Margin = new Thickness(0),
            TextAlignment = TextAlignment.Center
        };

        var cell = new TableCell(paragraph)
        {
            Padding = new Thickness(6),
            BorderThickness = new Thickness(0.5),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xE5, 0xE7, 0xEB))
        };

        if (isHeader)
        {
            cell.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17));
            cell.Foreground = System.Windows.Media.Brushes.White;
            cell.FontWeight = FontWeights.Bold;
        }

        return cell;
    }

    private string GetLocalizedPdfReportTitle()
    {
        return _localization.GetString("ScanbridgeHistoryReport");
    }

    private string GetLocalizedAllScansText()
    {
        return _localization.GetString("AllScans");
    }

    private string[] GetLocalizedPdfAllHeaders()
    {
        return _localization.GetStringArray("ARRAY:No");
    }

    private string[] GetLocalizedPdfTtTeckHeaders()
    {
        return _localization.GetStringArray("ARRAY:No2");
    }

    private string GetLocalizedNoDataTitle()
    {
        return _localization.GetString("NoData");
    }

    private string GetLocalizedNoDataMessage()
    {
        return _localization.GetString("ThereIsNoHistoryItemForTheCurrentFilter");
    }

    private string GetLocalizedPdfSuccessPathText()
    {
        return _localization.GetString("PDFReportWasCreatedSentSuccessfully");
    }

    private void ShowExportTypeSelection()
    {
        ExportTypeTitle.Text = _pendingExportTarget == ExportTarget.Pdf
            ? _localization.GetString("SelectPDFExportType")
            : _localization.GetString("SelectExportType");

        bool ttacExportAllowed = HasLicenseModule("ttac");
        if (ExportTtTeckOptionBorder != null)
            ExportTtTeckOptionBorder.Visibility = ttacExportAllowed ? Visibility.Visible : Visibility.Collapsed;
        if (!ttacExportAllowed)
        {
            ExportAllRadio.IsChecked = true;
            ExportTtTeckRadio.IsChecked = false;
        }

        ExportTypeSelectOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void CloseExportTypeSelection()
    {
        ExportTypeSelectOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void ExportTypeSelectOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseExportTypeSelection();
    }

    private void ExportTypeSelectCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void ExportAllOption_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        ExportAllRadio.IsChecked = true;
        ExportTtTeckRadio.IsChecked = false;
    }

    private void ExportTtTeckOption_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (!HasLicenseModule("ttac"))
        {
            ExportAllRadio.IsChecked = true;
            ExportTtTeckRadio.IsChecked = false;
            e.Handled = true;
            return;
        }

        ExportAllRadio.IsChecked = false;
        ExportTtTeckRadio.IsChecked = true;
        e.Handled = true;
    }

    private void ExportTypeOkButton_Click(object sender, RoutedEventArgs e)
    {
        _isExportingOnlyTtTeck = ExportTtTeckRadio.IsChecked ?? false;
        CloseExportTypeSelection();

        if (_pendingExportTarget == ExportTarget.Pdf)
            PrintHistoryPdfReport();
        else
            ExportHistoryToExcelFiltered();
    }

    private void ExportTypeCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseExportTypeSelection();
    }

    private void ExportHistoryToExcelFiltered()
    {
        var saveFileDialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"ScanHistory_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = _localization.GetString("SaveHistoryToExcel")
        };

        if (saveFileDialog.ShowDialog() != true)
            return;

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("History");

            var headers = _isExportingOnlyTtTeck
                ? _localization.GetStringArray("ARRAY:No3")
                : _localization.GetStringArray("ARRAY:No4");

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(245, 127, 23);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int rowNumber = 2;
            int itemNumber = 1;

            // قبلاً اینجا روی کل HistoryItems حلقه می‌زد و فیلتر تاریخ/جست‌وجوی فعال روی صفحه
            // (که خروجی PDF از GetFilteredHistoryRecords رعایت می‌کند) را نادیده می‌گرفت - کاربر
            // فیلتر می‌زد، تعداد کمی رکورد روی صفحه می‌دید، ولی اکسل کل تاریخچه را می‌گرفت.
            foreach (var item in GetFilteredHistoryRecords())
            {
                // قبلاً اینجا فقط item.Source == BarcodeSource.TtTeck چک می‌شد - یک تشخیص ساده‌تر
                // و ضعیف‌تر از IsTtTeckHistoryRecord که فیلتر روی صفحه و خروجی PDF از آن استفاده
                // می‌کنند (که علاوه‌بر Source، نوع بارکد و IsTtTeckLookupCandidate را هم چک
                // می‌کند). نتیجه: رکوردهایی که روی صفحه/PDF جزو «فقط تی‌تک» حساب می‌شدند، ممکن بود
                // در همین خروجی اکسل جا بمانند (باگ گزارش ممیزی). حالا هر سه از یک معیار استفاده
                // می‌کنند.
                if (_isExportingOnlyTtTeck && !IsTtTeckHistoryRecord(item))
                    continue;

                worksheet.Cell(rowNumber, 1).Value = itemNumber;
                worksheet.Cell(rowNumber, 2).Value = item.PersianDateText;
                worksheet.Cell(rowNumber, 3).Value = item.TimeText;
                worksheet.Cell(rowNumber, 4).Value = GetDeviceDisplayName(item.DeviceName);
                worksheet.Cell(rowNumber, 5).Value = item.Barcode; // ✅ بارکد کامل

                int totalColumns;
                if (_isExportingOnlyTtTeck)
                {
                    var (persianProductName, englishProductName) = SplitTtTeckProductNames(item.DrugName);
                    worksheet.Cell(rowNumber, 6).Value = persianProductName;
                    worksheet.Cell(rowNumber, 7).Value = englishProductName;
                    totalColumns = 7;
                }
                else
                {
                    worksheet.Cell(rowNumber, 6).Value = item.DrugName;
                    totalColumns = 6;
                }

                for (int col = 1; col <= totalColumns; col++)
                {
                    worksheet.Cell(rowNumber, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                rowNumber++;
                itemNumber++;
            }

            worksheet.Column(1).Width = 8;
            worksheet.Column(2).Width = 15;
            worksheet.Column(3).Width = 12;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 40; // ✅ بارکد بزرگتر
            worksheet.Column(6).Width = _isExportingOnlyTtTeck ? 35 : 30;
            if (_isExportingOnlyTtTeck)
                worksheet.Column(7).Width = 35;

            worksheet.RangeUsed()?.CreateTable();

            workbook.SaveAs(saveFileDialog.FileName);

            ShowSuccessDialog(saveFileDialog.FileName);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Error: {ex.Message}");
        }
    }

    private static (string PersianName, string EnglishName) SplitTtTeckProductNames(string? drugName)
    {
        if (string.IsNullOrWhiteSpace(drugName))
            return (string.Empty, string.Empty);

        var parts = drugName.Split(new[] { " | " }, 2, StringSplitOptions.None);
        if (parts.Length == 2)
            return (parts[0].Trim(), parts[1].Trim());

        // برای پیام‌های خطا یا رکوردهای قدیمی که هنوز نام جداگانه ندارند
        return (drugName.Trim(), string.Empty);
    }

    private void ShowSuccessDialog(string fileName, bool isPdf = false)
    {
        if (isPdf)
        {
            SuccessTitle.Text = _localization.GetString("PDFCreatedSuccessfully");

            SuccessMessage.Text = _localization.GetString("ThePDFReportWasCreatedOrSentToTheSelectedPrinterSuccessfully");
        }
        else
        {
            SuccessTitle.Text = _localization.GetString("ExportSuccessTitle");

            SuccessMessage.Text = _localization.GetString("ExportSuccessMessage");
        }

        SuccessMessageGrid.Visibility = Visibility.Visible;
        HistoryListBox.Visibility = Visibility.Collapsed;
        TtTeckHistoryListBox.Visibility = Visibility.Collapsed;
        SuccessFilePath.Text = fileName;

        Task.Delay(3000).ContinueWith(_ =>
        {
            Dispatcher.Invoke(() =>
            {
                SuccessMessageGrid.Visibility = Visibility.Collapsed;
                ApplyHistoryFilterMode();
            });
        });
    }

    // ---------- Styled Message Dialog ----------

    private bool ShouldFocusTtacMobileAfterMessage(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
            return false;

        return message.Contains("شماره تماس", StringComparison.OrdinalIgnoreCase)
               || message.Contains("شماره همراه", StringComparison.OrdinalIgnoreCase)
               || message.Contains("Mobile number", StringComparison.OrdinalIgnoreCase)
               || message.Contains("phone", StringComparison.OrdinalIgnoreCase);
    }

    private bool ShouldSwitchTtacToPrescriptionAfterMessage(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
            return false;

        return message.Contains("1093", StringComparison.OrdinalIgnoreCase)
               || message.Contains("نسخه الکترونیک", StringComparison.OrdinalIgnoreCase)
               || message.Contains("نسخه‌ الکترونیک", StringComparison.OrdinalIgnoreCase)
               || message.Contains("نسخه محور", StringComparison.OrdinalIgnoreCase)
               || message.Contains("نسخه‌محور", StringComparison.OrdinalIgnoreCase)
               || message.Contains("مربوط به زیر فرآورده دارو", StringComparison.OrdinalIgnoreCase)
               || message.Contains("مربوط به زیرفرآورده دارو", StringComparison.OrdinalIgnoreCase)
               || message.Contains("electronic prescription", StringComparison.OrdinalIgnoreCase)
               || message.Contains("prescription", StringComparison.OrdinalIgnoreCase);
    }

    private bool ShouldOpenReceiveStatusAfterMessage(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
            return false;

        string normalized = ToEnglishDigits(message)
            .Replace("ي", "ی")
            .Replace("ك", "ک")
            .Replace("‌", "")
            .ToLowerInvariant();

        bool hasZero = normalized.Contains("صفر", StringComparison.OrdinalIgnoreCase)
                       || normalized.Contains(" 0 ", StringComparison.OrdinalIgnoreCase)
                       || normalized.Contains(":0", StringComparison.OrdinalIgnoreCase)
                       || normalized.Contains("=0", StringComparison.OrdinalIgnoreCase);

        bool isInventoryZero = normalized.Contains("موجودی", StringComparison.OrdinalIgnoreCase)
                               && hasZero
                               && (normalized.Contains("داروخانه", StringComparison.OrdinalIgnoreCase)
                                   || normalized.Contains("فرآورده", StringComparison.OrdinalIgnoreCase)
                                   || normalized.Contains("فراورده", StringComparison.OrdinalIgnoreCase));

        return isInventoryZero
               || normalized.Contains("inventory is zero", StringComparison.OrdinalIgnoreCase)
               || normalized.Contains("zero inventory", StringComparison.OrdinalIgnoreCase)
               || normalized.Contains("pharmacy inventory", StringComparison.OrdinalIgnoreCase);
    }

    private static string ExtractFirstUrl(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var match = Regex.Match(text, @"https?://[^\s\)\]\}]+", RegexOptions.IgnoreCase);
        return match.Success ? match.Value.Trim().TrimEnd('.', '،', ',', ';') : string.Empty;
    }

    // پیام‌های موفقیت تی‌تک برای ثبت شیرخشک معمولاً عیناً از سرور می‌آیند و شامل یک جمله با دو
    // مبلغ ریالی هستند: «قیمت این فرآورده X ریال است که ... و Y ریال آن توسط متقاضی ... پرداخت شود».
    // این الگو آن دو مبلغ (قیمت اولیه فرآورده و مبلغ قابل پرداخت متقاضی) را برای رنگی کردن پیدا می‌کند.
    private static readonly Regex TtacPriceMessageRegex = new Regex(
        @"قیمت\s+این\s+فرآورده\s+(?<base>[\d,٬]+\s*ریال).*?(?<patient>[\d,٬]+\s*ریال)\s+آن\s+توسط\s+متقاضی",
        RegexOptions.Singleline);

    private static bool TtacAmountTextsEqual(string a, string b)
    {
        string na = new string(a.Where(char.IsDigit).ToArray());
        string nb = new string(b.Where(char.IsDigit).ToArray());
        return na.Length > 0 && na == nb;
    }

    // پیام را در StyledMessageText نشان می‌دهد؛ اگر پیام با الگوی بالا مطابقت داشته باشد، قیمت اولیه‌ی
    // فرآورده را قرمز پررنگ و مبلغ قابل پرداخت متقاضی را سبز پررنگ می‌کند (یا اگر بیمه چیزی پرداخت
    // نکرده و دو مبلغ برابرند، هر دو را قرمز پررنگ نشان می‌دهد). برای هر پیام دیگری که با این الگو
    // مطابقت نداشته باشد، بدون تغییر به‌صورت متن ساده نمایش داده می‌شود.
    private void SetStyledMessageBody(string message)
    {
        StyledMessageText.Inlines.Clear();
        AppendTtacColorizedRuns(StyledMessageText.Inlines, message ?? string.Empty);
    }

    private void AppendTtacColorizedRuns(InlineCollection inlines, string message)
    {
        var match = TtacPriceMessageRegex.Match(message);
        if (!match.Success)
        {
            inlines.Add(new Run(message));
            return;
        }

        var baseGroup = match.Groups["base"];
        var patientGroup = match.Groups["patient"];
        bool insurancePaidNothing = TtacAmountTextsEqual(baseGroup.Value, patientGroup.Value);

        var redBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xE5, 0x39, 0x35));
        var greenBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81));

        inlines.Add(new Run(message.Substring(0, baseGroup.Index)));
        inlines.Add(new Run(baseGroup.Value) { Foreground = redBrush, FontWeight = FontWeights.Bold });
        inlines.Add(new Run(message.Substring(baseGroup.Index + baseGroup.Length, patientGroup.Index - (baseGroup.Index + baseGroup.Length))));
        inlines.Add(new Run(patientGroup.Value) { Foreground = insurancePaidNothing ? redBrush : greenBrush, FontWeight = FontWeights.Bold });
        inlines.Add(new Run(message.Substring(patientGroup.Index + patientGroup.Length)));
    }

    private void ShowStyledMessage(string title, string message, bool isError = false, bool showFormulaRepeat = false, string? linkUrl = null, string? photoPath = null, string? customIcon = null, string? linkButtonText = null)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(new Action(() => ShowStyledMessage(title, message, isError, showFormulaRepeat, linkUrl, photoPath, customIcon, linkButtonText)));
            return;
        }

        if (string.IsNullOrWhiteSpace(photoPath))
        {
            StyledMessageProductPhoto.Source = null;
            StyledMessageProductPhotoBorder.Visibility = Visibility.Collapsed;
        }
        else
        {
            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(photoPath, UriKind.Absolute);
                bitmap.EndInit();
                bitmap.Freeze();
                StyledMessageProductPhoto.Source = bitmap;
                StyledMessageProductPhotoBorder.Visibility = Visibility.Visible;
            }
            catch
            {
                StyledMessageProductPhoto.Source = null;
                StyledMessageProductPhotoBorder.Visibility = Visibility.Collapsed;
            }
        }

        StyledMessageTitle.Text = title;
        SetStyledMessageBody(message);
        StyledMessageIcon.Text = !string.IsNullOrWhiteSpace(customIcon) ? customIcon! : (isError ? "!" : "✓");
        StyledMessageIcon.Foreground = new SolidColorBrush(isError
            ? System.Windows.Media.Color.FromRgb(0xE5, 0x39, 0x35)
            : System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81));
        StyledMessageIconCircle.Background = new SolidColorBrush(isError
            ? System.Windows.Media.Color.FromRgb(0xFE, 0xF2, 0xF2)
            : System.Windows.Media.Color.FromRgb(0xEC, 0xFD, 0xF5));
        StyledMessageOkButton.Background = new SolidColorBrush(isError
            ? System.Windows.Media.Color.FromRgb(0xE5, 0x39, 0x35)
            : System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81));
        StyledMessageOkButton.Content = _localization.GetString("OK");
        StyledMessageRepeatFormulaButton.Content = _localization.GetString("RegisterAnother");
        StyledMessageRepeatFormulaButton.Visibility = showFormulaRepeat ? Visibility.Visible : Visibility.Collapsed;
        // به‌طور پیش‌فرض لیست داروخانه‌ها برای تعویض را مخفی کن؛ فقط وقتی پیام «این فرآورده در
        // این داروخانه نیست» نشان داده شود، `PopulateStyledMessagePharmacyList` آن را آشکار می‌کند.
        StyledMessagePharmacyListBorder.Visibility = Visibility.Collapsed;
        _styledMessageLinkUrl = !string.IsNullOrWhiteSpace(linkUrl) ? linkUrl! : ExtractFirstUrl(message);
        StyledMessageLinkButton.Content = !string.IsNullOrWhiteSpace(linkButtonText) ? linkButtonText! : _localization.GetString("OpenLink3");
        StyledMessageLinkButton.Visibility = string.IsNullOrWhiteSpace(_styledMessageLinkUrl) ? Visibility.Collapsed : Visibility.Visible;

        _focusTtacMobileAfterStyledMessageClose = isError
            && TtTeckRegistrationOverlay.Visibility == Visibility.Visible
            && ShouldFocusTtacMobileAfterMessage(message);
        _switchTtacToPrescriptionAfterStyledMessageClose = isError
            && TtTeckRegistrationOverlay.Visibility == Visibility.Visible
            && ShouldSwitchTtacToPrescriptionAfterMessage(message);
        _openReceiveStatusAfterStyledMessageClose = isError
            && TtTeckRegistrationOverlay.Visibility == Visibility.Visible
            && ShouldOpenReceiveStatusAfterMessage(message);
        _receiveStatusBarcodeAfterStyledMessageClose = _openReceiveStatusAfterStyledMessageClose && _pendingRegistrationTtTeckRow != null
            ? _pendingRegistrationTtTeckRow.Barcode
            : string.Empty;
        _ttac5173FlowAfterStyledMessageClose = _openReceiveStatusAfterStyledMessageClose
            && isError
            && TtTeckRegistrationOverlay.Visibility == Visibility.Visible
            && _pendingRegistrationTtTeckRow != null
            && ShouldOpenReceiveStatusAfterMessage(message);
        _ttac5173BarcodeAfterStyledMessageClose = _ttac5173FlowAfterStyledMessageClose && _pendingRegistrationTtTeckRow != null
            ? _pendingRegistrationTtTeckRow.Barcode
            : string.Empty;
        _ttac5173ReturnContext = _ttac5173FlowAfterStyledMessageClose ? CreateCurrentFormulaRepeatContext() : null;
        if (_ttac5173FlowAfterStyledMessageClose)
        {
            StyledMessageText.Inlines.Add(new LineBreak());
            StyledMessageText.Inlines.Add(new LineBreak());
            StyledMessageText.Inlines.Add(new Run(_localization.GetString("ThisProductMayNotBeReceivedConfirmedForThisPharmacyPressOKToCheckReceiveStatus")));
        }

        System.Windows.Controls.Panel.SetZIndex(StyledMessageOverlay, 350);
        StyledMessageOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        Dispatcher.BeginInvoke(new Action(() => StyledMessageOkButton.Focus()), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
    }

    private void CloseStyledMessage()
    {
        bool focusMobile = _focusTtacMobileAfterStyledMessageClose;
        bool switchToPrescription = _switchTtacToPrescriptionAfterStyledMessageClose;
        bool openReceiveStatus = _openReceiveStatusAfterStyledMessageClose;
        string receiveStatusBarcode = _receiveStatusBarcodeAfterStyledMessageClose;
        bool run5173Flow = _ttac5173FlowAfterStyledMessageClose;
        string barcode5173 = _ttac5173BarcodeAfterStyledMessageClose;
        var context5173 = _ttac5173ReturnContext;
        bool returnToRegistration = _returnToRegistrationAfterStyledMessageClose;
        string returnBarcode = _returnRegistrationBarcodeAfterStyledMessageClose;
        var returnContext = _returnRegistrationContextAfterStyledMessageClose;
        _focusTtacMobileAfterStyledMessageClose = false;
        _switchTtacToPrescriptionAfterStyledMessageClose = false;
        _openReceiveStatusAfterStyledMessageClose = false;
        _receiveStatusBarcodeAfterStyledMessageClose = string.Empty;
        _ttac5173FlowAfterStyledMessageClose = false;
        _ttac5173BarcodeAfterStyledMessageClose = string.Empty;
        _ttac5173ReturnContext = null;
        _returnToRegistrationAfterStyledMessageClose = false;
        _returnRegistrationBarcodeAfterStyledMessageClose = string.Empty;
        _returnRegistrationContextAfterStyledMessageClose = null;
        _styledMessageLogoutAction = false;
        _styledMessageLinkUrl = string.Empty;
        StyledMessageLinkButton.Visibility = Visibility.Collapsed;
        StyledMessageRetryButton.Visibility = Visibility.Collapsed;

        StyledMessageOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = TtTeckRegistrationOverlay.Visibility == Visibility.Visible
            || TtacPanelOverlay.Visibility == Visibility.Visible
            || HistoryOverlay.Visibility == Visibility.Visible
            || TtTeckWebViewOverlay.Visibility == Visibility.Visible
            || TtacLoginOverlay.Visibility == Visibility.Visible
                ? new System.Windows.Media.Effects.BlurEffect { Radius = 18 }
                : null;

        if (returnToRegistration && !string.IsNullOrWhiteSpace(returnBarcode) && returnContext != null)
        {
            Dispatcher.BeginInvoke(new Action(async () =>
            {
                await ReturnToTtacRegistrationFormAsync(returnBarcode, returnContext);
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            return;
        }

        if (run5173Flow && !string.IsNullOrWhiteSpace(barcode5173) && context5173 != null)
        {
            Dispatcher.BeginInvoke(new Action(async () =>
            {
                TtTeckRegistrationOverlay.Visibility = Visibility.Collapsed;
                await Run5173ReceiveStatusFlowAsync(barcode5173, context5173);
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            return;
        }

        if (openReceiveStatus && !string.IsNullOrWhiteSpace(receiveStatusBarcode))
        {
            Dispatcher.BeginInvoke(new Action(async () =>
            {
                TtTeckRegistrationOverlay.Visibility = Visibility.Collapsed;
                OpenReceiveStatusPanelNow();
                await AddReceiveStatusBarcodeAsync(receiveStatusBarcode, showErrors: true);
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            return;
        }

        if (switchToPrescription && TtTeckRegistrationOverlay.Visibility == Visibility.Visible)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                bool changed = TtTeckRegistrationTypeComboBox.SelectedIndex != 1;
                TtTeckRegistrationTypeComboBox.SelectedIndex = 1;
                UpdateTtTeckRegistrationTypeButtons();
                if (changed)
                    ResetTtacRegistrationAfterTypeChange(focusMedicalCouncilAfterLoad: true);
                else
                    FocusAndSelect(TtTeckRegistrationMedicalCouncilTextBox);
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            return;
        }

        if (focusMobile && TtTeckRegistrationOverlay.Visibility == Visibility.Visible)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                FocusAndSelect(TtTeckRegistrationMobileTextBox);
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }
    }

    private async Task<ReceiveStatusRow?> AddReceiveStatusBarcodeAndReturnRowAsync(string barcode, bool showErrors)
    {
        int beforeCount = ReceiveStatusItems.Count;
        await AddReceiveStatusBarcodeAsync(barcode, showErrors);
        return ReceiveStatusItems.FirstOrDefault(x => x.Barcode.Equals(barcode, StringComparison.OrdinalIgnoreCase))
               ?? (ReceiveStatusItems.Count > beforeCount ? ReceiveStatusItems.FirstOrDefault() : null);
    }

    private async Task<bool> ConfirmReceiveStatusRowAutomaticallyAsync(ReceiveStatusRow row)
    {
        if (!row.IsConfirmable || row.ReceiveId <= 0)
            return false;

        using var doc = await SendTtacJsonAsync(HttpMethod.Post, "https://statisticsreports.ttac.ir/declaration/ConfirmReceive", new { Id = row.ReceiveId });
        string? message = doc == null ? null : ReadJsonString(doc.RootElement, "Message");
        row.StatusText = string.IsNullOrWhiteSpace(message) || message == "null"
            ? (_localization.GetString("ConfirmedSuccessfully"))
            : message;
        row.IsConfirmable = false;
        row.StatusBrush = System.Windows.Media.Brushes.Green;
        ReceiveStatusListBox.Items.Refresh();
        SaveCurrentReceiveStatusItemsForCurrentPharmacy();
        return true;
    }

    private async Task Run5173ReceiveStatusFlowAsync(string barcode, TtacRepeatFormulaContext returnContext)
    {
        OpenReceiveStatusPanelNow();
        var row = await AddReceiveStatusBarcodeAndReturnRowAsync(barcode, showErrors: true);
        if (row == null)
        {
            ShowProductNotInCurrentPharmacyLogoutMessage();
            return;
        }

        if (row.IsConfirmable && row.ReceiveId > 0)
        {
            try
            {
                await ConfirmReceiveStatusRowAutomaticallyAsync(row);
                _returnToRegistrationAfterStyledMessageClose = true;
                _returnRegistrationBarcodeAfterStyledMessageClose = barcode;
                _returnRegistrationContextAfterStyledMessageClose = returnContext;
                ShowStyledMessage(
                    _localization.GetString("ReceiveStatusCompleted"),
                    _localization.GetString("ReceiveStatusWasCompletedPressOKToReturnToTheRegistrationForm"));
                return;
            }
            catch
            {
                // The row was already confirmed to belong to the current pharmacy above
                // (row != null). An exception here means the confirmation call itself failed
                // (network timeout, transient TTAC server error, etc.) - it does NOT mean the
                // product belongs to a different pharmacy, so we must not show the wrong-pharmacy
                // / logout message here. Show a generic retry message instead.
                bool english = _localization.CurrentLanguage == AppLanguage.English;
                ShowStyledMessage(
                    english ? "Confirmation Failed" : "تایید انجام نشد",
                    english
                        ? "The receipt confirmation could not be completed due to a connection error. Please try again."
                        : "تایید دریافت به دلیل خطای ارتباطی انجام نشد. لطفاً دوباره تلاش کنید.",
                    true);
                return;
            }
        }

        if (row.StatusText.Contains("قبلاً", StringComparison.OrdinalIgnoreCase)
            || row.StatusText.Contains("Already", StringComparison.OrdinalIgnoreCase)
            || row.StatusBrush == System.Windows.Media.Brushes.Green)
        {
            _returnToRegistrationAfterStyledMessageClose = true;
            _returnRegistrationBarcodeAfterStyledMessageClose = barcode;
            _returnRegistrationContextAfterStyledMessageClose = returnContext;
            ShowStyledMessage(
                _localization.GetString("AlreadyReceived"),
                _localization.GetString("ThisItemWasAlreadyReceivedConfirmedPressOKToReturnToTheRegistrationForm"));
            return;
        }

        ShowProductNotInCurrentPharmacyLogoutMessage();
    }

    private void ShowProductNotInCurrentPharmacyLogoutMessage()
    {
        _styledMessageLogoutAction = true;
        ShowStyledMessage(
            _localization.GetString("WrongPharmacy"),
            _localization.GetString("ThisProductIsNotAvailableForTheCurrentPharmacyPressLogoutLogInToTheCorrectPharmacyThenScanTheBarcodeAndFillTheFormAgain"),
            true);
        StyledMessageRepeatFormulaButton.Content = _localization.GetString("Logout");
        StyledMessageRepeatFormulaButton.Visibility = Visibility.Visible;
        StyledMessageOkButton.Content = _localization.GetString("CancelButton");

        // Show pharmacy switch list (excluding current pharmacy)
        PopulateStyledMessagePharmacyList();
    }

    private void PopulateStyledMessagePharmacyList()
    {
        if (StyledMessagePharmacyListPanel == null || StyledMessagePharmacyListBorder == null)
            return;

        StyledMessagePharmacyListPanel.Children.Clear();
        var logins = LoadSavedTtacLogins();
        if (logins.Count <= 1)
        {
            // No other pharmacies to show - keep logout button visible
            StyledMessagePharmacyListBorder.Visibility = Visibility.Collapsed;
            StyledMessageRepeatFormulaButton.Visibility = Visibility.Visible;
            return;
        }

        // Get current pharmacy username
        string currentUsername = _ttacRetryUsername ?? string.Empty;
        var otherLogins = logins.Where(x => !string.Equals(x.Username, currentUsername, StringComparison.OrdinalIgnoreCase)).ToList();
        if (otherLogins.Count == 0)
        {
            // No other pharmacies to show - keep logout button visible
            StyledMessagePharmacyListBorder.Visibility = Visibility.Collapsed;
            StyledMessageRepeatFormulaButton.Visibility = Visibility.Visible;
            return;
        }

        // We have other pharmacies - hide logout button and show pharmacy list
        StyledMessageRepeatFormulaButton.Visibility = Visibility.Collapsed;
        StyledMessagePharmacyListTitle.Text = _localization.GetString("ProductNotInPharmacySwitchHint");

        foreach (var login in otherLogins.OrderByDescending(x => x.LastUsedUtc))
        {
            string label = !string.IsNullOrWhiteSpace(login.PharmacyName) ? login.PharmacyName : login.Username;
            string content = _localization.GetFormattedString("LoginToPharmacy", label);

            var button = new System.Windows.Controls.Button
            {
                Content = content,
                Tag = login.Username,
                Height = 48,
                Margin = new Thickness(0, 4, 0, 4),
                FontSize = 15,
                Style = (Style)FindResource("RoundedButtonStyle"),
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0))
            };
            button.Click += StyledMessagePharmacySwitchButton_Click;
            StyledMessagePharmacyListPanel.Children.Add(button);
        }

        StyledMessagePharmacyListBorder.Visibility = Visibility.Visible;
    }

    private async void StyledMessagePharmacySwitchButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var button = sender as System.Windows.Controls.Button;
            var username = button?.Tag as string;
            if (string.IsNullOrWhiteSpace(username))
                return;

            // بستن فوری همهی پنلها و اورلیها — قبل از هر کار دیگر
            StyledMessageOverlay.Visibility = Visibility.Collapsed;
            ReceiveStatusOverlay.Visibility = Visibility.Collapsed;
            CargoDeliveryOverlay.Visibility = Visibility.Collapsed;
            TtTeckRegistrationOverlay.Visibility = Visibility.Collapsed;
            TtacPanelOverlay.Visibility = Visibility.Collapsed;
            ExpiryWatchOverlay.Visibility = Visibility.Collapsed;
            TtTeckWebViewOverlay.Visibility = Visibility.Collapsed;
            MainContent.Effect = null;

            // ذخیرهی وضعیت داروخانه‌ی فعلی قبل از تغییر
            try
            {
                SaveCurrentReceiveStatusItemsForCurrentPharmacy();
                SaveCurrentCargoDeliveryItemsForCurrentPharmacy();
                SaveCurrentHistoryForCurrentPharmacy();
                SaveTtacRegistrationHistory();
            }
            catch { }

            // پاک‌سازی توکن و وضعیت فعلی
            _ttacAccessTokenOverride = null;
            _ttacPharmacyDisplayName = string.Empty;
            _ttacAccessTokenExpiresAtUtc = DateTime.MinValue;
            UpdateTtacTokenValidityTracking(false, suppressExpiryNotification: true);
            _pendingTtacRetryAction = null;
            _pendingTtacRetryLabel = null;
            _ttacQuickLoginInProgress = false;
            ReceiveStatusItems.Clear();
            _receiveStatusKnownBarcodes.Clear();
            _queuedReceiveStatusBarcodes.Clear();
            _receiveStatusLoadedPharmacyKey = string.Empty;
            _cargoDeliveryLoadedPharmacyKey = string.Empty;
            CargoDeliveryItems.Clear();
            _cargoDeliveryKnownBarcodes.Clear();
            _ttacRegistrationHistoryLoadedPharmacyKey = "default";
            _ttacRegistrationHistory.Clear();
            _historyLoadedPharmacyKey = "default";
            HistoryItems.Clear();
            ApplyHistoryFilters();

            // پاک‌سازی کوکیها و localStorage مرورگر داخلی تا نشست قبلی بسته شود
            try
            {
                if (_ttTeckWebView?.CoreWebView2 != null)
                {
                    await _ttTeckWebView.CoreWebView2.ExecuteScriptAsync("try{localStorage.clear();sessionStorage.clear();}catch(e){}");
                    _ttTeckWebView.CoreWebView2.CookieManager.DeleteAllCookies();
                    _ttTeckWebView.CoreWebView2.Navigate("about:blank");
                    await Task.Delay(500);
                }
            }
            catch { }

            // باز کردن مرورگر داخلی با حساب داروخانه‌ی جدید
            _pendingTtacAutofillUsername = username;
            _ttacRetryUsername = username;
            await OpenTtTeckInternalBrowserAsync("https://newstatisticsreports.ttac.ir/pharmacyDashboard");
            _ = MonitorTtacConnectionAfterBrowserOpenAsync();
        }
        catch (Exception ex)
        {
            // قبلاً این catch کاملاً خالی بود: تا اینجا همه‌ی داده‌ی داروخانه‌ی قبلی پاک شده
            // (توکن، تاریخچه، وضعیت دریافت و ...) - اگر هر بخش پیش‌بینی‌نشده‌ای از این جریان
            // خطا بدهد، کاربر با صفحه‌ای خالی می‌ماند و هیچ ردی برای پشتیبانی باقی نمی‌ماند
            // (باگ ۱۶ گزارش ممیزی). باز کردن خودِ مرورگر (OpenTtTeckInternalBrowserAsync) از قبل
            // خطای خودش را به کاربر نشان می‌دهد؛ این فقط برای هر خطای دیگر در همین جریان است -
            // حداقل در گزارش تشخیصی ثبت شود تا قابل پیگیری باشد.
            LogBackgroundHandlerError(ex, "StyledMessagePharmacySwitchButton_Click");
        }
    }

    private async Task ReturnToTtacRegistrationFormAsync(string barcode, TtacRepeatFormulaContext context)
    {
        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () => await ReturnToTtacRegistrationFormAsync(barcode, context);
            _pendingTtacRetryLabel = _localization.GetString("PendingReturnToRegistrationForm");
            ShowTtacLoginOverlay();
            return;
        }

        var record = HistoryItems.FirstOrDefault(x => x.Barcode.Equals(barcode, StringComparison.OrdinalIgnoreCase));
        if (record == null)
        {
            record = new ScanRecord(DateTime.Now, barcode, _localization.GetString("ReturnedFromReceiveStatus"))
            {
                Source = BarcodeDetector.DetectBarcodeType(barcode),
                DrugName = GetTtTeckLookupPendingText()
            };
            AddHistoryRecord(record);
            await LookupTtTeckForRecordAsync(record, false);
        }

        var row = CreateTtTeckHistoryRowFromRecord(record);
        bool forceElectronic = context.IsElectronic;
        bool forceNone = !forceElectronic;
        OpenTtTeckRegistrationForRow(row, forceNone, forceElectronic);
        await Task.Delay(650);
        if (TtTeckRegistrationOverlay.Visibility == Visibility.Visible)
        {
            ApplyFormulaRepeatContextToOpenForm(context);
            await LoadTtacCaptchaAsync(true);
            FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
        }
    }

    private void StyledMessageLinkButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(_styledMessageLinkUrl))
                Process.Start(new ProcessStartInfo(_styledMessageLinkUrl) { UseShellExecute = true });
        }
        catch { }
    }

    private async void StyledMessageRepeatFormulaButton_Click(object sender, RoutedEventArgs e)
    {
        if (_styledMessageLogoutAction)
        {
            _styledMessageLogoutAction = false;
            StyledMessageOverlay.Visibility = Visibility.Collapsed;
            await DisconnectTtacSessionAsync(showMessage: false);
            ShowTtacLoginOverlay();
            return;
        }

        StyledMessageOverlay.Visibility = Visibility.Collapsed;
        OpenRepeatFormulaBarcodeDialog();
    }

    private void OpenRepeatFormulaBarcodeDialog()
    {
        RepeatFormulaBarcodeTitle.Text = _localization.GetString("RegisterAnotherFormula");
        RepeatFormulaBarcodeDescription.Text = _localization.GetString("EnterOrScanTheNewFormulaBarcode");
        RepeatFormulaBarcodeSubmitButton.Content = _localization.GetString("Continue");
        RepeatFormulaBarcodeCancelButton.Content = _localization.GetString("CancelButton");
        RepeatFormulaBarcodeTextBox.Text = string.Empty;
        RepeatFormulaBarcodeOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        Dispatcher.BeginInvoke(new Action(() => FocusAndSelect(RepeatFormulaBarcodeTextBox)), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
    }

    private void CloseRepeatFormulaBarcodeDialog()
    {
        RepeatFormulaBarcodeOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void RepeatFormulaBarcodeOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseRepeatFormulaBarcodeDialog();
    }

    private void RepeatFormulaBarcodeCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void RepeatFormulaBarcodeCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseRepeatFormulaBarcodeDialog();
    }

    private void RepeatFormulaBarcodeSubmitButton_Click(object sender, RoutedEventArgs e)
    {
        _ = SubmitRepeatFormulaBarcodeAsync();
    }

    private void RepeatFormulaBarcodeTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != System.Windows.Input.Key.Enter)
            return;
        e.Handled = true;
        _ = SubmitRepeatFormulaBarcodeAsync();
    }

    private async Task SubmitRepeatFormulaBarcodeAsync()
    {
        string barcode = CleanBarcodeForExternalUse(RepeatFormulaBarcodeTextBox.Text);
        if (string.IsNullOrWhiteSpace(barcode))
            return;
        await OpenRepeatFormulaRegistrationForBarcodeAsync(barcode);
    }

    private async Task OpenRepeatFormulaRegistrationForBarcodeAsync(string barcode)
    {
        if (_lastFormulaRepeatContext == null)
        {
            ShowStyledMessage(_localization.GetString("RegisterAnother"), _localization.GetString("PreviousFormulaRegistrationInformationIsNotAvailable"), true);
            return;
        }

        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () => await OpenRepeatFormulaRegistrationForBarcodeAsync(barcode);
            _pendingTtacRetryLabel = _localization.GetString("PendingOpenRepeatFormulaRegistration");
            ShowTtacLoginOverlay();
            return;
        }

        RepeatFormulaBarcodeSubmitButton.IsEnabled = false;
        try
        {
            var record = new ScanRecord(DateTime.Now, barcode, _localization.GetString("RepeatFormula"))
            {
                Source = BarcodeDetector.DetectBarcodeType(barcode),
                DrugName = GetTtTeckLookupPendingText()
            };
            AddHistoryRecord(record);
            await LookupTtTeckForRecordAsync(record, false);
            SaveHistoryItemsToCsv();

            var row = CreateTtTeckHistoryRowFromRecord(record);
            FormulaRegistrationMode mode = GetFormulaRegistrationModeForRow(row);
            bool forceElectronic = mode == FormulaRegistrationMode.PrescriptionBased || _lastFormulaRepeatContext.IsElectronic;
            bool forceNone = !forceElectronic;

            CloseRepeatFormulaBarcodeDialog();
            OpenTtTeckRegistrationForRow(row, forceNone, forceElectronic);
            await Task.Delay(650);
            if (TtTeckRegistrationOverlay.Visibility == Visibility.Visible && _lastFormulaRepeatContext != null)
            {
                ApplyFormulaRepeatContextToOpenForm(_lastFormulaRepeatContext);
                await LoadTtacCaptchaAsync(true);
                FocusAndSelect(TtTeckRegistrationCaptchaTextBox);
            }
        }
        catch (Exception ex)
        {
            ShowStyledMessage(GetLocalizedLookupFailedTitle(), ex.Message, true);
        }
        finally
        {
            RepeatFormulaBarcodeSubmitButton.IsEnabled = true;
        }
    }

    private void StyledMessageOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseStyledMessage();
    }

    private void StyledMessageCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void StyledMessageOkButton_Click(object sender, RoutedEventArgs e)
    {
        CloseStyledMessage();
    }

    // ---------- Messages ----------

    private void LoadMessages()
    {
        try
        {
            string path = Path.Combine(AppContext.BaseDirectory, "notifications.json");
            Messages.Clear();

            if (File.Exists(path))
            {
                string json = File.ReadAllText(path);
                var items = JsonSerializer.Deserialize<List<AppMessage>>(json) ?? new List<AppMessage>();
                foreach (var item in items)
                {
                    // پیام‌های قدیمی‌تر (یا پیام‌هایی که از بیرون در notifications.json نوشته شده‌اند)
                    // متن دکمه ندارند - قبلاً این متن از یک property مشترک سطح Window خوانده می‌شد؛
                    // حالا هر پیام متن دکمه‌ی خودش را دارد (تا پیام بروزرسانی بتواند «دانلود و نصب»
                    // نشان دهد نه متن عمومی «مشاهده») - نگاه کنید به AppMessage.LinkButtonText.
                    if (string.IsNullOrEmpty(item.LinkButtonText))
                        item.LinkButtonText = _localization.GetString("OpenLink2");
                    Messages.Add(item);
                }
            }

            NoMessagesText.Visibility = Messages.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            UpdateMessagesBadgeCount();
        }
        catch
        {
            NoMessagesText.Visibility = Visibility.Visible;
        }
    }

    // شمارنده‌ی پیام‌های خوانده‌نشده را روی نشان‌گر قرمز دکمه‌ی «پیام‌ها» به‌روز می‌کند - هر جا لیست
    // پیام‌ها یا وضعیت خوانده‌شدن یکی از آن‌ها عوض شود باید این متد صدا زده شود (LoadMessages،
    // CheckForAppUpdateAsync، MessageAcknowledgeButton_Click، RemoveMessage، MessageLink_Click).
    private void UpdateMessagesBadgeCount()
    {
        int unread = Messages.Count(m => !m.IsRead);
        if (unread <= 0)
        {
            MessagesBadge.Visibility = Visibility.Collapsed;
        }
        else
        {
            MessagesBadgeText.Text = unread > 9 ? "9+" : unread.ToString();
            MessagesBadge.Visibility = Visibility.Visible;
        }
    }

    // کاربر دکمه‌ی «باشه» یکی از پیام‌ها را زده - جایگزین چک‌باکس قبلی «خوانده شد» طبق درخواست صریح
    // کاربر: پیام بلافاصله خوانده‌شده علامت می‌خورد (نشان‌گر قرمز فوراً به‌روز می‌شود)، بعد کارتِ همان
    // پیام با یک محوشدنِ نرم (fade-out ~350ms) کم‌رنگ می‌شود و در پایان انیمیشن کاملاً از لیست حذف
    // می‌شود - نگاه کنید به RemoveMessage که منطق حذف واقعی (مشترک با دکمه‌ی «حذف») را انجام می‌دهد.
    private void MessageAcknowledgeButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn || btn.DataContext is not AppMessage message)
            return;

        message.IsRead = true;
        UpdateMessagesBadgeCount();
        btn.IsEnabled = false;

        if (FindVisualAncestor<Border>(btn) is not Border card)
        {
            // اگر به هر دلیلی کارتِ پیام در visual tree پیدا نشد، بدون انیمیشن مستقیم حذفش کن - تا
            // پیام هیچ‌وقت روی صفحه گیر نکند.
            RemoveMessage(message);
            return;
        }

        var fadeOut = new DoubleAnimation
        {
            From = 1.0,
            To = 0.0,
            Duration = TimeSpan.FromMilliseconds(350),
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
        };
        fadeOut.Completed += (_, _) => RemoveMessage(message);
        card.BeginAnimation(UIElement.OpacityProperty, fadeOut);
    }

    // کاربر دکمه‌ی «حذف» یکی از پیام‌ها را زده - همان RemoveMessage مشترک را بدون انیمیشن (فوری) صدا
    // می‌زند.
    private void MessageDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn || btn.DataContext is not AppMessage message)
            return;

        RemoveMessage(message);
    }

    // حذف واقعی یک پیام - کاملاً از لیست (و در نتیجه از notifications.json، چون SaveMessagesToDisk
    // همیشه کل Messages فعلی را می‌نویسد) پاک می‌شود. چون Messages یک ObservableCollection است و
    // ItemsSource همان است (نگاه کنید به سازنده‌ی پنجره)، حذف از این مجموعه خودش MessagesList را
    // بدون نیاز به Items.Refresh() به‌روز می‌کند. هم دکمه‌ی «باشه» (بعد از fade) و هم دکمه‌ی «حذف»
    // (فوری) همین متد را صدا می‌زنند تا هر دو مسیر دقیقاً یک‌جور رفتار کنند.
    private void RemoveMessage(AppMessage message)
    {
        // اگر فایل نصب این پیام قبلاً دانلود شده بود، همراه با خودِ پیام از دیسک هم پاکش می‌کنیم - وگرنه
        // یک فایل نصبِ یتیم روی دیسک کاربر باقی می‌ماند که دیگر هیچ‌جا به آن اشاره نمی‌شود.
        if (!string.IsNullOrWhiteSpace(message.DownloadedInstallerPath))
        {
            try { if (File.Exists(message.DownloadedInstallerPath)) File.Delete(message.DownloadedInstallerPath); } catch { }
        }

        Messages.Remove(message);
        SaveMessagesToDisk();
        UpdateMessagesBadgeCount();
        NoMessagesText.Visibility = Messages.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    // در visual tree از یک عنصر به بالا می‌رود تا اولین نیای هم‌نوع T را پیدا کند - برای پیدا کردن
    // Border ریشه‌ی کارتِ پیام از روی دکمه‌ی داخلش (نگاه کنید به MessageAcknowledgeButton_Click).
    private static T? FindVisualAncestor<T>(DependencyObject start) where T : DependencyObject
    {
        DependencyObject? current = VisualTreeHelper.GetParent(start);
        while (current != null)
        {
            if (current is T typed)
                return typed;
            current = VisualTreeHelper.GetParent(current);
        }
        return null;
    }

    // بعد از این‌که CheckForAppUpdateAsync پیام بروزرسانی را به Messages اضافه/حذف می‌کند، همان
    // لیست کامل روی دیسک (notifications.json) هم بازنویسی می‌شود - تا اگر برنامه قبل از چک بعدی
    // بسته و باز شود، همان پیام (با همان دکمه‌ی «دانلود و نصب») دوباره در لیست باشد، نه این‌که تا
    // چک روزانه‌ی بعدی گم شود.
    private void SaveMessagesToDisk()
    {
        try
        {
            string path = Path.Combine(AppContext.BaseDirectory, "notifications.json");
            File.WriteAllText(path, JsonSerializer.Serialize(Messages.ToList(), new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private static string GetCurrentAppVersionString()
    {
        try
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            string? info = asm.GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrWhiteSpace(info))
            {
                int plus = info.IndexOf('+');
                if (plus >= 0)
                    info = info[..plus];
                return info.Trim();
            }

            string? product = System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location).ProductVersion;
            if (!string.IsNullOrWhiteSpace(product))
                return product.Split('+')[0].Trim();

            return asm.GetName().Version?.ToString() ?? "1.0.0";
        }
        catch
        {
            return "1.0.0";
        }
    }

    // بررسی بروزرسانی: یک فایل JSON ساده روی سایت با فرمت {"version": "...", "message": "...",
    // "url": "..."} می‌خواند (دقیقاً همان فرمتی که اپ اندروید هم برای همین کار استفاده می‌کند - هیچ
    // سرور اختصاصی/کد سمت سرور جدید لازم نیست، فقط کافی است بعد از هر انتشار نسخه‌ی جدید همین یک
    // فایل روی سایت، در مسیر AppUpdateCheckUrl، بازنویسی شود). اگر نسخه‌ی داخل فایل از نسخه‌ی
    // نصب‌شده‌ی همین سیستم جدیدتر باشد، یک پیام (با دکمه‌ی «دانلود و نصب») به «پیام‌ها» اضافه و
    // نشان‌گر قرمز روی دکمه‌ی پیام‌ها فعال می‌شود. اگر همان نسخه قبلاً اضافه شده (کاربر هنوز
    // آپدیت نکرده)، دوباره پیام تکراری اضافه نمی‌شود - فقط با انتشار نسخه‌ی جدیدتر دیگری جایگزین
    // می‌شود؛ همین‌طور اگر کاربر آپدیت کرد (نسخه‌ی این سیستم دیگر قدیمی‌تر از سایت نیست)، پیام
    // بروزرسانی قبلی از لیست پاک می‌شود.
    private async Task CheckForAppUpdateAsync(bool manualTrigger)
    {
        try
        {
            using var response = await _updateCheckHttpClient.GetAsync(AppUpdateCheckUrl);
            if (!response.IsSuccessStatusCode)
            {
                if (manualTrigger)
                    ShowUpdateCheckFailedMessage();
                return;
            }

            string content = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(content);
            string remoteVersionText = doc.RootElement.TryGetProperty("version", out var vProp) ? vProp.GetString() ?? string.Empty : string.Empty;
            string updateMessageBody = doc.RootElement.TryGetProperty("message", out var mProp) ? mProp.GetString() ?? string.Empty : string.Empty;
            string downloadUrl = doc.RootElement.TryGetProperty("url", out var uProp) ? uProp.GetString() ?? string.Empty : string.Empty;

            // فقط «version» برای این‌که بفهمیم درخواست واقعاً موفق بوده و قابل‌مقایسه است لازم است؛
            // «message»/«url» خالی یعنی سرور جواب داده ولی فعلاً بروزرسانی‌ای تعریف نشده (دقیقاً حالت
            // پیش‌فرض فایل روی سرور) - این را نباید با شکست واقعی اتصال/سرور یکی گرفت، وگرنه بررسی
            // دستی همیشه پیغام «ناموفق بود» می‌دهد حتی وقتی همه‌چیز درست کار می‌کند (این باگ همان
            // چیزی بود که کاربر گزارش داد).
            if (string.IsNullOrWhiteSpace(remoteVersionText)
                || !Version.TryParse(remoteVersionText, out var remoteVersion)
                || !Version.TryParse(GetCurrentAppVersionString(), out var currentVersion))
            {
                if (manualTrigger)
                    ShowUpdateCheckFailedMessage();
                return;
            }

            if (remoteVersion <= currentVersion)
            {
                var stale = Messages.Where(m => m.IsUpdateDownload).ToList();
                if (stale.Count > 0)
                {
                    foreach (var s in stale)
                    {
                        Messages.Remove(s);
                        // اگر فایل نصب این نسخه‌ی قدیمی روی دیسک مانده (دانلود شده ولی هیچ‌وقت نصب
                        // نشده)، پاکش می‌کنیم تا فضای دیسک کاربر بی‌دلیل اشغال نماند.
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(s.DownloadedInstallerPath) && File.Exists(s.DownloadedInstallerPath))
                                File.Delete(s.DownloadedInstallerPath);
                        }
                        catch { }
                    }
                    NoMessagesText.Visibility = Messages.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
                    SaveMessagesToDisk();
                    UpdateMessagesBadgeCount();
                }

                if (manualTrigger)
                {
                    ShowStyledMessage(
                        _localization.CurrentLanguage == AppLanguage.English ? "Up to date" : "بروزرسانی",
                        _localization.CurrentLanguage == AppLanguage.English ? "You already have the latest version." : "شما جدیدترین نسخه‌ی برنامه را دارید.",
                        false);
                }
                return;
            }

            // اینجا یعنی نسخه‌ی روی سایت واقعاً از نسخه‌ی نصب‌شده جدیدتر است - اگر لینک دانلود خالی
            // مانده (مثلاً نسخه در فایل بالا برده شده ولی لینک هنوز ست نشده)، به‌جای پیام گمراه‌کننده‌ی
            // «اتصال قطع است»، مشکل واقعی (تنظیمات ناقص روی سایت) گفته می‌شود.
            if (string.IsNullOrWhiteSpace(downloadUrl))
            {
                if (manualTrigger)
                    ShowStyledMessage(
                        _localization.CurrentLanguage == AppLanguage.English ? "Update not ready" : "بروزرسانی هنوز آماده نیست",
                        _localization.CurrentLanguage == AppLanguage.English
                            ? $"Version {remoteVersionText} is listed on the server but no download link is set yet."
                            : $"نسخه‌ی {remoteVersionText} روی سایت تعریف شده ولی لینک دانلودش هنوز خالی است.",
                        true);
                return;
            }

            bool alreadyKnown = Messages.Any(m => m.IsUpdateDownload && string.Equals(m.UpdateVersion, remoteVersionText, StringComparison.OrdinalIgnoreCase));
            if (!alreadyKnown)
            {
                var oldUpdates = Messages.Where(m => m.IsUpdateDownload).ToList();
                foreach (var old in oldUpdates)
                {
                    Messages.Remove(old);
                    // فایل دانلودشده‌ی نسخه‌ی قبلی‌تر (اگر بوده) دیگر لازم نیست - الان نسخه‌ی جدیدتری
                    // جایگزینش می‌شود.
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(old.DownloadedInstallerPath) && File.Exists(old.DownloadedInstallerPath))
                            File.Delete(old.DownloadedInstallerPath);
                    }
                    catch { }
                }

                var updateMessage = new AppMessage
                {
                    Title = (_localization.CurrentLanguage == AppLanguage.English ? "New version available: " : "نسخه‌ی جدید در دسترس است: ") + remoteVersionText,
                    Body = updateMessageBody,
                    Link = downloadUrl,
                    LinkButtonText = _localization.CurrentLanguage == AppLanguage.English ? "Download & install" : "دانلود و نصب",
                    IsUpdateDownload = true,
                    UpdateVersion = remoteVersionText
                };
                Messages.Insert(0, updateMessage);
                NoMessagesText.Visibility = Visibility.Collapsed;
                SaveMessagesToDisk();
                UpdateMessagesBadgeCount();
            }
            else
            {
                // پیام همان نسخه از قبل توی لیست هست - چیزی عوض نشده، فقط شمارنده را دوباره حساب
                // می‌کنیم (اگر کاربر قبلاً «خوانده شد» زده بود، نباید اینجا دوباره نشان‌گر روشن شود).
                UpdateMessagesBadgeCount();
            }

            if (manualTrigger)
                MessagesButton_Click(this, new RoutedEventArgs());
        }
        catch
        {
            // بدون اینترنت یا سایت موقتاً در دسترس نبود - بی‌سروصدا نادیده گرفته می‌شود؛ چک بعدی
            // (روزانه یا دستی) دوباره امتحان می‌کند. داروخانه‌ای که اینترنتش قطع است نباید با پیام
            // خطا مزاحمش شویم مگر خودش دستی زده باشد «بررسی بروزرسانی».
            if (manualTrigger)
                ShowUpdateCheckFailedMessage();
        }
    }

    private void ShowUpdateCheckFailedMessage()
    {
        ShowStyledMessage(
            _localization.CurrentLanguage == AppLanguage.English ? "Update check failed" : "بررسی بروزرسانی ناموفق بود",
            _localization.CurrentLanguage == AppLanguage.English ? "No internet connection, or the update server is unavailable." : "اتصال اینترنت برقرار نیست یا سرور بروزرسانی در دسترس نیست.",
            true);
    }

    private void CheckForUpdateButton_Click(object sender, RoutedEventArgs e)
    {
        // این دکمه داخل TtTeckSettingsOverlay است. اگر همین‌جا بدون بستن تنظیمات، پیام بروزرسانی
        // را باز کنیم (نگاه کنید به انتهای CheckForAppUpdateAsync)، هر دو overlay هم‌زمان Visible
        // می‌شوند و بسته به ترتیب آن‌ها در XAML، پیام‌ها ممکن است پشت پنجره‌ی تنظیمات پنهان بماند.
        // پس اول تنظیمات را می‌بندیم تا پیام همیشه در جلو دیده شود.
        CloseTtTeckSettings();
        _ = CheckForAppUpdateAsync(manualTrigger: true);
    }

    // کاربر روی دکمه‌ی پیام بروزرسانی کلیک کرده. مسیر (بر اساس تست کاربر):
    //   ۰) اگر همین نسخه قبلاً یک‌بار کامل دانلود شده و فایلش هنوز روی دیسک است (message.DownloadedInstallerPath)،
    //      اصلاً دوباره دانلود نمی‌کنیم - مستقیم می‌رویم سراغ تایید نصب.
    //   ۱) وگرنه، تایید قبل از شروع دانلود.
    //   ۲) دانلود با نمایش پیشرفت در باکسِ درون‌برنامه‌ایِ DownloadProgressPanel (داخل ستون «پنل
    //      کاربری»، درست زیر «تاریخ نزدیک») که دکمه‌ی «لغو» هم دارد (با CancellationTokenSource دانلود
    //      واقعاً متوقف و فایل ناقص حذف می‌شود).
    //   ۳) بعد از اتمام دانلود، یک تاییدِ دوم («نصب و راه‌اندازی مجدد») - نه نصب خودکار فوری - تا
    //      کاربر خودش زمان بستن برنامه را انتخاب کند؛ همان‌جا مسیر فایل روی پیام ذخیره و متن دکمه از
    //      «دانلود و نصب» به «نصب» عوض می‌شود تا دفعه‌ی بعد دیگر دانلود نشود.
    // نصب با فلگ‌های بی‌صدای Inno Setup انجام می‌شود (کاربر تأیید کرد اینستالرشان با Inno Setup ساخته
    // می‌شود) تا نیازی به کلیک روی صفحات ویزارد Setup نباشد - نگاه کنید به یادداشت‌های داخل
    // RunSilentInstallAndRestart برای جزئیات /CLOSEAPPLICATIONS و /RESTARTAPPLICATIONS.
    private async Task StartAppUpdateDownloadAsync(AppMessage message)
    {
        string? downloadUrl = message.Link;
        string? version = message.UpdateVersion;
        if (string.IsNullOrWhiteSpace(downloadUrl))
            return;

        // قبلاً همین نسخه دانلود شده - دیگر دوباره دانلود نکن، فقط تایید نصب را نشان بده.
        if (!string.IsNullOrWhiteSpace(message.DownloadedInstallerPath) && File.Exists(message.DownloadedInstallerPath))
        {
            CloseMessagesOverlay();
            await ConfirmAndRunInstallAsync(message.DownloadedInstallerPath, version);
            return;
        }

        var cts = new CancellationTokenSource();
        _downloadCts = cts;
        string? installerPathForCleanup = null;
        try
        {
            string title = _localization.CurrentLanguage == AppLanguage.English ? "Download update" : "دانلود بروزرسانی";
            string confirmMessage = _localization.CurrentLanguage == AppLanguage.English
                ? $"Version {version} will be downloaded. When the download finishes you will be asked to install and restart. Continue?"
                : $"نسخه‌ی {version} دانلود می‌شود. بعد از اتمام دانلود، گزینه‌ی «نصب و راه‌اندازی مجدد» به شما نشان داده خواهد شد. ادامه می‌دهید؟";

            bool confirmed = await ShowUpdateConfirmOverlayAsync(
                title,
                confirmMessage,
                confirmText: _localization.CurrentLanguage == AppLanguage.English ? "Download" : "دانلود",
                cancelText: _localization.CurrentLanguage == AppLanguage.English ? "Later" : "بعداً");
            if (!confirmed)
                return;

            CloseMessagesOverlay();

            string updateFolder = Path.Combine(Path.GetTempPath(), "ScanBridgeUpdate");
            Directory.CreateDirectory(updateFolder);
            string fileName = "ScanBridgeSetup" + (string.IsNullOrWhiteSpace(version) ? "" : $"-{version}") + ".exe";
            string installerPath = Path.Combine(updateFolder, fileName);
            installerPathForCleanup = installerPath;

            ShowDownloadProgressPanel(version);

            using (var response = await _updateCheckHttpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead, cts.Token))
            {
                response.EnsureSuccessStatusCode();
                long? totalBytes = response.Content.Headers.ContentLength;

                await using var httpStream = await response.Content.ReadAsStreamAsync(cts.Token);
                await using var fileStream = File.Create(installerPath);

                var buffer = new byte[81920];
                long readSoFar = 0;
                int bytesRead;
                while ((bytesRead = await httpStream.ReadAsync(buffer, cts.Token)) > 0)
                {
                    await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead), cts.Token);
                    readSoFar += bytesRead;

                    double readMb = readSoFar / 1024.0 / 1024.0;
                    if (totalBytes.HasValue && totalBytes.Value > 0)
                    {
                        double percent = Math.Min(100.0, readSoFar * 100.0 / totalBytes.Value);
                        double totalMb = totalBytes.Value / 1024.0 / 1024.0;
                        UpdateDownloadProgressPanel(percent, $"{readMb:0.0} / {totalMb:0.0} MB");
                    }
                    else
                    {
                        UpdateDownloadProgressPanel(null, $"{readMb:0.0} MB");
                    }
                }
            }

            HideDownloadProgressPanel();

            // دانلود کامل شد - مسیرش را روی خود پیام ذخیره می‌کنیم تا دفعه‌ی بعد که کاربر روی دکمه‌ی
            // همین پیام کلیک کرد، دیگر دوباره دانلود نشود؛ متن دکمه هم از «دانلود و نصب» به «نصب»
            // عوض می‌شود چون دیگر واقعاً فقط نصب مانده.
            message.DownloadedInstallerPath = installerPath;
            message.LinkButtonText = _localization.CurrentLanguage == AppLanguage.English ? "Install" : "نصب";
            SaveMessagesToDisk();
            MessagesList.Items.Refresh();

            await ConfirmAndRunInstallAsync(installerPath, version);
        }
        catch (OperationCanceledException)
        {
            try { HideDownloadProgressPanel(); } catch { }
            try { if (installerPathForCleanup != null && File.Exists(installerPathForCleanup)) File.Delete(installerPathForCleanup); } catch { }
            ShowStyledMessage(
                _localization.CurrentLanguage == AppLanguage.English ? "Download cancelled" : "دانلود لغو شد",
                _localization.CurrentLanguage == AppLanguage.English ? "You can start it again anytime from the same message." : "هر وقت خواستید، از همین پیام دوباره می‌توانید شروعش کنید.",
                false);
        }
        catch (Exception ex)
        {
            try { HideDownloadProgressPanel(); } catch { }
            try { CloseStyledMessage(); } catch { }
            ShowStyledMessage(
                _localization.CurrentLanguage == AppLanguage.English ? "Update failed" : "بروزرسانی ناموفق بود",
                (_localization.CurrentLanguage == AppLanguage.English ? "Could not download or start the update: " : "دانلود یا اجرای بروزرسانی ممکن نشد: ") + ex.Message,
                true);
        }
        finally
        {
            cts.Dispose();
            _downloadCts = null;
        }
    }

    // تایید «نصب و راه‌اندازی مجدد» را نشان می‌دهد - چه بلافاصله بعد از دانلود، چه (طبق درخواست
    // کاربر) دفعه‌ی بعدی که فایل از قبل روی دیسک آماده است و نیازی به دانلود دوباره نیست.
    private async Task ConfirmAndRunInstallAsync(string installerPath, string? version)
    {
        string doneTitle = _localization.CurrentLanguage == AppLanguage.English ? "Ready to install" : "آماده‌ی نصب";
        string doneMessage = _localization.CurrentLanguage == AppLanguage.English
            ? $"Version {version} has been downloaded. The installer window will open so you can click Next and finish setup."
            : $"نسخه‌ی {version} دانلود شد. با زدن نصب، پنجره‌ی Setup باز می‌شود تا خودتان Next بزنید و نصب را تمام کنید.";

        bool installConfirmed = await ShowUpdateConfirmOverlayAsync(
            doneTitle,
            doneMessage,
            confirmText: _localization.CurrentLanguage == AppLanguage.English ? "Install" : "نصب",
            cancelText: _localization.CurrentLanguage == AppLanguage.English ? "Later" : "بعداً");

        if (!installConfirmed)
        {
            ShowStyledMessage(
                _localization.CurrentLanguage == AppLanguage.English ? "Saved for later" : "برای بعد ذخیره شد",
                _localization.CurrentLanguage == AppLanguage.English ? "You can install it later from the same message in \"Messages\"." : "هر وقت خواستید، از همین پیام توی «پیام‌ها» می‌توانید نصبش کنید.",
                false);
            return;
        }

        RunSilentInstallAndRestart(installerPath);
    }

    // باکس پیشرفت دانلود دیگر یک پنجره‌ی جدا با موقعیت محاسبه‌شده نیست (آن روش با چند بار تست کاربر
    // ثابت شد قابل‌اعتماد نیست) - DownloadProgressPanel یک Border معمولی در MainWindow.xaml است که
    // به‌عنوان فرزندِ همان ستونِ دکمه‌های «پنل کاربری»، بلافاصله بعد از «تاریخ نزدیک» تعریف شده؛ پس
    // چیدمانِ طبیعیِ WPF خودش جایش را همیشه دقیقاً زیر آن دکمه نگه می‌دارد.
    private CancellationTokenSource? _downloadCts;

    private void ShowDownloadProgressPanel(string? version)
    {
        DownloadProgressTitleText.Text = "⬇ " + (_localization.CurrentLanguage == AppLanguage.English ? "Downloading update" : "دانلود بروزرسانی")
            + (string.IsNullOrWhiteSpace(version) ? "" : $" ({(_localization.CurrentLanguage == AppLanguage.English ? "version " : "نسخه‌ی ")}{version})");
        DownloadProgressBar.IsIndeterminate = true;
        DownloadProgressBar.Value = 0;
        DownloadProgressSizeText.Text = "";
        DownloadProgressPercentText.Text = _localization.CurrentLanguage == AppLanguage.English ? "Starting..." : "در حال شروع...";
        DownloadProgressCancelButton.IsEnabled = true;
        DownloadProgressPanel.Visibility = Visibility.Visible;
    }

    // percent=null یعنی اندازه‌ی فایل از سرور مشخص نبوده (هدر Content-Length نداشت) - نوار پیشرفت
    // به‌صورت نامعین (در حال حرکت، بدون درصد دقیق) نمایش داده می‌شود.
    private void UpdateDownloadProgressPanel(double? percent, string sizeText)
    {
        DownloadProgressSizeText.Text = sizeText;
        if (percent.HasValue)
        {
            DownloadProgressBar.IsIndeterminate = false;
            DownloadProgressBar.Value = percent.Value;
            DownloadProgressPercentText.Text = $"{percent.Value:0}%";
        }
        else
        {
            DownloadProgressBar.IsIndeterminate = true;
            DownloadProgressPercentText.Text = sizeText;
        }
    }

    private void HideDownloadProgressPanel()
    {
        DownloadProgressPanel.Visibility = Visibility.Collapsed;
    }

    private void DownloadProgressCancelButton_Click(object sender, RoutedEventArgs e)
    {
        if (_downloadCts == null)
            return;
        DownloadProgressCancelButton.IsEnabled = false;
        DownloadProgressPercentText.Text = _localization.CurrentLanguage == AppLanguage.English ? "Cancelling..." : "در حال لغو...";
        _downloadCts.Cancel();
    }

    private TaskCompletionSource<bool>? _updateConfirmTcs;

    // جایگزین HighUsageConfirmWindow برای تاییدهای مربوط به بروزرسانی دسکتاپ - طبق درخواست صریح
    // کاربر («این پیغام یه پنجره‌ی جدید نباشه، بیارش توی برنامه») به‌جای یک Window جدا، یک Overlay
    // داخل خودِ MainWindow است (UpdateConfirmOverlay در XAML، هم‌استایل با بقیه‌ی Overlayهای برنامه،
    // با بلور معمول پشت‌زمینه). چون این تابع وسط متدهای async صدا زده می‌شود و باید منتظر انتخاب
    // کاربر بماند، از TaskCompletionSource استفاده شده - دکمه‌های بله/بعداً و کلیک روی پس‌زمینه
    // (مثل «بعداً» رفتار می‌کند) همان Task را کامل می‌کنند.
    private Task<bool> ShowUpdateConfirmOverlayAsync(string title, string message, string confirmText, string cancelText)
    {
        // اگر یک تایید قبلی هنوز باز بود (نباید عملاً پیش بیاید)، به‌عنوان لغوشده بسته‌اش می‌کنیم تا
        // Task قبلی بی‌جواب نماند.
        _updateConfirmTcs?.TrySetResult(false);

        UpdateConfirmTitleText.Text = "⚠ " + title;
        UpdateConfirmMessageText.Text = message;
        UpdateConfirmYesButton.Content = confirmText;
        UpdateConfirmNoButton.Content = cancelText;

        UpdateConfirmOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };

        _updateConfirmTcs = new TaskCompletionSource<bool>();
        return _updateConfirmTcs.Task;
    }

    private void CloseUpdateConfirmOverlay(bool result)
    {
        UpdateConfirmOverlay.Visibility = Visibility.Collapsed;
        // اگر هنوز پنجره‌ی پیام‌ها یا تنظیمات باز است، بلورشان را نگه می‌داریم - فقط اگر هیچ Overlay
        // دیگری باز نیست، بلور پاک می‌شود.
        if (MessagesOverlay.Visibility != Visibility.Visible && TtTeckSettingsOverlay.Visibility != Visibility.Visible)
            MainContent.Effect = null;

        _updateConfirmTcs?.TrySetResult(result);
        _updateConfirmTcs = null;
    }

    private void UpdateConfirmYesButton_Click(object sender, RoutedEventArgs e) => CloseUpdateConfirmOverlay(true);

    private void UpdateConfirmNoButton_Click(object sender, RoutedEventArgs e) => CloseUpdateConfirmOverlay(false);

    private void UpdateConfirmOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e) => CloseUpdateConfirmOverlay(false);

    private void UpdateConfirmCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e) => e.Handled = true;

    private static void TryUnblockDownloadedFile(string path)
    {
        try { File.Delete(path + ":Zone.Identifier"); } catch { }
    }

    // بعد از تایید کاربر، فایل Setup دانلودشده را مثل یک نصب معمولی باز می‌کند
    // (پنجره‌ی Next اینستالر). برنامه را خودش نمی‌بندد.
    private void RunSilentInstallAndRestart(string installerPath)
    {
        try
        {
            if (!File.Exists(installerPath))
            {
                ShowStyledMessage(
                    _localization.CurrentLanguage == AppLanguage.English ? "Installer not found" : "فایل نصب پیدا نشد",
                    installerPath,
                    true);
                return;
            }
            TryUnblockDownloadedFile(installerPath);
            Process.Start(new ProcessStartInfo(installerPath)
            {
                UseShellExecute = true,
                Verb = "runas"
            });
        }
        catch (Exception ex)
        {
            ShowStyledMessage(
                _localization.CurrentLanguage == AppLanguage.English ? "Could not open installer" : "باز کردن فایل نصب ممکن نشد",
                ex.Message,
                true);
        }
    }

    private void MessagesButton_Click(object sender, RoutedEventArgs e)
    {
        MessagesOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        // طبق درخواست کاربر، دیگر صرفاً باز کردن این پنجره نشان‌گر قرمز را پاک نمی‌کند - فقط زدن دکمه‌ی
        // «باشه» هر پیام (نگاه کنید به MessageAcknowledgeButton_Click) این کار را می‌کند.
    }

    private void CloseMessagesButton_Click(object sender, RoutedEventArgs e)
    {
        CloseMessagesOverlay();
    }

    private void MessagesOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseMessagesOverlay();
    }

    private void MessagesCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseMessagesOverlay()
    {
        MessagesOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void MessageLink_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn || btn.Tag is not string link || string.IsNullOrWhiteSpace(link))
            return;

        // کلیک روی دکمه‌ی پیام یعنی کاربر با آن تعامل داشته - همین‌جا هم خوانده‌شده علامت می‌خورد،
        // مستقل از این‌که خودش چک‌باکس «خوانده شد» را هم بزند یا نه.
        if (btn.DataContext is AppMessage clickedMessage && !clickedMessage.IsRead)
        {
            clickedMessage.IsRead = true;
            SaveMessagesToDisk();
            UpdateMessagesBadgeCount();
            MessagesList.Items.Refresh();
        }

        // پیام‌های بروزرسانی (نگاه کنید به CheckForAppUpdateAsync) به‌جای بازکردن لینک در مرورگر،
        // خودشان فایل نصب را دانلود و اجرا می‌کنند و برنامه را می‌بندند - نه هر پیامی با Link.
        if (btn.DataContext is AppMessage message && message.IsUpdateDownload)
        {
            _ = StartAppUpdateDownloadAsync(message);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo(link) { UseShellExecute = true });
        }
        catch { }
    }

    // ---------- Device Aliases ----------

    private string GetDeviceAliasesPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "device-aliases.json");
    }

    private void LoadDeviceAliases()
    {
        try
        {
            string path = GetDeviceAliasesPath();
            if (!File.Exists(path))
                return;

            string json = File.ReadAllText(path);
            var aliases = JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
            _deviceAliases.Clear();
            foreach (var pair in aliases)
            {
                if (!string.IsNullOrWhiteSpace(pair.Key) && !string.IsNullOrWhiteSpace(pair.Value))
                    _deviceAliases[pair.Key] = pair.Value;
            }
        }
        catch { }
    }

    private void SaveDeviceAliases()
    {
        try
        {
            string json = JsonSerializer.Serialize(_deviceAliases, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(GetDeviceAliasesPath(), json);
        }
        catch { }
    }

    private string GetDeviceDisplayName(string? originalDeviceName)
    {
        if (string.IsNullOrWhiteSpace(originalDeviceName))
            return "";

        return _deviceAliases.TryGetValue(originalDeviceName, out var alias) && !string.IsNullOrWhiteSpace(alias)
            ? alias
            : originalDeviceName;
    }

    private void DeviceRow_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement element || element.Tag is not string originalDeviceName)
            return;

        OpenDeviceAliasEditor(originalDeviceName);
        e.Handled = true;
    }

    private void DisconnectDeviceButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.Tag is not string originalDeviceName)
            return;

        e.Handled = true;

        bool disconnected = TryDisconnectDevice(originalDeviceName);
        if (disconnected)
        {
            ShowStyledMessage(GetLocalizedDeviceDisconnectedTitle(), GetLocalizedDeviceDisconnectedMessage(originalDeviceName));
            return;
        }

        ShowStyledMessage(GetLocalizedDeviceDisconnectUnavailableTitle(), GetLocalizedDeviceDisconnectUnavailableMessage(), true);
    }

    private bool TryDisconnectDevice(string deviceName)
    {
        try
        {
            var serviceType = _service.GetType();
            string[] methodNames =
            {
                "DisconnectDevice",
                "DisconnectClient",
                "RemoveDevice",
                "KickDevice",
                "CloseDeviceConnection",
                "DisconnectByDeviceName"
            };

            foreach (string methodName in methodNames)
            {
                var method = serviceType.GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (method == null)
                    continue;

                var parameters = method.GetParameters();
                object? result;
                if (parameters.Length == 1 && parameters[0].ParameterType == typeof(string))
                {
                    result = method.Invoke(_service, new object[] { deviceName });
                }
                else if (parameters.Length == 0)
                {
                    result = method.Invoke(_service, null);
                }
                else
                {
                    continue;
                }

                if (result is bool boolResult)
                    return boolResult;

                return true;
            }
        }
        catch { }

        return false;
    }

    private string GetLocalizedDeviceDisconnectedTitle()
    {
        return _localization.GetString("DeviceDisconnected");
    }

    private string GetLocalizedDeviceDisconnectedMessage(string deviceName)
    {
        string displayName = GetDeviceDisplayName(deviceName);
        return _localization.GetFormattedString("DeviceDisconnectedMessage", displayName);
    }

    private string GetLocalizedDeviceDisconnectUnavailableTitle()
    {
        return _localization.GetString("DisconnectUnavailable");
    }

    private string GetLocalizedDeviceDisconnectUnavailableMessage()
    {
        return _localization.GetString("TheCurrentConnectionServiceDoesNotExposeADirectDisconnectMethodForThisDevice");
    }

    private void OpenDeviceAliasEditor(string originalDeviceName)
    {
        _editingDeviceOriginalName = originalDeviceName;
        DeviceAliasOriginalText.Text = originalDeviceName;
        DeviceAliasTextBox.Text = GetDeviceDisplayName(originalDeviceName);
        DeviceAliasOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        DeviceAliasTextBox.Focus();
        DeviceAliasTextBox.SelectAll();
    }

    private void DeviceAliasOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseDeviceAliasEditor();
    }

    private void DeviceAliasCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void DeviceAliasCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseDeviceAliasEditor();
    }

    private void DeviceAliasSaveButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_editingDeviceOriginalName))
            return;

        string oldDisplayName = GetDeviceDisplayName(_editingDeviceOriginalName);
        string newDisplayName = DeviceAliasTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(newDisplayName) || newDisplayName.Equals(_editingDeviceOriginalName, StringComparison.OrdinalIgnoreCase))
            _deviceAliases.Remove(_editingDeviceOriginalName);
        else
            _deviceAliases[_editingDeviceOriginalName] = newDisplayName;

        SaveDeviceAliases();
        RefreshDeviceRowsFromLastState();
        ApplyDeviceAliasToHistoryItems(_editingDeviceOriginalName, oldDisplayName, GetDeviceDisplayName(_editingDeviceOriginalName));
        CloseDeviceAliasEditor();
    }

    private void DeviceAliasResetButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_editingDeviceOriginalName))
            return;

        string oldDisplayName = GetDeviceDisplayName(_editingDeviceOriginalName);
        _deviceAliases.Remove(_editingDeviceOriginalName);
        SaveDeviceAliases();
        RefreshDeviceRowsFromLastState();
        ApplyDeviceAliasToHistoryItems(_editingDeviceOriginalName, oldDisplayName, _editingDeviceOriginalName);
        CloseDeviceAliasEditor();
    }

    private void CloseDeviceAliasEditor()
    {
        DeviceAliasOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void RefreshDeviceRowsFromLastState()
    {
        UpdateDeviceRows(_lastConnectedDevices);
    }

    private void ApplyDeviceAliasToHistoryItems(string originalName, string oldDisplayName, string newDisplayName)
    {
        if (HistoryItems.Count == 0)
            return;

        var updated = new List<ScanRecord>();
        foreach (var item in HistoryItems)
        {
            string deviceName = item.DeviceName;
            if (deviceName.Equals(originalName, StringComparison.OrdinalIgnoreCase) ||
                deviceName.Equals(oldDisplayName, StringComparison.OrdinalIgnoreCase))
            {
                deviceName = newDisplayName;
            }

            var newRecord = new ScanRecord(item.TimestampLocal, item.Barcode, deviceName)
            {
                Source = item.Source,
                DrugName = item.DrugName
            };
            updated.Add(newRecord);
        }

        HistoryItems.Clear();
        foreach (var item in updated.OrderByDescending(r => r.TimestampLocal))
            HistoryItems.Add(item);

        ApplyHistoryFilters();
    }

    // ---------- Connected Devices ----------

    private void UpdateDeviceRows(IReadOnlyList<ConnectedDeviceInfo> devices)
    {
        _lastConnectedDevices.Clear();
        _lastConnectedDevices.AddRange(devices);

        Dispatcher.BeginInvoke(new Action(() =>
        {
            DeviceRows.Clear();
            foreach (var device in devices)
            {
                string badge = device.LinkKind switch
                {
                    "USB" => "🔌 کابل",
                    "WiFi" => "📶 Wi-Fi",
                    _ => "🔗 LAN"
                };

                DeviceRows.Add(new DeviceRowDisplayViewModel
                {
                    OriginalDeviceName = device.DeviceName,
                    DeviceName = GetDeviceDisplayName(device.DeviceName),
                    LinkBadge = badge,
                    StatusColor = device.HasScanned
                        ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x4C, 0xAF, 0x50))
                        : new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0x98, 0x00))
                });

                if (device.LinkKind == "USB" && !_usbInternetTipShown)
                {
                    _usbInternetTipShown = true;
                    ShowStyledMessage(
                        "اتصال با کابل برقرار شد ✅",
                        "اگر اینترنت سیستم قطع شد، از تنظیمات دکمه‌ی «رفع تداخل اینترنت USB» را بزنید تا اینترنت از طریق وای‌فای/اترنت ادامه پیدا کند و اتصال کابل هم سالم بماند.",
                        false);
                }
            }

            NoDevicesText.Visibility = DeviceRows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }));
    }

    // ================= پنجره نتیجه قیمت =================

    private Services.PriceLookupService.PriceResult? _lastPriceResult;
    private List<Services.PriceLookupService.ProductSummary> _priceLookupCurrentList = new();
    private bool _priceLookupFilterUpdating;

    private void ShowPriceResultWindow(Services.PriceLookupService.PriceResult result)
    {
        _lastPriceResult = result;
        // عنوان: همیشه «💰 استعلام قیمت» نشان بده
        PriceResultTitleText.Text = "💰 استعلام قیمت";
        PriceResultFaName.Text = string.IsNullOrWhiteSpace(result.FaName) ? "—" : result.FaName;
        PriceResultEnName.Text = string.IsNullOrWhiteSpace(result.EnName) ? "—" : result.EnName;
        PriceResultGenericCode.Text = string.IsNullOrWhiteSpace(result.GenericCode) ? "—" : result.GenericCode;
        PriceResultPackageCount.Text = string.IsNullOrWhiteSpace(result.PackageCount) ? "—" : result.PackageCount;
        PriceResultBrandOwner.Text = string.IsNullOrWhiteSpace(result.BrandOwner) ? "—" : result.BrandOwner;
        PriceResultNotDrugWarning.Visibility = result.FoundButNotDrugSubgroup ? Visibility.Visible : Visibility.Collapsed;
        PriceResultPriceBox.Visibility = result.TotalPriceRial > 0 ? Visibility.Visible : Visibility.Collapsed;
        if (result.TotalPriceRial > 0)
            PriceResultPriceText.Text = result.TotalPriceRial.ToString("N0", System.Globalization.CultureInfo.InvariantCulture);
        PriceResultOverlay.Visibility = Visibility.Visible;
        BroadcastPriceLookupDetail(result);
    }

    private void ClosePriceResult()
    {
        ClosePriceCustomQtyResult();
        ClosePriceCustomQty();
        PriceResultOverlay.Visibility = Visibility.Collapsed;
    }

    private void PriceResultCloseButton_Click(object sender, RoutedEventArgs e) => ClosePriceResult();

    private void PriceResultOverlay_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (PriceCustomQtyOverlay.Visibility == Visibility.Visible || PriceCustomQtyResultOverlay.Visibility == Visibility.Visible)
            return;
        if (e.Key == System.Windows.Input.Key.Escape || e.Key == System.Windows.Input.Key.Enter)
        {
            e.Handled = true;
            ClosePriceResult();
        }
    }

    private void PriceCustomQtyButton_Click(object sender, RoutedEventArgs e)
    {
        if (_lastPriceResult == null || GetUnitConsumerPrice(_lastPriceResult) <= 0)
        {
            PriceCustomQtyErrorText.Text = "برای این فرآورده قیمت مصرف‌کننده در تی‌تک ثبت نشده است.";
            PriceCustomQtyErrorText.Visibility = Visibility.Visible;
        }
        else
        {
            PriceCustomQtyErrorText.Text = "";
            PriceCustomQtyErrorText.Visibility = Visibility.Collapsed;
        }

        PriceCustomQtyInput.Text = "";
        UpdatePriceCustomQtyPlaceholder();
        PriceCustomQtyOverlay.Visibility = Visibility.Visible;
        PriceCustomQtyInput.Focus();
    }

    private void ClosePriceCustomQty()
    {
        PriceCustomQtyOverlay.Visibility = Visibility.Collapsed;
    }

    private void PriceCustomQtyInput_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) => UpdatePriceCustomQtyPlaceholder();

    private void UpdatePriceCustomQtyPlaceholder()
    {
        if (PriceCustomQtyPlaceholder == null)
            return;
        PriceCustomQtyPlaceholder.Visibility = string.IsNullOrWhiteSpace(PriceCustomQtyInput.Text)
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private void PriceCustomQtyCloseButton_Click(object sender, RoutedEventArgs e) => ClosePriceCustomQty();

    private void PriceCustomQtyOverlay_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape)
        {
            e.Handled = true;
            ClosePriceCustomQty();
            return;
        }
        if (e.Key == System.Windows.Input.Key.Enter)
        {
            e.Handled = true;
            ConfirmPriceCustomQty();
        }
    }

    private void PriceCustomQtyConfirmButton_Click(object sender, RoutedEventArgs e) => ConfirmPriceCustomQty();

    /// <summary>قیمت مصرف‌کننده برای ۱ عدد — نه قیمت کل جعبه.</summary>
    private static decimal GetUnitConsumerPrice(Services.PriceLookupService.PriceResult result)
    {
        if (result.ConsumerPricePerUnit > 0)
            return result.ConsumerPricePerUnit;

        if (result.TotalPriceRial <= 0)
            return 0;

        string packRaw = ToEnglishDigits((result.PackageCount ?? "").Trim()).Replace(",", "");
        if (decimal.TryParse(packRaw, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out var pack) && pack > 1)
            return result.TotalPriceRial / pack;

        return 0;
    }

    private void ConfirmPriceCustomQty()
    {
        if (_lastPriceResult == null)
        {
            PriceCustomQtyErrorText.Text = "ابتدا یک فرآورده را استعلام کنید.";
            PriceCustomQtyErrorText.Visibility = Visibility.Visible;
            return;
        }

        decimal unit = GetUnitConsumerPrice(_lastPriceResult);
        if (unit <= 0)
        {
            PriceCustomQtyErrorText.Text = "قیمت یک عدد (مصرف‌کننده) برای این فرآورده موجود نیست.";
            PriceCustomQtyErrorText.Visibility = Visibility.Visible;
            return;
        }

        string raw = ToEnglishDigits((PriceCustomQtyInput.Text ?? "").Trim()).Replace(",", "");
        if (!decimal.TryParse(raw, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out var qty) || qty <= 0)
        {
            PriceCustomQtyErrorText.Text = "یک تعداد معتبر بزرگ‌تر از صفر وارد کنید.";
            PriceCustomQtyErrorText.Visibility = Visibility.Visible;
            PriceCustomQtyInput.Focus();
            PriceCustomQtyInput.SelectAll();
            return;
        }

        PriceCustomQtyErrorText.Visibility = Visibility.Collapsed;
        Services.PriceLookupService.ParseNameParts(_lastPriceResult.FaName ?? "", out _, out var form, out _);
        if (string.IsNullOrWhiteSpace(form))
            form = _lastPriceResult.ProductType;

        decimal total = qty * unit;
        PriceQtyResultFaName.Text = string.IsNullOrWhiteSpace(_lastPriceResult.FaName) ? "—" : _lastPriceResult.FaName;
        PriceQtyResultEnName.Text = string.IsNullOrWhiteSpace(_lastPriceResult.EnName) ? "—" : _lastPriceResult.EnName;
        PriceQtyResultForm.Text = string.IsNullOrWhiteSpace(form) ? "—" : form;
        PriceQtyResultCompany.Text = string.IsNullOrWhiteSpace(_lastPriceResult.BrandOwner) ? "—" : _lastPriceResult.BrandOwner;
        PriceQtyResultCount.Text = qty.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        PriceQtyResultTotal.Text = total.ToString("N0", System.Globalization.CultureInfo.InvariantCulture);
        PriceCustomQtyResultOverlay.Visibility = Visibility.Visible;
    }

    private void ClosePriceCustomQtyResult()
    {
        PriceCustomQtyResultOverlay.Visibility = Visibility.Collapsed;
    }

    private void PriceCustomQtyResultCloseButton_Click(object sender, RoutedEventArgs e) => ClosePriceCustomQtyResult();

    private void PriceCustomQtyResultOverlay_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape || e.Key == System.Windows.Input.Key.Enter)
        {
            e.Handled = true;
            ClosePriceCustomQtyResult();
        }
    }

    // ================= استعلام قیمت فرآورده از تی‌تک =================

    private Services.PriceLookupService? _priceLookupService;
    private string _lastScannedIrc = "";

    private Services.PriceLookupService PriceLookup => _priceLookupService ??= new Services.PriceLookupService();

    private string? PriceLookupToken => string.IsNullOrWhiteSpace(_ttacAccessTokenOverride) ? null : _ttacAccessTokenOverride;

    private void OpenPriceLookup(string subtitle)
    {
        PriceLookupSubtitleText.Text = subtitle;
        PriceLookupStatusText.Text = "";
        PriceLookupStatusText.Visibility = Visibility.Collapsed;
        PriceLookupResultsList.Visibility = Visibility.Collapsed;
        PriceLookupResultsList.Children.Clear();
        HidePriceLookupFilterPanel();
        PriceLookupDetailsPanel.Visibility = Visibility.Collapsed;
        PriceLookupNotDrugWarning.Visibility = Visibility.Collapsed;
        PriceLookupOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        // تا وقتی این پنجره باز است، اسکن‌ها دیگر به‌صورت تایپ کیبورد تزریق نمی‌شوند
        if (_service != null) _service.SuppressKeyboardInjection = true;
        PriceSearchInputsPanel.Visibility = Visibility.Visible;
        PriceNameInput.Focus();
        BroadcastPriceLookupOpen();
    }

    /// <summary>
    /// اسکن گوشی را در کادر فعال استعلام قیمت می‌گذارد (نام / بارکد / ژنریک) و همان را جست‌وجو می‌کند.
    /// </summary>
    private void ApplyScannedValueToPriceLookup(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return;

        PriceSearchInputsPanel.Visibility = Visibility.Visible;
        PriceLookupResultsList.Visibility = Visibility.Collapsed;
        HidePriceLookupFilterPanel();
        PriceLookupDetailsPanel.Visibility = Visibility.Collapsed;

        string mode = _priceActiveField;
        if (mode != "name" && mode != "barcode" && mode != "generic")
            mode = "barcode";

        System.Windows.Controls.TextBox target;
        if (mode == "name")
        {
            PriceBarcodeInput.Text = "";
            PriceGenericInput.Text = "";
            PriceNameInput.Text = value;
            target = PriceNameInput;
        }
        else if (mode == "generic")
        {
            PriceNameInput.Text = "";
            PriceBarcodeInput.Text = "";
            PriceGenericInput.Text = value;
            target = PriceGenericInput;
        }
        else
        {
            PriceNameInput.Text = "";
            PriceGenericInput.Text = "";
            PriceBarcodeInput.Text = value;
            target = PriceBarcodeInput;
            mode = "barcode";
        }

        _priceActiveField = mode;
        target.Focus();
        target.CaretIndex = value.Length;
        target.SelectAll();

        _ = RunActivePriceSearchAsync();
    }

    private void ClosePriceLookup()
    {
        PriceLookupOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
        if (_service != null) _service.SuppressKeyboardInjection = false;
        BroadcastPriceLookupCancel();
    }

    private void PriceLookupCloseButton_Click(object sender, RoutedEventArgs e) => ClosePriceLookup();

    /// <summary>دکمه‌ی «💰 استعلام قیمت» در پنل کاربری — فقط بعد از ورود به تی‌تک.</summary>
    private async void PriceLookupPanelButton_Click(object sender, RoutedEventArgs e)
    {
        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () =>
            {
                OpenPriceLookup("بر اساس نام، بارکد یا کد ژنریک جست‌وجو کنید");
                PriceResetButton_Click(PriceLookupSearchButton, new RoutedEventArgs());
                await Task.CompletedTask;
            };
            _pendingTtacRetryLabel = "استعلام قیمت";
            ShowTtacLoginOverlay();
            return;
        }

        OpenPriceLookup("بر اساس نام، بارکد یا کد ژنریک جست‌وجو کنید");
        PriceResetButton_Click(sender, e);
    }

    /// <summary>دکمه‌ی زرد «قیمت» کنار «ثبت در تی‌تک» در ردیف‌های تاریخچه تی‌تک.</summary>
    private async void TtTeckHistoryPriceButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn || btn.Tag is not TtTeckHistoryRow row || string.IsNullOrWhiteSpace(row.Barcode))
        {
            ShowStyledMessage("بارکدی نیست", "برای این ردیف بارکدی ثبت نشده است.", true);
            return;
        }

        string barcode = row.Barcode;
        string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
        if (string.IsNullOrWhiteSpace(token))
        {
            _pendingTtacRetryAction = async () => await LookupHistoryRowPriceAsync(barcode);
            _pendingTtacRetryLabel = "استعلام قیمت";
            ShowTtacLoginOverlay();
            return;
        }

        await LookupHistoryRowPriceAsync(barcode);
    }

    private async Task LookupHistoryRowPriceAsync(string barcode)
    {
        string? irc = _lastScannedIrc;
        if (string.IsNullOrWhiteSpace(irc))
            irc = await PriceLookup.GetIrcFromBarcodeAsync(barcode, PriceLookupToken);
        var result = !string.IsNullOrWhiteSpace(irc)
            ? await PriceLookup.LookupByIrcAsync(irc, PriceLookupToken)
            : await PriceLookup.LookupByBarcodeAsync(barcode, PriceLookupToken);
        ShowPriceResult(result);
    }

    private void PriceLookupOverlay_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape)
        {
            e.Handled = true;
            ClosePriceLookup();
        }
        else if (e.Key == System.Windows.Input.Key.Enter && PriceSearchInputsPanel.Visibility == Visibility.Visible)
        {
            e.Handled = true;
            _ = RunActivePriceSearchAsync();
        }
    }

    private string _priceActiveField = "name";

    private void PriceInput_GotFocus(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.TextBox tb && tb.Tag is string tag)
        {
            _priceActiveField = tag;
            if (tb.Parent is Border b)
                b.BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2563EB");
        }
    }

    private void PriceInputBorder_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is Border border)
        {
            var box = FindVisualChild<System.Windows.Controls.TextBox>(border);
            if (box != null)
            {
                box.Focus();
                e.Handled = true;
            }
        }
    }

    private void PriceInput_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.TextBox tb && tb.Parent is Border b)
            b.BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#CBD5E1");
    }

    private async void PriceInput_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape)
        {
            e.Handled = true;
            ClosePriceLookup();
            return;
        }
        if (e.Key == System.Windows.Input.Key.Enter)
        {
            e.Handled = true;
            await RunActivePriceSearchAsync();
        }
    }

    /// <summary>جست‌وجو بر اساس کادر فعال (نام / بارکد / کد ژنریک)</summary>
    private async Task RunActivePriceSearchAsync()
    {
        string query;
        string mode = _priceActiveField;

        if (mode == "barcode") query = PriceBarcodeInput.Text.Trim();
        else if (mode == "generic") query = PriceGenericInput.Text.Trim();
        else { query = PriceNameInput.Text.Trim(); mode = "name"; }

        // اگر فوکوس روی کادر دیگری مانده ولی فقط یکی از کادرها پر است، همان را جست‌وجو کن
        if (string.IsNullOrWhiteSpace(query))
        {
            if (!string.IsNullOrWhiteSpace(PriceGenericInput.Text))
            {
                mode = "generic";
                query = PriceGenericInput.Text.Trim();
            }
            else if (!string.IsNullOrWhiteSpace(PriceBarcodeInput.Text))
            {
                mode = "barcode";
                query = PriceBarcodeInput.Text.Trim();
            }
            else if (!string.IsNullOrWhiteSpace(PriceNameInput.Text))
            {
                mode = "name";
                query = PriceNameInput.Text.Trim();
            }
        }

        if (query.Length < 2)
        {
            BroadcastPriceLookupStatus("حداقل دو حرف وارد کنید یا بارکد را اسکن کنید.");
            if (!_priceLookupQueryFromPhone)
                ShowStyledMessage("کم است", "حداقل دو حرف وارد کنید یا بارکد را اسکن کنید.", true);
            return;
        }

        BroadcastPriceLookupStatus("🔍 در حال جست‌وجو در تی‌تک...");

        if (mode == "barcode")
        {
            await RunPriceLookupAsync(query, isFromScan: true);
            return;
        }

        if (mode == "generic")
        {
            PriceLookupStatusText.Visibility = Visibility.Visible;
            PriceLookupDetailsPanel.Visibility = Visibility.Collapsed;
            PriceLookupResultsList.Visibility = Visibility.Collapsed;
            HidePriceLookupFilterPanel();
            PriceLookupStatusText.Text = "🔍 در حال جست‌وجوی کد ژنریک در تی‌تک...";
            var products = await PriceLookup.SearchByGenericCodeAsync(query, PriceLookupToken);
            products = await PriceLookup.EnrichSummariesWithDetailsAsync(products, PriceLookupToken);
            await ShowProductSelectionListAsync(products);
            return;
        }

        await RunPriceLookupAsync(query, isFromScan: false);
    }

    /// <summary>نمایش لیست مرتب‌شده‌ی فرآورده‌ها برای انتخاب. فیلتر شکل/دوز فقط برای جست‌وجوی نام.</summary>
    private async Task ShowProductSelectionListAsync(List<Services.PriceLookupService.ProductSummary> products, bool showFormDoseFilter = false)
    {
        if (products.Count == 0)
        {
            HidePriceLookupFilterPanel();
            PriceLookupStatusText.Visibility = Visibility.Visible;
            PriceLookupStatusText.Text = "❌ فرآورده‌ای پیدا نشد. عبارت را بررسی کنید.";
            BroadcastPriceLookupList(products, "❌ فرآورده‌ای پیدا نشد. عبارت را بررسی کنید.");
            return;
        }

        if (products.Count == 1)
        {
            BroadcastPriceLookupList(products, "۱ فرآورده پیدا شد");
            var single = products[0];
            bool needsFetch = single.ProductId > 0
                && string.IsNullOrWhiteSpace(single.EnName)
                && string.IsNullOrWhiteSpace(single.GenericCode)
                && string.IsNullOrWhiteSpace(single.BrandOwner);
            if (needsFetch)
            {
                await RunProductDetailsAsync(single.ProductId, single.Title);
                return;
            }
            ShowPriceResult(PriceLookup.ToPriceResult(single));
            return;
        }

        foreach (var item in products)
        {
            Services.PriceLookupService.ParseNameParts(item.Title, item.EnName, out var brand, out var form, out var dose);
            if (!string.IsNullOrWhiteSpace(form) || !string.IsNullOrWhiteSpace(dose))
            {
                if (string.IsNullOrWhiteSpace(item.Brand) || item.Brand == item.Title)
                    item.Brand = brand;
                if (!string.IsNullOrWhiteSpace(form))
                    item.Form = form;
                if (!string.IsNullOrWhiteSpace(dose))
                    item.Dose = dose;
            }
        }

        // مرتب‌سازی طبق درخواست: شکل دارویی → اسم دارو → دوز → IRC
        var ordered = products
            .OrderBy(p => p.Title.StartsWith("فرآورده ", StringComparison.Ordinal) ? 1 : 0)
            .ThenBy(p => p.Form, StringComparer.OrdinalIgnoreCase)
            .ThenBy(p => string.IsNullOrWhiteSpace(p.Brand) ? p.Title : p.Brand, StringComparer.OrdinalIgnoreCase)
            .ThenBy(p => Services.PriceLookupService.DoseValue(p.Dose))
            .ThenBy(p => p.Subtitle, StringComparer.OrdinalIgnoreCase)
            .ToList();

        _priceLookupCurrentList = ordered;
        PriceLookupStatusText.Text = ordered.Count + " فرآورده پیدا شد — یکی را انتخاب کنید:";
        PriceLookupStatusText.Visibility = Visibility.Visible;
        BroadcastPriceLookupList(ordered, PriceLookupStatusText.Text);
        PriceLookupResultsList.Visibility = Visibility.Visible;
        PriceLookupDetailsPanel.Visibility = Visibility.Collapsed;
        if (showFormDoseFilter)
            PopulatePriceLookupFilters(ordered);
        else
            HidePriceLookupFilterPanel();
        RenderPriceLookupResultRows(ordered);
        return;
    }

    private void HidePriceLookupFilterPanel()
    {
        _priceLookupCurrentList = new List<Services.PriceLookupService.ProductSummary>();
        if (PriceLookupFilterPanel != null)
            PriceLookupFilterPanel.Visibility = Visibility.Collapsed;
    }

    private void PopulatePriceLookupFilters(List<Services.PriceLookupService.ProductSummary> items)
    {
        if (PriceLookupFilterPanel == null || PriceLookupFormFilterCombo == null || PriceLookupDoseFilterCombo == null)
            return;

        var forms = items
            .Select(p => (p.Form ?? string.Empty).Trim())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var doses = items
            .Select(p => (p.Dose ?? string.Empty).Trim())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(Services.PriceLookupService.DoseValue)
            .ThenBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToList();

        _priceLookupFilterUpdating = true;
        try
        {
            PriceLookupFormFilterCombo.Items.Clear();
            PriceLookupFormFilterCombo.Items.Add("همه شکل‌ها");
            foreach (var f in forms)
                PriceLookupFormFilterCombo.Items.Add(f);
            PriceLookupFormFilterCombo.SelectedIndex = 0;

            PriceLookupDoseFilterCombo.Items.Clear();
            PriceLookupDoseFilterCombo.Items.Add("همه دوزها");
            foreach (var d in doses)
                PriceLookupDoseFilterCombo.Items.Add(d);
            PriceLookupDoseFilterCombo.SelectedIndex = 0;
        }
        finally
        {
            _priceLookupFilterUpdating = false;
        }

        PriceLookupFilterPanel.Visibility = (forms.Count > 0 || doses.Count > 0) ? Visibility.Visible : Visibility.Collapsed;
    }

    private void PriceLookupFilter_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_priceLookupFilterUpdating)
            return;
        ApplyPriceLookupFilters();
    }

    private void PriceLookupFilterClearButton_Click(object sender, RoutedEventArgs e)
    {
        if (PriceLookupFormFilterCombo == null || PriceLookupDoseFilterCombo == null)
            return;
        _priceLookupFilterUpdating = true;
        try
        {
            if (PriceLookupFormFilterCombo.Items.Count > 0)
                PriceLookupFormFilterCombo.SelectedIndex = 0;
            if (PriceLookupDoseFilterCombo.Items.Count > 0)
                PriceLookupDoseFilterCombo.SelectedIndex = 0;
        }
        finally
        {
            _priceLookupFilterUpdating = false;
        }
        ApplyPriceLookupFilters();
    }

    private void ApplyPriceLookupFilters()
    {
        if (_priceLookupCurrentList.Count == 0)
            return;

        string formFilter = PriceLookupFormFilterCombo?.SelectedIndex > 0
            ? (PriceLookupFormFilterCombo.SelectedItem?.ToString() ?? string.Empty)
            : string.Empty;
        string doseFilter = PriceLookupDoseFilterCombo?.SelectedIndex > 0
            ? (PriceLookupDoseFilterCombo.SelectedItem?.ToString() ?? string.Empty)
            : string.Empty;

        var filtered = _priceLookupCurrentList.Where(p =>
        {
            if (!string.IsNullOrWhiteSpace(formFilter)
                && !string.Equals((p.Form ?? string.Empty).Trim(), formFilter, StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrWhiteSpace(doseFilter)
                && !string.Equals((p.Dose ?? string.Empty).Trim(), doseFilter, StringComparison.OrdinalIgnoreCase))
                return false;
            return true;
        }).ToList();

        PriceLookupStatusText.Text = filtered.Count == _priceLookupCurrentList.Count
            ? _priceLookupCurrentList.Count + " فرآورده پیدا شد — یکی را انتخاب کنید:"
            : filtered.Count + " از " + _priceLookupCurrentList.Count + " فرآورده (فیلتر شده):";
        PriceLookupStatusText.Visibility = Visibility.Visible;
        RenderPriceLookupResultRows(filtered);
    }

    private void RenderPriceLookupResultRows(List<Services.PriceLookupService.ProductSummary> rows)
    {
        PriceLookupResultsList.Children.Clear();
        PriceLookupResultsList.Visibility = Visibility.Visible;

        foreach (var p in rows)
        {
            string brandCell = string.IsNullOrWhiteSpace(p.Brand) ? p.Title : p.Brand;
            string formCell = string.IsNullOrWhiteSpace(p.Form) ? "—" : p.Form;
            string doseCell = string.IsNullOrWhiteSpace(p.Dose) ? "—" : p.Dose;
            string ircCell = p.Subtitle.StartsWith("IRC: ") ? p.Subtitle.Substring(5) : (p.Subtitle.Length > 0 ? p.Subtitle : "—");

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(190) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });

            AddPriceCell(grid, 0, formCell, "#EFF6FF", "#1E3A8A", 11.5, wrap: true);
            AddPriceCell(grid, 1, brandCell, "#FFFFFF", "#0F172A", 12.5, bold: true, wrap: true);
            AddPriceCell(grid, 2, doseCell, "#FFFFFF", "#0F172A", 11.5, ltr: true, wrap: true);
            AddPriceCell(grid, 3, ircCell, "#F8FAFC", "#475569", 11, ltr: true);

            var btn = new System.Windows.Controls.Button
            {
                Content = grid,
                Style = (Style)FindResource("RoundedButtonStyle"),
                Background = System.Windows.Media.Brushes.White,
                Margin = new Thickness(0, 0, 0, 6),
                HorizontalContentAlignment = System.Windows.HorizontalAlignment.Stretch,
                VerticalContentAlignment = System.Windows.VerticalAlignment.Stretch,
                Padding = new Thickness(6, 6, 6, 6),
                MinHeight = 52,
                Height = double.NaN,
                Tag = p,
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = BuildPriceRowTooltip(p.Title),
            };
            System.Windows.Controls.ToolTipService.SetInitialShowDelay(btn, 250);
            System.Windows.Controls.ToolTipService.SetShowDuration(btn, 30000);
            btn.Click += async (s, args) =>
            {
                if (s is System.Windows.Controls.Button b && b.Tag is Services.PriceLookupService.ProductSummary sel)
                {
                    bool needsFetch = sel.ProductId > 0
                        && string.IsNullOrWhiteSpace(sel.EnName)
                        && string.IsNullOrWhiteSpace(sel.GenericCode)
                        && string.IsNullOrWhiteSpace(sel.BrandOwner);
                    if (needsFetch)
                    {
                        await RunProductDetailsAsync(sel.ProductId, sel.Title);
                        return;
                    }
                    ShowPriceResult(PriceLookup.ToPriceResult(sel));
                }
            };
            PriceLookupResultsList.Children.Add(btn);
        }
    }

    /// <summary>یک باکس (سلول) داخل ردیف لیست فرآورده‌ها می‌سازد</summary>
    private static void AddPriceCell(Grid grid, int col, string text, string bg, string fg, double fontSize, bool bold = false, bool ltr = false, bool wrap = false)
    {
        var border = new Border
        {
            Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString(bg),
            CornerRadius = new CornerRadius(6),
            Margin = new Thickness(2),
            Padding = new Thickness(8, 5, 8, 5),
            VerticalAlignment = System.Windows.VerticalAlignment.Stretch,
        };
        var tb = new System.Windows.Controls.TextBlock
        {
            Text = text,
            FontSize = fontSize,
            Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString(fg),
            TextWrapping = wrap ? System.Windows.TextWrapping.Wrap : System.Windows.TextWrapping.NoWrap,
            TextTrimming = wrap ? System.Windows.TextTrimming.None : System.Windows.TextTrimming.CharacterEllipsis,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            FontWeight = bold ? System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal,
        };
        if (ltr)
            tb.FlowDirection = System.Windows.FlowDirection.LeftToRight;
        border.Child = tb;
        Grid.SetColumn(border, col);
        grid.Children.Add(border);
    }

    private static System.Windows.Controls.ToolTip BuildPriceRowTooltip(string fullName)
    {
        return new System.Windows.Controls.ToolTip
        {
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xCB, 0xD5, 0xE1)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(12, 8, 12, 8),
            Content = new System.Windows.Controls.TextBlock
            {
                Text = string.IsNullOrWhiteSpace(fullName) ? "—" : fullName,
                FontSize = 13.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x0F, 0x17, 0x2A)),
                TextWrapping = System.Windows.TextWrapping.Wrap,
                MaxWidth = 520,
            }
        };
    }

    private async void PriceLookupSearchButton_Click(object sender, RoutedEventArgs e)
    {
        await RunActivePriceSearchAsync();
    }

    /// <summary>✕ جست‌وجوی جدید: همه‌ی کادرها پاک و آماده‌ی سرچ تازه</summary>
    private void PriceResetButton_Click(object sender, RoutedEventArgs e)
    {
        PriceNameInput.Text = "";
        PriceBarcodeInput.Text = "";
        PriceGenericInput.Text = "";
        PriceLookupStatusText.Text = "";
        PriceLookupStatusText.Visibility = Visibility.Collapsed;
        PriceLookupResultsList.Visibility = Visibility.Collapsed;
        PriceLookupResultsList.Children.Clear();
        HidePriceLookupFilterPanel();
        PriceLookupDetailsPanel.Visibility = Visibility.Collapsed;
        PriceLookupNotDrugWarning.Visibility = Visibility.Collapsed;
        PriceSearchInputsPanel.Visibility = Visibility.Visible;
        PriceNameInput.Focus();
    }

    private static bool IsAllDigits(string s)
    {
        if (s.Length == 0) return false;
        foreach (char c in s) if (c < '0' || c > '9') return false;
        return true;
    }

    /// <summary>اجرای استعلام: عددی طولانی = بارکد (کاتالوگ→IRC→فرآورده)؛ در غیر این صورت جست‌وجوی نامی با لیست انتخاب.</summary>
    private async Task RunPriceLookupAsync(string query, bool isFromScan)
    {
        if (string.IsNullOrWhiteSpace(query))
            return;

        bool isBarcode = isFromScan || (IsAllDigits(query) && query.Length >= 8);

        PriceLookupStatusText.Visibility = Visibility.Visible;
        PriceLookupDetailsPanel.Visibility = Visibility.Collapsed;
        PriceLookupResultsList.Visibility = Visibility.Collapsed;
        PriceLookupResultsList.Children.Clear();
        HidePriceLookupFilterPanel();

        try
        {
            if (isBarcode)
            {
                PriceLookupStatusText.Text = "🔍 در حال استعلام از کاتالوگ تی‌تک... (بارکد → IRC → فرآورده)";
                var result = await PriceLookup.LookupByBarcodeAsync(query, PriceLookupToken);
                ShowPriceResult(result);
            }
            else
            {
                PriceLookupStatusText.Text = "🔍 در حال جست‌وجوی فرآورده در تی‌تک... (تا دو صفحه نتیجه)";
                var products = await PriceLookup.SearchProductsAsync(query, PriceLookupToken);
                products = await PriceLookup.EnrichSummariesWithDetailsAsync(products, PriceLookupToken);
                await ShowProductSelectionListAsync(products, showFormDoseFilter: true);
            }
        }
        catch (Exception ex)
        {
            PriceLookupStatusText.Visibility = Visibility.Visible;
            PriceLookupStatusText.Text = "❌ خطا: " + ex.Message;
        }
    }

    private async Task RunProductDetailsAsync(long productId, string title)
    {
        string previousStatus = PriceLookupStatusText.Text;
        var previousStatusVis = PriceLookupStatusText.Visibility;

        PriceLookupStatusText.Visibility = Visibility.Visible;
        PriceLookupStatusText.Text = "🔍 در حال دریافت اطلاعات «" + title + "» از تی‌تک...";
        // لیست را باز نگه دار تا بعد از بستن پنجره قیمت بشود مورد بعدی را زد

        var result = await PriceLookup.GetProductDetailsAsync(productId, PriceLookupToken);

        PriceLookupStatusText.Text = previousStatus;
        PriceLookupStatusText.Visibility = previousStatusVis;
        ShowPriceResult(result);
    }

    private void ShowPriceResult(Services.PriceLookupService.PriceResult result)
    {
        if (!result.Success)
        {
            if (PriceLookupResultsList.Visibility == Visibility.Visible)
            {
                ShowStyledMessage("استعلام قیمت", result.Message, true);
                return;
            }
            PriceLookupStatusText.Visibility = Visibility.Visible;
            PriceLookupStatusText.Text = result.Message;
            return;
        }

        ShowPriceResultWindow(result);
    }

    [System.Obsolete]
    private void ShowPriceResultOld(Services.PriceLookupService.PriceResult result)
    {
        PriceLookupStatusText.Visibility = Visibility.Visible;

        if (!result.Success)
        {
            PriceLookupStatusText.Text = result.Message;
            PriceLookupDetailsPanel.Visibility = Visibility.Collapsed;
            return;
        }

        if (result.FoundButNotDrugSubgroup)
        {
            PriceLookupDetailsPanel.Visibility = Visibility.Visible;
            PriceLookupNotDrugWarning.Visibility = Visibility.Visible;
            PriceLookupFaNameText.Text = result.FaName;
            PriceLookupEnNameText.Text = result.EnName;
            PriceLookupGenericCodeText.Text = result.GenericCode;
            PriceLookupPackageCountText.Text = result.PackageCount;
            PriceLookupBrandOwnerText.Text = result.BrandOwner;
            PriceLookupPriceText.Text = "—";
            PriceLookupStatusText.Text = "نوع فرآورده: " + result.ProductType;
            return;
        }

        PriceLookupDetailsPanel.Visibility = Visibility.Visible;
        PriceLookupNotDrugWarning.Visibility = Visibility.Collapsed;
        PriceLookupFaNameText.Text = string.IsNullOrWhiteSpace(result.FaName) ? "—" : result.FaName;
        PriceLookupEnNameText.Text = string.IsNullOrWhiteSpace(result.EnName) ? "—" : result.EnName;
        PriceLookupGenericCodeText.Text = string.IsNullOrWhiteSpace(result.GenericCode) ? "—" : result.GenericCode;
        PriceLookupPackageCountText.Text = string.IsNullOrWhiteSpace(result.PackageCount) ? "—" : result.PackageCount;
        PriceLookupBrandOwnerText.Text = string.IsNullOrWhiteSpace(result.BrandOwner) ? "—" : result.BrandOwner;

        if (result.TotalPriceRial > 0)
        {
            // سه‌رقم سه‌رقم جدا شده، به ریال
            PriceLookupPriceText.Text = result.TotalPriceRial.ToString("N0", System.Globalization.CultureInfo.InvariantCulture);
            PriceLookupStatusText.Text = "";
            PriceLookupStatusText.Visibility = Visibility.Collapsed;
        }
        else
        {
            PriceLookupPriceText.Text = "—";
            PriceLookupStatusText.Text = "⚠️ قیمت مصرف‌کننده برای این فرآورده در تی‌تک ثبت نشده است.";
        }
    }

    // «اینترنت جانبی» در اتصال USB: آداپتور تترینگِ کابل به‌صورت پیش‌فرض یک مسیر پیش‌فرضِ تازه
    // با متریک پایین‌تر می‌سازد و اینترنت ویندوز را به شبکه‌ی بدون‌اینترنتِ گوشی می‌بُرد. با
    // بالا بردن متریک همان آداپتور (نیازمند تأیید مدیر، یک‌بار)، اینترنت از مسیر وای‌فای/اترنت
    // ادامه پیدا می‌کند و ارتباطِ هم‌ساب‌نت با گوشی دست‌نخورده می‌ماند.
    private void FixUsbInternetButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var targets = new List<string>();
            foreach (var ni in System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces())
            {
                if (ni.OperationalStatus != System.Net.NetworkInformation.OperationalStatus.Up)
                    continue;
                var d = (ni.Description + " " + ni.Name).ToLowerInvariant();
                if (!(d.Contains("rndis") || d.Contains("usb") || d.Contains("ncm") || d.Contains("remote ndis")))
                    continue;
                if (!ni.GetIPProperties().UnicastAddresses.Any(ua =>
                        ua.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork))
                    continue;
                targets.Add(ni.Name);
            }

            if (targets.Count == 0)
            {
                ShowStyledMessage("پیدا نشد", "آداپتور شبکه‌ی USB فعالی پیدا نشد. اول کابل را وصل و USB Tethering را روشن کنید.", true);
                return;
            }

            var cmds = string.Join(" & ", targets.Select(n =>
                "netsh interface ipv4 set interface interface=\"" + n + "\" metric=4000"));

            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c " + cmds + " & pause",
                Verb = "runas",
                UseShellExecute = true,
            };
            Process.Start(psi);

            ShowStyledMessage(
                "در حال رفع تداخل اینترنت",
                "در پنجره‌ی باز‌شده اجازه‌ی مدیر (Yes) را بدهید. بعد از آن اینترنت سیستم از مسیر وای‌فای/اترنت برمی‌گردد و اتصال کابل هم کار می‌کند.",
                false);
        }
        catch (Exception ex)
        {
            ShowStyledMessage("خطا", "اجرای دستور ممکن نشد: " + ex.Message, true);
        }
    }

    // ---------- Support ----------

    private void SupportButton_Click(object sender, RoutedEventArgs e)
    {
        if (SupportPanel != null)
            SupportPanel.Visibility = Visibility.Collapsed;
        if (SupportOverlay != null)
        {
            SupportOverlay.Visibility = Visibility.Visible;
            MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        }
    }

    private void CloseSupportOverlay()
    {
        if (SupportOverlay != null)
            SupportOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void SupportOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseSupportOverlay();
    }

    private void SupportOverlayCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void SupportOverlayCloseButton_Click(object sender, RoutedEventArgs e)
    {
        CloseSupportOverlay();
    }

    private void WhatsAppButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo("https://wa.me/989136346309") { UseShellExecute = true });
        }
        catch { }
    }

    // گزارش تشخیصی برای پشتیبانی: یک فایل متنی از وضعیت این سیستم (نسخه‌ی برنامه، شناسه/نام
    // سیستم، وضعیت لایسنس، وضعیت اتصال تی‌تک، و آخرین رخدادها/خطاهای ثبت‌شده در
    // startup-trace.log/startup-error.log) می‌سازد و روی دسکتاپ ذخیره می‌کند - تا داروخانه بدون
    // نیاز به اتصال از راه دور (RustDesk) بتواند همین یک فایل را برای پشتیبانی در واتساپ بفرستد و
    // پشتیبانی بدون تماس/توضیح شفاهی بفهمد دقیقاً سیستم چه وضعیتی دارد.
    private void SupportDiagnosticsReportButton_Click(object sender, RoutedEventArgs e)
    {
        // اول پنجره‌ی «پشتیبانی» را می‌بندیم؛ چون Panel.ZIndex آن (۳۷۰) از StyledMessageOverlay
        // (۳۵۰) بالاتر است و اگر باز بماند، پیام «گزارش تشخیصی آماده شد» پشت همین پنجره پنهان
        // می‌شود (قابل مشاهده نیست، فقط یک لبه‌ی محو از متنش پشت کادر پشتیبانی دیده می‌شود).
        CloseSupportOverlay();

        try
        {
            string reportPath = GenerateDiagnosticsReport();

            // عمداً هیچ پنجره‌ی بیرونی (اکسپلورر/واتساپ/مرورگر) باز نمی‌شود - چون آن پنجره جلوی همین
            // پیام می‌آید و به نظر می‌رسد برنامه چیزی نشان نداده. همه‌ی اطلاعات (مسیر فایل + راهنمای
            // آپلود در پنل + شماره‌ی واتساپ پشتیبانی به‌عنوان جایگزین) همین‌جا در متن پیام نوشته می‌شود.
            // دکمه‌ی «ورود به پنل» هم اضافه شده تا با یک کلیک مستقیم صفحه‌ی ورود باز شود.
            string title = _localization.CurrentLanguage == AppLanguage.English ? "Diagnostics report ready" : "گزارش تشخیصی آماده شد";
            string message = _localization.CurrentLanguage == AppLanguage.English
                ? $"The report was saved to the desktop:\n{Path.GetFileName(reportPath)}\n\nGo to scanbridge.ir/panel/login, open the \"Support\" tab, and upload this file - our team will reply to you right here in \"Messages\". (Or send it on WhatsApp to support: 09136346309)"
                : $"گزارش روی دسکتاپ ذخیره شد:\n{Path.GetFileName(reportPath)}\n\nوارد scanbridge.ir/panel/login بشید، بخش «پشتیبانی» رو باز کنید و همین فایل رو آپلود کنید - پاسخ تیم پشتیبانی همین‌جا توی «پیام‌ها» براتون میاد. (یا می‌تونید همین فایل رو در واتساپ به شماره‌ی ۰۹۱۳۶۳۴۶۳۰۹ بفرستید.)";
            ShowStyledMessage(
                title,
                message,
                false,
                linkUrl: "https://scanbridge.ir/panel/login",
                linkButtonText: _localization.CurrentLanguage == AppLanguage.English ? "Go to login page" : "ورود به پنل");
        }
        catch (Exception ex)
        {
            ShowStyledMessage(
                _localization.CurrentLanguage == AppLanguage.English ? "Error" : "خطا",
                (_localization.CurrentLanguage == AppLanguage.English ? "Could not create the diagnostics report: " : "ساخت گزارش تشخیصی ممکن نشد: ") + ex.Message,
                true);
        }
    }

    private string GenerateDiagnosticsReport()
    {
        var sb = new StringBuilder();
        sb.AppendLine("=== گزارش تشخیصی Scanbridge ===");
        sb.AppendLine($"زمان تهیه‌ی گزارش: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        sb.AppendLine("--- سیستم ---");
        sb.AppendLine($"نسخه‌ی برنامه: {GetCurrentAppVersionString()}");
        sb.AppendLine($"نام سیستم (ویندوز): {Environment.MachineName}");
        try { sb.AppendLine($"نام سیستم (داخل برنامه): {_service?.ComputerName}"); } catch { }
        try { sb.AppendLine($"شناسه‌ی سیستم: {_service?.ComputerId}"); } catch { }
        sb.AppendLine($"نسخه‌ی ویندوز: {Environment.OSVersion}");
        sb.AppendLine($"معماری سیستم: {(Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit")}");
        sb.AppendLine();

        sb.AppendLine("--- لایسنس ---");
        try
        {
            sb.AppendLine($"وضعیت: {(IsLicenseValid() ? "معتبر" : "نامعتبر/موجود نیست")}");
            sb.AppendLine($"کد لایسنس: {_activeLicense.LicenseId}");
            sb.AppendLine($"پلن: {_activeLicense.Plan}");
            sb.AppendLine($"داروخانه: {_activeLicense.PharmacyName}");
            sb.AppendLine($"تاریخ انقضا: {_activeLicense.ExpiresAt:yyyy-MM-dd}");
        }
        catch (Exception ex) { sb.AppendLine("خطا در خواندن اطلاعات لایسنس: " + ex.Message); }
        sb.AppendLine();

        sb.AppendLine("--- اتصال ---");
        try { sb.AppendLine($"تعداد گوشی‌های متصل: {_service?.ConnectedClients}"); } catch { }
        try { sb.AppendLine($"وضعیت اتصال تی‌تک: {(HasValidTtacToken() ? "متصل" : "قطع")}"); } catch { }
        sb.AppendLine();

        try
        {
            string? tracePath = GetAppLogFilePathForReading("startup-trace.log");
            if (tracePath != null)
            {
                var traceLines = File.ReadAllLines(tracePath);
                var lastLines = traceLines.Length > 150 ? traceLines[^150..] : traceLines;
                sb.AppendLine("--- آخرین رخدادهای راه‌اندازی (۱۵۰ خط آخر) ---");
                foreach (var line in lastLines)
                    sb.AppendLine(line);
                sb.AppendLine();
            }
        }
        catch { }

        try
        {
            string? errorPath = GetAppLogFilePathForReading("startup-error.log");
            if (errorPath != null)
            {
                string errorText = File.ReadAllText(errorPath);
                // اگر فایل خیلی بزرگ شده (روزها/هفته‌ها جمع شده)، فقط بخش پایانی (جدیدترین خطاها)
                // نگه داشته می‌شود تا گزارش قابل‌فرستادن (نه چند مگابایت) بماند.
                const int maxChars = 30000;
                if (errorText.Length > maxChars)
                    errorText = "...(بخش ابتدایی به‌خاطر حجم زیاد حذف شد)...\n" + errorText.Substring(errorText.Length - maxChars);
                sb.AppendLine("--- خطاهای ثبت‌شده ---");
                sb.AppendLine(errorText);
            }
            else
            {
                sb.AppendLine("--- خطاهای ثبت‌شده ---");
                sb.AppendLine("(خطایی ثبت نشده است)");
            }
        }
        catch { }

        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string fileName = $"ScanBridge-Diagnostics-{Environment.MachineName}-{DateTime.Now:yyyyMMdd-HHmmss}.txt";
        string fullPath = Path.Combine(desktopPath, fileName);
        File.WriteAllText(fullPath, sb.ToString(), Encoding.UTF8);
        return fullPath;
    }

    private void SupportSiteButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo("https://scanbridge.ir") { UseShellExecute = true });
        }
        catch { }
    }

    private void RemoteSupportButton_Click(object sender, RoutedEventArgs e)
    {
        WhatsAppButton_Click(sender, e);
    }

    private void LaunchRemoteSupportTool(bool showErrorIfMissing)
    {
        try
        {
            string? exePath = FindRemoteSupportExecutable();
            if (!string.IsNullOrWhiteSpace(exePath) && File.Exists(exePath))
            {
                Process.Start(new ProcessStartInfo(exePath) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory });
                return;
            }

            if (showErrorIfMissing)
            {
                string title = _localization.GetString("SupportToolNotFound");
                string message = _localization.GetString("RustDeskWasNotFoundPutRustdeskExeInsideTheSupportFolderNextToScanbridgeOrIncludeItInTheInstaller");
                ShowStyledMessage(title, message, true);
            }
        }
        catch (Exception ex)
        {
            ShowStyledMessage(
                _localization.GetString("SupportError"),
                (_localization.GetString("ErrorOpeningSupportTool")) + ex.Message,
                true);
        }
    }

    private string? FindRemoteSupportExecutable()
    {
        string[] fileNames = { "rustdesk.exe", "RustDesk.exe", "rustdesk-host.exe", "RustDeskHost.exe" };
        var directories = new List<string>
        {
            Path.Combine(AppContext.BaseDirectory, "Support"),
            AppContext.BaseDirectory,
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Scanbridge", "Support"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "RustDesk"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "RustDesk")
        };

        foreach (string directory in directories.Where(d => !string.IsNullOrWhiteSpace(d)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            foreach (string fileName in fileNames)
            {
                try
                {
                    string path = Path.Combine(directory, fileName);
                    if (File.Exists(path))
                        return path;
                }
                catch { }
            }
        }

        return null;
    }

    // ---------- Auto Startup ----------

    private void StartupOnWindowsBootCheckBox_Checked(object sender, RoutedEventArgs e)
    {
        SetStartupRegistryEntry(true);
    }

    private void StartupOnWindowsBootCheckBox_Unchecked(object sender, RoutedEventArgs e)
    {
        SetStartupRegistryEntry(false);
    }

    private static string GetExecutablePath()
    {
        var assemblyLocation = Assembly.GetExecutingAssembly().Location;
        if (!string.IsNullOrWhiteSpace(assemblyLocation))
            return assemblyLocation;

        var mainModulePath = Process.GetCurrentProcess().MainModule?.FileName;
        if (!string.IsNullOrWhiteSpace(mainModulePath))
            return mainModulePath;

        return Environment.ProcessPath ?? string.Empty;
    }

    private static bool IsStartupRegistryEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(StartupRegistryPath, false);
        return key?.GetValue(StartupRegistryValueName) is not null;
    }

    private void UpdateStartupCheckboxFromRegistry()
    {
        StartupOnWindowsBootCheckBox.IsChecked = IsStartupRegistryEnabled();
    }

    private static void SetStartupRegistryEntry(bool enabled)
    {
        var executablePath = GetExecutablePath();
        if (string.IsNullOrWhiteSpace(executablePath))
            return;

        using var key = Registry.CurrentUser.OpenSubKey(StartupRegistryPath, true) ?? Registry.CurrentUser.CreateSubKey(StartupRegistryPath);
        if (enabled)
        {
            key.SetValue(StartupRegistryValueName, $"\"{executablePath}\" --startup");
            return;
        }

        key.DeleteValue(StartupRegistryValueName, false);
    }

    // ---------- Settings ----------

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        OpenTtTeckSettings();
    }

    private void OpenTtTeckSettings()
    {
        LoadTtTeckSettings();
        bool ttacAllowed = HasLicenseModule("ttac");

        if (AppUpdateCurrentVersionText != null)
            AppUpdateCurrentVersionText.Text = (_localization.CurrentLanguage == AppLanguage.English
                ? "Installed version: "
                : "نسخه‌ی نصب‌شده: ") + GetCurrentAppVersionString();

        // «فعال‌سازی جستجوی خودکار تی‌تک» دیگر تیک دستی ندارد؛ خودش به‌صورت خودکار فقط برای
        // پلن‌های تی‌تک/تی‌تک‌پلاس فعال است (در LoadTtTeckSettings/SaveTtTeckSettings)، پس این
        // کارت از تنظیمات همیشه مخفی است و ستون وسط جمع می‌شود تا فضای خالی ایجاد نشود.
        if (TtTeckSettingsSection != null)
            TtTeckSettingsSection.Visibility = Visibility.Collapsed;

        if (SettingsTopRowGrid != null && SettingsTopRowGrid.ColumnDefinitions.Count >= 5)
        {
            SettingsTopRowGrid.ColumnDefinitions[1].Width = new GridLength(0);
            SettingsTopRowGrid.ColumnDefinitions[2].Width = new GridLength(0);
        }

        bool ttacPlusAllowed = IsLicenseValid() && _activeLicense.Plan == "TtacPlus";
        if (ExpiryAlertSettingsSection != null)
            ExpiryAlertSettingsSection.Visibility = ttacPlusAllowed ? Visibility.Visible : Visibility.Collapsed;
        if (ttacPlusAllowed)
        {
            LoadExpiryAlertSettings();
            if (ExpiryThresholdMonthsTextBox != null)
                ExpiryThresholdMonthsTextBox.Text = _expiryAlertSettings.ThresholdMonths.ToString(CultureInfo.InvariantCulture);
            if (RepeatReminderDaysTextBox != null)
                RepeatReminderDaysTextBox.Text = _expiryAlertSettings.RepeatReminderDays.ToString(CultureInfo.InvariantCulture);
            UpdateExpiryAlertSettingsLocalizedTexts();
        }
        else
        {
            StopBaleActivationPolling();
        }

        // حساب‌های ذخیره‌شده در پلن‌های تی‌تک و تی‌تک پلاس نمایش داده شود
        if (TtacSavedLoginsSection != null)
            TtacSavedLoginsSection.Visibility = ttacAllowed ? Visibility.Visible : Visibility.Collapsed;
        if (ttacAllowed)
        {
            RefreshTtacSavedLoginsList();
            UpdateTtacSavedLoginsLocalizedTexts();
        }

        TtTeckSettingsOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

 // =========================================================================================
 // هماهنگ‌سازی تنظیمات TtTeck/آستانه‌های تاریخ نزدیک/وضعیت فعال بودن بله بین سیستم‌های هم‌شبکه
 // که همان لایسنس را دارند - کاملاً محلی (روی همان شبکه‌ی وای‌فای/لن)، بدون هیچ سرور ابری. جزئیات
 // پروتکل (UDP discovery + WebSocket peer-to-peer) در ScanBridgeService.cs پیاده شده؛ اینجا فقط
 // کلید گروه (هش لایسنس) را به سرویس می‌دهیم و پیام‌های رسیده/تغییرات محلی را با آن رد و بدل می‌کنیم.
 // =========================================================================================

 private string ComputeLicenseGroupKey()
 {
     string raw = _activeLicense?.LicenseId ?? string.Empty;
     if (string.IsNullOrWhiteSpace(raw) || !IsLicenseValid())
         return string.Empty;

     byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(raw.Trim().ToUpperInvariant()));
     return Convert.ToHexString(hash);
 }

 // این تابع وقتی صدا زده می‌شود که یک پیام همگام‌سازی از یک سیستم هم‌شبکه (با همان لایسنس) برسد -
 // یا از طریق broadcast خودکار وقتی کسی چیزی عوض کرده، یا در پاسخ به درخواست ما موقع پیدا شدن یک
 // سیستم جدید. اگر نسخه‌ی رسیده جدیدتر از آخرین نسخه‌ی محلی باشد اعمال و ذخیره می‌شود؛ وگرنه نادیده
 // گرفته می‌شود (تا تنظیمات قدیمی‌تر یک سیستم تازه‌فعال‌شده، تنظیمات جدیدتر بقیه را خراب نکند).
 private void ApplyPeerDesktopSettings(string payloadJson, long versionUtcMs)
 {
     if (!Dispatcher.CheckAccess())
     {
         Dispatcher.BeginInvoke(new Action(() => ApplyPeerDesktopSettings(payloadJson, versionUtcMs)));
         return;
     }

     if (versionUtcMs <= _expiryAlertSettings.DesktopSettingsSyncVersionUtcMs)
         return;

     try
     {
         using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(payloadJson) ? "{}" : payloadJson);
         var root = doc.RootElement;

         if (root.TryGetProperty("ttTeckEnabled", out var ttEnabledProp) && ttEnabledProp.ValueKind is JsonValueKind.True or JsonValueKind.False)
             _ttTeckSettings.IsEnabled = ttEnabledProp.GetBoolean();
         // این مقدار همیشه باید از پلن لایسنس همین سیستم مشتق شود (نه از سیستم هم‌شبکه)، چون در
         // تئوری امکان دارد پیام هم‌زمانی قدیمی‌تر/دستگاه دیگری با تنظیمات دستیِ گذشته برسد.
         _ttTeckSettings.IsEnabled = HasLicenseModule("ttac");

         if (root.TryGetProperty("expiryThresholdMonths", out var thresholdProp) && thresholdProp.TryGetInt32(out var threshold) && threshold > 0)
             _expiryAlertSettings.ThresholdMonths = threshold;

         if (root.TryGetProperty("expiryRepeatReminderDays", out var repeatProp) && repeatProp.TryGetInt32(out var repeat) && repeat > 0)
             _expiryAlertSettings.RepeatReminderDays = repeat;

         if (root.TryGetProperty("baleNotificationsEnabled", out var baleProp) && baleProp.ValueKind is JsonValueKind.True or JsonValueKind.False)
             _expiryAlertSettings.IsBaleNotificationsEnabled = baleProp.GetBoolean();

         _expiryAlertSettings.DesktopSettingsSyncVersionUtcMs = versionUtcMs;
         SaveExpiryAlertSettings();

         string ttTeckPath = Path.Combine(AppContext.BaseDirectory, "ttteck-settings.json");
         try
         {
             File.WriteAllText(ttTeckPath, JsonSerializer.Serialize(_ttTeckSettings, new JsonSerializerOptions { WriteIndented = true }));
         }
         catch { }

         // اگر پنل تنظیمات تی‌تک همین الان روی صفحه باز است، مقادیر زنده هم به‌روزرسانی شوند.
         if (TtTeckEnabledCheckBox != null)
             TtTeckEnabledCheckBox.IsChecked = _ttTeckSettings.IsEnabled;
         if (ExpiryThresholdMonthsTextBox != null)
             ExpiryThresholdMonthsTextBox.Text = _expiryAlertSettings.ThresholdMonths.ToString(CultureInfo.InvariantCulture);
         if (RepeatReminderDaysTextBox != null)
             RepeatReminderDaysTextBox.Text = _expiryAlertSettings.RepeatReminderDays.ToString(CultureInfo.InvariantCulture);
         UpdateBaleConnectionStatusText();
         RefreshExpiryWatchDisplayList();
     }
     catch (Exception ex)
     {
         Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to apply peer desktop settings: {ex.Message}");
     }
 }

 // بعد از هر تغییر محلیِ یکی از این سه گروه تنظیمات صدا زده می‌شود: مقدار فعلی را با یک نسخه‌ی
 // زمانی جدید ذخیره و به سرویس می‌دهد تا برای سیستم‌های دیگر (با همان لایسنس، روی همین شبکه) پخش کند.
 private void PublishDesktopSettingsForSync()
 {
     try
     {
         long version = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
         var payloadObj = new
         {
             ttTeckEnabled = _ttTeckSettings.IsEnabled,
             expiryThresholdMonths = _expiryAlertSettings.ThresholdMonths,
             expiryRepeatReminderDays = _expiryAlertSettings.RepeatReminderDays,
             baleNotificationsEnabled = _expiryAlertSettings.IsBaleNotificationsEnabled
         };
         string json = JsonSerializer.Serialize(payloadObj);

         _expiryAlertSettings.DesktopSettingsSyncVersionUtcMs = version;
         SaveExpiryAlertSettings();

         _service?.PublishDesktopSettings(json, version);
     }
     catch { }
 }

 private void LoadTtTeckSettings()
{
    string settingsPath = Path.Combine(AppContext.BaseDirectory, "ttteck-settings.json");

    if (File.Exists(settingsPath))
    {
        try
        {
            string json = File.ReadAllText(settingsPath);
            _ttTeckSettings = JsonSerializer.Deserialize<TtTeckSettings>(json) ?? new();
        }
        catch
        {
            _ttTeckSettings = new();
        }
    }
    else
    {
        _ttTeckSettings = new();
    }

    // «جستجوی خودکار تی‌تک» دیگر سوئیچ دستی ندارد؛ به‌صورت خودکار فقط برای پلن‌های تی‌تک و
    // تی‌تک‌پلاس فعال است (نه پلن عادی) - کاربر نیازی به فعال‌سازی دستی آن از تنظیمات ندارد.
    _ttTeckSettings.IsEnabled = HasLicenseModule("ttac");
    TtTeckEnabledCheckBox.IsChecked = _ttTeckSettings.IsEnabled;
}

private void SaveTtTeckSettings()
{
    _ttTeckSettings.IsEnabled = HasLicenseModule("ttac");

    bool ttacPlusAllowed = IsLicenseValid() && _activeLicense.Plan == "TtacPlus";
    if (ttacPlusAllowed && ExpiryAlertSettingsSection != null && ExpiryAlertSettingsSection.Visibility == Visibility.Visible)
    {
        int months = int.TryParse((ExpiryThresholdMonthsTextBox?.Text ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedMonths) && parsedMonths > 0
            ? parsedMonths
            : 6;
        _expiryAlertSettings.ThresholdMonths = months;

        int repeatDays = int.TryParse((RepeatReminderDaysTextBox?.Text ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedRepeatDays) && parsedRepeatDays > 0
            ? parsedRepeatDays
            : 30;
        _expiryAlertSettings.RepeatReminderDays = repeatDays;

        SaveExpiryAlertSettings();
    }

    string settingsPath = Path.Combine(AppContext.BaseDirectory, "ttteck-settings.json");
    try
    {
        string json = JsonSerializer.Serialize(_ttTeckSettings, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(settingsPath, json);

        // این تنظیمات (فعال/غیرفعال تی‌تک + آستانه‌های تاریخ نزدیک) را برای سیستم‌های دیگرِ هم‌شبکه
        // که همان لایسنس را دارند هم پخش کن.
        PublishDesktopSettingsForSync();

        string title = _localization.GetString("Saved");

        string message = _localization.GetString("TtTeckSettingsWereSavedSuccessfully");

        CloseTtTeckSettings();
        ShowStyledMessage(title, message);
    }
    catch (Exception ex)
    {
        string title = _localization.GetString("Error");

        CloseTtTeckSettings();
        ShowStyledMessage(title, $"خطا: {ex.Message}", true);
    }
}
    private void CloseTtTeckSettings()
    {
        TtTeckSettingsOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
        StopBaleActivationPolling();

        // پاک کردن فرم ذخیره حساب تی‌تک
        TtacSavedLoginUsernameTextBox.Text = string.Empty;
        TtacSavedLoginPasswordBox.Password = string.Empty;
        TtacSavedLoginPasswordTextBox.Text = string.Empty;
        TtacSavedLoginPharmacyNameTextBox.Text = string.Empty;
        TtacSavedLoginStatusText.Visibility = Visibility.Collapsed;
        // ریست حالت نمایش رمز
        _isPasswordVisible = false;
        TtacSavedLoginPasswordBox.Visibility = Visibility.Visible;
        TtacSavedLoginPasswordTextBox.Visibility = Visibility.Collapsed;
        UpdateTtacSavedLoginPasswordToggleIcon(false);
    }

    private void TtTeckSettingsOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseTtTeckSettings();
    }

    private void TtTeckSettingsCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void TtTeckEnabledCheckBox_Checked(object sender, RoutedEventArgs e) { }
    private void TtTeckEnabledCheckBox_Unchecked(object sender, RoutedEventArgs e) { }

    private void TtTeckSaveButton_Click(object sender, RoutedEventArgs e)
    {
        SaveTtTeckSettings();
    }

    private void TtTeckCancelButton_Click(object sender, RoutedEventArgs e)
    {
        CloseTtTeckSettings();
    }

    private void ChangeLanguageButton_Click(object sender, RoutedEventArgs e)
    {
        TtTeckSettingsOverlay.Visibility = Visibility.Collapsed;
        LanguageOverlay.Visibility = Visibility.Visible;
        MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
    }

    private void AllowDeviceReconnectButton_Click(object sender, RoutedEventArgs e)
    {
        _service.AllowAllDevicesToReconnect();

        string title = _localization.GetString("ReconnectEnabled");

        string message = _localization.GetString("PhonesThatWereManuallyDisconnectedCanConnectAgainNow");

        CloseTtTeckSettings();
        ShowStyledMessage(title, message);
    }

    // ---------- QR Code & Print ----------

    private void OnLanIpChanged()
    {
        // IP نمایش‌داده‌شده در «اطلاعات سیستم» و خود QR اتصال را با IP جدید شبکه به‌روز کن.
        UpdateSystemInfo();
        RefreshQrCode();
    }

    private void PrintQrButton_Click(object sender, RoutedEventArgs e)
    {
        var printDialog = new System.Windows.Controls.PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var printContainer = new Grid
        {
            Width = 500,
            Height = 620,
            Background = System.Windows.Media.Brushes.White,
            FlowDirection = System.Windows.FlowDirection.RightToLeft
        };

        var image = new System.Windows.Controls.Image
        {
            Source = QrImage.Source,
            Width = 420,
            Height = 420,
            Stretch = Stretch.Uniform,
            Margin = new Thickness(0, 10, 0, 20),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top
        };

        var title = new TextBlock
        {
            Text = _service.ComputerName,
            FontSize = 22,
            FontWeight = FontWeights.Bold,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            Margin = new Thickness(0, 440, 0, 0)
        };

        printContainer.Children.Add(image);
        printContainer.Children.Add(title);

        printDialog.PrintVisual(printContainer, "Scanbridge QR");
    }

    private void RefreshQrCode()
    {
        var pngBytes = _service.CreatePairingQrPng();
        var image = LoadBitmapImage(pngBytes);
        Dispatcher.BeginInvoke(new Action(() => QrImage.Source = image));
    }

    // ---------- TtTeck Internal Browser ----------

    private async Task OpenTtTeckInternalBrowserAsync(string url)
    {
        try
        {
            _lastTtTeckWebViewUrl = url;
            _ttacBrowserOpenedAtUtc = DateTime.UtcNow;
            _ttacBrowserSlowWarningShown = false;
            _ttacLoginSuccessHandled = false;
            await EnsureTtTeckWebViewAsync();
            TtTeckWebView.Source = new Uri(url);
            TtTeckWebViewAddressText.Text = (_localization.GetString("ConnectingToTTAC")) + url;
            TtTeckWebViewOverlay.Visibility = Visibility.Visible;
            MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
            _ = ShowTtacBrowserSlowWarningIfNeededAsync(url);
        }
        catch (Exception ex)
        {
            string message = ex.Message;
            if (message.Contains("WebView2", StringComparison.OrdinalIgnoreCase)
                || message.Contains("CoreWebView2", StringComparison.OrdinalIgnoreCase)
                || message.Contains("runtime", StringComparison.OrdinalIgnoreCase))
            {
                message = _localization.GetString("MicrosoftEdgeWebView2RuntimeIsRequiredForTheInternalTTACBrowserPleaseInstallWebView2RuntimeThenOpenScanbridgeAgainDownloadHttpsDeveloperMicrosoftComMicrosoftEdgeWebview2");
            }

            string webView2DownloadUrl = "https://developer.microsoft.com/microsoft-edge/webview2/";
            ShowStyledMessage(
                _localization.GetString("BrowserError"),
                message,
                true,
                false,
                message.Contains("WebView2", StringComparison.OrdinalIgnoreCase) ? webView2DownloadUrl : null);
        }
    }

    private async Task ShowTtacBrowserSlowWarningIfNeededAsync(string url)
    {
        try
        {
            await Task.Delay(18000);
            if (_ttacBrowserSlowWarningShown || TtTeckWebViewOverlay.Visibility != Visibility.Visible || HasValidTtacToken())
                return;

            _ttacBrowserSlowWarningShown = true;
            string text = _localization.GetString("TTACIsLoadingSlowlyThisCanHappenDuringHighTTACTrafficOrInternetVPNDNSProblemsPleaseWaitIfItDoesNotOpenCloseAndTryAgain");
            TtTeckWebViewAddressText.Text = text;
        }
        catch { }
    }

    // بعد از موفق شدن ورود تی‌تک صدا زده می‌شود: هشدار «نشست منقضی شده» داخل پنجره‌ی انتخاب
    // داروخانه پاک می‌شود تا دفعه‌ی بعد که پنجره باز شد، با اطلاعات کهنه نمایش داده نشود.
    private void ClearTtacSessionExpiredWarning()
    {
        if (TtacQuickLoginWarningBorder != null)
            TtacQuickLoginWarningBorder.Visibility = Visibility.Collapsed;
        _pendingTtacRetryLabel = null;
        _ttacRetryUsername = null;
    }

    private async Task MonitorTtacConnectionAfterBrowserOpenAsync()
    {
        if (_isTtacTokenMonitorRunning)
            return;

        _isTtacTokenMonitorRunning = true;
        try
        {
            bool closedByUser = false;
            bool stuckOnLoginPage = false;
            int loginPageSeconds = 0;
            for (int i = 0; i < 90; i++)
            {
                await Task.Delay(1000);

                // اول توکن را چک کن - اگر ورود موفق شده، حتی اگر مرورگر داخلی (به‌خاطر موفق
                // شدن ورود) بسته شده باشد، عملیات در انتظار باید ادامه پیدا کند.
                string? token = await GetTtacAccessTokenOnUiThreadAsync(false);
                if (!string.IsNullOrWhiteSpace(token))
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        CompleteTtacInternalBrowserLogin(token, 5400);
                    }));
                    return;
                }

                // اگر مرورگر داخلی بسته شده ولی ورود موفق نشده، دیگر منتظر نمانیم.
                if (TtTeckWebViewOverlay.Visibility != Visibility.Visible)
                {
                    closedByUser = true;
                    break;
                }

                // اگر مرورگر همچنان روی صفحه‌ی ورود idp.ttac.ir مانده باشد، یعنی ورود انجام
                // نشده و کاربر احتمالاً روی همان صفحه گیر کرده (کد/رمز اشتباه، کپچا، یا ورود
                // کامل نشده). بعد از ۴۵ ثانیه روی همان صفحه، زودتر از مهلت عادی پیام خطا نشان بده.
                var source = TtTeckWebView.Source;
                bool onLoginPage = source != null
                    && source.Host != null
                    && source.Host.EndsWith("idp.ttac.ir", StringComparison.OrdinalIgnoreCase);
                if (onLoginPage)
                {
                    loginPageSeconds++;
                    if (loginPageSeconds >= 45)
                    {
                        stuckOnLoginPage = true;
                        break;
                    }
                }
                else
                {
                    loginPageSeconds = 0;
                }
            }

            _ttacQuickLoginInProgress = false;

            // مرورگر هنوز باز است ولی توکنی به‌دست نیامد (یا کاربر ۴۵ ثانیه روی صفحه‌ی ورود
            // گیر کرده) - یعنی ورود با همان حساب انجام نشده (رمز اشتباه، کپچا، یا حساب مسدود).
            // یک پیام خطای واضح با دکمه‌ی «تلاش مجدد» نشان بده.
            if (!closedByUser && !HasValidTtacToken())
            {
                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    // چون ورود از یک حسابِ ذخیره‌شده شروع شده، اگر روی صفحه‌ی ورود گیر کرده
                    // باشیم، محتمل‌ترین دلیل این است که کد/رمز ذخیره‌شده‌ی همان داروخانه درست
                    // نیست - نام داروخانه را پیدا کن تا در پیام خطا نوشته شود.
                    string pharmacyName = string.Empty;
                    if (!string.IsNullOrWhiteSpace(_ttacRetryUsername))
                    {
                        var saved = LoadSavedTtacLogins()
                            .FirstOrDefault(x => string.Equals(x.Username, _ttacRetryUsername, StringComparison.OrdinalIgnoreCase));
                        if (saved != null && !string.IsNullOrWhiteSpace(saved.PharmacyName))
                            pharmacyName = saved.PharmacyName;
                    }

                    string message;
                    if (stuckOnLoginPage)
                    {
                        message = !string.IsNullOrWhiteSpace(pharmacyName)
                            ? _localization.GetFormattedString("TtacLoginStuckOnLoginPageWithPharmacy", pharmacyName)
                            : _localization.GetString("TtacLoginStuckOnLoginPage");
                    }
                    else
                    {
                        message = !string.IsNullOrWhiteSpace(pharmacyName)
                            ? _localization.GetFormattedString("TtacLoginFailedWithPharmacy", pharmacyName)
                            : _localization.GetString("TtacLoginFailedGeneric");
                    }

                    StyledMessageRetryButton.Content = _localization.GetString("RetryButton");
                    StyledMessageRetryButton.Visibility = Visibility.Visible;
                    ShowStyledMessage(_localization.GetString("TtacLoginFailedTitle"), message, true);
                }));
            }
        }
        catch { }
        finally
        {
            _isTtacTokenMonitorRunning = false;
        }
    }

    // دکمه‌ی «تلاش مجدد» در پیام خطای ورود: دیالوگ را می‌بندد و مرورگر داخلی را با همان
    // داروخانه‌ی قبلی دوباره باز می‌کند تا ورود دوباره تلاش شود.
    private async void StyledMessageRetryButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StyledMessageRetryButton.Visibility = Visibility.Collapsed;
            CloseStyledMessage();

            string username = _ttacRetryUsername ?? string.Empty;
            if (string.IsNullOrWhiteSpace(username))
                return;

            _pendingTtacAutofillUsername = username;
            _ttacQuickLoginInProgress = true;
            await OpenTtTeckInternalBrowserAsync("https://newstatisticsreports.ttac.ir/pharmacyDashboard");
            _ = MonitorTtacConnectionAfterBrowserOpenAsync();
        }
        catch { }
    }

    // بعد از موفق شدن ورود تی‌تک (وقتی از پنجره‌ی انتخاب داروخانه شروع شده باشد) صدا زده
    // می‌شود: پنجره‌ی لیست داروخانه‌ها را فوری می‌بندد؛ دیگر ۲٫۵ ثانیه با بنر روی همان لیست
    // نمی‌ماند.
    private void ShowTtacLoginSuccessBanner(string? pendingLabel = null)
    {
        try
        {
            if (TtacQuickLoginSuccessBorder != null)
                TtacQuickLoginSuccessBorder.Visibility = Visibility.Collapsed;
            CloseTtacQuickLoginOverlay();
            _ = Dispatcher.BeginInvoke(new Action(FocusTtacActivePanelAfterLogin), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }
        catch { }
    }

    // فوکوس را به ورودی اصلیِ پنلی که بعد از ورود باز شده می‌دهد (همان رفتار هر پنل هنگام
    // باز شدن عادی). پنل‌هایی که ورودی ندارند (تاریخ انقضا، پنل تی‌تک) دست نمی‌خورند.
    private void FocusTtacActivePanelAfterLogin()
    {
        try
        {
            if (ReceiveStatusOverlay.Visibility == Visibility.Visible)
            {
                ReceiveStatusManualTextBox.Focus();
                return;
            }
            if (CargoDeliveryOverlay.Visibility == Visibility.Visible)
            {
                CargoDeliveryManualTextBox.Focus();
                return;
            }
            if (TtTeckRegistrationOverlay.Visibility == Visibility.Visible)
            {
                FocusAndSelect(TtTeckRegistrationNationalIdTextBox);
            }
        }
        catch { }
    }

    private string GetTtTeckWebViewUserDataFolder()
    {
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(localAppData))
            localAppData = AppContext.BaseDirectory;

        return Path.Combine(localAppData, "Scanbridge", "TtacWebView2");
    }

    private static void TrySetWebViewBoolProperty(object? target, string propertyName, bool value)
    {
        if (target == null)
            return;

        try
        {
            var property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            if (property != null && property.CanWrite && property.PropertyType == typeof(bool))
                property.SetValue(target, value);
        }
        catch { }
    }

    // این تنظیم به‌صورت پیش‌فرض false است (طبق مستندات WebView2)؛ یعنی تا وقتی صریحاً true
    // نشود، نه پیشنهاد ذخیره‌ی رمز نشان داده می‌شود و نه رمز جدیدی ذخیره می‌شود - برای همین
    // «مرورگر داخلی» تا الان رفتار Chrome (که این تنظیم را از قبل روشن دارد) را نداشت.
    // قبلاً این مقداردهی از طریق reflection (که اگر پراپرتی به هر دلیلی پیدا نمی‌شد، کاملاً
    // بی‌صدا شکست می‌خورد) انجام می‌شد؛ چون نسخه‌ی نصب‌شده‌ی WebView2 SDK (۱٫۰٫۴۱۲۹٫۵۰) این
    // پراپرتی‌ها را مستقیماً دارد، الان مستقیم تنظیم می‌شوند تا اگر مشکلی بود در کامپایل معلوم شود.
    private void EnableTtacWebViewAutofillAndPasswordSave()
    {
        try
        {
            var core = TtTeckWebView.CoreWebView2;
            if (core == null)
                return;

            core.Settings.IsPasswordAutosaveEnabled = true;
            core.Settings.IsGeneralAutofillEnabled = true;

            if (core.Profile != null)
            {
                core.Profile.IsPasswordAutosaveEnabled = true;
                core.Profile.IsGeneralAutofillEnabled = true;
            }
        }
        catch
        {
            try
            {
                var core = TtTeckWebView.CoreWebView2;
                TrySetWebViewBoolProperty(core?.Settings, "IsPasswordAutosaveEnabled", true);
                TrySetWebViewBoolProperty(core?.Settings, "IsGeneralAutofillEnabled", true);
                TrySetWebViewBoolProperty(core?.Profile, "IsPasswordAutosaveEnabled", true);
                TrySetWebViewBoolProperty(core?.Profile, "IsGeneralAutofillEnabled", true);
            }
            catch { }
        }
    }

    // صفحه‌ی ورود واقعی تی‌تک (idp.ttac.ir) فیلد رمز را با autocomplete="off" می‌فرستد و به فیلد
    // نام‌کاربری هم هیچ autocomplete‌ای نمی‌دهد؛ در نتیجه Chromium/WebView2 این فرم را «فرم ورود
    // قابل‌ذخیره» تشخیص نمی‌دهد و هرگز پیشنهاد «ذخیره‌ی رمز» را نشان نمی‌دهد (دیتابیس Login Data
    // همین‌جا هم خالی است و با اینکه IsPasswordAutosaveEnabled روشن است هیچ رمزی ذخیره نمی‌شود).
    // این اسکریپت فقط روی همان دامنه‌ی idp.ttac.ir، attribute های autocomplete را به مقادیر استاندارد
    // credential (username / current-password) تغییر می‌دهد تا مدیر رمز WebView2 بتواند پیشنهاد
    // ذخیره/پرکردن بدهد. رفتار خود سایت (ارسال فرم، کپچا و ...) دست‌نخورده می‌ماند؛ فقط یک hint
    // به مرورگر است.
    private void EnableTtacWebViewPasswordSaveOnIdp()
    {
        try
        {
            var core = TtTeckWebView.CoreWebView2;
            if (core == null)
                return;

            // این اسکریپت برای همه‌ی اسناد آینده‌ی مرورگر داخلی تی‌تک اجرا می‌شود؛ داخلش فقط به
            // دامنه‌ی idp.ttac.ir محدود شده. چون Angular فرم را بعد از بارگذاری مدل دوباره رندر
            // می‌کند، از MutationObserver استفاده می‌کنیم تا به محض اضافه‌شدن فیلدها (خیلی قبل از
            // تشخیص فرم توسط Chromium) attribute ها اصلاح شوند؛ polling هر ۲۵۰ms ممکن است دیرتر از
            // تشخیص اولیه‌ی PasswordManager اجرا شود.
            AppendTtacWebViewDebugLog($"Password-save script registered. Runtime={core.Environment?.BrowserVersionString}, Settings.IsPasswordAutosaveEnabled={core.Settings?.IsPasswordAutosaveEnabled}, Profile.IsPasswordAutosaveEnabled={core.Profile?.IsPasswordAutosaveEnabled}");
            _ = core.AddScriptToExecuteOnDocumentCreatedAsync(@"
(function() {
  try {
    if (window.location.hostname !== 'idp.ttac.ir') return;
    var FIXED = 'data-scanbridge-autofix';
    function fix() {
      try {
        var u = document.getElementById('username') || document.querySelector('input[name=username]');
        var p = document.getElementById('password') || document.querySelector('input[name=password]');
        var changed = false;
        if (u && u.getAttribute('autocomplete') !== 'username') { u.setAttribute('autocomplete', 'username'); changed = true; }
        if (p && p.getAttribute('autocomplete') !== 'current-password') { p.setAttribute('autocomplete', 'current-password'); changed = true; }
        if (u && p) document.documentElement.setAttribute(FIXED, '1');
        return changed;
      } catch (e) { return false; }
    }
    fix();
    try {
      if (window.MutationObserver) {
        var mo = new MutationObserver(fix);
        mo.observe(document.documentElement || document, { childList: true, subtree: true });
      }
    } catch (e) {}
    document.addEventListener('DOMContentLoaded', fix);
    window.addEventListener('load', fix);

    // چک‌باکس «من را بخاطر بسپار» خودِ سایت فقط نشست را طولانی‌تر می‌کند و کد/رمز را جایی
    // ذخیره نمی‌کند. این شنونده موقع ارسال فرم، اگر کاربر تیک را زده باشد، کد و رمز را به
    // برنامه می‌دهد تا رمزنگاری‌شده (مثل «حساب‌های ذخیره‌شده‌ی تی‌تک» در تنظیمات) ذخیره شود و
    // دفعه‌ی بعد که صفحه‌ی ورود باز شد خودکار پر شود.
    try {
      function attachSave() {
        try {
          var form = document.querySelector('form');
          if (!form || form.getAttribute('data-scanbridge-submit') === '1') return;
          form.setAttribute('data-scanbridge-submit', '1');
          form.addEventListener('submit', function() {
            try {
              var u = document.getElementById('username') || document.querySelector('input[name=username]');
              var p = document.getElementById('password') || document.querySelector('input[name=password]');
              var remember = document.getElementById('rememberMe') || document.querySelector('input[name=rememberMe]');
              if (u && p && remember && remember.checked && u.value && p.value && window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage({ type: 'saveTtacLogin', username: u.value, password: p.value });
              }
            } catch (e2) {}
          }, true);
        } catch (e3) {}
      }
      attachSave();
      if (window.MutationObserver) {
        var mo2 = new MutationObserver(attachSave);
        mo2.observe(document.documentElement || document, { childList: true, subtree: true });
      }
    } catch (e4) {}
  } catch (e) {}
})();");

            // پیام «ذخیره‌ی ورود» از صفحه‌ی ورود تی‌تک (وقتی کاربر «من را بخاطر بسپار» را زده و
            // فرم را ارسال کرده) دریافت و در همان مخزن امن «حساب‌های ذخیره‌شده‌ی تی‌تک» ذخیره می‌شود.
            core.WebMessageReceived += (_, e) =>
            {
                try
                {
                    using var doc = JsonDocument.Parse(e.WebMessageAsJson);
                    var root = doc.RootElement;
                    if (!root.TryGetProperty("type", out var typeProp))
                        return;

                    string? messageType = typeProp.GetString();
                    if (messageType == "ttacTokenFromUrl")
                    {
                        string href = root.TryGetProperty("href", out var hrefProp) ? (hrefProp.GetString() ?? string.Empty) : string.Empty;
                        Dispatcher.BeginInvoke(new Action(() => TryFinishTtacInternalBrowserLoginFromUrl(href)));
                        return;
                    }

                    if (messageType == "saveTtacLogin"
                        && root.TryGetProperty("username", out var uProp) && root.TryGetProperty("password", out var pProp))
                    {
                        string username = (uProp.GetString() ?? string.Empty).Trim();
                        string password = pProp.GetString() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(username) && !string.IsNullOrWhiteSpace(password))
                        {
                            UpsertSavedTtacLogin(username, password, null);
                            AppendTtacWebViewDebugLog($"TTAC remember-me saved via webview: {username}");
                            // اگر همین الان صفحه‌ی اصلی باز است، دکمه‌ی ورود سریع همان
                            // داروخانه هم ساخته شود.
                            RefreshTtacQuickLoginButtons();
                        }
                    }
                }
                catch { }
            };
        }
        catch { }
    }

    // لاگ اشکال‌زدایی مرورگر داخلی تی‌تک - کنار فایل برنامه نوشته می‌شود تا بعد از تست قابل
    // بررسی باشد (ttac-webview-debug.log). فقط برای عیب‌یابی؛ اگر نویزی بود حذف می‌شود.
    private static void AppendTtacWebViewDebugLog(string message)
    {
        try
        {
            string path = Path.Combine(AppContext.BaseDirectory, "ttac-webview-debug.log");
            File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}\r\n");
        }
        catch { }
    }

    // به‌محض دیدن access_token در آدرس (ریدایرکت callback)، پنجره را می‌بندد؛ دیگر منتظر
    // لود کامل داشبورد سنگین تی‌تک نمی‌ماند.
    private void TryFinishTtacInternalBrowserLoginFromUrl(string? url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return;
        if (!TryExtractTokenFromUri(uri, out var urlTokenResult) || string.IsNullOrWhiteSpace(urlTokenResult.AccessToken))
            return;

        CompleteTtacInternalBrowserLogin(urlTokenResult.AccessToken, Math.Max(60, urlTokenResult.ExpiresInSeconds - 60));
    }

    private void CompleteTtacInternalBrowserLogin(string token, int expiresInSeconds)
    {
        if (string.IsNullOrWhiteSpace(token))
            return;

        // توکن داروخانه‌ی قبلی را تا دیدن صفحه‌ی ورود idp قبول نکن.
        if (ShouldIgnoreStaleTtacToken())
            return;

        if (_ttacLoginSuccessHandled
            && HasValidTtacToken()
            && TtTeckWebViewOverlay.Visibility != Visibility.Visible)
            return;

        _ttacLoginSuccessHandled = true;
        _ttacWaitingForFreshLogin = false;
        _ttacSawIdpLoginPage = false;

        bool wasQuickLogin = _ttacQuickLoginInProgress;
        _ttacQuickLoginInProgress = false;
        string? pendingLabel = _pendingTtacRetryLabel;

        // اول پنجره را ببند تا کاربر معطل لود داشبورد یا نوشتن فایل نشود.
        if (TtTeckWebViewOverlay.Visibility == Visibility.Visible)
        {
            TtTeckWebViewOverlay.Visibility = Visibility.Collapsed;
            MainContent.Effect = TtTeckRegistrationOverlay.Visibility == Visibility.Visible
                || TtacPanelOverlay.Visibility == Visibility.Visible
                || CargoDeliveryOverlay.Visibility == Visibility.Visible
                || ReceiveStatusOverlay.Visibility == Visibility.Visible
                ? new System.Windows.Media.Effects.BlurEffect { Radius = 18 }
                : null;
            if (TtTeckRegistrationOverlay.Visibility == Visibility.Visible)
                TtTeckRegistrationResultText.Text = _localization.GetString("TTACLoginCompletedContinuingTheOperation");
        }

        ClearTtacSessionExpiredWarning();
        ApplyTtacAccessToken(token, DateTime.UtcNow.AddSeconds(expiresInSeconds));

        if (_pendingTtacRetryAction != null)
            _ = RunPendingTtacRetryIfAnyAsync();

        if (wasQuickLogin)
            ShowTtacLoginSuccessBanner(pendingLabel);
    }

    private async Task EnsureTtTeckWebViewAsync()
    {
        if (TtTeckWebView.CoreWebView2 != null)
            return;

        if (_ttTeckWebViewEnvironment == null)
        {
            string userDataFolder = GetTtTeckWebViewUserDataFolder();
            Directory.CreateDirectory(userDataFolder);
            _ttTeckWebViewEnvironment = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: userDataFolder);
        }

        await TtTeckWebView.EnsureCoreWebView2Async(_ttTeckWebViewEnvironment);
        if (TtTeckWebView.CoreWebView2 == null)
            return;

        EnableTtacWebViewAutofillAndPasswordSave();
        EnableTtacWebViewPasswordSaveOnIdp();

        _ = TtTeckWebView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(@"
(function(){
  try {
    function report() {
      try {
        var href = String(location.href || '');
        if (href.indexOf('access_token') === -1) return;
        if (window.chrome && window.chrome.webview) {
          window.chrome.webview.postMessage({ type: 'ttacTokenFromUrl', href: href });
        }
      } catch (e) {}
    }
    report();
    window.addEventListener('hashchange', report);
    window.addEventListener('load', report);
  } catch (e) {}
})();");

        TtTeckWebView.CoreWebView2.NavigationStarting += (_, e) =>
        {
            string uri = e.Uri;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                MarkTtacIdpLoginSeen(uri);
                TryFinishTtacInternalBrowserLoginFromUrl(uri);
            }));
        };
        TtTeckWebView.CoreWebView2.SourceChanged += (_, _) =>
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                string? href = TtTeckWebView.Source?.ToString();
                MarkTtacIdpLoginSeen(href);
                TryFinishTtacInternalBrowserLoginFromUrl(href);
            }));
        };

        TtTeckWebView.CoreWebView2.NavigationCompleted += (_, args) =>
        {
            Dispatcher.BeginInvoke(new Action(async () =>
            {
                try
                {
                    _lastTtTeckWebViewUrl = TtTeckWebView.Source?.ToString() ?? _lastTtTeckWebViewUrl;

                    // فقط برای عیب‌یابی ذخیره‌ی رمز: روی صفحه‌ی ورود واقعی تی‌تک، وضعیت فیلدهای
                    // فرم و اینکه آیا اسکریپت اصلاح autocomplete اجرا شده را در لاگ می‌نویسیم.
                    try
                    {
                        if (TtTeckWebView.Source?.Host?.Equals("idp.ttac.ir", StringComparison.OrdinalIgnoreCase) == true && TtTeckWebView.CoreWebView2 != null)
                        {
                            string probe = await TtTeckWebView.CoreWebView2.ExecuteScriptAsync(@"(function(){try{var u=document.getElementById('username')||document.querySelector('input[name=username]');var p=document.getElementById('password')||document.querySelector('input[name=password]');return JSON.stringify({fixed:document.documentElement?document.documentElement.getAttribute('data-scanbridge-autofix'):null,u:u?u.getAttribute('autocomplete'):null,p:p?p.getAttribute('autocomplete'):null,hasForm:!!(u&&p),url:location.href});}catch(e){return JSON.stringify({err:String(e)});}})()");
                            AppendTtacWebViewDebugLog($"Probe idp login page => {probe}");
                        }
                    }
                    catch { }

                    if (!args.IsSuccess)
                    {
                        TtTeckWebViewAddressText.Text = _localization.GetFormattedString("TtacPageLoadError", args.WebErrorStatus);
                    }
                    else
                    {
                        TtTeckWebViewAddressText.Text = _lastTtTeckWebViewUrl;
                    }

                    string? token = null;
                    int expiresIn = 5400;
                    if (TtTeckWebView.Source != null && TryExtractTokenFromUri(TtTeckWebView.Source, out var urlTokenResult))
                    {
                        token = urlTokenResult.AccessToken;
                        expiresIn = Math.Max(60, urlTokenResult.ExpiresInSeconds - 60);
                    }
                    else
                    {
                        token = await GetTtacAccessTokenAsync(false);
                    }

                    if (!string.IsNullOrWhiteSpace(token))
                    {
                        CompleteTtacInternalBrowserLogin(token, expiresIn);
                    }
                    else
                    {
                        // هنوز توکن معتبری نداریم - یعنی احتمالاً همین الان صفحه‌ی ورود تی‌تک است.
                        // اگر کاربر از تنظیمات، یوزرنیم/رمزی برای این حساب‌ها ذخیره کرده، همین‌جا
                        // (فقط یک‌بار برای همین بارگذاری صفحه) پر می‌شود.
                        _ = TryAutofillTtacWebViewLoginAsync();

                        // فقط برای عیب‌یابی: ۲ ثانیه بعد، ببین فیلدهای فرم چه وضعیتی دارند (مقدار
                        // رمز لاگ نمی‌شود - فقط طولش) تا معلوم شود پر شدن خودکار ماندگار است یا
                        // صفحه آن را دوباره خالی می‌کند.
                        if (TtTeckWebView.Source?.Host?.Equals("idp.ttac.ir", StringComparison.OrdinalIgnoreCase) == true && TtTeckWebView.CoreWebView2 != null)
                        {
                            var wv2 = TtTeckWebView.CoreWebView2;
                            _ = Dispatcher.BeginInvoke(new Action(async () =>
                            {
                                try
                                {
                                    await Task.Delay(2000);
                                    if (wv2 == null) return;
                                    string v = await wv2.ExecuteScriptAsync(@"(function(){try{var u=document.getElementById('username')||document.querySelector('input[name=username]');var p=document.getElementById('password')||document.querySelector('input[name=password]');var form=document.querySelector('form');var pwds=Array.prototype.slice.call(document.querySelectorAll('input[type=password]')).map(function(x){return x.id||x.name;});function vis(el){if(!el)return null;try{if(el.offsetParent===null)return 'hidden';var r=el.getBoundingClientRect();var s=window.getComputedStyle(el);return 'w='+Math.round(r.width)+',h='+Math.round(r.height)+',disp='+s.display+',op='+s.opacity;}catch(e){return 'err';}}return JSON.stringify({u:u?u.value:'',pLen:p?p.value.length:-1,pwds:pwds,log:(window.__sbAutofillLog||[]).slice(-15),formAction:form?(form.getAttribute('action')||'').substring(0,60):null,uVis:vis(u),pVis:vis(p)});}catch(e){return JSON.stringify({err:String(e)});}})()");
                                    AppendTtacWebViewDebugLog("Autofill probe (2s) => " + v);
                                }
                                catch { }
                            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                        }
                    }
                }
                catch { }
            }));
        };
    }

    private async void TtTeckWebViewCloseButton_Click(object sender, RoutedEventArgs e)
    {
        // پنجره بلافاصله بسته می‌شود — خواندن توکن از WebView (که اگر صفحه در حال بارگذاری
        // باشد یا رندرر مشغول باشد ممکن است چند ثانیه یا بیشتر معطل شود) به بعد از بستن
        // موکول شده تا دکمه‌ی ✕ هرگز گیر نکند. WebView بعد از مخفی‌شدن پنجره هم پابرجاست و
        // خواندن توکن در پس‌زمینه همچنان کار می‌کند.
        TtTeckWebViewOverlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
        UpdateTtacConnectionStatusUI();

        try
        {
            await GetTtacAccessTokenOnUiThreadAsync(false);
        }
        catch { }
        UpdateTtacConnectionStatusUI();
    }

    private void TtTeckWebViewBackButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (TtTeckWebView.CoreWebView2?.CanGoBack == true)
                TtTeckWebView.CoreWebView2.GoBack();
        }
        catch { }
    }

    private void TtTeckWebViewRefreshButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TtTeckWebView.CoreWebView2?.Reload();
        }
        catch { }
    }

    private void TtTeckWebViewOpenExternalButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string url = TtTeckWebView.Source?.ToString() ?? _lastTtTeckWebViewUrl;
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch { }
    }

    // ---------- Scan Toast ----------

    // شمارنده نسل toast: هر بار ShowScanToast صدا زده می‌شود یک نسل جدید می‌گیرد. اگر تا پایان
    // تایمر ۲.۵ ثانیه‌ای یک toast قدیمی‌تر، یک اسکن جدیدتر رسیده و toast را با محتوای خودش
    // به‌روزرسانی کرده باشد، فید-اوت تایمر قدیمی نباید اجرا شود - وگرنه پیام اسکن جدید را
    // زودتر از موعد قطع می‌کند (باگ ۱۱ گزارش).
    private int _scanToastGeneration;

    private void ShowScanToast(ScanRecord record, bool lookupSuccess, string productName)
    {
        Dispatcher.BeginInvoke(new Action(async () =>
        {
            int myGeneration = ++_scanToastGeneration;

            ScanToast.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81));
            ScanToastTitle.Text = _localization.GetString("NewBarcodeReceived");
            ScanToastMessage.Text = record.Barcode;

            ScanToast.Visibility = Visibility.Visible;
            ScanToast.Opacity = 0;

            var fadeIn = new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(180)));
            ScanToast.BeginAnimation(UIElement.OpacityProperty, fadeIn);

            await Task.Delay(2500);

            // اگر در همین فاصله یک اسکن جدیدتر toast را دوباره باز کرده، این تایمر قدیمی
            // نباید آن را ببندد.
            if (myGeneration != _scanToastGeneration)
                return;

            var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(250)));
            fadeOut.Completed += (_, _) =>
            {
                if (myGeneration == _scanToastGeneration)
                    ScanToast.Visibility = Visibility.Collapsed;
            };
            ScanToast.BeginAnimation(UIElement.OpacityProperty, fadeOut);
        }));
    }

    private string GetScanToastTitle(bool lookupSuccess, bool isTtTeck)
    {
        if (lookupSuccess)
        {
            return _localization.GetString("TtTeckProductFound");
        }

        if (isTtTeck && _ttTeckSettings.IsEnabled)
        {
            return _localization.GetString("ScanSavedTtTeckResultUnavailable");
        }

        return _localization.GetString("NewScanSaved");
    }

    // ---------- History Loading ----------

    private static DateTime ToLocalTimestamp(DateTime timestampUtc)
    {
        if (timestampUtc.Kind == DateTimeKind.Local)
            return timestampUtc;

        return DateTime.SpecifyKind(timestampUtc, DateTimeKind.Utc).ToLocalTime();
    }

    private static List<string> ParseCsvLine(string line)
    {
        var result = new List<string>();
        if (line is null)
            return result;

        var current = new System.Text.StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                result.Add(current.ToString().Trim());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        result.Add(current.ToString().Trim());
        return result;
    }

    private static bool IsLikelyScannedCode(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        string trimmed = value.Trim();
        if (trimmed.StartsWith("{", StringComparison.Ordinal))
            return true;

        string digitsOnly = new string(trimmed.Where(char.IsDigit).ToArray());
        if (digitsOnly.Length >= 8 && digitsOnly.Length >= trimmed.Length * 0.55)
            return true;

        return trimmed.StartsWith("01", StringComparison.Ordinal) && trimmed.Length >= 20;
    }

    private static (string Barcode, string DeviceName, string DrugName) ResolveBarcodeAndDeviceFromCsvParts(List<string> parts)
    {
        if (parts.Count < 3)
            return (parts.Count > 1 ? parts[1] : "", "", "");

        // فرمت جدید: timestamp,deviceName,barcode,drugName - چون هر فیلد به‌درستی escape شده
        // (EscapeCsvLocal)، دقیقاً ۴ ستون خواهیم داشت و نیازی به join کردن نیست. این‌طوری
        // کاما داخل نام دارو باعث به‌هم‌ریختن ستون بارکد/نام دستگاه نمی‌شود.
        // موقعیت ستون‌ها اینجا قطعی است (چون نوشتن همیشه با EscapeCsvLocal و به همین ترتیب انجام
        // می‌شود)، پس دیگر نباید heuristic حدسی (IsLikelyScannedCode) این دو ستون را جابه‌جا کند -
        // همان heuristic باعث می‌شد وقتی نام دستگاه کاملاً عددی (مثلاً شماره‌تلفن به‌عنوان اسم گوشی)
        // یا بارکد کوتاه/غیرعددی بود، بارکد و نام دستگاه در بارگذاری مجدد جابه‌جا نمایش داده شوند.
        if (parts.Count == 4)
            return (parts[2], parts[1], parts[3]);

        // فرمت قدیمی (فایل‌های ساخته‌شده با نسخه‌های قبلی برنامه): اگر نام دستگاه شامل
        // کامای escape‌نشده بود، ممکن بود بیش از ۳ ستون ایجاد شود؛ رفتار قبلی (join) حفظ می‌شود
        // و در این حالت نام دارو در دسترس نیست.
        string first = parts[1];
        string second = string.Join(",", parts.Skip(2));

        bool firstLooksBarcode = IsLikelyScannedCode(first);
        bool secondLooksBarcode = IsLikelyScannedCode(second);

        if (!firstLooksBarcode && secondLooksBarcode)
            return (second, first, "");

        return (first, second, "");
    }

    private void LoadHistoryFromCsv()
    {
        var rows = new List<ScanRecord>();
        var csvPath = GetActiveHistoryCsvPath();

        if (!File.Exists(csvPath))
        {
            Dispatcher.Invoke(() =>
            {
                HistoryItems.Clear();
                ApplyHistoryFilters();
            });
            return;
        }

        foreach (var line in File.ReadLines(csvPath).Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            var parts = ParseCsvLine(line);
            if (parts.Count < 2)
                continue;

            try
            {
                var ts = DateTimeOffset.Parse(parts[0].Trim(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal);
                var resolved = ResolveBarcodeAndDeviceFromCsvParts(parts);
                string barcode = resolved.Barcode.Trim();
                string deviceName = GetDeviceDisplayName(resolved.DeviceName.Trim());

                if (IsScanBridgePairPayload(barcode))
                    continue;

                var record = new ScanRecord(ts.LocalDateTime, barcode, deviceName);
                record.Source = BarcodeDetector.DetectBarcodeType(barcode);
                record.DrugName = (resolved.DrugName ?? "").Trim();
                rows.Add(record);
            }
            catch { }
        }

        rows = rows.OrderByDescending(r => r.TimestampLocal).ToList();
        Dispatcher.Invoke(() =>
        {
            HistoryItems.Clear();
            foreach (var row in rows)
            {
                HistoryItems.Add(row);
            }

            ApplyHistoryFilters();
        });
    }

    private void AddHistoryRecord(ScanRecord record)
    {
        Dispatcher.Invoke(() =>
        {
            foreach (var item in HistoryItems)
            {
                if (item.Barcode == record.Barcode && item.TimestampLocal == record.TimestampLocal)
                    return;
            }

            HistoryItems.Insert(0, record);
            ApplyHistoryFilters();
            SaveHistoryItemsToCsv();
        });
    }

    private void UpdateConnectionStatus(ConnectionState state, int connectedClients)
    {
        // این متد قبلاً فقط یک اکشن خالی صف می‌کرد و هیچ‌کاری با state/connectedClients
        // نمی‌کرد (باگ ۱۳ گزارش). حالا هم عنوان پنجره (که همیشه در دسترس است، بدون وابستگی
        // به نام کنترل‌های خاص در XAML) و هم - در صورت وجود - یک پنل وضعیت اتصال در XAML
        // به‌روزرسانی می‌شود.
        Dispatcher.BeginInvoke(new Action(() =>
        {
            bool english = _localization.CurrentLanguage == AppLanguage.English;
            string statusText = state switch
            {
                ConnectionState.Offline => _localization.GetString("NoPhoneConnected"),
                ConnectionState.Ready => _localization.GetString("Connected2"),
                ConnectionState.Busy => _localization.GetFormattedString("ConnectedCount", connectedClients),
                _ => _localization.GetString("Unknown")
            };

            string appTitle = _localization.GetString("ScanBridge");
            Title = $"{appTitle} - {statusText}";

            // در صورتی که یک پنل وضعیت اتصال با همین نام‌گذاری الگوی پنل تی‌تک در XAML وجود
            // داشته باشد، آن هم به‌روزرسانی می‌شود؛ در غیر این صورت این بخش بی‌اثر و بی‌خطر است.
            if (FindName("ConnectionStatusPanel") is Border connPanel &&
                FindName("ConnectionStatusIcon") is TextBlock connIcon &&
                FindName("ConnectionStatusText") is TextBlock connText)
            {
                connPanel.Visibility = Visibility.Visible;
                connText.Text = statusText;

                System.Windows.Media.Color dotColor = state switch
                {
                    ConnectionState.Offline => System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF),
                    ConnectionState.Ready => System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81),
                    ConnectionState.Busy => System.Windows.Media.Color.FromRgb(0xF5, 0x9E, 0x0B),
                    _ => System.Windows.Media.Color.FromRgb(0x9C, 0xA3, 0xAF)
                };
                connIcon.Text = "●";
                connIcon.Foreground = new SolidColorBrush(dotColor);
            }
        }));
    }

    private static BitmapImage LoadBitmapImage(byte[] pngBytes)
    {
        using var stream = new MemoryStream(pngBytes);
        var bitmap = new BitmapImage();
        bitmap.BeginInit();
        bitmap.CacheOption = BitmapCacheOption.OnLoad;
        bitmap.StreamSource = stream;
        bitmap.EndInit();
        bitmap.Freeze();
        return bitmap;
    }

    // ---------- License ----------

    private string GetLicenseFilePath()
    {
        return Path.Combine(AppContext.BaseDirectory, "scanbridge-license.json");
    }

    private bool IsLicenseValid()
    {
        return _activeLicense.IsValidForSystem(_service?.ComputerId ?? Environment.MachineName);
    }

    private bool HasLicenseModule(string moduleName)
    {
        return IsLicenseValid() && _activeLicense.Modules.IsEnabled(moduleName);
    }

    private void LoadAndApplyLicense()
    {
        try
        {
            string path = GetLicenseFilePath();
            if (File.Exists(path))
            {
                string code = File.ReadAllText(path).Trim();
                if (TryReadLicenseCode(code, out var license, out _))
                {
                    _activeLicense = license;
                    _activeLicenseCode = code;
                }
                else
                {
                    _activeLicense = ScanbridgeLicense.Missing();
                    _activeLicenseCode = string.Empty;
                }
            }
            else
            {
                _activeLicense = ScanbridgeLicense.Missing();
                _activeLicenseCode = string.Empty;
            }
        }
        catch
        {
            _activeLicense = ScanbridgeLicense.Missing();
            _activeLicenseCode = string.Empty;
        }

        ApplyLicenseToUi();

        // اگر لایسنس عوض شده باشد (نه تمدید)، حساب‌های ذخیره‌شده تی‌تک پاک می‌شوند.
        ClearSavedLoginsIfLicenseChanged();
    }

    private void SaveActiveLicense()
    {
        try
        {
            string code = !string.IsNullOrWhiteSpace(_activeLicenseCode) ? _activeLicenseCode : CreateLicenseCode(_activeLicense);
            File.WriteAllText(GetLicenseFilePath(), code, Encoding.UTF8);
        }
        catch { }
    }

    private void ApplyLicenseToUi()
    {
        bool valid = IsLicenseValid();
        bool barcode = HasLicenseModule("barcodeBridge");
        bool history = HasLicenseModule("history");
        bool excelPdf = HasLicenseModule("excelPdf");
        bool ttac = HasLicenseModule("ttac");
        bool cargo = HasLicenseModule("cargoDelivery");

        if (HistoryButton != null) HistoryButton.IsEnabled = valid && history;
        if (PrintQrButton != null) PrintQrButton.IsEnabled = valid && barcode;
        if (TtacPanelButton != null)
        {
            TtacPanelButton.IsEnabled = valid && ttac;
            TtacPanelButton.Visibility = valid && ttac ? Visibility.Visible : Visibility.Collapsed;
        }
        if (CargoDeliveryPanelButton != null)
        {
            CargoDeliveryPanelButton.IsEnabled = valid && cargo;
            CargoDeliveryPanelButton.Visibility = valid && cargo ? Visibility.Visible : Visibility.Collapsed;
        }
        if (TtacPanelReceiveStatusButton != null)
            TtacPanelReceiveStatusButton.Visibility = HasLicenseModule("receiveStatus") ? Visibility.Visible : Visibility.Collapsed;
        if (TtTeckSettingsSection != null)
            TtTeckSettingsSection.Visibility = valid && ttac ? Visibility.Visible : Visibility.Collapsed;
        // حساب‌های ذخیره‌شده در پلن‌های تی‌تک و تی‌تک پلاس
        if (TtacSavedLoginsSection != null)
            TtacSavedLoginsSection.Visibility = valid && ttac ? Visibility.Visible : Visibility.Collapsed;
        if (ExportTtTeckOptionBorder != null)
            ExportTtTeckOptionBorder.Visibility = valid && ttac ? Visibility.Visible : Visibility.Collapsed;

        // بانک بارکد پرمصرف فقط برای پلن تی‌تک‌پلاس مجاز است.
        bool highUsageBarcodeAllowed = valid && _activeLicense.Plan == "TtacPlus";
        if (HighUsageBarcodePanelButton != null)
            HighUsageBarcodePanelButton.Visibility = highUsageBarcodeAllowed ? Visibility.Visible : Visibility.Collapsed;
        if (HighUsageWidgetEnableCheckBox != null)
            HighUsageWidgetEnableCheckBox.Visibility = highUsageBarcodeAllowed ? Visibility.Visible : Visibility.Collapsed;
        if (!highUsageBarcodeAllowed && _highUsageSettings.WidgetEnabled)
        {
            // اگر لایسنس دیگر تی‌تک‌پلاس نیست (تغییر/تمدید با پلن پایین‌تر)، آیکون شناور را خاموش کن.
            _highUsageSettings.WidgetEnabled = false;
            SaveHighUsageBarcodeSettings();
            HideHighUsageWidget();
            if (HighUsageWidgetEnableCheckBox != null)
                HighUsageWidgetEnableCheckBox.IsChecked = false;
        }

        if (ExportExcelButton != null) ExportExcelButton.IsEnabled = valid && excelPdf;
        if (ExportPdfButton != null) ExportPdfButton.IsEnabled = valid && excelPdf;
        if (TtacPanelExportExcelButton != null) TtacPanelExportExcelButton.IsEnabled = valid && excelPdf;
        if (TtacPanelExportPdfButton != null) TtacPanelExportPdfButton.IsEnabled = valid && excelPdf;
        if (ReceiveStatusExportExcelButton != null) ReceiveStatusExportExcelButton.IsEnabled = valid && excelPdf;
        if (ReceiveStatusExportPdfButton != null) ReceiveStatusExportPdfButton.IsEnabled = valid && excelPdf;
        if (CargoDeliveryExportExcelButton != null) CargoDeliveryExportExcelButton.IsEnabled = valid && excelPdf;
        if (CargoDeliveryExportPdfButton != null) CargoDeliveryExportPdfButton.IsEnabled = valid && excelPdf;

        if (valid)
            StartLicenseHeartbeatTimer();
        else
            _licenseHeartbeatTimer.Stop();

        UpdateTtacConnectionStatusUI();
        UpdateLicenseOverlayTextsSafe();

        // نمایش/عدم‌نمایش دکمه‌ی «تاریخ نزدیک» (که فقط پلن تی‌تک‌پلاس مجاز است) داخل
        // RefreshExpiryWatchDisplayList محاسبه می‌شود؛ اگر اینجا هم صدا زده نشود، با تغییر پلن
        // لایسنس (فعال‌سازی/تمدید/غیرفعال شدن) دکمه تا باز شدن بعدیِ یک پنل دیگر یا تغییر زبان
        // به‌روز نمی‌شود و همچنان با پلن قبلی روی صفحه می‌ماند.
        RefreshExpiryWatchDisplayList();

        // به سرویس می‌گوییم این سیستم عضو کدام «گروه لایسنس» است تا فقط با سیستم‌های هم‌شبکه‌ای که
        // همان لایسنس را دارند (نه هر سیستم دیگری روی همان وای‌فای) تنظیمات را رد و بدل کند؛ و بلافاصل�کند؛ و بلافاصله
        // یک اعلان می‌فرستد تا اگر همکاری روی همین شبکه از قبل بالا بوده، زودتر (نه بعد از ۱۵ ثانیه)
        // همدیگر را پیدا و تنظیمات را هماهنگ کنند.
        _service?.SetLicenseGroupKey(ComputeLicenseGroupKey());
        _service?.AnnounceNow();
    }

    private void ShowLicenseOverlayStrict()
    {
        UpdateLicenseOverlayTextsSafe();
        if (FindName("LicenseOverlay") is FrameworkElement overlay)
        {
            overlay.Visibility = Visibility.Visible;
            MainContent.Effect = new System.Windows.Media.Effects.BlurEffect { Radius = 18 };
        }
    }

    private static string ToBase64Url(byte[] bytes)
    {
        return Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    private static byte[] FromBase64Url(string value)
    {
        value = value.Replace('-', '+').Replace('_', '/');
        value = value.PadRight(value.Length + ((4 - value.Length % 4) % 4), '=');
        return Convert.FromBase64String(value);
    }

    private string SignLicensePayload(string payloadBase64Url)
    {
        using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(LicenseHmacSecret));
        return ToBase64Url(hmac.ComputeHash(Encoding.UTF8.GetBytes(payloadBase64Url)));
    }

    private string CreateLicenseCode(ScanbridgeLicense license)
    {
        // فقط برای سازگاری با لایسنس‌های قدیمی/تست داخلی نگه داشته شده است.
        // لایسنس فروش عمومی از سرور با RSA و فرمت SCB2 صادر می‌شود.
        license.Signature = string.Empty;
        string json = JsonSerializer.Serialize(license);
        string payload = ToBase64Url(Encoding.UTF8.GetBytes(json));
        string signature = SignLicensePayload(payload);
        return $"{LegacyLicenseCodePrefix}.{payload}.{signature}";
    }

    private bool TryReadLicenseCode(string code, out ScanbridgeLicense license, out string error)
    {
        license = ScanbridgeLicense.Missing();
        error = string.Empty;
        try
        {
            code = (code ?? string.Empty).Trim();
            if (File.Exists(code))
                code = File.ReadAllText(code).Trim();

            if (code.StartsWith("{", StringComparison.Ordinal))
            {
                error = _localization.GetString("RawJSONLicenseFilesAreNotAcceptedUseTheSignedActivationCode");
                return false;
            }

            var parts = code.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (parts.Length != 3)
            {
                error = _localization.GetString("LicenseCodeFormatIsInvalid");
                return false;
            }

            if (parts[0].Equals(ServerLicenseCodePrefix, StringComparison.OrdinalIgnoreCase))
                return TryReadServerSignedLicense(parts, out license, out error);

            // فرمت قدیمی SCB1 دیگر پذیرفته نمی‌شود - طبق ممیزی امنیتی، این فرمت با یک رمز HMAC
            // که به‌صورت ثابت داخل خودِ فایل exe نوشته شده بود (LicenseHmacSecret) هم امضا و هم
            // تایید می‌شد؛ یعنی هرکس آن رشته را از فایل اجرایی استخراج می‌کرد (با یک decompiler
            // معمولی) می‌توانست خودش یک کد لایسنس کاملاً معتبر با هر پلن/تاریخ انقضایی بسازد،
            // بدون هیچ تماسی با سرور. لایسنس فروش واقعی همیشه فرمت SCB2 با امضای RSA سرور است
            // (TryReadServerSignedLicense) که این مشکل را ندارد. اگر یک کد SCB1 واقعی و قدیمی
            // جایی دست کسی مانده، دیگر پذیرفته نمی‌شود و باید یک کد SCB2 جدید از پشتیبانی گرفته
            // شود.
            if (parts[0].Equals(LegacyLicenseCodePrefix, StringComparison.OrdinalIgnoreCase))
            {
                error = _localization.CurrentLanguage == AppLanguage.English
                    ? "This activation code format is no longer supported. Please request a new activation code from support."
                    : "این فرمت کد فعال‌سازی دیگر پشتیبانی نمی‌شود. لطفاً یک کد فعال‌سازی جدید از پشتیبانی درخواست کنید.";
                return false;
            }

            error = _localization.GetString("LicenseCodeFormatIsInvalid");
            return false;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private bool TryReadLegacyHmacLicense(string[] parts, out ScanbridgeLicense license, out string error)
    {
        license = ScanbridgeLicense.Missing();
        error = string.Empty;

        string expectedSignature = SignLicensePayload(parts[1]);
        if (!CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(expectedSignature), Encoding.UTF8.GetBytes(parts[2])))
        {
            error = _localization.GetString("LicenseSignatureIsInvalid");
            return false;
        }

        string json = Encoding.UTF8.GetString(FromBase64Url(parts[1]));
        license = JsonSerializer.Deserialize<ScanbridgeLicense>(json) ?? ScanbridgeLicense.Missing();
        license.Signature = parts[2];
        return true;
    }

    private bool TryReadServerSignedLicense(string[] parts, out ScanbridgeLicense license, out string error)
    {
        license = ScanbridgeLicense.Missing();
        error = string.Empty;

        byte[] signature = FromBase64Url(parts[2]);
        using var rsa = RSA.Create();
        rsa.ImportFromPem(LicenseServerPublicKeyPem.AsSpan());

        bool ok = rsa.VerifyData(
            Encoding.UTF8.GetBytes(parts[1]),
            signature,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);

        if (!ok)
        {
            error = _localization.GetString("LicenseSignatureIsInvalid");
            return false;
        }

        string json = Encoding.UTF8.GetString(FromBase64Url(parts[1]));
        using var doc = JsonDocument.Parse(json);
        license = BuildLicenseFromServerPayload(doc.RootElement);
        license.Signature = parts[2];
        return true;
    }

    private static string JsonString(JsonElement element, string name, string fallback = "")
    {
        return element.TryGetProperty(name, out var prop) && prop.ValueKind != JsonValueKind.Null ? (prop.ToString() ?? fallback) : fallback;
    }

    private static DateTime JsonDate(JsonElement element, string name)
    {
        string value = JsonString(element, name);
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt))
            return dt;
        return DateTime.MinValue;
    }

    private ScanbridgeLicense BuildLicenseFromServerPayload(JsonElement root)
    {
        string plan = JsonString(root, "plan", "Missing");
        string status = JsonString(root, "status", "inactive");
        string deviceId = JsonString(root, "device_id", "");
        string licenseKey = JsonString(root, "license_key", "");
        string customerName = JsonString(root, "customer_name", "");
        string pharmacyName = JsonString(root, "pharmacy_name", "");
        string customerPhone = JsonString(root, "mobile", "");

        var modules = LicensedModules.ForPlan(plan);
        if (root.TryGetProperty("features", out var features) && features.ValueKind == JsonValueKind.Object)
        {
            bool GetFeature(string name)
                => features.TryGetProperty(name, out var prop) && prop.ValueKind == JsonValueKind.True;

            modules.BarcodeBridge = GetFeature("barcodeBridge");
            modules.History = GetFeature("history");
            modules.ExcelPdf = GetFeature("exports") || GetFeature("excelPdf");
            modules.Ttac = GetFeature("ttac");
            modules.TtacRegistration = modules.Ttac;
            modules.Formula = GetFeature("formula");
            modules.ReceiveStatus = GetFeature("receiveStatus");
            modules.CargoDelivery = GetFeature("cargoDelivery");
            modules.DeviceManagement = modules.BarcodeBridge;
            modules.MonthlyArchive = modules.Ttac;
            modules.PharmacyMemory = modules.Ttac;
        }

        if (!string.Equals(status, "active", StringComparison.OrdinalIgnoreCase))
            modules = new LicensedModules();

        return new ScanbridgeLicense
        {
            LicenseId = JsonString(root, "license_id", licenseKey),
            CustomerName = !string.IsNullOrWhiteSpace(customerName) ? customerName : pharmacyName,
            PharmacyName = pharmacyName,
            CustomerPhone = customerPhone,
            CustomerType = plan.StartsWith("Ttac", StringComparison.OrdinalIgnoreCase) || !string.IsNullOrWhiteSpace(pharmacyName) ? "Pharmacy" : "General",
            Product = plan switch
            {
                "Normal" => "Scanbridge Barcode",
                "Ttac" => "Scanbridge Ttac",
                "TtacPlus" => "Scanbridge Ttac Plus",
                "Trial" => "Scanbridge Trial",
                _ => "Scanbridge"
            },
            Plan = plan,
            SystemId = deviceId,
            IssuedAt = JsonDate(root, "issued_at"),
            ExpiresAt = JsonDate(root, "expires_at"),
            MaxDevices = 1,
            Modules = modules,
            Signature = string.Empty
        };
    }

    private static string ExtractApiMessage(string content, string fallback)
    {
        try
        {
            using var doc = JsonDocument.Parse(content);
            if (doc.RootElement.TryGetProperty("message", out var message) && message.ValueKind != JsonValueKind.Null)
                return message.ToString();
        }
        catch { }
        return fallback;
    }

    private async Task<string> ActivateLicenseOnlineAsync(string licenseKey)
    {
        string systemId = _service?.ComputerId ?? Environment.MachineName;
        var body = new Dictionary<string, object?>
        {
            ["license_key"] = licenseKey.Trim(),
            ["device_id"] = systemId,
            ["device_name"] = Environment.MachineName,
            ["app_version"] = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0"
        };

        using var requestContent = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
        using var response = await _licenseHttpClient.PostAsync($"{LicenseApiBaseUrl}/license/activate", requestContent);
        string content = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException(ExtractApiMessage(content, $"HTTP {(int)response.StatusCode}"));

        using var doc = JsonDocument.Parse(content);
        if (!doc.RootElement.TryGetProperty("success", out var successProp) || successProp.ValueKind != JsonValueKind.True)
            throw new InvalidOperationException(ExtractApiMessage(content, _localization.GetString("ActivationFailed")));

        if (!doc.RootElement.TryGetProperty("signed_license", out var signedProp))
            throw new InvalidOperationException(_localization.GetString("SignedLicenseWasNotReturnedByServer"));

        _lastLicenseOnlineValidationUtc = DateTime.UtcNow;
        return signedProp.GetString() ?? string.Empty;
    }

    private void StartLicenseHeartbeatTimer()
    {
        if (!_licenseHeartbeatTimer.IsEnabled)
            _licenseHeartbeatTimer.Start();
    }

    private async Task ValidateActiveLicenseOnlineBestEffortAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_activeLicenseCode) || !_activeLicenseCode.StartsWith(ServerLicenseCodePrefix + ".", StringComparison.OrdinalIgnoreCase))
                return;

            string systemId = _service?.ComputerId ?? Environment.MachineName;
            var body = new Dictionary<string, object?>
            {
                ["license_key"] = _activeLicense.LicenseId.StartsWith("SCB-", StringComparison.OrdinalIgnoreCase) ? _activeLicense.LicenseId : string.Empty,
                ["device_id"] = systemId
            };

            // در Payload فعلی، license_id ممکن است عدد باشد؛ پس کلید را از خود کد امضاشده می‌خوانیم.
            using (var doc = JsonDocument.Parse(Encoding.UTF8.GetString(FromBase64Url(_activeLicenseCode.Split('.')[1]))))
            {
                body["license_key"] = JsonString(doc.RootElement, "license_key", body["license_key"]?.ToString() ?? string.Empty);
            }

            using var requestContent = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
            using var response = await _licenseHttpClient.PostAsync($"{LicenseApiBaseUrl}/license/validate", requestContent);
            string content = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                // فقط کدهای HTTP که یعنی سرور صراحتاً این لایسنس را رد/باطل کرده (نه یک خطای
                // موقت شبکه/سرور) باعث خروج فوری کاربر شوند. کدهای دیگر (500/502/503/504/429 و...)
                // یعنی سرور یا شبکه موقتاً در دسترس نیست - نباید لایسنس معتبرِ کش‌شده را باطل کند،
                // وگرنه یک خطای موقت سرور می‌تواند مشتریانی که پول داده‌اند را قفل کند.
                bool isDefiniteRejection = response.StatusCode == System.Net.HttpStatusCode.Unauthorized
                    || response.StatusCode == System.Net.HttpStatusCode.Forbidden
                    || response.StatusCode == System.Net.HttpStatusCode.NotFound
                    || response.StatusCode == System.Net.HttpStatusCode.Gone;

                if (isDefiniteRejection)
                {
                    _activeLicense = ScanbridgeLicense.Missing();
                    _activeLicenseCode = string.Empty;
                    ApplyLicenseToUi();
                    Dispatcher.BeginInvoke(new Action(ShowLicenseOverlayStrict));
                }
                // در غیر این صورت (خطای موقت سرور/شبکه): لایسنس کش‌شده دست‌نخورده باقی می‌ماند و
                // دفعه‌ی بعدِ heartbeat (۶ ساعت دیگر) دوباره تلاش می‌شود.
                return;
            }

            _lastLicenseOnlineValidationUtc = DateTime.UtcNow;

            using var responseDoc = JsonDocument.Parse(content);
            if (responseDoc.RootElement.TryGetProperty("signed_license", out var signedProp))
            {
                string signed = signedProp.GetString() ?? string.Empty;
                if (TryReadLicenseCode(signed, out var refreshed, out _))
                {
                    _activeLicense = refreshed;
                    _activeLicenseCode = signed;
                    SaveActiveLicense();
                    ApplyLicenseToUi();

                    // اگر لایسنس در اعتبارسنجی آنلاین تغییر کرده باشد (تمدید نباشد)، حساب‌ها پاک شوند.
                    ClearSavedLoginsIfLicenseChanged();
                }
            }

            // پیام‌های اختصاصیِ پشتیبانی - وقتی ادمین توی صفحه‌ی «پشتیبانی» (روی سرور) برای همین لایسنس
            // جواب می‌نویسد، همین درخواست /license/validate (که هر ۶ ساعت + موقع بالا آمدن برنامه صدا
            // زده می‌شود) آن را در فیلد support_messages برمی‌گرداند. سرور خودش بعد از یک‌بار فرستادن
            // این پیام را «تحویل‌شده» علامت می‌زند، پس اینجا نیازی به هیچ dedup ای نیست - هر آیتمی که
            // برسد، تازه و ندیده است.
            if (responseDoc.RootElement.TryGetProperty("support_messages", out var supportMsgsProp)
                && supportMsgsProp.ValueKind == JsonValueKind.Array
                && supportMsgsProp.GetArrayLength() > 0)
            {
                bool anyAdded = false;
                foreach (var item in supportMsgsProp.EnumerateArray())
                {
                    string messageBody = JsonString(item, "body", "");
                    if (string.IsNullOrWhiteSpace(messageBody))
                        continue;

                    string title = JsonString(item, "title", _localization.CurrentLanguage == AppLanguage.English ? "Support reply" : "پاسخ پشتیبانی اسکن‌بریج");
                    Messages.Add(new AppMessage
                    {
                        Title = title,
                        Body = messageBody,
                        IsRead = false
                    });
                    anyAdded = true;
                }

                if (anyAdded)
                {
                    SaveMessagesToDisk();
                    UpdateMessagesBadgeCount();
                }
            }
        }
        catch
        {
            // اگر اینترنت/سرور موقتاً در دسترس نبود، لایسنس کش‌شده معتبر می‌ماند.
        }
    }

    private async Task ActivateLicenseFromCodeAsync(string code)
    {
        try
        {
            code = (code ?? string.Empty).Trim();
            LicenseMessageText.Text = _localization.GetString("ActivatingLicense");
            LicenseActivateButton.IsEnabled = false;

            string signedCode = code;
            if (code.StartsWith("SCB-", StringComparison.OrdinalIgnoreCase) && !code.StartsWith(ServerLicenseCodePrefix + ".", StringComparison.OrdinalIgnoreCase))
                signedCode = await ActivateLicenseOnlineAsync(code);

            if (!TryReadLicenseCode(signedCode, out var license, out string error))
            {
                LicenseMessageText.Text = error;
                LicenseStatusText.Text = _localization.GetString("InvalidLicense");
                return;
            }

            string systemId = _service?.ComputerId ?? Environment.MachineName;
            if (!string.Equals(license.SystemId, systemId, StringComparison.OrdinalIgnoreCase) && !string.Equals(license.SystemId, "*", StringComparison.OrdinalIgnoreCase))
            {
                LicenseMessageText.Text = _localization.GetString("ThisLicenseIsNotIssuedForThisSystem");
                LicenseStatusText.Text = _localization.GetString("WrongSystem");
                return;
            }

            _activeLicense = license;
            _activeLicenseCode = signedCode;
            StartLicenseHeartbeatTimer();
            SaveActiveLicense();
            ApplyLicenseToUi();

            // اگر لایسنس جدید باشد (نه تمدید)، حساب‌های ذخیره‌شده تی‌تک پاک می‌شوند.
            ClearSavedLoginsIfLicenseChanged();

            LicenseMessageText.Text = _localization.GetString("LicenseActivatedSuccessfully");
        }
        catch (Exception ex)
        {
            LicenseMessageText.Text = _localization.GetFormattedString("ActivationFailedFormat", ex.Message);
            LicenseStatusText.Text = _localization.GetString("ActivationFailed2");
        }
        finally
        {
            LicenseActivateButton.IsEnabled = true;
        }
    }

    // CreateDemoLicense/ActivateDemoLicense و دکمه‌های LicenseDemo*Button_Click از اینجا حذف
    // شدند - طبق ممیزی امنیتی، این‌ها یک لایسنس واقعی (از جمله «تی‌تک‌پلاس» یک‌ساله‌ی کامل) با
    // امضای HMAC قدیمی (SCB1) خودشان می‌ساختند و فعال می‌کردند، بدون هیچ تماسی با سرور. مسیر
    // پذیرش SCB1 هم در TryReadLicenseCode غیرفعال شد (نگاه کنید به یادداشت همان‌جا).

    private void LicenseButton_Click(object sender, RoutedEventArgs e)
    {
        ShowLicenseOverlayStrict();
    }

    private void UpdateLicenseOverlayTextsSafe()
    {
        if (LicenseSystemIdText == null)
            return;

        string systemId = _service?.ComputerId ?? Environment.MachineName;
        LicenseSystemIdText.Text = systemId;
        LicenseTitleText.Text = _localization.GetString("ScanbridgeLicense");
        LicenseSystemIdLabel.Text = _localization.GetString("SystemID");
        LicenseCodeLabel.Text = _localization.GetString("LicenseKey");
        LicenseCopySystemIdButton.Content = _localization.GetString("Copy");
        LicenseActivateButton.Content = _localization.GetString("ActivateOnline");
        LicenseImportFileButton.Content = _localization.GetString("ImportFile");
        LicenseExportFileButton.Content = _localization.GetString("ExportID");
        CloseLicenseOverlayButton.Content = _localization.GetString("Close");

        bool english = _localization.CurrentLanguage == AppLanguage.English;
        if (LicenseInfoPlanLabel != null) LicenseInfoPlanLabel.Text = _localization.GetString("Plan");
        if (LicenseInfoPharmacyLabel != null) LicenseInfoPharmacyLabel.Text = _localization.GetString("User");
        if (LicenseInfoExpiryLabel != null) LicenseInfoExpiryLabel.Text = _localization.GetString("Expires");
        if (LicenseInfoLastCheckLabel != null) LicenseInfoLastCheckLabel.Text = _localization.GetString("LastUpdate");
        if (LicenseInfoPhoneLabel != null) LicenseInfoPhoneLabel.Text = _localization.GetString("Phone");

        if (IsLicenseValid())
        {
            LicenseStatusText.Foreground = System.Windows.Media.Brushes.Green;
            LicenseStatusText.Text = _localization.GetString("LicenseIsActive");

            string pharmacy = !string.IsNullOrWhiteSpace(_activeLicense.PharmacyName) ? _activeLicense.PharmacyName : _activeLicense.CustomerName;
            DateTime localLastCheck = _lastLicenseOnlineValidationUtc == DateTime.MinValue
                ? DateTime.MinValue
                : DateTime.SpecifyKind(_lastLicenseOnlineValidationUtc, DateTimeKind.Utc).ToLocalTime();

            if (LicenseInfoTable != null) LicenseInfoTable.Visibility = Visibility.Visible;
            if (LicenseInfoPlanValue != null) LicenseInfoPlanValue.Text = english ? _activeLicense.Plan : _activeLicense.GetPersianPlanName();
            if (LicenseInfoPharmacyValue != null) LicenseInfoPharmacyValue.Text = string.IsNullOrWhiteSpace(pharmacy) ? "-" : pharmacy;
            if (LicenseInfoExpiryValue != null)
                LicenseInfoExpiryValue.Text = _activeLicense.ExpiresAt == DateTime.MinValue
                    ? "-"
                    : (english ? _activeLicense.ExpiresAt.ToString("yyyy/MM/dd HH:mm") : FormatPersianDateTime(_activeLicense.ExpiresAt));
            if (LicenseInfoLastCheckValue != null)
                LicenseInfoLastCheckValue.Text = localLastCheck == DateTime.MinValue
                    ? "-"
                    : (english ? localLastCheck.ToString("yyyy/MM/dd HH:mm") : FormatPersianDateTime(localLastCheck));

            bool hasPhone = !string.IsNullOrWhiteSpace(_activeLicense.CustomerPhone);
            if (LicenseInfoPhoneLabel != null) LicenseInfoPhoneLabel.Visibility = hasPhone ? Visibility.Visible : Visibility.Collapsed;
            if (LicenseInfoPhoneValue != null)
            {
                LicenseInfoPhoneValue.Visibility = hasPhone ? Visibility.Visible : Visibility.Collapsed;
                LicenseInfoPhoneValue.Text = hasPhone ? _activeLicense.CustomerPhone : "-";
            }
        }
        else
        {
            if (LicenseInfoTable != null) LicenseInfoTable.Visibility = Visibility.Collapsed;
            LicenseStatusText.Foreground = System.Windows.Media.Brushes.Firebrick;
            LicenseStatusText.Text = _activeLicense.IsExpired()
                ? (_localization.GetString("LicenseExpired"))
                : (_localization.GetString("NotActivated"));
        }
    }

    private void CloseLicenseOverlaySafe()
    {
        if (!IsLicenseValid())
        {
            UpdateLicenseOverlayTextsSafe();
            return;
        }

        if (FindName("LicenseOverlay") is FrameworkElement overlay)
            overlay.Visibility = Visibility.Collapsed;
        MainContent.Effect = null;
    }

    private void LicenseOverlay_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        CloseLicenseOverlaySafe();
    }

    private void LicenseCard_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void CloseLicenseOverlay_Click(object sender, RoutedEventArgs e)
    {
        CloseLicenseOverlaySafe();
    }

    private void LicenseCopySystemIdButton_Click(object sender, RoutedEventArgs e)
    {
        System.Windows.Clipboard.SetText(_service?.ComputerId ?? Environment.MachineName);
        LicenseMessageText.Text = _localization.GetString("SystemIDCopied");
    }

    private async void LicenseActivateButton_Click(object sender, RoutedEventArgs e)
    {
        await ActivateLicenseFromCodeAsync(LicenseCodeTextBox.Text);
    }

    private async void LicenseImportFileButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "License Files (*.lic;*.json;*.txt)|*.lic;*.json;*.txt|All Files (*.*)|*.*",
                Title = _localization.GetString("ImportLicenseFile")
            };
            if (dialog.ShowDialog() == true)
            {
                LicenseCodeTextBox.Text = File.ReadAllText(dialog.FileName);
                await ActivateLicenseFromCodeAsync(LicenseCodeTextBox.Text);
            }
        }
        catch (Exception ex)
        {
            LicenseMessageText.Text = ex.Message;
        }
    }

    private void LicenseExportFileButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = "Scanbridge-SystemId.txt",
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                Title = _localization.GetString("ExportSystemID")
            };
            if (dialog.ShowDialog() == true)
            {
                File.WriteAllText(dialog.FileName, _service?.ComputerId ?? Environment.MachineName, Encoding.UTF8);
                LicenseMessageText.Text = _localization.GetString("SystemIDExported");
            }
        }
        catch (Exception ex)
        {
            LicenseMessageText.Text = ex.Message;
        }
    }

    private void MainWindow_OnClosing(object sender, CancelEventArgs e)
    {
        // قبلاً اینجا با e.Cancel = true و Hide() برنامه فقط مخفی می‌شد و پشت صحنه (و سرویس
        // اسکن‌بریج/پورت وب‌ساکتش) روشن می‌ماند - یعنی با زدن ضربدر برنامه واقعاً بسته
        // نمی‌شد. کاربر فکر می‌کرد بسته شده، دوباره برنامه را اجرا می‌کرد، و چون نمونه‌ی قبلی
        // هنوز در پس‌زمینه/تری فعال بود، نمونه‌ی جدید یا اصلاً بالا نمی‌آمد یا گوشی به آن
        // نمونه‌ی جدید متصل نمی‌شد (چون سرور واقعی همچنان همان نمونه‌ی قبلی بود).
        // حالا با زدن ضربدر، برنامه واقعاً و به‌طور کامل بسته می‌شود: سرویس اسکن‌بریج و آیکون
        // تری هم در App.OnExit به‌درستی Dispose می‌شوند.
        System.Windows.Application.Current.Shutdown();
    }

}

public class ScanbridgeLicense
{
    public string LicenseId { get; set; } = Guid.NewGuid().ToString("N");
    public string CustomerName { get; set; } = "";
    public string PharmacyName { get; set; } = "";
    public string CustomerPhone { get; set; } = "";
    public string CustomerType { get; set; } = "";
    public string Product { get; set; } = "Scanbridge";
    public string Plan { get; set; } = "Missing";
    public string SystemId { get; set; } = "";
    public DateTime IssuedAt { get; set; } = DateTime.Now;
    public DateTime ExpiresAt { get; set; } = DateTime.MinValue;
    public int MaxDevices { get; set; } = 1;
    public LicensedModules Modules { get; set; } = new();
    public string Signature { get; set; } = "";

    public static ScanbridgeLicense Missing() => new()
    {
        Plan = "Missing",
        ExpiresAt = DateTime.MinValue,
        Modules = new LicensedModules()
    };

    public bool IsExpired() => ExpiresAt != DateTime.MinValue && DateTime.Now.Date > ExpiresAt.Date;

    public bool IsValidForSystem(string systemId)
    {
        if (string.IsNullOrWhiteSpace(Plan) || Plan == "Missing")
            return false;
        if (IsExpired())
            return false;
        if (!string.Equals(SystemId, "*", StringComparison.OrdinalIgnoreCase) && !string.Equals(SystemId, systemId, StringComparison.OrdinalIgnoreCase))
            return false;
        return Modules.BarcodeBridge;
    }

    public string GetPersianPlanName() => Plan switch
    {
        "Normal" => "پایه",
        "Ttac" => "تی‌تک",
        "TtacPlus" => "حرفه‌ای",
        "Trial" => "آزمایشی",
        _ => Plan
    };

    public static ScanbridgeLicense ForPlan(string plan, string systemId, DateTime expiresAt)
    {
        var modules = LicensedModules.ForPlan(plan);
        return new ScanbridgeLicense
        {
            Plan = plan,
            SystemId = systemId,
            ExpiresAt = expiresAt,
            CustomerType = plan.StartsWith("Ttac", StringComparison.OrdinalIgnoreCase) ? "Pharmacy" : "General",
            Product = plan switch
            {
                "Normal" => "Scanbridge Barcode",
                "Ttac" => "Scanbridge Ttac",
                "TtacPlus" => "Scanbridge Ttac Plus",
                _ => "Scanbridge Trial"
            },
            Modules = modules
        };
    }
}

public class LicensedModules
{
    public bool BarcodeBridge { get; set; }
    public bool History { get; set; }
    public bool ExcelPdf { get; set; }
    public bool DeviceManagement { get; set; }
    public bool Ttac { get; set; }
    public bool TtacRegistration { get; set; }
    public bool Formula { get; set; }
    public bool ReceiveStatus { get; set; }
    public bool CargoDelivery { get; set; }
    public bool MonthlyArchive { get; set; }
    public bool PharmacyMemory { get; set; }

    public bool IsEnabled(string name) => name switch
    {
        "barcodeBridge" => BarcodeBridge,
        "history" => History,
        "excelPdf" => ExcelPdf,
        "deviceManagement" => DeviceManagement,
        "ttac" => Ttac,
        "ttacRegistration" => TtacRegistration,
        "formula" => Formula,
        "receiveStatus" => ReceiveStatus,
        "cargoDelivery" => CargoDelivery,
        "monthlyArchive" => MonthlyArchive,
        "pharmacyMemory" => PharmacyMemory,
        _ => false
    };

    public static LicensedModules ForPlan(string plan)
    {
        var m = new LicensedModules();
        switch (plan)
        {
            case "Trial":
                m.BarcodeBridge = m.History = m.ExcelPdf = m.DeviceManagement = true;
                m.Ttac = m.TtacRegistration = m.Formula = m.ReceiveStatus = m.CargoDelivery = m.MonthlyArchive = m.PharmacyMemory = true;
                break;
            case "Normal":
                m.BarcodeBridge = m.History = m.DeviceManagement = true;
                break;
            case "Ttac":
                m.BarcodeBridge = m.History = m.ExcelPdf = m.DeviceManagement = true;
                m.Ttac = m.TtacRegistration = m.Formula = m.MonthlyArchive = m.PharmacyMemory = true;
                break;
            case "TtacPlus":
                m.BarcodeBridge = m.History = m.ExcelPdf = m.DeviceManagement = true;
                m.Ttac = m.TtacRegistration = m.Formula = m.ReceiveStatus = m.CargoDelivery = m.MonthlyArchive = m.PharmacyMemory = true;
                break;
        }
        return m;
    }
}

public class HistoryDisplayRow
{
    public HistoryDisplayRow(int rowNumber, ScanRecord record, string deviceName)
    {
        RowNumber = rowNumber;
        Record = record;
        DeviceName = deviceName;
    }

    public int RowNumber { get; }
    public ScanRecord Record { get; }
    public string Barcode => Record.Barcode;
    public string TimeText => Record.TimeText;
    public string PersianDateText => Record.PersianDateText;
    public string DeviceName { get; }
    public string DrugName => Record.DrugName;
    public DateTime TimestampLocal => Record.TimestampLocal;
}

public class TtTeckHistoryRow
{
    public int RowNumber { get; set; }
    public bool IsRegistered { get; set; }
    public bool IsInfantFormula { get; set; }
    public string RegistrationButtonText { get; set; } = "ثبت در تی‌تک";
    public string ProductDisplayName { get; set; } = "";
    public System.Windows.Media.Brush RegistrationButtonBackground { get; set; } = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0));
    public string PersianProductName { get; set; } = "";
    public string EnglishProductName { get; set; } = "";
    public string Barcode { get; set; } = "";
    public string TimeText { get; set; } = "";
    public string PersianDateText { get; set; } = "";
    public DateTime TimestampLocal { get; set; }
    public string DeviceName { get; set; } = "";
    public string StatusText { get; set; } = "";
    public string RetryReason { get; set; } = "";
    public Visibility RetryButtonVisibility { get; set; } = Visibility.Collapsed;
}

public class HistoryReportRow
{
    public string Date { get; set; } = "";
    public string Time { get; set; } = "";
    public string Barcode { get; set; } = "";
    public string DeviceName { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string EnglishProductName { get; set; } = "";
}


public class ProductDetailField
{
    public string Label { get; set; } = "";
    public string Value { get; set; } = "";
}


public class DeviceRowDisplayViewModel
{
    public string OriginalDeviceName { get; set; } = "";
    public string DeviceName { get; set; } = "";
    public string LinkBadge { get; set; } = "";
    public System.Windows.Media.Brush StatusColor { get; set; } = System.Windows.Media.Brushes.Transparent;
}


public class TtacRegistrationLogRow
{
    public string Title { get; set; } = "";
    public string Message { get; set; } = "";
    public System.Windows.Media.Brush StatusBrush { get; set; } = System.Windows.Media.Brushes.Green;
}


public class ReceiveStatusRow
{
    public int RowNumber { get; set; }
    public long ReceiveId { get; set; }
    public bool IsConfirmable { get; set; } = true;
    public string Barcode { get; set; } = "";
    public string Irc { get; set; } = "";
    public string UID { get; set; } = "";
    public string GTIN { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string ProductEnName { get; set; } = "";
    public string GenericCode { get; set; } = "";
    public string GenericName { get; set; } = "";
    public string LotNumber { get; set; } = "";
    public string Expiration { get; set; } = "";
    public string SenderName { get; set; } = "";
    public string Quantity { get; set; } = "";
    public string SentDatePersian { get; set; } = "";
    public string StatusText { get; set; } = "";
    public System.Windows.Media.Brush StatusBrush { get; set; } = System.Windows.Media.Brushes.DodgerBlue;
    public string DetailText => $"IRC: {Irc} | سری: {LotNumber} | تعداد: {Quantity} | ارسال‌کننده: {SenderName} | تاریخ ارسال: {SentDatePersian}";
}

public class CargoDeliveryStorageRow
{
    public long ReceiveId { get; set; }
    public bool IsSelected { get; set; }
    public bool IsActionEnabled { get; set; }
    public string ActionButtonText { get; set; } = "";
    public string ActionKind { get; set; } = "Error";
    public string Barcode { get; set; } = "";
    public string Irc { get; set; } = "";
    public string UID { get; set; } = "";
    public string GTIN { get; set; } = "";
    public string GenericCode { get; set; } = "";
    public string GenericName { get; set; } = "";
    public string Expiration { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string ProductEnName { get; set; } = "";
    public string LotNumber { get; set; } = "";
    public string Quantity { get; set; } = "";
    public string SenderName { get; set; } = "";
    public string SentDatePersian { get; set; } = "";
    public string StatusText { get; set; } = "";
    public string StatusKind { get; set; } = "Info";

    public static CargoDeliveryStorageRow FromRow(CargoDeliveryRow row)
    {
        string statusKind = row.StatusBrush == System.Windows.Media.Brushes.Green ? "Success" : row.StatusBrush == System.Windows.Media.Brushes.Firebrick ? "Error" : row.IsActionEnabled ? "Info" : "Warning";
        string actionKind = row.ActionButtonBrush == System.Windows.Media.Brushes.Green ? "Success" : row.ActionButtonBrush == System.Windows.Media.Brushes.Gray ? "Disabled" : row.IsActionEnabled ? "Error" : "Warning";
        return new CargoDeliveryStorageRow
        {
            ReceiveId = row.ReceiveId,
            IsSelected = row.IsSelected,
            IsActionEnabled = row.IsActionEnabled,
            ActionButtonText = row.ActionButtonText,
            ActionKind = actionKind,
            Barcode = row.Barcode,
            Irc = row.Irc,
            UID = row.UID,
            GTIN = row.GTIN,
            GenericCode = row.GenericCode,
            GenericName = row.GenericName,
            Expiration = row.Expiration,
            ProductName = row.ProductName,
            ProductEnName = row.ProductEnName,
            LotNumber = row.LotNumber,
            Quantity = row.Quantity,
            SenderName = row.SenderName,
            SentDatePersian = row.SentDatePersian,
            StatusText = row.StatusText,
            StatusKind = statusKind
        };
    }

    public CargoDeliveryRow ToRow(bool english = false)
    {
        var statusBrush = StatusKind switch
        {
            "Success" => System.Windows.Media.Brushes.Green,
            "Error" => System.Windows.Media.Brushes.Firebrick,
            "Warning" => new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17)),
            _ => System.Windows.Media.Brushes.DodgerBlue
        };
        var actionBrush = ActionKind switch
        {
            "Success" => System.Windows.Media.Brushes.Green,
            "Disabled" => System.Windows.Media.Brushes.Gray,
            "Warning" => new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17)),
            _ => System.Windows.Media.Brushes.Firebrick
        };
        return new CargoDeliveryRow
        {
            English = english,
            ReceiveId = ReceiveId,
            IsSelected = IsSelected,
            IsActionEnabled = IsActionEnabled,
            ActionButtonText = string.IsNullOrWhiteSpace(ActionButtonText) ? (LocalizationManager.Instance.GetString("AddToPharmacy")) : ActionButtonText,
            ActionButtonBrush = actionBrush,
            Barcode = Barcode,
            Irc = Irc,
            UID = UID,
            GTIN = GTIN,
            GenericCode = GenericCode,
            GenericName = GenericName,
            Expiration = Expiration,
            ProductName = ProductName,
            ProductEnName = ProductEnName,
            LotNumber = LotNumber,
            Quantity = Quantity,
            SenderName = SenderName,
            SentDatePersian = SentDatePersian,
            StatusText = StatusText,
            StatusBrush = statusBrush
        };
    }
}

public class CargoDeliveryRow
{
    public int RowNumber { get; set; }
    public long ReceiveId { get; set; }
    public bool IsSelected { get; set; }
    public bool IsActionEnabled { get; set; }
    public string ActionButtonText { get; set; } = "افزودن به داروخانه";
    public System.Windows.Media.Brush ActionButtonBrush { get; set; } = System.Windows.Media.Brushes.Firebrick;
    public string Barcode { get; set; } = "";
    public string Irc { get; set; } = "";
    public string UID { get; set; } = "";
    public string GTIN { get; set; } = "";
    public string GenericCode { get; set; } = "";
    public string GenericName { get; set; } = "";
    public string Expiration { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string ProductEnName { get; set; } = "";
    public string LotNumber { get; set; } = "";
    public string Quantity { get; set; } = "";
    public string SenderName { get; set; } = "";
    public string SentDatePersian { get; set; } = "";
    public string StatusText { get; set; } = "";
    public System.Windows.Media.Brush StatusBrush { get; set; } = System.Windows.Media.Brushes.DodgerBlue;
    public bool English { get; set; }
    public string QuantityDisplay => string.IsNullOrWhiteSpace(Quantity) || Quantity == "-" ? "" : $"{LocalizationManager.Instance.GetString("Qty")}: {Quantity}";
    public string DetailText => LocalizationManager.Instance.GetFormattedString("DetailTextFormat", Irc, LotNumber, Quantity, SenderName, SentDatePersian);
    public string DeleteButtonText => LocalizationManager.Instance.GetString("Delete");

    public static CargoDeliveryRow FromReceiveStatus(ReceiveStatusRow source, bool english)
    {
        var row = new CargoDeliveryRow
        {
            English = english,
            ReceiveId = source.ReceiveId,
            IsSelected = source.IsConfirmable,
            IsActionEnabled = source.IsConfirmable && source.ReceiveId > 0,
            Barcode = source.Barcode,
            Irc = source.Irc,
            UID = source.UID,
            GTIN = source.GTIN,
            GenericCode = source.GenericCode,
            GenericName = source.GenericName,
            Expiration = source.Expiration,
            ProductName = source.ProductName,
            ProductEnName = source.ProductEnName,
            LotNumber = source.LotNumber,
            Quantity = source.Quantity,
            SenderName = source.SenderName,
            SentDatePersian = source.SentDatePersian,
            StatusText = source.StatusText,
            StatusBrush = source.StatusBrush
        };

        bool isProbable = source.StatusText.Contains("احتمالاً", StringComparison.OrdinalIgnoreCase)
                          || source.StatusText.Contains("Probably", StringComparison.OrdinalIgnoreCase);
        if (!isProbable && (source.StatusBrush == System.Windows.Media.Brushes.Green || source.StatusText.Contains("قبلاً", StringComparison.OrdinalIgnoreCase) || source.StatusText.Contains("Already", StringComparison.OrdinalIgnoreCase)))
        {
            row.MarkAdded(english, source.StatusText);
        }
        else if (isProbable)
        {
            row.ActionButtonText = LocalizationManager.Instance.GetString("NeedsReview");
            row.ActionButtonBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17));
            row.IsActionEnabled = false;
            row.StatusText = source.StatusText;
            row.StatusBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17));
        }
        else if (row.IsActionEnabled)
        {
            row.ActionButtonText = LocalizationManager.Instance.GetString("AddToPharmacy");
            row.ActionButtonBrush = System.Windows.Media.Brushes.Firebrick;
            row.StatusText = LocalizationManager.Instance.GetString("ReadyToAdd");
            row.StatusBrush = System.Windows.Media.Brushes.DodgerBlue;
        }
        else
        {
            row.ActionButtonText = LocalizationManager.Instance.GetString("NotAvailable");
            row.ActionButtonBrush = System.Windows.Media.Brushes.Gray;
            row.IsActionEnabled = false;
        }
        return row;
    }

    public void MarkAdded(bool english, string? message = null)
    {
        English = english;
        IsActionEnabled = false;
        IsSelected = false;
        ActionButtonText = LocalizationManager.Instance.GetString("InPharmacy");
        ActionButtonBrush = System.Windows.Media.Brushes.Green;
        StatusText = string.IsNullOrWhiteSpace(message) ? (LocalizationManager.Instance.GetString("AddedToPharmacy")) : message;
        StatusBrush = System.Windows.Media.Brushes.Green;
    }
}

public class MonthlyArchiveState
{
    public string LastSeenMonth { get; set; } = "";
    public string LastArchivedMonth { get; set; } = "";
    public string LastDismissedDate { get; set; } = "";
}

public class ReceiveStatusStorageRow
{
    public long ReceiveId { get; set; }
    public bool IsConfirmable { get; set; }
    public string Barcode { get; set; } = "";
    public string Irc { get; set; } = "";
    public string UID { get; set; } = "";
    public string GTIN { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string ProductEnName { get; set; } = "";
    public string GenericCode { get; set; } = "";
    public string GenericName { get; set; } = "";
    public string LotNumber { get; set; } = "";
    public string Expiration { get; set; } = "";
    public string SenderName { get; set; } = "";
    public string Quantity { get; set; } = "";
    public string SentDatePersian { get; set; } = "";
    public string StatusText { get; set; } = "";
    public string StatusKind { get; set; } = "Info";

    public static ReceiveStatusStorageRow FromRow(ReceiveStatusRow row)
    {
        string kind = row.StatusBrush == System.Windows.Media.Brushes.Green
            ? "Success"
            : row.StatusBrush == System.Windows.Media.Brushes.Firebrick
                ? "Error"
                : row.IsConfirmable ? "Info" : "Warning";

        return new ReceiveStatusStorageRow
        {
            ReceiveId = row.ReceiveId,
            IsConfirmable = row.IsConfirmable,
            Barcode = row.Barcode,
            Irc = row.Irc,
            UID = row.UID,
            GTIN = row.GTIN,
            ProductName = row.ProductName,
            ProductEnName = row.ProductEnName,
            GenericCode = row.GenericCode,
            GenericName = row.GenericName,
            LotNumber = row.LotNumber,
            Expiration = row.Expiration,
            SenderName = row.SenderName,
            Quantity = row.Quantity,
            SentDatePersian = row.SentDatePersian,
            StatusText = row.StatusText,
            StatusKind = kind
        };
    }

    public ReceiveStatusRow ToRow()
    {
        var brush = StatusKind switch
        {
            "Success" => System.Windows.Media.Brushes.Green,
            "Error" => System.Windows.Media.Brushes.Firebrick,
            "Warning" => new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF5, 0x7F, 0x17)),
            _ => new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0))
        };

        return new ReceiveStatusRow
        {
            ReceiveId = ReceiveId,
            IsConfirmable = IsConfirmable,
            Barcode = Barcode,
            Irc = Irc,
            UID = UID,
            GTIN = GTIN,
            ProductName = ProductName,
            ProductEnName = ProductEnName,
            GenericCode = GenericCode,
            GenericName = GenericName,
            LotNumber = LotNumber,
            Expiration = Expiration,
            SenderName = SenderName,
            Quantity = Quantity,
            SentDatePersian = SentDatePersian,
            StatusText = StatusText,
            StatusBrush = brush
        };
    }
}

public class TtacRegistrationHistoryEntry
{
    public DateTime RegisteredAt { get; set; }
    public string Barcode { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string RegistrationType { get; set; } = "";
    public long? PrescriptionId { get; set; }
    public string Amount { get; set; } = "";
    public string NationalIdFull { get; set; } = "";
    public string MobileFull { get; set; } = "";
    public string NationalIdMasked { get; set; } = "";
    public string MobileMasked { get; set; } = "";
    public string PatientFullName { get; set; } = "";
    public bool Success { get; set; }
    public string Message { get; set; } = "";
}

// =============================================================================================
// هشدار تاریخ انقضای نزدیک - مدل داده و پنجره‌های مربوطه (ویژگی جدید)
// =============================================================================================

public enum ExpiryWatchStatus
{
    Watching,
    Sold
}

public class ExpiryWatchItem
{
    public string Barcode { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string ProductEnName { get; set; } = "";
    public string BatchCode { get; set; } = "";
    public string Quantity { get; set; } = "";
    public string ExpirationRaw { get; set; } = "";
    public DateTime ExpirationDate { get; set; }
    public ExpiryWatchStatus Status { get; set; } = ExpiryWatchStatus.Watching;
    public DateTime NextAlertDueUtc { get; set; } = DateTime.UtcNow;
    public DateTime? LastAlertedUtc { get; set; }
    public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;
    public DateTime? SoldAtUtc { get; set; }
    // آیا این قلم هنوز منتظر پاسخ کاربر است (نه «فروخته شد» زده شده نه «حواسم هست»)؟ این پرچم
    // مستقل از NextAlertDueUtc است - آن فقط زمان چرخه‌ی خودکار بعدی را کنترل می‌کند، ولی همین
    // پرچم است که نشان قرمز روی دکمه‌ی «تاریخ نزدیک» و پین‌شدن در لیست را کنترل می‌کند. پیش‌فرض
    // true است تا اقلام قدیمی (از قبل از افزوده‌شدن این پرچم) هم بلافاصله دیده شوند.
    public bool NeedsResponse { get; set; } = true;

    public string PersianExpirationText => ExpirationDate == default
        ? ""
        : ExpirationDate.ToString("yyyy/MM/dd", new CultureInfo("fa-IR"));
}

public class ExpiryAlertSettings
{
    public int ThresholdMonths { get; set; } = 6;
    // دوباره بعد از چند روز یادآوری کند (وقتی کاربر «فروخته شد» نزده باشد).
    public int RepeatReminderDays { get; set; } = 30;
    // توکن ربات دیگر اینجا ذخیره نمی‌شود - یک ربات مشترک برای کل برنامه است (SharedBaleBotToken).
    // فقط چت‌آیدی مخصوص همین داروخانه اینجا ذخیره می‌شود که با دکمه‌ی «فعال‌سازی یادآور بله» به‌دست می‌آید.
    public string BaleChatId { get; set; } = "";

    // آیا ارسال هشدارها به بله فعال است؟ این جدا از BaleChatId است: کاربر می‌تواند بدون از دست دادن
    // اتصال بله (بدون نیاز به فعال‌سازی دوباره از صفر)، فقط ارسال پیام‌ها را موقتاً خاموش/روشن کند.
    public bool IsBaleNotificationsEnabled { get; set; } = true;

    // آخرین زمان (میلی‌ثانیه یونیکس UTC) که تنظیمات قابل‌همگام‌سازی بین سیستم‌ها (فعال/غیرفعال
    // بودن تی‌تک، آستانه‌های تاریخ نزدیک، فعال/غیرفعال بودن هشدار بله) روی همین سیستم تغییر کرد یا
    // از یک سیستم هم‌شبکه با همان لایسنس دریافت شد. صفر یعنی هنوز هیچ تغییر/همگام‌سازی‌ای رخ نداده.
    public long DesktopSettingsSyncVersionUtcMs { get; set; } = 0;

    [JsonIgnore]
    public bool IsBaleConfigured => !string.IsNullOrWhiteSpace(BaleChatId);
}

/// <summary>
/// پنجره‌ی هشدار «تاریخ نزدیک» - کاملاً با کد ساخته می‌شود (بدون XAML جداگانه) تا وابسته به
/// نام کنترل‌های موجود در MainWindow.xaml نباشد و همیشه کامپایل شود.
/// </summary>
public class ExpiryAlertWindow : Window
{
    public event Action<string>? ItemMarkedSold;
    public event Action<string>? ItemAcknowledged;
    private readonly StackPanel _listPanel;
    private readonly bool _english;

    public ExpiryAlertWindow(List<ExpiryWatchItem> items, bool english = false)
    {
        _english = english;

        // این پنجره فقط باید داخل خودِ برنامه (روی MainWindow، که با Owner ست می‌شود) دیده شود، نه
        // روی همه‌ی پنجره‌های ویندوز. قبلاً Topmost=true باعث می‌شد حتی وقتی کاربر می‌رفت سراغ یک
        // برنامه‌ی دیگر، این هشدار همچنان روی آن هم نمایش داده شود. Owner به‌تنهایی کافی است تا
        // این پنجره همیشه بالای MainWindow بماند - بدون این‌که سراسر ویندوز را اشغال کند.
        const double cardWidth = 460;
        const double shadowMargin = 40; // فضای لازم دور کارت برای این‌که سایه (BlurRadius+ShadowDepth) کامل و بدون بریدگی رندر شود

        Title = LocalizationManager.Instance.GetString("NearExpiryAlert2");
        Width = cardWidth + shadowMargin * 2;
        SizeToContent = System.Windows.SizeToContent.Height;
        MaxHeight = 640 + shadowMargin * 2;
        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
        FlowDirection = english ? System.Windows.FlowDirection.LeftToRight : System.Windows.FlowDirection.RightToLeft;
        WindowStyle = System.Windows.WindowStyle.None;
        AllowsTransparency = true;
        Background = System.Windows.Media.Brushes.Transparent;
        ResizeMode = System.Windows.ResizeMode.NoResize;
        Topmost = false;

        var windowCard = new Border
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(24),
            Padding = new Thickness(20),
            // بدون این مارجین، سایه‌ی زیر درست روی لبه‌ی پنجره‌ی AllowsTransparency می‌افتد و به‌جای
            // یک سایه‌ی نرم، به شکل یک هاله‌ی خاکستری مستطیلی و بریده دیده می‌شود؛ این فاصله اجازه
            // می‌دهد سایه کامل و به‌صورت طبیعی محو شود.
            Margin = new Thickness(shadowMargin),
            Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                Color = System.Windows.Media.Colors.Black,
                Opacity = 0.28,
                BlurRadius = 26,
                ShadowDepth = 6
            }
        };

        var root = new DockPanel();

        var headerRow = new Grid { Margin = new Thickness(0, 0, 0, 12) };
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        headerRow.MouseLeftButtonDown += (_, _) => { try { DragMove(); } catch { } };

        var header = new TextBlock
        {
            Text = LocalizationManager.Instance.GetFormattedString("NearExpiryHeader", items.Count),
            FontSize = 16,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x1A, 0x23, 0x7E)),
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        System.Windows.Controls.Grid.SetColumn(header, 0);
        headerRow.Children.Add(header);

        var closeButton = CreateRoundedButton("✕", System.Windows.Media.Color.FromRgb(0xF3, 0xF4, 0xF6), System.Windows.Media.Color.FromRgb(0x37, 0x41, 0x51), 30, 30, 12);
        closeButton.Click += (_, _) => Close();
        System.Windows.Controls.Grid.SetColumn(closeButton, 1);
        headerRow.Children.Add(closeButton);

        DockPanel.SetDock(headerRow, Dock.Top);
        root.Children.Add(headerRow);

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, MaxHeight = 480 };
        _listPanel = new StackPanel();
        scroll.Content = _listPanel;
        root.Children.Add(scroll);

        windowCard.Child = root;
        Content = windowCard;

        foreach (var item in items)
            AddRow(item);
    }

    /// <summary>
    /// دکمه‌ی گردگوشه با همون منطق RoundedButtonStyle خودِ اپ (چون این پنجره XAML جدا نداره و
    /// نمی‌تواند به Window.Resources اصلی دسترسی داشته باشد).
    /// </summary>
    private static System.Windows.Controls.Button CreateRoundedButton(string content, System.Windows.Media.Color background, System.Windows.Media.Color foreground, double width, double height, double fontSize)
    {
        var button = new System.Windows.Controls.Button
        {
            Content = content,
            Width = width,
            Height = height,
            FontSize = fontSize,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(foreground),
            Background = new SolidColorBrush(background),
            BorderThickness = new Thickness(0),
            Cursor = System.Windows.Input.Cursors.Hand,
            HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        };

        var template = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
        var borderFactory = new System.Windows.FrameworkElementFactory(typeof(Border));
        borderFactory.Name = "Bd";
        borderFactory.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background")
        {
            RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent)
        });
        borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(Math.Min(14, Math.Min(width, height) / 2)));
        borderFactory.SetValue(Border.SnapsToDevicePixelsProperty, true);

        var contentFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
        contentFactory.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        contentFactory.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        borderFactory.AppendChild(contentFactory);
        template.VisualTree = borderFactory;

        var hoverTrigger = new System.Windows.Trigger { Property = System.Windows.Controls.Button.IsMouseOverProperty, Value = true };
        hoverTrigger.Setters.Add(new System.Windows.Setter(Border.OpacityProperty, 0.88) { TargetName = "Bd" });
        template.Triggers.Add(hoverTrigger);

        var pressTrigger = new System.Windows.Trigger { Property = System.Windows.Controls.Button.IsPressedProperty, Value = true };
        pressTrigger.Setters.Add(new System.Windows.Setter(Border.OpacityProperty, 0.75) { TargetName = "Bd" });
        template.Triggers.Add(pressTrigger);

        button.Template = template;
        return button;
    }

    private void AddRow(ExpiryWatchItem item)
    {
        var card = new Border
        {
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xF9, 0xFA, 0xFB)),
            BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xD1, 0xFA, 0xE5)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(14),
            Margin = new Thickness(0, 0, 0, 10)
        };

        string name = string.IsNullOrWhiteSpace(item.ProductName)
            ? (string.IsNullOrWhiteSpace(item.ProductEnName) ? item.Barcode : item.ProductEnName)
            : item.ProductName;

        var stack = new StackPanel();
        stack.Children.Add(new TextBlock { Text = name, FontWeight = FontWeights.SemiBold, FontSize = 14, TextWrapping = TextWrapping.Wrap });
        string batchCode = string.IsNullOrWhiteSpace(item.BatchCode) ? "-" : item.BatchCode;
        stack.Children.Add(new TextBlock
        {
            Text = LocalizationManager.Instance.GetFormattedString("LotExpiryWide", batchCode, item.PersianExpirationText),
            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x6B, 0x72, 0x80)),
            FontSize = 12,
            Margin = new Thickness(0, 4, 0, 10)
        });

        var buttonsPanel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Left };

        var soldButton = CreateRoundedButton(LocalizationManager.Instance.GetString("Sold"), System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81), System.Windows.Media.Colors.White, 100, 36, 12);
        soldButton.Margin = new Thickness(0, 0, 8, 0);
        soldButton.Click += (_, _) =>
        {
            ItemMarkedSold?.Invoke(item.Barcode);
            _listPanel.Children.Remove(card);
            CloseIfEmpty();
        };

        var ackButton = CreateRoundedButton(LocalizationManager.Instance.GetString("GotItThanks"), System.Windows.Media.Color.FromRgb(0xF3, 0xF4, 0xF6), System.Windows.Media.Color.FromRgb(0x37, 0x41, 0x51), 190, 36, 12);
        ackButton.Click += (_, _) =>
        {
            ItemAcknowledged?.Invoke(item.Barcode);
            _listPanel.Children.Remove(card);
            CloseIfEmpty();
        };

        buttonsPanel.Children.Add(soldButton);
        buttonsPanel.Children.Add(ackButton);
        stack.Children.Add(buttonsPanel);
        card.Child = stack;
        _listPanel.Children.Add(card);
    }

    private void CloseIfEmpty()
    {
        if (_listPanel.Children.Count == 0)
            Close();
    }
}

public class ExpiryWatchDisplayRow
{
    public int RowNumber { get; set; }
    public string Barcode { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string BatchCode { get; set; } = "";
    public string ExpirationText { get; set; } = "";
    public string DetailText { get; set; } = "";
    public string StatusText { get; set; } = "";
    public System.Windows.Media.Brush StatusBrush { get; set; } = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x92, 0x40, 0x0E));
    public bool IsPinned { get; set; }
    public string SoldButtonText { get; set; } = "";
    public string AckButtonText { get; set; } = "";
    public System.Windows.Media.Brush CardBackground { get; set; } = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xFB, 0xEB));
    public System.Windows.Media.Brush CardBorder { get; set; } = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFD, 0xE6, 0x8A));
}

// ردیف نمایشی برای لیست «حساب‌های ذخیره‌شده‌ی تی‌تک» در تنظیمات - عمداً هیچ‌وقت رمز عبور را در
// خودش نگه نمی‌دارد (فقط یوزرنیم و یک متن توضیحی)، چون این کلاس مستقیماً به UI بایند می‌شود و
public sealed class TtacSavedLoginDisplayRow
{
    public string Username { get; set; } = string.Empty;
    public string PharmacyName { get; set; } = string.Empty;
    public string DisplayLabel => string.IsNullOrWhiteSpace(PharmacyName) ? Username : PharmacyName;
}
