using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Fleck;
using QRCoder;

namespace ScanBridgeTest;

public enum ConnectionState
{
    Offline,
    Ready,
    Busy
}

public sealed class ConnectedDeviceInfo
{
    public ConnectedDeviceInfo(string deviceName, bool hasScanned)
    {
        DeviceName = deviceName;
        HasScanned = hasScanned;
    }

    public string DeviceName { get; }
    public bool HasScanned { get; }
}

public sealed class ConnectedDevicesChangedEventArgs : EventArgs
{
    public ConnectedDevicesChangedEventArgs(IReadOnlyList<ConnectedDeviceInfo> devices)
    {
        Devices = devices;
    }

    public IReadOnlyList<ConnectedDeviceInfo> Devices { get; }
}

public sealed class ScanBridgeService : IDisposable
{
    public const int Port = 5050;
    // پورت جداگانه‌ای که فقط بین خودِ سیستم‌های دسکتاپ (نه گوشی‌ها) برای هماهنگ‌سازی تنظیمات
    // استفاده می‌شود - عمداً از سرور اصلی گوشی‌ها (Port) جدا است تا هیچ اتصال بین دو دسکتاپ در
    // لیست «دستگاه‌های وصل‌شده»ی گوشی‌ها ظاهر نشود.
    public const int PeerSyncPort = Port + 1;
    // پورت UDP broadcast برای پیدا کردن سیستم‌های دیگر روی همین شبکه‌ی محلی که همان لایسنس را دارند.
    private const int DiscoveryPort = 45059;
    private const string ScanFileName = "scans.csv";
    private readonly object _scanFileLock = new();
    private readonly object _connectionLock = new();
    private readonly System.Timers.Timer _pruneTimer = new(TimeSpan.FromHours(1).TotalMilliseconds);
    private readonly System.Timers.Timer _connectionHealthTimer = new(TimeSpan.FromSeconds(20).TotalMilliseconds);
    // هر چند ثانیه یک‌بار IP شبکه را چک می‌کند تا اگر عوض شد، QR اتصال خودکار دوباره ساخته شود.
    private readonly System.Timers.Timer _lanIpWatchTimer = new(TimeSpan.FromSeconds(5).TotalMilliseconds);
    private readonly BlockingCollection<(IWebSocketConnection Socket, string Barcode, string DeviceName)> _keyboardQueue = new();
    private readonly ConcurrentDictionary<IWebSocketConnection, bool> _erroredConnections = new();
    private readonly ConcurrentDictionary<IWebSocketConnection, DeviceState> _connectedDevices = new();
    // ack اسکن (صف پردازش)، پینگ سلامت (تایمر جدا)، و broadcastهای هشدار/ورود از راه دور همگی
    // می‌توانند هم‌زمان بخواهند روی یک اتصال WebSocket بنویسند - چون IWebSocketConnection.Send
    // خودش async است (Task برمی‌گرداند، نوشتن واقعی روی سوکت زیرین کامل نمی‌شود که برگردد)، دو
    // Send هم‌زمان روی یک اتصال می‌توانستند بایت‌هایشان روی سیم قاطی شوند و فریم WebSocket را
    // خراب کنند (باگ ۱۳ گزارش ممیزی). یک SemaphoreSlim به‌ازای هر سوکت (نگاه کنید به SafeSend)
    // مطمئن می‌شود در هر لحظه فقط یک Send واقعاً در حال تکمیل‌شدن است.
    private readonly ConcurrentDictionary<IWebSocketConnection, SemaphoreSlim> _socketSendGates = new();
    private readonly ConcurrentDictionary<string, DateTime> _manuallyBlockedDevices = new(StringComparer.OrdinalIgnoreCase);
    // یک دستگاه بلاک‌شده می‌تواند با اتصال مجدد و فرستادن بارکد به‌صورت متن‌ساده (پروتکل قدیمی،
    // بدون فیلد deviceName در JSON) بلاک روی نام را دور بزند - چون تشخیص بلاک فقط با deviceName
    // پیام فعلی مقایسه می‌شود که در این حالت همیشه خالی است (باگ ۱۱ گزارش ممیزی). به‌عنوان یک
    // لایه‌ی دفاعی مستقل از deviceName، IP اتصال هم هنگام بلاک‌کردن ذخیره می‌شود.
    private readonly ConcurrentDictionary<string, DateTime> _manuallyBlockedIps = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, List<string>> _blockedDeviceNameToIps = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, PeerInfo> _knownPeers = new(StringComparer.OrdinalIgnoreCase);
    private WebSocketServer? _server;
    private WebSocketServer? _peerServer;
    private UdpClient? _discoveryUdp;
    private CancellationTokenSource? _discoveryCts;
    private Task? _queueTask;
    private CancellationTokenSource? _queueCts;
    private string _licenseGroupKey = string.Empty;
    private string _lastDesktopSettingsJson = string.Empty;
    private long _lastDesktopSettingsVersionUtcMs;
    // عکسِ کامل آخرین بانک بارکد پرمصرف - فقط برای هماهنگ‌سازی اولیه‌ی یک سیستم تازه‌وصل‌شده
    // (bootstrap) استفاده می‌شود؛ به‌روزرسانی زنده/آنی از طریق عملیات‌های تکی رد و بدل می‌شود.
    private string _lastHighUsageSnapshotJson = string.Empty;
    private long _lastHighUsageSnapshotVersionUtcMs;

    public event EventHandler<ScanReceivedEventArgs>? ScanReceived;
    public event EventHandler<ConnectionStateChangedEventArgs>? ConnectionStatusChanged;
    public event EventHandler<ConnectedDevicesChangedEventArgs>? ConnectedDevicesChanged;
    public event EventHandler? UnexpectedDisconnection;
    // وقتی یک سیستم دیگر روی همین شبکه (با همان لایسنس) تنظیمات TtTeck/تاریخ نزدیک/بله را عوض
    // کرده و برای این سیستم فرستاده باشد - یا در پاسخ به درخواست ما - این رویداد فایر می‌شود.
    public event EventHandler<PeerDesktopSettingsEventArgs>? PeerDesktopSettingsReceived;
    // هماهنگ‌سازی بانک بارکد پرمصرف بین سیستم‌های هم‌شبکه‌ی هم‌لایسنس: یک رویداد برای عملیات‌های
    // تکی و آنی (اضافه/حذف/دیسپنس یک بارکد - همان لحظه که روی سیستم دیگر اتفاق افتاده)، و یک رویداد
    // برای عکسِ کامل بانک (فقط موقع هماهنگ‌سازی اولیه‌ی یک سیستم تازه‌وصل‌شده یا تازه‌روشن‌شده).
    public event EventHandler<HighUsageBarcodeOperationEventArgs>? HighUsageBarcodeOperationReceived;
    public event EventHandler<HighUsageBarcodeSnapshotEventArgs>? HighUsageBarcodeSnapshotReceived;
    // وقتی IP شبکه‌ی محلی عوض می‌شود (مثلاً از وای‌فای به کابل یا شبکه‌ی دیگر) فایر می‌شود تا
    // برنامه QR اتصال را خودکار دوباره بسازد و IP نمایش‌داده‌شده را به‌روز کند.
    public event EventHandler? LanIpChanged;
    // ویژگی «ورود اطلاعات از راه دور» (فرم ثبت شیرخشک روی گوشی): وقتی کاربر روی گوشی یک مقدار
    // (کد ملی، تاریخ تولد و ...) وارد کرده و «بعدی» را زده، یا وقتی دکمه‌ی نهایی «ثبت» را زده.
    public event EventHandler<RemoteEntryValueEventArgs>? RemoteEntryValueReceived;
    public event EventHandler<RemoteEntrySubmitEventArgs>? RemoteEntrySubmitReceived;
    // وقتی کاربر روی گوشی دکمه‌ی «قبلی» را زده (می‌خواهد به مرحله‌ی قبلِ ویزارد برگردد - مثلاً
    // چون یک کادر را اشتباه زده).
    public event EventHandler<RemoteEntryBackEventArgs>? RemoteEntryBackReceived;

    public string ComputerId { get; }
    public string ComputerName { get; }
    public string LanIp { get; private set; }
    public int ConnectedClients { get; private set; }
    public ConnectionState ConnectionState { get; private set; } = ConnectionState.Offline;

    private sealed class DeviceState
    {
        public string DeviceName = "دستگاه ناشناس";
        public bool HasScanned;
        public DateTime LastSeenUtc = DateTime.UtcNow;
    }

    private sealed class PeerInfo
    {
        public string ComputerId = string.Empty;
        public string Ip = string.Empty;
        public DateTime LastSeenUtc = DateTime.UtcNow;
    }

    public ScanBridgeService()
    {
        LanIp = GetPrimaryLanIp() ?? "127.0.0.1";
        ComputerId = ReadOrCreateComputerId(AppContext.BaseDirectory);
        ComputerName = ReadOptionalComputerName(AppContext.BaseDirectory);
        _pruneTimer.AutoReset = true;
        _pruneTimer.Elapsed += (_, _) => PruneOldScans();
        _connectionHealthTimer.AutoReset = true;
        _connectionHealthTimer.Elapsed += (_, _) => CheckConnectionHealth();
        _lanIpWatchTimer.AutoReset = true;
        _lanIpWatchTimer.Elapsed += (_, _) => CheckLanIpChange();
    }

    public void Start()
    {
        if (_server is not null)
        {
            return;
        }

        _lanIpWatchTimer.Start();

        _queueCts = new CancellationTokenSource();
        _queueTask = Task.Factory.StartNew(() => ProcessQueue(_queueCts.Token), TaskCreationOptions.LongRunning);

        var listenUrl = $"ws://0.0.0.0:{Port}";
        // عمداً روی یک متغیر محلی ساخته می‌شود، نه مستقیم روی _server: اگر بایند پورت زیر
        // (server.Start پایین) شکست بخورد (مثلاً پورت از قبل توسط یک نمونه‌ی دیگر برنامه یا یک
        // فرآیند نیمه‌بسته اشغال شده باشد)، نباید _server ست شود - وگرنه چک بالای همین متد
        // (if (_server is not null) return;) هر تلاش بعدیِ Start() را برای همیشه بی‌صدا no-op
        // می‌کند، حتی بعد از اینکه پورت آزاد شد (باگ ۱۹ گزارش ممیزی).
        var server = new WebSocketServer(listenUrl);
        try
        {
            server.Start(socket =>
            {
            socket.OnOpen = () =>
            {
                lock (_connectionLock)
                {
                    ConnectedClients++;
                }

                _connectedDevices[socket] = new DeviceState { LastSeenUtc = DateTime.UtcNow };

                PublishConnectionState();
                PublishConnectedDevices();
            };

            socket.OnClose = () =>
            {
                _erroredConnections.TryRemove(socket, out bool wasAbnormal);
                RemoveTrackedSocket(socket, wasAbnormal);
            };

            socket.OnError = ex =>
            {
                _erroredConnections[socket] = true;
                Console.WriteLine($"[{DateTime.UtcNow:O}] Connection error: {ex.Message}");
            };

            socket.OnMessage = message =>
            {
                string deviceName = "";
                string barcode = string.Empty;
                string messageType = string.Empty;
                string remoteEntryStepId = string.Empty;
                string remoteEntryValue = string.Empty;

                if (_connectedDevices.TryGetValue(socket, out var heartbeatState))
                    heartbeatState.LastSeenUtc = DateTime.UtcNow;

                // قبلاً این بخش یک try/catch واحد داشت: هر استثنایی - چه «اصلاً JSON نیست» (فرمت
                // متن‌سادهٔ قدیمی‌تر که هنوز پشتیبانی می‌شود) چه «JSON هست ولی یک فیلدش نوع
                // اشتباه دارد» (مثلاً barcode به‌جای رشته، عدد فرستاده شود - از یک نسخه‌ی متفاوت
                // اپ گوشی) - به یک نتیجه می‌رسید: کل متن خام پیام (شامل آکولاد/کوتیشن‌های JSON) به
                // barcode می‌رفت و همان‌طور تایپ می‌شد. حالا این دو حالت را از هم جدا می‌کنیم: فقط
                // وقتی واقعاً JSON نیست به فرمت متن‌ساده برمی‌گردیم؛ اگر JSON معتبر بود ولی یک فیلد
                // نوعش اشتباه بود، فقط همان یک فیلد نادیده گرفته می‌شود (خالی می‌ماند - که پایین‌تر
                // با IsNullOrWhiteSpace(barcode) هم‌اکنون به‌درستی نادیده گرفته می‌شود)، نه کل پیام.
                try
                {
                    using var doc = System.Text.Json.JsonDocument.Parse(message);
                    deviceName = GetJsonMessageString(doc.RootElement, "deviceName");
                    barcode = GetJsonMessageString(doc.RootElement, "barcode");
                    messageType = GetJsonMessageString(doc.RootElement, "type");
                    remoteEntryStepId = GetJsonMessageString(doc.RootElement, "stepId");
                    remoteEntryValue = GetJsonMessageString(doc.RootElement, "value");
                }
                catch (System.Text.Json.JsonException)
                {
                    // این پیام اصلاً JSON نیست (فرمت قدیمی‌تر: خودِ متن پیام همان بارکد است).
                    barcode = message?.Trim() ?? string.Empty;
                }

                // پیام‌های ویژگی «ورود اطلاعات از راه دور» (فرم شیرخشک روی گوشی) کاملاً مستقل از
                // مسیر عادی بارکد/اسکنر هستند - نه صف می‌شوند نه با اسکن واقعی قاطی می‌شوند؛ در این
                // پیام‌ها فیلد "barcode" همان بارکد قلم شیرخشکی است که فرمش روی گوشی باز است.
                if (messageType == "REMOTE_ENTRY_VALUE")
                {
                    if (!string.IsNullOrWhiteSpace(barcode) && !string.IsNullOrWhiteSpace(remoteEntryStepId))
                        RemoteEntryValueReceived?.Invoke(this, new RemoteEntryValueEventArgs(barcode, remoteEntryStepId, remoteEntryValue));
                    return;
                }
                if (messageType == "REMOTE_ENTRY_SUBMIT")
                {
                    if (!string.IsNullOrWhiteSpace(barcode))
                        RemoteEntrySubmitReceived?.Invoke(this, new RemoteEntrySubmitEventArgs(barcode));
                    return;
                }
                if (messageType == "REMOTE_ENTRY_BACK")
                {
                    if (!string.IsNullOrWhiteSpace(barcode))
                        RemoteEntryBackReceived?.Invoke(this, new RemoteEntryBackEventArgs(barcode));
                    return;
                }

                string clientIp = socket.ConnectionInfo?.ClientIpAddress ?? string.Empty;
                if (IsDeviceManuallyBlocked(deviceName) || IsIpManuallyBlocked(clientIp))
                {
                    try
                    {
                        SafeSend(socket, "DISCONNECT");
                    }
                    catch { }

                    socket.Close();
                    return;
                }

                if (_connectedDevices.TryGetValue(socket, out var state))
                {
                    state.LastSeenUtc = DateTime.UtcNow;
                    if (!string.IsNullOrWhiteSpace(deviceName))
                    {
                        state.DeviceName = deviceName;
                    }
                    state.HasScanned = true;
                    PublishConnectedDevices();
                }

                if (!string.IsNullOrWhiteSpace(barcode))
                {
                    _keyboardQueue.Add((socket, barcode, deviceName));
                }
            };
            });
        }
        catch
        {
            // بایند پورت شکست خورد - هر چیزی که بالاتر از این نقطه استارت شده بود (تایمر IP
            // شبکه، صف پردازش اسکن) هم متوقف/پاک شود تا وضعیت نیمه‌راه نماند و تلاش بعدیِ Start()
            // بتواند از صفر و کامل انجام شود.
            _lanIpWatchTimer.Stop();
            try { _queueCts?.Cancel(); } catch { }
            _queueCts?.Dispose();
            _queueCts = null;
            _queueTask = null;
            throw;
        }

        _server = server;

        StartPeerSync();

        _pruneTimer.Start();
        _connectionHealthTimer.Start();
        PruneOldScans();
    }

    // JsonElement.GetString() خودش exception می‌دهد اگر فیلد موجود باشد ولی نوعش رشته/null نباشد
    // (مثلاً یک نسخه‌ی دیگر اپ گوشی barcode را به‌جای رشته، عدد بفرستد). این helper چنین حالتی را
    // exception نمی‌دهد - فقط همان یک فیلد را نادیده می‌گیرد (رشته‌ی خالی برمی‌گرداند) تا خطای یک
    // فیلد باعث نشود کل پیامِ JSON خام به‌جای بارکد استفاده شود.
    private static string GetJsonMessageString(System.Text.Json.JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out var prop))
            return string.Empty;

        return prop.ValueKind == System.Text.Json.JsonValueKind.String ? (prop.GetString() ?? string.Empty) : string.Empty;
    }

    // =========================================================================================
    // هماهنگ‌سازی تنظیمات بین سیستم‌های دسکتاپ هم‌شبکه که همان لایسنس را دارند - کاملاً محلی
    // (LAN)، بدون هیچ سرور ابری. دو بخش دارد: کشف همکار با UDP broadcast (پورت DiscoveryPort)، و
    // رد و بدل خودِ تنظیمات با یک سرور WebSocket جدا (PeerSyncPort) که فقط بین دسکتاپ‌ها استفاده
    // می‌شود - عمداً از سرور اصلی گوشی‌ها جدا نگه داشته شده تا هیچ‌وقت به‌عنوان یک «گوشی وصل‌شده»
    // در رابط کاربری دیده نشود.
    // =========================================================================================

    public void SetLicenseGroupKey(string groupKeyHash)
    {
        _licenseGroupKey = groupKeyHash ?? string.Empty;
    }

    /// <summary>
    /// آخرین تنظیمات محلی را به‌عنوان مبنا ذخیره می‌کند و بلافاصله برای همه‌ی همکارهای شناخته‌شده
    /// (روی همین شبکه، با همان لایسنس) می‌فرستد.
    /// </summary>
    public void PublishDesktopSettings(string payloadJson, long versionUtcMs)
    {
        _lastDesktopSettingsJson = payloadJson ?? "{}";
        _lastDesktopSettingsVersionUtcMs = versionUtcMs;

        if (string.IsNullOrWhiteSpace(_licenseGroupKey))
            return;

        foreach (var peer in _knownPeers.Values.ToArray())
        {
            _ = PushDesktopSettingsToPeerAsync(peer, _lastDesktopSettingsJson, _lastDesktopSettingsVersionUtcMs);
        }
    }

    /// <summary>یک اعلان فوری می‌فرستد (به‌جای صبر کردن برای چرخه‌ی خودکار بعدی).</summary>
    public void AnnounceNow()
    {
        _ = SendAnnounceOnceAsync();
    }

    /// <summary>
    /// عکسِ کامل فعلی بانک بارکد پرمصرف را به‌عنوان مبنا ذخیره می‌کند (برای پاسخ به درخواست
    /// هماهنگ‌سازی اولیه‌ی سیستم‌های تازه‌وصل‌شده) - خودش چیزی را «پوش» نمی‌کند، فقط کش می‌کند.
    /// پوش زنده به همکارها از طریق BroadcastHighUsageBarcodeOperation انجام می‌شود.
    /// </summary>
    public void PublishHighUsageBarcodeSnapshot(string payloadJson, long versionUtcMs)
    {
        _lastHighUsageSnapshotJson = payloadJson ?? "[]";
        _lastHighUsageSnapshotVersionUtcMs = versionUtcMs;
    }

    /// <summary>
    /// یک عملیات تکی (اضافه/حذف/دیسپنس یک بارکد) را همین الان برای همه‌ی همکارهای شناخته‌شده
    /// (روی همین شبکه، با همان لایسنس) می‌فرستد - جدا از عکسِ کامل بانک، برای این‌که تغییرات با
    /// حداقل تاخیر به بقیه‌ی سیستم‌ها برسد (مهم برای جلوگیری از دیسپنس دوباره‌ی یک بارکد فیزیکی
    /// یکسان روی دو سیستم هم‌زمان).
    /// </summary>
    public void BroadcastHighUsageBarcodeOperation(string opJson)
    {
        if (string.IsNullOrWhiteSpace(_licenseGroupKey) || string.IsNullOrWhiteSpace(opJson))
            return;

        foreach (var peer in _knownPeers.Values.ToArray())
        {
            _ = PushHighUsageOperationToPeerAsync(peer, opJson);
        }
    }

    private void StartPeerSync()
    {
        try
        {
            var peerListenUrl = $"ws://0.0.0.0:{PeerSyncPort}";
            _peerServer = new WebSocketServer(peerListenUrl);
            _peerServer.Start(socket =>
            {
                socket.OnMessage = message => HandlePeerMessage(socket, message);
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to start peer-sync server: {ex.Message}");
        }

        try
        {
            _discoveryCts = new CancellationTokenSource();
            _discoveryUdp = new UdpClient();
            _discoveryUdp.EnableBroadcast = true;
            try { _discoveryUdp.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true); } catch { }
            _discoveryUdp.Client.Bind(new IPEndPoint(IPAddress.Any, DiscoveryPort));

            _ = Task.Run(() => DiscoveryReceiveLoopAsync(_discoveryCts.Token));
            _ = Task.Run(() => DiscoveryAnnounceLoopAsync(_discoveryCts.Token));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to start LAN discovery: {ex.Message}");
        }
    }

    private void HandlePeerMessage(IWebSocketConnection socket, string message)
    {
        try
        {
            using var doc = JsonDocument.Parse(message);
            string type = doc.RootElement.TryGetProperty("type", out var t) ? t.GetString() ?? string.Empty : string.Empty;
            string groupKey = doc.RootElement.TryGetProperty("groupKey", out var gk) ? gk.GetString() ?? string.Empty : string.Empty;

            if (string.IsNullOrWhiteSpace(_licenseGroupKey) || !string.Equals(groupKey, _licenseGroupKey, StringComparison.Ordinal))
                return;

            if (type == "SCANBRIDGE_SETTINGS_SYNC")
            {
                string fromComputerId = doc.RootElement.TryGetProperty("computerId", out var cid) ? cid.GetString() ?? string.Empty : string.Empty;
                long version = doc.RootElement.TryGetProperty("versionUtcMs", out var v) && v.TryGetInt64(out var vv) ? vv : 0;
                string payload = doc.RootElement.TryGetProperty("payload", out var p) ? p.GetRawText() : "{}";

                if (version > _lastDesktopSettingsVersionUtcMs)
                {
                    _lastDesktopSettingsJson = payload;
                    _lastDesktopSettingsVersionUtcMs = version;
                    PeerDesktopSettingsReceived?.Invoke(this, new PeerDesktopSettingsEventArgs(payload, version, fromComputerId));
                }
            }
            else if (type == "SCANBRIDGE_SETTINGS_REQUEST" && _lastDesktopSettingsVersionUtcMs > 0)
            {
                var responseObj = new
                {
                    type = "SCANBRIDGE_SETTINGS_SYNC",
                    groupKey = _licenseGroupKey,
                    computerId = ComputerId,
                    versionUtcMs = _lastDesktopSettingsVersionUtcMs,
                    payload = JsonDocument.Parse(string.IsNullOrWhiteSpace(_lastDesktopSettingsJson) ? "{}" : _lastDesktopSettingsJson).RootElement
                };
                try { SafeSend(socket, JsonSerializer.Serialize(responseObj)); } catch { }
            }
            // یک عملیات تکی و آنی (اضافه/حذف/دیسپنس یک بارکد) از یک سیستم هم‌شبکه رسیده - بدون
            // مقایسه‌ی نسخه (خودِ عملیات‌ها با شناسه اعمال می‌شوند، پس تکراری بودن بی‌ضرر است؛
            // نگاه کنید به ApplyHighUsageBarcodeOperation در MainWindow).
            else if (type == "SCANBRIDGE_HIGHUSAGE_OP")
            {
                string fromComputerId = doc.RootElement.TryGetProperty("computerId", out var opCid) ? opCid.GetString() ?? string.Empty : string.Empty;
                string opPayload = doc.RootElement.TryGetProperty("op", out var opProp) ? opProp.GetRawText() : "{}";
                HighUsageBarcodeOperationReceived?.Invoke(this, new HighUsageBarcodeOperationEventArgs(opPayload, fromComputerId));
            }
            // عکسِ کامل بانک از یک سیستم هم‌شبکه رسیده - فقط برای هماهنگ‌سازی اولیه استفاده می‌شود؛
            // نسخه‌ی بزرگ‌تر برنده است (مثل تنظیمات).
            else if (type == "SCANBRIDGE_HIGHUSAGE_SNAPSHOT_SYNC")
            {
                string fromComputerId = doc.RootElement.TryGetProperty("computerId", out var snCid) ? snCid.GetString() ?? string.Empty : string.Empty;
                long version = doc.RootElement.TryGetProperty("versionUtcMs", out var snV) && snV.TryGetInt64(out var snVv) ? snVv : 0;
                string payload = doc.RootElement.TryGetProperty("payload", out var snP) ? snP.GetRawText() : "[]";

                if (version > _lastHighUsageSnapshotVersionUtcMs)
                {
                    _lastHighUsageSnapshotJson = payload;
                    _lastHighUsageSnapshotVersionUtcMs = version;
                    HighUsageBarcodeSnapshotReceived?.Invoke(this, new HighUsageBarcodeSnapshotEventArgs(payload, version, fromComputerId));
                }
            }
            else if (type == "SCANBRIDGE_HIGHUSAGE_SNAPSHOT_REQUEST" && _lastHighUsageSnapshotVersionUtcMs > 0)
            {
                var responseObj = new
                {
                    type = "SCANBRIDGE_HIGHUSAGE_SNAPSHOT_SYNC",
                    groupKey = _licenseGroupKey,
                    computerId = ComputerId,
                    versionUtcMs = _lastHighUsageSnapshotVersionUtcMs,
                    payload = JsonDocument.Parse(string.IsNullOrWhiteSpace(_lastHighUsageSnapshotJson) ? "[]" : _lastHighUsageSnapshotJson).RootElement
                };
                try { SafeSend(socket, JsonSerializer.Serialize(responseObj)); } catch { }
            }
        }
        catch { }
    }

    private async Task DiscoveryReceiveLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try
            {
                if (_discoveryUdp is null)
                    return;

                var result = await _discoveryUdp.ReceiveAsync(ct);
                string json = Encoding.UTF8.GetString(result.Buffer);
                using var doc = JsonDocument.Parse(json);

                if (!doc.RootElement.TryGetProperty("type", out var t) || t.GetString() != "SCANBRIDGE_PEER_ANNOUNCE")
                    continue;

                string groupKey = doc.RootElement.TryGetProperty("groupKey", out var gk) ? gk.GetString() ?? string.Empty : string.Empty;
                string peerComputerId = doc.RootElement.TryGetProperty("computerId", out var cid) ? cid.GetString() ?? string.Empty : string.Empty;
                string ip = doc.RootElement.TryGetProperty("ip", out var ipProp) ? ipProp.GetString() ?? string.Empty : string.Empty;

                if (string.IsNullOrWhiteSpace(_licenseGroupKey) || string.IsNullOrWhiteSpace(groupKey) || !string.Equals(groupKey, _licenseGroupKey, StringComparison.Ordinal))
                    continue;
                if (string.IsNullOrWhiteSpace(peerComputerId) || string.Equals(peerComputerId, ComputerId, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (string.IsNullOrWhiteSpace(ip))
                    ip = result.RemoteEndPoint.Address.ToString();

                bool isNewPeer = !_knownPeers.ContainsKey(peerComputerId);
                var peer = new PeerInfo { ComputerId = peerComputerId, Ip = ip, LastSeenUtc = DateTime.UtcNow };
                _knownPeers[peerComputerId] = peer;

                if (isNewPeer)
                {
                    // یک سیستم تازه‌پیدا‌شده - هم از او درخواست آخرین تنظیمات را می‌کنیم، هم تنظیمات
                    // خودمان را برایش می‌فرستیم؛ هرکدام نسخه‌ی جدیدتر بود، همان برنده می‌شود.
                    _ = RequestDesktopSettingsFromPeerAsync(peer);
                    if (_lastDesktopSettingsVersionUtcMs > 0)
                        _ = PushDesktopSettingsToPeerAsync(peer, _lastDesktopSettingsJson, _lastDesktopSettingsVersionUtcMs);

                    // همان کار برای بانک بارکد پرمصرف - هماهنگ‌سازی اولیه‌ی دوطرفه با یک سیستم
                    // تازه‌پیدا‌شده (بعد از این، به‌روزرسانی‌های زنده از طریق عملیات‌های تکی می‌آیند).
                    _ = RequestHighUsageSnapshotFromPeerAsync(peer);
                    if (_lastHighUsageSnapshotVersionUtcMs > 0)
                        _ = PushHighUsageSnapshotToPeerAsync(peer, _lastHighUsageSnapshotJson, _lastHighUsageSnapshotVersionUtcMs);
                }
            }
            catch (OperationCanceledException) { }
            catch (ObjectDisposedException) { return; }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.UtcNow:O}] LAN discovery receive error: {ex.Message}");
                try { await Task.Delay(1000, ct); } catch { }
            }
        }
    }

    private async Task DiscoveryAnnounceLoopAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            await SendAnnounceOnceAsync();
            PrunePeers();

            try { await Task.Delay(TimeSpan.FromSeconds(15), ct); }
            catch (OperationCanceledException) { return; }
        }
    }

    private async Task SendAnnounceOnceAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_licenseGroupKey) || _discoveryUdp is null)
                return;

            var announce = new
            {
                type = "SCANBRIDGE_PEER_ANNOUNCE",
                groupKey = _licenseGroupKey,
                computerId = ComputerId,
                computerName = ComputerName,
                ip = GetPrimaryLanIp() ?? LanIp
            };
            byte[] bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(announce));
            await _discoveryUdp.SendAsync(bytes, bytes.Length, new IPEndPoint(IPAddress.Broadcast, DiscoveryPort));
        }
        catch { }
    }

    private void PrunePeers()
    {
        var cutoff = DateTime.UtcNow.AddMinutes(-2);
        foreach (var kvp in _knownPeers.ToArray())
        {
            if (kvp.Value.LastSeenUtc < cutoff)
                _knownPeers.TryRemove(kvp.Key, out _);
        }
    }

    private async Task PushDesktopSettingsToPeerAsync(PeerInfo peer, string payloadJson, long versionUtcMs)
    {
        if (string.IsNullOrWhiteSpace(_licenseGroupKey) || string.IsNullOrWhiteSpace(payloadJson))
            return;

        try
        {
            using var ws = new ClientWebSocket();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(4));
            await ws.ConnectAsync(new Uri($"ws://{peer.Ip}:{PeerSyncPort}"), cts.Token);

            var messageObj = new
            {
                type = "SCANBRIDGE_SETTINGS_SYNC",
                groupKey = _licenseGroupKey,
                computerId = ComputerId,
                versionUtcMs,
                payload = JsonDocument.Parse(payloadJson).RootElement
            };
            byte[] bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(messageObj));
            await ws.SendAsync(bytes, WebSocketMessageType.Text, true, cts.Token);
            try { await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", cts.Token); } catch { }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to push settings to peer {peer.ComputerId}: {ex.Message}");
        }
    }

    private async Task RequestDesktopSettingsFromPeerAsync(PeerInfo peer)
    {
        if (string.IsNullOrWhiteSpace(_licenseGroupKey))
            return;

        try
        {
            using var ws = new ClientWebSocket();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            await ws.ConnectAsync(new Uri($"ws://{peer.Ip}:{PeerSyncPort}"), cts.Token);

            var requestObj = new { type = "SCANBRIDGE_SETTINGS_REQUEST", groupKey = _licenseGroupKey, computerId = ComputerId };
            byte[] requestBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(requestObj));
            await ws.SendAsync(requestBytes, WebSocketMessageType.Text, true, cts.Token);

            var buffer = new byte[16384];
            var result = await ws.ReceiveAsync(buffer, cts.Token);
            if (result.MessageType == WebSocketMessageType.Text && result.Count > 0)
            {
                string json = Encoding.UTF8.GetString(buffer, 0, result.Count);
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("type", out var t) && t.GetString() == "SCANBRIDGE_SETTINGS_SYNC")
                {
                    string groupKey = doc.RootElement.TryGetProperty("groupKey", out var gk) ? gk.GetString() ?? string.Empty : string.Empty;
                    if (string.Equals(groupKey, _licenseGroupKey, StringComparison.Ordinal))
                    {
                        long version = doc.RootElement.TryGetProperty("versionUtcMs", out var v) && v.TryGetInt64(out var vv) ? vv : 0;
                        string payload = doc.RootElement.TryGetProperty("payload", out var p) ? p.GetRawText() : "{}";
                        string fromComputerId = doc.RootElement.TryGetProperty("computerId", out var cid) ? cid.GetString() ?? string.Empty : string.Empty;

                        if (version > _lastDesktopSettingsVersionUtcMs)
                        {
                            _lastDesktopSettingsJson = payload;
                            _lastDesktopSettingsVersionUtcMs = version;
                            PeerDesktopSettingsReceived?.Invoke(this, new PeerDesktopSettingsEventArgs(payload, version, fromComputerId));
                        }
                    }
                }
            }
            try { await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", cts.Token); } catch { }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to request settings from peer {peer.ComputerId}: {ex.Message}");
        }
    }

    private async Task PushHighUsageOperationToPeerAsync(PeerInfo peer, string opJson)
    {
        if (string.IsNullOrWhiteSpace(_licenseGroupKey) || string.IsNullOrWhiteSpace(opJson))
            return;

        try
        {
            using var ws = new ClientWebSocket();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(4));
            await ws.ConnectAsync(new Uri($"ws://{peer.Ip}:{PeerSyncPort}"), cts.Token);

            var messageObj = new
            {
                type = "SCANBRIDGE_HIGHUSAGE_OP",
                groupKey = _licenseGroupKey,
                computerId = ComputerId,
                op = JsonDocument.Parse(opJson).RootElement
            };
            byte[] bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(messageObj));
            await ws.SendAsync(bytes, WebSocketMessageType.Text, true, cts.Token);
            try { await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", cts.Token); } catch { }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to push high-usage-barcode operation to peer {peer.ComputerId}: {ex.Message}");
        }
    }

    private async Task PushHighUsageSnapshotToPeerAsync(PeerInfo peer, string payloadJson, long versionUtcMs)
    {
        if (string.IsNullOrWhiteSpace(_licenseGroupKey) || string.IsNullOrWhiteSpace(payloadJson))
            return;

        try
        {
            using var ws = new ClientWebSocket();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(8));
            await ws.ConnectAsync(new Uri($"ws://{peer.Ip}:{PeerSyncPort}"), cts.Token);

            var messageObj = new
            {
                type = "SCANBRIDGE_HIGHUSAGE_SNAPSHOT_SYNC",
                groupKey = _licenseGroupKey,
                computerId = ComputerId,
                versionUtcMs,
                payload = JsonDocument.Parse(payloadJson).RootElement
            };
            byte[] bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(messageObj));
            await ws.SendAsync(bytes, WebSocketMessageType.Text, true, cts.Token);
            try { await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", cts.Token); } catch { }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to push high-usage-barcode snapshot to peer {peer.ComputerId}: {ex.Message}");
        }
    }

    // بانک بارکد پرمصرف می‌تواند خیلی بزرگ باشد (مثلاً هزار بارکد سرم) - برخلاف پیام‌های کوچک
    // تنظیمات، ممکن است در یک فریم/بافر ۱۶ کیلوبایتی جا نشود. برخلاف RequestDesktopSettingsFromPeerAsync
    // (که یک ReceiveAsync تکی کافی بود چون همیشه کوچک است)، اینجا تا رسیدن به انتهای پیام
    // (EndOfMessage) در یک حلقه می‌خوانیم تا هیچ داده‌ای برش نخورد.
    private static async Task<string?> ReceiveFullTextMessageAsync(ClientWebSocket ws, CancellationToken ct)
    {
        var buffer = new byte[16384];
        using var stream = new MemoryStream();
        while (true)
        {
            var result = await ws.ReceiveAsync(buffer, ct);
            if (result.MessageType == WebSocketMessageType.Close)
                return null;
            if (result.Count > 0)
                stream.Write(buffer, 0, result.Count);
            if (result.EndOfMessage)
                break;
        }
        return stream.Length > 0 ? Encoding.UTF8.GetString(stream.ToArray()) : null;
    }

    private async Task RequestHighUsageSnapshotFromPeerAsync(PeerInfo peer)
    {
        if (string.IsNullOrWhiteSpace(_licenseGroupKey))
            return;

        try
        {
            using var ws = new ClientWebSocket();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            await ws.ConnectAsync(new Uri($"ws://{peer.Ip}:{PeerSyncPort}"), cts.Token);

            var requestObj = new { type = "SCANBRIDGE_HIGHUSAGE_SNAPSHOT_REQUEST", groupKey = _licenseGroupKey, computerId = ComputerId };
            byte[] requestBytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(requestObj));
            await ws.SendAsync(requestBytes, WebSocketMessageType.Text, true, cts.Token);

            string? json = await ReceiveFullTextMessageAsync(ws, cts.Token);
            if (!string.IsNullOrEmpty(json))
            {
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("type", out var t) && t.GetString() == "SCANBRIDGE_HIGHUSAGE_SNAPSHOT_SYNC")
                {
                    string groupKey = doc.RootElement.TryGetProperty("groupKey", out var gk) ? gk.GetString() ?? string.Empty : string.Empty;
                    if (string.Equals(groupKey, _licenseGroupKey, StringComparison.Ordinal))
                    {
                        long version = doc.RootElement.TryGetProperty("versionUtcMs", out var v) && v.TryGetInt64(out var vv) ? vv : 0;
                        string payload = doc.RootElement.TryGetProperty("payload", out var p) ? p.GetRawText() : "[]";
                        string fromComputerId = doc.RootElement.TryGetProperty("computerId", out var cid) ? cid.GetString() ?? string.Empty : string.Empty;

                        if (version > _lastHighUsageSnapshotVersionUtcMs)
                        {
                            _lastHighUsageSnapshotJson = payload;
                            _lastHighUsageSnapshotVersionUtcMs = version;
                            HighUsageBarcodeSnapshotReceived?.Invoke(this, new HighUsageBarcodeSnapshotEventArgs(payload, version, fromComputerId));
                        }
                    }
                }
            }
            try { await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", cts.Token); } catch { }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to request high-usage-barcode snapshot from peer {peer.ComputerId}: {ex.Message}");
        }
    }

    private void ProcessQueue(CancellationToken ct)
    {
        try
        {
            foreach (var item in _keyboardQueue.GetConsumingEnumerable(ct))
            {
                // مهم: هرچه داخل این بلوک اتفاق بیفتد (از جمله استثناهای پرتاب‌شده توسط مشترکین
                // رویداد ScanReceived در AppendScan) باید همین‌جا گرفته شود. اگر استثنایی از این
                // حلقه خارج شود، ترد پردازش صف برای همیشه متوقف می‌شود و هیچ اسکن بعدی (تا
                // ری‌استارت برنامه) نه تایپ می‌شود، نه ذخیره، نه به رابط کاربری می‌رسد — بدون
                // هیچ پیام خطایی به کاربر.
                try
                {
                    string barcode = item.Barcode;
                    string deviceName = item.DeviceName;
                    try
                    {
                        KeyboardInjector.TypeText(barcode);
                        KeyboardInjector.PressEnter();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[{DateTime.UtcNow:O}] SendInput failed: {ex.Message}");
                        try
                        {
                            KeyboardInjector.SendKeysFallback(barcode);
                            KeyboardInjector.PressEnter();
                        }
                        catch (Exception fallbackEx)
                        {
                            Console.WriteLine($"[{DateTime.UtcNow:O}] Keyboard injection fallback failed: {fallbackEx.Message}");
                        }
                    }

                    try
                    {
                        AppendScan(barcode, deviceName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[{DateTime.UtcNow:O}] AppendScan/ScanReceived handling failed: {ex}");
                    }

                    Console.WriteLine($"Received scan from {deviceName}: {barcode}");

                    try
                    {
                        SafeSend(item.Socket, $"OK {barcode}", ok =>
                        {
                            if (!ok)
                                Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to send ack for {barcode}.");
                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to send ack: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    // شبکه‌ی ایمنی نهایی: هر خطای غیرمنتظره‌ی دیگر هم صف پردازش را متوقف نکند.
                    Console.WriteLine($"[{DateTime.UtcNow:O}] Unexpected error while processing a scan: {ex}");
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown
        }
    }

    /// <summary>
    /// تمام Sendهای روی یک اتصال WebSocket (ack اسکن، پینگ سلامت، broadcastهای هشدار/ورود از
    /// راه دور، پیام DISCONNECT) باید از همین متد رد شوند - نگاه کنید به توضیح بالای
    /// _socketSendGates. هر Send واقعی داخل نوبت خودش (SemaphoreSlim مخصوص همان سوکت) اجرا
    /// می‌شود، پس دیگر دو نوشتن روی یک اتصال هم‌زمان کامل نمی‌شوند. غیرهمزمان/fire-and-forget
    /// است تا رفتار محل‌های فراخوانی قبلی (که همه `_ = socket.Send(...)` بودند) عوض نشود؛
    /// onSettled اختیاری برای مواردی است که قبلاً به نتیجه‌ی Send نیاز داشتند (مثلاً پینگ سلامت
    /// که روی شکست باید اتصال را ببندد).
    /// </summary>
    private void SafeSend(IWebSocketConnection socket, string payload, Action<bool>? onSettled = null)
    {
        var gate = _socketSendGates.GetOrAdd(socket, _ => new SemaphoreSlim(1, 1));
        _ = Task.Run(async () =>
        {
            bool ok = false;
            try
            {
                await gate.WaitAsync();
                try
                {
                    await socket.Send(payload);
                    ok = true;
                }
                finally
                {
                    // اگر درست همین بین اتصال بسته و RemoveTrackedSocket این gate را Dispose کرده
                    // باشد، Release می‌تواند ObjectDisposedException بدهد - بی‌ضرر است، فقط نادیده گرفته شود.
                    try { gate.Release(); } catch (ObjectDisposedException) { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.UtcNow:O}] SafeSend failed: {ex.Message}");
            }

            onSettled?.Invoke(ok);
        });
    }

    private void RemoveTrackedSocket(IWebSocketConnection socket, bool abnormal)
    {
        bool removed = _connectedDevices.TryRemove(socket, out _);
        _erroredConnections.TryRemove(socket, out _);
        if (_socketSendGates.TryRemove(socket, out var gate))
            gate.Dispose();

        if (removed)
        {
            lock (_connectionLock)
            {
                if (ConnectedClients > 0)
                    ConnectedClients--;
            }

            PublishConnectionState();
            PublishConnectedDevices();

            if (abnormal)
                UnexpectedDisconnection?.Invoke(this, EventArgs.Empty);
        }
    }

    private void CheckConnectionHealth()
    {
        foreach (var pair in _connectedDevices.ToArray())
        {
            var socket = pair.Key;
            var state = pair.Value;

            try
            {
                // پیام سبک سلامت باعث می‌شود اتصال‌های مرده/گوشی خاموش‌شده زودتر توسط TCP/WebSocket مشخص شوند.
                SafeSend(socket, "{\"type\":\"SCANBRIDGE_PING\"}", ok =>
                {
                    if (!ok)
                    {
                        try { socket.Close(); } catch { }
                        RemoveTrackedSocket(socket, abnormal: true);
                    }
                });

                // اگر اتصال ساعت‌ها هیچ واکنشی نداشته باشد، برای جلوگیری از نمایش دستگاه روحی حذف می‌شود.
                if (DateTime.UtcNow - state.LastSeenUtc > TimeSpan.FromHours(6))
                {
                    try { socket.Close(); } catch { }
                    RemoveTrackedSocket(socket, abnormal: false);
                }
            }
            catch
            {
                try { socket.Close(); } catch { }
                RemoveTrackedSocket(socket, abnormal: true);
            }
        }
    }

    public bool DisconnectDevice(string deviceName)
    {
        if (string.IsNullOrWhiteSpace(deviceName))
            return false;

        // جلوگیری از اتصال مجدد خودکار اپ گوشی برای چند دقیقه بعد از قطع دستی
        var blockedUntilUtc = DateTime.UtcNow.AddMinutes(10);
        _manuallyBlockedDevices[deviceName] = blockedUntilUtc;

        var sockets = _connectedDevices
            .Where(pair => string.Equals(pair.Value.DeviceName, deviceName, StringComparison.OrdinalIgnoreCase))
            .Select(pair => pair.Key)
            .ToList();

        if (sockets.Count == 0)
            return false;

        bool disconnected = false;
        var blockedIps = new List<string>();
        foreach (var socket in sockets)
        {
            try
            {
                // این قطع اتصال توسط کاربر است، پس نباید هشدار قطع غیرمنتظره نشان داده شود.
                _erroredConnections.TryRemove(socket, out _);

                // IP اتصال هم بلاک می‌شود، مستقل از deviceName - نگاه کنید به توضیح بالای
                // _manuallyBlockedIps.
                string ip = socket.ConnectionInfo?.ClientIpAddress ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(ip))
                {
                    _manuallyBlockedIps[ip] = blockedUntilUtc;
                    blockedIps.Add(ip);
                }

                try
                {
                    SafeSend(socket, "DISCONNECT");
                }
                catch { }

                socket.Close();
                disconnected = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to disconnect device '{deviceName}': {ex.Message}");
            }
        }

        if (blockedIps.Count > 0)
            _blockedDeviceNameToIps[deviceName] = blockedIps;

        return disconnected;
    }

    public void AllowAllDevicesToReconnect()
    {
        _manuallyBlockedDevices.Clear();
        _manuallyBlockedIps.Clear();
        _blockedDeviceNameToIps.Clear();
    }

    /// <summary>
    /// یک هشدار (مثلاً نتیجه‌ی ثبت شیرخشک - موفق، ناموفق یا هر پیام دیگر) را برای همه‌ی
    /// گوشی‌های وصل‌شده می‌فرستد. اپ اندروید این پیام را با فیلد "type": "SCANBRIDGE_ALERT"
    /// می‌شناسد و به‌صورت یک دیالوگ با دکمه‌ی «باشه» نشان می‌دهد. اگر گوشی‌ای وصل نباشد، این کار
    /// بی‌اثر است (نه خطا می‌دهد نه چیزی صف می‌شود - فقط هشدارهای لحظه‌ای که گوشی آنلاین است ارسال می‌شوند).
    /// اگر <paramref name="photoPath"/> داده شود (مسیر فایل عکس شیرخشک روی دیسک)، محتوای فایل
    /// به Base64 تبدیل و به‌عنوان "photoBase64" داخل همین پیام فرستاده می‌شود تا گوشی هم عکس را
    /// کنار پیام نشان دهد؛ اگر فایل پیدا نشود یا خواندنش خطا بدهد، پیام بدون عکس (فقط متن) فرستاده
    /// می‌شود - این خطا نباید جلوی رسیدن خودِ پیام را بگیرد.
    /// </summary>
    public void BroadcastAlert(string title, string body, bool success, string? photoPath = null)
    {
        string? photoBase64 = null;
        if (!string.IsNullOrWhiteSpace(photoPath))
        {
            try
            {
                byte[] photoBytes = File.ReadAllBytes(photoPath);
                photoBase64 = Convert.ToBase64String(photoBytes);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to read formula photo for phone alert: {ex.Message}");
            }
        }

        var payloadObj = new
        {
            type = "SCANBRIDGE_ALERT",
            title,
            body,
            success,
            photoBase64
        };

        string json = JsonSerializer.Serialize(payloadObj);
        foreach (var socket in _connectedDevices.Keys.ToArray())
        {
            try
            {
                SafeSend(socket, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to send alert to a device: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// یک مرحله از فرم ثبت شیرخشک (کد ملی، تاریخ تولد، کپچا و ...) را برای نمایش روی گوشی
    /// می‌فرستد - بخشی از ویژگی «ورود اطلاعات از راه دور». اگر گوشی‌ای وصل نباشد، بی‌اثر است.
    /// </summary>
    public void BroadcastRemoteEntryStep(string barcode, string stepId, string label, string hint,
        string? photoBase64 = null, string? captchaImageBase64 = null, string inputType = "text", string? prefillValue = null)
    {
        var payloadObj = new
        {
            type = "REMOTE_ENTRY_STEP",
            barcode,
            stepId,
            label,
            hint,
            inputType,
            photoBase64,
            captchaImageBase64,
            prefillValue
        };

        string json = JsonSerializer.Serialize(payloadObj);
        foreach (var socket in _connectedDevices.Keys.ToArray())
        {
            try { SafeSend(socket, json); }
            catch (Exception ex) { Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to send remote-entry step to a device: {ex.Message}"); }
        }
    }

    /// <summary>
    /// جریان «ورود اطلاعات از راه دور» را روی گوشی لغو می‌کند (مثلاً چون کاربر خودش فرم را روی
    /// دسکتاپ بست) - بدون هیچ پیام قابل‌مشاهده‌ای، فقط ویزارد را از روی گوشی پاک می‌کند.
    /// </summary>
    public void BroadcastRemoteEntryCancel(string barcode)
    {
        var payloadObj = new { type = "REMOTE_ENTRY_CANCEL", barcode };
        string json = JsonSerializer.Serialize(payloadObj);
        foreach (var socket in _connectedDevices.Keys.ToArray())
        {
            try { SafeSend(socket, json); }
            catch (Exception ex) { Console.WriteLine($"[{DateTime.UtcNow:O}] Failed to send remote-entry cancel to a device: {ex.Message}"); }
        }
    }

    public bool AllowDeviceToReconnect(string deviceName)
    {
        if (string.IsNullOrWhiteSpace(deviceName))
            return false;

        bool removedName = _manuallyBlockedDevices.TryRemove(deviceName, out _);
        if (_blockedDeviceNameToIps.TryRemove(deviceName, out var ips))
        {
            foreach (var ip in ips)
                _manuallyBlockedIps.TryRemove(ip, out _);
        }

        return removedName;
    }

    public void Dispose()
    {
        _pruneTimer.Stop();
        _connectionHealthTimer.Stop();
        _lanIpWatchTimer.Stop();
        _pruneTimer.Dispose();
        _connectionHealthTimer.Dispose();
        _lanIpWatchTimer.Dispose();
        _queueCts?.Cancel();

        if (_server is not null)
        {
            _server.Dispose();
            _server = null;
        }

        _discoveryCts?.Cancel();
        try { _discoveryUdp?.Close(); } catch { }
        _discoveryUdp?.Dispose();
        _discoveryUdp = null;

        if (_peerServer is not null)
        {
            _peerServer.Dispose();
            _peerServer = null;
        }

        ConnectedClients = 0;
        PublishConnectionState();
    }

    public byte[] CreatePairingQrPng()
    {
        try
        {
            string? ip = GetPrimaryLanIp();
            if (string.IsNullOrEmpty(ip))
            {
                ip = LanIp;
            }

            var payloadObj = new
            {
                type = "SCANBRIDGE_PAIR",
                protocolVersion = "1.0",
                computerId = ComputerId,
                computerName = ComputerName,
                ip,
                port = Port
            };

            string payloadJson = JsonSerializer.Serialize(payloadObj);
            using var qrGenerator = new QRCodeGenerator();
            using var qrData = qrGenerator.CreateQrCode(payloadJson, QRCodeGenerator.ECCLevel.M);
            var png = new PngByteQRCode(qrData);
            var bytes = png.GetGraphic(20);

            string path = Path.Combine(AppContext.BaseDirectory, "pairing-qr.png");
            File.WriteAllBytes(path, bytes);
            return bytes;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.UtcNow:O}] خطا هنگام تولید کد QR: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    public IReadOnlyList<ScanEntry> GetTodayHistory()
    {
        var path = GetScanFilePath();

        // خواندن فایل هم باید همان قفلی را بگیرد که AppendScan/PruneOldScans می‌گیرند، وگرنه
        // ممکن است هم‌زمان با بازنویسی کامل فایل توسط PruneOldScans یا نوشتن AppendScan اجرا
        // شود و یا استثنا بدهد یا داده‌ی نصفه‌نوشته‌شده را بخواند.
        List<string> lines;
        lock (_scanFileLock)
        {
            if (!File.Exists(path))
            {
                return Array.Empty<ScanEntry>();
            }

            lines = File.ReadLines(path).Skip(1).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();
        }

        var today = DateTime.UtcNow.Date;
        var result = new List<ScanEntry>();

        foreach (var line in lines)
        {
            try
            {
                // فایل سه‌ستونی است: timestamp_iso,deviceName,barcode (نگاه کنید به AppendScan).
                // قبلاً اینجا Split(',', 2) استفاده می‌شد که سطر را فقط به ۲ تکه می‌شکست - یعنی
                // parts[1] در واقع "نام‌دستگاه,بارکد" چسبیده‌به‌هم بود، نه فقط بارکد. حالا با
                // ParseCsvLine (که quote را هم درست می‌فهمد، برای نام‌دستگاه/بارکدی که به‌ندرت
                // ممکن است خودش کاما داشته باشد) هر سه ستون جدا استخراج می‌شود.
                var parts = ParseCsvLine(line);
                if (parts.Count < 3)
                {
                    continue;
                }

                var timestamp = DateTimeOffset.Parse(parts[0].Trim(), System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal);
                var barcode = parts[2].Trim();
                if (timestamp.UtcDateTime.Date == today)
                {
                    result.Add(new ScanEntry(timestamp.UtcDateTime, barcode));
                }
            }
            catch
            {
                // ignored intentionally to keep the history view resilient to malformed rows
            }
        }

        return result.OrderByDescending(x => x.TimestampUtc).ToList();
    }

    // یک CSV-parser ساده و quote-آگاه: کاما داخل یک فیلدِ داخل گیومه را جدا-کننده حساب نمی‌کند و
    // گیومه‌ی دوتایی ("") را به یک گیومه‌ی تکی تبدیل می‌کند - دقیقاً معکوس EscapeCsv در همین فایل.
    private static List<string> ParseCsvLine(string line)
    {
        var result = new List<string>();
        if (line is null)
            return result;

        var current = new StringBuilder();
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

    public void PruneOldScans()
    {
        var path = GetScanFilePath();

        // کل خواندن + بازنویسی فایل باید داخل همان قفلی باشد که AppendScan استفاده می‌کند؛
        // وگرنه یک اسکن تازه که هم‌زمان با این متد نوشته می‌شود ممکن است یا با خطای
        // «فایل توسط پردازش دیگری استفاده می‌شود» مواجه شود یا وسط بازنویسی این متد گم شود.
        lock (_scanFileLock)
        {
            if (!File.Exists(path))
            {
                return;
            }

            var rows = new List<string>();
            var lines = File.ReadAllLines(path);
            if (lines.Length == 0)
            {
                return;
            }

            rows.Add(lines[0]);
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                {
                    continue;
                }

                var row = lines[i];
                string[] parts = row.Split(',', 2);
                if (parts.Length < 2)
                {
                    continue;
                }

                try
                {
                    var timestamp = DateTimeOffset.Parse(parts[0].Trim(), System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal);
                    if (DateTime.UtcNow - timestamp.UtcDateTime <= TimeSpan.FromHours(24))
                    {
                        rows.Add(row);
                    }
                }
                catch
                {
                    // ignore malformed rows while pruning
                }
            }

            File.WriteAllLines(path, rows, Encoding.UTF8);
        }
    }

    public void ClearHistoryFile()
    {
        lock (_scanFileLock)
        {
            var path = GetScanFilePath();
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    private void AppendScan(string barcode, string deviceName = "")
    {
        lock (_scanFileLock)
        {
            string path = GetScanFilePath();
            bool writeHeader = !File.Exists(path);

            using var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read);
            using var writer = new StreamWriter(stream, Encoding.UTF8);
            if (writeHeader)
            {
                writer.WriteLine("timestamp_iso,deviceName,barcode");
            }

            writer.WriteLine($"{DateTime.UtcNow:O},{EscapeCsv(deviceName)},{EscapeCsv(barcode)}");
            writer.Flush();
        }

        ScanReceived?.Invoke(this, new ScanReceivedEventArgs(barcode, DateTime.UtcNow, deviceName));
    }

    private static string EscapeCsv(string value)
    {
        if (value.Contains('"') || value.Contains(',') || value.Contains('\n') || value.Contains('\r'))
        {
            return '"' + value.Replace("\"", "\"\"") + '"';
        }

        return value;
    }

    private static string GetScanFilePath()
    {
        return Path.Combine(AppContext.BaseDirectory, ScanFileName);
    }

    private void PublishConnectionState()
    {
        var state = ConnectedClients switch
        {
            0 => ConnectionState.Offline,
            1 => ConnectionState.Ready,
            _ => ConnectionState.Busy
        };

        ConnectionState = state;
        ConnectionStatusChanged?.Invoke(this, new ConnectionStateChangedEventArgs(state, ConnectedClients));
    }

    private bool IsDeviceManuallyBlocked(string deviceName)
    {
        if (string.IsNullOrWhiteSpace(deviceName))
            return false;

        if (!_manuallyBlockedDevices.TryGetValue(deviceName, out var blockedUntilUtc))
            return false;

        if (DateTime.UtcNow <= blockedUntilUtc)
            return true;

        _manuallyBlockedDevices.TryRemove(deviceName, out _);
        return false;
    }

    private bool IsIpManuallyBlocked(string ip)
    {
        if (string.IsNullOrWhiteSpace(ip))
            return false;

        if (!_manuallyBlockedIps.TryGetValue(ip, out var blockedUntilUtc))
            return false;

        if (DateTime.UtcNow <= blockedUntilUtc)
            return true;

        _manuallyBlockedIps.TryRemove(ip, out _);
        return false;
    }

    private void PublishConnectedDevices()
    {
        var snapshot = _connectedDevices.Values
            .Select(s => new ConnectedDeviceInfo(s.DeviceName, s.HasScanned))
            .ToList();

        ConnectedDevicesChanged?.Invoke(this, new ConnectedDevicesChangedEventArgs(snapshot));
    }

    private static List<string> GetLanIpCandidates()
    {
        try
        {
            return NetworkInterface.GetAllNetworkInterfaces()
                .Where(n => n.OperationalStatus == OperationalStatus.Up && n.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .SelectMany(n => n.GetIPProperties().UnicastAddresses)
                .Where(a => a.Address.AddressFamily == AddressFamily.InterNetwork)
                .Select(a => a.Address.ToString())
                .Distinct()
                .ToList();
        }
        catch
        {
            return new List<string>();
        }
    }

    private static string? GetPrimaryLanIp()
    {
        return GetLanIpCandidates().FirstOrDefault();
    }

    // اگر IP فعلی عوض شده (و IP قبلی دیگر در دسترس نیست - تا وقتی همان آداپتور هنوز هست ثابت
    // بماند و بین دو آداپتور بالا/پایین نپرد) آن را به‌روز می‌کند و رویداد LanIpChanged را
    // صدا می‌زند تا QR اتصال دوباره ساخته شود.
    private void CheckLanIpChange()
    {
        try
        {
            var candidates = GetLanIpCandidates();
            if (candidates.Count == 0)
                return;

            string current = LanIp;
            if (candidates.Contains(current, StringComparer.Ordinal))
                return;

            LanIp = candidates[0];
            LanIpChanged?.Invoke(this, EventArgs.Empty);
        }
        catch { }
    }

    private static string ReadOrCreateComputerId(string folder)
    {
        string file = Path.Combine(folder, "computer-id.txt");
        if (File.Exists(file))
        {
            var existing = File.ReadAllText(file, Encoding.UTF8).Trim();
            if (!string.IsNullOrEmpty(existing))
            {
                return existing;
            }
        }

        string id = Guid.NewGuid().ToString();
        File.WriteAllText(file, id, Encoding.UTF8);
        return id;
    }

    private static string ReadOptionalComputerName(string folder)
    {
        string file = Path.Combine(folder, "computer-name.txt");
        if (File.Exists(file))
        {
            var existing = File.ReadAllText(file, Encoding.UTF8).Trim();
            if (!string.IsNullOrEmpty(existing))
            {
                return existing;
            }
        }

        return Environment.MachineName;
    }
}

public sealed class ScanEntry
{
    public ScanEntry(DateTime timestampUtc, string barcode)
    {
        TimestampUtc = timestampUtc;
        Barcode = barcode;
    }

    public DateTime TimestampUtc { get; }
    public string Barcode { get; }
    public string DisplayText => $"{TimestampUtc:HH:mm:ss} · {Barcode}";
}

public sealed class ScanReceivedEventArgs : EventArgs
{
    public ScanReceivedEventArgs(string barcode, DateTime timestampUtc, string deviceName = "")
    {
        Barcode = barcode;
        TimestampUtc = timestampUtc;
        DeviceName = deviceName;
    }

    public string Barcode { get; }
    public DateTime TimestampUtc { get; }
    public string DeviceName { get; }
}

public sealed class RemoteEntryValueEventArgs : EventArgs
{
    public RemoteEntryValueEventArgs(string barcode, string stepId, string value)
    {
        Barcode = barcode;
        StepId = stepId;
        Value = value;
    }

    public string Barcode { get; }
    public string StepId { get; }
    public string Value { get; }
}

public sealed class RemoteEntrySubmitEventArgs : EventArgs
{
    public RemoteEntrySubmitEventArgs(string barcode)
    {
        Barcode = barcode;
    }

    public string Barcode { get; }
}

public sealed class RemoteEntryBackEventArgs : EventArgs
{
    public RemoteEntryBackEventArgs(string barcode)
    {
        Barcode = barcode;
    }

    public string Barcode { get; }
}

public sealed class ConnectionStateChangedEventArgs : EventArgs
{
    public ConnectionStateChangedEventArgs(ConnectionState state, int connectedClients)
    {
        State = state;
        ConnectedClients = connectedClients;
    }

    public ConnectionState State { get; }
    public int ConnectedClients { get; }
}

public sealed class PeerDesktopSettingsEventArgs : EventArgs
{
    public PeerDesktopSettingsEventArgs(string payloadJson, long versionUtcMs, string fromComputerId)
    {
        PayloadJson = payloadJson;
        VersionUtcMs = versionUtcMs;
        FromComputerId = fromComputerId;
    }

    public string PayloadJson { get; }
    public long VersionUtcMs { get; }
    public string FromComputerId { get; }
}

public sealed class HighUsageBarcodeOperationEventArgs : EventArgs
{
    public HighUsageBarcodeOperationEventArgs(string opJson, string fromComputerId)
    {
        OpJson = opJson;
        FromComputerId = fromComputerId;
    }

    public string OpJson { get; }
    public string FromComputerId { get; }
}

public sealed class HighUsageBarcodeSnapshotEventArgs : EventArgs
{
    public HighUsageBarcodeSnapshotEventArgs(string payloadJson, long versionUtcMs, string fromComputerId)
    {
        PayloadJson = payloadJson;
        VersionUtcMs = versionUtcMs;
        FromComputerId = fromComputerId;
    }

    public string PayloadJson { get; }
    public long VersionUtcMs { get; }
    public string FromComputerId { get; }
}

public static class KeyboardInjector
{
    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint Type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT Mi;
        [FieldOffset(0)] public KEYBDINPUT Ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int Dx;
        public int Dy;
        public uint MouseData;
        public uint DwFlags;
        public uint Time;
        public IntPtr DwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort WVk;
        public ushort WScan;
        public uint DwFlags;
        public uint Time;
        public IntPtr DwExtraInfo;
    }

    private const uint InputKeyboard = 1;
    private const uint KeyeventfUnicode = 0x0004;
    private const uint KeyeventfKeyup = 0x0002;
    private const ushort VkReturn = 0x0D;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, [In] INPUT[] pInputs, int cbSize);

    public static void TypeText(string text)
    {
        foreach (char c in text)
        {
            var down = new INPUT
            {
                Type = InputKeyboard,
                U = new InputUnion { Ki = new KEYBDINPUT { WVk = 0, WScan = c, DwFlags = KeyeventfUnicode, Time = 0, DwExtraInfo = IntPtr.Zero } }
            };
            var up = new INPUT
            {
                Type = InputKeyboard,
                U = new InputUnion { Ki = new KEYBDINPUT { WVk = 0, WScan = c, DwFlags = KeyeventfUnicode | KeyeventfKeyup, Time = 0, DwExtraInfo = IntPtr.Zero } }
            };
            SendOne(down);
            SendOne(up);
        }
    }

    public static void PressEnter()
    {
        var down = new INPUT
        {
            Type = InputKeyboard,
            U = new InputUnion { Ki = new KEYBDINPUT { WVk = VkReturn, WScan = 0, DwFlags = 0, Time = 0, DwExtraInfo = IntPtr.Zero } }
        };
        var up = new INPUT
        {
            Type = InputKeyboard,
            U = new InputUnion { Ki = new KEYBDINPUT { WVk = VkReturn, WScan = 0, DwFlags = KeyeventfKeyup, Time = 0, DwExtraInfo = IntPtr.Zero } }
        };
        SendOne(down);
        SendOne(up);
    }

    public static void SendKeysFallback(string text)
    {
        SendKeys.SendWait(EscapeSendKeysText(text));
    }

    private static void SendOne(INPUT input)
    {
        uint size = (uint)Marshal.SizeOf<INPUT>();
        uint result = SendInput(1, new[] { input }, (int)size);
        if (result == 0)
        {
            int error = Marshal.GetLastWin32Error();
            throw new InvalidOperationException($"SendInput failed with Win32 error {error}.");
        }
    }

    private static string EscapeSendKeysText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        foreach (char c in text)
        {
            switch (c)
            {
                case '+':
                case '^':
                case '%':
                case '~':
                case '(':
                case ')':
                case '{':
                case '}':
                    builder.Append('{').Append(c).Append('}');
                    break;
                default:
                    builder.Append(c);
                    break;
            }
        }

        return builder.ToString();
    }
}