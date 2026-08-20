using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using ClosedXML.Excel;

// پروژه UseWindowsForms=true دارد که باعث می‌شود System.Windows.Forms و System.Drawing به‌صورت
// implicit global using در کل پروژه اضافه شوند - و چون این‌ها هم‌نام‌هایی دقیقاً مثل کنترل‌های
// WPF دارند (Button، TextBox، Color، Point، MouseEventArgs، Cursors، CheckBox)، بدون این alias‌ها،
// هر استفاده‌ی ساده از این نام‌ها با خطای CS0104 (ambiguous reference) مواجه می‌شود.
using Button = System.Windows.Controls.Button;
using TextBox = System.Windows.Controls.TextBox;
using CheckBox = System.Windows.Controls.CheckBox;
using Color = System.Windows.Media.Color;
using Point = System.Windows.Point;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using Cursors = System.Windows.Input.Cursors;
using Orientation = System.Windows.Controls.Orientation;

namespace ScanBridgeTest;

// =====================================================================================
// ویژگی «بارکد پرمصرف»: یک بانک از بارکدهایی که هر روز زیاد استفاده می‌شوند (مثلاً سرم‌ها یا
// آنتی‌بیوتیک‌ها)، به‌صورت گروه/زیرگروه. کاربر با اسکن گوشی، بارکدها را در صف یک زیرگروه ذخیره
// می‌کند؛ بعداً با یک آیکون شناور روی صفحه (که حتی وقتی برنامه‌ی نسخه‌نویسی دیگری فوکوس دارد هم
// دیده می‌شود) با یک کلیک، قدیمی‌ترین بارکد آن زیرگروه (FIFO) در همان برنامه‌ی بیرونی وارد می‌شود.
//
// این فایل عمداً از بقیه‌ی MainWindow.xaml.cs جدا نگه داشته شده (partial class) تا ریسک تغییر در
// فایل بزرگ و پرکاربرد اصلی به حداقل برسد. تمام رابط کاربری این ویژگی (پنجره‌ی مدیریت، پاپ‌آپ
// انتخاب زیرگروه، آیکون شناور) کاملاً با کد ساخته شده‌اند (بدون XAML جداگانه) - دقیقاً با همان
// روشی که ExpiryAlertWindow در MainWindow.xaml.cs از قبل استفاده می‌کند.
// =====================================================================================

public sealed class HighUsageBarcodeEntry
{
    // شناسه‌ی پایدار - برای هماهنگ‌سازی بین چند سیستم (تا هر سیستم بتواند دقیقاً همین بارکد را با
    // شناسه، نه با برابری مرجع/آبجکت که فقط داخل همان پردازه معنی دارد، پیدا کند).
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Barcode { get; set; } = string.Empty;
    public DateTime ScannedAtUtc { get; set; } = DateTime.UtcNow;
    // چند واحد از این بارکد هنوز استفاده نشده - برای بارکدهایی که روی جعبه هستند و چند واحد را
    // با هم پوشش می‌دهند (مثلاً یک جعبه‌ی ۲۰ ویالی). هر بار دیسپنس یکی کم می‌شود؛ وقتی به صفر
    // برسد، این بارکد از صف حذف و نوبت به بعدی می‌رسد. مقدار اولیه از UnitsPerBarcode زیرگروه در
    // لحظه‌ی اسکن گرفته می‌شود (نه لحظه‌ی مصرف) تا تغییر بعدی تنظیم زیرگروه، جعبه‌های قبلاً
    // اسکن‌شده را عقب برنگرداند.
    public int RemainingUses { get; set; } = 1;
}

public sealed class HighUsageBarcodeSubgroup
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = string.Empty;
    // چند واحد فیزیکی را یک بارکد پوشش می‌دهد. ۱ یعنی هر واحد بارکد خودش را دارد (مثلاً هر سرم).
    // بیشتر از ۱ یعنی یک بارکد (مثلاً روی جعبه) چند واحد را با هم پوشش می‌دهد (مثلاً جعبه‌ی
    // ۲۰ ویالی پنتوپرازول که فقط یک بارکد کلی روی جعبه دارد).
    public int UnitsPerBarcode { get; set; } = 1;
    public List<HighUsageBarcodeEntry> Entries { get; set; } = new();
}

public sealed class HighUsageBarcodeGroup
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = string.Empty;
    public List<HighUsageBarcodeSubgroup> Subgroups { get; set; } = new();
}

public sealed class HighUsageBarcodeSettings
{
    public bool WidgetEnabled { get; set; }
    public double WidgetLeft { get; set; } = -1;
    public double WidgetTop { get; set; } = -1;
}

// فایل ذخیره‌ی بانک روی دیسک این پوششِ خودش را دارد (نه فقط یک آرایه‌ی خام گروه‌ها) تا یک شماره‌ی
// نسخه‌ی زمانی هم همراهش ذخیره شود - همان چیزی که برای هماهنگ‌سازی اولیه (bootstrap) با سیستم‌های
// هم‌شبکه‌ی هم‌لایسنس لازم است (پایین‌تر توضیح داده شده).
public sealed class HighUsageBarcodeBankFile
{
    public long VersionUtcMs { get; set; }
    public List<HighUsageBarcodeGroup> Groups { get; set; } = new();
}

public partial class MainWindow
{
    private List<HighUsageBarcodeGroup> _highUsageGroups = new();
    // نسخه‌ی زمانی کل بانک - با هر تغییر محلی (اضافه/حذف/دیسپنس) و با هر عملیات دریافتی از یک
    // سیستم هم‌شبکه بالا می‌رود. فقط برای هماهنگ‌سازی اولیه‌ی یک سیستم تازه‌وصل‌شده استفاده می‌شود؛
    // خودِ به‌روزرسانی زنده از طریق عملیات‌های تکی (نه این نسخه) رد و بدل می‌شود - نگاه کنید به
    // توضیح بالای BroadcastHighUsageOperation.
    private long _highUsageBankVersionUtcMs;
    private HighUsageBarcodeSettings _highUsageSettings = new();
    private HighUsageBarcodeWidgetWindow? _highUsageWidgetWindow;
    private HighUsageBarcodeManagerWindow? _highUsageManagerWindow;
    private HighUsageBarcodePickerWindow? _highUsagePickerWindow;
    private string? _highUsageCaptureSubgroupId;
    private IntPtr _highUsageCapturedForegroundWindow = IntPtr.Zero;

    private string GetHighUsageBarcodeDataPath() => Path.Combine(AppContext.BaseDirectory, "high-usage-barcodes.dat");
    private string GetHighUsageBarcodeSettingsPath() => Path.Combine(AppContext.BaseDirectory, "high-usage-barcode-settings.dat");

    private void InitializeHighUsageBarcodeFeature()
    {
        LoadHighUsageBarcodeSettings();
        LoadHighUsageBarcodeGroups();

        try
        {
            if (HighUsageWidgetEnableCheckBox != null)
                HighUsageWidgetEnableCheckBox.IsChecked = _highUsageSettings.WidgetEnabled;
        }
        catch { }

        if (_highUsageSettings.WidgetEnabled)
        {
            try { ShowHighUsageWidget(); } catch { }
        }
    }

    private void LoadHighUsageBarcodeSettings()
    {
        try
        {
            string path = GetHighUsageBarcodeSettingsPath();
            if (File.Exists(path))
            {
                string json = File.ReadAllText(path);
                if (!string.IsNullOrWhiteSpace(json))
                    _highUsageSettings = JsonSerializer.Deserialize<HighUsageBarcodeSettings>(json) ?? new HighUsageBarcodeSettings();
            }
        }
        catch { _highUsageSettings = new HighUsageBarcodeSettings(); }
    }

    private void SaveHighUsageBarcodeSettings()
    {
        try
        {
            File.WriteAllText(GetHighUsageBarcodeSettingsPath(), JsonSerializer.Serialize(_highUsageSettings, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    private void LoadHighUsageBarcodeGroups()
    {
        try
        {
            string path = GetHighUsageBarcodeDataPath();
            if (File.Exists(path))
            {
                string json = File.ReadAllText(path);
                if (!string.IsNullOrWhiteSpace(json))
                {
                    // فرمت جدید: یک آبجکت با VersionUtcMs + Groups. برای سازگاری با فایل‌های قدیمی‌تر
                    // (که فقط یک آرایه‌ی خام گروه‌ها بودند، بدون نسخه)، اگر فرمت جدید جواب نداد یا
                    // Groups خالی برگشت در حالی که JSON با «[» شروع می‌شود، به فرمت قدیمی برمی‌گردیم.
                    bool looksLikeOldArrayFormat = json.TrimStart().StartsWith("[");
                    if (!looksLikeOldArrayFormat)
                    {
                        var wrapper = JsonSerializer.Deserialize<HighUsageBarcodeBankFile>(json);
                        if (wrapper != null)
                        {
                            _highUsageGroups = wrapper.Groups ?? new List<HighUsageBarcodeGroup>();
                            _highUsageBankVersionUtcMs = wrapper.VersionUtcMs;
                            return;
                        }
                    }

                    _highUsageGroups = JsonSerializer.Deserialize<List<HighUsageBarcodeGroup>>(json) ?? new List<HighUsageBarcodeGroup>();
                    _highUsageBankVersionUtcMs = 0;
                }
            }
        }
        catch { _highUsageGroups = new List<HighUsageBarcodeGroup>(); }
    }

    private void SaveHighUsageBarcodeGroups()
    {
        try
        {
            var wrapper = new HighUsageBarcodeBankFile { VersionUtcMs = _highUsageBankVersionUtcMs, Groups = _highUsageGroups };
            File.WriteAllText(GetHighUsageBarcodeDataPath(), JsonSerializer.Serialize(wrapper, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    // بعد از هر تغییر محلیِ بانک (اضافه/حذف/دیسپنس) صدا زده می‌شود: نسخه را بالا می‌برد، ذخیره
    // می‌کند، و یک عکسِ کامل از کل بانک را برای سیستم‌های هم‌شبکه‌ی هم‌لایسنس آماده‌ی درخواست
    // نگه می‌دارد (برای هماهنگ‌سازی اولیه‌ی یک سیستم تازه‌وصل‌شده - نگاه کنید به BroadcastHighUsageOperation
    // برای بخش زنده/آنی‌ که با عملیات‌های تکی انجام می‌شود، نه با این عکسِ کامل).
    private void PublishHighUsageBankSnapshotForSync()
    {
        _highUsageBankVersionUtcMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        SaveHighUsageBarcodeGroups();
        try
        {
            string json = JsonSerializer.Serialize(_highUsageGroups);
            _service?.PublishHighUsageBarcodeSnapshot(json, _highUsageBankVersionUtcMs);
        }
        catch { }
    }

    // هر بار که بانک به‌صورت محلی تغییر می‌کند (اضافه/حذف/دیسپنس)، این تابع به‌جای فقط ذخیره‌ی روی
    // دیسک، یک «عملیات» کوچک (نه کل بانک) برای سیستم‌های هم‌شبکه‌ی هم‌لایسنس هم broadcast می‌کند -
    // چون فقط همان یک واحد تغییر رد و بدل می‌شود (نه کل لیست)، اعمال آن روی سیستم‌های دیگر تقریباً
    // آنی است؛ این یعنی اگر دو سیستم/دو گوشی هم‌زمان با هم کار کنند، هر دو تقریباً بلافاصله از کار
    // هم باخبر می‌شوند (مثلاً یک بارکد دیسپنس‌شده روی یک سیستم، سریع از صف سیستم دیگر هم کم می‌شود
    // و دوباره دیسپنس نمی‌شود). علاوه بر broadcast عملیات، یک نسخه‌ی snapshot هم برای هماهنگ‌سازی
    // اولیه (bootstrap یک سیستم تازه‌وصل‌شده یا تازه از آفلاین برگشته) ذخیره و آماده‌ی درخواست
    // نگه داشته می‌شود.
    private void BroadcastHighUsageOperation(object opPayload)
    {
        try
        {
            string opJson = JsonSerializer.Serialize(opPayload);
            _service?.BroadcastHighUsageBarcodeOperation(opJson);
        }
        catch { }
        PublishHighUsageBankSnapshotForSync();
    }

    private static string GetHighUsageOpString(JsonElement root, string propertyName)
        => root.TryGetProperty(propertyName, out var prop) ? prop.GetString() ?? string.Empty : string.Empty;

    // یک عملیات دریافتی از یک سیستم هم‌شبکه را روی بانک محلی اعمال می‌کند - بدون broadcast دوباره
    // (تا حلقه‌ی بی‌پایان رد و بدل پیام پیش نیاید). چون هر سیستم شناسه‌ی پایدار گروه/زیرگروه/رکورد
    // را می‌فرستد (نه اندیس)، اعمال آن مستقل از ترتیب رسیدن پیام‌ها روی همه‌ی سیستم‌ها یکسان جواب
    // می‌دهد.
    private void ApplyHighUsageBarcodeOperation(string opJson, string fromComputerId)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(new Action(() => ApplyHighUsageBarcodeOperation(opJson, fromComputerId)));
            return;
        }

        if (!string.IsNullOrEmpty(fromComputerId) && _service != null && string.Equals(fromComputerId, _service.ComputerId, StringComparison.OrdinalIgnoreCase))
            return;

        try
        {
            using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(opJson) ? "{}" : opJson);
            var root = doc.RootElement;
            string kind = GetHighUsageOpString(root, "kind");
            bool changed = false;

            switch (kind)
            {
                case "addGroup":
                    {
                        string groupId = GetHighUsageOpString(root, "groupId");
                        if (!string.IsNullOrEmpty(groupId) && !_highUsageGroups.Any(g => g.Id == groupId))
                        {
                            _highUsageGroups.Add(new HighUsageBarcodeGroup { Id = groupId, Name = GetHighUsageOpString(root, "name") });
                            changed = true;
                        }
                        break;
                    }
                case "addSubgroup":
                    {
                        string groupId = GetHighUsageOpString(root, "groupId");
                        string subgroupId = GetHighUsageOpString(root, "subgroupId");
                        int units = root.TryGetProperty("unitsPerBarcode", out var unitsProp) && unitsProp.TryGetInt32(out var u) ? Math.Max(1, u) : 1;
                        var group = _highUsageGroups.FirstOrDefault(g => g.Id == groupId);
                        if (group != null && !string.IsNullOrEmpty(subgroupId) && !group.Subgroups.Any(s => s.Id == subgroupId))
                        {
                            group.Subgroups.Add(new HighUsageBarcodeSubgroup { Id = subgroupId, Name = GetHighUsageOpString(root, "name"), UnitsPerBarcode = units });
                            changed = true;
                        }
                        break;
                    }
                case "deleteGroup":
                    {
                        string groupId = GetHighUsageOpString(root, "groupId");
                        if (_highUsageGroups.RemoveAll(g => g.Id == groupId) > 0)
                        {
                            if (_highUsageCaptureSubgroupId != null && FindHighUsageSubgroup(_highUsageCaptureSubgroupId) == null)
                                _highUsageCaptureSubgroupId = null;
                            changed = true;
                        }
                        break;
                    }
                case "deleteSubgroup":
                    {
                        string groupId = GetHighUsageOpString(root, "groupId");
                        string subgroupId = GetHighUsageOpString(root, "subgroupId");
                        var group = _highUsageGroups.FirstOrDefault(g => g.Id == groupId);
                        if (group != null && group.Subgroups.RemoveAll(s => s.Id == subgroupId) > 0)
                        {
                            if (_highUsageCaptureSubgroupId == subgroupId)
                                _highUsageCaptureSubgroupId = null;
                            changed = true;
                        }
                        break;
                    }
                case "addEntry":
                    {
                        string subgroupId = GetHighUsageOpString(root, "subgroupId");
                        string entryId = GetHighUsageOpString(root, "entryId");
                        var subgroup = FindHighUsageSubgroup(subgroupId);
                        if (subgroup != null && !string.IsNullOrEmpty(entryId) && !subgroup.Entries.Any(x => x.Id == entryId))
                        {
                            DateTime scannedAtUtc = root.TryGetProperty("scannedAtUtc", out var scannedProp) && scannedProp.TryGetDateTime(out var dt) ? dt : DateTime.UtcNow;
                            int remaining = root.TryGetProperty("remainingUses", out var remProp) && remProp.TryGetInt32(out var r) ? Math.Max(0, r) : 1;
                            subgroup.Entries.Add(new HighUsageBarcodeEntry { Id = entryId, Barcode = GetHighUsageOpString(root, "barcode"), ScannedAtUtc = scannedAtUtc, RemainingUses = remaining });
                            changed = true;
                        }
                        break;
                    }
                case "deleteEntry":
                    {
                        string subgroupId = GetHighUsageOpString(root, "subgroupId");
                        string entryId = GetHighUsageOpString(root, "entryId");
                        var subgroup = FindHighUsageSubgroup(subgroupId);
                        if (subgroup != null && subgroup.Entries.RemoveAll(x => x.Id == entryId) > 0)
                            changed = true;
                        break;
                    }
                case "dispenseEntry":
                    {
                        string subgroupId = GetHighUsageOpString(root, "subgroupId");
                        string entryId = GetHighUsageOpString(root, "entryId");
                        int remainingAfter = root.TryGetProperty("remainingUsesAfter", out var raProp) && raProp.TryGetInt32(out var ra) ? Math.Max(0, ra) : 0;
                        var subgroup = FindHighUsageSubgroup(subgroupId);
                        var entry = subgroup?.Entries.FirstOrDefault(x => x.Id == entryId);
                        if (entry != null)
                        {
                            if (remainingAfter <= 0)
                                subgroup!.Entries.Remove(entry);
                            else
                                entry.RemainingUses = remainingAfter;
                            changed = true;
                        }
                        break;
                    }
            }

            if (changed)
            {
                SaveHighUsageBarcodeGroups();
                try { _highUsageManagerWindow?.RefreshFromSource(); } catch { }
            }
        }
        catch { }
    }

    // یک عکسِ کامل از بانک یک سیستم هم‌شبکه (برای هماهنگ‌سازی اولیه‌ی یک سیستم تازه‌وصل‌شده یا در
    // جواب درخواست ما). فقط additive است - یعنی هیچ‌وقت رکوردی که همین‌جا محلی وجود دارد را حذف یا
    // جایگزین نمی‌کند، فقط چیزی که اینجا نیست (بر اساس شناسه) را اضافه می‌کند؛ این‌طور اگر یک سیستم
    // مدتی آفلاین بوده و بعد وصل شده، چیزی که کاربر همین سیستم محلی زده از بین نمی‌رود. این طراحی
    // عمداً یک CRDT کامل نیست - یعنی حذف‌هایی که وقتی یک سیستم آفلاین بوده روی سیستم دیگر انجام
    // شده، با این snapshot به آن سیستم برنمی‌گردد (رکورد حذف‌شده روی سیستم دیگر دوباره از snapshot
    // این سیستم اضافه می‌شود). برای همین بخش زنده/آنلاین (BroadcastHighUsageOperation) طراحی اصلی
    // است و این فقط برای «کم نیفتادن چیزی» موقع اتصال دوباره است، نه یک تضمین ریاضی کامل.
    private void ApplyHighUsageBarcodeSnapshot(string payloadJson, long versionUtcMs, string fromComputerId)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(new Action(() => ApplyHighUsageBarcodeSnapshot(payloadJson, versionUtcMs, fromComputerId)));
            return;
        }

        if (!string.IsNullOrEmpty(fromComputerId) && _service != null && string.Equals(fromComputerId, _service.ComputerId, StringComparison.OrdinalIgnoreCase))
            return;

        try
        {
            var remoteGroups = JsonSerializer.Deserialize<List<HighUsageBarcodeGroup>>(string.IsNullOrWhiteSpace(payloadJson) ? "[]" : payloadJson) ?? new List<HighUsageBarcodeGroup>();
            bool changed = false;

            foreach (var remoteGroup in remoteGroups)
            {
                var localGroup = _highUsageGroups.FirstOrDefault(g => g.Id == remoteGroup.Id);
                if (localGroup == null)
                {
                    _highUsageGroups.Add(remoteGroup);
                    changed = true;
                    continue;
                }

                foreach (var remoteSubgroup in remoteGroup.Subgroups)
                {
                    var localSubgroup = localGroup.Subgroups.FirstOrDefault(s => s.Id == remoteSubgroup.Id);
                    if (localSubgroup == null)
                    {
                        localGroup.Subgroups.Add(remoteSubgroup);
                        changed = true;
                        continue;
                    }

                    foreach (var remoteEntry in remoteSubgroup.Entries)
                    {
                        if (!localSubgroup.Entries.Any(e => e.Id == remoteEntry.Id))
                        {
                            localSubgroup.Entries.Add(remoteEntry);
                            changed = true;
                        }
                    }
                }
            }

            if (versionUtcMs > _highUsageBankVersionUtcMs)
                _highUsageBankVersionUtcMs = versionUtcMs;

            if (changed)
            {
                SaveHighUsageBarcodeGroups();
                try { _highUsageManagerWindow?.RefreshFromSource(); } catch { }
            }
        }
        catch { }
    }

    private void HighUsageWidgetEnableCheckBox_Checked(object sender, RoutedEventArgs e)
    {
        _highUsageSettings.WidgetEnabled = true;
        SaveHighUsageBarcodeSettings();
        ShowHighUsageWidget();
    }

    private void HighUsageWidgetEnableCheckBox_Unchecked(object sender, RoutedEventArgs e)
    {
        _highUsageSettings.WidgetEnabled = false;
        SaveHighUsageBarcodeSettings();
        HideHighUsageWidget();
    }

    private void ShowHighUsageWidget()
    {
        if (_highUsageWidgetWindow != null)
        {
            try { _highUsageWidgetWindow.Show(); } catch { }
            return;
        }

        _highUsageWidgetWindow = new HighUsageBarcodeWidgetWindow(_highUsageSettings.WidgetLeft, _highUsageSettings.WidgetTop);
        _highUsageWidgetWindow.WidgetClicked += (_, _) => OnHighUsageWidgetClicked();
        _highUsageWidgetWindow.WidgetMoved += (_, _) =>
        {
            if (_highUsageWidgetWindow == null) return;
            _highUsageSettings.WidgetLeft = _highUsageWidgetWindow.Left;
            _highUsageSettings.WidgetTop = _highUsageWidgetWindow.Top;
            SaveHighUsageBarcodeSettings();
        };
        // دکمه‌ی ✕ روی آیکون شناور دقیقاً همان مسیر خاموش‌کردن تیک تنظیمات را طی می‌کند تا وضعیت
        // (تیک صفحه‌ی اصلی + فایل تنظیمات) با بسته‌شدن آیکون هماهنگ بماند.
        _highUsageWidgetWindow.CloseRequested += (_, _) =>
        {
            if (HighUsageWidgetEnableCheckBox != null)
                HighUsageWidgetEnableCheckBox.IsChecked = false;
            else
            {
                _highUsageSettings.WidgetEnabled = false;
                SaveHighUsageBarcodeSettings();
                HideHighUsageWidget();
            }
        };
        try { _highUsageWidgetWindow.Show(); } catch { }
    }

    private void HideHighUsageWidget()
    {
        try { _highUsageWidgetWindow?.Close(); } catch { }
        _highUsageWidgetWindow = null;
        try { _highUsagePickerWindow?.Close(); } catch { }
        _highUsagePickerWindow = null;
    }

    private void HighUsageBarcodePanelButton_Click(object sender, RoutedEventArgs e)
    {
        OpenHighUsageManagerWindow();
    }

    internal void OpenHighUsageManagerWindow()
    {
        if (_highUsageManagerWindow == null)
        {
            _highUsageManagerWindow = new HighUsageBarcodeManagerWindow(this) { Owner = this };
            _highUsageManagerWindow.Closed += (_, _) => _highUsageManagerWindow = null;
        }

        _highUsageManagerWindow.RefreshFromSource();
        if (_highUsageManagerWindow.WindowState == System.Windows.WindowState.Minimized)
            _highUsageManagerWindow.WindowState = System.Windows.WindowState.Normal;
        _highUsageManagerWindow.Show();
        _highUsageManagerWindow.Activate();
    }

    // این متد وقتی آیکون شناور کلیک می‌شود صدا زده می‌شود. آیکون و پاپ‌آپ هر دو با استایل
    // «NoActivate» ساخته شده‌اند (نگاه کنید به HighUsageNativeInterop.ApplyNoActivate) یعنی
    // کلیک روی آن‌ها فوکوس ویندوز را از برنامه‌ی بیرونی (نرم‌افزار نسخه‌نویسی) نمی‌گیرد؛ پس
    // پنجره‌ای که همین الان فوکوس دارد همان پنجره‌ی هدف نهایی برای پیست بارکد است.
    private void OnHighUsageWidgetClicked()
    {
        _highUsageCapturedForegroundWindow = HighUsageNativeInterop.GetForegroundWindow();

        try { _highUsagePickerWindow?.Close(); } catch { }
        _highUsagePickerWindow = new HighUsageBarcodePickerWindow(this);
        _highUsagePickerWindow.Closed += (_, _) => _highUsagePickerWindow = null;

        if (_highUsageWidgetWindow != null)
        {
            var workArea = SystemParameters.WorkArea;
            double pickerWidth = _highUsagePickerWindow.Width;
            double left = _highUsageWidgetWindow.Left + _highUsageWidgetWindow.Width + 12;
            if (left + pickerWidth > workArea.Right)
                left = _highUsageWidgetWindow.Left - pickerWidth - 12;
            if (left < workArea.Left)
                left = Math.Max(workArea.Left, workArea.Right - pickerWidth - 12);

            double top = Math.Min(Math.Max(_highUsageWidgetWindow.Top, workArea.Top), Math.Max(workArea.Top, workArea.Bottom - 480));

            _highUsagePickerWindow.WindowStartupLocation = System.Windows.WindowStartupLocation.Manual;
            _highUsagePickerWindow.Left = left;
            _highUsagePickerWindow.Top = top;
        }

        _highUsagePickerWindow.Show();
    }

    // ---- CRUD گروه/زیرگروه ----

    internal List<HighUsageBarcodeGroup> HighUsageGroups => _highUsageGroups;

    internal string? HighUsageCaptureSubgroupId => _highUsageCaptureSubgroupId;

    internal void SetHighUsageCaptureTarget(string? subgroupId)
    {
        _highUsageCaptureSubgroupId = subgroupId;
    }

    internal HighUsageBarcodeGroup AddHighUsageGroup(string name)
    {
        var group = new HighUsageBarcodeGroup { Name = name.Trim() };
        _highUsageGroups.Add(group);
        BroadcastHighUsageOperation(new { kind = "addGroup", groupId = group.Id, name = group.Name });
        return group;
    }

    internal HighUsageBarcodeSubgroup? AddHighUsageSubgroup(string groupId, string name, int unitsPerBarcode = 1)
    {
        var group = _highUsageGroups.FirstOrDefault(g => g.Id == groupId);
        if (group == null) return null;
        var subgroup = new HighUsageBarcodeSubgroup { Name = name.Trim(), UnitsPerBarcode = Math.Max(1, unitsPerBarcode) };
        group.Subgroups.Add(subgroup);
        BroadcastHighUsageOperation(new { kind = "addSubgroup", groupId, subgroupId = subgroup.Id, name = subgroup.Name, unitsPerBarcode = subgroup.UnitsPerBarcode });
        return subgroup;
    }

    internal void DeleteHighUsageGroup(string groupId)
    {
        var group = _highUsageGroups.FirstOrDefault(g => g.Id == groupId);
        if (group != null && _highUsageCaptureSubgroupId != null && group.Subgroups.Any(s => s.Id == _highUsageCaptureSubgroupId))
            _highUsageCaptureSubgroupId = null;

        _highUsageGroups.RemoveAll(g => g.Id == groupId);
        BroadcastHighUsageOperation(new { kind = "deleteGroup", groupId });
    }

    internal void DeleteHighUsageSubgroup(string groupId, string subgroupId)
    {
        var group = _highUsageGroups.FirstOrDefault(g => g.Id == groupId);
        group?.Subgroups.RemoveAll(s => s.Id == subgroupId);
        if (_highUsageCaptureSubgroupId == subgroupId)
            _highUsageCaptureSubgroupId = null;
        BroadcastHighUsageOperation(new { kind = "deleteSubgroup", groupId, subgroupId });
    }

    internal void DeleteHighUsageEntry(string subgroupId, HighUsageBarcodeEntry entry)
    {
        var subgroup = FindHighUsageSubgroup(subgroupId);
        subgroup?.Entries.Remove(entry);
        BroadcastHighUsageOperation(new { kind = "deleteEntry", subgroupId, entryId = entry.Id });
    }

    internal HighUsageBarcodeSubgroup? FindHighUsageSubgroup(string subgroupId)
    {
        foreach (var g in _highUsageGroups)
        {
            var s = g.Subgroups.FirstOrDefault(x => x.Id == subgroupId);
            if (s != null) return s;
        }
        return null;
    }

    internal void ShowHighUsageMessage(string title, string message, bool isError = false)
    {
        try { ShowStyledMessage(title, message, isError); } catch { }
    }

    // اگر «حالت دریافت» برای یک زیرگروه فعال باشد، بارکد اسکن‌شده را به‌جای مسیر عادی
    // (تاریخچه/تی‌تک) به انتهای صف همان زیرگروه اضافه می‌کند و true برمی‌گرداند تا مسیر عادی
    // پردازش اسکن (ScanReceived) همان‌جا متوقف شود.
    private bool TryCaptureHighUsageBarcode(string barcode)
    {
        if (string.IsNullOrEmpty(_highUsageCaptureSubgroupId))
            return false;

        var subgroup = FindHighUsageSubgroup(_highUsageCaptureSubgroupId);
        if (subgroup == null)
        {
            _highUsageCaptureSubgroupId = null;
            return false;
        }

        int unitsPerBarcode = Math.Max(1, subgroup.UnitsPerBarcode);
        var newEntry = new HighUsageBarcodeEntry { Barcode = barcode, ScannedAtUtc = DateTime.UtcNow, RemainingUses = unitsPerBarcode };
        subgroup.Entries.Add(newEntry);
        BroadcastHighUsageOperation(new { kind = "addEntry", subgroupId = subgroup.Id, entryId = newEntry.Id, barcode = newEntry.Barcode, scannedAtUtc = newEntry.ScannedAtUtc, remainingUses = newEntry.RemainingUses });

        try { _highUsageManagerWindow?.RefreshFromSource(); } catch { }
        try
        {
            var record = new ScanRecord(DateTime.Now, barcode, subgroup.Name);
            string countText = HighUsageUi.FormatSubgroupCount(subgroup);
            ShowScanToast(record, true, $"اضافه شد به «{subgroup.Name}» ({countText})");
        }
        catch { }

        return true;
    }

    // قدیمی‌ترین بارکد صف (FIFO) یک زیرگروه را برمی‌دارد، فوکوس ویندوز را به همان پنجره‌ای که
    // قبل از باز شدن آیکون/پاپ‌آپ فعال بود برمی‌گرداند و بارکد را مستقیماً همان‌جا تایپ می‌کند
    // (با همان مکانیزم SendInput که در ScanBridgeService برای تایپ اسکن‌های عادی استفاده می‌شود)
    // و در پایان یک Enter می‌زند.
    internal async Task<(bool success, string barcode, int remaining, string subgroupName)> DispenseHighUsageBarcodeAsync(string subgroupId)
    {
        var subgroup = FindHighUsageSubgroup(subgroupId);
        if (subgroup == null || subgroup.Entries.Count == 0)
            return (false, string.Empty, 0, subgroup?.Name ?? string.Empty);

        // قدیمی‌ترین بارکد صف؛ اگر بیش از یک واحد باقی داشته باشد (بارکد جعبه‌ای/چندتایی)، فقط یک
        // واحد از آن کم می‌شود و خودِ بارکد در صف می‌ماند (همچنان قدیمی‌ترین است، پس دفعه‌ی بعد هم
        // همین انتخاب می‌شود) تا وقتی واحدهایش تمام شود؛ آن‌وقت از صف حذف و نوبت به بعدی می‌رسد.
        var oldest = subgroup.Entries.OrderBy(x => x.ScannedAtUtc).First();
        string dispensedBarcode = oldest.Barcode;
        string dispensedEntryId = oldest.Id;
        oldest.RemainingUses = Math.Max(0, oldest.RemainingUses - 1);
        int remainingUsesAfterDispense = oldest.RemainingUses;
        if (oldest.RemainingUses <= 0)
            subgroup.Entries.Remove(oldest);
        BroadcastHighUsageOperation(new { kind = "dispenseEntry", subgroupId = subgroup.Id, entryId = dispensedEntryId, remainingUsesAfter = remainingUsesAfterDispense });

        try
        {
            if (_highUsageCapturedForegroundWindow != IntPtr.Zero)
                HighUsageNativeInterop.ForceSetForegroundWindow(_highUsageCapturedForegroundWindow);
        }
        catch { }

        // کمی مکث تا فوکوس واقعاً به پنجره‌ی بیرونی برگردد، قبل از تایپ.
        await Task.Delay(150);

        try
        {
            KeyboardInjector.TypeText(dispensedBarcode);
            KeyboardInjector.PressEnter();
        }
        catch
        {
            try
            {
                System.Windows.Clipboard.SetText(dispensedBarcode);
                System.Windows.Forms.SendKeys.SendWait("^v");
                await Task.Delay(100);
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            }
            catch { }
        }

        try { _highUsageManagerWindow?.RefreshFromSource(); } catch { }

        int remainingUnits = subgroup.Entries.Sum(x => x.RemainingUses);
        return (true, dispensedBarcode, remainingUnits, subgroup.Name);
    }

    // خروجی اکسل بانک: یک شیت جدا برای هر زیرگروه (با همان الگوی ClosedXML که در بقیه‌ی خروجی‌های
    // اکسل برنامه استفاده شده). اگر dialogOwner داده شود، قبل از خروجی یک صفحه‌ی انتخاب گروه/زیرگروه
    // روی همان پنجره نشان داده می‌شود؛ اگر notify داده شود، پیغام‌ها به‌جای overlay داخل MainWindow
    // (که ممکن است پشت پنجره‌ی مدیریت پنهان بماند) از همان مسیر (مثلاً toast داخل پنجره‌ی فراخواننده)
    // نمایش داده می‌شوند.
    internal void ExportHighUsageBarcodesToExcel(Window? dialogOwner = null, Action<string, string, bool>? notify = null)
    {
        void Notify(string title, string message, bool isError = false)
        {
            if (notify != null) { try { notify(title, message, isError); } catch { } }
            else { try { ShowStyledMessage(title, message, isError); } catch { } }
        }

        if (_highUsageGroups.Count == 0 || _highUsageGroups.All(g => g.Subgroups.All(s => s.Entries.Count == 0)))
        {
            Notify(GetLocalizedNoDataTitle(), GetLocalizedNoDataMessage(), true);
            return;
        }

        HashSet<string>? selectedSubgroupIds = null;
        if (dialogOwner != null)
        {
            selectedSubgroupIds = HighUsageExportSelectionWindow.PromptSelection(dialogOwner, _highUsageGroups);
            if (selectedSubgroupIds == null)
                return; // کاربر انصراف داد

            if (selectedSubgroupIds.Count == 0)
            {
                Notify("خروجی اکسل", "هیچ زیرگروهی برای خروجی انتخاب نشده است.", true);
                return;
            }
        }

        var saveFileDialog = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"Scanbridge_HighUsageBarcodes_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx",
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            Title = "خروجی اکسل بانک بارکد پرمصرف"
        };
        if (saveFileDialog.ShowDialog() != true)
            return;

        try
        {
            using var workbook = new XLWorkbook();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var group in _highUsageGroups)
            {
                foreach (var sub in group.Subgroups)
                {
                    if (sub.Entries.Count == 0) continue;
                    if (selectedSubgroupIds != null && !selectedSubgroupIds.Contains(sub.Id)) continue;

                    string sheetName = SanitizeExcelSheetName($"{group.Name}-{sub.Name}", usedNames);
                    var ws = workbook.Worksheets.Add(sheetName);
                    string[] headers = { "ردیف", "گروه", "زیرگروه", "بارکد", "واحد در هر بارکد", "واحد باقی‌مانده از این بارکد", "تاریخ و ساعت ثبت" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = ws.Cell(1, i + 1);
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.FromArgb(234, 88, 12);
                        cell.Style.Font.FontColor = XLColor.White;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    int r = 2, idx = 1;
                    foreach (var entry in sub.Entries.OrderBy(x => x.ScannedAtUtc))
                    {
                        ws.Cell(r, 1).Value = idx++;
                        ws.Cell(r, 2).Value = group.Name;
                        ws.Cell(r, 3).Value = sub.Name;
                        ws.Cell(r, 4).Value = entry.Barcode;
                        ws.Cell(r, 5).Value = sub.UnitsPerBarcode;
                        ws.Cell(r, 6).Value = entry.RemainingUses;
                        ws.Cell(r, 7).Value = entry.ScannedAtUtc.ToLocalTime().ToString("yyyy/MM/dd HH:mm:ss");
                        r++;
                    }
                    ws.Columns().AdjustToContents();
                }
            }

            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add("بانک بارکد پرمصرف");

            workbook.SaveAs(saveFileDialog.FileName);
            Notify("خروجی اکسل", "فایل اکسل با موفقیت ذخیره شد.");
        }
        catch (Exception ex)
        {
            Notify("خطا", $"ذخیره فایل اکسل با خطا مواجه شد: {ex.Message}", true);
        }
    }

    private static string SanitizeExcelSheetName(string raw, HashSet<string> used)
    {
        string cleaned = new string((raw ?? "Sheet").Select(c => "[]:*?/\\".Contains(c) ? '-' : c).ToArray()).Trim();
        if (string.IsNullOrWhiteSpace(cleaned)) cleaned = "Sheet";
        if (cleaned.Length > 31) cleaned = cleaned.Substring(0, 31);

        string candidate = cleaned;
        int n = 1;
        while (used.Contains(candidate))
        {
            string suffix = $"_{n++}";
            int keep = Math.Max(1, Math.Min(cleaned.Length, 31 - suffix.Length));
            candidate = cleaned.Substring(0, keep) + suffix;
        }
        used.Add(candidate);
        return candidate;
    }
}

// =====================================================================================
// Win32 interop کمکی: هم برای این‌که آیکون شناور/پاپ‌آپ هرگز فوکوس ویندوز را از برنامه‌ی
// بیرونی نگیرند (WS_EX_NOACTIVATE)، و هم به‌عنوان یک شبکه‌ی ایمنی برای برگرداندن فوکوس در صورت
// نیاز (ForceSetForegroundWindow با ترفند شناخته‌شده‌ی AttachThreadInput).
// =====================================================================================
internal static class HighUsageNativeInterop
{
    private const int GwlExstyle = -20;
    private const int WsExNoactivate = 0x08000000;
    private const int WsExToolwindow = 0x00000080;

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    internal static void ApplyNoActivate(Window window)
    {
        var helper = new WindowInteropHelper(window);
        IntPtr handle = helper.Handle;
        if (handle == IntPtr.Zero) return;
        int exStyle = GetWindowLong(handle, GwlExstyle);
        SetWindowLong(handle, GwlExstyle, exStyle | WsExNoactivate | WsExToolwindow);
    }

    internal static bool ForceSetForegroundWindow(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
            return false;

        IntPtr foreground = GetForegroundWindow();
        if (foreground == hWnd)
            return true;

        uint foregroundThreadId = GetWindowThreadProcessId(foreground, out _);
        uint currentThreadId = GetCurrentThreadId();

        bool attached = false;
        try
        {
            if (foregroundThreadId != 0 && foregroundThreadId != currentThreadId)
                attached = AttachThreadInput(currentThreadId, foregroundThreadId, true);

            return SetForegroundWindow(hWnd);
        }
        finally
        {
            if (attached)
                AttachThreadInput(currentThreadId, foregroundThreadId, false);
        }
    }
}

// دکمه‌ی گردگوشه‌ی مشترک برای پنجره‌های این ویژگی - دقیقاً با همان منطق CreateRoundedButton که
// ExpiryAlertWindow در MainWindow.xaml.cs از قبل استفاده می‌کند (چون این پنجره‌ها هم XAML جدا
// ندارند و به Window.Resources اصلی دسترسی ندارند).
internal static class HighUsageUi
{
    internal static Button CreateRoundedButton(string content, Color background, Color foreground, double width, double height, double fontSize)
    {
        var button = new Button
        {
            Content = content,
            Height = height,
            FontSize = fontSize,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(foreground),
            Background = new SolidColorBrush(background),
            BorderThickness = new Thickness(0),
            Cursor = Cursors.Hand,
            HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalContentAlignment = System.Windows.VerticalAlignment.Center,
            // برای دکمه‌های کوچک مربعی/آیکونی (مثل ✕ یا 🗑 با عرض کم، مثلاً ۲۴-۲۸px)، همان پدینگ
            // افقی ۱۰ پیکسلیِ دکمه‌های متنی باعث می‌شد محتوا (ایموجی) در آن عرض کم بریده/جمع دیده
            // شود - دقیقاً همان چیزی که باعث می‌شد دکمه‌ی حذف گروه «خوب ننشیند». برای دکمه‌های کوچک
            // پدینگ خیلی کمتر است.
            Padding = (width > 0 && width <= 40) ? new Thickness(2, 0, 2, 0) : new Thickness(10, 0, 10, 0)
        };
        if (width > 0)
            button.Width = width;

        var template = new ControlTemplate(typeof(Button));
        var borderFactory = new FrameworkElementFactory(typeof(Border));
        borderFactory.Name = "Bd";
        borderFactory.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background")
        {
            RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent)
        });
        borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(Math.Min(14, height / 2)));
        borderFactory.SetValue(Border.SnapsToDevicePixelsProperty, true);

        var contentFactory = new FrameworkElementFactory(typeof(ContentPresenter));
        contentFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        contentFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        contentFactory.SetBinding(ContentPresenter.MarginProperty, new System.Windows.Data.Binding("Padding")
        {
            RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent)
        });
        borderFactory.AppendChild(contentFactory);
        template.VisualTree = borderFactory;

        var hoverTrigger = new Trigger { Property = Button.IsMouseOverProperty, Value = true };
        hoverTrigger.Setters.Add(new Setter(Border.OpacityProperty, 0.88) { TargetName = "Bd" });
        template.Triggers.Add(hoverTrigger);

        var pressTrigger = new Trigger { Property = Button.IsPressedProperty, Value = true };
        pressTrigger.Setters.Add(new Setter(Border.OpacityProperty, 0.75) { TargetName = "Bd" });
        template.Triggers.Add(pressTrigger);

        var disabledTrigger = new Trigger { Property = Button.IsEnabledProperty, Value = false };
        disabledTrigger.Setters.Add(new Setter(Border.OpacityProperty, 0.45) { TargetName = "Bd" });
        template.Triggers.Add(disabledTrigger);

        button.Template = template;
        return button;
    }

    internal static void SetButtonColors(Button button, Color background, Color foreground)
    {
        button.Background = new SolidColorBrush(background);
        button.Foreground = new SolidColorBrush(foreground);
    }

    // متن تعداد باقی‌مانده‌ی یک زیرگروه برای نمایش روی دکمه‌ها/هدرها. اگر هر بارکد فقط یک واحد
    // پوشش می‌دهد (حالت معمولِ قبلی)، فقط همان عدد ساده نشان داده می‌شود؛ اگر بیشتر از یک واحد
    // پوشش می‌دهد (بارکد جعبه‌ای)، هم تعداد بارکد (جعبه) هم مجموع واحدهای باقی‌مانده نشان داده می‌شود.
    internal static string FormatSubgroupCount(HighUsageBarcodeSubgroup sub)
    {
        int totalUnits = sub.Entries.Sum(e => e.RemainingUses);
        if (sub.UnitsPerBarcode <= 1)
            return $"{totalUnits} عدد";
        return $"{sub.Entries.Count} بارکد، {totalUnits} واحد";
    }
}

// =====================================================================================
// آیکون شناور: یک دایره‌ی کوچک، همیشه‌بالا (Topmost)، قابل‌کشیدن، که هرگز فوکوس ویندوز را از
// برنامه‌ی بیرونی نمی‌گیرد (WS_EX_NOACTIVATE). کلیک ساده (بدون جابه‌جایی) رویداد WidgetClicked
// را صدا می‌زند؛ کشیدن، رویداد WidgetMoved را (برای ذخیره‌ی موقعیت جدید).
// =====================================================================================
public class HighUsageBarcodeWidgetWindow : Window
{
    public event EventHandler? WidgetClicked;
    public event EventHandler? WidgetMoved;
    // با کلیک روی دکمه‌ی کوچک ✕ گوشه‌ی آیکون شناور صدا زده می‌شود (برای بستن/غیرفعال کردن آیکون).
    public event EventHandler? CloseRequested;

    private Point _dragStartPoint;
    private bool _isDragging;
    private bool _dragMoved;

    // اندازه‌ی خودِ آیکون ۶۰x۶۰ است؛ ولی پنجره کمی بزرگ‌تر (۶۸x۶۸) گرفته می‌شود و آیکون به گوشه‌ی
    // پایین-چپ آن می‌چسبد تا فضای گوشه‌ی بالا-راست برای دکمه‌ی ✕ آزاد بماند - وگرنه چون خودِ پنجره
    // یک AllowsTransparency=True با اندازه‌ی دقیق ۶۰x۶۰ بود، هر چیزی با مارجین منفی (بیرون از آن
    // ۶۰x۶۰) توسط خودِ پنجره برش می‌خورد و ناقص/بدشکل دیده می‌شد (باگ گزارش‌شده).
    private const double IconSize = 60;
    private const double WindowPadding = 8;
    private const double WidgetWindowSize = IconSize + WindowPadding;

    public HighUsageBarcodeWidgetWindow(double savedLeft, double savedTop)
    {
        Width = WidgetWindowSize;
        Height = WidgetWindowSize;
        WindowStyle = System.Windows.WindowStyle.None;
        AllowsTransparency = true;
        Background = System.Windows.Media.Brushes.Transparent;
        ResizeMode = System.Windows.ResizeMode.NoResize;
        ShowInTaskbar = false;
        Topmost = true;
        WindowStartupLocation = System.Windows.WindowStartupLocation.Manual;

        var workArea = SystemParameters.WorkArea;
        Left = (savedLeft >= workArea.Left && savedLeft <= workArea.Right - Width) ? savedLeft : workArea.Right - Width - 24;
        Top = (savedTop >= workArea.Top && savedTop <= workArea.Bottom - Height) ? savedTop : workArea.Bottom - Height - 24;

        // مربعی (نه دایره‌ای) با گوشه‌های کمی گرد، هم‌استایل بقیه‌ی عناصر گردگوشه‌ی برنامه.
        var border = new Border
        {
            Width = IconSize,
            Height = IconSize,
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(Color.FromRgb(0xEA, 0x58, 0x0C)),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
            VerticalAlignment = System.Windows.VerticalAlignment.Bottom,
            // طراحی «فلت» طبق درخواست کاربر - بدون سایه/هاله‌ی تیره‌ی سه‌بعدی دور آیکون.
            Cursor = Cursors.Hand
        };

        // به‌جای ایموجی، از لوگوی خود برنامه (Assets/app-icon.ico) استفاده می‌شود.
        var logoImage = new System.Windows.Controls.Image
        {
            Width = 36,
            Height = 36,
            Stretch = System.Windows.Media.Stretch.Uniform,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        try
        {
            logoImage.Source = new System.Windows.Media.Imaging.BitmapImage(
                new Uri("pack://application:,,,/Assets/app-icon.ico", UriKind.Absolute));
        }
        catch { }
        border.Child = logoImage;

        // دکمه‌ی کوچک ✕ گوشه‌ی بالا-راست، برای بستن آیکون شناور - هم‌استایل بقیه‌ی دکمه‌های گردگوشه‌ی
        // این ویژگی (HighUsageUi.CreateRoundedButton). فرزندِ هم‌سطحِ border داخل همان Grid است (نه
        // فرزند خودِ border) تا کلیک روی آن با منطق کشیدن/کلیکِ آیکون اصلی تداخل نکند؛ کاملاً داخل
        // محدوده‌ی پنجره (بدون مارجین منفی) تا هرگز برش نخورد؛ و طبق درخواست کاربر، فقط وقتی موس
        // روی آیکون شناور می‌رود نمایان می‌شود (نه همیشه).
        var closeButton = HighUsageUi.CreateRoundedButton("✕", Color.FromRgb(0x37, 0x41, 0x51), Colors.White, 20, 20, 11);
        closeButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Right;
        closeButton.VerticalAlignment = System.Windows.VerticalAlignment.Top;
        closeButton.Margin = new Thickness(0);
        closeButton.ToolTip = "بستن";
        closeButton.Visibility = System.Windows.Visibility.Collapsed;
        closeButton.Click += (_, _) => CloseRequested?.Invoke(this, EventArgs.Empty);

        var root = new Grid { Width = WidgetWindowSize, Height = WidgetWindowSize };
        root.Children.Add(border);
        root.Children.Add(closeButton);
        Content = root;

        // ✕ فقط وقتی موس روی کل آیکون شناور (خودِ آیکون یا دکمه) است نشان داده شود؛ در حالت عادی
        // مخفی است. IsMouseOver روی Grid تا وقتی موس روی هر کدام از فرزندانش باشد true می‌ماند، پس
        // بردن موس از روی آیکون به روی خودِ دکمه‌ی ✕ باعث مخفی‌شدنش نمی‌شود.
        root.MouseEnter += (_, _) => closeButton.Visibility = System.Windows.Visibility.Visible;
        root.MouseLeave += (_, _) => closeButton.Visibility = System.Windows.Visibility.Collapsed;

        border.PreviewMouseLeftButtonDown += Border_PreviewMouseLeftButtonDown;
        border.PreviewMouseMove += Border_PreviewMouseMove;
        border.PreviewMouseLeftButtonUp += Border_PreviewMouseLeftButtonUp;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        try { HighUsageNativeInterop.ApplyNoActivate(this); } catch { }
    }

    private void Border_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _dragStartPoint = e.GetPosition(this);
        _isDragging = true;
        _dragMoved = false;
        (sender as UIElement)?.CaptureMouse();
    }

    private void Border_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (!_isDragging || e.LeftButton != MouseButtonState.Pressed) return;
        var pos = e.GetPosition(this);
        double dx = pos.X - _dragStartPoint.X;
        double dy = pos.Y - _dragStartPoint.Y;
        if (!_dragMoved && (Math.Abs(dx) > 4 || Math.Abs(dy) > 4))
            _dragMoved = true;
        if (_dragMoved)
        {
            Left += dx;
            Top += dy;
        }
    }

    private void Border_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isDragging) return;
        _isDragging = false;
        (sender as UIElement)?.ReleaseMouseCapture();
        if (_dragMoved)
            WidgetMoved?.Invoke(this, EventArgs.Empty);
        else
            WidgetClicked?.Invoke(this, EventArgs.Empty);
    }
}

// =====================================================================================
// پاپ‌آپ انتخاب زیرگروه: با کلیک روی آیکون شناور باز می‌شود؛ برای هر گروه، دکمه‌های زیرگروه
// (با تعداد باقی‌مانده) نشان می‌دهد. کلیک روی یک زیرگروه، دیسپنس (FIFO) را انجام می‌دهد و
// پاپ‌آپ بسته می‌شود.
// =====================================================================================
public class HighUsageBarcodePickerWindow : Window
{
    private readonly MainWindow _owner;
    private readonly StackPanel _groupsPanel;
    // null یعنی «لیست گروه‌ها» نمایش داده شود؛ غیر null یعنی داخل آن گروه هستیم و فقط
    // زیرگروه‌های همان گروه (به‌همراه دکمه‌ی بازگشت) نشان داده می‌شود.
    private string? _currentGroupId;

    public HighUsageBarcodePickerWindow(MainWindow owner)
    {
        _owner = owner;

        Width = 420;
        SizeToContent = System.Windows.SizeToContent.Height;
        MaxHeight = 560;
        WindowStyle = System.Windows.WindowStyle.None;
        AllowsTransparency = true;
        Background = System.Windows.Media.Brushes.Transparent;
        ResizeMode = System.Windows.ResizeMode.NoResize;
        ShowInTaskbar = false;
        Topmost = true;
        FlowDirection = System.Windows.FlowDirection.RightToLeft;

        var card = new Border
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(20),
            Padding = new Thickness(16),
            Margin = new Thickness(20),
            Effect = new DropShadowEffect { Color = Colors.Black, Opacity = 0.3, BlurRadius = 22, ShadowDepth = 5 }
        };

        var root = new DockPanel();

        var headerRow = new Grid { Margin = new Thickness(0, 0, 0, 10) };
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var title = new TextBlock
        {
            Text = "📦 بانک بارکد پرمصرف",
            FontSize = 15,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0xC2, 0x41, 0x0C)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        Grid.SetColumn(title, 0);
        headerRow.Children.Add(title);
        var closeBtn = HighUsageUi.CreateRoundedButton("✕", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 28, 28, 12);
        closeBtn.Click += (_, _) => Close();
        Grid.SetColumn(closeBtn, 1);
        headerRow.Children.Add(closeBtn);
        DockPanel.SetDock(headerRow, Dock.Top);
        root.Children.Add(headerRow);

        var manageBtn = HighUsageUi.CreateRoundedButton("⚙ مدیریت کامل بانک", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 0, 34, 12);
        manageBtn.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
        manageBtn.Margin = new Thickness(0, 0, 0, 10);
        manageBtn.Click += (_, _) => { _owner.OpenHighUsageManagerWindow(); Close(); };
        DockPanel.SetDock(manageBtn, Dock.Top);
        root.Children.Add(manageBtn);

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto, MaxHeight = 420 };
        _groupsPanel = new StackPanel();
        scroll.Content = _groupsPanel;
        root.Children.Add(scroll);

        card.Child = root;
        Content = card;

        BuildGroups();
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        try { HighUsageNativeInterop.ApplyNoActivate(this); } catch { }
    }

    private void BuildGroups()
    {
        _groupsPanel.Children.Clear();
        var groups = _owner.HighUsageGroups.Where(g => g.Subgroups.Count > 0).ToList();

        if (groups.Count == 0)
        {
            _groupsPanel.Children.Add(new TextBlock
            {
                Text = "هنوز گروه/زیرگروهی نساخته‌اید. از «مدیریت کامل بانک» یک گروه و زیرگروه بسازید و چند بارکد در آن ذخیره کنید.",
                TextWrapping = TextWrapping.Wrap,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
                Margin = new Thickness(4, 8, 4, 8)
            });
            return;
        }

        if (_currentGroupId == null)
        {
            BuildGroupList(groups);
            return;
        }

        var currentGroup = groups.FirstOrDefault(g => g.Id == _currentGroupId);
        if (currentGroup == null)
        {
            _currentGroupId = null;
            BuildGroupList(groups);
            return;
        }
        BuildSubgroupList(currentGroup);
    }

    // فقط لیست گروه‌ها (کلیک اول) - طبق درخواست کاربر.
    private void BuildGroupList(List<HighUsageBarcodeGroup> groups)
    {
        foreach (var group in groups)
        {
            var btn = HighUsageUi.CreateRoundedButton($"{group.Name}  ({group.Subgroups.Count} زیرگروه)", Color.FromRgb(0xEA, 0x58, 0x0C), Colors.White, 0, 46, 13);
            btn.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            btn.Margin = new Thickness(0, 0, 0, 8);
            string groupId = group.Id;
            btn.Click += (_, _) => { _currentGroupId = groupId; BuildGroups(); };
            _groupsPanel.Children.Add(btn);
        }
    }

    // فقط زیرگروه‌های همان گروه به‌همراه دکمه‌ی بازگشت (کلیک روی یک گروه) - لیست در صورت زیاد
    // بودن، از طریق همان ScrollViewer بیرونی قابل اسکرول است.
    private void BuildSubgroupList(HighUsageBarcodeGroup group)
    {
        var backBtn = HighUsageUi.CreateRoundedButton("⬅ بازگشت به گروه‌ها", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 0, 36, 12);
        backBtn.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
        backBtn.Margin = new Thickness(0, 0, 0, 10);
        backBtn.Click += (_, _) => { _currentGroupId = null; BuildGroups(); };
        _groupsPanel.Children.Add(backBtn);

        var groupHeader = new TextBlock
        {
            Text = group.Name,
            FontSize = 13,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0xC2, 0x41, 0x0C)),
            Margin = new Thickness(2, 0, 2, 6)
        };
        _groupsPanel.Children.Add(groupHeader);

        var wrap = new WrapPanel();
        foreach (var sub in group.Subgroups)
        {
            var btn = HighUsageUi.CreateRoundedButton($"{sub.Name}  ({HighUsageUi.FormatSubgroupCount(sub)})", Color.FromRgb(0x0E, 0xA5, 0xE9), Colors.White, 168, 44, 13);
            btn.Margin = new Thickness(0, 0, 8, 8);
            btn.IsEnabled = sub.Entries.Count > 0;
            if (sub.Entries.Count == 0)
                btn.Opacity = 0.5;
            string subgroupId = sub.Id;
            btn.Click += async (s, _) => await OnSubgroupClickedAsync((Button)s, subgroupId);
            wrap.Children.Add(btn);
        }
        _groupsPanel.Children.Add(wrap);
    }

    private async Task OnSubgroupClickedAsync(Button btn, string subgroupId)
    {
        btn.IsEnabled = false;
        try
        {
            var result = await _owner.DispenseHighUsageBarcodeAsync(subgroupId);
            if (!result.success)
            {
                _owner.ShowHighUsageMessage("بارکد پرمصرف", "بارکدی در این زیرگروه باقی نمانده است.", true);
                btn.IsEnabled = true;
                return;
            }
            Close();
        }
        catch
        {
            btn.IsEnabled = true;
        }
    }
}

// =====================================================================================
// پنجره‌ی مدیریت کامل بانک: ساخت/حذف گروه و زیرگروه، فعال/غیرفعال‌کردن «دریافت بارکد» برای یک
// زیرگروه، دیدن/حذف بارکدهای هر زیرگروه، و خروجی اکسل کل بانک.
// =====================================================================================
public class HighUsageBarcodeManagerWindow : Window
{
    private readonly MainWindow _owner;
    private readonly StackPanel _groupsPanel;
    private readonly StackPanel _entriesPanel;
    private readonly TextBlock _entriesHeaderText;
    private readonly TextBlock _remainingCountText;
    private readonly Button _captureToggleButton;
    private readonly Border _toastBorder;
    private readonly TextBlock _toastText;
    private DispatcherTimer? _toastTimer;
    private string? _selectedSubgroupId;

    public HighUsageBarcodeManagerWindow(MainWindow owner)
    {
        _owner = owner;

        Title = "بانک بارکد پرمصرف";
        Width = 920;
        Height = 640;
        MinWidth = 720;
        MinHeight = 480;
        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
        FlowDirection = System.Windows.FlowDirection.RightToLeft;
        Background = new SolidColorBrush(Color.FromRgb(0xF3, 0xF4, 0xF6));

        var outerGrid = new Grid { Margin = new Thickness(16) };
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var headerGrid = new Grid { Margin = new Thickness(0, 0, 0, 12) };
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var titleText = new TextBlock
        {
            Text = "📦 بانک بارکد پرمصرف",
            FontSize = 20,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0xC2, 0x41, 0x0C)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        Grid.SetColumn(titleText, 0);
        headerGrid.Children.Add(titleText);

        var exportBtn = HighUsageUi.CreateRoundedButton("⬇ خروجی اکسل", Color.FromRgb(0x10, 0xB9, 0x81), Colors.White, 0, 40, 13);
        exportBtn.Click += (_, _) => _owner.ExportHighUsageBarcodesToExcel(this, (t, m, isErr) => ShowToast(t, m, isErr));
        Grid.SetColumn(exportBtn, 1);
        headerGrid.Children.Add(exportBtn);

        Grid.SetRow(headerGrid, 0);
        outerGrid.Children.Add(headerGrid);

        var bodyGrid = new Grid();
        bodyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        bodyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
        bodyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.3, GridUnitType.Star) });
        Grid.SetRow(bodyGrid, 1);
        outerGrid.Children.Add(bodyGrid);

        // ستون راست (چون RightToLeft است، ستون ۰ همان سمتی است که بصری «راست» دیده می‌شود): گروه‌ها/زیرگروه‌ها
        var leftCard = new Border { Background = System.Windows.Media.Brushes.White, CornerRadius = new CornerRadius(16), Padding = new Thickness(14) };
        Grid.SetColumn(leftCard, 0);
        var leftStack = new DockPanel();
        var leftTitle = new TextBlock
        {
            Text = "گروه‌ها و زیرگروه‌ها",
            FontSize = 15,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
            Margin = new Thickness(0, 0, 0, 10)
        };
        DockPanel.SetDock(leftTitle, Dock.Top);
        leftStack.Children.Add(leftTitle);

        var addGroupBtn = HighUsageUi.CreateRoundedButton("+ گروه جدید", Color.FromRgb(0xEA, 0x58, 0x0C), Colors.White, 0, 40, 13);
        addGroupBtn.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
        addGroupBtn.Margin = new Thickness(0, 10, 0, 0);
        addGroupBtn.Click += (_, _) => PromptAddGroup();
        DockPanel.SetDock(addGroupBtn, Dock.Bottom);
        leftStack.Children.Add(addGroupBtn);

        var groupsScroll = new ScrollViewer { VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto };
        _groupsPanel = new StackPanel();
        groupsScroll.Content = _groupsPanel;
        leftStack.Children.Add(groupsScroll);

        leftCard.Child = leftStack;
        bodyGrid.Children.Add(leftCard);

        // ستون چپ: بارکدهای زیرگروه انتخاب‌شده
        var rightCard = new Border { Background = System.Windows.Media.Brushes.White, CornerRadius = new CornerRadius(16), Padding = new Thickness(14) };
        Grid.SetColumn(rightCard, 2);
        var rightStack = new DockPanel();

        var rightHeaderStack = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
        _entriesHeaderText = new TextBlock
        {
            Text = "یک زیرگروه را از لیست گروه‌ها انتخاب کنید",
            FontSize = 15,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
            TextWrapping = TextWrapping.Wrap
        };
        rightHeaderStack.Children.Add(_entriesHeaderText);
        _remainingCountText = new TextBlock
        {
            FontSize = 12,
            Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
            Margin = new Thickness(0, 4, 0, 0)
        };
        rightHeaderStack.Children.Add(_remainingCountText);

        _captureToggleButton = HighUsageUi.CreateRoundedButton("", Color.FromRgb(0x10, 0xB9, 0x81), Colors.White, 0, 44, 14);
        _captureToggleButton.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
        _captureToggleButton.Margin = new Thickness(0, 10, 0, 0);
        _captureToggleButton.Click += CaptureToggleButton_Click;
        _captureToggleButton.Visibility = System.Windows.Visibility.Collapsed;
        rightHeaderStack.Children.Add(_captureToggleButton);

        var captureHint = new TextBlock
        {
            Text = "وقتی «دریافت بارکد» فعال باشد، تا وقتی آن را متوقف نکنید، هر بارکدی که از گوشی اسکن شود، به‌جای تاریخچه، به همین زیرگروه اضافه می‌شود.",
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 6, 0, 0)
        };
        rightHeaderStack.Children.Add(captureHint);

        DockPanel.SetDock(rightHeaderStack, Dock.Top);
        rightStack.Children.Add(rightHeaderStack);

        var entriesScroll = new ScrollViewer { VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto };
        _entriesPanel = new StackPanel();
        entriesScroll.Content = _entriesPanel;
        rightStack.Children.Add(entriesScroll);

        rightCard.Child = rightStack;
        bodyGrid.Children.Add(rightCard);

        // این پنجره یک پنجره‌ی مستقل ویندوزی جدا از MainWindow است، پس پیغام‌های overlay داخل
        // MainWindow (ShowStyledMessage) همیشه پشت این پنجره پنهان می‌مانند («خروجی اکسل گرفتم اما
        // پیغامش رفت پشت پنجره»). به همین دلیل یک toast مستقل و کوچک داخل همین پنجره ساخته می‌شود.
        var rootGrid = new Grid();
        rootGrid.Children.Add(outerGrid);

        _toastBorder = new Border
        {
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(18, 12, 18, 12),
            Margin = new Thickness(24, 0, 24, 24),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
            VerticalAlignment = System.Windows.VerticalAlignment.Bottom,
            Visibility = System.Windows.Visibility.Collapsed,
            Effect = new DropShadowEffect { Color = Colors.Black, Opacity = 0.25, BlurRadius = 16, ShadowDepth = 3 }
        };
        _toastText = new TextBlock
        {
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = System.Windows.Media.Brushes.White,
            TextWrapping = TextWrapping.Wrap,
            TextAlignment = TextAlignment.Center
        };
        _toastBorder.Child = _toastText;
        rootGrid.Children.Add(_toastBorder);

        Content = rootGrid;

        RefreshFromSource();
    }

    // پیغام کوتاه شناور داخل همین پنجره (نه overlay داخل MainWindow) - چند ثانیه دیده می‌شود و
    // خودش پنهان می‌شود. برای پیغام‌های خروجی اکسل (موفق/خطا/بدون داده) از همین استفاده می‌شود.
    internal void ShowToast(string title, string message, bool isError = false)
    {
        try
        {
            _toastBorder.Background = new SolidColorBrush(isError ? Color.FromRgb(0xDC, 0x26, 0x26) : Color.FromRgb(0x10, 0xB9, 0x81));
            _toastText.Text = string.IsNullOrWhiteSpace(title) ? message : $"{title}: {message}";
            _toastBorder.Visibility = System.Windows.Visibility.Visible;

            _toastTimer?.Stop();
            _toastTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3.2) };
            _toastTimer.Tick += (_, _) =>
            {
                _toastTimer?.Stop();
                _toastBorder.Visibility = System.Windows.Visibility.Collapsed;
            };
            _toastTimer.Start();
        }
        catch { }
    }

    public void RefreshFromSource()
    {
        RebuildGroupsPanel();
        RebuildEntriesPanel();

        // چرا این‌جا لازم است: کل برنامه برای هر بارکد اسکن‌شده (صرف‌نظر از این ویژگی)، همیشه یک
        // Enter شبیه‌سازی‌شده به پنجره‌ای که همان لحظه فوکوس کیبورد دارد می‌فرستد (تا در نرم‌افزار
        // نسخه‌نویسی خارجی هم تایید شود). دکمه‌ی «شروع/توقف دریافت بارکد» بر خلاف دکمه‌های
        // گروه/زیرگروه بازسازی نمی‌شود (همیشه همان یک نمونه است)، پس اگر فوکوس کیبورد رویش مانده
        // باشد (که بعد از هر کلیک روی آن پیش می‌آید)، همان Enter شبیه‌سازی‌شده دوباره خودِ همین
        // دکمه را کلیک می‌کند و «دریافت بارکد» را خاموش می‌کند - دقیقاً همان باگی که باعث می‌شد
        // بعد از هر اسکن، دریافت بارکد خودش قطع شود. پاک‌کردن فوکوس بعد از هر رفرش از این جلوگیری می‌کند.
        try { Keyboard.ClearFocus(); } catch { }
    }

    private void RebuildGroupsPanel()
    {
        _groupsPanel.Children.Clear();

        foreach (var group in _owner.HighUsageGroups)
        {
            var groupCard = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xF7, 0xED)),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(10),
                Margin = new Thickness(0, 0, 0, 10)
            };
            var stack = new StackPanel();

            var headerRow = new Grid();
            headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            var nameText = new TextBlock
            {
                Text = group.Name,
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                Foreground = new SolidColorBrush(Color.FromRgb(0xC2, 0x41, 0x0C)),
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetColumn(nameText, 0);
            headerRow.Children.Add(nameText);

            var deleteGroupBtn = HighUsageUi.CreateRoundedButton("🗑", Color.FromRgb(0xFE, 0xE2, 0xE2), Color.FromRgb(0xB9, 0x1C, 0x1C), 26, 26, 11);
            deleteGroupBtn.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            deleteGroupBtn.Margin = new Thickness(6, 0, 0, 0);
            string groupId = group.Id;
            string groupName = group.Name;
            var groupSubIds = group.Subgroups.Select(s => s.Id).ToList();
            deleteGroupBtn.Click += (_, _) =>
            {
                if (HighUsageConfirmWindow.Confirm(this, "حذف گروه", $"گروه «{groupName}» و همه‌ی زیرگروه‌ها و بارکدهای داخل آن حذف شود؟"))
                {
                    _owner.DeleteHighUsageGroup(groupId);
                    if (_selectedSubgroupId != null && groupSubIds.Contains(_selectedSubgroupId))
                        _selectedSubgroupId = null;
                    RefreshFromSource();
                }
            };
            Grid.SetColumn(deleteGroupBtn, 1);
            headerRow.Children.Add(deleteGroupBtn);
            stack.Children.Add(headerRow);

            var subWrap = new WrapPanel { Margin = new Thickness(0, 8, 0, 0) };
            foreach (var sub in group.Subgroups)
            {
                bool isSelected = sub.Id == _selectedSubgroupId;
                bool isCapturing = _owner.HighUsageCaptureSubgroupId == sub.Id;
                var subBtn = HighUsageUi.CreateRoundedButton(
                    $"{sub.Name} ({HighUsageUi.FormatSubgroupCount(sub)}){(isCapturing ? " 🔴" : "")}",
                    isSelected ? Color.FromRgb(0x0E, 0x74, 0x90) : Color.FromRgb(0x0E, 0xA5, 0xE9),
                    Colors.White, 0, 38, 12);
                string subgroupId = sub.Id;
                subBtn.Click += (_, _) => { _selectedSubgroupId = subgroupId; RefreshFromSource(); };

                // دکمه‌ی کوچک ضربدر برای حذف مستقیم همین زیرگروه (طبق درخواست کاربر).
                var deleteSubBtn = HighUsageUi.CreateRoundedButton("✕", Color.FromRgb(0xFE, 0xE2, 0xE2), Color.FromRgb(0xB9, 0x1C, 0x1C), 22, 22, 9);
                deleteSubBtn.VerticalAlignment = System.Windows.VerticalAlignment.Center;
                deleteSubBtn.Margin = new Thickness(2, 0, 0, 0);
                string subName = sub.Name;
                string parentGroupId = group.Id;
                deleteSubBtn.Click += (_, _) =>
                {
                    if (HighUsageConfirmWindow.Confirm(this, "حذف زیرگروه", $"زیرگروه «{subName}» و بارکدهای داخل آن حذف شود؟"))
                    {
                        _owner.DeleteHighUsageSubgroup(parentGroupId, subgroupId);
                        if (_selectedSubgroupId == subgroupId)
                            _selectedSubgroupId = null;
                        RefreshFromSource();
                    }
                };

                var subChip = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 6, 6) };
                subChip.Children.Add(subBtn);
                subChip.Children.Add(deleteSubBtn);
                subWrap.Children.Add(subChip);
            }

            var addSubBtn = HighUsageUi.CreateRoundedButton("+ زیرگروه", Color.FromRgb(0xE5, 0xE7, 0xEB), Color.FromRgb(0x37, 0x41, 0x51), 0, 38, 12);
            addSubBtn.Margin = new Thickness(0, 0, 6, 6);
            string addSubGroupId = group.Id;
            addSubBtn.Click += (_, _) => PromptAddSubgroup(addSubGroupId);
            subWrap.Children.Add(addSubBtn);

            stack.Children.Add(subWrap);
            groupCard.Child = stack;
            _groupsPanel.Children.Add(groupCard);
        }

        if (_owner.HighUsageGroups.Count == 0)
        {
            _groupsPanel.Children.Add(new TextBlock
            {
                Text = "هنوز گروهی نساخته‌اید. با دکمه‌ی «+ گروه جدید» شروع کنید (مثلاً «سرم‌ها» یا «آنتی‌بیوتیک‌ها»).",
                TextWrapping = TextWrapping.Wrap,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80))
            });
        }
    }

    private void RebuildEntriesPanel()
    {
        _entriesPanel.Children.Clear();
        var subgroup = _selectedSubgroupId != null ? _owner.FindHighUsageSubgroup(_selectedSubgroupId) : null;

        if (subgroup == null)
        {
            _entriesHeaderText.Text = "یک زیرگروه را از لیست گروه‌ها انتخاب کنید";
            _remainingCountText.Text = string.Empty;
            _captureToggleButton.Visibility = System.Windows.Visibility.Collapsed;
            return;
        }

        _entriesHeaderText.Text = $"بارکدهای «{subgroup.Name}»" + (subgroup.UnitsPerBarcode > 1 ? $" (هر بارکد {subgroup.UnitsPerBarcode} واحد)" : "");
        _remainingCountText.Text = $"{HighUsageUi.FormatSubgroupCount(subgroup)} باقی مانده در صف";

        bool isCapturing = _owner.HighUsageCaptureSubgroupId == subgroup.Id;
        _captureToggleButton.Visibility = System.Windows.Visibility.Visible;
        _captureToggleButton.Content = isCapturing ? "⏹ توقف دریافت بارکد" : "▶ شروع دریافت بارکد برای این زیرگروه";
        HighUsageUi.SetButtonColors(_captureToggleButton, isCapturing ? Color.FromRgb(0xDC, 0x26, 0x26) : Color.FromRgb(0x10, 0xB9, 0x81), Colors.White);

        var entries = subgroup.Entries.OrderBy(x => x.ScannedAtUtc).ToList();
        string subgroupId = subgroup.Id;
        int i = 1;
        foreach (var entry in entries)
        {
            var row = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var rowBg = new Border { Background = new SolidColorBrush(Color.FromRgb(0xF9, 0xFA, 0xFB)), CornerRadius = new CornerRadius(8) };
            Grid.SetColumnSpan(rowBg, 4);
            row.Children.Add(rowBg);

            var idxText = new TextBlock
            {
                Text = (i++).ToString(),
                Width = 28,
                TextAlignment = TextAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)),
                Margin = new Thickness(8, 8, 0, 8)
            };
            Grid.SetColumn(idxText, 0);
            row.Children.Add(idxText);

            var barcodeText = new TextBlock
            {
                Text = subgroup.UnitsPerBarcode > 1
                    ? $"{entry.Barcode}   ({entry.RemainingUses} از {subgroup.UnitsPerBarcode} واحد باقی مانده)"
                    : entry.Barcode,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(6, 8, 6, 8),
                FlowDirection = System.Windows.FlowDirection.LeftToRight,
                TextAlignment = TextAlignment.Left
            };
            Grid.SetColumn(barcodeText, 1);
            row.Children.Add(barcodeText);

            var timeText = new TextBlock
            {
                Text = entry.ScannedAtUtc.ToLocalTime().ToString("yyyy/MM/dd HH:mm:ss"),
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
                Margin = new Thickness(0, 8, 6, 8),
                FlowDirection = System.Windows.FlowDirection.LeftToRight
            };
            Grid.SetColumn(timeText, 2);
            row.Children.Add(timeText);

            var deleteBtn = HighUsageUi.CreateRoundedButton("✕", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x9C, 0xA3, 0xAF), 24, 24, 10);
            deleteBtn.Margin = new Thickness(0, 0, 8, 0);
            var entryRef = entry;
            deleteBtn.Click += (_, _) =>
            {
                _owner.DeleteHighUsageEntry(subgroupId, entryRef);
                RefreshFromSource();
            };
            Grid.SetColumn(deleteBtn, 3);
            row.Children.Add(deleteBtn);

            _entriesPanel.Children.Add(row);
        }

        if (entries.Count == 0)
        {
            _entriesPanel.Children.Add(new TextBlock
            {
                Text = "این زیرگروه هنوز بارکدی ندارد. «دریافت بارکد» را فعال کنید و با گوشی، بارکدها را اسکن کنید.",
                TextWrapping = TextWrapping.Wrap,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
                Margin = new Thickness(4, 10, 4, 4)
            });
        }
    }

    private void CaptureToggleButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedSubgroupId == null) return;
        bool isCapturing = _owner.HighUsageCaptureSubgroupId == _selectedSubgroupId;
        _owner.SetHighUsageCaptureTarget(isCapturing ? null : _selectedSubgroupId);
        RefreshFromSource();
    }

    private void PromptAddGroup()
    {
        var input = HighUsageTextInputWindow.Prompt(this, "گروه جدید", "نام گروه (مثلاً سرم‌ها):");
        if (string.IsNullOrWhiteSpace(input)) return;
        _owner.AddHighUsageGroup(input.Trim());
        RefreshFromSource();
    }

    private void PromptAddSubgroup(string groupId)
    {
        var (name, unitsText) = HighUsageTextInputWindow.PromptWithSecondary(
            this,
            "زیرگروه جدید",
            "نام زیرگروه (مثلاً رینگر):",
            "هر بارکد چند واحد را پوشش می‌دهد؟",
            "1",
            "برای بارکدهایی که هر واحد بارکد جداگانه‌ی خودش را دارد (مثلاً هر سرم)، عدد ۱ بگذارید. " +
            "برای بارکدی که فقط روی جعبه است و چند واحد را با هم پوشش می‌دهد (مثلاً یک جعبه‌ی ۲۰ ویالی " +
            "پنتوپرازول که خودِ ویال‌ها بارکد جدا ندارند)، همان تعداد (مثلاً ۲۰) را بگذارید - هر بار " +
            "استفاده از آن بارکد یک واحد کم می‌شود و وقتی تمام شد، خودش از صف حذف می‌شود.");
        if (string.IsNullOrWhiteSpace(name)) return;
        int units = int.TryParse((unitsText ?? "").Trim(), out var parsedUnits) && parsedUnits > 0 ? parsedUnits : 1;
        var sub = _owner.AddHighUsageSubgroup(groupId, name.Trim(), units);
        if (sub != null)
            _selectedSubgroupId = sub.Id;
        RefreshFromSource();
    }
}

// =====================================================================================
// یک پنجره‌ی کوچک ورودی متن (برای نام گروه/زیرگروه) - چون این ویژگی هیچ XAML جداگانه‌ای ندارد.
// =====================================================================================
public class HighUsageTextInputWindow : Window
{
    private readonly TextBox _textBox;
    private readonly TextBox? _secondaryTextBox;
    public string? ResultText { get; private set; }
    public string? SecondaryResultText { get; private set; }

    // secondaryLabel/secondaryDefaultValue/secondaryHint: وقتی secondaryLabel داده شود، یک فیلد
    // دومِ اختیاری (مثلاً «هر بارکد چند واحد را پوشش می‌دهد؟») زیر فیلد اصلی اضافه می‌شود - برای
    // این‌که پنجره‌ی مشترک ورودی متن (که قبلاً فقط یک فیلد داشت) بدون تکرار کد، برای «زیرگروه جدید»
    // هم قابل استفاده باشد.
    public HighUsageTextInputWindow(string title, string label, string? secondaryLabel = null, string secondaryDefaultValue = "", string? secondaryHint = null)
    {
        Title = title;
        Width = 380;
        SizeToContent = System.Windows.SizeToContent.Height;
        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
        FlowDirection = System.Windows.FlowDirection.RightToLeft;
        WindowStyle = System.Windows.WindowStyle.None;
        AllowsTransparency = true;
        Background = System.Windows.Media.Brushes.Transparent;
        ResizeMode = System.Windows.ResizeMode.NoResize;

        var card = new Border
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(18),
            Padding = new Thickness(20),
            Margin = new Thickness(20),
            Effect = new DropShadowEffect { Color = Colors.Black, Opacity = 0.3, BlurRadius = 20, ShadowDepth = 4 }
        };

        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = label,
            FontSize = 14,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 10)
        });

        _textBox = new TextBox
        {
            FontSize = 14,
            Padding = new Thickness(8),
            Height = 38,
            VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        };
        _textBox.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter) { AcceptAndClose(); e.Handled = true; }
            else if (e.Key == Key.Escape) { Close(); e.Handled = true; }
        };
        stack.Children.Add(_textBox);

        if (secondaryLabel != null)
        {
            stack.Children.Add(new TextBlock
            {
                Text = secondaryLabel,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 14, 0, 10)
            });

            _secondaryTextBox = new TextBox
            {
                Text = secondaryDefaultValue,
                FontSize = 14,
                Padding = new Thickness(8),
                Height = 38,
                VerticalContentAlignment = System.Windows.VerticalAlignment.Center,
                FlowDirection = System.Windows.FlowDirection.LeftToRight,
                TextAlignment = TextAlignment.Left
            };
            _secondaryTextBox.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter) { AcceptAndClose(); e.Handled = true; }
                else if (e.Key == Key.Escape) { Close(); e.Handled = true; }
            };
            stack.Children.Add(_secondaryTextBox);

            if (!string.IsNullOrWhiteSpace(secondaryHint))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = secondaryHint,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 6, 0, 0)
                });
            }
        }

        var buttonsRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Left, Margin = new Thickness(0, 16, 0, 0) };
        var okBtn = HighUsageUi.CreateRoundedButton("تایید", Color.FromRgb(0xEA, 0x58, 0x0C), Colors.White, 90, 38, 13);
        okBtn.Margin = new Thickness(0, 0, 8, 0);
        okBtn.Click += (_, _) => AcceptAndClose();
        var cancelBtn = HighUsageUi.CreateRoundedButton("انصراف", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 90, 38, 13);
        cancelBtn.Click += (_, _) => Close();
        buttonsRow.Children.Add(okBtn);
        buttonsRow.Children.Add(cancelBtn);
        stack.Children.Add(buttonsRow);

        card.Child = stack;
        Content = card;

        Loaded += (_, _) => { _textBox.Focus(); };
    }

    private void AcceptAndClose()
    {
        ResultText = _textBox.Text;
        SecondaryResultText = _secondaryTextBox?.Text;
        DialogResult = true;
    }

    public static string? Prompt(Window owner, string title, string label)
    {
        var win = new HighUsageTextInputWindow(title, label) { Owner = owner };
        bool? result = win.ShowDialog();
        return result == true ? win.ResultText : null;
    }

    public static (string? name, string? secondary) PromptWithSecondary(Window owner, string title, string label, string secondaryLabel, string secondaryDefaultValue, string? secondaryHint = null)
    {
        var win = new HighUsageTextInputWindow(title, label, secondaryLabel, secondaryDefaultValue, secondaryHint) { Owner = owner };
        bool? result = win.ShowDialog();
        return result == true ? (win.ResultText, win.SecondaryResultText) : (null, null);
    }
}

// =====================================================================================
// یک دیالوگ تایید/لغو با همان استایل گردگوشه‌ی نرم‌افزار - جایگزین MessageBox.Show پیش‌فرض
// ویندوز برای پیغام‌های اخطاری این ویژگی (مثلاً تایید حذف گروه/زیرگروه).
// =====================================================================================
public class HighUsageConfirmWindow : Window
{
    public bool Confirmed { get; private set; }

    public HighUsageConfirmWindow(string title, string message, string confirmText = "حذف", string cancelText = "انصراف")
    {
        Title = title;
        Width = 380;
        SizeToContent = System.Windows.SizeToContent.Height;
        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
        FlowDirection = System.Windows.FlowDirection.RightToLeft;
        WindowStyle = System.Windows.WindowStyle.None;
        AllowsTransparency = true;
        Background = System.Windows.Media.Brushes.Transparent;
        ResizeMode = System.Windows.ResizeMode.NoResize;

        var card = new Border
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(18),
            Padding = new Thickness(20),
            Margin = new Thickness(20),
            Effect = new DropShadowEffect { Color = Colors.Black, Opacity = 0.3, BlurRadius = 20, ShadowDepth = 4 }
        };

        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = "⚠ " + title,
            FontSize = 15,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0xB9, 0x1C, 0x1C)),
            Margin = new Thickness(0, 0, 0, 8)
        });
        stack.Children.Add(new TextBlock
        {
            Text = message,
            FontSize = 13,
            Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 4)
        });

        var buttonsRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Left, Margin = new Thickness(0, 16, 0, 0) };
        var okBtn = HighUsageUi.CreateRoundedButton(confirmText, Color.FromRgb(0xDC, 0x26, 0x26), Colors.White, 90, 38, 13);
        okBtn.Margin = new Thickness(0, 0, 8, 0);
        okBtn.Click += (_, _) => { Confirmed = true; DialogResult = true; };
        var cancelBtn = HighUsageUi.CreateRoundedButton(cancelText, Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 90, 38, 13);
        cancelBtn.Click += (_, _) => Close();
        buttonsRow.Children.Add(okBtn);
        buttonsRow.Children.Add(cancelBtn);
        stack.Children.Add(buttonsRow);

        card.Child = stack;
        Content = card;
    }

    public static bool Confirm(Window owner, string title, string message, string confirmText = "حذف", string cancelText = "انصراف")
    {
        var win = new HighUsageConfirmWindow(title, message, confirmText, cancelText) { Owner = owner };
        bool? result = win.ShowDialog();
        return result == true && win.Confirmed;
    }
}

// =====================================================================================
// پنجره‌ی انتخاب گروه/زیرگروه قبل از خروجی اکسل: چون معمولاً کاربر فقط چند زیرگروه خاص را
// می‌خواهد نه کل بانک را - قبل از باز شدن SaveFileDialog روی پنجره‌ی مدیریت نشان داده می‌شود.
// =====================================================================================
public class HighUsageExportSelectionWindow : Window
{
    private readonly List<CheckBox> _checkBoxes = new();
    public HashSet<string>? SelectedSubgroupIds { get; private set; }

    public HighUsageExportSelectionWindow(List<HighUsageBarcodeGroup> groups)
    {
        Title = "انتخاب برای خروجی اکسل";
        Width = 420;
        SizeToContent = System.Windows.SizeToContent.Height;
        MaxHeight = 560;
        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
        FlowDirection = System.Windows.FlowDirection.RightToLeft;
        WindowStyle = System.Windows.WindowStyle.None;
        AllowsTransparency = true;
        Background = System.Windows.Media.Brushes.Transparent;
        ResizeMode = System.Windows.ResizeMode.NoResize;

        var card = new Border
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(18),
            Padding = new Thickness(20),
            Margin = new Thickness(20),
            Effect = new DropShadowEffect { Color = Colors.Black, Opacity = 0.3, BlurRadius = 20, ShadowDepth = 4 }
        };

        var root = new DockPanel();

        var titleText = new TextBlock
        {
            Text = "کدام گروه/زیرگروه‌ها خروجی گرفته شوند؟",
            FontSize = 14,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 10)
        };
        DockPanel.SetDock(titleText, Dock.Top);
        root.Children.Add(titleText);

        var quickRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
        var selectAllBtn = HighUsageUi.CreateRoundedButton("انتخاب همه", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 0, 30, 11);
        selectAllBtn.Margin = new Thickness(0, 0, 8, 0);
        selectAllBtn.Click += (_, _) => { foreach (var cb in _checkBoxes) cb.IsChecked = true; };
        var selectNoneBtn = HighUsageUi.CreateRoundedButton("هیچ‌کدام", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 0, 30, 11);
        selectNoneBtn.Click += (_, _) => { foreach (var cb in _checkBoxes) cb.IsChecked = false; };
        quickRow.Children.Add(selectAllBtn);
        quickRow.Children.Add(selectNoneBtn);
        DockPanel.SetDock(quickRow, Dock.Top);
        root.Children.Add(quickRow);

        var buttonsRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Left, Margin = new Thickness(0, 14, 0, 0) };
        var okBtn = HighUsageUi.CreateRoundedButton("خروجی اکسل", Color.FromRgb(0x10, 0xB9, 0x81), Colors.White, 110, 38, 13);
        okBtn.Margin = new Thickness(0, 0, 8, 0);
        okBtn.Click += (_, _) => AcceptAndClose();
        var cancelBtn = HighUsageUi.CreateRoundedButton("انصراف", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x37, 0x41, 0x51), 90, 38, 13);
        cancelBtn.Click += (_, _) => Close();
        buttonsRow.Children.Add(okBtn);
        buttonsRow.Children.Add(cancelBtn);
        DockPanel.SetDock(buttonsRow, Dock.Bottom);
        root.Children.Add(buttonsRow);

        var scroll = new ScrollViewer { VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto, MaxHeight = 360 };
        var listPanel = new StackPanel();

        foreach (var group in groups)
        {
            var withEntries = group.Subgroups.Where(s => s.Entries.Count > 0).ToList();
            if (withEntries.Count == 0) continue;

            listPanel.Children.Add(new TextBlock
            {
                Text = group.Name,
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(0xC2, 0x41, 0x0C)),
                Margin = new Thickness(2, 8, 2, 4)
            });

            foreach (var sub in withEntries)
            {
                var cb = new CheckBox
                {
                    Content = $"{sub.Name} ({HighUsageUi.FormatSubgroupCount(sub)})",
                    IsChecked = true,
                    FontSize = 12,
                    Tag = sub.Id,
                    Margin = new Thickness(10, 2, 2, 2)
                };
                _checkBoxes.Add(cb);
                listPanel.Children.Add(cb);
            }
        }

        if (_checkBoxes.Count == 0)
        {
            listPanel.Children.Add(new TextBlock
            {
                Text = "هیچ زیرگروهی با بارکد ثبت‌شده وجود ندارد.",
                TextWrapping = TextWrapping.Wrap,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80))
            });
        }

        scroll.Content = listPanel;
        root.Children.Add(scroll);

        card.Child = root;
        Content = card;
    }

    private void AcceptAndClose()
    {
        SelectedSubgroupIds = _checkBoxes.Where(c => c.IsChecked == true).Select(c => (string)c.Tag).ToHashSet();
        DialogResult = true;
    }

    public static HashSet<string>? PromptSelection(Window owner, List<HighUsageBarcodeGroup> groups)
    {
        var win = new HighUsageExportSelectionWindow(groups) { Owner = owner };
        bool? result = win.ShowDialog();
        return result == true ? win.SelectedSubgroupIds : null;
    }
}
