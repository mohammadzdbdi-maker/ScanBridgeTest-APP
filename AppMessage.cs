using System.Windows;
using System.Windows.Media;
// پروژه UseWindowsForms=true دارد، پس System.Drawing هم به‌صورت implicit global using در کل
// پروژه هست و «Brush» بین System.Drawing.Brush و System.Windows.Media.Brush مبهم می‌شود - دقیقاً
// همان مشکلی که در MainWindow.HighUsageBarcode.cs برای Button/TextBox/Color/... حل شده بود.
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;

namespace ScanBridgeTest;

public sealed class AppMessage
{
    public string Title { get; set; } = "";
    public string Body { get; set; } = "";
    public string? Link { get; set; }

    // متن دکمه‌ی پایین پیام. پیش‌فرض خالی است - وقتی از notifications.json محلی خوانده می‌شود،
    // LoadMessages آن را با متن عمومی «مشاهده» پر می‌کند؛ برای پیام بروزرسانی که خودِ برنامه
    // می‌سازد (نگاه کنید به CheckForAppUpdateAsync در MainWindow.xaml.cs)، صراحتاً «دانلود و
    // نصب» ست می‌شود. جدا کردن این متن به هر پیام (به‌جای یک متن ثابت مشترک برای همه‌ی پیام‌ها که
    // قبلاً بود) همین امکان را می‌دهد که دکمه‌ی پیام بروزرسانی برچسب متفاوتی داشته باشد.
    public string LinkButtonText { get; set; } = "";

    // اگر true باشد، کلیک روی دکمه (به‌جای باز کردن Link در مرورگر) دانلود فایل نصب از روی Link و
    // اجرای آن را شروع می‌کند - نگاه کنید به MessageLink_Click.
    public bool IsUpdateDownload { get; set; }

    // شماره‌نسخه‌ای که این پیامِ بروزرسانی به آن اشاره دارد - برای «بعداً یادم بنداز» (اگر کاربر
    // این پیام را نادیده گرفت)، تا وقتی نسخه‌ی جدیدتری منتشر نشده دوباره برایش پاپ‌آپ/پیام تکراری
    // ساخته نشود؛ نگاه کنید به AppUpdateCheckSettings.DismissedVersion.
    public string? UpdateVersion { get; set; }

    // اگر این نسخه قبلاً یک‌بار با موفقیت دانلود شده، مسیر فایل نصب دانلودشده اینجا ذخیره می‌شود -
    // طبق درخواست صریح کاربر («یک بار که دانلود کردم دیگه نباید دوباره بیاد»)، دفعه‌ی بعد که کاربر
    // روی دکمه‌ی این پیام کلیک می‌کند (حتی بعد از بستن و باز کردن مجدد برنامه، چون همراه با بقیه‌ی
    // AppMessage در notifications.json ذخیره می‌شود)، اگر این فایل هنوز روی دیسک باشد، دوباره دانلود
    // نمی‌شود - مستقیم به تایید نصب می‌رود؛ نگاه کنید به StartAppUpdateDownloadAsync.
    public string? DownloadedInstallerPath { get; set; }

    // خوانده شده یا نه - پیش‌فرض false (خوانده‌نشده). کاربر با دکمه‌ی «باشه» روی هر پیام (نگاه کنید
    // به MessageAcknowledgeButton_Click در MainWindow.xaml.cs) این را true می‌کند و بلافاصله بعد از
    // آن هم پیام با یک fade-out کاملاً حذف می‌شود. تا وقتی یک پیام false بماند، در شمارش نشان‌گر
    // قرمز روی دکمه‌ی «پیام‌ها» حساب می‌شود - فقط باز کردن پنجره‌ی پیام‌ها دیگر به‌تنهایی نشان‌گر را
    // پاک نمی‌کند (طبق درخواست صریح کاربر).
    public bool IsRead { get; set; }

    public Visibility HasLinkVisibility => string.IsNullOrWhiteSpace(Link) ? Visibility.Collapsed : Visibility.Visible;

    // رنگ پس‌زمینه/حاشیه‌ی کارت پیام بر اساس خوانده‌شده بودن - پیام‌های خوانده‌نشده کمی آبی‌رنگ و
    // پررنگ‌تر دیده می‌شوند تا بدون باز کردن جزئیات هم مشخص باشد کدام‌ها هنوز دیده نشده‌اند. چون
    // AppMessage از INotifyPropertyChanged پیروی نمی‌کند، بعد از عوض شدن IsRead باید MessagesList
    // را با Items.Refresh() دوباره رسم کرد تا این دو مقدار دوباره خوانده شوند.
    public Brush CardBackground => IsRead
        ? new SolidColorBrush(Color.FromRgb(0xF9, 0xFA, 0xFB))
        : new SolidColorBrush(Color.FromRgb(0xEF, 0xF6, 0xFF));

    public Brush CardBorderBrush => IsRead
        ? Brushes.Transparent
        : new SolidColorBrush(Color.FromRgb(0x93, 0xC5, 0xFD));
}
