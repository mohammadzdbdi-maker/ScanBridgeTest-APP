using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
// پروژه UseWindowsForms=true دارد، پس System.Drawing/System.Windows.Forms هم implicit global
// using هستند - این alias‌ها دقیقاً همان مشکل ابهام CS0104 (Color/ProgressBar/Brushes/Button) را حل می‌کنند.
using Color = System.Windows.Media.Color;
using ProgressBar = System.Windows.Controls.ProgressBar;
using Brushes = System.Windows.Media.Brushes;
using Button = System.Windows.Controls.Button;
using Point = System.Windows.Point;

namespace ScanBridgeTest;

// =====================================================================================
// پنجره‌ی کوچک و غیرمودالِ کنار پنجره‌ی اصلی که پیشرفت دانلود بروزرسانی را نشان می‌دهد - درخواست
// کاربر: «یه صفحه کنار صفحه و کوچولو بیاد که نشون بده چقدر دانلود شده». از StartAppUpdateDownloadAsync
// (در MainWindow.xaml.cs) ساخته و با SetProgress در حین دانلود آپدیت می‌شود.
//
// دو نکته که بعد از تست کاربر اصلاح شد:
//   - Owner ست می‌شود ولی Topmost دیگر true نیست - قبلاً چون Topmost=true بود، این پنجره روی همه‌ی
//     برنامه‌های باز روی دسکتاپ (نه فقط روی خودِ اسکن‌بریج) نشان داده می‌شد؛ یک پنجره‌ی owned بدون
//     Topmost همچنان بالای پنجره‌ی صاحبش (owner) می‌ماند، ولی مزاحم برنامه‌های دیگر نمی‌شود.
//   - یک دکمه‌ی «لغو» دارد که رویداد CancelRequested را صدا می‌زند - StartAppUpdateDownloadAsync با
//     یک CancellationTokenSource به آن گوش می‌دهد تا دانلود را نیمه‌کاره متوقف و فایل ناقص را حذف کند.
// =====================================================================================
public class DownloadProgressWindow : Window
{
    private readonly ProgressBar _bar;
    private readonly TextBlock _percentText;
    private readonly TextBlock _sizeText;
    private readonly Button _cancelBtn;
    private bool _cancelRequested;
    private readonly Point? _anchorRelativeToOwner;

    public event EventHandler? CancelRequested;

    // anchorRelativeToOwner (اختیاری): نقطه‌ای که گوشه‌ی بالا-چپ این پنجره باید دقیقاً همان‌جا قرار
    // بگیرد، نسبت به گوشه‌ی بالا-چپ ناحیه‌ی محتوای پنجره‌ی صاحب (Owner) - نه کل صفحه‌نمایش. طبق
    // درخواست کاربر («دقیقا بیارش زیر تاریخ نزدیک»)، MainWindow این نقطه را از پایین ستون دکمه‌های
    // «پنل کاربری» حساب و پاس می‌دهد (نگاه کنید به GetDownloadProgressAnchorPoint). اگر داده نشود،
    // به حالت قبلی (گوشه‌ی پایین-چپِ کل پنجره) برمی‌گردد.
    public DownloadProgressWindow(string? version, Point? anchorRelativeToOwner = null)
    {
        _anchorRelativeToOwner = anchorRelativeToOwner;
        Title = "دانلود بروزرسانی";
        Width = 300;
        SizeToContent = SizeToContent.Height;
        WindowStartupLocation = WindowStartupLocation.Manual;
        FlowDirection = System.Windows.FlowDirection.RightToLeft;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;

        var card = new Border
        {
            Background = Brushes.White,
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(18),
            Margin = new Thickness(14),
            Effect = new DropShadowEffect { Color = Colors.Black, Opacity = 0.25, BlurRadius = 18, ShadowDepth = 3 }
        };

        var stack = new StackPanel();

        var headerRow = new Grid();
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var titleText = new TextBlock
        {
            Text = "⬇ دانلود بروزرسانی" + (string.IsNullOrWhiteSpace(version) ? "" : $" (نسخه‌ی {version})"),
            FontSize = 13,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x1A, 0x23, 0x7E)),
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(titleText, 0);
        headerRow.Children.Add(titleText);

        _cancelBtn = HighUsageUi.CreateRoundedButton("✕", Color.FromRgb(0xF3, 0xF4, 0xF6), Color.FromRgb(0x6B, 0x72, 0x80), 26, 26, 11);
        _cancelBtn.Margin = new Thickness(8, 0, 0, 0);
        _cancelBtn.Click += (_, _) =>
        {
            if (_cancelRequested)
                return;
            _cancelRequested = true;
            _cancelBtn.IsEnabled = false;
            _percentText.Text = "در حال لغو...";
            CancelRequested?.Invoke(this, EventArgs.Empty);
        };
        Grid.SetColumn(_cancelBtn, 1);
        headerRow.Children.Add(_cancelBtn);

        headerRow.Margin = new Thickness(0, 0, 0, 10);
        stack.Children.Add(headerRow);

        _bar = new ProgressBar
        {
            Height = 10,
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            IsIndeterminate = true,
            Foreground = new SolidColorBrush(Color.FromRgb(0x15, 0x65, 0xC0)),
            Background = new SolidColorBrush(Color.FromRgb(0xE5, 0xE7, 0xEB)),
            BorderThickness = new Thickness(0)
        };
        stack.Children.Add(_bar);

        var row = new Grid { Margin = new Thickness(0, 8, 0, 0) };
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        _sizeText = new TextBlock
        {
            Text = "",
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80))
        };
        Grid.SetColumn(_sizeText, 0);
        row.Children.Add(_sizeText);

        _percentText = new TextBlock
        {
            Text = "در حال شروع...",
            FontSize = 11,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(0x15, 0x65, 0xC0))
        };
        Grid.SetColumn(_percentText, 1);
        row.Children.Add(_percentText);

        stack.Children.Add(row);

        card.Child = stack;
        Content = card;

        // نکته‌ی مهم: از Loaded به ContentRendered عوض شد - چون SizeToContent=Height است، در لحظه‌ی
        // Loaded هنوز یک پاس Layout کامل انجام نشده و ActualHeight/ActualWidth ممکن است هنوز صفر
        // باشند؛ محاسبه‌ی موقعیت با ارتفاع صفر باعث می‌شد پنجره عملاً از پایین لبه‌ی صفحه بیرون بزند
        // (باگی که کاربر گزارش داد: «باکس دانلود رفته زیر»). ContentRendered تضمین می‌کند اندازه‌ی
        // واقعی محتوا حاضر است.
        ContentRendered += (_, _) => PositionBesideOwner();
    }

    // اگر anchorRelativeToOwner داده شده (حالت معمول - نگاه کنید به GetDownloadProgressAnchorPoint
    // در MainWindow.xaml.cs)، پنجره دقیقاً زیر همان نقطه (پایین ستون دکمه‌های «پنل کاربری») قرار
    // می‌گیرد. وگرنه (fallback) گوشه‌ی پایین-چپِ خودِ پنجره‌ی اصلی. چون Topmost=false است و Owner
    // ست شده، این پنجره فقط بالای پنجره‌ی اصلیِ اسکن‌بریج می‌ماند - نه بالای بقیه‌ی برنامه‌های باز
    // روی دسکتاپ کاربر.
    private void PositionBesideOwner()
    {
        try
        {
            if (Owner == null)
            {
                var workArea = SystemParameters.WorkArea;
                Left = workArea.Left + 24;
                Top = workArea.Bottom - ActualHeight - 24;
                return;
            }

            double left, top;
            if (_anchorRelativeToOwner.HasValue)
            {
                left = Owner.Left + _anchorRelativeToOwner.Value.X;
                top = Owner.Top + _anchorRelativeToOwner.Value.Y;
            }
            else
            {
                left = Owner.Left + 24;
                top = Owner.Top + Owner.ActualHeight - ActualHeight - 24;
            }

            // اگر (مثلاً به‌خاطر متن طولانی) پایین‌تر از کف پنجره‌ی اصلی بیفتد، به داخل پنجره برش می‌خورد.
            if (Owner.ActualHeight > 0)
                top = Math.Min(top, Owner.Top + Owner.ActualHeight - ActualHeight - 12);

            Left = left;
            Top = top;
        }
        catch
        {
            Left = 40;
            Top = 40;
        }
    }

    // percent=null یعنی اندازه‌ی فایل از سرور مشخص نبوده (هدر Content-Length نداشت) - نوار پیشرفت
    // به‌صورت نامعین (در حال حرکت، بدون درصد دقیق) نمایش داده می‌شود.
    public void SetProgress(double? percent, string sizeText)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.Invoke(() => SetProgress(percent, sizeText));
            return;
        }

        if (_cancelRequested)
            return;

        _sizeText.Text = sizeText;

        if (percent.HasValue)
        {
            _bar.IsIndeterminate = false;
            _bar.Value = percent.Value;
            _percentText.Text = $"{percent.Value:0}%";
        }
        else
        {
            _bar.IsIndeterminate = true;
            _percentText.Text = sizeText;
        }
    }
}
