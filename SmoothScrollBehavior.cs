using System;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace ScanBridgeTest;

// =====================================================================================
// رفتار (attached behavior) اسکرول نرم: به‌جای پرش پله‌ای پیش‌فرض ویندوز روی چرخ ماوس، هر بار که
// چرخ ماوس می‌چرخد یک «هدف» جدید محاسبه می‌شود و موقعیت واقعی اسکرول با یک ایزینگ نمایی (هر تیک فقط
// بخشی از فاصله‌ی باقی‌مانده تا هدف را طی می‌کند) به آن هدف نزدیک می‌شود - همین باعث حس نرم و کشسان
// اسکرول به‌جای پرش‌های بریده‌بریده می‌شود.
//
// این کلاس فقط با یک Style سراسری (در App.xaml، TargetType="ScrollViewer") روی همه‌ی ScrollViewer های
// برنامه اعمال می‌شود - چه آن‌هایی که مستقیم در XAML نوشته شده‌اند، چه آن‌هایی که داخل تمپلیت پیش‌فرض
// ListBox/ComboBox/... هستند، و چه آن‌هایی که در پنجره‌های ساخته‌شده با کد (مثل بانک بارکد پرمصرف)
// به‌صورت پویا ساخته می‌شوند - چون Style های سطح Application بدون نیاز به ارجاع صریح روی کل درخت
// بصری اعمال می‌شوند.
// =====================================================================================
public static class SmoothScrollBehavior
{
    // مقدار جابه‌جایی (پیکسل) به ازای هر «دندانه»ی استاندارد چرخ ماوس (Delta = 120).
    private const double PixelsPerWheelNotch = 90.0;

    // نسبت طی‌شدن فاصله‌ی باقی‌مانده در هر تیک انیمیشن (هرچه بزرگ‌تر، سریع‌تر و کم‌نرم‌تر).
    private const double EasingFactor = 0.28;

    // وقتی فاصله تا هدف کمتر از این مقدار برسد، انیمیشن متوقف و مستقیم روی هدف نهایی می‌نشیند.
    private const double SnapThreshold = 0.6;

    public static readonly DependencyProperty IsEnabledProperty =
        DependencyProperty.RegisterAttached(
            "IsEnabled",
            typeof(bool),
            typeof(SmoothScrollBehavior),
            new PropertyMetadata(false, OnIsEnabledChanged));

    public static bool GetIsEnabled(DependencyObject obj) => (bool)obj.GetValue(IsEnabledProperty);

    public static void SetIsEnabled(DependencyObject obj, bool value) => obj.SetValue(IsEnabledProperty, value);

    private sealed class ScrollAnimationState
    {
        public double TargetOffset;
        public DispatcherTimer? Timer;
    }

    private static readonly ConditionalWeakTable<ScrollViewer, ScrollAnimationState> States = new();

    private static void OnIsEnabledChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not ScrollViewer scrollViewer)
            return;

        if ((bool)e.NewValue)
        {
            // اسکرول پیکسلی (نه آیتمی) - پیش‌نیاز انیمیشن نرم، حتی روی لیست‌های مجازی‌سازی‌شده.
            scrollViewer.CanContentScroll = false;
            scrollViewer.PreviewMouseWheel += OnPreviewMouseWheel;

            // اگر این ScrollViewer (یا پنجره‌اش) در همان لحظه که تایمر انیمیشن در حال اجراست
            // Unload شود، باید تایمر همین‌جا صریحاً متوقف شود - وگرنه DispatcherTimer.Tick
            // خودش را روی scrollViewer نگه می‌دارد و مانع GC شدن آن (و پنجره‌اش) می‌شود، حتی بعد
            // از بسته‌شدن پنجره (بخشی از باگ ۲۰ گزارش ممیزی).
            scrollViewer.Unloaded += OnScrollViewerUnloaded;
        }
        else
        {
            scrollViewer.PreviewMouseWheel -= OnPreviewMouseWheel;
            scrollViewer.Unloaded -= OnScrollViewerUnloaded;
            StopTimer(scrollViewer);
        }
    }

    private static void OnScrollViewerUnloaded(object sender, RoutedEventArgs e)
    {
        if (sender is ScrollViewer scrollViewer)
            StopTimer(scrollViewer);
    }

    private static void StopTimer(ScrollViewer scrollViewer)
    {
        if (States.TryGetValue(scrollViewer, out var state))
        {
            state.Timer?.Stop();
            state.Timer = null;
        }
    }

    private static void OnPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not ScrollViewer scrollViewer)
            return;

        // اگر اصلاً چیزی برای اسکرول نیست، بگذار رویداد عادی مسیر خودش را برود.
        if (scrollViewer.ScrollableHeight <= 0)
            return;

        e.Handled = true;

        var state = States.GetOrCreateValue(scrollViewer);

        // اگر انیمیشن قبلی هنوز در حال اجراست، از هدفِ همان انیمیشن ادامه بده (نه از موقعیت لحظه‌ای
        // اسکرول که هنوز به هدف قبلی نرسیده) - این‌طور چرخش‌های پشت‌سرهم چرخ ماوس روی هم جمع می‌شوند
        // و اسکرول یک‌دست باقی می‌ماند.
        double baseline = state.Timer != null ? state.TargetOffset : scrollViewer.VerticalOffset;
        double delta = -(e.Delta / 120.0) * PixelsPerWheelNotch;
        double target = baseline + delta;
        if (target < 0) target = 0;
        if (target > scrollViewer.ScrollableHeight) target = scrollViewer.ScrollableHeight;
        state.TargetOffset = target;

        if (state.Timer == null)
        {
            state.Timer = new DispatcherTimer(DispatcherPriority.Render) { Interval = TimeSpan.FromMilliseconds(15) };
            state.Timer.Tick += (_, _) => AnimateStep(scrollViewer, state);
            state.Timer.Start();
        }
    }

    private static void AnimateStep(ScrollViewer scrollViewer, ScrollAnimationState state)
    {
        // اگر لیست همین حین انیمیشن کوچک شود (مثلاً فیلتر شدن)، ScrollableHeight می‌تواند از
        // TargetOffset قدیمی کمتر شود. ScrollToVerticalOffset مقدار واقعی اعمال‌شده را خودش
        // clamp می‌کند، ولی TargetOffset ذخیره‌شده‌ی ما را نه - بدون این تصحیح، current هیچ‌وقت
        // به TargetOffset نمی‌رسد، diff هیچ‌وقت به SnapThreshold نمی‌رسد، و این تایمر تا ابد
        // (هر ۱۵ میلی‌ثانیه، برای همیشه) تیک می‌زند (باگ ۲۰ گزارش ممیزی).
        if (state.TargetOffset > scrollViewer.ScrollableHeight)
            state.TargetOffset = Math.Max(0, scrollViewer.ScrollableHeight);

        double current = scrollViewer.VerticalOffset;
        double diff = state.TargetOffset - current;

        if (Math.Abs(diff) < SnapThreshold)
        {
            scrollViewer.ScrollToVerticalOffset(state.TargetOffset);
            state.Timer?.Stop();
            state.Timer = null;
            return;
        }

        scrollViewer.ScrollToVerticalOffset(current + diff * EasingFactor);
    }
}
