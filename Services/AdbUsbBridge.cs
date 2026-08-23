using System.Diagnostics;

namespace ScanBridgeTest.Services;

/// <summary>
/// پل ADB برای اتصال صفر-ضربه‌ی گوشی با کابل (بدون USB Tethering):
/// هر ۵ ثانیه «adb devices» چک می‌شود؛ هرگاه گوشیِ دارای USB Debugging وصل و تأیید شد،
/// «adb reverse tcp:5050 tcp:5050» اجرا می‌شود تا ws://127.0.0.1:5050 روی خودِ گوشی به همین
/// سیستمِ ویندوز فوروارد شود — یعنی اپ گوشی بدون هیچ تترینگی فقط با وصل کردن کابل می‌تواند
/// متصل شود. با کشیده‌شدن کابل، reverse می‌افتد و پل در دورِ بعدی دوباره آن را برمی‌گرداند.
/// </summary>
public sealed class AdbUsbBridge : IDisposable
{
    private readonly System.Timers.Timer _timer = new(TimeSpan.FromSeconds(5).TotalMilliseconds) { AutoReset = true };
    private readonly string _adbPath;
    private bool _unauthorizedTipShown;
    private bool _reverseActive;

    /// <summary>رویداد وضعیت برای نمایش راهنما در رابط کاربری: UNAUTHORIZED / REVERSE_ON / REVERSE_OFF</summary>
    public event Action<string>? StatusTip;

    public AdbUsbBridge()
    {
        var bundled = Path.Combine(AppContext.BaseDirectory, "adb", "adb.exe");
        _adbPath = File.Exists(bundled) ? bundled : "adb";
        _timer.Elapsed += async (_, _) => await TickAsync();
    }

    public void Start() => _timer.Start();

    private async Task TickAsync()
    {
        try
        {
            string output = await RunAsync("devices");
            bool hasAuthorized = false, hasUnauthorized = false;

            foreach (var raw in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var line = raw.Trim();
                if (line.EndsWith("\tdevice", StringComparison.Ordinal)) hasAuthorized = true;
                if (line.EndsWith("\tunauthorized", StringComparison.Ordinal)) hasUnauthorized = true;
            }

            if (hasUnauthorized && !_unauthorizedTipShown)
            {
                _unauthorizedTipShown = true;
                StatusTip?.Invoke("UNAUTHORIZED");
            }
            else if (!hasUnauthorized)
            {
                _unauthorizedTipShown = false;
            }

            if (hasAuthorized)
            {
                if (!_reverseActive)
                {
                    await RunAsync("reverse tcp:5050 tcp:5050");
                    _reverseActive = true;
                    StatusTip?.Invoke("REVERSE_ON");
                }
            }
            else if (_reverseActive)
            {
                // کابل قطع شد (یا گوشی رفت) — در اتصال بعدی دوباره reverse می‌گیریم
                _reverseActive = false;
                StatusTip?.Invoke("REVERSE_OFF");
            }
        }
        catch
        {
            // adb در دسترس نیست یا خطای موقت — بی‌صدا رد می‌شود؛ اتصال Tethering همچنان کار می‌کند
        }
    }

    private async Task<string> RunAsync(string args)
    {
        var psi = new ProcessStartInfo
        {
            FileName = _adbPath,
            Arguments = args,
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };

        using var process = Process.Start(psi);
        if (process is null)
            return string.Empty;

        string stdout = await process.StandardOutput.ReadToEndAsync();
        await process.WaitForExitAsync();
        return stdout;
    }

    public void Dispose()
    {
        _timer.Stop();
        _timer.Dispose();
    }
}
