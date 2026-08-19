using System;

namespace ScanBridgeTest;

/// <summary>
/// تنظیمات مربوط به جستجو در تی‌تک
/// </summary>
public class TtTeckSettings
{
    /// <summary>
    /// آیا جستجو در تی تک فعال است؟
    /// </summary>
    public bool IsEnabled { get; set; } = false;

    /// <summary>
    /// تاریخ آخرین همگام‌سازی
    /// </summary>
    public DateTime? LastSyncDate { get; set; }
}