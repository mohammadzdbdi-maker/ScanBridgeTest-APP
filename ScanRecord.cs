using System;
using System.Globalization;

namespace ScanBridgeTest;

public class ScanRecord
{
    public DateTime TimestampLocal { get; set; }
    public string Barcode { get; set; } = "";
    public string DeviceName { get; set; } = "";
    public string DrugName { get; set; } = "";
    public BarcodeSource Source { get; set; }

    public ScanRecord() { }

    public ScanRecord(DateTime timestamp, string barcode, string deviceName)
    {
        TimestampLocal = timestamp;
        Barcode = barcode;
        DeviceName = deviceName;
    }

    public string PersianDateText => TimestampLocal.ToString("yyyy/MM/dd", new CultureInfo("fa-IR"));
    public string TimeText => TimestampLocal.ToString("HH:mm:ss");
}