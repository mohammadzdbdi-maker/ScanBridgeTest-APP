using System;
using System.Linq;

namespace ScanBridgeTest;

public enum BarcodeSource
{
    Unknown,
    TtTeck,
    QRCode,
    Manual,
    Linear,
    Other
}

public static class BarcodeDetector
{
    public static BarcodeSource DetectBarcodeType(string barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return BarcodeSource.Unknown;

        barcode = barcode.Trim();

        // بارکد GS1 (شروع با 01) - تی تک
        if (barcode.StartsWith("01"))
        {
            return BarcodeSource.TtTeck;
        }

        // تشخیص QR Code
        if (IsQRCode(barcode))
        {
            return BarcodeSource.QRCode;
        }

        // تشخیص بارکد خطی
        if (IsLinearBarcode(barcode))
        {
            return BarcodeSource.Linear;
        }

        return BarcodeSource.Unknown;
    }

    private static bool IsQRCode(string barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return false;

        return barcode.Length > 20 && barcode.Any(char.IsLetter);
    }

    private static bool IsLinearBarcode(string barcode)
    {
        if (string.IsNullOrWhiteSpace(barcode))
            return false;

        return barcode.All(char.IsDigit) && barcode.Length >= 8 && barcode.Length <= 50;
    }
}