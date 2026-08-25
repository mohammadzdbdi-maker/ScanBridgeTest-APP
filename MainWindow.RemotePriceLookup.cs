using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace ScanBridgeTest;

public partial class MainWindow
{
    private List<Services.PriceLookupService.ProductSummary> _priceLookupPhoneList = new();
    private bool _priceLookupQueryFromPhone;

    private void BroadcastPriceLookupOpen()
    {
        _service?.BroadcastJson(new { type = "PRICE_LOOKUP_OPEN" });
    }

    private void BroadcastPriceLookupCancel()
    {
        _service?.BroadcastJson(new { type = "PRICE_LOOKUP_CANCEL" });
    }

    private void BroadcastPriceLookupStatus(string status)
    {
        _service?.BroadcastJson(new { type = "PRICE_LOOKUP_STATUS", status = status ?? "" });
    }

    private void BroadcastPriceLookupList(IReadOnlyList<Services.PriceLookupService.ProductSummary> products, string status)
    {
        _priceLookupPhoneList = products?.ToList() ?? new List<Services.PriceLookupService.ProductSummary>();
        var items = _priceLookupPhoneList.Select((p, i) =>
        {
            string title = string.IsNullOrWhiteSpace(p.Brand) ? p.Title : p.Brand;
            return new
            {
                index = i,
                title = title ?? "",
                form = p.Form ?? "",
                dose = p.Dose ?? "",
                price = p.TotalPriceRial,
                productId = p.ProductId
            };
        }).ToList();
        _service?.BroadcastJson(new { type = "PRICE_LOOKUP_LIST", status = status ?? "", items });
    }

    private void BroadcastPriceLookupDetail(Services.PriceLookupService.PriceResult result)
    {
        if (result == null)
            return;
        Services.PriceLookupService.ParseNameParts(result.FaName ?? "", out _, out var form, out _);
        if (string.IsNullOrWhiteSpace(form))
            form = result.ProductType ?? "";
        _service?.BroadcastJson(new
        {
            type = "PRICE_LOOKUP_DETAIL",
            faName = result.FaName ?? "",
            enName = result.EnName ?? "",
            form,
            genericCode = result.GenericCode ?? "",
            packageCount = result.PackageCount ?? "",
            brandOwner = result.BrandOwner ?? "",
            price = result.TotalPriceRial,
            unitPrice = result.ConsumerPricePerUnit
        });
    }

    private void HandlePriceLookupPhoneMessage(string type, string json)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(new Action(() => HandlePriceLookupPhoneMessage(type, json)));
            return;
        }

        _ = HandlePriceLookupPhoneMessageAsync(type, json);
    }

    private async Task HandlePriceLookupPhoneMessageAsync(string type, string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(json) ? "{}" : json);
            var root = doc.RootElement;

            if (string.Equals(type, "PRICE_LOOKUP_CLOSE", StringComparison.OrdinalIgnoreCase))
            {
                ClosePriceLookup();
                return;
            }

            if (string.Equals(type, "PRICE_LOOKUP_REQUEST_OPEN", StringComparison.OrdinalIgnoreCase))
            {
                if (!HasValidTtacToken())
                {
                    _pendingTtacRetryAction = async () =>
                    {
                        OpenPriceLookup("بر اساس نام، بارکد یا کد ژنریک جست‌وجو کنید");
                        await Task.CompletedTask;
                    };
                    _pendingTtacRetryLabel = "استعلام قیمت";
                    ShowTtacLoginOverlay();
                    BroadcastPriceLookupStatus("برای استعلام قیمت وارد تی‌تک شوید.");
                    return;
                }
                OpenPriceLookup("بر اساس نام، بارکد یا کد ژنریک جست‌وجو کنید");
                return;
            }

            if (string.Equals(type, "PRICE_LOOKUP_SEARCH", StringComparison.OrdinalIgnoreCase))
            {
                string mode = ReadJsonStringLocal(root, "mode") ?? "name";
                string query = ReadJsonStringLocal(root, "query") ?? "";
                await RunPriceLookupFromPhoneAsync(mode, query);
                return;
            }

            if (string.Equals(type, "PRICE_LOOKUP_SELECT", StringComparison.OrdinalIgnoreCase))
            {
                int index = -1;
                if (root.TryGetProperty("index", out var ix))
                {
                    if (ix.TryGetInt32(out var n)) index = n;
                    else if (ix.TryGetInt64(out var l)) index = (int)l;
                }
                long productId = 0;
                if (root.TryGetProperty("productId", out var pid))
                {
                    if (pid.TryGetInt64(out var pl)) productId = pl;
                    else if (pid.TryGetInt32(out var pi)) productId = pi;
                }
                await SelectPriceLookupItemFromPhoneAsync(index, productId);
                return;
            }

            if (string.Equals(type, "PRICE_LOOKUP_CUSTOM_QTY", StringComparison.OrdinalIgnoreCase))
            {
                decimal qty = 0;
                if (root.TryGetProperty("qty", out var qel))
                {
                    if (qel.TryGetDecimal(out var qd)) qty = qd;
                    else if (qel.TryGetDouble(out var qf)) qty = (decimal)qf;
                }
                ApplyCustomQtyFromPhone(qty);
            }
        }
        catch (Exception ex)
        {
            BroadcastPriceLookupStatus("❌ " + ex.Message);
        }
    }

    private async Task RunPriceLookupFromPhoneAsync(string mode, string query)
    {
        if (PriceLookupOverlay.Visibility != Visibility.Visible)
        {
            if (!HasValidTtacToken())
            {
                BroadcastPriceLookupStatus("ابتدا روی سیستم وارد تی‌تک شوید.");
                return;
            }
            OpenPriceLookup("بر اساس نام، بارکد یا کد ژنریک جست‌وجو کنید");
        }

        query = (query ?? "").Trim();
        mode = (mode ?? "name").Trim().ToLowerInvariant();
        PriceNameInput.Text = "";
        PriceBarcodeInput.Text = "";
        PriceGenericInput.Text = "";
        if (mode == "barcode")
        {
            PriceBarcodeInput.Text = query;
            _priceActiveField = "barcode";
        }
        else if (mode == "generic")
        {
            PriceGenericInput.Text = query;
            _priceActiveField = "generic";
        }
        else
        {
            PriceNameInput.Text = query;
            _priceActiveField = "name";
        }

        _priceLookupQueryFromPhone = true;
        try
        {
            await RunActivePriceSearchAsync();
        }
        finally
        {
            _priceLookupQueryFromPhone = false;
        }
    }

    private async Task SelectPriceLookupItemFromPhoneAsync(int index, long productId = 0)
    {
        Services.PriceLookupService.ProductSummary? sel = null;
        if (index >= 0 && index < _priceLookupPhoneList.Count)
            sel = _priceLookupPhoneList[index];
        if (sel == null && productId > 0)
            sel = _priceLookupPhoneList.FirstOrDefault(p => p.ProductId == productId);

        if (sel == null)
        {
            if (productId > 0)
            {
                await RunProductDetailsAsync(productId, "فرآورده");
                return;
            }
            BroadcastPriceLookupStatus("این فرآورده در لیست نیست. دوباره جست‌وجو کنید.");
            return;
        }

        bool needsFetch = sel.ProductId > 0
            && (string.IsNullOrWhiteSpace(sel.EnName)
                || string.IsNullOrWhiteSpace(sel.BrandOwner)
                || sel.TotalPriceRial <= 0);
        if (needsFetch)
        {
            await RunProductDetailsAsync(sel.ProductId, sel.Title);
            return;
        }
        ShowPriceResult(PriceLookup.ToPriceResult(sel));
    }

    private void ApplyCustomQtyFromPhone(decimal qty)
    {
        if (qty <= 0)
        {
            BroadcastPriceLookupStatus("یک تعداد معتبر وارد کنید.");
            return;
        }
        if (_lastPriceResult == null)
        {
            BroadcastPriceLookupStatus("ابتدا یک فرآورده را انتخاب کنید.");
            return;
        }
        PriceCustomQtyInput.Text = qty.ToString(CultureInfo.InvariantCulture);
        ConfirmPriceCustomQty();
    }

    private static string? ReadJsonStringLocal(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var el))
            return null;
        if (el.ValueKind == JsonValueKind.String)
            return el.GetString();
        if (el.ValueKind == JsonValueKind.Number)
            return el.GetRawText();
        return null;
    }
}
