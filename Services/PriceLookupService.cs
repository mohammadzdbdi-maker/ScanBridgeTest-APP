using System.Net.Http;
using System.Text.Json;

namespace ScanBridgeTest.Services;

/// <summary>
/// استعلام قیمت فرآورده از پنل آمار تی‌تک (newstatisticsreports.ttac.ir):
/// ۱) بارکد → InstanceCatalog → IRC
/// ۲) IRC یا نام فرآورده → GetCompactProducts → ProductId (لیست انتخاب)
/// ۳) ProductId → GetProductsForPharmacies → اطلاعات کامل + قیمت مصرف‌کننده
/// همه‌ی فراخوانی‌ها بدون نیاز به ورود اجرا می‌شوند؛ فقط هدرهای Origin/Referer کافی است
/// (توکن ورود تی‌تک اگر موجود باشد اختیاری ارسال می‌شود).
/// </summary>
public sealed class PriceLookupService
{
    private const string Base = "https://statisticsreports.ttac.ir";
    private const string InstanceCatalogUrl = Base + "/product/InstanceCatalog";
    private const string CompactProductsUrl = Base + "/Product/GetCompactProducts";
    private const string ProductsForPharmaciesUrl = Base + "/product/GetProductsForPharmacies";

    private static readonly HttpClient _http = new HttpClient(new HttpClientHandler
    {
        UseCookies = true,
        AllowAutoRedirect = true,
    })
    { Timeout = TimeSpan.FromSeconds(25) };

    static PriceLookupService()
    {
        _http.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json, text/plain, */*");
        _http.DefaultRequestHeaders.TryAddWithoutValidation("Origin", "https://newstatisticsreports.ttac.ir");
        _http.DefaultRequestHeaders.Referrer = new Uri("https://newstatisticsreports.ttac.ir/");
    }

    // ---------- مدل‌ها ----------

    public sealed class ProductSummary
    {
        public long ProductId { get; set; }
        public string Title { get; set; } = "";
        public string Subtitle { get; set; } = "";
    }

    public sealed class PriceResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public bool FoundButNotDrugSubgroup { get; set; }
        public string ProductType { get; set; } = "";

        public string FaName { get; set; } = "";
        public string EnName { get; set; } = "";
        public string GenericCode { get; set; } = "";
        public string PackageCount { get; set; } = "";
        public string BrandOwner { get; set; } = "";
        public decimal ConsumerPricePerUnit { get; set; }
        public decimal TotalPriceRial { get; set; }
    }

    // ---------- API ----------

    private static HttpRequestMessage BuildGet(string url, string? token)
    {
        var req = new HttpRequestMessage(HttpMethod.Get, url);
        if (!string.IsNullOrWhiteSpace(token))
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        return req;
    }

    /// <summary>بارکد → IRC (از کاتالوگ)</summary>
    public async Task<string?> GetIrcFromBarcodeAsync(string barcode, string? token = null)
    {
        try
        {
            string url = $"{InstanceCatalogUrl}?uid={Uri.EscapeDataString(barcode)}";
            using var req = BuildGet(url, token);
            using var resp = await _http.SendAsync(req);
            string body = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode || string.IsNullOrWhiteSpace(body))
                return null;

            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;
            if (root.ValueKind == JsonValueKind.Object)
            {
                if (root.TryGetProperty("Result", out var result) && result.ValueKind == JsonValueKind.Object)
                {
                    var irc = GetString(result, "Irc", "irc", "IRC");
                    if (!string.IsNullOrWhiteSpace(irc))
                        return irc;
                }
                var rootIrc = GetString(root, "Irc", "irc", "IRC");
                if (!string.IsNullOrWhiteSpace(rootIrc))
                    return rootIrc;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// جست‌وجوی فرآورده با نام یا IRC — تا دو صفحه نتیجه برمی‌گرداند
    /// (کاربر خواست اگر لیست در دو صفحه بود هر دو را ببیند و خودش انتخاب کند).
    /// </summary>
    public async Task<List<ProductSummary>> SearchProductsAsync(string query, string? token = null)
    {
        var all = new List<ProductSummary>();
        for (int page = 1; page <= 2; page++)
        {
            string url = $"{CompactProductsUrl}?Name={Uri.EscapeDataString(query)}&PageSize=50&PageNumber={page}";
            List<ProductSummary> items;
            try
            {
                using var req = BuildGet(url, token);
                using var resp = await _http.SendAsync(req);
                string body = await resp.Content.ReadAsStringAsync();
                if (!resp.IsSuccessStatusCode || string.IsNullOrWhiteSpace(body))
                    break;
                items = ParseCompactList(body);
            }
            catch
            {
                break;
            }

            all.AddRange(items);
            if (items.Count < 50)
                break;
        }

        // حذف تکراری‌ها بر اساس ProductId
        return all
            .GroupBy(p => p.ProductId)
            .Select(g => g.First())
            .ToList();
    }

    /// <summary>جزئیات و قیمت فرآورده با ProductId</summary>
    public async Task<PriceResult> GetProductDetailsAsync(long productId, string? token = null)
    {
        try
        {
            string url = $"{ProductsForPharmaciesUrl}?searchExp=&ProductId={productId}&PageSize=50&PageNumber=1";
            using var req = BuildGet(url, token);
            using var resp = await _http.SendAsync(req);
            string body = await resp.Content.ReadAsStringAsync();

            if (!resp.IsSuccessStatusCode || string.IsNullOrWhiteSpace(body))
                return new PriceResult { Success = false, Message = $"❌ خطای سرور تی‌تک: {resp.StatusCode}" };

            JsonElement first;
            try
            {
                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;
                var items = ExtractArray(root, "data", "Data", "items", "Items", "result", "Result", "list", "List");
                if (items is null)
                {
                    if (root.ValueKind == JsonValueKind.Array && root.GetArrayLength() > 0)
                        items = root;
                    else if (root.ValueKind == JsonValueKind.Object)
                        items = ExtractAnyArray(root);
                    else
                        return new PriceResult { Success = false, Message = "❌ فرآورده‌ای با این مشخصات یافت نشد" };
                }

                if (items is null || items.Value.GetArrayLength() == 0)
                    return new PriceResult { Success = false, Message = "❌ فرآورده‌ای با این مشخصات یافت نشد" };

                first = items.Value[0];
            }
            catch (Exception ex)
            {
                return new PriceResult { Success = false, Message = "❌ پاسخ نامعتبر از سرور: " + ex.Message };
            }

            var result = new PriceResult
            {
                Success = true,
                FaName = GetString(first, "FaBrandName", "faBrandName", "PersianName", "persianName", "NameFa"),
                EnName = GetString(first, "EnBrandName", "enBrandName", "EnglishName", "englishName", "NameEn"),
                GenericCode = GetString(first, "DrugGenericCode", "drugGenericCode", "GenericCode", "genericCode"),
                PackageCount = GetString(first, "PackageCount", "packageCount", "PackCount"),
                BrandOwner = GetString(first, "FaBrandOwnerName", "faBrandOwnerName", "BrandOwnerFa"),
                ProductType = GetString(first, "ProductType", "productType"),
            };

            result.ConsumerPricePerUnit = GetDecimal(first, "ConsumerPrice", "consumerPrice");
            decimal pack = GetDecimal(first, "PackageCount", "packageCount", "PackCount");
            if (result.ConsumerPricePerUnit > 0)
            {
                // طبق درخواست کاربر: قیمت مصرف‌کننده = تعداد در بسته × قیمت هر واحد (به ریال)
                result.TotalPriceRial = pack > 0 ? result.ConsumerPricePerUnit * pack : result.ConsumerPricePerUnit;
            }

            // نوع فرآورده باید «زیرفرآورده دارویی» باشد؛ در غیر این صورت قیمت استنادی ندارد
            string normalized = NormalizePersian(result.ProductType);
            bool isDrugSub = normalized.Contains(NormalizePersian("زیرفرآورده دارویی"))
                          || normalized.Contains(NormalizePersian("زیر فرآورده دارویی"));
            result.FoundButNotDrugSubgroup = !isDrugSub && !string.IsNullOrWhiteSpace(result.ProductType);

            return result;
        }
        catch (Exception ex)
        {
            return new PriceResult { Success = false, Message = "❌ خطا در ارتباط با تی‌تک: " + ex.Message };
        }
    }

    /// <summary>مسیر کامل بارکد → قیمت (کاتالوگ + انتخاب اولین فرآورده)</summary>
    public async Task<PriceResult> LookupByBarcodeAsync(string barcode, string? token = null)
    {
        string? irc = await GetIrcFromBarcodeAsync(barcode, token);
        if (string.IsNullOrWhiteSpace(irc))
            return new PriceResult { Success = false, Message = "❌ این بارکد در کاتالوگ فرآورده‌های تی‌تک پیدا نشد" };

        var products = await SearchProductsAsync(irc, token);
        if (products.Count == 0)
            return new PriceResult { Success = false, Message = "❌ فرآورده‌ای برای IRC «" + irc + "» یافت نشد" };

        return await GetProductDetailsAsync(products[0].ProductId, token);
    }

    // ---------- Parse helpers ----------

    private List<ProductSummary> ParseCompactList(string json)
    {
        var list = new List<ProductSummary>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement? items = ExtractArray(root, "data", "Data", "items", "Items", "result", "Result", "list", "List");
            if (items is null)
            {
                if (root.ValueKind == JsonValueKind.Array)
                    items = root;
                else if (root.ValueKind == JsonValueKind.Object)
                    items = ExtractAnyArray(root);
            }
            if (items is null)
                return list;

            foreach (var item in items.Value.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                    continue;

                long id = (long)GetDecimal(item, "ProductId", "productId", "id", "Id");
                if (id <= 0)
                    continue;

                string fa = GetString(item, "FaBrandName", "faBrandName", "PersianName", "persianName", "nameFa");
                string en = GetString(item, "EnBrandName", "enBrandName", "EnglishName", "englishName");
                string irc = GetString(item, "Irc", "irc");
                string owner = GetString(item, "FaBrandOwnerName", "faBrandOwnerName");

                string title = fa;
                if (string.IsNullOrWhiteSpace(title))
                    title = en;
                if (string.IsNullOrWhiteSpace(title))
                    title = "فرآورده " + id;

                string subtitle = owner;
                if (!string.IsNullOrWhiteSpace(irc))
                    subtitle = string.IsNullOrWhiteSpace(subtitle) ? "IRC: " + irc : subtitle + " | IRC: " + irc;

                list.Add(new ProductSummary { ProductId = id, Title = title, Subtitle = subtitle });
            }
        }
        catch
        {
        }
        return list;
    }

    /// <summary>اولین آرایه‌ای که داخل آبجکت پیدا شود (برای فرمت‌های نامعلوم)</summary>
    private static JsonElement? ExtractAnyArray(JsonElement obj)
    {
        foreach (var prop in obj.EnumerateObject())
        {
            if (prop.Value.ValueKind == JsonValueKind.Array && prop.Value.GetArrayLength() > 0)
            {
                // آرایه‌های metadata مثل totalCount را رد کن — فقط آرایه‌ای که آبجکت دارد
                if (prop.Value[0].ValueKind == JsonValueKind.Object)
                    return prop.Value;
            }
        }
        return null;
    }

    private static JsonElement? ExtractArray(JsonElement obj, params string[] names)
    {
        foreach (var name in names)
        {
            if (obj.ValueKind == JsonValueKind.Object && obj.TryGetProperty(name, out var el))
            {
                if (el.ValueKind == JsonValueKind.Array)
                    return el;
                if (el.ValueKind == JsonValueKind.Object)
                {
                    // بعضی APIها { data: { items: [...] } } برمی‌گردانند
                    foreach (var inner in names)
                    {
                        if (el.TryGetProperty(inner, out var innerEl) && innerEl.ValueKind == JsonValueKind.Array)
                            return innerEl;
                    }
                }
            }
        }
        return null;
    }

    private static string GetString(JsonElement obj, params string[] names)
    {
        foreach (var name in names)
        {
            if (obj.TryGetProperty(name, out var el))
            {
                if (el.ValueKind == JsonValueKind.String)
                    return el.GetString() ?? "";
                if (el.ValueKind == JsonValueKind.Number)
                    return el.GetRawText();
            }
        }
        return "";
    }

    private static decimal GetDecimal(JsonElement obj, params string[] names)
    {
        foreach (var name in names)
        {
            if (obj.TryGetProperty(name, out var el))
            {
                if (el.ValueKind == JsonValueKind.Number && el.TryGetDecimal(out var d))
                    return d;
                if (el.ValueKind == JsonValueKind.String)
                {
                    string s = (el.GetString() ?? "").Replace(",", "").Trim();
                    // قیمت‌های لاتین یا فارسی با اعداد
                    s = ToEnglishDigits(s);
                    if (decimal.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                        return parsed;
                }
            }
        }
        return 0m;
    }

    private static string ToEnglishDigits(string s)
    {
        var sb = new System.Text.StringBuilder(s.Length);
        foreach (char c in s)
        {
            if (c >= '۰' && c <= '۹') sb.Append((char)(c - '۰' + '0'));
            else if (c >= '٠' && c <= '٩') sb.Append((char)(c - '٠' + '0'));
            else sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>نرمال‌سازی متن فارسی برای مقایسه (حذف نیم‌فاصله، فاصله، ی/ک عربی)</summary>
    private static string NormalizePersian(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        return s
            .Replace("\u200c", "")
            .Replace(" ", "")
            .Replace("ي", "ی")
            .Replace("ك", "ک")
            .Trim();
    }
}
