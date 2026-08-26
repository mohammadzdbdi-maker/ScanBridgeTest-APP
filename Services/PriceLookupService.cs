using System.Linq;
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
        // اجزای تفکیک‌شده‌ی اسم برای نمایش باکس‌به‌باکس و مرتب‌سازی
        public string Brand { get; set; } = "";
        public string Form { get; set; } = "";
        public string Dose { get; set; } = "";
        // فیلدهای کامل محصول — باید جدا ذخیره شوند؛ نمی‌شود به JsonElement بعد از Dispose سند تکیه کرد
        public string EnName { get; set; } = "";
        public string GenericCode { get; set; } = "";
        public string PackageCount { get; set; } = "";
        public string BrandOwner { get; set; } = "";
        public string ProductType { get; set; } = "";
        public string Irc { get; set; } = "";
        // اگر از GetProductsForPharmacies آمده باشد، قیمت و اطلاعات کامل همین‌جا موجود است
        public decimal ConsumerPricePerUnit { get; set; }
        public decimal TotalPriceRial { get; set; }
        public JsonElement FullInfo { get; set; }
        public bool HasDirectPrice => ConsumerPricePerUnit > 0;
        public bool HasProductInfo =>
            !string.IsNullOrWhiteSpace(Title)
            && !Title.StartsWith("فرآورده ", StringComparison.Ordinal);
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

    /// <summary>
    /// لیست کومپکت تی‌تک بعضی وقت‌ها فقط ProductId و Irc دارد (بدون اسم). اگر اسمی نیامده بود،
    /// برای حداکثر ۱۲ فرآورده‌ی اول، جزئیات را موازی می‌گیریم و اسم فارسی/مالک برند را پر می‌کنیم.
    /// </summary>
    public async Task<List<ProductSummary>> EnrichSummariesWithDetailsAsync(List<ProductSummary> items, string? token = null)
    {
        var needs = items.Take(12).Where(i => i.Title.StartsWith("فرآورده ")).ToList();
        if (needs.Count == 0)
            return items;

        var tasks = needs.Select(async it =>
        {
            var d = await GetProductDetailsAsync(it.ProductId, token);
            if (d.Success)
            {
                string fa = d.FaName;
                if (string.IsNullOrWhiteSpace(fa)) fa = d.EnName;
                if (!string.IsNullOrWhiteSpace(fa))
                {
                    it.Title = fa;
                    ParseNameParts(fa, d.EnName, out var brand, out var form, out var dose);
                    it.Brand = brand;
                    it.Form = form;
                    it.Dose = dose;
                }
                if (!string.IsNullOrWhiteSpace(d.BrandOwner))
                    it.Subtitle = string.IsNullOrWhiteSpace(it.Subtitle)
                        ? d.BrandOwner
                        : it.Subtitle + " | " + d.BrandOwner;
                it.EnName = d.EnName;
                it.GenericCode = d.GenericCode;
                it.PackageCount = d.PackageCount;
                it.BrandOwner = d.BrandOwner;
                it.ProductType = d.ProductType;
                it.ConsumerPricePerUnit = d.ConsumerPricePerUnit;
                it.TotalPriceRial = d.TotalPriceRial;
            }
        }).ToList();

        try { await Task.WhenAll(tasks); } catch { }
        return items;
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

            // همه‌ی خواندن‌ها باید داخل using انجام شود؛ JsonElement بعد از Dispose سند والد
            // «Cannot access a disposed object» می‌دهد (نتیجه‌ی محاسبه از قبل داخل سند آماده می‌شود).
            PriceResult result;
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

                var first = items.Value[0];

                // ✅ لاگ دیباگ: کلیدهای موجود در پاسخ API را ذخیره کن
                try
                {
                    string debugKeys = string.Join(", ", first.EnumerateObject().Select(p => p.Name));
                    System.Diagnostics.Debug.WriteLine($"[PriceLookup] API keys: {debugKeys}");
                    System.Diagnostics.Debug.WriteLine($"[PriceLookup] Raw first item: {first.GetRawText()}");
                } catch { }

                result = ParsePriceResultFromItem(first);
            }
            catch (Exception ex)
            {
                return new PriceResult { Success = false, Message = "❌ پاسخ نامعتبر از سرور: " + ex.Message };
            }
            // نوع فرآورده باید «زیر فرآورده دارو(یی)» باشد؛ مقدار واقعی تی‌تک «زیر فرآورده دارو»
            // (بدون ی پایانی) است، پس با پیشوند نرمال‌شده‌ی «زیرفرآوردهدارو» چک می‌کنیم.
            string normalized = NormalizePersian(result.ProductType);
            bool isDrugSub = normalized.Contains(NormalizePersian("زیر فرآورده دارو"));
            result.FoundButNotDrugSubgroup = !isDrugSub && !string.IsNullOrWhiteSpace(result.ProductType);

            return result;
        }
        catch (Exception ex)
        {
            return new PriceResult { Success = false, Message = "❌ خطا در ارتباط با تی‌تک: " + ex.Message };
        }
    }

    /// <summary>
    /// ساخت نتیجهٔ کامل از یک ردیف لیست (جست‌وجوی ژنریک/نام) بدون مراجعهٔ دوباره به API.
    /// </summary>
    public PriceResult ToPriceResult(ProductSummary item)
    {
        string ptype = item.ProductType ?? "";
        string normalized = NormalizePersian(ptype);
        bool isDrugSub = string.IsNullOrWhiteSpace(ptype)
                         || normalized.Contains(NormalizePersian("زیر فرآورده دارو"));
        return new PriceResult
        {
            Success = true,
            FaName = item.Title,
            EnName = item.EnName,
            GenericCode = item.GenericCode,
            PackageCount = item.PackageCount,
            BrandOwner = item.BrandOwner,
            ProductType = ptype,
            ConsumerPricePerUnit = item.ConsumerPricePerUnit,
            TotalPriceRial = item.TotalPriceRial,
            FoundButNotDrugSubgroup = !isDrugSub && !string.IsNullOrWhiteSpace(ptype),
        };
    }

    /// <summary>
    /// جست‌وجوی کد ژنریک: مثل سایت تی‌تک، فیلتر روی «کد ژنریک». چند نام پارامتر محتمل را
    /// به‌ترتیب امتحان می‌کنیم تا یکی جواب دهد (GenericCode، DrugGenericCode، Name=کد، searchExp).
    /// </summary>
    public async Task<List<ProductSummary>> SearchByGenericCodeAsync(string code, string? token = null)
    {
        code = ToEnglishDigits(code ?? "").Trim();
        if (string.IsNullOrWhiteSpace(code))
            return new List<ProductSummary>();

        // پارامتر درست که از HAR واقعی درآمد: DrugGenericCode روی GetProductsForPharmacies
        var list = await TryFetchPharmaciesListAsync($"searchExp=&DrugGenericCode={Uri.EscapeDataString(code)}", token);
        if (list.Count > 0)
            return list;

        list = await TryFetchPharmaciesListAsync($"searchExp=&GenericCode={Uri.EscapeDataString(code)}", token);
        if (list.Count > 0)
            return list;

        // اگر پاسخ خالی بود، شاید کد به‌صورت Name در لیست فشرده پیدا شود
        try
        {
            string url = $"{CompactProductsUrl}?Name={Uri.EscapeDataString(code)}&PageSize=50&PageNumber=1";
            using var req = BuildGet(url, token);
            using var resp = await _http.SendAsync(req);
            string body = await resp.Content.ReadAsStringAsync();
            if (resp.IsSuccessStatusCode && !string.IsNullOrWhiteSpace(body))
            {
                list = ParseCompactList(body);
                if (list.Count > 0)
                    return list;
            }
        }
        catch { }

        return new List<ProductSummary>();
    }

    private async Task<List<ProductSummary>> TryFetchPharmaciesListAsync(string query, string? token)
    {
        try
        {
            string url = $"{ProductsForPharmaciesUrl}?{query}&PageSize=50&PageNumber=1";
            using var req = BuildGet(url, token);
            using var resp = await _http.SendAsync(req);
            string body = await resp.Content.ReadAsStringAsync();
            if (resp.IsSuccessStatusCode && !string.IsNullOrWhiteSpace(body))
                return ParseFullProductsList(body);
        }
        catch { }
        return new List<ProductSummary>();
    }

    /// <summary>پارس پاسخ GetProductsForPharmacies — همه‌ی فیلدها (اسم، شکل، دوز، IRC، ژنریک، مالک) ذخیره می‌شود</summary>
    private List<ProductSummary> ParseFullProductsList(string json)
    {
        var list = new List<ProductSummary>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var items = ExtractArray(root, "Result", "result", "data", "Data", "items", "Items")
                        ?? (root.ValueKind == JsonValueKind.Array ? root : ExtractAnyArray(root));
            if (items is null)
                return list;

            foreach (var item in items.Value.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                    continue;

                var parsed = ParsePriceResultFromItem(item);
                long id = (long)GetDecimal(item, "Id", "id", "ProductId", "productId");
                string irc = GetString(item, "Irc", "irc", "IRC");

                string title = !string.IsNullOrWhiteSpace(parsed.FaName)
                    ? parsed.FaName
                    : (!string.IsNullOrWhiteSpace(parsed.EnName)
                        ? parsed.EnName
                        : (!string.IsNullOrWhiteSpace(irc) ? "فرآورده " + irc : (id > 0 ? "فرآورده " + id : "فرآورده")));
                ParseNameParts(title, parsed.EnName, out var brand, out var form, out var dose);

                JsonElement full = default;
                try { full = item.Clone(); } catch { }

                string subtitle = !string.IsNullOrWhiteSpace(irc) ? "IRC: " + irc : parsed.BrandOwner;

                list.Add(new ProductSummary
                {
                    ProductId = id,
                    Title = title,
                    Subtitle = subtitle,
                    Brand = brand,
                    Form = form,
                    Dose = dose,
                    EnName = parsed.EnName,
                    GenericCode = parsed.GenericCode,
                    PackageCount = parsed.PackageCount,
                    BrandOwner = parsed.BrandOwner,
                    ProductType = parsed.ProductType,
                    Irc = irc,
                    ConsumerPricePerUnit = parsed.ConsumerPricePerUnit,
                    TotalPriceRial = parsed.TotalPriceRial,
                    FullInfo = full,
                });
            }
        }
        catch { }
        return list;
    }

    /// <summary>خواندن همهٔ فیلدهای نمایشی از یک آیتم JSON تی‌تک (قبل از Dispose سند)</summary>
    private static PriceResult ParsePriceResultFromItem(JsonElement first)
    {
        decimal unitPrice = GetDecimal(first, "ConsumerPrice", "consumerPrice", "UnitPrice", "unitPrice", "Price", "price");
        decimal pack = GetDecimal(first, "PackageCount", "packageCount", "PackCount", "packCount", "PackageQty", "packageQty", "Qty", "qty");

        return new PriceResult
        {
            Success = true,
            FaName = GetString(first, "FaBrandName", "faBrandName", "PersianName", "persianName", "NameFa", "PersianProductName", "persianProductName", "ProductNameFa", "productNameFa", "Name", "name", "Title", "title", "FaName", "faName"),
            EnName = GetString(first, "EnBrandName", "enBrandName", "EnglishName", "englishName", "NameEn", "EnglishProductName", "englishProductName", "ProductNameEn", "productNameEn", "EnName", "enName", "TitleEn", "titleEn"),
            GenericCode = GetString(first, "DrugGenericCode", "drugGenericCode", "GenericCode", "genericCode", "DrugCode", "drugCode", "GenericCodeStr", "genericCodeStr"),
            PackageCount = GetString(first, "PackageCount", "packageCount", "PackCount", "packCount", "PackageQty", "packageQty", "PackSize", "packSize", "Qty", "qty", "PackageQuantity", "packageQuantity"),
            BrandOwner = GetString(first, "FaBrandOwnerName", "faBrandOwnerName", "BrandOwnerFa", "BrandOwner", "brandOwner", "Manufacturer", "manufacturer", "CompanyName", "companyName", "OwnerName", "ownerName", "FaCompanyName", "faCompanyName"),
            ProductType = GetString(first, "ProductType", "productType", "Type", "type", "Category", "category"),
            ConsumerPricePerUnit = unitPrice,
            TotalPriceRial = unitPrice > 0 && pack > 0 ? unitPrice * pack : unitPrice,
        };
    }

    /// <summary>
    /// جست‌وجوی آزاد (برای هر عبارت): اول با compact Name؛ اگر نتیجه نداشت
    /// با searchExp روی GetProductsForPharmacies.
    /// </summary>
    public async Task<List<ProductSummary>> SearchByExpressionAsync(string expr, string? token = null)
    {
        var viaCompact = await SearchProductsAsync(expr, token);
        if (viaCompact.Count > 0)
            return viaCompact;

        var list = new List<ProductSummary>();
        try
        {
            string url = $"{ProductsForPharmaciesUrl}?searchExp={Uri.EscapeDataString(expr)}&ProductId=&PageSize=50&PageNumber=1";
            using var req = BuildGet(url, token);
            using var resp = await _http.SendAsync(req);
            string body = await resp.Content.ReadAsStringAsync();
            if (!resp.IsSuccessStatusCode || string.IsNullOrWhiteSpace(body))
                return list;

            using var doc = JsonDocument.Parse(body);
            var root = doc.RootElement;
            var items = ExtractArray(root, "data", "Data", "items", "Items", "result", "Result", "list", "List")
                        ?? (root.ValueKind == JsonValueKind.Array ? root : ExtractAnyArray(root));
            if (items is null)
                return list;

            foreach (var item in items.Value.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                    continue;
                long id = (long)GetDecimal(item, "Id", "id", "ProductId", "productId");
                if (id <= 0)
                    continue;
                string fa = GetString(item, "FaBrandName", "faBrandName", "PersianName", "persianName");
                string en = GetString(item, "EnBrandName", "enBrandName", "EnglishName", "englishName");
                string irc = GetString(item, "Irc", "irc");
                string title = !string.IsNullOrWhiteSpace(fa) ? fa : (!string.IsNullOrWhiteSpace(en) ? en : "فرآورده " + id);
                ParseNameParts(title, en, out var brand, out var form, out var dose);
                list.Add(new ProductSummary
                {
                    ProductId = id,
                    Title = title,
                    Subtitle = string.IsNullOrWhiteSpace(irc) ? "" : "IRC: " + irc,
                    Brand = brand,
                    Form = form,
                    Dose = dose,
                    EnName = en,
                    Irc = irc,
                });
            }
            return list;
        }
        catch
        {
            return list;
        }
    }

    /// <summary>
    /// مسیر سریع: اگر IRC از قبل معلوم است (مثلاً از استعلام اولیه تی‌تک بعد از اسکن)،
    /// مستقیم جست‌وجو کن بدون رفتن به InstanceCatalog.
    /// </summary>
    public async Task<PriceResult> LookupByIrcAsync(string irc, string? token = null)
    {
        var products = await SearchProductsAsync(irc, token);
        if (products.Count == 0)
            return new PriceResult { Success = false, Message = "❌ فرآورده‌ای برای IRC «" + irc + "» یافت نشد" };
        return await GetProductDetailsAsync(products[0].ProductId, token);
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

    // ---------- تفکیک اسم فرآورده به برند / شکل دارویی / دوز ----------

    private static readonly string[] KnownForms =
    {
        "سرنگ پیش‌پر شده", "سرنگ پیش پر شده", "محلول تزریقی", "پودر تزریقی", "سوسپانسیون تزریقی",
        "قطره چشمی", "قطره گوشی", "قطره بینی", "پماد چشمی", "ژل چشمی",
        "اسپری بینی", "اسپری دهانی", "اسپری تنفسی", "محلول استنشاقی",
        "قرص پراکنده شونده", "قرص آهسته رهش", "پیوسته رهش", "سافت ژل",
        "قرص روکشدار", "قرص جوشان", "قرص جویدنی", "قرص زیرزبانی",
        "شیاف واژینال", "کرم واژینال", "قرص واژینال", "انما مقعدی", "کپسول نرم",
        "پرنترال", "Parenteral", "تزریقی", "Injection", "انما", "Enema", "Rectal Enema",
        "لیوفیلیزه", "امولسیون", "الگزیر", "الیکسیر", "Elixir",
        "افشانه", "دهانشویه", "شامپو", "فوم", "خمیر", "روغن", "پچ",
        "کارتریج", "آئروسل", "نبولایزر",
        "قرص", "کپسول", "شربت", "آمپول", "قطره", "پماد", "کرم", "ژل",
        "اسپری", "سوسپانسیون", "محلول", "پودر", "شیاف", "ویال", "مایع", "انفوزیون", "تری گرم",
        "ماندگار", "مواد مؤثره", "اشکال", "ساشه", "سرم", "گرانول", "لوسیون", "اینهیلر",
        "Tablet", "Capsule", "Ampoule", "Ampule", "Vial", "Syrup", "Suspension",
        "Cream", "Ointment", "Gel", "Drop", "Spray", "Powder", "Solution",
        "Infusion", "Sachet", "Suppository", "Inhaler", "Lotion", "Granule"
    };

    private const string DoseUnitPattern =
        @"mg|mcg|µg|ug|μg|g|kg|ml|µl|μl|ul|l|iu|ui|cc|mmol|meq|%|٪|میلی‌گرم|میلی گرم|میلی‌لیتر|میلی لیتر|واحد";

    /// <summary>
    /// دوز کامل را برمی‌گرداند؛ ترجیح با اسم انگلیسی است تا چیزهایی مثل «5 mg/ml 2 ml»
    /// به «5 mg» قیچی نشود.
    /// </summary>
    public static string ExtractFullDose(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return "";

        string pattern =
            @"(?i)(?:\d+(?:[.,]\d+)?(?:\s*/\s*\d+(?:[.,]\d+)?)*\s*(?:" + DoseUnitPattern + @")"
            + @"(?:\s*/\s*(?:\d+(?:[.,]\d+)?\s*)?(?:" + DoseUnitPattern + @"))?"
            + @"(?:\s+\d+(?:[.,]\d+)?(?:\s*/\s*\d+(?:[.,]\d+)?)*\s*(?:" + DoseUnitPattern + @"))*)";

        var matches = System.Text.RegularExpressions.Regex.Matches(text, pattern);
        if (matches.Count == 0)
            return "";

        string dose = matches[matches.Count - 1].Value;
        for (int i = matches.Count - 2; i >= 0; i--)
        {
            int gapStart = matches[i].Index + matches[i].Length;
            int gapLen = matches[i + 1].Index - gapStart;
            if (gapLen < 0)
                break;
            string gap = text.Substring(gapStart, gapLen);
            if (!System.Text.RegularExpressions.Regex.IsMatch(gap, @"^[\s/×xX]*$"))
                break;
            dose = matches[i].Value + gap + dose;
        }

        return NormalizeDose(dose);
    }

    private static string NormalizeDose(string dose)
    {
        dose = System.Text.RegularExpressions.Regex.Replace(dose.Trim(), @"\s+", " ");
        dose = System.Text.RegularExpressions.Regex.Replace(dose, @"\s*/\s*", "/");
        dose = System.Text.RegularExpressions.Regex.Replace(
            dose,
            @"(?<=\d)(?=" + DoseUnitPattern + @")",
            " ",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return dose.Trim();
    }

    /// <summary>
    /// اسم تی‌تک به شکل «برند + شکل دارویی + دوز» است.
    /// دوز ترجیحاً از اسم انگلیسی خوانده می‌شود تا کامل بماند.
    /// </summary>
    public static void ParseNameParts(string title, out string brand, out string form, out string dose)
        => ParseNameParts(title, null, out brand, out form, out dose);

    public static void ParseNameParts(string? title, string? enName, out string brand, out string form, out string dose)
    {
        title = title?.Trim() ?? "";
        enName = enName?.Trim() ?? "";
        brand = title;
        form = "";
        dose = "";

        if (!string.IsNullOrWhiteSpace(enName))
            dose = ExtractFullDose(enName);
        if (string.IsNullOrWhiteSpace(dose))
            dose = ExtractFullDose(title);

        string formSource = !string.IsNullOrWhiteSpace(title) ? title : enName;
        if (string.IsNullOrWhiteSpace(formSource))
            return;

        int formIdx = -1;
        string foundForm = "";
        foreach (var f in KnownForms)
        {
            int idx = formSource.IndexOf(f, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0 && (formIdx < 0 || idx < formIdx))
            {
                formIdx = idx;
                foundForm = f;
            }
        }

        if (formIdx < 0 && !string.IsNullOrWhiteSpace(enName) && !string.Equals(formSource, enName, StringComparison.Ordinal))
        {
            foreach (var f in KnownForms)
            {
                int idx = enName.IndexOf(f, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0 && (formIdx < 0 || idx < formIdx))
                {
                    formIdx = idx;
                    foundForm = f;
                    formSource = enName;
                }
            }
        }

        string splitDose = ExtractFullDose(formSource);

        if (formIdx < 0)
        {
            brand = string.IsNullOrWhiteSpace(splitDose)
                ? formSource
                : formSource.Replace(splitDose, "", StringComparison.OrdinalIgnoreCase).Trim();
            return;
        }

        brand = formSource.Substring(0, formIdx).Trim();
        string rest = formSource.Substring(formIdx);
        int doseIdx = -1;
        if (!string.IsNullOrWhiteSpace(splitDose))
            doseIdx = rest.IndexOf(splitDose, StringComparison.OrdinalIgnoreCase);
        if (doseIdx < 0)
        {
            var num = System.Text.RegularExpressions.Regex.Match(rest, @"\d");
            if (num.Success)
                doseIdx = num.Index;
        }

        form = doseIdx >= 0 ? rest.Substring(0, doseIdx).Trim() : rest.Trim();
        if (string.IsNullOrWhiteSpace(form))
            form = foundForm;
    }

    /// <summary>عدد دوز برای مرتب‌سازی عددی (مثلاً «37.5 mg» → 37.5)</summary>
    public static decimal DoseValue(string dose)
    {
        var m = System.Text.RegularExpressions.Regex.Match(dose ?? "", @"\d+(?:[.,]\d+)?");
        if (m.Success && decimal.TryParse(m.Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
            return v;
        return 0m;
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

                string fa = GetString(item, "Name", "name", "FaBrandName", "faBrandName", "PersianName", "persianName", "nameFa", "FaName", "faName", "ProductName", "productName", "Title", "title");
                string en = GetString(item, "EnBrandName", "enBrandName", "EnglishName", "englishName", "EnName", "enName");
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

                ParseNameParts(title, en, out var brand, out var form, out var dose);

                list.Add(new ProductSummary
                {
                    ProductId = id,
                    Title = title,
                    Subtitle = subtitle,
                    Brand = brand,
                    Form = form,
                    Dose = dose,
                    EnName = en,
                    BrandOwner = owner,
                    Irc = irc,
                });
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
