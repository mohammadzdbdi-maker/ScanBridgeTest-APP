using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.Json;
using System.Linq;
using System.Text.RegularExpressions;

namespace ScanBridgeTest;

public class DrugLookupService
{
    private static HttpClient? _httpClient;
    private static CookieContainer? _cookieContainer;

    // صفحه‌ای که کاربر در سایت تی‌تک می‌بیند:
    private const string MOBILE_ENTER_UID_URL = "https://mobile.ttac.ir/enterUID";

    // API پنل جدید تی‌تک برای کاتالوگ فرآورده. اگر پاسخ بدهد، اطلاعات کامل‌تری می‌دهد.
    private const string INSTANCE_CATALOG_URL = "https://statisticsreports.ttac.ir/product/InstanceCatalog";

    // API پشت همان صفحه enterUID برای دریافت اطلاعات محصول:
    private const string NEW_API_URL = "https://newapi.ttac.ir/irfdamobile/v1/checkuId";
    private const string API_KEY = "e4a3abe1-9725-4aea-9eb8-3843aad4061f";

    public DrugLookupService()
    {
    }

    private void InitializeHttpClient()
    {
        _cookieContainer = new CookieContainer();
        var handler = new HttpClientHandler
        {
            CookieContainer = _cookieContainer,
            AllowAutoRedirect = true,
            UseCookies = true,
            ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };

        _httpClient = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        SetDefaultHeaders();
    }

    private void SetDefaultHeaders()
    {
        if (_httpClient == null) return;

        _httpClient.DefaultRequestHeaders.Clear();
        _httpClient.DefaultRequestHeaders.Add("User-Agent",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36");
        _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
        _httpClient.DefaultRequestHeaders.Add("Accept-Language", "fa-IR,fa;q=0.9");
        _httpClient.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate, br");
        _httpClient.DefaultRequestHeaders.Add("Origin", "https://mobile.ttac.ir");
        _httpClient.DefaultRequestHeaders.Add("Referer", MOBILE_ENTER_UID_URL);
        _httpClient.DefaultRequestHeaders.Add("x-ssp-api-key", API_KEY);
        _httpClient.DefaultRequestHeaders.Add("Sec-Fetch-Dest", "empty");
        _httpClient.DefaultRequestHeaders.Add("Sec-Fetch-Mode", "cors");
        _httpClient.DefaultRequestHeaders.Add("Sec-Fetch-Site", "same-site");
    }

    /// <summary>
    /// استخراج UID بیست رقمی تی‌تک بدون دستکاری مخرب بارکد اصلی.
    /// چند ساختار رایج اسکنرها را پوشش می‌دهد: UID خالص، GS1 با AI های 01 و 21،
    /// وجود کاراکترهای کنترلی FNC1/GS، و پیشوندهای اسکنر مثل ]d2 یا ]C1.
    /// </summary>
    private string ExtractUIDFromBarcode(string barcode)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(barcode))
                return "";

            string raw = barcode.Trim();

            // فقط کاراکترهای کنترلی اسکنر مانند GS/FNC1 حذف می‌شوند؛ خود رشته اصلی برای لاگ حفظ می‌شود.
            string cleaned = new string(raw.Where(c => !char.IsControl(c)).ToArray()).Trim();
            cleaned = cleaned
                .Replace("]d2", "", StringComparison.OrdinalIgnoreCase)
                .Replace("]C1", "", StringComparison.OrdinalIgnoreCase)
                .Replace("]e0", "", StringComparison.OrdinalIgnoreCase)
                .Replace(" ", "")
                .Trim();

            // نسخه فقط عددی برای الگوهای GS1. پرانتزها، جداکننده‌ها و پیشوندهای اسکنر حذف می‌شوند.
            string digitsOnly = new string(cleaned.Where(char.IsDigit).ToArray());

            LogMessage($"🧾 بارکد خام دریافتی: {raw}");
            LogMessage($"🧾 بارکد عددی برای استخراج UID: {digitsOnly}");

            // اگر خروجی اسکنر خودش UID بیست رقمی باشد.
            if (TryReturnUid(digitsOnly, "بارکد ورودی خودش UID ۲۰ رقمی است", out var uid))
                return uid;

            // الگوی استاندارد: AI 01 + GTIN چهارده رقمی + AI 21 + UID بیست رقمی
            // مثال: 01xxxxxxxxxxxxxx21yyyyyyyyyyyyyyyyyyyy
            var gs1Match = Regex.Match(digitsOnly, @"01\d{14}21(?<uid>\d{20})");
            if (gs1Match.Success && TryReturnUid(gs1Match.Groups["uid"].Value, "UID از ساختار استاندارد GS1/AI01/AI21 استخراج شد", out uid))
                return uid;

            // بعضی اسکنرها یا تنظیمات، بخش قبل از AI21 را متفاوت ارسال می‌کنند؛ در این حالت بعد از AI21 را می‌خوانیم.
            var ai21Matches = Regex.Matches(digitsOnly, @"21(?<uid>\d{20})");
            foreach (Match match in ai21Matches)
            {
                if (TryReturnUid(match.Groups["uid"].Value, "UID از AI21 استخراج شد", out uid))
                    return uid;
            }

            // سازگاری با منطق قبلی: ۱۸ کاراکتر بعد از 01، UID بیست رقمی شروع می‌شود.
            int idx01 = digitsOnly.IndexOf("01", StringComparison.Ordinal);
            if (idx01 != -1 && digitsOnly.Length >= idx01 + 38)
            {
                string legacyUid = digitsOnly.Substring(idx01 + 18, 20);
                if (TryReturnUid(legacyUid, "UID با روش قبلی استخراج شد", out uid))
                    return uid;
            }

            // آخرین راه‌حل: اگر فقط یک توالی ۲۰ رقمی قابل قبول در بارکد وجود داشت، همان UID در نظر گرفته می‌شود.
            var twentyDigitCandidates = Regex.Matches(digitsOnly, @"\d{20}")
                .Cast<Match>()
                .Select(m => m.Value)
                .Distinct()
                .ToList();

            if (twentyDigitCandidates.Count == 1 && TryReturnUid(twentyDigitCandidates[0], "تنها توالی ۲۰ رقمی موجود به عنوان UID انتخاب شد", out uid))
                return uid;

            LogMessage($"⚠️ ساختار بارکد معتبر تی‌تک یافت نشد. cleaned={cleaned}, digits={digitsOnly}");
            return "";
        }
        catch (Exception ex)
        {
            LogMessage($"❌ خطا در استخراج UID: {ex.Message}");
            return "";
        }
    }

    private static bool IsValidUid(string uid)
    {
        return uid.Length == 20 && uid.All(char.IsDigit);
    }

    private bool TryReturnUid(string candidate, string reason, out string uid)
    {
        uid = candidate?.Trim() ?? "";
        if (!IsValidUid(uid))
            return false;

        LogMessage($"✅ {reason}: {uid}");
        LogMessage($"  • کادر اول (۱۰ رقم): {uid.Substring(0, 10)}");
        LogMessage($"  • کادر دوم (۱۰ رقم): {uid.Substring(10, 10)}");
        return true;
    }

    public async Task<DrugInfo> GetDrugNameAsync(string barcode, string? ttacAccessToken = null)
    {
        try
        {
            LogMessage("");
            LogMessage(new string('=', 60));
            LogMessage($"🔍 شروع جستجو برای بارکد: {barcode}");
            LogMessage(new string('=', 60));

            InitializeHttpClient();

            // اول از API جدید پنل کاتالوگ استفاده می‌کنیم؛ اگر در دسترس نبود، با API موبایل قبلی ادامه می‌دهیم.
            var instanceCatalogResult = await SearchDrugByInstanceCatalogAsync(barcode, ttacAccessToken);
            if (instanceCatalogResult.Success)
            {
                LogMessage(new string('=', 60));
                return instanceCatalogResult;
            }

            string uid = ExtractUIDFromBarcode(barcode);

            if (string.IsNullOrWhiteSpace(uid))
            {
                return new DrugInfo
                {
                    Success = false,
                    Message = "❌ ساختار UID در بارکد معتبر نیست",
                    OriginalBarcode = barcode
                };
            }

            var result = await SearchDrugByUIDAsync(uid, barcode);

            LogMessage(new string('=', 60));
            return result;
        }
        catch (Exception ex)
        {
            LogMessage($"❌ خطا: {ex.Message}");
            return new DrugInfo
            {
                Success = false,
                Message = $"❌ خطا: {ex.Message}",
                OriginalBarcode = barcode
            };
        }
    }

    private async Task<DrugInfo> SearchDrugByInstanceCatalogAsync(string originalBarcode, string? ttacAccessToken = null)
    {
        try
        {
            string url = $"{INSTANCE_CATALOG_URL}?uid={Uri.EscapeDataString(originalBarcode)}";
            LogMessage($"🔗 درخواست به API کاتالوگ پنل تی‌تک: {url}");

            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.TryAddWithoutValidation("Accept", "application/json, text/plain, */*");
            request.Headers.TryAddWithoutValidation("Origin", "https://newstatisticsreports.ttac.ir");
            request.Headers.TryAddWithoutValidation("Referer", "https://newstatisticsreports.ttac.ir/pharmacyDashboard/instanceCatalog");
            if (!string.IsNullOrWhiteSpace(ttacAccessToken))
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", ttacAccessToken);
                LogMessage("🔐 درخواست کاتالوگ با توکن ورود تی‌تک ارسال می‌شود.");
            }

            var response = await _httpClient!.SendAsync(request);
            string responseContent = await response.Content.ReadAsStringAsync();
            LogMessage($"📊 کد وضعیت پاسخ کاتالوگ: {response.StatusCode}");

            if (response.IsSuccessStatusCode && !string.IsNullOrWhiteSpace(responseContent))
            {
                var parsed = ParseInstanceCatalogResponse(responseContent, originalBarcode);
                if (parsed.Success)
                    return parsed;
            }

            return new DrugInfo
            {
                Success = false,
                Message = $"❌ پاسخ کاتالوگ: {response.StatusCode}",
                OriginalBarcode = originalBarcode
            };
        }
        catch (Exception ex)
        {
            LogMessage($"⚠️ API کاتالوگ در دسترس نبود، تلاش با API موبایل ادامه می‌یابد: {ex.Message}");
            return new DrugInfo
            {
                Success = false,
                Message = $"❌ خطا در API کاتالوگ: {ex.Message}",
                OriginalBarcode = originalBarcode
            };
        }
    }

    private DrugInfo ParseInstanceCatalogResponse(string json, string originalBarcode)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            bool success = root.TryGetProperty("Success", out var successElement) &&
                           successElement.ValueKind == JsonValueKind.True;
            if (!success)
            {
                string? message = ExtractString(root, "Message");
                return new DrugInfo { Success = false, Message = message ?? "❌ کاتالوگ پاسخ موفق نداد", OriginalBarcode = originalBarcode };
            }

            if (!root.TryGetProperty("Result", out var data) || data.ValueKind != JsonValueKind.Object)
                return new DrugInfo { Success = false, Message = "❌ داده کاتالوگ یافت نشد", OriginalBarcode = originalBarcode };

            string? persianName = ExtractString(data, "PersianName");
            string? englishName = ExtractString(data, "EnglishName");
            string? gtin = ExtractString(data, "GTIN");
            string? irc = ExtractString(data, "Irc");
            string? uid = ExtractString(data, "UID");
            string? batchCode = ExtractString(data, "BatchCode");
            string? genericName = ExtractString(data, "GenericName");
            string? genericCode = ExtractString(data, "GenericCode");
            string? packageCount = ExtractString(data, "PackageCount");
            string? expiration = ExtractFirstString(data, "Expiration", "expiration", "ExpireDate", "expireDate", "ExpiryDate", "expiryDate", "ExpDate", "expDate");
            string? manufacturing = ExtractFirstString(data, "Manufacturing", "manufacturing", "ProductionDate", "productionDate", "ManufactureDate", "manufactureDate", "MfgDate", "mfgDate");

            var extraFields = ExtractSimpleFields(data);
            if (!string.IsNullOrWhiteSpace(uid)) extraFields["UID"] = uid;
            if (!string.IsNullOrWhiteSpace(batchCode)) extraFields["سری ساخت"] = batchCode;
            if (!string.IsNullOrWhiteSpace(genericName)) extraFields["نام ژنریک"] = genericName;
            if (!string.IsNullOrWhiteSpace(genericCode)) extraFields["کد ژنریک"] = genericCode;
            if (!string.IsNullOrWhiteSpace(packageCount)) extraFields["تعداد در بسته"] = packageCount;
            if (!string.IsNullOrWhiteSpace(expiration)) extraFields["تاریخ انقضا"] = expiration;
            if (!string.IsNullOrWhiteSpace(manufacturing)) extraFields["تاریخ ساخت"] = manufacturing;

            LogMessage("📦 اطلاعات کاتالوگ استخراج شده:");
            LogMessage($"  • نام فارسی: {persianName}");
            LogMessage($"  • نام انگلیسی: {englishName}");
            LogMessage($"  • GTIN: {gtin}");
            LogMessage($"  • IRC: {irc}");
            LogMessage($"  • UID: {uid}");

            if (string.IsNullOrWhiteSpace(persianName) && string.IsNullOrWhiteSpace(englishName))
                return new DrugInfo { Success = false, Message = "❌ نام محصول در کاتالوگ مشخص نیست", OriginalBarcode = originalBarcode };

            return new DrugInfo
            {
                Success = true,
                PersianName = persianName ?? "نامشخص",
                EnglishName = englishName ?? "Unknown",
                GTIN = gtin,
                IRC = irc,
                ExtraFields = extraFields,
                Message = $"✅ {persianName ?? englishName}",
                OriginalBarcode = originalBarcode
            };
        }
        catch (Exception ex)
        {
            LogMessage($"❌ خطا در تحلیل پاسخ کاتالوگ: {ex.Message}");
            return new DrugInfo { Success = false, Message = $"❌ خطا در کاتالوگ: {ex.Message}", OriginalBarcode = originalBarcode };
        }
    }

    private async Task<DrugInfo> SearchDrugByUIDAsync(string uid, string originalBarcode)
    {
        try
        {
            string url = $"{NEW_API_URL}?uidCode={uid}&latitude=0&longitude=0&device=2&deviceIdentifier=0301d88c-e06f-40cf-a07f-873c560b1c1c&Platform=0&PlatformVersion=1&VersionCode=&AppVersion=";

            LogMessage($"🔗 درخواست به API پشت صفحه {MOBILE_ENTER_UID_URL}: {url}");

            var response = await _httpClient!.GetAsync(url);
            string responseContent = await response.Content.ReadAsStringAsync();

            LogMessage($"📊 کد وضعیت پاسخ: {response.StatusCode}");

            if (response.IsSuccessStatusCode && !string.IsNullOrWhiteSpace(responseContent))
            {
                return ParseDrugResponse(responseContent, originalBarcode);
            }

            return new DrugInfo
            {
                Success = false,
                Message = $"❌ پاسخ سرور: {response.StatusCode}",
                OriginalBarcode = originalBarcode
            };
        }
        catch (TaskCanceledException ex)
        {
            string message = "❌ اتصال به تی‌تک زمان‌بر شد و پاسخی دریافت نشد. اینترنت یا در دسترس بودن سایت تی‌تک را بررسی کنید.";
            LogMessage($"❌ خطا در ارسال درخواست: {ex.Message}");
            return new DrugInfo
            {
                Success = false,
                Message = message,
                OriginalBarcode = originalBarcode
            };
        }
        catch (HttpRequestException ex)
        {
            string message = BuildFriendlyNetworkErrorMessage(ex);
            LogMessage($"❌ خطا در ارسال درخواست: {ex.Message}");
            return new DrugInfo
            {
                Success = false,
                Message = message,
                OriginalBarcode = originalBarcode
            };
        }
        catch (Exception ex)
        {
            string message = BuildFriendlyNetworkErrorMessage(ex);
            LogMessage($"❌ خطا در ارسال درخواست: {ex.Message}");
            return new DrugInfo
            {
                Success = false,
                Message = message,
                OriginalBarcode = originalBarcode
            };
        }
    }

    private static string BuildFriendlyNetworkErrorMessage(Exception ex)
    {
        string raw = ex.ToString();

        if (raw.Contains("No such host", StringComparison.OrdinalIgnoreCase) ||
            raw.Contains("Name or service not known", StringComparison.OrdinalIgnoreCase) ||
            raw.Contains("nodename nor servname", StringComparison.OrdinalIgnoreCase))
        {
            return "❌ اتصال به سرور تی‌تک برقرار نشد. سیستم نمی‌تواند آدرس newapi.ttac.ir را پیدا کند؛ اینترنت، DNS یا فیلترشکن/پروکسی را بررسی کنید.";
        }

        if (raw.Contains("timed out", StringComparison.OrdinalIgnoreCase) ||
            raw.Contains("timeout", StringComparison.OrdinalIgnoreCase))
        {
            return "❌ تی‌تک در زمان مناسب پاسخ نداد. چند دقیقه بعد دوباره استعلام بگیرید.";
        }

        if (raw.Contains("SSL", StringComparison.OrdinalIgnoreCase) ||
            raw.Contains("certificate", StringComparison.OrdinalIgnoreCase))
        {
            return "❌ خطای ارتباط امن با تی‌تک رخ داد. تاریخ و ساعت ویندوز، اینترنت یا فیلترشکن را بررسی کنید.";
        }

        if (raw.Contains("actively refused", StringComparison.OrdinalIgnoreCase) ||
            raw.Contains("Unable to connect", StringComparison.OrdinalIgnoreCase))
        {
            return "❌ ارتباط با سرور تی‌تک برقرار نشد. ممکن است سایت تی‌تک یا اینترنت شما موقتاً در دسترس نباشد.";
        }

        return $"❌ خطا در اتصال به تی‌تک: {ex.Message}";
    }

    private DrugInfo ParseDrugResponse(string json, string originalBarcode)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (!root.TryGetProperty("data", out var dataElement))
            {
                return new DrugInfo { Success = false, Message = "❌ داده یافت نشد", OriginalBarcode = originalBarcode };
            }

            var data = dataElement;

            if (data.TryGetProperty("statusCode", out var statusCode) && statusCode.GetInt32() != 0)
            {
                string? statusMsg = ExtractString(data, "statusMessage");
                LogMessage($"⚠️ سرور خطای وضعیت داد: {statusMsg}");
                return new DrugInfo { Success = false, Message = "❌ دارو یافت نشد", OriginalBarcode = originalBarcode };
            }

            string? persianName = ExtractString(data, "persianName");
            string? englishName = ExtractString(data, "englishName");
            string? gtin = ExtractString(data, "gtin");
            string? irc = ExtractString(data, "irc");
            string? distributer = ExtractString(data, "distributer");
            var extraFields = ExtractSimpleFields(data);

            LogMessage("📦 اطلاعات استخراج شده:");
            LogMessage($"  • نام فارسی: {persianName}");
            LogMessage($"  • نام انگلیسی: {englishName}");
            LogMessage($"  • GTIN: {gtin}");
            foreach (var field in extraFields)
            {
                LogMessage($"  • {field.Key}: {field.Value}");
            }

            if (string.IsNullOrWhiteSpace(persianName) && string.IsNullOrWhiteSpace(englishName))
            {
                return new DrugInfo { Success = false, Message = "❌ نام دارو مشخص نیست", OriginalBarcode = originalBarcode };
            }

            return new DrugInfo
            {
                Success = true,
                PersianName = persianName ?? "نامشخص",
                EnglishName = englishName ?? "Unknown",
                GTIN = gtin,
                IRC = irc,
                Distributer = distributer,
                ExtraFields = extraFields,
                Message = $"✅ {persianName ?? englishName}",
                OriginalBarcode = originalBarcode
            };
        }
        catch (Exception ex)
        {
            LogMessage($"❌ خطا در تحلیل پاسخ JSON: {ex.Message}");
            return new DrugInfo { Success = false, Message = $"❌ خطا: {ex.Message}", OriginalBarcode = originalBarcode };
        }
    }

    private Dictionary<string, string> ExtractSimpleFields(JsonElement element)
    {
        var fields = new Dictionary<string, string>();

        try
        {
            foreach (var property in element.EnumerateObject())
            {
                string? value = JsonElementToDisplayString(property.Value);
                if (string.IsNullOrWhiteSpace(value))
                    continue;

                fields[GetFriendlyFieldName(property.Name)] = value;
            }
        }
        catch { }

        return fields;
    }

    private string? JsonElementToDisplayString(JsonElement element)
    {
        try
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number => element.ToString(),
                JsonValueKind.True => "True",
                JsonValueKind.False => "False",
                JsonValueKind.Array => string.Join("، ", element.EnumerateArray()
                    .Select(JsonElementToDisplayString)
                    .Where(v => !string.IsNullOrWhiteSpace(v))),
                JsonValueKind.Object => element.ToString(),
                _ => null
            };
        }
        catch
        {
            return null;
        }
    }

    private string GetFriendlyFieldName(string propertyName)
    {
        return propertyName switch
        {
            "persianName" => "نام فارسی",
            "englishName" => "نام انگلیسی",
            "gtin" => "GTIN",
            "irc" => "IRC",
            "uid" => "UID",
            "uidCode" => "UID",
            "distributer" => "توزیع‌کننده",
            "distributor" => "توزیع‌کننده",
            "producer" => "تولیدکننده",
            "manufacturer" => "تولیدکننده",
            "brand" => "برند",
            "genericName" => "نام ژنریک",
            "licenseOwner" => "صاحب پروانه",
            "statusMessage" => "پیام وضعیت",
            "statusCode" => "کد وضعیت",
            "Expiration" => "تاریخ انقضا",
            "expiration" => "تاریخ انقضا",
            "ExpireDate" => "تاریخ انقضا",
            "expireDate" => "تاریخ انقضا",
            "ExpiryDate" => "تاریخ انقضا",
            "expiryDate" => "تاریخ انقضا",
            "ExpDate" => "تاریخ انقضا",
            "expDate" => "تاریخ انقضا",
            "Manufacturing" => "تاریخ ساخت",
            "manufacturing" => "تاریخ ساخت",
            "ProductionDate" => "تاریخ تولید",
            "productionDate" => "تاریخ تولید",
            "ManufactureDate" => "تاریخ تولید",
            "manufactureDate" => "تاریخ تولید",
            "MfgDate" => "تاریخ تولید",
            "mfgDate" => "تاریخ تولید",
            "BatchCode" => "سری ساخت",
            "batchCode" => "سری ساخت",
            "batchNumber" => "شماره بچ",
            "lotNumber" => "شماره لات",
            _ => propertyName
        };
    }

    private string? ExtractFirstString(JsonElement element, params string[] propertyNames)
    {
        foreach (var propertyName in propertyNames)
        {
            string? value = ExtractString(element, propertyName);
            if (!string.IsNullOrWhiteSpace(value))
                return value;
        }

        return null;
    }

    private string? ExtractString(JsonElement element, string propertyName)
    {
        try
        {
            if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind == JsonValueKind.String)
            {
                return prop.GetString();
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static void LogMessage(string message)
    {
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        string logLine = $"[{timestamp}] {message}";
        Console.WriteLine(logLine);

        try
        {
            string logPath = System.IO.Path.Combine(AppContext.BaseDirectory, "drug_lookup.log");
            System.IO.File.AppendAllText(logPath, logLine + Environment.NewLine);
        }
        catch { }
    }

    public static void OpenLogFile()
    {
        try
        {
            string logPath = System.IO.Path.Combine(AppContext.BaseDirectory, "drug_lookup.log");
            if (System.IO.File.Exists(logPath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "notepad.exe",
                    Arguments = logPath,
                    UseShellExecute = true
                });
            }
        }
        catch { }
    }
}

public class DrugInfo
{
    public bool Success { get; set; }
    public string Message { get; set; } = "";
    public string? PersianName { get; set; }
    public string? EnglishName { get; set; }
    public string? GTIN { get; set; }
    public string? IRC { get; set; }
    public string? Distributer { get; set; }
    public Dictionary<string, string> ExtraFields { get; set; } = new();
    public string OriginalBarcode { get; set; } = "";

    public override string ToString() => Message;
}