using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace ScanBridgeTest;

// =====================================================================================
// ویژگی «ورود اطلاعات از راه دور» برای فرم ثبت شیرخشک (TtTeckRegistrationOverlay).
//
// وقتی این فرم برای یک قلم شیرخشک باز می‌شود و حداقل یک گوشی وصل است، همان مراحل فرم - عکس/نام
// محصول، کد ملی، تاریخ تولد، شماره نظام‌پزشکی (فقط نسخه‌محور)، شماره تماس، و در آخر کپچا - یکی‌یکی
// روی گوشی هم نشان داده می‌شود. کاربر می‌تواند از روی گوشی تایپ کند؛ مقدار وارد‌شده مستقیم در همان
// TextBox دسکتاپ گذاشته می‌شود - دقیقاً مثل این‌که پشت سیستم تایپ شده باشد - پس تمام منطق
// اعتبارسنجی/ارسالِ موجود در MainWindow.xaml.cs بدون هیچ تغییری برای هر دو مسیر (محلی/از راه دور)
// استفاده می‌شود. اگر هم‌زمان کسی هم پشت سیستم تایپ کند، هر دو باز می‌مانند - این عمدی است (تصمیم
// صریح کاربر): هیچ قفلی روی فیلدها گذاشته نمی‌شود.
//
// این جریان هیچ‌وقت روی گوشی ذخیره نمی‌شود؛ فقط یک‌بار مصرف است - با هر اسکن جدید از نو شروع
// می‌شود و با رسیدن نتیجه (SCANBRIDGE_ALERT، چه موفق چه خطا) یا بسته‌شدن دستی فرم روی دسکتاپ، از
// روی گوشی پاک می‌شود.
//
// پروتکل پیام (نگاه کنید به Services/ScanBridgeService.cs):
//   دسکتاپ → گوشی:  {"type":"REMOTE_ENTRY_STEP","barcode","stepId","label","hint","inputType",
//                     "photoBase64","captchaImageBase64"}
//                    {"type":"REMOTE_ENTRY_CANCEL","barcode"}
//   گوشی → دسکتاپ:   {"type":"REMOTE_ENTRY_VALUE","barcode","stepId","value"}
//                    {"type":"REMOTE_ENTRY_SUBMIT","barcode"}
//                    {"type":"REMOTE_ENTRY_BACK","barcode"}
// =====================================================================================

public partial class MainWindow
{
    private static class RemoteEntryStepIds
    {
        public const string Info = "info";
        public const string NationalId = "nationalId";
        public const string BirthDate = "birthDate";
        public const string MedicalCouncil = "medicalCouncil";
        public const string Mobile = "mobile";
        public const string Captcha = "captcha";
        public const string Submit = "submit";
    }

    // یک اسنپ‌شات کامل از یک مرحله (همان چیزی که برای گوشی فرستاده شده) - برای این‌که دکمه‌ی
    // «قبلی» روی گوشی بتواند دقیقاً همان مرحله‌ی قبلی را (با همان عکس/کپچا) دوباره نشان دهد.
    private sealed class RemoteEntryStepSnapshot
    {
        public string StepId = string.Empty;
        public string Label = string.Empty;
        public string? Hint;
        public string? PhotoBase64;
        public string? CaptchaImageBase64;
        public string InputType = "text";
    }

    // بارکد قلمی که همین الان ورود از راه دور برایش فعال است؛ null یعنی غیرفعال.
    private string? _remoteEntryBarcode;
    // هویت (DeviceId یا در نبودش DeviceName خام) گوشیِ اسکن‌کننده - وقتی این جریان شروع شد از
    // GetLastScanDeviceKey() گرفته می‌شود (نگاه کنید به MainWindow.xaml.cs) و تا شروع جریانِ بعدی
    // همین می‌ماند (حتی بعد از EndRemoteFormulaEntry - چون پیام موفقیتِ نهایی، چند خط بعد از
    // EndRemoteFormulaEntry فرستاده می‌شود و هنوز باید بداند برای کدام گوشی بود). همه‌ی پیام‌های
    // این ویژگی (مرحله/لغو/نتیجه) با همین کلید فقط برای همین یک گوشی فرستاده می‌شوند - نه برای
    // همه‌ی گوشی‌های وصل - تا وقتی چند گوشی هم‌زمان وصل‌اند، اسکنِ یکی روی فرمِ بقیه ظاهر نشود.
    private string _remoteEntryDeviceKey = string.Empty;
    private string? _remoteEntryCurrentStepId;
    // مرحله‌ای که همین الان روی گوشی نشان داده شده - برای بازسازی دقیق مرحله‌ی قبلی وقتی کاربر
    // دکمه‌ی «قبلی» را می‌زند.
    private RemoteEntryStepSnapshot? _remoteEntryCurrentStepSnapshot;
    // پشته‌ی مراحلِ قبلی (به ترتیبی که کاربر روی گوشی جلو رفته) - برای دکمه‌ی «قبلی».
    private readonly List<RemoteEntryStepSnapshot> _remoteEntryStepHistory = new();
    // true یعنی نوبت مرحله‌ی کپچا رسیده ولی کپچا هنوز روی دسکتاپ لود نشده - به محض لود شدن
    // (NotifyRemoteEntryCaptchaLoaded) خودش فرستاده می‌شود.
    private bool _remoteEntryWaitingForCaptcha;
    // آخرین تصویر کپچای بارگذاری‌شده به‌صورت Base64 (همان چیزی که LoadTtacCaptchaAsync از سرور
    // گرفته) - برای این‌که وقتی نوبت مرحله‌ی کپچا می‌رسد و کپچا از قبل لود شده، دوباره درخواست
    // گرفتن کپچای تازه نزنیم.
    private string? _lastLoadedCaptchaImageBase64;

    // «ثبت مجدد از روی گوشی»: کاربر روی دیالوگ موفقیتِ ثبت، دکمه‌ی «ثبت مجدد» را زده - یعنی
    // می‌خواهد قلم بعدی را با همان اطلاعات بیمار (_lastFormulaRepeatContext در MainWindow.xaml.cs)
    // فقط با اسکن بارکد + کپچای تازه ثبت کند. یک پرچم یک‌بارمصرف است؛ اولین اسکن بعدی
    // (TryAutoOpenInfantFormulaRegistration در MainWindow.xaml.cs) آن را مصرف می‌کند - چه بارکد
    // معتبر باشد چه قبلاً ثبت‌شده باشد.
    private bool _remoteEntryRepeatArmed;

    private bool IsRemoteFormulaEntryActive =>
        !string.IsNullOrEmpty(_remoteEntryBarcode)
        && _pendingRegistrationTtTeckRow != null
        && string.Equals(_remoteEntryBarcode, _pendingRegistrationTtTeckRow.Barcode, StringComparison.OrdinalIgnoreCase);

    private bool IsRemoteFormulaEntryActiveFor(string barcode) =>
        IsRemoteFormulaEntryActive && string.Equals(_remoteEntryBarcode, barcode, StringComparison.OrdinalIgnoreCase);

    // فراخوانی از OpenTtTeckRegistrationForRow، بلافاصله بعد از آماده‌شدن فرم برای یک قلم. اگر قلم
    // شیرخشک نباشد یا گوشی‌ای وصل نباشد، بی‌اثر است (هیچ پیامی فرستاده نمی‌شود).
    private void StartRemoteFormulaEntryIfPossible(TtTeckHistoryRow row)
    {
        _remoteEntryBarcode = null;
        _remoteEntryCurrentStepId = null;
        _remoteEntryCurrentStepSnapshot = null;
        _remoteEntryStepHistory.Clear();
        _remoteEntryWaitingForCaptcha = false;
        _lastLoadedCaptchaImageBase64 = null;

        if (_service == null || _service.ConnectedClients <= 0)
            return;
        if (GetFormulaRegistrationModeForRow(row) == FormulaRegistrationMode.Unknown)
            return;

        _remoteEntryBarcode = row.Barcode;
        // همین‌جا هویت گوشیِ اسکن‌کننده برای کل این جریان «قفل» می‌شود - نگاه کنید به توضیح بالای
        // فیلد _remoteEntryDeviceKey.
        _remoteEntryDeviceKey = GetLastScanDeviceKey();

        string productName = string.IsNullOrWhiteSpace(row.ProductDisplayName) ? row.Barcode : row.ProductDisplayName;
        string? photoBase64 = ReadFileAsBase64OrNull(GetFormulaPhotoPathForBarcode(row.Barcode));

        PushRemoteEntryStep(RemoteEntryStepIds.Info, productName, "برای شروع، دکمه‌ی بعدی را بزنید.", photoBase64: photoBase64, inputType: "info");
    }

    // مرحله‌ی فعلی نمایش‌داده‌شده روی گوشی را عوض می‌کند؛ اگر ورود از راه دور فعال نباشد، بی‌اثر است.
    // isBack=true وقتی است که این تغییر مرحله به‌خاطر زدن دکمه‌ی «قبلی» روی گوشی است - آن‌وقت خودِ
    // مرحله‌ای که داریم ترکش می‌کنیم نباید دوباره به پشته‌ی history اضافه شود (وگرنه «قبلی» زدن
    // چیزی جز رفت‌وبرگشت بین دو مرحله‌ی آخر ممکن نمی‌شد).
    private void PushRemoteEntryStep(string stepId, string label, string? hint, string? photoBase64 = null, string? captchaImageBase64 = null, string inputType = "text", bool isBack = false)
    {
        if (!IsRemoteFormulaEntryActive || _service == null)
            return;

        if (!isBack && _remoteEntryCurrentStepSnapshot != null)
            _remoteEntryStepHistory.Add(_remoteEntryCurrentStepSnapshot);

        // مقدار فعلیِ همان TextBox روی دسکتاپ برای این مرحله (اگر از قبل چیزی در آن هست) - تا
        // وقتی این مرحله دوباره روی گوشی نشان داده می‌شود (چه با «قبلی»، چه بعد از عبور دوباره از
        // یک مرحله‌ی میانی برای رسیدن به مرحله‌ای که باید اصلاح شود)، کادر گوشی خالی نباشد و
        // کاربر مجبور نشود چیزی را که قبلاً درست وارد کرده دوباره تایپ کند - همان مقدار از قبل در
        // کادر هست، فقط اگر لازم بود اصلاحش می‌کند.
        string? prefillValue = GetRemoteEntryFieldPrefill(stepId);

        _remoteEntryCurrentStepId = stepId;
        _remoteEntryCurrentStepSnapshot = new RemoteEntryStepSnapshot
        {
            StepId = stepId,
            Label = label,
            Hint = hint,
            PhotoBase64 = photoBase64,
            CaptchaImageBase64 = captchaImageBase64,
            InputType = inputType
        };
        _service.BroadcastRemoteEntryStep(_remoteEntryBarcode!, stepId, label, hint ?? string.Empty, photoBase64, captchaImageBase64, inputType, prefillValue, targetDeviceKey: _remoteEntryDeviceKey);
    }

    // مقدار فعلیِ TextBox دسکتاپ متناظر با یک stepId را برمی‌گرداند (برای پرکردن از قبلِ کادر
    // گوشی) - برای مراحلی که کادر متنی ندارند (info/submit) همیشه null است.
    private string? GetRemoteEntryFieldPrefill(string stepId)
    {
        switch (stepId)
        {
            case RemoteEntryStepIds.NationalId:
                return NullIfBlank(TtTeckRegistrationNationalIdTextBox.Text);

            case RemoteEntryStepIds.BirthDate:
                string dayText = ToEnglishDigits(TtTeckBirthDayTextBox.Text.Trim());
                string monthText = ToEnglishDigits(TtTeckBirthMonthTextBox.Text.Trim());
                string yearText = ToEnglishDigits(TtTeckBirthYearTextBox.Text.Trim());
                if (int.TryParse(dayText, out int day) && int.TryParse(monthText, out int month) && int.TryParse(yearText, out int year))
                    return $"{year:0000}/{month:00}/{day:00}";
                return null;

            case RemoteEntryStepIds.MedicalCouncil:
                return NullIfBlank(TtTeckRegistrationMedicalCouncilTextBox.Text);

            case RemoteEntryStepIds.Mobile:
                return NullIfBlank(TtTeckRegistrationMobileTextBox.Text);

            case RemoteEntryStepIds.Captcha:
                return NullIfBlank(TtTeckRegistrationCaptchaTextBox.Text);

            default:
                return null;
        }
    }

    private static string? NullIfBlank(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string? ReadFileAsBase64OrNull(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return null;
        try
        {
            return Convert.ToBase64String(File.ReadAllBytes(path));
        }
        catch
        {
            return null;
        }
    }

    // ورود از راه دور را غیرفعال می‌کند. notifyPhone=true (پیش‌فرض) وقتی درست است که فرم بدون هیچ
    // نتیجه‌ی نهایی‌ای بسته می‌شود (کاربر خودش بستش) - آن‌وقت باید صریحاً به گوشی گفت ویزارد را پاک
    // کند. notifyPhone=false در مسیرهایی استفاده می‌شود که خودِ BroadcastAlert (موفق یا ناموفق)
    // چند لحظه بعد نتیجه را روی گوشی نشان می‌دهد - آن‌جا فرستادن یک پیام لغوِ جداگانه زائد است.
    private void EndRemoteFormulaEntry(bool notifyPhone = true)
    {
        bool wasActive = IsRemoteFormulaEntryActive;
        string? barcode = _remoteEntryBarcode;

        _remoteEntryBarcode = null;
        _remoteEntryCurrentStepId = null;
        _remoteEntryCurrentStepSnapshot = null;
        _remoteEntryStepHistory.Clear();
        _remoteEntryWaitingForCaptcha = false;

        if (wasActive && notifyPhone && !string.IsNullOrEmpty(barcode))
        {
            try { _service?.BroadcastRemoteEntryCancel(barcode, targetDeviceKey: _remoteEntryDeviceKey); } catch { }
        }
    }

    // فراخوانی از LoadTtacCaptchaAsync بعد از این‌که یک کپچای جدید با موفقیت لود شد (چه در همان
    // لحظه‌ی باز شدن فرم، چه بعداً با «بازنشانی کپچا»).
    private void NotifyRemoteEntryCaptchaLoaded(string captchaImageBase64)
    {
        _lastLoadedCaptchaImageBase64 = captchaImageBase64;

        if (!IsRemoteFormulaEntryActive)
            return;

        // اگر گوشی همین الان منتظر کپچا بود (چه چون تازه به این مرحله رسیده، چه چون کپچا با
        // «بازنشانی» عوض شد درحالی‌که گوشی داشت مرحله‌ی کپچای قبلی را نشان می‌داد)، تصویر تازه را
        // بفرست.
        if (_remoteEntryWaitingForCaptcha || _remoteEntryCurrentStepId == RemoteEntryStepIds.Captcha)
        {
            _remoteEntryWaitingForCaptcha = false;
            PushRemoteEntryStep(RemoteEntryStepIds.Captcha, "کد کپچا", "کد داخل عکس را وارد کنید.", captchaImageBase64: captchaImageBase64, inputType: "captcha");
        }
    }

    // فراخوانی می‌شود وقتی از گوشی یک REMOTE_ENTRY_VALUE می‌رسد (کاربر یک فیلد را پر کرده و
    // «بعدی» را زده). این پیام از یک ترد وب‌سوکت پس‌زمینه می‌رسد، پس همیشه با Dispatcher به ترد UI
    // منتقل می‌شود.
    private void HandleRemoteEntryValueFromPhone(string barcode, string stepId, string value)
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            if (!IsRemoteFormulaEntryActiveFor(barcode))
                return; // پیام مربوط به یک فرم قبلی/بسته‌شده است - نادیده گرفته می‌شود.

            switch (stepId)
            {
                case RemoteEntryStepIds.Info:
                    PushRemoteEntryStep(RemoteEntryStepIds.NationalId, "کد ملی", "کد ملی بیمار را وارد کنید.", inputType: "number");
                    break;

                case RemoteEntryStepIds.NationalId:
                    TtTeckRegistrationNationalIdTextBox.Text = ToEnglishDigits(value.Trim());
                    PushRemoteEntryStep(RemoteEntryStepIds.BirthDate, "تاریخ تولد", "روز، ماه و سال تولد (شمسی) را وارد کنید.", inputType: "birthDate");
                    break;

                case RemoteEntryStepIds.BirthDate:
                    ApplyRemoteEntryBirthDate(ToEnglishDigits(value.Trim()));
                    if (TtTeckRegistrationTypeComboBox.SelectedIndex == 1)
                        PushRemoteEntryStep(RemoteEntryStepIds.MedicalCouncil, "شماره نظام پزشکی", "برای ثبت نسخه‌محور الزامی است.", inputType: "number");
                    else
                        PushRemoteEntryStep(RemoteEntryStepIds.Mobile, "شماره تماس", "شماره موبایل بیمار (برای پیامک نتیجه).", inputType: "number");
                    break;

                case RemoteEntryStepIds.MedicalCouncil:
                    TtTeckRegistrationMedicalCouncilTextBox.Text = ToEnglishDigits(value.Trim());
                    PushRemoteEntryStep(RemoteEntryStepIds.Mobile, "شماره تماس", "شماره موبایل بیمار (برای پیامک نتیجه).", inputType: "number");
                    break;

                case RemoteEntryStepIds.Mobile:
                    TtTeckRegistrationMobileTextBox.Text = ToEnglishDigits(value.Trim());
                    AdvanceRemoteEntryToCaptcha();
                    break;

                case RemoteEntryStepIds.Captcha:
                    TtTeckRegistrationCaptchaTextBox.Text = value.Trim();
                    PushRemoteEntryStep(RemoteEntryStepIds.Submit, "ایجاد نسخه و ثبت", "اطلاعات کامل شد. برای ثبت نهایی دکمه را بزنید.", inputType: "button");
                    break;
            }
        }));
    }

    // مقدار تاریخ تولد که از گوشی می‌رسد به فرمت YYYY/MM/DD است (همان چیزی که ویزارد گوشی از سه
    // کادر جدای روز/ماه/سال می‌سازد - نگاه کنید به RemoteFormulaEntryDialog در MainActivity.kt).
    // این مقدار مستقیماً در همان سه TextBox روز/ماه/سال دسکتاپ گذاشته می‌شود - دقیقاً مثل این‌که
    // پشت سیستم تایپ شده باشد - تا هم UI دسکتاپ درست به‌روز شود و هم GetTtacBirthDateText() (که
    // این سه کادر را نسبت به کادر ترکیبی متنی در اولویت می‌گذارد) مستقیماً از همین‌ها بخواند.
    private void ApplyRemoteEntryBirthDate(string value)
    {
        var parts = value.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 3 &&
            int.TryParse(parts[0], out int year) &&
            int.TryParse(parts[1], out int month) &&
            int.TryParse(parts[2], out int day))
        {
            TtTeckBirthYearTextBox.Text = year.ToString();
            TtTeckBirthMonthTextBox.Text = month.ToString();
            TtTeckBirthDayTextBox.Text = day.ToString();
            return;
        }

        // فرمت غیرمنتظره - برای احتیاط لااقل در کادر ترکیبی متنی می‌گذاریم تا NormalizeDateInput
        // به‌عنوان راه دوم امتحانش کند (همان مسیر fallback موجود در GetTtacBirthDateText).
        TtTeckRegistrationBirthDateTextBox.Text = NormalizeDateInput(value);
    }

    // فراخوانی می‌شود وقتی ایجاد نسخه یا ثبت قلمِ از راه دور با خطا مواجه شود (مثلاً «کد ملی
    // نامعتبر است» یا «احتمال ثبت تکراری»). به‌جای پایان‌دادن به کل جریان (که یعنی ویزارد از روی
    // گوشی پاک می‌شد و برای اصلاح باید از اول - از عکس محصول - دوباره شروع می‌شد)، همان مرحله‌ی
    // «ثبت نهایی» را با متنِ خطا دوباره می‌فرستیم؛ کاربر می‌تواند از همین‌جا با دکمه‌ی «قبلی»
    // (که چون isBack:true است، مرحله‌ی فعلی را دوباره به پشته اضافه نمی‌کند) به هر مرحله‌ای که
    // لازم است برگردد، فیلد را اصلاح کند، دوباره جلو بیاید و همین دکمه را بزند.
    private void ShowRemoteEntryErrorAndAllowRetry(string title, string message)
    {
        if (!IsRemoteFormulaEntryActive)
            return;

        string hint = string.IsNullOrWhiteSpace(message)
            ? "برای اصلاح، با «قبلی» به مرحله‌ی موردنظر برگردید."
            : $"{message} برای اصلاح، با «قبلی» به مرحله‌ی موردنظر برگردید و دوباره امتحان کنید.";

        PushRemoteEntryStep(RemoteEntryStepIds.Submit, title, hint, inputType: "button", isBack: true);
    }

    private void AdvanceRemoteEntryToCaptcha()
    {
        if (!string.IsNullOrWhiteSpace(_ttacCurrentCaptchaId) && !string.IsNullOrWhiteSpace(_lastLoadedCaptchaImageBase64))
        {
            _remoteEntryWaitingForCaptcha = false;
            PushRemoteEntryStep(RemoteEntryStepIds.Captcha, "کد کپچا", "کد داخل عکس را وارد کنید.", captchaImageBase64: _lastLoadedCaptchaImageBase64, inputType: "captcha");
        }
        else
        {
            // کپچا هنوز لود نشده (نادر - معمولاً تا این مرحله از گوشی، کپچا از قبل روی دسکتاپ آماده
            // است)؛ NotifyRemoteEntryCaptchaLoaded به‌محض لود شدن خودش این مرحله را می‌فرستد.
            _remoteEntryWaitingForCaptcha = true;
        }
    }

    // فراخوانی می‌شود وقتی از گوشی یک REMOTE_ENTRY_SUBMIT می‌رسد (دکمه‌ی نهایی «ایجاد نسخه و ثبت»
    // روی گوشی). همان دو مرحله‌ی دکمه‌های محلی دسکتاپ - «ایجاد نسخه» و بعد «ثبت قلم» - پشت‌سرهم و
    // با همان منطق اعتبارسنجی/خطا اجرا می‌شوند؛ نتیجه (موفق یا ناموفق) از مسیر موجود BroadcastAlert
    // به گوشی می‌رسد.
    private void HandleRemoteEntrySubmitFromPhone(string barcode)
    {
        Dispatcher.BeginInvoke(new Action(async () =>
        {
            if (!IsRemoteFormulaEntryActiveFor(barcode))
                return;

            await CreateCurrentTtacPrescriptionAsync();

            if (_ttacCurrentPrescriptionId.HasValue && IsRemoteFormulaEntryActiveFor(barcode))
            {
                await SubmitCurrentTtacItemAsync(viaRemoteEntry: true);
            }
            // اگر ایجاد نسخه ناموفق بود یا فرم در این فاصله عوض/بسته شد، همین‌جا متوقف می‌شود -
            // خطای احتمالی از قبل با onFailureMessageShown → BroadcastAlert به گوشی رسیده است.
        }));
    }

    // فراخوانی می‌شود وقتی از گوشی یک REMOTE_ENTRY_BACK می‌رسد (کاربر دکمه‌ی «قبلی» را زده - مثلاً
    // چون یک کادر را اشتباه زده و می‌خواهد برگردد اصلاحش کند). مرحله‌ی قبلی از پشته‌ی
    // _remoteEntryStepHistory با همان عکس/کپچای وقتی که اول نشان داده شده بود دوباره فرستاده
    // می‌شود؛ چیزی از فیلدهای دسکتاپ پاک نمی‌شود - فقط خودِ صفحه‌ی گوشی یک قدم برمی‌گردد.
    private void HandleRemoteEntryBackFromPhone(string barcode)
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            if (!IsRemoteFormulaEntryActiveFor(barcode))
                return;

            if (_remoteEntryStepHistory.Count == 0)
                return; // همین الان روی اولین مرحله است - «قبلی» بی‌اثر است.

            var previous = _remoteEntryStepHistory[^1];
            _remoteEntryStepHistory.RemoveAt(_remoteEntryStepHistory.Count - 1);

            PushRemoteEntryStep(previous.StepId, previous.Label, previous.Hint,
                previous.PhotoBase64, previous.CaptchaImageBase64, previous.InputType, isBack: true);
        }));
    }

    // فراخوانی می‌شود وقتی از گوشی یک REMOTE_ENTRY_REPEAT_ARM می‌رسد (کاربر روی دیالوگ موفقیتِ ثبت
    // شیرخشک، دکمه‌ی «ثبت مجدد» را زده). این پیام از یک ترد وب‌سوکت پس‌زمینه می‌رسد. اگر از ثبت
    // قبلی هیچ context ای موجود نباشد (مثلاً برنامه همین الان استارت شده و هنوز چیزی ثبت نشده)،
    // چیزی مسلح نمی‌شود - اسکن بعدی مثل همیشه (فرم خالی) باز می‌شود.
    private void HandleRemoteEntryRepeatArmFromPhone()
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            _remoteEntryRepeatArmed = _lastFormulaRepeatContext != null;
        }));
    }
}
