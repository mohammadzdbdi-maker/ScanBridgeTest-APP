using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ScanBridgeTest;

public enum AppLanguage
{
    Persian,
    English
}

public sealed class LocalizationManager
{
    private static LocalizationManager? _instance;

    private AppLanguage _currentLanguage = AppLanguage.Persian;

    public event EventHandler? LanguageChanged;

    public static LocalizationManager Instance
    {
        get
        {
            _instance ??= new LocalizationManager();
            return _instance;
        }
    }

    public AppLanguage CurrentLanguage
    {
        get => _currentLanguage;
        set
        {
            if (_currentLanguage == value)
            {
                return;
            }

            _currentLanguage = value;
            LanguageChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private LocalizationManager()
    {
        LoadSettings();
    }

    private static string SettingsPath =>
        Path.Combine(AppContext.BaseDirectory, "settings.json");

    private void LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsPath))
            {
                _currentLanguage = AppLanguage.Persian;
                return;
            }

            string json = File.ReadAllText(SettingsPath);
            var settings = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

            if (settings is null)
            {
                _currentLanguage = AppLanguage.Persian;
                return;
            }

            if (settings.TryGetValue("Language", out string? languageText))
            {
                if (Enum.TryParse<AppLanguage>(languageText, out var language))
                {
                    _currentLanguage = language;
                    return;
                }
            }

            _currentLanguage = AppLanguage.Persian;
        }
        catch
        {
            _currentLanguage = AppLanguage.Persian;
        }
    }

    public void SaveSettings(AppLanguage language)
    {
        try
        {
            var settings = new Dictionary<string, string>
            {
                { "Language", language.ToString() }
            };

            string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            File.WriteAllText(SettingsPath, json);
        }
        catch
        {
            // اگر ذخیره تنظیمات خطا داد، برنامه خراب نشود
        }
    }

    private static readonly Dictionary<string, (string English, string Persian)> _translations = new()
    {
    ["AppTitle"] = ("Barcode Bridge", "بارکد برج"),
    ["Messages"] = ("Messages", "پیام‌ها"),
    ["BarcodeTitle"] = ("Scan the QR code below on your phone to connect", "جهت اتصال بارکد زیر را روی گوشی اسکن کنید"),
    ["ConnectedDevices"] = ("Connected Devices", "دستگاه های متصل"),
    ["NoDevices"] = ("No devices connected", "هیچ دستگاهی متصل نیست"),
    ["SystemInfo"] = ("System Information", "اطلاعات سیستم"),
    ["AutoStartup"] = ("Auto-start on Windows boot", "اجرای خودکار هنگام روشن شدن ویندوز"),
    ["UserPanel"] = ("User Panel", "پنل کاربری"),
    ["History"] = ("History", "تاریخچه"),
    ["Print"] = ("Print", "چاپ"),
    ["Support"] = ("Support", "پشتیبانی"),
    ["SupportMessage"] = ("Contact us via WhatsApp:", "برای پشتیبانی با شماره تماس زیر در واتساپ در تماس باشید:"),
    ["RemoteSupport"] = ("Remote Support Connection", "اتصال پشتیبانی از راه دور"),
    ["NoMessages"] = ("No messages", "پیامی وجود ندارد"),
    ["OpenLink"] = ("Open Link", "مشاهده لینک"),
    ["TtTeckPanel"] = ("TtTeck", "تی‌تک"),
    ["CargoDelivery"] = ("Cargo delivery", "تحویل بار"),
    ["ReceiveStatus"] = ("Receive status", "تعیین وضعیت"),
    ["SupportMessageText"] = ("Contact us via WhatsApp, or open the website to buy/renew:", "برای پشتیبانی با شماره زیر در واتساپ در تماس باشید یا برای خرید/تمدید وارد سایت شوید:"),
    ["Website"] = ("Website", "سایت خرید و تمدید"),
    ["SupportOverlayTitle"] = ("Scanbridge Support", "پشتیبانی Scanbridge"),
    ["WhatsAppContact"] = ("WhatsApp: 09136346309", "واتساپ: 09136346309"),
    ["Close"] = ("Close", "بستن"),
    ["LanguageSelectTitle"] = ("🌐 Select Language", "🌐 انتخاب زبان"),
    ["SelectLanguage"] = ("Select Language", "انتخاب زبان"),
    ["PersianOption"] = ("Persian", "فارسی"),
    ["EnglishOption"] = ("English", "انگلیسی"),
    ["HistoryTitle"] = ("Scan History", "تاریخچه اسکن‌ها"),
    ["ClearHistory"] = ("🗑️ Clear all", "🗑️ پاک کردن همه"),
    ["Settings"] = ("Settings", "تنظیمات"),
    ["TtTeckEnableText"] = ("Enable auto search on TtTeck", "فعال‌سازی جستجوی خودکار تی‌تک"),
    ["SelectExportType"] = ("Select Export Type", "انتخاب نوع خروجی"),
    ["ExportTypeDescription"] = ("Which barcodes do you want to export?", "کدام بارکدها را می‌خواهید صادر کنید؟"),
    ["ExportAllTitle"] = ("All Barcodes", "تمام بارکدها"),
    ["ExportAllDesc"] = ("Export all scanned barcodes", "تمام بارکدهای اسکن‌شده را صادر کن"),
    ["ExportTtTeckTitle"] = ("TtTeck Only", "فقط بارکدهای تی‌تک"),
    ["ExportTtTeckDesc"] = ("Only barcodes searched from TtTeck", "فقط بارکدهایی که از تی‌تک جستجو شدند"),
    ["LanguageCardTitle"] = ("🌐 Language", "🌐 زبان"),
    ["ChangeLanguage"] = ("Change Language", "تغییر زبان"),
    ["ConfirmDeleteTitle"] = ("Clear History", "حذف تاریخچه"),
    ["ConfirmDeleteMessage"] = ("Are you sure you want to clear the entire history?", "آیا مطمئن هستید که می‌خواهید کل تاریخچه پاک شود؟"),
    ["ConfirmDeleteWarning"] = ("This action cannot be undone.", "این عملیات قابل بازگشت نیست."),
    ["ConfirmDeleteYes"] = ("Yes, clear it", "بله، حذف کن"),
    ["ConfirmDeleteNo"] = ("Cancel", "منصرف شو"),
    ["ExportSuccessTitle"] = ("Export successful!", "صادر کردن موفق!"),
    ["ExportSuccessMessage"] = ("History was exported to Excel successfully", "تاریخچه با موفقیت به Excel صادر شد"),
    ["OkConfirm"] = ("OK", "تایید"),
    ["CancelButton"] = ("Cancel", "انصراف"),
    ["Save"] = ("Save", "ذخیره"),
    ["SearchHistory"] = ("Search history", "جستجو در تاریخچه"),
    ["SearchHistoryTooltip"] = ("Search by barcode, product name, device name, date or time", "جستجو بر اساس بارکد، نام محصول، نام دستگاه، تاریخ یا ساعت"),
    ["DateAndTimeFilter"] = ("Date and time filter", "فیلتر تاریخ و ساعت"),
    ["SelectDateRange"] = ("Select date and time range", "انتخاب بازه تاریخ و ساعت"),
    ["FromDateTime"] = ("From date and time", "از تاریخ و ساعت"),
    ["ToDateTime"] = ("To date and time", "تا تاریخ و ساعت"),
    ["Year"] = ("Year", "سال"),
    ["Month"] = ("Month", "ماه"),
    ["Day"] = ("Day", "روز"),
    ["Hour"] = ("Hour", "ساعت"),
    ["Minute"] = ("Minute", "دقیقه"),
    ["ApplyFilter"] = ("Apply filter", "اعمال فیلتر"),
    ["ClearFilter"] = ("Clear filter", "حذف فیلتر"),
    ["RetryTtTeckTitle"] = ("Retry TtTeck lookup", "استعلام مجدد تی‌تک"),
    ["RetryTtTeckReason"] = ("Reason the previous lookup failed:", "دلیل ناموفق بودن استعلام قبلی:"),
    ["RetryTtTeckConfirm"] = ("Retry lookup", "استعلام مجدد"),
    ["ProductDetailsTitle"] = ("Product details", "جزئیات محصول"),
    ["PersianNameLabel"] = ("Persian name:", "نام فارسی:"),
    ["EnglishNameLabel"] = ("English name:", "نام انگلیسی:"),
    ["BarcodeLabel"] = ("Barcode:", "بارکد:"),
    ["DateTimeLabel"] = ("Date and time:", "تاریخ و ساعت:"),
    ["DeviceLabel"] = ("Device:", "دستگاه:"),
    ["StatusLabel"] = ("Status:", "وضعیت:"),
    ["CopyBarcode"] = ("Copy barcode", "کپی بارکد"),
    ["AllTtTeckInfo"] = ("All TtTeck information", "همه اطلاعات تی‌تک"),
    ["EditDeviceName"] = ("Edit device name", "ویرایش نام دستگاه"),
    ["CurrentDeviceName"] = ("Current device name:", "نام فعلی دستگاه:"),
    ["CustomName"] = ("Custom name", "نام دلخواه"),
    ["RemoveCustomName"] = ("Remove custom name", "حذف نام دلخواه"),
    ["DeviceReconnectTitle"] = ("🔌 Device reconnect", "🔌 اتصال مجدد دستگاه‌ها"),
    ["DeviceReconnectDesc"] = ("If you disconnected a phone and want it to connect again, click this option.", "اگر گوشی را قطع کرده‌اید و می‌خواهید دوباره وصل شود، این گزینه را بزنید."),
    ["AllowPhonesReconnect"] = ("Allow phones to reconnect", "فعال‌سازی اتصال مجدد گوشی‌ها"),
    ["SystemInfoComputer"] = ("Computer: {0}", "نام سیستم: {0}"),
    ["SystemInfoIp"] = ("IP: {0}", "آی‌پی: {0}"),
    ["SystemInfoPort"] = ("Port: {0}", "پورت: {0}"),
    ["DateRangeFromTo"] = ("From {0} to {1}", "از {0} تا {1}"),
    ["DeviceDisconnectedMessage"] = ("{0} was disconnected.", "اتصال {0} قطع شد."),
    ["RialSuffix"] = ("", " ریال"),
    ["TokenRemaining"] = ("token: {0}", "اعتبار توکن: {0}"),
    ["ConnectedWithName"] = ("Connected - {0}", "متصل است - {0}"),
    ["MonthsCount"] = ("{0} months", "{0} ماه"),
    ["DaysLeft"] = ("{0} days left", "{0} روز تا انقضا"),
    ["LotExpiry"] = ("Lot: {0} | Expiry: {1}", "سری: {0} | انقضا: {1}"),
    ["AddingCount"] = ("Adding {0} of {1}...", "در حال افزودن {0} از {1}..."),
    ["SuccessFailed"] = ("Successful: {0}\nFailed: {1}", "موفق: {0}\nناموفق: {1}"),
    ["ResultStatus"] = ("Result: {0}", "نتیجه: {0}"),
    ["Successful"] = ("Successful", "موفق"),
    ["Failed"] = ("Failed", "ناموفق"),
    ["LoginDeleted"] = ("{0} was deleted.", "«{0}» حذف شد."),
    ["EditingLogin"] = ("Editing: {0} — re-enter the password and press Save.", "در حال ویرایش: {0} — رمز را دوباره وارد کنید و ذخیره بزنید."),
    ["CodeUsedFor"] = ("This code was used for: {0}", "این کد قبلاً برای داروخانه «{0}» استفاده شده"),
    ["SavedCodesMatching"] = ("Saved codes matching: {0}", "کدهای ذخیره‌شده مشابه: {0}"),
    ["LoginTo"] = ("Login to {0}", "ورود به {1}"),
    ["OpenTtacLoginPage"] = ("Open the TTAC login page with the saved code {0}", "باز کردن صفحه‌ی ورود تی‌تک با کد ذخیره‌شده {0}"),
    ["StatusCodeFormat"] = ("{0}\nStatus code: {1}", "{0}\nکد وضعیت: {1}"),
    ["TtacQuotaExhaustedFriendly"] = ("The quota of this product in this pharmacy is exhausted; all of its serials have already been received (status code 5173). Choose another product to register.", "سهمیه‌ی این محصول در این داروخانه تمام شده است؛ همه‌ی سری‌های آن قبلاً دریافت شده‌اند (کد وضعیت 5173). برای ثبت، محصول دیگری را انتخاب کنید."),
    ["BirthDateInfo"] = ("Entered birth date: {0} | API birth date: {1}", "تاریخ تولد واردشده: {0} | تاریخ ارسالی به API: {1}"),
    ["PrescriptionCreated"] = ("Prescription created. Prescription ID: {0}", "نسخه ایجاد شد. شناسه نسخه: {0}"),
    ["PrescriptionIdFormat"] = ("Prescription ID: {0}", "شناسه نسخه: {0}"),
    ["TtacPageLoadError"] = ("TTAC page did not load successfully: {0}. Check internet/VPN/DNS or try again later.", "صفحه تی‌تک درست بارگذاری نشد: {0}. اینترنت/VPN/DNS را بررسی کنید یا کمی بعد دوباره تلاش کنید."),
    ["ConnectedCount"] = ("Connected ({0} devices)", "متصل ({0} دستگاه)"),
    ["ActivationFailedFormat"] = ("Activation failed: {0}", "فعال‌سازی ناموفق بود: {0}"),
    ["DetailTextFormat"] = ("IRC: {0} | Lot: {1} | Qty: {2} | Sender: {3} | Sent: {4}", "IRC: {0} | سری: {1} | تعداد: {2} | ارسال‌کننده: {3} | تاریخ ارسال: {4}"),
    ["NearExpiryHeader"] = ("⏰ {0} item(s) are near their expiration date", "⏰ {0} قلم دارو در آستانه‌ی تاریخ انقضا هستند"),
    ["LotExpiryWide"] = ("Lot: {0}    |    Expiry: {1}", "سری ساخت: {0}    |    تاریخ انقضا: {1}"),
    ["Qty"] = ("Qty", "تعداد"),
    ["OpenLink2"] = ("Open link", "مشاهده لینک"),
    ["Delete"] = ("Delete", "حذف"),
    ["Register"] = ("Register", "ثبت در تی‌تک"),
    ["Copy"] = ("Copy", "کپی"),
    ["Copied"] = ("Copied", "کپی شد"),
    ["NotATtTeckBarcode"] = ("Not a TtTeck barcode", "بارکد تی‌تک نیست"),
    ["TtTeckPanel2"] = ("TtTeck Panel", "پنل تی‌تک"),
    ["AllTtTeckItems"] = ("All TtTeck items", "همه تی‌تک‌ها"),
    ["Formula"] = ("🍼 Formula", "🍼 شیر خشک"),
    ["InvalidRange"] = ("Invalid range", "بازه نامعتبر"),
    ["TheStartDateTimeMustBeBeforeTheEndDateTime"] = ("The start date/time must be before the end date/time.", "تاریخ و ساعت شروع باید قبل از تاریخ و ساعت پایان باشد."),
    ["RegistrationTypeChangedGettingANewCaptcha"] = ("Registration type changed. Getting a new captcha...", "نوع ثبت تغییر کرد. در حال دریافت کپچای جدید..."),
    ["ReadyForANewPrescription"] = ("Ready for a new prescription.", "برای نسخه جدید آماده شد."),
    ["Registered"] = ("Registered", "ثبت شد"),
    ["TtTeckSessionExpiredPleaseLoginAgain"] = ("TtTeck session expired. Please login again.", "نشست تی‌تک منقضی شده است. لطفاً دوباره وارد شوید."),
    ["WaitingForTtTeckLookup"] = ("⏳ Waiting for TtTeck lookup", "⏳ در انتظار استعلام تی‌تک"),
    ["LookupSuccessful"] = ("Lookup successful", "استعلام موفق"),
    ["LookupFailed"] = ("Lookup failed", "استعلام ناموفق"),
    ["ThePreviousLookupResultIsNotAvailable"] = ("The previous lookup result is not available.", "نتیجه استعلام قبلی در دسترس نیست."),
    ["ConnectionToTtTeckFailedWindowsCannotResolveNewapiTtacIrCheckInternetDNSVPNOrProxySettings"] = ("Connection to TtTeck failed. Windows cannot resolve newapi.ttac.ir. Check internet, DNS, VPN or proxy settings.", "اتصال به سرور تی‌تک برقرار نشد. ویندوز نمی‌تواند آدرس newapi.ttac.ir را پیدا کند؛ اینترنت، DNS یا فیلترشکن/پروکسی را بررسی کنید."),
    ["TtTeckDidNotRespondInTimePleaseRetryAFewMinutesLater"] = ("TtTeck did not respond in time. Please retry a few minutes later.", "تی‌تک در زمان مناسب پاسخ نداد. چند دقیقه بعد دوباره استعلام بگیرید."),
    ["ShowAll"] = ("Show All", "نمایش همه"),
    ["TtTeckFilter"] = ("🧪 TtTeck Filter", "🧪 فیلتر تی‌تک"),
    ["FormulaRegistration"] = ("🍼 Formula Registration", "🍼 ثبت شیر خشک"),
    ["RegularBarcode"] = ("Regular barcode", "بارکد عادی"),
    ["TtTeckProduct"] = ("TtTeck product", "محصول تی‌تک"),
    ["ScanbridgeHistoryReport"] = ("Scanbridge History Report", "گزارش تاریخچه Scanbridge"),
    ["AllScans"] = ("All scans", "همه اسکن‌ها"),
    ["NoData"] = ("No data", "داده‌ای وجود ندارد"),
    ["ThereIsNoHistoryItemForTheCurrentFilter"] = ("There is no history item for the current filter.", "برای فیلتر فعلی موردی در تاریخچه وجود ندارد."),
    ["PDFReportWasCreatedSentSuccessfully"] = ("PDF report was created/sent successfully", "گزارش PDF با موفقیت ساخته/ارسال شد"),
    ["SelectPDFExportType"] = ("Select PDF Export Type", "انتخاب نوع خروجی PDF"),
    ["SaveHistoryToExcel"] = ("Save History to Excel", "صادر کردن تاریخچه به Excel"),
    ["PDFCreatedSuccessfully"] = ("PDF created successfully!", "PDF با موفقیت ساخته شد!"),
    ["ThePDFReportWasCreatedOrSentToTheSelectedPrinterSuccessfully"] = ("The PDF report was created or sent to the selected printer successfully.", "گزارش PDF با موفقیت ساخته یا به چاپگر انتخاب‌شده ارسال شد."),
    ["OK"] = ("OK", "باشه"),
    ["DeviceDisconnected"] = ("Device disconnected", "اتصال دستگاه قطع شد"),
    ["DisconnectUnavailable"] = ("Disconnect unavailable", "قطع اتصال در دسترس نیست"),
    ["TheCurrentConnectionServiceDoesNotExposeADirectDisconnectMethodForThisDevice"] = ("The current connection service does not expose a direct disconnect method for this device.", "در سرویس فعلی برنامه، متد مستقیم برای قطع اتصال این دستگاه پیدا نشد."),
    ["Saved"] = ("Saved", "ثبت شد"),
    ["TtTeckSettingsWereSavedSuccessfully"] = ("TtTeck settings were saved successfully.", "تنظیمات تی‌تک با موفقیت ذخیره شد."),
    ["Error"] = ("Error", "خطا"),
    ["ReconnectEnabled"] = ("Reconnect enabled", "اتصال مجدد فعال شد"),
    ["PhonesThatWereManuallyDisconnectedCanConnectAgainNow"] = ("Phones that were manually disconnected can connect again now.", "گوشی‌هایی که دستی قطع شده بودند، الان می‌توانند دوباره وصل شوند."),
    ["BrowserError"] = ("Browser error", "خطای مرورگر"),
    ["NewBarcodeReceived"] = ("New barcode received", "بارکد جدید دریافت شد"),
    ["TtTeckProductFound"] = ("TtTeck product found", "محصول تی‌تک پیدا شد"),
    ["ScanSavedTtTeckResultUnavailable"] = ("Scan saved, TtTeck result unavailable", "اسکن ثبت شد، نتیجه تی‌تک در دسترس نیست"),
    ["NewScanSaved"] = ("New scan saved", "اسکن جدید ثبت شد"),
    ["TtTeckPlusOnlyWhenAnItemRegisteredInCargoDeliveryHasANearExpirationDateItIsFlaggedInTheSamePanel"] = ("TtTeck Plus only - when an item registered in \"Cargo delivery\" has a near expiration date, it is flagged in the same panel.", "مخصوص پلن تی‌تک پلاس - وقتی کالایی در «تحویل بار» با تاریخ انقضای نزدیک ثبت شود، در همان پنل هشدار داده می‌شود."),
    ["TappingThisButtonOpensBaleJustTapStartThereAndNearExpiryAlertsWillAlsoBeSentToYouOnBaleFromThenOn"] = ("Tapping this button opens Bale - just tap \"Start\" there, and near-expiry alerts will also be sent to you on Bale from then on.", "با زدن این دکمه، بله باز می‌شود؛ کافیست روی «شروع» بزنید تا از این پس هشدار تاریخ نزدیک برایتان در بله هم ارسال شود."),
    ["TTACSessionExpiredLoginAgainCargoDeliveryWillContinueAutomatically"] = ("TTAC session expired. Login again; cargo delivery will continue automatically.", "نشست تی‌تک منقضی شده است. دوباره وارد شوید؛ تحویل بار خودکار ادامه پیدا می‌کند."),
    ["TTACReturnedAnEmptyResultForThisReceiveStatusRequest"] = ("TTAC returned an empty result for this receive-status request.", "سامانه تی‌تک برای این درخواست تعیین وضعیت، پاسخ خالی برگرداند."),
    ["NoReceivableItemWasFoundForThisProductBatch"] = ("No receivable item was found for this product/batch.", "برای این فرآورده/سری ساخت مورد قابل تعیین وضعیت پیدا نشد."),
    ["TTACReturnedAnEmptyResponseAfterConfirmationPleaseRefreshCheckThePortal"] = ("TTAC returned an empty response after confirmation. Please refresh/check the portal.", "تی‌تک بعد از تأیید، پاسخ خالی برگرداند. لطفاً در سامانه بررسی/تازه‌سازی کنید."),
    ["ScanbridgeFormulaItemsReport"] = ("Scanbridge Formula Items Report", "گزارش شیر خشک Scanbridge"),
    ["ScanbridgeTtTeckItemsReport"] = ("Scanbridge TtTeck Items Report", "گزارش تی‌تک Scanbridge"),
    ["ItemRegistrationRequestWasSentSuccessfully"] = ("Item registration request was sent successfully.", "درخواست ثبت قلم با موفقیت ارسال شد."),
    ["TheTTACSessionHasExpiredOrYouHavenTLoggedInYetPleaseLogInUsingTheInternalBrowser"] = ("The TTAC session has expired or you haven't logged in yet. Please log in using the internal browser.", "نشست تی‌تک منقضی شده یا هنوز وارد نشده‌اید. برای ورود از مرورگر داخلی استفاده کنید."),
    ["YourTTACSessionExpiredLoginAgainThePreviousOperationWillContinueAutomatically"] = ("Your TTAC session expired. Login again; the previous operation will continue automatically.", "نشست تی‌تک منقضی شده است. دوباره وارد شوید؛ عملیات قبلی خودکار ادامه پیدا می‌کند."),
    ["SaveTheUsernamePasswordAndPharmacyNameALoginToButtonWillAppearOnTheMainScreenForEachSavedPharmacyAndItFillsTheCodeAndPasswordAutomaticallyOnTheTTACLoginPage"] = ("Save the username, password and pharmacy name; a \"Login to ...\" button will appear on the main screen for each saved pharmacy, and it fills the code and password automatically on the TTAC login page.", "نام کاربری، رمز عبور و نام داروخانه را ذخیره کنید؛ روی صفحهٔ اصلی برای هر داروخانه یک دکمهٔ «ورود به داروخانه ...» ساخته می‌شود که خودش کد و رمز را در صفحهٔ ورود تی‌تک پر می‌کند."),
    ["PleaseLogInUsingTheInternalBrowser"] = ("Please log in using the internal browser.", "لطفاً از مرورگر داخلی برای ورود استفاده کنید."),
    ["EnterUsernameAndPasswordOrUseTheInternalBrowser"] = ("Enter username and password, or use the internal browser.", "نام کاربری و رمز عبور را وارد کنید، یا از مرورگر داخلی استفاده کنید."),
    ["LoginTokenWasNotReceivedTheSiteMayRequireAnExtraStepThisFormCanTHandlePleaseUseInternalBrowserLogin"] = ("Login token was not received (the site may require an extra step this form can't handle). Please use internal browser login.", "توکن ورود دریافت نشد (ممکن است سایت یک مرحله‌ی اضافی مثل کد تایید بخواهد که این فرم نمی‌تواند انجام دهد). لطفاً از ورود با مرورگر داخلی استفاده کنید."),
    ["BaleRemindersForScanbridgeWereTurnedBackOn"] = ("Bale reminders for Scanbridge were turned back on.", "یادآور بله دوباره برای اسکن‌بریج فعال شد."),
    ["BaleBotIsNotConfiguredInThisAppBuildYet"] = ("Bale bot is not configured in this app build yet.", "ربات بله هنوز در این نسخه از برنامه تنظیم نشده است."),
    ["WaitingForConfirmationInBalePleaseWaitAFewSecondsAfterTappingStart"] = ("... waiting for confirmation in Bale - please wait a few seconds after tapping Start", "... منتظر تایید در بله - بعد از زدن «شروع» چند ثانیه صبر کنید"),
    ["TimedOutTapActivateBaleReminderAgain"] = ("Timed out. Tap \"Activate Bale reminder\" again.", "زمان تمام شد. دوباره روی «فعال‌سازی یادآور بله» بزنید."),
    ["ANewMonthHasStartedDoYouWantScanbridgeToCreateAMultiSheetExcelArchiveForThePreviousDataAndStartThisMonthWithACleanWorkspaceRawDataWillAlsoBeKeptInArchive"] = ("A new month has started. Do you want Scanbridge to create a multi-sheet Excel archive for the previous data and start this month with a clean workspace? Raw data will also be kept in Archive.", "ماه جدید شروع شده است. آیا می‌خواهید Scanbridge از اطلاعات قبلی یک خروجی Excel چندشیتی بگیرد و ماه جدید را با حافظه خالی شروع کند؟ داده خام هم در پوشه Archive نگه داشته می‌شود."),
    ["TheArchiveCouldNotBeFullyBackedUpAFileMayBeLockedOrTheDiskIsFullSoNothingWasDeletedPleaseTryAgain"] = ("The archive could not be fully backed up (a file may be locked or the disk is full), so nothing was deleted. Please try again.", "امکان پشتیبان‌گیری کامل از آرشیو نبود (ممکن است فایلی قفل باشد یا دیسک پر باشد)؛ به همین دلیل چیزی پاک نشد. لطفاً دوباره تلاش کنید."),
    ["TTACIsRespondingSlowlyOrDidNotRespondInTimeThisIsUsuallyCausedByHighTTACTrafficSlowInternetVPNProxyDNSIssuesOrTemporaryTTACOutagePleaseWaitAMomentAndTryAgain"] = ("TTAC is responding slowly or did not respond in time. This is usually caused by high TTAC traffic, slow internet, VPN/proxy/DNS issues, or temporary TTAC outage. Please wait a moment and try again.", "تی‌تک کند پاسخ می‌دهد یا در زمان مناسب پاسخ نداد. معمولاً به‌خاطر ترافیک بالای تی‌تک، کندی اینترنت، مشکل VPN/Proxy/DNS یا قطعی موقت تی‌تک است. کمی صبر کنید و دوباره تلاش کنید."),
    ["ConnectionToTTACFailedCheckInternetDNSVPNProxyOrTTACAvailability"] = ("Connection to TTAC failed. Check internet, DNS, VPN/proxy, or TTAC availability.", "اتصال به تی‌تک برقرار نشد. اینترنت، DNS، VPN/Proxy یا در دسترس بودن تی‌تک را بررسی کنید."),
    ["TTACLoginIsRequired"] = ("TTAC login is required.", "ورود به تی‌تک لازم است."),
    ["CaptchaReceivedEnterTheCodeAndCreatePrescription"] = ("Captcha received. Enter the code and create prescription.", "کپچا دریافت شد. کد را وارد کنید و نسخه را ایجاد کنید."),
    ["CompletionSMSWasSent"] = ("Completion SMS was sent.", "پیامک تکمیل نسخه ارسال شد."),
    ["ARegistrationRequestForThisExactItemAndPrescriptionMayHaveAlreadyReachedTTACOnceBeforeEGTheSessionExpiredRightAfterSendingSendingItAgainCouldRegisterItTwiceSendAgainAnyway"] = ("A registration request for this exact item and prescription may have already reached TTAC once before (e.g. the session expired right after sending). Sending it again could register it twice. Send again anyway?", "درخواست ثبت همین قلم برای همین نسخه احتمالاً قبلاً یک‌بار به تی‌تک ارسال شده (مثلاً درست بعد از ارسال، نشست منقضی شده). ارسال دوباره ممکن است باعث ثبت دوباره‌ی همین قلم در تی‌تک شود. مطمئنید می‌خواهید دوباره ارسال شود؟"),
    ["AutomaticResendSkippedByTheUserToAvoidADuplicateTTACRegistration"] = ("Automatic resend skipped by the user to avoid a duplicate TTAC registration.", "ارسال دوباره برای جلوگیری از ثبت تکراری در تی‌تک، توسط کاربر لغو شد."),
    ["ThisProductMayNotBeReceivedConfirmedForThisPharmacyPressOKToCheckReceiveStatus"] = ("This product may not be received/confirmed for this pharmacy. Press OK to check Receive Status.", "این فرآورده ممکن است تعیین وضعیت نشده باشد یا مربوط به داروخانه دیگری باشد. با زدن «باشه» به بخش تعیین وضعیت می‌روم."),
    ["ReceiveStatusWasCompletedPressOKToReturnToTheRegistrationForm"] = ("Receive status was completed. Press OK to return to the registration form.", "تعیین وضعیت انجام شد. با زدن «باشه» به فرم ثبت برمی‌گردید."),
    ["ThisItemWasAlreadyReceivedConfirmedPressOKToReturnToTheRegistrationForm"] = ("This item was already received/confirmed. Press OK to return to the registration form.", "این فرآورده قبلاً تعیین وضعیت شده است. با زدن «باشه» به فرم ثبت برمی‌گردید."),
    ["AlreadyRegisteredFormulaTitleForPhone"] = ("Already Registered", "قبلاً ثبت شده"),
    ["ThisProductWasAlreadyRegisteredForThisPharmacyForPhone"] = ("This product has already been registered for this pharmacy.", "این فرآورده قبلاً برای این داروخانه ثبت شده است."),
    ["ThisProductIsNotAvailableForTheCurrentPharmacyPressLogoutLogInToTheCorrectPharmacyThenScanTheBarcodeAndFillTheFormAgain"] = ("This product is not available for the current pharmacy. Press Logout, log in to the correct pharmacy, then scan the barcode and fill the form again.", "این فرآورده در این داروخانه نیست. لطفاً دکمه خروج را بزنید، به داروخانه مورد نظر وارد شوید و مجدد بارکد را اسکن و فرم را پر کنید."),
    ["RustDeskWasNotFoundPutRustdeskExeInsideTheSupportFolderNextToScanbridgeOrIncludeItInTheInstaller"] = ("RustDesk was not found. Put rustdesk.exe inside the Support folder next to Scanbridge, or include it in the installer.", "RustDesk پیدا نشد. فایل rustdesk.exe را داخل پوشه Support کنار Scanbridge بگذارید یا آن را داخل Installer قرار دهید."),
    ["MicrosoftEdgeWebView2RuntimeIsRequiredForTheInternalTTACBrowserPleaseInstallWebView2RuntimeThenOpenScanbridgeAgainDownloadHttpsDeveloperMicrosoftComMicrosoftEdgeWebview2"] = ("Microsoft Edge WebView2 Runtime is required for the internal TTAC browser. Please install WebView2 Runtime, then open Scanbridge again. Download: https://developer.microsoft.com/microsoft-edge/webview2/", "برای مرورگر داخلی تی‌تک، Microsoft Edge WebView2 Runtime لازم است. لطفاً WebView2 Runtime را نصب کنید و دوباره Scanbridge را باز کنید. لینک دانلود: https://developer.microsoft.com/microsoft-edge/webview2/"),
    ["TTACIsLoadingSlowlyThisCanHappenDuringHighTTACTrafficOrInternetVPNDNSProblemsPleaseWaitIfItDoesNotOpenCloseAndTryAgain"] = ("TTAC is loading slowly. This can happen during high TTAC traffic or internet/VPN/DNS problems. Please wait; if it does not open, close and try again.", "تی‌تک کند باز می‌شود. این حالت معمولاً هنگام ترافیک بالای تی‌تک یا مشکل اینترنت/VPN/DNS رخ می‌دهد. لطفاً کمی صبر کنید؛ اگر باز نشد، پنجره را ببندید و دوباره تلاش کنید."),
    ["RawJSONLicenseFilesAreNotAcceptedUseTheSignedActivationCode"] = ("Raw JSON license files are not accepted. Use the signed activation code.", "فایل JSON خام پذیرفته نمی‌شود. از کد فعال‌سازی امضاشده استفاده کنید."),
    ["NearExpiryAlert"] = ("📅 Near-expiry alert", "📅 هشدار تاریخ انقضای نزدیک"),
    ["HowManyMonthsBeforeExpiryToAlert"] = ("How many months before expiry to alert?", "چند ماه قبل از انقضا هشدار بدهد؟"),
    ["RepeatTheReminderEveryHowManyDays"] = ("Repeat the reminder every how many days?", "دوباره هر چند روز یادآوری کند؟"),
    ["BaleReminder"] = ("🔔 Bale reminder", "🔔 یادآور بله"),
    ["ActivateBaleReminder"] = ("Activate Bale reminder", "فعال‌سازی یادآور بله"),
    ["SendTestMessage"] = ("Send test message", "ارسال پیام آزمایشی"),
    ["DisplayRange"] = ("Display range:", "بازه‌ی نمایش:"),
    ["From"] = ("From", "از"),
    ["To"] = ("To", "تا"),
    ["ClearFilter2"] = ("Clear filter", "پاک کردن فیلتر"),
    ["ExportExcel"] = ("Export Excel", "خروجی اکسل"),
    ["TTACConnectionStatus"] = ("TTAC connection status", "وضعیت اتصال تی‌تک"),
    ["DisconnectTTAC"] = ("Disconnect TTAC", "قطع اتصال تی‌تک"),
    ["OpenTTACSite"] = ("Open TTAC site", "ورود به سایت تی‌تک"),
    ["OpenTTACWebsiteInTheInternalBrowser"] = ("Open TTAC website in the internal browser", "باز کردن سایت تی‌تک در مرورگر داخلی"),
    ["Connected"] = ("Connected", "متصل است"),
    ["NotConnected"] = ("Not connected", "متصل نیست"),
    ["TTACDisconnected"] = ("TTAC disconnected", "قطع اتصال تی‌تک"),
    ["TTACDisconnectedSuccess"] = ("Disconnected successfully", "با موفقیت قطع شد"),
    ["TTACTokenAndInternalBrowserSessionWereCleared"] = ("TTAC token and internal browser session were cleared.", "توکن تی‌تک و نشست مرورگر داخلی پاک شد."),
    ["Unlimited"] = ("Unlimited", "نامحدود"),
    ["Today"] = ("Today", "امروز"),
    ["ExpirationDateHasPassed"] = ("Expiration date has passed", "تاریخ انقضا گذشته است"),
    ["Sold"] = ("Sold", "فروخته شد"),
    ["GotIt"] = ("Got it", "حواسم هست"),
    ["UpcomingExpiry"] = ("Upcoming expiry", "تاریخ نزدیک"),
    ["SaveNearExpiryReport"] = ("Save near-expiry report", "ذخیره گزارش تاریخ نزدیک"),
    ["ExportSuccessful"] = ("Export successful", "خروجی موفق"),
    ["ExportFailed"] = ("Export failed", "خروجی ناموفق"),
    ["ManualUIDBarcode"] = ("Manual UID / barcode", "ورود دستی UID / بارکد"),
    ["InCargoModeScansAreAddedOnlyHere"] = ("In cargo mode, scans are added only here", "در حالت تحویل بار، اسکن‌ها فقط همین‌جا اضافه می‌شوند"),
    ["AddManually"] = ("Add manually", "افزودن دستی"),
    ["SelectAll"] = ("Select all", "انتخاب همه"),
    ["AddSelected"] = ("Add selected", "افزودن موارد انتخاب‌شده"),
    ["AddAll"] = ("Add all", "افزودن همه"),
    ["ClearAll"] = ("Clear all", "پاک کردن همه"),
    ["Checking"] = ("Checking...", "در حال بررسی..."),
    ["AddingToPharmacy"] = ("Adding to pharmacy...", "در حال افزودن به داروخانه..."),
    ["CargoDeliveryFailed"] = ("Cargo delivery failed", "تحویل بار ناموفق"),
    ["NoItemIsReadyToAdd"] = ("No item is ready to add.", "مورد آماده‌ای برای افزودن وجود ندارد."),
    ["CargoDeliveryCompleted"] = ("Cargo delivery completed", "تحویل بار انجام شد"),
    ["Product"] = ("Product", "نام فرآورده"),
    ["EnglishProduct"] = ("English product", "نام انگلیسی"),
    ["BarcodeUID"] = ("Barcode / UID", "بارکد / UID کامل"),
    ["LotBatch"] = ("Lot / batch", "سری ساخت"),
    ["Expiration"] = ("Expiration", "تاریخ انقضا"),
    ["GenericCode"] = ("Generic code", "کد ژنریک"),
    ["GenericName"] = ("Generic name", "نام ژنریک"),
    ["SenderDistributor"] = ("Sender / distributor", "ارسال‌کننده / توزیع‌کننده"),
    ["Quantity"] = ("Quantity", "تعداد"),
    ["SentDate"] = ("Sent date", "تاریخ ارسال"),
    ["Status"] = ("Status", "وضعیت"),
    ["ReceiveItemID"] = ("Receive item ID", "شناسه ردیف دریافت"),
    ["CargoItemDetails"] = ("Cargo item details", "جزئیات قلم تحویل بار"),
    ["ReceiveStatusDetails"] = ("Receive status details", "جزئیات تعیین وضعیت"),
    ["SaveCargoDeliveryReport"] = ("Save cargo delivery report", "ذخیره گزارش تحویل بار"),
    ["ScanbridgeCargoDelivery"] = ("Scanbridge Cargo Delivery", "گزارش تحویل بار Scanbridge"),
    ["PDFReport"] = ("PDF report", "گزارش PDF"),
    ["EnterUIDOrFullBarcodeToReceiveConfirm"] = ("Enter UID or full barcode to receive/confirm", "UID یا بارکد کامل را برای تعیین وضعیت وارد کنید"),
    ["Filter"] = ("Filter", "فیلتر"),
    ["SearchByProductBarcodeIRCLotDistributorDateOrStatus"] = ("Search by product, barcode, IRC, lot, distributor, date or status", "جستجو بر اساس محصول، بارکد، IRC، سری ساخت، پخش، تاریخ یا وضعیت"),
    ["Clear"] = ("Clear", "پاک کردن"),
    ["ManualEntry"] = ("Manual entry", "ورود دستی"),
    ["EnterTheUIDOrFullBarcodeFirst"] = ("Enter the UID or full barcode first.", "ابتدا UID یا بارکد کامل را وارد کنید."),
    ["ThisBarcodeIsAlreadyInTheReceiveStatusList"] = ("This barcode is already in the receive status list.", "این بارکد قبلاً در لیست تعیین وضعیت اضافه شده است."),
    ["ReceiveStatusFailed"] = ("Receive status failed", "تعیین وضعیت ناموفق"),
    ["ProductWasNotFoundInTTACCatalog"] = ("Product was not found in TTAC catalog.", "محصول در کاتالوگ تی‌تک پیدا نشد."),
    ["CatalogResponseIsIncomplete"] = ("Catalog response is incomplete.", "پاسخ کاتالوگ برای تعیین وضعیت کامل نیست."),
    ["ProductIDWasNotFound"] = ("Product ID was not found.", "شناسه فرآورده برای تعیین وضعیت پیدا نشد."),
    ["ReceiveItemIDWasNotFound"] = ("Receive item ID was not found.", "شناسه ردیف تعیین وضعیت پیدا نشد."),
    ["ReadyToConfirm"] = ("Ready to confirm", "آماده تعیین وضعیت"),
    ["AlreadyReceivedConfirmed"] = ("Already received/confirmed", "قبلاً تعیین وضعیت شده"),
    ["ProbablyAlreadyReceivedConfirmedNeedsReview"] = ("Probably already received/confirmed - needs review", "احتمالاً قبلاً تعیین وضعیت شده - نیاز به بررسی"),
    ["ThisItemIsNotConfirmable"] = ("This item is not confirmable.", "این مورد قابل تعیین وضعیت نیست."),
    ["ConfirmedSuccessfully"] = ("Confirmed successfully", "تعیین وضعیت با موفقیت انجام شد"),
    ["SaveReceiveStatusReport"] = ("Save receive status report", "ذخیره گزارش تعیین وضعیت"),
    ["ScanbridgeReceiveStatusReport"] = ("Scanbridge Receive Status Report", "گزارش تعیین وضعیت Scanbridge"),
    ["Excel"] = ("Excel", "Excel"),
    ["PDF"] = ("PDF", "PDF"),
    ["IfScanningIsNotPossibleEnterTheUIDOrFullBarcodeHere"] = ("If scanning is not possible, enter the UID or full barcode here", "اگر اسکن ممکن نبود، UID یا بارکد کامل را اینجا وارد کنید"),
    ["SearchByProductBarcodeNationalIDMobilePatientNameDateOrTime"] = ("Search by product, barcode, national ID, mobile, patient name, date or time", "جستجو بر اساس نام، بارکد، کد ملی، شماره همراه، نام بیمار، تاریخ یا ساعت"),
    ["SaveReportToExcel"] = ("Save report to Excel", "ذخیره گزارش در Excel"),
    ["ExcelReportWasSavedSuccessfullyN"] = ("Excel report was saved successfully:\n", "گزارش Excel با موفقیت ذخیره شد:\n"),
    ["InvalidUID"] = ("Invalid UID", "UID نامعتبر"),
    ["TheEnteredValueDoesNotLookLikeATTACUIDBarcode"] = ("The entered value does not look like a TTAC UID/barcode.", "مقدار واردشده شبیه UID یا بارکد تی‌تک نیست."),
    ["Adding"] = ("Adding...", "در حال افزودن..."),
    ["NationalIDMustBe10Digits"] = ("National ID must be 10 digits.", "کد ملی باید ۱۰ رقم باشد."),
    ["MobileNumberIsNotValid"] = ("Mobile number is not valid.", "شماره همراه معتبر نیست."),
    ["TTACRegistrationInquiryWasSuccessful"] = ("TTAC registration/inquiry was successful.", "استعلام/ثبت تی‌تک با موفقیت انجام شد."),
    ["PersianProductName"] = ("Persian product name", "نام فارسی فرآورده"),
    ["EnglishProductName"] = ("English product name", "نام انگلیسی فرآورده"),
    ["ProductName"] = ("Product name", "نام فرآورده"),
    ["BatchCode"] = ("Batch code", "سری ساخت"),
    ["PackageCount"] = ("Package count", "تعداد بسته"),
    ["PrescriptionItemID"] = ("Prescription item ID", "شناسه قلم نسخه"),
    ["InquiredAmount"] = ("Inquired amount", "مقدار استعلام‌شده"),
    ["ProductPrice"] = ("Product price", "قیمت فرآورده"),
    ["InsurancePayment"] = ("Insurance payment", "پرداختی بیمه"),
    ["PatientPayment"] = ("Patient payment", "پرداختی متقاضی"),
    ["CurrencyDifference"] = ("Currency difference", "مابه‌التفاوت ارزی"),
    ["TotalPrice"] = ("Total price", "قیمت کل"),
    ["Price"] = ("Price", "قیمت"),
    ["Payment"] = ("Payment", "پرداخت"),
    ["PrescriptionBased"] = ("Prescription-based", "نسخه‌محور"),
    ["WithoutPrescription"] = ("Without prescription", "فاقد نسخه"),
    ["RegistrationHistory"] = ("Registration history", "تاریخچه ثبت"),
    ["NoRegistrationHistoryWasFoundForThisProduct"] = ("No registration history was found for this product.", "برای این محصول سابقه ثبت پیدا نشد."),
    ["PrescriptionID"] = ("Prescription ID", "شناسه نسخه"),
    ["Amount"] = ("Amount", "تعداد"),
    ["NationalID"] = ("National ID", "کد ملی"),
    ["Patient"] = ("Patient", "بیمار"),
    ["Mobile"] = ("Mobile", "همراه"),
    ["RegistrationHistoryForThisProduct"] = ("Registration history for this product", "تاریخچه ثبت این محصول"),
    ["TTACLogin"] = ("TTAC login", "ورود به تی‌تک"),
    ["TtacQuickLoginSubtitle"] = ("Choose which pharmacy to log in with; the code and password are filled automatically.", "انتخاب کنید با کدام داروخانه وارد شوید؛ کد و رمز آن خودکار پر می‌شود."),
    ["SaveAccountToLoginFirst"] = ("No pharmacy saved yet. Save your TtTeck account here so you can log in with one click.", "هنوز داروخانه‌ای ذخیره نشده است. برای ورود به تی‌تک، ابتدا حساب خود را در همین بخش ذخیره کنید."),
    ["TtacSessionExpiredWarning"] = ("Your TtTeck session has expired. Choose a pharmacy to log in again.", "نشست تی‌تک منقضی شده است. برای ورود مجدد، یکی از داروخانه‌ها را انتخاب کنید."),
    ["TtacSessionExpiredWithPending"] = ("Your TtTeck session expired. After logging in, \"{0}\" will continue automatically.", "نشست تی‌تک منقضی شده است. بعد از ورود، «{0}» به‌طور خودکار ادامه پیدا می‌کند."),
    ["TtacSessionExpiredPendingGeneric"] = ("Your TtTeck session expired. After logging in, the pending operation will continue automatically.", "نشست تی‌تک منقضی شده است. بعد از ورود، عملیات در انتظار به‌طور خودکار ادامه پیدا می‌کند."),
    ["PendingOpenCargoDelivery"] = ("Opening the cargo delivery panel", "باز کردن پنل تحویل بار"),
    ["PendingOpenExpiryWatch"] = ("Opening the near-expiry panel", "باز کردن پنل تاریخ انقضای نزدیک"),
    ["PendingAddCargoBarcode"] = ("Adding the barcode to cargo delivery", "افزودن بارکد به تحویل بار"),
    ["PendingContinueCargoDelivery"] = ("Continuing cargo delivery", "ادامه‌ی تحویل بار"),
    ["PendingOpenReceiveStatus"] = ("Opening the receive-status panel", "باز کردن پنل تعیین وضعیت"),
    ["PendingAddReceiveStatusBarcode"] = ("Adding the barcode to receive status", "افزودن بارکد به تعیین وضعیت"),
    ["PendingConfirmReceiveStatus"] = ("Confirming receive status", "تأیید تعیین وضعیت"),
    ["PendingOpenTtacPanel"] = ("Opening the TtTeck panel", "باز کردن پنل تی‌تک"),
    ["PendingOpenRegistrationForm"] = ("Opening the TtTeck registration form", "باز کردن فرم ثبت تی‌تک"),
    ["PendingReturnToRegistrationForm"] = ("Returning to the registration form", "بازگشت به فرم ثبت"),
    ["TtacLoginFailedTitle"] = ("TtTeck login failed", "ورود به تی‌تک ناموفق بود"),
    ["TtacLoginFailedWithPharmacy"] = ("Logging in with \"{0}\" failed in the internal browser. Check the code and password of this pharmacy, or try again.", "ورود با «{0}» در مرورگر داخلی انجام نشد. کد و رمز این داروخانه را بررسی کنید یا دوباره تلاش کنید."),
    ["TtacLoginFailedGeneric"] = ("Logging in with the selected pharmacy failed in the internal browser. Check the code and password, or try again.", "ورود با داروخانه‌ی انتخاب‌شده در مرورگر داخلی انجام نشد. کد و رمز را بررسی کنید یا دوباره تلاش کنید."),
    ["RetryButton"] = ("Try again", "تلاش مجدد"),
    ["LoginWithoutSaving"] = ("Log in without saving an account", "ورود بدون ذخیره حساب"),
    ["LoginWithoutSavingHint"] = ("If you don't want to save an account, log in with the internal browser and enter the code and password manually.", "اگر نمی‌خواهید حسابی ذخیره کنید، با مرورگر داخلی وارد شوید و کد و رمز را دستی وارد کنید."),
    ["TtacLoginSuccessBanner"] = ("Login successful", "ورود موفق شد"),
    ["TtacLoginSuccessWithPending"] = ("Login successful; \"{0}\" has started.", "ورود موفق شد؛ «{0}» شروع شد."),
    ["TtacLoginStuckOnLoginPage"] = ("You are still on the TtTeck login page. The saved code or password may be wrong. Check the account in Settings or try again.", "هنوز روی صفحه‌ی ورود تی‌تک هستید. ممکن است کد یا رمز ذخیره‌شده درست نباشد. حساب را در تنظیمات بررسی کنید یا دوباره تلاش کنید."),
    ["TtacLoginStuckOnLoginPageWithPharmacy"] = ("You are still on the TtTeck login page. The saved code or password of \"{0}\" may be wrong. Check or update it in Settings, or try again.", "هنوز روی صفحه‌ی ورود تی‌تک هستید. به نظر می‌رسد کد یا رمز ذخیره‌شده‌ی «{0}» درست نیست. آن را در تنظیمات بررسی یا اصلاح کنید یا دوباره تلاش کنید."),
    ["PendingOpenRepeatFormulaRegistration"] = ("Opening repeat formula registration", "باز کردن ثبت مجدد شیرخشک"),
    ["PendingConfirmCargoRow"] = ("Confirming the cargo delivery item", "تأیید قلم تحویل بار"),
    ["PendingLoadCaptcha"] = ("Getting the captcha", "دریافت کپچا"),
    ["PendingCreatePrescription"] = ("Creating the prescription", "ایجاد نسخه"),
    ["PendingSubmitItem"] = ("Submitting the item", "ثبت قلم"),
    ["InternalBrowser"] = ("Internal browser", "مرورگر داخلی"),
    ["SavedTTACAccounts"] = ("🔑 Saved TTAC accounts", "🔑 حساب‌های ذخیره‌شده تی‌تک"),
    ["UsernamePharmacyCode"] = ("Username / pharmacy code", "نام کاربری / کد داروخانه"),
    ["Password"] = ("Password", "رمز عبور"),
    ["PharmacyName"] = ("Pharmacy name", "نام داروخانه"),
    ["NoSavedAccountsYet"] = ("No saved accounts yet.", "هنوز هیچ حسابی ذخیره نشده است."),
    ["EnterTheUsernameCodeAndThePassword"] = ("Enter the username/code and the password.", "نام کاربری (کد داروخانه) و رمز عبور را وارد کنید."),
    ["SavedAQuickLoginButtonWasAddedToTheMainScreen"] = ("Saved. A quick-login button was added to the main screen.", "ذخیره شد. دکمهٔ ورود سریع به صفحهٔ اصلی اضافه شد."),
    ["ThisCodeIsSaved"] = ("This code is saved.", "این کد قبلاً ذخیره شده است."),
    ["LoggingIn"] = ("Logging in...", "در حال ورود..."),
    ["TTACLoginPageWasNotReceived"] = ("TTAC login page was not received.", "صفحه ورود تی‌تک دریافت نشد."),
    ["TheLoginPageSecurityTokenWasNotFoundPleaseUseInternalBrowserLogin"] = ("The login page security token was not found. Please use internal browser login.", "توکن امنیتی صفحه ورود پیدا نشد. لطفاً از ورود با مرورگر داخلی استفاده کنید."),
    ["LoginFailedCheckUsernameOrPassword"] = ("Login failed. Check username or password.", "ورود ناموفق بود. نام کاربری یا رمز عبور را بررسی کنید."),
    ["Inactive"] = ("○ Inactive", "○ غیرفعال"),
    ["ActiveAlertsAreAlsoSentOnBale"] = ("● Active - alerts are also sent on Bale", "● فعال - هشدارها در بله هم ارسال می‌شود"),
    ["PausedBaleAlertsAreTurnedOff"] = ("⏸ Paused - Bale alerts are turned off", "⏸ متوقف‌شده - ارسال هشدار در بله خاموش است"),
    ["SendBaleAlerts"] = ("Send Bale alerts", "ارسال هشدار در بله فعال باشد"),
    ["BaleTest"] = ("Bale test", "تست یادآور بله"),
    ["TestMessageSentCheckBale"] = ("Test message sent - check Bale.", "پیام آزمایشی ارسال شد؛ بله را چک کنید."),
    ["SendingFailedCheckYourInternetConnection"] = ("Sending failed. Check your internet connection.", "ارسال پیام ناموفق بود. اتصال اینترنت را بررسی کنید."),
    ["BaleReminder2"] = ("Bale reminder", "یادآور بله"),
    ["MonthlyArchive"] = ("Monthly archive", "آرشیو ماهانه"),
    ["MonthlyArchiveWasCreatedSuccessfully"] = ("Monthly archive was created successfully.", "آرشیو ماهانه با موفقیت ساخته شد."),
    ["ArchiveFailed"] = ("Archive failed", "آرشیو ناموفق"),
    ["RequestFailed"] = ("Request failed.", "درخواست ناموفق بود."),
    ["CaptchaFailed"] = ("Captcha failed", "دریافت کپچا ناموفق"),
    ["NationalIDAndCaptchaAreRequired"] = ("National ID and captcha are required.", "کد ملی و کپچا الزامی است."),
    ["EnterBirthDateInTheCorrectFormatPersianExample13790620"] = ("Enter birth date in the correct format. Persian example: 1379/06/20", "تاریخ تولد را با فرمت صحیح وارد کنید. مثال شمسی: 1379/06/20"),
    ["MedicalCouncilNumberIsRequired"] = ("Medical council number is required.", "شماره نظام پزشکی الزامی است."),
    ["CreatingPrescriptionFailed"] = ("Creating prescription failed.", "ایجاد نسخه ناموفق بود."),
    ["PrescriptionIDWasNotFoundInThePortalResponse"] = ("Prescription ID was not found in the portal response.", "شناسه نسخه در پاسخ سامانه پیدا نشد."),
    ["CreatePrescription"] = ("Create prescription", "ایجاد نسخه"),
    ["CreatePrescriptionFailed"] = ("Create prescription failed", "ایجاد نسخه ناموفق"),
    ["NoBarcodeIsSelected"] = ("No barcode is selected.", "هیچ بارکدی انتخاب نشده است."),
    ["CreateThePrescriptionFirst"] = ("Create the prescription first.", "ابتدا نسخه را ایجاد کنید."),
    ["MobileNumberIsRequiredForFormulaRegistration"] = ("Mobile number is required for formula registration.", "برای ثبت شیر خشک، شماره تماس الزامی است."),
    ["PossibleDuplicateSubmission"] = ("Possible duplicate submission", "احتمال ثبت تکراری"),
    ["SubmitItem"] = ("Submit item", "ثبت قلم"),
    ["SubmitItemFailed"] = ("Submit item failed", "ثبت قلم ناموفق"),
    ["RegisterAnother"] = ("Register another", "ثبت مجدد"),
    ["OpenLink3"] = ("Open link", "باز کردن لینک"),
    ["ReceiveStatusCompleted"] = ("Receive status completed", "تعیین وضعیت انجام شد"),
    ["AlreadyReceived"] = ("Already received", "قبلاً تعیین وضعیت شده"),
    ["WrongPharmacy"] = ("Wrong pharmacy", "فرآورده در این داروخانه نیست"),
    ["ProductNotInPharmacySwitchHint"] = ("This product is not in this pharmacy. Select another pharmacy to check:", "این فرآورده در این داروخانه نیست. داروخانه دیگری را برای بررسی انتخاب کنید:"),
    ["LoginToPharmacy"] = ("Login to {0}", "ورود به {0}"),
    ["Logout"] = ("Logout", "خروج"),
    ["ReturnedFromReceiveStatus"] = ("Returned from receive status", "بازگشت از تعیین وضعیت"),
    ["RegisterAnotherFormula"] = ("Register another formula", "ثبت مجدد شیر خشک"),
    ["EnterOrScanTheNewFormulaBarcode"] = ("Enter or scan the new formula barcode", "بارکد شیرخشک جدید را وارد یا اسکن کنید"),
    ["Continue"] = ("Continue", "ادامه"),
    ["PreviousFormulaRegistrationInformationIsNotAvailable"] = ("Previous formula registration information is not available.", "اطلاعات ثبت قبلی شیر خشک در دسترس نیست."),
    ["RepeatFormula"] = ("Repeat formula", "ثبت مجدد شیر خشک"),
    ["SupportToolNotFound"] = ("Support tool not found", "ابزار پشتیبانی پیدا نشد"),
    ["SupportError"] = ("Support error", "خطای پشتیبانی"),
    ["ErrorOpeningSupportTool"] = ("Error opening support tool: ", "خطا در اجرای ابزار پشتیبانی: "),
    ["ConnectingToTTAC"] = ("Connecting to TTAC... ", "در حال اتصال به تی‌تک... "),
    ["TTACLoginCompletedContinuingTheOperation"] = ("TTAC login completed. Continuing the operation.", "ورود تی‌تک انجام شد. ادامه عملیات انجام می‌شود."),
    ["NoPhoneConnected"] = ("No phone connected", "گوشی متصل نیست"),
    ["Connected2"] = ("Connected", "متصل"),
    ["Unknown"] = ("Unknown", "نامشخص"),
    ["ScanBridge"] = ("ScanBridge", "اسکن بریج"),
    ["LicenseCodeFormatIsInvalid"] = ("License code format is invalid.", "فرمت کد لایسنس نامعتبر است."),
    ["LicenseSignatureIsInvalid"] = ("License signature is invalid.", "امضای لایسنس معتبر نیست."),
    ["ActivationFailed"] = ("Activation failed.", "فعال‌سازی ناموفق بود."),
    ["SignedLicenseWasNotReturnedByServer"] = ("Signed license was not returned by server.", "سرور کد امضاشده لایسنس را برنگرداند."),
    ["ActivatingLicense"] = ("Activating license...", "در حال فعال‌سازی لایسنس..."),
    ["InvalidLicense"] = ("Invalid license", "لایسنس نامعتبر"),
    ["ThisLicenseIsNotIssuedForThisSystem"] = ("This license is not issued for this system.", "این لایسنس برای این سیستم صادر نشده است."),
    ["WrongSystem"] = ("Wrong system", "سیستم نامعتبر"),
    ["LicenseActivatedSuccessfully"] = ("License activated successfully.", "لایسنس با موفقیت فعال شد."),
    ["ActivationFailed2"] = ("Activation failed", "فعال‌سازی ناموفق"),
    ["ScanbridgeLicense"] = ("Scanbridge License", "لایسنس Scanbridge"),
    ["SystemID"] = ("System ID", "شناسه سیستم"),
    ["LicenseKey"] = ("License key", "کلید لایسنس"),
    ["ActivateOnline"] = ("Activate online", "فعال‌سازی آنلاین"),
    ["ImportFile"] = ("Import file", "ورود فایل"),
    ["ExportID"] = ("Export ID", "خروجی شناسه"),
    ["Plan"] = ("Plan", "پلن"),
    ["User"] = ("User", "کاربر"),
    ["Expires"] = ("Expires", "تاریخ انقضا"),
    ["LastUpdate"] = ("Last update", "آخرین بروزرسانی"),
    ["Phone"] = ("Phone", "شماره تماس"),
    ["LicenseIsActive"] = ("License is active", "لایسنس فعال است"),
    ["LicenseExpired"] = ("License expired", "لایسنس منقضی شده"),
    ["NotActivated"] = ("Not activated", "فعال نشده"),
    ["SystemIDCopied"] = ("System ID copied.", "شناسه سیستم کپی شد."),
    ["ImportLicenseFile"] = ("Import license file", "وارد کردن فایل لایسنس"),
    ["ExportSystemID"] = ("Export system ID", "خروجی شناسه سیستم"),
    ["SystemIDExported"] = ("System ID exported.", "شناسه سیستم ذخیره شد."),
    ["AddToPharmacy"] = ("Add to pharmacy", "افزودن به داروخانه"),
    ["NeedsReview"] = ("Needs review", "نیاز به بررسی"),
    ["ReadyToAdd"] = ("Ready to add", "آماده افزودن به داروخانه"),
    ["NotAvailable"] = ("Not available", "قابل افزودن نیست"),
    ["InPharmacy"] = ("In pharmacy", "موجود در داروخانه"),
    ["AddedToPharmacy"] = ("Added to pharmacy", "به داروخانه اضافه شد"),
    ["NearExpiryAlert2"] = ("Near-expiry alert", "هشدار تاریخ انقضای نزدیک"),
    ["GotItThanks"] = ("Got it, thanks", "حواسم هست، مرسی از یادآوری"),
    };

    private static readonly Dictionary<string, (string[] English, string[] Persian)> _arrayTranslations = new()
    {
    ["ARRAY:No"] = (new[] { "No.", "Date", "Time", "Device", "Barcode", "Product / Status" }, new[] { "ردیف", "تاریخ", "ساعت", "دستگاه", "بارکد", "محصول / وضعیت" }),
    ["ARRAY:No2"] = (new[] { "No.", "Date", "Time", "Device", "Barcode", "Persian Product Name", "English Product Name" }, new[] { "ردیف", "تاریخ", "ساعت", "دستگاه", "بارکد", "نام فارسی محصول", "نام انگلیسی محصول" }),
    ["ARRAY:No3"] = (new[] { "No.", "Date", "Time", "Device Name", "Barcode", "Persian Product Name", "English Product Name" }, new[] { "ردیف", "تاریخ", "ساعت", "اسم دستگاه", "بارکد", "نام فارسی محصول", "نام انگلیسی محصول" }),
    ["ARRAY:No4"] = (new[] { "No.", "Date", "Time", "Device Name", "Barcode", "Drug Name" }, new[] { "ردیف", "تاریخ", "ساعت", "اسم دستگاه", "بارکد", "نام دارو" }),
    ["ExpiryWatchHeaders"] = (new[] { "No.", "Product", "Barcode / UID", "Lot", "Expiry date", "Status" }, new[] { "ردیف", "محصول", "بارکد / UID", "سری ساخت", "تاریخ انقضا", "وضعیت" }),
    ["CargoDeliveryExportHeaders"] = (new[] { "No.", "Selected", "Product", "English product", "Barcode / UID", "IRC", "Lot", "Qty", "Sender", "Status" }, new[] { "ردیف", "انتخاب", "محصول", "نام انگلیسی", "بارکد / UID", "IRC", "سری ساخت", "تعداد", "ارسال‌کننده", "وضعیت" }),
    ["ReceiveStatusExportHeaders"] = (new[] { "No.", "Product", "English product", "Barcode / UID", "IRC", "Generic code", "Lot", "Qty", "Sender", "Sent date", "Status" }, new[] { "ردیف", "محصول", "نام انگلیسی", "بارکد / UID", "IRC", "کد ژنریک", "سری ساخت", "تعداد", "ارسال‌کننده", "تاریخ ارسال", "وضعیت" }),
    ["TtacPanelExportHeaders"] = (new[] { "No.", "Date", "Time", "Device", "Barcode / UID", "Persian product", "English product", "Registration status" }, new[] { "ردیف", "تاریخ", "ساعت", "دستگاه", "بارکد / UID", "نام فارسی محصول", "نام انگلیسی محصول", "وضعیت ثبت" }),
    };

    public string GetString(string key)
    {
        if (_translations.TryGetValue(key, out var pair))
            return CurrentLanguage == AppLanguage.English ? pair.English : pair.Persian;
        return key;
    }

    public string GetFormattedString(string key, params object[] args)
    {
        string template = GetString(key);
        return args.Length == 0 ? template : string.Format(template, args);
    }

    public string[] GetStringArray(string key)
    {
        if (_arrayTranslations.TryGetValue(key, out var pair))
            return CurrentLanguage == AppLanguage.English ? pair.English : pair.Persian;
        return new[] { key };
    }}
