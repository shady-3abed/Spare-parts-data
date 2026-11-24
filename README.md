# لوحة قطع الغيار — بيانات الموقع

## إيه ده؟

المستودع ده فيه بيانات لوحة قطع الغيار بتاعة قسم الصيانة.
اللوحة نفسها بتقرأ ملف `website_data.json` من هنا، فأي تحديث للملف ده
بيظهر على اللوحة على طول من غير ما حد يبني الموقع من جديد.

رابط اللوحة:
https://muse.ai/s/spare-parts-dashboard-qt-lgyr-xzxb73xkj6xgmxz1/_92EhcGySROjm9_7edisxBVc4ki_e8qTztsIRufEokw

## إزاي تحدّث البيانات؟ (من غير ما تستنى حد)

1. افتح المستودع ده على موقع github.com.
2. ادخل على فولدر `workbook`.
3. دوس على **Add file** وبعدين **Upload files**.
4. حط ملف الشيت المحدّث `All_Machines_Spareparts_Final.xlsm` — **لازم نفس الاسم بالظبط**،
   ولو سألك عن استبدال الملف القديم دوس تأكيد.
5. دوس على **Commit changes** (الزرار الأخضر تحت).
6. استنى حوالي دقيقة: فيه مهمة أوتوماتيك بتقرأ الشيت الجديد وتحدّث ملف
   `website_data.json` لوحدها.
7. افتح رابط اللوحة وحدّث الصفحة — هتلاقي الداتا الجديدة ظاهرة للفريق كله.

## ملاحظات مهمة

- الملف اللي ترفعه لازم يكون نسخة كاملة من شيت قطع الغيار، بنفس أسماء الشيتات
  (الماكينات الـ 30 + Stock_Analysis + Data_Integrity_Check + غيرها).
  لو الملف ناقص شيتات، المهمة الأوتوماتيك هتفشل ومش هيتحدّث حاجة.
- المستودع ده **عام (public)** عشان صفحة اللوحة تقدر تقرأ ملف البيانات منه.
  أي حد معاه رابط اللوحة يقدر يشوف الداتا — واللوحة نفسها محمية بباسوورد.

## الملفات

- `workbook/All_Machines_Spareparts_Final.xlsm` — ملف الشيت (أنت بترفعه).
- `website_data.json` — بيانات اللوحة (بيتولّد أوتوماتيك، متعدّلش فيه بإيدك).
- `scripts/build_data.py` — السكربت اللي بيحوّل الشيت لملف البيانات.
- `.github/workflows/update-data.yml` — المهمة الأوتوماتيك اللي بتشغّل السكربت
  كل مرة يترفع فيها شيت جديد.

---

# Spare Parts Dashboard — data repo

## What is this?

This repo holds the data for the plant maintenance spare-parts dashboard.
The dashboard page reads `website_data.json` from here, so updating that
file updates the dashboard for everyone — no site rebuild needed.

Dashboard link:
https://muse.ai/s/spare-parts-dashboard-qt-lgyr-xzxb73xkj6xgmxz1/_92EhcGySROjm9_7edisxBVc4ki_e8qTztsIRufEokw

## How to update the data (no waiting on anyone)

1. Open this repo on github.com.
2. Open the `workbook` folder.
3. Click **Add file** → **Upload files**.
4. Drop the updated `All_Machines_Spareparts_Final.xlsm` — **exactly the same
   file name** — and confirm replacing the old one if asked.
5. Click **Commit changes** (the green button).
6. Wait about a minute: an automatic job reads the new workbook and
   regenerates `website_data.json`.
7. Refresh the dashboard link — the team will see the new data.

## Notes

- The uploaded workbook must be the complete spare-parts workbook with the
  same sheet names (the 30 machine sheets + Stock_Analysis +
  Data_Integrity_Check, etc.). If sheets are missing, the automatic job
  fails and nothing updates.
- This repo is **public** so the dashboard page can read the data file.
  Anyone with the dashboard link can see the data — the dashboard page
  itself is protected by a password.

## Files

- `workbook/All_Machines_Spareparts_Final.xlsm` — the workbook (you upload it).
- `website_data.json` — dashboard data (auto-generated, do not edit by hand).
- `scripts/build_data.py` — the script that converts the workbook to the data file.
- `.github/workflows/update-data.yml` — the automation that runs the script
  whenever a new workbook is uploaded.
