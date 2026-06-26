import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from PIL import Image, ImageDraw, ImageFont
import os
import sys
import re
import json
import sqlite3
import datetime

# <--- تلاش برای ایمپورت کتابخانه اکسل --->
try:
    import openpyxl

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("کتابخانه openpyxl نصب نیست. خروجی اکسل غیرفعال شد.")
    print("برای نصب: pip install openpyxl")

# <--- START: افزودن فولدر دیتا --->
DATA_FOLDER = "data"

if not os.path.exists(DATA_FOLDER):
    try:
        os.makedirs(DATA_FOLDER)
        print(f"فولدر '{DATA_FOLDER}' ساخته شد.")
    except OSError as e:
        messagebox.showerror("خطای ساخت فولدر", f"امکان ساخت فولدر '{DATA_FOLDER}' وجود ندارد.\n{e}")


# <--- END: افزودن فولدر دیتا --->


# <--- START: تابع کمکی برای PyInstaller --->
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# <--- END: تابع کمکی برای PyInstaller --->


# --- 1. تنظیمات اصلی ---
BASE_IMAGE_PATH = resource_path(os.path.join(DATA_FOLDER, "Main.jpg"))
# FONT_PATH = resource_path(os.path.join(DATA_FOLDER, "B-NAZANIN.TTF"))
# FONT_PATH = resource_path(os.path.join(DATA_FOLDER, "Vazir.ttf"))
FONT_PATH = resource_path(os.path.join(DATA_FOLDER, "XB Niloofar.ttf"))


MAIN_COLOR = "navy"
FONT_SIZE = 51
ADDRESS_FONT_SIZE = 48
CHECKMARK_FONT_SIZE = 75
INVOICE_FONT_SIZE = 75
COUNTER_FILE = os.path.join(DATA_FOLDER, "counter.txt")
DB_FILE = os.path.join(DATA_FOLDER, "transactions.db")
EXCEL_FILE = os.path.join(DATA_FOLDER, "transactions.xlsx")

DESC_MAX_WIDTH_PX = 3164

# --- تنظیمات دیتابیس و اکسل (فقط فیلدهای درخواستی) ---
# شماره قولنامه (سریال) به صورت جداگانه در تابع اضافه می‌شود
DB_COLUMN_MAPPING = {
    "trans_date_year_xx": "سال_معامله",
    "trans_date_month": "ماه_معامله",
    "trans_date_day": "روز_معامله",
    "seller_name": "نام_فروشنده",
    "buyer_name": "نام_خریدار",
    "car_type": "نوع_اتومبیل",
    "car_tip": "تیپ",
    "car_model": "مدل",
    "trans_total_price": "قیمت_کل"
}

# --- 2. مختصات (x, y) ---
COORDS = {
    # --- فروشنده ---
    "seller_name": (2831, 494),
    "seller_father": (2125, 497),
    "seller_sh_sh": (1168, 509),
    "seller_birth_year": (280, 519),
    "seller_sodoor": (674, 519),
    "seller_code_melli": (2626, 581),
    "seller_shenase_melli": (3108, 581),
    "seller_address": (2115, 581),
    "seller_zip_code": (986, 596),
    "seller_phone": (542, 609),
    "seller_owner": (1900, 666),
    "seller_mojieb": (1329, 673),

    # --- خریدار ---
    "buyer_name": (2886, 729),
    "buyer_father": (2115, 736),
    "buyer_sh_sh": (1176, 739),
    "buyer_birth_year": (264, 739),
    "buyer_sodoor": (674, 748),
    "buyer_code_melli": (2629, 799),
    "buyer_shenase_melli": (3113, 801),
    "buyer_address": (2118, 801),
    "buyer_zip_code": (989, 806),
    "buyer_phone": (541, 819),
    "buyer_owner": (1901, 884),
    "buyer_mojieb": (1340, 890),

    # --- اتومبیل ---
    "car_dang": (2988, 1029),
    "car_type": (2385, 1005),
    "car_system": (2076, 1034),
    "car_tip": (1792, 1031),
    "car_model": (1575, 1034),
    "car_color": (1356, 1029),
    "car_capacity": (1024, 1031),
    "car_cylinders": (668, 1029),
    "car_fuel": (301, 1036),
    "car_axles": (3108, 1101),
    "car_wheels": (2746, 1098),
    "car_motor_no": (2380, 1106),
    "car_shasi_no": (1543, 1111),
    "car_pelak": (568, 1111),
    "car_owner_mojeb": (1136, 1189),
    "car_lavazem": (3220, 1240),

    # --- معامله ---
    "trans_total_price": (2696, 1396),
    "trans_total_price_words": (1591, 1398),
    "trans_biane_words": (2658, 1473),
    "trans_baghimandeh_words": (2522, 1555),

    "check_number": (1361, 1492),
    "check_bank": (825, 1482),
    "check_branch": (540, 1482),
    "check_code": (331, 1492),

    "doc_day": (2135, 1856),
    "doc_month": (2045, 1856),
    "doc_year_xx": (1904, 1856),
    "doc_date_words": (1583, 1842),

    "delivery_day": (2842, 2815),
    "delivery_month": (2759, 2815),
    "delivery_year_xx": (2653, 2815),
    "delivery_time": (1630, 2817),
    "delivery_date_words": (2367, 2807),

    "hazine_enteghal_text": (1360, 2438),

    "after_doc_day": (917, 2828),
    "after_doc_month": (819, 2828),
    "after_doc_year_xx": (699, 2828),

    "trans_time": (1768, 3405),
    "trans_desc_1": (3324, 3757),
    "trans_desc_2": (3324, 3827),
    "trans_desc_3": (3324, 3897),
    "trans_desc_4": (3324, 3967),

    "trans_date_day": (2969, 3403),
    "trans_date_month": (2865, 3403),
    "trans_date_year_xx": (2757, 3403),
    "trans_date_words": (2468, 3401),

    "delay_cost_doc": (3120, 3038),
    "daily_penalty": (1332, 2682),

    "commission_amount": (2203, 3187),
    "vahed_senfi_1": (1595, 3183),
    "vahed_senfi_2": (818, 3421),
    "commission_pay_day": (374, 3190),
    "commission_pay_month": (294, 3190),
    "commission_pay_year_xx": (217, 3190),
}

# --- بارگذاری فونت‌ها ---
try:
    FONT_MAIN = ImageFont.truetype(FONT_PATH, size=FONT_SIZE)
    FONT_ADDRESS = ImageFont.truetype(FONT_PATH, size=ADDRESS_FONT_SIZE)
    FONT_CHECKMARK = ImageFont.truetype(FONT_PATH, size=CHECKMARK_FONT_SIZE)
    FONT_INVOICE = ImageFont.truetype(FONT_PATH, size=INVOICE_FONT_SIZE)

    try:
        if sys.platform == "win32":
            FONT_ENGLISH_MAIN = ImageFont.truetype("arial.ttf", size=FONT_SIZE)
            FONT_ENGLISH_ADDRESS = ImageFont.truetype("arial.ttf", size=ADDRESS_FONT_SIZE)
        else:
            FONT_ENGLISH_MAIN = ImageFont.load_default()
            FONT_ENGLISH_ADDRESS = ImageFont.load_default()
    except:
        FONT_ENGLISH_MAIN = ImageFont.load_default()
        FONT_ENGLISH_ADDRESS = ImageFont.load_default()

except IOError:
    messagebox.showerror("خطای فونت", f"فایل فونت {FONT_PATH} یافت نشد. برنامه بسته خواهد شد.")
    FONT_MAIN = ImageFont.load_default()
    FONT_ADDRESS = FONT_MAIN
    FONT_CHECKMARK = FONT_MAIN
    FONT_INVOICE = FONT_MAIN
    FONT_ENGLISH_MAIN = FONT_MAIN
    FONT_ENGLISH_ADDRESS = FONT_MAIN

# --- 3. آماده‌سازی متن فارسی ---
try:
    from arabic_reshaper import reshape
    from bidi.algorithm import get_display

    HAS_BIDI = True
except ImportError:
    HAS_BIDI = False
    print("هشدار: کتابخانه‌های bidi و reshaper نصب نیستند.")


def is_persian(text):
    if not text:
        return False
    return bool(re.search(r'[\u0600-\u06FF\uFB8A\u067E\u0686\u06AF]', text))


def prepare_text(text):
    if HAS_BIDI:
        if is_persian(text):
            try:
                reshaped_text = reshape(text)
                bidi_text = get_display(reshaped_text)
                return bidi_text
            except Exception:
                return text
        else:
            return text
    return text


# --- 4. توابع تبدیل عدد به حروف ---
def _3_digit_to_persian(num_str):
    ones = ["", "یک", "دو", "سه", "چهار", "پنج", "شش", "هفت", "هشت", "نه"]
    tens = ["", "ده", "بیست", "سی", "چهل", "پنجاه", "شصت", "هفتاد", "هشتاد", "نود"]
    hundreds = ["", "صد", "دویست", "سیصد", "چهارصد", "پانصد", "ششصد", "هفتصد", "هشتصد", "نهصد"]
    teens = ["ده", "یازده", "دوازده", "سیزده", "چهارده", "پانزده", "شانزده", "هفده", "هجده", "نوزده"]
    num = int(num_str)
    parts = []
    if num >= 100:
        parts.append(hundreds[num // 100])
        num %= 100
        if num > 0:
            parts.append("و")
    if num >= 20:
        parts.append(tens[num // 10])
        num %= 10
        if num > 0:
            parts.append("و")
    if 10 <= num <= 19:
        parts.append(teens[num - 10])
    elif num > 0:
        parts.append(ones[num])
    return " ".join(parts)


def num_to_persian_words(num_str):
    num_str = "".join(filter(str.isdigit, num_str))
    if not num_str:
        return ""
    try:
        num = int(num_str)
        if num == 0:
            return "صفر"
    except ValueError:
        return ""
    magnitudes = ["", "هزار", "میلیون", "میلیارد", "تریلیون"]
    num_str = str(num)
    groups = []
    while len(num_str) > 0:
        groups.append(num_str[-3:])
        num_str = num_str[:-3]
    result_parts = []
    for i, group in enumerate(groups):
        if int(group) > 0:
            group_words = _3_digit_to_persian(group)
            magnitude_word = magnitudes[i]
            if result_parts:
                result_parts.insert(0, "و")
            result_parts.insert(0, f"{group_words} {magnitude_word}".strip())
    return " ".join(result_parts).strip()


# --- 5. توابع فرمت‌دهی خودکار ---
_formatting_flags = {}
_validating_flags = {}


def calculate_and_update_baghimandeh():
    try:
        total_str = string_vars["trans_total_price"].get()
        biane_str = string_vars["trans_biane"].get()
        clean_total = int("".join(filter(str.isdigit, total_str))) if total_str else 0
        clean_biane = int("".join(filter(str.isdigit, biane_str))) if biane_str else 0
        remaining = clean_total - clean_biane
        string_vars["trans_baghimandeh"].set(f'{remaining:,}')
        persian_words = num_to_persian_words(str(remaining))
        if persian_words:
            string_vars["trans_baghimandeh_words"].set(f"{persian_words} ریال")
        else:
            string_vars["trans_baghimandeh_words"].set("")
    except Exception as e:
        print(f"Error calculating remaining: {e}")
        string_vars["trans_baghimandeh"].set("خطا")
        string_vars["trans_baghimandeh_words"].set("")


def on_generic_price_change(value_var, words_var, widget_key, needs_remaining_update=False):
    if _formatting_flags.get(widget_key, False):
        return
    _formatting_flags[widget_key] = True

    widget = entries.get(widget_key)
    if not widget:
        _formatting_flags[widget_key] = False
        return

    cursor_pos = widget.index(tk.INSERT)
    text_content = value_var.get()

    digits_to_left = 0
    for i, char in enumerate(text_content):
        if i >= cursor_pos:
            break
        if char.isdigit():
            digits_to_left += 1

    clean_num = "".join(filter(str.isdigit, text_content))
    persian_words = num_to_persian_words(clean_num)
    words_var.set(f"{persian_words} ریال" if persian_words else "")
    formatted_num = f'{int(clean_num):,}' if clean_num else ""
    value_var.set(formatted_num)

    new_cursor_pos = 0
    current_digit_count = 0

    if digits_to_left > 0:
        for index, char in enumerate(formatted_num):
            if char.isdigit():
                current_digit_count += 1
            if current_digit_count == digits_to_left:
                new_cursor_pos = index + 1
                break
    else:
        new_cursor_pos = 0

    def set_cursor():
        try:
            widget.icursor(new_cursor_pos)
        except tk.TclError:
            pass

    widget.after_idle(set_cursor)
    if needs_remaining_update:
        calculate_and_update_baghimandeh()
    _formatting_flags[widget_key] = False


def validate_date_part(var, max_val, max_len=2):
    current_val = var.get()
    clean_val = "".join(filter(str.isdigit, current_val))
    if clean_val:
        try:
            num = int(clean_val)
            if num > max_val:
                clean_val = str(max_val)
            elif num == 0 and len(clean_val) > 1:
                clean_val = "0"
            elif max_len and len(clean_val) > max_len:
                clean_val = clean_val[-max_len:]
            if current_val != clean_val:
                var.set(clean_val)
        except ValueError:
            clean_val = ""
            var.set(clean_val)
    return clean_val


def on_generic_date_change(flag_key, day_var, month_var, year_var, words_var=None):
    if _validating_flags.get(flag_key, False):
        return
    _validating_flags[flag_key] = True

    day_clean = validate_date_part(day_var, 31, 2)
    month_clean = validate_date_part(month_var, 12, 2)
    year_clean = validate_date_part(year_var, 99, 2)

    if words_var is not None:
        persian_date_str = convert_date_to_persian_words(day_clean, month_clean, year_clean)
        words_var.set(persian_date_str)

    _validating_flags[flag_key] = False


def convert_date_to_persian_words(day_str, month_str, year_xx_str):
    try:
        day_clean = "".join(filter(str.isdigit, day_str))
        month_clean = "".join(filter(str.isdigit, month_str))
        year_clean = "".join(filter(str.isdigit, year_xx_str))
        if not day_clean or not month_clean or not year_clean:
            return ""
        day_num = int(day_clean)
        month_num = int(month_clean)
        year_num = int(year_clean)
    except ValueError:
        return ""
    months = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]
    if not (1 <= month_num <= 12):
        return ""
    if not (1 <= day_num <= 31):
        return ""
    day_words = _3_digit_to_persian(str(day_num))
    month_name = months[month_num - 1]
    year_words = num_to_persian_words(str(year_num))
    if not day_words or year_words == "":
        return ""
    return f"{day_words} {month_name} ماه چهارصد و {year_words}"


def wrap_text_by_pixel_width(text, font, max_width_px):
    wrapped_lines = []
    paragraphs = text.split('\n')
    for paragraph in paragraphs:
        if not paragraph:
            wrapped_lines.append("")
            continue
        words = paragraph.split()
        if not words:
            continue
        current_line = ""
        for word in words:
            if not current_line:
                test_line = word
            else:
                test_line = current_line + " " + word
            prepared_test_line = prepare_text(test_line)
            try:
                line_width = font.getlength(prepared_test_line)
            except AttributeError:
                line_width = font.getsize(prepared_test_line)[0]
            if line_width <= max_width_px:
                current_line = test_line
            else:
                if not current_line:
                    wrapped_lines.append(test_line)
                    current_line = ""
                else:
                    wrapped_lines.append(current_line)
                    current_line = word
        if current_line:
            wrapped_lines.append(current_line)
    return wrapped_lines


def sanitize_filename(name):
    if not name:
        return "بدون_نام"
    invalid_chars = r'\/:*?"<>|'
    clean_name = name
    for char in invalid_chars:
        clean_name = clean_name.replace(char, '')
    clean_name = " ".join(clean_name.split()).strip().strip('.')
    if not clean_name:
        return "بدون_نام"
    return clean_name


def load_counter():
    try:
        with open(COUNTER_FILE, 'r') as f:
            return int(f.read().strip())
    except (FileNotFoundError, ValueError, IOError):
        return 0


def save_counter(value):
    try:
        with open(COUNTER_FILE, 'w') as f:
            f.write(str(value))
    except IOError as e:
        print(f"خطا در ذخیره شمارنده: {e}")
        messagebox.showwarning("خطا", f"شماره سریال در فایل {COUNTER_FILE} ذخیره نشد.\n{e}")


# <<<<<<<<<<<<<< START: توابع دیتابیس و اکسل >>>>>>>>>>>>>>
def initialize_database():
    """ساخت فایل دیتابیس و جدول با ستون‌های فارسی اگر وجود نداشته باشد"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # ستون id (شماره قولنامه هم به عنوان فیلد شماره_قولنامه اضافه میشه)
    columns_sql = ["id INTEGER PRIMARY KEY AUTOINCREMENT"]

    # اضافه کردن ستون شماره قولنامه
    columns_sql.append("[شماره_قولنامه] TEXT")

    # ساخت ستون‌ها بر اساس نگاشت فارسی
    created_columns = set()
    for pers_col in DB_COLUMN_MAPPING.values():
        if pers_col not in created_columns:
            columns_sql.append(f"[{pers_col}] TEXT")
            created_columns.add(pers_col)

    create_table_sql = f"CREATE TABLE IF NOT EXISTS contracts ({', '.join(columns_sql)})"

    try:
        cursor.execute(create_table_sql)
        conn.commit()
    except sqlite3.Error as e:
        messagebox.showerror("خطای دیتابیس", f"خطا در ساخت دیتابیس:\n{e}")
    finally:
        conn.close()


def save_data_to_excel(raw_data, serial_no):
    """ذخیره اطلاعات محدود در فایل اکسل"""
    if not HAS_OPENPYXL:
        return

    # فقط هدرهای درخواستی: شماره قولنامه + موارد موجود در دیکشنری
    headers = ["شماره_قولنامه"] + list(DB_COLUMN_MAPPING.values())

    # ساخت ردیف داده
    row_data = [serial_no]
    for eng_key in DB_COLUMN_MAPPING.keys():
        val = raw_data.get(eng_key, "")
        if val is None: val = ""
        row_data.append(str(val))

    try:
        if not os.path.exists(EXCEL_FILE):
            # اگر فایل نیست، یکی بساز و هدرها را بنویس
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Transactions"
            ws.append(headers)
        else:
            # اگر هست، بازش کن
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active

        # افزودن ردیف جدید
        ws.append(row_data)

        wb.save(EXCEL_FILE)
        print(f"Data saved to Excel: {serial_no}")

    except PermissionError:
        messagebox.showerror("خطای اکسل",
                             "فایل اکسل باز است!\nلطفاً فایل transactions.xlsx را ببندید و دوباره تلاش کنید.")
    except Exception as e:
        messagebox.showerror("خطای اکسل", f"خطا در ذخیره فایل اکسل:\n{e}")


def save_data_to_db(raw_data, serial_no):
    """ذخیره اطلاعات محدود در دیتابیس"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    data_to_insert = {}

    # فقط شماره قولنامه
    data_to_insert["شماره_قولنامه"] = serial_no
    # تاریخ سیستم حذف شد

    # نگاشت فیلدهای محدود شده
    for eng_key, pers_key in DB_COLUMN_MAPPING.items():
        # دریافت مقدار از داده‌های خام
        value = raw_data.get(eng_key, "")
        if value is None: value = ""
        data_to_insert[pers_key] = str(value)

    # ساخت کوئری INSERT
    columns = list(data_to_insert.keys())
    values = list(data_to_insert.values())

    cols_str = ", ".join([f"[{c}]" for c in columns])
    q_marks = ", ".join(["?" for _ in columns])

    sql = f"INSERT INTO contracts ({cols_str}) VALUES ({q_marks})"

    try:
        cursor.execute(sql, values)
        conn.commit()
        print(f"Data saved to DB: {serial_no}")
    except sqlite3.Error as e:
        messagebox.showerror("خطای ذخیره دیتابیس", f"اطلاعات در دیتابیس ذخیره نشد:\n{e}")
    finally:
        conn.close()


# <<<<<<<<<<<<<< END: توابع دیتابیس و اکسل >>>>>>>>>>>>>>


# --- 6. تابع اصلی ذخیره ---
current_counter_int = load_counter()


def fill_and_save():
    global current_counter_int
    serial_number = counter_display_var.get()
    try:
        seller_name_raw = entries["seller_name"].get()
        buyer_name_raw = entries["buyer_name"].get()
        car_type_raw = entries["car_type"].get()
    except KeyError as e:
        messagebox.showerror("خطای فیلد", f"فیلد مورد نیاز یافت نشد: {e}")
        return

    s_seller_name = sanitize_filename(seller_name_raw)
    s_buyer_name = sanitize_filename(buyer_name_raw)
    s_car_type = sanitize_filename(car_type_raw)
    if not s_seller_name: s_seller_name = "فروشنده"
    if not s_buyer_name: s_buyer_name = "خریدار"
    if not s_car_type: s_car_type = "خودرو"

    form_data = {}
    special_word_fields = {
        "trans_total_price_words", "trans_biane_words", "trans_baghimandeh_words", "commission_words"
    }

    # جمع‌آوری اطلاعات برای عکس
    for key in COORDS.keys():
        text = ""
        if key == "car_pelak":
            try:
                part1 = entries["pelak_part1"].get()
                part2 = entries["pelak_part2"].get()
                part3 = entries["pelak_part3"].get()
                part4 = entries["pelak_part4"].get()
                if part1 or part2 or part3 or part4:
                    text = f"\u200E{part1} {part2} {part3} - {part4}"
                else:
                    text = ""
                form_data[key] = text
                continue
            except KeyError:
                form_data[key] = ""
                continue

        if key.startswith("vahed_senfi_") or key.startswith("trans_desc_"):
            continue

        if key in string_vars:
            text = string_vars[key].get()
            if key in special_word_fields and text.endswith(" ریال"):
                text = text[:-5]
        elif key in entries:
            widget = entries[key]
            if isinstance(widget, scrolledtext.ScrolledText):
                text = widget.get("1.0", tk.END).strip()
            elif isinstance(widget, tk.Entry):
                text = widget.get()

        form_data[key] = text

    # جمع‌آوری اطلاعات تکمیلی برای دیتابیس و اکسل
    full_db_data = form_data.copy()

    # اضافه کردن توضیحات کامل (نه خط به خط)
    if "trans_desc" in entries:
        full_db_data["trans_desc"] = entries["trans_desc"].get("1.0", tk.END).strip()

    if "vahed_senfi" in entries and isinstance(entries["vahed_senfi"], tk.Entry):
        full_db_data["vahed_senfi"] = entries["vahed_senfi"].get()

    # اضافه کردن مقادیر خام متغیرها
    full_db_data["trans_biane"] = string_vars["trans_biane"].get()
    full_db_data["trans_baghimandeh"] = string_vars["trans_baghimandeh"].get()
    full_db_data["hazine_enteghal"] = string_vars["hazine_enteghal"].get()

    # <<<<<<<< ذخیره در دیتابیس و اکسل >>>>>>>>
    save_data_to_db(full_db_data, serial_number)
    save_data_to_excel(full_db_data, serial_number)
    # <<<<<<<<<<<<<<< >>>>>>>>>>>>>>>

    if not os.path.exists(BASE_IMAGE_PATH):
        messagebox.showerror("خطا", f"فایل عکس فرم یافت نشد: {BASE_IMAGE_PATH}")
        return
    if not os.path.exists(FONT_PATH):
        messagebox.showerror("خطا", f"فایل فونت یافت نشد: {FONT_PATH}")
        return

    try:
        with Image.open(BASE_IMAGE_PATH) as img:
            draw = ImageDraw.Draw(img)

            for key, text in form_data.items():
                if text and key in COORDS:
                    coord = COORDS[key]
                    if coord == (0, 0):
                        continue

                    is_text_persian = is_persian(text)

                    if key == "car_pelak":
                        prepared_text = text
                        current_font = FONT_MAIN
                    else:
                        prepared_text = prepare_text(text)

                        if is_text_persian:
                            if key in ["seller_address", "buyer_address", "car_lavazem",
                                       "trans_total_price_words", "trans_biane_words", "trans_baghimandeh_words",
                                       "commission_words", "seller_mojieb", "buyer_mojieb", "car_owner_mojeb",
                                       "check_number", "check_bank", "check_branch", "check_code",
                                       "doc_date_words", "delivery_date_words", "trans_date_words"]:
                                current_font = FONT_ADDRESS
                            else:
                                current_font = FONT_MAIN
                        else:
                            if key in ["seller_address", "buyer_address", "car_lavazem",
                                       "trans_total_price_words", "trans_biane_words", "trans_baghimandeh_words",
                                       "commission_words", "seller_mojieb", "buyer_mojieb", "car_owner_mojeb",
                                       "check_number", "check_bank", "check_branch", "check_code",
                                       "doc_date_words", "delivery_date_words", "trans_date_words"]:
                                current_font = FONT_ENGLISH_ADDRESS
                            else:
                                current_font = FONT_ENGLISH_MAIN

                    draw.text(coord, prepared_text, fill=MAIN_COLOR, font=current_font, anchor="ra", align="right")

            hazine_shakhs = string_vars["hazine_enteghal"].get()
            if hazine_shakhs and "hazine_enteghal_text" in COORDS:
                coord = COORDS["hazine_enteghal_text"]
                prepared_text = prepare_text(hazine_shakhs)
                draw.text(coord, prepared_text, fill=MAIN_COLOR, font=FONT_ADDRESS, anchor="ra", align="right")

            vahed_senfi_text = ""
            if "vahed_senfi" in entries:
                widget = entries["vahed_senfi"]
                if isinstance(widget, tk.Entry):
                    vahed_senfi_text = widget.get()

            if vahed_senfi_text:
                prepared_text = prepare_text(vahed_senfi_text)
                if "vahed_senfi_1" in COORDS:
                    draw.text(COORDS["vahed_senfi_1"], prepared_text, fill=MAIN_COLOR, font=FONT_ADDRESS, anchor="ra",
                              align="right")
                if "vahed_senfi_2" in COORDS:
                    draw.text(COORDS["vahed_senfi_2"], prepared_text, fill=MAIN_COLOR, font=FONT_ADDRESS, anchor="ra",
                              align="right")

            desc_keys = ["trans_desc_1", "trans_desc_2", "trans_desc_3", "trans_desc_4"]
            full_text = ""
            if "trans_desc" in entries:
                widget = entries["trans_desc"]
                if isinstance(widget, scrolledtext.ScrolledText):
                    full_text = widget.get("1.0", tk.END).strip()

            if full_text:
                wrapped_lines = wrap_text_by_pixel_width(full_text, FONT_ADDRESS, DESC_MAX_WIDTH_PX)
                for i, line_text in enumerate(wrapped_lines):
                    if i >= len(desc_keys): break
                    key = desc_keys[i]
                    if key in COORDS:
                        coord = COORDS[key]
                        if coord != (0, 0):
                            prepared_text = prepare_text(line_text.strip())
                            desc_font = FONT_ADDRESS if is_persian(line_text) else FONT_ENGLISH_ADDRESS
                            if prepared_text:
                                draw.text(coord, prepared_text, fill=MAIN_COLOR, font=desc_font, anchor="ra",
                                          align="right")

            new_folder_name = f"{serial_number} - {s_buyer_name} - {s_seller_name}"
            new_folder_path = new_folder_name
            new_file_name = f"{s_buyer_name} - {s_car_type}.jpg"
            final_output_path = os.path.join(new_folder_path, new_file_name)

            try:
                os.makedirs(new_folder_path, exist_ok=True)
            except OSError as e:
                messagebox.showerror("خطای ساخت فولدر", f"امکان ساخت فولدر مقصد وجود نداشت.\n{e}")
                return

            img.save(final_output_path)

            new_file_name_pdf = f"{s_buyer_name} - {s_car_type}.pdf"
            final_output_path_pdf = os.path.join(new_folder_path, new_file_name_pdf)

            if img.mode != 'RGB':
                img = img.convert('RGB')

            img.save(final_output_path_pdf, "PDF", resolution=100.0, save_all=True)

            if increment_counter_var.get():
                current_counter_int += 1
                save_counter(current_counter_int)
                counter_display_var.set(f"{current_counter_int:04d}")

            messagebox.showinfo("موفقیت",
                                f"فرم پر شد.\nفایل‌ها ذخیره شدند در: {final_output_path}\nدیتابیس و اکسل بروز شدند.")
            img.show()

    except Exception as e:
        messagebox.showerror("خطای کلی", f"یک خطا رخ داد: {e}")


def clear_form():
    if not messagebox.askyesno("پاک کردن فرم", "آیا مطمئن هستید که می‌خواهید تمام اطلاعات فرم را پاک کنید؟"):
        return
    for key, var in string_vars.items():
        if key == "hazine_enteghal":
            var.set("خریدار")
        else:
            var.set("")
    for key, widget in entries.items():
        if isinstance(widget, tk.Entry):
            try:
                widget.delete(0, tk.END)
            except tk.TclError:
                pass
        elif isinstance(widget, scrolledtext.ScrolledText):
            try:
                widget.delete('1.0', tk.END)
            except tk.TclError:
                pass
    calculate_and_update_baghimandeh()
    messagebox.showinfo("موفقیت", "فرم با موفقیت پاک شد.")


# <--- START: توابع ذخیره و بازیابی (JSON) --->
def save_draft_json():
    data = {}
    for key, var in string_vars.items():
        data['VAR_' + key] = var.get()
    for key, widget in entries.items():
        if isinstance(widget, tk.Entry):
            data['ENTRY_' + key] = widget.get()
        elif isinstance(widget, scrolledtext.ScrolledText):
            data['TEXT_' + key] = widget.get("1.0", tk.END)

    filename = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        title="ذخیره موقت فرم"
    )

    if filename:
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("ذخیره موقت", "وضعیت فرم با موفقیت ذخیره شد.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره فایل: {e}")


def load_draft_json():
    filename = filedialog.askopenfilename(
        filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        title="باز کردن فرم ذخیره شده"
    )

    if not filename:
        return

    try:
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)

        for key, val in data.items():
            if key.startswith('VAR_'):
                real_key = key[4:]
                if real_key in string_vars:
                    string_vars[real_key].set(val)
            elif key.startswith('ENTRY_'):
                real_key = key[6:]
                if real_key in entries:
                    widget = entries[real_key]
                    if isinstance(widget, tk.Entry):
                        widget.delete(0, tk.END)
                        widget.insert(0, val)
            elif key.startswith('TEXT_'):
                real_key = key[5:]
                if real_key in entries:
                    widget = entries[real_key]
                    if isinstance(widget, scrolledtext.ScrolledText):
                        widget.delete("1.0", tk.END)
                        widget.insert("1.0", val)

        calculate_and_update_baghimandeh()
        messagebox.showinfo("بارگذاری", "اطلاعات با موفقیت بارگذاری شد.")

    except Exception as e:
        messagebox.showerror("خطا", f"خطا در بارگذاری فایل: {e}")


# <--- END: توابع ذخیره و بازیابی (JSON) --->


# --- 7. ساخت رابط کاربری (GUI) ---
root = tk.Tk()
root.title("نمایشگاه آزاد")

counter_display_var = tk.StringVar(value=f"{current_counter_int:04d}")
increment_counter_var = tk.BooleanVar(value=True)

top_frame = ttk.Frame(root)
top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(5, 0))

counter_label_display = ttk.Label(top_frame, textvariable=counter_display_var, font=('tahoma', 12, 'bold'),
                                  foreground='red')
counter_label_display.pack(side=tk.LEFT, padx=(5, 10))
counter_label_text = ttk.Label(top_frame, text=":شماره سریال", font=('tahoma', 11, 'bold'))
counter_label_text.pack(side=tk.LEFT)
sep_top = ttk.Separator(root, orient='horizontal')
sep_top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(2, 5))

entries = {}
string_vars = {}
notebook = ttk.Notebook(root)
notebook.pack(pady=(0, 10), padx=10, fill="both", expand=True)


def create_field(parent, key, label_text, row, textvariable=None, state=tk.NORMAL):
    tk.Label(parent, text=label_text).grid(row=row, column=1, padx=5, pady=5, sticky="e")
    entry = tk.Entry(parent, width=50, justify='right', textvariable=textvariable, state=state)
    if state == 'readonly':
        entry.configure(readonlybackground="#f0f0f0")
    entry.grid(row=row, column=0, padx=5, pady=2, sticky="ew")
    entries[key] = entry


def create_multiline_field(parent, key, label_text, row, height=3):
    tk.Label(parent, text=label_text).grid(row=row, column=1, padx=5, pady=5, sticky="e")
    entry = scrolledtext.ScrolledText(parent, width=50, height=height, wrap=tk.WORD)
    entry.configure(font=('Tahoma', 10))
    entry.grid(row=row, column=0, padx=5, pady=2, sticky="ew")
    entries[key] = entry


tab_seller = ttk.Frame(notebook, padding="10")
notebook.add(tab_seller, text='     فروشنده     ')
tab_seller.grid_columnconfigure(0, weight=1)
create_field(tab_seller, "seller_name", ":نام و نام خانوادگی", 0)
create_field(tab_seller, "seller_father", ":فرزند", 1)
create_field(tab_seller, "seller_sh_sh", ":شماره شناسنامه/اقتصادی", 2)
create_field(tab_seller, "seller_sodoor", ":صادره از", 3)
create_field(tab_seller, "seller_birth_year", ":تاریخ ثبت/متولد", 4)
create_field(tab_seller, "seller_shenase_melli", ":شناسه ملی", 5)
create_field(tab_seller, "seller_code_melli", ":کد ملی", 6)
create_multiline_field(tab_seller, "seller_address", ":ساکن", 7, height=2)
create_field(tab_seller, "seller_zip_code", ":کد پستی", 8)
create_field(tab_seller, "seller_phone", ":تلفن", 9)
create_field(tab_seller, "seller_owner", ":از طرف /وکالتا/قیوما", 10)
create_field(tab_seller, "seller_mojieb", ":به موجب", 11)

tab_buyer = ttk.Frame(notebook, padding="10")
notebook.add(tab_buyer, text='     خریدار     ')
tab_buyer.grid_columnconfigure(0, weight=1)
create_field(tab_buyer, "buyer_name", ":نام و نام خانوادگی", 0)
create_field(tab_buyer, "buyer_father", ":فرزند", 1)
create_field(tab_buyer, "buyer_sh_sh", ":شماره شناسنامه/اقتصادی", 2)
create_field(tab_buyer, "buyer_sodoor", ":صادره از", 3)
create_field(tab_buyer, "buyer_birth_year", ":تاریخ ثبت/متولد", 4)
create_field(tab_buyer, "buyer_shenase_melli", ":شناسه ملی", 5)
create_field(tab_buyer, "buyer_code_melli", ":کد ملی", 6)
create_multiline_field(tab_buyer, "buyer_address", ":ساکن", 7, height=2)
create_field(tab_buyer, "buyer_zip_code", ":کد پستی", 8)
create_field(tab_buyer, "buyer_phone", ":تلفن", 9)
create_field(tab_buyer, "buyer_owner", ":از طرف/وکالتا/قیوما", 10)
create_field(tab_buyer, "buyer_mojieb", ":به موجب", 11)

tab_car = ttk.Frame(notebook, padding="10")
notebook.add(tab_car, text='     اتومبیل     ')
tab_car.grid_columnconfigure(0, weight=1)
create_field(tab_car, "car_dang", ":انتقال (دانگ)", 0)
create_field(tab_car, "car_type", ":نوع اتومبیل", 1)
create_field(tab_car, "car_system", ":سیستم", 2)
create_field(tab_car, "car_tip", ":تیپ", 3)
create_field(tab_car, "car_model", ":مدل", 4)
create_field(tab_car, "car_color", ":رنگ", 5)
create_field(tab_car, "car_capacity", ":ظرفیت", 6)
create_field(tab_car, "car_cylinders", ":تعداد سیلندر", 7)
create_field(tab_car, "car_fuel", ":نوع سوخت", 8)
create_field(tab_car, "car_axles", ":تعداد محور", 9)
create_field(tab_car, "car_wheels", ":تعداد چرخ", 10)
create_field(tab_car, "car_motor_no", ":شماره موتور", 11)
create_field(tab_car, "car_shasi_no", ":شماره شاسی", 12)

tk.Label(tab_car, text=":شماره انتظامی ناجا").grid(row=13, column=1, padx=5, pady=5, sticky="e")
pelak_frame = ttk.Frame(tab_car)
pelak_frame.grid(row=13, column=0, padx=5, pady=2, sticky="ew")
entries["pelak_part4"] = tk.Entry(pelak_frame, width=4, justify='center', font=('tahoma', 10, 'bold'))
entries["pelak_part4"].pack(side=tk.RIGHT, padx=(0, 2))
tk.Label(pelak_frame, text="-", font=('tahoma', 10, 'bold')).pack(side=tk.RIGHT, padx=2)
entries["pelak_part3"] = tk.Entry(pelak_frame, width=5, justify='center', font=('tahoma', 10, 'bold'))
entries["pelak_part3"].pack(side=tk.RIGHT, padx=2)
entries["pelak_part2"] = tk.Entry(pelak_frame, width=4, justify='center', font=('tahoma', 10, 'bold'))
entries["pelak_part2"].pack(side=tk.RIGHT, padx=2)
entries["pelak_part1"] = tk.Entry(pelak_frame, width=4, justify='center', font=('tahoma', 10, 'bold'))
entries["pelak_part1"].pack(side=tk.RIGHT, padx=2)

create_field(tab_car, "car_owner_mojeb", ":مالکیت فروشنده به موجب", 14)
create_multiline_field(tab_car, "car_lavazem", ":لوازم/مدارک", 115, height=2)

tab_trans = ttk.Frame(notebook, padding="10")
notebook.add(tab_trans, text='     معامله     ')
canvas = tk.Canvas(tab_trans, borderwidth=0, highlightthickness=0)
scrollbar = ttk.Scrollbar(tab_trans, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
tab_trans.grid_rowconfigure(0, weight=1)
tab_trans.grid_columnconfigure(0, weight=1)
canvas.grid(row=0, column=0, sticky="nsew")
scrollbar.grid(row=0, column=1, sticky="ns")


def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))


def on_canvas_configure(event):
    canvas.itemconfig(scrollable_frame_window, width=event.width)


canvas.bind("<Configure>", on_canvas_configure)
scrollable_frame.bind("<Configure>", on_frame_configure)
scrollable_frame.grid_columnconfigure(0, weight=1)


def _on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def _on_mousewheel_linux(event):
    if event.num == 4:
        canvas.yview_scroll(-1, "units")
    elif event.num == 5:
        canvas.yview_scroll(1, "units")


def _bind_mouse_to_canvas(event=None):
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", _on_mousewheel_linux)
    canvas.bind_all("<Button-5>", _on_mousewheel_linux)


def _unbind_mouse_from_canvas(event=None):
    canvas.unbind_all("<MouseWheel>")
    canvas.unbind_all("<Button-4>")
    canvas.unbind_all("<Button-5>")


canvas.bind('<Enter>', _bind_mouse_to_canvas)
scrollable_frame.bind('<Enter>', _bind_mouse_to_canvas)
canvas.bind('<Leave>', _unbind_mouse_from_canvas)
scrollable_frame.bind('<Leave>', _unbind_mouse_from_canvas)

string_vars["trans_total_price"] = tk.StringVar()
string_vars["trans_total_price_words"] = tk.StringVar()
string_vars["trans_biane"] = tk.StringVar()
string_vars["trans_biane_words"] = tk.StringVar()
string_vars["trans_baghimandeh"] = tk.StringVar()
string_vars["trans_baghimandeh_words"] = tk.StringVar()
string_vars["doc_day"] = tk.StringVar()
string_vars["doc_month"] = tk.StringVar()
string_vars["doc_year_xx"] = tk.StringVar()
string_vars["doc_date_words"] = tk.StringVar()
string_vars["delivery_day"] = tk.StringVar()
string_vars["delivery_month"] = tk.StringVar()
string_vars["delivery_year_xx"] = tk.StringVar()
string_vars["delivery_time"] = tk.StringVar()
string_vars["delivery_date_words"] = tk.StringVar()
string_vars["after_doc_day"] = tk.StringVar()
string_vars["after_doc_month"] = tk.StringVar()
string_vars["after_doc_year_xx"] = tk.StringVar()
string_vars["hazine_enteghal"] = tk.StringVar(value="خریدار")
string_vars["delay_cost_doc"] = tk.StringVar()
string_vars["delay_cost_words"] = tk.StringVar()
string_vars["daily_penalty"] = tk.StringVar()
string_vars["daily_penalty_words"] = tk.StringVar()
string_vars["commission_amount"] = tk.StringVar()
string_vars["commission_words"] = tk.StringVar()
string_vars["commission_pay_day"] = tk.StringVar()
string_vars["commission_pay_month"] = tk.StringVar()
string_vars["commission_pay_year_xx"] = tk.StringVar()
string_vars["trans_date_day"] = tk.StringVar()
string_vars["trans_date_month"] = tk.StringVar()
string_vars["trans_date_year_xx"] = tk.StringVar()
string_vars["trans_time"] = tk.StringVar()
string_vars["trans_date_words"] = tk.StringVar()

string_vars["trans_total_price"].trace_add("write", lambda *a: on_generic_price_change(string_vars["trans_total_price"],
                                                                                       string_vars[
                                                                                           "trans_total_price_words"],
                                                                                       "trans_total_price", True))
string_vars["trans_biane"].trace_add("write", lambda *a: on_generic_price_change(string_vars["trans_biane"],
                                                                                 string_vars["trans_biane_words"],
                                                                                 "trans_biane", True))
string_vars["delay_cost_doc"].trace_add("write", lambda *a: on_generic_price_change(string_vars["delay_cost_doc"],
                                                                                    string_vars["delay_cost_words"],
                                                                                    "delay_cost_doc", False))
string_vars["daily_penalty"].trace_add("write", lambda *a: on_generic_price_change(string_vars["daily_penalty"],
                                                                                   string_vars["daily_penalty_words"],
                                                                                   "daily_penalty", False))
string_vars["commission_amount"].trace_add("write", lambda *a: on_generic_price_change(string_vars["commission_amount"],
                                                                                       string_vars["commission_words"],
                                                                                       "commission_amount", False))

doc_date_lambda = lambda *a: on_generic_date_change("doc_date", string_vars["doc_day"], string_vars["doc_month"],
                                                    string_vars["doc_year_xx"], string_vars["doc_date_words"])
string_vars["doc_day"].trace_add("write", doc_date_lambda)
string_vars["doc_month"].trace_add("write", doc_date_lambda)
string_vars["doc_year_xx"].trace_add("write", doc_date_lambda)

delivery_date_lambda = lambda *a: on_generic_date_change("delivery_date", string_vars["delivery_day"],
                                                         string_vars["delivery_month"], string_vars["delivery_year_xx"],
                                                         string_vars["delivery_date_words"])
string_vars["delivery_day"].trace_add("write", delivery_date_lambda)
string_vars["delivery_month"].trace_add("write", delivery_date_lambda)
string_vars["delivery_year_xx"].trace_add("write", delivery_date_lambda)

after_doc_date_lambda = lambda *a: on_generic_date_change("after_doc_date", string_vars["after_doc_day"],
                                                          string_vars["after_doc_month"],
                                                          string_vars["after_doc_year_xx"], None)
string_vars["after_doc_day"].trace_add("write", after_doc_date_lambda)
string_vars["after_doc_month"].trace_add("write", after_doc_date_lambda)
string_vars["after_doc_year_xx"].trace_add("write", after_doc_date_lambda)

commission_date_lambda = lambda *a: on_generic_date_change("commission_date", string_vars["commission_pay_day"],
                                                           string_vars["commission_pay_month"],
                                                           string_vars["commission_pay_year_xx"], None)
string_vars["commission_pay_day"].trace_add("write", commission_date_lambda)
string_vars["commission_pay_month"].trace_add("write", commission_date_lambda)
string_vars["commission_pay_year_xx"].trace_add("write", commission_date_lambda)

trans_date_words_lambda = lambda *a: on_generic_date_change("trans_date_words", string_vars["trans_date_day"],
                                                            string_vars["trans_date_month"],
                                                            string_vars["trans_date_year_xx"],
                                                            string_vars["trans_date_words"])
string_vars["trans_date_day"].trace_add("write", trans_date_words_lambda)
string_vars["trans_date_month"].trace_add("write", trans_date_words_lambda)
string_vars["trans_date_year_xx"].trace_add("write", trans_date_words_lambda)

create_field(scrollable_frame, "trans_total_price", ":قیمت کل", 0, textvariable=string_vars["trans_total_price"])
create_field(scrollable_frame, "trans_total_price_words", "", 1, textvariable=string_vars["trans_total_price_words"],
             state='readonly')
create_field(scrollable_frame, "trans_biane", ":بیعانه", 2, textvariable=string_vars["trans_biane"])
create_field(scrollable_frame, "trans_biane_words", "", 3, textvariable=string_vars["trans_biane_words"],
             state='readonly')

sep1 = ttk.Separator(scrollable_frame, orient='horizontal')
sep1.grid(row=4, column=0, columnspan=2, sticky='ew', pady=(10, 5))
create_field(scrollable_frame, "check_number", ":یا طی چک شماره", 5)
create_field(scrollable_frame, "check_bank", ":عهده بانک", 6)
create_field(scrollable_frame, "check_branch", ":شعبه", 7)
create_field(scrollable_frame, "check_code", ":کد", 8)
sep2 = ttk.Separator(scrollable_frame, orient='horizontal')
sep2.grid(row=9, column=0, columnspan=2, sticky='ew', pady=(5, 10))

create_field(scrollable_frame, "trans_baghimandeh", ":باقی مانده", 10, textvariable=string_vars["trans_baghimandeh"],
             state='readonly')
create_field(scrollable_frame, "trans_baghimandeh_words", "", 11, textvariable=string_vars["trans_baghimandeh_words"],
             state='readonly')

date_frame = tk.Frame(scrollable_frame)
date_frame.grid(row=12, column=0, columnspan=2, pady=(5, 0), sticky="e")
tk.Label(date_frame, text=":وکالت تا تاریخ").pack(side=tk.RIGHT, padx=5)
entry_day = tk.Entry(date_frame, width=3, justify='right', textvariable=string_vars["doc_day"])
entry_day.pack(side=tk.RIGHT)
entries["doc_day"] = entry_day
tk.Label(date_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_month = tk.Entry(date_frame, width=3, justify='right', textvariable=string_vars["doc_month"])
entry_month.pack(side=tk.RIGHT)
entries["doc_month"] = entry_month
tk.Label(date_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_year_xx = tk.Entry(date_frame, width=3, justify='right', textvariable=string_vars["doc_year_xx"])
entry_year_xx.pack(side=tk.RIGHT)
entries["doc_year_xx"] = entry_year_xx
tk.Label(date_frame, text="14").pack(side=tk.RIGHT, padx=(2, 0))

create_field(scrollable_frame, "doc_date_words", "", 13, textvariable=string_vars["doc_date_words"], state='readonly')

delivery_frame = tk.Frame(scrollable_frame)
delivery_frame.grid(row=14, column=0, columnspan=2, pady=(5, 0), sticky="e")
tk.Label(delivery_frame, text=":ساعت").pack(side=tk.RIGHT, padx=(5, 0))
entry_time = tk.Entry(delivery_frame, width=5, justify='right', textvariable=string_vars["delivery_time"])
entry_time.pack(side=tk.RIGHT)
entries["delivery_time"] = entry_time
tk.Label(delivery_frame, text=":تاریخ تحویل").pack(side=tk.RIGHT, padx=(10, 5))
entry_delivery_day = tk.Entry(delivery_frame, width=3, justify='right', textvariable=string_vars["delivery_day"])
entry_delivery_day.pack(side=tk.RIGHT)
entries["delivery_day"] = entry_delivery_day
tk.Label(delivery_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_delivery_month = tk.Entry(delivery_frame, width=3, justify='right', textvariable=string_vars["delivery_month"])
entry_delivery_month.pack(side=tk.RIGHT)
entries["delivery_month"] = entry_delivery_month
tk.Label(delivery_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_delivery_year_xx = tk.Entry(delivery_frame, width=3, justify='right',
                                  textvariable=string_vars["delivery_year_xx"])
entry_delivery_year_xx.pack(side=tk.RIGHT)
entries["delivery_year_xx"] = entry_delivery_year_xx
tk.Label(delivery_frame, text="14").pack(side=tk.RIGHT, padx=(2, 0))

create_field(scrollable_frame, "delivery_date_words", "", 15, textvariable=string_vars["delivery_date_words"],
             state='readonly')

radio_frame = tk.Frame(scrollable_frame)
radio_frame.grid(row=16, column=0, columnspan=2, pady=5, sticky="e")
tk.Label(radio_frame, text=":هزینه وکالت سند").pack(side=tk.RIGHT, padx=5)
hazine_var = string_vars["hazine_enteghal"]
entries["hazine_enteghal"] = hazine_var
tk.Radiobutton(radio_frame, text="فروشنده", variable=hazine_var, value="فروشنده").pack(side=tk.RIGHT)
tk.Radiobutton(radio_frame, text="خریدار", variable=hazine_var, value="خریدار").pack(side=tk.RIGHT)

after_doc_delivery_frame = tk.Frame(scrollable_frame)
after_doc_delivery_frame.grid(row=17, column=0, columnspan=2, pady=(5, 0), sticky="e")
tk.Label(after_doc_delivery_frame, text=":تاریخ تحویل پس از سند (اختیاری)").pack(side=tk.RIGHT, padx=5)
entry_after_doc_day = tk.Entry(after_doc_delivery_frame, width=3, justify='right',
                               textvariable=string_vars["after_doc_day"])
entry_after_doc_day.pack(side=tk.RIGHT)
entries["after_doc_day"] = entry_after_doc_day
tk.Label(after_doc_delivery_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_after_doc_month = tk.Entry(after_doc_delivery_frame, width=3, justify='right',
                                 textvariable=string_vars["after_doc_month"])
entry_after_doc_month.pack(side=tk.RIGHT)
entries["after_doc_month"] = entry_after_doc_month
tk.Label(after_doc_delivery_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_after_doc_year_xx = tk.Entry(after_doc_delivery_frame, width=3, justify='right',
                                   textvariable=string_vars["after_doc_year_xx"])
entry_after_doc_year_xx.pack(side=tk.RIGHT)
entries["after_doc_year_xx"] = entry_after_doc_year_xx
tk.Label(after_doc_delivery_frame, text="14").pack(side=tk.RIGHT, padx=(2, 0))

create_field(scrollable_frame, "delay_cost_doc", ":هزینه دیرکرد از تاریخ وکالت", 18,
             textvariable=string_vars["delay_cost_doc"])
create_field(scrollable_frame, "delay_cost_words", "", 19, textvariable=string_vars["delay_cost_words"],
             state='readonly')
create_field(scrollable_frame, "daily_penalty", ":خسارت روزانه", 20, textvariable=string_vars["daily_penalty"])
create_field(scrollable_frame, "daily_penalty_words", "", 21, textvariable=string_vars["daily_penalty_words"],
             state='readonly')

create_field(scrollable_frame, "commission_amount", ":کمیسیون", 22, textvariable=string_vars["commission_amount"])
create_field(scrollable_frame, "commission_words", "", 23, textvariable=string_vars["commission_words"],
             state='readonly')

create_field(scrollable_frame, "vahed_senfi", ":واحد صنفی", 24)

commission_date_frame = tk.Frame(scrollable_frame)
commission_date_frame.grid(row=25, column=0, columnspan=2, pady=(5, 0), sticky="e")
tk.Label(commission_date_frame, text=":تاریخ پرداخت کمیسیون").pack(side=tk.RIGHT, padx=5)
entry_comm_day = tk.Entry(commission_date_frame, width=3, justify='right',
                          textvariable=string_vars["commission_pay_day"])
entry_comm_day.pack(side=tk.RIGHT)
entries["commission_pay_day"] = entry_comm_day
tk.Label(commission_date_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_comm_month = tk.Entry(commission_date_frame, width=3, justify='right',
                            textvariable=string_vars["commission_pay_month"])
entry_comm_month.pack(side=tk.RIGHT)
entries["commission_pay_month"] = entry_comm_month
tk.Label(commission_date_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_comm_year_xx = tk.Entry(commission_date_frame, width=3, justify='right',
                              textvariable=string_vars["commission_pay_year_xx"])
entry_comm_year_xx.pack(side=tk.RIGHT)
entries["commission_pay_year_xx"] = entry_comm_year_xx
tk.Label(commission_date_frame, text="14").pack(side=tk.RIGHT, padx=(2, 0))

trans_date_time_frame = tk.Frame(scrollable_frame)
trans_date_time_frame.grid(row=26, column=0, columnspan=2, pady=(5, 0), sticky="e")

tk.Label(trans_date_time_frame, text=":ساعت").pack(side=tk.RIGHT, padx=(5, 0))
entry_trans_time = tk.Entry(trans_date_time_frame, width=5, justify='right', textvariable=string_vars["trans_time"])
entry_trans_time.pack(side=tk.RIGHT)
entries["trans_time"] = entry_trans_time

tk.Label(trans_date_time_frame, text=":تاریخ معامله").pack(side=tk.RIGHT, padx=(10, 5))
entry_trans_day = tk.Entry(trans_date_time_frame, width=3, justify='right', textvariable=string_vars["trans_date_day"])
entry_trans_day.pack(side=tk.RIGHT)
entries["trans_date_day"] = entry_trans_day
tk.Label(trans_date_time_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_trans_month = tk.Entry(trans_date_time_frame, width=3, justify='right',
                             textvariable=string_vars["trans_date_month"])
entry_trans_month.pack(side=tk.RIGHT)
entries["trans_date_month"] = entry_trans_month
tk.Label(trans_date_time_frame, text="/").pack(side=tk.RIGHT, padx=2)
entry_trans_year_xx = tk.Entry(trans_date_time_frame, width=3, justify='right',
                               textvariable=string_vars["trans_date_year_xx"])
entry_trans_year_xx.pack(side=tk.RIGHT)
entries["trans_date_year_xx"] = entry_trans_year_xx
tk.Label(trans_date_time_frame, text="14").pack(side=tk.RIGHT, padx=(2, 0))

create_field(scrollable_frame, "trans_date_words", "", 27, textvariable=string_vars["trans_date_words"],
             state='readonly')
create_multiline_field(scrollable_frame, "trans_desc", ":توضیحات", 28, height=5)

check_frame = ttk.Frame(scrollable_frame)
check_frame.grid(row=29, column=0, columnspan=2, pady=(10, 5), sticky="e")
cb_increment = ttk.Checkbutton(check_frame, text="افزایش شماره سریال پس از ذخیره", variable=increment_counter_var)
cb_increment.pack(side=tk.RIGHT, padx=5)

# <--- START: دکمه‌ها --->
button_frame = tk.Frame(root)
button_frame.pack(pady=15, padx=10, fill=tk.X)

clear_button = tk.Button(button_frame, text="پاک کردن فرم", command=clear_form, font=('tahoma', 12, 'bold'),
                         bg="#ffcccc", fg="red")
clear_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))

load_draft_button = tk.Button(button_frame, text="باز کردن فرم", command=load_draft_json, font=('tahoma', 12, 'bold'),
                              bg="#e0e0e0", fg="black")
load_draft_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

save_draft_button = tk.Button(button_frame, text="ذخیره موقت", command=save_draft_json, font=('tahoma', 12, 'bold'),
                              bg="#e0e0e0", fg="black")
save_draft_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

save_button = tk.Button(button_frame, text="ذخیره و پر کردن فرم", command=fill_and_save, font=('tahoma', 12, 'bold'),
                        bg="#d0f0c0", fg="black")
save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))
# <--- END: دکمه‌ها --->

initialize_database()
calculate_and_update_baghimandeh()
root.mainloop()