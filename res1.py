
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import openpyxl
from PIL import Image, ImageTk
import os
import shutil
from datetime import datetime
import sqlite3
import subprocess
import sys
import re
from tkcalendar import DateEntry
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl import Workbook, load_workbook
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display
import sqlite3
import sys
import os
def close_splash():
    """إغلاق شاشة التحميل (تعمل فقط في EXE المبني بـ --splash)"""
    try:
        import pyi_splash
        if pyi_splash.is_alive():
            pyi_splash.close()
    except ImportError:
        pass

def resource_path(relative_path):
    """الحصول على المسار الصحيح للملفات داخل EXE أو وضع التطوير"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)



def migrate_existing_treasury_paids():
    """ترحيل المصاريف القديمة بطريقة 'الخزينة' إلى treasury_actions"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        c.execute("""
            SELECT p.id, p.date, p.main_name, p.sub_name,
                   p.value, p.description
            FROM paids p
            LEFT JOIN treasury_actions t ON t.paid_id = p.id
            WHERE p.payment_method = 'الخزينة'
              AND t.id IS NULL
              AND p.value > 0
        """)
        rows = c.fetchall()

        count = 0
        for paid_id, date, main_name, sub_name, value, description in rows:
            action_type, beneficiary = classify_treasury_action(
                main_name, sub_name)

            c.execute('''INSERT INTO treasury_actions
                (date, action_type, beneficiary, amount,
                 description, paid_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)''',
                (date, action_type, beneficiary, value,
                 description or f'مصروف من الخزينة #{paid_id}',
                 paid_id,
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            count += 1

        conn.commit()
        conn.close()
        return count
    except Exception as e:
        print(f"❌ خطأ في migrate_existing_treasury_paids: {e}")
        return 0

 #################################################

# اسم المطعم/الشركة (يظهر في التقارير)
RESTAURANT_NAME = "مطعم الحديقة"

# ???????????????????????????????????????????????????????????????????
#                       AUTOCOMPLETE COMBOBOX
# ??????????????????????????????????????????????????????????????????
class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._full_values = []
        self.bind('<KeyRelease>', self._on_keyrelease)

    def set_values(self, values):
        self._full_values = list(values)
        self['values'] = self._full_values

    def _on_keyrelease(self, event):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'Escape'):
            return
        text = self.get().strip()
        if not text:
            self['values'] = self._full_values
            return
        matches = [v for v in self._full_values if text.lower() in v.lower()]
        self['values'] = matches
        if matches:
            self.event_generate('<Down>')


def calc_trader_balance(conn, trader_name):
    """رصيد التاجر = فواتير - paids عامة - treasury_actions"""
    c = conn.cursor()

    # 1. المتبقي من الفواتير
    c.execute("""SELECT COALESCE(SUM(total - COALESCE(paid_amount, 0)), 0)
                 FROM invoices WHERE supplier = ?""", (trader_name,))
    inv_rem = c.fetchone()[0] or 0

    # 2. الدفعات العامة من paids (بدون ربط بفاتورة)
    # ⚠️ نستثني 'الخزينة' لأنها محسوبة في treasury_actions
    c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                 WHERE sub_name = ?
                 AND (invoice_id IS NULL OR invoice_id = 0)
                 AND COALESCE(payment_method, 'نقدي') != 'الخزينة'""",
              (trader_name,))
    from_paids = c.fetchone()[0] or 0

    # 3. سدادات الخزينة (تشمل المُسجَّل من صفحة المصاريف تلقائياً)
    c.execute("""SELECT COALESCE(SUM(amount), 0) FROM treasury_actions
                 WHERE action_type = 'سداد تاجر' AND beneficiary = ?""",
              (trader_name,))
    from_treasury = c.fetchone()[0] or 0

    balance = inv_rem - from_paids - from_treasury
    return max(0, balance)



# بعد
if hasattr(sys, '_MEIPASS'):
    DB_NAME = os.path.join(os.path.dirname(sys.executable), "restaurant.db")
else:
    DB_NAME = "restaurant.db"



def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # 1. المبيعات
    c.execute('''CREATE TABLE IF NOT EXISTS sales (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    pound INTEGER DEFAULT 0,
                    vodafone INTEGER DEFAULT 0,
                    visa INTEGER DEFAULT 0,
                    dollar INTEGER DEFAULT 0,
                    euro INTEGER DEFAULT 0,
                    dollar_rate REAL DEFAULT 48.0,
                    euro_rate REAL DEFAULT 51.0,
                    image_path TEXT)''')

    # 2. التبويبات الرئيسية
    c.execute('''CREATE TABLE IF NOT EXISTS main_categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL)''')

    # 3. التبويبات الفرعية
    c.execute('''CREATE TABLE IF NOT EXISTS sub_categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    main_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    FOREIGN KEY (main_id) REFERENCES main_categories (id) ON DELETE CASCADE,
                    UNIQUE(main_id, name))''')

    # 4. المصروفات
    # 4. المصروفات (نسخة نظيفة)
    c.execute('''CREATE TABLE IF NOT EXISTS paids (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    value REAL NOT NULL DEFAULT 0,
                    main_name TEXT,
                    sub_name TEXT,
                    description TEXT,
                    payment_method TEXT DEFAULT 'نقدي',
                    image_path TEXT,
                    invoice_id INTEGER)''')

    # 5. الفواتير
    c.execute('''CREATE TABLE IF NOT EXISTS invoices (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    invoice_no TEXT,
                    date TEXT NOT NULL,
                    supplier TEXT,
                    notes TEXT,
                    total REAL DEFAULT 0,
                    created_at TEXT,
                    trader_id INTEGER,
                    payment_type TEXT DEFAULT 'أجل',
                    paid_amount REAL DEFAULT 0)''')

    # 6. بنود الفاتورة
    c.execute('''CREATE TABLE IF NOT EXISTS invoice_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    invoice_id INTEGER NOT NULL,
                    item_id INTEGER,
                    item_name TEXT NOT NULL,
                    quantity REAL NOT NULL,
                    price REAL NOT NULL,
                    total REAL NOT NULL,
                    FOREIGN KEY (invoice_id) REFERENCES invoices (id) ON DELETE CASCADE)''')

    # 7. الأصناف
    c.execute('''CREATE TABLE IF NOT EXISTS items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    unit TEXT,
                    opening_balance REAL DEFAULT 0,
                    closing_balance REAL DEFAULT 0,
                    created_at TEXT)''')

    # 8. أرصدة المخزن الشهرية
    c.execute('''CREATE TABLE IF NOT EXISTS inventory_balances (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_id INTEGER NOT NULL,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    opening_balance REAL DEFAULT 0,
                    closing_balance REAL DEFAULT 0,
                    created_at TEXT,
                    updated_at TEXT,
                    UNIQUE(item_id, year, month),
                    FOREIGN KEY (item_id) REFERENCES items (id) ON DELETE CASCADE)''')

    # 9. التجار
    c.execute('''CREATE TABLE IF NOT EXISTS traders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    phone TEXT,
                    notes TEXT,
                    created_at TEXT)''')

    # 10. مدفوعات التجار
    c.execute('''CREATE TABLE IF NOT EXISTS trader_payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    trader_id INTEGER NOT NULL,
                    invoice_id INTEGER,
                    date TEXT NOT NULL,
                    amount REAL NOT NULL,
                    payment_method TEXT,
                    notes TEXT,
                    created_at TEXT,
                    FOREIGN KEY (trader_id) REFERENCES traders (id) ON DELETE CASCADE)''')

    # 11. العاملون
    c.execute('''CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    phone TEXT,
                    position TEXT,
                    daily_wage REAL DEFAULT 0,
                    first_work_date TEXT,
                    vacations INTEGER DEFAULT 0,
                    deductions REAL DEFAULT 0,
                    bonus REAL DEFAULT 0,
                    status TEXT DEFAULT 'نشط',
                    notes TEXT,
                    created_at TEXT)''')

    # 12. مدد عمل العاملين
    c.execute('''CREATE TABLE IF NOT EXISTS employee_periods (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_id INTEGER NOT NULL,
                    first_date TEXT,
                    last_date TEXT,
                    vacations INTEGER DEFAULT 0,
                    deductions REAL DEFAULT 0,
                    bonus REAL DEFAULT 0,
                    notes TEXT,
                    created_at TEXT,
                    FOREIGN KEY (employee_id) REFERENCES employees (id) ON DELETE CASCADE)''')

    # 13. حركات الخزينة
    c.execute('''CREATE TABLE IF NOT EXISTS treasury_actions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    action_type TEXT NOT NULL,
                    beneficiary TEXT,
                    amount REAL NOT NULL DEFAULT 0,
                    euro REAL DEFAULT 0,
                    usd REAL DEFAULT 0,
                    description TEXT,
                    paid_id INTEGER,
                    created_at TEXT)''')

    # 14. حركات محفظة فودافون كاش
    c.execute('''CREATE TABLE IF NOT EXISTS vodafone_wallet_actions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    action_type TEXT NOT NULL,
                    beneficiary TEXT,
                    amount REAL NOT NULL DEFAULT 0,
                    description TEXT,
                    paid_id INTEGER,
                    created_at TEXT)''')

    conn.commit()

    # التبويبات الرئيسية الثابتة
    c.execute("INSERT OR IGNORE INTO main_categories (name) VALUES ('مشتريات')")
    c.execute("INSERT OR IGNORE INTO main_categories (name) VALUES ('أجور ومرتبات')")
    conn.commit()

    conn.close()
    print(f"✅ تم إنشاء قاعدة البيانات: {DB_NAME} — 13 جدول")



def migrate_add_sub_name2():
    """إضافة عمود sub_name2 لجدول paids (لمرة واحدة)"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("PRAGMA table_info(paids)")
        cols = [r[1] for r in c.fetchall()]
        if 'sub_name2' not in cols:
            c.execute("ALTER TABLE paids ADD COLUMN sub_name2 TEXT")
            conn.commit()
            print("✅ تم إضافة عمود sub_name2 إلى جدول paids")
        conn.close()
    except Exception as e:
        print(f"⚠️ migration error: {e}")




def export_statement_to_image(file_path, title, subtitle, summary_items,
                                net_label, net_value, movements,
                                movements_title="الحركات التفصيلية",
                                page_width=1240, page_height=1754,
                                logo_path=None):
    """إنشاء صورة PNG لكشف الحساب مع دعم العربية والشعار"""
    try:
        from PIL import Image, ImageDraw, ImageFont
        import arabic_reshaper
        from bidi.algorithm import get_display
    except ImportError as e:
        msg = (f"مكتبة مفقودة: {e}\n\n"
               "افتح Command Prompt ونفذ:\n"
               "pip install arabic-reshaper python-bidi Pillow")
        messagebox.showerror("مكتبات مطلوبة", msg)
        return False

    # البحث عن خط عربي
    font_candidates = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
    ]
    font_path = None
    for fp in font_candidates:
        if os.path.exists(fp):
            font_path = fp
            break

    if not font_path:
        messagebox.showerror("خطأ", "لم يتم العثور على خط عربي")
        return False

    try:
        font_title = ImageFont.truetype(font_path, 42)
        font_subtitle = ImageFont.truetype(font_path, 26)
        font_section = ImageFont.truetype(font_path, 32)
        font_label = ImageFont.truetype(font_path, 26)
        font_value = ImageFont.truetype(font_path, 26)
        font_net = ImageFont.truetype(font_path, 38)
        font_table_header = ImageFont.truetype(font_path, 24)
        font_table_body = ImageFont.truetype(font_path, 22)
        font_footer = ImageFont.truetype(font_path, 18)
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل تحميل الخط: {e}")
        return False

    def ar(text):
        if text is None:
            return ''
        text = str(text)
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)

    # ??????????? تجهيز الشعار ???????????
    logo_img = None
    logo_size = 100  # حجم الشعار (مربع)
    if logo_path and os.path.exists(logo_path):
        try:
            logo_img = Image.open(logo_path)
            # تحويل الصورة إلى RGBA للحفاظ على الشفافية
            if logo_img.mode != 'RGBA':
                logo_img = logo_img.convert('RGBA')
            # تصغير الشعار
            logo_img.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)
        except Exception as e:
            print(f"?? فشل تحميل الشعار: {e}")
            logo_img = None

    img = Image.new('RGB', (page_width, page_height), 'white')
    draw = ImageDraw.Draw(img)

    COLOR_HEADER_BG = (44, 62, 80)
    COLOR_HEADER_FG = (255, 255, 255)
    COLOR_SUBTITLE = (241, 196, 15)
    COLOR_SECTION_BG = (139, 69, 19)
    COLOR_SECTION_FG = (255, 255, 255)
    COLOR_LABEL = (52, 73, 94)
    COLOR_VALUE = (44, 62, 80)
    COLOR_GREEN = (39, 174, 96)
    COLOR_RED = (231, 76, 60)
    COLOR_BORDER = (189, 195, 199)
    COLOR_ROW_ALT = (248, 249, 250)
    COLOR_TABLE_HEADER_BG = (52, 73, 94)

    margin = 60
    y = 0

    # ??????????? الهيدر ???????????
    header_height = 130
    draw.rectangle([(0, 0), (page_width, header_height)], fill=COLOR_HEADER_BG)

    # الشعار على اليمين (قبل العنوان بقليل)
    if logo_img:
        logo_x = page_width - logo_size - 30
        logo_y = (header_height - logo_size) // 2
        # لصق الشعار مع احترام الشفافية
        if logo_img.mode == 'RGBA':
            img.paste(logo_img, (logo_x, logo_y), logo_img)
        else:
            img.paste(logo_img, (logo_x, logo_y))

    # الشعار على اليسار أيضاً (اختياري - إذا أردت الشعار في الجهتين)
    # if logo_img:
    #     img.paste(logo_img, (30, (header_height - logo_size) // 2), logo_img if logo_img.mode == 'RGBA' else None)

    # العنوان الرئيسي (وسط - يتم إزاحته تلقائياً إذا كان الشعار كبيراً)
    title_text = ar(title)
    bbox = draw.textbbox((0, 0), title_text, font=font_title)
    title_w = bbox[2] - bbox[0]

    if logo_img:
        # إزاحة العنوان لليسار بمقدار نصف عرض الشعار
        center_x = (page_width - logo_size - 60) / 2 + 30
    else:
        center_x = page_width / 2

    draw.text((center_x - title_w / 2, 25), title_text,
              fill=COLOR_HEADER_FG, font=font_title)

    if subtitle:
        sub_text = ar(subtitle)
        bbox = draw.textbbox((0, 0), sub_text, font=font_subtitle)
        sub_w = bbox[2] - bbox[0]
        draw.text((center_x - sub_w / 2, 85), sub_text,
                  fill=COLOR_SUBTITLE, font=font_subtitle)

    y = header_height + 30

    # ملخص الحساب
    section_h = 50
    draw.rectangle([(margin, y), (page_width - margin, y + section_h)],
                   fill=COLOR_SECTION_BG)
    section_text = ar("ملخص الحساب")
    bbox = draw.textbbox((0, 0), section_text, font=font_section)
    sec_w = bbox[2] - bbox[0]
    draw.text(((page_width - sec_w) / 2, y + 8), section_text,
              fill=COLOR_SECTION_FG, font=font_section)
    y += section_h + 20

    # صفوف الملخص
    row_h = 45
    for label, value in summary_items:
        draw.rectangle([(margin, y), (page_width - margin, y + row_h)],
                       fill=(250, 250, 250), outline=COLOR_BORDER)
        label_text = ar(str(label))
        draw.text((page_width - margin - 20, y + 10), label_text,
                  fill=COLOR_LABEL, font=font_label, anchor='ra')
        value_text = ar(str(value))
        draw.text((margin + 20, y + 10), value_text,
                  fill=COLOR_VALUE, font=font_value, anchor='la')
        y += row_h

    # صافي المستحق
    y += 15
    net_h = 70
    try:
        net_float = float(str(net_value).replace(' ج.م', '').replace(',', ''))
    except:
        net_float = 0
    net_color = COLOR_GREEN if net_float >= 0 else COLOR_RED
    net_bg = (232, 248, 240) if net_float >= 0 else (253, 237, 235)

    draw.rectangle([(margin, y), (page_width - margin, y + net_h)],
                   fill=net_bg, outline=net_color, width=3)

    net_label_text = ar(net_label)
    draw.text((page_width - margin - 25, y + 15), net_label_text,
              fill=COLOR_SECTION_BG, font=font_net, anchor='ra')
    net_value_text = ar(str(net_value))
    draw.text((margin + 25, y + 15), net_value_text,
              fill=net_color, font=font_net, anchor='la')
    y += net_h + 30

    # جدول الحركات
    if movements:
        section_h2 = 45
        draw.rectangle([(margin, y), (page_width - margin, y + section_h2)],
                       fill=COLOR_SECTION_BG)
        title_text = ar(movements_title)
        bbox = draw.textbbox((0, 0), title_text, font=font_section)
        t_w = bbox[2] - bbox[0]
        draw.text(((page_width - t_w) / 2, y + 5), title_text,
                  fill=COLOR_SECTION_FG, font=font_section)
        y += section_h2 + 10

        # رؤوس الأعمدة
        header_h = 45
        draw.rectangle([(margin, y), (page_width - margin, y + header_h)],
                       fill=COLOR_TABLE_HEADER_BG)

        headers = ["المبلغ", "البيان", "النوع", "التاريخ"]
        positions = [
            margin + 20,
            margin + 180,
            page_width - margin - 320,
            page_width - margin - 140,
        ]
        anchors = ['la', 'la', 'ra', 'ra']

        for h_text, x_pos, anch in zip(headers, positions, anchors):
            draw.text((x_pos, y + 10), ar(h_text),
                      fill=(255, 255, 255), font=font_table_header, anchor=anch)
        y += header_h

        # صفوف الحركات
        body_h = 38
        for i, movement in enumerate(movements):
            if y + body_h > img.size[1] - 60:
                new_img = Image.new('RGB',
                                     (page_width, img.size[1] + 400),
                                     'white')
                new_img.paste(img, (0, 0))
                img = new_img
                draw = ImageDraw.Draw(img)

            if i % 2 == 0:
                draw.rectangle([(margin, y), (page_width - margin, y + body_h)],
                               fill=COLOR_ROW_ALT)
            draw.rectangle([(margin, y), (page_width - margin, y + body_h)],
                           outline=COLOR_BORDER)

            if len(movement) == 4:
                date, type_, desc, value = movement
            else:
                date, type_, desc, value = movement[1], movement[2], movement[3], movement[4]

            try:
                value_str = f"{float(value):.2f}"
            except:
                value_str = str(value or '0.00')

            draw.text((margin + 20, y + 7), ar(value_str),
                      fill=COLOR_VALUE, font=font_table_body, anchor='la')

            desc_str = str(desc or '')[:40]
            draw.text((margin + 180, y + 7), ar(desc_str),
                      fill=COLOR_VALUE, font=font_table_body, anchor='la')

            type_str = str(type_ or '')[:20]
            draw.text((page_width - margin - 320, y + 7), ar(type_str),
                      fill=COLOR_VALUE, font=font_table_body, anchor='ra')

            date_str = str(date or '')
            draw.text((page_width - margin - 140, y + 7), ar(date_str),
                      fill=COLOR_VALUE, font=font_table_body, anchor='ra')

            y += body_h

    # التذييل
    y += 20
    footer = ar("شركة الزهراء للبرمجيات -المحاسب/حسام الدين سيد ربيع 01003796984")
    bbox = draw.textbbox((0, 0), footer, font=font_footer)
    f_w = bbox[2] - bbox[0]
    draw.text(((page_width - f_w) / 2, y), footer,
              fill=(149, 165, 166), font=font_footer)

    try:
        img.save(file_path, 'PNG')
        return True
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل حفظ الصورة: {e}")
        return False



def export_table_to_image(file_path, title, subtitle, columns, rows,
                          summary_items=None, page_width=1400, logo_path=None):
    """
    تصدير أي جدول إلى صورة PNG

    - file_path: مسار الحفظ
    - title: العنوان الرئيسي
    - subtitle: العنوان الفرعي
    - columns: قائمة رؤوس الأعمدة (strings)
    - rows: قائمة صفوف، كل صف قائمة قيم (بنفس ترتيب columns)
    - summary_items: قائمة (label, value) للإجماليات في الأسفل (اختياري)
    - page_width: عرض الصورة
    - logo_path: مسار الشعار (اختياري)
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        import arabic_reshaper
        from bidi.algorithm import get_display
    except ImportError as e:
        messagebox.showerror("مكتبات مطلوبة",
                             f"مكتبة مفقودة: {e}\n\n"
                             "نفذ:\n"
                             "pip install arabic-reshaper python-bidi Pillow")
        return False

    # البحث عن خط عربي
    font_candidates = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
    ]
    font_path = None
    for fp in font_candidates:
        if os.path.exists(fp):
            font_path = fp
            break

    if not font_path:
        messagebox.showerror("خطأ", "لم يتم العثور على خط عربي")
        return False

    try:
        font_title = ImageFont.truetype(font_path, 40)
        font_subtitle = ImageFont.truetype(font_path, 22)
        font_header = ImageFont.truetype(font_path, 22)
        font_body = ImageFont.truetype(font_path, 20)
        font_summary = ImageFont.truetype(font_path, 24)
        font_footer = ImageFont.truetype(font_path, 16)
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل تحميل الخط: {e}")
        return False

    def ar(text):
        if text is None:
            return ''
        text = str(text)
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)

    # ????????? حساب الأبعاد ?????????
    margin = 40
    num_cols = len(columns)
    row_height = 40
    header_height = 55

    # ارتفاع الصورة = هيدر + جدول + إجماليات + تذييل
    header_total = 130
    table_height = header_height + (len(rows) * row_height) + 20
    summary_height = (len(summary_items) * 50) if summary_items else 0
    footer_height = 60

    page_height = header_total + 40 + table_height + summary_height + footer_height + 40

    # ضبط العرض ليتناسب مع عدد الأعمدة
    if num_cols > 8:
        page_width = max(page_width, num_cols * 130 + 2 * margin)

    img = Image.new('RGB', (page_width, page_height), 'white')
    draw = ImageDraw.Draw(img)

    # الألوان
    COLOR_HEADER_BG = (44, 62, 80)
    COLOR_HEADER_FG = (255, 255, 255)
    COLOR_SUBTITLE = (241, 196, 15)
    COLOR_TABLE_HEADER_BG = (52, 73, 94)
    COLOR_TABLE_HEADER_FG = (255, 255, 255)
    COLOR_ROW_ALT = (248, 249, 250)
    COLOR_BORDER = (189, 195, 199)
    COLOR_TEXT = (44, 62, 80)
    COLOR_SUMMARY_BG = (255, 243, 205)
    COLOR_SUMMARY_FG = (139, 69, 19)

    # ????????? الهيدر ?????????
    header_h = 110
    draw.rectangle([(0, 0), (page_width, header_h)], fill=COLOR_HEADER_BG)

    # الشعار
    logo_size = 80
    logo_img = None
    if logo_path and os.path.exists(logo_path):
        try:
            logo_img = Image.open(logo_path)
            if logo_img.mode != 'RGBA':
                logo_img = logo_img.convert('RGBA')
            logo_img.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)
            logo_x = page_width - logo_size - 30
            logo_y = (header_h - logo_size) // 2
            img.paste(logo_img, (logo_x, logo_y), logo_img)
        except:
            logo_img = None

    # العنوان
    title_text = ar(title)
    bbox = draw.textbbox((0, 0), title_text, font=font_title)
    title_w = bbox[2] - bbox[0]
    center_x = (page_width - (logo_size + 60 if logo_img else 0)) / 2 + 30
    draw.text((center_x - title_w / 2, 20), title_text,
              fill=COLOR_HEADER_FG, font=font_title)

    if subtitle:
        sub_text = ar(subtitle)
        bbox = draw.textbbox((0, 0), sub_text, font=font_subtitle)
        sub_w = bbox[2] - bbox[0]
        draw.text((center_x - sub_w / 2, 70), sub_text,
                  fill=COLOR_SUBTITLE, font=font_subtitle)

    y = header_h + 30

    # ????????? توزيع الأعمدة ?????????
    available_width = page_width - 2 * margin
    col_width = available_width // num_cols

    # ????????? رؤوس الأعمدة ?????????
    draw.rectangle([(margin, y), (page_width - margin, y + header_height)],
                   fill=COLOR_TABLE_HEADER_BG)

    for i, col_name in enumerate(columns):
        # الأعمدة من اليمين إلى اليسار
        x_right = page_width - margin - (i * col_width)
        x_center = x_right - col_width / 2
        text = ar(col_name)
        bbox = draw.textbbox((0, 0), text, font=font_header)
        text_w = bbox[2] - bbox[0]
        draw.text((x_center - text_w / 2, y + 14), text,
                  fill=COLOR_TABLE_HEADER_FG, font=font_header)

    y += header_height

    # ????????? صفوف البيانات ?????????
    for row_idx, row in enumerate(rows):
        bg = COLOR_ROW_ALT if row_idx % 2 == 0 else 'white'
        draw.rectangle([(margin, y), (page_width - margin, y + row_height)],
                       fill=bg, outline=COLOR_BORDER)

        for col_idx, value in enumerate(row):
            x_right = page_width - margin - (col_idx * col_width)
            x_center = x_right - col_width / 2
            text = ar(value if value is not None else '')
            bbox = draw.textbbox((0, 0), text, font=font_body)
            text_w = bbox[2] - bbox[0]
            draw.text((x_center - text_w / 2, y + 10), text,
                      fill=COLOR_TEXT, font=font_body)

        y += row_height

    y += 15

    # ????????? الإجماليات ?????????
    if summary_items:
        for label, value in summary_items:
            draw.rectangle([(margin, y), (page_width - margin, y + 45)],
                           fill=COLOR_SUMMARY_BG, outline=COLOR_SUMMARY_FG, width=2)
            # التسمية يمين
            draw.text((page_width - margin - 20, y + 12), ar(str(label)),
                      fill=COLOR_SUMMARY_FG, font=font_summary, anchor='ra')
            # القيمة يسار
            draw.text((margin + 20, y + 12), ar(str(value)),
                      fill=(231, 76, 60), font=font_summary, anchor='la')
            y += 50

    # ????????? التذييل ?????????
    y += 20
    footer = ar("شركة الزهراء للبرمجيات -المحاسب/حسام الدين سيد ربيع 01003796984")
    bbox = draw.textbbox((0, 0), footer, font=font_footer)
    f_w = bbox[2] - bbox[0]
    draw.text(((page_width - f_w) / 2, y), footer,
              fill=(149, 165, 166), font=font_footer)

    # ????????? الحفظ ?????????
    try:
        img.save(file_path, 'PNG')
        return True
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل حفظ الصورة: {e}")
        return False


def export_daily_report_to_image(file_path, report_date, print_time,
                                   total_pound, total_visa, total_euro,
                                   total_dollar, total_vodafone_sales,
                                   total_euro_egp, total_dollar_egp,
                                   total_invoices, total_cash_expenses,
                                   total_vodafone_expenses, net_revenue,
                                   expense_items, voda_expense_items=None,
                                   logo_path=None):
    """تصدير تقرير اليومية إلى صورة PNG مع جدولين للمصاريف"""
    if voda_expense_items is None:
        voda_expense_items = []

    try:
        from PIL import Image, ImageDraw, ImageFont
        import arabic_reshaper
        from bidi.algorithm import get_display
    except ImportError as e:
        messagebox.showerror("مكتبات مطلوبة",
            f"مكتبة مفقودة: {e}\n\nنفذ:\n"
            "pip install arabic-reshaper python-bidi Pillow")
        return False

    font_candidates = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
    ]
    font_path = None
    for fp in font_candidates:
        if os.path.exists(fp):
            font_path = fp
            break

    if not font_path:
        messagebox.showerror("خطأ", "لم يتم العثور على خط عربي")
        return False

    try:
        font_company_en = ImageFont.truetype(font_path, 46)
        font_company_ar = ImageFont.truetype(font_path, 38)
        font_section = ImageFont.truetype(font_path, 28)
        font_row = ImageFont.truetype(font_path, 24)
        font_row_bold = ImageFont.truetype(font_path, 26)
        font_value = ImageFont.truetype(font_path, 24)
        font_value_bold = ImageFont.truetype(font_path, 26)
        font_footer = ImageFont.truetype(font_path, 18)
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل تحميل الخط: {e}")
        return False

    def ar(text):
        if text is None:
            return ''
        text = str(text)
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)

    # الشعار
    logo_img = None
    logo_height_allocated = 0
    if logo_path and os.path.exists(logo_path):
        try:
            logo_img = Image.open(logo_path)
            if logo_img.mode != 'RGBA':
                logo_img = logo_img.convert('RGBA')
            max_logo_h = 110
            if logo_img.height > max_logo_h:
                ratio = max_logo_h / logo_img.height
                new_w = int(logo_img.width * ratio)
                logo_img = logo_img.resize((new_w, max_logo_h),
                                            Image.Resampling.LANCZOS)
            logo_height_allocated = logo_img.height + 15
        except Exception as e:
            print(f"⚠️ فشل تحميل الشعار: {e}")
            logo_img = None

    # حساب عدد الصفوف المطلوبة
    sales_rows = 1
    if total_visa > 0: sales_rows += 1
    if total_euro > 0: sales_rows += 1
    if total_dollar > 0: sales_rows += 1
    if total_vodafone_sales > 0: sales_rows += 1

    expenses_rows = 1
    if total_vodafone_expenses > 0: expenses_rows += 1

    # الأبعاد
    page_width = 900
    margin = 30
    inner_margin = 55
    GAP = 15

    h_company_en = 85
    h_company_ar = 70
    h_summary_title = 55
    h_info_row = 50
    h_section = 45
    h_row = 55
    h_row_bold = 65
    h_net_row = 75
    h_items_title = 55
    h_item_row = 50
    h_total_bar = 50

    items_count = max(len(expense_items), 1)
    voda_count = max(len(voda_expense_items), 1)

    page_height = (
        margin +
        logo_height_allocated +
        h_company_en + h_company_ar + GAP +
        h_summary_title +
        (h_info_row * 2) + GAP +
        h_section + (h_row * sales_rows) + h_row_bold + GAP +
        h_section + (h_row * expenses_rows) + GAP +
        h_net_row + GAP +
        h_items_title + (h_item_row * items_count) + h_total_bar + GAP +
        h_items_title + (h_item_row * voda_count) + h_total_bar +
        GAP + 40 + margin
    )

    img = Image.new('RGB', (page_width, page_height), 'white')
    draw = ImageDraw.Draw(img)

    border_color = (44, 62, 80)
    border_width = 2

    COLOR_CASH   = (234, 250, 241)
    COLOR_VISA   = (244, 236, 247)
    COLOR_EURO   = (232, 244, 248)
    COLOR_DOLLAR = (240, 248, 232)
    COLOR_VODA   = (253, 242, 233)
    COLOR_INV    = (254, 249, 231)
    COLOR_EXP    = (253, 237, 236)
    COLOR_NET    = (213, 244, 230)
    COLOR_SECTION = (139, 69, 19)

    def draw_box(x1, y1, x2, y2, fill=None, outline=border_color, width=border_width):
        draw.rectangle([(x1, y1), (x2, y2)],
                       fill=fill, outline=outline, width=width)

    def draw_center_text(text, y_top, height, font, color=(0, 0, 0)):
        text_ar = ar(text)
        draw.text((page_width / 2, y_top + height / 2),
                  text_ar, fill=color, font=font, anchor='mm')

    def draw_label_value(label, value, y_top, height,
                          label_color=(0, 0, 0), value_color=(0, 0, 0),
                          bg_color=None, font_label=None, font_val=None):
        if bg_color:
            draw_box(margin, y_top, page_width - margin,
                     y_top + height, fill=bg_color)
        fl = font_label or font_row
        fv = font_val or font_value
        draw.text((page_width - inner_margin, y_top + height / 2),
                  ar(label), fill=label_color, font=fl, anchor='rm')
        draw.text((inner_margin, y_top + height / 2),
                  ar(str(value)), fill=value_color, font=fv, anchor='lm')

    def draw_section_title(title, y_top, height, bg_color=COLOR_SECTION):
        draw_box(margin, y_top, page_width - margin,
                 y_top + height, fill=bg_color)
        draw_center_text(title, y_top, height, font_section, color=(255, 255, 255))

    y = margin

    # ═══════════════════════════════════════════════════════
    #                    الشعار
    # ═══════════════════════════════════════════════════════
    if logo_img:
        logo_x = (page_width - logo_img.width) // 2
        if logo_img.mode == 'RGBA':
            img.paste(logo_img, (logo_x, y), logo_img)
        else:
            img.paste(logo_img, (logo_x, y))
        y += logo_height_allocated

    # ═══════════════════════════════════════════════════════
    #                  اسم الكافيتريا
    # ═══════════════════════════════════════════════════════
    draw_box(margin, y, page_width - margin, y + h_company_en)
    draw_center_text("مطعم حديقة العائلات", y, h_company_en, font_company_en)
    y += h_company_en

    draw_box(margin, y, page_width - margin, y + h_company_ar)
    draw_center_text("Marsa Pizza", y, h_company_ar, font_company_ar)
    y += h_company_ar + GAP

    # ═══════════════════════════════════════════════════════
    #                ملخص الوردية
    # ═══════════════════════════════════════════════════════
    draw_box(margin, y, page_width - margin, y + h_summary_title,
             fill=(44, 62, 80))
    draw_center_text("ملخص الوردية", y, h_summary_title,
                     font_section, color=(255, 255, 255))
    y += h_summary_title

    draw_box(margin, y, page_width - margin, y + h_info_row,
             fill=(248, 249, 250))
    draw_label_value("التاريخ", report_date, y, h_info_row,
                     label_color=(44, 62, 80))
    y += h_info_row

    draw_box(margin, y, page_width - margin, y + h_info_row,
             fill=(248, 249, 250))
    # ✅ وقت الطباعة نُقل للتذييل
    y += GAP

    # ═══════════════════════════════════════════════════════
    #                💰 قسم المبيعات
    # ═══════════════════════════════════════════════════════
    draw_section_title("💰 المبيعات", y, h_section)
    y += h_section

    draw_label_value("💵 نقدي (L.E)", f"{total_pound:,.2f}",
                     y, h_row,
                     label_color=(39, 174, 96),
                     value_color=(26, 82, 118),
                     bg_color=COLOR_CASH)
    y += h_row

    if total_visa > 0:
        draw_label_value("💳 فيزا", f"{total_visa:,.2f}",
                         y, h_row,
                         label_color=(142, 68, 173),
                         value_color=(26, 82, 118),
                         bg_color=COLOR_VISA)
        y += h_row

    if total_euro > 0:
        draw_label_value("💶 يورو (€)",
                         f"{total_euro:,.2f} €  =  {total_euro_egp:,.2f} ج.م",
                         y, h_row,
                         label_color=(41, 128, 185),
                         value_color=(26, 82, 118),
                         bg_color=COLOR_EURO)
        y += h_row

    if total_dollar > 0:
        draw_label_value("💵 دولار ($)",
                         f"{total_dollar:,.2f} $  =  {total_dollar_egp:,.2f} ج.م",
                         y, h_row,
                         label_color=(22, 160, 133),
                         value_color=(26, 82, 118),
                         bg_color=COLOR_DOLLAR)
        y += h_row

    if total_vodafone_sales > 0:
        draw_label_value("📱 فودافون كاش", f"{total_vodafone_sales:,.2f}",
                         y, h_row,
                         label_color=(230, 126, 34),
                         value_color=(26, 82, 118),
                         bg_color=COLOR_VODA)
        y += h_row

    draw_label_value("🔷 إجمالي الفواتير", f"{total_invoices:,.2f}",
                     y, h_row_bold,
                     label_color=(192, 57, 43),
                     value_color=(192, 57, 43),
                     bg_color=COLOR_INV,
                     font_label=font_row_bold,
                     font_val=font_value_bold)
    y += h_row_bold + GAP

    # ═══════════════════════════════════════════════════════
    #                💸 قسم المصروفات
    # ═══════════════════════════════════════════════════════
    draw_section_title("💸 المصروفات", y, h_section)
    y += h_section

    draw_label_value("🧾 مصروفات نقدية", f"{total_cash_expenses:,.2f}",
                     y, h_row,
                     label_color=(192, 57, 43),
                     value_color=(26, 82, 118),
                     bg_color=COLOR_EXP)
    y += h_row

    if total_vodafone_expenses > 0:
        draw_label_value("📱 مصاريف فودافون (منفصل)",
                         f"{total_vodafone_expenses:,.2f}",
                         y, h_row,
                         label_color=(230, 126, 34),
                         value_color=(26, 82, 118),
                         bg_color=COLOR_VODA)
        y += h_row

    y += GAP

    # ═══════════════════════════════════════════════════════
    # 💰 صافي الإيراد النقدي — قسم بارز
    # ═══════════════════════════════════════════════════════
    # شريط عنوان
    draw_section_title("💰 صافي الإيراد النقدي", y, h_section,
                       bg_color=(26, 82, 118))
    y += h_section

    # شريط القيمة
    net_color = (22, 101, 52) if net_revenue >= 0 else (192, 57, 43)
    net_bg = COLOR_NET if net_revenue >= 0 else (253, 237, 236)

    draw_box(margin, y, page_width - margin, y + h_net_row,
             fill=net_bg, outline=net_color, width=3)

    # ✅ التسمية على اليمين (مثل باقي البنود)
    label_text = ar("💰 صافي الإيراد النقدي")
    draw.text((page_width - inner_margin, y + h_net_row / 2),
              label_text, fill=net_color,
              font=font_value_bold, anchor='rm')

    # ✅ القيمة على اليسار (مثل باقي البنود)
    value_text = ar(f"{net_revenue:,.2f} ج.م")
    draw.text((inner_margin, y + h_net_row / 2),
              value_text, fill=net_color,
              font=font_value_bold, anchor='lm')

    y += h_net_row + GAP

    # ═══════════════════════════════════════════════════════
    #    🧾 الجدول 1: بنود المصروفات النقدية
    # ═══════════════════════════════════════════════════════
    draw_section_title("🧾 بنود المصروفات النقدية", y, h_items_title,
                       bg_color=(139, 69, 19))
    y += h_items_title

    if expense_items:
        for i, (desc, value) in enumerate(expense_items):
            row_bg = (248, 249, 250) if i % 2 == 0 else (255, 255, 255)
            draw_box(margin, y, page_width - margin, y + h_item_row,
                     fill=row_bg)
            desc_text = ar(str(desc)[:55] if desc else '')
            draw.text((page_width - inner_margin, y + h_item_row / 2),
                      desc_text, fill=(44, 62, 80),
                      font=font_row, anchor='rm')
            value_text = ar(f"{value:,.2f}")
            draw.text((inner_margin, y + h_item_row / 2),
                      value_text, fill=(192, 57, 43),
                      font=font_row_bold, anchor='lm')
            y += h_item_row
    else:
        draw_box(margin, y, page_width - margin, y + h_item_row,
                 fill=(248, 249, 250))
        draw_center_text("لا توجد مصروفات نقدية لهذا اليوم",
                         y, h_item_row,
                         font_row, color=(128, 128, 128))
        y += h_item_row

    # شريط إجمالي بنود نقدية
    items_total = sum(v for _, v in expense_items)
    draw_box(margin, y, page_width - margin, y + h_total_bar,
             fill=(26, 82, 118))
    draw.text((page_width - inner_margin, y + h_total_bar / 2),
              ar("✅ إجمالي المصروفات النقدية"),
              fill=(241, 196, 15), font=font_row_bold, anchor='rm')
    draw.text((inner_margin, y + h_total_bar / 2),
              ar(f"{items_total:,.2f} ج.م"),
              fill=(46, 204, 113), font=font_value_bold, anchor='lm')
    y += h_total_bar + GAP

    # ═══════════════════════════════════════════════════════
    #    📱 الجدول 2: بنود مصاريف فودافون كاش
    # ═══════════════════════════════════════════════════════
    draw_section_title("📱 بنود مصاريف فودافون كاش", y, h_items_title,
                       bg_color=(192, 57, 43))
    y += h_items_title

    if voda_expense_items:
        for i, (desc, value) in enumerate(voda_expense_items):
            row_bg = (253, 242, 233) if i % 2 == 0 else (255, 255, 255)
            draw_box(margin, y, page_width - margin, y + h_item_row,
                     fill=row_bg)
            desc_text = ar(str(desc)[:55] if desc else '')
            draw.text((page_width - inner_margin, y + h_item_row / 2),
                      desc_text, fill=(44, 62, 80),
                      font=font_row, anchor='rm')
            value_text = ar(f"{value:,.2f}")
            draw.text((inner_margin, y + h_item_row / 2),
                      value_text, fill=(192, 57, 43),
                      font=font_row_bold, anchor='lm')
            y += h_item_row
    else:
        draw_box(margin, y, page_width - margin, y + h_item_row,
                 fill=(253, 242, 233))
        draw_center_text("لا توجد مصاريف فودافون لهذا اليوم",
                         y, h_item_row,
                         font_row, color=(128, 128, 128))
        y += h_item_row

    # شريط إجمالي بنود فودافون
    voda_total = sum(v for _, v in voda_expense_items)
    draw_box(margin, y, page_width - margin, y + h_total_bar,
             fill=(123, 36, 28))
    draw.text((page_width - inner_margin, y + h_total_bar / 2),
              ar("✅ إجمالي مصاريف فودافون"),
              fill=(241, 196, 15), font=font_row_bold, anchor='rm')
    draw.text((inner_margin, y + h_total_bar / 2),
              ar(f"{voda_total:,.2f} ج.م"),
              fill=(245, 183, 177), font=font_value_bold, anchor='lm')
    y += h_total_bar

    # ═══════════════════════════════════════════════════════
    # التذييل
    # ═══════════════════════════════════════════════════════
    y += 20

    # ✅ وقت الطباعة فوق سطر الشركة
    time_text = ar(f"🕐 وقت الطباعة: {print_time}")
    bbox = draw.textbbox((0, 0), time_text, font=font_footer)
    t_w = bbox[2] - bbox[0]
    draw.text(((page_width - t_w) / 2, y), time_text,
              fill=(44, 62, 80), font=font_footer)
    y += 25

    footer = ar("شركة الزهراء للبرمجيات - المحاسب/حسام الدين سيد ربيع 01003796984 - برنامج حسابات المطاعم")

    bbox = draw.textbbox((0, 0), footer, font=font_footer)
    f_w = bbox[2] - bbox[0]
    draw.text(((page_width - f_w) / 2, y), footer,
              fill=(149, 165, 166), font=font_footer)

    try:
        img.save(file_path, 'PNG')
        return True
    except Exception as e:
        messagebox.showerror("خطأ", f"فشل حفظ الصورة: {e}")

        return False





# ═══════════════════════════════════════════════════════════════════
#                       HELPERS
# ═══════════════════════════════════════════════════════════════════
COMPANY_AD = "🏢 شركة الزهراء للبرمجيات المحاسب/ حسام الدين سيد ربيع - أنظمة إدارة متكاملة | جودة واحترافية| للاستفسار: 01003796984 📞 "
AD_BG_COLOR = "#2c3e50"
AD_FG_COLOR = "#f1c40f"
AD_SPEED = 80
AD_DISPLAY_LENGTH = 70
COMPANY_LOGO_PATH = resource_path("m.png")


# ✅ قائمة عامة بليبلات الإعلان في الشريط العلوي (تحدّثها MainApp)
BANNER_AD_LABELS = []

# القائمة الجانبية (بنفس ألوان الشريط العلوي)
# ═══════════════════════════════════════════════════════
SIDEBAR_BG = '#154360'  # أزرق أغمق (مطابق للبانر)
SIDEBAR_BTN_BG = '#2471a3'  # أزرق متوسط للأزرار
SIDEBAR_BTN_HOVER = '#f1c40f'  # ذهبي عند المرور
SIDEBAR_FG = 'white'  # لون النص
SIDEBAR_ACCENT = '#f1c40f'  # لون الفواصل والعنوان


def copy_image_to_project(source_path, record_id, target_dir='bill_images/'):
    src = clean_path(source_path)
    if not src or not os.path.exists(src) or record_id is None:
        return None
    os.makedirs(target_dir, exist_ok=True)
    ext = os.path.splitext(src)[1]
    if not ext:
        return None
    target_path = os.path.join(target_dir, f"{record_id}{ext}")
    try:
        shutil.copy2(src, target_path)
        return target_path
    except Exception as e:
        print(f"? فشل نسخ الصورة: {e}")
        return None


def clean_path(path):
    if path is None:
        return None
    if isinstance(path, tuple):
        path = path[0] if path else None
    if path is not None:
        path = str(path).strip()
        return path if path else None
    return None


def safe_float(value, default=0):
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        value = value.strip()
        if value == '':
            return default
        try:
            return float(value)
        except ValueError:
            return default
    return default


def get_prev_month(year, month):
    if month == 1:
        return year - 1, 12
    return year, month - 1


def get_next_month(year, month):
    if month == 12:
        return year + 1, 1
    return year, month + 1


ARABIC_MONTHS = {1: 'يناير', 2: 'فبراير', 3: 'مارس', 4: 'أبريل', 5: 'مايو', 6: 'يونيو',
                 7: 'يوليو', 8: 'أغسطس', 9: 'سبتمبر', 10: 'أكتوبر', 11: 'نوفمبر', 12: 'ديسمبر'}


def calculate_work_days(first_work_date, vacations=0, end_date=None):
    """
    حساب أيام العمل الفعلية

    المعادلة: عدد الأيام الكلي - الإجازات

    - first_work_date: أول يوم عمل (YYYY-MM-DD)
    - vacations: عدد أيام الإجازات
    - end_date: تاريخ النهاية (افتراضياً اليوم)
    """
    if not first_work_date:
        return 0

    try:
        start = datetime.strptime(first_work_date, '%Y-%m-%d')
    except:
        return 0

    if end_date:
        try:
            end = datetime.strptime(end_date, '%Y-%m-%d')
        except:
            end = datetime.now()
    else:
        end = datetime.now()

    # عدد الأيام من أول يوم عمل حتى اليوم
    total_days = (end - start).days + 1  # +1 لاحتساب يوم البداية

    if total_days < 0:
        total_days = 0

    # طرح الإجازات
    work_days = total_days - (vacations or 0)

    return max(0, work_days)


    # ===================== دوال العمليات على قاعدة البيانات المبيعات =====================


def add_sale(date, pound, vodafone, visa, dollar, euro, dollar_rate, euro_rate, image_source_path):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''
             INSERT INTO sales (date, pound, vodafone, visa, dollar, euro,
                                dollar_rate, euro_rate, image_path)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
         ''', (date, pound, vodafone, visa, dollar, euro,
               dollar_rate, euro_rate, None))

        sale_id = c.lastrowid
        conn.commit()

        final_path = None
        if image_source_path and os.path.exists(image_source_path):
            final_path = copy_image_to_project(image_source_path, sale_id, 'sales_images/')
            if final_path:
                c.execute('UPDATE sales SET image_path = ? WHERE id = ?', (final_path, sale_id))
                conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ في add_sale: {e}")
        return False




def update_sale(sale_id, date, pound, vodafone, visa, dollar, euro,
                dollar_rate, euro_rate, new_image_path):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('SELECT image_path FROM sales WHERE id = ?', (sale_id,))
        row = c.fetchone()
        old_path = clean_path(row[0]) if row else None
        new_path = clean_path(new_image_path)
        final_path = old_path

        if new_path == "REMOVE":
            if old_path and os.path.exists(old_path):
                try: os.remove(old_path)
                except: pass
            final_path = None
        elif new_path is not None and new_path != old_path:
            if old_path and os.path.exists(old_path):
                try: os.remove(old_path)
                except: pass
            final_path = copy_image_to_project(new_path, sale_id, 'sales_images/')
        elif new_path is not None and old_path is None:
            final_path = copy_image_to_project(new_path, sale_id, 'sales_images/')

        c.execute('''
             UPDATE sales
             SET date=?, pound=?, vodafone=?, visa=?, dollar=?, euro=?,
                 dollar_rate=?, euro_rate=?, image_path=?
             WHERE id=?
         ''', (date, pound, vodafone, visa, dollar, euro,
               dollar_rate, euro_rate, final_path, sale_id))

        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"❌ خطأ في update_sale: {e}")
        return False


def delete_sale(sale_id):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('SELECT image_path FROM sales WHERE id = ?', (sale_id,))
        row = c.fetchone()
        image_path = clean_path(row[0]) if row else None

        c.execute('DELETE FROM sales WHERE id = ?', (sale_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()

        if affected > 0:
            if image_path and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                    print(f"??? تم حذف الصورة: {image_path}")
                except Exception as e:
                    print(f"?? فشل حذف الصورة: {e}")
            return True
        else:
            print(f"?? لم يتم العثور على صف برقم {sale_id}")
            return False

    except Exception as e:
        print(f"? خطأ في delete_sale: {e}")
        return False


def get_all_sales():
    try:
        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM sales ORDER BY id DESC')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في get_all_sales: {e}")
        return []


def search_sales(search_term=None, date=None):
    try:
        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        query = "SELECT * FROM sales WHERE 1=1"
        params = []

        if search_term and search_term.strip():
            query += """ AND (date LIKE ? OR pound LIKE ? OR vodafone LIKE ?
                              OR visa LIKE ? OR dollar LIKE ? OR euro LIKE ?)"""
            term = f"%{search_term.strip()}%"
            params.extend([term]*6)
        if date and date.strip():
            query += " AND date = ?"
            params.append(date.strip())

        query += " ORDER BY id"
        c.execute(query, params)
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في search_sales: {e}")
        return []
def search_sales_filtered(search_term=None, date_from=None, date_to=None):
    """بحث مفلتر بالمدة والنص"""
    try:
        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        query = "SELECT * FROM sales WHERE 1=1"
        params = []

        if search_term and search_term.strip():
            query += """ AND (date LIKE ? OR pound LIKE ? OR vodafone LIKE ?
                              OR visa LIKE ? OR dollar LIKE ? OR euro LIKE ?)"""
            term = f"%{search_term.strip()}%"
            params.extend([term] * 6)

        if date_from and date_from.strip():
            query += " AND date >= ?"
            params.append(date_from.strip())

        if date_to and date_to.strip():
            query += " AND date <= ?"
            params.append(date_to.strip())

        query += " ORDER BY date DESC, id DESC"
        c.execute(query, params)
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في search_sales_filtered: {e}")
        return []

# ═══════════════════════════════════════════════════════════════════
# ✅ مزامنة الخزينة — دوال مساعدة
# ═══════════════════════════════════════════════════════════════════

def classify_treasury_action(main_name, sub_name):
    """تصنيف حركة الخزينة بناءً على التصنيف الرئيسي"""
    if main_name == 'مشتريات':
        return 'سداد تاجر', (sub_name or '')
    elif main_name == 'أجور ومرتبات':
        return 'سداد عامل', (sub_name or '')
    else:
        return 'مصروف آخر', (sub_name or main_name or '')

def cleanup_wrong_treasury_actions():
    """حذف حركات الخزينة المرتبطة بمصروفات ليست بطريقة 'الخزينة'"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""
            DELETE FROM treasury_actions
            WHERE paid_id IS NOT NULL AND paid_id > 0
              AND paid_id IN (
                  SELECT id FROM paids
                  WHERE COALESCE(payment_method, 'نقدي') != 'الخزينة'
              )
        """)
        deleted = c.rowcount
        conn.commit()
        conn.close()
        if deleted > 0:
            print(f"🧹 تم تنظيف {deleted} حركة خزينة خاطئة")
        return deleted
    except Exception as e:
        print(f"⚠️ خطأ في التنظيف: {e}")
        return 0

def add_treasury_action_for_paid(paid_id, date, main_name, sub_name,
                                  value, description):
    """إنشاء حركة خزينة مرتبطة بمصروف — فقط لو الطريقة 'الخزينة'"""
    if value <= 0:
        return True

    # ✅ التحقق الفعلي من قاعدة البيانات
    try:
        conn_check = sqlite3.connect(DB_NAME)
        c_check = conn_check.cursor()
        c_check.execute("SELECT payment_method FROM paids WHERE id=?",
                        (paid_id,))
        row = c_check.fetchone()
        conn_check.close()

        actual_method = (row[0] if row else '') or 'نقدي'

        if actual_method != 'الخزينة':
            print(f"⏭️ تم تخطي مزامنة الخزينة للقيد #{paid_id} "
                  f"(طريقة الدفع: {actual_method})")
            return False
    except Exception as e:
        print(f"⚠️ فشل التحقق من طريقة الدفع: {e}")
        return False

    # ─── بقية الدالة كما هي ───
    action_type, beneficiary = classify_treasury_action(main_name, sub_name)
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''INSERT INTO treasury_actions
            (date, action_type, beneficiary, amount,
             description, paid_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)''',
            (date, action_type, beneficiary, value,
             description or f'مصروف من الخزينة #{paid_id}',
             paid_id,
             datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ في add_treasury_action_for_paid: {e}")
        return False



def delete_treasury_action_for_paid(paid_id):
    """حذف حركة الخزينة المرتبطة بمصروف"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM treasury_actions WHERE paid_id=?", (paid_id,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ في delete_treasury_action_for_paid: {e}")
        return False


def update_treasury_action_for_paid(paid_id, date, main_name, sub_name,
                                     value, description):
    """تحديث حركة الخزينة المرتبطة بمصروف (حذف + إعادة إنشاء)"""
    delete_treasury_action_for_paid(paid_id)
    if value > 0:
        add_treasury_action_for_paid(paid_id, date, main_name, sub_name,
                                      value, description)
    return True


# =====================  جدول المصاريف دوال العمليات على قاعدة البيانات =====================

# ═══════════════════════════════════════════════════════════════════
#                    EXTERNAL PAIDS DB FUNCTIONS
# ═══════════════════════════════════════════════════════════════════
# ============================================================
#                    دوال جدول المصاريف
# ============================================================

def add_paid(date, main_name, sub_name, description, value,
             payment_method, transfer_fee, image_source_path,
             sub_name2=None):
    """إضافة مصروف مع نوع الحركة الاختياري"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        c.execute('''INSERT INTO paids
        (date, value, main_name, sub_name, description,
        payment_method, image_path, sub_name2)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
        (date, value, main_name, sub_name, description,
        payment_method or 'نقدي', None, sub_name2))
        paid_id = c.lastrowid
        conn.commit()

        # الصورة
        if image_source_path and os.path.exists(image_source_path):
            final_path = copy_image_to_project(image_source_path, paid_id)
            if final_path:
                c.execute('UPDATE paids SET image_path = ? WHERE id = ?',
                          (final_path, paid_id))
                conn.commit()
        conn.close()

        # الرسوم
        if payment_method == 'فودافون كاش' and transfer_fee and transfer_fee > 0:
            add_fee_for_paid(paid_id, date, transfer_fee,
                             sub_name or main_name,
                             f"رسوم تحويل: {description or ''}")

        # مزامنة الخزينة
        if payment_method == 'الخزينة':
            add_treasury_action_for_paid(paid_id, date, main_name, sub_name,
                                          value, description)

        return True
    except Exception as e:
        print(f"❌ خطأ في add_paid: {e}")
        return False




def update_paid(paid_id, date, main_name, sub_name, description,
                value, payment_method, transfer_fee, new_image_path,
                sub_name2=None):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        c.execute('SELECT image_path FROM paids WHERE id=?', (paid_id,))
        row = c.fetchone()
        old_path = clean_path(row[0]) if row else None
        new_path = clean_path(new_image_path)
        final_path = old_path

        if new_path == "REMOVE":
            if old_path and os.path.exists(old_path):
                try: os.remove(old_path)
                except: pass
            final_path = None
        elif new_path is not None and new_path != old_path:
            if old_path and os.path.exists(old_path):
                try: os.remove(old_path)
                except: pass
            final_path = copy_image_to_project(new_path, paid_id)

        c.execute('''UPDATE paids
        SET date=?, value=?, main_name=?, sub_name=?,
        description=?, payment_method=?, image_path=?, sub_name2=?
        WHERE id=?''',
        (date, value, main_name, sub_name, description,
        payment_method or 'نقدي', final_path, sub_name2, paid_id))
        affected = c.rowcount

        conn.commit()
        conn.close()

        # تحديث الرسوم
        update_fee_for_paid(paid_id, date, transfer_fee or 0,
                            sub_name or main_name,
                            f"رسوم تحويل: {description or ''}")

        # مزامنة الخزينة
        update_treasury_action_for_paid(paid_id, date, main_name, sub_name,
                                         value, description)

        return affected > 0
    except Exception as e:
        print(f"❌ خطأ في update_paid: {e}")
        return False



def delete_paid(paid_id):
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        c.execute('SELECT image_path FROM paids WHERE id=?', (paid_id,))
        row = c.fetchone()
        image_path = clean_path(row[0]) if row else None

        c.execute('DELETE FROM paids WHERE id=?', (paid_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()

        # ✅ حذف الرسوم المرتبطة
        delete_fee_for_paid(paid_id)

        # ✅ حذف حركة الخزينة المرتبطة
        delete_treasury_action_for_paid(paid_id)

        if affected > 0 and image_path and os.path.exists(image_path):
            try: os.remove(image_path)
            except: pass

        return affected > 0
    except Exception as e:
        print(f"❌ خطأ في delete_paid: {e}")
        return False


def get_all_paids():
    """جلب كل المصاريف"""
    try:
        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM paids ORDER BY date DESC, id DESC')
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في get_all_paids: {e}")
        return []


def search_paids(search_term=None, start_date=None, end_date=None,
                  main_name=None, sub_name=None, payment_method=None):
    """بحث مفلتر في المصاريف"""
    try:
        conn = sqlite3.connect(DB_NAME)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        query = "SELECT * FROM paids WHERE 1=1"
        params = []

        if search_term and search_term.strip():
            query += """ AND (description LIKE ? OR main_name LIKE ?
                              OR sub_name LIKE ? OR value LIKE ?
                              OR date LIKE ?)"""
            term = f"%{search_term.strip()}%"
            params.extend([term] * 5)

        if start_date and start_date.strip():
            query += " AND date >= ?"
            params.append(start_date.strip())

        if end_date and end_date.strip():
            query += " AND date <= ?"
            params.append(end_date.strip())

        if main_name and main_name.strip() and main_name != "ALL":
            query += " AND main_name = ?"
            params.append(main_name.strip())

        if sub_name and sub_name.strip() and sub_name != "ALL":
            query += " AND sub_name = ?"
            params.append(sub_name.strip())

        if payment_method and payment_method.strip() and payment_method != "الكل":
            query += " AND payment_method = ?"
            params.append(payment_method.strip())

        query += " ORDER BY date DESC, id DESC"
        c.execute(query, params)
        rows = c.fetchall()
        result = [dict(row) for row in rows]
        conn.close()
        return result
    except Exception as e:
        print(f"خطأ في search_paids: {e}")
        return []



def add_fee_for_paid(paid_id, date, amount, beneficiary, description):
    """إنشاء سجل رسوم مرتبط بمصروف فودافون في vodafone_wallet_actions"""
    if amount <= 0:
        return True  # لا داعي لإنشاء سجل للرسوم الصفرية
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''INSERT INTO vodafone_wallet_actions
                      (date, action_type, beneficiary, amount,
                       description, paid_id, created_at)
                      VALUES (?, 'رسوم تحويل', ?, ?, ?, ?, ?)''',
                  (date, beneficiary or '', amount,
                   description or f'رسوم على مصروف #{paid_id}',
                   paid_id,
                   datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ في add_fee_for_paid: {e}")
        return False


def update_fee_for_paid(paid_id, date, amount, beneficiary, description):
    """تحديث رسوم مرتبطة بمصروف (حذف + إعادة إنشاء)"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM vodafone_wallet_actions WHERE paid_id=?",
                  (paid_id,))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"⚠️ خطأ في حذف الرسوم القديمة: {e}")

    if amount > 0:
        add_fee_for_paid(paid_id, date, amount, beneficiary, description)
    return True


def delete_fee_for_paid(paid_id):
    """حذف الرسوم المرتبطة بمصروف"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM vodafone_wallet_actions WHERE paid_id=?",
                  (paid_id,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ: {e}")
        return False


def get_fee_for_paid(paid_id):
    """جلب رسوم مصروف معيّن"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""SELECT COALESCE(amount, 0) FROM vodafone_wallet_actions
                      WHERE paid_id=? AND action_type='رسوم تحويل'""",
                  (paid_id,))
        row = c.fetchone()
        conn.close()
        return row[0] if row else 0
    except Exception as e:
        print(f"خطأ: {e}")
        return 0


def add_program_banner(parent, bg='#1a5276', fg_title='#f1c40f',
                       logo_path=None, logo_size=70):

    """[شريط علوي احترافي: شعار يسار + إعلان منتصف + اسم البرنامج يمين]"""
    if logo_path is None:
        logo_path = COMPANY_LOGO_PATH

    # ═══════════════════════════════════════════════════════
    # الألوان
    # ═══════════════════════════════════════════════════════
    BG_MAIN     = '#1a5276'      # أزرق داكن (رئيسي)
    BG_LOGO     = '#154360'      # أغمق للشعار
    BG_AD       = '#154360'      # أغمق للمنتصف
    BG_TEXT     = '#154360'      # للمعلومات
    SEPARATOR   = '#f1c40f'      # فاصل ذهبي

    # ارتفاع البانر
    banner = tk.Frame(parent, bg=BG_MAIN, height=90)
    banner.pack(fill='x', side='top')
    banner.pack_propagate(False)

    # ═══════════════════════════════════════════════════════
    # 1) الشعار على اليسار (بخلفية داكنة قليلًا)
    # ═══════════════════════════════════════════════════════
    logo_frame = tk.Frame(banner, bg=BG_LOGO, width=120)
    logo_frame.pack(side='left', fill='y')
    logo_frame.pack_propagate(False)

    if logo_path and os.path.exists(logo_path):
        try:
            logo_img = Image.open(logo_path)
            if logo_img.mode != 'RGBA':
                logo_img = logo_img.convert('RGBA')
            logo_img.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)
            logo_frame._logo_photo = ImageTk.PhotoImage(logo_img)
            tk.Label(logo_frame, image=logo_frame._logo_photo,
                     bg=BG_LOGO).pack(expand=True)
        except Exception as e:
            print(f"⚠️ [لم يتم تحميل اللوجو]: {e}")

    # ─── فاصل ذهبي ───
   # tk.Frame(banner, bg=SEPARATOR, width=2).pack(side='left', fill='y')

    # ═══════════════════════════════════════════════════════
    # 2) اسم البرنامج والشركة على اليمين
    # ═══════════════════════════════════════════════════════
    text_frame = tk.Frame(banner, bg=BG_TEXT)
    text_frame.pack(side='right', fill='y', padx=0)

    # سطر أول: اسم البرنامج
    tk.Label(text_frame, text="🍽️  RestoBooks",
             font=('Arial', 17, 'bold'),
             bg=BG_TEXT, fg=fg_title).pack(anchor='c', pady=(15, 2))

    # خط فاصل أفقي رفيع
#    tk.Frame(text_frame, bg=fg_title, height=1).pack(
  #      fill='x', pady=4)

    # سطر ثاني: اسم الشركة
    tk.Label(text_frame,
             text="شركة الزهراء للبرمجيات - المحاسب / حسام الدين سيد ربيع 📞  01003796984",
             font=('Arial', 11, 'bold'),
             bg=BG_TEXT, fg='white').pack(anchor='e')

    # سطر ثالث: رقم الهاتف
    tk.Label(text_frame,
             text="📞  01003796984",
             font=('Arial', 10, 'bold'),
             bg=BG_TEXT, fg='#bdc3c7').pack(anchor='e', pady=(2, 15))

    # ─── فاصل ذهبي ───
   # tk.Frame(banner, bg=SEPARATOR, width=2).pack(side='right', fill='y')

    # ═══════════════════════════════════════════════════════
    # 3) الإعلان المتحرك في المنتصف
    # ═══════════════════════════════════════════════════════
    ad_frame = tk.Frame(banner, bg=BG_AD)
    ad_frame.pack(side='left', fill='both', expand=True)

    ad_label = tk.Label(
        ad_frame,
        text=" " * AD_DISPLAY_LENGTH,
        font=('Arial', 11, 'bold'),
        bg=BG_AD,
        fg=fg_title,
        anchor='center',
    )
    ad_label.pack(expand=True, fill='both')

    # ✅ سجّل الليبل لتحديثه من MainApp
    BANNER_AD_LABELS.append(ad_label)

    return banner


def make_scrollable(page_frame, bg_color='#ecf0f1'):
    """يحوّل الإطار إلى منطقة قابلة للتمرير — النسخة المضمونة"""
    page_frame.configure(bg=bg_color)

    # الحاوية الرئيسية للتمرير
    scroll_holder = tk.Frame(page_frame, bg=bg_color)
    scroll_holder.pack(fill='both', expand=True, side='top')

    canvas = tk.Canvas(scroll_holder, bg=bg_color,
                       highlightthickness=0, bd=0)
    vsb = ttk.Scrollbar(scroll_holder, orient='vertical',
                        command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)

    vsb.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)

    # ✅ الإطار الداخلي
    content = tk.Frame(canvas, bg=bg_color)
    win_id = canvas.create_window((0, 0), window=content,
                                  anchor='nw', tags='content')

    # ✅ ضبط العرض ليطابق عرض الـ canvas
    def _sync_width(event):
        canvas.itemconfig(win_id, width=event.width)
    canvas.bind('<Configure>', _sync_width)

    # ✅ ضبط منطقة التمرير
    def _update_scrollregion(event=None):
        canvas.configure(scrollregion=canvas.bbox('all'))
    content.bind('<Configure>', _update_scrollregion)

    # ✅ عجلة الفأرة
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
    canvas.bind_all('<MouseWheel>', _on_mousewheel)

    page_frame._canvas = canvas
    page_frame._content = content
    return content


# ============================================================
#              دوال محفظة فودافون كاش
# ============================================================

def add_vodafone_action(date, action_type, beneficiary, amount, description):
    """إضافة حركة جديدة على محفظة فودافون كاش"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''INSERT INTO vodafone_wallet_actions
                     (date, action_type, beneficiary, amount,
                      description, created_at)
                     VALUES (?, ?, ?, ?, ?, ?)''',
                  (date, action_type, beneficiary or '', amount,
                   description or '',
                   datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطأ في add_vodafone_action: {e}")
        return False


def delete_vodafone_action(action_id):
    """حذف حركة من محفظة فودافون"""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("DELETE FROM vodafone_wallet_actions WHERE id=?",
                  (action_id,))
        affected = c.rowcount
        conn.commit()
        conn.close()
        return affected > 0
    except Exception as e:
        print(f"❌ خطأ في delete_vodafone_action: {e}")
        return False





def get_vodafone_wallet_balance():
    """
    رصيد فودافون = مبيعات فودافون
                  − مصاريف paids (فودافون كاش)
                  − كل حركات vodafone_wallet_actions (بما فيها الرسوم)
    """
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        # 1. إيراد فودافون
        c.execute("SELECT COALESCE(SUM(vodafone), 0) FROM sales")
        total_income = c.fetchone()[0] or 0

        # 2. مصاريف من paids
        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids 
                     WHERE payment_method = 'فودافون كاش'""")
        paids_expenses = c.fetchone()[0] or 0

        # 3. من vodafone_wallet_actions (يشمل الحركات اليدوية + الرسوم)
        c.execute("""SELECT 
                       COALESCE(SUM(CASE WHEN action_type != 'رسوم تحويل' 
                                         THEN amount ELSE 0 END), 0),
                       COALESCE(SUM(CASE WHEN action_type = 'رسوم تحويل' 
                                         THEN amount ELSE 0 END), 0)
                     FROM vodafone_wallet_actions""")
        wallet_expenses, total_fees = c.fetchone()
        wallet_expenses = wallet_expenses or 0
        total_fees = total_fees or 0

        conn.close()
        return {
            'income': total_income,
            'paids_expenses': paids_expenses,
            'wallet_expenses': wallet_expenses,
            'fees': total_fees,
            'expenses': paids_expenses + wallet_expenses,
            'balance': total_income - paids_expenses - wallet_expenses - total_fees
        }
    except Exception as e:
        print(f"خطأ: {e}")
        return {'income': 0, 'paids_expenses': 0, 'wallet_expenses': 0,
                'fees': 0, 'expenses': 0, 'balance': 0}

class VodafoneActionWindow:
    """نافذة تسجيل حركة جديدة على محفظة فودافون كاش"""
    def __init__(self, parent_page, action_type):
        self.parent_page = parent_page
        self.action_type = action_type

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title(f"إضافة حركة: {action_type}")
        self.window.geometry("520x480")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # العنوان
        title_frame = tk.Frame(self.window, bg='#c0392b', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icons = {
            'سداد تاجر': '👨‍💼',
            'سداد عامل': '👷',
            'مصروف آخر': '💸'
        }
        icon = icons.get(action_type, '📱')

        tk.Label(title_frame, text=f"{icon} {action_type}",
                 font=('Arial', 15, 'bold'),
                 bg='#c0392b', fg='white').pack(pady=15)

        # الإطار الرئيسي
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # التاريخ
        tk.Label(main, text="التاريخ: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                    row=0, column=0, padx=5, pady=8, sticky='e')
        self.entry_date = DateEntry(main, width=22,
                                     date_pattern='yyyy-mm-dd')
        self.entry_date.grid(row=0, column=1, padx=5, pady=8, sticky='w')
        self.entry_date.set_date(datetime.now())

        # المستفيد (تاجر / عامل)
        if action_type in ('سداد تاجر', 'سداد عامل'):
            if action_type == 'سداد تاجر':
                label = "التاجر: *"
                table = 'traders'
            else:
                label = "العامل: *"
                table = 'employees'

            tk.Label(main, text=label, font=('Arial', 11, 'bold'),
                     bg='#ecf0f1', anchor='e').grid(
                        row=1, column=0, padx=5, pady=8, sticky='e')

            self.combo_beneficiary = AutocompleteCombobox(
                main, width=22, font=('Arial', 11))
            self.combo_beneficiary.grid(row=1, column=1, padx=5,
                                         pady=8, sticky='w')

            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute(f"SELECT name FROM {table} ORDER BY name")
            names = [r[0] for r in c.fetchall()]
            conn.close()
            self.combo_beneficiary.set_values(names)
            if names:
                self.combo_beneficiary.set(names[0])
        else:
            self.combo_beneficiary = None

        # المبلغ
        row = 2 if self.combo_beneficiary else 1
        tk.Label(main, text="المبلغ (ج.م): *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#c0392b', anchor='e').grid(
                    row=row, column=0, padx=5, pady=8, sticky='e')

        self.entry_amount = tk.Entry(main, width=22, font=('Arial', 13),
                                      justify='center', bg='#fadbd8')
        self.entry_amount.grid(row=row, column=1, padx=5,
                                pady=8, sticky='w')

        # البيان
        row += 1
        tk.Label(main, text="البيان:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='ne').grid(
                    row=row, column=0, padx=5, pady=8, sticky='ne')

        self.entry_desc = tk.Text(main, width=22, height=4,
                                   font=('Arial', 11))
        self.entry_desc.grid(row=row, column=1, padx=5,
                              pady=8, sticky='w')

        # عرض رصيد المحفظة الحالي
        row += 1
        info_frame = tk.Frame(main, bg='#fff9e6', relief='solid', bd=1)
        info_frame.grid(row=row, column=0, columnspan=2,
                         pady=15, sticky='ew')

        balance_info = get_vodafone_wallet_balance()
        tk.Label(info_frame,
                 text=f"📱 رصيد المحفظة الحالي: {balance_info['balance']:,.2f} ج.م",
                 font=('Arial', 12, 'bold'),
                 bg='#fff9e6', fg='#c0392b').pack(pady=10, padx=15)

        # أزرار الحفظ
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        self.entry_amount.focus()

    def save(self):
        date_str = self.entry_date.get().strip()
        desc = self.entry_desc.get('1.0', 'end').strip()

        if not date_str:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return

        try:
            amount = float(self.entry_amount.get().strip() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "المبلغ يجب أن يكون رقماً")
            return

        if amount <= 0:
            messagebox.showwarning("تنبيه", "المبلغ يجب أن يكون أكبر من صفر")
            return

        beneficiary = ''
        if self.combo_beneficiary:
            beneficiary = self.combo_beneficiary.get().strip()
            if not beneficiary:
                messagebox.showwarning("تنبيه", "يجب اختيار المستفيد")
                return

        # تحذير لو المبلغ أكبر من الرصيد
        balance_info = get_vodafone_wallet_balance()
        if amount > balance_info['balance']:
            if not messagebox.askyesno(
                "تحذير",
                f"المبلغ ({amount:,.2f}) أكبر من رصيد المحفظة "
                f"({balance_info['balance']:,.2f}).\n\n"
                f"هل تريد المتابعة على أي حال؟"):
                return

        ok = add_vodafone_action(date_str, self.action_type,
                                  beneficiary, amount, desc)

        if ok:
            messagebox.showinfo("نجاح",
                f"✅ تم حفظ الحركة\n"
                f"المبلغ: {amount:,.2f} ج.م")
            self.window.destroy()
            self.parent_page.load_actions()
        else:
            messagebox.showerror("خطأ", "فشل حفظ الحركة")
class VodafoneWalletPage(tk.Frame):
    """صفحة إدارة محفظة فودافون كاش"""
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.selected_id = None
        # 1️⃣ الشريط العلوي
        add_program_banner(self)

        # 2️⃣ منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # العنوان
        tk.Label(content, text="📱 محفظة فودافون كاش",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#c0392b').pack(pady=12)

        # ============ بطاقة الرصيد ============
        balance_card = tk.Frame(content, bg='#c0392b', height=140)
        balance_card.pack(fill='x', padx=15, pady=(0, 10))
        balance_card.pack_propagate(False)

        tk.Label(balance_card, text="💳 رصيد محفظة فودافون الحالي",
                 font=('Arial', 13, 'bold'),
                 bg='#c0392b', fg='#f9e79f').pack(pady=(12, 5))

        self.lbl_balance = tk.Label(balance_card, text="0.00",
                                     font=('Arial', 26, 'bold'),
                                     bg='#c0392b', fg='white')
        self.lbl_balance.pack()

        details_row = tk.Frame(balance_card, bg='#c0392b')
        details_row.pack(pady=(5, 0))

        self.lbl_income = tk.Label(details_row,
                                    text="إجمالي الإيراد: 0.00",
                                    font=('Arial', 11, 'bold'),
                                    bg='#c0392b', fg='#a9dfbf')
        self.lbl_income.pack(side='right', padx=15)

        self.lbl_expenses = tk.Label(details_row,
                                      text="إجمالي المصاريف: 0.00",
                                      font=('Arial', 11, 'bold'),
                                      bg='#c0392b', fg='#f5b7b1')
        self.lbl_expenses.pack(side='right', padx=15)

        # ============ فلاتر المدة ============
        filter_frame = tk.LabelFrame(content, text="فلترة الحركات",
                                      font=('Arial', 11, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=5, pady=10)
        filter_frame.pack(fill='x', padx=15, pady=(0, 5))

        date_row = tk.Frame(filter_frame, bg='#ecf0f1')
        date_row.pack(fill='x', pady=(0, 8))

        tk.Label(date_row, text="من تاريخ:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='right', padx=5)
        today = datetime.now()
        self.date_from = DateEntry(date_row, width=13,
                                    date_pattern='yyyy-mm-dd')
        self.date_from.pack(side='right', padx=5)
        self.date_from.set_date(today.replace(day=1))

        tk.Label(date_row, text="إلى تاريخ:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='right', padx=5)
        self.date_to = DateEntry(date_row, width=13,
                                  date_pattern='yyyy-mm-dd')
        self.date_to.pack(side='right', padx=5)
        self.date_to.set_date(today)

        tk.Button(date_row, text="🔍 عرض", command=self.load_actions,
                  bg='#2980b9', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=12, pady=5).pack(side='left', padx=5)

        # فلتر أنواع الحركة
        type_row = tk.Frame(filter_frame, bg='#ecf0f1')
        type_row.pack(fill='x')

        tk.Label(type_row, text="نوع الحركة:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#34495e').pack(side='right', padx=5)

        self.var_trader = tk.BooleanVar(value=True)
        self.var_employee = tk.BooleanVar(value=True)
        self.var_other = tk.BooleanVar(value=True)

        tk.Checkbutton(type_row, text="👨‍💼 سداد تاجر",
                       variable=self.var_trader,
                       bg='#ecf0f1', fg='#c0392b',
                       font=('Arial', 11, 'bold'),
                       activebackground='#ecf0f1',
                       command=self.load_actions).pack(side='right', padx=5)

        tk.Checkbutton(type_row, text="👷 سداد عامل",
                       variable=self.var_employee,
                       bg='#ecf0f1', fg='#16a085',
                       font=('Arial', 11, 'bold'),
                       activebackground='#ecf0f1',
                       command=self.load_actions).pack(side='right', padx=5)

        tk.Checkbutton(type_row, text="💸 مصروف آخر",
                       variable=self.var_other,
                       bg='#ecf0f1', fg='#f39c12',
                       font=('Arial', 11, 'bold'),
                       activebackground='#ecf0f1',
                       command=self.load_actions).pack(side='right', padx=5)

        # ============ أزرار العمليات ============
        action_frame = tk.LabelFrame(content, text="إضافة حركة جديدة",
                                       font=('Arial', 11, 'bold'),
                                       bg='#ecf0f1', fg='#34495e',
                                       padx=5, pady=12)
        action_frame.pack(fill='x', padx=15, pady=(0, 5))

        row1 = tk.Frame(action_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="👨‍💼 سداد تاجر",
                  command=lambda: self.add_action('سداد تاجر'),
                  bg='#c0392b', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=18).pack(side='right', padx=5)

        tk.Button(row1, text="👷 سداد عامل",
                  command=lambda: self.add_action('سداد عامل'),
                  bg='#16a085', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=18).pack(side='right', padx=5)

        tk.Button(row1, text="💸 مصروف آخر",
                  command=lambda: self.add_action('مصروف آخر'),
                  bg='#f39c12', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=18).pack(side='right', padx=5)

        tk.Button(row1, text="🗑 حذف المحدد",
                  command=self.delete_action,
                  bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=15).pack(side='right', padx=5)

        tk.Button(row1, text="🔄 تحديث",
                  command=self.load_actions,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=12).pack(side='right', padx=5)

        # ============ جدول الحركات ============
        tree_frame = tk.LabelFrame(content, text="حركات المحفظة",
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#34495e',
                                     padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=5)

        columns = ('id', 'date', 'type', 'beneficiary',
                   'amount', 'fee', 'total', 'desc')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=15)

        self.tree.heading('id', text='ID')
        self.tree.heading('date', text='التاريخ')
        self.tree.heading('type', text='نوع الحركة')
        self.tree.heading('beneficiary', text='المستفيد')
        self.tree.heading('amount', text='المبلغ')
        self.tree.heading('fee', text='الرسوم')
        self.tree.heading('total', text='الإجمالي')
        self.tree.heading('desc', text='البيان')

        self.tree.column('id', width=40, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('type', width=130, anchor='center')
        self.tree.column('beneficiary', width=130, anchor='center')
        self.tree.column('amount', width=90, anchor='center')
        self.tree.column('fee', width=80, anchor='center')
        self.tree.column('total', width=100, anchor='center')
        self.tree.column('desc', width=180, anchor='e')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        self.tree.tag_configure('trader', background='#fadbd8')
        self.tree.tag_configure('employee', background='#d1f2eb')
        self.tree.tag_configure('other', background='#fef5e7')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # ============ شريط الإجماليات السفلي ============
        summary = tk.Frame(content, bg='#34495e')
        summary.pack(fill='x', padx=15, pady=(5, 15))

        self.lbl_total_trader = tk.Label(summary,
                                          text="سدادات تجار: 0.00",
                                          font=('Arial', 12, 'bold'),
                                          bg='#34495e', fg='#e74c3c',
                                          padx=15, pady=8)
        self.lbl_total_trader.pack(side='right', padx=10)

        self.lbl_total_emp = tk.Label(summary,
                                       text="سدادات عمال: 0.00",
                                       font=('Arial', 12, 'bold'),
                                       bg='#34495e', fg='#1abc9c',
                                       padx=15, pady=8)
        self.lbl_total_emp.pack(side='right', padx=10)

        self.lbl_total_other = tk.Label(summary,
                                         text="مصروفات أخرى: 0.00",
                                         font=('Arial', 12, 'bold'),
                                         bg='#34495e', fg='#f39c12',
                                         padx=15, pady=8)
        self.lbl_total_other.pack(side='right', padx=10)

        self.lbl_total_period = tk.Label(summary,
                                          text="إجمالي الفترة: 0.00",
                                          font=('Arial', 13, 'bold'),
                                          bg='#f1c40f', fg='#2c3e50',
                                          padx=20, pady=8)
        self.lbl_total_period.pack(side='left', padx=15)

        # التحميل الأولي
        self.load_actions()

    # ---------- المساعدات ----------
    def get_connection(self):
        return sqlite3.connect(DB_NAME)



    def add_action(self, action_type):
        VodafoneActionWindow(self, action_type)

    def on_select(self, event):
        sel = self.tree.selection()
        self.selected_id = sel[0] if sel else None  # ✅ iid نصي

    def delete_action(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر حركة من الجدول أولاً")
            return

        iid = self.selected_id
        source = iid[:1]  # 'w' أو 'p'
        real_id = int(iid[2:])

        if source == 'w':
            msg = "هل أنت متأكد من حذف هذه الحركة اليدوية؟"
        else:
            msg = ("هذه الحركة مسجّلة من صفحة المصاريف.\n\n"
                   "هل تريد حذفها من صفحة المصاريف أيضاً؟\n"
                   "(سيؤثر ذلك على إجمالي المصاريف)")

        if not messagebox.askyesno("تأكيد", msg):
            return

        try:
            conn = self.get_connection()
            c = conn.cursor()

            if source == 'w':
                c.execute("DELETE FROM vodafone_wallet_actions WHERE id=?",
                          (real_id,))
            else:
                c.execute("DELETE FROM paids WHERE id=?", (real_id,))

            conn.commit()
            conn.close()
            messagebox.showinfo("نجاح", "✅ تم حذف الحركة")
            self.selected_id = None
            self.load_actions()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")

    # ---------- التحميل ----------
    def load_actions(self):
        """تحميل الحركات من المصدرين + تحديث الرصيد"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        # 1. الرصيد الحالي
        balance_info = get_vodafone_wallet_balance()

        self.lbl_balance.config(
            text=f"{balance_info['balance']:,.2f} ج.م",
            fg='white' if balance_info['balance'] >= 0 else '#f5b7b1')
        self.lbl_income.config(
            text=f"إجمالي الإيراد: {balance_info['income']:,.2f}")
        self.lbl_expenses.config(
            text=f"إجمالي المصاريف: {balance_info['expenses']:,.2f}")

        # 2. تحديد الأنواع المفعلة
        selected_types = []
        if self.var_trader.get():
            selected_types.append('سداد تاجر')
        if self.var_employee.get():
            selected_types.append('سداد عامل')
        if self.var_other.get():
            selected_types.append('مصروف آخر')

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()

        totals = {'سداد تاجر': 0, 'سداد عامل': 0, 'مصروف آخر': 0}

        # قائمة موحّدة
        all_rows = []

        if selected_types:
            conn = self.get_connection()
            c = conn.cursor()

            # 1. حركات يدوية من vodafone_wallet_actions (بدون الرسوم)
            placeholders = ','.join('?' * len(selected_types))
            query = f"""SELECT id, date, action_type, beneficiary,
                               amount, description
                        FROM vodafone_wallet_actions
                        WHERE date >= ? AND date <= ?
                        AND action_type IN ({placeholders})
                        ORDER BY date DESC, id DESC"""
            params = [date_from, date_to] + selected_types
            c.execute(query, params)
            for rid, date, atype, benef, amt, desc in c.fetchall():
                all_rows.append({
                    'iid': f'w_{rid}', 'source': 'wallet',
                    'id': rid, 'date': date, 'type': atype,
                    'beneficiary': benef or '',
                    'amount': amt or 0, 'fee': 0,
                    'desc': desc or ''
                })

            # 2. مصاريف من paids
            c.execute("""SELECT id, date, sub_name, main_name,
                                description, value
                         FROM paids
                         WHERE date >= ? AND date <= ?
                         AND payment_method = 'فودافون كاش'
                         ORDER BY date DESC, id DESC""",
                      (date_from, date_to))
            for rid, date, sub_name, main_name, desc, value in c.fetchall():
                if main_name == 'مشتريات':
                    atype = 'سداد تاجر'
                elif main_name == 'أجور ومرتبات':
                    atype = 'سداد عامل'
                else:
                    atype = 'مصروف آخر'

                # ✅ جلب الرسوم المرتبطة
                fee = get_fee_for_paid(rid)

                all_rows.append({
                    'iid': f'p_{rid}', 'source': 'paids',
                    'id': rid, 'date': date, 'type': atype,
                    'beneficiary': sub_name or main_name or '',
                    'amount': value or 0, 'fee': fee,
                    'desc': desc or ''
                })

            conn.close()

        all_rows.sort(key=lambda r: (r['date'] or '', r['iid']), reverse=True)

        # عرض الجدول
        for r in all_rows:
            atype = r['type']
            tag = 'trader' if 'تاجر' in atype else \
                'employee' if 'عامل' in atype else 'other'
            source_icon = '📱' if r['source'] == 'paids' else '✏'
            total_row = r['amount'] + r['fee']

            self.tree.insert('', 'end', iid=r['iid'], tags=(tag,), values=(
                r['id'],
                r['date'],
                f"{source_icon} {atype}",
                r['beneficiary'],
                f"{r['amount']:,.2f}",
                f"{r['fee']:,.2f}",
                f"{total_row:,.2f}",
                r['desc']))

            if atype in totals:
                totals[atype] += total_row

        # 7. تحديث الإجماليات
        total_period = sum(totals.values())
        self.lbl_total_trader.config(
            text=f"سدادات تجار: {totals['سداد تاجر']:,.2f}")
        self.lbl_total_emp.config(
            text=f"سدادات عمال: {totals['سداد عامل']:,.2f}")
        self.lbl_total_other.config(
            text=f"مصروفات أخرى: {totals['مصروف آخر']:,.2f}")
        self.lbl_total_period.config(
            text=f"إجمالي الفترة: {total_period:,.2f}")


class SaleFormWindow:
    """نافذة إدخال / تعديل بيانات المبيعات"""
    def __init__(self, parent_page, sale_id=None):
        self.parent_page = parent_page
        self.sale_id = sale_id
        self.is_edit = sale_id is not None
        self.image_path = None
        self.image_preview = None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل مبيعات" if self.is_edit else "إضافة مبيعات جديدة"
        self.window.title(title)
        self.window.geometry("760x620")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ============ العنوان ============
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#2c3e50', fg='white').pack(pady=15)

        # ============ الإطار الرئيسي ============
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # --- التاريخ ---
        tk.Label(main, text="التاريخ: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(row=0, column=0,
                                                 padx=5, pady=8, sticky='e')
        self.entry_date = DateEntry(main, width=18, date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11), justify='center')
        self.entry_date.grid(row=0, column=1, padx=5, pady=8, sticky='w')
        self.entry_date.set_date(datetime.now())

        # --- المبيعات النقدية والعملات المحلية ---
        tk.Label(main, text="نقدي (ج.م):", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#27ae60', anchor='e').grid(
                    row=1, column=0, padx=5, pady=8, sticky='e')
        self.entry_pound = tk.Entry(main, width=20, font=('Arial', 12),
                                     justify='center', bg='#eafaf1')
        self.entry_pound.grid(row=1, column=1, padx=5, pady=8, sticky='w')
        self.entry_pound.insert(0, "0")

        tk.Label(main, text="فودافون كاش:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#e67e22', anchor='e').grid(
                    row=2, column=0, padx=5, pady=8, sticky='e')
        self.entry_vodafone = tk.Entry(main, width=20, font=('Arial', 12),
                                        justify='center', bg='#fef5e7')
        self.entry_vodafone.grid(row=2, column=1, padx=5, pady=8, sticky='w')
        self.entry_vodafone.insert(0, "0")

        tk.Label(main, text="فيزا:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#8e44ad', anchor='e').grid(
                    row=3, column=0, padx=5, pady=8, sticky='e')
        self.entry_visa = tk.Entry(main, width=20, font=('Arial', 12),
                                    justify='center', bg='#f4ecf7')
        self.entry_visa.grid(row=3, column=1, padx=5, pady=8, sticky='w')
        self.entry_visa.insert(0, "0")

        # --- العملات الأجنبية ---
        tk.Label(main, text="دولار ($):", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#16a085', anchor='e').grid(
                    row=1, column=2, padx=5, pady=8, sticky='e')
        self.entry_dollar = tk.Entry(main, width=20, font=('Arial', 12),
                                      justify='center', bg='#e8f8f5')
        self.entry_dollar.grid(row=1, column=3, padx=5, pady=8, sticky='w')
        self.entry_dollar.insert(0, "0")

        tk.Label(main, text="سعر الدولار:", font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(row=2, column=2, padx=5,
                                                 pady=8, sticky='e')
        self.entry_dollar_rate = tk.Entry(main, width=20, font=('Arial', 12),
                                           justify='center', bg='#d5f4e6')
        self.entry_dollar_rate.grid(row=2, column=3, padx=5, pady=8, sticky='w')
        self.entry_dollar_rate.insert(0, "48")

        tk.Label(main, text="يورو (€):", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#2980b9', anchor='e').grid(
                    row=3, column=2, padx=5, pady=8, sticky='e')
        self.entry_euro = tk.Entry(main, width=20, font=('Arial', 12),
                                    justify='center', bg='#e8f4f8')
        self.entry_euro.grid(row=3, column=3, padx=5, pady=8, sticky='w')
        self.entry_euro.insert(0, "0")

        tk.Label(main, text="سعر اليورو:", font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(row=4, column=2, padx=5,
                                                 pady=8, sticky='e')
        self.entry_euro_rate = tk.Entry(main, width=20, font=('Arial', 12),
                                         justify='center', bg='#d6eaf8')
        self.entry_euro_rate.grid(row=4, column=3, padx=5, pady=8, sticky='w')
        self.entry_euro_rate.insert(0, "51")

        # --- الإجمالي الفوري ---
        tk.Label(main, text="الإجمالي المتوقع:", font=('Arial', 12, 'bold'),
                 bg='#ecf0f1', fg='#c0392b', anchor='e').grid(
                    row=5, column=0, columnspan=2, padx=5, pady=15, sticky='e')
        self.lbl_row_total = tk.Label(main, text="0.00 ج.م",
                                       font=('Arial', 16, 'bold'),
                                       bg='#f1c40f', fg='#2c3e50',
                                       width=18, relief='ridge', bd=2)
        self.lbl_row_total.grid(row=5, column=2, columnspan=2,
                                 padx=5, pady=15, sticky='w')

        # ربط الحساب الفوري بكل الحقول
        for e in (self.entry_pound, self.entry_vodafone, self.entry_visa,
                  self.entry_dollar, self.entry_euro,
                  self.entry_dollar_rate, self.entry_euro_rate):
            e.bind('<KeyRelease>', lambda ev: self.update_total())

        # ============ الصورة ============
        img_frame = tk.LabelFrame(main, text="صورة الفاتورة (اختياري)",
                                    font=('Arial', 10, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=8, pady=8)
        img_frame.grid(row=0, column=4, rowspan=6, padx=15, pady=5, sticky='n')

        img_container = tk.Frame(img_frame, width=140, height=140,
                                  bg='lightgray', relief='ridge', bd=2)
        img_container.pack()
        img_container.grid_propagate(False)

        self.image_label = tk.Label(img_container, bg='lightgray',
                                     text='\nلا توجد صورة\n',
                                     font=('Arial', 11))
        self.image_label.pack(fill='both', expand=True)
        self.image_label.bind("<Double-Button-1>", self.open_image)

        tk.Button(img_frame, text="📁 اختيار صورة",
                  command=self.select_image,
                  bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(pady=5)

        tk.Button(img_frame, text="🗑 إزالة الصورة",
                  command=self.remove_image,
                  bg='#e74c3c', fg='white', font=('Arial', 9),
                  padx=10, pady=3).pack()

        # ============ أزرار الحفظ ============
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        # تحميل البيانات إن كان تعديلاً
        if self.is_edit:
            self.load_data()
        else:
            self.update_total()

        self.entry_pound.focus()
        self.entry_pound.select_range(0, tk.END)

    # ---------- حساب الإجمالي الفوري ----------
    def update_total(self):
        try:
            cash = float(self.entry_pound.get() or 0)
            voda = float(self.entry_vodafone.get() or 0)
            visa = float(self.entry_visa.get() or 0)
            dollar = float(self.entry_dollar.get() or 0)
            euro = float(self.entry_euro.get() or 0)
            d_rate = float(self.entry_dollar_rate.get() or 0)
            e_rate = float(self.entry_euro_rate.get() or 0)
        except ValueError:
            self.lbl_row_total.config(text="أرقام فقط!", fg='#e74c3c')
            return

        total = cash + voda + visa + (dollar * d_rate) + (euro * e_rate)
        self.lbl_row_total.config(text=f"{total:,.2f} ج.م", fg='#2c3e50')

    # ---------- تحميل بيانات التعديل ----------
    def load_data(self):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("""SELECT date, pound, vodafone, visa, dollar, euro,
                            dollar_rate, euro_rate, image_path
                     FROM sales WHERE id=?""", (self.sale_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "الصف غير موجود")
            self.window.destroy()
            return

        date, pound, vodafone, visa, dollar, euro, d_rate, e_rate, image_path = row

        # التاريخ
        try:
            if date:
                self.entry_date.set_date(datetime.strptime(date, '%Y-%m-%d'))
        except Exception as e:
            print(f"⚠️ خطأ في التاريخ: {e}")

        # الحقول
        fields = [
            (self.entry_pound, pound),
            (self.entry_vodafone, vodafone),
            (self.entry_visa, visa),
            (self.entry_dollar, dollar),
            (self.entry_euro, euro),
            (self.entry_dollar_rate, d_rate),
            (self.entry_euro_rate, e_rate),
        ]
        for entry, val in fields:
            entry.delete(0, tk.END)
            entry.insert(0, '' if val in (None, '') else str(val))

        # الصورة
        image_path = clean_path(image_path)
        if image_path and os.path.exists(image_path):
            self.image_path = image_path
            self.show_preview(image_path)
        else:
            self.image_path = None

        self.update_total()

    # ---------- إدارة الصورة ----------
    def select_image(self):
        path = filedialog.askopenfilename(
            title="اختر الصورة",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                       ("All files", "*.*")])
        if path:
            self.image_path = path
            self.show_preview(path)

    def show_preview(self, path):
        try:
            img = Image.open(path)
            img.thumbnail((140, 140), Image.Resampling.LANCZOS)
            self.image_preview = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.image_preview, text='')
        except Exception as e:
            print(f"⚠️ خطأ في الصورة: {e}")

    def remove_image(self):
        if self.image_path:
            self.image_path = "REMOVE"
        self.image_label.config(image='', text='\nلا توجد صورة\n')
        self.image_preview = None

    def open_image(self, event=None):
        if not self.image_path or self.image_path == "REMOVE":
            return
        if not os.path.exists(self.image_path):
            messagebox.showerror("خطأ", "الملف غير موجود")
            return
        try:
            os.startfile(self.image_path)
        except:
            subprocess.Popen(['start', self.image_path], shell=True)

    # ---------- الحفظ ----------
    def save(self):
        # التاريخ
        date = self.entry_date.get().strip()
        if not date:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return

        # قراءة الأرقام
        try:
            pound = int(float(self.entry_pound.get() or 0))
            vodafone = int(float(self.entry_vodafone.get() or 0))
            visa = int(float(self.entry_visa.get() or 0))
            dollar = int(float(self.entry_dollar.get() or 0))
            euro = int(float(self.entry_euro.get() or 0))
            d_rate = float(self.entry_dollar_rate.get() or 0)
            e_rate = float(self.entry_euro_rate.get() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "القيم المالية يجب أن تكون أرقاماً")
            return

        # يجب إدخال قيمة واحدة على الأقل
        if pound == 0 and vodafone == 0 and visa == 0 and dollar == 0 and euro == 0:
            messagebox.showwarning("تنبيه", "أدخل قيمة واحدة على الأقل")
            return

        # التنفيذ
        if self.is_edit:
            ok = update_sale(self.sale_id, date, pound, vodafone, visa,
                             dollar, euro, d_rate, e_rate, self.image_path)
            msg_ok = "تم تعديل المبيعات"
            msg_err = "فشل تعديل المبيعات"
        else:
            ok = add_sale(date, pound, vodafone, visa, dollar, euro,
                          d_rate, e_rate, self.image_path)
            msg_ok = "تمت إضافة المبيعات"
            msg_err = "فشل إضافة المبيعات"

        if ok:
            messagebox.showinfo("نجاح", msg_ok)
            self.window.destroy()
            self.parent_page.load_sales()
        else:
            messagebox.showerror("خطأ", msg_err)


class salesPage(tk.Frame):
    """صفحة إدارة المبيعات (أزرار + بحث + جدول فقط)"""

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.selected_id = None

        # ═══════════════════════════════════════════════════════
        # 1️⃣ الشريط العلوي (اسم البرنامج + الشركة + اللوجو)
        # ═══════════════════════════════════════════════════════
        add_program_banner(self)

        # ═══════════════════════════════════════════════════════
        # 2️⃣ شريط الحالة (ثابت أسفل الصفحة)
        # ═══════════════════════════════════════════════════════
        self.status_bar = tk.Label(self, text="جاهز",
                                   font=('Arial', 10),
                                   bg='#2c3e50', fg='white',
                                   anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ═══════════════════════════════════════════════════════
        # 3️⃣ منطقة التمرير (كل العناصر داخلها)
        # ═══════════════════════════════════════════════════════
        content = make_scrollable(self)
        self._page_content = content

        # ============ العنوان ============
        tk.Label(content, text="📊 إدارة المبيعات",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=12)

        # ============ إطار الأزرار ============
        btn_frame = tk.LabelFrame(content, font=('Arial', 11, 'bold'),
                                  bg='#ecf0f1', fg='#34495e',
                                  padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة مبيعات", command=self.add_sale,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=16).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                  command=self.edit_sale,
                                  bg='#f39c12', fg='white',
                                  font=('Arial', 11, 'bold'),
                                  padx=20, pady=8, width=14,
                                  state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                    command=self.delete_sale,
                                    bg='#e74c3c', fg='white',
                                    font=('Arial', 11, 'bold'),
                                    padx=20, pady=8, width=14,
                                    state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث", command=self.load_sales,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=14).pack(side='left', padx=5)

        tk.Button(row1, text="📊 ملخص المبيعات",
                  command=self.show_summary,
                  bg='#2980b9', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=18).pack(side='left', padx=5)

        # ============ إطار البحث ============
        search_frame = tk.Frame(content, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=15, pady=5)

        tk.Label(search_frame, text="🔍 بحث:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_entry = tk.Entry(search_frame, width=20, font=('Arial', 11))
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind('<KeyRelease>',
                               lambda e: self.load_sales(self.search_entry.get().strip()))

        tk.Label(search_frame, text="من تاريخ:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)
        self.date_from = DateEntry(search_frame, width=12,
                                   date_pattern='yyyy-mm-dd',
                                   font=('Arial', 10))
        self.date_from.pack(side='left', padx=5)

        tk.Label(search_frame, text="إلى تاريخ:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)
        self.date_to = DateEntry(search_frame, width=12,
                                 date_pattern='yyyy-mm-dd',
                                 font=('Arial', 10))
        self.date_to.pack(side='left', padx=5)

        tk.Button(search_frame, text="🔎 فلترة",
                  command=self.filter_sales,
                  bg='#9b59b6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(search_frame, text="🗑 مسح",
                  command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(search_frame, text="💾 Excel",
                  command=self.export_to_excel,
                  bg='#16a085', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        tk.Button(search_frame, text="🖼️ طباعه",
                  command=self.export_to_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        # ============ جدول العرض ============
        tree_frame = tk.LabelFrame(content, text="سجل المبيعات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        columns = ('id', 'date', 'cash', 'vodafone', 'visa',
                   'dollar', 'd_rate', 'euro', 'e_rate', 'total', 'image')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=15)

        headers = {
            'id': 'ID', 'date': 'التاريخ',
            'cash': 'نقدي (ج.م)', 'vodafone': 'فودافون كاش',
            'visa': 'فيزا', 'dollar': 'دولار ($)',
            'd_rate': 'سعر $', 'euro': 'يورو (€)',
            'e_rate': 'سعر €', 'total': 'الإجمالي (ج.م)',
            'image': 'الصورة'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=40, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('cash', width=110, anchor='center')
        self.tree.column('vodafone', width=110, anchor='center')
        self.tree.column('visa', width=95, anchor='center')
        self.tree.column('dollar', width=95, anchor='center')
        self.tree.column('d_rate', width=80, anchor='center')
        self.tree.column('euro', width=95, anchor='center')
        self.tree.column('e_rate', width=80, anchor='center')
        self.tree.column('total', width=130, anchor='center')
        self.tree.column('image', width=140, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        # ═══════════════════════════════════════════════════════════
        # 🎨 تنسيق صفوف الجدول (Zebra + Hover)
        # ═══════════════════════════════════════════════════════════
        # 1) Zebra Stripes
        self.tree.tag_configure('odd_row', background='#eafaf1')
        self.tree.tag_configure('even_row', background='#ffffff')


        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', self.on_tree_double_click)

        # ============ شريط الإجماليات ============
        summary = tk.Frame(content, bg='#34495e', height=90)
        summary.pack(fill='x', padx=15, pady=(0, 5))
        summary.pack_propagate(False)

        row1 = tk.Frame(summary, bg='#34495e')
        row1.pack(fill='x', pady=6)

        self.lbl_cash = self._make_card(row1, "نقدي", "0", '#2ecc71')
        self.lbl_voda = self._make_card(row1, "فودافون", "0", '#e67e22')
        self.lbl_visa = self._make_card(row1, "فيزا", "0", '#9b59b6')
        self.lbl_dollar = self._make_card(row1, "دولار ($)", "0", '#16a085')
        self.lbl_euro = self._make_card(row1, "يورو (€)", "0", '#3498db')

        row2 = tk.Frame(summary, bg='#34495e')
        row2.pack(fill='x', pady=4)

        tk.Label(row1, text="💰 إجمالي المبيعات (ج.م):",
                 font=('Arial', 14, 'bold'),
                 bg='#34495e', fg='#f1c40f').pack(side='right', padx=20)

        self.lbl_grand = tk.Label(row2, text="0.00",
                                  font=('Arial', 18, 'bold'),
                                  bg='#f1c40f', fg='#2c3e50',
                                  padx=25, pady=3)
        self.lbl_grand.pack(side='left', padx=20)

        # ============ التحميل الأولي ============
        self.load_sales()

    # ---------- مساعدات ----------
    def _make_card(self, parent, title, value, color):
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8)
        tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                 bg='#34495e', fg=color).pack()
        lbl = tk.Label(card, text=value, font=('Arial', 14, 'bold'),
                       bg='#34495e', fg='white')
        lbl.pack()
        return lbl

    def _calc_row_total(self, row):
        cash = safe_float(row.get('pound', 0))
        voda = safe_float(row.get('vodafone', 0))
        visa = safe_float(row.get('visa', 0))
        dollar = safe_float(row.get('dollar', 0))
        euro = safe_float(row.get('euro', 0))
        d_rate = safe_float(row.get('dollar_rate', 0))
        e_rate = safe_float(row.get('euro_rate', 0))
        return cash + voda + visa + (dollar * d_rate) + (euro * e_rate)

    # ---------- التحميل والعرض ----------
    def load_sales(self, search=''):
        """تحميل كل الصفوف أو بحث عام"""
        if search:
            rows = search_sales(search, None)
        else:
            rows = get_all_sales()

        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"عدد السجلات: {len(rows)}")

    def filter_sales(self):
        """فلترة بالمدة"""
        search_term = self.search_entry.get().strip()
        d1 = self.date_from.get().strip()
        d2 = self.date_to.get().strip()

        rows = search_sales_filtered(search_term, d1, d2)
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(
            text=f"نتائج الفلترة من {d1} إلى {d2} | {len(rows)} سجل")

        if not rows:
            messagebox.showinfo("بحث", "لا توجد نتائج في هذه المدة")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        today = datetime.now()
        self.date_from.set_date(today.replace(day=1))
        self.date_to.set_date(today)
        self.load_sales()

    def _populate_tree(self, rows):
        """تحميل الصفوف مع تخطيط Zebra"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not rows:
            return

        # ─── إدراج الصفوف ───
        for i, row in enumerate(rows):
            total_egp = self._calc_row_total(row)

            img_path = row.get('image_path') or ''
            img_display = os.path.basename(img_path) if img_path else '---'

            values = (
                row.get('id', ''),
                row.get('date', ''),
                f"{safe_float(row.get('pound', 0)):,.0f}",
                f"{safe_float(row.get('vodafone', 0)):,.0f}",
                f"{safe_float(row.get('visa', 0)):,.0f}",
                f"{safe_float(row.get('dollar', 0)):,.2f}",
                f"{safe_float(row.get('dollar_rate', 0)):,.2f}",
                f"{safe_float(row.get('euro', 0)):,.2f}",
                f"{safe_float(row.get('euro_rate', 0)):,.2f}",
                f"{total_egp:,.2f}",
                img_display,
            )

            # ✅ تخطيط Zebra
            tag = 'odd_row' if i % 2 == 0 else 'even_row'

            self.tree.insert('', 'end',
                             iid=str(row.get('id')),
                             tags=(tag,),
                             values=values)

    def on_tree_double_click(self, event):
        """لو العمود اللي اتضغط عليه هو الصورة → افتحها، غير كده → افتح التعديل"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)  # "#11" للعمود الحادي عشر
        if col == '#11':  # عمود الصورة (الترتيب: id=1, date=2, ... image=11)
            self.open_row_image()
        else:
            self.edit_sale()

    def open_row_image(self, event=None):
        """فتح صورة الفاتورة للسجل المحدد"""
        if self.selected_id is None:
            return
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT image_path FROM sales WHERE id=?", (self.selected_id,))
        row = c.fetchone()
        conn.close()
        if not row or not row[0]:
            messagebox.showinfo("معلومة", "لا توجد صورة لهذا السجل")
            return
        path = clean_path(row[0])
        if not os.path.exists(path):
            messagebox.showerror("خطأ", "ملف الصورة غير موجود")
            return
        try:
            os.startfile(path)
        except:
            subprocess.Popen(['start', path], shell=True)

    def _update_totals(self, rows):
        t_cash = sum(safe_float(r.get('pound', 0)) for r in rows)
        t_voda = sum(safe_float(r.get('vodafone', 0)) for r in rows)
        t_visa = sum(safe_float(r.get('visa', 0)) for r in rows)
        t_dollar = sum(safe_float(r.get('dollar', 0)) for r in rows)
        t_euro = sum(safe_float(r.get('euro', 0)) for r in rows)
        t_grand = sum(self._calc_row_total(r) for r in rows)

        self.lbl_cash.config(text=f"{t_cash:,.0f}")
        self.lbl_voda.config(text=f"{t_voda:,.0f}")
        self.lbl_visa.config(text=f"{t_visa:,.0f}")
        self.lbl_dollar.config(text=f"{t_dollar:,.2f}")
        self.lbl_euro.config(text=f"{t_euro:,.2f}")
        self.lbl_grand.config(text=f"{t_grand:,.2f}")

    # ---------- التحديد ----------
    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    # ---------- CRUD ----------
    def add_sale(self):
        SaleFormWindow(self, sale_id=None)

    def edit_sale(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً من الجدول أولاً")
            return
        SaleFormWindow(self, sale_id=self.selected_id)

    def delete_sale(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً من الجدول أولاً")
            return

        if not messagebox.askyesno("تأكيد",
                f"هل أنت متأكد من حذف السجل رقم {self.selected_id}؟"):
            return

        if delete_sale(self.selected_id):
            messagebox.showinfo("نجاح", "تم حذف السجل")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_sales()
        else:
            messagebox.showerror("خطأ", "فشل حذف السجل")

    # ---------- الملخص ----------
    def show_summary(self):
        """ملخص المبيعات الحالية"""
        rows = get_all_sales()
        if not rows:
            messagebox.showinfo("ملخص", "لا توجد بيانات")
            return

        t_cash = sum(safe_float(r.get('pound', 0)) for r in rows)
        t_voda = sum(safe_float(r.get('vodafone', 0)) for r in rows)
        t_visa = sum(safe_float(r.get('visa', 0)) for r in rows)
        t_dollar = sum(safe_float(r.get('dollar', 0)) for r in rows)
        t_euro = sum(safe_float(r.get('euro', 0)) for r in rows)
        t_grand = sum(self._calc_row_total(r) for r in rows)

        msg = (
            f"📊 ملخص إجمالي المبيعات\n"
            f"{'─' * 40}\n"
            f"💵 نقدي (ج.م)      : {t_cash:>15,.2f}\n"
            f"📱 فودافون كاش     : {t_voda:>15,.2f}\n"
            f"💳 فيزا             : {t_visa:>15,.2f}\n"
            f"💵 دولار ($)        : {t_dollar:>15,.2f}\n"
            f"💶 يورو (€)         : {t_euro:>15,.2f}\n"
            f"{'─' * 40}\n"
            f"💰 الإجمالي (ج.م)  : {t_grand:>15,.2f}\n"
            f"{'─' * 40}\n"
            f"📄 عدد السجلات     : {len(rows)}"
        )
        messagebox.showinfo("ملخص المبيعات", msg)

    # ─────────────────────────────────────────────
    # تصدير نتائج الفلترة الحالية
    # ─────────────────────────────────────────────

    def _get_visible_rows(self):
        """جلب الصفوف الظاهرة حالياً في الجدول"""
        items = self.tree.get_children()
        rows = []
        for item in items:
            rows.append(self.tree.item(item)['values'])
        return rows

    def export_to_excel(self):
        """تصدير نتائج الفلترة إلى Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        rows_data = self._get_visible_rows()
        if not rows_data:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"تقرير_المبيعات_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            title="حفظ تقرير المبيعات Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "المبيعات"
            ws.sheet_view.rightToLeft = True

            headers = ['ID', 'التاريخ', 'نقدي (ج.م)', 'فودافون كاش', 'فيزا',
                       'دولار ($)', 'سعر $', 'يورو (€)', 'سعر €',
                       'الإجمالي (ج.م)']

            # ─── العنوان ───
            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=len(headers))
            c = ws.cell(row=1, column=1, value="📊 تقرير المبيعات")
            c.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            c.fill = PatternFill("solid", fgColor="2C3E50")
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ─── الرؤوس ───
            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=col_idx, value=h)
                cell.font = Font(name='Arial', size=11, bold=True,
                                 color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            # ─── البيانات ───
            row_num = 3
            for r in rows_data:
                for col_idx, val in enumerate(r, start=1):
                    cell = ws.cell(row=row_num, column=col_idx, value=val)
                    cell.font = Font(name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center',
                                               vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
                row_num += 1

            # ─── صف الإجمالي ───
            row_num += 1
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=len(headers) - 1)
            cell = ws.cell(row=row_num, column=1, value="الإجمالي العام")
            cell.font = Font(name='Arial', size=12, bold=True, color='8B4513')
            cell.fill = PatternFill("solid", fgColor="FFF3CD")
            cell.alignment = Alignment(horizontal='center', vertical='center')

            grand_txt = self.lbl_grand.cget('text')
            cell = ws.cell(row=row_num, column=len(headers), value=grand_txt)
            cell.font = Font(name='Arial', size=12, bold=True, color='C0392B')
            cell.fill = PatternFill("solid", fgColor="FFF3CD")
            cell.alignment = Alignment(horizontal='center', vertical='center')

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")
            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    def export_to_image(self):
        """تصدير نتائج الفلترة إلى صورة PNG"""
        rows_data = self._get_visible_rows()
        if not rows_data:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"تقرير_المبيعات_{datetime.now().strftime('%Y-%m-%d')}.png",
            title="حفظ تقرير المبيعات صورة"
        )
        if not file_path:
            return

        columns = ['ID', 'التاريخ', 'نقدي (ج.م)', 'فودافون كاش', 'فيزا',
                   'دولار ($)', 'سعر $', 'يورو (€)', 'سعر €',
                   'الإجمالي (ج.م)']

        # ─── الإجماليات أسفل الصورة ───
        summary_items = [
            ("عدد السجلات", str(len(rows_data))),
            ("إجمالي نقدي (ج.م)", self.lbl_cash.cget('text')),
            ("إجمالي فودافون", self.lbl_voda.cget('text')),
            ("إجمالي فيزا", self.lbl_visa.cget('text')),
            ("إجمالي دولار ($)", self.lbl_dollar.cget('text')),
            ("إجمالي يورو (€)", self.lbl_euro.cget('text')),
            ("─────", ""),
            ("💰 الإجمالي العام (ج.م)", self.lbl_grand.cget('text')),
        ]

        success = export_table_to_image(
            file_path=file_path,
            title="📊 تقرير المبيعات",
            subtitle=f"عدد السجلات: {len(rows_data)}",
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass


class PaidFormWindow:
    """نافذة إدخال / تعديل مصروف"""
    def __init__(self, parent_page, paid_id=None):
        self.parent_page = parent_page
        self.paid_id = paid_id
        self.is_edit = paid_id is not None
        self.sub2_var = tk.StringVar(value='سلفة')
        self.image_path = None
        self.image_preview = None
        self.main_var = tk.StringVar()
        self.sub_var = tk.StringVar()
        self.payment_var = tk.StringVar(value='نقدي')

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل مصروف" if self.is_edit else "إضافة مصروف جديد"
        self.window.title(title)
        self.window.geometry("700x620")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # العنوان
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        icon = "✏" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#2c3e50', fg='white').pack(pady=15)

        # الإطار الرئيسي
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # التاريخ
        tk.Label(main, text="التاريخ: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                    row=0, column=0, padx=5, pady=8, sticky='e')
        self.entry_date = DateEntry(main, width=22,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11), justify='center')
        self.entry_date.grid(row=0, column=1, padx=5, pady=8, sticky='w')
        self.entry_date.set_date(datetime.now())

        # التبويب الرئيسي
        tk.Label(main, text="التبويب الرئيسي: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                    row=1, column=0, padx=5, pady=8, sticky='e')

        main_frame = tk.Frame(main, bg='#ecf0f1')
        main_frame.grid(row=1, column=1, padx=5, pady=8, sticky='w')

        self.main_combobox = ttk.Combobox(main_frame,
                                            textvariable=self.main_var,
                                            width=18, font=('Arial', 11),
                                            state='readonly')
        self.main_combobox.pack(side='left')
        self.main_combobox.bind('<<ComboboxSelected>>',
                                 lambda e: self.update_sub_categories())

        tk.Button(main_frame, text="⚙", command=self.open_category_manager,
                  bg='#3498db', fg='white', width=3).pack(side='left', padx=3)

        # التبويب الفرعي (الأول + الثاني بجانب بعضهما)
        tk.Label(main, text="التبويب الفرعي:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
            row=2, column=0, padx=5, pady=8, sticky='e')

        sub_container = tk.Frame(main, bg='#ecf0f1')
        sub_container.grid(row=2, column=1, padx=5, pady=8, sticky='w')

        self.sub_combobox = ttk.Combobox(sub_container,
                                         textvariable=self.sub_var,
                                         width=12, font=('Arial', 11))
        self.sub_combobox.pack(side='left')

        # ✅ التبويب الفرعي الثاني (يظهر فقط لأجور ومرتبات)
        self.sub2_var = tk.StringVar(value='سلفة')
        self.sub2_combo = ttk.Combobox(sub_container,
                                       textvariable=self.sub2_var,
                                       values=['سلفة', 'باقي حساب'],
                                       state='readonly',
                                       width=12, font=('Arial', 11))
        self.sub2_combo.pack(side='left', padx=(10, 0))

        # المبلغ
        tk.Label(main, text="المبلغ (ج.م): *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#c0392b', anchor='e').grid(
                    row=3, column=0, padx=5, pady=8, sticky='e')
        self.entry_value = tk.Entry(main, width=22, font=('Arial', 13),
                                     justify='center', bg='#fadbd8')
        self.entry_value.grid(row=3, column=1, padx=5, pady=8, sticky='w')

        # طريقة الدفع
        tk.Label(main, text="طريقة الدفع: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#8e44ad', anchor='e').grid(
            row=4, column=0, padx=5, pady=8, sticky='e')
        self.payment_combo = ttk.Combobox(main, textvariable=self.payment_var,
                                          values=['نقدي', 'فودافون كاش', 'الخزينة'],
                                          state='readonly',
                                          width=20, font=('Arial', 11))

        self.payment_combo.grid(row=4, column=1, padx=5, pady=8, sticky='w')
        self.payment_combo.bind('<<ComboboxSelected>>',
                                self.on_payment_change)

        # ✅ حقل الرسوم (يظهر فقط لفودافون كاش)
        self.lbl_fee = tk.Label(main, text="رسوم التحويل (ج.م):",
                                font=('Arial', 11, 'bold'),
                                bg='#ecf0f1', fg='#e67e22', anchor='e')
        self.lbl_fee.grid(row=5, column=0, padx=5, pady=8, sticky='e')

        self.entry_fee = tk.Entry(main, width=22, font=('Arial', 12),
                                  justify='center', bg='#fef5e7')
        self.entry_fee.grid(row=5, column=1, padx=5, pady=8, sticky='w')
        self.entry_fee.insert(0, "0")

        # إخفاء مبدئي
        self.lbl_fee.grid_remove()
        self.entry_fee.grid_remove()

        # معلومة
        self.lbl_fee_info = tk.Label(main,
                                     text="💡 الرسوم تؤثر على رصيد محفظة فودافون فقط",
                                     font=('Arial', 9, 'italic'), bg='#ecf0f1', fg='#7f8c8d',
                                     anchor='w')
        self.lbl_fee_info.grid(row=6, column=0, columnspan=2,
                               padx=5, pady=(0, 5), sticky='w')
        self.lbl_fee_info.grid_remove()

        # الوصف
        tk.Label(main, text="البيان:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='ne').grid(
                    row=5, column=0, padx=5, pady=8, sticky='ne')
        self.entry_description = tk.Text(main, width=22, height=3,
                                          font=('Arial', 11))
        self.entry_description.grid(row=7, column=1, padx=5,
                                     pady=8, sticky='w')

        # الصورة
        img_frame = tk.LabelFrame(main, text="صورة الفاتورة",
                                    font=('Arial', 10, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=8, pady=8)
        img_frame.grid(row=0, column=2, rowspan=6, padx=15, pady=5, sticky='n')

        img_container = tk.Frame(img_frame, width=150, height=150,
                                  bg='lightgray', relief='ridge', bd=2)
        img_container.pack()
        img_container.grid_propagate(False)

        self.image_label = tk.Label(img_container, bg='lightgray',
                                     text='\nلا توجد صورة\n',
                                     font=('Arial', 11))
        self.image_label.pack(fill='both', expand=True)
        self.image_label.bind("<Double-Button-1>", self.open_image)

        tk.Button(img_frame, text="📁 اختيار صورة",
                  command=self.select_image,
                  bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(pady=5)

        tk.Button(img_frame, text="🗑 إزالة",
                  command=self.remove_image,
                  bg='#e74c3c', fg='white', font=('Arial', 9),
                  padx=10, pady=3).pack()

        # أزرار الحفظ
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=30, pady=12,
                  font=('Arial', 13, 'bold')).pack(side='left', padx=5)

        self.load_main_categories()
        if self.is_edit:
            self.load_data()
        self.entry_value.focus()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def on_payment_change(self, event=None):
        """إظهار/إخفاء حقل الرسوم"""
        if self.payment_var.get() == 'فودافون كاش':
            self.lbl_fee.grid()
            self.entry_fee.grid()
            self.lbl_fee_info.grid()
        else:
            self.lbl_fee.grid_remove()
            self.entry_fee.grid_remove()
            self.lbl_fee_info.grid_remove()
            self.entry_fee.delete(0, tk.END)
            self.entry_fee.insert(0, "0")

    def load_main_categories(self):
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT id, name FROM main_categories ORDER BY name")
        self.main_categories = c.fetchall()
        conn.close()
        names = [name for _, name in self.main_categories]
        self.main_combobox['values'] = names
        if names and not self.main_var.get():
            self.main_var.set(names[0])
        self.update_sub_categories()

    def update_sub_categories(self):
        main_name = self.main_var.get().strip()
        if not main_name:
            self.sub_combobox['values'] = []
            self.sub_var.set('')
            self.sub2_combo.pack_forget()
            return

        if main_name == 'مشتريات':
            conn = self.get_connection()
            c = conn.cursor()
            c.execute("SELECT name FROM traders ORDER BY name")
            names = [r[0] for r in c.fetchall()]
            conn.close()
            self.sub_combobox['values'] = names
            self.sub_var.set(names[0] if names else '')
            self.sub2_combo.pack_forget()
            return

        if main_name == 'أجور ومرتبات':
            conn = self.get_connection()
            c = conn.cursor()
            c.execute("SELECT name FROM employees ORDER BY name")
            names = [r[0] for r in c.fetchall()]
            conn.close()
            self.sub_combobox['values'] = names
            self.sub_var.set(names[0] if names else '')
            # ✅ إظهار حقل نوع الحركة
            self.sub2_combo.pack(side='left', padx=(10, 0))
            if not self.sub2_var.get():
                self.sub2_var.set('سلفة')
            return

        main_id = None
        for mid, mname in self.main_categories:
            if mname == main_name:
                main_id = mid
                break

        if main_id is None:
            self.sub_combobox['values'] = []
            self.sub_var.set('')
            self.sub2_combo.pack_forget()
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM sub_categories WHERE main_id=? ORDER BY name",
                  (main_id,))
        subs = [r[0] for r in c.fetchall()]
        conn.close()
        self.sub_combobox['values'] = subs
        self.sub_var.set(subs[0] if subs else '')
        self.sub2_combo.pack_forget()

    def open_category_manager(self):
        CategoryManager(self.parent_page)

    def load_data(self):
        """تحميل بيانات المصروف للتعديل (نسخة آمنة 100%)"""

        # ═════════ اتصال واحد لكل العمليات ═════════
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        # ─── فحص وجود sub_name2 ───
        has_sub2 = False
        try:
            c.execute("PRAGMA table_info(paids)")
            cols = [r[1] for r in c.fetchall()]
            has_sub2 = 'sub_name2' in cols
        except Exception as e:
            print(f"⚠️ فشل PRAGMA: {e}")

        # ─── الاستعلام الرئيسي ───
        try:
            if has_sub2:
                c.execute("""SELECT date, main_name, sub_name, description,
                                    value, payment_method, image_path, sub_name2
                             FROM paids WHERE id=?""", (self.paid_id,))
            else:
                c.execute("""SELECT date, main_name, sub_name, description,
                                    value, payment_method, image_path
                             FROM paids WHERE id=?""", (self.paid_id,))
            row = c.fetchone()
        except Exception as e:
            print(f"❌ خطأ في SELECT: {e}")
            row = None
        finally:
            # ✅ الإغلاق في النهاية فقط — بعد كل الاستعلامات
            conn.close()

        if not row:
            messagebox.showerror("خطأ", "الصف غير موجود")
            self.window.destroy()
            return

        # ─── استخراج البيانات ───
        if has_sub2:
            (date, main_name, sub_name, description, value,
             payment_method, image_path, sub_name2) = row
        else:
            (date, main_name, sub_name, description, value,
             payment_method, image_path) = row
            sub_name2 = None

        # ─── التاريخ ───
        try:
            if date:
                self.entry_date.set_date(datetime.strptime(date, '%Y-%m-%d'))
        except Exception as e:
            print(f"⚠️ خطأ في تحميل التاريخ: {e}")

        # ─── التصنيف الرئيسي ───
        main_name = (main_name or '').strip()
        if main_name:
            self.main_var.set(main_name)
            self.update_sub_categories()
            if sub_name:
                self.sub_var.set(sub_name.strip())

        # ─── المبلغ ───
        self.entry_value.delete(0, tk.END)
        if value is not None:
            self.entry_value.insert(0, f"{value}")
        else:
            self.entry_value.insert(0, "0")

        # ─── طريقة الدفع ───
        payment_method = (payment_method or 'نقدي').strip()
        if payment_method not in ['نقدي', 'فودافون كاش', 'الخزينة']:
            payment_method = 'نقدي'
        self.payment_var.set(payment_method)
        self.on_payment_change()

        # ─── نوع الحركة ───
        if hasattr(self, 'sub2_var'):
            if sub_name2:
                self.sub2_var.set(sub_name2)
            else:
                self.sub2_var.set('سلفة')

        # ─── الرسوم ───
        try:
            fee = get_fee_for_paid(self.paid_id)
        except:
            fee = 0
        if hasattr(self, 'entry_fee'):
            self.entry_fee.delete(0, tk.END)
            self.entry_fee.insert(0, str(fee or 0))

        # ─── البيان ───
        self.entry_description.delete('1.0', tk.END)
        if description:
            self.entry_description.insert('1.0', description)

        # ─── الصورة ───
        image_path = clean_path(image_path)
        if image_path and os.path.exists(image_path):
            self.image_path = image_path
            self.show_preview(image_path)
        else:
            self.image_path = None

    def select_image(self):
        path = filedialog.askopenfilename(
            title="اختر الصورة",
            filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                       ("All files", "*.*")])
        if path:
            self.image_path = path
            self.show_preview(path)

    def show_preview(self, path):
        try:
            img = Image.open(path)
            img.thumbnail((150, 150), Image.Resampling.LANCZOS)
            self.image_preview = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.image_preview, text='')
        except Exception as e:
            print(f"⚠️ خطأ في الصورة: {e}")

    def remove_image(self):
        if self.image_path:
            self.image_path = "REMOVE"
        self.image_label.config(image='', text='\nلا توجد صورة\n')
        self.image_preview = None

    def open_image(self, event=None):
        if not self.image_path or self.image_path == "REMOVE":
            return
        if not os.path.exists(self.image_path):
            messagebox.showerror("خطأ", "الملف غير موجود")
            return
        try:
            os.startfile(self.image_path)
        except:
            subprocess.Popen(['start', self.image_path], shell=True)

    def save(self):
        date = self.entry_date.get().strip()
        value_str = self.entry_value.get().strip()
        main_name = self.main_var.get().strip()
        sub_name = self.sub_var.get().strip()
        description = self.entry_description.get('1.0', 'end').strip()
        payment_method = self.payment_var.get().strip() or 'نقدي'

        if not date:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return
       # if not value_str:
          #  messagebox.showwarning("تنبيه", "المبلغ مطلوب")
          #  return
        if not main_name:
            messagebox.showwarning("تنبيه", "التبويب الرئيسي مطلوب")
            return

        try:
            value = float(value_str)
            if value <= -1:
                raise ValueError
        except ValueError:
            messagebox.showerror("خطأ", "المبلغ يجب أن يكون رقماً موجباً")
            return
        # ✅ قراءة نوع الحركة (فقط لأجور ومرتبات)
        sub_name2 = None
        if main_name == 'أجور ومرتبات':
            sub_name2 = self.sub2_var.get().strip()
            if not sub_name2:
                sub_name2 = 'سلفة'

        # ✅ قراءة رسوم التحويل (فقط لو فودافون كاش)
        transfer_fee = 0.0
        if payment_method == 'فودافون كاش':
            try:
                transfer_fee = float(self.entry_fee.get().strip() or 0)
                if transfer_fee < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("خطأ", "الرسوم يجب أن تكون رقماً موجباً")
                return

        # ✅ التنفيذ مع تمرير transfer_fee
        if self.is_edit:
            ok = update_paid(self.paid_id, date, main_name, sub_name,
                             description, value, payment_method,
                             transfer_fee, self.image_path,
                             sub_name2)  # ✅
            msg_ok = "تم تعديل المصروف"
            msg_err = "فشل تعديل المصروف"
        else:
            ok = add_paid(date, main_name, sub_name, description, value,
                          payment_method, transfer_fee, self.image_path,
                          sub_name2)  # ✅
            msg_ok = "تمت إضافة المصروف"
            msg_err = "فشل إضافة المصروف"

        if ok:
            messagebox.showinfo("نجاح", msg_ok)
            self.window.destroy()
            self.parent_page.load_paids()
        else:
            messagebox.showerror("خطأ", msg_err)


# 2. صفحة المصاريف (نموذج آخر)
# =============================================
# ═══════════════════════════════════════════════════════════════════
#                        PAIDS PAGE
# ═══════════════════════════════════════════════════════════════════
class paidsPage(tk.Frame):
    """صفحة إدارة المصاريف"""

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.selected_id = None
        self.selected_invoice_id = None

        # 1️⃣ الشريط العلوي
        add_program_banner(self)

        # 2️⃣ شريط الحالة (أسفل)
        self.status_bar = tk.Label(self, text="جاهز",
                                   font=('Arial', 10),
                                   bg='#2c3e50', fg='white',
                                   anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # 3️⃣ منطقة التمرير (تأخذ كل المتبقي)
        content = make_scrollable(self)
        self._page_content = content

        # ============ العنوان ============
        tk.Label(content, text="💸 إدارة المصاريف",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=8)

        # ============ إطار الأزرار (صف واحد) ============
        btn_frame = tk.LabelFrame(content, font=('Arial', 11, 'bold'),
                                  bg='#ecf0f1', fg='#34495e',
                                  padx=10, pady=8)
        btn_frame.pack(fill='x', padx=15, pady=(0, 8))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack()

        tk.Button(row1, text="➕ إضافة مصروف", command=self.add_paid,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=6, width=15).pack(side='left', padx=3)

        self.btn_edit = tk.Button(row1, text="✏ تعديل",
                                  command=self.edit_paid,
                                  bg='#f39c12', fg='white',
                                  font=('Arial', 11, 'bold'),
                                  padx=15, pady=6, width=12,
                                  state='disabled')
        self.btn_edit.pack(side='left', padx=3)

        self.btn_delete = tk.Button(row1, text="🗑 حذف",
                                    command=self.delete_paid,
                                    bg='#e74c3c', fg='white',
                                    font=('Arial', 11, 'bold'),
                                    padx=15, pady=6, width=12,
                                    state='disabled')
        self.btn_delete.pack(side='left', padx=3)

        tk.Button(row1, text="🔄 تحديث", command=self.load_paids,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=6, width=12).pack(side='left', padx=3)

        tk.Button(row1, text="⚙ إدارة التبويبات",
                  command=self.open_category_manager,
                  bg='#3498db', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=6, width=16).pack(side='left', padx=3)

        tk.Button(row1, text="📄 إضافة فاتورة",
                  command=self.open_invoice_window,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=6, width=15).pack(side='left', padx=3)

        # ✅ زر عرض الفاتورة (مخفي افتراضياً - لا pack)
        self.view_invoice_btn = tk.Button(row1, text="👁 عرض الفاتورة",
                                          command=self.view_invoice,
                                          bg='#8e44ad', fg='white',
                                          font=('Arial', 11, 'bold'),
                                          padx=15, pady=6, width=15)

        # ============ إطار البحث ============
        search_frame = tk.Frame(content, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=15, pady=4)

        tk.Label(search_frame, text="🔍 بحث:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=3)

        self.search_entry = tk.Entry(search_frame, width=15,
                                     font=('Arial', 11))
        self.search_entry.pack(side='left', padx=3)
        self.search_entry.bind('<KeyRelease>',
                               lambda e: self.load_paids(self.search_entry.get().strip()))

        tk.Label(search_frame, text="من:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=3)
        self.start_date_entry = DateEntry(search_frame, width=11,
                                          date_pattern='yyyy-mm-dd',
                                          font=('Arial', 10))
        self.start_date_entry.pack(side='left', padx=3)

        tk.Label(search_frame, text="إلى:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=3)
        self.end_date_entry = DateEntry(search_frame, width=11,
                                        date_pattern='yyyy-mm-dd',
                                        font=('Arial', 10))
        self.end_date_entry.pack(side='left', padx=3)

        tk.Label(search_frame, text="الرئيسي:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=3)
        self.search_main = ttk.Combobox(search_frame, width=10,
                                        font=('Arial', 11))
        self.search_main.pack(side='left', padx=3)

        tk.Label(search_frame, text="الفرعي:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=3)
        self.search_sub = ttk.Combobox(search_frame, width=10,
                                       font=('Arial', 11))
        self.search_sub.pack(side='left', padx=3)

        tk.Label(search_frame, text="الدفع:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=3)
        self.search_payment = ttk.Combobox(search_frame,
                                           values=['الكل', 'نقدي', 'فودافون كاش', 'الخزينة'],
                                           width=12, font=('Arial', 11),
                                           state='readonly')

        self.search_payment.pack(side='left', padx=3)
        self.search_payment.set('الكل')

        tk.Button(search_frame, text="🔎 فلترة",
                  command=self.filter_paids,
                  bg='#9b59b6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=8).pack(side='left', padx=3)

        tk.Button(search_frame, text="🗑 مسح",
                  command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50',
                  font=('Arial', 10, 'bold'),
                  padx=8).pack(side='left', padx=3)

        tk.Button(search_frame, text="💾 Excel",
                  command=self.export_to_excel,
                  bg='#16a085', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=8).pack(side='left', padx=3)

        tk.Button(search_frame, text="🖼️ طباعه",
                  command=self.export_to_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=8).pack(side='left', padx=3)

        self.load_search_combos()

        # ============ جدول العرض ============
        tree_frame = tk.LabelFrame(content, text="سجل المصاريف",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True,
                        padx=15, pady=(4, 4))

        columns = ('id', 'date', 'main_name', 'sub_name',
                   'description', 'value', 'payment_method', 'image_path')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=12)

        headers = {
            'id': 'ID', 'date': 'التاريخ',
            'main_name': 'الرئيسي', 'sub_name': 'الفرعي',
            'description': 'البيان', 'value': 'القيمة',
            'payment_method': 'طريقة الدفع', 'image_path': 'الصورة'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=40, anchor='center')
        self.tree.column('date', width=95, anchor='center')
        self.tree.column('main_name', width=95, anchor='center')
        self.tree.column('sub_name', width=110, anchor='center')
        self.tree.column('description', width=230, anchor='center')
        self.tree.column('value', width=90, anchor='center')
        self.tree.column('payment_method', width=100, anchor='center')
        self.tree.column('image_path', width=130, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        self.tree.tag_configure('cash_method', background='#eafaf1')
        self.tree.tag_configure('wallet_method', background='#fef5e7')
        self.tree.tag_configure('treasury_method', background='#fdebd0')  # ✅ جديد

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', self.on_tree_double_click)

        # ============ شريط الإجماليات ============
        summary = tk.Frame(content, bg='#34495e', height=85)
        summary.pack(fill='x', padx=15, pady=(0, 8))
        summary.pack_propagate(False)

        s_row1 = tk.Frame(summary, bg='#34495e')
        s_row1.pack(fill='x', pady=4)

        self.lbl_total_cash = self._make_card(s_row1, "نقدي", "0", '#2ecc71')
        self.lbl_total_treasury = self._make_card(s_row1, "خزينة", "0", '#f39c12')  # ✅ جديد
        self.lbl_total_wallet = self._make_card(s_row1, "فودافون كاش", "0", '#e67e22')
        self.lbl_total_invoice = self._make_card(s_row1, "من الفواتير", "0", '#8e44ad')
        self.lbl_count = self._make_card(s_row1, "عدد السجلات", "0", '#95a5a6')

        s_row2 = tk.Frame(summary, bg='#34495e')
        s_row2.pack(fill='x', pady=2)

        tk.Label(s_row1, text="💰 إجمالي المصاريف:",
                 font=('Arial', 13, 'bold'),
                 bg='#34495e', fg='#f1c40f').pack(side='right', padx=20)

        self.lbl_total_grand = tk.Label(s_row2, text="0.00",
                                        font=('Arial', 17, 'bold'),
                                        bg='#f1c40f', fg='#2c3e50',
                                        padx=20, pady=2)
        self.lbl_total_grand.pack(side='left', padx=20)

        # ============ التحميل الأولي ============
        self.load_paids()
        self.update_view_invoice_button_state()

    # ---------- المساعدات ----------
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def _make_card(self, parent, title, value, color):
        card = tk.Frame(parent, bg='#34495e')
        card.pack(side='right', expand=True, padx=8)
        tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                 bg='#34495e', fg=color).pack()
        lbl = tk.Label(card, text=value, font=('Arial', 14, 'bold'),
                       bg='#34495e', fg='white')
        lbl.pack()
        return lbl

    def load_search_combos(self):
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM main_categories ORDER BY name")
        main_names = [row[0] for row in c.fetchall()]
        c.execute("SELECT name FROM sub_categories ORDER BY name")
        sub_names = [row[0] for row in c.fetchall()]
        conn.close()
        self.search_main['values'] = [''] + main_names
        self.search_sub['values'] = [''] + sub_names
        self.search_main.set('')
        self.search_sub.set('')

    # ---------- التحميل ----------
    def load_paids(self, search=''):
        if search:
            rows = search_paids(search, None, None, None, None, None)
        else:
            rows = get_all_paids()

        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(text=f"عدد السجلات: {len(rows)}")

    def filter_paids(self):
        search_term = self.search_entry.get().strip()
        d1 = self.start_date_entry.get().strip()
        d2 = self.end_date_entry.get().strip()
        main_name = self.search_main.get().strip()
        sub_name = self.search_sub.get().strip()
        payment = self.search_payment.get().strip()

        rows = search_paids(search_term, d1, d2, main_name, sub_name, payment)
        self._populate_tree(rows)
        self._update_totals(rows)
        self.status_bar.config(
            text=f"فلترة: {len(rows)} سجل من {d1} إلى {d2}")

        if not rows:
            messagebox.showinfo("بحث", "لا توجد نتائج")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.search_main.set('')
        self.search_sub.set('')
        self.search_payment.set('الكل')
        today = datetime.now()
        self.start_date_entry.set_date(today.replace(day=1))
        self.end_date_entry.set_date(today)
        self.load_paids()

    def _populate_tree(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not rows:
            return

        for row in rows:
            rid = row.get('id', '')
            rdate = row.get('date', '') or ''
            rmain = row.get('main_name', '') or ''
            rsub = row.get('sub_name', '') or ''
            rdesc = row.get('description', '') or ''
            rvalue = row.get('value', 0) or 0
            rmethod = row.get('payment_method', '') or 'نقدي'
            rimage = row.get('image_path', '') or ''


            image_name = os.path.basename(rimage) if rimage else ''
            values = (rid, rdate, rmain, rsub, rdesc,
                      f"{rvalue:,.2f}" if rvalue else '0.00',
                      rmethod, image_name)

            if rmethod == 'فودافون كاش':
                tag = 'wallet_method'
            elif rmethod == 'الخزينة':
                tag = 'treasury_method'
            else:
                tag = 'cash_method'

            try:
                self.tree.insert('', 'end', iid=str(rid),
                                  values=values, tags=(tag,))
            except Exception as e:
                print(f"خطأ في إدراج {rid}: {e}")

    def _update_totals(self, rows):
        t_cash = 0
        t_wallet = 0
        t_treasury = 0  # ✅ جديد
        t_invoice = 0
        t_grand = 0

        for r in rows:
            val = safe_float(r.get('value', 0))
            method = r.get('payment_method') or 'نقدي'
            inv_id = r.get('invoice_id') or 0

            t_grand += val

            if method == 'فودافون كاش':
                t_wallet += val
            elif method == 'الخزينة':  # ✅ جديد
                t_treasury += val
            else:
                t_cash += val

            if inv_id:
                t_invoice += val

        self.lbl_total_cash.config(text=f"{t_cash:,.2f}")
        self.lbl_total_treasury.config(text=f"{t_treasury:,.2f}")  # ✅ جديد
        self.lbl_total_wallet.config(text=f"{t_wallet:,.2f}")
        self.lbl_total_invoice.config(text=f"{t_invoice:,.2f}")
        self.lbl_count.config(text=str(len(rows)))
        self.lbl_total_grand.config(text=f"{t_grand:,.2f}")

    # ---------- التحديد ----------
    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                self.update_view_invoice_button_state()  # ✅
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

        # ✅ تُستدعى دائماً في نهاية الدالة (تحديد أو إلغاء)
        self.update_view_invoice_button_state()

    def on_tree_double_click(self, event):
        """لو العمود اللي اتضغط عليه هو الصورة → افتحها، غير كده → افتح التعديل"""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        col = self.tree.identify_column(event.x)
        # ترتيب الأعمدة: id=#1, date=#2, main=#3, sub=#4,
        # description=#5, value=#6, payment=#7, image=#8
        if col == '#8':
            self.open_row_image()
        else:
            self.edit_paid()

    def open_row_image(self):
        """فتح صورة القيد المحدد"""
        sel = self.tree.selection()
        if not sel:
            return

        try:
            paid_id = int(sel[0])
        except (ValueError, TypeError):
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT image_path FROM paids WHERE id=?", (paid_id,))
        row = c.fetchone()
        conn.close()

        if not row or not row[0]:
            messagebox.showinfo("معلومة", "لا توجد صورة لهذا القيد")
            return

        path = clean_path(row[0])
        if not path or not os.path.exists(path):
            messagebox.showerror("خطأ", "ملف الصورة غير موجود")
            return

        try:
            os.startfile(path)
        except:
            try:
                subprocess.Popen(['start', path], shell=True)
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل فتح الصورة:\n{e}")

    def update_view_invoice_button_state(self):
        """إظهار زر (عرض الفاتورة) فقط عند تحديد قيد مرتبط بفاتورة"""
        if self.selected_id is not None:
            conn = self.get_connection()
            c = conn.cursor()
            c.execute("SELECT invoice_id FROM paids WHERE id=?",
                      (self.selected_id,))
            result = c.fetchone()
            conn.close()

            if result and result[0]:
                self.selected_invoice_id = result[0]
                if not self.view_invoice_btn.winfo_ismapped():
                    self.view_invoice_btn.pack(side='left', padx=5)
                return

        # إخفاء تام للزر
        self.selected_invoice_id = None
        self.view_invoice_btn.pack_forget()

    # ---------- CRUD ----------
    def add_paid(self):
        PaidFormWindow(self, paid_id=None)

    def edit_paid(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً من الجدول أولاً")
            return
        PaidFormWindow(self, paid_id=self.selected_id)

    def delete_paid(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر سجلاً أولاً")
            return
        if messagebox.askyesno("تأكيد",
                f"هل أنت متأكد من حذف السجل رقم {self.selected_id}؟"):
            if delete_paid(self.selected_id):
                messagebox.showinfo("نجاح", "تم الحذف")
                self.selected_id = None
                self.btn_edit.config(state='disabled')
                self.btn_delete.config(state='disabled')
                self.load_paids()
            else:
                messagebox.showerror("خطأ", "فشل الحذف")

    # ---------- الفواتير والتقارير ----------
    def open_invoice_window(self):
        InvoiceWindow(self)

    def view_invoice(self):
        if self.selected_invoice_id is None:
            messagebox.showwarning("تنبيه", "الفاتورة غير موجودة")
            return
        InvoiceEditWindow(self, self.selected_invoice_id,
                          on_save_callback=self.load_paids)



    def open_category_manager(self):
        CategoryManager(self)

    # ---------- التصدير ----------
    def export_to_excel(self):
        search_term = self.search_entry.get().strip()
        d1 = self.start_date_entry.get().strip()
        d2 = self.end_date_entry.get().strip()
        main_name = self.search_main.get().strip()
        sub_name = self.search_sub.get().strip()
        payment = self.search_payment.get().strip()

        if not d1 or not d2:
            messagebox.showwarning("تنبيه", "حدد تاريخ البداية والنهاية")
            return

        rows = search_paids(search_term, d1, d2, main_name, sub_name, payment)
        if not rows:
            messagebox.showwarning("تنبيه", "لا توجد بيانات")
            return

        total = sum((r.get('value') or 0) for r in rows)

        data_rows = []
        for r in rows:
            data_rows.append([
                r.get('date', ''),
                r.get('main_name', ''),
                r.get('sub_name', ''),
                r.get('description', ''),
                r.get('payment_method', 'نقدي'),
                r.get('value', 0)
            ])
        data_rows.append(["الإجمالي", "", "", "", "", total])

        headers = ["التاريخ", "الرئيسي", "الفرعي", "البيان",
                   "طريقة الدفع", "القيمة"]

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"تقرير_المصاريف_{d1}_{d2}.xlsx"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "المصاريف"
            ws.sheet_view.rightToLeft = True
            ws.append(headers)
            for row in data_rows:
                ws.append(row)
            wb.save(file_path)
            messagebox.showinfo("نجاح", f"تم التصدير إلى:\n{file_path}")
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل التصدير: {e}")
    def export_to_image(self):
        """تصدير نتائج فلتر البحث الحالي إلى صورة PNG"""

        # ─── قراءة نفس فلاتر البحث ───
        search_term = self.search_entry.get().strip()
        d1 = self.start_date_entry.get().strip()
        d2 = self.end_date_entry.get().strip()
        main_name = self.search_main.get().strip()
        sub_name = self.search_sub.get().strip()
        payment = self.search_payment.get().strip()

        # ─── جلب نفس البيانات المعروضة في الجدول ───
        rows = search_paids(search_term, d1, d2, main_name, sub_name, payment)

        if not rows:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        # ─── اختيار مسار الحفظ ───
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"تقرير_المصاريف_{d1}_{d2}.png",
            title="حفظ تقرير المصاريف صورة"
        )
        if not file_path:
            return

        # ─── تجهيز الصفوف ───
        columns = ['ID', 'التاريخ', 'الرئيسي', 'الفرعي',
                   'البيان', 'القيمة', 'طريقة الدفع']

        rows_data = []
        total_cash = 0
        total_wallet = 0
        total_grand = 0

        for r in rows:
            rid = r.get('id', '')
            rdate = r.get('date', '') or ''
            rmain = r.get('main_name', '') or ''
            rsub = r.get('sub_name', '') or ''
            rdesc = r.get('description', '') or ''
            rvalue = safe_float(r.get('value', 0))
            rmethod = r.get('payment_method', '') or 'نقدي'

            rows_data.append((
                str(rid),
                rdate,
                rmain,
                rsub,
                rdesc[:60],   # قص البيان الطويل
                f"{rvalue:,.2f}",
                rmethod
            ))

            total_grand += rvalue
            if rmethod == 'فودافون كاش':
                total_wallet += rvalue
            else:
                total_cash += rvalue

        # ─── الإجماليات أسفل الصورة ───
        summary_items = [
            ("عدد السجلات", str(len(rows_data))),
            ("إجمالي نقدي", f"{total_cash:,.2f}"),
            ("إجمالي فودافون كاش", f"{total_wallet:,.2f}"),
            ("─────", ""),
            ("💰 إجمالي المصاريف", f"{total_grand:,.2f}"),
        ]

        # ─── العنوان الفرعي ───
        subtitle_parts = []
        if d1 or d2:
            subtitle_parts.append(f"من {d1 or '...'} إلى {d2 or '...'}")
        if main_name:
            subtitle_parts.append(f"الرئيسي: {main_name}")
        if sub_name:
            subtitle_parts.append(f"الفرعي: {sub_name}")
        if payment and payment != 'الكل':
            subtitle_parts.append(f"طريقة الدفع: {payment}")
        if search_term:
            subtitle_parts.append(f"بحث: {search_term}")

        subtitle = " | ".join(subtitle_parts) if subtitle_parts else "كل المصاريف"

        # ─── استدعاء دالة التصدير المشتركة ───
        success = export_table_to_image(
            file_path=file_path,
            title="💸 تقرير المصاريف",
            subtitle=subtitle,
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                    f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

##########    report_sales_page  #################
class PeriodReportPage(tk.Frame):
    """صفحة تقرير مدة (من تاريخ إلى تاريخ)"""
    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        # 1️⃣ الشريط العلوي
        add_program_banner(self)

        # 2️⃣ منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ═════════ العنوان ═════════
        tk.Label(content, text="📊 تقرير مدة",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=12)

        # ═════════ فلاتر التاريخ ═════════
        filter_frame = tk.LabelFrame(content, text="اختر المدة",
                                       font=('Arial', 11, 'bold'),
                                       bg='#ecf0f1', fg='#34495e',
                                       padx=15, pady=12)
        filter_frame.pack(fill='x', padx=15, pady=(0, 10))

        # من تاريخ
        tk.Label(filter_frame, text="من تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(side='right', padx=5)

        self.date_from = DateEntry(filter_frame, width=13,
                                     date_pattern='yyyy-mm-dd',
                                     font=('Arial', 11))
        self.date_from.pack(side='right', padx=5)
        # افتراضياً: أول يوم من الشهر الحالي
        today = datetime.now()
        first_of_month = today.replace(day=1)
        self.date_from.set_date(first_of_month)

        # إلى تاريخ
        tk.Label(filter_frame, text="إلى تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(side='right', padx=5)

        self.date_to = DateEntry(filter_frame, width=13,
                                   date_pattern='yyyy-mm-dd',
                                   font=('Arial', 11))
        self.date_to.pack(side='right', padx=5)
        self.date_to.set_date(today)

        # أزرار سريعة
        quick_btns = tk.Frame(filter_frame, bg='#ecf0f1')
        quick_btns.pack(side='right', padx=10)

        tk.Button(quick_btns, text="اليوم", command=self.set_today,
                  bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)
        tk.Button(quick_btns, text="الشهر", command=self.set_this_month,
                  bg='#16a085', fg='white', font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)
        tk.Button(quick_btns, text="آخر 7 أيام", command=self.set_last_7_days,
                  bg='#f39c12', fg='white', font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)
        tk.Button(quick_btns, text="آخر 30 يوم", command=self.set_last_30_days,
                  bg='#8e44ad', fg='white', font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)

        # أزرار العرض والتصدير
        action_frame = tk.Frame(content, bg='#ecf0f1')
        action_frame.pack(fill='x', padx=15, pady=5)

        tk.Button(action_frame, text="🔍 عرض التقرير",
                  command=self.load_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=25, pady=10).pack(side='right', padx=5)

        tk.Button(action_frame, text="🖼️ طباعه",
                  command=self.export_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=25, pady=10).pack(side='right', padx=5)

        tk.Button(action_frame, text="💾 Excel",
                  command=self.export_excel,
                  bg='#2980b9', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=25, pady=10).pack(side='right', padx=5)

        # ═════════ جدول التقرير ═════════
        tree_frame = tk.LabelFrame(content, text="تفاصيل المدة",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        # ✅ الترتيب الجديد: اليورو والدولار في النهاية
        columns = ('date', 'day', 'sales', 'expenses', 'net', 'euro', 'dollar')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=18)

        self.tree.heading('date', text='التاريخ')
        self.tree.heading('day', text='اليوم')
        self.tree.heading('sales', text='المبيعات (ج.م)')
        self.tree.heading('expenses', text='المصروفات')
        self.tree.heading('net', text='صافي الإيراد')
        self.tree.heading('euro', text='يورو (€)')
        self.tree.heading('dollar', text='دولار ($)')

        self.tree.column('date', width=110, anchor='center')
        self.tree.column('day', width=90, anchor='center')
        self.tree.column('sales', width=120, anchor='center')
        self.tree.column('expenses', width=110, anchor='center')
        self.tree.column('net', width=120, anchor='center')
        self.tree.column('euro', width=100, anchor='center')
        self.tree.column('dollar', width=100, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        # تنسيقات
        self.tree.tag_configure('positive', background='#eafaf1')
        self.tree.tag_configure('negative', background='#fdedec')
        self.tree.tag_configure('total_row', background='#2c3e50',
                                foreground='#f1c40f',
                                font=('Arial', 11, 'bold'))

        # ═════════ شريط الإجماليات ═════════
        summary = tk.Frame(content, bg='#34495e', height=70)
        summary.pack(fill='x', padx=15, pady=(0, 15))
        summary.pack_propagate(False)

        self.lbl_days = tk.Label(summary, text="عدد الأيام: 0",
                                   font=('Arial', 12, 'bold'),
                                   bg='#34495e', fg='white')
        self.lbl_days.pack(side='right', padx=15, pady=20)

        self.lbl_total_sales = tk.Label(summary, text="إجمالي المبيعات: 0.00",
                                          font=('Arial', 12, 'bold'),
                                          bg='#34495e', fg='#f1c40f')
        self.lbl_total_sales.pack(side='right', padx=15, pady=20)

        self.lbl_total_exp = tk.Label(summary, text="إجمالي المصروفات: 0.00",
                                        font=('Arial', 12, 'bold'),
                                        bg='#34495e', fg='#e74c3c')
        self.lbl_total_exp.pack(side='right', padx=15, pady=20)

        self.lbl_total_net = tk.Label(summary, text="إجمالي الصافي: 0.00",
                                        font=('Arial', 14, 'bold'),
                                        bg='#34495e', fg='#2ecc71')
        self.lbl_total_net.pack(side='left', padx=20, pady=20)

        # شريط الحالة
        self.status_bar = tk.Label(content, text="جاهز",
                                     font=('Arial', 10),
                                     bg='#2c3e50', fg='white',
                                     anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # تحميل أولي
        self.load_report()

    # ────────────────────────────────────────────────
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def set_today(self):
        today = datetime.now()
        self.date_from.set_date(today)
        self.date_to.set_date(today)
        self.load_report()

    def set_this_month(self):
        today = datetime.now()
        self.date_from.set_date(today.replace(day=1))
        self.date_to.set_date(today)
        self.load_report()

    def set_last_7_days(self):
        from datetime import timedelta
        today = datetime.now()
        self.date_from.set_date(today - timedelta(days=6))
        self.date_to.set_date(today)
        self.load_report()

    def set_last_30_days(self):
        from datetime import timedelta
        today = datetime.now()
        self.date_from.set_date(today - timedelta(days=29))
        self.date_to.set_date(today)
        self.load_report()

    # ────────────────────────────────────────────────
    def get_day_name_ar(self, date_str):
        """إرجاع اسم اليوم بالعربية"""
        try:
            d = datetime.strptime(date_str, '%Y-%m-%d')
            days_ar = {
                0: 'الإثنين', 1: 'الثلاثاء', 2: 'الأربعاء',
                3: 'الخميس', 4: 'الجمعة', 5: 'السبت', 6: 'الأحد'
            }
            return days_ar.get(d.weekday(), '')
        except:
            return ''

    def load_report(self):
        """تحميل التقرير"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()

        if not date_from or not date_to:
            messagebox.showwarning("تنبيه", "اختر تاريخ البداية والنهاية")
            return

        try:
            d1 = datetime.strptime(date_from, '%Y-%m-%d')
            d2 = datetime.strptime(date_to, '%Y-%m-%d')
            if d2 < d1:
                messagebox.showerror("خطأ", "تاريخ النهاية يجب أن يكون بعد البداية")
                return
        except:
            messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة")
            return

        conn = self.get_connection()
        c = conn.cursor()

        from datetime import timedelta
        days_list = []
        current = d1
        while current <= d2:
            days_list.append(current.strftime('%Y-%m-%d'))
            current += timedelta(days=1)

        # ✅ متغيرات الإجماليات
        total_sales_sum = 0
        total_euro_sum = 0
        total_dollar_sum = 0
        total_exp_sum = 0
        total_net_sum = 0

        for date_str in days_list:
            # ═══ المبيعات (جنيه + فيزا) ═══
            c.execute("""SELECT COALESCE(SUM(pound), 0), COALESCE(SUM(visa), 0),
                                COALESCE(SUM(euro), 0), COALESCE(SUM(dollar), 0)
                         FROM sales WHERE date = ?""", (date_str,))
            row = c.fetchone()
            sales_pound = row[0] or 0
            sales_visa = row[1] or 0
            sales_euro = row[2] or 0
            sales_dollar = row[3] or 0

            sales = sales_pound + sales_visa

            # ═══ المصروفات ═══
            c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids 
                         WHERE date = ?""", (date_str,))
            expenses = c.fetchone()[0] or 0

            net = sales - expenses

            # ✅ تحديث الإجماليات
            total_sales_sum += sales
            total_euro_sum += sales_euro
            total_dollar_sum += sales_dollar
            total_exp_sum += expenses
            total_net_sum += net

            # ✅ تخطي الأيام الفارغة
            if sales == 0 and expenses == 0 and sales_euro == 0 and sales_dollar == 0:
                continue

            # التاريخ بصيغة مقروءة
            try:
                date_display = datetime.strptime(date_str, '%Y-%m-%d').strftime('%d-%m-%Y')
            except:
                date_display = date_str

            day_name = self.get_day_name_ar(date_str)
            tag = 'positive' if net >= 0 else 'negative'

            # ✅ إظهار/إخفاء أعمدة العملات
            euro_str = f"{sales_euro:,.2f}" if sales_euro > 0 else ''
            dollar_str = f"{sales_dollar:,.2f}" if sales_dollar > 0 else ''

            # ✅ الترتيب الجديد: اليورو والدولار في النهاية
            self.tree.insert('', 'end', values=(
                date_display, day_name,
                f"{sales:,.2f}",
                f"{expenses:,.2f}",
                f"{net:,.2f}",
                euro_str,
                dollar_str
            ), tags=(tag,))

        # ═══ صف الإجمالي ═══
        visible_days = len(self.tree.get_children())
        if visible_days > 0:
            # ✅ أعمدة العملات في الإجمالي
            euro_total_str = f"{total_euro_sum:,.2f} €" if total_euro_sum > 0 else ''
            dollar_total_str = f"{total_dollar_sum:,.2f} $" if total_dollar_sum > 0 else ''

            # ✅ الإجمالي بنفس الترتيب
            self.tree.insert('', 'end', values=(
                'الإجمالي', '',
                f"{total_sales_sum:,.2f}",
                f"{total_exp_sum:,.2f}",
                f"{total_net_sum:,.2f}",
                euro_total_str,
                dollar_total_str
            ), tags=('total_row',))

        conn.close()

        # ═══ إخفاء/إظهار أعمدة العملات حسب البيانات ═══
        if total_euro_sum > 0:
            self.tree.column('euro', width=100)
            self.tree.heading('euro', text='يورو (€)')
        else:
            self.tree.column('euro', width=0, minwidth=0, stretch=False)
            self.tree.heading('euro', text='')

        if total_dollar_sum > 0:
            self.tree.column('dollar', width=100)
            self.tree.heading('dollar', text='دولار ($)')
        else:
            self.tree.column('dollar', width=0, minwidth=0, stretch=False)
            self.tree.heading('dollar', text='')

        # ═══ تحديث شريط الإجماليات ═══
        self.lbl_days.config(text=f"عدد الأيام: {visible_days}")
        self.lbl_total_sales.config(text=f"إجمالي المبيعات: {total_sales_sum:,.2f}")
        self.lbl_total_exp.config(text=f"إجمالي المصروفات: {total_exp_sum:,.2f}")
        self.lbl_total_net.config(text=f"إجمالي الصافي: {total_net_sum:,.2f}")

        # ✅ إظهار إجماليات العملات (إن وُجدت)
        if total_euro_sum > 0:
            if not hasattr(self, 'lbl_total_euro'):
                self.lbl_total_euro = tk.Label(
                    self, text=f"إجمالي اليورو: {total_euro_sum:,.2f} €",
                    font=('Arial', 12, 'bold'),
                    bg='#34495e', fg='#3498db')
                self.lbl_total_euro.pack(in_=self.status_bar.master,
                                         side='left', padx=15, pady=5)
            else:
                self.lbl_total_euro.config(
                    text=f"إجمالي اليورو: {total_euro_sum:,.2f} €")
        else:
            if hasattr(self, 'lbl_total_euro'):
                self.lbl_total_euro.destroy()
                del self.lbl_total_euro

        if total_dollar_sum > 0:
            if not hasattr(self, 'lbl_total_dollar'):
                self.lbl_total_dollar = tk.Label(
                    self, text=f"إجمالي الدولار: {total_dollar_sum:,.2f} $",
                    font=('Arial', 12, 'bold'),
                    bg='#34495e', fg='#2ecc71')
                self.lbl_total_dollar.pack(in_=self.status_bar.master,
                                           side='left', padx=15, pady=5)
            else:
                self.lbl_total_dollar.config(
                    text=f"إجمالي الدولار: {total_dollar_sum:,.2f} $")
        else:
            if hasattr(self, 'lbl_total_dollar'):
                self.lbl_total_dollar.destroy()
                del self.lbl_total_dollar

        self.status_bar.config(
            text=f"من {date_from} إلى {date_to}  |  عدد الأيام المعروضة: {visible_days}")

        # حفظ البيانات للتصدير
        self.last_data = {
            'date_from': date_from,
            'date_to': date_to,
            'total_sales': total_sales_sum,
            'total_euro': total_euro_sum,
            'total_dollar': total_dollar_sum,
            'total_expenses': total_exp_sum,
            'total_net': total_net_sum,
            'days_count': visible_days
        }

    # ────────────────────────────────────────────────
    def export_image(self):
        """تصدير التقرير إلى صورة PNG"""
        from tkinter import filedialog

        if not hasattr(self, 'last_data') or not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض التقرير أولاً")
            return

        data = self.last_data
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"تقرير_مدة_{data['date_from']}_{data['date_to']}.png",
            title="حفظ التقرير صورة"
        )
        if not file_path:
            return

        # جمع بيانات الجدول
        # ✅ 7 أعمدة
        # ✅ 7 أعمدة (الترتيب الجديد)
        columns = ['التاريخ', 'اليوم', 'المبيعات (ج.م)',
                   'المصروفات', 'صافي الإيراد',
                   'يورو (€)', 'دولار ($)']

        rows_data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            rows_data.append(values)

        # ✅ الإجماليات (مع العملات)
        summary_items = [
            ("إجمالي المبيعات", f"{data['total_sales']:,.2f} ج.م"),
            ("إجمالي المصروفات", f"{data['total_expenses']:,.2f} ج.م"),
            ("إجمالي صافي الإيراد", f"{data['total_net']:,.2f} ج.م"),
        ]

        # إضافة العملات فقط إن وُجدت
        if data.get('total_euro', 0) > 0:
            summary_items.append(
                ("إجمالي اليورو", f"{data['total_euro']:,.2f} €"))
        if data.get('total_dollar', 0) > 0:
            summary_items.append(
                ("إجمالي الدولار", f"{data['total_dollar']:,.2f} $"))

        success = export_table_to_image(
            file_path=file_path,
            title="📊 تقرير مدة",
            subtitle=f"من {data['date_from']} إلى {data['date_to']}",
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                    f"✅ تم حفظ الصورة:\n{file_path}\n\nفتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

    # ────────────────────────────────────────────────
    def export_excel(self):
        """تصدير التقرير إلى Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        if not hasattr(self, 'last_data') or not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض التقرير أولاً")
            return

        data = self.last_data
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"تقرير_مدة_{data['date_from']}_{data['date_to']}.xlsx",
            title="حفظ التقرير Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "تقرير مدة"
            ws.sheet_view.rightToLeft = True

            # ═══ العنوان ═══
            ws.merge_cells('A1:E1')
            c = ws['A1']
            c.value = f"تقرير مدة من {data['date_from']} إلى {data['date_to']}"
            c.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            c.fill = PatternFill("solid", fgColor="2C3E50")
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ✅ رؤوس الأعمدة (7 أعمدة)
            # ✅ رؤوس الأعمدة (الترتيب الجديد)
            headers = ['التاريخ', 'اليوم', 'المبيعات (ج.م)',
                       'المصروفات', 'صافي الإيراد',
                       'يورو (€)', 'دولار ($)']

            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=col_idx, value=h)
                cell.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            # ✅ دمج العنوان لـ 7 أعمدة
            ws.merge_cells('A1:G1')
            c = ws['A1']
            c.value = f"تقرير مدة من {data['date_from']} إلى {data['date_to']}"
            c.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            c.fill = PatternFill("solid", fgColor="2C3E50")
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ✅ البيانات
            row_num = 3
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                is_total = (values[0] == 'الإجمالي')

                for col_idx, val in enumerate(values, start=1):
                    cell = ws.cell(row=row_num, column=col_idx, value=val)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
                    if is_total:
                        cell.font = Font(name='Arial', size=11, bold=True, color='8B4513')
                        cell.fill = PatternFill("solid", fgColor="FFF3CD")
                    else:
                        cell.font = Font(name='Arial', size=10)
                row_num += 1

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            messagebox.showerror("خطأ", f"فشل التصدير: {e}")


class TreasuryPage(tk.Frame):
    """صفحة إدارة صافي الإيراد والخزينة"""

    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None
        self._filtered_mode = False   # ← جديد: وضع الفلترة
        # 1️⃣ الشريط العلوي
        add_program_banner(self)

        # 2️⃣ منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ═════════ العنوان ═════════
        tk.Label(content, text="💰 إدارة صافي الإيراد والخزينة",
                 font=('Arial', 15, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=12)

        # ═════════ بطاقة الرصيد الحالي (3 عملات) ═════════
        balance_card = tk.Frame(content, bg='#2c3e50', height=130)
        balance_card.pack(fill='x', padx=15, pady=(0, 10))
        balance_card.pack_propagate(False)

        tk.Label(balance_card, text="💰 رصيد الخزينة الحالي",
                 font=('Arial', 13, 'bold'),
                 bg='#2c3e50', fg='#bdc3c7').pack(pady=(10, 5))

        balances_row = tk.Frame(balance_card, bg='#2c3e50')
        balances_row.pack(fill='x', pady=(0, 10))

        # الرصيد بالجنيه
        egp_frame = tk.Frame(balances_row, bg='#2c3e50')
        egp_frame.pack(side='right', expand=True, padx=15)
        tk.Label(egp_frame, text="🇪🇬 ج.م",
                 font=('Arial', 11, 'bold'),
                 bg='#2c3e50', fg='#bdc3c7').pack()
        self.lbl_balance = tk.Label(egp_frame, text="0.00",
                                    font=('Arial', 20, 'bold'),
                                    bg='#2c3e50', fg='#2ecc71')
        self.lbl_balance.pack()

        # الرصيد باليورو
        euro_frame = tk.Frame(balances_row, bg='#2c3e50')
        euro_frame.pack(side='right', expand=True, padx=15)
        tk.Label(euro_frame, text="🇪🇺 يورو",
                 font=('Arial', 11, 'bold'),
                 bg='#2c3e50', fg='#bdc3c7').pack()
        self.lbl_balance_euro = tk.Label(euro_frame, text="0.00",
                                         font=('Arial', 20, 'bold'),
                                         bg='#2c3e50', fg='#3498db')
        self.lbl_balance_euro.pack()

        # الرصيد بالدولار
        usd_frame = tk.Frame(balances_row, bg='#2c3e50')
        usd_frame.pack(side='right', expand=True, padx=15)
        tk.Label(usd_frame, text="🇺🇸 دولار",
                 font=('Arial', 11, 'bold'),
                 bg='#2c3e50', fg='#bdc3c7').pack()
        self.lbl_balance_usd = tk.Label(usd_frame, text="0.00",
                                        font=('Arial', 20, 'bold'),
                                        bg='#2c3e50', fg='#16a085')
        self.lbl_balance_usd.pack()

        # ═════════ فلاتر التاريخ ═════════
        filter_frame = tk.LabelFrame(content, text="فلترة الحركات",
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#34495e',
                                     padx=5, pady=10)
        filter_frame.pack(fill='x', padx=5, pady=(0, 5))

        # ─── الصف الأول: التاريخ ───
        date_row = tk.Frame(filter_frame, bg='#ecf0f1')
        date_row.pack(fill='x', pady=(0, 8))

        tk.Label(date_row, text="من تاريخ:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='right', padx=5)
        today = datetime.now()
        self.date_from = DateEntry(date_row, width=13,
                                   date_pattern='yyyy-mm-dd')
        self.date_from.pack(side='right', padx=5)
        self.date_from.set_date(today.replace(day=1))

        tk.Label(date_row, text="إلى تاريخ:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='right', padx=5)
        self.date_to = DateEntry(date_row, width=13,
                                 date_pattern='yyyy-mm-dd')
        self.date_to.pack(side='right', padx=5)
        self.date_to.set_date(today)

        # ✅ زرّان بدلاً من زر واحد
        tk.Button(date_row, text="🔍 عرض / فلترة",
                  command=self.apply_filter,
                  bg='#2980b9', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=12, pady=5).pack(side='left', padx=5)

        tk.Button(date_row, text="🔄 عرض الكل",
                  command=self.show_all,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=12, pady=5).pack(side='left', padx=5)

        # ─── الصف الثاني: فلترة نوع الحركة (بدون command) ───
        type_row = tk.Frame(filter_frame, bg='#ecf0f1')
        type_row.pack(fill='x')

        tk.Label(type_row, text="نوع الحركة:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#34495e').pack(side='right', padx=5)

        self.var_owner = tk.BooleanVar(value=True)
        self.var_trader = tk.BooleanVar(value=True)
        self.var_employee = tk.BooleanVar(value=True)
        self.var_other = tk.BooleanVar(value=True)

        tk.Checkbutton(type_row, text="🏢 تسليم لصاحب المكان",
                       variable=self.var_owner,
                       bg='#ecf0f1', fg='#8e44ad',
                       font=('Arial', 11, 'bold'),
                       activebackground='#ecf0f1').pack(side='right', padx=5)

        tk.Checkbutton(type_row, text="🧑‍💼 سداد تاجر",
                       variable=self.var_trader,
                       bg='#ecf0f1', fg='#c0392b',
                       font=('Arial', 11, 'bold'),
                       activebackground='#ecf0f1').pack(side='right', padx=5)

        tk.Checkbutton(type_row, text="👷 سداد عامل",
                       variable=self.var_employee,
                       bg='#ecf0f1', fg='#16a085',
                       font=('Arial', 11, 'bold'),
                       activebackground='#ecf0f1').pack(side='right', padx=5)

        tk.Checkbutton(type_row, text="💸 مصروف آخر",
                       variable=self.var_other,
                       bg='#ecf0f1', fg='#f39c12',
                       font=('Arial', 11, 'bold'),
                       activebackground='#ecf0f1').pack(side='right', padx=5)

        # ─── الصف الثالث: أزرار الاختيار الجماعي والتصدير ───
        buttons_row = tk.Frame(filter_frame, bg='#ecf0f1')
        buttons_row.pack(fill='x', pady=(8, 0))

        tk.Button(buttons_row, text="✅ تحديد الكل",
                  command=self.select_all_types,
                  bg='#27ae60', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='right', padx=3)

        tk.Button(buttons_row, text="❌ إلغاء الكل",
                  command=self.deselect_all_types,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='right', padx=3)

        tk.Button(buttons_row, text="📊 Excel",
                  command=self.export_excel,
                  bg='#27ae60', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='left', padx=3)

        tk.Button(buttons_row, text="🖼️ طباعه",
                  command=self.export_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='left', padx=3)

        tk.Button(buttons_row, text="🖼️ تقرير اعمال الخزينه",
                  command=self.open_treasury_report,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=3).pack(side='left', padx=3)

        # ═════════ أزرار العمليات ═════════
        action_frame = tk.LabelFrame(content, text="العمليات",
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#34495e',
                                     padx=5, pady=12)
        action_frame.pack(fill='x', padx=15, pady=(0, 5))

        row1 = tk.Frame(action_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="🏢 تسليم لصاحب المكان",
                  command=lambda: self.add_action('تسليم لصاحب المكان'),
                  bg='#8e44ad', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=22).pack(side='right', padx=5)

        tk.Button(row1, text="🧑‍💼 سداد تاجر",
                  command=lambda: self.add_action('سداد تاجر'),
                  bg='#c0392b', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=18).pack(side='right', padx=5)

        tk.Button(row1, text="👷 سداد عامل",
                  command=lambda: self.add_action('سداد عامل'),
                  bg='#16a085', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=18).pack(side='right', padx=5)

        row2 = tk.Frame(action_frame, bg='#ecf0f1')
        row2.pack(pady=3)

        tk.Button(row1, text="💸 مصروف آخر",
                  command=lambda: self.add_action('مصروف آخر'),
                  bg='#f39c12', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=22).pack(side='right', padx=5)

        tk.Button(row1, text="➖ حذف الحركة المحددة",
                  command=self.delete_action,
                  bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=22).pack(side='right', padx=5)

        tk.Button(row1, text="🔄 تحديث الرصيد",
                  command=self.load_actions,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, width=22).pack(side='right', padx=5)

        # ═════════ جدول الحركات ═════════
        tree_frame = tk.LabelFrame(content, text="حركات الخزينة",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)

        columns = ('id', 'date', 'type', 'beneficiary',
                   'amount', 'euro', 'usd', 'desc')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=15)
        self.tree.heading('id', text='ID')
        self.tree.heading('date', text='التاريخ')
        self.tree.heading('type', text='نوع الحركة')
        self.tree.heading('beneficiary', text='المستفيد')
        self.tree.heading('amount', text='ج.م')
        self.tree.heading('euro', text='يورو €')
        self.tree.heading('usd', text='دولار $')
        self.tree.heading('desc', text='البيان')

        self.tree.column('id', width=40, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('type', width=150, anchor='center')
        self.tree.column('beneficiary', width=150, anchor='center')
        self.tree.column('amount', width=100, anchor='center')
        self.tree.column('euro', width=90, anchor='center')
        self.tree.column('usd', width=90, anchor='center')
        self.tree.column('desc', width=250, anchor='e')

        self.tree.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        self.tree.tag_configure('owner', background='#f4ecf7')
        self.tree.tag_configure('trader', background='#fadbd8')
        self.tree.tag_configure('employee', background='#d1f2eb')
        self.tree.tag_configure('other', background='#fef5e7')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # ═════════ شريط الملخص ═════════
        summary = tk.Frame(content, bg='#34495e')
        summary.pack(fill='x', padx=15, pady=(5, 15))

        sum_row1 = tk.Frame(summary, bg='#34495e')
        sum_row1.pack(fill='x', pady=5)

        self.lbl_total_revenue = tk.Label(sum_row1,
                                          text="إجمالي صافي الإيراد: 0.00",
                                          font=('Arial', 13, 'bold'),
                                          bg='#F1C40F', fg='#2c3e50',
                                          padx=15, pady=8)
        self.lbl_total_revenue.pack(side='right', padx=8)

        self.lbl_total_owner = tk.Label(sum_row1, text="مسلّم لصاحب المكان: 0.00",
                                        font=('Arial', 11, 'bold'),
                                        bg='#34495e', fg='#9b59b6')
        self.lbl_total_owner.pack(side='right', padx=12)

        self.lbl_total_trader = tk.Label(sum_row1, text="سدادات تجار: 0.00",
                                         font=('Arial', 11, 'bold'),
                                         bg='#34495e', fg='#e74c3c')
        self.lbl_total_trader.pack(side='right', padx=12)

        self.lbl_total_emp = tk.Label(sum_row1, text="سدادات عمال: 0.00",
                                      font=('Arial', 11, 'bold'),
                                      bg='#34495e', fg='#1abc9c')
        self.lbl_total_emp.pack(side='right', padx=12)

        self.lbl_total_other = tk.Label(sum_row1, text="مصروفات أخرى: 0.00",
                                        font=('Arial', 11, 'bold'),
                                        bg='#34495e', fg='#f39c12')
        self.lbl_total_other.pack(side='right', padx=12)

        self.lbl_total_balance = tk.Label(sum_row1,
                                          text="رصيد الخزينة: 0.00",
                                          font=('Arial', 13, 'bold'),
                                          bg='#F1C40F', fg='#2c3e50',
                                          padx=15, pady=8)

        # التحميل الأولي: عرض كل الحركات بدون فلترة
        self.load_actions(filtered=False)

    # ────────────────────────────────────────────────
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def on_select(self, event):
        sel = self.tree.selection()
        self.selected_id = int(sel[0]) if sel else None

    # ────────────────────────────────────────────────
    def calc_net_revenue(self, conn, date_str):
        """صافي الإيراد النقدي ليوم = مبيعات نقدية − مصاريف نقدية"""
        c = conn.cursor()
        c.execute("""SELECT COALESCE(SUM(pound), 0)
                     FROM sales WHERE date = ?""", (date_str,))
        sales = c.fetchone()[0] or 0

        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                     WHERE date = ?
                     AND COALESCE(payment_method, 'نقدي') = 'نقدي'""",
                  (date_str,))
        expenses = c.fetchone()[0] or 0
        return sales - expenses

    def calc_total_net_revenue(self, conn, date_from, date_to):
        """مجموع صافي الإيراد النقدي خلال مدة"""
        c = conn.cursor()
        c.execute("""SELECT COALESCE(SUM(pound), 0)
                     FROM sales WHERE date >= ? AND date <= ?""",
                  (date_from, date_to))
        total_sales = c.fetchone()[0] or 0

        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                     WHERE date >= ? AND date <= ?
                     AND COALESCE(payment_method, 'نقدي') = 'نقدي'""",
                  (date_from, date_to))
        total_expenses = c.fetchone()[0] or 0
        return total_sales - total_expenses

    # ────────────────────────────────────────────────
    def apply_filter(self):
        """تطبيق الفلاتر (تاريخ + أنواع) عند الضغط على الزر"""
        self._filtered_mode = True
        self.load_actions(filtered=True)

    def show_all(self):
        """عرض كل الحركات بدون فلترة"""
        self._filtered_mode = False
        self.load_actions(filtered=False)

    # ────────────────────────────────────────────────
    def load_actions(self, filtered=None):
        """تحميل حركات الخزينة + حساب الأرصدة (3 عملات)"""
        if filtered is None:
            filtered = getattr(self, '_filtered_mode', False)
        self._filtered_mode = filtered

        for row in self.tree.get_children():
            self.tree.delete(row)

        # ─── قراءة الفلاتر حسب الوضع ───
        if filtered:
            date_from = self.date_from.get().strip()
            date_to = self.date_to.get().strip()
            selected_types = []
            if self.var_owner.get(): selected_types.append('تسليم لصاحب المكان')
            if self.var_trader.get(): selected_types.append('سداد تاجر')
            if self.var_employee.get(): selected_types.append('سداد عامل')
            if self.var_other.get(): selected_types.append('مصروف آخر')
        else:
            date_from = None
            date_to = None
            selected_types = [
                'تسليم لصاحب المكان', 'سداد تاجر',
                'سداد عامل', 'مصروف آخر'
            ]

        conn = self.get_connection()
        c = conn.cursor()

        # ═══════════════════════════════════════════════════
        # 1. الأرصدة الحالية (3 عملات)
        # ═══════════════════════════════════════════════════
        # المبيعات النقدية فقط
        c.execute("""SELECT COALESCE(SUM(pound), 0),
                            COALESCE(SUM(euro), 0),
                            COALESCE(SUM(dollar), 0)
                     FROM sales""")
        sales_egp, sales_euro, sales_usd = c.fetchone()
        sales_egp = sales_egp or 0
        sales_euro = sales_euro or 0
        sales_usd = sales_usd or 0

        # مصاريف paids النقدية فقط (بدون فودافون كاش وبدون خزينة)
        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                     WHERE COALESCE(payment_method, 'نقدي') = 'نقدي'""")
        paids_egp = c.fetchone()[0] or 0

        # حركات الخزينة (تشمل المُسجَّل من صفحة المصاريف تلقائياً)
        c.execute("""SELECT COALESCE(SUM(amount), 0),
                            COALESCE(SUM(euro), 0),
                            COALESCE(SUM(usd), 0)
                     FROM treasury_actions""")
        treasury_egp, treasury_euro, treasury_usd = c.fetchone()
        treasury_egp = treasury_egp or 0
        treasury_euro = treasury_euro or 0
        treasury_usd = treasury_usd or 0

        # الأرصدة النهائية (بدون ازدواج)
        balance_egp = sales_egp - paids_egp - treasury_egp
        balance_euro = sales_euro - treasury_euro
        balance_usd = sales_usd - treasury_usd

        # تحديث البطاقات
        self.lbl_balance.config(
            text=f"{balance_egp:,.2f}",
            fg='#2ecc71' if balance_egp >= 0 else '#e74c3c')
        self.lbl_balance_euro.config(
            text=f"{balance_euro:,.2f}",
            fg='#3498db' if balance_euro >= 0 else '#e74c3c')
        self.lbl_balance_usd.config(
            text=f"{balance_usd:,.2f}",
            fg='#16a085' if balance_usd >= 0 else '#e74c3c')

        # ═══════════════════════════════════════════════════
        # 2. عرض الحركات
        # ═══════════════════════════════════════════════════
        totals = {
            'تسليم لصاحب المكان': {'egp': 0, 'euro': 0, 'usd': 0},
            'سداد تاجر': {'egp': 0, 'euro': 0, 'usd': 0},
            'سداد عامل': {'egp': 0, 'euro': 0, 'usd': 0},
            'مصروف آخر': {'egp': 0, 'euro': 0, 'usd': 0}
        }

        if selected_types:
            placeholders = ','.join('?' * len(selected_types))

            if filtered and date_from and date_to:
                query = f"""SELECT id, date, action_type, beneficiary,
                                   amount, euro, usd, description
                            FROM treasury_actions
                            WHERE date >= ? AND date <= ?
                              AND action_type IN ({placeholders})
                            ORDER BY date DESC, id DESC"""
                params = [date_from, date_to] + selected_types
            else:
                query = f"""SELECT id, date, action_type, beneficiary,
                                   amount, euro, usd, description
                            FROM treasury_actions
                            WHERE action_type IN ({placeholders})
                            ORDER BY date DESC, id DESC"""
                params = selected_types

            c.execute(query, params)
            rows = c.fetchall()
            self.last_loaded_rows = rows

            for rid, date, action_type, beneficiary, amount, euro, usd, desc in rows:
                tag = 'owner' if 'صاحب' in action_type else \
                    'trader' if 'تاجر' in action_type else \
                        'employee' if 'عامل' in action_type else 'other'

                egp_str = f"{amount:,.2f}" if amount else ''
                euro_str = f"{euro:,.2f}" if euro else ''
                usd_str = f"{usd:,.2f}" if usd else ''

                self.tree.insert('', 'end', iid=rid, tags=(tag,), values=(
                    rid, date, action_type, beneficiary or '',
                    egp_str, euro_str, usd_str, desc or ''))

                if action_type in totals:
                    totals[action_type]['egp'] += amount or 0
                    totals[action_type]['euro'] += euro or 0
                    totals[action_type]['usd'] += usd or 0
        else:
            self.last_loaded_rows = []

        # ═══════════════════════════════════════════════════
        # 3. حسابات الفترة
        # ═══════════════════════════════════════════════════
        if filtered and date_from and date_to:
            c.execute("""SELECT COALESCE(SUM(pound), 0),
                                COALESCE(SUM(euro), 0),
                                COALESCE(SUM(dollar), 0)
                         FROM sales WHERE date >= ? AND date <= ?""",
                      (date_from, date_to))
            period_sales_egp, period_sales_euro, period_sales_usd = c.fetchone()

            c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                         WHERE date >= ? AND date <= ?
                           AND COALESCE(payment_method, 'نقدي') = 'نقدي'""",
                      (date_from, date_to))
            period_paids = c.fetchone()[0] or 0
        else:
            c.execute("""SELECT COALESCE(SUM(pound), 0),
                                COALESCE(SUM(euro), 0),
                                COALESCE(SUM(dollar), 0)
                         FROM sales""")
            period_sales_egp, period_sales_euro, period_sales_usd = c.fetchone()

            c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                         WHERE COALESCE(payment_method, 'نقدي') = 'نقدي'""")
            period_paids = c.fetchone()[0] or 0

        conn.close()

        # إجماليات الفترة
        period_revenue = (period_sales_egp or 0) - (period_paids or 0)
        treasury_total_egp = sum(t['egp'] for t in totals.values())
        treasury_total_euro = sum(t['euro'] for t in totals.values())
        treasury_total_usd = sum(t['usd'] for t in totals.values())

        period_balance = period_revenue - treasury_total_egp

        self.last_period_data = {
            'period_revenue': period_revenue,
            'treasury_total': treasury_total_egp,
            'treasury_total_euro': treasury_total_euro,
            'treasury_total_usd': treasury_total_usd,
            'period_balance': period_balance,
            'current_balance': balance_egp,
            'current_balance_euro': balance_euro,
            'current_balance_usd': balance_usd,
            'totals': totals,
            'date_from': date_from or 'البداية',
            'date_to': date_to or 'اليوم'
        }

        # ═══════════════════════════════════════════════════
        # 4. تحديث الإجماليات السفلية
        # ═══════════════════════════════════════════════════
        self.lbl_total_revenue.config(
            text=f"إجمالي صافي الإيراد: {period_revenue:,.2f}")
        self.lbl_total_owner.config(
            text=f"مسلّم لصاحب المكان: {totals['تسليم لصاحب المكان']['egp']:,.2f}")
        self.lbl_total_trader.config(
            text=f"سدادات تجار: {totals['سداد تاجر']['egp']:,.2f}")
        self.lbl_total_emp.config(
            text=f"سدادات عمال: {totals['سداد عامل']['egp']:,.2f}")
        self.lbl_total_other.config(
            text=f"مصروفات أخرى: {totals['مصروف آخر']['egp']:,.2f}")
        self.lbl_total_balance.config(
            text=f"رصيد الخزينة: {period_balance:,.2f}")

    # ────────────────────────────────────────────────
    def add_action(self, action_type):
        """فتح نافذة إضافة حركة جديدة"""
        TreasuryActionWindow(self, action_type)

    def delete_action(self):
        """حذف حركة محددة"""
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر حركة من الجدول أولاً")
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT action_type, beneficiary, amount, euro, usd,
                            paid_id, description
                     FROM treasury_actions WHERE id=?""", (self.selected_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return

        action_type, beneficiary, amount, euro, usd, paid_id, desc = row

        # بناء رسالة التأكيد
        msg = f"هل أنت متأكد من حذف الحركة؟\n\n"
        msg += f"• النوع: {action_type}\n"
        msg += f"• المستفيد: {beneficiary or '---'}\n"
        if amount:
            msg += f"• ج.م: {amount:,.2f}\n"
        if euro:
            msg += f"• يورو: {euro:,.2f} €\n"
        if usd:
            msg += f"• دولار: {usd:,.2f} $\n"
        if paid_id:
            msg += f"\n⚠️ سيتم حذف السجل المرتبط في المصاريف"
        msg += "\n\nهل تريد المتابعة؟"

        if not messagebox.askyesno("تأكيد الحذف", msg):
            conn.close()
            return

        try:
            if paid_id:
                c.execute("DELETE FROM paids WHERE id=?", (paid_id,))
            c.execute("DELETE FROM treasury_actions WHERE id=?",
                      (self.selected_id,))
            conn.commit()
            messagebox.showinfo("نجاح", "✅ تم حذف الحركة بنجاح")
            self.selected_id = None
            # إعادة التحميل بنفس الوضع الحالي
            self.load_actions(filtered=self._filtered_mode)
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()

    def select_all_types(self):
        """تحديد كل الأنواع (تُطبَّق عند الضغط على زر الفلترة)"""
        self.var_owner.set(True)
        self.var_trader.set(True)
        self.var_employee.set(True)
        self.var_other.set(True)

    def deselect_all_types(self):
        """إلغاء تحديد كل الأنواع (تُطبَّق عند الضغط على زر الفلترة)"""
        self.var_owner.set(False)
        self.var_trader.set(False)
        self.var_employee.set(False)
        self.var_other.set(False)

    # ═══════════════════════════════════════════════
    # تصدير Excel
    # ═══════════════════════════════════════════════
    def export_excel(self):
        """تصدير حركات الخزينة إلى Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        if not hasattr(self, 'last_loaded_rows') or not self.last_loaded_rows:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        rows = self.last_loaded_rows
        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"حركات_الخزينة_{date_from}_{date_to}.xlsx",
            title="حفظ تقرير الخزينة Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "حركات الخزينة"
            ws.sheet_view.rightToLeft = True

            # عرض الأعمدة (8 أعمدة الآن)
            col_widths = {'A': 8, 'B': 14, 'C': 22, 'D': 22,
                          'E': 14, 'F': 14, 'G': 14, 'H': 40}
            for col, width in col_widths.items():
                ws.column_dimensions[col].width = width

            # ═══ العنوان ═══
            ws.merge_cells('A1:H1')
            title_cell = ws['A1']
            title_cell.value = f"تقرير حركات الخزينة ({date_from} - {date_to})"
            title_cell.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            title_cell.fill = PatternFill("solid", fgColor="2C3E50")
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ═══ رؤوس الأعمدة (8 أعمدة) ═══
            headers = ['ID', 'التاريخ', 'نوع الحركة', 'المستفيد',
                       'ج.م', 'يورو €', 'دولار $', 'البيان']
            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=col_idx, value=header)
                cell.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            # ═══ البيانات — تفريغ 8 متغيرات ═══
            row_num = 3
            total_amount = 0
            total_euro = 0
            total_usd = 0

            for rid, date, action_type, beneficiary, amount, euro, usd, desc in rows:
                ws.cell(row=row_num, column=1, value=rid)
                ws.cell(row=row_num, column=2, value=date or '')
                ws.cell(row=row_num, column=3, value=action_type or '')
                ws.cell(row=row_num, column=4, value=beneficiary or '')
                ws.cell(row=row_num, column=5, value=amount or 0)
                ws.cell(row=row_num, column=6, value=euro or 0)
                ws.cell(row=row_num, column=7, value=usd or 0)
                ws.cell(row=row_num, column=8, value=desc or '')

                for col_idx in range(1, 9):
                    cell = ws.cell(row=row_num, column=col_idx)
                    cell.font = Font(name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center',
                                               vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

                total_amount += amount or 0
                total_euro += euro or 0
                total_usd += usd or 0
                row_num += 1

            # ═══ صف الإجمالي ═══
            row_num += 1
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            cell = ws.cell(row=row_num, column=1, value="الإجمالي")
            cell.font = Font(name='Arial', size=12, bold=True, color='8B4513')
            cell.fill = PatternFill("solid", fgColor="FFF3CD")
            cell.alignment = Alignment(horizontal='center', vertical='center')

            for col_idx, val in zip([5, 6, 7], [total_amount, total_euro, total_usd]):
                cell = ws.cell(row=row_num, column=col_idx, value=val)
                cell.font = Font(name='Arial', size=12, bold=True, color='C0392B')
                cell.fill = PatternFill("solid", fgColor="FFF3CD")
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # ═══ ملخص الفترة (إن وُجد) ═══
            if hasattr(self, 'last_period_data'):
                data = self.last_period_data
                totals = data['totals']
                row_num += 2

                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=8)
                cell = ws.cell(row=row_num, column=1, value="📊 ملخص الفترة")
                cell.font = Font(name='Arial', size=13, bold=True, color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="8B4513")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                ws.row_dimensions[row_num].height = 25
                row_num += 1

                summary_rows = [
                    ("إجمالي صافي الإيراد", data['period_revenue'], True),
                    ("مسلّم لصاحب المكان", totals['تسليم لصاحب المكان']['egp'], False),
                    ("سداد إلى التجار", totals['سداد تاجر']['egp'], False),
                    ("سداد إلى العمال", totals['سداد عامل']['egp'], False),
                    ("مصاريف أخرى", totals['مصروف آخر']['egp'], False),
                    ("رصيد الخزينة", data['period_balance'], True),
                ]

                for label, value, highlight in summary_rows:
                    ws.merge_cells(start_row=row_num, start_column=1,
                                   end_row=row_num, end_column=4)
                    cell = ws.cell(row=row_num, column=1, value=value)
                    cell.font = Font(name='Arial', size=12, bold=True,
                                     color='000000' if highlight else '2C3E50')
                    cell.fill = PatternFill("solid",
                                            fgColor="FFFF00" if highlight else "F8F9FA")
                    cell.alignment = Alignment(horizontal='center',
                                               vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

                    ws.merge_cells(start_row=row_num, start_column=5,
                                   end_row=row_num, end_column=8)
                    cell = ws.cell(row=row_num, column=5, value=label)
                    cell.font = Font(name='Arial', size=12, bold=True,
                                     color='000000' if highlight else '2C3E50')
                    cell.fill = PatternFill("solid",
                                            fgColor="FFFF00" if highlight else "F8F9FA")
                    cell.alignment = Alignment(horizontal='right',
                                               vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

                    row_num += 1

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم حفظ الملف:\n{file_path}")
            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    # ═══════════════════════════════════════════════
    # طباعه
    # ═══════════════════════════════════════════════
    def export_image(self):
        """تصدير حركات الخزينة إلى صورة PNG"""
        from tkinter import filedialog

        if not hasattr(self, 'last_loaded_rows') or not self.last_loaded_rows:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"حركات_الخزينة_{date_from}_{date_to}.png",
            title="حفظ تقرير الخزينة صورة"
        )
        if not file_path:
            return

        columns = ['ID', 'التاريخ', 'نوع الحركة', 'المستفيد', 'المبلغ', 'البيان']
        rows_data = []
        total_amount = 0
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            rows_data.append(values)
            try:
                total_amount += float(str(values[4]).replace(',', ''))
            except:
                pass

        summary_items = []
        summary_items.append(("عدد الحركات المعروضة", str(len(rows_data))))

        if hasattr(self, 'last_period_data'):
            data = self.last_period_data
            totals = data['totals']
            summary_items.append(("─────", ""))
            summary_items.append(("⭐ إجمالي صافي الإيراد",
                                  f"{data['period_revenue']:,.2f} ج.م"))
            summary_items.append(("─────", ""))
            summary_items.append(("مسلّم لصاحب المكان",
                                  f"{totals['تسليم لصاحب المكان']['egp']:,.2f} ج.م"))
            summary_items.append(("سداد إلى التجار",
                                  f"{totals['سداد تاجر']['egp']:,.2f} ج.م"))
            summary_items.append(("سداد إلى العمال",
                                  f"{totals['سداد عامل']['egp']:,.2f} ج.م"))
            summary_items.append(("مصاريف أخرى",
                                  f"{totals['مصروف آخر']['egp']:,.2f} ج.م"))
            summary_items.append(("─────", ""))
            summary_items.append(("⭐ رصيد الخزينة",
                                  f"{data['period_balance']:,.2f} ج.م"))

        success = export_table_to_image(
            file_path=file_path,
            title="💰 حركات الخزينة",
            subtitle=f"من {date_from} إلى {date_to}",
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم حفظ الصورة:\n{file_path}\n\nفتح الصورة؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

    def open_treasury_report(self):
        """فتح تقرير أعمال الخزينة عن مدة"""
        TreasuryReportWindow(self)




class TreasuryActionWindow:
    """نافذة تسجيل حركة خزينة جديدة"""
    def __init__(self, parent_page, action_type):
        self.parent_page = parent_page
        self.action_type = action_type

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title(f"إضافة حركة: {action_type}")
        self.window.geometry("520x550")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ═══ العنوان ═══
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icons = {
            'تسليم لصاحب المكان': '🏢',
            'سداد تاجر': '🧑‍💼',
            'سداد عامل': '👷',
            'مصروف آخر': '💸'
        }
        icon = icons.get(action_type, '💰')

        tk.Label(title_frame, text=f"{icon} {action_type}",
                 font=('Arial', 15, 'bold'),
                 bg='#2c3e50', fg='white').pack(pady=15)

        # ═══ الإطار الرئيسي ═══
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        # التاريخ
        tk.Label(main, text="التاريخ: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(row=0, column=0, padx=5, pady=8, sticky='e')
        self.entry_date = DateEntry(main, width=22, date_pattern='yyyy-mm-dd')
        self.entry_date.grid(row=0, column=1, padx=5, pady=8, sticky='w')
        self.entry_date.set_date(datetime.now())

        # المستفيد
        if action_type in ('سداد تاجر', 'سداد عامل'):
            if action_type == 'سداد تاجر':
                label = "التاجر: *"
                table = 'traders'
            else:
                label = "العامل: *"
                table = 'employees'

            tk.Label(main, text=label, font=('Arial', 11, 'bold'),
                     bg='#ecf0f1', anchor='e').grid(row=1, column=0, padx=5, pady=8, sticky='e')

            self.combo_beneficiary = AutocompleteCombobox(main, width=22,
                                                            font=('Arial', 11))
            self.combo_beneficiary.grid(row=1, column=1, padx=5, pady=8, sticky='w')

            # جلب الأسماء
            conn = self.parent_page.get_connection()
            c = conn.cursor()
            c.execute(f"SELECT name FROM {table} ORDER BY name")
            names = [r[0] for r in c.fetchall()]
            conn.close()
            self.combo_beneficiary.set_values(names)
            if names:
                self.combo_beneficiary.set(names[0])
        else:
            self.combo_beneficiary = None

        # المبلغ
        # ═════════ حقول المبالغ الثلاثة ═════════
        row = 2 if self.combo_beneficiary else 1

        # حاوية المبالغ
        amounts_frame = tk.LabelFrame(main, text="المبلغ (أدخل واحد على الأقل)",
                                      font=('Arial', 10, 'bold'),
                                      bg='#ecf0f1', fg='#34495e',
                                      padx=10, pady=8)
        amounts_frame.grid(row=row, column=0, columnspan=2,
                           padx=5, pady=10, sticky='ew')

        # ─── ج.م ───
        tk.Label(amounts_frame, text="ج.م:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#27ae60').grid(row=0, column=0,
                                                  padx=5, pady=5, sticky='e')
        self.entry_amount_egp = tk.Entry(amounts_frame, width=14,
                                         font=('Arial', 12), justify='center')
        self.entry_amount_egp.grid(row=0, column=1, padx=5, pady=5)
        self.entry_amount_egp.insert(0, "0")

        # ─── يورو ───
        tk.Label(amounts_frame, text="يورو €:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#2980b9').grid(row=1, column=0,
                                                  padx=5, pady=5, sticky='e')
        self.entry_amount_euro = tk.Entry(amounts_frame, width=14,
                                          font=('Arial', 12), justify='center')
        self.entry_amount_euro.grid(row=1, column=1, padx=5, pady=5)
        self.entry_amount_euro.insert(0, "0")

        # ─── دولار ───
        tk.Label(amounts_frame, text="دولار $:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#16a085').grid(row=2, column=0,
                                                  padx=5, pady=5, sticky='e')
        self.entry_amount_usd = tk.Entry(amounts_frame, width=14,
                                         font=('Arial', 12), justify='center')
        self.entry_amount_usd.grid(row=2, column=1, padx=5, pady=5)
        self.entry_amount_usd.insert(0, "0")

        # الوصف
        row += 1
        tk.Label(main, text="البيان:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(row=row, column=0, padx=5, pady=8, sticky='ne')

        self.entry_desc = tk.Text(main, width=25, height=4, font=('Arial', 11))
        self.entry_desc.grid(row=row, column=1, padx=5, pady=8, sticky='w')

        # ═══ معلومة: رصيد الخزينة الحالي ═══
        row += 1
        info_frame = tk.Frame(main, bg='#fff9e6', relief='solid', bd=1)
        info_frame.grid(row=row, column=0, columnspan=2, pady=15, sticky='ew')

        self.lbl_balance_info = tk.Label(
            info_frame,
            text=f"💰 رصيد الخزينة الحالي: {self.parent_page.lbl_balance.cget('text')}",
            font=('Arial', 11, 'bold'), bg='#fff9e6', fg='#8B4513')
        self.lbl_balance_info.pack(pady=10, padx=15)

        # ═══ أزرار الحفظ والإلغاء ═══
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        # ✅ تركيز على أول حقل مبلغ (ج.م)
        self.entry_amount_egp.focus()
        self.entry_amount_egp.select_range(0, tk.END)

    def save(self):
        """حفظ الحركة في treasury_actions (3 عملات)"""
        date_str = self.entry_date.get().strip()
        desc = self.entry_desc.get('1.0', 'end').strip()

        if not date_str:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return

        # ═════════ قراءة المبالغ الثلاثة ═════════
        try:
            amount_egp = float(self.entry_amount_egp.get().strip() or 0)
            amount_euro = float(self.entry_amount_euro.get().strip() or 0)
            amount_usd = float(self.entry_amount_usd.get().strip() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "المبالغ يجب أن تكون أرقاماً")
            return

        # ✅ يجب إدخال مبلغ واحد على الأقل
        if amount_egp <= 0 and amount_euro <= 0 and amount_usd <= 0:
            messagebox.showwarning("تنبيه",
                                   "يجب إدخال مبلغ واحد على الأقل (ج.م / يورو / دولار)")
            return

        # ✅ التحقق من القيم السالبة
        if amount_egp < 0 or amount_euro < 0 or amount_usd < 0:
            messagebox.showerror("خطأ", "المبالغ يجب أن تكون موجبة")
            return

        beneficiary = ''
        if self.combo_beneficiary:
            beneficiary = self.combo_beneficiary.get().strip()
            if not beneficiary:
                messagebox.showwarning("تنبيه", "يجب اختيار المستفيد")
                return

        # ═════════ حفظ في قاعدة البيانات ═════════
        conn = self.parent_page.get_connection()
        c = conn.cursor()
        try:
            c.execute("""INSERT INTO treasury_actions 
                         (date, action_type, beneficiary, 
                          amount, euro, usd, description, created_at)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                      (date_str, self.action_type, beneficiary,
                       amount_egp, amount_euro, amount_usd, desc,
                       datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            conn.commit()

            # رسالة النجاح
            msg = f"✅ تم حفظ الحركة:\n"
            if amount_egp > 0:
                msg += f"  • ج.م: {amount_egp:,.2f}\n"
            if amount_euro > 0:
                msg += f"  • يورو: {amount_euro:,.2f} €\n"
            if amount_usd > 0:
                msg += f"  • دولار: {amount_usd:,.2f} $\n"

            messagebox.showinfo("نجاح", msg)

            self.window.destroy()
            self.parent_page.load_actions()
        except Exception as e:
            conn.rollback()
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}")
        finally:
            conn.close()


class TreasuryReportWindow:
    """تقرير أعمال الخزينة عن مدة"""

    def __init__(self, parent_page):
        self.parent_page = parent_page

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("📋 تقرير أعمال الخزينة")
        self.window.geometry("1200x750")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # ═════════ العنوان ═════════
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="📋 تقرير أعمال الخزينة",
                 font=('Arial', 17, 'bold'),
                 bg='#2c3e50', fg='white').pack(side='right', padx=5, pady=15)

        # ═════════ فلاتر الفترة ═════════
        filter_frame = tk.LabelFrame(self.window, text="الفترة",
                                       font=('Arial', 11, 'bold'),
                                       bg='#ecf0f1', fg='#34495e',
                                       padx=5, pady=10)
        filter_frame.pack(fill='x', padx=5, pady=10)

        # من تاريخ
        tk.Label(filter_frame, text="من تاريخ:",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=5)
        self.date_from = DateEntry(filter_frame, width=13,
                                     date_pattern='yyyy-mm-dd')
        self.date_from.pack(side='right', padx=5)
        # أول يوم من الشهر الحالي
        today = datetime.now()
        self.date_from.set_date(today.replace(day=1))

        # إلى تاريخ
        tk.Label(filter_frame, text="إلى تاريخ:",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1').pack(side='right', padx=5)
        self.date_to = DateEntry(filter_frame, width=13,
                                   date_pattern='yyyy-mm-dd')
        self.date_to.pack(side='right', padx=5)
        self.date_to.set_date(today)

        # زر عرض
        tk.Button(filter_frame, text="🔍 عرض التقرير",
                  command=self.load_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='left', padx=5)

        # أزرار سريعة
        tk.Button(filter_frame, text="الشهر الحالي",
                  command=self.set_this_month,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(side='left', padx=3)
        tk.Button(filter_frame, text="آخر 7 أيام",
                  command=self.set_last_7_days,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=10, pady=5).pack(side='left', padx=3)

        # ═════════ الجدول ═════════
        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)

        columns = ('date', 'revenue', 'amount', 'action_type', 'beneficiary')

        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                   show='headings', height=15)

        # رؤوس الأعمدة (من اليمين إلى اليسار)
        # ✅ الترتيب الجديد: التاريخ - إيراد - مبلغ - نوع - مستفيد
        headers = {
            'date': 'التاريخ',
            'revenue': 'إيراد الخزينة',
            'amount': 'المبلغ',
            'action_type': 'نوع الحركة',
            'beneficiary': 'المستفيد'
        }

        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('date', width=120, anchor='center')
        self.tree.column('revenue', width=140, anchor='center')
        self.tree.column('amount', width=140, anchor='center')
        self.tree.column('action_type', width=200, anchor='center')
        self.tree.column('beneficiary', width=180, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        # تنسيقات الصفوف
        self.tree.tag_configure('total_row', background='#2c3e50',
                                 foreground='#f1c40f',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('balance_row', background='#f1c40f',
                                 foreground='#2c3e50',
                                 font=('Arial', 11, 'bold'))
        self.tree.tag_configure('alt_row', background='#f8f9fa')
        # ═════════ ✅ شريط الملخص السفلي ═════════
        summary_frame = tk.Frame(self.window, bg='#34495e')
        summary_frame.pack(fill='x', padx=5, pady=(5, 0))

        # الصف الأول
        sum_row1 = tk.Frame(summary_frame, bg='#34495e')
        sum_row1.pack(fill='x', pady=5)

        self.lbl_sum_owner = tk.Label(sum_row1,
                                      text="مسلّم لصاحب المكان: 0.00 ج.م",
                                      font=('Arial', 12, 'bold'),
                                      bg='#34495e', fg='#9b59b6')
        self.lbl_sum_owner.pack(side='right', padx=10)

        self.lbl_sum_trader = tk.Label(sum_row1,
                                       text="سداد التجار: 0.00 ج.م",
                                       font=('Arial', 12, 'bold'),
                                       bg='#34495e', fg='#e74c3c')
        self.lbl_sum_trader.pack(side='right', padx=10)

        # الصف الثاني
        sum_row2 = tk.Frame(summary_frame, bg='#34495e')
        sum_row2.pack(fill='x', pady=5)

        self.lbl_sum_employee = tk.Label(sum_row1,
                                         text="سداد العمال: 0.00 ج.م",
                                         font=('Arial', 12, 'bold'),
                                         bg='#34495e', fg='#1abc9c')
        self.lbl_sum_employee.pack(side='right', padx=10)

        self.lbl_sum_other = tk.Label(sum_row1,
                                      text="مصاريف أخرى: 0.00 ج.م",
                                      font=('Arial', 12, 'bold'),
                                      bg='#34495e', fg='#f39c12')
        self.lbl_sum_other.pack(side='right', padx=10)

        # ═════════ الأزرار ═════════
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="📊 Excel",
                  command=self.export_excel,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='left', padx=5)

        tk.Button(btn_frame, text="🖼️ طباعه",
                  command=self.export_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إغلاق",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='left', padx=5)

        # تحميل أولي
        self.load_report()

    # ═══════════════════════════════════════════
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def set_this_month(self):
        today = datetime.now()
        self.date_from.set_date(today.replace(day=1))
        self.date_to.set_date(today)
        self.load_report()

    def set_last_7_days(self):
        from datetime import timedelta
        today = datetime.now()
        self.date_from.set_date(today - timedelta(days=6))
        self.date_to.set_date(today)
        self.load_report()

    # ═══════════════════════════════════════════
    def update_summary_bar(self, owner, trader, employee, other):
        """تحديث شريط الملخص السفلي"""
        self.lbl_sum_owner.config(
            text=f"مسلّم لصاحب المكان: {owner:,.2f} ج.م")
        self.lbl_sum_trader.config(
            text=f"سداد التجار: {trader:,.2f} ج.م")
        self.lbl_sum_employee.config(
            text=f"سداد العمال: {employee:,.2f} ج.م")
        self.lbl_sum_other.config(
            text=f"مصاريف أخرى: {other:,.2f} ج.م")

    # ═══════════════════════════════════════════
    def load_report(self):
        """تحميل تقرير أعمال الخزينة (نقدي فقط)"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()

        if not date_from or not date_to:
            messagebox.showwarning("تنبيه", "اختر الفترة")
            return

        conn = self.get_connection()
        c = conn.cursor()

        from datetime import timedelta
        d1 = datetime.strptime(date_from, '%Y-%m-%d')
        d2 = datetime.strptime(date_to, '%Y-%m-%d')

        all_dates = []
        current = d1
        while current <= d2:
            all_dates.append(current.strftime('%Y-%m-%d'))
            current += timedelta(days=1)

        total_amount = 0
        total_revenue = 0
        totals_owner = 0
        totals_trader = 0
        totals_employee = 0
        totals_other = 0
        report_rows = []

        for date_str in all_dates:
            # ✅ 1. المبيعات النقدية فقط (بدون فيزا)
            c.execute("""SELECT COALESCE(SUM(pound), 0)
                         FROM sales WHERE date = ?""", (date_str,))
            day_sales = c.fetchone()[0] or 0

            # ✅ 2. المصاريف النقدية فقط
            c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids 
                         WHERE date = ?
                         AND COALESCE(payment_method, 'نقدي') = 'نقدي'""",
                      (date_str,))
            day_paids = c.fetchone()[0] or 0

            # ✅ 3. صافي الإيراد النقدي
            day_revenue = day_sales - day_paids

            # 4. حركات الخزينة لهذا اليوم
            c.execute("""SELECT action_type, beneficiary, amount
                         FROM treasury_actions
                         WHERE date = ?
                         ORDER BY id""", (date_str,))
            day_actions = c.fetchall()

            # 5. بناء الصفوف
            if day_actions:
                first_row = True
                for action_type, beneficiary, amount in day_actions:
                    report_rows.append({
                        'date': date_str,
                        'revenue': day_revenue if first_row else None,
                        'amount': amount or 0,
                        'action_type': action_type or '',
                        'beneficiary': beneficiary or ''
                    })
                    total_amount += amount or 0
                    first_row = False

                    # تصنيف الإجماليات حسب النوع
                    if 'صاحب' in (action_type or ''):
                        totals_owner += amount or 0
                    elif 'تاجر' in (action_type or ''):
                        totals_trader += amount or 0
                    elif 'عامل' in (action_type or ''):
                        totals_employee += amount or 0
                    else:
                        totals_other += amount or 0

                total_revenue += day_revenue
            elif day_revenue != 0:
                report_rows.append({
                    'date': date_str,
                    'revenue': day_revenue,
                    'amount': None,
                    'action_type': '',
                    'beneficiary': ''
                })
                total_revenue += day_revenue

        conn.close()

        # حساب المتبقي
        remaining = total_revenue - total_amount

        # عرض البيانات في الجدول
        for i, r in enumerate(report_rows):
            date_display = r['date']
            try:
                date_display = datetime.strptime(r['date'], '%Y-%m-%d').strftime('%d-%m-%Y')
            except:
                pass

            amount_str = f"{r['amount']:,.0f}" if r['amount'] is not None else ''
            revenue_str = f"{r['revenue']:,.0f}" if r['revenue'] is not None else ''
            tag = 'alt_row' if i % 2 == 0 else ''

            self.tree.insert('', 'end', tags=(tag,), values=(
                date_display,
                revenue_str,
                amount_str,
                r['action_type'],
                r['beneficiary']
            ))

        # حفظ البيانات للتصدير
        self.last_report_data = {
            'date_from': date_from,
            'date_to': date_to,
            'rows': report_rows,
            'remaining': remaining,
            'total_amount': total_amount,
            'total_revenue': total_revenue,
            'totals_owner': totals_owner,
            'totals_trader': totals_trader,
            'totals_employee': totals_employee,
            'totals_other': totals_other
        }

        # صف المتبقي
        self.tree.insert('', 'end', tags=('balance_row',), values=(
            'المتبقي',
            '',
            f"{remaining:,.0f}",
            '',
            ''
        ))

        # صف الإجمالي
        self.tree.insert('', 'end', tags=('total_row',), values=(
            'الإجمالي',
            f"{total_revenue:,.0f}",
            f"{total_amount + remaining:,.0f}",
            '',
            ''
        ))

        # تحديث شريط الملخص السفلي
        self.update_summary_bar(totals_owner, totals_trader,
                                totals_employee, totals_other)

    # ═══════════════════════════════════════════
    def export_excel(self):
        """تصدير تقرير أعمال الخزينة إلى Excel"""
        from tkinter import filedialog
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        if not hasattr(self, 'last_report_data'):
            messagebox.showwarning("تنبيه", "اعرض التقرير أولاً")
            return

        data = self.last_report_data

        # اختيار مسار الحفظ
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"تقرير_الخزينة_{data['date_from']}_{data['date_to']}.xlsx",
            title="حفظ التقرير Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "تقرير الخزينة"
            ws.sheet_view.rightToLeft = True

            # ═════════ عرض الأعمدة ═════════
            for col, w in {'A': 15, 'B': 16, 'C': 14, 'D': 22, 'E': 22}.items():
                ws.column_dimensions[col].width = w

            # ═════════ العنوان ═════════
            ws.merge_cells('A1:E1')
            cell = ws['A1']
            cell.value = f"تقرير أعمال الخزينة - {RESTAURANT_NAME} ({data['date_from']} - {data['date_to']})"

            cell.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            cell.fill = PatternFill("solid", fgColor="2C3E50")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ═════════ رؤوس الأعمدة ═════════
            headers = ['التاريخ', 'إيراد الخزينة', 'المبلغ', 'نوع الحركة', 'المستفيد']
            for i, h in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=i, value=h)
                cell.font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            # ═════════ البيانات ═════════
            row_num = 3
            for r in data['rows']:
                date_display = r['date']
                try:
                    date_display = datetime.strptime(r['date'], '%Y-%m-%d').strftime('%d-%m-%Y')
                except:
                    pass

                ws.cell(row=row_num, column=1, value=date_display)
                ws.cell(row=row_num, column=2, value=r['revenue'] if r['revenue'] is not None else '')
                ws.cell(row=row_num, column=3, value=r['amount'] if r['amount'] is not None else '')
                ws.cell(row=row_num, column=4, value=r['action_type'])
                ws.cell(row=row_num, column=5, value=r['beneficiary'])

                for col in range(1, 6):
                    cell = ws.cell(row=row_num, column=col)
                    cell.font = Font(name='Arial', size=11)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
                row_num += 1

            # ═════════ صف المتبقي (قبل الأخير) ═════════
            ws.cell(row=row_num, column=1, value='المتبقي')
            ws.cell(row=row_num, column=3, value=data['remaining'])
            for col in range(1, 6):
                cell = ws.cell(row=row_num, column=col)
                cell.font = Font(name='Arial', size=12, bold=True, color='2C3E50')
                cell.fill = PatternFill("solid", fgColor="F1C40F")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
            row_num += 1

            # ═════════ صف الإجمالي (الأخير) ═════════
            ws.cell(row=row_num, column=1, value='الإجمالي')
            ws.cell(row=row_num, column=2, value=data['total_revenue'])
            ws.cell(row=row_num, column=3, value=data['total_amount'] + data['remaining'])
            for col in range(1, 6):
                cell = ws.cell(row=row_num, column=col)
                cell.font = Font(name='Arial', size=12, bold=True, color='F1C40F')
                cell.fill = PatternFill("solid", fgColor="2C3E50")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
            row_num += 2

            # ═════════ ملخص الأنواع ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            cell = ws.cell(row=row_num, column=1, value="📊 ملخص الحركات")
            cell.font = Font(name='Arial', size=13, bold=True, color='FFFFFF')
            cell.fill = PatternFill("solid", fgColor="8B4513")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            row_num += 1

            summary_data = [
                ("مسلّم لصاحب المكان", data['totals_owner'], "9B59B6"),
                ("سداد التجار", data['totals_trader'], "E74C3C"),
                ("سداد العمال", data['totals_employee'], "16A085"),
                ("مصاريف أخرى", data['totals_other'], "F39C12"),
            ]

            for label, value, color in summary_data:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=3)
                cell = ws.cell(row=row_num, column=1, value=label)
                cell.font = Font(name='Arial', size=11, bold=True, color=color)
                cell.fill = PatternFill("solid", fgColor="F8F9FA")
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

                ws.merge_cells(start_row=row_num, start_column=4,
                               end_row=row_num, end_column=5)
                cell = ws.cell(row=row_num, column=4, value=f"{value:,.2f} ج.م")
                cell.font = Font(name='Arial', size=11, bold=True, color='2C3E50')
                cell.fill = PatternFill("solid", fgColor="F8F9FA")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
                row_num += 1

            # ═════════ حفظ الملف ═════════
            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم حفظ الملف:\n{file_path}")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    # ═══════════════════════════════════════════
    def export_image(self):
        """تصدير التقرير إلى صورة"""
        from tkinter import filedialog

        if not hasattr(self, 'last_report_data'):
            messagebox.showwarning("تنبيه", "اعرض التقرير أولاً")
            return

        data = self.last_report_data
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"تقرير_الخزينة_{data['date_from']}_{data['date_to']}.png",
            title="حفظ التقرير صورة"
        )
        if not file_path:
            return

        # جمع بيانات الجدول
        columns = ['التاريخ','إيراد الخزينة', 'المبلغ','نوع الحركة', 'المستفيد',  ]
        rows_data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            rows_data.append(values)

        summary_items = [
            ("مسلّم لصاحب المكان", f"{data['totals_owner']:,.2f} ج.م"),
            ("سداد التجار", f"{data['totals_trader']:,.2f} ج.م"),
            ("سداد العمال", f"{data['totals_employee']:,.2f} ج.م"),
            ("مصاريف أخرى", f"{data['totals_other']:,.2f} ج.م"),
            ("─────", ""),
            ("إجمالي الإيراد", f"{data['total_revenue']:,.2f} ج.م"),
            ("إجمالي المبلغ", f"{data['total_amount'] + data['remaining']:,.2f} ج.م"),
            ("المتبقي", f"{data['remaining']:,.2f} ج.م"),
        ]

        success = export_table_to_image(
            file_path=file_path,
            title=f"📋 تقرير أعمال الخزينة - {RESTAURANT_NAME}",
            subtitle=f"من {data['date_from']} إلى {data['date_to']}",
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح", f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass



# --------------------- مدير التبويبات ---------------------
# ═══════════════════════════════════════════════════════════════════
#                       CATEGORY MANAGER
# ═══════════════════════════════════════════════════════════════════
class CategoryManager:
    def __init__(self, parent_page):
        self.parent_page = parent_page
        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("إدارة التبويبات")
        self.window.geometry("800x550")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')
        # ✅ ربط النافذة بالنافذة الأم (تظل فوقها)
        self.window.transient(parent_page.winfo_toplevel())
        self.window.attributes('-topmost', True)

        self.selected_main_id = None

        main_frame = tk.Frame(self.window, bg='#ecf0f1')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # القائمة الرئيسية
        left_frame = tk.LabelFrame(main_frame, text="التبويبات الرئيسية",
                                    font=('Arial', 12, 'bold'), bg='#ecf0f1')
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.main_listbox = tk.Listbox(left_frame, font=('Arial', 11), height=15)
        self.main_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.main_listbox.bind('<<ListboxSelect>>', self.on_main_select)

        btn_main = tk.Frame(left_frame, bg='#ecf0f1')
        btn_main.pack(pady=5)
        tk.Button(btn_main, text="➕ إضافة", command=self.add_main_category,
                  bg='#27ae60', fg='white', width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_main, text="✏️ تعديل", command=self.edit_main_category,
                  bg='#f39c12', fg='white', width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_main, text="🗑️ حذف", command=self.delete_main_category,
                  bg='#e74c3c', fg='white', width=10).pack(side=tk.LEFT, padx=2)

        # القائمة الفرعية
        right_frame = tk.LabelFrame(main_frame, text="التبويبات الفرعية",
                                     font=('Arial', 12, 'bold'), bg='#ecf0f1')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        self.sub_listbox = tk.Listbox(right_frame, font=('Arial', 11), height=15)
        self.sub_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        btn_sub = tk.Frame(right_frame, bg='#ecf0f1')
        btn_sub.pack(pady=5)
        tk.Button(btn_sub, text="➕ إضافة", command=self.add_sub_category,
                  bg='#27ae60', fg='white', width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_sub, text="✏️ تعديل", command=self.edit_sub_category,
                  bg='#f39c12', fg='white', width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_sub, text="🗑️ حذف", command=self.delete_sub_category,
                  bg='#e74c3c', fg='white', width=10).pack(side=tk.LEFT, padx=2)

        bottom = tk.Frame(self.window, bg='#ecf0f1')
        bottom.pack(pady=5)
        tk.Button(bottom, text="🔄 تحديث", command=self.refresh_all,
                  bg='#3498db', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom, text="❌ إغلاق", command=self.on_close,
                  bg='#95a5a6', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).pack(side=tk.LEFT, padx=5)

        self.load_main_list()

    def _lift_window(self):
        """إعادة النافذة للأمام بعد أي رسالة"""
        try:
            self.window.deiconify()  # استعادة النافذة لو مصغّرة
            self.window.lift()  # رفعها فوق الكل
            self.window.focus_force()  # إعادة التركيز
            self.window.grab_set()  # إعادة القفل
        except Exception as e:
            print(f"⚠️ _lift_window: {e}")

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def load_main_list(self):
        self.main_listbox.delete(0, tk.END)
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT id, name FROM main_categories ORDER BY name")
        self.main_categories = c.fetchall()
        conn.close()
        for _, name in self.main_categories:
            self.main_listbox.insert(tk.END, name)
        if self.main_listbox.size() > 0:
            self.main_listbox.selection_set(0)
            self.on_main_select(None)

    def on_main_select(self, event):
        selection = self.main_listbox.curselection()
        if not selection:
            self.selected_main_id = None
            return
        idx = selection[0]
        self.selected_main_id = self.main_categories[idx][0]
        self.load_sub_list(self.selected_main_id)

    def load_sub_list(self, main_id):
        self.sub_listbox.delete(0, tk.END)
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT id, name FROM sub_categories WHERE main_id=? ORDER BY name", (main_id,))
        self.sub_categories = c.fetchall()
        conn.close()
        for _, name in self.sub_categories:
            self.sub_listbox.insert(tk.END, name)

    def add_main_category(self):
        name = simpledialog.askstring("إضافة تبويب رئيسي",
                                      "أدخل اسم التبويب الرئيسي:",
                                      parent=self.window)
        if not name or not name.strip():
            self._lift_window()
            return
        name = name.strip()
        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("INSERT INTO main_categories (name) VALUES (?)", (name,))
            conn.commit()
            messagebox.showinfo("نجاح", f"تمت إضافة: {name}", parent=self.window)
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "هذا الاسم موجود مسبقاً", parent=self.window)
        finally:
            conn.close()
        self.load_main_list()
        self.refresh_parent_combos()
        self._lift_window()  # ✅ إعادة النافذة للأمام

    def edit_main_category(self):
        selection = self.main_listbox.curselection()
        if not selection:
            messagebox.showwarning("تنبيه", "اختر تبويباً للتعديل", parent=self.window)
            self._lift_window()
            return
        idx = selection[0]
        main_id = self.main_categories[idx][0]
        old_name = self.main_categories[idx][1]
        new_name = simpledialog.askstring("تعديل تبويب رئيسي",
                                          "أدخل الاسم الجديد:",
                                          initialvalue=old_name,
                                          parent=self.window)
        if not new_name or not new_name.strip() or new_name == old_name:
            self._lift_window()
            return
        new_name = new_name.strip()
        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("UPDATE main_categories SET name=? WHERE id=?",
                      (new_name, main_id))
            c.execute("UPDATE paids SET main_name=? WHERE main_name=?",
                      (new_name, old_name))
            conn.commit()
            messagebox.showinfo("نجاح", "تم التعديل", parent=self.window)
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "هذا الاسم موجود مسبقاً", parent=self.window)
        finally:
            conn.close()
        self.load_main_list()
        self.refresh_parent_combos()
        self._lift_window()

    def delete_main_category(self):
        selection = self.main_listbox.curselection()
        if not selection:
            messagebox.showwarning("تنبيه", "اختر تبويباً للحذف", parent=self.window)
            self._lift_window()
            return
        idx = selection[0]
        main_id = self.main_categories[idx][0]
        name = self.main_categories[idx][1]

        protected = ['مشتريات', 'أجور ومرتبات']
        if name in protected:
            messagebox.showerror("ممنوع",
                                 f"لا يمكن حذف التبويب '{name}' لأنه تبويب أساسي.\n"
                                 f"يمكنك فقط إضافة وتعديل التبويبات الفرعية تحته.",
                                 parent=self.window)
            self._lift_window()
            return

        if not messagebox.askyesno("تأكيد",
                                   f"هل أنت متأكد من حذف '{name}' وجميع فروعه؟",
                                   parent=self.window):
            self._lift_window()
            return
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("DELETE FROM main_categories WHERE id=?", (main_id,))
        conn.commit()
        conn.close()
        messagebox.showinfo("نجاح", "تم الحذف", parent=self.window)
        self.load_main_list()
        self.refresh_parent_combos()
        self._lift_window()

    def add_sub_category(self):
        if self.selected_main_id is None:
            messagebox.showwarning("تنبيه", "اختر تبويباً رئيسياً أولاً",
                                   parent=self.window)
            self._lift_window()
            return
        main_name = ""
        for mid, mname in self.main_categories:
            if mid == self.selected_main_id:
                main_name = mname
                break
        name = simpledialog.askstring("إضافة تبويب فرعي",
                                      f"أدخل اسم التبويب الفرعي (تحت '{main_name}'):",
                                      parent=self.window)
        if not name or not name.strip():
            self._lift_window()
            return
        name = name.strip()
        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("INSERT INTO sub_categories (main_id, name) VALUES (?, ?)",
                      (self.selected_main_id, name))
            conn.commit()
            messagebox.showinfo("نجاح", f"تمت إضافة: {name}", parent=self.window)
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "هذا الاسم موجود مسبقاً لهذا الرئيسي",
                                 parent=self.window)
        finally:
            conn.close()
        self.load_sub_list(self.selected_main_id)
        self.refresh_parent_combos()
        self._lift_window()

    def edit_sub_category(self):
        selection = self.sub_listbox.curselection()
        if not selection:
            messagebox.showwarning("تنبيه", "اختر تبويباً فرعياً للتعديل",
                                   parent=self.window)
            self._lift_window()
            return
        idx = selection[0]
        sub_id = self.sub_categories[idx][0]
        old_name = self.sub_categories[idx][1]
        new_name = simpledialog.askstring("تعديل تبويب فرعي",
                                          "أدخل الاسم الجديد:",
                                          initialvalue=old_name,
                                          parent=self.window)
        if not new_name or not new_name.strip() or new_name == old_name:
            self._lift_window()
            return
        new_name = new_name.strip()
        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("UPDATE sub_categories SET name=? WHERE id=?",
                      (new_name, sub_id))
            c.execute("UPDATE paids SET sub_name=? WHERE sub_name=?",
                      (new_name, old_name))
            conn.commit()
            messagebox.showinfo("نجاح", "تم التعديل", parent=self.window)
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "هذا الاسم موجود مسبقاً", parent=self.window)
        finally:
            conn.close()
        self.load_sub_list(self.selected_main_id)
        self.refresh_parent_combos()
        self._lift_window()

    def delete_sub_category(self):
        selection = self.sub_listbox.curselection()
        if not selection:
            messagebox.showwarning("تنبيه", "اختر تبويباً فرعياً للحذف",
                                   parent=self.window)
            self._lift_window()
            return
        idx = selection[0]
        sub_id = self.sub_categories[idx][0]
        name = self.sub_categories[idx][1]
        if not messagebox.askyesno("تأكيد",
                                   f"هل أنت متأكد من حذف '{name}'؟",
                                   parent=self.window):
            self._lift_window()
            return
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("DELETE FROM sub_categories WHERE id=?", (sub_id,))
        conn.commit()
        conn.close()
        messagebox.showinfo("نجاح", "تم الحذف", parent=self.window)
        self.load_sub_list(self.selected_main_id)
        self.refresh_parent_combos()
        self._lift_window()

    def refresh_parent_combos(self):
        try:
            self.parent_page.load_main_categories()
            self.parent_page.load_search_combos()
        except Exception as e:
            print(f"خطأ في refresh_parent_combos: {e}")

    def refresh_all(self):
        self.load_main_list()
        self.refresh_parent_combos()

    def on_close(self):
        self.refresh_parent_combos()
        self.window.destroy()
def get_or_create_trader(trader_name, phone='', notes=''):
    """
    إرجاع trader_id للتاجر.
    إذا لم يكن موجوداً، يتم إنشاؤه تلقائياً في جدول traders
    + إضافته كتصنيف فرعي تحت 'مشتريات'.
    """
    if not trader_name or not trader_name.strip():
        return None
    name = trader_name.strip()

    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        # 1. هل موجود في جدول traders؟
        c.execute("SELECT id FROM traders WHERE name=?", (name,))
        row = c.fetchone()
        if row:
            trader_id = row[0]
        else:
            # غير موجود ← أنشئه
            c.execute("""INSERT INTO traders (name, phone, notes, created_at)
                         VALUES (?, ?, ?, ?)""",
                      (name, phone or '', notes or '',
                       datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            trader_id = c.lastrowid
            print(f"✅ تم إنشاء تاجر جديد تلقائياً: {name} (ID: {trader_id})")

        # 2. تأكد من إضافته كتصنيف فرعي تحت 'مشتريات'
        c.execute("SELECT id FROM main_categories WHERE name='مشتريات'")
        main_row = c.fetchone()
        if main_row:
            main_id = main_row[0]
            c.execute("SELECT id FROM sub_categories WHERE main_id=? AND name=?",
                      (main_id, name))
            if not c.fetchone():
                c.execute("""INSERT INTO sub_categories (main_id, name)
                             VALUES (?, ?)""", (main_id, name))
                print(f"✅ تم ربط '{name}' تحت تبويب 'مشتريات'")

        conn.commit()
        conn.close()
        return trader_id

    except Exception as e:
        print(f"❌ خطأ في get_or_create_trader: {e}")
        return None

def get_or_create_item(item_name, unit=''):
    """
    إرجاع item_id للصنف.
    إذا لم يكن موجوداً، يتم إنشاؤه تلقائياً في جدول items.
    """
    if not item_name or not item_name.strip():
        return None
    name = item_name.strip()
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()

        # هل موجود؟
        c.execute("SELECT id FROM items WHERE name=?", (name,))
        row = c.fetchone()
        if row:
            conn.close()
            return row[0]

        # غير موجود ← أنشئه
        c.execute("""INSERT INTO items (name, unit, created_at)
                     VALUES (?, ?, ?)""",
                  (name, unit or '', datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        new_id = c.lastrowid
        conn.commit()
        conn.close()
        print(f"✅ تم إنشاء صنف جديد تلقائياً: {name} (ID: {new_id})")
        return new_id
    except Exception as e:
        print(f"❌ خطأ في get_or_create_item: {e}")
        return None

# ═══════════════════════════════════════════════════════════════════
#                       INVOICE WINDOW
# ═══════════════════════════════════════════════════════════════════
class InvoiceWindow:
    def __init__(self, parent_page):
        self.parent_page = parent_page
        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("إضافة فاتورة جديدة")
        self.window.geometry("900x700")
        self.window.grab_set()
        self.invoice_items = []
        self.selected_item_index = None
        self.invoice_number = self.generate_invoice_number()
        self.create_widgets()

    def generate_invoice_number(self):
        try:
            conn = self.parent_page.get_connection()
            c = conn.cursor()
            c.execute("SELECT MAX(CAST(invoice_no AS INTEGER)) FROM invoices")
            max_num = c.fetchone()[0]
            conn.close()
            return 1 if max_num is None else max_num + 1
        except Exception as e:
            print(f"خطأ في توليد الرقم: {e}")
            return 1

    def create_widgets(self):
        info_frame = tk.LabelFrame(self.window, text="معلومات الفاتورة", font=('Arial', 12))
        info_frame.pack(fill=tk.X, padx=10, pady=5)

        # رقم الفاتورة
        tk.Label(info_frame, text="رقم الفاتورة:", font=('Arial', 11)).grid(row=0, column=0, padx=5, pady=5, sticky='e')
        tk.Label(info_frame, text=str(self.invoice_number), font=('Arial', 11, 'bold'),
                 fg='#2c3e50').grid(row=0, column=1, padx=5, pady=5, sticky='w')

        # التاريخ
        tk.Label(info_frame, text="التاريخ:", font=('Arial', 11)).grid(row=0, column=2, padx=5, pady=5, sticky='e')
        self.invoice_date = DateEntry(info_frame, width=18, date_pattern='yyyy-mm-dd')
        self.invoice_date.grid(row=0, column=3, padx=5, pady=5, sticky='w')

        # التاجر
        tk.Label(info_frame, text="التاجر:", font=('Arial', 11)).grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.trader_combo = AutocompleteCombobox(info_frame, width=28, font=('Arial', 11))
        self.trader_combo.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        self.load_traders_to_combo()

        # نوع الدفع
        tk.Label(info_frame, text="نوع الدفع:", font=('Arial', 11)).grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.payment_type_combo = ttk.Combobox(info_frame, values=['كاش', 'أجل', 'جزئي'],
                                                width=15, font=('Arial', 11), state='readonly')
        self.payment_type_combo.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        self.payment_type_combo.set('كاش')
        self.payment_type_combo.bind('<<ComboboxSelected>>', self.on_payment_type_change)

        # المبلغ المدفوع
        tk.Label(info_frame, text="المبلغ المدفوع:", font=('Arial', 11)).grid(row=2, column=2, padx=5, pady=5, sticky='e')
        self.entry_paid = tk.Entry(info_frame, width=15, font=('Arial', 11))
        self.entry_paid.grid(row=2, column=3, padx=5, pady=5, sticky='w')
        self.entry_paid.config(state='disabled')

        # المتبقي
        tk.Label(info_frame, text="المتبقي (أجل):", font=('Arial', 11, 'bold')).grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.lbl_remaining = tk.Label(info_frame, text="0.00", font=('Arial', 12, 'bold'), fg='#e74c3c')
        self.lbl_remaining.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        # ملاحظات
        tk.Label(info_frame, text="ملاحظات:", font=('Arial', 11)).grid(row=3, column=2, padx=5, pady=5, sticky='e')
        self.invoice_notes = tk.Entry(info_frame, width=28, font=('Arial', 11))
        self.invoice_notes.grid(row=3, column=3, padx=5, pady=5, sticky='w')

        # جدول الأصناف
        items_frame = tk.LabelFrame(self.window, text="أصناف الفاتورة", font=('Arial', 12))
        items_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ('الصنف', 'الكمية', 'السعر', 'الإجمالي')
        self.tree = ttk.Treeview(items_frame, columns=columns, show='headings', height=8)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150 if col == 'الصنف' else 100, anchor='center')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(items_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)

        # إدخال الصنف
        add_frame = tk.Frame(self.window)
        add_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(add_frame, text="الصنف:", font=('Arial', 11)).pack(side=tk.LEFT, padx=5)

        self.item_combo = AutocompleteCombobox(add_frame, width=25, font=('Arial', 11))
        self.item_combo.pack(side=tk.LEFT, padx=5)
        self.load_items_to_combo()

        # ✅ ربط حدث اختيار الصنف لجلب آخر سعر
        self.item_combo.bind('<<ComboboxSelected>>', self.on_item_selected)
        self.item_combo.bind('<Return>', self.on_item_selected)
        self.item_combo.bind('<Tab>', self.on_item_selected)

        tk.Label(add_frame, text="الكمية:", font=('Arial', 11)).pack(side=tk.LEFT, padx=5)
        self.item_qty = tk.Entry(add_frame, width=10, font=('Arial', 11))
        self.item_qty.pack(side=tk.LEFT, padx=5)

        tk.Label(add_frame, text="السعر:", font=('Arial', 11)).pack(side=tk.LEFT, padx=5)
        self.item_price = tk.Entry(add_frame, width=10, font=('Arial', 11),
                                   bg='#fef9e7', justify='center')
        self.item_price.pack(side=tk.LEFT, padx=5)

        # ✅ ليبل يعرض آخر سعر مسجل للصنف (مرجعي دائم)
        self.lbl_last_price = tk.Label(add_frame, text="",
                                       font=('Arial', 9, 'italic'),
                                       fg='#7f8c8d')
        self.lbl_last_price.pack(side=tk.LEFT, padx=5)

        tk.Button(add_frame, text="➕ إضافة", command=self.add_item,
                  bg='#27ae60', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(add_frame, text="✏️ تعديل", command=self.edit_item,
                  bg='#f39c12', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(add_frame, text="🗑️ حذف", command=self.delete_item,
                  bg='#e74c3c', fg='white').pack(side=tk.LEFT, padx=5)
        tk.Button(add_frame, text="🧹 مسح", command=self.clear_item_fields,
                  bg='#95a5a6', fg='white').pack(side=tk.LEFT, padx=5)

        # الإجمالي والأزرار
        bottom_frame = tk.Frame(self.window)
        bottom_frame.pack(fill=tk.X, padx=10, pady=5)

        self.total_label = tk.Label(bottom_frame, text="إجمالي الفاتورة: 0.00",
                                     font=('Arial', 14, 'bold'), fg='#e74c3c')
        self.total_label.pack(side=tk.LEFT, padx=20)

        tk.Button(bottom_frame, text="💾 حفظ الفاتورة", command=self.save_invoice,
                  bg='#2980b9', fg='white', font=('Arial', 12), padx=15, pady=5).pack(side=tk.RIGHT, padx=10)
        tk.Button(bottom_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', font=('Arial', 12), padx=15, pady=5).pack(side=tk.RIGHT, padx=10)

    def load_traders_to_combo(self):
        conn = self.parent_page.get_connection()
        c = conn.cursor()
        c.execute("SELECT id, name FROM traders ORDER BY name")
        rows = c.fetchall()
        conn.close()
        self.traders_dict = {name: tid for tid, name in rows}
        self.trader_combo.set_values([name for _, name in rows])

    def reload_traders_combo(self):
        """إعادة تحميل قائمة التجار في الكومبوبوكس"""
        try:
            conn = self.parent_page.get_connection()
            c = conn.cursor()
            c.execute("SELECT id, name FROM traders ORDER BY name")
            rows = c.fetchall()
            conn.close()
            self.traders_dict = {name: tid for tid, name in rows}
            self.trader_combo.set_values([name for _, name in rows])
        except Exception as e:
            print(f"⚠️ خطأ في reload_traders_combo: {e}")

    def load_items_to_combo(self):
        conn = self.parent_page.get_connection()
        c = conn.cursor()
        c.execute("SELECT id, name FROM items ORDER BY name")
        rows = c.fetchall()
        conn.close()
        self.items_dict = {name: iid for iid, name in rows}
        self.item_combo.set_values([name for _, name in rows])

    def reload_items_combo(self):
        """إعادة تحميل قائمة الأصناف في الكومبوبوكس"""
        try:
            conn = self.parent_page.get_connection()
            c = conn.cursor()
            c.execute("SELECT id, name FROM items ORDER BY name")
            rows = c.fetchall()
            conn.close()
            self.items_dict = {name: iid for iid, name in rows}
            self.item_combo.set_values([name for _, name in rows])
        except Exception as e:
            print(f"⚠️ خطأ في reload_items_combo: {e}")

    def get_last_price(self, item_name):
        """جلب آخر سعر مسجل للصنف من جدول بنود الفواتير"""
        if not item_name or not item_name.strip():
            return None
        try:
            conn = self.parent_page.get_connection()
            c = conn.cursor()
            c.execute("""
                SELECT ii.price
                FROM invoice_items ii
                JOIN invoices inv ON ii.invoice_id = inv.id
                WHERE ii.item_name = ?
                ORDER BY inv.date DESC, ii.id DESC
                LIMIT 1
            """, (item_name.strip(),))
            row = c.fetchone()
            conn.close()
            if row and row[0] is not None:
                return float(row[0])
            return None
        except Exception as e:
            print(f"⚠️ خطأ في جلب آخر سعر: {e}")
            return None

    def on_item_selected(self, event=None):
        """عند اختيار صنف → جلب آخر سعر وإظهاره + ملء الحقل تلقائياً"""
        name = self.item_combo.get().strip()
        if not name:
            self.lbl_last_price.config(text='')
            return

        last_price = self.get_last_price(name)

        if last_price is None:
            # صنف جديد بلا سجل سابق
            self.lbl_last_price.config(text="🆕 صنف جديد (لا يوجد سعر سابق)",
                                       fg='#95a5a6')
            return

        # ✅ تحديث الليبل المرجعي دائماً
        self.lbl_last_price.config(
            text=f"💡 آخر سعر: {last_price:,.2f} ج.م",
            fg='#27ae60'
        )

        # ✅ ملء حقل السعر (فقط إذا كان فارغاً أو صفراً — لا نمسح ما كتبه المستخدم)
        current = self.item_price.get().strip()
        if not current or current in ('0', '0.0', '0.00'):
            self.item_price.delete(0, tk.END)
            self.item_price.insert(0, f"{last_price:.2f}")
            # نضع المؤشر في حقل الكمية لأنه جاهز للكتابة
            self.item_qty.focus()

    def on_payment_type_change(self, event=None):
        ptype = self.payment_type_combo.get()
        total = sum(i['total'] for i in self.invoice_items)
        if ptype == 'كاش':
            self.entry_paid.config(state='normal')
            self.entry_paid.delete(0, tk.END)
            self.entry_paid.insert(0, f"{total:.2f}")
            self.lbl_remaining.config(text="0.00", fg='#27ae60')
        elif ptype == 'أجل':
            self.entry_paid.delete(0, tk.END)
            self.entry_paid.config(state='disabled')
            self.lbl_remaining.config(text=f"{total:.2f}", fg='#e74c3c')
        else:
            self.entry_paid.config(state='normal')
            self.entry_paid.delete(0, tk.END)

    def add_item(self):
        """إضافة صنف للفاتورة — ينشئ الصنف تلقائياً في المخزن إن لم يكن موجوداً"""
        name = self.item_combo.get().strip()
        qty = self.item_qty.get().strip()
        price = self.item_price.get().strip()

        if not name or not qty or not price:
            messagebox.showerror("خطأ", "يرجى ملء جميع الحقول")
            return

        try:
            qty = float(qty)
            price = float(price)
            if qty <= 0 or price <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("خطأ", "الكمية والسعر يجب أن يكونا رقمين موجبين")
            return

        # ✅ احصل على ID الصنف — أو أنشئه لو جديد
        item_id = get_or_create_item(name)
        if item_id is None:
            messagebox.showerror("خطأ", "فشل في إنشاء الصنف")
            return

        # التحقق من عدم التكرار في نفس الفاتورة
        for item in self.invoice_items:
            if item.get('item_id') == item_id:
                messagebox.showwarning("تنبيه", "الصنف مضاف بالفعل")
                return

        total = qty * price
        self.invoice_items.append({
            'item_id': item_id,
            'name': name,
            'qty': qty,
            'price': price,
            'total': total
        })

        # ✅ حدّث قائمة الأصناف في الـ combobox
        self.reload_items_combo()

        self.refresh_tree()
        self.update_total()
        self.clear_item_fields()

    def edit_item(self):
        if self.selected_item_index is None:
            messagebox.showwarning("تنبيه", "اختر صنفاً")
            return
        name = self.item_combo.get().strip()
        qty = self.item_qty.get().strip()
        price = self.item_price.get().strip()
        if not name or not qty or not price or name not in self.items_dict:
            messagebox.showerror("خطأ", "بيانات غير صحيحة")
            return
        try:
            qty = float(qty)
            price = float(price)
        except ValueError:
            messagebox.showerror("خطأ", "أرقام غير صحيحة")
            return
        self.invoice_items[self.selected_item_index] = {
            'item_id': self.items_dict[name], 'name': name,
            'qty': qty, 'price': price, 'total': qty * price}
        self.refresh_tree()
        self.update_total()
        self.clear_item_fields()

    def delete_item(self):
        if self.selected_item_index is None:
            return
        if messagebox.askyesno("تأكيد", "حذف الصنف؟"):
            del self.invoice_items[self.selected_item_index]
            self.refresh_tree()
            self.update_total()
            self.clear_item_fields()

    def on_item_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0])['values']
        for i, item in enumerate(self.invoice_items):
            if item['name'] == values[0]:
                self.selected_item_index = i
                self.item_combo.set(item['name'])
                self.item_qty.delete(0, tk.END)
                self.item_qty.insert(0, str(item['qty']))
                self.item_price.delete(0, tk.END)
                self.item_price.insert(0, str(item['price']))
                break

    def clear_item_fields(self):
        self.item_combo.set('')
        self.item_qty.delete(0, tk.END)
        self.item_price.delete(0, tk.END)
        self.selected_item_index = None
        # ✅ إعادة تعيين ليبل آخر سعر
        if hasattr(self, 'lbl_last_price'):
            self.lbl_last_price.config(text='')

    def refresh_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for item in self.invoice_items:
            self.tree.insert('', tk.END, values=(item['name'], item['qty'],
                                                  item['price'], item['total']))

    def update_total(self):
        total = sum(item['total'] for item in self.invoice_items)
        self.total_label.config(text=f"إجمالي الفاتورة: {total:.2f}")
        self.on_payment_type_change()

    # ─────────────────────────────────────────────────────────────
    # SAVE INVOICE - ← تم تصحيح الأخطاء هنا
    # ─────────────────────────────────────────────────────────────
    def save_invoice(self):
        if not self.invoice_items:
            messagebox.showwarning("تنبيه", "الفاتورة فارغة")
            return

        trader_name = self.trader_combo.get().strip()
        if not trader_name:
            messagebox.showwarning("تنبيه", "يجب اختيار أو كتابة اسم التاجر")
            return

        # ✅ احصل على التاجر — أو أنشئه لو جديد
        trader_id = get_or_create_trader(trader_name)
        if trader_id is None:
            messagebox.showerror("خطأ", "فشل في إنشاء التاجر")
            return

        # (اختياري) حدّث الـ dict لو التاجر جديد
        if trader_name not in self.traders_dict:
            self.traders_dict[trader_name] = trader_id

        total = sum(item['total'] for item in self.invoice_items)
        payment_type = self.payment_type_combo.get()

        if payment_type == 'كاش':
            paid_amount = total
        elif payment_type == 'أجل':
            paid_amount = 0
        else:
            try:
                paid_amount = float(self.entry_paid.get().strip() or 0)
                if paid_amount > total:
                    messagebox.showerror("خطأ", "المبلغ المدفوع أكبر من الإجمالي")
                    return
            except ValueError:
                messagebox.showerror("خطأ", "المبلغ المدفوع غير صحيح")
                return

        invoice_no = str(self.invoice_number)
        date = self.invoice_date.get().strip() or datetime.now().strftime('%Y-%m-%d')
        notes = self.invoice_notes.get().strip()
        created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        conn = self.parent_page.get_connection()
        c = conn.cursor()
        try:
            # 1. حفظ الفاتورة
            c.execute('''INSERT INTO invoices
                         (invoice_no, date, supplier, notes, total, created_at,
                          trader_id, payment_type, paid_amount)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                      (invoice_no, date, trader_name, notes, total, created_at,
                       trader_id, payment_type, paid_amount))
            invoice_id = c.lastrowid

            # 2. حفظ البنود
            for item in self.invoice_items:
                c.execute('''INSERT INTO invoice_items
                (invoice_id, item_id, item_name, quantity, price, total)
                VALUES (?, ?, ?, ?, ?, ?)''',
                          (invoice_id, item.get('item_id'), item['name'],
                           item['qty'], item['price'], item['total']))

            # ═════════ 3. قيد واحد فقط في paids (خارج الحلقة) ═════════
            # ✅ تحديد الوصف حسب نوع الدفع
            if payment_type == 'كاش':
                desc = f"فاتورة كاش بمبلغ {total:.2f}"
            elif payment_type == 'أجل':
                desc = f"فاتورة أجل بمبلغ {total:.2f}"
            else:  # جزئي
                desc = f"فاتورة رقم {invoice_no} - دفعة جزئية ({paid_amount:.2f} من {total:.2f})"

            # 🛡️ حماية مزدوجة: التحقق أن الفاتورة لم تُسجَّل مسبقاً في paids
            c.execute("SELECT id FROM paids WHERE invoice_id=?", (invoice_id,))
            existing = c.fetchone()

            if existing:
                # لو موجود (حالة نادرة) → تحديث بدلاً من إدراج قيد مكرر
                c.execute("""UPDATE paids
                             SET date=?, value=?, main_name=?, sub_name=?,
                                 description=?, payment_method=?
                             WHERE id=?""",
                          (date, paid_amount, 'مشتريات', trader_name, desc,
                           'نقدي', existing[0]))
            else:
                c.execute('''INSERT INTO paids
                (date, value, main_name, sub_name, description,
                payment_method, invoice_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)''',
                          (date, paid_amount, 'مشتريات', trader_name, desc,
                           'نقدي', invoice_id))

            conn.commit()

            msg = f"تم حفظ الفاتورة رقم {invoice_no}\n"
            msg += f"الإجمالي: {total:.2f}\n"
            msg += f"المدفوع: {paid_amount:.2f}\n"
            msg += f"المتبقي (دين): {total - paid_amount:.2f}"
            if paid_amount > 0:
                msg += f"\n\n✅ تم تسجيل {paid_amount:.2f} في صفحة المصاريف"
            else:
                msg += f"\n\nℹ️ لم يُسجَّل في المصاريف (آجلة بالكامل)"

            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_page.load_paids()
        except Exception as e:
            conn.rollback()
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")
        finally:
            conn.close()


# ═══════════════════════════════════════════════════════════════════
#                  INVOICE DETAILS WINDOW
# ═══════════════════════════════════════════════════════════════════
class InvoiceEditWindow:
    """نافذة تعديل فاتورة كاملة (ماعدا المبلغ المدفوع)"""
    def __init__(self, parent_page, invoice_id, on_save_callback=None):
        self.parent_page = parent_page
        self.invoice_id = invoice_id
        self.on_save_callback = on_save_callback

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title(f"تعديل الفاتورة رقم {invoice_id}")
        self.window.geometry("950x700")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        self.items = []  # قائمة البنود الحالية
        self.selected_item_id = None
        self.selected_item_index = None

        self.create_widgets()
        self.load_invoice_data()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def create_widgets(self):
        # ============ إطار المعلومات ============
        info_frame = tk.LabelFrame(self.window, text="معلومات الفاتورة",
                                    font=('Arial', 12, 'bold'), bg='#ecf0f1',
                                    padx=10, pady=10)
        info_frame.pack(fill='x', padx=10, pady=5)

        # رقم الفاتورة
        tk.Label(info_frame, text="رقم الفاتورة:", font=('Arial', 11), bg='#ecf0f1').grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.entry_invoice_no = tk.Entry(info_frame, width=20, font=('Arial', 11))
        self.entry_invoice_no.grid(row=0, column=1, padx=5, pady=5, sticky='w')

        # التاريخ
        tk.Label(info_frame, text="التاريخ:", font=('Arial', 11), bg='#ecf0f1').grid(row=0, column=2, padx=5, pady=5, sticky='e')
        self.entry_date = tk.Entry(info_frame, width=15, font=('Arial', 11))
        self.entry_date.grid(row=0, column=3, padx=5, pady=5, sticky='w')

        # التاجر
        tk.Label(info_frame, text="التاجر:", font=('Arial', 11), bg='#ecf0f1').grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.combo_trader = AutocompleteCombobox(info_frame, width=25, font=('Arial', 11))
        self.combo_trader.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        self.load_traders()

        # نوع الدفع
        tk.Label(info_frame, text="نوع الدفع:", font=('Arial', 11), bg='#ecf0f1').grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.combo_ptype = ttk.Combobox(info_frame, values=['كاش', 'أجل', 'جزئي'],
                                         width=15, font=('Arial', 11), state='readonly')
        self.combo_ptype.grid(row=2, column=1, padx=5, pady=5, sticky='w')

        # المبلغ المدفوع - للقراءة فقط
        tk.Label(info_frame, text="المبلغ المدفوع:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', fg='#8B4513').grid(row=2, column=2, padx=5, pady=5, sticky='e')
        self.lbl_paid_amount = tk.Label(info_frame, text="0.00",
                                          font=('Arial', 11, 'bold'),
                                          bg='#fff3cd', fg='#c0392b',
                                          width=15, relief='sunken', anchor='center')
        self.lbl_paid_amount.grid(row=2, column=3, padx=5, pady=5, sticky='w')

        # ملاحظة توضيحية
        tk.Label(info_frame,
                 text="💡 المبلغ المدفوع لا يمكن تعديله إلا من صفحة المصاريف",
                 font=('Arial', 9, 'italic'), bg='#ecf0f1', fg='#7f8c8d').grid(
                    row=3, column=0, columnspan=4, pady=3, sticky='w')

        # ملاحظات
        tk.Label(info_frame, text="ملاحظات:", font=('Arial', 11), bg='#ecf0f1').grid(row=4, column=0, padx=5, pady=5, sticky='e')
        self.entry_notes = tk.Entry(info_frame, width=50, font=('Arial', 11))
        self.entry_notes.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky='w')

        # ============ جدول الأصناف ============
        items_frame = tk.LabelFrame(self.window, text="بنود الفاتورة (تعديل / إضافة / حذف)",
                                     font=('Arial', 12, 'bold'), bg='#ecf0f1',
                                     padx=5, pady=5)
        items_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ('id', 'الصنف', 'الكمية', 'السعر', 'الإجمالي')
        self.tree = ttk.Treeview(items_frame, columns=columns, show='headings', height=10)
        for col in columns:
            self.tree.heading(col, text=col)
            width = 50 if col == 'id' else (200 if col == 'الصنف' else 100)
            self.tree.column(col, width=width, anchor='center')
        self.tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        sb = ttk.Scrollbar(items_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)

        # ============ إدخال / تعديل صنف ============
        edit_frame = tk.Frame(self.window, bg='#ecf0f1')
        edit_frame.pack(fill='x', padx=10, pady=5)

        tk.Label(edit_frame, text="الصنف:", font=('Arial', 11), bg='#ecf0f1').pack(side='left', padx=5)
        self.combo_item = AutocompleteCombobox(edit_frame, width=25, font=('Arial', 11))
        self.combo_item.pack(side='left', padx=5)
        self.load_items()

        tk.Label(edit_frame, text="الكمية:", font=('Arial', 11), bg='#ecf0f1').pack(side='left', padx=5)
        self.entry_qty = tk.Entry(edit_frame, width=10, font=('Arial', 11))
        self.entry_qty.pack(side='left', padx=5)

        tk.Label(edit_frame, text="السعر:", font=('Arial', 11), bg='#ecf0f1').pack(side='left', padx=5)
        self.entry_price = tk.Entry(edit_frame, width=10, font=('Arial', 11))
        self.entry_price.pack(side='left', padx=5)

        tk.Button(edit_frame, text="➕ إضافة", command=self.add_item,
                  bg='#27ae60', fg='white', padx=10).pack(side='left', padx=3)
        tk.Button(edit_frame, text="✏️ تعديل", command=self.edit_item,
                  bg='#f39c12', fg='white', padx=10).pack(side='left', padx=3)
        tk.Button(edit_frame, text="🗑️ حذف", command=self.delete_item,
                  bg='#e74c3c', fg='white', padx=10).pack(side='left', padx=3)
        tk.Button(edit_frame, text="🧹 مسح", command=self.clear_item_fields,
                  bg='#95a5a6', fg='white', padx=10).pack(side='left', padx=3)

        # ============ الإجمالي والأزرار ============
        bottom_frame = tk.Frame(self.window, bg='#34495e', height=70)
        bottom_frame.pack(fill='x', padx=10, pady=10)
        bottom_frame.pack_propagate(False)

        self.lbl_total = tk.Label(bottom_frame, text="إجمالي الفاتورة: 0.00",
                                    font=('Arial', 14, 'bold'), bg='#34495e', fg='#f1c40f')
        self.lbl_total.pack(side='left', padx=20, pady=20)

        self.lbl_remaining = tk.Label(bottom_frame, text="المتبقي: 0.00",
                                        font=('Arial', 12, 'bold'), bg='#34495e', fg='#e74c3c')
        self.lbl_remaining.pack(side='left', padx=20, pady=20)

        tk.Button(bottom_frame, text="💾 حفظ التعديلات", command=self.save_all,
                  bg='#27ae60', fg='white', font=('Arial', 12, 'bold'),
                  padx=20, pady=10).pack(side='right', padx=10)

        tk.Button(bottom_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', font=('Arial', 12, 'bold'),
                  padx=20, pady=10).pack(side='right', padx=5)

    # ─────────────────────────────────────────────────────────────
    def load_traders(self):
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM traders ORDER BY name")
        names = [r[0] for r in c.fetchall()]
        conn.close()
        self.combo_trader.set_values(names)

    def load_items(self):
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM items ORDER BY name")
        names = [r[0] for r in c.fetchall()]
        conn.close()
        self.combo_item.set_values(names)

    def load_invoice_data(self):
        """تحميل بيانات الفاتورة والبنود"""
        conn = self.get_connection()
        c = conn.cursor()

        # بيانات الفاتورة
        c.execute("""SELECT invoice_no, date, supplier, notes,
                            COALESCE(payment_type, 'أجل'),
                            COALESCE(paid_amount, 0)
                     FROM invoices WHERE id=?""", (self.invoice_id,))
        row = c.fetchone()

        if not row:
            messagebox.showerror("خطأ", "الفاتورة غير موجودة")
            conn.close()
            self.window.destroy()
            return

        inv_no, date, supplier, notes, ptype, paid = row

        self.entry_invoice_no.delete(0, tk.END)
        self.entry_invoice_no.insert(0, inv_no or '')

        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, date or '')

        self.combo_trader.set(supplier or '')
        self.combo_ptype.set(ptype)
        self.lbl_paid_amount.config(text=f"{paid:.2f}")

        self.entry_notes.delete(0, tk.END)
        self.entry_notes.insert(0, notes or '')

        # البنود
        c.execute("""SELECT id, item_name, quantity, price, total
                     FROM invoice_items WHERE invoice_id=?
                     ORDER BY id""", (self.invoice_id,))
        self.items = c.fetchall()
        conn.close()

        self.refresh_items_tree()
        self.update_total_label()
        self.clear_item_fields()

    def refresh_items_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for item in self.items:
            self.tree.insert('', 'end', iid=str(item[0]), values=item)

    def update_total_label(self):
        total = sum(i[4] for i in self.items)
        self.lbl_total.config(text=f"إجمالي الفاتورة: {total:.2f}")

        # حساب المتبقي
        try:
            paid = float(self.lbl_paid_amount.cget('text'))
        except:
            paid = 0
        remaining = total - paid
        self.lbl_remaining.config(text=f"المتبقي: {remaining:.2f}",
                                    fg='#e74c3c' if remaining > 0 else '#2ecc71')

    def on_item_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0])['values']
        if not values:
            return
        self.selected_item_id = values[0]
        for i, item in enumerate(self.items):
            if item[0] == self.selected_item_id:
                self.selected_item_index = i
                self.combo_item.set(item[1])
                self.entry_qty.delete(0, tk.END)
                self.entry_qty.insert(0, str(item[2]))
                self.entry_price.delete(0, tk.END)
                self.entry_price.insert(0, str(item[3]))
                break

    def clear_item_fields(self):
        self.combo_item.set('')
        self.entry_qty.delete(0, tk.END)
        self.entry_price.delete(0, tk.END)
        self.selected_item_id = None
        self.selected_item_index = None

    def add_item(self):
        name = self.combo_item.get().strip()
        qty_str = self.entry_qty.get().strip()
        price_str = self.entry_price.get().strip()

        if not name or not qty_str or not price_str:
            messagebox.showerror("خطأ", "يرجى ملء جميع الحقول")
            return

        try:
            qty = float(qty_str)
            price = float(price_str)
        except ValueError:
            messagebox.showerror("خطأ", "الكمية والسعر أرقام")
            return

        total = qty * price

        conn = self.get_connection()
        c = conn.cursor()
        try:
            # جلب item_id
            c.execute("SELECT id FROM items WHERE name=?", (name,))
            item_row = c.fetchone()
            item_id = item_row[0] if item_row else None

            c.execute("""INSERT INTO invoice_items
                         (invoice_id, item_id, item_name, quantity, price, total)
                         VALUES (?, ?, ?, ?, ?, ?)""",
                      (self.invoice_id, item_id, name, qty, price, total))
            conn.commit()
            messagebox.showinfo("نجاح", "تمت إضافة البند")
            self.load_invoice_data()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الإضافة: {e}")
        finally:
            conn.close()

    def edit_item(self):
        if self.selected_item_id is None:
            messagebox.showwarning("تنبيه", "اختر بنداً للتعديل")
            return

        name = self.combo_item.get().strip()
        qty_str = self.entry_qty.get().strip()
        price_str = self.entry_price.get().strip()

        if not name or not qty_str or not price_str:
            messagebox.showerror("خطأ", "يرجى ملء جميع الحقول")
            return

        try:
            qty = float(qty_str)
            price = float(price_str)
        except ValueError:
            messagebox.showerror("خطأ", "الكمية والسعر أرقام")
            return

        total = qty * price

        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("SELECT id FROM items WHERE name=?", (name,))
            item_row = c.fetchone()
            item_id = item_row[0] if item_row else None

            c.execute("""UPDATE invoice_items
                         SET item_id=?, item_name=?, quantity=?, price=?, total=?
                         WHERE id=?""",
                      (item_id, name, qty, price, total, self.selected_item_id))
            conn.commit()
            messagebox.showinfo("نجاح", "تم تعديل البند")
            self.load_invoice_data()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل التعديل: {e}")
        finally:
            conn.close()

    def delete_item(self):
        if self.selected_item_id is None:
            messagebox.showwarning("تنبيه", "اختر بنداً للحذف")
            return

        if not messagebox.askyesno("تأكيد", "هل أنت متأكد من حذف هذا البند؟"):
            return

        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM invoice_items WHERE id=?", (self.selected_item_id,))
            conn.commit()
            messagebox.showinfo("نجاح", "تم حذف البند")
            self.load_invoice_data()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()

    def save_all(self):
        """حفظ كل تعديلات الفاتورة (ماعدا المبلغ المدفوع)"""
        inv_no = self.entry_invoice_no.get().strip()
        date = self.entry_date.get().strip()
        supplier = self.combo_trader.get().strip()
        ptype = self.combo_ptype.get()
        notes = self.entry_notes.get().strip()

        if not inv_no:
            messagebox.showwarning("تنبيه", "رقم الفاتورة مطلوب")
            return
        if not date:
            messagebox.showwarning("تنبيه", "التاريخ مطلوب")
            return
        if not supplier:
            messagebox.showwarning("تنبيه", "التاجر مطلوب")
            return

        # ✅ احصل أو أنشئ التاجر الجديد
        trader_id = get_or_create_trader(supplier)
        if trader_id is None:
            messagebox.showerror("خطأ", "فشل في إنشاء التاجر")
            return

        # حساب الإجمالي من البنود
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT SUM(total) FROM invoice_items WHERE invoice_id=?",
                  (self.invoice_id,))
        new_total = c.fetchone()[0] or 0

        # جلب المبلغ المدفوع الحالي (لا نلمسه)
        c.execute("SELECT COALESCE(paid_amount, 0) FROM invoices WHERE id=?",
                  (self.invoice_id,))
        paid = c.fetchone()[0]

        # تحذير إذا كان الإجمالي أقل من المدفوع
        if new_total < paid:
            if not messagebox.askyesno("تحذير",
                    f"الإجمالي الجديد ({new_total:.2f}) أقل من المبلغ المدفوع ({paid:.2f}).\n\n"
                    f"هل تريد المتابعة؟ (المبلغ المدفوع سيبقى كما هو)"):
                conn.close()
                return

        try:
            c.execute("""UPDATE invoices
                         SET invoice_no=?, date=?, supplier=?, notes=?,
                             payment_type=?, total=?, trader_id=?
                         WHERE id=?""",
                      (inv_no, date, supplier, notes, ptype, new_total,
                       trader_id, self.invoice_id))

            conn.commit()
            messagebox.showinfo("نجاح", "تم حفظ التعديلات")

            if self.on_save_callback:
                try:
                    self.on_save_callback()
                except:
                    pass

            self.window.destroy()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}")
        finally:
            conn.close()

# ═══════════════════════════════════════════════════════════════════
#                       ITEMS REPORT WINDOW
# ═══════════════════════════════════════════════════════════════════
class ItemsReportWindow:
    def __init__(self, parent_page):
        self.parent_page = parent_page
        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("تقرير المشتريات حسب الأصناف")
        self.window.geometry("950x650")
        self.window.grab_set()
        self.create_widgets()
        self.generate_report()

    def create_widgets(self):
        filters = tk.LabelFrame(self.window, text="فلاتر التقرير",
                                font=('Arial', 12))
        filters.pack(fill='x', padx=10, pady=5)

        tk.Label(filters, text="من تاريخ:", font=('Arial', 11)).grid(
            row=0, column=0, padx=5, pady=5, sticky='e')
        self.date_from = DateEntry(filters, width=15,
                                   date_pattern='yyyy-mm-dd',
                                   font=('Arial', 11), justify='center')
        self.date_from.grid(row=0, column=1, padx=5, pady=5, sticky='w')





        tk.Label(filters, text="إلى تاريخ:", font=('Arial', 11)).grid(
            row=0, column=2, padx=5, pady=5, sticky='e')
        self.date_to = DateEntry(filters, width=15,
                                   date_pattern='yyyy-mm-dd',
                                   font=('Arial', 11), justify='center')
        self.date_to.grid(row=0, column=3, padx=5, pady=5, sticky='w')


        tk.Label(filters, text="الصنف:", font=('Arial', 11)).grid(
            row=0, column=4, padx=5, pady=5, sticky='e')
        self.item_filter = ttk.Combobox(filters, width=20,
                                        font=('Arial', 11))
        self.item_filter.grid(row=0, column=5, padx=5, pady=5, sticky='w')
        self.load_items_combo()

        # ─── أزرار العمليات ───
        tk.Button(filters, text="📊 عرض التقرير",
                  command=self.generate_report,
                  bg='#16a085', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).grid(
                      row=0, column=6, padx=10)

        tk.Button(filters, text="💾 Excel",
                  command=self.export_excel,
                  bg='#27ae60', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).grid(
                      row=0, column=7, padx=5)

        tk.Button(filters, text="🖼️ صورة",
                  command=self.export_image,
                  bg='#8e44ad', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).grid(
                      row=0, column=8, padx=5)

        # ─── جدول العرض ───
        tree_frame = tk.Frame(self.window)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(tree_frame,
                                  columns=('item', 'qty', 'total', 'count'),
                                  show='headings', height=15)
        self.tree.heading('item', text='الصنف')
        self.tree.heading('qty', text='إجمالي الكمية')
        self.tree.heading('total', text='إجمالي القيمة')
        self.tree.heading('count', text='عدد الفواتير')

        self.tree.column('item', width=250, anchor='center')
        self.tree.column('qty', width=120, anchor='center')
        self.tree.column('total', width=120, anchor='center')
        self.tree.column('count', width=120, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        self.total_label = tk.Label(self.window,
                                     text="الإجمالي العام: 0.00",
                                     font=('Arial', 14, 'bold'),
                                     fg='#e74c3c')
        self.total_label.pack(pady=10)

        tk.Button(self.window, text="إغلاق",
                  command=self.window.destroy,
                  bg='#e74c3c', fg='white',
                  padx=20, pady=5).pack(pady=5)

    # ─────────────────────────────────────────────
    # تحميل قائمة الأصناف
    # ─────────────────────────────────────────────
    def load_items_combo(self):
        conn = self.parent_page.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM items ORDER BY name")
        names = [r[0] for r in c.fetchall()]
        conn.close()
        self.item_filter['values'] = ['الكل'] + names
        self.item_filter.set('الكل')

    # ─────────────────────────────────────────────
    # توليد التقرير
    # ─────────────────────────────────────────────
    def generate_report(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()
        item_name = self.item_filter.get().strip()

        conn = self.parent_page.get_connection()
        c = conn.cursor()

        query = '''SELECT ii.item_name,
        SUM(ii.quantity) as total_qty,
        SUM(ii.total) as total_value,
        COUNT(DISTINCT ii.invoice_id) as invoice_count
        FROM invoice_items ii
        JOIN invoices inv ON ii.invoice_id = inv.id
        WHERE inv.date >= ? AND inv.date <= ?'''
        params = [date_from, date_to]

        if item_name and item_name != 'الكل':
            query += " AND ii.item_name = ?"
            params.append(item_name)

        query += " GROUP BY ii.item_name ORDER BY total_value DESC"

        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        grand_total = 0
        total_qty_sum = 0
        total_invoices = 0

        for row in rows:
            self.tree.insert('', 'end', values=row)
            grand_total += row[2] if row[2] else 0
            total_qty_sum += row[1] if row[1] else 0
            total_invoices += row[3] if row[3] else 0

        self.total_label.config(text=f"الإجمالي العام: {grand_total:.2f}")

        # ✅ حفظ البيانات للتصدير
        self.last_data = {
            'date_from': date_from,
            'date_to': date_to,
            'item_filter': item_name if item_name else 'الكل',
            'rows': rows,
            'grand_total': grand_total,
            'total_qty': total_qty_sum,
            'total_invoices': total_invoices,
            'items_count': len(rows),
        }

    # ─────────────────────────────────────────────
    # التحقق من وجود بيانات
    # ─────────────────────────────────────────────
    def _check_data(self):
        if not hasattr(self, 'last_data') or not self.last_data['rows']:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return False
        return True

    # ─────────────────────────────────────────────
    # تصدير Excel
    # ─────────────────────────────────────────────
    def export_excel(self):
        if not self._check_data():
            return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import (Font, Alignment, PatternFill,
                                         Border, Side)
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        data = self.last_data
        default_name = f"تقرير_أصناف_{data['date_from']}_{data['date_to']}.xlsx"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="حفظ تقرير الأصناف Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "تقرير الأصناف"
            ws.sheet_view.rightToLeft = True

            # ─── عرض الأعمدة ───
            col_widths = {'A': 30, 'B': 16, 'C': 18, 'D': 16}
            for col, w in col_widths.items():
                ws.column_dimensions[col].width = w

            # ─── العنوان ───
            ws.merge_cells('A1:D1')
            c = ws['A1']
            c.value = (f"📊 تقرير المشتريات حسب الأصناف "
                       f"({data['date_from']} - {data['date_to']})")
            c.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            c.fill = PatternFill("solid", fgColor="2C3E50")
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ─── سطر الصنف المفلتر ───
            ws.merge_cells('A2:D2')
            c = ws['A2']
            c.value = f"الصنف: {data['item_filter']}"
            c.font = Font(name='Arial', size=11, bold=True, color='8B4513')
            c.fill = PatternFill("solid", fgColor="FFF3CD")
            c.alignment = Alignment(horizontal='center', vertical='center')

            # ─── الرؤوس ───
            headers = ['الصنف', 'إجمالي الكمية',
                       'إجمالي القيمة', 'عدد الفواتير']
            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=3, column=col_idx, value=h)
                cell.font = Font(name='Arial', size=11, bold=True,
                                 color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            # ─── البيانات ───
            row_num = 4
            for r in data['rows']:
                item_name = r[0] if r[0] else ''
                qty = r[1] if r[1] else 0
                total = r[2] if r[2] else 0
                count = r[3] if r[3] else 0

                ws.cell(row=row_num, column=1, value=item_name)
                ws.cell(row=row_num, column=2, value=qty)
                ws.cell(row=row_num, column=3, value=total)
                ws.cell(row=row_num, column=4, value=count)

                for col_idx in range(1, 5):
                    cell = ws.cell(row=row_num, column=col_idx)
                    cell.font = Font(name='Arial', size=11)
                    cell.alignment = Alignment(horizontal='center',
                                               vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin'))
                row_num += 1

            # ═════════ صفوف الإجماليات ═════════
            row_num += 1
            totals_data = [
                ("👥 عدد الأصناف", data['items_count'], False),
                ("📦 إجمالي الكميات", data['total_qty'], False),
                ("🧾 إجمالي عدد الفواتير", data['total_invoices'], False),
                ("💰 الإجمالي العام", data['grand_total'], True),
            ]

            for label, value, is_grand in totals_data:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=3)

                bg = "F1C40F" if is_grand else "FFF3CD"
                fg = "2C3E50" if is_grand else "8B4513"

                # خلية التسمية
                cell = ws.cell(row=row_num, column=1, value=label)
                cell.font = Font(name='Arial', size=12, bold=True,
                                 color=fg)
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal='right',
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

                # خلية القيمة
                cell = ws.cell(row=row_num, column=4, value=value)
                cell.font = Font(name='Arial', size=12, bold=True,
                                 color='C0392B')
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

                row_num += 1

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    # ─────────────────────────────────────────────
    # تصدير صورة
    # ─────────────────────────────────────────────
    def export_image(self):
        """تصدير تقرير الأصناف إلى صورة PNG"""
        if not self._check_data():
            return

        from tkinter import filedialog

        data = self.last_data
        default_name = f"تقرير_أصناف_{data['date_from']}_{data['date_to']}.png"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=default_name,
            title="حفظ تقرير الأصناف صورة"
        )
        if not file_path:
            return

        columns = ['الصنف', 'إجمالي الكمية',
                   'إجمالي القيمة', 'عدد الفواتير']

        rows_data = []
        for r in data['rows']:
            rows_data.append((
                r[0] if r[0] else '',
                f"{r[1]:,.2f}" if r[1] else '0.00',
                f"{r[2]:,.2f}" if r[2] else '0.00',
                str(r[3]) if r[3] else '0',
            ))

        summary_items = [
            ("👥 عدد الأصناف", str(data['items_count'])),
            ("📦 إجمالي الكميات", f"{data['total_qty']:,.2f}"),
            ("🧾 إجمالي عدد الفواتير", str(data['total_invoices'])),
            ("─────", ""),
            ("💰 الإجمالي العام (ج.م)", f"{data['grand_total']:,.2f}"),
        ]

        # ✅ نقلنا "عدد الأصناف" إلى subtitle بدلاً من bottom_note
        subtitle = (f"من {data['date_from']} إلى {data['date_to']} | "
                    f"الصنف: {data['item_filter']} | "
                    f"عدد الأصناف: {data['items_count']}")

        success = export_table_to_image(
            file_path=file_path,
            title="📊 تقرير المشتريات حسب الأصناف",
            subtitle=subtitle,
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass
# ═══════════════════════════════════════════════════════════════════
# ITEM FORM WINDOW (نافذة إضافة/تعديل صنف)
# ═══════════════════════════════════════════════════════════════════
class ItemFormWindow:
    """نافذة إدخال/تعديل صنف"""
    def __init__(self, parent_page, item_id=None):
        self.parent_page = parent_page
        self.item_id = item_id
        self.is_edit = item_id is not None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل صنف" if self.is_edit else "إضافة صنف جديد"
        self.window.title(title)
        self.window.geometry("500x300")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ─── العنوان ───
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        icon = "✏️" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'),
                 bg='#2c3e50', fg='white').pack(pady=15)

        # ─── الإطار الرئيسي ───
        main = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main.pack(fill='both', expand=True)

        tk.Label(main, text="اسم الصنف: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=0, column=0, padx=5, pady=10, sticky='e')
        self.entry_name = tk.Entry(main, width=28, font=('Arial', 11))
        self.entry_name.grid(row=0, column=1, padx=5, pady=10, sticky='w')

        tk.Label(main, text="الوحدة:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(
                 row=1, column=0, padx=5, pady=10, sticky='e')
        self.entry_unit = tk.Entry(main, width=28, font=('Arial', 11))
        self.entry_unit.grid(row=1, column=1, padx=5, pady=10, sticky='w')

        # ─── أزرار الحفظ ───
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        if self.is_edit:
            self.load_data()

        self.entry_name.focus()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def load_data(self):
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name, unit FROM items WHERE id=?", (self.item_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            messagebox.showerror("خطأ", "الصنف غير موجود")
            self.window.destroy()
            return
        self.entry_name.insert(0, row[0] or '')
        self.entry_unit.insert(0, row[1] or '')

    def save(self):
        name = self.entry_name.get().strip()
        unit = self.entry_unit.get().strip()

        if not name:
            messagebox.showwarning("تنبيه", "اسم الصنف مطلوب")
            self.entry_name.focus()
            return

        conn = self.get_connection()
        c = conn.cursor()
        try:
            if self.is_edit:
                c.execute("UPDATE items SET name=?, unit=? WHERE id=?",
                          (name, unit, self.item_id))
                msg = f"✅ تم تعديل الصنف '{name}'"
            else:
                c.execute(
                    "INSERT INTO items (name, unit, created_at) VALUES (?, ?, ?)",
                    (name, unit,
                     datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                msg = f"✅ تمت إضافة الصنف '{name}'"
            conn.commit()
            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_page.load_items()
        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "هذا الاسم موجود مسبقاً")
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}")
        finally:
            conn.close()


# ═══════════════════════════════════════════════════════════════════
# INVENTORY PAGE (بنمط موحّد مع باقي الصفحات)
# ═══════════════════════════════════════════════════════════════════
class InventoryPage(tk.Frame):
    """صفحة إدارة المخزن"""

    def __init__(self, parent, controller=None):
        super().__init__(parent)
        self.controller = controller
        self.configure(bg='#ecf0f1')
        self.selected_id = None

        # 1) الشريط العلوي
        add_program_banner(self)

        # 2) شريط الحالة (أسفل)
        self.status_bar = tk.Label(self, text="جاهز",
                                   font=('Arial', 10),
                                   bg='#2c3e50', fg='white',
                                   anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # 3) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ═══════════════ العنوان ═══════════════
        tk.Label(content, text="📦 إدارة المخزن",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=12)

        # ═══════════════ إطار الأزرار ═══════════════
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                  font=('Arial', 11, 'bold'),
                                  bg='#ecf0f1', fg='#34495e',
                                  padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة صنف",
                  command=self.add_item,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏️ تعديل",
                                  command=self.edit_item,
                                  bg='#f39c12', fg='white',
                                  font=('Arial', 11, 'bold'),
                                  padx=20, pady=8, width=15,
                                  state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(row1, text="🗑️ حذف",
                                    command=self.delete_item,
                                    bg='#e74c3c', fg='white',
                                    font=('Arial', 11, 'bold'),
                                    padx=20, pady=8, width=15,
                                    state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث",
                  command=self.load_items,
                  bg='#95a5a6', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        tk.Button(row1, text="📊 جرد المخزن",
                  command=self.open_inventory_report,
                  bg='#16a085', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        # ═══════════════ إطار البحث ═══════════════
        search_frame = tk.Frame(content, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=15, pady=5)

        tk.Label(search_frame, text="🔍 بحث:",
                 font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_entry = tk.Entry(search_frame, width=35,
                                     font=('Arial', 11))
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind(
            '<KeyRelease>',
            lambda e: self.load_items(self.search_entry.get().strip()))

        tk.Button(search_frame, text="🗑 مسح",
                  command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50',
                  font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        # ═══════════════ جدول العرض ═══════════════
        tree_frame = tk.LabelFrame(content, text="قائمة الأصناف",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 15))

        columns = ('id', 'name', 'unit', 'created_at')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=15)

        self.tree.heading('id', text='ID')
        self.tree.heading('name', text='اسم الصنف')
        self.tree.heading('unit', text='الوحدة')
        self.tree.heading('created_at', text='تاريخ الإضافة')

        self.tree.column('id', width=60, anchor='center')
        self.tree.column('name', width=300, anchor='center')
        self.tree.column('unit', width=150, anchor='center')
        self.tree.column('created_at', width=180, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree.yview)

        # ═══════════════════════════════════════════════════════════
        # 🎨 تنسيق الصفوف
        # ═══════════════════════════════════════════════════════════
        # 1) Zebra Stripes (صفوف متبادلة)
        self.tree.tag_configure('odd_row', background='#ebf5fb')
        self.tree.tag_configure('even_row', background='#ffffff')
        self.tree.tag_configure('hover', background='#d6eaf8', foreground='#1a5276')
        self.tree.tag_configure('used_item', background='#fef9e7', foreground='#7d6608')

        # 2) لون التحذير (الأصناف المستخدمة في الفواتير)
        self.tree.tag_configure('used_item',
                                background='#fdebd0',
                                foreground='#a04000')

        # 3) لون التمييز عند المرور بالماوس (سنستخدم متغير)
        self.hovered_item = None

        # ─── دالة تأثير Hover ───
        def on_tree_hover(event):
            item = self.tree.identify_row(event.y)
            if item != self.hovered_item:
                # إزالة تأثير الصف السابق
                if self.hovered_item:
                    try:
                        prev_values = self.tree.item(self.hovered_item)
                        tags = prev_values.get('tags', [])
                        # استبعاد tag الهوفر فقط
                        tags = [t for t in tags if t != 'hover']
                        self.tree.item(self.hovered_item, tags=tags)
                    except tk.TclError:
                        pass
                # تطبيق التأثير على الصف الجديد
                self.hovered_item = item
                if item:
                    try:
                        curr = self.tree.item(item)
                        tags = list(curr.get('tags', []))
                        if 'hover' not in tags:
                            tags.append('hover')
                        self.tree.item(item, tags=tags)
                    except tk.TclError:
                        pass

        def on_tree_leave(event):
            """إزالة تأثير Hover عند الخروج من الجدول"""
            if self.hovered_item:
                try:
                    curr = self.tree.item(self.hovered_item)
                    tags = [t for t in curr.get('tags', []) if t != 'hover']
                    self.tree.item(self.hovered_item, tags=tags)
                except tk.TclError:
                    pass
                self.hovered_item = None

        # لون الـ hover
        self.tree.tag_configure('hover',
                                background='#d6eaf8',  # أزرق فاتح
                                foreground='#1a5276')

        self.tree.bind('<Motion>', on_tree_hover)
        self.tree.bind('<Leave>', on_tree_leave)

        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_item())

        # التحميل الأولي
        self.load_items()

    # ═══════════════════════════════════════════════════════════
    # المساعدات
    # ═══════════════════════════════════════════════════════════
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def load_traders(self):
        """(يُستدعى تلقائياً من MainApp عند فتح الصفحة)"""
        self.load_items()

    def load_items(self, search=''):
        """تحميل قائمة الأصناف مع تلوين الصفوف"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = self.get_connection()
        c = conn.cursor()
        if search:
            c.execute("""SELECT id, name, unit, created_at FROM items
                         WHERE name LIKE ? ORDER BY name""",
                      (f'%{search}%',))
        else:
            c.execute("""SELECT id, name, unit, created_at
                         FROM items ORDER BY name""")

        rows = c.fetchall()

        # ─── اجلب عدد مرات الاستخدام لكل صنف ───
        c.execute("""SELECT item_id, COUNT(*) FROM invoice_items
                     WHERE item_id IS NOT NULL
                     GROUP BY item_id""")
        usage_map = dict(c.fetchall())
        conn.close()

        # ─── إدراج الصفوف مع التلوين ───
        for i, row in enumerate(rows):
            item_id, name, unit, created_at = row
            used_count = usage_map.get(item_id, 0)

            base_tag = 'odd_row' if i % 2 == 0 else 'even_row'

            # إدراج الصف
            self.tree.insert('', 'end',
                             iid=str(item_id),
                             values=(item_id, name, unit or '', created_at or ''),
                             tags=(base_tag,))

        self.status_bar.config(text=f"عدد الأصناف: {len(rows)}")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.load_items()

    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None
                return
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    # ═══════════════════════════════════════════════════════════
    # CRUD
    # ═══════════════════════════════════════════════════════════
    def add_item(self):
        ItemFormWindow(self, item_id=None)

    def edit_item(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر صنفاً من الجدول أولاً")
            return
        ItemFormWindow(self, item_id=self.selected_id)

    def delete_item(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر صنفاً من الجدول أولاً")
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM items WHERE id=?", (self.selected_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return
        item_name = row[0]

        # فحص الاستخدام في الفواتير
        c.execute("SELECT COUNT(*) FROM invoice_items WHERE item_id=?",
                  (self.selected_id,))
        count = c.fetchone()[0]
        conn.close()

        if count > 0:
            messagebox.showerror(
                "خطأ",
                f"لا يمكن حذف الصنف '{item_name}'،\n"
                f"لأنه مستخدم في {count} فاتورة")
            return

        if not messagebox.askyesno(
                "تأكيد",
                f"هل أنت متأكد من حذف الصنف '{item_name}'؟"):
            return

        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM items WHERE id=?", (self.selected_id,))
            conn.commit()
            messagebox.showinfo("نجاح", f"تم حذف الصنف '{item_name}'")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_items()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()

    def open_inventory_report(self):
        InventoryReportWindow(self)

class InventoryReportWindow:
    def __init__(self, parent_page):
        self.parent_page = parent_page
        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("📊 جرد المخزن الشهري")
        self.window.geometry("1400x750")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')
        self.selected_item_id = None
        self.create_widgets()
        self.load_report()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def get_opening_balance(self, conn, item_id, year, month):
        c = conn.cursor()
        c.execute("""SELECT opening_balance FROM inventory_balances
                     WHERE item_id=? AND year=? AND month=?""", (item_id, year, month))
        row = c.fetchone()
        if row:
            return row[0] or 0
        prev_y, prev_m = get_prev_month(year, month)
        c.execute("""SELECT closing_balance FROM inventory_balances
                     WHERE item_id=? AND year=? AND month=?""", (item_id, prev_y, prev_m))
        row = c.fetchone()
        return (row[0] or 0) if row else 0

    def get_closing_balance(self, conn, item_id, year, month):
        c = conn.cursor()
        c.execute("""SELECT closing_balance FROM inventory_balances
                     WHERE item_id=? AND year=? AND month=?""", (item_id, year, month))
        row = c.fetchone()
        return (row[0] or 0) if row else 0

    def save_month_balance(self, item_id, year, month, opening, closing):
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT id FROM inventory_balances
                     WHERE item_id=? AND year=? AND month=?""", (item_id, year, month))
        existing = c.fetchone()
        if existing:
            c.execute("""UPDATE inventory_balances
                         SET opening_balance=?, closing_balance=?, updated_at=?
                         WHERE id=?""", (opening, closing, now, existing[0]))
        else:
            c.execute("""INSERT INTO inventory_balances
                         (item_id, year, month, opening_balance, closing_balance, created_at, updated_at)
                         VALUES (?, ?, ?, ?, ?, ?, ?)""",
                      (item_id, year, month, opening, closing, now, now))
        # الترحيل للشهر التالي
        next_y, next_m = get_next_month(year, month)
        c.execute("""SELECT id FROM inventory_balances
                     WHERE item_id=? AND year=? AND month=?""", (item_id, next_y, next_m))
        next_row = c.fetchone()
        if next_row:
            c.execute("""UPDATE inventory_balances SET opening_balance=?, updated_at=?
                         WHERE id=?""", (closing, now, next_row[0]))
        else:
            c.execute("""INSERT INTO inventory_balances
                         (item_id, year, month, opening_balance, closing_balance, created_at, updated_at)
                         VALUES (?, ?, ?, ?, 0, ?, ?)""",
                      (item_id, next_y, next_m, closing, now, now))
        conn.commit()
        conn.close()

    def create_widgets(self):
        current_year = datetime.now().year
        current_month = datetime.now().month

        title_frame = tk.Frame(self.window, bg='#2c3e50', height=50)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text="📊 جرد المخزن الشهري",
                 font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white').pack(pady=12)

        period_frame = tk.LabelFrame(self.window, text="📅 اختر الفترة",
                                      font=('Arial', 12, 'bold'), bg='#d5f4e6', fg='#145a32',
                                      padx=10, pady=10)
        period_frame.pack(fill='x', padx=10, pady=5)

        tk.Label(period_frame, text="السنة:", font=('Arial', 11, 'bold'), bg='#d5f4e6').grid(row=0, column=0, padx=5, pady=5)
        self.year_var = tk.StringVar(value=str(current_year))
        ttk.Combobox(period_frame, textvariable=self.year_var,
                     values=[str(y) for y in range(current_year - 2, current_year + 3)],
                     width=10, font=('Arial', 11), state='readonly').grid(row=0, column=1, padx=5, pady=5)

        tk.Label(period_frame, text="الشهر:", font=('Arial', 11, 'bold'), bg='#d5f4e6').grid(row=0, column=2, padx=5, pady=5)
        month_values = [f"{i} - {ARABIC_MONTHS[i]}" for i in range(1, 13)]
        self.month_combo = ttk.Combobox(period_frame, values=month_values,
                                         width=15, font=('Arial', 11), state='readonly')
        self.month_combo.set(f"{current_month} - {ARABIC_MONTHS[current_month]}")
        self.month_combo.grid(row=0, column=3, padx=5, pady=5)

        tk.Button(period_frame, text="🔍 عرض الشهر", command=self.load_report,
                  bg='#16a085', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=5).grid(row=0, column=4, padx=10)

        tk.Button(period_frame, text="◀ السابق", command=self.go_prev_month,
                  bg='#7f8c8d', fg='white', padx=10, pady=5).grid(row=0, column=5, padx=5)
        tk.Button(period_frame, text="التالي ▶", command=self.go_next_month,
                  bg='#7f8c8d', fg='white', padx=10, pady=5).grid(row=0, column=6, padx=5)

        # إطار تعديل الرصيد
        balance_frame = tk.LabelFrame(self.window, text="✏️ إدخال رصيد آخر المدة",
                                       font=('Arial', 12, 'bold'), bg='#fff9e6', fg='#8B4513',
                                       padx=10, pady=10)
        balance_frame.pack(fill='x', padx=10, pady=5)

        tk.Label(balance_frame, text="الصنف:", font=('Arial', 11, 'bold'), bg='#fff9e6').grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.lbl_selected_item = tk.Label(balance_frame, text="(لم يتم تحديد صنف)",
                                           font=('Arial', 11, 'bold'), bg='#fff9e6', fg='#c0392b',
                                           width=25, anchor='w')
        self.lbl_selected_item.grid(row=0, column=1, padx=5, pady=5, sticky='w')

        tk.Label(balance_frame, text="رصيد أول المدة:", font=('Arial', 11), bg='#fff9e6').grid(row=0, column=2, padx=5, pady=5, sticky='e')
        self.entry_opening = tk.Entry(balance_frame, width=12, font=('Arial', 11, 'bold'),
                                       bg='#ecf0f1', justify='center', state='readonly')
        self.entry_opening.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(balance_frame, text="رصيد آخر المدة:", font=('Arial', 11, 'bold'),
                 bg='#fff9e6', fg='#c0392b').grid(row=0, column=4, padx=5, pady=5, sticky='e')
        self.entry_closing = tk.Entry(balance_frame, width=12, font=('Arial', 12, 'bold'),
                                       bg='#fffacd', justify='center')
        self.entry_closing.grid(row=0, column=5, padx=5, pady=5)

        tk.Button(balance_frame, text="💾 حفظ وترحيل", command=self.save_balances,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8).grid(row=0, column=6, padx=15)

        # جدول العرض
        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ('id', 'name', 'unit', 'opening', 'purchases',
                   'closing', 'sold', 'price', 'sold_value', 'stock_value')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=14)

        headers = {'id': 'ID', 'name': 'اسم الصنف', 'unit': 'الوحدة',
                   'opening': 'رصيد أول', 'purchases': 'المشتريات',
                   'closing': 'رصيد آخر', 'sold': 'المباع', 'price': 'السعر',
                   'sold_value': 'قيمة المباع', 'stock_value': 'قيمة المخزون'}
        for col, header in headers.items():
            self.tree.heading(col, text=header)
        for col in columns:
            self.tree.column(col, width=100, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)

        tk.Button(self.window, text="إغلاق", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=20, pady=5).pack(pady=5)

    def get_selected_period(self):
        year = int(self.year_var.get())
        month = int(self.month_combo.get().split(' - ')[0])
        return year, month

    def go_prev_month(self):
        year, month = self.get_selected_period()
        year, month = get_prev_month(year, month)
        self.year_var.set(str(year))
        self.month_combo.set(f"{month} - {ARABIC_MONTHS[month]}")
        self.load_report()

    def go_next_month(self):
        year, month = self.get_selected_period()
        year, month = get_next_month(year, month)
        self.year_var.set(str(year))
        self.month_combo.set(f"{month} - {ARABIC_MONTHS[month]}")
        self.load_report()

    def load_report(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        year, month = self.get_selected_period()
        if month == 12:
            next_year, next_month = year + 1, 1
        else:
            next_year, next_month = year, month + 1
        date_from = f"{year}-{month:02d}-01"
        date_to = f"{next_year}-{next_month:02d}-01"

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT id, name, unit FROM items ORDER BY name")
        items = c.fetchall()

        for item_id, item_name, unit in items:
            opening = self.get_opening_balance(conn, item_id, year, month)
            closing = self.get_closing_balance(conn, item_id, year, month)

            c.execute("""SELECT COALESCE(SUM(ii.quantity), 0)
                         FROM invoice_items ii
                         JOIN invoices inv ON ii.invoice_id = inv.id
                         WHERE ii.item_id = ? AND inv.date >= ? AND inv.date < ?""",
                      (item_id, date_from, date_to))
            purchases = c.fetchone()[0] or 0

            c.execute("""SELECT ii.price FROM invoice_items ii
                         JOIN invoices inv ON ii.invoice_id = inv.id
                         WHERE ii.item_id = ?
                         ORDER BY inv.date DESC, ii.id DESC LIMIT 1""", (item_id,))
            price_row = c.fetchone()
            last_price = price_row[0] if price_row else 0

            sold = opening + purchases - closing
            values = (item_id, item_name, unit or '',
                      f"{opening:.2f}", f"{purchases:.2f}", f"{closing:.2f}",
                      f"{sold:.2f}", f"{last_price:.2f}",
                      f"{sold * last_price:.2f}", f"{closing * last_price:.2f}")
            self.tree.insert('', 'end', iid=str(item_id), values=values)

        conn.close()

    def on_item_select(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        item_id = int(selected[0])
        self.selected_item_id = item_id
        year, month = self.get_selected_period()
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM items WHERE id=?", (item_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return
        item_name = row[0]
        opening = self.get_opening_balance(conn, item_id, year, month)
        closing = self.get_closing_balance(conn, item_id, year, month)
        conn.close()
        self.lbl_selected_item.config(text=item_name, fg='#27ae60')
        self.entry_opening.config(state='normal')
        self.entry_opening.delete(0, tk.END)
        self.entry_opening.insert(0, f"{opening:.2f}")
        self.entry_opening.config(state='readonly')
        self.entry_closing.delete(0, tk.END)
        self.entry_closing.insert(0, f"{closing:.2f}")

    def save_balances(self):
        if self.selected_item_id is None:
            messagebox.showwarning("تنبيه", "اختر صنفاً أولاً")
            return
        try:
            closing = float(self.entry_closing.get().strip() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "أدخل رقماً صحيحاً")
            return
        year, month = self.get_selected_period()
        conn = self.get_connection()
        opening = self.get_opening_balance(conn, self.selected_item_id, year, month)
        conn.close()
        self.save_month_balance(self.selected_item_id, year, month, opening, closing)
        messagebox.showinfo("✅ نجاح", f"تم حفظ رصيد آخر المدة: {closing:.2f}")
        self.load_report()
class TraderFormWindow:
    """نافذة إدخال/تعديل بيانات تاجر"""
    def __init__(self, parent_page, trader_id=None):
        self.parent_page = parent_page
        self.trader_id = trader_id
        self.is_edit = trader_id is not None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل بيانات تاجر" if self.is_edit else "إضافة تاجر جديد"
        self.window.title(title)
        self.window.geometry("520x450")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ============ العنوان ============
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏️" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white').pack(pady=15)

        # ============ الإطار الرئيسي ============
        main_frame = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main_frame.pack(fill='both', expand=True)

        # اسم التاجر
        tk.Label(main_frame, text="اسم التاجر: *",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='e').grid(row=0, column=0, padx=5, pady=10, sticky='e')
        self.entry_name = tk.Entry(main_frame, width=28, font=('Arial', 11))
        self.entry_name.grid(row=0, column=1, padx=5, pady=10, sticky='w')

        # الهاتف
        tk.Label(main_frame, text="الهاتف:",
                 font=('Arial', 11), bg='#ecf0f1',
                 anchor='e').grid(row=1, column=0, padx=5, pady=10, sticky='e')
        self.entry_phone = tk.Entry(main_frame, width=28, font=('Arial', 11))
        self.entry_phone.grid(row=1, column=1, padx=5, pady=10, sticky='w')

        # ملاحظات
        tk.Label(main_frame, text="ملاحظات:",
                 font=('Arial', 11), bg='#ecf0f1',
                 anchor='ne').grid(row=2, column=0, padx=5, pady=10, sticky='ne')
        self.entry_notes = tk.Text(main_frame, width=28, height=5, font=('Arial', 11))
        self.entry_notes.grid(row=2, column=1, padx=5, pady=10, sticky='w')

        # معلومة توضيحية
        info_frame = tk.Frame(self.window, bg='#e8f8f0', pady=8)
        info_frame.pack(fill='x', padx=20, pady=5)
        tk.Label(info_frame,
                 text="💡 سيُضاف التاجر تلقائياً تحت تبويب 'مشتريات'",
                 font=('Arial', 10, 'italic'), bg='#e8f8f0', fg='#27ae60').pack()

        # ============ أزرار الحفظ والإلغاء ============
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        # تحميل البيانات إن كان تعديلاً
        if self.is_edit:
            self.load_trader_data()

        self.entry_name.focus()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def load_trader_data(self):
        """جلب بيانات التاجر لتعبئة الحقول"""
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT name, phone, notes FROM traders WHERE id=?""",
                  (self.trader_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "التاجر غير موجود")
            self.window.destroy()
            return

        name, phone, notes = row

        self.entry_name.insert(0, name or '')
        self.entry_phone.insert(0, phone or '')
        self.entry_notes.insert('1.0', notes or '')

    def save(self):
        """حفظ بيانات التاجر (إضافة أو تعديل)"""
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("تنبيه", "اسم التاجر مطلوب")
            self.entry_name.focus()
            return

        phone = self.entry_phone.get().strip()
        notes = self.entry_notes.get('1.0', 'end').strip()

        conn = self.get_connection()
        c = conn.cursor()

        try:
            if self.is_edit:
                # ✅ تعديل
                c.execute("SELECT name FROM traders WHERE id=?", (self.trader_id,))
                old_row = c.fetchone()
                old_name = old_row[0] if old_row else None

                c.execute("""UPDATE traders 
                             SET name=?, phone=?, notes=?
                             WHERE id=?""",
                          (name, phone, notes, self.trader_id))

                # تحديث الاسم في sub_categories و paids و invoices
                if old_name and old_name != name:
                    c.execute("UPDATE sub_categories SET name=? WHERE name=?",
                              (name, old_name))
                    c.execute("UPDATE paids SET sub_name=? WHERE sub_name=?",
                              (name, old_name))
                    c.execute("UPDATE invoices SET supplier=? WHERE supplier=?",
                              (name, old_name))

                msg = f"✅ تم تعديل بيانات '{name}'"
            else:
                # ✅ إضافة جديدة
                c.execute("""INSERT INTO traders (name, phone, notes, created_at)
                             VALUES (?, ?, ?, ?)""",
                          (name, phone, notes,
                           datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

                # إضافته تحت "مشتريات"
                c.execute("SELECT id FROM main_categories WHERE name='مشتريات'")
                main_row = c.fetchone()
                if main_row:
                    main_id = main_row[0]
                    c.execute("SELECT id FROM sub_categories WHERE main_id=? AND name=?",
                              (main_id, name))
                    if not c.fetchone():
                        c.execute("INSERT INTO sub_categories (main_id, name) VALUES (?, ?)",
                                  (main_id, name))

                msg = f"✅ تمت إضافة التاجر '{name}'"

            conn.commit()
            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_page.load_traders()

        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "هذا الاسم موجود مسبقاً")
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}")
        finally:
            conn.close()


# ═══════════════════════════════════════════════════════════════════
#                       TRADERS PAGE
# ═══════════════════════════════════════════════════════════════════
class TradersPage(tk.Frame):
    """صفحة إدارة التجار (أزرار + جدول فقط)"""
    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None
        # 1️⃣ الشريط العلوي
        add_program_banner(self)

        # 2️⃣ منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ============ العنوان ============
        tk.Label(content, text="🧑‍💼 إدارة التجار والموردين",
                 font=('Arial', 18, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=12)

        # ============ شريط الأزرار ============
        btn_frame = tk.LabelFrame(content, text="العمليات",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        # الصف الأول من الأزرار
        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة تاجر", command=self.add_trader,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏️ تعديل", command=self.edit_trader,
                                    bg='#f39c12', fg='white', font=('Arial', 11, 'bold'),
                                    padx=20, pady=8, width=15, state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(row1, text="🗑️ حذف", command=self.delete_trader,
                                      bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                                      padx=20, pady=8, width=15, state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        # الصف الثاني من الأزرار
        row2 = tk.Frame(btn_frame, bg='#ecf0f1')
        row2.pack(pady=3)

        self.btn_statement = tk.Button(row1, text="📄 كشف الحساب",
                                         command=self.open_statement,
                                         bg='#8e44ad', fg='white',
                                         font=('Arial', 11, 'bold'),
                                         padx=20, pady=8, width=15, state='disabled')
        self.btn_statement.pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث", command=self.load_traders,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        tk.Button(row1, text="📊 إجمالي المديونيات",
                  command=self.show_total_summary,
                  bg='#2980b9', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=18).pack(side='left', padx=5)

        # ============ شريط البحث ============
        search_frame = tk.Frame(content, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=15, pady=5)

        tk.Label(search_frame, text="🔎 بحث:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_entry = tk.Entry(search_frame, width=30, font=('Arial', 11))
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind('<KeyRelease>',
                                lambda e: self.load_traders(self.search_entry.get().strip()))

        tk.Button(search_frame, text="🔄 مسح", command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50', font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        # ============ جدول العرض ============
        tree_frame = tk.LabelFrame(content, text="قائمة التجار",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 10))

        columns = ('id', 'name', 'phone', 'notes', 'debt')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)

        headers = {
            'id': 'ID',
            'name': 'اسم التاجر',
            'phone': 'الهاتف',
            'notes': 'ملاحظات',
            'debt': 'الرصيد المستحق (ج.م)'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=50, anchor='center')
        self.tree.column('name', width=250, anchor='center')
        self.tree.column('phone', width=140, anchor='center')
        self.tree.column('notes', width=300, anchor='w')
        self.tree.column('debt', width=180, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        # ✅ تلوين: أحمر فاتح (له مديونية) / أخضر فاتح (مسدد)
        self.tree.tag_configure('owe_them', background='#fadbd8')
        self.tree.tag_configure('paid', background='#d4edda')

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.open_statement())

        # شريط الحالة
        self.status_bar = tk.Label(content, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#34495e', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # التحميل الأولي
        self.load_traders()

    # ─────────────────────────────────────────────────────────
    def get_connection(self):
        return sqlite3.connect(DB_NAME)



    def load_traders(self, search=''):
        """عرض كل التجار مع رصيدهم الصحيح"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = self.get_connection()
        c = conn.cursor()
        if search:
            c.execute("""SELECT id, name, phone, notes FROM traders 
    	                     WHERE name LIKE ? ORDER BY name""",
                      (f'%{search}%',))
        else:
            c.execute("SELECT id, name, phone, notes FROM traders ORDER BY name")
        rows = c.fetchall()

        total_we_owe = 0  # إجمالي المديونيات لنا (المتبقي على كل التجار)

        for tid, name, phone, notes in rows:
            # ✅ استخدام الدالة المُصححة لحساب الرصيد
            debt = calc_trader_balance(conn, name)

            # ✅ تحديد اللون حسب الحالة
            if debt > 0.01:
                tag = 'owe_them'  # له مديونية (علينا)
                total_we_owe += debt
            else:
                tag = 'paid'  # لا مديونية (مسدد)

            notes_display = (notes or '').replace('\n', ' ')[:50]

            self.tree.insert('', 'end', iid=tid, tags=(tag,), values=(
                tid, name, phone or '', notes_display, f"{debt:,.2f}"))

        conn.close()

        # ✅ شريط الحالة: عدد كل التجار + إجمالي المديونيات
        self.status_bar.config(
            text=f"عدد التجار: {len(rows)}  |  "
                 f"إجمالي المديونيات المتبقية: {total_we_owe:,.2f} ج.م")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.load_traders()

    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            self.selected_id = int(sel[0])
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
            self.btn_statement.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.btn_statement.config(state='disabled')

    def add_trader(self):
        TraderFormWindow(self, trader_id=None)

    def edit_trader(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً من الجدول أولاً")
            return
        TraderFormWindow(self, trader_id=self.selected_id)

    def delete_trader(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً من الجدول أولاً")
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM traders WHERE id=?", (self.selected_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return
        trader_name = row[0]

        # التحقق من الفواتير
        c.execute("SELECT COUNT(*) FROM invoices WHERE trader_id=?", (self.selected_id,))
        inv_count = c.fetchone()[0]

        # التحقق من الحركات في paids
        c.execute("SELECT COUNT(*) FROM paids WHERE sub_name=?", (trader_name,))
        paid_count = c.fetchone()[0]

        conn.close()

        warn_msg = f"هل أنت متأكد من حذف التاجر '{trader_name}'؟"
        if inv_count > 0 or paid_count > 0:
            warn_msg += f"\n\n⚠️ لديه:\n• {inv_count} فاتورة\n• {paid_count} حركة في المصاريف"
            warn_msg += "\n\n(لن تُحذف الفواتير أو الحركات، فقط التاجر)"

        if not messagebox.askyesno("تأكيد", warn_msg):
            return

        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM traders WHERE id=?", (self.selected_id,))
            conn.commit()
            messagebox.showinfo("نجاح", f"تم حذف '{trader_name}'")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.btn_statement.config(state='disabled')
            self.load_traders()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()

    def open_statement(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر تاجراً من الجدول أولاً")
            return
        TraderStatementWindow(self, self.selected_id)

    def show_total_summary(self):
        """ملخص إجمالي المديونيات المتبقية علينا"""
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM traders ORDER BY name")
        traders = c.fetchall()

        total_due = 0
        details = []

        for (name,) in traders:
            debt = calc_trader_balance(conn, name)
            if debt > 0.01:  # فقط من له مبلغ متبقٍ
                total_due += debt
                details.append((name, debt))
        conn.close()

        if not details:
            messagebox.showinfo("لا توجد مديونيات",
                                "لا يوجد أي تاجر له مبالغ متبقية علينا. 🎉")
            return

        win = tk.Toplevel(self.winfo_toplevel())
        win.title("ملخص المديونيات المتبقية")
        win.geometry("650x550")
        win.grab_set()
        win.configure(bg='#ecf0f1')

        # ═════════ العنوان ═════════
        tk.Label(win, text="💰 إجمالي المديونيات المتبقية علينا",
                 font=('Arial', 16, 'bold'),
                 bg='#2c3e50', fg='white').pack(fill='x', ipady=12)

        # ═════════ الإجمالي ═════════
        total_frame = tk.Frame(win, bg='#e8f4f8', pady=15,
                               highlightbackground='#2980b9',
                               highlightthickness=2)
        total_frame.pack(fill='x', padx=15, pady=15)

        tk.Label(total_frame,
                 text=f"المبلغ الإجمالي المتبقي للتجار: {total_due:,.2f} ج.م",
                 font=('Arial', 15, 'bold'),
                 bg='#e8f4f8', fg='#c0392b').pack()
        tk.Label(total_frame,
                 text=f"عدد التجار المدينين: {len(details)}",
                 font=('Arial', 11),
                 bg='#e8f4f8', fg='#34495e').pack(pady=(5, 0))

        # ═════════ جدول التفاصيل ═════════
        tree_frame = tk.Frame(win, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=15, pady=10)

        tree = ttk.Treeview(tree_frame,
                            columns=('name', 'amount'),
                            show='headings', height=15)
        tree.heading('name', text='التاجر')
        tree.heading('amount', text='المبلغ المتبقي علينا')
        tree.column('name', width=300, anchor='center')
        tree.column('amount', width=200, anchor='center')
        tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        for name, amount in details:
            tree.insert('', 'end', values=(name, f"{amount:,.2f}"))

        tree.tag_configure('row', background='#fadbd8')
        for item in tree.get_children():
            tree.item(item, tags=('row',))

        # ═════════ الأزرار ═════════
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="🖼️ طباعه",
                  command=lambda: self.export_summary_image(details, total_due),
                  bg='#8e44ad', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إغلاق", command=win.destroy,
                  bg='#e74c3c', fg='white', padx=20, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

    def export_summary_image(self, details, total_due):
        """تصدير تقرير المديونيات المتبقية إلى صورة PNG"""
        from tkinter import filedialog

        today = datetime.now().strftime('%Y-%m-%d')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"المديونيات_المتبقية_{today}.png",
            title="حفظ التقرير صورة"
        )
        if not file_path:
            return

        columns = ['اسم التاجر', 'المبلغ المتبقي علينا']
        rows_data = [(name, f"{amount:,.2f} ج.م") for name, amount in details]

        summary_items = [
            ("إجمالي المديونيات المتبقية", f"{total_due:,.2f} ج.م"),
            ("عدد التجار المدينين", str(len(details))),
        ]

        success = export_table_to_image(
            file_path=file_path,
            title="💰 المديونيات المتبقية علينا",
            subtitle=f"تاريخ التقرير: {today}",
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم حفظ الصورة:\n{file_path}\n\nفتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass


class TraderStatementWindow:
    """كشف حساب تفصيلي للتاجر"""
    def __init__(self, parent_page, trader_id):
        self.parent_page = parent_page
        self.trader_id = trader_id
        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("كشف حساب التاجر")
        self.window.geometry("1000x700")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        conn = parent_page.get_connection()
        c = conn.cursor()
        c.execute("SELECT name, phone FROM traders WHERE id=?", (trader_id,))
        row = c.fetchone()
        conn.close()
        trader_name = row[0] if row else 'غير معروف'
        phone = row[1] if row else ''
        # ✅ حفظ الاسم والهاتف في الكلاس للاستخدام في التصدير
        self.trader_name = trader_name
        self.trader_phone = phone
        self.trader_id = trader_id

        title_frame = tk.Frame(self.window, bg='#2c3e50', height=70)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text=f"📄 كشف حساب: {trader_name}",
                 font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white').pack(side='left', padx=20, pady=15)
        if phone:
            tk.Label(title_frame, text=f"📞 {phone}",
                     font=('Arial', 11), bg='#2c3e50', fg='#bdc3c7').pack(side='right', padx=20)

        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)

        columns = ('date', 'type', 'description', 'debit', 'credit', 'balance')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=18)
        self.tree.heading('date', text='التاريخ')
        self.tree.heading('type', text='النوع')
        self.tree.heading('description', text='البيان')
        self.tree.heading('debit', text='مدفوع')
        self.tree.heading('credit', text='مشتريات')
        self.tree.heading('balance', text='الرصيد')

        self.tree.column('date', width=100, anchor='center')
        self.tree.column('type', width=100, anchor='center')
        self.tree.column('description', width=350, anchor='w')
        self.tree.column('debit', width=110, anchor='center')
        self.tree.column('credit', width=110, anchor='center')
        self.tree.column('balance', width=120, anchor='center')

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # ملخص
        summary = tk.Frame(self.window, bg='#34495e', height=80)
        summary.pack(fill='x', padx=10, pady=5)
        summary.pack_propagate(False)

        self.lbl_total_inv = tk.Label(summary, text="إجمالي المشتريات: 0.00",
                                       font=('Arial', 12, 'bold'), bg='#34495e', fg='white')
        self.lbl_total_inv.pack(side='left', padx=20, pady=25)

        self.lbl_total_paid = tk.Label(summary, text="إجمالي المدفوع: 0.00",
                                        font=('Arial', 12, 'bold'), bg='#34495e', fg='#2ecc71')
        self.lbl_total_paid.pack(side='left', padx=20, pady=25)

        self.lbl_balance = tk.Label(summary, text="الرصيد الحالي: 0.00",
                                     font=('Arial', 14, 'bold'), bg='#34495e', fg='#f1c40f')
        self.lbl_balance.pack(side='right', padx=20, pady=25)

        btn_frame= tk.Frame(self.window, bg='white', height=80)
        btn_frame.pack()
        #btn_frame.pack_propagate(False)

        tk.Button(btn_frame, text="❌ إغلاق", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=15, pady=5).pack(side='left', padx=5)

        tk.Button(btn_frame, text="🖼️ طباعه",
                  command=self.export_image,
                  bg='#8e44ad', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)



        self.load_statement()

    def load_statement(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = self.parent_page.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM traders WHERE id=?", (self.trader_id,))
        trader_name = c.fetchone()[0]

        movements = []

        # 1. الفواتير (دائن - علينا)
        c.execute("""SELECT date, invoice_no, total,
                            COALESCE(payment_type, 'أجل'),
                            COALESCE(paid_amount, 0), id
                     FROM invoices WHERE supplier = ?""", (trader_name,))
        for inv in c.fetchall():
            date, inv_no, total, ptype, paid, inv_id = inv
            movements.append({
                'date': date, 'type': 'فاتورة',
                'description': f"فاتورة رقم {inv_no} ({ptype})",
                'debit': 0, 'credit': total,
                'sort': f"{date}_{inv_id}_1"
            })
            if paid > 0:
                movements.append({
                    'date': date, 'type': 'دفع',
                    'description': f"دفعة عند الفاتورة {inv_no}",
                    'debit': paid, 'credit': 0,
                    'sort': f"{date}_{inv_id}_2"
                })

        # 2. أي دفعات إضافية في paids (تحت اسم التاجر ولم تكن مرتبطة بفاتورة)
        # 2. أي دفعات إضافية في paids (تحت اسم التاجر ولم تكن مرتبطة بفاتورة)
        # ⚠️ نستثني 'الخزينة' — لأنها ستظهر من treasury_actions (منع الازدواج)
        c.execute("""SELECT date, value, description, id
        FROM paids
        WHERE sub_name = ?
          AND (invoice_id IS NULL OR invoice_id = 0)
          AND COALESCE(payment_method, 'نقدي') != 'الخزينة'""",
                  (trader_name,))

        for p in c.fetchall():
            date, value, desc, pid = p
            movements.append({
                'date': date, 'type': 'دفعة',
                'description': desc or 'دفعة على الحساب',
                'debit': value, 'credit': 0,
                'sort': f"{date}_{pid}_3"
            })

        # 3. ✅ سدادات الخزينة (سداد تاجر)
        c.execute("""SELECT date, amount, description, id
                     FROM treasury_actions
                     WHERE action_type = 'سداد تاجر' AND beneficiary = ?
                     ORDER BY date, id""", (trader_name,))
        for t in c.fetchall():
            date, value, desc, tid = t
            movements.append({
                'date': date, 'type': 'سداد خزينة',
                'description': desc or f'سداد من الخزينة',
                'debit': value, 'credit': 0,
                'sort': f"{date}_{tid}_4"
            })

        movements.sort(key=lambda x: x['sort'])

        total_purchases = 0
        total_payments = 0
        running = 0

        for m in movements:
            running += m['credit'] - m['debit']
            total_purchases += m['credit']
            total_payments += m['debit']
            self.tree.insert('', 'end', values=(
                m['date'], m['type'], m['description'],
                f"{m['debit']:.2f}" if m['debit'] else '',
                f"{m['credit']:.2f}" if m['credit'] else '',
                f"{running:.2f}"
            ))

        conn.close()
        self.lbl_total_inv.config(text=f"إجمالي المشتريات: {total_purchases:.2f}")
        self.lbl_total_paid.config(text=f"إجمالي المدفوع: {total_payments:.2f}")
        self.lbl_balance.config(text=f"الرصيد الحالي: {running:.2f}")


    def export_image(self):
        """تصدير كشف حساب التاجر إلى صورة"""
        from tkinter import filedialog

        safe_name = self.trader_name.replace(' ', '_').replace('/', '_')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"كشف_حساب_تاجر_{safe_name}.png",
            title="حفظ كشف الحساب صورة"
        )
        if not file_path:
            return

        columns = ['التاريخ', 'النوع', 'البيان', 'مدفوع', 'مشتريات', 'الرصيد']
        rows_data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            rows_data.append(values)

        success = export_table_to_image(
            file_path=file_path,
            title=f"كشف حساب التاجر: {self.trader_name}",
            subtitle=f"📞 {self.trader_phone}" if self.trader_phone else "",
            columns=columns,
            rows=rows_data,
            summary_items=[
                ("إجمالي المشتريات", self.lbl_total_inv.cget('text').replace('إجمالي المشتريات:', '').strip()),
                ("إجمالي المدفوع", self.lbl_total_paid.cget('text').replace('إجمالي المدفوع:', '').strip()),
                ("الرصيد الحالي", self.lbl_balance.cget('text').replace('الرصيد الحالي:', '').strip()),
            ],
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح", f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass




class DeferredInvoicesWindow:
    """عرض وإدارة الفواتير الآجلة"""
    def __init__(self, parent_page):
        self.parent_page = parent_page
        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("📋 الفواتير الآجلة")
        self.window.geometry("1100x700")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # العنوان
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=50)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text="📋 الفواتير الآجلة (التي عليها متبقي)",
                 font=('Arial', 16, 'bold'),
                 bg='#2c3e50', fg='white').pack(pady=12)

        # الفلاتر
        filters = tk.LabelFrame(self.window, text="فلاتر",
                                 font=('Arial', 12, 'bold'),
                                 bg='#ecf0f1', padx=10, pady=10)
        filters.pack(fill='x', padx=10, pady=5)

        tk.Label(filters, text="من تاريخ:", font=('Arial', 11),
                 bg='#ecf0f1').grid(row=0, column=0, padx=5, pady=5)
        self.date_from = tk.Entry(filters, width=15, font=('Arial', 11))
        self.date_from.grid(row=0, column=1, padx=5, pady=5)
        self.date_from.insert(0, datetime.now().strftime('%Y-%m-01'))

        tk.Label(filters, text="إلى تاريخ:", font=('Arial', 11),
                 bg='#ecf0f1').grid(row=0, column=2, padx=5, pady=5)
        self.date_to = tk.Entry(filters, width=15, font=('Arial', 11))
        self.date_to.grid(row=0, column=3, padx=5, pady=5)
        self.date_to.insert(0, datetime.now().strftime('%Y-%m-%d'))

        tk.Label(filters, text="التاجر:", font=('Arial', 11),
                 bg='#ecf0f1').grid(row=0, column=4, padx=5, pady=5)
        self.trader_filter = AutocompleteCombobox(filters, width=20,
                                                    font=('Arial', 11))
        self.trader_filter.grid(row=0, column=5, padx=5, pady=5)

        tk.Button(filters, text="🔍 بحث", command=self.load_data,
                  bg='#2980b9', fg='white', padx=15, pady=5).grid(row=0, column=6, padx=10)
        tk.Button(filters, text="🔄 عرض الكل", command=self.load_all,
                  bg='#7f8c8d', fg='white', padx=15, pady=5).grid(row=0, column=7, padx=5)

        # الجدول
        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ('id', 'invoice_no', 'date', 'supplier', 'total',
                   'paid', 'remaining', 'payment_type')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                   show='headings', height=15)

        headers = {'id': 'ID', 'invoice_no': 'رقم الفاتورة', 'date': 'التاريخ',
                   'supplier': 'التاجر', 'total': 'الإجمالي',
                   'paid': 'المدفوع', 'remaining': 'المتبقي',
                   'payment_type': 'نوع الدفع'}
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=40, anchor='center')
        self.tree.column('invoice_no', width=100, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('supplier', width=180, anchor='center')
        self.tree.column('total', width=100, anchor='center')
        self.tree.column('paid', width=100, anchor='center')
        self.tree.column('remaining', width=100, anchor='center')
        self.tree.column('payment_type', width=100, anchor='center')

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self.tree.bind('<Double-Button-1>', lambda e: self.edit_invoice())
        self.tree.bind('<Delete>', lambda e: self.delete_invoice())
        self.tree.bind('<F2>', lambda e: self.edit_invoice())

        # الملخص
        summary = tk.Frame(self.window, bg='#34495e', height=60)
        summary.pack(fill='x', padx=10, pady=5)
        summary.pack_propagate(False)

        self.lbl_count = tk.Label(summary, text="عدد الفواتير: 0",
                                    font=('Arial', 12, 'bold'),
                                    bg='#34495e', fg='white')
        self.lbl_count.pack(side='left', padx=20, pady=15)

        self.lbl_total = tk.Label(summary, text="إجمالي المتبقي: 0.00",
                                    font=('Arial', 12, 'bold'),
                                    bg='#34495e', fg='#e74c3c')
        self.lbl_total.pack(side='right', padx=20, pady=15)

        # الأزرار
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=8)

        tk.Button(btn_frame, text="✏️ تعديل الفاتورة (F2)",
                  command=self.edit_invoice,
                  bg='#f39c12', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="🗑️ حذف الفاتورة (Delete)",
                  command=self.delete_invoice,
                  bg='#e74c3c', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="💰 سداد المتبقي",
                  command=self.settle_remaining,
                  bg='#16a085', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إغلاق",
                  command=self.window.destroy,
                  bg='#95a5a6', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        # شريط معلومات
        tk.Label(self.window,
                 text="💡 نصيحة: دوبل كليك = تعديل  |  Delete = حذف  |  F2 = تعديل",
                 font=('Arial', 10, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=5)

        self.load_all()

    # ────────────────────────────────────────────────
    def load_all(self):
        self.date_from.delete(0, tk.END)
        self.date_to.delete(0, tk.END)
        self.trader_filter.set('')
        self.load_data()

    def load_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()
        trader = self.trader_filter.get().strip()

        conn = self.parent_page.get_connection()
        c = conn.cursor()

        query = """SELECT id, invoice_no, date, supplier, total,
                          COALESCE(paid_amount, 0),
                          (total - COALESCE(paid_amount, 0)),
                          COALESCE(payment_type, 'أجل')
                   FROM invoices
                   WHERE (total - COALESCE(paid_amount, 0)) > 0.01"""
        params = []

        if date_from:
            query += " AND date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND date <= ?"
            params.append(date_to)
        if trader:
            query += " AND supplier LIKE ?"
            params.append(f'%{trader}%')

        query += " ORDER BY date DESC, id DESC"
        c.execute(query, params)
        rows = c.fetchall()

        # قائمة التجار
        c.execute("SELECT name FROM traders ORDER BY name")
        names = [r[0] for r in c.fetchall()]
        conn.close()

        self.trader_filter.set_values(names)

        total_remaining = 0
        for row in rows:
            rid, inv_no, date, supplier, total, paid, remaining, ptype = row
            self.tree.insert('', 'end', iid=str(rid), values=(
                rid, inv_no, date, supplier,
                f"{total:.2f}", f"{paid:.2f}", f"{remaining:.2f}", ptype))
            total_remaining += remaining

        self.lbl_count.config(text=f"عدد الفواتير: {len(rows)}")
        self.lbl_total.config(text=f"إجمالي المتبقي: {total_remaining:.2f}")

    # ────────────────────────────────────────────────
    def edit_invoice(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("تنبيه", "اختر فاتورة أولاً")
            return
        invoice_id = int(selected[0])
        InvoiceEditWindow(
            self.parent_page,
            invoice_id,
            on_save_callback=self.load_data
        )

    def delete_invoice(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("تنبيه", "اختر فاتورة أولاً")
            return

        invoice_id = int(selected[0])
        values = self.tree.item(selected[0])['values']
        inv_no = values[1]
        date = values[2]
        supplier = values[3]
        total = values[4]
        paid = values[5]
        remaining = values[6]

        msg = (
            f"⚠️ هل أنت متأكد من حذف الفاتورة؟\n\n"
            f"📄 رقم الفاتورة: {inv_no}\n"
            f"📅 التاريخ: {date}\n"
            f"🧑‍💼 التاجر: {supplier}\n"
            f"💰 الإجمالي: {total}\n"
            f"✅ المدفوع: {paid}\n"
            f"⏳ المتبقي: {remaining}\n\n"
            f"🔴 سيتم حذف الفاتورة وكل بنودها والسجلات المرتبطة."
        )

        if not messagebox.askyesno("تأكيد الحذف", msg):
            return

        conn = self.parent_page.get_connection()
        c = conn.cursor()

        try:
            c.execute("DELETE FROM invoice_items WHERE invoice_id=?", (invoice_id,))
            c.execute("DELETE FROM paids WHERE invoice_id=?", (invoice_id,))
            c.execute("DELETE FROM invoices WHERE id=?", (invoice_id,))

            try:
                c.execute("""DELETE FROM treasury_actions 
                             WHERE paid_id IN (SELECT id FROM paids WHERE invoice_id=?)""",
                          (invoice_id,))
            except:
                pass

            conn.commit()
            messagebox.showinfo("نجاح", "✅ تم حذف الفاتورة بنجاح")
            self.load_data()

        except Exception as e:
            conn.rollback()
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()

    def settle_remaining(self):
        # (نفس الكود السابق)
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("تنبيه", "اختر فاتورة أولاً")
            return

        invoice_id = int(selected[0])
        values = self.tree.item(selected[0])['values']
        remaining = float(values[6])
        supplier = values[3]
        invoice_no = values[1]
        date = values[2]

        if not messagebox.askyesno("تأكيد",
                f"سيتم تقييد المبلغ المتبقي ({remaining:.2f}) في صفحة المصاريف\n"
                f"كدفعة للتاجر '{supplier}' عن الفاتورة رقم {invoice_no}.\n\n"
                f"هل تريد المتابعة؟"):
            return

        conn = self.parent_page.get_connection()
        c = conn.cursor()
        try:
            c.execute("SELECT id FROM main_categories WHERE name='مشتريات'")
            main_row = c.fetchone()
            if not main_row:
                messagebox.showerror("خطأ", "التصنيف 'مشتريات' غير موجود")
                conn.close()
                return
            main_id = main_row[0]

            c.execute("SELECT id FROM sub_categories WHERE main_id=? AND name=?",
                      (main_id, supplier))
            if not c.fetchone():
                c.execute("INSERT INTO sub_categories (main_id, name) VALUES (?, ?)",
                          (main_id, supplier))

            c.execute('''INSERT INTO paids
                         (date, value, main_name, sub_name, description, invoice_id)
                         VALUES (?, ?, ?, ?, ?, ?)''',
                      (date, remaining, 'مشتريات', supplier,
                       f"سداد متبقي الفاتورة {invoice_no}", invoice_id))

            c.execute("UPDATE invoices SET paid_amount = total WHERE id=?", (invoice_id,))

            conn.commit()
            messagebox.showinfo("نجاح",
                f"✅ تم تقييد {remaining:.2f} في صفحة المصاريف")
            self.load_data()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل التقييد: {e}")
        finally:
            conn.close()

class InvoicesPage(tk.Frame):
    """صفحة عرض وإدارة الفواتير"""
    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None
        # 1️⃣ الشريط العلوي
        add_program_banner(self)

        # 2️⃣ منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        tk.Label(content, text="📄 إدارة الفواتير",
                 font=('Arial', 16, 'bold'), bg='#ecf0f1').pack(pady=10)

        # ============ فلاتر البحث ============
        filters = tk.LabelFrame(content, text="فلاتر البحث",
                                 font=('Arial', 12, 'bold'), bg='#f0f0f0',
                                 padx=10, pady=10)
        filters.pack(fill='x', padx=10, pady=5)

        tk.Label(filters, text="من تاريخ:", font=('Arial', 11), bg='#f0f0f0').grid(row=0, column=0, padx=5, pady=5)
        self.date_from = tk.Entry(filters, width=15, font=('Arial', 11))
        self.date_from.grid(row=0, column=1, padx=5, pady=5)
        self.date_from.insert(0, datetime.now().strftime('%Y-%m-01'))

        tk.Label(filters, text="إلى تاريخ:", font=('Arial', 11), bg='#f0f0f0').grid(row=0, column=2, padx=5, pady=5)
        self.date_to = tk.Entry(filters, width=15, font=('Arial', 11))
        self.date_to.grid(row=0, column=3, padx=5, pady=5)
        self.date_to.insert(0, datetime.now().strftime('%Y-%m-%d'))

        tk.Label(filters, text="التاجر:", font=('Arial', 11), bg='#f0f0f0').grid(row=0, column=4, padx=5, pady=5)
        self.trader_filter = AutocompleteCombobox(filters, width=20, font=('Arial', 11))
        self.trader_filter.grid(row=0, column=5, padx=5, pady=5)

        tk.Label(filters, text="نوع الدفع:", font=('Arial', 11), bg='#f0f0f0').grid(row=0, column=6, padx=5, pady=5)
        self.ptype_filter = ttk.Combobox(filters, values=['الكل', 'كاش', 'أجل', 'جزئي'],
                                          width=12, font=('Arial', 11), state='readonly')
        self.ptype_filter.set('الكل')
        self.ptype_filter.grid(row=0, column=7, padx=5, pady=5)

        tk.Button(filters, text="🔍 بحث", command=self.load_invoices,
                  bg='#2980b9', fg='white', padx=15, pady=5).grid(row=0, column=8, padx=10)
        tk.Button(filters, text="🔄 عرض الكل", command=self.show_all,
                  bg='#7f8c8d', fg='white', padx=15, pady=5).grid(row=0, column=9, padx=5)

        # ============ أزرار العمليات ============
        btn_frame = tk.Frame(content, bg='#ecf0f1')
        btn_frame.pack(pady=8)

        self.btn_edit = tk.Button(btn_frame, text="✏️ تعديل الفاتورة",
                                    command=self.edit_invoice,
                                    bg='#f39c12', fg='white', padx=15, pady=8,
                                    font=('Arial', 11, 'bold'), state='disabled')
        self.btn_edit.grid(row=0, column=0, padx=5)

        self.btn_delete = tk.Button(btn_frame, text="🗑️ حذف الفاتورة",
                                      command=self.delete_invoice,
                                      bg='#e74c3c', fg='white', padx=15, pady=8,
                                      font=('Arial', 11, 'bold'), state='disabled')
        self.btn_delete.grid(row=0, column=1, padx=5)

        self.btn_deferred = tk.Button(btn_frame, text="📋 الآجلة فقط",
                                        command=self.show_deferred,
                                        bg='#c0392b', fg='white', padx=15, pady=8,
                                        font=('Arial', 11, 'bold'))
        self.btn_deferred.grid(row=0, column=2, padx=5)
        # ✅ الأزرار المنقولة من صفحة المصاريف
        tk.Button(btn_frame, text="⏳ الفواتير الآجلة",
                  command=self.open_deferred_invoices,
                  bg='#c0392b', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).grid(row=0, column=3, padx=5)

        tk.Button(btn_frame, text="📊 تقرير أصناف",
                  command=self.open_items_report,
                  bg='#16a085', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).grid(row=0, column=4, padx=5)

        # ============ جدول الفواتير ============
        tree_frame = tk.Frame(content, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ('id', 'invoice_no', 'date', 'supplier', 'total',
                   'paid', 'remaining', 'payment_type', 'items_count')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)

        headers = {
            'id': 'ID', 'invoice_no': 'رقم الفاتورة', 'date': 'التاريخ',
            'supplier': 'التاجر', 'total': 'الإجمالي', 'paid': 'المدفوع',
            'remaining': 'المتبقي', 'payment_type': 'نوع الدفع',
            'items_count': 'عدد الأصناف'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=40, anchor='center')
        self.tree.column('invoice_no', width=100, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('supplier', width=180, anchor='center')
        self.tree.column('total', width=100, anchor='center')
        self.tree.column('paid', width=100, anchor='center')
        self.tree.column('remaining', width=100, anchor='center')
        self.tree.column('payment_type', width=100, anchor='center')
        self.tree.column('items_count', width=100, anchor='center')

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # تلوين الصفوف حسب الحالة
        self.tree.tag_configure('paid_full', background='#d4edda')     # مسدد
        self.tree.tag_configure('deferred', background='#ffe6e6')      # عليه متبقي
        self.tree.tag_configure('partial', background='#fff3cd')       # جزئي

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_invoice())

        # ============ الإجماليات ============
        summary = tk.Frame(content, bg='#34495e', height=60)
        summary.pack(fill='x', padx=10, pady=5)
        summary.pack_propagate(False)

        self.lbl_count = tk.Label(summary, text="عدد الفواتير: 0",
                                    font=('Arial', 12, 'bold'), bg='#34495e', fg='white')
        self.lbl_count.pack(side='left', padx=20, pady=15)

        self.lbl_total = tk.Label(summary, text="إجمالي الفواتير: 0.00",
                                    font=('Arial', 12, 'bold'), bg='#34495e', fg='#f1c40f')
        self.lbl_total.pack(side='left', padx=20, pady=15)

        self.lbl_remaining = tk.Label(summary, text="إجمالي المتبقي: 0.00",
                                        font=('Arial', 12, 'bold'), bg='#34495e', fg='#e74c3c')
        self.lbl_remaining.pack(side='right', padx=20, pady=15)

        self.load_invoices()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def show_all(self):
        self.date_from.delete(0, tk.END)
        self.date_to.delete(0, tk.END)
        self.trader_filter.set('')
        self.ptype_filter.set('الكل')
        self.load_invoices()

    def show_deferred(self):
        """عرض الفواتير التي عليها متبقي فقط"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT id, invoice_no, date, supplier, total,
                            COALESCE(paid_amount, 0),
                            (total - COALESCE(paid_amount, 0)),
                            COALESCE(payment_type, 'أجل')
                     FROM invoices
                     WHERE (total - COALESCE(paid_amount, 0)) > 0.01
                     ORDER BY date DESC, id DESC""")
        rows = c.fetchall()
        conn.close()

        self._populate_tree(rows)

    def load_invoices(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()
        trader = self.trader_filter.get().strip()
        ptype = self.ptype_filter.get().strip()

        conn = self.get_connection()
        c = conn.cursor()

        query = """SELECT id, invoice_no, date, supplier, total,
                          COALESCE(paid_amount, 0),
                          (total - COALESCE(paid_amount, 0)),
                          COALESCE(payment_type, 'أجل')
                   FROM invoices WHERE 1=1"""
        params = []

        if date_from:
            query += " AND date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND date <= ?"
            params.append(date_to)
        if trader:
            query += " AND supplier LIKE ?"
            params.append(f'%{trader}%')
        if ptype and ptype != 'الكل':
            query += " AND payment_type = ?"
            params.append(ptype)

        query += " ORDER BY date DESC, id DESC"
        c.execute(query, params)
        rows = c.fetchall()

        # جلب أسماء التجار للفلتر
        c.execute("SELECT name FROM traders ORDER BY name")
        names = [r[0] for r in c.fetchall()]
        conn.close()

        self.trader_filter.set_values(names)
        self._populate_tree(rows)

    def _populate_tree(self, rows):
        total_sum = 0
        remaining_sum = 0

        for row in rows:
            rid, inv_no, date, supplier, total, paid, remaining, ptype = row

            # جلب عدد الأصناف
            conn = self.get_connection()
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM invoice_items WHERE invoice_id=?", (rid,))
            items_count = c.fetchone()[0]
            conn.close()

            # تحديد لون الصف
            if remaining <= 0.01:
                tag = 'paid_full'
            elif paid > 0:
                tag = 'partial'
            else:
                tag = 'deferred'

            self.tree.insert('', 'end', iid=str(rid), tags=(tag,), values=(
                rid, inv_no or '', date or '', supplier or '',
                f"{total:.2f}", f"{paid:.2f}", f"{remaining:.2f}",
                ptype, items_count))

            total_sum += total
            remaining_sum += remaining

        self.lbl_count.config(text=f"عدد الفواتير: {len(rows)}")
        self.lbl_total.config(text=f"إجمالي الفواتير: {total_sum:.2f}")
        self.lbl_remaining.config(text=f"إجمالي المتبقي: {remaining_sum:.2f}")

    def on_select(self, event):
        selected = self.tree.selection()
        if selected:
            self.selected_id = int(selected[0])
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')

    def edit_invoice(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر فاتورة أولاً")
            return
        InvoiceEditWindow(self, self.selected_id, on_save_callback=self.load_invoices)

    def delete_invoice(self):
        if self.selected_id is None:
            return

        # جلب معلومات الفاتورة
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT invoice_no, supplier, total, COALESCE(paid_amount, 0) FROM invoices WHERE id=?",
                  (self.selected_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            return

        inv_no, supplier, total, paid = row

        if not messagebox.askyesno("تأكيد",
                f"هل أنت متأكد من حذف الفاتورة رقم {inv_no}؟\n\n"
                f"التاجر: {supplier}\n"
                f"الإجمالي: {total:.2f}\n"
                f"المدفوع: {paid:.2f}\n\n"
                f"⚠️ سيتم حذف جميع بنود الفاتورة أيضاً."):
            return

        conn = self.get_connection()
        c = conn.cursor()
        try:
            # حذف البنود أولاً
            c.execute("DELETE FROM invoice_items WHERE invoice_id=?", (self.selected_id,))
            # حذف السجلات المرتبطة في paids
            c.execute("DELETE FROM paids WHERE invoice_id=?", (self.selected_id,))
            # حذف الفاتورة
            c.execute("DELETE FROM invoices WHERE id=?", (self.selected_id,))
            conn.commit()
            messagebox.showinfo("نجاح", "تم حذف الفاتورة")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.load_invoices()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()

    def open_deferred_invoices(self):
        """فتح نافذة الفواتير الآجلة"""
        DeferredInvoicesWindow(self)

    def open_items_report(self):
        """فتح نافذة تقرير الأصناف"""
        ItemsReportWindow(self)


class EmployeeFormWindow:
    """نافذة إدخال/تعديل بيانات عامل"""
    def __init__(self, parent_page, employee_id=None):
        self.parent_page = parent_page
        self.employee_id = employee_id  # None = إضافة، غير None = تعديل
        self.is_edit = employee_id is not None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        title = "تعديل بيانات عامل" if self.is_edit else "إضافة عامل جديد"
        self.window.title(title)
        self.window.geometry("550x480")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ============ العنوان ============
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏️" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white').pack(pady=15)

        # ============ الإطار الرئيسي ============
        main_frame = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main_frame.pack(fill='both', expand=True)

        # الاسم
        tk.Label(main_frame, text="اسم العامل: *", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', anchor='e').grid(row=0, column=0, padx=5, pady=8, sticky='e')
        self.entry_name = tk.Entry(main_frame, width=28, font=('Arial', 11))
        self.entry_name.grid(row=0, column=1, padx=5, pady=8, sticky='w')

        # الهاتف
        tk.Label(main_frame, text="الهاتف:", font=('Arial', 11),
                 bg='#ecf0f1', anchor='e').grid(row=1, column=0, padx=5, pady=8, sticky='e')
        self.entry_phone = tk.Entry(main_frame, width=28, font=('Arial', 11))
        self.entry_phone.grid(row=1, column=1, padx=5, pady=8, sticky='w')

        # الوظيفة
        tk.Label(main_frame, text="الوظيفة:", font=('Arial', 11),
                 bg='#ecf0f1', anchor='e').grid(row=2, column=0, padx=5, pady=8, sticky='e')
        self.entry_position = tk.Entry(main_frame, width=28, font=('Arial', 11))
        self.entry_position.grid(row=2, column=1, padx=5, pady=8, sticky='w')

        # اليومية
        tk.Label(main_frame, text="اليومية (ج.م):", font=('Arial', 11),
                 bg='#ecf0f1', anchor='e').grid(row=3, column=0, padx=5, pady=8, sticky='e')
        self.entry_daily_wage = tk.Entry(main_frame, width=28, font=('Arial', 11))
        self.entry_daily_wage.grid(row=3, column=1, padx=5, pady=8, sticky='w')



        # ملاحظات
        tk.Label(main_frame, text="ملاحظات:", font=('Arial', 11),
                 bg='#ecf0f1', anchor='e').grid(row=5, column=0, padx=5, pady=8, sticky='ne')
        self.entry_notes = tk.Text(main_frame, width=28, height=4, font=('Arial', 11))
        self.entry_notes.grid(row=5, column=1, padx=5, pady=8, sticky='w')

        # ============ أزرار الحفظ والإلغاء ============
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        # إذا كان تعديلاً، حمّل البيانات
        if self.is_edit:
            self.load_employee_data()

        # تركيز على الاسم
        self.entry_name.focus()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def load_employee_data(self):
        """جلب بيانات العامل لتعبئة الحقول"""
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT name, phone, position, daily_wage,  notes
                     FROM employees WHERE id=?""", (self.employee_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "العامل غير موجود")
            self.window.destroy()
            return

        name, phone, position, daily_wage, notes = row

        self.entry_name.insert(0, name or '')
        self.entry_phone.insert(0, phone or '')
        self.entry_position.insert(0, position or '')
        self.entry_daily_wage.insert(0, str(daily_wage or 0))

        self.entry_notes.insert('1.0', notes or '')

    def save(self):
        """حفظ بيانات العامل (الحالة تُحسب تلقائياً)"""
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("تنبيه", "اسم العامل مطلوب")
            self.entry_name.focus()
            return

        try:
            daily_wage = float(self.entry_daily_wage.get().strip() or 0)
        except ValueError:
            messagebox.showerror("خطأ", "اليومية يجب أن تكون رقماً")
            self.entry_daily_wage.focus()
            return

        phone = self.entry_phone.get().strip()
        position = self.entry_position.get().strip()
        notes = self.entry_notes.get('1.0', 'end').strip()

        conn = self.get_connection()
        c = conn.cursor()

        try:
            if self.is_edit:
                c.execute("SELECT name FROM employees WHERE id=?", (self.employee_id,))
                old_row = c.fetchone()
                old_name = old_row[0] if old_row else None

                c.execute("""UPDATE employees 
                             SET name=?, phone=?, position=?, daily_wage=?, notes=?
                             WHERE id=?""",
                          (name, phone, position, daily_wage, notes, self.employee_id))

                if old_name and old_name != name:
                    c.execute("UPDATE sub_categories SET name=? WHERE name=?",
                              (name, old_name))
                    c.execute("UPDATE paids SET sub_name=? WHERE sub_name=?",
                              (name, old_name))

                msg = f"✅ تم تعديل بيانات '{name}'"
            else:
                # ✅ عند الإضافة، الحالة الافتراضية 'لم يبدأ' (ستُحدَّث عند إضافة أول مدة)
                c.execute("""INSERT INTO employees 
                             (name, phone, position, daily_wage, status, notes, created_at)
                             VALUES (?, ?, ?, ?, ?, ?, ?)""",
                          (name, phone, position, daily_wage, 'لم يبدأ', notes,
                           datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

                c.execute("SELECT id FROM main_categories WHERE name='أجور ومرتبات'")
                main_row = c.fetchone()
                if main_row:
                    main_id = main_row[0]
                    c.execute("SELECT id FROM sub_categories WHERE main_id=? AND name=?",
                              (main_id, name))
                    if not c.fetchone():
                        c.execute("INSERT INTO sub_categories (main_id, name) VALUES (?, ?)",
                                  (main_id, name))

                msg = f"✅ تمت إضافة العامل '{name}'"

            conn.commit()
            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_page.load_employees()

        except sqlite3.IntegrityError:
            messagebox.showerror("خطأ", "هذا الاسم موجود مسبقاً")
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}")
        finally:
            conn.close()


class EmployeesPage(tk.Frame):
    """صفحة إدارة العاملين (أزرار + جدول فقط)"""
    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        self.selected_id = None
        # 1️⃣ الشريط العلوي
        add_program_banner(self)

        # 2️⃣ منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ============ العنوان ============
        tk.Label(content, text="👷 إدارة العاملين",
                 font=('Arial', 18, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=12)

        # ============ شريط الأزرار ============
        btn_frame = tk.LabelFrame(content,
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(0, 10))

        # الصف الأول من الأزرار
        row1 = tk.Frame(btn_frame, bg='#ecf0f1')
        row1.pack(pady=3)

        tk.Button(row1, text="➕ إضافة عامل", command=self.add_employee,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(row1, text="✏️ تعديل", command=self.edit_employee,
                                    bg='#f39c12', fg='white', font=('Arial', 11, 'bold'),
                                    padx=20, pady=8, width=15, state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        tk.Button(row1, text="🗑️ حذف", command=self.delete_employee,
                  bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        # الصف الثاني من الأزرار
        row2 = tk.Frame(btn_frame, bg='#ecf0f1')
        row2.pack(pady=3)



        tk.Button(row1, text="📄 كشف الحساب", command=self.open_statement,
                  bg='#2980b9', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        tk.Button(row1, text="🔄 تحديث", command=self.load_employees,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)
        tk.Button(row1, text="📊 تقرير النشطين",
                  command=self.show_active_report,
                  bg='#c0392b', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=18).pack(side='left', padx=5)

        # ============ شريط البحث ============
        search_frame = tk.Frame(content, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=15, pady=5)

        tk.Label(search_frame, text="🔎 بحث:", font=('Arial', 11, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=5)

        self.search_entry = tk.Entry(search_frame, width=30, font=('Arial', 11))
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind('<KeyRelease>', lambda e: self.load_employees(self.search_entry.get().strip()))

        tk.Button(search_frame, text="🔄 مسح", command=self.clear_search,
                  bg='#bdc3c7', fg='#2c3e50', font=('Arial', 10, 'bold'),
                  padx=10).pack(side='left', padx=5)

        # ============ جدول العرض ============
        tree_frame = tk.LabelFrame(content, text="قائمة العاملين",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 15))

        columns = ('id', 'name', 'phone', 'position', 'daily_wage',
                   'status', 'periods_count')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)

        headers = {
            'id': 'ID',
            'name': 'اسم العامل',
            'phone': 'الهاتف',
            'position': 'الوظيفة',
            'daily_wage': 'اليومية (ج.م)',
            'status': 'الحالة',
            'periods_count': 'عدد مدد العمل'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=50, anchor='center')
        self.tree.column('name', width=220, anchor='center')
        self.tree.column('phone', width=130, anchor='center')
        self.tree.column('position', width=170, anchor='center')
        self.tree.column('daily_wage', width=120, anchor='center')
        self.tree.column('status', width=110, anchor='center')
        self.tree.column('periods_count', width=140, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        # تلوين الصفوف
        self.tree.tag_configure('active', background='#d4edda')
        self.tree.tag_configure('inactive', background='#f8d7da')
        self.tree.tag_configure('not_started', background='#fff3cd')

        # ربط الأحداث
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_employee())

        # شريط الحالة (عدد العاملين)
        self.status_bar = tk.Label(content, text="جاهز",
                                    font=('Arial', 10),
                                    bg='#34495e', fg='white',
                                    anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # التحميل الأولي
        self.load_employees()

    # ─────────────────────────────────────────────────────────
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def get_employee_status(self, conn, employee_id):
        """
        حساب حالة العامل تلقائياً من مدد العمل:
        - نشط: يوجد مدة عمل حالية (آخر يوم >= اليوم) أو مدة بلا تاريخ نهاية
        - غير نشط: كل المدد انتهت قبل اليوم
        - لم يبدأ: لا توجد مدد
        """
        c = conn.cursor()
        today = datetime.now().strftime('%Y-%m-%d')

        # هل يوجد مدة نشطة (آخر يوم >= اليوم)؟
        c.execute("""SELECT COUNT(*) FROM employee_periods 
                     WHERE employee_id=? AND (last_date >= ? OR last_date IS NULL OR last_date = '')""",
                  (employee_id, today))
        active = c.fetchone()[0]

        if active > 0:
            return 'نشط'

        # هل يوجد أي مدة؟
        c.execute("SELECT COUNT(*) FROM employee_periods WHERE employee_id=?",
                  (employee_id,))
        total = c.fetchone()[0]

        if total == 0:
            return 'لم يبدأ'

        return 'غير نشط'

    def show_active_report(self):
        ActiveEmployeesReportWindow(self)

    def load_employees(self, search=''):
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = self.get_connection()
        c = conn.cursor()

        if search:
            c.execute("""SELECT id, name, phone, position, daily_wage
                         FROM employees WHERE name LIKE ? ORDER BY name""",
                      (f'%{search}%',))
        else:
            c.execute("""SELECT id, name, phone, position, daily_wage
                         FROM employees ORDER BY name""")
        rows = c.fetchall()

        for emp_id, name, phone, position, daily_wage in rows:
            # ✅ حساب الحالة تلقائياً
            status = self.get_employee_status(conn, emp_id)

            # عدد مدد العمل
            c.execute("SELECT COUNT(*) FROM employee_periods WHERE employee_id=?", (emp_id,))
            periods_count = c.fetchone()[0]

            # ✅ تحديث الحالة في قاعدة البيانات (للتزامن مع أي كود آخر)
            c.execute("UPDATE employees SET status=? WHERE id=?", (status, emp_id))

            # التلوين حسب الحالة
            if status == 'نشط':
                tag = 'active'
            elif status == 'غير نشط':
                tag = 'inactive'
            else:
                tag = 'not_started'

            self.tree.insert('', 'end', iid=emp_id, tags=(tag,), values=(
                emp_id, name, phone or '', position or '',
                f"{daily_wage or 0:.2f}", status, periods_count))

        conn.commit()
        conn.close()
        # ✅ العدد الكلي
        total = len(rows)

        # ✅ عدد النشطين (من قاعدة البيانات)
        conn2 = self.get_connection()
        c2 = conn2.cursor()
        c2.execute("SELECT COUNT(*) FROM employees WHERE status='نشط'")
        active_count = c2.fetchone()[0]
        conn2.close()

        self.status_bar.config(
            text=f"👥 العاملون النشطون: {active_count} من {total}")

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.load_employees()

    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            self.selected_id = int(sel[0])
            self.btn_edit.config(state='normal')
        else:
            self.selected_id = None
            self.btn_edit.config(state='disabled')

    # ─────────────────────────────────────────────────────────
    def add_employee(self):
        EmployeeFormWindow(self, employee_id=None)

    def edit_employee(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر عاملاً من الجدول أولاً")
            return
        EmployeeFormWindow(self, employee_id=self.selected_id)

    def delete_employee(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر عاملاً من الجدول أولاً")
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM employees WHERE id=?", (self.selected_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return
        emp_name = row[0]
        conn.close()

        if not messagebox.askyesno("تأكيد",
                f"هل أنت متأكد من حذف العامل '{emp_name}'؟\n\n"
                f"⚠️ سيتم حذف جميع مدد عمله أيضاً."):
            return

        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM employee_periods WHERE employee_id=?",
                      (self.selected_id,))
            c.execute("DELETE FROM employees WHERE id=?", (self.selected_id,))
            conn.commit()
            messagebox.showinfo("نجاح", f"تم حذف '{emp_name}'")
            self.selected_id = None
            self.btn_edit.config(state='disabled')
            self.load_employees()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()



    def open_statement(self):
        if self.selected_id is None:
            messagebox.showwarning("تنبيه", "اختر عاملاً من الجدول أولاً")
            return
        # فتح نافذة مدد العمل (منها يمكن فتح كشف الحساب)
        EmployeePeriodsWindow(self, self.selected_id)

class ActiveEmployeesReportWindow:
    """تقرير العاملين النشطين مع إجمالي السلف والمستحقات"""
    def __init__(self, parent_page):
        self.parent_page = parent_page

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("📊 تقرير العاملين النشطين")
        self.window.geometry("1400x750")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # ============ العنوان ============
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=70)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="📊 تقرير العاملين النشطين",
                 font=('Arial', 17, 'bold'), bg='#2c3e50', fg='white').pack(side='left', padx=20, pady=15)

        today = datetime.now().strftime('%Y-%m-%d')
        tk.Label(title_frame, text=f"📅 تاريخ التقرير: {today}",
                 font=('Arial', 12), bg='#2c3e50', fg='#f1c40f').pack(side='right', padx=20)

        # ============ جدول التقرير ============
        tree_frame = tk.Frame(self.window, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)

        columns = ('id', 'name', 'position', 'first_date', 'last_date',
                   'work_days', 'daily_wage', 'base_due',
                   'bonus', 'deductions', 'advances', 'net_due')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=18)

        headers = {
            'id': 'ID',
            'name': 'اسم العامل',
            'position': 'الوظيفة',
            'first_date': 'أول يوم',
            'last_date': 'آخر يوم',
            'work_days': 'أيام العمل',
            'daily_wage': 'اليومية',
            'base_due': 'المستحق الأساسي',
            'bonus': 'المكافآت',
            'deductions': 'الخصومات',
            'advances': 'السلف',
            'net_due': 'الصافي المستحق'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        widths = {
            'id': 40, 'name': 150, 'position': 100,
            'first_date': 100, 'last_date': 100, 'work_days': 80,
            'daily_wage': 80, 'base_due': 110, 'bonus': 80,
            'deductions': 80, 'advances': 110, 'net_due': 120
        }
        for col, w in widths.items():
            self.tree.column(col, width=w, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        # تلوين صف الإجمالي
        self.tree.tag_configure('total_row',
                                 background='#34495e',
                                 foreground='#f1c40f',
                                 font=('Arial', 11, 'bold'))

        # ============ الإجماليات في الأسفل ============
        summary_frame = tk.Frame(self.window, bg='#34495e', height=80)
        summary_frame.pack(fill='x', padx=10, pady=5)
        summary_frame.pack_propagate(False)

        self.lbl_count = tk.Label(summary_frame,
                                    text="عدد العاملين النشطين: 0",
                                    font=('Arial', 13, 'bold'),
                                    bg='#34495e', fg='white')
        self.lbl_count.pack(side='left', padx=25, pady=25)

        self.lbl_total_advances = tk.Label(summary_frame,
                                             text="إجمالي السلف: 0.00 ج.م",
                                             font=('Arial', 13, 'bold'),
                                             bg='#34495e', fg='#e74c3c')
        self.lbl_total_advances.pack(side='left', padx=25, pady=25)

        self.lbl_total_due = tk.Label(summary_frame,
                                        text="إجمالي المستحقات: 0.00 ج.م",
                                        font=('Arial', 14, 'bold'),
                                        bg='#34495e', fg='#2ecc71')
        self.lbl_total_due.pack(side='right', padx=25, pady=25)

        # ============ الأزرار ============
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=8)

        tk.Button(btn_frame, text="🖨️ Excel",
                  command=self.export_excel,
                  bg='#27ae60', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="🔄 تحديث", command=self.load_report,
                  bg='#3498db', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)


        tk.Button(btn_frame, text="🖼️ طباعه",
                  command=self.export_image,
                  bg='#8e44ad', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إغلاق", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=15, pady=8,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)
        # تحميل التقرير
        self.load_report()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def get_last_active_period(self, conn, employee_id):
        """
        جلب آخر مدة عمل نشطة للعامل:
        - المدة التي last_date >= اليوم
        - أو المدة ذات أكبر last_date
        - إذا لم يكن هناك last_date، نعتبره تاريخ اليوم
        """
        c = conn.cursor()
        today = datetime.now().strftime('%Y-%m-%d')

        # 1. ابحث عن مدة نشطة (last_date >= اليوم)
        c.execute("""SELECT id, first_date, last_date, vacations, deductions, bonus
                     FROM employee_periods 
                     WHERE employee_id=? AND (last_date >= ? OR last_date IS NULL OR last_date = '')
                     ORDER BY first_date DESC LIMIT 1""",
                  (employee_id, today))
        row = c.fetchone()
        if row:
            return row

        # 2. إذا لم توجد، جلب آخر مدة (بأكبر first_date)
        c.execute("""SELECT id, first_date, last_date, vacations, deductions, bonus
                     FROM employee_periods 
                     WHERE employee_id=?
                     ORDER BY first_date DESC LIMIT 1""",
                  (employee_id,))
        return c.fetchone()

    def get_period_advances(self, conn, employee_name, first_date, last_date,
                            period_id=None, employee_id=None,
                            period_last_date=None):
        """السلف + باقي الحساب خلال الفترة"""
        if not first_date or not last_date:
            return 0

        c = conn.cursor()

        # 1. السلف من paids
        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                     WHERE sub_name=? AND date >= ? AND date <= ?
                       AND COALESCE(payment_method, 'نقدي') != 'الخزينة'
                       AND (sub_name2 IS NULL OR sub_name2 = ''
                            OR sub_name2 = 'سلفة')""",
                  (employee_name, first_date, last_date))
        from_paids = c.fetchone()[0] or 0

        # 2. السلف من treasury_actions
        c.execute("""SELECT COALESCE(SUM(t.amount), 0)
                     FROM treasury_actions t
                     LEFT JOIN paids p ON t.paid_id = p.id
                     WHERE t.action_type='سداد عامل'
                       AND t.beneficiary=?
                       AND t.date >= ? AND t.date <= ?
                       AND (p.id IS NULL OR p.sub_name2 IS NULL
                            OR p.sub_name2 = '' OR p.sub_name2 = 'سلفة')""",
                  (employee_name, first_date, last_date))
        from_treasury = c.fetchone()[0] or 0

        # 3. ✅ باقي الحساب (لو كنا نعرف المدة)
        remaining = 0
        if period_id is not None and employee_id is not None:
            remaining = self._get_remaining_for_period(
                c, period_id, employee_id, employee_name, period_last_date)

        return from_paids + from_treasury + remaining

    def _get_remaining_for_period(self, c, period_id, employee_id,
                                  employee_name, period_last_date):
        """باقي الحساب المنتمي لهذه المدة"""
        if not period_last_date:
            return 0

        c.execute("""SELECT p.id, p.date, p.value FROM paids p
                     WHERE p.sub_name = ?
                       AND p.sub_name2 = 'باقي حساب'
                       AND p.date > ?""",
                  (employee_name, period_last_date))
        payments = c.fetchall()

        total = 0
        for p_id, p_date, p_value in payments:
            c.execute("""SELECT id FROM employee_periods
                         WHERE employee_id = ?
                           AND last_date IS NOT NULL
                           AND last_date <= ?
                         ORDER BY last_date DESC LIMIT 1""",
                      (employee_id, p_date))
            row = c.fetchone()
            if row and row[0] == period_id:
                total += p_value or 0

        return total

    def export_image(self):
        """تصدير التقرير إلى صورة PNG"""
        from tkinter import filedialog

        today = datetime.now().strftime('%Y-%m-%d')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"تقرير_العاملين_النشطين_{today}.png",
            title="حفظ التقرير صورة"
        )
        if not file_path:
            return

        # جمع بيانات الجدول
        columns = ['ID', 'اسم العامل', 'الوظيفة', 'أول يوم', 'آخر يوم',
                   'أيام العمل', 'اليومية', 'المستحق الأساسي',
                   'المكافآت', 'الخصومات', 'السلف', 'الصافي المستحق']

        rows_data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if values[1] == 'الإجمالي':  # تخطي صف الإجمالي
                continue
            rows_data.append(values)

        # الإجماليات
        summary_items = [
            (self.lbl_count.cget('text'),
             ""),
            (self.lbl_total_advances.cget('text'), ""),
            (self.lbl_total_due.cget('text'), ""),
        ]

        # تبسيط الإجماليات
        summary_items = [
            (self.lbl_count.cget('text'), ''),
            (self.lbl_total_advances.cget('text'), ''),
            (self.lbl_total_due.cget('text'), ''),
        ]

        success = export_table_to_image(
            file_path=file_path,
            title="تقرير العاملين النشطين",
            subtitle=f"تاريخ التقرير: {today}",
            columns=columns,
            rows=rows_data,
            summary_items=[
                ("عدد العاملين النشطين",
                 self.lbl_count.cget('text').replace('عدد العاملين النشطين:', '').strip()),
                ("إجمالي السلف",
                 self.lbl_total_advances.cget('text').replace('إجمالي السلف:', '').strip()),
                ("إجمالي المستحقات",
                 self.lbl_total_due.cget('text').replace('إجمالي المستحقات:', '').strip()),
            ],
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح", f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass


    def calculate_work_days(self, first_date, last_date, vacations):
        """حساب أيام العمل الفعلية"""
        if not first_date or not last_date:
            return 0
        try:
            start = datetime.strptime(first_date, '%Y-%m-%d')
            end = datetime.strptime(last_date, '%Y-%m-%d')
            total_days = (end - start).days + 1
            return max(0, total_days - (vacations or 0))
        except:
            return 0

    def load_report(self):
        """تحميل تقرير العاملين النشطين"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = self.get_connection()
        c = conn.cursor()

        today = datetime.now().strftime('%Y-%m-%d')

        # جلب كل العاملين
        c.execute("""SELECT id, name, position, daily_wage, status
                     FROM employees ORDER BY name""")
        employees = c.fetchall()

        total_advances_sum = 0
        total_due_sum = 0
        active_count = 0

        for emp_id, name, position, daily_wage, status in employees:
            daily_wage = daily_wage or 0

            # جلب آخر مدة عمل نشطة
            period = self.get_last_active_period(conn, emp_id)

            if not period:
                # لا يوجد مدد — العامل لم يبدأ
                continue

            pid, first_date, last_date, vacations, deductions, bonus = period
            vacations = vacations or 0
            deductions = deductions or 0
            bonus = bonus or 0

            # ✅ إذا لم يكن هناك last_date، نستخدم تاريخ اليوم
            effective_last_date = last_date if last_date else today

            # حساب أيام العمل
            work_days = self.calculate_work_days(first_date, effective_last_date, vacations)
            base_due = work_days * daily_wage

            # ✅ حساب السلف خلال الفترة
            advances = self.get_period_advances(
                conn, name, first_date, effective_last_date,
                period_id=pid,
                employee_id=emp_id,
                period_last_date=last_date
            )

            # الصافي المستحق
            net_due = base_due + bonus - deductions - advances

            # ✅ هل العامل "نشط"؟
            # الشرط: آخر يوم عمل >= اليوم، أو لم يُدخل آخر يوم
            is_active = False
            if not last_date:
                # لم يُدخل آخر يوم → يعمل حالياً
                is_active = True
            elif last_date >= today:
                is_active = True

            if not is_active:
                continue

            active_count += 1
            total_advances_sum += advances
            total_due_sum += net_due

            self.tree.insert('', 'end', iid=emp_id, values=(
                emp_id,
                name,
                position or '',
                first_date or '',
                last_date or today + " (اليوم)",
                work_days,
                f"{daily_wage:,.2f}",
                f"{base_due:,.2f}",
                f"{bonus:,.2f}",
                f"{deductions:,.2f}",
                f"{advances:,.2f}",
                f"{net_due:,.2f}"
            ))

        conn.close()

        # ═══════════ صف الإجمالي في الجدول ═══════════
        if active_count > 0:
            self.tree.insert('', 'end', values=(
                '', 'الإجمالي', '', '', '',
                '', '', '', '', '',
                f"{total_advances_sum:,.2f}",
                f"{total_due_sum:,.2f}"
            ), tags=('total_row',))

        # تحديث الإجماليات في الأسفل
        self.lbl_count.config(text=f"عدد العاملين النشطين: {active_count}")
        self.lbl_total_advances.config(
            text=f"إجمالي السلف: {total_advances_sum:,.2f} ج.م")
        self.lbl_total_due.config(
            text=f"إجمالي المستحقات: {total_due_sum:,.2f} ج.م")

    def export_excel(self):
        """تصدير التقرير إلى Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        today = datetime.now().strftime('%Y-%m-%d')
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"تقرير_العاملين_النشطين_{today}.xlsx",
            title="حفظ تقرير العاملين النشطين"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "العاملون النشطون"

            # الاتجاه من اليمين لليسار
            ws.sheet_view.rightToLeft = True

            # عرض الأعمدة
            col_widths = {'A': 6, 'B': 22, 'C': 15, 'D': 13, 'E': 13,
                          'F': 12, 'G': 12, 'H': 15, 'I': 12, 'J': 12,
                          'K': 15, 'L': 15}
            for col, width in col_widths.items():
                ws.column_dimensions[col].width = width

            # ═══ العنوان ═══
            ws.merge_cells('A1:L1')
            title_cell = ws['A1']
            title_cell.value = f"تقرير العاملين النشطين — {today}"
            title_cell.font = Font(name='Arial', size=16, bold=True, color='FFFFFF')
            title_cell.fill = PatternFill("solid", fgColor="2C3E50")
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ═══ رؤوس الأعمدة ═══
            headers = ['ID', 'اسم العامل', 'الوظيفة', 'أول يوم', 'آخر يوم',
                       'أيام العمل', 'اليومية', 'المستحق الأساسي',
                       'المكافآت', 'الخصومات', 'السلف', 'الصافي المستحق']

            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=col_idx, value=header)
                cell.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            # ═══ البيانات ═══
            row_num = 3
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                for col_idx, val in enumerate(values, start=1):
                    cell = ws.cell(row=row_num, column=col_idx, value=val)
                    cell.font = Font(name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
                    # تلوين صف الإجمالي
                    if 'الإجمالي' in str(values[1]):
                        cell.fill = PatternFill("solid", fgColor="FFF3CD")
                        cell.font = Font(name='Arial', size=11, bold=True, color='8B4513')
                row_num += 1

            # ═══ الإجماليات في الأسفل ═══
            row_num += 1
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=12)
            cell = ws.cell(row=row_num, column=1, value=self.lbl_total_advances.cget('text'))
            cell.font = Font(name='Arial', size=12, bold=True, color='E74C3C')
            cell.alignment = Alignment(horizontal='center')

            row_num += 1
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=12)
            cell = ws.cell(row=row_num, column=1, value=self.lbl_total_due.cget('text'))
            cell.font = Font(name='Arial', size=12, bold=True, color='27AE60')
            cell.alignment = Alignment(horizontal='center')

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            messagebox.showerror("خطأ", f"فشل التصدير: {e}")




class EmployeePeriodFormWindow:
    """نافذة إدخال/تعديل مدة عمل"""
    def __init__(self, parent_window, period_id=None):
        self.parent_window = parent_window  # نافذة EmployeePeriodsWindow
        self.period_id = period_id
        self.is_edit = period_id is not None

        self.window = tk.Toplevel(parent_window.window)
        title = "تعديل مدة عمل" if self.is_edit else "إضافة مدة عمل جديدة"
        self.window.title(title)
        self.window.geometry("560x520")
        self.window.grab_set()
        self.window.resizable(False, False)
        self.window.configure(bg='#ecf0f1')

        # ============ العنوان ============
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        icon = "✏️" if self.is_edit else "➕"
        tk.Label(title_frame, text=f"{icon} {title}",
                 font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white').pack(pady=15)

        tk.Label(title_frame, text=f"العامل: {parent_window.emp_name}",
                 font=('Arial', 11), bg='#2c3e50', fg='#f1c40f').pack()

        # ============ الإطار الرئيسي ============
        main_frame = tk.Frame(self.window, bg='#ecf0f1', padx=25, pady=20)
        main_frame.pack(fill='both', expand=True)

        # أول يوم عمل
        tk.Label(main_frame, text="أول يوم عمل: *",
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='e').grid(row=0, column=0, padx=5, pady=8, sticky='e')
        self.entry_first = DateEntry(main_frame, width=20,
                                       date_pattern='yyyy-mm-dd')
        self.entry_first.grid(row=0, column=1, padx=5, pady=8, sticky='w')

        # آخر يوم عمل (اختياري)
        tk.Label(main_frame, text="آخر يوم عمل:",
                 font=('Arial', 11), bg='#ecf0f1',
                 anchor='e').grid(row=1, column=0, padx=5, pady=8, sticky='e')

        last_frame = tk.Frame(main_frame, bg='#ecf0f1')
        last_frame.grid(row=1, column=1, padx=5, pady=8, sticky='w')

        self.entry_last = DateEntry(last_frame, width=15, date_pattern='yyyy-mm-dd')
        self.entry_last.pack(side='left')

        tk.Button(last_frame, text="✖ مسح", command=self.clear_last_date,
                  bg='#e74c3c', fg='white', font=('Arial', 9),
                  padx=5, pady=2).pack(side='left', padx=5)

        tk.Label(last_frame, text="(اتركه فارغاً للمدة الجارية)",
                 font=('Arial', 9, 'italic'), bg='#ecf0f1', fg='#7f8c8d').pack(side='left')

        # حالة افتراضية: فارغ (مدة جارية)
        self.entry_last.delete(0, tk.END)

        # الإجازات
        tk.Label(main_frame, text="الإجازات (بالأيام):",
                 font=('Arial', 11), bg='#ecf0f1',
                 anchor='e').grid(row=2, column=0, padx=5, pady=8, sticky='e')
        self.entry_vacations = tk.Entry(main_frame, width=20, font=('Arial', 11))
        self.entry_vacations.grid(row=2, column=1, padx=5, pady=8, sticky='w')
        self.entry_vacations.insert(0, "0")

        # الخصومات
        tk.Label(main_frame, text="الخصومات (ج.م):",
                 font=('Arial', 11), bg='#ecf0f1',
                 anchor='e').grid(row=3, column=0, padx=5, pady=8, sticky='e')
        self.entry_deductions = tk.Entry(main_frame, width=20, font=('Arial', 11))
        self.entry_deductions.grid(row=3, column=1, padx=5, pady=8, sticky='w')
        self.entry_deductions.insert(0, "0")

        # المكافآت
        tk.Label(main_frame, text="المكافآت (ج.م):",
                 font=('Arial', 11), bg='#ecf0f1',
                 anchor='e').grid(row=4, column=0, padx=5, pady=8, sticky='e')
        self.entry_bonus = tk.Entry(main_frame, width=20, font=('Arial', 11))
        self.entry_bonus.grid(row=4, column=1, padx=5, pady=8, sticky='w')
        self.entry_bonus.insert(0, "0")

        # ملاحظات
        tk.Label(main_frame, text="ملاحظات:", font=('Arial', 11),
                 bg='#ecf0f1', anchor='e').grid(row=5, column=0, padx=5, pady=8, sticky='ne')
        self.entry_notes = tk.Text(main_frame, width=20, height=3, font=('Arial', 11))
        self.entry_notes.grid(row=5, column=1, padx=5, pady=8, sticky='w')

        # معلومة توضيحية
        info_frame = tk.Frame(self.window, bg='#fff9e6', pady=8)
        info_frame.pack(fill='x', padx=20, pady=5)
        tk.Label(info_frame,
                 text="💡 السلف تُحسب تلقائياً من صفحة المصاريف خلال هذه الفترة",
                 font=('Arial', 10, 'italic'), bg='#fff9e6', fg='#8B4513').pack()

        # ============ أزرار الحفظ والإلغاء ============
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="💾 حفظ", command=self.save,
                  bg='#27ae60', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ إلغاء", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=25, pady=10,
                  font=('Arial', 12, 'bold')).pack(side='left', padx=5)

        # تحميل البيانات إن كان تعديلاً
        if self.is_edit:
            self.load_period_data()

        self.entry_first.focus()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def update_employee_status(self, conn, employee_id):
        """تحديث حقل الحالة في جدول employees بناءً على مدد العمل"""
        c = conn.cursor()
        today = datetime.now().strftime('%Y-%m-%d')

        c.execute("""SELECT COUNT(*) FROM employee_periods 
                     WHERE employee_id=? AND (last_date >= ? OR last_date IS NULL OR last_date = '')""",
                  (employee_id, today))
        active = c.fetchone()[0]

        if active > 0:
            status = 'نشط'
        else:
            c.execute("SELECT COUNT(*) FROM employee_periods WHERE employee_id=?",
                      (employee_id,))
            total = c.fetchone()[0]
            status = 'لم يبدأ' if total == 0 else 'غير نشط'

        c.execute("UPDATE employees SET status=? WHERE id=?", (status, employee_id))

    def clear_last_date(self):
        """مسح حقل آخر يوم عمل (لمدة جارية)"""
        self.entry_last.delete(0, tk.END)

    def load_period_data(self):
        """جلب بيانات المدة لتعبئة الحقول"""
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT first_date, last_date, vacations, deductions, bonus, notes
                     FROM employee_periods WHERE id=?""", (self.period_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "المدة غير موجودة")
            self.window.destroy()
            return

        first_date, last_date, vacations, deductions, bonus, notes = row

        if first_date:
            try:
                self.entry_first.set_date(first_date)
            except:
                pass
        if last_date:
            try:
                self.entry_last.set_date(last_date)
            except:
                pass
        else:
            # مدة جارية — اترك الحقل فارغاً
            self.entry_last.delete(0, tk.END)

        self.entry_vacations.delete(0, tk.END)
        self.entry_vacations.insert(0, str(vacations or 0))

        self.entry_deductions.delete(0, tk.END)
        self.entry_deductions.insert(0, str(deductions or 0))

        self.entry_bonus.delete(0, tk.END)
        self.entry_bonus.insert(0, str(bonus or 0))

        self.entry_notes.insert('1.0', notes or '')

    def save(self):
        """حفظ المدة (إضافة أو تعديل) - آخر يوم اختياري"""
        first_date = self.entry_first.get().strip()

        if not first_date:
            messagebox.showwarning("تنبيه", "أول يوم عمل مطلوب")
            return

        # ✅ آخر يوم اختياري
        last_date = self.entry_last.get().strip()
        is_open_period = not last_date  # مدة جارية

        # التحقق من التاريخ
        try:
            d1 = datetime.strptime(first_date, '%Y-%m-%d')
            if last_date:
                d2 = datetime.strptime(last_date, '%Y-%m-%d')
                if d2 < d1:
                    messagebox.showerror("خطأ",
                                         "آخر يوم عمل يجب أن يكون بعد أول يوم عمل أو نفس التاريخ")
                    return
        except Exception as e:
            messagebox.showerror("خطأ", f"صيغة التاريخ غير صحيحة: {e}")
            return

        try:
            vacations = int(self.entry_vacations.get().strip() or 0)
            if vacations < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("خطأ", "الإجازات يجب أن تكون رقماً موجباً")
            return

        try:
            deductions = float(self.entry_deductions.get().strip() or 0)
            if deductions < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("خطأ", "الخصومات يجب أن تكون رقماً موجباً")
            return

        try:
            bonus = float(self.entry_bonus.get().strip() or 0)
            if bonus < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("خطأ", "المكافآت يجب أن تكون رقماً موجباً")
            return

        notes = self.entry_notes.get('1.0', 'end').strip()
        employee_id = self.parent_window.employee_id

        conn = self.get_connection()
        c = conn.cursor()

        try:
            if self.is_edit:
                # ✅ تعديل
                # التحقق من التداخل فقط إذا كان last_date موجوداً
                if last_date:
                    c.execute("""SELECT id FROM employee_periods 
                                 WHERE employee_id=? AND id != ? AND
                                       ((first_date <= ? AND (last_date >= ? OR last_date IS NULL OR last_date = '')) OR
                                        (first_date >= ? AND first_date <= ?))""",
                              (employee_id, self.period_id,
                               last_date, first_date,
                               first_date, last_date))
                    if c.fetchone():
                        if not messagebox.askyesno("تحذير",
                                                   "المدة قد تتداخل مع مدة أخرى للعامل نفسه.\n\n"
                                                   "هل تريد المتابعة على أي حال؟"):
                            conn.close()
                            return

                c.execute("""UPDATE employee_periods 
                             SET first_date=?, last_date=?, vacations=?,
                                 deductions=?, bonus=?, notes=?
                             WHERE id=?""",
                          (first_date, last_date if last_date else None,
                           vacations, deductions, bonus, notes, self.period_id))

                msg = "✅ تم تعديل المدة بنجاح"
            else:
                # ✅ إضافة
                if last_date:
                    c.execute("""SELECT id FROM employee_periods 
                                 WHERE employee_id=? AND 
                                       ((first_date <= ? AND (last_date >= ? OR last_date IS NULL OR last_date = '')) OR
                                        (first_date >= ? AND first_date <= ?))""",
                              (employee_id,
                               last_date, first_date,
                               first_date, last_date))
                    if c.fetchone():
                        if not messagebox.askyesno("تحذير",
                                                   "المدة قد تتداخل مع مدة أخرى للعامل نفسه.\n\n"
                                                   "هل تريد المتابعة على أي حال؟"):
                            conn.close()
                            return

                c.execute("""INSERT INTO employee_periods 
                             (employee_id, first_date, last_date, vacations,
                              deductions, bonus, notes, created_at)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                          (employee_id, first_date,
                           last_date if last_date else None,
                           vacations, deductions, bonus, notes,
                           datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

                if is_open_period:
                    msg = "✅ تمت إضافة المدة (جارية بدون تاريخ نهاية)"
                else:
                    msg = "✅ تمت إضافة المدة بنجاح"

            # ✅ تحديث حالة العامل
            self.update_employee_status(conn, employee_id)

            conn.commit()
            messagebox.showinfo("نجاح", msg)
            self.window.destroy()
            self.parent_window.load_periods()

        except Exception as e:
            conn.rollback()
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل الحفظ: {e}")
        finally:
            conn.close()

#============================================================================

class EmployeePeriodsWindow:
    """نافذة مدد عمل العامل (أزرار + جدول فقط)"""
    def __init__(self, parent_page, employee_id):
        self.parent_page = parent_page
        self.employee_id = employee_id
        self.selected_period_id = None

        self.window = tk.Toplevel(parent_page.winfo_toplevel())
        self.window.title("مدد العمل")
        self.window.geometry("1250x700")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        # جلب بيانات العامل
        conn = parent_page.get_connection()
        c = conn.cursor()
        c.execute("""SELECT name, phone, position, daily_wage, status 
                     FROM employees WHERE id=?""", (employee_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            messagebox.showerror("خطأ", "العامل غير موجود")
            self.window.destroy()
            return

        self.emp_name, self.emp_phone, self.emp_position, self.daily_wage, self.emp_status = row
        self.daily_wage = self.daily_wage or 0

        # ============ العنوان ============
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=80)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text=f"📅 مدد عمل: {self.emp_name}",
                 font=('Arial', 17, 'bold'), bg='#2c3e50', fg='white').pack(side='left', padx=25, pady=18)

        info_parts = [f"💰 اليومية: {self.daily_wage:.2f} ج.م"]
        if self.emp_position:
            info_parts.insert(0, f"💼 {self.emp_position}")
        info_parts.append(f"الحالة: {self.emp_status}")

        tk.Label(title_frame, text=" | ".join(info_parts),
                 font=('Arial', 11), bg='#2c3e50', fg='#bdc3c7').pack(side='right', padx=25)

        # ============ شريط الأزرار ============
        btn_frame = tk.LabelFrame(self.window, text="العمليات",
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#34495e',
                                    padx=15, pady=12)
        btn_frame.pack(fill='x', padx=15, pady=(10, 5))

        tk.Button(btn_frame, text="➕ إضافة مدة", command=self.add_period,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        self.btn_edit = tk.Button(btn_frame, text="✏️ تعديل",
                                    command=self.edit_period,
                                    bg='#f39c12', fg='white',
                                    font=('Arial', 11, 'bold'),
                                    padx=20, pady=8, width=15, state='disabled')
        self.btn_edit.pack(side='left', padx=5)

        self.btn_delete = tk.Button(btn_frame, text="🗑️ حذف",
                                      command=self.delete_period,
                                      bg='#e74c3c', fg='white',
                                      font=('Arial', 11, 'bold'),
                                      padx=20, pady=8, width=15, state='disabled')
        self.btn_delete.pack(side='left', padx=5)

        self.btn_statement = tk.Button(btn_frame, text="📄 كشف الحساب",
                                         command=self.open_statement,
                                         bg='#2980b9', fg='white',
                                         font=('Arial', 11, 'bold'),
                                         padx=20, pady=8, width=15, state='disabled')
        self.btn_statement.pack(side='left', padx=5)

        tk.Button(btn_frame, text="🔄 تحديث", command=self.load_periods,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=8, width=15).pack(side='left', padx=5)

        # ============ جدول العرض ============
        tree_frame = tk.LabelFrame(self.window, text="مدد العمل المسجلة",
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#34495e',
                                     padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 10))

        columns = ('id', 'first_date', 'last_date', 'work_days', 'vacations',
                   'deductions', 'bonus', 'advances', 'net_due')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=12)

        headers = {
            'id': 'ID',
            'first_date': 'أول يوم عمل',
            'last_date': 'آخر يوم عمل',
            'work_days': 'أيام العمل',
            'vacations': 'الإجازات',
            'deductions': 'الخصومات',
            'bonus': 'المكافآت',
            'advances': 'السلف (محسوبة)',
            'net_due': 'صافي المستحق'
        }
        for col, h in headers.items():
            self.tree.heading(col, text=h)

        self.tree.column('id', width=50, anchor='center')
        self.tree.column('first_date', width=120, anchor='center')
        self.tree.column('last_date', width=120, anchor='center')
        self.tree.column('work_days', width=100, anchor='center')
        self.tree.column('vacations', width=90, anchor='center')
        self.tree.column('deductions', width=110, anchor='center')
        self.tree.column('bonus', width=110, anchor='center')
        self.tree.column('advances', width=140, anchor='center')
        self.tree.column('net_due', width=140, anchor='center')
        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        self.tree.tag_configure('positive', background='#d4edda')
        self.tree.tag_configure('negative', background='#f8d7da')
        self.tree.tag_configure('open_period', background='#d1ecf1')  # أزرق فاتح

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-Button-1>', lambda e: self.edit_period())

        # ============ شريط الإجماليات ============
        summary = tk.Frame(self.window, bg='#34495e', height=60)
        summary.pack(fill='x', padx=15, pady=(0, 10))
        summary.pack_propagate(False)

        self.lbl_count = tk.Label(summary, text="عدد المدد: 0",
                                    font=('Arial', 12, 'bold'),
                                    bg='#34495e', fg='white')
        self.lbl_count.pack(side='left', padx=25, pady=15)

        self.lbl_total_due = tk.Label(summary, text="إجمالي المستحق: 0.00",
                                        font=('Arial', 14, 'bold'),
                                        bg='#34495e', fg='#f1c40f')
        self.lbl_total_due.pack(side='right', padx=25, pady=15)

        # شريط الحالة
        self.status_bar = tk.Label(self.window, text="جاهز",
                                     font=('Arial', 10),
                                     bg='#2c3e50', fg='white',
                                     anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        self.load_periods()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def get_period_advances(self, first_date, last_date,
                            period_id=None, employee_id=None,
                            period_last_date=None):
        """
        حساب السلف خلال فترة + باقي الحساب المنتمي للمدة.

        - سلفة: date in [first_date, last_date]
        - باقي حساب: يُحسب فقط إذا كانت هذه المدة آخر مدة
          منتهية قبل تاريخ الدفع
        """
        if not first_date:
            return 0
        if not last_date:
            last_date = datetime.now().strftime('%Y-%m-%d')

        conn = self.get_connection()
        c = conn.cursor()

        # 1. السلف من paids
        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids
                     WHERE sub_name=? AND date >= ? AND date <= ?
                       AND COALESCE(payment_method, 'نقدي') != 'الخزينة'
                       AND (sub_name2 IS NULL OR sub_name2 = ''
                            OR sub_name2 = 'سلفة')""",
                  (self.emp_name, first_date, last_date))
        from_paids = c.fetchone()[0] or 0

        # 2. السلف من treasury_actions
        c.execute("""SELECT COALESCE(SUM(t.amount), 0)
                     FROM treasury_actions t
                     LEFT JOIN paids p ON t.paid_id = p.id
                     WHERE t.action_type='سداد عامل'
                       AND t.beneficiary=?
                       AND t.date >= ? AND t.date <= ?
                       AND (p.id IS NULL OR p.sub_name2 IS NULL
                            OR p.sub_name2 = '' OR p.sub_name2 = 'سلفة')""",
                  (self.emp_name, first_date, last_date))
        from_treasury = c.fetchone()[0] or 0

        # 3. ✅ باقي الحساب (المرتبط بهذه المدة)
        remaining = 0
        if period_id is not None and employee_id is not None:
            remaining = self._get_remaining_for_period(
                c, period_id, employee_id, self.emp_name, period_last_date)

        conn.close()
        return from_paids + from_treasury + remaining

    def _get_remaining_for_period(self, c, period_id, employee_id,
                                  employee_name, period_last_date):
        """
        حساب "باقي الحساب" المنتمي لمدة معينة.
        القاعدة: باقي حساب ينتمي للمدة إذا كانت هذه المدة هي
        آخر مدة منتهية قبل تاريخ الدفع.
        """
        if not period_last_date:
            return 0

        # جلب كل مدفوعات "باقي حساب" لهذا العامل بعد نهاية المدة
        c.execute("""SELECT p.id, p.date, p.value FROM paids p
                     WHERE p.sub_name = ?
                       AND p.sub_name2 = 'باقي حساب'
                       AND p.date > ?""",
                  (employee_name, period_last_date))
        payments = c.fetchall()

        total = 0
        for p_id, p_date, p_value in payments:
            # نجد المدة المالكة لهذا الدفع
            c.execute("""SELECT id FROM employee_periods
                         WHERE employee_id = ?
                           AND last_date IS NOT NULL
                           AND last_date <= ?
                         ORDER BY last_date DESC LIMIT 1""",
                      (employee_id, p_date))
            row = c.fetchone()
            if row and row[0] == period_id:
                total += p_value or 0

        return total

    def calculate_work_days(self, first_date, last_date, vacations):
        """حساب أيام العمل الفعلية"""
        if not first_date or not last_date:
            return 0
        try:
            start = datetime.strptime(first_date, '%Y-%m-%d')
            end = datetime.strptime(last_date, '%Y-%m-%d')
            total_days = (end - start).days + 1
            return max(0, total_days - (vacations or 0))
        except:
            return 0

    def load_periods(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT id, first_date, last_date, vacations, deductions, bonus
                     FROM employee_periods WHERE employee_id=?
                     ORDER BY first_date DESC""", (self.employee_id,))
        rows = c.fetchall()

        total_due_sum = 0

        for pid, first_date, last_date, vacations, deductions, bonus in rows:
            vacations = vacations or 0
            deductions = deductions or 0
            bonus = bonus or 0

            # ═════════ تاريخ نهاية أيام العمل (الأصلي) ═════════
            if last_date:
                work_end_date = last_date
            else:
                work_end_date = datetime.now().strftime('%Y-%m-%d')

            # ═════════ تاريخ نهاية المدفوعات (ممتد للفترة المفتوحة) ═════════
            # ✅ السلف تُحسب داخل المدة فقط (من first_date إلى last_date)
            if last_date:
                # المدة مغلقة → نتوقف عند last_date بالضبط
                payment_end_date = last_date
            else:
                # مدة جارية → حتى اليوم
                payment_end_date = datetime.now().strftime('%Y-%m-%d')

            # ═════════ الحسابات ═════════
            work_days = self.calculate_work_days(first_date, work_end_date, vacations)
            base_due = work_days * self.daily_wage

            # ✅ استخدام payment_end_date بدلاً من work_end_date
            advances = self.get_period_advances(
                first_date, payment_end_date,
                period_id=pid,
                employee_id=self.employee_id,
                period_last_date=last_date
            )

            net_due = base_due + bonus - deductions - advances

            total_due_sum += net_due

            tag = 'positive' if net_due >= 0 else 'negative'
            if not last_date:
                tag = 'open_period'

            last_date_display = last_date if last_date else "🟢 جارية"

            self.tree.insert('', 'end', iid=pid, tags=(tag,), values=(
                pid,
                first_date or '',
                last_date_display,
                work_days,
                vacations,
                f"{deductions:.2f}",
                f"{bonus:.2f}",
                f"{advances:.2f}",
                f"{net_due:.2f}"
            ))

        conn.close()

        self.lbl_count.config(text=f"عدد المدد: {len(rows)}")
        self.lbl_total_due.config(text=f"إجمالي المستحق: {total_due_sum:,.2f}")
        self.status_bar.config(text=f"عدد مدد العمل: {len(rows)}")

    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            self.selected_period_id = int(sel[0])
            self.btn_edit.config(state='normal')
            self.btn_delete.config(state='normal')
            self.btn_statement.config(state='normal')
        else:
            self.selected_period_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.btn_statement.config(state='disabled')

    def add_period(self):
        EmployeePeriodFormWindow(self, period_id=None)

    def edit_period(self):
        if self.selected_period_id is None:
            messagebox.showwarning("تنبيه", "اختر مدة من الجدول أولاً")
            return
        EmployeePeriodFormWindow(self, period_id=self.selected_period_id)

    def delete_period(self):
        if self.selected_period_id is None:
            messagebox.showwarning("تنبيه", "اختر مدة من الجدول أولاً")
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""SELECT first_date, last_date FROM employee_periods 
                     WHERE id=?""", (self.selected_period_id,))
        row = c.fetchone()
        conn.close()

        if not row:
            return

        first_date, last_date = row

        if not messagebox.askyesno("تأكيد",
                                   f"هل أنت متأكد من حذف المدة:\n"
                                   f"من {first_date} إلى {last_date}؟"):
            return

        conn = self.get_connection()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM employee_periods WHERE id=?",
                      (self.selected_period_id,))

            # ✅ إعادة حساب الحالة بعد الحذف
            c2 = conn.cursor()
            today = datetime.now().strftime('%Y-%m-%d')

            c2.execute("""SELECT COUNT(*) FROM employee_periods 
                          WHERE employee_id=? AND (last_date >= ? OR last_date IS NULL OR last_date = '')""",
                       (self.employee_id, today))
            active = c2.fetchone()[0]

            if active > 0:
                new_status = 'نشط'
            else:
                c2.execute("SELECT COUNT(*) FROM employee_periods WHERE employee_id=?",
                           (self.employee_id,))
                total = c2.fetchone()[0]
                new_status = 'لم يبدأ' if total == 0 else 'غير نشط'

            c2.execute("UPDATE employees SET status=? WHERE id=?",
                       (new_status, self.employee_id))

            conn.commit()
            messagebox.showinfo("نجاح", f"تم حذف المدة (الحالة الجديدة: {new_status})")
            self.selected_period_id = None
            self.btn_edit.config(state='disabled')
            self.btn_delete.config(state='disabled')
            self.btn_statement.config(state='disabled')
            self.load_periods()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("خطأ", f"فشل الحذف: {e}")
        finally:
            conn.close()

    def open_statement(self):
        if self.selected_period_id is None:
            messagebox.showwarning("تنبيه", "اختر مدة من الجدول أولاً")
            return
        PeriodStatementWindow(self, self.selected_period_id)

    # ================================================================================================================

class PeriodStatementWindow:
    """كشف حساب مدة عمل محددة"""
    def __init__(self, parent_window, period_id):
        self.parent_window = parent_window
        self.period_id = period_id

        self.window = tk.Toplevel(parent_window.window)
        self.window.title("كشف حساب المدة")
        self.window.geometry("1000x720")
        self.window.grab_set()
        self.window.configure(bg='#ecf0f1')

        emp_name = parent_window.emp_name
        daily_wage = parent_window.daily_wage

        # [جلب بيانات المدة]
        conn = parent_window.get_connection()
        c = conn.cursor()
        c.execute("""SELECT first_date, last_date, vacations, deductions, bonus, notes
        FROM employee_periods WHERE id=?""", (period_id,))
        row = c.fetchone()
        if not row:
            messagebox.showerror("خطأ", "المدة غير موجودة")
            conn.close()
            self.window.destroy()
            return
        first_date, last_date, vacations, deductions, bonus, notes = row
        notes = notes or ""  # ✅ الملاحظات (قد تكون فارغة)

        vacations = vacations or 0
        deductions = deductions or 0
        bonus = bonus or 0


        # ═════════ تاريخ نهاية أيام العمل ═════════
        if last_date:
            work_end_date = last_date
        else:
            work_end_date = datetime.now().strftime('%Y-%m-%d')

        # ═════════ ✅ السلف تُحسب داخل المدة فقط ═════════
        if last_date:
            # المدة مغلقة → نتوقف عند last_date بالضبط
            payment_end_date = last_date
        else:
            # مدة جارية → حتى اليوم
            payment_end_date = datetime.now().strftime('%Y-%m-%d')

        is_open_period = not last_date
        effective_last_date = payment_end_date

        # ═════════ الحسابات ═════════
        work_days = parent_window.calculate_work_days(first_date, work_end_date, vacations)
        base_due = work_days * daily_wage
        advances = parent_window.get_period_advances(
            first_date, payment_end_date,
            period_id=period_id,
            employee_id=parent_window.employee_id,
            period_last_date=last_date
        )

        net_due = base_due + bonus - deductions - advances

        # العنوان
        title_frame = tk.Frame(self.window, bg='#2c3e50', height=70)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text=f"📄 كشف حساب: {emp_name}",
                 font=('Arial', 15, 'bold'), bg='#2c3e50', fg='white').pack(side='left', padx=20, pady=15)
        period_text = f"من {first_date}"
        if last_date:
            period_text += f" إلى {last_date}"
        else:
            period_text += f" إلى اليوم ({effective_last_date}) 🟢 مدة جارية"

        tk.Label(title_frame, text=period_text,
                 font=('Arial', 11), bg='#2c3e50', fg='#f1c40f').pack(side='right', padx=20)

        # ملخص الحساب
        # ═══════════════════════════════════════════════════
        #        ملخص المدة - بتنسيق مطابق للصورة
        # ═══════════════════════════════════════════════════
        calc_outer = tk.Frame(self.window, bg='#ecf0f1')
        calc_outer.pack(fill='x', padx=10, pady=5)

        # ─── شريط العنوان البني ───
        title_bar = tk.Frame(calc_outer, bg='#8B4513', height=42)
        title_bar.pack(fill='x')
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text="📊 ملخص الحساب",
                 font=('Arial', 14, 'bold'),
                 bg='#8B4513', fg='white').pack(pady=8)

        # ─── جدول الملخص ───
        table_frame = tk.Frame(calc_outer, bg='white')
        table_frame.pack(fill='x', padx=0, pady=0)

        # حساب عدد أيام الفترة (باستخدام المتغيرات المحلية)
        try:
            total_days = (datetime.strptime(last_date, '%Y-%m-%d') -
                          datetime.strptime(first_date, '%Y-%m-%d')).days + 1
        except:
            total_days = work_days + vacations

        summary_rows = [
            ("اليومية", f"{daily_wage:,.2f} ج.م"),
            ("عدد أيام المدة", f"{total_days} يوم"),
            ("الإجازات", f"{vacations} يوم"),
            ("أيام العمل الفعلية", f"{work_days} يوم"),
            ("المستحق الأساسي", f"{base_due:,.2f} ج.م"),
            ("➕ المكافآت", f"{bonus:,.2f} ج.م"),
            ("➖ الخصومات", f"{deductions:,.2f} ج.م"),
            ("➖ السلف", f"{advances:,.2f} ج.م"),
        ]

        # ✅ إضافة صف الملاحظات فقط إذا كانت موجودة
        if notes.strip():
            summary_rows.append(("📝 ملاحظات", notes.strip()))

        for i, (label, value) in enumerate(summary_rows):
            bg_color = '#fafafa' if i % 2 == 0 else 'white'

            row_frame = tk.Frame(table_frame, bg=bg_color, height=40,
                                 highlightbackground='#d5d8dc',
                                 highlightthickness=1)
            row_frame.pack(fill='x')
            row_frame.pack_propagate(False)

            tk.Label(row_frame, text=value,
                     font=('Arial', 12, 'bold'),
                     bg=bg_color, fg='#2c3e50',
                     anchor='w').pack(side='left', padx=20)

            tk.Label(row_frame, text=label,
                     font=('Arial', 12, 'bold'),
                     bg=bg_color, fg='#34495e',
                     anchor='e').pack(side='right', padx=20)

        # ─── صف صافي المستحق (مميز) ───
        net_color = '#27ae60' if net_due >= 0 else '#e74c3c'
        net_bg = '#e8f8f0' if net_due >= 0 else '#fdedeb'

        net_row = tk.Frame(table_frame, bg=net_bg, height=60,
                           highlightbackground=net_color,
                           highlightthickness=3)
        net_row.pack(fill='x', pady=(6, 0))
        net_row.pack_propagate(False)

        tk.Label(net_row, text=f"{net_due:,.2f} ج.م",
                 font=('Arial', 18, 'bold'),
                 bg=net_bg, fg=net_color,
                 anchor='w').pack(side='left', padx=25)

        tk.Label(net_row, text="💰 صافي المستحق",
                 font=('Arial', 16, 'bold'),
                 bg=net_bg, fg='#8B4513',
                 anchor='e').pack(side='right', padx=25)

        # جدول الحركات خلال المدة
        tree_frame = tk.LabelFrame(self.window, text="📋 الحركات خلال المدة",
                                    font=('Arial', 12, 'bold'), bg='#ecf0f1',
                                    padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ('id', 'date', 'type_info', 'description', 'value')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=4)
        self.tree.heading('id', text='ID')
        self.tree.heading('date', text='التاريخ')
        self.tree.heading('type_info', text='النوع / طريقة الدفع')  # ✅ جديد
        self.tree.heading('description', text='البيان')
        self.tree.heading('value', text='المبلغ')

        self.tree.column('id', width=50, anchor='center')
        self.tree.column('date', width=100, anchor='center')
        self.tree.column('type_info', width=180, anchor='center')  # ✅ جديد
        self.tree.column('description', width=350, anchor='w')
        self.tree.column('value', width=120, anchor='center')

        self.tree.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

        # ═════════ 1. جلب السلف العادية من paids ═════════
        c.execute("""SELECT id, date, main_name, description, value,
                            payment_method, sub_name2
                     FROM paids
                     WHERE sub_name=? AND date >= ? AND date <= ?
                       AND (sub_name2 IS NULL OR sub_name2 = ''
                            OR sub_name2 = 'سلفة')
                     ORDER BY date, id""",
                  (emp_name, first_date, payment_end_date))
        paids_movements = c.fetchall()

        # ═════════ 2. جلب 'باقي حساب' المنتمي لهذه المدة ═════════
        remaining_movements = []
        if last_date:
            c.execute("""SELECT id, date, main_name, description, value,
                                payment_method, sub_name2
                         FROM paids
                         WHERE sub_name = ?
                           AND sub_name2 = 'باقي حساب'
                           AND date > ?
                         ORDER BY date, id""",
                      (emp_name, last_date))
            remaining_candidates = c.fetchall()

            for row in remaining_candidates:
                p_id, p_date = row[0], row[1]
                # نجد المدة المالكة لهذا الدفع
                c.execute("""SELECT id FROM employee_periods
                             WHERE employee_id = ?
                               AND last_date IS NOT NULL
                               AND last_date <= ?
                             ORDER BY last_date DESC LIMIT 1""",
                          (parent_window.employee_id, p_date))
                owner = c.fetchone()
                if owner and owner[0] == period_id:
                    remaining_movements.append(row)

        # ═════════ 3. جلب حركات الخزينة المباشرة (بدون paid_id) ═════════
        c.execute("""SELECT id, date,
                            'سداد خزينة' as main_name,
                            COALESCE(description, 'سداد من الخزينة') as description,
                            amount as value
                     FROM treasury_actions
                     WHERE action_type='سداد عامل'
                       AND beneficiary=?
                       AND date >= ? AND date <= ?
                       AND (paid_id IS NULL OR paid_id = 0)
                     ORDER BY date, id""",
                  (emp_name, first_date, payment_end_date))
        treasury_movements = c.fetchall()

        conn.close()

        # ═════════ دمج + ترتيب ═════════
        movements = []

        # 1. السلف العادية
        for row in paids_movements:
            rid, date, main_name, description, value, p_method, sub2 = row
            movements.append({
                'sort': (date or '', 0, rid or 0),
                'id': rid,
                'date': date,
                'type_info': self._build_type_display(sub2, p_method),
                'description': description or '',
                'value': value,
                'tag': ''
            })

        # 2. باقي الحساب
        for row in remaining_movements:
            rid, date, main_name, description, value, p_method, sub2 = row
            movements.append({
                'sort': (date or '', 1, rid or 0),
                'id': rid,
                'date': date,
                'type_info': self._build_type_display(sub2, p_method),
                'description': description or '',
                'value': value,
                'tag': 'remaining_row'
            })

        # 3. حركات الخزينة المباشرة
        for row in treasury_movements:
            rid, date, main_name, description, value = row
            movements.append({
                'sort': (date or '', 2, rid or 0),
                'id': rid,
                'date': date,
                'type_info': '💸 سلفة - 🏦 خزينة',
                'description': description or '',
                'value': value,
                'tag': 'treasury_row'
            })

        # ترتيب: الأحدث أولاً
        movements.sort(key=lambda m: m['sort'], reverse=True)

        # ═════════ إدراج في الجدول ═════════
        for m in movements:
            tag = m['tag']
            self.tree.insert('', 'end', tags=(tag,) if tag else (), values=(
                m['id'],
                m['date'] or '',
                m['type_info'],
                m['description'],
                f"{m['value']:.2f}" if m['value'] else '0.00'
            ))



        self.tree.tag_configure('treasury_row', background='#e8f4f8')
        self.tree.tag_configure('remaining_row', background='#fef5e7')  # ✅ جديد

        # أزرار
        btn_frame = tk.Frame(self.window, bg='#ecf0f1')
        btn_frame.pack(pady=8)

        tk.Button(btn_frame, text="💾 [Excel]", command=self.export_excel,
                  bg='#27ae60', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="🖼️ [طباعه]", command=self.export_image,
                  bg='#8e44ad', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        tk.Button(btn_frame, text="❌ [إغلاق]", command=self.window.destroy,
                  bg='#e74c3c', fg='white', padx=15, pady=5,
                  font=('Arial', 11, 'bold')).pack(side='left', padx=5)

        # حفظ البيانات للطباعة
        self.emp_name = emp_name
        self.first_date = first_date
        self.last_date = last_date
        self.daily_wage = daily_wage
        self.work_days = work_days
        self.vacations = vacations
        self.base_due = base_due
        self.bonus = bonus
        self.deductions = deductions
        self.advances = advances
        self.net_due = net_due
        self.movements = movements
        self.notes = notes  # ✅ حفظ الملاحظات للاستخدام في التصدير

    def _build_type_display(self, sub_name2, payment_method):
        """بناء نص 'النوع / طريقة الدفع'"""
        parts = []

        # النوع
        if sub_name2 == 'باقي حساب':
            parts.append('💰 باقي حساب')
        else:
            parts.append('💸 سلفة')

        # طريقة الدفع (فقط لو ليست نقدي)
        if payment_method == 'فودافون كاش':
            parts.append('📱 فودافون كاش')
        elif payment_method == 'الخزينة':
            parts.append('🏦 خزينة')
        # نقدي = لا نضيف شي

        return ' - '.join(parts)


    def export_excel(self):
        """تصدير كشف حساب المدة إلى Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import (Font, Alignment, PatternFill,
                                         Border, Side)
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        safe_name = self.emp_name.replace(' ', '_').replace('/', '_')
        default_name = f"كشف_حساب_{safe_name}_{self.first_date}.xlsx"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="حفظ كشف الحساب Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "كشف حساب"
            ws.sheet_view.rightToLeft = True

            # ─── عرض الأعمدة ───
            col_widths = {'A': 12, 'B': 15, 'C': 25, 'D': 40, 'E': 14}
            for col, w in col_widths.items():
                ws.column_dimensions[col].width = w

            def style_cell(cell, bg='FFFFFF', fg='000000',
                           bold=False, size=11, align='center'):
                cell.font = Font(name='Arial', size=size, bold=bold, color=fg)
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal=align,
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            row_num = 1

            # ═════════ العنوان ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            c = ws.cell(row=row_num, column=1,
                        value=f"📄 كشف حساب : {self.emp_name}")
            style_cell(c, bg='2C3E50', fg='FFFFFF', bold=True, size=14)
            ws.row_dimensions[row_num].height = 30
            row_num += 1

            # ─── الفترة ───
            period_text = f"من {self.first_date}"
            if self.last_date:
                period_text += f" إلى {self.last_date}"
            else:
                period_text += f" إلى اليوم"

            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            c = ws.cell(row=row_num, column=1, value=period_text)
            style_cell(c, bg='34495E', fg='F1C40F', bold=True, size=12)
            ws.row_dimensions[row_num].height = 25
            row_num += 2

            # ═════════ ملخص الحساب ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            c = ws.cell(row=row_num, column=1, value="📊 ملخص الحساب")
            style_cell(c, bg='8B4513', fg='FFFFFF', bold=True, size=13)
            ws.row_dimensions[row_num].height = 25
            row_num += 1

            # ─── صفوف الملخص ───
            try:
                total_days = (datetime.strptime(self.last_date, '%Y-%m-%d') -
                              datetime.strptime(self.first_date, '%Y-%m-%d')
                              ).days + 1
            except:
                total_days = self.work_days + self.vacations

            summary_rows = [
                ("اليومية", f"{self.daily_wage:,.2f} ج.م"),
                ("عدد أيام المدة", f"{total_days} يوم"),
                ("الإجازات", f"{self.vacations} يوم"),
                ("أيام العمل الفعلية", f"{self.work_days} يوم"),
                ("المستحق الأساسي", f"{self.base_due:,.2f} ج.م"),
                ("➕ المكافآت", f"{self.bonus:,.2f} ج.م"),
                ("➖ الخصومات", f"{self.deductions:,.2f} ج.م"),
                ("➖ السلف", f"{self.advances:,.2f} ج.م"),
            ]

            for label, value in summary_rows:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=4)
                c = ws.cell(row=row_num, column=1, value=label)
                style_cell(c, bg='F8F9FA', fg='2C3E50', bold=True,
                           align='right')
                c = ws.cell(row=row_num, column=5, value=value)
                style_cell(c, bg='F8F9FA', fg='1A5276', bold=True)
                row_num += 1

            # ─── صافي المستحق ───
            net_bg = 'D5F4E6' if self.net_due >= 0 else 'FDEDEC'
            net_fg = '1E8449' if self.net_due >= 0 else 'C0392B'

            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1, value="💰 صافي المستحق")
            style_cell(c, bg=net_bg, fg=net_fg, bold=True, size=12,
                       align='right')
            c = ws.cell(row=row_num, column=5,
                        value=f"{self.net_due:,.2f} ج.م")
            style_cell(c, bg=net_bg, fg=net_fg, bold=True, size=12)
            ws.row_dimensions[row_num].height = 30
            row_num += 2

            # ═════════ جدول الحركات ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            c = ws.cell(row=row_num, column=1,
                        value="📋 الحركات خلال المدة")
            style_cell(c, bg='8B4513', fg='FFFFFF', bold=True, size=13)
            ws.row_dimensions[row_num].height = 25
            row_num += 1

            # ─── رؤوس الأعمدة ───
            headers = ['ID', 'التاريخ', 'النوع / طريقة الدفع',
                       'البيان', 'المبلغ']
            for col_idx, h in enumerate(headers, start=1):
                c = ws.cell(row=row_num, column=col_idx, value=h)
                style_cell(c, bg='34495E', fg='FFFFFF', bold=True, size=11)
            row_num += 1

            # ─── بيانات الحركات ───
            has_data = False
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                has_data = True
                for col_idx, val in enumerate(values, start=1):
                    c = ws.cell(row=row_num, column=col_idx, value=val)
                    style_cell(c, bg='FFFFFF', fg='2C3E50', size=10)
                row_num += 1

            if not has_data:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=5)
                c = ws.cell(row=row_num, column=1,
                            value="لا توجد حركات خلال المدة")
                style_cell(c, bg='F8F9FA', fg='95A5A6', align='center')
                row_num += 1

            # ─── التذييل ───
            row_num += 1
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=5)
            c = ws.cell(row=row_num, column=1,
                        value="شركة الزهراء للبرمجيات - المحاسب/حسام الدين "
                              "سيد ربيع 01003796984")
            style_cell(c, bg='ECF0F1', fg='7F8C8D', bold=True, size=10)

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    def export_image(self):
        """تصدير كشف الحساب إلى صورة PNG"""
        from tkinter import filedialog

        safe_name = self.emp_name.replace(' ', '_').replace('/', '_')
        default_name = f"كشف_حساب_{safe_name}_{self.first_date}.png"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg")],
            initialfile=default_name,
            title="حفظ كشف الحساب صورة"
        )
        if not file_path:
            return

        # ═════════ تجهيز البيانات ═════════
        try:
            total_days = (datetime.strptime(self.last_date, '%Y-%m-%d') -
                          datetime.strptime(self.first_date, '%Y-%m-%d')
                          ).days + 1
        except:
            total_days = self.work_days + self.vacations

        summary_items = [
            ("اليومية", f"{self.daily_wage:.2f} ج.م"),
            ("عدد أيام المدة", f"{total_days} يوم"),
            ("الإجازات", f"{self.vacations} يوم"),
            ("أيام العمل الفعلية", f"{self.work_days} يوم"),
            ("المستحق الأساسي", f"{self.base_due:.2f} ج.م"),
            ("المكافآت", f"{self.bonus:.2f} ج.م"),
            ("الخصومات", f"{self.deductions:.2f} ج.م"),
            ("السلف", f"{self.advances:.2f} ج.م"),
        ]

        # ✅ movements الآن list of dicts → استخرج القيم
        movements_list = []
        for m in self.movements:
            movements_list.append((
                m.get('date', ''),
                m.get('type_info', ''),
                m.get('description', ''),
                m.get('value', 0)
            ))

        success = export_statement_to_image(
            file_path=file_path,
            title=f"كشف حساب : {self.emp_name}",
            subtitle=f"الفترة: من {self.first_date} إلى {self.last_date}",
            summary_items=summary_items,
            net_label="صافي المستحق",
            net_value=f"{self.net_due:.2f} ج.م",
            movements=movements_list,
            movements_title="الحركات خلال المدة",
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم حفظ الصورة:\n{file_path}\n\nفتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    try:
                        subprocess.Popen(['xdg-open', file_path])
                    except:
                        pass


#=====================================================================================

class DailyReportPage(tk.Frame):
    """صفحة اليومية - تقرير يومي بالمبيعات والمصروفات"""
    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller
        add_program_banner(self)   # ✅ سطر واحد فقط

        # ═════════ العنوان ═════════
        tk.Label(self, text="📅 اليومية", font=('Arial', 14, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=5)

        # ═════════ فلاتر التاريخ ═════════
        filter_frame = tk.LabelFrame(self, text="اختر اليوم",
                                       font=('Arial', 11, 'bold'),
                                       bg='#ecf0f1', fg='#34495e',
                                       padx=5, pady=12)
        filter_frame.pack(fill='x', padx=20, pady=(0, 10))

        tk.Label(filter_frame, text="التاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(side='right', padx=20)

        self.date_entry = DateEntry(filter_frame, width=15,
                                      date_pattern='yyyy-mm-dd',
                                      font=('Arial', 11))
        self.date_entry.pack(side='right', padx=25)

        # أزرار سريعة
        tk.Button(filter_frame, text="📅 اليوم", command=self.set_today,
                  bg='#3498db', fg='white', font=('Arial', 11, 'bold'),
                  padx=12, pady=5).pack(side='right', padx=3)
        tk.Button(filter_frame, text="⬅️ أمس", command=self.set_yesterday,
                  bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                  padx=12, pady=5).pack(side='right', padx=3)

        # زر عرض
        tk.Button(filter_frame, text="🔍 عرض اليومية",
                  command=self.load_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=5, pady=8).pack(side='left', padx=5)

        # ═════════ زر التصدير ═════════
        tk.Button(filter_frame, text="🖼️ طباعه",
                  command=self.export_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=5, pady=8).pack(side='left', padx=5)

        # ✅ زر تصدير Excel
        tk.Button(filter_frame, text="💾 Excel",
                  command=self.export_excel,
                  bg='#16a085', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=5, pady=8).pack(side='left', padx=5)

        # ═════════ [بطاقة التقرير — قابلة للتمرير] ═════════
        self.show_euro_row = False
        self.show_dollar_row = False
        self.last_data = {}

        # حاوية خارجية
        container = tk.Frame(self, bg='#ecf0f1')
        container.pack(fill='both', expand=True, padx=5, pady=(0, 15))

        # ✅ Canvas + Scrollbar
        self._canvas = tk.Canvas(container, bg='white',
                                 highlightthickness=0)
        self._scrollbar = ttk.Scrollbar(container, orient='vertical',
                                        command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._scrollbar.set)

        self._scrollbar.pack(side='right', fill='y')
        self._canvas.pack(side='left', fill='both', expand=True)

        # ✅ الإطار الداخلي (بديل report_frame)
        self.report_frame = tk.Frame(self._canvas, bg='white',
                                     relief='solid', bd=2)
        self._canvas_window = self._canvas.create_window(
            (0, 0), window=self.report_frame, anchor='nw')

        # ✅ تحديث منطقة التمرير عند تغيّر حجم الإطار الداخلي
        def _on_frame_configure(event=None):
            self._canvas.configure(scrollregion=self._canvas.bbox('all'))

        def _on_canvas_configure(event):
            self._canvas.itemconfig(self._canvas_window, width=event.width)

        self.report_frame.bind('<Configure>', _on_frame_configure)
        self._canvas.bind('<Configure>', _on_canvas_configure)

        # ✅ تمرير بعجلة الفأرة
        def _on_mousewheel(event):
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        self._canvas.bind_all('<MouseWheel>', _on_mousewheel)

        # ✅ تحديث منطقة التمرير عند تغيير حجم النافذة الرئيسية
        self.bind('<Configure>', lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox('all')))

        self.build_report_ui()

        # تحميل بيانات اليوم
        self.set_today()

    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def get_day_name_ar(self, date_str):
        """إرجاع اسم اليوم بالعربية من تاريخ YYYY-MM-DD"""
        try:
            d = datetime.strptime(date_str, '%Y-%m-%d')
            days_ar = {
                0: 'الإثنين', 1: 'الثلاثاء', 2: 'الأربعاء',
                3: 'الخميس', 4: 'الجمعة', 5: 'السبت', 6: 'الأحد'
            }
            return days_ar.get(d.weekday(), '')
        except:
            return ''

    def build_report_ui(self):
        """بناء واجهة التقرير (تُحدَّث ديناميكياً)"""
        # مسح الإطار
        for w in self.report_frame.winfo_children():
            w.destroy()

        # ═══════════════════════════════════════════
        # الهيدر
        # ═══════════════════════════════════════════


        header2 = tk.Frame(self.report_frame, bg='#34495e', height=50)
        header2.pack(fill='x')
        header2.pack_propagate(False)
        tk.Label(header2, text="Marsa Piza",
                 font=('Arial', 14, 'bold'),
                 bg='#34495e', fg='white').pack(pady=8)

        # ═══════════════════════════════════════════
        # عنوان ملخص الوردية
        # ═══════════════════════════════════════════
        section = tk.Frame(self.report_frame, bg='#ecf0f1', height=30)
        section.pack(fill='x')
        section.pack_propagate(False)
        tk.Label(section, text="ملخص الوردية",
                 font=('Arial', 14, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=8)

        # ═══════════════════════════════════════════
        # جدول المعلومات
        # ═══════════════════════════════════════════
        self.info_frame = tk.Frame(self.report_frame, bg='white')
        self.info_frame.pack(fill='x', padx=0, pady=10)
        self.add_info_row(self.info_frame, "اليوم والتاريخ", "---")

        # ✅ وقت الطباعة نُقل إلى التذييل

        # ═══════════════════════════════════════════
        # الإجماليات
        # ═══════════════════════════════════════════
        totals_frame = tk.Frame(self.report_frame, bg='#ecf0f1', height=15)
        totals_frame.pack(fill='x', pady=(15, 0))
        totals_frame.pack_propagate(False)

        self.totals_inner = tk.Frame(self.report_frame, bg='white')
        self.totals_inner.pack(fill='x', padx=0, pady=5)

        # الصفوف الأساسية (سيتم إعادة بنائها في load_report)
        self.add_total_row(self.totals_inner, "اجمالي الفواتير", "—", '#fef9e7')
        self.add_total_row(self.totals_inner, "اجمالي المصروفات", "—", '#fdedec')
        self.add_total_row(self.totals_inner, "صافي الايراد (ج.م)", "—", '#eafaf1', bold=True)

        # ═══════════════════════════════════════════
        # ✅ إطار 1: بنود المصروفات النقدية
        # ═══════════════════════════════════════════
        items_section = tk.Frame(self.report_frame, bg='#ecf0f1')
        items_section.pack(fill='both', expand=True, padx=10, pady=10)

        items_header = tk.Frame(items_section, bg='#8B4513', height=40)
        items_header.pack(fill='x')
        items_header.pack_propagate(False)
        tk.Label(items_header,
                 text="🧾 بنود المصروفات النقدية",
                 font=('Arial', 13, 'bold'),
                 bg='#8B4513', fg='white').pack(pady=8)

        items_body = tk.Frame(items_section, bg='white',
                              relief='solid', bd=2)
        items_body.pack(fill='both', expand=True, pady=(2, 0))

        cols = ('description', 'value')
        self.items_tree = ttk.Treeview(items_body, columns=cols,
                                       show='headings', height=4)
        self.items_tree.heading('description',
                                text='📋 البيان',
                                anchor='e')
        self.items_tree.heading('value',
                                text='💰 المبلغ (ج.م)',
                                anchor='center')
        self.items_tree.column('description', width=500, anchor='e')
        self.items_tree.column('value', width=180, anchor='center')

        self.items_tree.tag_configure('odd_row', background='#f8f9fa')
        self.items_tree.tag_configure('even_row', background='white')

        items_scroll = ttk.Scrollbar(items_body, orient='vertical',
                                     command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=items_scroll.set)
        self.items_tree.pack(side='left', fill='both', expand=True)
        items_scroll.pack(side='right', fill='y')

        items_total_bar = tk.Frame(items_section, bg='#1a5276',
                                   relief='solid', bd=2)
        items_total_bar.pack(fill='x', pady=(4, 0))

        tk.Label(items_total_bar, text="📊 عدد البنود:",
                 font=('Arial', 11, 'bold'),
                 bg='#1a5276', fg='#95a5a6').pack(side='right', padx=15, pady=8)

        self.items_count_label = tk.Label(items_total_bar,
                                          text="0",
                                          font=('Arial', 11, 'bold'),
                                          bg='#1a5276', fg='white')
        self.items_count_label.pack(side='right', padx=5, pady=8)

        tk.Label(items_total_bar,
                 text="✅ إجمالي المصروفات النقدية:",
                 font=('Arial', 12, 'bold'),
                 bg='#1a5276', fg='#f1c40f').pack(side='right', padx=20, pady=8)

        self.items_total_label = tk.Label(items_total_bar,
                                          text="0.00 ج.م",
                                          font=('Arial', 14, 'bold'),
                                          bg='#1a5276', fg='#2ecc71')
        self.items_total_label.pack(side='left', padx=20, pady=8)

        # ═══════════════════════════════════════════
        # ✅ إطار 2: بنود مصاريف فودافون كاش
        # ═══════════════════════════════════════════
        voda_section = tk.Frame(self.report_frame, bg='#ecf0f1')
        voda_section.pack(fill='both', expand=True, padx=10, pady=(5, 10))

        voda_header = tk.Frame(voda_section, bg='#c0392b', height=40)
        voda_header.pack(fill='x')
        voda_header.pack_propagate(False)
        tk.Label(voda_header,
                 text="📱 بنود مصاريف فودافون كاش",
                 font=('Arial', 13, 'bold'),
                 bg='#c0392b', fg='white').pack(pady=8)

        voda_body = tk.Frame(voda_section, bg='white',
                             relief='solid', bd=2)
        voda_body.pack(fill='both', expand=True, pady=(2, 0))

        voda_cols = ('description', 'value')
        self.voda_tree = ttk.Treeview(voda_body, columns=voda_cols,
                                      show='headings', height=4)
        self.voda_tree.heading('description',
                               text='📋 البيان',
                               anchor='e')
        self.voda_tree.heading('value',
                               text='💰 المبلغ (ج.م)',
                               anchor='center')
        self.voda_tree.column('description', width=500, anchor='e')
        self.voda_tree.column('value', width=180, anchor='center')

        self.voda_tree.tag_configure('odd_row', background='#fdf2e9')
        self.voda_tree.tag_configure('even_row', background='white')

        voda_scroll = ttk.Scrollbar(voda_body, orient='vertical',
                                    command=self.voda_tree.yview)
        self.voda_tree.configure(yscrollcommand=voda_scroll.set)
        self.voda_tree.pack(side='left', fill='both', expand=True)
        voda_scroll.pack(side='right', fill='y')

        voda_total_bar = tk.Frame(voda_section, bg='#7b241c',
                                  relief='solid', bd=2)
        voda_total_bar.pack(fill='x', pady=(4, 0))

        tk.Label(voda_total_bar, text="📊 عدد البنود:",
                 font=('Arial', 11, 'bold'),
                 bg='#7b241c', fg='#f5b7b1').pack(side='right', padx=15, pady=8)

        self.voda_count_label = tk.Label(voda_total_bar,
                                         text="0",
                                         font=('Arial', 11, 'bold'),
                                         bg='#7b241c', fg='white')
        self.voda_count_label.pack(side='right', padx=5, pady=8)

        tk.Label(voda_total_bar,
                 text="✅ إجمالي مصاريف فودافون:",
                 font=('Arial', 12, 'bold'),
                 bg='#7b241c', fg='#f1c40f').pack(side='right', padx=20, pady=8)

        self.voda_total_label = tk.Label(voda_total_bar,
                                         text="0.00 ج.م",
                                         font=('Arial', 14, 'bold'),
                                         bg='#7b241c', fg='#f5b7b1')
        self.voda_total_label.pack(side='left', padx=20, pady=8)

        # ═══════════════════════════════════════════
        # ✅ التذييل (وقت الطباعة)
        # ═══════════════════════════════════════════
        footer_frame = tk.Frame(self.report_frame, bg='#2c3e50', height=50)
        footer_frame.pack(fill='x', pady=(15, 5))
        footer_frame.pack_propagate(False)

        self.lbl_print_time_footer = tk.Label(
            footer_frame,
            text="🕐 وقت الطباعة: ---",
            font=('Arial', 12, 'bold'),
            bg='#2c3e50', fg='#f1c40f'
        )
        self.lbl_print_time_footer.pack(pady=15)

    def add_info_row(self, parent, label, value):
        """إضافة صف معلومات (label + value)"""
        row = tk.Frame(parent, bg='white', relief='solid', bd=1)
        row.pack(fill='x', pady=1)

        # التسمية يمين
        tk.Label(row, text=label, font=('Arial', 12, 'bold'),
                 bg='white', fg='#2c3e50', width=20,
                 anchor='e').pack(side='right', padx=15, pady=8)

        # القيمة
        lbl = tk.Label(row, text=value, font=('Arial', 12),
                       bg='white', fg='#34495e', anchor='w')
        lbl.pack(side='left', padx=15, pady=8)

        return lbl

    def add_total_row(self, parent, label, value, bg_color, bold=False):
        row = tk.Frame(parent, bg=bg_color, relief='solid', bd=1)
        row.pack(fill='x', pady=1, padx=3)
        font_style = ('Arial', 13, 'bold') if bold else ('Arial', 11, 'bold')

        tk.Label(row, text=label, font=font_style,
                 bg=bg_color, fg='#2c3e50', width=22,
                 anchor='e').pack(side='right', padx=15, pady=4)  # ✅ 8→4

        value_color = '#c0392b' if bold else '#1a5276'
        lbl = tk.Label(row, text=value,
                       font=('Arial', 13, 'bold') if bold else ('Arial', 12),
                       bg=bg_color, fg=value_color,
                       anchor='w')
        lbl.pack(side='left', padx=15, pady=4)  # ✅ 8→4
        return lbl

    def set_today(self):
        self.date_entry.set_date(datetime.now())
        self.load_report()

    def set_yesterday(self):
        from datetime import timedelta
        yesterday = datetime.now() - timedelta(days=1)
        self.date_entry.set_date(yesterday)
        self.load_report()

    def load_report(self):
        """تحميل بيانات اليوم المحدد"""
        report_date = self.date_entry.get().strip()
        if not report_date:
            messagebox.showwarning("تنبيه", "اختر تاريخاً")
            return

        conn = self.get_connection()
        c = conn.cursor()

        # ═══════════════════════════════════════════
        # 1. جلب مبيعات اليوم + تحويل العملات
        # ═══════════════════════════════════════════
        c.execute("""SELECT 
                       COALESCE(SUM(pound), 0),
                       COALESCE(SUM(visa), 0),
                       COALESCE(SUM(euro), 0),
                       COALESCE(SUM(dollar), 0),
                       COALESCE(SUM(vodafone), 0),
                       COALESCE(SUM(euro * euro_rate), 0),
                       COALESCE(SUM(dollar * dollar_rate), 0)
                     FROM sales WHERE date = ?""", (report_date,))
        row = c.fetchone()
        total_pound = row[0] or 0
        total_visa = row[1] or 0
        total_euro = row[2] or 0
        total_dollar = row[3] or 0
        total_vodafone_sales = row[4] or 0
        total_euro_egp = row[5] or 0
        total_dollar_egp = row[6] or 0

        # إجمالي الفواتير = كل شيء + العملات محوّلة
        total_invoices = (total_pound + total_visa + total_vodafone_sales +
                          total_euro_egp + total_dollar_egp)

        net_euro = total_euro
        net_dollar = total_dollar

        self.show_euro_row = net_euro > 0
        self.show_dollar_row = net_dollar > 0

        # ═══════════════════════════════════════════
        # 2. المصروفات
        # ═══════════════════════════════════════════
        # مصروفات نقدية
        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids 
                     WHERE date = ?
                     AND COALESCE(payment_method, 'نقدي') = 'نقدي'""",
                  (report_date,))
        total_cash_expenses = c.fetchone()[0] or 0

        # مصروفات فودافون (منفصلة)
        c.execute("""SELECT COALESCE(SUM(value), 0) FROM paids 
                     WHERE date = ?
                     AND payment_method = 'فودافون كاش'""",
                  (report_date,))
        total_vodafone_expenses = c.fetchone()[0] or 0

        # صافي الإيراد = نقدي فقط
        net_revenue = total_pound - total_cash_expenses

        # ═══════════════════════════════════════════
        # 3. بنود المصروفات (نقدي + فودافون)
        # ═══════════════════════════════════════════
        # 3-أ. بنود المصروفات النقدية
        c.execute("""SELECT main_name, sub_name, description, value
                     FROM paids 
                     WHERE date = ?
                     AND COALESCE(payment_method, 'نقدي') = 'نقدي'
                     ORDER BY id""", (report_date,))
        raw_items = c.fetchall()

        expense_items = []
        for main_name, sub_name, description, value in raw_items:
            parts = []
            if main_name: parts.append(str(main_name))
            if sub_name: parts.append(str(sub_name))
            if description: parts.append(str(description))
            full_desc = ' - '.join(parts) if parts else '(بدون بيان)'
            expense_items.append((full_desc, value))

        # 3-ب. بنود مصاريف فودافون كاش
        c.execute("""SELECT main_name, sub_name, description, value
                     FROM paids 
                     WHERE date = ?
                     AND payment_method = 'فودافون كاش'
                     ORDER BY id""", (report_date,))
        raw_voda_items = c.fetchall()

        voda_expense_items = []
        for main_name, sub_name, description, value in raw_voda_items:
            parts = []
            if main_name: parts.append(str(main_name))
            if sub_name: parts.append(str(sub_name))
            if description: parts.append(str(description))
            full_desc = ' - '.join(parts) if parts else '(بدون بيان)'
            voda_expense_items.append((full_desc, value))

        conn.close()

        # ═══════════════════════════════════════════
        # 4. إعادة بناء الواجهة
        # ═══════════════════════════════════════════
        self.build_report_ui()

        try:
            date_obj = datetime.strptime(report_date, '%Y-%m-%d')
            day_name = self.get_day_name_ar(report_date)
            date_display = f"{day_name} - {date_obj.strftime('%d-%m-%Y')}"
        except:
            date_display = report_date

        print_time = datetime.now().strftime('%d/%m/%Y - %I:%M%p')

        # تحديث صفوف المعلومات
        # تحديث صف التاريخ فقط (وقت الطباعة انتقل للتذييل)
        info_rows = self.info_frame.winfo_children()
        if info_rows:
            for child in info_rows[0].winfo_children():
                if isinstance(child, tk.Label) and child.cget('anchor') == 'w':
                    child.config(text=date_display)

        # ✅ تحديث وقت الطباعة في التذييل
        if hasattr(self, 'lbl_print_time_footer'):
            self.lbl_print_time_footer.config(
                text=f"🕐 وقت الطباعة: {print_time}"
            )

        # ═══════════════════════════════════════════
        # 5. إعادة بناء الإجماليات
        # ═══════════════════════════════════════════
        for child in self.totals_inner.winfo_children():
            child.destroy()

        # 💰 قسم المبيعات
        self._add_section_title(self.totals_inner, "💰 المبيعات")

        self.add_total_row(self.totals_inner, "💵 نقدي (L.E)",
                           f"{total_pound:,.2f}", '#eafaf1')

        if total_visa > 0:
            self.add_total_row(self.totals_inner, "💳 فيزا",
                               f"{total_visa:,.2f}", '#f4ecf7')

        if total_euro > 0:
            self.add_total_row(self.totals_inner, "💶 يورو (€)",
                               f"{total_euro:,.2f} €  =  {total_euro_egp:,.2f} ج.م",
                               '#e8f4f8')

        if total_dollar > 0:
            self.add_total_row(self.totals_inner, "💵 دولار ($)",
                               f"{total_dollar:,.2f} $  =  {total_dollar_egp:,.2f} ج.م",
                               '#f0f8e8')

        if total_vodafone_sales > 0:
            self.add_total_row(self.totals_inner, "📱 فودافون كاش",
                               f"{total_vodafone_sales:,.2f}", '#fdf2e9')

        self.add_total_row(self.totals_inner,
                           "🔷 إجمالي الفواتير",
                           f"{total_invoices:,.2f}",
                           '#fef9e7', bold=True)

        # 💸 قسم المصروفات
        self._add_section_title(self.totals_inner, "💸 المصروفات")

        self.add_total_row(self.totals_inner,
                           "🧾 مصروفات نقدية",
                           f"{total_cash_expenses:,.2f}",
                           '#fdedec')

        if total_vodafone_expenses > 0:
            self.add_total_row(self.totals_inner,
                               "📱 مصاريف فودافون (منفصل)",
                               f"{total_vodafone_expenses:,.2f}",
                               '#fdf2e9')

        # 💰 صافي الإيراد
        # 💰 صافي الإيراد — قسم منفصل بارز
        self._add_net_revenue_section(self.totals_inner, net_revenue)

        # ═══════════════════════════════════════════
        # 6. جدول بنود المصروفات النقدية
        # ═══════════════════════════════════════════
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)

        if expense_items:
            for i, (desc, value) in enumerate(expense_items):
                tag = 'odd_row' if i % 2 == 0 else 'even_row'
                self.items_tree.insert('', 'end',
                                       values=(desc or '', f"{value:,.2f}"),
                                       tags=(tag,))
        else:
            self.items_tree.insert('', 'end',
                                   values=('لا توجد مصروفات نقدية لهذا اليوم', ''),
                                   tags=('odd_row',))

        items_total = sum(v for _, v in expense_items)
        self.items_total_label.config(text=f"{items_total:,.2f} ج.م")
        self.items_count_label.config(text=str(len(expense_items)))

        # ═══════════════════════════════════════════
        # 7. جدول بنود مصاريف فودافون كاش
        # ═══════════════════════════════════════════
        for item in self.voda_tree.get_children():
            self.voda_tree.delete(item)

        if voda_expense_items:
            for i, (desc, value) in enumerate(voda_expense_items):
                tag = 'odd_row' if i % 2 == 0 else 'even_row'
                self.voda_tree.insert('', 'end',
                                      values=(desc or '', f"{value:,.2f}"),
                                      tags=(tag,))
        else:
            self.voda_tree.insert('', 'end',
                                  values=('لا توجد مصاريف فودافون لهذا اليوم', ''),
                                  tags=('odd_row',))

        voda_total = sum(v for _, v in voda_expense_items)
        self.voda_total_label.config(text=f"{voda_total:,.2f} ج.م")
        self.voda_count_label.config(text=str(len(voda_expense_items)))

        # ═══════════════════════════════════════════
        # 8. حفظ البيانات للتصدير
        # ═══════════════════════════════════════════
        self.last_data = {
            'report_date': date_display,
            'print_time': print_time,
            'total_pound': total_pound,
            'total_visa': total_visa,
            'total_euro': total_euro,
            'total_dollar': total_dollar,
            'total_vodafone_sales': total_vodafone_sales,
            'total_euro_egp': total_euro_egp,
            'total_dollar_egp': total_dollar_egp,
            'total_invoices': total_invoices,
            'total_cash_expenses': total_cash_expenses,
            'total_vodafone_expenses': total_vodafone_expenses,
            'net_revenue': net_revenue,
            'net_euro': total_euro,
            'net_dollar': total_dollar,
            'expense_items': expense_items,
            'voda_expense_items': voda_expense_items,
        }

        # ✅ تحديث منطقة التمرير بعد إعادة البناء
        try:
            self.report_frame.update_idletasks()
            self._canvas.configure(scrollregion=self._canvas.bbox('all'))
            # ✅ إعادة تعيين موضع التمرير للأعلى عند كل استعلام
            self._canvas.yview_moveto(0)

        except Exception as e:
            print(f"⚠️ scroll update: {e}")

    def _add_section_title(self, parent, title):
        """إضافة عنوان قسم بين الصفوف"""
        section = tk.Frame(parent, bg='#8B4513', height=32)
        section.pack(fill='x', pady=(6, 2))
        section.pack_propagate(False)
        tk.Label(section, text=title,
                 font=('Arial', 12, 'bold'),
                 bg='#8B4513', fg='white').pack(pady=5)

    def _add_net_revenue_section(self, parent, value):
        """إضافة قسم صافي الإيراد بتصميم بارز ومميز"""
        # ─── شريط العنوان ───
        header = tk.Frame(parent, bg='#1a5276', height=32)
        header.pack(fill='x', pady=(10, 0))
        header.pack_propagate(False)
        tk.Label(header, text="💰 صافي الإيراد النقدي",
                 font=('Arial', 12, 'bold'),
                 bg='#1a5276', fg='white').pack(pady=5)

        # ─── شريط القيمة ───
        value_bar = tk.Frame(parent, bg='#d5f4e6', height=60,
                             relief='solid', bd=2,
                             highlightbackground='#1e8449',
                             highlightthickness=2)
        value_bar.pack(fill='x', pady=(0, 5))
        value_bar.pack_propagate(False)

        color = '#1e8449' if value >= 0 else '#c0392b'
        bg_color = '#d5f4e6' if value >= 0 else '#fdedec'

        value_bar.config(bg=bg_color)

        tk.Label(value_bar, text="💰",
                 font=('Arial', 20, 'bold'),
                 bg=bg_color, fg=color).pack(side='right', padx=15)

        tk.Label(value_bar, text=f"{value:,.2f} ج.م",
                 font=('Arial', 18, 'bold'),
                 bg=bg_color, fg=color).pack(side='left', padx=20)

        tk.Label(value_bar, text="صافي الإيراد النقدي",
                 font=('Arial', 14, 'bold'),
                 bg=bg_color, fg='#1a5276').pack(side='right', padx=10)

    def export_image(self):
        """تصدير اليومية إلى صورة PNG"""
        from tkinter import filedialog

        if not hasattr(self, 'last_data') or not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض اليومية أولاً")
            return

        data = self.last_data
        safe_date = data['report_date'].replace('/', '-').replace(' ', '_')

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"يومية_{safe_date}.png",
            title="حفظ اليومية صورة"
        )
        if not file_path:
            return

        success = export_daily_report_to_image(
            file_path=file_path,
            report_date=data['report_date'],
            print_time=data['print_time'],
            total_pound=data.get('total_pound', 0),
            total_visa=data.get('total_visa', 0),
            total_euro=data.get('total_euro', 0),
            total_dollar=data.get('total_dollar', 0),
            total_vodafone_sales=data.get('total_vodafone_sales', 0),
            total_euro_egp=data.get('total_euro_egp', 0),
            total_dollar_egp=data.get('total_dollar_egp', 0),
            total_invoices=data.get('total_invoices', 0),
            total_cash_expenses=data.get('total_cash_expenses', 0),
            total_vodafone_expenses=data.get('total_vodafone_expenses', 0),
            net_revenue=data.get('net_revenue', 0),
            expense_items=data.get('expense_items', []),
            voda_expense_items=data.get('voda_expense_items', []),
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم حفظ اليومية:\n{file_path}\n\nفتح الصورة الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass
    def export_excel(self):
        """تصدير اليومية إلى Excel"""
        if not hasattr(self, 'last_data') or not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض اليومية أولاً")
            return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import (Font, Alignment, PatternFill,
                                         Border, Side)
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        data = self.last_data
        safe_date = data['report_date'].replace('/', '-').replace(' ', '_')

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"يومية_{safe_date}.xlsx",
            title="حفظ اليومية Excel"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "اليومية"
            ws.sheet_view.rightToLeft = True

            # ─── عرض الأعمدة ───
            col_widths = {'A': 28, 'B': 22, 'C': 18, 'D': 18}
            for col, w in col_widths.items():
                ws.column_dimensions[col].width = w

            def style_cell(cell, bg='FFFFFF', fg='000000',
                           bold=False, size=11, align='center'):
                cell.font = Font(name='Arial', size=size, bold=bold,
                                 color=fg)
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal=align,
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            row_num = 1

            # ═════════ عنوان الكافيتريا ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1,
                        value="MARSA PIZZA - مطعم حديقة العائلات")
            style_cell(c, bg='2C3E50', fg='FFFFFF', bold=True, size=16)
            ws.row_dimensions[row_num].height = 35
            row_num += 1

            # ═════════ ملخص الوردية ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1, value="ملخص الوردية")
            style_cell(c, bg='8B4513', fg='FFFFFF', bold=True, size=14)
            ws.row_dimensions[row_num].height = 28
            row_num += 1

            # التاريخ + وقت الطباعة
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1,
                        value=f"التاريخ: {data['report_date']}")
            style_cell(c, bg='F8F9FA', fg='2C3E50', bold=True, size=12,
                       align='right')
            row_num += 1

            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1,
                        value=f"وقت الطباعة: {data['print_time']}")
            style_cell(c, bg='F8F9FA', fg='2C3E50', bold=True, size=12,
                       align='right')
            row_num += 2

            # ═════════ قسم المبيعات ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1, value="💰 المبيعات")
            style_cell(c, bg='8B4513', fg='FFFFFF', bold=True, size=13)
            ws.row_dimensions[row_num].height = 25
            row_num += 1

            sales_rows = [
                ("💵 نقدي (L.E)", f"{data.get('total_pound', 0):,.2f}",
                 'EAF7F1'),
                ("💳 فيزا", f"{data.get('total_visa', 0):,.2f}",
                 'F4ECF7'),
            ]
            if data.get('total_euro', 0) > 0:
                sales_rows.append(
                    ("💶 يورو (€)",
                     f"{data['total_euro']:,.2f} € = "
                     f"{data['total_euro_egp']:,.2f} ج.م",
                     'E8F4F8'))
            if data.get('total_dollar', 0) > 0:
                sales_rows.append(
                    ("💵 دولار ($)",
                     f"{data['total_dollar']:,.2f} $ = "
                     f"{data['total_dollar_egp']:,.2f} ج.م",
                     'F0F8E8'))
            if data.get('total_vodafone_sales', 0) > 0:
                sales_rows.append(
                    ("📱 فودافون كاش",
                     f"{data['total_vodafone_sales']:,.2f}", 'FDF2E9'))

            for label, value, bg in sales_rows:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=3)
                c = ws.cell(row=row_num, column=1, value=label)
                style_cell(c, bg=bg, fg='2C3E50', bold=True, align='right')

                c = ws.cell(row=row_num, column=4, value=value)
                style_cell(c, bg=bg, fg='1A5276', bold=True)

                row_num += 1

            # إجمالي الفواتير
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=3)
            c = ws.cell(row=row_num, column=1, value="🔷 إجمالي الفواتير")
            style_cell(c, bg='FEF9E7', fg='8B4513', bold=True, align='right')
            c = ws.cell(row=row_num, column=4,
                        value=f"{data.get('total_invoices', 0):,.2f}")
            style_cell(c, bg='FEF9E7', fg='C0392B', bold=True)
            row_num += 2

            # ═════════ قسم المصروفات ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1, value="💸 المصروفات")
            style_cell(c, bg='8B4513', fg='FFFFFF', bold=True, size=13)
            ws.row_dimensions[row_num].height = 25
            row_num += 1

            # مصروفات نقدية
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=3)
            c = ws.cell(row=row_num, column=1, value="🧾 مصروفات نقدية")
            style_cell(c, bg='FDEDEC', fg='2C3E50', bold=True, align='right')
            c = ws.cell(row=row_num, column=4,
                        value=f"{data.get('total_cash_expenses', 0):,.2f}")
            style_cell(c, bg='FDEDEC', fg='1A5276', bold=True)
            row_num += 1

            if data.get('total_vodafone_expenses', 0) > 0:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=3)
                c = ws.cell(row=row_num, column=1,
                            value="📱 مصاريف فودافون (منفصل)")
                style_cell(c, bg='FDF2E9', fg='2C3E50', bold=True,
                           align='right')
                c = ws.cell(row=row_num, column=4,
                            value=f"{data['total_vodafone_expenses']:,.2f}")
                style_cell(c, bg='FDF2E9', fg='1A5276', bold=True)
                row_num += 1

            row_num += 1

            # ═════════ صافي الإيراد النقدي ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1, value="💰 صافي الإيراد النقدي")
            style_cell(c, bg='1A5276', fg='FFFFFF', bold=True, size=13)
            ws.row_dimensions[row_num].height = 25
            row_num += 1

            net_revenue = data.get('net_revenue', 0)
            net_bg = 'D5F4E6' if net_revenue >= 0 else 'FDEDEC'
            net_fg = '1E8449' if net_revenue >= 0 else 'C0392B'

            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=3)
            c = ws.cell(row=row_num, column=1, value="💰 صافي الإيراد النقدي")
            style_cell(c, bg=net_bg, fg=net_fg, bold=True, size=12,
                       align='right')
            c = ws.cell(row=row_num, column=4,
                        value=f"{net_revenue:,.2f} ج.م")
            style_cell(c, bg=net_bg, fg=net_fg, bold=True, size=12)
            ws.row_dimensions[row_num].height = 30
            row_num += 2

            # ═════════ بنود المصروفات النقدية ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1,
                        value="🧾 بنود المصروفات النقدية")
            style_cell(c, bg='8B4513', fg='FFFFFF', bold=True, size=13)
            ws.row_dimensions[row_num].height = 25
            row_num += 1

            # الرؤوس
            c = ws.cell(row=row_num, column=1, value="البيان")
            style_cell(c, bg='34495E', fg='FFFFFF', bold=True)
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=3)
            c = ws.cell(row=row_num, column=4, value="المبلغ (ج.م)")
            style_cell(c, bg='34495E', fg='FFFFFF', bold=True)
            row_num += 1

            expense_items = data.get('expense_items', [])
            items_total = 0
            if expense_items:
                for desc, value in expense_items:
                    ws.merge_cells(start_row=row_num, start_column=1,
                                   end_row=row_num, end_column=3)
                    c = ws.cell(row=row_num, column=1,
                                value=desc or 'بدون بيان')
                    style_cell(c, bg='FFFFFF', fg='2C3E50', align='right')
                    c = ws.cell(row=row_num, column=4, value=value)
                    style_cell(c, bg='FFFFFF', fg='C0392B', bold=True)
                    items_total += value
                    row_num += 1
            else:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=4)
                c = ws.cell(row=row_num, column=1,
                            value="لا توجد مصروفات نقدية لهذا اليوم")
                style_cell(c, bg='F8F9FA', fg='95A5A6', align='center')
                row_num += 1

            # إجمالي البنود النقدية
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=3)
            c = ws.cell(row=row_num, column=1,
                        value="✅ إجمالي المصروفات النقدية")
            style_cell(c, bg='1A5276', fg='F1C40F', bold=True, align='right')
            c = ws.cell(row=row_num, column=4,
                        value=f"{items_total:,.2f} ج.م")
            style_cell(c, bg='1A5276', fg='2ECC71', bold=True)
            row_num += 2

            # ═════════ بنود مصاريف فودافون كاش ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1,
                        value="📱 بنود مصاريف فودافون كاش")
            style_cell(c, bg='C0392B', fg='FFFFFF', bold=True, size=13)
            ws.row_dimensions[row_num].height = 25
            row_num += 1

            # الرؤوس
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=3)
            c = ws.cell(row=row_num, column=1, value="البيان")
            style_cell(c, bg='34495E', fg='FFFFFF', bold=True)
            c = ws.cell(row=row_num, column=4, value="المبلغ (ج.م)")
            style_cell(c, bg='34495E', fg='FFFFFF', bold=True)
            row_num += 1

            voda_items = data.get('voda_expense_items', [])
            voda_total = 0
            if voda_items:
                for desc, value in voda_items:
                    ws.merge_cells(start_row=row_num, start_column=1,
                                   end_row=row_num, end_column=3)
                    c = ws.cell(row=row_num, column=1,
                                value=desc or 'بدون بيان')
                    style_cell(c, bg='FDF2E9', fg='2C3E50', align='right')
                    c = ws.cell(row=row_num, column=4, value=value)
                    style_cell(c, bg='FDF2E9', fg='C0392B', bold=True)
                    voda_total += value
                    row_num += 1
            else:
                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=4)
                c = ws.cell(row=row_num, column=1,
                            value="لا توجد مصاريف فودافون لهذا اليوم")
                style_cell(c, bg='FDF2E9', fg='95A5A6', align='center')
                row_num += 1

            # إجمالي فودافون
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=3)
            c = ws.cell(row=row_num, column=1,
                        value="✅ إجمالي مصاريف فودافون")
            style_cell(c, bg='7B241C', fg='F1C40F', bold=True, align='right')
            c = ws.cell(row=row_num, column=4,
                        value=f"{voda_total:,.2f} ج.م")
            style_cell(c, bg='7B241C', fg='F5B7B1', bold=True)
            row_num += 2

            # ═════════ التذييل ═════════
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=4)
            c = ws.cell(row=row_num, column=1,
                        value="شركة الزهراء للبرمجيات - المحاسب/حسام الدين "
                              "سيد ربيع 01003796984")
            style_cell(c, bg='ECF0F1', fg='7F8C8D', bold=True, size=10)

            # ─── الحفظ ───
            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")
# ═══════════════════════════════════════════════════════════════════
# FINAL REPORT PAGE (تقرير ختامي)
# ═══════════════════════════════════════════════════════════════════
class FinalReportPage(tk.Frame):
    """التقرير الختامي - مبيعات ومشتريات ومصاريف وصافي"""

    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # 1) الشريط العلوي
        add_program_banner(self)

        # 2) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ═══════════════ العنوان ═══════════════
        tk.Label(content, text="📊 التقرير الختامي",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=12)

        # ═══════════════ فلاتر التاريخ ═══════════════
        filter_frame = tk.LabelFrame(content, text="اختر المدة",
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#34495e',
                                     padx=15, pady=12)
        filter_frame.pack(fill='x', padx=15, pady=(0, 10))

        # ─── من تاريخ ───
        tk.Label(filter_frame, text="من تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.date_from = DateEntry(filter_frame, width=13,
                                   date_pattern='yyyy-mm-dd',
                                   font=('Arial', 11))
        self.date_from.pack(side='right', padx=5)

        today = datetime.now()
        self.date_from.set_date(today.replace(day=1))

        # ─── إلى تاريخ ───
        tk.Label(filter_frame, text="إلى تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.date_to = DateEntry(filter_frame, width=13,
                                 date_pattern='yyyy-mm-dd',
                                 font=('Arial', 11))
        self.date_to.pack(side='right', padx=5)
        self.date_to.set_date(today)

        # ─── أزرار سريعة ───
        quick = tk.Frame(filter_frame, bg='#ecf0f1')
        quick.pack(side='right', padx=10)

        tk.Button(quick, text="اليوم", command=self.set_today,
                  bg='#3498db', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)

        tk.Button(quick, text="هذا الشهر", command=self.set_this_month,
                  bg='#16a085', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)

        tk.Button(quick, text="آخر 7 أيام", command=self.set_last_7_days,
                  bg='#f39c12', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)

        tk.Button(quick, text="آخر 30 يوم", command=self.set_last_30_days,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 10, 'bold'),
                  padx=8, pady=3).pack(side='right', padx=2)

        # ═══════════════ أزرار العمليات ═══════════════
        action_frame = tk.Frame(content, bg='#ecf0f1')
        action_frame.pack(fill='x', padx=15, pady=5)

        tk.Button(action_frame, text="🔍 عرض التقرير",
                  command=self.load_report,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=25, pady=10).pack(side='right', padx=5)

        tk.Button(action_frame, text="💾 تصدير Excel",
                  command=self.export_excel,
                  bg='#2980b9', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=25, pady=10).pack(side='right', padx=5)

        tk.Button(action_frame, text="🖼️ تصدير صورة",
                  command=self.export_image,
                  bg='#8e44ad', fg='white',
                  font=('Arial', 12, 'bold'),
                  padx=25, pady=10).pack(side='right', padx=5)

        # ═══════════════ جدول التقرير ═══════════════
        tree_frame = tk.LabelFrame(content, text="ملخص التقرير",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))

        columns = ('statement', 'cash', 'currency',
                   'visa', 'vodafone', 'total', 'percentage')

        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=8)

        # ─── العناوين ───
        self.tree.heading('statement',  text='البيان')
        self.tree.heading('cash',       text='نقدي')
        self.tree.heading('currency',   text='عملة')
        self.tree.heading('visa',       text='فيزا')
        self.tree.heading('vodafone',   text='فودافون كاش')
        self.tree.heading('total',      text='إجمالي')
        self.tree.heading('percentage', text='النسبة المئوية')

        # ─── العرض ───
        self.tree.column('statement',  width=180, anchor='center')
        self.tree.column('cash',       width=120, anchor='center')
        self.tree.column('currency',   width=100, anchor='center')
        self.tree.column('visa',       width=100, anchor='center')
        self.tree.column('vodafone',   width=120, anchor='center')
        self.tree.column('total',      width=130, anchor='center')
        self.tree.column('percentage', width=110, anchor='center')

        # ─── التنسيقات (3 صفوف بألوان مختلفة) ───
        self.tree.tag_configure('normal_row',
                                background='#ffffff',
                                foreground='#2c3e50',
                                font=('Arial', 12, 'bold'))
        self.tree.tag_configure('net_row',
                                background='#d5d8dc',
                                foreground='#2c3e50',
                                font=('Arial', 12, 'bold'))
        self.tree.tag_configure('final_row',
                                background='#2874a6',
                                foreground='#ffffff',
                                font=('Arial', 13, 'bold'))

        self.tree.pack(fill='both', expand=True)

        # ═══════════════ شريط الحالة ═══════════════
        self.status_bar = tk.Label(content, text="جاهز",
                                   font=('Arial', 10),
                                   bg='#2c3e50', fg='white',
                                   anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ─── تحميل أولي ───
        self.last_data = None
        self.load_report()

    # ═══════════════════════════════════════════════════════════
    # المساعدات
    # ═══════════════════════════════════════════════════════════
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def load_traders(self):
        """يُستدعى تلقائيًا عند فتح الصفحة"""
        pass

    # ─── أزرار التواريخ السريعة ───
    def set_today(self):
        today = datetime.now()
        self.date_from.set_date(today)
        self.date_to.set_date(today)
        self.load_report()

    def set_this_month(self):
        today = datetime.now()
        self.date_from.set_date(today.replace(day=1))
        self.date_to.set_date(today)
        self.load_report()

    def set_last_7_days(self):
        from datetime import timedelta
        today = datetime.now()
        self.date_from.set_date(today - timedelta(days=6))
        self.date_to.set_date(today)
        self.load_report()

    def set_last_30_days(self):
        from datetime import timedelta
        today = datetime.now()
        self.date_from.set_date(today - timedelta(days=29))
        self.date_to.set_date(today)
        self.load_report()

    # ═══════════════════════════════════════════════════════════
    # حساب التقرير
    # ═══════════════════════════════════════════════════════════
    def _calc_report(self, date_from, date_to):
        """حساب كل قيم التقرير"""
        conn = self.get_connection()
        c = conn.cursor()

        # ─── 1) المبيعات ───
        # نقدي (pound فقط) → نقدي
        # يورو + دولار (محوّلة) → عملة
        # فيزا → فيزا (عمود مستقل)
        # فودافون → فودافون كاش
        c.execute("""
            SELECT
                COALESCE(SUM(pound), 0),
                COALESCE(SUM(euro * euro_rate), 0)
                  + COALESCE(SUM(dollar * dollar_rate), 0),
                COALESCE(SUM(visa), 0),
                COALESCE(SUM(vodafone), 0)
            FROM sales
            WHERE date >= ? AND date <= ?
        """, (date_from, date_to))
        sales_cash, sales_currency, sales_visa, sales_voda = c.fetchone()

        # ─── 2) المشتريات (main_name = 'مشتريات') ───
        c.execute("""
            SELECT
                COALESCE(SUM(CASE
                    WHEN COALESCE(payment_method, 'نقدي') IN ('نقدي', 'الخزينة')
                    THEN value ELSE 0 END), 0),
                COALESCE(SUM(CASE
                    WHEN payment_method = 'فودافون كاش'
                    THEN value ELSE 0 END), 0)
            FROM paids
            WHERE main_name = 'مشتريات'
              AND date >= ? AND date <= ?
        """, (date_from, date_to))
        purch_cash, purch_voda = c.fetchone()

        # ─── 3) المصاريف (main_name != 'مشتريات') ───
        c.execute("""
            SELECT
                COALESCE(SUM(CASE
                    WHEN COALESCE(payment_method, 'نقدي') IN ('نقدي', 'الخزينة')
                    THEN value ELSE 0 END), 0),
                COALESCE(SUM(CASE
                    WHEN payment_method = 'فودافون كاش'
                    THEN value ELSE 0 END), 0)
            FROM paids
            WHERE main_name != 'مشتريات'
              AND date >= ? AND date <= ?
        """, (date_from, date_to))
        exp_cash, exp_voda = c.fetchone()

        conn.close()

        # ─── الحسابات ───
        # المشتريات: عملة = 0، فيزا = 0
        # المصاريف:   عملة = 0، فيزا = 0
        sales_total = sales_cash + sales_currency + sales_visa + sales_voda
        purch_total = purch_cash + purch_voda
        exp_total   = exp_cash + exp_voda

        # صافي المبيعات
        net_sales_cash = sales_cash - purch_cash
        net_sales_curr = sales_currency
        net_sales_visa = sales_visa
        net_sales_voda = sales_voda - purch_voda
        net_sales_total = sales_total - purch_total

        # الصافي
        final_cash = net_sales_cash - exp_cash
        final_curr = net_sales_curr
        final_visa = net_sales_visa
        final_voda = net_sales_voda - exp_voda
        final_total = net_sales_total - exp_total

        # ─── النسب المئوية ───
        def pct(value):
            if sales_total == 0:
                return 0
            return round((value / sales_total) * 100, 0)

        return {
            'sales': {
                'cash': sales_cash, 'currency': sales_currency,
                'visa': sales_visa, 'voda': sales_voda,
                'total': sales_total,
            },
            'purchases': {
                'cash': purch_cash, 'currency': 0,
                'visa': 0, 'voda': purch_voda,
                'total': purch_total,
                'pct': pct(purch_total),
            },
            'net_sales': {
                'cash': net_sales_cash, 'currency': net_sales_curr,
                'visa': net_sales_visa, 'voda': net_sales_voda,
                'total': net_sales_total,
            },
            'expenses': {
                'cash': exp_cash, 'currency': 0,
                'visa': 0, 'voda': exp_voda,
                'total': exp_total,
                'pct': pct(exp_total),
            },
            'final': {
                'cash': final_cash, 'currency': final_curr,
                'visa': final_visa, 'voda': final_voda,
                'total': final_total,
                'pct': pct(final_total),
            },
        }

    # ═══════════════════════════════════════════════════════════
    # عرض التقرير
    # ═══════════════════════════════════════════════════════════
    def load_report(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()

        if not date_from or not date_to:
            messagebox.showwarning("تنبيه", "اختر تاريخ البداية والنهاية")
            return

        try:
            d1 = datetime.strptime(date_from, '%Y-%m-%d')
            d2 = datetime.strptime(date_to, '%Y-%m-%d')
            if d2 < d1:
                messagebox.showerror("خطأ", "تاريخ النهاية يجب أن يكون بعد البداية")
                return
        except:
            messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة")
            return

        # ─── حساب ───
        r = self._calc_report(date_from, date_to)

        # ─── إدراج الصفوف ───
        def fmt(v):
            return f"{v:,.0f}"

        # 1) المبيعات
        self.tree.insert('', 'end', tags=('normal_row',), values=(
            'المبيعات',
            fmt(r['sales']['cash']),
            fmt(r['sales']['currency']),
            fmt(r['sales']['visa']),
            fmt(r['sales']['voda']),
            fmt(r['sales']['total']),
            '',
        ))

        # 2) مشتريات
        self.tree.insert('', 'end', tags=('normal_row',), values=(
            'مشتريات',
            fmt(r['purchases']['cash']),
            fmt(r['purchases']['currency']),
            fmt(r['purchases']['visa']),
            fmt(r['purchases']['voda']),
            fmt(r['purchases']['total']),
            f"{int(r['purchases']['pct'])}",
        ))

        # 3) صافي المبيعات (رمادي)
        self.tree.insert('', 'end', tags=('net_row',), values=(
            'صافي المبيعات',
            fmt(r['net_sales']['cash']),
            fmt(r['net_sales']['currency']),
            fmt(r['net_sales']['visa']),
            fmt(r['net_sales']['voda']),
            fmt(r['net_sales']['total']),
            '',
        ))

        # 4) مصاريف
        self.tree.insert('', 'end', tags=('normal_row',), values=(
            'مصاريف',
            fmt(r['expenses']['cash']),
            fmt(r['expenses']['currency']),
            fmt(r['expenses']['visa']),
            fmt(r['expenses']['voda']),
            fmt(r['expenses']['total']),
            f"{int(r['expenses']['pct'])}",
        ))

        # 5) الصافي (أزرق)
        self.tree.insert('', 'end', tags=('final_row',), values=(
            'الصافي',
            fmt(r['final']['cash']),
            fmt(r['final']['currency']),
            fmt(r['final']['visa']),
            fmt(r['final']['voda']),
            fmt(r['final']['total']),
            f"{int(r['final']['pct'])}",
        ))

        # ─── الحفظ للتصدير ───
        self.last_data = {
            'date_from': date_from,
            'date_to': date_to,
            **r,
        }

        self.status_bar.config(
            text=f"التقرير من {date_from} إلى {date_to}")

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى Excel
    # ═══════════════════════════════════════════════════════════
    def export_excel(self):
        if not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض التقرير أولاً")
            return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import (Font, Alignment,
                                         PatternFill, Border, Side)
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog

        data = self.last_data
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"التقرير_الختامي_{data['date_from']}_{data['date_to']}.xlsx",
            title="حفظ التقرير الختامي Excel"
        )

        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "التقرير الختامي"
            ws.sheet_view.rightToLeft = True

            # ─── عرض الأعمدة ───
            for col, w in {'A': 20, 'B': 15, 'C': 13,
                           'D': 13, 'E': 15, 'F': 15, 'G': 15}.items():
                ws.column_dimensions[col].width = w

            # ─── العنوان ───
            ws.merge_cells('A1:G1')
            c = ws['A1']
            c.value = f"📊 التقرير الختامي - مطعم الحديقة ({data['date_from']} - {data['date_to']})"
            c.font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            c.fill = PatternFill("solid", fgColor="2C3E50")
            c.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # ─── الرؤوس ───
            headers = ['البيان', 'نقدي', 'عملة', 'فيزا',
                       'فودافون كاش', 'إجمالي', 'النسبة المئوية']
            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=col_idx, value=h)
                cell.font = Font(name='Arial', size=12, bold=True,
                                 color='FFFFFF')
                cell.fill = PatternFill("solid", fgColor="2874A6")
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
            ws.row_dimensions[2].height = 25

            # ─── البيانات ───
            def fmt(v):
                return int(v) if v == int(v) else v

            rows_data = [
                ('المبيعات',
                 fmt(data['sales']['cash']),
                 fmt(data['sales']['currency']),
                 fmt(data['sales']['visa']),
                 fmt(data['sales']['voda']),
                 fmt(data['sales']['total']),
                 '', '#FFFFFF'),
                ('مشتريات',
                 fmt(data['purchases']['cash']),
                 fmt(data['purchases']['currency']),
                 fmt(data['purchases']['visa']),
                 fmt(data['purchases']['voda']),
                 fmt(data['purchases']['total']),
                 f"{int(data['purchases']['pct'])}", '#FFFFFF'),
                ('صافي المبيعات',
                 fmt(data['net_sales']['cash']),
                 fmt(data['net_sales']['currency']),
                 fmt(data['net_sales']['visa']),
                 fmt(data['net_sales']['voda']),
                 fmt(data['net_sales']['total']),
                 '', '#D5D8DC'),
                ('مصاريف',
                 fmt(data['expenses']['cash']),
                 fmt(data['expenses']['currency']),
                 fmt(data['expenses']['visa']),
                 fmt(data['expenses']['voda']),
                 fmt(data['expenses']['total']),
                 f"{int(data['expenses']['pct'])}", '#FFFFFF'),
                ('الصافي',
                 fmt(data['final']['cash']),
                 fmt(data['final']['currency']),
                 fmt(data['final']['visa']),
                 fmt(data['final']['voda']),
                 fmt(data['final']['total']),
                 f"{int(data['final']['pct'])}", '#2874A6'),
            ]

            row_num = 3
            for row_data in rows_data:
                bg = row_data[-1]
                text_color = 'FFFFFF' if bg == '#2874A6' else '2C3E50'
                for col_idx, val in enumerate(row_data[:-1], start=1):
                    cell = ws.cell(row=row_num, column=col_idx, value=val)
                    cell.font = Font(name='Arial', size=12, bold=True,
                                     color=text_color)
                    cell.fill = PatternFill("solid", fgColor=bg.lstrip('#'))
                    cell.alignment = Alignment(horizontal='center',
                                               vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin'))
                ws.row_dimensions[row_num].height = 28
                row_num += 1

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")

            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    # ═══════════════════════════════════════════════════════════
    # التصدير إلى صورة
    # ═══════════════════════════════════════════════════════════
    def export_image(self):
        if not self.last_data:
            messagebox.showwarning("تنبيه", "اعرض التقرير أولاً")
            return

        from tkinter import filedialog

        data = self.last_data
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"التقرير_الختامي_{data['date_from']}_{data['date_to']}.png",
            title="حفظ التقرير الختامي صورة"
        )

        if not file_path:
            return

        # ─── جمع الصفوف من الجدول ───
        columns = ['البيان', 'نقدي', 'عملة', 'فيزا',
                   'فودافون كاش', 'إجمالي', 'النسبة المئوية']
        rows_data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            rows_data.append(values)

        success = export_table_to_image(
            file_path=file_path,
            title="📊 التقرير الختامي - مطعم الحديقة",
            subtitle=f"من {data['date_from']} إلى {data['date_to']}",
            columns=columns,
            rows=rows_data,
            summary_items=None,
            logo_path=COMPANY_LOGO_PATH
        )

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass
# ═══════════════════════════════════════════════════════════════════
# REPORTS PAGE (3 تبويبات: الإجمالي + التفصيلي + المعاملات التفصيلية)
# ═══════════════════════════════════════════════════════════════════
class DetailedReportPage(tk.Frame):
    """
    1) التقرير الإجمالي: كل الأبواب الرئيسية + تفكيك "مصاريف" إلى بنودها
    2) التقرير التفصيلي: اختر باباً → تظهر بنوده الفرعية
    3) المعاملات التفصيلية: اختر رئيسي + فرعي → كل معاملة بتفاصيلها
    """

    # الأبواب الرئيسية التي تُفكَّك إلى بنودها في التقرير الإجمالي
    FLATTEN_MAINS = ['مصاريف']

    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # 1) الشريط العلوي
        add_program_banner(self)

        # 2) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ═══════════════ العنوان ═══════════════
        tk.Label(content, text="📊 التقارير",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(pady=12)

        # ═══════════════ فلتر التاريخ ═══════════════
        filter_frame = tk.LabelFrame(content, text="اختر المدة",
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#34495e',
                                     padx=15, pady=12)
        filter_frame.pack(fill='x', padx=15, pady=(0, 10))

        tk.Label(filter_frame, text="من تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.date_from = DateEntry(filter_frame, width=13,
                                   date_pattern='yyyy-mm-dd',
                                   font=('Arial', 11))
        self.date_from.pack(side='right', padx=5)
        today = datetime.now()
        self.date_from.set_date(today.replace(day=1))

        tk.Label(filter_frame, text="إلى تاريخ:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(
                 side='right', padx=5)

        self.date_to = DateEntry(filter_frame, width=13,
                                 date_pattern='yyyy-mm-dd',
                                 font=('Arial', 11))
        self.date_to.pack(side='right', padx=5)
        self.date_to.set_date(today)

        tk.Button(filter_frame, text="🔍 عرض الكل",
                  command=self.load_all_tabs,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='left', padx=5)

        # ═══════════════ نظام التبويبات (3 تبويبات) ═══════════════
        self.notebook = ttk.Notebook(content)
        self.notebook.pack(fill='both', expand=True, padx=15, pady=(0, 10))

        # ─── التبويب 1: التقرير الإجمالي ───
        self.tab_total = tk.Frame(self.notebook, bg='#ecf0f1')
        self.notebook.add(self.tab_total, text='  💰 التقرير الإجمالي  ')
        self._build_total_tab()

        # ─── التبويب 2: التقرير التفصيلي ───
        self.tab_detail = tk.Frame(self.notebook, bg='#ecf0f1')
        self.notebook.add(self.tab_detail, text='  📋 التقرير التفصيلي  ')
        self._build_detail_tab()

        # ─── التبويب 3: المعاملات التفصيلية ───
        self.tab_trans = tk.Frame(self.notebook, bg='#ecf0f1')
        self.notebook.add(self.tab_trans, text='  🧾 المعاملات التفصيلية  ')
        self._build_transactions_tab()

        # ─── بيانات التصدير ───
        self.total_data = None
        self.detail_data = None
        self.trans_data = None

        # ─── شريط الحالة ───
        self.status_bar = tk.Label(content, text="جاهز",
                                   font=('Arial', 10),
                                   bg='#2c3e50', fg='white',
                                   anchor='w', padx=15, pady=5)
        self.status_bar.pack(fill='x', side='bottom')

        # ─── تحميل أولي ───
        self.load_all_tabs()

    # ═══════════════════════════════════════════════════════════
    # بناء التبويب 1: التقرير الإجمالي
    # ═══════════════════════════════════════════════════════════
    def _build_total_tab(self):
        btn_frame = tk.Frame(self.tab_total, bg='#ecf0f1')
        btn_frame.pack(fill='x', padx=10, pady=8)

        tk.Button(btn_frame, text="💾 تصدير Excel",
                  command=lambda: self._export_excel('total'),
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        tk.Button(btn_frame, text="🖼️ تصدير صورة",
                  command=lambda: self._export_image('total'),
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=8).pack(side='right', padx=5)

        tree_frame = tk.LabelFrame(self.tab_total,
                                   text="التقرير الإجمالي لمصاريف التشغيل",
                                   font=('Arial', 11, 'bold'),
                                   bg='#ecf0f1', fg='#34495e',
                                   padx=5, pady=5)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        columns = ('name', 'amount')
        self.tree_total = ttk.Treeview(tree_frame, columns=columns,
                                       show='headings', height=18)

        self.tree_total.heading('name', text='البند')
        self.tree_total.heading('amount', text='المبلغ')

        self.tree_total.column('name',   width=400, anchor='e')
        self.tree_total.column('amount', width=200, anchor='center')

        self.tree_total.tag_configure('main_row',
                                      background='#fff3cd',
                                      foreground='#7d6608',
                                      font=('Arial', 12, 'bold'))
        self.tree_total.tag_configure('sub_row',
                                      background='#ffffff',
                                      foreground='#2c3e50',
                                      font=('Arial', 11))
        self.tree_total.tag_configure('total_row',
                                      background='#ffff00',
                                      foreground='#000000',
                                      font=('Arial', 13, 'bold'))

        self.tree_total.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree_total.yview)
        self.tree_total.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

    # ═══════════════════════════════════════════════════════════
    # بناء التبويب 2: التقرير التفصيلي
    # ═══════════════════════════════════════════════════════════
    def _build_detail_tab(self):
        top = tk.Frame(self.tab_detail, bg='#ecf0f1')
        top.pack(fill='x', padx=10, pady=8)

        tk.Label(top, text="اختر الباب:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1',
                 fg='#8B4513').pack(side='right', padx=5)

        self.detail_main_combo = ttk.Combobox(top, width=28,
                                              font=('Arial', 11),
                                              state='readonly')
        self.detail_main_combo.pack(side='right', padx=5)
        self.detail_main_combo.bind('<<ComboboxSelected>>',
                                    lambda e: self._load_detail_tab())

        tk.Button(top, text="🔍 عرض",
                  command=self._load_detail_tab,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='right', padx=5)

        tk.Button(top, text="💾 تصدير Excel",
                  command=lambda: self._export_excel('detail'),
                  bg='#2980b9', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='left', padx=5)

        tk.Button(top, text="🖼️ تصدير صورة",
                  command=lambda: self._export_image('detail'),
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='left', padx=5)

        title_bar = tk.Frame(self.tab_detail, bg='#27ae60', height=45)
        title_bar.pack(fill='x', padx=10, pady=(5, 0))
        title_bar.pack_propagate(False)

        tk.Label(title_bar,
                 text="التقرير التفصيلي للمصاريف",
                 font=('Arial', 14, 'bold'),
                 bg='#27ae60', fg='white').pack(pady=10)

        tree_frame = tk.Frame(self.tab_detail, bg='#ecf0f1',
                              relief='solid', bd=1)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        columns = ('name', 'amount')
        self.tree_detail = ttk.Treeview(tree_frame, columns=columns,
                                        show='headings', height=18)

        self.tree_detail.heading('name', text='البند')
        self.tree_detail.heading('amount', text='المبلغ')

        self.tree_detail.column('name',   width=400, anchor='e')
        self.tree_detail.column('amount', width=200, anchor='center')

        self.tree_detail.tag_configure('row',
                                       background='#ffffff',
                                       foreground='#2c3e50',
                                       font=('Arial', 11))
        self.tree_detail.tag_configure('total_row',
                                       background='#ffff00',
                                       foreground='#000000',
                                       font=('Arial', 13, 'bold'))

        self.tree_detail.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree_detail.yview)
        self.tree_detail.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

    # ═══════════════════════════════════════════════════════════
    # بناء التبويب 3: المعاملات التفصيلية
    # ═══════════════════════════════════════════════════════════
    def _build_transactions_tab(self):
        top = tk.Frame(self.tab_trans, bg='#ecf0f1')
        top.pack(fill='x', padx=10, pady=8)

        tk.Label(top, text="الباب الرئيسي:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1',
                 fg='#8B4513').pack(side='right', padx=5)

        self.trans_main_combo = ttk.Combobox(top, width=22,
                                             font=('Arial', 11),
                                             state='readonly')
        self.trans_main_combo.pack(side='right', padx=5)
        self.trans_main_combo.bind('<<ComboboxSelected>>',
                                   lambda e: self._on_trans_main_change())

        tk.Label(top, text="الباب الفرعي:",
                 font=('Arial', 12, 'bold'), bg='#ecf0f1',
                 fg='#8B4513').pack(side='right', padx=5)

        self.trans_sub_combo = ttk.Combobox(top, width=22,
                                            font=('Arial', 11),
                                            state='readonly')
        self.trans_sub_combo.pack(side='right', padx=5)
        self.trans_sub_combo.bind('<<ComboboxSelected>>',
                                  lambda e: self._load_trans_tab())

        tk.Button(top, text="🔍 عرض",
                  command=self._load_trans_tab,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='right', padx=5)

        tk.Button(top, text="💾 تصدير Excel",
                  command=lambda: self._export_excel('trans'),
                  bg='#2980b9', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='left', padx=5)

        tk.Button(top, text="🖼️ تصدير صورة",
                  command=lambda: self._export_image('trans'),
                  bg='#8e44ad', fg='white',
                  font=('Arial', 11, 'bold'),
                  padx=15, pady=5).pack(side='left', padx=5)

        title_bar = tk.Frame(self.tab_trans, bg='#27ae60', height=45)
        title_bar.pack(fill='x', padx=10, pady=(5, 0))
        title_bar.pack_propagate(False)

        tk.Label(title_bar,
                 text="🧾 المعاملات التفصيلية",
                 font=('Arial', 14, 'bold'),
                 bg='#27ae60', fg='white').pack(pady=10)

        tree_frame = tk.Frame(self.tab_trans, bg='#ecf0f1',
                              relief='solid', bd=1)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        columns = ('date', 'description', 'payment', 'amount')
        self.tree_trans = ttk.Treeview(tree_frame, columns=columns,
                                       show='headings', height=18)

        self.tree_trans.heading('date',        text='التاريخ')
        self.tree_trans.heading('description', text='الوصف')
        self.tree_trans.heading('payment',     text='طريقة الدفع')
        self.tree_trans.heading('amount',      text='المبلغ')

        self.tree_trans.column('date',        width=120, anchor='center')
        self.tree_trans.column('description', width=400, anchor='e')
        self.tree_trans.column('payment',     width=130, anchor='center')
        self.tree_trans.column('amount',      width=150, anchor='center')

        self.tree_trans.tag_configure('odd_row',
                                      background='#f8f9fa',
                                      foreground='#2c3e50',
                                      font=('Arial', 11))
        self.tree_trans.tag_configure('even_row',
                                      background='#ffffff',
                                      foreground='#2c3e50',
                                      font=('Arial', 11))
        self.tree_trans.tag_configure('total_row',
                                      background='#ffff00',
                                      foreground='#000000',
                                      font=('Arial', 13, 'bold'))

        self.tree_trans.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(tree_frame, orient='vertical',
                           command=self.tree_trans.yview)
        self.tree_trans.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')

    # ═══════════════════════════════════════════════════════════
    # المساعدات
    # ═══════════════════════════════════════════════════════════
    def get_connection(self):
        return sqlite3.connect(DB_NAME)

    def load_traders(self):
        pass

    def _check_dates(self):
        date_from = self.date_from.get().strip()
        date_to = self.date_to.get().strip()
        if not date_from or not date_to:
            messagebox.showwarning("تنبيه", "اختر تاريخ البداية والنهاية")
            return None, None
        try:
            d1 = datetime.strptime(date_from, '%Y-%m-%d')
            d2 = datetime.strptime(date_to, '%Y-%m-%d')
            if d2 < d1:
                messagebox.showerror("خطأ",
                                     "تاريخ النهاية يجب أن يكون بعد البداية")
                return None, None
        except:
            messagebox.showerror("خطأ", "صيغة التاريخ غير صحيحة")
            return None, None
        return date_from, date_to

    # ═══════════════════════════════════════════════════════════
    # التحميل
    # ═══════════════════════════════════════════════════════════
    def load_all_tabs(self):
        self._load_total_tab()
        self._refresh_detail_combo()
        self._load_detail_tab()
        self._refresh_trans_combo()
        self._load_trans_tab()
        self.status_bar.config(text="تم تحديث التقارير")

    # ═══════════════════════════════════════════════════════════
    # التبويب 1: التقرير الإجمالي
    # ═══════════════════════════════════════════════════════════
    def _load_total_tab(self):
        for item in self.tree_total.get_children():
            self.tree_total.delete(item)

        date_from, date_to = self._check_dates()
        if not date_from:
            return

        conn = self.get_connection()
        c = conn.cursor()

        c.execute("""
            SELECT main_name, COALESCE(SUM(value), 0)
            FROM paids
            WHERE date >= ? AND date <= ?
              AND main_name != 'مشتريات'
            GROUP BY main_name
            ORDER BY SUM(value) DESC
        """, (date_from, date_to))
        main_rows = c.fetchall()

        flat_subs = {}
        if self.FLATTEN_MAINS:
            placeholders = ','.join('?' * len(self.FLATTEN_MAINS))
            c.execute(f"""
                SELECT main_name, sub_name, COALESCE(SUM(value), 0)
                FROM paids
                WHERE date >= ? AND date <= ?
                  AND main_name IN ({placeholders})
                GROUP BY main_name, sub_name
                ORDER BY main_name, SUM(value) DESC
            """, [date_from, date_to] + self.FLATTEN_MAINS)

            for main_name, sub_name, total in c.fetchall():
                flat_subs.setdefault(main_name, []).append(
                    (sub_name, total))

        conn.close()

        display_rows = []
        grand_total = 0

        for main_name, main_total in main_rows:
            if main_name in self.FLATTEN_MAINS:
                subs = flat_subs.get(main_name, [])
                for sub_name, s_total in subs:
                    display_rows.append({
                        'level': 'sub',
                        'name': sub_name or '(بدون تصنيف)',
                        'amount': s_total,
                    })
                    grand_total += s_total
            else:
                display_rows.append({
                    'level': 'main',
                    'name': main_name,
                    'amount': main_total,
                })
                grand_total += main_total

        for r in display_rows:
            tag = 'main_row' if r['level'] == 'main' else 'sub_row'
            self.tree_total.insert('', 'end', tags=(tag,), values=(
                r['name'],
                f"{r['amount']:,.0f}",
            ))

        if display_rows:
            self.tree_total.insert('', 'end', tags=('total_row',), values=(
                'الإجمالي',
                f"{grand_total:,.0f}",
            ))

        self.total_data = {
            'rows': display_rows,
            'grand_total': grand_total,
        }

    # ═══════════════════════════════════════════════════════════
    # التبويب 2: التقرير التفصيلي
    # ═══════════════════════════════════════════════════════════
    def _refresh_detail_combo(self):
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM main_categories ORDER BY name")
        mains = [r[0] for r in c.fetchall()]
        conn.close()
        self.detail_main_combo['values'] = mains
        if mains and not self.detail_main_combo.get():
            self.detail_main_combo.set(mains[0])

    def _load_detail_tab(self):
        selected_main = self.detail_main_combo.get().strip()
        if not selected_main:
            return

        for item in self.tree_detail.get_children():
            self.tree_detail.delete(item)

        date_from, date_to = self._check_dates()
        if not date_from:
            return

        conn = self.get_connection()
        c = conn.cursor()

        c.execute("""
            SELECT sub_name, COALESCE(SUM(value), 0)
            FROM paids
            WHERE date >= ? AND date <= ?
              AND main_name = ?
            GROUP BY sub_name
            ORDER BY SUM(value) DESC
        """, (date_from, date_to, selected_main))
        sub_rows = c.fetchall()
        conn.close()

        grand_total = 0
        for sub_name, sub_total in sub_rows:
            self.tree_detail.insert('', 'end', tags=('row',), values=(
                sub_name or '(بدون تصنيف)',
                f"{sub_total:,.0f}",
            ))
            grand_total += sub_total

        self.tree_detail.insert('', 'end', tags=('total_row',), values=(
            'الإجمالي',
            f"{grand_total:,.0f}",
        ))

        self.detail_data = {
            'selected_main': selected_main,
            'subs': sub_rows,
            'grand_total': grand_total,
        }

    # ═══════════════════════════════════════════════════════════
    # التبويب 3: المعاملات التفصيلية
    # ═══════════════════════════════════════════════════════════
    def _refresh_trans_combo(self):
        conn = self.get_connection()
        c = conn.cursor()
        c.execute("SELECT name FROM main_categories ORDER BY name")
        mains = [r[0] for r in c.fetchall()]
        conn.close()
        self.trans_main_combo['values'] = mains
        if mains and not self.trans_main_combo.get():
            self.trans_main_combo.set(mains[0])
            self._on_trans_main_change()

    def _on_trans_main_change(self):
        selected_main = self.trans_main_combo.get().strip()
        if not selected_main:
            return

        conn = self.get_connection()
        c = conn.cursor()
        c.execute("""
            SELECT sc.name
            FROM sub_categories sc
            JOIN main_categories mc ON sc.main_id = mc.id
            WHERE mc.name = ?
            ORDER BY sc.name
        """, (selected_main,))
        subs = [r[0] for r in c.fetchall()]
        conn.close()

        self.trans_sub_combo['values'] = ['الكل'] + subs
        self.trans_sub_combo.set('الكل')
        self._load_trans_tab()

    def _load_trans_tab(self):
        selected_main = self.trans_main_combo.get().strip()
        selected_sub = self.trans_sub_combo.get().strip()

        if not selected_main:
            return

        for item in self.tree_trans.get_children():
            self.tree_trans.delete(item)

        date_from, date_to = self._check_dates()
        if not date_from:
            return

        conn = self.get_connection()
        c = conn.cursor()

        query = """
            SELECT date, description, payment_method, value
            FROM paids
            WHERE date >= ? AND date <= ?
              AND main_name = ?
        """
        params = [date_from, date_to, selected_main]

        if selected_sub and selected_sub != 'الكل':
            query += " AND sub_name = ?"
            params.append(selected_sub)

        query += " ORDER BY date DESC, id DESC"

        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        total = 0
        for i, (date, desc, payment, value) in enumerate(rows):
            tag = 'odd_row' if i % 2 == 0 else 'even_row'
            self.tree_trans.insert('', 'end', tags=(tag,), values=(
                date or '',
                desc or '(بدون وصف)',
                payment or 'نقدي',
                f"{value:,.2f}",
            ))
            total += value or 0

        if rows:
            self.tree_trans.insert('', 'end', tags=('total_row',), values=(
                '',
                f'الإجمالي ({len(rows)} معاملة)',
                '',
                f"{total:,.2f}",
            ))

        self.trans_data = {
            'selected_main': selected_main,
            'selected_sub': selected_sub,
            'rows': rows,
            'grand_total': total,
        }

    # ═══════════════════════════════════════════════════════════
    # التصدير Excel (3 أنواع)
    # ═══════════════════════════════════════════════════════════
    def _export_excel(self, tab):
        # ─── جلب البيانات ───
        if tab == 'total':
            data = self.total_data
            label = "التقرير_الإجمالي"
            title_text = "📊 التقرير الإجمالي لمصاريف التشغيل"
        elif tab == 'detail':
            data = self.detail_data
            if not data:
                messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
                return
            label = f"تفصيلي_{data['selected_main']}"
            title_text = f"📋 التقرير التفصيلي - {data['selected_main']}"
        else:  # trans
            data = self.trans_data
            if not data or not data['rows']:
                messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
                return
            sub_part = ""
            if data['selected_sub'] and data['selected_sub'] != 'الكل':
                sub_part = f"_{data['selected_sub']}"
            label = f"معاملات_{data['selected_main']}{sub_part}"
            title_text = (f"🧾 المعاملات التفصيلية - {data['selected_main']}"
                          + (f" - {data['selected_sub']}"
                             if data['selected_sub'] != 'الكل' else ""))

        if not data:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        try:
            from openpyxl import Workbook
            from openpyxl.styles import (Font, Alignment,
                                         PatternFill, Border, Side)
        except ImportError:
            messagebox.showerror("خطأ", "مكتبة openpyxl غير مثبتة")
            return

        from tkinter import filedialog
        df = self.date_from.get().strip()
        dt = self.date_to.get().strip()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{label}_{df}_{dt}.xlsx",
            title=f"حفظ {label} Excel")
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = label[:30]
            ws.sheet_view.rightToLeft = True

            # ─── عرض الأعمدة ───
            if tab in ('total', 'detail'):
                ws.column_dimensions['A'].width = 40
                ws.column_dimensions['B'].width = 20
            else:
                ws.column_dimensions['A'].width = 15
                ws.column_dimensions['B'].width = 45
                ws.column_dimensions['C'].width = 18
                ws.column_dimensions['D'].width = 18

            # ─── العنوان ───
            n_cols = 2 if tab in ('total', 'detail') else 4
            last_col = chr(ord('A') + n_cols - 1)
            ws.merge_cells(f'A1:{last_col}1')
            cell = ws['A1']
            cell.value = f"{title_text} | مطعم الحديقة ({df} - {dt})"
            cell.font = Font(name='Arial', size=14, bold=True,
                             color='FFFFFF')
            cell.fill = PatternFill("solid", fgColor="27AE60")
            cell.alignment = Alignment(horizontal='center',
                                       vertical='center')
            ws.row_dimensions[1].height = 30

            # ─── الرؤوس ───
            if tab in ('total', 'detail'):
                headers = ['البند', 'المبلغ']
            else:
                headers = ['التاريخ', 'الوصف', 'طريقة الدفع', 'المبلغ']

            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=2, column=col_idx, value=h)
                cell.font = Font(name='Arial', size=12, bold=True,
                                 color='000000')
                cell.fill = PatternFill("solid", fgColor="FFFF00")
                cell.alignment = Alignment(horizontal='center',
                                           vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
            ws.row_dimensions[2].height = 25

            # ─── البيانات ───
            row_num = 3

            if tab == 'total':
                for r in data['rows']:
                    bg = "FFF3CD" if r['level'] == 'main' else "FFFFFF"
                    bold = (r['level'] == 'main')

                    c1 = ws.cell(row=row_num, column=1, value=r['name'])
                    c1.font = Font(name='Arial', size=11, bold=bold)
                    c1.fill = PatternFill("solid", fgColor=bg)
                    c1.alignment = Alignment(horizontal='right',
                                             vertical='center')
                    c1.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

                    c2 = ws.cell(row=row_num, column=2, value=r['amount'])
                    c2.font = Font(name='Arial', size=11, bold=bold)
                    c2.fill = PatternFill("solid", fgColor=bg)
                    c2.alignment = Alignment(horizontal='center',
                                             vertical='center')
                    c2.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
                    row_num += 1

                c1 = ws.cell(row=row_num, column=1, value='الإجمالي')
                c1.font = Font(name='Arial', size=13, bold=True)
                c1.fill = PatternFill("solid", fgColor="FFFF00")
                c1.alignment = Alignment(horizontal='right',
                                         vertical='center')
                c1.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

                c2 = ws.cell(row=row_num, column=2,
                             value=data['grand_total'])
                c2.font = Font(name='Arial', size=13, bold=True)
                c2.fill = PatternFill("solid", fgColor="FFFF00")
                c2.alignment = Alignment(horizontal='center',
                                         vertical='center')
                c2.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            elif tab == 'detail':
                for sub_name, sub_total in data['subs']:
                    c1 = ws.cell(row=row_num, column=1,
                                 value=sub_name or '(بدون تصنيف)')
                    c1.font = Font(name='Arial', size=11)
                    c1.alignment = Alignment(horizontal='right',
                                             vertical='center')
                    c1.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

                    c2 = ws.cell(row=row_num, column=2, value=sub_total)
                    c2.font = Font(name='Arial', size=11)
                    c2.alignment = Alignment(horizontal='center',
                                             vertical='center')
                    c2.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
                    row_num += 1

                c1 = ws.cell(row=row_num, column=1, value='الإجمالي')
                c1.font = Font(name='Arial', size=13, bold=True)
                c1.fill = PatternFill("solid", fgColor="FFFF00")
                c1.alignment = Alignment(horizontal='right',
                                         vertical='center')
                c1.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

                c2 = ws.cell(row=row_num, column=2,
                             value=data['grand_total'])
                c2.font = Font(name='Arial', size=13, bold=True)
                c2.fill = PatternFill("solid", fgColor="FFFF00")
                c2.alignment = Alignment(horizontal='center',
                                         vertical='center')
                c2.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            else:  # trans
                for date, desc, payment, value in data['rows']:
                    values = [date or '', desc or '(بدون وصف)',
                              payment or 'نقدي', value]
                    for col_idx, val in enumerate(values, start=1):
                        cell = ws.cell(row=row_num, column=col_idx,
                                       value=val)
                        cell.font = Font(name='Arial', size=11)
                        cell.alignment = Alignment(
                            horizontal='right' if col_idx == 2 else 'center',
                            vertical='center')
                        cell.border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin'))
                    row_num += 1

                ws.merge_cells(start_row=row_num, start_column=1,
                               end_row=row_num, end_column=3)
                c1 = ws.cell(row=row_num, column=1, value='الإجمالي')
                c1.font = Font(name='Arial', size=13, bold=True)
                c1.fill = PatternFill("solid", fgColor="FFFF00")
                c1.alignment = Alignment(horizontal='right',
                                         vertical='center')
                c1.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

                c2 = ws.cell(row=row_num, column=4,
                             value=data['grand_total'])
                c2.font = Font(name='Arial', size=13, bold=True)
                c2.fill = PatternFill("solid", fgColor="FFFF00")
                c2.alignment = Alignment(horizontal='center',
                                         vertical='center')
                c2.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

            wb.save(file_path)
            messagebox.showinfo("نجاح", f"✅ تم الحفظ:\n{file_path}")
            if messagebox.askyesno("فتح الملف", "هل تريد فتح الملف الآن؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ", f"فشل التصدير:\n{e}")

    # ═══════════════════════════════════════════════════════════
    # التصدير صورة (3 أنواع)
    # ═══════════════════════════════════════════════════════════
    def _export_image(self, tab):
        # ─── جلب البيانات ───
        if tab == 'total':
            data = self.total_data
            label = "التقرير_الإجمالي"
            title = "📊 التقرير الإجمالي لمصاريف التشغيل - مطعم الحديقة"
        elif tab == 'detail':
            data = self.detail_data
            if not data:
                messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
                return
            label = f"تفصيلي_{data['selected_main']}"
            title = (f"📋 التقرير التفصيلي - {data['selected_main']}"
                     f" - مطعم الحديقة")
        else:  # trans
            data = self.trans_data
            if not data or not data['rows']:
                messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
                return
            sub_part = ""
            if data['selected_sub'] and data['selected_sub'] != 'الكل':
                sub_part = f"_{data['selected_sub']}"
            label = f"معاملات_{data['selected_main']}{sub_part}"
            title = (f"🧾 المعاملات التفصيلية - {data['selected_main']}"
                     + (f" - {data['selected_sub']}"
                        if data['selected_sub'] != 'الكل' else "")
                     + " - مطعم الحديقة")

        if not data:
            messagebox.showwarning("تنبيه", "لا توجد بيانات للتصدير")
            return

        from tkinter import filedialog
        df = self.date_from.get().strip()
        dt = self.date_to.get().strip()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            initialfile=f"{label}_{df}_{dt}.png",
            title="حفظ التقرير صورة")
        if not file_path:
            return

        # ─── تجهيز الأعمدة والصفوف ───
        if tab == 'total':
            columns = ['البند', 'المبلغ']
            rows_data = []
            for r in data['rows']:
                rows_data.append((r['name'], f"{r['amount']:,.0f}"))
            rows_data.append(('الإجمالي', f"{data['grand_total']:,.0f}"))
            summary_items = None

        elif tab == 'detail':
            columns = ['البند', 'المبلغ']
            rows_data = []
            for sub_name, sub_total in data['subs']:
                rows_data.append((sub_name or '(بدون تصنيف)',
                                  f"{sub_total:,.0f}"))
            rows_data.append(('الإجمالي', f"{data['grand_total']:,.0f}"))
            summary_items = None

        else:  # trans
            columns = ['التاريخ', 'الوصف', 'طريقة الدفع', 'المبلغ']
            rows_data = []
            for date, desc, payment, value in data['rows']:
                rows_data.append((date or '',
                                  desc or '(بدون وصف)',
                                  payment or 'نقدي',
                                  f"{value:,.2f}"))
            summary_items = [
                ('عدد المعاملات', str(len(data['rows']))),
                ('─────', ''),
                ('الإجمالي', f"{data['grand_total']:,.2f} ج.م"),
            ]

        success = export_table_to_image(
            file_path=file_path,
            title=title,
            subtitle=f"من {df} إلى {dt}",
            columns=columns,
            rows=rows_data,
            summary_items=summary_items,
            logo_path=COMPANY_LOGO_PATH)

        if success:
            if messagebox.askyesno("نجاح",
                                   f"✅ تم الحفظ:\n{file_path}\n\nفتح الصورة؟"):
                try:
                    os.startfile(file_path)
                except:
                    pass

# ═══════════════════════════════════════════════════════════════════
# ABOUT PAGE (عن الشركة)
# ═══════════════════════════════════════════════════════════════════
class AboutPage(tk.Frame):
    """صفحة التعريف بالشركة والمسؤول والبرامج"""

    def __init__(self, parent, controller=None):
        super().__init__(parent, bg='#ecf0f1')
        self.controller = controller

        # 1) الشريط العلوي
        add_program_banner(self)

        # 2) منطقة التمرير
        content = make_scrollable(self)
        self._page_content = content

        # ═══════════════════════════════════════════════════════
        # العنوان الرئيسي
        # ═══════════════════════════════════════════════════════
        tk.Label(content, text="🏢 عن الشركة",
                 font=('Arial', 22, 'bold'),
                 bg='#ecf0f1', fg='#1a5276').pack(pady=(20, 10))

        # ═══════════════════════════════════════════════════════
        # الشعار الكبير
        # ═══════════════════════════════════════════════════════
        logo_frame = tk.Frame(content, bg='#ecf0f1')
        logo_frame.pack(pady=(10, 20))

        logo_displayed = False
        if os.path.exists(COMPANY_LOGO_PATH):
            try:
                logo_img = Image.open(COMPANY_LOGO_PATH)
                if logo_img.mode != 'RGBA':
                    logo_img = logo_img.convert('RGBA')
                # ✅ تكبير الشعار إلى 220×220
                logo_img.thumbnail((220, 220), Image.Resampling.LANCZOS)
                logo_photo = ImageTk.PhotoImage(logo_img)
                lbl = tk.Label(logo_frame, image=logo_photo,
                               bg='#ecf0f1')
                lbl.image = logo_photo  # ✅ لتجنب حذفها من الذاكرة
                lbl.pack()
                logo_displayed = True
            except Exception as e:
                print(f"⚠️ [فشل تحميل الشعار]: {e}")

        if not logo_displayed:
            tk.Label(logo_frame, text="🏢",
                     font=('Arial', 80),
                     bg='#ecf0f1', fg='#1a5276').pack()

        # ─── اسم الشركة تحت الشعار ───
        tk.Label(content,
                 text="شركة الزهراء للبرمجيات",
                 font=('Arial', 24, 'bold'),
                 bg='#ecf0f1', fg='#1a5276').pack(pady=(0, 5))

        tk.Label(content,
                 text="Zahraa Software Company",
                 font=('Arial', 14, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 25))

        # ═══════════════════════════════════════════════════════
        # قسم: عن الشركة
        # ═══════════════════════════════════════════════════════
        company_section = tk.LabelFrame(
            content, text="  🏢  عن الشركة  ",
            font=('Arial', 14, 'bold'),
            bg='#ecf0f1', fg='#1a5276',
            padx=25, pady=20)
        company_section.pack(fill='x', padx=40, pady=(0, 20))

        company_info = (
            "شركة الزهراء للبرمجيات هي شركة رائدة في مجال تصميم وتطوير "
            "الأنظمة والبرامج المحاسبية والإدارية المتكاملة.\n\n"
            "نقدم حلولاً برمجية احترافية تلبي احتياجات الشركات والمؤسسات "
            "والمطاعم والمحلات التجارية، مع التركيز على:\n\n"
            "    ✅  جودة البرمجة والاحترافية في التنفيذ\n\n"
            "    ✅  سهولة الاستخدام وبساطة الواجهات\n\n"
            "    ✅  الدعم الفني المستمر بعد التسليم\n\n"
            "    ✅  التخصيص حسب احتياجات كل عميل\n\n"
        )
        tk.Label(company_section, text=company_info,
                 font=('Arial', 12),
                 bg='#ecf0f1', fg='#2c3e50',
                 justify='right', anchor='e',
                 wraplength=900).pack(fill='x')

        # ═══════════════════════════════════════════════════════
        # قسم: المسؤول عن الشركة
        # ═══════════════════════════════════════════════════════
        manager_section = tk.LabelFrame(
            content, text="  👨‍💼  المسؤول عن الشركة  ",
            font=('Arial', 14, 'bold'),
            bg='#ecf0f1', fg='#c0392b',
            padx=25, pady=20)
        manager_section.pack(fill='x', padx=40, pady=(0, 20))

        # ─── الاسم ───
        tk.Label(manager_section,
                 text="المحاسب / حسام الدين سيد ربيع",
                 font=('Arial', 18, 'bold'),
                 bg='#ecf0f1', fg='#1a5276').pack(pady=(0, 5))

        # ─── المنصب ───
        tk.Label(manager_section,
                 text="المدير التنفيذي ومصمم البرامج",
                 font=('Arial', 12, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(0, 15))

        # ─── بيانات التواصل ───
        contact_info = (
            "📞  الهاتف / واتساب:  01003796984\n"
            "✉️  البريد الإلكتروني:  hussein@zahraa-soft.com\n"
            "📍  العنوان:  مصر"
        )
        tk.Label(manager_section, text=contact_info,
                 font=('Arial', 12, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50',
                 justify='right', anchor='e').pack()

        # ═══════════════════════════════════════════════════════
        # قسم: البرامج المصممة
        # ═══════════════════════════════════════════════════════
        programs_section = tk.LabelFrame(
            content, text="  💻  البرامج المصممة من الشركة  ",
            font=('Arial', 14, 'bold'),
            bg='#ecf0f1', fg='#16a085',
            padx=25, pady=20)
        programs_section.pack(fill='x', padx=40, pady=(0, 20))

        # ─── قائمة البرامج ───
        programs = [
            ("🍽️", "برنامج حسابات المطاعم",
             "نظام متكامل لإدارة المطاعم والمبيعات والمصاريف والمخزون"),
            ("🏪", "برنامج إدارة المحلات التجارية",
             "نظام شامل لإدارة المبيعات، المخزون، الفواتير، والعملاء"),
            ("📦", "برنامج إدارة المخازن",
             "نظام متطور لتتبع المخزون والجرد الشهري والتنبيهات"),
            ("💰", "برنامج الحسابات العامة",
             "نظام محاسبي متكامل للقيود والتقارير المالية"),
            ("👥", "برنامج إدارة الموارد البشرية",
             "نظام لإدارة العاملين، الرواتب، السلف، والحضور"),
            ("📊", "برنامج التقارير والتحليلات",
             "نظام تقارير ذكي مع رسوم بيانية وتصدير Excel"),
        ]

        for icon, name, desc in programs:
            row = tk.Frame(programs_section, bg='#ecf0f1')
            row.pack(fill='x', pady=6)

            # ✅ إطار ثابت العرض للأيقونة (60 بكسل)
            icon_holder = tk.Frame(row, bg='#ecf0f1', width=60, height=40)
            icon_holder.pack(side='right', padx=(0, 10))
            icon_holder.pack_propagate(False)  # ← منع الإطار من التمدد

            tk.Label(icon_holder, text=icon,
                     font=('Arial', 20),
                     bg='#ecf0f1').pack(expand=True)

            # ─── نصوص البرنامج ───
            text_frame = tk.Frame(row, bg='#ecf0f1')
            text_frame.pack(side='right', fill='x', expand=True)

            tk.Label(text_frame, text=name,
                     font=('Arial', 13, 'bold'),
                     bg='#ecf0f1', fg='#2c3e50',
                     anchor='e').pack(fill='x')

            tk.Label(text_frame, text=desc,
                     font=('Arial', 11),
                     bg='#ecf0f1', fg='#7f8c8d',
                     anchor='e').pack(fill='x')

        # ═══════════════════════════════════════════════════════
        # التذييل
        # ═══════════════════════════════════════════════════════
        footer = tk.Frame(content, bg='#2c3e50', height=60)
        footer.pack(fill='x', pady=(20, 0))
        footer.pack_propagate(False)

        tk.Label(footer,
                 text="© 2025 شركة الزهراء للبرمجيات — جميع الحقوق محفوظة",
                 font=('Arial', 11, 'bold'),
                 bg='#2c3e50', fg='#f1c40f').pack(pady=20)

    def load_traders(self):
        """يُستدعى تلقائيًا من MainApp (لا يحتاج تحديث)"""
        pass


# =============================================
# 3. الصفحة الرئيسية (التحكم)
# =============================================
# ═══════════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════
class MainApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("RestoBooks - نظام حسابات المطاعم")
        self.geometry("1100x700")

        # ═══════════════════════════════════════════════════════
        # 1) الإعلان المتحرك
        # ═══════════════════════════════════════════════════════
        self.ad_full_text = COMPANY_AD * 6
        self.ad_display_length = AD_DISPLAY_LENGTH
        self.ad_index = 0
        self.is_ad_mode = True
        self.scroll_ad()

        # ═══════════════════════════════════════════════════════
        # 2) القائمة الجانبية
        # ═══════════════════════════════════════════════════════
        sidebar = tk.Frame(self, bg=SIDEBAR_BG, width=210)
        sidebar.pack(side='left', fill='y')
        sidebar.pack_propagate(False)

        # ─── عنوان القائمة ───
        title_frame = tk.Frame(sidebar, bg=SIDEBAR_BG, height=70)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="📋  القائمة الرئيسية",
                 font=('Arial', 14, 'bold'),
                 bg=SIDEBAR_BG, fg=SIDEBAR_ACCENT).pack(pady=20)

        # ─── خط ذهبي فاصل ───
        tk.Frame(sidebar, bg=SIDEBAR_ACCENT, height=2).pack(fill='x', padx=15)

        # ─── قاموس الأزرار + الصفحة النشطة ───
        self.menu_buttons = {}
        self.active_page_key = None

        # ─── دالة إنشاء زر موحّد ───
        def make_menu_button(text, command, page_key=None):
            """إنشاء زر قائمة مع تمييز الصفحة النشطة"""
            btn = tk.Button(
                sidebar,
                text=f"  {text}",
                command=command,
                bg=SIDEBAR_BTN_BG,
                fg=SIDEBAR_FG,
                font=('Arial', 12, 'bold'),
                relief='flat',
                bd=0,
                activebackground=SIDEBAR_BTN_HOVER,
                activeforeground='#1a5276',
                cursor='hand2',
                anchor='w',
                padx=10,
                pady=8,
            )
            btn.pack(pady=3, padx=10, fill='x')

            # ✅ تأثير hover (لا يعمل على الزر النشط)
            def on_enter(e):
                if self.active_page_key != page_key:
                    btn.config(bg=SIDEBAR_BTN_HOVER, fg='#1a5276')

            def on_leave(e):
                if self.active_page_key != page_key:
                    btn.config(bg=SIDEBAR_BTN_BG, fg=SIDEBAR_FG)

            btn.bind('<Enter>', on_enter)
            btn.bind('<Leave>', on_leave)

            # خزّن الزر في القاموس
            if page_key:
                self.menu_buttons[page_key] = btn

            return btn

        # ═══════════════════════════════════════════════════════
        # 3) إنشاء الأزرار
        # ═══════════════════════════════════════════════════════
        self.btn_Sales_page = make_menu_button(
            "📚  المبيعات",
            lambda: self.show_page(salesPage),
            page_key='salesPage'
        )

        self.btn_Expenses = make_menu_button(
            "💸  المصاريف",
            lambda: self.show_page(paidsPage),
            page_key='paidsPage'
        )

        make_menu_button(
            "📅  اليومية",
            lambda: self.show_page(DailyReportPage),
            page_key='DailyReportPage'
        )

        make_menu_button(
            "📊  تقرير مدة",
            lambda: self.show_page(PeriodReportPage),
            page_key='PeriodReportPage'
        )

        make_menu_button(
            "💰  الخزينة",
            lambda: self.show_page(TreasuryPage),
            page_key='TreasuryPage'
        )

        self.btn_inventroy = make_menu_button(
            "📦  المخزن",
            lambda: self.show_page(InventoryPage),
            page_key='InventoryPage'
        )

        make_menu_button(
            "🧑‍💼  التجار",
            lambda: self.show_page(TradersPage),
            page_key='TradersPage'
        )

        make_menu_button(
            "📄  الفواتير",
            lambda: self.show_page(InvoicesPage),
            page_key='InvoicesPage'
        )

        make_menu_button(
            "👷  العاملين",
            lambda: self.show_page(EmployeesPage),
            page_key='EmployeesPage'
        )

        make_menu_button(
            "📱  فودافون كاش",
            lambda: self.show_page(VodafoneWalletPage),
            page_key='VodafoneWalletPage'
        )
        make_menu_button(
            "📈  التقرير الختامي",
            lambda: self.show_page(FinalReportPage),
            page_key='FinalReportPage'
        )
        make_menu_button(
            "📑  تقرير تفصيلي",
            lambda: self.show_page(DetailedReportPage),
            page_key='DetailedReportPage'
        )
        make_menu_button(
            "ℹ️  عن الشركة",
            lambda: self.show_page(AboutPage),
            page_key='AboutPage'
        )

        # ═══════════════════════════════════════════════════════
        # 4) إطار عرض الصفحات
        # ═══════════════════════════════════════════════════════
        self.container = tk.Frame(self, bg='#ecf0f1')
        self.container.pack(side='right', fill='both', expand=True)

        # قاموس الصفحات
        self.pages = {}

        # عرض الصفحة الافتراضية
        self.show_page(salesPage)
        # ✅ إغلاق شاشة التحميل
        close_splash()   # إغلاق فوري


        # ═══════════════════════════════════════════════════════
        # 5) تنسيق الجداول (Treeview)
        # ═══════════════════════════════════════════════════════
        style = ttk.Style()
        style.theme_use('clam')

        # رأس الجدول
        style.configure(
            "Treeview.Heading",
            font=('Arial', 12, 'bold'),
            background='#2c3e50',
            foreground='white',
            relief='raised',
            borderwidth=1,
        )
        style.map("Treeview.Heading",
                  background=[('active', '#34495e')])

        # صفوف الجدول
        style.configure(
            "Treeview",
            font=('Arial', 11),
            rowheight=30,
            background='white',
            fieldbackground='white',
            foreground='black',
            borderwidth=0,
            relief='flat',
        )
        style.map("Treeview",
                  background=[('selected', '#3498db')],
                  foreground=[('selected', 'white')])

    # ═══════════════════════════════════════════════════════════
    # دوال التمرير والتمييز والعرض
    # ═══════════════════════════════════════════════════════════
    def scroll_ad(self):
        """تحريك الإعلان داخل الشريط العلوي لكل الصفحات"""
        if not self.is_ad_mode:
            return

        # تنظيف الليبلات المحذوفة
        global BANNER_AD_LABELS
        BANNER_AD_LABELS = [lbl for lbl in BANNER_AD_LABELS
                            if lbl.winfo_exists()]

        full_text = self.ad_full_text
        start = self.ad_index
        end = start + self.ad_display_length

        if end >= len(full_text):
            self.ad_index = 0
            start = self.ad_index
            end = start + self.ad_display_length

        display_part = full_text[start:end]

        # ✅ تحديث كل ليبلات الشريط العلوي
        for lbl in BANNER_AD_LABELS:
            try:
                lbl.config(text=display_part)
            except tk.TclError:
                pass

        self.ad_index += 1
        self.after(AD_SPEED, self.scroll_ad)

    def _highlight_active_button(self, active_key):
        """تمييز الزر النشط بلون ذهبي، وإعادة الباقي للأزرق"""
        self.active_page_key = active_key
        for key, btn in self.menu_buttons.items():
            try:
                if key == active_key:
                    btn.config(bg=SIDEBAR_BTN_HOVER, fg='#1a5276')  # ذهبي
                else:
                    btn.config(bg=SIDEBAR_BTN_BG, fg=SIDEBAR_FG)     # أزرق
            except tk.TclError:
                pass

    def show_page(self, page_class):
        page_name = page_class.__name__

        # ─── إنشاء الصفحة أو استرجاعها ───
        if page_name not in self.pages:
            self.pages[page_name] = page_class(
                parent=self.container,
                controller=self
            )
            self.pages[page_name].pack(fill='both', expand=True)
        else:
            self.pages[page_name].pack(fill='both', expand=True)

        # ─── إعادة تحميل البيانات عند كل عرض ───
        if hasattr(self.pages[page_name], 'load_traders'):
            try:
                self.pages[page_name].load_traders()
            except Exception:
                pass

        # ─── إخفاء باقي الصفحات ───
        for name, page in self.pages.items():
            if name != page_name:
                page.pack_forget()

        # ─── تمييز الزر النشط ───
        self._highlight_active_button(page_name)


# =============================================
# 4. تشغيل التطبيق
# =============================================
if __name__ == '__main__':
    init_db()
    migrate_add_sub_name2()
    cleanup_wrong_treasury_actions()   # ✅ جديد — تنظيف تلقائي
    app = MainApp()
    app.mainloop()

