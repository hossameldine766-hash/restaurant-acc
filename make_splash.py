# ═══════════════════════════════════════════════════════════
# إنشاء شاشة Splash لـ RestoBooks
# ═══════════════════════════════════════════════════════════
from PIL import Image, ImageDraw, ImageFont
import os

# ─── الأبعاد ───
W, H = 500, 400
BG_COLOR = '#1a5276'       # أزرق داكن
GOLD = '#f1c40f'           # ذهبي
WHITE = '#ffffff'
GRAY = '#bdc3c7'

# ─── إنشاء الصورة ───
img = Image.new('RGB', (W, H), BG_COLOR)
draw = ImageDraw.Draw(img)

# ─── الخطوط ───
def load_font(name, size):
    paths = [
        f"C:/Windows/Fonts/{name}",
        f"C:/Windows/Fonts/arial.ttf",
    ]
    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except:
            pass
    return ImageFont.load_default()

font_logo_name = load_font("arialbd.ttf", 42)
font_sub = load_font("arial.ttf", 16)
font_desc = load_font("arial.ttf", 13)
font_small = load_font("arial.ttf", 11)

# ─── الشعار (في المنتصف) ───
if os.path.exists("m.png"):
    logo = Image.open("m.png").convert("RGBA")
    logo.thumbnail((130, 130), Image.Resampling.LANCZOS)
    logo_x = (W - logo.width) // 2
    img.paste(logo, (logo_x, 35), logo)

# ─── خط ذهبي فاصل (أعلى) ───
draw.line([(80, 185), (W - 80, 185)], fill=GOLD, width=2)

# ─── اسم البرنامج ───
title = "RestoBooks"
title_w = draw.textlength(title, font=font_logo_name)
draw.text(((W - title_w) / 2, 200), title, fill=GOLD, font=font_logo_name)

# ─── الوصف ───
sub = "Restaurant Accounting System"
sub_w = draw.textlength(sub, font=font_sub)
draw.text(((W - sub_w) / 2, 255), sub, fill=WHITE, font=font_sub)

# ─── الوصف العربي ───
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    ar = lambda t: get_display(arabic_reshaper.reshape(t))
    ar_sub = ar("نظام حسابات المطاعم المتكامل")
    ar_w = draw.textlength(ar_sub, font=font_desc)
    draw.text(((W - ar_w) / 2, 280), ar_sub, fill=WHITE, font=font_desc)
except:
    pass

# ─── خط ذهبي فاصل (أسفل) ───
draw.line([(80, 310), (W - 80, 310)], fill=GOLD, width=2)

# ─── حالة التحميل ───
try:
    ar_load = ar("جارٍ التحميل، يرجى الانتظار...")
    ar_load_w = draw.textlength(ar_load, font=font_desc)
    draw.text(((W - ar_load_w) / 2, 325), ar_load, fill=GRAY, font=font_desc)
except:
    pass

# ─── الشركة ───
company = "Zahraa Software © 2025"
company_w = draw.textlength(company, font=font_small)
draw.text(((W - company_w) / 2, 370), company, fill=GRAY, font=font_small)

# ─── الحفظ ───
img.save("splash.png", "PNG")
print("✅ تم إنشاء splash.png بنجاح!")
print("📐 الأبعاد: 500 × 400")
print("📁 المكان: splash.png")