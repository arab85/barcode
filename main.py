import tkinter as tk
import jdatetime
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
import os

# تنظیمات اصلی پنجره
root = tk.Tk()
root.title("نیرو انسانی")
root.geometry("1600x800")
root.configure(bg="#f0f0f0")

# رنگ‌ها و فونت‌ها
COLORS = {
    "primary": "#2c3e50",
    "secondary": "#34495e",
    "accent": "#3498db",
    "success": "#2ecc71",
    "warning": "#f39c12",
    "danger": "#e74c3c",
    "light": "#ecf0f1",
    "dark": "#2c3e50",
    "background": "#f5f5f5"
}

FONTS = {
    "title": ("B Nazanin", 40, "bold"),
    "heading": ("B Nazanin", 30, "bold"),
    "subheading": ("B Nazanin", 24),
    "normal": ("B Nazanin", 18),
    "small": ("B Nazanin", 14)
}

def update_time():
    now_gregorian = datetime.now()
    now_jalali = jdatetime.datetime.now()

    date_str = now_jalali.strftime("%Y/%m/%d")
    time_str = now_gregorian.strftime("%H:%M:%S")

    label.config(text=f"{date_str}    {time_str}")
    label.after(1000, update_time)  

# ایجاد ویجت زمان
label = tk.Label(root, font=FONTS["normal"], fg=COLORS["primary"], bg=COLORS["background"])
label.place(x=50, y=20)

# فعال‌سازی کشیدن ردیف و ستون اصلی
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# ایجاد فریم‌ها
def create_frame():
    frame = tk.Frame(root, bg=COLORS["background"])
    frame.place(x=0, y=80, width=1600, height=720)
    return frame


ramz = create_frame()
home = create_frame()
search = create_frame()
search_b = create_frame()
search_d = create_frame()
add = create_frame()


# متغیرهای ورودی
fname_var = tk.StringVar()
lname_var = tk.StringVar()
pedar_var = tk.StringVar()
meli_var = tk.IntVar()
shenasi_var = tk.IntVar()
vaz_var = tk.StringVar()
#تاریخ تولد و عضویت
roz_var = tk.IntVar()
mah_var = tk.IntVar()
sal_var = tk.IntVar()

def save_to_excel(filename, data_dict, sheet_name="داده‌ها"):
    """ذخیره داده‌ها در اکسل"""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        
        # اضافه کردن هدر
        ws.append(["نام آیتم", "تعداد"])
        
        # اضافه کردن داده‌ها
        for item, count in data_dict.items():
            ws.append([item, count])
        
        wb.save(filename)
        return True
    except Exception as e:
        print(f"خطا در ذخیره فایل {filename}: {e}")
        return False
    
# --- ایجاد ویجت‌های مشترک ---
def create_button(parent, text, command, x=None, y=None, width=20, height=2):
    btn = tk.Button(parent, text=text, font=FONTS["normal"], bg=COLORS["accent"], 
                   fg=COLORS["light"], relief="raised", bd=2, command=command,
                   width=width, height=height)
    if x is not None and y is not None:
        btn.place(x=x, y=y)
    else:
        btn.pack(pady=10)
    return btn

def create_label(parent, text, font_style="normal", x=None, y=None):
    label = tk.Label(parent, text=text, font=FONTS[font_style], 
                    bg=COLORS["background"], fg=COLORS["dark"])
    if x is not None and y is not None:
        label.place(x=x, y=y)
    else:
        label.pack(pady=10)
    return label

def create_entry(parent, textvariable, x=None, y=None, width=20):
    entry = tk.Entry(parent, textvariable=textvariable, font=FONTS["normal"], 
                    width=width, justify="center", bd=2, relief="sunken")
    if x is not None and y is not None:
        entry.place(x=x, y=y)
    else:
        entry.pack(pady=5)
    return entry