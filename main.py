import tkinter as tk
import jdatetime
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
import os
import barcode
from barcode.writer import ImageWriter


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
delat = create_frame()

def show_frame(frame):
    frame.tkraise()


# متغیرهای ورودی
fname_var = tk.StringVar()
lname_var = tk.StringVar()
pedar_var = tk.StringVar()
a = tk.StringVar()
meli_var = tk.IntVar()
shenasi_var = tk.IntVar()
vaz_var = tk.StringVar()
#تاریخ تولد و عضویت
troz_var = tk.IntVar()
tmah_var = tk.IntVar()
tsal_var = tk.IntVar()
oroz_var = tk.IntVar()
omah_var = tk.IntVar()
osal_var = tk.IntVar()
delete_meli_var = tk.StringVar()
search_meli_var = StringVar()

def save_to_excel(filename, data_dict, sheet_name="داده‌ها"):
    """ذخیره اطلاعات هر فرد به‌صورت ردیفی (افقی)"""
    try:
        file_exists = os.path.exists(filename)

        if file_exists:
            wb = load_workbook(filename)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_name
            # نوشتن عنوان‌ها فقط یک‌بار
            ws.append(list(data_dict.keys()))

        # نوشتن مقادیر (ردیف جدید برای هر فرد)
        ws.append(list(data_dict.values()))

        wb.save(filename)
        return True
    except Exception as e:
        print(f"❌ خطا در ذخیره فایل {filename}: {e}")
        return False


def save_data_and_generate_barcode():
    result_label.config(text="⏳ لطفاً صبر کنید...", fg="blue")
    add.after(1000, do_save_and_generate)  # بعد از ۱ ثانیه تابع اصلی رو اجرا کنه

def do_save_and_generate():
    data_dict = {
        "نام": str(fname_var.get()).strip(),
        "فامیلی": str(lname_var.get()).strip(),
        "نام پدر": str(pedar_var.get()).strip(),
        "کد ملی": str(meli_var.get()).strip(),
        "شناسه بسیج": str(shenasi_var.get()).strip(),
        "وضعیت عضویت": str(vaz_var.get()).strip(),
        "روز تولد": str(troz_var.get()).strip(),
        "ماه تولد": str(tmah_var.get()).strip(),
        "سال تولد": str(tsal_var.get()).strip(),
        "روز عضویت": str(oroz_var.get()).strip(),
        "ماه عضویت": str(omah_var.get()).strip(),
        "سال عضویت": str(osal_var.get()).strip(),
    }

    # بررسی اینکه هیچ‌کدام از فیلدها خالی نباشد
    empty_fields = [key for key, value in data_dict.items() if value == ""]
    if empty_fields:
        result_label.config(text=f"❌ لطفاً همه فیلدها را پر کنید (خالی: {', '.join(empty_fields)})", fg="red")
        return

    # ذخیره در اکسل
    success = save_to_excel("member_data.xlsx", data_dict)
    
    if success:
        result_label.config(text="✅ ذخیره با موفقیت انجام شد", fg="green")
    else:
        result_label.config(text="❌ خطا در ذخیره اطلاعات", fg="red")

    # ساخت بارکد از کد ملی
    code = data_dict["کد ملی"]
    if code:
        try:
            barcode_class = barcode.get_barcode_class('code128')
            my_barcode = barcode_class(code, writer=ImageWriter())
            barcode_filename = my_barcode.save("barcode_image")
            # اگر خواستی: show_barcode_image(barcode_filename + ".png")
        except Exception as e:
            print(f"خطا در ساخت بارکد: {e}")


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
#حذف کاربر

def delete_person_by_national_code(widget):
    national_code = str(widget.get()).strip()
    result_labe.config(text="صبر ", fg="blue")
    delat.after(1000, lambda: show_frame(home)) 
    if not national_code:
        result_labe.config(text="❌ لطفاً کد ملی را وارد کنید")
        return

    try:
        wb = load_workbook("member_data.xlsx")
        ws = wb.active
        
        row_to_delete = None
        for row in ws.iter_rows(min_row=2, values_only=False):  # فرض بر اینه که ردیف اول عنوان‌هاست
            for cell in row:
                if cell.value == national_code:
                    row_to_delete = cell.row
                    break
            if row_to_delete:
                break

        if row_to_delete:
            ws.delete_rows(row_to_delete, 1)
            wb.save("member_data.xlsx")
            result_labe.config(text="✅ فرد با موفقیت حذف شد")
        else:
            result_labe.config(text="❌ فردی با این کد ملی پیدا نشد")

    except FileNotFoundError:
        result_labe.config(text="❌ فایل اکسل پیدا نشد")
    except Exception as e:
        result_labe.config(text=f"❌ خطا در حذف: {e}")

# ورود
def check():
    password = a.get().strip()
    if password:
        if password == "1234":
            result_label_a.config(text="خوش آمدید", fg="blue")
            ramz.after(1000, lambda: show_frame(home)) 
        else:
            result_label_a.config(text="رمز اشتباه است.", fg="red")
    else:
        result_label_a.config(text="لطفاً همه فیلدها را پر کنید.", fg="red")



create_label(ramz, "ورود", "title", x=1200, y=300)
create_label(ramz, "رمز را وارد کنید", "title", x=800, y=300)
create_entry(ramz, a)
create_button(ramz, " ورود", check)
result_label_a = create_label(ramz, "")
create_button(ramz, "خروج", root.quit, x=100, y=500)

#خانه
create_label(home, "خانه", "title", x=1200, y=300)
create_button(home, "جستجو اعضا ", lambda: show_frame(search), x=100, y=100)
create_button(home, "اضافه کردن اعضا ", lambda: show_frame(add), x=100, y=300)
create_button(home, "خروج", root.quit, x=100, y=500)

#سرچ

create_label(add, "کد ملی برای جستجو", x=600, y=400)
create_entry(add, search_meli_var, x=500, y=440)
create_button(add, "جستجو", lambda: search_person_by_national_code(search_meli_var), x=600, y=480)


#افزودن اعضا
create_label(add, "اضافه نمودن عضو", "title", x=650, y=10)
create_label(add, "اسم", x=210, y=120)
create_entry(add, fname_var, x=100, y=160)

create_label(add, "فامیلی", x=560, y=120)
create_entry(add, lname_var, x=450, y=160)

create_label(add, "نام پدر", x=960, y=120)
create_entry(add, pedar_var, x=850, y=160)

create_label(add, "کد ملی", x=1360, y=120)
create_entry(add, meli_var, x=1250, y=160)

create_label(add, "شناسه بسیج", x=210, y=220)
create_entry(add, shenasi_var, x=100, y=260)

create_label(add, "وضعیت عضویت", x=560, y=220)
create_entry(add, vaz_var, x=450, y=260)

create_label(add, "روز تولد", x=960, y=220)
create_entry(add, troz_var, x=850, y=260)

create_label(add, "ماه تولد", x=1360, y=220)
create_entry(add, tmah_var, x=1250, y=260)

create_label(add, "سال تولد", x=210, y=320)
create_entry(add, tsal_var, x=100, y=360)

create_label(add, "روز عضویت", x=560, y=320)
create_entry(add, oroz_var, x=450, y=360)

create_label(add, "ماه عضویت", x=960, y=320)
create_entry(add, omah_var, x=850, y=360)

create_label(add, "سال عضویت", x=1360, y=320)
create_entry(add, osal_var, x=1250, y=360)

create_button(add, "ذخیره و ساخت بارکد", save_data_and_generate_barcode, x=600, y=500)
create_button(add, "حذف کاربر", lambda: show_frame(delat), x=300, y=500)
result_label = create_label(add, "",x=700,y=650)
create_button(add, "بازگشت", lambda: show_frame(home), x=100, y=500)

#حذف فرد

create_label(delat, "کد ملی برای حذف", x=600, y=400)
create_entry(delat, delete_meli_var, x=500, y=440)
result_labe = create_label(delat, "",x=700,y=650)
create_button(delat, "حذف عضو", lambda: delete_person_by_national_code(delete_meli_var), x=600, y=480)
create_button(delat, "بازگشت", lambda: show_frame(add), x=100, y=100)


# برای نمایش اولیه فریم خانه
show_frame(ramz)
update_time()

root.mainloop()