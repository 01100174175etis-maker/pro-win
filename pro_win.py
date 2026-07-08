import os
import glob
import re
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
import pandas as pd

# استيراد أداة التقويم الرسومية بآلية مستقرة
try:
    from tkcalendar import DateEntry
    HAVE_TKCAL = True
except ImportError:
    HAVE_TKCAL = False
    # بديل بسيط لـ DateEntry عندما لا تكون tkcalendar مثبّتة
    class DateEntry(ttk.Entry):
        def __init__(self, master=None, **kwargs):
            # قبول وسيلة التهيئة الشائعة ولكن تجاهلها إن لم تكن مطلوبة
            kwargs.pop('date_pattern', None)
            kwargs.pop('year', None)
            kwargs.pop('month', None)
            kwargs.pop('day', None)
            kwargs.pop('mindate', None)
            kwargs.pop('maxdate', None)
            super().__init__(master, **kwargs)

        def get_date(self):
            # محاولة قراءة تاريخ بصيغة YYYY-MM-DD من النص، وإلا إرجاع الآن
            try:
                txt = self.get().strip()
                return datetime.strptime(txt, '%Y-%m-%d')
            except Exception:
                return datetime.now()

ITEMS_FILE = "items_list.txt"

# مساعدة: تنظيف أسماء شيتات Excel لجعلها صالحة (<=31 حرف، بدون أحرف ممنوعة) وضمان التفرد
def _sanitize_sheet_name(name, used_names=None, max_len=31):
    if used_names is None:
        used_names = set()

    if not isinstance(name, str):
        name = str(name)

    # إزالة الأحرف الممنوعة في أسماء sheets: : \ / ? * [ ]
    clean = re.sub(r'[:\\/\?\*\[\]]', '_', name)
    clean = clean.strip()
    if not clean:
        clean = "Sheet"

    # اقتطاع للطول الأقصى
    if len(clean) > max_len:
        clean = clean[:max_len]

    base = clean
    suffix = 1
    # ضمان التفرد ضمن used_names
    while clean in used_names:
        # احجز بعض المساحة للبادئة الرقمية
        tail = f"_{suffix}"
        avail = max_len - len(tail)
        clean = (base[:avail] + tail) if avail > 0 else (base[:max_len])
        suffix += 1

    used_names.add(clean)
    return clean

# 1. دالة جلب قائمة العملاء بأمان من أسماء الملفات فقط (تم تصحيحها لتعمل خارج IDLE)
def load_customers():
    try:
        files = glob.glob("*.xlsx")
        customer_list = []
        for file in files:
            if not file.startswith("~$") and file != "الملف_الرئيسي_المحاسبي.xlsx":
                name = os.path.splitext(file)[0]
                customer_list.append(name)
        combo_customer['values'] = customer_list
    except Exception as e:
        print("load_customers error:", e)

# 2. دالة جلب وتحديث قائمة الأصناف من ملف نصي خارجي
def load_items(new_item=None):
    items = set()
    if os.path.exists(ITEMS_FILE):
        try:
            with open(ITEMS_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        items.add(line.strip())
        except Exception as e:
            print("load_items read error:", e)
            # استمرار العمل مع مجموعة فارغة
            items = set()
            
    if new_item and new_item.strip():
        items.add(new_item.strip())
        try:
            with open(ITEMS_FILE, "w", encoding="utf-8") as f:
                for item in sorted(list(items)):
                    f.write(f"{item}\n")
        except Exception as e:
            print("load_items write error:", e)
            # لا نوقف التطبيق بسبب فشل حفظ قائمة الأصناف
            pass
            
    combo_type['values'] = sorted(list(items))

# 3. دالة قراءة وعرض جدول العميل عند اختياره
def show_customer_table(event=None):
    customer = combo_customer.get().strip()
    file_name = f"{customer}.xlsx"

    for item in tree.get_children():
        tree.delete(item)

    if not customer or not os.path.exists(file_name):
        return

    try:
        df = pd.read_excel(file_name)
        for _, row in df.iterrows():
            tree.insert("", tk.END, values=(
                row.get("التاريخ", ""),
                row.get("الإجمالي", 0),
                row.get("السعر", 0),
                row.get("الكمية", 0),
                row.get("نوع الصنف", "")
            ))
    except Exception as e:
        messagebox.showerror("خطأ", f"تعذر فتح ملف العميل: {e}")

# 4. دالة حفظ البيانات الأساسية المستقرة
def save_data():
    customer = combo_customer.get().strip()
    item_type = combo_type.get().strip()
    qty_str = ent_qty.get().strip()
    price_str = ent_price.get().strip()
    
    # جلب التاريخ بصيغة نصية YYYY-MM-DD من مربع التقويم مباشرة
    try:
        # إذا كان DateEntry حقيقياً فإنه يحتوي get_date()
        date_obj = cal_date.get_date()
        try:
            date_str = date_obj.strftime('%Y-%m-%d')
        except Exception:
            date_str = datetime.now().strftime('%Y-%m-%d')
    except Exception:
        date_str = datetime.now().strftime('%Y-%m-%d')

    if not customer or not item_type or not qty_str or not price_str:
        messagebox.showwarning("تنبيه", "الرجاء كتابة اسم العميل، الصنف، الكمية والسعر!")
        return

    try:
        qty = int(qty_str)
        price = float(price_str)
    except ValueError:
        messagebox.showerror("خطأ", "يجب إدخال أرقام صحيحة للكمية والسعر.")
        return

    total = qty * price
    file_name = f"{customer}.xlsx"

    new_row = {
        "نوع الصنف": item_type,
        "الكمية": qty,
        "السعر": price,
        "الإجمالي": total,
        "التاريخ": date_str
    }

    if os.path.exists(file_name):
        try:
            df = pd.read_excel(file_name)
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        except Exception as e:
            messagebox.showerror("خطأ", f"تأكد من إغلاق ملف إكسل العميل أولاً!\n{e}")
            return
    else:
        df = pd.DataFrame([new_row])

    try:
        df.to_excel(file_name, index=False)
        messagebox.showinfo("نجاح", "تم حفظ الصنف بنجاح.")
        
        load_items(item_type)
        load_customers()
        show_customer_table()
        
        combo_type.set('')
        ent_qty.delete(0, tk.END)
        ent_price.delete(0, tk.END)
    except Exception as e:
        messagebox.showerror("خطأ في الحفظ", f"لم يتم الحفظ: {e}")

# 5. دالة حذف السطر المحدد من الجدول ومن ملف الإكسل (تم استعادتها وتصحيحها)
def delete_selected_item():
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("تنبيه", "الرجاء تحديد السطر المراد حذفه من الجدول أولاً!")
        return
        
    # خذ العنصر الأول المحدد
    sel = selected_items[0]
    customer = combo_customer.get().strip()
    file_name = f"{customer}.xlsx"
    
    if not customer or not os.path.exists(file_name):
        return
        
    confirm = messagebox.askyesno("تأكيد الحذف", "هل أنت متأكد من رغبتك في حذف هذا الصنف نهائياً؟")
    if not confirm:
        return
        
    try:
        item_values = tree.item(sel)['values']
        target_date = str(item_values[0])
        target_total = float(item_values[1])
        target_price = float(item_values[2])
        target_qty = int(item_values[3])
        target_type = str(item_values[4])
        
        df = pd.read_excel(file_name)
        
        condition = (
            (df["نوع الصنف"].astype(str) == target_type) & 
            (df["الكمية"] == target_qty) & 
            (df["السعر"] == target_price) & 
            (df["التاريخ"].astype(str) == target_date)
        )
        
        if condition.any():
            idx_to_drop = df[condition].index
            df = df.drop(idx_to_drop).reset_index(drop=True)
            df.to_excel(file_name, index=False)
            messagebox.showinfo("نجاح", "تم حذف الصنف وتحديث الملف بنجاح.")
            
            load_customers()
            show_customer_table()
        else:
            messagebox.showerror("خطأ", "تعذر العثور على السطر المطابق في ملف الـ Excel.")
            
    except Exception as e:
        messagebox.showerror("خطأ في الحذف", f"فشل حذف الصنف: {e}")

# 6. دالة إنشاء الملف الرئيسي بصفحة لكل عميل + صفحة الملخص العام (تم تدعيمها بتنظيف أسماء الشيتات)
def make_master_file():
    files = glob.glob("*.xlsx")
    master_file = "الملف_الرئيسي_المحاسبي.xlsx"
    valid_files = [f for f in files if not f.startswith("~$") and f != master_file]
    
    if not valid_files:
        messagebox.showwarning("تنبيه", "لا توجد ملفات عملاء حالياً!")
        return

    summary = []
    used_sheet_names = set()
    try:
        with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
            for file in valid_files:
                try:
                    name = os.path.splitext(file)[0]
                    df = pd.read_excel(file)
                except Exception as e:
                    # سجل وتخطى الملفات التي تعطي أخطاء عند القراءة
                    print(f"skipping file {file} due read error: {e}")
                    continue
                
                total_sales = df["الإجمالي"].sum() if "الإجمالي" in df.columns else 0
                total_qty = df["الكمية"].sum() if "الكمية" in df.columns else 0
                
                summary.append({
                    "اسم العميل": name,
                    "إجمالي الكميات": total_qty,
                    "إجمالي الحساب": total_sales
                })
            
            # اكتب ملخص عام بصيغة آمنة لاسم الشيت
            summary_sheet = _sanitize_sheet_name("الملخص_العام", used_names=used_sheet_names)
            pd.DataFrame(summary).to_excel(writer, sheet_name=summary_sheet, index=False)
            
            # الآن أضف كل شيت عميل مع أسماء شيتات مُنقّحة ومتفردة
            for file in valid_files:
                try:
                    name = os.path.splitext(file)[0]
                    df = pd.read_excel(file)
                except Exception as e:
                    print(f"skipping file {file} due read error: {e}")
                    continue

                sheet_name = _sanitize_sheet_name(name, used_names=used_sheet_names)
                try:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    # إذا فشل الكتابة باسم الشيت (نادر)، سجّل وفرض اسم افتراضي
                    fallback = _sanitize_sheet_name(f"Sheet_{len(used_sheet_names)+1}", used_names=used_sheet_names)
                    print(f"failed to write sheet {sheet_name}, using fallback {fallback}: {e}")
                    df.to_excel(writer, sheet_name=fallback, index=False)
                
        messagebox.showinfo("نجاح", f"تم إنشاء الملف الرئيسي الشامل بنجاح باسم:\n{master_file}")
    except Exception as e:
        messagebox.showerror("خطأ", f"تأكد من إغلاق الملف الرئيسي أولاً: {e}")

# بناء الواجهة الرسومية الثابتة والمستقرة تماماً
root = tk.Tk()
root.title("النظام المحاسبي - إصدار تقويم ويندوز الرسومي")
root.geometry("400x660")

# إطار العميل
frame_c = tk.LabelFrame(root, text=" العميل ", padx=10, pady=5)
frame_c.pack(fill="x", padx=10, pady=5)

combo_customer = ttk.Combobox(frame_c, justify="right", width=30)
combo_customer.pack(side="right", padx=5)
combo_customer.bind("<<ComboboxSelected>>", show_customer_table)

lbl_c = tk.Label(frame_c, text="اسم العميل:")
lbl_c.pack(side="right")

btn_ref = tk.Button(frame_c, text="عرض الجدول", command=show_customer_table)
btn_ref.pack(side="left")

# إطار تفاصيل الصنف
frame_i = tk.LabelFrame(root, text=" تفاصيل الصنف ", padx=10, pady=5)
frame_i.pack(fill="x", padx=10, pady=5)

combo_type = ttk.Combobox(frame_i, justify="right", width=25)
combo_type.grid(row=0, column=0, pady=5, padx=5)
try:
    tk.Label(frame_i, text="نوع الصنف:").grid(row=0, column=1, sticky="e", padx=5)
except Exception:
    pass

ent_qty = tk.Entry(frame_i, justify="right", width=28)
ent_qty.grid(row=1, column=0, pady=5, padx=5)
try:
    tk.Label(frame_i, text="الكمية:").grid(row=1, column=1, sticky="e", padx=5)
except Exception:
    pass

ent_price = tk.Entry(frame_i, justify="right", width=28)
ent_price.grid(row=2, column=0, pady=5, padx=5)
try:
    tk.Label(frame_i, text="السعر لِلْوحدة:").grid(row=2, column=1, sticky="e", padx=5)
except Exception:
    pass

# مربع التقويم المالي والرسومي (DateEntry)
cal_date = DateEntry(frame_i, width=25, background='darkgreen',
                     foreground='white', borderwidth=2, 
                     year=2024, month=7, day=1,
                     mindate=datetime(2024, 7, 1), maxdate=datetime.now(),
                     date_pattern='yyyy-mm-dd', justify="center")
cal_date.grid(row=3, column=0, pady=5, padx=5)
try:
    tk.Label(frame_i, text="اختر التاريخ:").grid(row=3, column=1, sticky="e", padx=5)
except Exception:
    pass

# الأزرار الرئيسية
btn_save = tk.Button(root, text="تسجيل و حفظ", bg="#5a376e", fg="white", width=40, command=save_data)
btn_save.pack(pady=4)

btn_master = tk.Button(root, text="تحديث وتصدير الملف الرئيسي الشامل", bg="#0065a3", fg="white", width=40, command=make_master_file)
btn_master.pack(pady=4)

# جدول العرض بالواجهة
frame_t = tk.LabelFrame(root, text=" جدول حسابات العميل الحالي ")
frame_t.pack(fill="both", expand=True, padx=10, pady=5)

columns = ("التاريخ", "الإجمالي", "السعر", "الكمية", "نوع الصنف")
tree = ttk.Treeview(frame_t, columns=columns, show="headings", height=6)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=110)
tree.pack(fill="both", expand=True, padx=5, pady=5)

# زر الحذف (تم استعادته وتثبيته في أسفل الواجهة بدقة)
btn_delete = tk.Button(root, text="حذف الصنف المحدد من ملف العميل", bg="#e74c3c", fg="white", width=40, command=delete_selected_item)
btn_delete.pack(pady=8)

# استدعاء الإعدادات الأولية
load_customers()
load_items()

root.mainloop()
