"""new_pro.py
نسخة معاد تصميمها من التطبيق: إدارة عملاء وبيانات المبيعات في ملفات Excel، إضافة/حذف صفوف، وإنشاء ملف رئيسي آمن.
ميزات مضافة:
- اسم ملف بايثون صالح: new_pro.py
- بديل لـ DateEntry عند غياب tkcalendar
- تنظيف أسماء شيتات Excel وضمان التفرد
- كتابة آمنة atomically عبر ملف مؤقت + os.replace
- رسائل أخطاء أوضح وفلاتر لملفات مؤقتة (~$)
- نفس واجهة المستخدم والعمليات الأساسية كما في النسخة الأصلية

التشغيل: python new_pro.py
المتطلبات: pandas, openpyxl (tkcalendar اختياري)
"""

import os
import glob
import re
import tempfile
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
import pandas as pd

# Attempt to import tkcalendar's DateEntry; fallback to simple entry-based
try:
    from tkcalendar import DateEntry
    HAVE_TKCAL = True
except Exception:
    HAVE_TKCAL = False
    class DateEntry(ttk.Entry):
        def __init__(self, master=None, **kwargs):
            kwargs.pop('date_pattern', None)
            kwargs.pop('year', None)
            kwargs.pop('month', None)
            kwargs.pop('day', None)
            kwargs.pop('mindate', None)
            kwargs.pop('maxdate', None)
            super().__init__(master, **kwargs)

        def get_date(self):
            txt = self.get().strip()
            try:
                return datetime.strptime(txt, '%Y-%m-%d')
            except Exception:
                return datetime.now()

ITEMS_FILE = "items_list.txt"
MASTER_FILE = "الملف_الرئيسي_المحاسبي.xlsx"

# Helpers
def _sanitize_sheet_name(name, used_names=None, max_len=31):
    if used_names is None:
        used_names = set()
    if not isinstance(name, str):
        name = str(name)
    clean = re.sub(r'[:\\/\?\*\[\]]', '_', name).strip()
    if not clean:
        clean = 'Sheet'
    if len(clean) > max_len:
        clean = clean[:max_len]
    base = clean
    suffix = 1
    while clean in used_names:
        tail = f"_{suffix}"
        avail = max_len - len(tail)
        clean = (base[:avail] + tail) if avail > 0 else base[:max_len]
        suffix += 1
    used_names.add(clean)
    return clean

# UI functions

def load_customers():
    try:
        files = glob.glob('*.xlsx')
        customers = [os.path.splitext(f)[0] for f in files if not f.startswith('~$') and f != MASTER_FILE]
        combo_customer['values'] = customers
    except Exception as e:
        print('load_customers error:', e)

def load_items(new_item=None):
    items = set()
    try:
        if os.path.exists(ITEMS_FILE):
            with open(ITEMS_FILE, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.strip():
                        items.add(line.strip())
    except Exception as e:
        print('load_items read error:', e)
    if new_item and new_item.strip():
        items.add(new_item.strip())
        try:
            with open(ITEMS_FILE, 'w', encoding='utf-8') as f:
                for it in sorted(items):
                    f.write(it + '\n')
        except Exception as e:
            print('load_items write error:', e)
    combo_type['values'] = sorted(items)

def show_customer_table(event=None):
    customer = combo_customer.get().strip()
    file_name = f"{customer}.xlsx"
    for it in tree.get_children():
        tree.delete(it)
    if not customer or not os.path.exists(file_name):
        return
    try:
        df = pd.read_excel(file_name)
        for _, row in df.iterrows():
            tree.insert('', tk.END, values=(
                row.get('التاريخ', ''),
                row.get('الإجمالي', 0),
                row.get('السعر', 0),
                row.get('الكمية', 0),
                row.get('نوع الصنف', '')
            ))
    except Exception as e:
        messagebox.showerror('خطأ', f'تعذر فتح ملف العميل: {e}')

def save_data():
    customer = combo_customer.get().strip()
    item_type = combo_type.get().strip()
    qty_str = ent_qty.get().strip()
    price_str = ent_price.get().strip()
    try:
        date_obj = cal_date.get_date()
        date_str = date_obj.strftime('%Y-%m-%d')
    except Exception:
        date_str = datetime.now().strftime('%Y-%m-%d')
    if not customer or not item_type or not qty_str or not price_str:
        messagebox.showwarning('تنبيه', 'الرجاء كتابة اسم العميل، الصنف، الكمية والسعر!')
        return
    try:
        qty = int(qty_str)
        price = float(price_str)
    except ValueError:
        messagebox.showerror('خطأ', 'يجب إدخال أرقام صحيحة للكمية والسعر.')
        return
    total = qty * price
    file_name = f"{customer}.xlsx"
    new_row = {'نوع الصنف': item_type, 'الكمية': qty, 'السعر': price, 'الإجمالي': total, 'التاريخ': date_str}
    try:
        if os.path.exists(file_name):
            df = pd.read_excel(file_name)
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            df = pd.DataFrame([new_row])
        # atomic write
        tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(tmp_fd)
        try:
            with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            try:
                os.replace(tmp_path, file_name)
            except PermissionError:
                os.remove(tmp_path)
                raise PermissionError('ملف العميل مفتوح؛ الرجاء إغلاق الملف ثم المحاولة مرة أخرى.')
        finally:
            if os.path.exists(tmp_path):
                try: os.remove(tmp_path)
                except: pass
        messagebox.showinfo('نجاح', 'تم حفظ الصنف بنجاح.')
        load_items(item_type)
        load_customers()
        show_customer_table()
        combo_type.set('')
        ent_qty.delete(0, tk.END)
        ent_price.delete(0, tk.END)
    except Exception as e:
        messagebox.showerror('خطأ في الحفظ', f'لم يتم الحفظ: {e}')

def delete_selected_item():
    sel = tree.selection()
    if not sel:
        messagebox.showwarning('تنبيه', 'الرجاء تحديد السطر المراد حذفه من الجدول أولاً!')
        return
    item = sel[0]
    customer = combo_customer.get().strip()
    file_name = f"{customer}.xlsx"
    if not customer or not os.path.exists(file_name):
        return
    if not messagebox.askyesno('تأكيد الحذف', 'هل أنت متأكد من رغبتك في حذف هذا الصنف نهائياً؟'):
        return
    try:
        vals = tree.item(item)['values']
        target_date = str(vals[0])
        target_total = float(vals[1])
        target_price = float(vals[2])
        target_qty = int(vals[3])
        target_type = str(vals[4])
        df = pd.read_excel(file_name)
        condition = (
            (df['نوع الصنف'].astype(str) == target_type) &
            (df['الكمية'] == target_qty) &
            (df['السعر'] == target_price) &
            (df['التاريخ'].astype(str) == target_date)
        )
        if condition.any():
            idx = df[condition].index
            df = df.drop(idx).reset_index(drop=True)
            # atomic save
            tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
            os.close(tmp_fd)
            try:
                with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                try:
                    os.replace(tmp_path, file_name)
                except PermissionError:
                    os.remove(tmp_path)
                    raise PermissionError('ملف العميل مفتوح؛ الرجاء إغلاق الملف ثم المحاولة مرة أخرى.')
            finally:
                if os.path.exists(tmp_path):
                    try: os.remove(tmp_path)
                    except: pass
            messagebox.showinfo('نجاح', 'تم حذف الصنف وتحديث الملف بنجاح.')
            load_customers()
            show_customer_table()
        else:
            messagebox.showerror('خطأ', 'تعذر العثور على السطر المطابق في ملف الـ Excel.')
    except Exception as e:
        messagebox.showerror('خطأ في الحذف', f'فشل حذف الصنف: {e}')

def make_master_file():
    files = glob.glob('*.xlsx')
    valid_files = [f for f in files if not f.startswith('~$') and f != MASTER_FILE]
    if not valid_files:
        messagebox.showwarning('تنبيه', 'لا توجد ملفات عملاء حالياً!')
        return
    summary = []
    used_names = set()
    # prepare sheets content
    sheets = []
    for fpath in valid_files:
        try:
            name = os.path.splitext(fpath)[0]
            df = pd.read_excel(fpath)
        except Exception as e:
            print(f'skipping {fpath} due read error: {e}')
            continue
        total_sales = df['الإجمالي'].sum() if 'الإجمالي' in df.columns else 0
        total_qty = df['الكمية'].sum() if 'الكمية' in df.columns else 0
        summary.append({'اسم العميل': name, 'إجمالي الكميات': total_qty, 'إجمالي الحساب': total_sales})
        sheets.append((name, df))
    # atomic write to temp file
    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(tmp_fd)
    try:
        with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
            summary_df = pd.DataFrame(summary)
            summary_sheet = _sanitize_sheet_name('الملخص_العام', used_names)
            summary_df.to_excel(writer, sheet_name=summary_sheet, index=False)
            for name, df in sheets:
                sheet_name = _sanitize_sheet_name(name, used_names)
                try:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    fallback = _sanitize_sheet_name(f'Sheet_{len(used_names)+1}', used_names)
                    print(f'failed to write sheet {sheet_name}, using fallback {fallback}: {e}')
                    df.to_excel(writer, sheet_name=fallback, index=False)
        try:
            os.replace(tmp_path, MASTER_FILE)
        except PermissionError:
            os.remove(tmp_path)
            messagebox.showerror('خطأ', 'الملف الرئيسي مفتوح؛ الرجاء إغلاق الملف ثم المحاولة مرة أخرى.')
            return
        messagebox.showinfo('نجاح', f'تم إنشاء الملف الرئيسي الشامل بنجاح باسم:\n{MASTER_FILE}')
    except Exception as e:
        if os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except: pass
        messagebox.showerror('خطأ', f'فشل إنشاء الملف الرئيسي: {e}')

# Build UI
root = tk.Tk()
root.title('النظام المحاسبي - new pro')
root.geometry('480x720')

frame_c = tk.LabelFrame(root, text=' العميل ', padx=10, pady=5)
frame_c.pack(fill='x', padx=10, pady=5)
combo_customer = ttk.Combobox(frame_c, justify='right', width=34)
combo_customer.pack(side='right', padx=5)
combo_customer.bind('<<ComboboxSelected>>', show_customer_table)

lbl_c = tk.Label(frame_c, text='اسم العميل:')
lbl_c.pack(side='right')
btn_ref = tk.Button(frame_c, text='عرض الجدول', command=show_customer_table)
btn_ref.pack(side='left')

frame_i = tk.LabelFrame(root, text=' تفاصيل الصنف ', padx=10, pady=5)
frame_i.pack(fill='x', padx=10, pady=5)

combo_type = ttk.Combobox(frame_i, justify='right', width=30)
combo_type.grid(row=0, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='نوع الصنف:').grid(row=0, column=1, sticky='e', padx=5)

ent_qty = tk.Entry(frame_i, justify='right', width=28)
ent_qty.grid(row=1, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='الكمية:').grid(row=1, column=1, sticky='e', padx=5)

ent_price = tk.Entry(frame_i, justify='right', width=28)
ent_price.grid(row=2, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='السعر لِلْوحدة:').grid(row=2, column=1, sticky='e', padx=5)

cal_date = DateEntry(frame_i, width=28, background='darkgreen', foreground='white', borderwidth=2,
                     year=2024, month=7, day=1, mindate=datetime(2024,7,1), maxdate=datetime.now(),
                     date_pattern='yyyy-mm-dd', justify='center')
cal_date.grid(row=3, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='اختر التاريخ:').grid(row=3, column=1, sticky='e', padx=5)

btn_save = tk.Button(root, text='تسجيل و حفظ', bg='#5a376e', fg='white', width=40, command=save_data)
btn_save.pack(pady=6)

btn_master = tk.Button(root, text='تحديث وتصدير الملف الرئيسي الشامل', bg='#0065a3', fg='white', width=40, command=make_master_file)
btn_master.pack(pady=6)

frame_t = tk.LabelFrame(root, text=' جدول حسابات العميل الحالي ')
frame_t.pack(fill='both', expand=True, padx=10, pady=5)

columns = ('التاريخ', 'الإجمالي', 'السعر', 'الكمية', 'نوع الصنف')
tree = ttk.Treeview(frame_t, columns=columns, show='headings', height=8)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor='center', width=120)
tree.pack(fill='both', expand=True, padx=5, pady=5)

btn_delete = tk.Button(root, text='حذف الصنف المحدد من ملف العميل', bg='#e74c3c', fg='white', width=40, command=delete_selected_item)
btn_delete.pack(pady=8)

# Initialize
load_customers()
load_items()

root.mainloop()
