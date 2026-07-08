"""new_pro.py
نسخة معاد تصميمها من التطبيق: إدارة عملاء وبيانات المبيعات في ملفات Excel، إضافة/حذف صفوف، وإنشاء ملف رئيسي آمن.
تحسينات إضافية في هذا التعديل:
- سجل تفصيلي باستخدام logging إلى ملف app.log
- شريط حالة أسفل الواجهة لعرض رسائل فورية مع تسجيلها
- شريط قوائم بسيط (ملف، مساعدة)
- اختصارات لوحة المفاتيح: Ctrl+S لحفظ، Delete لحذف السطر المحدد، Ctrl+M لإنشاء الملف الرئيسي
- شريط تمرير رأسي لعرض Treeview
- فحص اسم العميل قبل الحفظ للتأكد من صلاحية اسم الملف
- تسجيل الاستثناءات مع traceback لمساعدة التصحيح

التشغيل: python new_pro.py
المتطلبات: pandas, openpyxl (tkcalendar اختياري)
"""

import os
import glob
import re
import tempfile
import logging
import traceback
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
import pandas as pd

# Logging setup
LOG_FILE = 'app.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

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
INVALID_FILENAME_CHARS = r'[:\\/\?\*\[\]<>|]'

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


def _valid_filename(name):
    if not name or not name.strip():
        return False
    # disallow only-space names and invalid chars
    if re.search(INVALID_FILENAME_CHARS, name):
        return False
    # keep it short
    if len(name) > 200:
        return False
    return True

# UI functions

def set_status(msg, level='info'):
    try:
        status_var.set(msg)
    except Exception:
        pass
    if level == 'error':
        logging.error(msg)
    else:
        logging.info(msg)


def load_customers():
    try:
        files = glob.glob('*.xlsx')
        customers = [os.path.splitext(f)[0] for f in files if not f.startswith('~$') and f != MASTER_FILE]
        combo_customer['values'] = customers
        set_status(f'Loaded {len(customers)} customers')
    except Exception as e:
        logging.exception('load_customers error')
        set_status('خطأ في تحميل العملاء', level='error')


def load_items(new_item=None):
    items = set()
    try:
        if os.path.exists(ITEMS_FILE):
            with open(ITEMS_FILE, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.strip():
                        items.add(line.strip())
    except Exception:
        logging.exception('load_items read error')
        set_status('خطأ في قراءة قائمة الأصناف', level='error')
    if new_item and new_item.strip():
        items.add(new_item.strip())
        try:
            with open(ITEMS_FILE, 'w', encoding='utf-8') as f:
                for it in sorted(items):
                    f.write(it + '\n')
        except Exception:
            logging.exception('load_items write error')
            set_status('فشل حفظ قائمة الأصناف', level='error')
    combo_type['values'] = sorted(items)


def show_customer_table(event=None):
    customer = combo_customer.get().strip()
    file_name = f"{customer}.xlsx"
    for it in tree.get_children():
        tree.delete(it)
    if not customer:
        set_status('لم يتم اختيار عميل')
        return
    if not os.path.exists(file_name):
        set_status('ملف العميل غير موجود')
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
        set_status(f'عرض {len(df)} صفوف من {customer}')
        logging.info('showed table for %s (%d rows)', customer, len(df))
    except Exception:
        logging.exception('Failed to show customer table')
        messagebox.showerror('خطأ', 'تعذر فتح ملف العميل — تحقق من الملف أو صلاحيات الوصول')
        set_status('فشل عرض جدول العميل', level='error')


def save_data(event=None):
    try:
        customer = combo_customer.get().strip()
        item_type = combo_type.get().strip()
        qty_str = ent_qty.get().strip()
        price_str = ent_price.get().strip()
        try:
            date_obj = cal_date.get_date()
            date_str = date_obj.strftime('%Y-%m-%d')
        except Exception:
            date_str = datetime.now().strftime('%Y-%m-%d')

        if not _valid_filename(customer):
            messagebox.showwarning('تنبيه', 'اسم العميل غير صالح كاسم ملف. اختر اسماً أبسط دون أحرف مثل : \/?*[]')
            set_status('اسم العميل غير صالح', level='error')
            return

        if not customer or not item_type or not qty_str or not price_str:
            messagebox.showwarning('تنبيه', 'الرجاء كتابة اسم العميل، الصنف، الكمية والسعر!')
            set_status('حقول ناقصة', level='error')
            return
        try:
            qty = int(qty_str)
            price = float(price_str)
        except ValueError:
            messagebox.showerror('خطأ', 'يجب إدخال أرقام صحيحة للكمية والسعر.')
            set_status('قيمة كمية/سعر غير صحيحة', level='error')
            return
        total = qty * price
        file_name = f"{customer}.xlsx"
        new_row = {'نوع الصنف': item_type, 'الكمية': qty, 'السعر': price, 'الإجمالي': total, 'التاريخ': date_str}
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
                messagebox.showerror('خطأ', 'ملف العميل مفتوح؛ الرجاء إغلاق الملف ثم المحاولة مرة أخرى.')
                set_status('الملف مفتوح - أغلقه وحاول مرة أخرى', level='error')
                return
        finally:
            if os.path.exists(tmp_path):
                try: os.remove(tmp_path)
                except: pass
        messagebox.showinfo('نجاح', 'تم حفظ الصنف بنجاح.')
        logging.info('Saved row for %s: %s x %s', customer, qty, item_type)
        load_items(item_type)
        load_customers()
        show_customer_table()
        combo_type.set('')
        ent_qty.delete(0, tk.END)
        ent_price.delete(0, tk.END)
        set_status('تم الحفظ بنجاح')
    except Exception:
        logging.exception('save_data failed')
        messagebox.showerror('خطأ في الحفظ', 'تعذر حفظ البيانات — راجع السجل app.log')
        set_status('فشل الحفظ', level='error')


def delete_selected_item(event=None):
    try:
        sel = tree.selection()
        if not sel:
            messagebox.showwarning('تنبيه', 'الرجاء تحديد السطر المراد حذفه من الجدول أولاً!')
            set_status('لم يتم تحديد صف للحذف', level='error')
            return
        item = sel[0]
        customer = combo_customer.get().strip()
        file_name = f"{customer}.xlsx"
        if not customer or not os.path.exists(file_name):
            set_status('ملف العميل غير موجود للحذف', level='error')
            return
        if not messagebox.askyesno('تأكيد الحذف', 'هل أنت متأكد من رغبتك في حذف هذا الصنف نهائياً؟'):
            set_status('إلغاء عملية الحذف')
            return
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
                    messagebox.showerror('خطأ', 'ملف العميل مفتوح؛ الرجاء إغلاقه ثم المحاولة مرة أخرى.')
                    set_status('الملف مفتوح - فشل الحذف', level='error')
                    return
            finally:
                if os.path.exists(tmp_path):
                    try: os.remove(tmp_path)
                    except: pass
            messagebox.showinfo('نجاح', 'تم حذف الصنف وتحديث الملف بنجاح.')
            logging.info('Deleted row for %s: %s x %s', customer, target_qty, target_type)
            load_customers()
            show_customer_table()
            set_status('تم الحذف بنجاح')
        else:
            messagebox.showerror('خطأ', 'تعذر العثور على السطر المطابق في ملف الـ Excel.')
            set_status('السطر غير موجود في الملف', level='error')
    except Exception:
        logging.exception('delete_selected_item failed')
        messagebox.showerror('خطأ في الحذف', 'فشل حذف الصنف — راجع السجل app.log')
        set_status('فشل الحذف', level='error')


def make_master_file(event=None):
    try:
        files = glob.glob('*.xlsx')
        valid_files = [f for f in files if not f.startswith('~$') and f != MASTER_FILE]
        if not valid_files:
            messagebox.showwarning('تنبيه', 'لا توجد ملفات عملاء حالياً!')
            set_status('لا ملفات لإنشاء الملف الرئيسي')
            return
        summary = []
        used_names = set()
        # prepare sheets content
        sheets = []
        for fpath in valid_files:
            try:
                name = os.path.splitext(fpath)[0]
                df = pd.read_excel(fpath)
            except Exception:
                logging.exception('skipping file due read error: %s', fpath)
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
                    except Exception:
                        logging.exception('failed to write sheet %s', sheet_name)
                        fallback = _sanitize_sheet_name(f'Sheet_{len(used_names)+1}', used_names)
                        df.to_excel(writer, sheet_name=fallback, index=False)
            try:
                os.replace(tmp_path, MASTER_FILE)
            except PermissionError:
                os.remove(tmp_path)
                messagebox.showerror('خطأ', 'الملف الرئيسي مفتوح؛ الرجاء إغلاق الملف ثم المحاولة مرة أخرى.')
                set_status('الملف الرئيسي مفتوح - أعد المحاولة بعد الإغلاق', level='error')
                return
            messagebox.showinfo('نجاح', f'تم إنشاء الملف الرئيسي الشامل بنجاح باسم:\n{MASTER_FILE}')
            logging.info('Created master file: %s', MASTER_FILE)
            set_status('تم إنشاء الملف الرئيسي')
        finally:
            if os.path.exists(tmp_path):
                try: os.remove(tmp_path)
                except: pass
    except Exception:
        logging.exception('make_master_file failed')
        messagebox.showerror('خطأ', 'فشل إنشاء الملف الرئيسي — راجع السجل app.log')
        set_status('فشل إنشاء الملف الرئيسي', level='error')

# Build UI
root = tk.Tk()
root.title('النظام المحاسبي - new pro')
root.geometry('640x720')

# Menu
menubar = tk.Menu(root)
filemenu = tk.Menu(menubar, tearoff=0)
filemenu.add_command(label='تحديث الملف الرئيسي (Ctrl+M)', command=make_master_file)
filemenu.add_separator()
filemenu.add_command(label='خروج', command=root.quit)
menubar.add_cascade(label='ملف', menu=filemenu)
helpmenu = tk.Menu(menubar, tearoff=0)
helpmenu.add_command(label='حول', command=lambda: messagebox.showinfo('حول', 'new_pro - نظام محاسبي مبسط'))
menubar.add_cascade(label='مساعدة', menu=helpmenu)
root.config(menu=menubar)

frame_c = tk.LabelFrame(root, text=' العميل ', padx=10, pady=5)
frame_c.pack(fill='x', padx=10, pady=5)
combo_customer = ttk.Combobox(frame_c, justify='right', width=40)
combo_customer.pack(side='right', padx=5)
combo_customer.bind('<<ComboboxSelected>>', show_customer_table)

lbl_c = tk.Label(frame_c, text='اسم العميل:')
lbl_c.pack(side='right')
btn_ref = tk.Button(frame_c, text='عرض الجدول', command=show_customer_table)
btn_ref.pack(side='left')

frame_i = tk.LabelFrame(root, text=' تفاصيل الصنف ', padx=10, pady=5)
frame_i.pack(fill='x', padx=10, pady=5)

combo_type = ttk.Combobox(frame_i, justify='right', width=36)
combo_type.grid(row=0, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='نوع الصنف:').grid(row=0, column=1, sticky='e', padx=5)

ent_qty = tk.Entry(frame_i, justify='right', width=28)
ent_qty.grid(row=1, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='الكمية:').grid(row=1, column=1, sticky='e', padx=5)

ent_price = tk.Entry(frame_i, justify='right', width=28)
ent_price.grid(row=2, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='السعر لِلْوحدة:').grid(row=2, column=1, sticky='e', padx=5)

cal_date = DateEntry(frame_i, width=36, background='darkgreen', foreground='white', borderwidth=2,
                     year=2024, month=7, day=1, mindate=datetime(2024,7,1), maxdate=datetime.now(),
                     date_pattern='yyyy-mm-dd', justify='center')
cal_date.grid(row=3, column=0, pady=5, padx=5)
 tk.Label(frame_i, text='اختر التاريخ:').grid(row=3, column=1, sticky='e', padx=5)

btn_save = tk.Button(root, text='تسجيل و حفظ (Ctrl+S)', bg='#5a376e', fg='white', width=40, command=save_data)
btn_save.pack(pady=6)

btn_master = tk.Button(root, text='تحديث وتصدير الملف الرئيسي الشامل (Ctrl+M)', bg='#0065a3', fg='white', width=40, command=make_master_file)
btn_master.pack(pady=6)

frame_t = tk.LabelFrame(root, text=' جدول حسابات العميل الحالي ')
frame_t.pack(fill='both', expand=True, padx=10, pady=5)

# Treeview with scrollbar
columns = ('التاريخ', 'الإجمالي', 'السعر', 'الكمية', 'نوع الصنف')
container = tk.Frame(frame_t)
container.pack(fill='both', expand=True)
tree = ttk.Treeview(container, columns=columns, show='headings', height=12)
vsb = ttk.Scrollbar(container, orient='vertical', command=tree.yview)
tree.configure(yscrollcommand=vsb.set)
vsb.pack(side='right', fill='y')
tree.pack(fill='both', expand=True, side='left')
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor='center', width=120)

btn_delete = tk.Button(root, text='حذف الصنف المحدد من ملف العميل (Delete)', bg='#e74c3c', fg='white', width=40, command=delete_selected_item)
btn_delete.pack(pady=8)

# Status bar
status_var = tk.StringVar()
status_bar = tk.Label(root, textvariable=status_var, bd=1, relief='sunken', anchor='w')
status_bar.pack(side='bottom', fill='x')

# Keyboard shortcuts
root.bind('<Control-s>', save_data)
root.bind('<Control-S>', save_data)
root.bind('<Delete>', delete_selected_item)
root.bind('<Control-m>', make_master_file)
root.bind('<Control-M>', make_master_file)

# Initialize
load_customers()
load_items()
set_status('جاهز')

root.mainloop()
