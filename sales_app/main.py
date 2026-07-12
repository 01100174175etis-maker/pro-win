"""Main GUI for sales_app
Uses ttkbootstrap for a modern look and tkcalendar for date picking.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from tkcalendar import DateEntry
from datetime import datetime
import pandas as pd
import os
from pathlib import Path
import sales_app.db as db

APP_TITLE = "Sales Manager"

class SalesApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x600")
        # Styles are provided by ttkbootstrap
        self.style = tb.Style()

        # Top frame: form
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill='x')

        ttk.Label(frm, text="Customer:").grid(row=0, column=0, sticky='e')
        self.ent_customer = ttk.Entry(frm, width=30)
        self.ent_customer.grid(row=0, column=1, padx=6, pady=6)

        ttk.Label(frm, text="Item:").grid(row=0, column=2, sticky='e')
        self.ent_item = ttk.Entry(frm, width=30)
        self.ent_item.grid(row=0, column=3, padx=6, pady=6)

        ttk.Label(frm, text="Quantity:").grid(row=1, column=0, sticky='e')
        self.ent_qty = ttk.Entry(frm, width=10)
        self.ent_qty.grid(row=1, column=1, sticky='w', padx=6, pady=6)

        ttk.Label(frm, text="Date:").grid(row=1, column=2, sticky='e')
        self.date_entry = DateEntry(frm, width=20, date_pattern='yyyy-mm-dd')
        self.date_entry.set_date(datetime.now())
        self.date_entry.grid(row=1, column=3, sticky='w', padx=6, pady=6)

        self.btn_add = ttk.Button(frm, text="Add Sale", command=self.add_sale)
        self.btn_add.grid(row=0, column=4, rowspan=2, padx=10)

        # Middle: treeview
        mid = ttk.Frame(self.root)
        mid.pack(fill='both', expand=True, padx=10, pady=10)

        columns = ("id", "customer", "item", "qty", "sale_date", "created_at")
        self.tree = ttk.Treeview(mid, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='center')
        # hide id small
        self.tree.column('id', width=50)
        self.tree.column('customer', width=180)
        self.tree.column('item', width=180)
        self.tree.column('qty', width=80)
        self.tree.column('sale_date', width=120)
        self.tree.column('created_at', width=180)

        vsb = ttk.Scrollbar(mid, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)

        # Bottom controls
        bot = ttk.Frame(self.root, padding=8)
        bot.pack(fill='x')

        self.btn_delete = ttk.Button(bot, text='Delete Selected', command=self.delete_selected)
        self.btn_delete.pack(side='left')

        self.btn_export = ttk.Button(bot, text='Export to Excel', command=self.export_excel)
        self.btn_export.pack(side='left', padx=8)

        self.status_var = tk.StringVar(value='Ready')
        self.status = ttk.Label(self.root, textvariable=self.status_var, anchor='w')
        self.status.pack(fill='x', side='bottom')

        self.load_data()

    def set_status(self, text: str):
        self.status_var.set(text)

    def load_data(self):
        df = db.get_sales()
        # clear
        for r in self.tree.get_children():
            self.tree.delete(r)
        for _, row in df.iterrows():
            self.tree.insert('', 'end', values=(row['id'], row['customer'], row['item'], row['qty'], row['sale_date'], row['created_at']))
        self.set_status(f'Loaded {len(df)} records')

    def add_sale(self):
        customer = self.ent_customer.get().strip()
        item = self.ent_item.get().strip()
        qty_text = self.ent_qty.get().strip()
        date_text = self.date_entry.get_date().strftime('%Y-%m-%d')
        if not customer or not item or not qty_text:
            messagebox.showwarning('Missing', 'Please fill customer, item and quantity')
            return
        try:
            qty = int(qty_text)
        except ValueError:
            messagebox.showerror('Invalid', 'Quantity must be integer')
            return
        db.add_sale(customer, item, qty, date_text)
        self.load_data()
        self.ent_item.delete(0, 'end')
        self.ent_qty.delete(0, 'end')
        self.set_status('Sale added')

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = self.tree.item(item, 'values')
        sale_id = int(vals[0])
        if messagebox.askyesno('Confirm', f'Delete sale id {sale_id}?'):
            db.delete_sale(sale_id)
            self.load_data()
            self.set_status('Deleted')

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files','*.xlsx')])
        if not path:
            return
        try:
            db.export_to_excel(path)
            messagebox.showinfo('Exported', f'Exported to {path}')
            self.set_status(f'Exported to {os.path.basename(path)}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed export: {e}')
            self.set_status('Export failed')


def main():
    db.init_db()
    root = tb.Window(title=APP_TITLE) if hasattr(tb, 'Window') else tb.Style().master
    # If ttkbootstrap Window available, otherwise fallback to Tk
    try:
        app = SalesApp(root)
        root.mainloop()
    except Exception as e:
        import traceback; traceback.print_exc()

if __name__ == '__main__':
    main()
