"""Database helper for sales_app
Provides simple sqlite3-backed storage and export to Excel.
"""
import sqlite3
from pathlib import Path
import pandas as pd
import logging

DB_FILENAME = Path(__file__).parent / "sales.db"

logger = logging.getLogger(__name__)

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS sales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer TEXT NOT NULL,
    item TEXT NOT NULL,
    qty INTEGER NOT NULL,
    sale_date TEXT NOT NULL,
    created_at TEXT NOT NULL
);
"""


def init_db(db_path: Path = DB_FILENAME):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute(CREATE_TABLE_SQL)
    conn.commit()
    conn.close()
    logger.info("Initialized database at %s", db_path)


def add_sale(customer: str, item: str, qty: int, sale_date: str, db_path: Path = DB_FILENAME):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO sales (customer, item, qty, sale_date, created_at) VALUES (?, ?, ?, ?, datetime('now'))",
        (customer, item, qty, sale_date),
    )
    conn.commit()
    inserted_id = cursor.lastrowid
    conn.close()
    logger.info("Added sale id=%s customer=%s item=%s qty=%s date=%s", inserted_id, customer, item, qty, sale_date)
    return inserted_id


def get_sales(db_path: Path = DB_FILENAME):
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM sales ORDER BY id DESC", conn)
    conn.close()
    return df


def delete_sale(sale_id: int, db_path: Path = DB_FILENAME):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sales WHERE id = ?", (sale_id,))
    conn.commit()
    rows = cursor.rowcount
    conn.close()
    logger.info("Deleted sale id=%s (rows=%s)", sale_id, rows)
    return rows


def export_to_excel(path: str, db_path: Path = DB_FILENAME):
    df = get_sales(db_path)
    # Convert sale_date to string if needed and drop internal created_at
    out_df = df[['id', 'customer', 'item', 'qty', 'sale_date', 'created_at']]
    out_df.to_excel(path, index=False)
    logger.info("Exported %d rows to %s", len(out_df), path)
    return path
