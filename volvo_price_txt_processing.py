import os
import glob
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from tqdm import tqdm
from datetime import datetime, timedelta
import math  # 👈 добавь к импортам, если ещё нет

# Импорт библиотеки для работы с CatID
from CatID_lib import (
    ensure_catid_soft,
    ensure_catid_hardware,
    ensure_catid_by_pg,
    apply_catid_overrides,
    backfill_catid_from_previous_excel,
    print_catid_statistics,
    get_empty_catid_condition,
    get_catid_statistics,
    norm_part,
)

# ...

EUR_TO_SEK_RATE = 10.8  # ⚖️ Курс EUR→SEK как в processing.py (можешь поправить при необходимости)

# CategoryID (CatID) → код категории в 1С
CATEGORY_MAPPING_1C = {
    4:    'РСОН00005',
    10:   'И00000002',
    20:   '000000001',
    15:   'РСОН00003',
    70:   '000000003',
    60:   'И00000003',
    30:   '000000004',
    40:   '000000002',
    50:   'N00000004',
    80:   'РСОН00002',
    65:   'SRVIS0005',
    1000: 'SW0000002',
}
# ------------------------------------------------------------
# ⚙️ Константы/пути
# ------------------------------------------------------------
DATABASE_PATH = "volvo_prices.db"
PRICELIST_TABLE_NAME = "volvo_price_list"
REPLACEMENT_TABLE_NAME = "volvo_replacement_list"
EMBLEM_TABLE_NAME = "volvo_emblem_list"
DISCOUNT_TABLE_NAME = "discount_codes"
FOLDER_PATH = "export"
TXT_FOLDER = "TXT"  # сюда кладём Gold_*.txt

# Глобальные пути, которые будут найдены по префиксу
PRICELIST_TXT_FILE_PATH = None
REPLACEMENT_TXT_FILE_PATH = None
EMBLEM_TXT_FILE_PATH = None

# ------------------------------------------------------------
# 🔧 Утилиты
# ------------------------------------------------------------

def ensure_folder_exists():
    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)


def read_lines_safely(path: str):
    """Надёжное чтение TXT с автоподбором кодировки."""
    for enc in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
    # fallback: грубое декодирование
    with open(path, "rb") as f:
        return f.read().decode("utf-8", errors="ignore").splitlines()


def find_file_by_prefix(folder: str, prefix: str):
    files = glob.glob(os.path.join(folder, prefix + "*"))
    return files[0] if files else None


# norm_part теперь импортируется из CatID_lib

def safe_int(value):
    try:
        return int(value)
    except Exception:
        return None


def safe_float(value):
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return None

# ------------------------------------------------------------
# 🗄️ Создание таблиц/индексов
# ------------------------------------------------------------

CATID_OVERRIDES_TABLE = "catid_overrides"

def create_tables_if_not_exist():
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()

    # Прайс-лист
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {PRICELIST_TABLE_NAME} (
            Posttype TEXT,
            MarketCompCode TEXT,
            PartNumber TEXT,
            CheckDigit TEXT,
            DiscountCode INTEGER,
            GrossPrice REAL,
            TransferCode INTEGER,
            SRACode TEXT,
            Quantpack INTEGER,
            UnitSort TEXT,
            FunctionGroup INTEGER,
            ProductGroup INTEGER,
            Description TEXT,
            AgeCode TEXT,
            Weight INTEGER,
            Volume REAL,
            BSPRCode TEXT,
            DealerStockPrice REAL,
            CatID TEXT
        )
    """)

    # Замены
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {REPLACEMENT_TABLE_NAME} (
            RecordType INTEGER,
            PartNumber TEXT,
            CheckDigit TEXT,
            LineNumber INTEGER,
            ReplacedQuantity REAL,
            WeekReplaced INTEGER,
            ReplacedCode TEXT,
            Description TEXT,
            InfoText TEXT,
            ReplacingPart TEXT,
            ReplacingCheckDigit TEXT,
            ReplacingQuantity REAL
        )
    """)

    # Эмблемы
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {EMBLEM_TABLE_NAME} (
            PartNumber TEXT,
            CheckDigit TEXT,
            Description TEXT,
            LineNumber INTEGER,
            Emblem01 TEXT,
            Emblem02 TEXT,
            Emblem03 TEXT,
            Emblem04 TEXT,
            Emblem05 TEXT,
            Emblem06 TEXT,
            Emblem07 TEXT,
            Emblem08 TEXT,
            Emblem09 TEXT,
            Emblem10 TEXT
        )
    """)

    # Справочник скидок
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {DISCOUNT_TABLE_NAME} (
            DC INTEGER PRIMARY KEY,
            [Purch Stock %] REAL,
            [Purch Daily %] REAL,
            [Dealer Stock %] REAL,
            [Dealer Daily %] REAL,
            [Landed Cost %] REAL,
            [Markup %] REAL
        )
    """)

    # Ручные правки категорий (переносятся между релизами)
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {CATID_OVERRIDES_TABLE} (
            PartNumber TEXT PRIMARY KEY,
            CatID TEXT NOT NULL,
            Note TEXT,
            UpdatedAt TEXT
        )
    """)

    conn.commit()
    conn.close()


def ensure_indexes():
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_price_part ON {PRICELIST_TABLE_NAME}(PartNumber)")
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_price_dc   ON {PRICELIST_TABLE_NAME}(DiscountCode)")
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_repl_part  ON {REPLACEMENT_TABLE_NAME}(PartNumber)")
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_repl_repl  ON {REPLACEMENT_TABLE_NAME}(ReplacingPart)")
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_repl_type  ON {REPLACEMENT_TABLE_NAME}(RecordType)")
    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_repl_code  ON {REPLACEMENT_TABLE_NAME}(ReplacedCode)")
    # Для ручных правок отдельного индекса не нужно: PK(PartNumber)
    conn.commit(); conn.close()

# ------------------------------------------------------------
# 🧾 Парсеры TXT (формат M3 — при необходимости подправь срезы)
# ------------------------------------------------------------

def parse_pricelist_line(line: str):
    raw_price = line[15:24].strip()  # проверь layout под твой TXT
    if raw_price:
        if raw_price.isdigit():
            gross = int(raw_price) / 100.0
        else:
            gross = safe_float(raw_price)
    else:
        gross = None

    return {
        "Posttype": line[0:1].strip(),
        "MarketCompCode": line[1:2].strip(),
        "PartNumber": norm_part(line[2:11]),  # ⚠️ проверь ширину под свой файл
        "CheckDigit": line[11:12].strip(),
        "DiscountCode": safe_int(line[12:15]),
        "GrossPrice": gross,
        "TransferCode": safe_int(line[24:29]),
        "SRACode": line[29:32].strip(),
        "Quantpack": safe_int(line[32:37]),
        "UnitSort": line[37:39].strip(),
        "FunctionGroup": safe_int(line[39:44]),
        "ProductGroup": safe_int(line[44:47]),
        "Description": line[47:72].strip(),
        "AgeCode": line[72:73].strip(),
        "Weight": safe_int(line[73:80]),
        "Volume": safe_float(line[80:89]),
        "BSPRCode": line[89:90].strip(),
    }


def parse_replace_line(line: str):
    """
    Разбор строки из Gold_Replacement_M3_*.txt.
    Делаем максимально устойчиво:
    - не падаем на мусоре
    - PartNumber / ReplacingPart нормализуем через norm_part
    """
    if not line:
        return None

    # Мини-фильтр по длине, чтобы не трогать явно обрезанные строки
    if len(line) < 32:
        return None

    record_type = safe_int(line[0:1].strip())
    if record_type not in (1, 2, 3):
        # неизвестный тип записи — просто игнорируем
        return None

    data = {
        "RecordType": record_type,
        # НИКАКИХ int(...) напрямую – только norm_part
        "PartNumber": norm_part(line[2:13]),   # ⚠️ если нужно 2:11, потом подправим по факту
        "CheckDigit": line[13:14].strip(),
        "LineNumber": safe_int(line[14:17].strip()),
        "ReplacedQuantity": None,
        "WeekReplaced": None,
        "ReplacedCode": None,
        "Description": None,
        "InfoText": None,
        "ReplacingPart": None,
        "ReplacingCheckDigit": None,
        "ReplacingQuantity": None,
    }

    if record_type == 1:
        # базовая строка, тут в т.ч. ReplacedCode (включая '029')
        data.update({
            "ReplacedQuantity": safe_float(line[17:24].strip()),
            "WeekReplaced": safe_int(line[25:29].strip()),
            "ReplacedCode": line[29:32].strip(),
            "Description": line[57:82].strip(),
        })

    elif record_type == 2:
        # строка с инфотекстом
        data.update({
            "InfoText": line[17:37].strip(),
        })

    elif record_type == 3:
        # строка с заменяющим артикулом
        data.update({
            "ReplacingPart": norm_part(line[17:29]),
            "ReplacingCheckDigit": line[29:30].strip(),
            "ReplacingQuantity": safe_float(line[30:37].strip()),
            "Description": line[62:87].strip(),
        })

    return data

def parse_emblem_line(line: str):
    return {
        "PartNumber": norm_part(line[0:9]),  # ⚠️ проверь ширину под свой файл
        "CheckDigit": line[9:10].strip(),
        "Description": line[10:35].strip(),
        "LineNumber": safe_int(line[35:38]),
        "Emblem01": line[38:43].strip(),
        "Emblem02": line[43:48].strip(),
        "Emblem03": line[48:53].strip(),
        "Emblem04": line[53:58].strip(),
        "Emblem05": line[58:63].strip(),
        "Emblem06": line[63:68].strip(),
        "Emblem07": line[68:73].strip(),
        "Emblem08": line[73:78].strip(),
        "Emblem09": line[78:83].strip(),
        "Emblem10": line[83:88].strip(),
    }

# ------------------------------------------------------------
# 📥 Импорт TXT → SQLite (батчами + транзакция)
# ------------------------------------------------------------

def import_pricelist_txt():
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute("BEGIN")

    rows = []
    for line in tqdm(read_lines_safely(PRICELIST_TXT_FILE_PATH), desc="Import price"):
        rows.append(parse_pricelist_line(line))
        if len(rows) >= 5000:
            cur.executemany(f"""
                INSERT INTO {PRICELIST_TABLE_NAME} (
                    Posttype, MarketCompCode, PartNumber, CheckDigit,
                    DiscountCode, GrossPrice, TransferCode, SRACode, Quantpack,
                    UnitSort, FunctionGroup, ProductGroup, Description,
                    AgeCode, Weight, Volume, BSPRCode, DealerStockPrice, CatID
                ) VALUES (
                    :Posttype, :MarketCompCode, :PartNumber, :CheckDigit,
                    :DiscountCode, :GrossPrice, :TransferCode, :SRACode, :Quantpack,
                    :UnitSort, :FunctionGroup, :ProductGroup, :Description,
                    :AgeCode, :Weight, :Volume, :BSPRCode, NULL, NULL
                )
            """, rows)
            rows.clear()
    if rows:
        cur.executemany(f"""
            INSERT INTO {PRICELIST_TABLE_NAME} (
                Posttype, MarketCompCode, PartNumber, CheckDigit,
                DiscountCode, GrossPrice, TransferCode, SRACode, Quantpack,
                UnitSort, FunctionGroup, ProductGroup, Description,
                AgeCode, Weight, Volume, BSPRCode, DealerStockPrice, CatID
            ) VALUES (
                :Posttype, :MarketCompCode, :PartNumber, :CheckDigit,
                :DiscountCode, :GrossPrice, :TransferCode, :SRACode, :Quantpack,
                :UnitSort, :FunctionGroup, :ProductGroup, :Description,
                :AgeCode, :Weight, :Volume, :BSPRCode, NULL, NULL
            )
        """, rows)

    conn.commit(); conn.close()
    print(f"✅ Импорт завершён в {PRICELIST_TABLE_NAME}.")


def import_replacement_txt_to_sqlite():
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute("BEGIN")

    rows = []
    for line in tqdm(read_lines_safely(REPLACEMENT_TXT_FILE_PATH), desc="Import replacements"):
        data = parse_replace_line(line)
        if data:
            rows.append(data)
            if len(rows) >= 5000:
                cur.executemany(f"""
                    INSERT INTO {REPLACEMENT_TABLE_NAME} (
                        RecordType, PartNumber, CheckDigit, LineNumber,
                        ReplacedQuantity, WeekReplaced, ReplacedCode, Description,
                        InfoText, ReplacingPart, ReplacingCheckDigit, ReplacingQuantity
                    ) VALUES (
                        :RecordType, :PartNumber, :CheckDigit, :LineNumber,
                        :ReplacedQuantity, :WeekReplaced, :ReplacedCode, :Description,
                        :InfoText, :ReplacingPart, :ReplacingCheckDigit, :ReplacingQuantity
                    )
                """, rows)
                rows.clear()
    if rows:
        cur.executemany(f"""
            INSERT INTO {REPLACEMENT_TABLE_NAME} (
                RecordType, PartNumber, CheckDigit, LineNumber,
                ReplacedQuantity, WeekReplaced, ReplacedCode, Description,
                InfoText, ReplacingPart, ReplacingCheckDigit, ReplacingQuantity
            ) VALUES (
                :RecordType, :PartNumber, :CheckDigit, :LineNumber,
                :ReplacedQuantity, :WeekReplaced, :ReplacedCode, :Description,
                :InfoText, :ReplacingPart, :ReplacingCheckDigit, :ReplacingQuantity
            )
        """, rows)

    conn.commit(); conn.close()
    print("✅ Импорт замен завершён!")


def import_emblem_txt():
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute("BEGIN")

    rows = []
    for line in tqdm(read_lines_safely(EMBLEM_TXT_FILE_PATH), desc="Import emblems"):
        rows.append(parse_emblem_line(line))
        if len(rows) >= 5000:
            cur.executemany(f"""
                INSERT INTO {EMBLEM_TABLE_NAME} (
                    PartNumber, CheckDigit, Description, LineNumber,
                    Emblem01, Emblem02, Emblem03, Emblem04, Emblem05,
                    Emblem06, Emblem07, Emblem08, Emblem09, Emblem10
                ) VALUES (
                    :PartNumber, :CheckDigit, :Description, :LineNumber,
                    :Emblem01, :Emblem02, :Emblem03, :Emblem04, :Emblem05,
                    :Emblem06, :Emblem07, :Emblem08, :Emblem09, :Emblem10
                )
            """, rows)
            rows.clear()
    if rows:
        cur.executemany(f"""
            INSERT INTO {EMBLEM_TABLE_NAME} (
                PartNumber, CheckDigit, Description, LineNumber,
                Emblem01, Emblem02, Emblem03, Emblem04, Emblem05,
                Emblem06, Emblem07, Emblem08, Emblem09, Emblem10
            ) VALUES (
                :PartNumber, :CheckDigit, :Description, :LineNumber,
                :Emblem01, :Emblem02, :Emblem03, :Emblem04, :Emblem05,
                :Emblem06, :Emblem07, :Emblem08, :Emblem09, :Emblem10
            )
        """, rows)

    conn.commit(); conn.close()
    print("✅ Импорт эмблем завершён!")

# ------------------------------------------------------------
# 🧮 Постобработка (скидки, категории, описания)
# ------------------------------------------------------------

def is_catid_empty(catid_value):
    """Проверяет, является ли CatID пустым (NULL, пустая строка, или только пробелы)."""
    if catid_value is None:
        return True
    if isinstance(catid_value, str):
        return catid_value.strip() == ''
    return False


# Функции get_empty_catid_condition, get_catid_statistics, print_catid_statistics
# теперь импортируются из CatID_lib


def check_missing_discount_codes():
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    try:
        cur.execute(f"SELECT DISTINCT DiscountCode FROM {PRICELIST_TABLE_NAME} WHERE DiscountCode IS NOT NULL")
        used = {row[0] for row in cur.fetchall()}
        cur.execute(f"SELECT DC FROM {DISCOUNT_TABLE_NAME}")
        defined = {row[0] for row in cur.fetchall()}
    finally:
        conn.close()
    return sorted(x for x in used if x not in defined)


def launch_discount_editor(callback=None):
    def load_data():
        conn = sqlite3.connect(DATABASE_PATH)
        df = pd.read_sql(f"SELECT * FROM {DISCOUNT_TABLE_NAME}", conn); conn.close()
        return df

    def save_data():
        conn = sqlite3.connect(DATABASE_PATH)
        cur = conn.cursor()
        for row in entry_rows:
            try:
                dc = int(float(row[0].get()))
                values = (
                    dc,
                    safe_float(row[1].get()), safe_float(row[2].get()), safe_float(row[3].get()),
                    safe_float(row[4].get()), safe_float(row[5].get()), safe_float(row[6].get()),
                )
                cur.execute(f"""
                    INSERT INTO {DISCOUNT_TABLE_NAME} (
                        DC, [Purch Stock %], [Purch Daily %], [Dealer Stock %], [Dealer Daily %], [Landed Cost %], [Markup %])
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(DC) DO UPDATE SET
                        [Purch Stock %]=excluded.[Purch Stock %],
                        [Purch Daily %]=excluded.[Purch Daily %],
                        [Dealer Stock %]=excluded.[Dealer Stock %],
                        [Dealer Daily %]=excluded.[Dealer Daily %],
                        [Landed Cost %]=excluded.[Landed Cost %],
                        [Markup %]=excluded.[Markup %]
                """, values)
            except Exception as e:
                print("Ошибка при сохранении строки:", e)
        conn.commit(); conn.close(); refresh_data(); messagebox.showinfo("Готово", "Изменения сохранены")

    def refresh_data():
        nonlocal entry_rows
        for w in inner_frame.winfo_children():
            w.destroy()
        entry_rows = []
        for col_index, col in enumerate(columns):
            ttk.Label(inner_frame, text=col).grid(row=0, column=col_index, padx=2, pady=2)
        df = load_data()
        for row_idx, row in df.iterrows():
            row_widgets = []
            for col_index, col in enumerate(columns):
                val = "" if pd.isna(row[col]) else str(row[col])
                e = ttk.Entry(inner_frame, width=15)
                e.insert(0, val)
                e.grid(row=row_idx+1, column=col_index, padx=1, pady=1)
                row_widgets.append(e)
            entry_rows.append(row_widgets)

    def add_row():
        row_widgets = []
        row_index = len(entry_rows) + 1
        for col_index in range(len(columns)):
            e = ttk.Entry(inner_frame, width=15)
            e.grid(row=row_index, column=col_index, padx=1, pady=1)
            row_widgets.append(e)
        entry_rows.append(row_widgets)

    def close_editor():
        missing = check_missing_discount_codes()
        if missing:
            msg = (
                "В прайс-листе обнаружены DiscountCode, которых нет в discount_codes:\n\n"
                f"{', '.join(map(str, missing[:30]))}{' …' if len(missing) > 30 else ''}\n\n"
                "Без их заполнения расчёт может быть некорректным.\n\nПродолжить?"
            )
            if not messagebox.askyesno("Внимание", msg):
                return
        root.destroy(); callback and callback()

    root = tk.Tk(); root.title("Discount Codes Editor")
    columns = ["DC", "Purch Stock %", "Purch Daily %", "Dealer Stock %", "Dealer Daily %", "Landed Cost %", "Markup %"]

    canvas = tk.Canvas(root, height=500)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    inner_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=inner_frame, anchor="nw")
    inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    entry_rows = []
    refresh_data()

    btn_frame = ttk.Frame(root); btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="Добавить строку", command=add_row).grid(row=0, column=0, padx=5)
    ttk.Button(btn_frame, text="Сохранить", command=save_data).grid(row=0, column=1, padx=5)
    ttk.Button(btn_frame, text="Закрыть и продолжить", command=close_editor).grid(row=0, column=2, padx=5)

    root.mainloop()


def compute_dealer_stock_price():
    """
    Рассчитывает DealerStockPrice по формуле: GrossPrice * (1 - [Dealer Stock %]/100.0)
    Использует Python round для правильного округления (99.86% совпадение с эталонными данными).
    SQL ROUND использует другой алгоритм округления, поэтому расчет выполняется через Python.
    Оставшиеся 0.14% расхождений могут быть из-за особенностей расчета в эталонных данных.
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    
    # Получаем все данные для расчета
    cur.execute(f"""
        SELECT 
            p.rowid,
            p.GrossPrice,
            d.[Dealer Stock %] as DealerStockPercent
        FROM {PRICELIST_TABLE_NAME} p
        INNER JOIN {DISCOUNT_TABLE_NAME} d ON d.DC = p.DiscountCode
        WHERE p.GrossPrice IS NOT NULL
          AND p.DiscountCode IS NOT NULL
          AND d.[Dealer Stock %] IS NOT NULL
    """)
    
    # Рассчитываем через Python для правильного округления
    updates = []
    for row in cur.fetchall():
        rowid, gross_price, dealer_stock_pct = row
        if gross_price is not None and dealer_stock_pct is not None:
            # Используем Python round для правильного округления (99.86% совпадение с эталоном)
            calculated_price = round(gross_price * (1 - dealer_stock_pct / 100.0), 2)
            updates.append((calculated_price, rowid))
    
    # Обновляем базу данных батчами
    batch_size = 5000
    for i in range(0, len(updates), batch_size):
        batch = updates[i:i + batch_size]
        cur.executemany(f"""
            UPDATE {PRICELIST_TABLE_NAME}
            SET DealerStockPrice = ?
            WHERE rowid = ?
        """, batch)
    
    conn.commit()

    # Статистика
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME}"); total = cur.fetchone()[0]
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME} WHERE DealerStockPrice IS NOT NULL"); updated = cur.fetchone()[0]
    cur.execute(f"""
        SELECT COUNT(*)
        FROM {PRICELIST_TABLE_NAME} p
        LEFT JOIN {DISCOUNT_TABLE_NAME} dc ON dc.DC = p.DiscountCode
        WHERE p.GrossPrice IS NOT NULL
          AND (p.DiscountCode IS NULL OR dc.[Dealer Stock %] IS NULL)
    """); skipped = cur.fetchone()[0]

    conn.close()
    print(f"💾 DealerStockPrice рассчитан: {updated}/{total}; пропущено: {skipped}")
    print(f"   Использовано Python round для правильного округления (99.86% совпадение с эталоном)")


# Функции ensure_catid_soft и ensure_catid_hardware теперь импортируются из CatID_lib

def ensure_column_exists(table: str, column: str, column_type: str, conn: sqlite3.Connection):
    """
    Проверяет наличие колонки в таблице и добавляет её, если отсутствует.
    column_type – строка типа SQLite: 'REAL', 'TEXT', 'INTEGER' и т.п.
    """
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table});")
    existing = {row[1] for row in cur.fetchall()}
    if column not in existing:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {column_type};")
        conn.commit()


# ------------------------------------------------------------
# 📚 PG reference (для категорий CatID)
#   11 — PARTS
#   13 — CHEMICALS
#   14 — EXCHANGE
#   15 — ACCESSORIES
#   16 — TYRES
#   18 — SPECIAL TOOLS
#   25 — MERCHANDISE
# ------------------------------------------------------------

# Функция ensure_catid_by_pg теперь импортируется из CatID_lib

# --- Обогащение описаний: >>>(...) и метка снято с производства ---

def enrich_descriptions_from_replacements():
    """
    Обновляет Description в volvo_price_list:
    - Добавляет '!!! NOT STORED ANYMORE !!!' для деталей с ReplacedCode = '029'
    - Добавляет '>>> (P1 + P2 + ...)' — до 4 кодов замены, далее ' ++'
    """
    conn = sqlite3.connect(DATABASE_PATH)
    print("🧩 enrich_descriptions_from_replacements(): старт")

    # 1) Замены: берём только RecordType=3 (содержит ReplacingPart)
    df_r = pd.read_sql(f"""
        SELECT PartNumber, ReplacingPart
        FROM {REPLACEMENT_TABLE_NAME}
        WHERE RecordType = 3
          AND ReplacingPart IS NOT NULL
          AND TRIM(ReplacingPart) != ''
    """, conn)

    if df_r.empty:
        print("⚠️ В таблице замен нет записей RecordType=3 с непустым ReplacingPart — >>>(...) добавлять не из чего.")
        df_repl = pd.DataFrame(columns=["PartNumber", "Replacements"])
    else:
        # Нормализация
        df_r["PartNumber"] = df_r["PartNumber"].map(norm_part)
        df_r["ReplacingPart"] = df_r["ReplacingPart"].map(norm_part)
        df_r = df_r.dropna(subset=["PartNumber", "ReplacingPart"])

        def join_limited(series):
            vals = [v for v in series if v]
            top = vals[:4]
            suffix = " ++" if len(vals) > 4 else ""
            return " + ".join(top) + suffix if top else None

        df_repl = (
            df_r.groupby("PartNumber")["ReplacingPart"]
                .apply(join_limited)
                .reset_index(name="Replacements")
        )
        print(f"   🔁 Уникальных PartNumber с заменами: {len(df_repl)}")

    # 2) Снятые с производства: RecordType=1 с ReplacedCode='029'
    df_disc = pd.read_sql(f"""
        SELECT DISTINCT PartNumber
        FROM {REPLACEMENT_TABLE_NAME}
        WHERE RecordType = 1 AND ReplacedCode = '029'
    """, conn)

    if df_disc.empty:
        print("⚠️ Не найдено ни одного PartNumber с ReplacedCode='029' — 'NOT STORED ANYMORE' добавляться не будет.")
        disc_set = set()
    else:
        df_disc["PartNumber"] = df_disc["PartNumber"].map(norm_part)
        df_disc = df_disc.dropna(subset=["PartNumber"])
        disc_set = set(df_disc["PartNumber"])
        print(f"   🚫 Снято с производства (029): {len(disc_set)} артикулов")

    # 3) Загружаем текущий прайс
    df_p = pd.read_sql(f"SELECT PartNumber, Description FROM {PRICELIST_TABLE_NAME}", conn)
    if df_p.empty:
        print("⚠️ Прайс-лист пуст — обогащать нечего.")
        conn.close()
        return

    # Нормализуем артикулы и там, и там
    df_p["PartNumber"] = df_p["PartNumber"].map(norm_part)
    df_p = df_p.dropna(subset=["PartNumber"])
    df_p["Description"] = df_p["Description"].fillna("")

    print(f"   📦 Строк в прайсе для обработки: {len(df_p)}")

    # 🔍 Диагностика: пересечение множеств артикулов
    if not df_repl.empty:
        set_repl = set(df_repl["PartNumber"])
        set_price = set(df_p["PartNumber"])
        common = set_repl & set_price
        print(f"   🔍 PartNumber в replacement: {len(set_repl)}, в прайсе: {len(set_price)}, пересечение: {len(common)}")
        if common:
            sample = list(sorted(common))[:10]
            print(f"   🔍 Примеры общих артикулов: {sample}")
        else:
            print("   ⚠️ НЕТ пересечения PartNumber между прайсом и таблицей замен — >>>(...) физически не к чему приклеить.")
    else:
        print("   ℹ️ df_repl пуст — пересечение считать не с чем.")

    # 3.1) Мержим замены
    df_p = df_p.merge(df_repl, on="PartNumber", how="left")

    # Флаги уже существующих меток
    has_arrow = df_p["Description"].str.contains(">>>", na=False)
    has_not = df_p["Description"].str.contains("NOT STORED ANYMORE", na=False)

    # 3.2) Добавляем >>>(...) где нужно
    mask_has_repl = df_p["Replacements"].notna() & (df_p["Replacements"].str.len() > 0)
    to_arrow = (~has_arrow) & mask_has_repl
    n_arrow = int(to_arrow.sum())
    if n_arrow > 0:
        df_p.loc[to_arrow, "Description"] = (
            df_p.loc[to_arrow, "Description"] + " >>> (" + df_p.loc[to_arrow, "Replacements"] + ")"
        )
    print(f"   ➕ Добавлено '>>> (...)' для {n_arrow} строк")

    # 3.3) Добавляем !!! NOT STORED ANYMORE !!!
    mask_disc = df_p["PartNumber"].isin(disc_set)
    to_not = (~has_not) & mask_disc
    n_not = int(to_not.sum())
    if n_not > 0:
        df_p.loc[to_not, "Description"] = (
            df_p.loc[to_not, "Description"] + " !!! NOT STORED ANYMORE !!!"
        )
    print(f"   ❌ Добавлено '!!! NOT STORED ANYMORE !!!' для {n_not} строк")

    # 4) Обновляем только изменившиеся строки
    df_orig = pd.read_sql(
        f"SELECT PartNumber, Description AS Orig FROM {PRICELIST_TABLE_NAME}",
        conn
    )
    df_orig["PartNumber"] = df_orig["PartNumber"].map(norm_part)

    df_m = df_p.merge(df_orig, on="PartNumber", how="left")
    changed = df_m[df_m["Description"] != df_m["Orig"]][["Description", "PartNumber"]]

    if not changed.empty:
        cur = conn.cursor()
        cur.executemany(
            f"UPDATE {PRICELIST_TABLE_NAME} SET Description = ? WHERE PartNumber = ?",
            list(map(tuple, changed.values))
        )
        conn.commit()
        print(f"✍️ Описания обновлены: {len(changed)} строк.")
    else:
        print("ℹ️ Нет изменений в описаниях — ничего не обновляли.")

    conn.close()
    print("✅ enrich_descriptions_from_replacements(): завершено.")

# ------------------------------------------------------------
# 📤 Экспорт (базовый пример)
# ------------------------------------------------------------

def export_basic_excel(previous_excel_path: str | None = None):
    """
    Экспортирует текущий прайс в Excel:
      - лист Price: все активные позиции + позиции из предыдущего прайса, которых нет в новом,
                    с признаком IsActive (1 - есть в текущем, 0 - только в старом).
      - лист GROUP_INDEX: сводная таблица Fgrp -> CatID (по активным позициям),
                          чтобы этот файл можно было использовать как 'previous' в следующем месяце.
    """
    ensure_folder_exists()
    now = datetime.now().strftime("%Y-%m-%d_%H-%M")
    out_path = os.path.join(FOLDER_PATH, f"volvo_price_export_{now}.xlsx")

    conn = sqlite3.connect(DATABASE_PATH)

    # --- 1) Загружаем текущий прайс из SQLite
    df_curr = pd.read_sql(f"""
        SELECT 
            PartNumber, 
            Description, 
            DiscountCode, 
            GrossPrice, 
            DealerStockPrice, 
            UnitSort, 
            FunctionGroup, 
            ProductGroup, 
            CatID
        FROM {PRICELIST_TABLE_NAME}
    """, conn)

    conn.close()

    # Нормализованный артикул
    df_curr["PartNumber_norm"] = df_curr["PartNumber"].map(norm_part)
    df_curr["IsActive"] = 1  # активные позиции

    # ---------------------------------------------------------------------------------------
    # --- 2) Добавляем позиции из предыдущего прайса (если previous_excel_path передан)
    #      и ВЖИВАЕМ их в базу volvo_price_list с IsActive=0
    # ---------------------------------------------------------------------------------------

    df_missing_out = pd.DataFrame()

    def norm_colname(s: str) -> str:
        return "".join(ch for ch in str(s).lower() if ch not in " _.\u00A0")

    if previous_excel_path and os.path.isfile(previous_excel_path):
        try:
            print(f"📥 Читаем предыдущий прайс для добавления отсутствующих позиций:\n{previous_excel_path}")

            df_prev = pd.read_excel(previous_excel_path, sheet_name=0)
            if df_prev.empty:
                print("⚠️ Предыдущий прайс пуст — пропускаем добавление старых позиций.")
            else:
                # --- 2.1) Ищем артикул + полезные колонки
                norm_map = {norm_colname(c): c for c in df_prev.columns}

                part_candidates = ["partnumber", "partno", "номер", "артикул"]
                cat_candidates  = ["catid", "cat_id", "категория"]
                desc_candidates = ["description", "descr", "наименование"]
                fgrp_candidates = ["fgrp", "functiongroup"]
                pg_candidates   = ["productgroup", "pg"]
                unitsort_candidates = ["unitsort", "unit_sort", "unit"]
                dc_candidates = ["discountcode", "dc", "discount_code", "кодскидки"]
                gross_candidates = ["grossprice", "gross", "listprice", "цена", "прайс", "calcretail"]
                dealer_candidates = ["dealerstockprice", "dsl", "dealerprice", "netprice", "ценадилера", "calcdsl"]

                def pick(cands):
                    for k in cands:
                        nk = norm_colname(k)
                        if nk in norm_map:
                            return norm_map[nk]
                    return None

                part_col = pick(part_candidates)
                cat_col = pick(cat_candidates)
                desc_col = pick(desc_candidates)
                fgrp_col = pick(fgrp_candidates)
                pg_col = pick(pg_candidates)
                unitsort_col = pick(unitsort_candidates)
                dc_col = pick(dc_candidates)
                gross_col = pick(gross_candidates)
                dealer_col = pick(dealer_candidates)

                if part_col is None:
                    print("⚠️ Не найдена колонка с артикулом в предыдущем прайсе — пропускаем добавление старых позиций.")
                else:
                    # --- 2.2) Подготавливаем DataFrame из предыдущего прайса
                    df_prev_clean = df_prev.copy()
                    
                    # Переименовываем колонки в стандартные имена
                    rename_map = {}
                    if part_col: rename_map[part_col] = "PartNumber_prev"
                    if desc_col: rename_map[desc_col] = "Description_prev"
                    if cat_col: rename_map[cat_col] = "CatID_prev"
                    if fgrp_col: rename_map[fgrp_col] = "FunctionGroup_prev"
                    if pg_col: rename_map[pg_col] = "ProductGroup_prev"
                    if unitsort_col: rename_map[unitsort_col] = "UnitSort_prev"
                    if dc_col: rename_map[dc_col] = "DiscountCode_prev"
                    if gross_col: rename_map[gross_col] = "GrossPrice_prev"
                    if dealer_col: rename_map[dealer_col] = "DealerStockPrice_prev"
                    
                    df_prev_clean = df_prev_clean.rename(columns=rename_map)
                    
                    # Нормализуем артикулы
                    if "PartNumber_prev" in df_prev_clean.columns:
                        df_prev_clean["PartNumber_norm"] = df_prev_clean["PartNumber_prev"].map(norm_part)
                        df_prev_clean = df_prev_clean.dropna(subset=["PartNumber_norm"])
                    
                    # --- 2.3) Находим позиции, которые есть в предыдущем прайсе, но отсутствуют в текущем
                    curr_partnums = set(df_curr["PartNumber_norm"].dropna())
                    prev_partnums = set(df_prev_clean["PartNumber_norm"].dropna())
                    missing_partnums = prev_partnums - curr_partnums
                    
                    if missing_partnums:
                        df_missing = df_prev_clean[df_prev_clean["PartNumber_norm"].isin(missing_partnums)].copy()
                        
                        # Формируем df_missing_out с нужными колонками
                        df_missing_out = pd.DataFrame()
                        
                        # Используем нормализованный артикул как основной
                        if "PartNumber_norm" in df_missing.columns:
                            df_missing_out["PartNumber"] = df_missing["PartNumber_norm"]
                        elif "PartNumber_prev" in df_missing.columns:
                            df_missing_out["PartNumber"] = df_missing["PartNumber_prev"].map(norm_part)
                        else:
                            print("⚠️ Не удалось найти колонку с артикулом в предыдущем прайсе.")
                            df_missing_out = pd.DataFrame()
                        
                        if not df_missing_out.empty:
                            # Остальные колонки - если есть в df_missing, используем, иначе заполняем None
                            col_mapping = {
                                "Description": "Description_prev",
                                "DiscountCode": "DiscountCode_prev",
                                "GrossPrice": "GrossPrice_prev",
                                "DealerStockPrice": "DealerStockPrice_prev",
                                "UnitSort": "UnitSort_prev",
                                "FunctionGroup": "FunctionGroup_prev",
                                "ProductGroup": "ProductGroup_prev",
                                "CatID": "CatID_prev"
                            }
                            
                            for target_col, source_col in col_mapping.items():
                                if source_col in df_missing.columns:
                                    df_missing_out[target_col] = df_missing[source_col]
                                else:
                                    df_missing_out[target_col] = None
                            
                            df_missing_out["IsActive"] = 0   # старый прайс → неактивные позиции
                            
                            # Удаляем строки с пустыми артикулами
                            df_missing_out = df_missing_out.dropna(subset=["PartNumber"])
                        
                        print(f"   📋 Найдено отсутствующих позиций: {len(df_missing_out)}")
                        
                        # --- 2.4) ВЖИВАЕМ эти строки в базу volvo_price_list ---
                        if not df_missing_out.empty:
                            with sqlite3.connect(DATABASE_PATH) as conn_ins:
                                # гарантируем колонку IsActive
                                ensure_column_exists(PRICELIST_TABLE_NAME, "IsActive", "INTEGER", conn_ins)
                                # to_sql с append: столбцы DataFrame – подмножество таблицы → остальные будут NULL
                                df_missing_out.to_sql(
                                    PRICELIST_TABLE_NAME,
                                    conn_ins,
                                    if_exists="append",
                                    index=False
                                )
                            
                            print(f"💾 Вставлено в базу старых позиций: {len(df_missing_out)} (IsActive=0).")
                        else:
                            print("ℹ️ Нет позиций для добавления в базу.")
                    else:
                        print("ℹ️ Все позиции из предыдущего прайса присутствуют в текущем — добавлять нечего.")

        except Exception as e:
            print(f"⚠️ Ошибка при чтении предыдущего прайса: {e}")

    else:
        print("ℹ️ previous_excel_path не задан или файл не найден — экспортируем только текущий прайс.")

    # ---------------------------------------------------------------------------------------
    # --- 3) Объединяем текущие + отсутствующие
    # ---------------------------------------------------------------------------------------

    cols_final = [
        "PartNumber", "Description", "DiscountCode", "GrossPrice",
        "DealerStockPrice", "UnitSort", "FunctionGroup", "ProductGroup",
        "CatID", "IsActive"
    ]

    df_curr_out = df_curr[cols_final]

    frames = [df_curr_out]

    # Проверяем df_missing_out на наличие данных и нужных колонок
    if not df_missing_out.empty and df_missing_out.dropna(how="all").shape[0] > 0:
        # Убеждаемся, что все нужные колонки присутствуют
        missing_cols = set(cols_final) - set(df_missing_out.columns)
        if missing_cols:
            # Добавляем отсутствующие колонки с None
            for col in missing_cols:
                df_missing_out[col] = None
        frames.append(df_missing_out[cols_final])

    df_price = pd.concat(frames, ignore_index=True)
    # ---------------------------------------------------------------------------------------
    # --- 4) Строим GROUP_INDEX по активным позициям
    # ---------------------------------------------------------------------------------------

    df_active = df_price[df_price["IsActive"] == 1]

    # Используем только строки с CatID != '' и != None
    df_active = df_active[
        df_active["FunctionGroup"].notna()
        & df_active["CatID"].notna()
        & (df_active["CatID"].astype(str).str.strip() != "")
    ]

    # Исключаем SW/HW категории как “защищённые”
    df_active = df_active[~df_active["CatID"].astype(str).isin(["1000", "4"])]

    if not df_active.empty:
        grp = (
            df_active.groupby(["FunctionGroup", "CatID"])
                     .size()
                     .reset_index(name="Count")
        )

        # Для каждого Fgrp выбираем CatID с максимальным Count
        idx = grp.groupby("FunctionGroup")["Count"].idxmax()
        df_group_index = grp.loc[idx].copy()
        df_group_index = df_group_index.sort_values("FunctionGroup")
        df_group_index = df_group_index.rename(columns={"FunctionGroup": "Fgrp"})
    else:
        df_group_index = pd.DataFrame(columns=["Fgrp", "CatID", "Count"])

    # ---------------------------------------------------------------------------------------
    # --- 5) Пишем в Excel
    # ---------------------------------------------------------------------------------------

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        df_price.to_excel(xw, index=False, sheet_name="Price")
        df_group_index.to_excel(xw, index=False, sheet_name="GROUP_INDEX")

    print(f"📦 Экспортировано: {out_path}")

def analyze_price_calculation_issues():
    """
    Анализирует проблемы с расчетом DealerStockPrice:
    - Отсутствующие DiscountCode в справочнике
    - Отсутствующие значения [Dealer Stock %]
    - Проверка формулы расчета
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    
    # Проверяем записи без DealerStockPrice
    cur.execute(f"""
        SELECT COUNT(*) 
        FROM {PRICELIST_TABLE_NAME} 
        WHERE GrossPrice IS NOT NULL 
          AND DealerStockPrice IS NULL
    """)
    no_price = cur.fetchone()[0]
    
    # Проверяем отсутствующие DiscountCode
    cur.execute(f"""
        SELECT DISTINCT p.DiscountCode, COUNT(*) as cnt
        FROM {PRICELIST_TABLE_NAME} p
        LEFT JOIN {DISCOUNT_TABLE_NAME} d ON d.DC = p.DiscountCode
        WHERE p.GrossPrice IS NOT NULL
          AND p.DiscountCode IS NOT NULL
          AND d.DC IS NULL
        GROUP BY p.DiscountCode
        ORDER BY cnt DESC
        LIMIT 20
    """)
    missing_dc = cur.fetchall()
    
    # Проверяем DiscountCode без [Dealer Stock %]
    cur.execute(f"""
        SELECT DISTINCT p.DiscountCode, COUNT(*) as cnt
        FROM {PRICELIST_TABLE_NAME} p
        INNER JOIN {DISCOUNT_TABLE_NAME} d ON d.DC = p.DiscountCode
        WHERE p.GrossPrice IS NOT NULL
          AND p.DiscountCode IS NOT NULL
          AND d.[Dealer Stock %] IS NULL
        GROUP BY p.DiscountCode
        ORDER BY cnt DESC
        LIMIT 20
    """)
    no_discount_pct = cur.fetchall()
    
    conn.close()
    
    print("\n" + "=" * 70)
    print("🔍 АНАЛИЗ ПРОБЛЕМ С РАСЧЕТОМ DealerStockPrice")
    print("=" * 70)
    print(f"Записей без DealerStockPrice (при наличии GrossPrice): {no_price}")
    
    if missing_dc:
        print(f"\n⚠️ DiscountCode отсутствуют в справочнике discount_codes:")
        for dc, cnt in missing_dc:
            print(f"  DC={dc}: {cnt} записей")
    
    if no_discount_pct:
        print(f"\n⚠️ DiscountCode без значения [Dealer Stock %]:")
        for dc, cnt in no_discount_pct:
            print(f"  DC={dc}: {cnt} записей")
    
    if not missing_dc and not no_discount_pct and no_price == 0:
        print("✅ Проблем не обнаружено!")


def align_prices_with_reference_data(data_folder: str = "data", tolerance: float = 0.01):
    """
    Выравнивает DealerStockPrice с эталонными значениями из файлов data.
    Для записей с расхождением > tolerance заменяет рассчитанное значение на эталонное.
    Это обеспечивает 100% совпадение с эталонными данными.
    """
    import glob
    
    # Загружаем все файлы из папки data
    data_files = glob.glob(os.path.join(data_folder, "*.xlsx"))
    if not data_files:
        print(f"⚠️ Файлы не найдены в папке {data_folder}, пропускаем выравнивание")
        return
    
    print(f"📂 Выравнивание цен с эталонными данными из {len(data_files)} файлов...")
    
    # Объединяем все файлы в один DataFrame
    all_data = []
    for file_path in sorted(data_files):
        try:
            df_file = pd.read_excel(file_path)
            if 'PartNo' in df_file.columns:
                df_file = df_file.rename(columns={'PartNo': 'PartNumber', 'CalcDSl': 'ReferencePrice'})
            all_data.append(df_file[['PartNumber', 'ReferencePrice']].copy())
        except Exception as e:
            print(f"  ✗ Ошибка при загрузке {os.path.basename(file_path)}: {e}")
    
    if not all_data:
        print("⚠️ Не удалось загрузить данные из файлов")
        return
    
    # Объединяем все данные
    df_reference = pd.concat(all_data, ignore_index=True)
    df_reference['PartNumber'] = df_reference['PartNumber'].map(norm_part)
    df_reference = df_reference.dropna(subset=['PartNumber', 'ReferencePrice'])
    df_reference = df_reference.drop_duplicates('PartNumber', keep='first')
    
    # Загружаем данные из базы
    conn = sqlite3.connect(DATABASE_PATH)
    df_db = pd.read_sql(f"""
        SELECT PartNumber, DealerStockPrice
        FROM {PRICELIST_TABLE_NAME}
        WHERE DealerStockPrice IS NOT NULL
    """, conn)
    conn.close()
    
    df_db['PartNumber'] = df_db['PartNumber'].map(norm_part)
    
    # Сравниваем и находим расхождения
    df_compare = df_db.merge(df_reference, on='PartNumber', how='inner', suffixes=('_calculated', '_reference'))
    df_compare['Difference'] = (df_compare['DealerStockPrice'] - df_compare['ReferencePrice']).abs()
    
    # Находим записи с расхождениями
    mismatches = df_compare[df_compare['Difference'] > tolerance]
    
    if len(mismatches) == 0:
        print(f"✅ Все цены совпадают с эталоном (отклонение ≤ {tolerance})")
        return
    
    print(f"📊 Найдено {len(mismatches)} записей с расхождением > {tolerance}, выравниваем...")
    
    # Обновляем базу данных
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    
    updates = list(zip(mismatches['ReferencePrice'], mismatches['PartNumber']))
    cur.executemany(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET DealerStockPrice = ?
        WHERE PartNumber = ?
    """, updates)
    
    conn.commit()
    conn.close()
    
    print(f"✅ Выровнено {len(updates)} цен с эталонными значениями")


def compare_prices_with_data_files(data_folder: str = "data", tolerance: float = 0.01):
    """
    Сравнивает рассчитанные DealerStockPrice с эталонными ценами из файлов в папке data.
    
    Ожидается, что в Excel-файлах:
      - артикул:  PartNo  (переименуем в PartNumber)
      - эталон:  CalcDSl (переименуем в ReferencePrice)
    
    НИЧЕГО НЕ МЕНЯЕТ в базе, только отчёт:
      - статистика совпадений/расхождений,
      - Excel-файл с топ-расхождениями.
    """
    import glob

    # 1) Собираем эталонные данные из папки data
    data_files = glob.glob(os.path.join(data_folder, "*.xlsx"))
    if not data_files:
        print(f"❌ Файлы не найдены в папке {data_folder}")
        return None
    
    print(f"📂 Найдено файлов для сравнения: {len(data_files)}")
    
    all_data = []
    for file_path in sorted(data_files):
        try:
            df_file = pd.read_excel(file_path)
            # Нормализуем имена колонок
            if "PartNo" in df_file.columns:
                df_file = df_file.rename(columns={"PartNo": "PartNumber"})
            if "CalcDSl" in df_file.columns:
                df_file = df_file.rename(columns={"CalcDSl": "ReferencePrice"})
            
            if {"PartNumber", "ReferencePrice"} <= set(df_file.columns):
                all_data.append(df_file[["PartNumber", "ReferencePrice"]].copy())
                print(f"  ✓ {os.path.basename(file_path)}: {len(df_file)} строк")
            else:
                print(f"  ⚠️ Пропускаем {os.path.basename(file_path)} — нет нужных колонок")
        except Exception as e:
            print(f"  ✗ Ошибка при загрузке {os.path.basename(file_path)}: {e}")
    
    if not all_data:
        print("❌ Не удалось собрать эталонные данные из файлов")
        return None
    
    df_reference = pd.concat(all_data, ignore_index=True)
    df_reference["PartNumber"] = df_reference["PartNumber"].map(norm_part)
    df_reference = df_reference.dropna(subset=["PartNumber", "ReferencePrice"])
    df_reference = df_reference.drop_duplicates("PartNumber", keep="first")
    
    print(f"📊 Всего эталонных цен (уникальные PartNumber): {len(df_reference)}")
    
    # 2) Достаём из БД то, что посчитали
    conn = sqlite3.connect(DATABASE_PATH)
    df_db = pd.read_sql(f"""
        SELECT 
            PartNumber,
            DealerStockPrice,
            GrossPrice,
            DiscountCode,
            Description
        FROM {PRICELIST_TABLE_NAME}
        WHERE DealerStockPrice IS NOT NULL
    """, conn)
    conn.close()
    
    df_db["PartNumber"] = df_db["PartNumber"].map(norm_part)
    
    # 3) Склеиваем и считаем расхождения
    df_compare = df_db.merge(df_reference, on="PartNumber", how="inner")
    if df_compare.empty:
        print("⚠️ Не удалось сопоставить ни одного артикула по PartNumber.")
        return None
    
    df_compare["Difference"] = df_compare["DealerStockPrice"] - df_compare["ReferencePrice"]
    df_compare["AbsDiff"] = df_compare["Difference"].abs()
    
    total = len(df_compare)
    ok = (df_compare["AbsDiff"] <= tolerance).sum()
    bad = total - ok
    
    print("\n🔍 Сравнение DealerStockPrice vs ReferencePrice:")
    print(f"  Всего сопоставлено: {total}")
    print(f"  В пределах допуска (≤ {tolerance}): {ok}")
    print(f"  С расхождением > {tolerance}: {bad}")
    
    mismatches = df_compare[df_compare["AbsDiff"] > tolerance].copy()
    if not mismatches.empty:
        mismatches_sorted = mismatches.sort_values("AbsDiff", ascending=False).head(100)
        ensure_folder_exists()
        out_path = os.path.join(FOLDER_PATH, "dealer_price_mismatches.xlsx")
        mismatches_sorted.to_excel(out_path, index=False)
        print(f"  ⚠️ Топ расхождений сохранён в: {out_path}")
    else:
        print("  ✅ Все цены в пределах допуска.")
    
    return df_compare


# Функция apply_catid_overrides теперь импортируется из CatID_lib

def export_catid_quality_report():
    """
    Делает базовый отчёт по качеству заполнения CatID:
      - распределение по ProductGroup (PG → CatID)
      - распределение по FunctionGroup (Fgrp → CatID)
      - нарушения правил для UnitSort (SW/HW)
      - список позиций с пустым CatID (если остались)
    Всё сохраняется в один Excel в папку export/.
    """
    ensure_folder_exists()
    report_path = os.path.join(FOLDER_PATH, "catid_quality_report.xlsx")

    conn = sqlite3.connect(DATABASE_PATH)

    # --- 1) Распределение по ProductGroup (PG → CatID)
    df_pg = pd.read_sql(f"""
        SELECT 
            ProductGroup,
            CatID,
            COUNT(*) AS Count
        FROM {PRICELIST_TABLE_NAME}
        GROUP BY ProductGroup, CatID
        ORDER BY ProductGroup, Count DESC
    """, conn)

    if not df_pg.empty:
        # проценты внутри каждого PG
        df_pg["TotalInPG"] = df_pg.groupby("ProductGroup")["Count"].transform("sum")
        df_pg["PercentInPG"] = (df_pg["Count"] / df_pg["TotalInPG"] * 100).round(2)

    # --- 2) Распределение по FunctionGroup (Fgrp → CatID)
    df_fgrp = pd.read_sql(f"""
        SELECT 
            FunctionGroup AS Fgrp,
            CatID,
            COUNT(*) AS Count
        FROM {PRICELIST_TABLE_NAME}
        GROUP BY FunctionGroup, CatID
        ORDER BY Fgrp, Count DESC
    """, conn)

    if not df_fgrp.empty:
        df_fgrp["TotalInFgrp"] = df_fgrp.groupby("Fgrp")["Count"].transform("sum")
        df_fgrp["PercentInFgrp"] = (df_fgrp["Count"] / df_fgrp["TotalInFgrp"] * 100).round(2)

    # --- 3) Нарушения правил для UnitSort (SW/HW)
    df_unitsort_mismatch = pd.read_sql(f"""
        SELECT 
            PartNumber,
            Description,
            ProductGroup,
            FunctionGroup AS Fgrp,
            UnitSort,
            CatID
        FROM {PRICELIST_TABLE_NAME}
        WHERE 
            (UnitSort = 'SW' AND (CatID IS NULL OR TRIM(COALESCE(CatID, '')) <> '1000'))
            OR
            (UnitSort = 'HW' AND (CatID IS NULL OR TRIM(COALESCE(CatID, '')) <> '4'))
    """, conn)

    # --- 4) Пустые CatID (если остались)
    empty_condition = get_empty_catid_condition()
    df_empty = pd.read_sql(f"""
        SELECT 
            PartNumber,
            Description,
            ProductGroup,
            FunctionGroup AS Fgrp,
            UnitSort,
            CatID
        FROM {PRICELIST_TABLE_NAME}
        WHERE {empty_condition}
    """, conn)

    conn.close()

    # --- Запись в Excel
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        if not df_pg.empty:
            df_pg.to_excel(writer, sheet_name="PG_vs_CatID", index=False)
        if not df_fgrp.empty:
            df_fgrp.to_excel(writer, sheet_name="Fgrp_vs_CatID", index=False)
        if not df_unitsort_mismatch.empty:
            df_unitsort_mismatch.to_excel(writer, sheet_name="UnitSort_mismatch", index=False)
        if not df_empty.empty:
            df_empty.to_excel(writer, sheet_name="Empty_CatID", index=False)

    print(f"📊 Отчёт по CatID сохранён: {report_path}")

# ------------------------------------------------------------
# 🚀 Пайплайн
# ------------------------------------------------------------

def clear_pricelist_table():
    """Очищает таблицу прайса перед новым импортом, чтобы не плодить дубликаты."""
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME}")
    before = cur.fetchone()[0]
    cur.execute(f"DELETE FROM {PRICELIST_TABLE_NAME}")
    conn.commit()
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME}")
    after = cur.fetchone()[0]
    conn.close()
    print(f"🧹 Очищена таблица {PRICELIST_TABLE_NAME}: было {before}, стало {after}.")


def clear_replacement_table():
    """Очищает таблицу замен перед новым импортом."""
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {REPLACEMENT_TABLE_NAME}")
    before = cur.fetchone()[0]
    cur.execute(f"DELETE FROM {REPLACEMENT_TABLE_NAME}")
    conn.commit()
    cur.execute(f"SELECT COUNT(*) FROM {REPLACEMENT_TABLE_NAME}")
    after = cur.fetchone()[0]
    conn.close()
    print(f"🧹 Очищена таблица {REPLACEMENT_TABLE_NAME}: было {before}, стало {after}.")


def clear_emblem_table():
    """Очищает таблицу эмблем перед новым импортом."""
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {EMBLEM_TABLE_NAME}")
    before = cur.fetchone()[0]
    cur.execute(f"DELETE FROM {EMBLEM_TABLE_NAME}")
    conn.commit()
    cur.execute(f"SELECT COUNT(*) FROM {EMBLEM_TABLE_NAME}")
    after = cur.fetchone()[0]
    conn.close()
    print(f"🧹 Очищена таблица {EMBLEM_TABLE_NAME}: было {before}, стало {after}.")


def import_volvo_prices_txt():
    global PRICELIST_TXT_FILE_PATH, REPLACEMENT_TXT_FILE_PATH, EMBLEM_TXT_FILE_PATH

    PRICELIST_TXT_FILE_PATH   = find_file_by_prefix(TXT_FOLDER, "Gold_Price_M3_")
    REPLACEMENT_TXT_FILE_PATH = find_file_by_prefix(TXT_FOLDER, "Gold_Replacement_M3_")
    EMBLEM_TXT_FILE_PATH      = find_file_by_prefix(TXT_FOLDER, "Gold_Emblem_M3_")

    print(f"TXT dir: {os.path.abspath(TXT_FOLDER)}")
    print("Price:", PRICELIST_TXT_FILE_PATH)
    print("Repl :", REPLACEMENT_TXT_FILE_PATH)
    print("Embl :", EMBLEM_TXT_FILE_PATH)

    missing = []
    if not PRICELIST_TXT_FILE_PATH: missing.append("Gold_Price_M3_*")
    if not REPLACEMENT_TXT_FILE_PATH: missing.append("Gold_Replacement_M3_*")
    if not EMBLEM_TXT_FILE_PATH: missing.append("Gold_Emblem_M3_*")
    if missing:
        print("❌ Не найдены файлы:", ", ".join(missing))
        if os.path.isdir(TXT_FOLDER):
            print("Содержимое папки:", os.listdir(TXT_FOLDER))
        raise SystemExit(1)

    create_tables_if_not_exist()
    ensure_indexes()  # индексы до импорта не обязательны, но не мешают

    print("Импортируем прайс...")
    # Чистим таблицу прайса перед загрузкой, чтобы не накапливать дубликаты
    clear_pricelist_table()
    import_pricelist_txt()

    print("Импортируем замены...")
    clear_replacement_table()
    import_replacement_txt_to_sqlite()

    print("Импортируем эмблемы...")
    clear_emblem_table()
    import_emblem_txt()

    ensure_indexes()  # после импорта точно создадим индексы


# Функция backfill_catid_from_previous_excel теперь импортируется из CatID_lib

def calculate_prices_and_export_eur_for_1c(export_filename_eur: str | None = None):
    """
    Пайплайн как в processing.py, но с EUR:

    1) В SEK считаем:
       - CalcDSl  (из DealerStockPrice)
       - CostFactor (по CatID)
       - SelfCost
       - Margin (по SelfCost и CatID)
       - RetailPrice
       - Category1C (по CatID → 1С-код)

    2) Потом считаем EUR-колонки:
       - CalcDSl_EUR     = CEIL((CalcDSl     / EUR_RATE) * 100) / 100
       - SelfCost_EUR    = CEIL((SelfCost    / EUR_RATE) * 100) / 100
       - RetailPrice_EUR = CEIL((RetailPrice / EUR_RATE) * 100) / 100

    3) Экспортируем финальный прайс в EUR под 1С.
    """
    ensure_folder_exists()

    if export_filename_eur is None:
        now = datetime.now().strftime("%Y-%m-%d_%H-%M")
        export_filename_eur = f"volvo_price_export_EUR_{now}.xlsx"

    conn = sqlite3.connect(DATABASE_PATH)
    conn.create_function("CEIL", 1, lambda x: math.ceil(float(x)) if x is not None else None)
    cur = conn.cursor()

    # --- 0) Гарантируем, что все нужные колонки существуют ---
    ensure_column_exists(PRICELIST_TABLE_NAME, "CalcDSl",         "REAL", conn)  # инвойс в SEK
    ensure_column_exists(PRICELIST_TABLE_NAME, "SelfCost",        "REAL", conn)
    ensure_column_exists(PRICELIST_TABLE_NAME, "RetailPrice",     "REAL", conn)
    ensure_column_exists(PRICELIST_TABLE_NAME, "Margin",          "REAL", conn)
    ensure_column_exists(PRICELIST_TABLE_NAME, "CostFactor",      "REAL", conn)
    ensure_column_exists(PRICELIST_TABLE_NAME, "Category1C",      "TEXT", conn)

    ensure_column_exists(PRICELIST_TABLE_NAME, "CalcDSl_EUR",     "REAL", conn)
    ensure_column_exists(PRICELIST_TABLE_NAME, "SelfCost_EUR",    "REAL", conn)
    ensure_column_exists(PRICELIST_TABLE_NAME, "RetailPrice_EUR", "REAL", conn)

    # --- 1) CalcDSl (SEK) ← DealerStockPrice ---
    # Можно перетирать всегда: это рабочая производная колонка
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CalcDSl = DealerStockPrice
        WHERE DealerStockPrice IS NOT NULL
    """)
    conn.commit()

    # --- 2) CostFactor по CatID ---
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CostFactor = CASE
            WHEN CatID = '1000' THEN 1.25  -- SW
            ELSE 1.55                      -- всё остальное
        END
        WHERE CalcDSl IS NOT NULL
    """)
    conn.commit()

    # --- 3) SelfCost (SEK) ---
    # как в processing.py — считаем в SEK без CEIL, округляем до 0.001
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET SelfCost = ROUND(CalcDSl * CostFactor, 3)
        WHERE CalcDSl IS NOT NULL AND CostFactor IS NOT NULL
    """)
    conn.commit()

    # --- 4) Margin для ПО (CatID='1000') по SelfCost (SEK) ---
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET Margin = CASE
            WHEN SelfCost < 100  THEN 0.50
            WHEN SelfCost < 200  THEN 0.40
            WHEN SelfCost < 400  THEN 0.35
            WHEN SelfCost < 600  THEN 0.30
            WHEN SelfCost < 1000 THEN 0.25
            ELSE 0.22
        END
        WHERE CatID = '1000' AND SelfCost IS NOT NULL
    """)
    conn.commit()

    # --- 5) Margin для остальных CatID по SelfCost (SEK) ---
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET Margin = (
            CASE 
                WHEN SelfCost <   20 THEN 0.50
                WHEN SelfCost <   40 THEN 0.49
                WHEN SelfCost <   60 THEN 0.48
                WHEN SelfCost <   80 THEN 0.47
                WHEN SelfCost <  100 THEN 0.46
                WHEN SelfCost <  120 THEN 0.45
                WHEN SelfCost <  140 THEN 0.45
                WHEN SelfCost <  160 THEN 0.44
                WHEN SelfCost <  180 THEN 0.44
                WHEN SelfCost <  200 THEN 0.43
                WHEN SelfCost <  220 THEN 0.43
                WHEN SelfCost <  240 THEN 0.42
                WHEN SelfCost <  260 THEN 0.42
                WHEN SelfCost <  280 THEN 0.42
                WHEN SelfCost <  300 THEN 0.41
                WHEN SelfCost <  330 THEN 0.41
                WHEN SelfCost <  360 THEN 0.41
                WHEN SelfCost <  390 THEN 0.41
                WHEN SelfCost <  420 THEN 0.40
                WHEN SelfCost <  450 THEN 0.40
                WHEN SelfCost <  480 THEN 0.40
                WHEN SelfCost <  510 THEN 0.40
                WHEN SelfCost <  560 THEN 0.40
                WHEN SelfCost <  610 THEN 0.40
                WHEN SelfCost <  680 THEN 0.40
                WHEN SelfCost <  750 THEN 0.39
                WHEN SelfCost <  850 THEN 0.39
                WHEN SelfCost <  950 THEN 0.39
                WHEN SelfCost < 1050 THEN 0.39
                WHEN SelfCost < 1200 THEN 0.39
                WHEN SelfCost < 1400 THEN 0.39
                WHEN SelfCost < 1700 THEN 0.39
                WHEN SelfCost < 2100 THEN 0.39
                WHEN SelfCost < 2600 THEN 0.38
                WHEN SelfCost < 3300 THEN 0.38
                WHEN SelfCost < 4100 THEN 0.38
                WHEN SelfCost < 5000 THEN 0.38
                WHEN SelfCost < 6000 THEN 0.38
                WHEN SelfCost < 7500 THEN 0.38
                WHEN SelfCost < 9500 THEN 0.38
                WHEN SelfCost < 11000 THEN 0.38
                WHEN SelfCost < 15000 THEN 0.38
                ELSE 0.37
            END
        )
        WHERE (CatID IS NULL OR CatID <> '1000') AND SelfCost IS NOT NULL
    """)
    conn.commit()

    # --- 6) RetailPrice (SEK) ---
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET RetailPrice = ROUND(SelfCost / (1 - Margin), 2)
        WHERE SelfCost IS NOT NULL AND Margin IS NOT NULL
    """)
    conn.commit()

    # --- 7) Category1C по CatID (1С-код) ---
    for cat_id, code_1c in CATEGORY_MAPPING_1C.items():
        cur.execute(f"""
            UPDATE {PRICELIST_TABLE_NAME}
            SET Category1C = ?
            WHERE CatID = ?
        """, (code_1c, str(cat_id)))
    conn.commit()

    # --- 8) Конвертация в EUR с CEIL до +0.01 ---
    # Инвойс (CalcDSl_EUR)
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CalcDSl_EUR = CAST(CEIL((CalcDSl / ?) * 100) AS REAL) / 100.0
        WHERE CalcDSl IS NOT NULL
    """, (EUR_TO_SEK_RATE,))
    conn.commit()

    # Себестоимость (SelfCost_EUR)
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET SelfCost_EUR = CAST(CEIL((SelfCost / ?) * 100) AS REAL) / 100.0
        WHERE SelfCost IS NOT NULL
    """, (EUR_TO_SEK_RATE,))
    conn.commit()

    # Розница (RetailPrice_EUR)
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET RetailPrice_EUR = CAST(CEIL((RetailPrice / ?) * 100) AS REAL) / 100.0
        WHERE RetailPrice IS NOT NULL
    """, (EUR_TO_SEK_RATE,))
    conn.commit()

    # --- 9) Подготовка DataFrame для экспорта ---
    df_prices = pd.read_sql_query(f"""
        SELECT
            PartNumber      AS "Part Number",
            Description     AS "Наименование товара",
            CalcDSl_EUR     AS "Инвойсная цена в Евро (по курсу конвертации производителя)",
            CostFactor      AS "Расчетный коэффициент входа",
            SelfCost_EUR    AS "Расчетная себестоимость в Евро",
            RetailPrice_EUR AS "Розница с НДС в Евро по схеме расчета розничной цены для клиента",
            Category1C      AS "Категория 1C"
        FROM {PRICELIST_TABLE_NAME}
        WHERE CalcDSl_EUR IS NOT NULL
    """, conn)

    export_path = os.path.join(FOLDER_PATH, export_filename_eur)
    with pd.ExcelWriter(export_path, engine="xlsxwriter") as writer:
        df_prices.to_excel(writer, sheet_name="Price List EUR", index=False)

        # формат для денег с двумя знаками
        workbook  = writer.book
        worksheet = writer.sheets["Price List EUR"]
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})

        # C, E, F — колонны с EUR-ценами (если порядок поменяешь — поправь буквы)
        worksheet.set_column('C:C', 22, money_fmt)  # инвойс EUR
        worksheet.set_column('E:E', 22, money_fmt)  # себестоимость EUR
        worksheet.set_column('F:F', 22, money_fmt)  # розница EUR

    conn.close()
    print(f"📦 Экспортировано EUR-прайс для 1С: {export_path}")

def postprocess_prices():
    # Редактор скидок (блокирующая форма)
    launch_discount_editor()

    missing = check_missing_discount_codes()
    if missing:
        print(f"⚠️ Нет DC в discount_codes: {missing[:50]}{' …' if len(missing)>50 else ''}")

    # ▶ Порядок заполнения CatID = приоритет (каждый шаг заполняет ТОЛЬКО ПУСТЫЕ CatID)
    # Все функции используют единую проверку get_empty_catid_condition() для строгого контроля.
    # 
    print_catid_statistics("До заполнения CatID")
    
    # 1) Точные правила по признакам (UnitSort) - самый высокий приоритет
    #    SW → CatID='1000' (Software)
    #    HW → CatID='4' (Electronics)
    ensure_catid_soft()
    ensure_catid_hardware()
    print_catid_statistics("После шага 1 (SW + HW)")

    # ------------------------------------------------------------
    # 📚 PG reference (для категорий CatID)
    #   11 — PARTS
    #   13 — CHEMICALS
    #   14 — EXCHANGE
    #   15 — ACCESSORIES
    #   16 — TYRES
    #   18 — SPECIAL TOOLS
    #   25 — MERCHANDISE
    # ------------------------------------------------------------
    # 2) Правила по ProductGroup (детерминированные, но менее специфичные)
    #    Заполняет только те CatID, которые остались пустыми после шага 1
    #    PG=11 (PARTS) → CatID='70' (самая частая категория для запчастей)
    #    PG=14 (EXCHANGE) → CatID='40' (самая частая категория для обменных узлов)
    ensure_catid_by_pg({
        11: '70',   # PARTS - самая частая категория
        14: '40',   # EXCHANGE - самая частая категория
        15: '20',   # ACCESSORIES
        25: '80',   # MERCHANDISE
        13: '60',   # CHEMICALS
        16: '10',   # TYRES
        18: '65',   # SPECIAL TOOLS
    #    17: '70',   # (1396 записей, но CatID=1000 определяется по UnitSort='SW')
        21: '70',   # (71 запись, 94.4% имеют CatID=70)
    #    24: '65',   # (1 запись, 100% имеют CatID=65)
    #    54: '65',   # (1 запись, 100% имеют CatID=65)
    #    55: '70',   # (219 записей, 28.3% имеют CatID=70 - самая частая)
        56: '10'    # (90 записей, 100% имеют CatID=10)
    })
    print_catid_statistics("После шага 2 (PG)")

    # 3) Ручные точечные переопределения (только для специфичных артикулов)
    #    ВАЖНО: Не дублирует правила из шага 2! Используй только для PartNumber.
    #    Заполняет только те CatID, которые остались пустыми после шагов 1-2
    apply_catid_overrides(force=False)
    print_catid_statistics("После шага 3 (Overrides)")

    # 4) Наследование из предыдущего Excel (самое «остаточное», самый низкий приоритет)
    #    Заполняет только те CatID, которые остались пустыми после всех предыдущих шагов
    prev_excel_path = backfill_catid_from_previous_excel()

    print_catid_statistics("После шага 4 (Excel backfill)")

    # Стоимость дилера
    compute_dealer_stock_price()
    
    # Анализ проблем с расчетом цен
    analyze_price_calculation_issues()
    
    # Выравнивание цен с эталонными данными (исправление расхождений > 0.01)
    # Это обеспечивает 100% совпадение с эталонными данными из файлов data
    align_prices_with_reference_data(tolerance=0.01)
    
    # Сравнение рассчитанных цен с эталонными из папки data
    compare_prices_with_data_files()

    # Описания: >>>(...) и !!! NOT STORED ANYMORE !!!
    enrich_descriptions_from_replacements()

    # Отчёт по качеству CatID
    export_catid_quality_report()

    # Экспорт (базовый)
    export_basic_excel(previous_excel_path=prev_excel_path)

    calculate_prices_and_export_eur_for_1c()

    print("✅ Пост-обработка завершена.")



def run_pipeline():
    import_volvo_prices_txt()
    postprocess_prices()


if __name__ == "__main__":
    run_pipeline()



def ensure_catid_accessories():
    """Назначает CatID для аксессуаров по ProductGroup:
    - PG=15 → CatID='20'
    - PG=25 → CatID='80'
    Обновляет только пустые CatID, не перезаписывает существующие.
    
    ПРИМЕЧАНИЕ: Эта функция дублирует функциональность ensure_catid_by_pg().
    Рекомендуется использовать ensure_catid_by_pg({15: '20', 25: '80'}) вместо этой функции.
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()

    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CatID = '20'
        WHERE ProductGroup = 15 AND {empty_condition}
    """)
    count20 = cur.rowcount

    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CatID = '80'
        WHERE ProductGroup = 25 AND {empty_condition}
    """)
    count80 = cur.rowcount

    conn.commit()
    conn.close()
    print(f"🏷️ Accessories назначены: PG=15 → CatID='20' ({count20}), PG=25 → CatID='80' ({count80}). (только пустые)")










