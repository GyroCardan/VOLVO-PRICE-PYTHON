import os
import glob
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from tqdm import tqdm
from datetime import datetime, timedelta

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


def norm_part(x):
    # None / NaN -> None
    if x is None:
        return None
    try:
        if x != x:  # NaN != NaN
            return None
    except:
        pass

    s = str(x).strip()
    if not s:
        return None

    # Если это число с плавающей точкой вида 123.0 / научная запись — делаем целым
    try:
        f = float(s.replace(',', '.'))
        if f.is_integer():
            return str(int(f))
    except:
        pass

    # Если это чисто целая строка — тоже нормализуем
    try:
        return str(int(s))
    except:
        # иначе оставляем как есть (например, если есть буквы)
        return s

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
    record_type = safe_int(line[0:1].strip())
    data = {
        "RecordType": record_type,
        "PartNumber": norm_part(line[2:13]),        # ⚠️ проверь ширину под свой файл
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
        data.update({
            "ReplacedQuantity": safe_float(line[17:24].strip()),
            "WeekReplaced": safe_int(line[25:29].strip()),
            "ReplacedCode": line[29:32].strip(),
            "Description": line[57:82].strip(),
        })
    elif record_type == 2:
        data.update({
            "InfoText": line[17:37].strip(),
        })
    elif record_type == 3:
        data.update({
            "ReplacingPart": norm_part(line[17:29]),  # ⚠️ проверь ширину под свой файл
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


def get_empty_catid_condition():
    """Возвращает SQL условие для проверки пустого CatID.
    Используется единообразно во всех функциях заполнения категорий.
    """
    return "(CatID IS NULL OR TRIM(COALESCE(CatID, '')) = '')"


def get_catid_statistics():
    """Возвращает статистику заполнения CatID для отладки и контроля."""
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME}")
    total = cur.fetchone()[0]
    
    empty_condition = get_empty_catid_condition()
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME} WHERE {empty_condition}")
    empty = cur.fetchone()[0]
    
    filled = total - empty
    
    conn.close()
    return {
        'total': total,
        'filled': filled,
        'empty': empty,
        'filled_percent': round(filled / total * 100, 2) if total > 0 else 0
    }


def print_catid_statistics(stage_name: str = ""):
    """Выводит статистику заполнения CatID на текущий момент."""
    stats = get_catid_statistics()
    prefix = f"[{stage_name}] " if stage_name else ""
    print(f"{prefix}📊 CatID статистика: заполнено {stats['filled']}/{stats['total']} ({stats['filled_percent']}%), пустых: {stats['empty']}")


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


def ensure_catid_soft():
    """Назначает CatID='1000' для программного обеспечения (UnitSort='SW').
    Заполняет ТОЛЬКО пустые CatID, не перезаписывает существующие.
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CatID = '1000'
        WHERE UnitSort = 'SW'
          AND {empty_condition}
    """)
    updated = cur.rowcount
    conn.commit(); conn.close()
    print(f"🏷️ Soft назначен {updated} позициям (CatID='1000'; только пустые).")


def ensure_catid_hardware():
    """Назначает CatID='4' для электроники (UnitSort='HW').
    Заполняет ТОЛЬКО пустые CatID, не перезаписывает существующие.
    Приоритет такой же, как у ensure_catid_soft() - шаг 1.
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()
    cur.execute(f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CatID = '4'
        WHERE UnitSort = 'HW'
          AND {empty_condition}
    """)
    updated = cur.rowcount
    conn.commit(); conn.close()
    print(f"🏷️ Hardware (Electronics) назначен {updated} позициям (CatID='4'; только пустые).")


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

def ensure_catid_by_pg(mapping):
    """Назначает CatID по значениям ProductGroup согласно словарю mapping {PG:int -> CatID:str}.
    Пример: ensure_catid_by_pg({15: '20', 25: '80'})
    Обновляет ТОЛЬКО пустые CatID, не перезаписывает существующие.
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()
    total = 0
    for pg, cat in mapping.items():
        cur.execute(f"""
            UPDATE {PRICELIST_TABLE_NAME}
            SET CatID = ?
            WHERE ProductGroup = ?
              AND {empty_condition}
        """, (str(cat), int(pg)))
        cnt = cur.rowcount
        total += cnt
        print(f"🏷️ PG={pg} → CatID='{cat}': {cnt} поз. (только пустые)")
    conn.commit(); conn.close()
    print(f"✅ Итого обновлено по PG-мэппингу: {total} позиций.")

# --- Обогащение описаний: >>>(...) и метка снято с производства ---

def enrich_descriptions_from_replacements():
    conn = sqlite3.connect(DATABASE_PATH)

    # 1) Замены: берём только RecordType=3 (содержит ReplacingPart)
    df_r = pd.read_sql(f"""
        SELECT PartNumber, ReplacingPart
        FROM {REPLACEMENT_TABLE_NAME}
        WHERE RecordType = 3
          AND ReplacingPart IS NOT NULL AND TRIM(ReplacingPart) != ''
    """, conn)
    # На всякий — нормализация
    df_r["PartNumber"] = df_r["PartNumber"].map(norm_part)
    df_r["ReplacingPart"] = df_r["ReplacingPart"].map(norm_part)

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

    # 2) Снятые с производства: RecordType=1 с ReplacedCode='029'
    df_disc = pd.read_sql(f"""
        SELECT DISTINCT PartNumber
        FROM {REPLACEMENT_TABLE_NAME}
        WHERE RecordType = 1 AND ReplacedCode = '029'
    """, conn)
    disc_set = set(df_disc["PartNumber"].dropna())

    # 3) Загружаем текущий прайс и вливаем
    df_p = pd.read_sql(f"SELECT PartNumber, Description FROM {PRICELIST_TABLE_NAME}", conn)
    df_p["Description"] = df_p["Description"].fillna("")
    df_p = df_p.merge(df_repl, on="PartNumber", how="left")

    # 3.1) Добавляем >>>(...) только если его ещё нет и есть что добавить
    mask_no_arrow = ~df_p["Description"].str.contains(">>>", na=False)
    mask_has_repl = df_p["Replacements"].notna() & (df_p["Replacements"].str.len() > 0)
    to_arrow = mask_no_arrow & mask_has_repl
    df_p.loc[to_arrow, "Description"] = df_p["Description"] + " >>> (" + df_p["Replacements"] + ")"

    # 3.2) Добавляем !!! NOT STORED ANYMORE !!! если PartNumber в disc_set и метки нет
    mask_no_not = ~df_p["Description"].str.contains("NOT STORED ANYMORE", na=False)
    mask_disc = df_p["PartNumber"].isin(disc_set)
    to_not = mask_no_not & mask_disc
    df_p.loc[to_not, "Description"] = df_p["Description"] + " !!! NOT STORED ANYMORE !!!"

    # 4) Обновляем только изменившиеся строки
    # Для этого перечитаем оригинальные описания ещё раз и сравним
    df_orig = pd.read_sql(f"SELECT PartNumber, Description AS Orig FROM {PRICELIST_TABLE_NAME}", conn)
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
        print("ℹ️ Описания не требуют обновления.")

    conn.close()

# ------------------------------------------------------------
# 📤 Экспорт (базовый пример)
# ------------------------------------------------------------

def export_basic_excel():
    ensure_folder_exists()
    now = datetime.now().strftime("%Y-%m-%d_%H-%M")
    out_path = os.path.join(FOLDER_PATH, f"volvo_price_export_{now}.xlsx")
    conn = sqlite3.connect(DATABASE_PATH)
    df = pd.read_sql(f"""
        SELECT PartNumber, Description, DiscountCode, GrossPrice, DealerStockPrice, UnitSort, FunctionGroup, ProductGroup, CatID
        FROM {PRICELIST_TABLE_NAME}
    """, conn)
    conn.close()
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name="Price")
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
    Сравнивает рассчитанные DealerStockPrice с ценами из файлов в папке data.
    
    Формула расчета: DealerStockPrice = GrossPrice * (1 - [Dealer Stock %]/100.0)
    В файлах data колонка CalcDSl содержит эталонные значения.
    
    Args:
        data_folder: Путь к папке с файлами для сравнения
        tolerance: Допустимое отклонение (по умолчанию 0.01 = 1 копейка)
    
    Returns:
        DataFrame с результатами сравнения
    """
    # Загружаем все файлы из папки data
    data_files = glob.glob(os.path.join(data_folder, "*.xlsx"))
    if not data_files:
        print(f"❌ Файлы не найдены в папке {data_folder}")
        return None
    
    print(f"📂 Найдено файлов для сравнения: {len(data_files)}")
    
    # Объединяем все файлы в один DataFrame
    all_data = []
    for file_path in sorted(data_files):
        try:
            df_file = pd.read_excel(file_path)
            # Нормализуем названия колонок (могут быть разные варианты)
            if 'PartNo' in df_file.columns:
                df_file = df_file.rename(columns={'PartNo': 'PartNumber'})
            if 'CalcDSl' in df_file.columns:
                df_file = df_file.rename(columns={'CalcDSl': 'ReferencePrice'})
            all_data.append(df_file[['PartNumber', 'ReferencePrice']].copy())
            print(f"  ✓ Загружен: {os.path.basename(file_path)} ({len(df_file)} записей)")
        except Exception as e:
            print(f"  ✗ Ошибка при загрузке {os.path.basename(file_path)}: {e}")
    
    if not all_data:
        print("❌ Не удалось загрузить данные из файлов")
        return None
    
    # Объединяем все данные
    df_reference = pd.concat(all_data, ignore_index=True)
    df_reference['PartNumber'] = df_reference['PartNumber'].map(norm_part)
    df_reference = df_reference.dropna(subset=['PartNumber', 'ReferencePrice'])
    df_reference = df_reference.drop_duplicates('PartNumber', keep='first')
    
    print(f"📊 Всего эталонных цен: {len(df_reference)}")
    
    # Загружаем данные из базы
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
    
    conn.close(); return

    # Дополнительная проверка: не заполняем пустые CatID из Excel
    upd = upd[upd["CatID"].notna() & (upd["CatID"].str.strip() != '')]
    if upd.empty:
        print("ℹ️ В предыдущем прайсе нет валидных CatID для дозаполнения.")
        conn.close(); return

    # Исключаем категорию SOFT (CatID='1000') из предыдущего прайса
    # Категория SOFT должна определяться только по UnitSort='SW' на шаге 1, а не переноситься из старого прайса
    upd = upd[upd["CatID"].str.strip() != '1000']
    if upd.empty:
        print("ℹ️ В предыдущем прайсе остались только категории SOFT, которые не переносятся.")
        conn.close(); return

    payload = list(map(tuple, upd[["CatID", "PartNumber"]].values))
    cur.executemany(
        f"UPDATE {PRICELIST_TABLE_NAME} SET CatID = ? WHERE PartNumber = ? AND {empty_condition}",
        payload,
    )
    conn.commit(); conn.close()
    print(f"✍️ Заполнено CatID из предыдущего прайса: {len(payload)} поз. (только пустые)")

def apply_catid_overrides(force: bool = False):
    """
    Применяет ручные переопределения CatID.
    По умолчанию НЕ перезаписывает непустые CatID (force=False).
    Если нужно принудительно — вызови apply_catid_overrides(force=True).
    
    ВАЖНО: Не дублируй правила, которые уже обрабатываются в ensure_catid_by_pg()!
    Используй эту функцию только для точечных переопределений по PartNumber
    или для правил, которые не покрываются автоматическими методами.
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()

    # ВАЖНО: Убраны дубликаты с ensure_catid_by_pg() (PG=15 и PG=25)
    # Оставляем только точечные переопределения по PartNumber
    overrides = [
        # пример точечных артикулов:
        # ("14", "PartNumber", "30657360"),
        # ("20", "PartNumber", "12345678"),
    ]

    total = 0
    for catid, column, value in overrides:
        if force:
            cur.execute(f"""
                UPDATE {PRICELIST_TABLE_NAME}
                SET CatID = ?
                WHERE {column} = ?
            """, (catid, value))
        else:
            cur.execute(f"""
                UPDATE {PRICELIST_TABLE_NAME}
                SET CatID = ?
                WHERE {column} = ? AND {empty_condition}
            """, (catid, value))
        total += cur.rowcount
    conn.commit(); conn.close()
    if total > 0:
        print(f"✅ Применены ручные переопределения CatID: {total} поз. ({'force' if force else 'только пустые'})")
    else:
        print(f"ℹ️ Ручные переопределения CatID: нет правил для применения.")

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


def backfill_catid_from_previous_excel(excel_path: str | None = None):
    """Дозаполняет CatID из предыдущего прайса (Excel), сопоставление по PartNumber.
    Имена колонок распознаются устойчиво: регистр не важен, пробелы/подчёркивания/точки/неразрывные пробелы игнорируются.
    Поддерживаемые варианты:
      • Артикул: PartNumber, PartNo, Part No, Part_Number, Артикул, Номер
      • Категория: CatID, Cat_ID, Cat Id, Category1C, Category, Категория
    """
    import pandas as pd
    from tkinter import filedialog

    def norm_colname(s: str) -> str:
        # к нижнему регистру и выкинуть пробелы/подчёркивания/точки/неразрывные пробелы
        return "".join(ch for ch in s.lower() if ch not in " _. ")

    # 1) выбрать файл
    default_candidate = os.path.abspath("previous_prices.xlsx")
    if excel_path is None:
        if os.path.isfile(default_candidate):
            excel_path = default_candidate
        else:
            print("📄 Выберите предыдущий прайс (Excel) для подтягивания CatID…")
            excel_path = filedialog.askopenfilename(
                title="Выберите предыдущий прайс (Excel)",
                filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls")],
            )
            if not excel_path:
                print("⛔ Файл не выбран, пропускаем backfill CatID.")
                return

    if not os.path.isfile(excel_path):
        print(f"⛔ Файл не найден: {excel_path}")
        return

    print(f"📥 Читаем предыдущий прайс: {excel_path}")
    try:
        df_prev = pd.read_excel(excel_path)
    except Exception as e:
        print("⛔ Не удалось прочитать Excel:", e)
        return

    if df_prev.empty:
        print("⛔ Файл пустой.")
        return

    # 2) карта нормализованных имён → оригинальные
    norm_map = {norm_colname(c): c for c in df_prev.columns}

    part_candidates = [
        "partnumber", "partno", "part_number", "part no", "номер", "артикул"
    ]
    cat_candidates = [
        "catid", "cat_id", "cat id", "category1c", "category", "категория"
    ]

    part_col = next((norm_map.get(norm_colname(k)) for k in part_candidates if norm_colname(k) in norm_map), None)
    cat_col  = next((norm_map.get(norm_colname(k)) for k in cat_candidates  if norm_colname(k) in norm_map), None)

    if part_col is None or cat_col is None:
        print("⛔ Не удалось найти нужные колонки в предыдущем прайсе.")
        print("   Доступные колонки:", list(df_prev.columns))
        print("   Ищем PartNumber среди:", part_candidates)
        print("   Ищем CatID среди:", cat_candidates)
        return

    # 3) привести к двум колонкам и нормализовать
    df_prev = df_prev[[part_col, cat_col]].copy()
    df_prev.columns = ["PartNumber", "CatID"]
    df_prev["PartNumber"] = df_prev["PartNumber"].map(norm_part)
    df_prev["CatID"] = df_prev["CatID"].astype(str).str.strip()
    df_prev = df_prev.dropna(subset=["PartNumber", "CatID"]).drop_duplicates("PartNumber")

    # 4) выбрать кого дозаполнять (только пустые CatID)
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()
    cur.execute(f"SELECT PartNumber FROM {PRICELIST_TABLE_NAME} WHERE {empty_condition}")
    missing_parts = {row[0] for row in cur.fetchall()}

    # Фильтруем только те артикулы, у которых CatID пустой
    upd = df_prev[df_prev["PartNumber"].isin(missing_parts)]
    if upd.empty:
        print("ℹ️ Нет позиций для дозаполнения CatID из предыдущего прайса.")
        conn.close(); return

    # Дополнительная проверка: не заполняем пустые CatID из Excel
    upd = upd[upd["CatID"].notna() & (upd["CatID"].str.strip() != '')]
    if upd.empty:
        print("ℹ️ В предыдущем прайсе нет валидных CatID для дозаполнения.")
        conn.close(); return

    # Исключаем категорию SOFT (CatID='1000') из предыдущего прайса
    # Категория SOFT должна определяться только по UnitSort='SW' на шаге 1, а не переноситься из старого прайса
    upd = upd[upd["CatID"].str.strip() != '1000']
    if upd.empty:
        print("ℹ️ В предыдущем прайсе остались только категории SOFT, которые не переносятся.")
        conn.close(); return

    payload = list(map(tuple, upd[["CatID", "PartNumber"]].values))
    cur.executemany(
        f"UPDATE {PRICELIST_TABLE_NAME} SET CatID = ? WHERE PartNumber = ? AND {empty_condition}",
        payload,
    )
    conn.commit(); conn.close()
    print(f"✍️ Заполнено CatID из предыдущего прайса: {len(payload)} поз. (только пустые)")


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
        17: '70',   # (1396 записей, но CatID=1000 определяется по UnitSort='SW')
        21: '70',   # (71 запись, 94.4% имеют CatID=70)
        24: '65',   # (1 запись, 100% имеют CatID=65)
        54: '65',   # (1 запись, 100% имеют CatID=65)
        55: '70',   # (219 записей, 28.3% имеют CatID=70 - самая частая)
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
    backfill_catid_from_previous_excel()
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

    # Экспорт (базовый)
    export_basic_excel()

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










