"""
CatID_lib - Библиотека для работы с категориями CatID в прайс-листах Volvo

Библиотека предоставляет функции для:
- Заполнения категорий CatID по различным правилам (UnitSort, ProductGroup, FunctionGroup)
- Импорта/экспорта матриц категорий
- Анализа качества заполнения категорий
- Полного пайплайна выравнивания CatID

Пример использования:
    from CatID_lib import align_catid_pipeline
    
    align_catid_pipeline(
        previous_excel_path="previous.xlsx",
        fgrp_matrix_excel_path="matrix.xlsx"
    )
"""

import os
import sqlite3
import pandas as pd
from datetime import datetime
from tkinter import filedialog

# ============================================================================
# Константы базы данных
# ============================================================================

DATABASE_PATH = "volvo_prices.db"
PRICELIST_TABLE_NAME = "volvo_price_list"
FOLDER_PATH = "export"
FGRP_TABLE_NAME = "fgrp_category"

# ============================================================================
# Конфигурация правил заполнения CatID
# ============================================================================

# Точные маппинги ProductGroup -> CatID (высокий приоритет)
PG_MAP_HARD = {
    15: '20',  # ACCESSORIES
    25: '80',  # MERCHANDISE
    # сюда добавляем PG, которые статистически «точные»
}

# Сомнительные маппинги ProductGroup -> CatID (низкий приоритет, включаются флагом)
PG_MAP_SOFT = {
    # сюда складываем PG, которые «сомнительные» и включаются флагом
    # 11: '70',
    # 14: '40',
    13: '60', # CHEMICALS
    16: '10', # TYRES
    18: '65', # SPECIAL TOOLS
}

USE_PG_SOFT = True  # Флаг включения мягких правил

# Защищённые категории (не перезаписываются)
FROZEN_CATIDS = {'1000', '4'}

# ============================================================================
# Утилиты
# ============================================================================

def ensure_folder_exists():
    """Создаёт папку для экспорта, если её нет."""
    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)


def norm_part(x):
    """Нормализует артикул (PartNumber).
    
    Args:
        x: Значение артикула (может быть строкой, числом, None)
    
    Returns:
        Нормализованный артикул или None
    """
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


def get_empty_catid_condition():
    """Возвращает SQL условие для проверки пустого CatID.
    
    Используется единообразно во всех функциях заполнения категорий.
    
    Returns:
        SQL условие в виде строки
    """
    return "(CatID IS NULL OR TRIM(COALESCE(CatID, '')) = '')"


# ============================================================================
# Статистика и мониторинг
# ============================================================================

def get_catid_statistics():
    """Возвращает статистику заполнения CatID для отладки и контроля.
    
    Returns:
        Словарь со статистикой: {'total', 'filled', 'empty', 'filled_percent'}
    """
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
    """Выводит статистику заполнения CatID на текущий момент.
    
    Args:
        stage_name: Название этапа для отображения в логе
    """
    stats = get_catid_statistics()
    prefix = f"[{stage_name}] " if stage_name else ""
    print(f"{prefix}📊 CatID статистика: заполнено {stats['filled']}/{stats['total']} ({stats['filled_percent']}%), пустых: {stats['empty']}")


# ============================================================================
# Основные функции заполнения CatID
# ============================================================================

def ensure_catid_soft():
    """Назначает CatID='1000' для программного обеспечения (UnitSort='SW').
    
    Заполняет ТОЛЬКО пустые CatID, не перезаписывает существующие.
    Приоритет: самый высокий (шаг 1 в пайплайне).
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
    conn.commit()
    conn.close()
    print(f"🏷️ Soft назначен {updated} позициям (CatID='1000'; только пустые).")


def ensure_catid_hardware():
    """Назначает CatID='4' для электроники (UnitSort='HW').
    
    Заполняет ТОЛЬКО пустые CatID, не перезаписывает существующие.
    Приоритет: самый высокий (шаг 1 в пайплайне).
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
    conn.commit()
    conn.close()
    print(f"🏷️ Hardware (Electronics) назначен {updated} позициям (CatID='4'; только пустые).")


def ensure_catid_by_pg(mapping):
    """Назначает CatID по значениям ProductGroup согласно словарю mapping.
    
    Args:
        mapping: Словарь {PG: int -> CatID: str}
                 Пример: {15: '20', 25: '80'}
    
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
    conn.commit()
    conn.close()
    print(f"✅ Итого обновлено по PG-мэппингу: {total} позиций.")


def apply_catid_overrides(force: bool = False):
    """Применяет ручные переопределения CatID.
    
    Args:
        force: Если True, перезаписывает все CatID (включая заполненные).
               Если False, заполняет только пустые CatID (по умолчанию).
    
    ВАЖНО: Не дублируй правила, которые уже обрабатываются в ensure_catid_by_pg()!
    Используй эту функцию только для точечных переопределений по PartNumber.
    """
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()
    
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
    conn.commit()
    conn.close()
    if total > 0:
        print(f"✅ Применены ручные переопределения CatID: {total} поз. ({'force' if force else 'только пустые'})")
    else:
        print(f"ℹ️ Ручные переопределения CatID: нет правил для применения.")


def backfill_catid_from_previous_excel(excel_path: str | None = None):
    """Дозаполняет CatID из предыдущего прайса (Excel), сопоставление по PartNumber.
    
    Имена колонок распознаются устойчиво: регистр не важен, пробелы/подчёркивания/точки/неразрывные пробелы игнорируются.
    
    Args:
        excel_path: Путь к Excel файлу с предыдущим прайсом.
                    Если None, пытается найти previous_prices.xlsx или открывает диалог выбора.
    
    Returns:
        Путь к использованному Excel (str), если файл успешно прочитан.
        None, если файл не выбран/не прочитан/не содержит нужных колонок.
    
    ВАЖНО:
      - Заполняет только ПУСТЫЕ CatID в базе.
      - НЕ перезаписывает уже заполненные (в т.ч. HW/SW).
      - НЕ переносит CatID='1000' (SOFT) из старого прайса.
    """
    def norm_colname(s: str) -> str:
        return "".join(ch for ch in str(s).lower() if ch not in " _.\u00A0.")
    
    # 1) Определяем файл: либо параметр, либо previous_prices.xlsx, либо диалог выбора
    if excel_path is None:
        default_candidate = os.path.abspath("previous_prices.xlsx")
        if os.path.isfile(default_candidate):
            excel_path = default_candidate
            print(f"📄 Используем previous_prices.xlsx как предыдущий прайс: {excel_path}")
        else:
            print("📄 Выберите предыдущий прайс (Excel) для подтягивания CatID…")
            excel_path = filedialog.askopenfilename(
                title="Выберите предыдущий прайс (Excel)",
                filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls")],
            )
            if not excel_path:
                print("⛔ Файл не выбран, пропускаем backfill CatID.")
                return None
    
    if not os.path.isfile(excel_path):
        print(f"⛔ Файл не найден: {excel_path}")
        return None
    
    excel_path = os.path.abspath(excel_path)
    print(f"📥 Читаем предыдущий прайс: {excel_path}")
    try:
        df_prev = pd.read_excel(excel_path)
    except Exception as e:
        print("⛔ Не удалось прочитать Excel:", e)
        return None
    
    if df_prev.empty:
        print("⛔ Файл предыдущего прайса пустой. Backfill CatID пропущен.")
        return excel_path
    
    # 2) Определяем колонки
    norm_map = {norm_colname(c): c for c in df_prev.columns}
    
    part_candidates = [
        "partnumber", "partno", "part_number", "part no", "номер", "артикул"
    ]
    cat_candidates = [
        "catid", "cat_id", "cat id", "category1c", "category", "категория"
    ]
    
    def pick(cands):
        for k in cands:
            nk = norm_colname(k)
            if nk in norm_map:
                return norm_map[nk]
        return None
    
    part_col = pick(part_candidates)
    cat_col  = pick(cat_candidates)
    
    if part_col is None or cat_col is None:
        print("⛔ Не удалось найти нужные колонки в предыдущем прайсе.")
        print("   Доступные колонки:", list(df_prev.columns))
        print("   Ищем PartNumber среди:", part_candidates)
        print("   Ищем CatID среди:", cat_candidates)
        return None
    
    # 3) Приводим к двум колонкам и нормализуем
    df_prev = df_prev[[part_col, cat_col]].copy()
    df_prev.columns = ["PartNumber", "CatID"]
    df_prev["PartNumber"] = df_prev["PartNumber"].map(norm_part)
    df_prev["CatID"] = df_prev["CatID"].astype(str).str.strip()
    df_prev = df_prev.dropna(subset=["PartNumber", "CatID"]).drop_duplicates("PartNumber")
    
    if df_prev.empty:
        print("⛔ В предыдущем прайсе нет валидных пар PartNumber+CatID.")
        return excel_path
    
    # 4) Выбираем, кого дозаполнять (только пустые CatID в БД)
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    empty_condition = get_empty_catid_condition()
    
    try:
        cur.execute(f"SELECT PartNumber FROM {PRICELIST_TABLE_NAME} WHERE {empty_condition}")
        missing_parts = {row[0] for row in cur.fetchall()}
    finally:
        pass
    
    if not missing_parts:
        print("ℹ️ В текущем прайсе нет позиций с пустым CatID. Backfill не требуется.")
        conn.close()
        return excel_path
    
    # Фильтруем только те артикулы, у которых CatID пустой в текущем прайсе
    upd = df_prev[df_prev["PartNumber"].isin(missing_parts)]
    
    if upd.empty:
        print("ℹ️ Нет пересечения артикулов между предыдущим прайсом и текущими пустыми CatID.")
        conn.close()
        return excel_path
    
    # Дополнительная защита: не заполняем пустые/плохие CatID из Excel
    upd = upd[upd["CatID"].notna() & (upd["CatID"].str.strip() != "")]
    if upd.empty:
        print("ℹ️ В предыдущем прайсе нет валидных CatID для дозаполнения.")
        conn.close()
        return excel_path
    
    # Исключаем категорию SOFT (CatID='1000') из предыдущего прайса:
    upd = upd[upd["CatID"].str.strip() != "1000"]
    if upd.empty:
        print("ℹ️ В предыдущем прайсе остались только категории SOFT (1000), которые не переносятся.")
        conn.close()
        return excel_path
    
    payload = list(map(tuple, upd[["CatID", "PartNumber"]].values))
    cur.executemany(
        f"""
        UPDATE {PRICELIST_TABLE_NAME}
        SET CatID = ?
        WHERE PartNumber = ?
          AND {empty_condition}
        """,
        payload,
    )
    conn.commit()
    conn.close()
    print(f"✍️ Заполнено CatID из предыдущего прайса: {len(payload)} поз. (только пустые)")
    
    return excel_path


# ============================================================================
# Работа с матрицей Fgrp -> CatID
# ============================================================================

def ensure_fgrp_table_exists():
    """Создаёт таблицу fgrp_category, если её нет."""
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {FGRP_TABLE_NAME} (
            Fgrp INTEGER PRIMARY KEY,
            CatID TEXT
        )
    """)
    conn.commit()
    conn.close()


def sync_new_fgrp_to_table():
    """Добавляет новые FunctionGroup из прайса в fgrp_category с CatID=NULL."""
    ensure_fgrp_table_exists()
    
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    
    # Находим новые FunctionGroup, которых нет в таблице fgrp_category
    cur.execute(f"""
        INSERT INTO {FGRP_TABLE_NAME} (Fgrp, CatID)
        SELECT DISTINCT FunctionGroup, NULL
        FROM {PRICELIST_TABLE_NAME}
        WHERE FunctionGroup IS NOT NULL
          AND FunctionGroup NOT IN (SELECT Fgrp FROM {FGRP_TABLE_NAME})
    """)
    added_count = cur.rowcount
    conn.commit()
    conn.close()
    
    if added_count > 0:
        print(f"🔄 Синхронизировано новых FunctionGroup: {added_count} (CatID=NULL)")
    else:
        print("ℹ️ Новых FunctionGroup не найдено.")


def import_fgrp_matrix_from_excel(path: str):
    """Импортирует матрицу Fgrp->CatID из Excel в таблицу fgrp_category.
    
    Args:
        path: Путь к Excel файлу с матрицей
    
    Ожидаемый формат Excel:
    - Лист должен содержать колонки: Fgrp (или FunctionGroup) и CatID
    - Пробует прочитать лист GROUP_INDEX, если его нет - читает первый лист
    """
    if not os.path.isfile(path):
        print(f"⚠️ Файл не найден: {path}")
        return
    
    ensure_fgrp_table_exists()
    
    try:
        # Пробуем прочитать лист GROUP_INDEX, если его нет - читаем первый лист
        try:
            df = pd.read_excel(path, sheet_name="GROUP_INDEX")
        except:
            df = pd.read_excel(path, sheet_name=0)
        
        if df.empty:
            print("⚠️ Excel файл пуст.")
            return
        
        # Нормализуем имена колонок
        def norm_colname(s: str) -> str:
            return "".join(ch for ch in str(s).lower() if ch not in " _.\u00A0")
        
        norm_map = {norm_colname(c): c for c in df.columns}
        
        # Ищем колонки Fgrp и CatID
        fgrp_col = None
        catid_col = None
        
        for cand in ["fgrp", "functiongroup"]:
            if cand in norm_map:
                fgrp_col = norm_map[cand]
                break
        
        for cand in ["catid", "cat_id", "category"]:
            if cand in norm_map:
                catid_col = norm_map[cand]
                break
        
        if fgrp_col is None or catid_col is None:
            print(f"⚠️ Не найдены нужные колонки в Excel. Доступные: {list(df.columns)}")
            print(f"   Ищем Fgrp среди: {list(norm_map.keys())}")
            return
        
        # Подготавливаем данные
        df_clean = df[[fgrp_col, catid_col]].copy()
        df_clean.columns = ["Fgrp", "CatID"]
        
        # Удаляем пустые строки
        df_clean = df_clean.dropna(subset=["Fgrp"])
        
        # Преобразуем CatID в строку (может быть числом или строкой)
        df_clean["CatID"] = df_clean["CatID"].astype(str).str.strip()
        df_clean["CatID"] = df_clean["CatID"].replace("nan", None)
        df_clean["CatID"] = df_clean["CatID"].replace("", None)
        
        # Преобразуем Fgrp в int
        df_clean["Fgrp"] = pd.to_numeric(df_clean["Fgrp"], errors="coerce")
        df_clean = df_clean.dropna(subset=["Fgrp"])
        df_clean["Fgrp"] = df_clean["Fgrp"].astype(int)
        
        # Обновляем таблицу: INSERT OR REPLACE для существующих Fgrp, INSERT для новых
        conn = sqlite3.connect(DATABASE_PATH)
        cur = conn.cursor()
        
        updated_count = 0
        inserted_count = 0
        
        for _, row in df_clean.iterrows():
            fgrp = int(row["Fgrp"])
            catid = row["CatID"] if pd.notna(row["CatID"]) else None
            
            # Проверяем, существует ли запись
            cur.execute(f"SELECT CatID FROM {FGRP_TABLE_NAME} WHERE Fgrp = ?", (fgrp,))
            existing = cur.fetchone()
            
            if existing:
                # Обновляем существующую запись
                cur.execute(f"""
                    UPDATE {FGRP_TABLE_NAME}
                    SET CatID = ?
                    WHERE Fgrp = ?
                """, (catid, fgrp))
                updated_count += 1
            else:
                # Вставляем новую запись
                cur.execute(f"""
                    INSERT INTO {FGRP_TABLE_NAME} (Fgrp, CatID)
                    VALUES (?, ?)
                """, (fgrp, catid))
                inserted_count += 1
        
        conn.commit()
        conn.close()
        
        print(f"📥 Импортировано матрицы Fgrp->CatID: обновлено {updated_count}, добавлено {inserted_count}")
        
    except Exception as e:
        print(f"❌ Ошибка при импорте матрицы из Excel: {e}")


def ensure_catid_by_fgrp(mode: str = "fill"):
    """Заполняет CatID по матрице FunctionGroup -> CatID из таблицы fgrp_category.
    
    Args:
        mode: "fill" - заполняет только пустые CatID (по умолчанию)
              "force" - перезаписывает все CatID (включая заполненные)
    """
    ensure_fgrp_table_exists()
    
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    
    if mode == "fill":
        empty_condition = get_empty_catid_condition()
        cur.execute(f"""
            UPDATE {PRICELIST_TABLE_NAME}
            SET CatID = (
                SELECT CatID FROM {FGRP_TABLE_NAME} f
                WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
            )
            WHERE FunctionGroup IS NOT NULL
              AND EXISTS (
                  SELECT 1 FROM {FGRP_TABLE_NAME} f
                  WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
                    AND f.CatID IS NOT NULL
              )
              AND {empty_condition}
        """)
    else:  # force
        cur.execute(f"""
            UPDATE {PRICELIST_TABLE_NAME}
            SET CatID = (
                SELECT CatID FROM {FGRP_TABLE_NAME} f
                WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
            )
            WHERE FunctionGroup IS NOT NULL
              AND EXISTS (
                  SELECT 1 FROM {FGRP_TABLE_NAME} f
                  WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
                    AND f.CatID IS NOT NULL
              )
        """)
    
    updated_count = cur.rowcount
    conn.commit()
    conn.close()
    
    if updated_count > 0:
        print(f"🏷️ Заполнено CatID по Fgrp матрице: {updated_count} поз. (mode={mode})")
    else:
        print(f"ℹ️ CatID по Fgrp матрице: нет позиций для обновления (mode={mode})")


# ============================================================================
# Отчёты и анализ
# ============================================================================

def export_catid_quality_report(previous_excel_path: str | None = None):
    """Сохраняет отчёт качества CatID (coverage, top-empty Fgrp, changes, pg effectiveness).
    
    Args:
        previous_excel_path: Путь к предыдущему прайсу для анализа изменений (опционально)
    
    Отчёт включает:
    - Coverage: статистика заполнения CatID
    - Top-empty Fgrp: FunctionGroup с наибольшим количеством пустых CatID
    - Changes: изменения CatID по сравнению с предыдущим прайсом (если передан)
    - PG effectiveness: эффективность правил по ProductGroup
    - Fgrp matrix status: статус матрицы Fgrp -> CatID
    """
    ensure_folder_exists()
    report_path = os.path.join(FOLDER_PATH, f"catid_quality_report_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx")
    
    conn = sqlite3.connect(DATABASE_PATH)
    
    # --- 1) Coverage: общая статистика заполнения
    empty_condition = get_empty_catid_condition()
    cur = conn.cursor()
    
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME}")
    total = cur.fetchone()[0]
    
    cur.execute(f"SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME} WHERE {empty_condition}")
    empty_count = cur.fetchone()[0]
    
    filled_count = total - empty_count
    coverage_pct = (filled_count / total * 100) if total > 0 else 0
    
    df_coverage = pd.DataFrame({
        "Metric": ["Total", "Filled", "Empty", "Coverage %"],
        "Value": [total, filled_count, empty_count, round(coverage_pct, 2)]
    })
    
    # --- 2) Top-empty Fgrp: FunctionGroup с наибольшим количеством пустых CatID
    df_empty_fgrp = pd.read_sql(f"""
        SELECT 
            FunctionGroup AS Fgrp,
            COUNT(*) AS EmptyCount,
            COUNT(*) * 100.0 / (SELECT COUNT(*) FROM {PRICELIST_TABLE_NAME} WHERE FunctionGroup = p.FunctionGroup) AS EmptyPercent
        FROM {PRICELIST_TABLE_NAME} p
        WHERE {empty_condition}
          AND FunctionGroup IS NOT NULL
        GROUP BY FunctionGroup
        ORDER BY EmptyCount DESC
        LIMIT 50
    """, conn)
    
    # --- 3) Changes: изменения CatID по сравнению с предыдущим прайсом
    df_changes = pd.DataFrame()
    if previous_excel_path and os.path.isfile(previous_excel_path):
        try:
            # Читаем предыдущий прайс
            df_prev = pd.read_excel(previous_excel_path, sheet_name=0)
            
            # Нормализуем имена колонок
            def norm_colname(s: str) -> str:
                return "".join(ch for ch in str(s).lower() if ch not in " _.\u00A0")
            
            norm_map = {norm_colname(c): c for c in df_prev.columns}
            
            part_col = None
            catid_col = None
            
            for cand in ["partnumber", "partno", "part_number"]:
                if cand in norm_map:
                    part_col = norm_map[cand]
                    break
            
            for cand in ["catid", "cat_id", "category"]:
                if cand in norm_map:
                    catid_col = norm_map[cand]
                    break
            
            if part_col and catid_col:
                df_prev_clean = df_prev[[part_col, catid_col]].copy()
                df_prev_clean.columns = ["PartNumber", "CatID_prev"]
                df_prev_clean["PartNumber"] = df_prev_clean["PartNumber"].astype(str).str.strip()
                df_prev_clean["CatID_prev"] = df_prev_clean["CatID_prev"].astype(str).str.strip()
                
                # Текущий прайс
                df_curr = pd.read_sql(f"""
                    SELECT PartNumber, CatID AS CatID_curr
                    FROM {PRICELIST_TABLE_NAME}
                """, conn)
                df_curr["PartNumber"] = df_curr["PartNumber"].astype(str).str.strip()
                
                # Объединяем и находим изменения
                df_merged = df_curr.merge(df_prev_clean, on="PartNumber", how="inner")
                df_merged = df_merged[
                    (df_merged["CatID_prev"].notna()) & 
                    (df_merged["CatID_curr"].notna()) &
                    (df_merged["CatID_prev"] != df_merged["CatID_curr"])
                ]
                
                if not df_merged.empty:
                    df_changes = df_merged[["PartNumber", "CatID_prev", "CatID_curr"]].copy()
                    df_changes.columns = ["PartNumber", "Previous CatID", "Current CatID"]
        except Exception as e:
            print(f"⚠️ Ошибка при анализе изменений: {e}")
    
    # --- 4) PG effectiveness: эффективность правил по ProductGroup
    df_pg_effectiveness = pd.read_sql(f"""
        SELECT 
            ProductGroup AS PG,
            COUNT(*) AS Total,
            SUM(CASE WHEN {empty_condition} THEN 1 ELSE 0 END) AS Empty,
            SUM(CASE WHEN NOT ({empty_condition}) THEN 1 ELSE 0 END) AS Filled,
            ROUND(SUM(CASE WHEN NOT ({empty_condition}) THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2) AS FilledPercent
        FROM {PRICELIST_TABLE_NAME}
        WHERE ProductGroup IS NOT NULL
        GROUP BY ProductGroup
        ORDER BY Empty DESC, Total DESC
    """, conn)
    
    # --- 5) Fgrp matrix status: статус матрицы Fgrp -> CatID
    ensure_fgrp_table_exists()
    df_fgrp_matrix = pd.read_sql(f"""
        SELECT 
            f.Fgrp,
            f.CatID AS MatrixCatID,
            COUNT(DISTINCT p.PartNumber) AS UsageCount,
            COUNT(DISTINCT CASE WHEN NOT ({empty_condition}) THEN p.PartNumber END) AS FilledCount
        FROM {FGRP_TABLE_NAME} f
        LEFT JOIN {PRICELIST_TABLE_NAME} p ON p.FunctionGroup = f.Fgrp
        GROUP BY f.Fgrp, f.CatID
        ORDER BY UsageCount DESC
    """, conn)
    
    conn.close()
    
    # --- Запись в Excel
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        df_coverage.to_excel(writer, sheet_name="Coverage", index=False)
        
        if not df_empty_fgrp.empty:
            df_empty_fgrp.to_excel(writer, sheet_name="Top Empty Fgrp", index=False)
        
        if not df_changes.empty:
            df_changes.to_excel(writer, sheet_name="CatID Changes", index=False)
        
        if not df_pg_effectiveness.empty:
            df_pg_effectiveness.to_excel(writer, sheet_name="PG Effectiveness", index=False)
        
        if not df_fgrp_matrix.empty:
            df_fgrp_matrix.to_excel(writer, sheet_name="Fgrp Matrix Status", index=False)
    
    print(f"📊 Отчёт качества CatID сохранён: {report_path}")
    print(f"   Coverage: {filled_count}/{total} ({coverage_pct:.2f}%)")


# ============================================================================
# Главный пайплайн выравнивания CatID
# ============================================================================

def align_catid_pipeline(
    previous_excel_path: str | None = None,
    fgrp_matrix_excel_path: str | None = None,
    pg_soft_enabled: bool = True,
    overrides_force: bool = False,
):
    """Единая точка входа для выравнивания CatID по текущему контракту.
    
    Выполняет полный пайплайн заполнения категорий CatID в следующем порядке:
    1. UnitSort (freeze): SW -> 1000, HW -> 4
    2. PG-high: точные маппинги ProductGroup -> CatID
    3. Fgrp matrix: заполнение по матрице FunctionGroup -> CatID
    4. PG-low: сомнительные маппинги ProductGroup -> CatID (если включены)
    5. Overrides: ручные переопределения по PartNumber
    6. Backfill: дозаполнение из предыдущего прайса
    7. Quality report: генерация отчёта качества
    
    Args:
        previous_excel_path: Путь к предыдущему прайсу для backfill (опционально)
        fgrp_matrix_excel_path: Путь к Excel с матрицей Fgrp->CatID (опционально)
        pg_soft_enabled: Включить ли мягкие правила PG (по умолчанию True)
        overrides_force: Принудительно применять overrides ко всем CatID (по умолчанию False)
    
    Пример:
        align_catid_pipeline(
            previous_excel_path="previous.xlsx",
            fgrp_matrix_excel_path="matrix.xlsx"
        )
    """
    print_catid_statistics("До CatID")

    # 1) UnitSort (freeze)
    ensure_catid_soft()  # SW -> 1000
    ensure_catid_hardware()  # HW -> 4
    print_catid_statistics("После шага 1 (SW/HW freeze)")

    # 2) PG-high (точные)
    if PG_MAP_HARD:
        ensure_catid_by_pg(PG_MAP_HARD)
        print_catid_statistics("После шага 2 (PG-high)")

    # 3) Fgrp -> CatID (матрица)
    # a) синхронизируем новые Fgrp в таблицу
    # b) при необходимости импортируем правки матрицы из Excel
    sync_new_fgrp_to_table()
    if fgrp_matrix_excel_path:
        import_fgrp_matrix_from_excel(fgrp_matrix_excel_path)
    ensure_catid_by_fgrp(mode="fill")  # fill-only
    print_catid_statistics("После шага 3 (Fgrp matrix)")

    # 4) PG-low (сомнительные)
    if pg_soft_enabled and PG_MAP_SOFT:
        ensure_catid_by_pg(PG_MAP_SOFT)
        print_catid_statistics("После шага 4 (PG-low)")

    # 5) Overrides (PartNumber)
    apply_catid_overrides(force=overrides_force)
    print_catid_statistics("После шага 5 (Overrides)")

    # 6) Backfill (previous)
    if previous_excel_path:
        backfill_catid_from_previous_excel(previous_excel_path)
        print_catid_statistics("После шага 6 (Backfill)")

    # 7) Quality report
    export_catid_quality_report(previous_excel_path=previous_excel_path)

