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

    
0.Administration - General information


1.Standard components, service materials and consumables


2.Engine with mountings and equipment


3.Electrical system


4.Power transmission


5.Brakes


6.Suspension and steering


7.Springs and wheels


8.Body and interior


9.Other - Special vehicles, e.g., ambulances, police vehicles. Components that deviatefrom standard 


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

# Правила ProductGroup -> CatID (единый источник всех правил)
# Все правила применяются на шаге 2 пайплайна
PG_MAP = {
    #11: '70',   # PARTS - самая частая категория
    14: '40',   # EXCHANGE - самая частая категория
    15: '20',   # ACCESSORIES
    25: '80',   # MERCHANDISE
    13: '60',   # CHEMICALS
    16: '10',   # TYRES
    18: '65',   # SPECIAL TOOLS
    21: '70',   # (71 запись, 94.4% имеют CatID=70)
    56: '10'    # (90 записей, 100% имеют CatID=10)
}

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
    applied_rules = []  # Для трассировки: какие правила были применены
    for pg, cat in mapping.items():
        cur.execute(f"""
            UPDATE {PRICELIST_TABLE_NAME}
            SET CatID = ?
            WHERE ProductGroup = ?
              AND {empty_condition}
        """, (str(cat), int(pg)))
        cnt = cur.rowcount
        total += cnt
        if cnt > 0:
            applied_rules.append((pg, cat, cnt))
        print(f"🏷️ PG={pg} → CatID='{cat}': {cnt} поз. (только пустые)")
    conn.commit()
    conn.close()
    print(f"✅ Итого обновлено по PG-мэппингу: {total} позиций.")
    return applied_rules  # Возвращаем для трассировки


def apply_pg_rules():
    """Применяет все правила ProductGroup из CatID_lib.
    
    Использует единый словарь PG_MAP со всеми правилами.
    Все правила применяются только к пустым CatID.
    Эта функция используется в пайплайне вместо прямых вызовов ensure_catid_by_pg().
    
    Returns:
        Список применённых правил [(PG, CatID, count), ...] для трассировки
    """
    if PG_MAP:
        return ensure_catid_by_pg(PG_MAP)
    return []


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


def trace_pg_catid_source(pg: int, catid: str, export_to_excel: bool = True):
    """Трассирует источник заполнения CatID для заданного ProductGroup.
    
    Детально анализирует, откуда могли взяться позиции с указанным PG и CatID:
    - Проверяет активные правила в PG_MAP
    - Анализирует распределение по Fgrp и матрицу Fgrp
    - Показывает статистику по UnitSort
    - Выводит список позиций для ручной проверки
    
    Args:
        pg: ProductGroup для проверки
        catid: CatID для проверки
        export_to_excel: Если True, экспортирует детальный отчёт в Excel
    
    Пример:
        trace_pg_catid_source(11, '70')  # Трассировка PG=11 → CatID=70
    """
    ensure_folder_exists()
    conn = sqlite3.connect(DATABASE_PATH)
    empty_condition = get_empty_catid_condition()
    
    # Находим все позиции с указанным PG и CatID
    df_positions = pd.read_sql(f"""
        SELECT 
            PartNumber,
            FunctionGroup AS Fgrp,
            UnitSort,
            CatID,
            Description
        FROM {PRICELIST_TABLE_NAME}
        WHERE ProductGroup = ?
          AND CatID = ?
    """, conn, params=(pg, catid))
    
    if df_positions.empty:
        print(f"ℹ️ Нет позиций с PG={pg} и CatID={catid}")
        conn.close()
        return None
    
    print(f"\n🔍 ТРАССИРОВКА: PG={pg} → CatID={catid}")
    print(f"   Найдено позиций: {len(df_positions)}")
    print("="*70)
    
    # 1. Проверяем правило в PG_MAP
    expected_catid = PG_MAP.get(pg)
    print(f"\n1️⃣ Проверка правила в PG_MAP:")
    if expected_catid == catid:
        print(f"   ✅ Правило ЕСТЬ в PG_MAP: PG={pg} → CatID={catid}")
        print(f"   ⚠️  ВНИМАНИЕ: Правило закомментировано в коде!")
        print(f"   💡 Это значит, что правило было применено РАНЬШЕ (до комментирования)")
        print(f"      или позиции были заполнены из другого источника.")
    elif expected_catid is not None:
        print(f"   ❌ Правило ЕСТЬ, но CatID НЕ СОВПАДАЕТ:")
        print(f"      PG={pg} → CatID={expected_catid} (ожидалось)")
        print(f"      PG={pg} → CatID={catid} (фактически)")
    else:
        print(f"   ❌ Правила НЕТ в PG_MAP для PG={pg}")
    
    # 2. Анализ по Fgrp
    print(f"\n2️⃣ Анализ по FunctionGroup:")
    df_fgrp_stats = df_positions.groupby("Fgrp").agg({
        "PartNumber": "count",
        "UnitSort": lambda x: ", ".join(x.unique())
    }).reset_index()
    df_fgrp_stats.columns = ["Fgrp", "Count", "UnitSorts"]
    
    # Загружаем матрицу Fgrp
    ensure_fgrp_table_exists()
    df_matrix = pd.read_sql(f"SELECT Fgrp, CatID AS MatrixCatID FROM {FGRP_TABLE_NAME}", conn)
    matrix_dict = {}
    for _, row in df_matrix.iterrows():
        fgrp_val = int(row["Fgrp"])
        catid_val = str(row["MatrixCatID"]).strip() if pd.notna(row["MatrixCatID"]) else None
        matrix_dict[fgrp_val] = catid_val if catid_val and catid_val != "" else None
    
    # Добавляем информацию о матрице
    df_fgrp_stats["MatrixCatID"] = df_fgrp_stats["Fgrp"].map(lambda x: matrix_dict.get(int(x), ""))
    df_fgrp_stats["MatchesMatrix"] = df_fgrp_stats.apply(
        lambda row: str(row["MatrixCatID"]) == catid if row["MatrixCatID"] else False, axis=1
    )
    
    matching_fgrp = df_fgrp_stats[df_fgrp_stats["MatchesMatrix"]]["Fgrp"].tolist()
    non_matching_fgrp = df_fgrp_stats[~df_fgrp_stats["MatchesMatrix"]]["Fgrp"].tolist()
    
    print(f"   Всего разных Fgrp: {len(df_fgrp_stats)}")
    if matching_fgrp:
        print(f"   ✅ Соответствуют матрице: {len(matching_fgrp)} Fgrp")
        print(f"      Fgrp: {matching_fgrp[:10]}{'...' if len(matching_fgrp) > 10 else ''}")
    if non_matching_fgrp:
        print(f"   ❌ НЕ соответствуют матрице: {len(non_matching_fgrp)} Fgrp")
        for fgrp in non_matching_fgrp[:5]:
            matrix_val = matrix_dict.get(int(fgrp), "не задана")
            print(f"      Fgrp={fgrp}: матрица → {matrix_val}, факт → {catid}")
    
    # 3. Проверка UnitSort
    print(f"\n3️⃣ Проверка UnitSort:")
    sw_count = len(df_positions[df_positions["UnitSort"] == "SW"])
    hw_count = len(df_positions[df_positions["UnitSort"] == "HW"])
    other_count = len(df_positions[~df_positions["UnitSort"].isin(["SW", "HW"])])
    
    if catid == "1000" and sw_count > 0:
        print(f"   ✅ UnitSort='SW' → CatID=1000: {sw_count} позиций")
    elif catid == "4" and hw_count > 0:
        print(f"   ✅ UnitSort='HW' → CatID=4: {hw_count} позиций")
    else:
        print(f"   ❌ UnitSort НЕ является причиной:")
        print(f"      SW={sw_count}, HW={hw_count}, Other={other_count}")
    
    # 4. Вывод наиболее вероятных источников
    print(f"\n4️⃣ Наиболее вероятные источники:")
    sources = []
    
    if expected_catid == catid:
        sources.append("✅ Правило PG_MAP было применено ранее (до комментирования)")
    if matching_fgrp:
        sources.append(f"✅ Матрица Fgrp ({len(matching_fgrp)} Fgrp соответствуют)")
    sources.append("✅ Backfill из предыдущего прайса (вероятнее всего)")
    sources.append("❓ Ручное заполнение или другой источник")
    
    for i, source in enumerate(sources, 1):
        print(f"   {i}. {source}")
    
    # 5. Экспорт детального отчёта
    if export_to_excel:
        report_path = os.path.join(FOLDER_PATH, f"trace_pg{pg}_catid{catid}_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx")
        
        with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
            # Детали позиций
            df_positions.to_excel(writer, sheet_name="Positions", index=False)
            
            # Статистика по Fgrp
            df_fgrp_stats.to_excel(writer, sheet_name="Fgrp_Stats", index=False)
            
            # Сводка
            df_summary = pd.DataFrame({
                "Параметр": [
                    "ProductGroup",
                    "CatID",
                    "Всего позиций",
                    "Правило в PG_MAP",
                    "Соответствует матрице Fgrp",
                    "Не соответствует матрице Fgrp",
                    "UnitSort=SW",
                    "UnitSort=HW",
                    "Другие UnitSort"
                ],
                "Значение": [
                    pg,
                    catid,
                    len(df_positions),
                    "Да (закомментировано)" if expected_catid == catid else "Нет",
                    len(matching_fgrp),
                    len(non_matching_fgrp),
                    sw_count,
                    hw_count,
                    other_count
                ]
            })
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
        
        print(f"\n📄 Детальный отчёт сохранён: {report_path}")
    
    conn.close()
    return df_positions


def diagnose_catid_source(pg: int, catid: str):
    """Диагностирует источник заполнения CatID для заданного ProductGroup.
    
    Показывает, откуда могли взяться позиции с указанным PG и CatID:
    - Из правил PG_MAP (если правило активно)
    - Из backfill (если есть предыдущий прайс)
    - Из матрицы Fgrp
    - Из overrides
    - Неизвестный источник
    
    Args:
        pg: ProductGroup для проверки
        catid: CatID для проверки
    
    Пример:
        diagnose_catid_source(11, '70')  # Проверит откуда взялись PG=11 → CatID=70
    
    Примечание: Для детальной трассировки используйте trace_pg_catid_source()
    """
    # Используем упрощённую версию для быстрой проверки
    trace_pg_catid_source(pg, catid, export_to_excel=False)
    conn = sqlite3.connect(DATABASE_PATH)
    empty_condition = get_empty_catid_condition()
    
    # Находим все позиции с указанным PG и CatID
    df_positions = pd.read_sql(f"""
        SELECT 
            PartNumber,
            FunctionGroup AS Fgrp,
            UnitSort,
            CatID
        FROM {PRICELIST_TABLE_NAME}
        WHERE ProductGroup = ?
          AND CatID = ?
    """, conn, params=(pg, catid))
    
    if df_positions.empty:
        print(f"ℹ️ Нет позиций с PG={pg} и CatID={catid}")
        conn.close()
        return
    
    print(f"\n🔍 Диагностика: PG={pg} → CatID={catid}")
    print(f"   Найдено позиций: {len(df_positions)}")
    
    # 1. Проверяем правило в PG_MAP
    expected_catid = PG_MAP.get(pg)
    if expected_catid == catid:
        print(f"   ✅ Правило есть в PG_MAP: PG={pg} → CatID={catid}")
        print(f"      ⚠️ Но правило закомментировано! Значит было применено ранее или из другого источника.")
    else:
        print(f"   ❌ Правила НЕТ в PG_MAP (ожидалось: {expected_catid if expected_catid else 'нет правила'})")
    
    # 2. Проверяем матрицу Fgrp
    ensure_fgrp_table_exists()
    unique_fgrps = df_positions["Fgrp"].dropna().unique()
    
    if len(unique_fgrps) > 0:
        fgrp_list = [int(f) for f in unique_fgrps if pd.notna(f)]
        if fgrp_list:
            placeholders = ','.join(['?'] * len(fgrp_list))
            df_matrix_check = pd.read_sql(f"""
                SELECT Fgrp, CatID AS MatrixCatID
                FROM {FGRP_TABLE_NAME}
                WHERE Fgrp IN ({placeholders})
            """, conn, params=tuple(fgrp_list))
            
            matching_fgrp = []
            for fgrp in fgrp_list:
                matrix_rows = df_matrix_check[df_matrix_check["Fgrp"] == fgrp]
                if not matrix_rows.empty:
                    matrix_catid_val = matrix_rows["MatrixCatID"].iloc[0]
                    matrix_catid = str(matrix_catid_val).strip() if pd.notna(matrix_catid_val) else None
                    if matrix_catid == catid:
                        matching_fgrp.append(fgrp)
            
            if matching_fgrp:
                print(f"   ✅ Матрица Fgrp: {len(matching_fgrp)} Fgrp имеют CatID={catid} в матрице")
                print(f"      Fgrp: {matching_fgrp[:10]}{'...' if len(matching_fgrp) > 10 else ''}")
            else:
                print(f"   ❌ Матрица Fgrp: CatID={catid} НЕ соответствует матрице для этих Fgrp")
    
    # 3. Проверяем UnitSort
    sw_count = len(df_positions[df_positions["UnitSort"] == "SW"])
    hw_count = len(df_positions[df_positions["UnitSort"] == "HW"])
    if catid == "1000" and sw_count > 0:
        print(f"   ✅ UnitSort: {sw_count} позиций с UnitSort='SW' → CatID=1000")
    elif catid == "4" and hw_count > 0:
        print(f"   ✅ UnitSort: {hw_count} позиций с UnitSort='HW' → CatID=4")
    else:
        print(f"   ❌ UnitSort: не является причиной (SW={sw_count}, HW={hw_count})")
    
    # 4. Вывод: наиболее вероятный источник
    print(f"\n   💡 Наиболее вероятный источник:")
    if expected_catid == catid:
        print(f"      - Правило PG_MAP было применено ранее (до того как его закомментировали)")
    if matching_fgrp:
        print(f"      - Матрица Fgrp (для {len(matching_fgrp)} Fgrp)")
    print(f"      - Backfill из предыдущего прайса (вероятнее всего, если правило закомментировано)")
    print(f"      - Ручное заполнение или другой источник")
    
    conn.close()


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


def fill_fgrp_matrix_from_current_prices(only_empty: bool = True):
    """Заполняет матрицу Fgrp -> CatID на основе текущих значений CatID в прайсе.
    
    Для каждого FunctionGroup выбирается самый частый CatID из текущего прайса.
    
    Args:
        only_empty: Если True (по умолчанию), заполняет пустые CatID и обновляет только изменённые значения
                   (сравнивает с текущим значением в матрице и обновляет только если отличается).
                   Если False, обновляет все CatID в матрице на основе текущего прайса без сравнения.
    """
    ensure_fgrp_table_exists()
    
    conn = sqlite3.connect(DATABASE_PATH)
    
    # Получаем статистику: для каждого Fgrp находим самый частый CatID
    empty_condition = get_empty_catid_condition()
    df_stats = pd.read_sql(f"""
        SELECT 
            FunctionGroup AS Fgrp,
            CatID,
            COUNT(*) AS Count
        FROM {PRICELIST_TABLE_NAME}
        WHERE FunctionGroup IS NOT NULL
          AND NOT ({empty_condition})
        GROUP BY FunctionGroup, CatID
        ORDER BY FunctionGroup, Count DESC
    """, conn)
    
    if df_stats.empty:
        print("ℹ️ В текущем прайсе нет заполненных CatID для анализа.")
        conn.close()
        return
    
    # Для каждого Fgrp выбираем CatID с максимальным Count
    df_top_catid = df_stats.groupby("Fgrp").first().reset_index()
    df_top_catid = df_top_catid[["Fgrp", "CatID"]]
    
    # Загружаем текущую матрицу для сравнения
    df_matrix = pd.read_sql(f"""
        SELECT Fgrp, CatID FROM {FGRP_TABLE_NAME}
    """, conn)
    matrix_dict = {}
    for _, row in df_matrix.iterrows():
        fgrp = int(row["Fgrp"])
        catid = str(row["CatID"]).strip() if pd.notna(row["CatID"]) else None
        matrix_dict[fgrp] = catid if catid and catid != "" else None
    
    # Обновляем матрицу
    cur = conn.cursor()
    updated_count = 0
    inserted_count = 0
    unchanged_count = 0
    
    for _, row in df_top_catid.iterrows():
        fgrp = int(row["Fgrp"])
        catid_new = str(row["CatID"]).strip() if pd.notna(row["CatID"]) else None
        
        if catid_new is None:
            continue
        
        # Проверяем, существует ли запись в матрице
        catid_existing = matrix_dict.get(fgrp)
        
        if catid_existing is not None:
            # Запись существует в матрице
            if only_empty:
                # Режим "fill": обновляем только если значение изменилось
                if catid_existing != catid_new:
                    cur.execute(f"""
                        UPDATE {FGRP_TABLE_NAME}
                        SET CatID = ?
                        WHERE Fgrp = ?
                    """, (catid_new, fgrp))
                    updated_count += 1
                else:
                    unchanged_count += 1
            else:
                # Режим "force": обновляем всегда
                cur.execute(f"""
                    UPDATE {FGRP_TABLE_NAME}
                    SET CatID = ?
                    WHERE Fgrp = ?
                """, (catid_new, fgrp))
                updated_count += 1
        else:
            # Записи нет в матрице - добавляем новую
            cur.execute(f"""
                INSERT INTO {FGRP_TABLE_NAME} (Fgrp, CatID)
                VALUES (?, ?)
            """, (fgrp, catid_new))
            inserted_count += 1
    
    conn.commit()
    conn.close()
    
    total = updated_count + inserted_count
    if total > 0:
        print(f"📊 Обновлена матрица Fgrp -> CatID из текущего прайса:")
        print(f"   - Обновлено (изменённые значения): {updated_count}")
        print(f"   - Добавлено (новые Fgrp): {inserted_count}")
        if only_empty and unchanged_count > 0:
            print(f"   - Без изменений (совпадают): {unchanged_count}")
        print(f"   Использованы самые частые CatID для каждого FunctionGroup")
    else:
        if only_empty:
            print(f"ℹ️ Матрица Fgrp -> CatID: нет изменений (все значения совпадают или пустые)")
        else:
            print(f"ℹ️ Матрица Fgrp -> CatID: нет данных для обновления")


def import_fgrp_matrix_from_excel(path: str, mode: str = "fill", initial_matrix_dict: dict[int, str | None] | None = None):
    """Импортирует матрицу Fgrp->CatID из Excel в таблицу fgrp_category.
    
    Args:
        path: Путь к Excel файлу с матрицей
        mode: Режим импорта:
              "fill" (по умолчанию) - обновляет только изменённые значения (сравнивает с начальной матрицей)
              "force" - обновляет все значения без сравнения
        initial_matrix_dict: Опциональный словарь {Fgrp: CatID} с начальным состоянием матрицы
                            для сравнения. Если передан, используется вместо чтения из Excel.
                            Если None, пытается прочитать из листа _Initial_Matrix в Excel.
    
    Returns:
        Список изменённых Fgrp (если mode="fill"), иначе None.
        Используется для применения изменений ко всем позициям группы.
    
    Ожидаемый формат Excel:
    - Лист должен содержать колонки: Fgrp (или FunctionGroup) и CatID
    - Пробует прочитать лист Fgrp_Matrix, если его нет - читает первый лист
    - Если есть лист _Initial_Matrix и initial_matrix_dict не передан, используется для сравнения в режиме "fill"
    """
    if not os.path.isfile(path):
        print(f"⚠️ Файл не найден: {path}")
        return
    
    ensure_fgrp_table_exists()
    
    try:
        # Пробуем прочитать лист Fgrp_Matrix, если его нет - читаем первый лист
        try:
            df = pd.read_excel(path, sheet_name="Fgrp_Matrix")
        except:
            try:
                df = pd.read_excel(path, sheet_name="GROUP_INDEX")
            except:
                df = pd.read_excel(path, sheet_name=0)
        
        if df.empty:
            print("⚠️ Excel файл пуст.")
            return
        
        # Загружаем начальную матрицу для сравнения (если есть)
        if mode == "fill":
            if initial_matrix_dict is not None:
                # Используем переданное начальное состояние матрицы
                initial_dict = initial_matrix_dict.copy()
                print("ℹ️ Используется переданное начальное состояние матрицы для сравнения")
            else:
                # Пытаемся прочитать из Excel
                try:
                    df_initial = pd.read_excel(path, sheet_name="_Initial_Matrix")
                    # Нормализуем начальную матрицу
                    df_initial["CatID"] = df_initial["CatID"].astype(str).str.strip()
                    df_initial["CatID"] = df_initial["CatID"].replace("nan", None)
                    df_initial["CatID"] = df_initial["CatID"].replace("", None)
                    df_initial["Fgrp"] = pd.to_numeric(df_initial["Fgrp"], errors="coerce")
                    df_initial = df_initial.dropna(subset=["Fgrp"])
                    df_initial["Fgrp"] = df_initial["Fgrp"].astype(int)
                    # Создаём словарь для быстрого поиска
                    initial_dict = {}
                    for _, row in df_initial.iterrows():
                        fgrp = int(row["Fgrp"])
                        catid = row["CatID"] if pd.notna(row["CatID"]) else None
                        initial_dict[fgrp] = catid if catid and catid != "" else None
                except:
                    # Начальной матрицы нет - будем сравнивать с текущей в базе
                    initial_dict = None
                    print("ℹ️ Начальная матрица не найдена, сравниваем с текущей в базе")
        else:
            initial_dict = None
        
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
        
        # Если нет начальной матрицы в Excel, загружаем текущую из базы для сравнения
        if mode == "fill" and initial_dict is None:
            conn_temp = sqlite3.connect(DATABASE_PATH)
            df_current = pd.read_sql(f"SELECT Fgrp, CatID FROM {FGRP_TABLE_NAME}", conn_temp)
            conn_temp.close()
            initial_dict = {}
            for _, row in df_current.iterrows():
                fgrp = int(row["Fgrp"])
                catid = str(row["CatID"]).strip() if pd.notna(row["CatID"]) else None
                initial_dict[fgrp] = catid if catid and catid != "" else None
        
        # Обновляем таблицу
        conn = sqlite3.connect(DATABASE_PATH)
        cur = conn.cursor()
        
        updated_count = 0
        inserted_count = 0
        unchanged_count = 0
        changed_fgrps_list = []  # Список изменённых Fgrp для применения ко всем позициям группы
        
        for _, row in df_clean.iterrows():
            fgrp = int(row["Fgrp"])
            catid_new = row["CatID"] if pd.notna(row["CatID"]) else None
            catid_new = catid_new if catid_new and catid_new != "" else None
            
            # Проверяем, существует ли запись
            cur.execute(f"SELECT CatID FROM {FGRP_TABLE_NAME} WHERE Fgrp = ?", (fgrp,))
            existing_db = cur.fetchone()
            
            if mode == "fill" and initial_dict is not None:
                # Режим "fill": сравниваем с начальной матрицей
                catid_initial = initial_dict.get(fgrp)
                
                if catid_initial != catid_new:
                    # Значение изменилось - обновляем и добавляем в список изменённых
                    if existing_db:
                        cur.execute(f"""
                            UPDATE {FGRP_TABLE_NAME}
                            SET CatID = ?
                            WHERE Fgrp = ?
                        """, (catid_new, fgrp))
                        updated_count += 1
                    else:
                        cur.execute(f"""
                            INSERT INTO {FGRP_TABLE_NAME} (Fgrp, CatID)
                            VALUES (?, ?)
                        """, (fgrp, catid_new))
                        inserted_count += 1
                    changed_fgrps_list.append(fgrp)  # Добавляем в список изменённых
                else:
                    # Значение не изменилось
                    unchanged_count += 1
            else:
                # Режим "force": обновляем всё без сравнения
                if existing_db:
                    cur.execute(f"""
                        UPDATE {FGRP_TABLE_NAME}
                        SET CatID = ?
                        WHERE Fgrp = ?
                    """, (catid_new, fgrp))
                    updated_count += 1
                else:
                    cur.execute(f"""
                        INSERT INTO {FGRP_TABLE_NAME} (Fgrp, CatID)
                        VALUES (?, ?)
                    """, (fgrp, catid_new))
                    inserted_count += 1
        
        conn.commit()
        conn.close()
        
        if mode == "fill":
            print(f"✅ Импортировано из Excel (mode=fill): обновлено {updated_count}, добавлено {inserted_count}, без изменений {unchanged_count}")
            if changed_fgrps_list:
                print(f"   Изменено значений в матрице: {len(changed_fgrps_list)} Fgrp")
                print(f"   Эти Fgrp будут применены ко ВСЕМ позициям группы при применении матрицы")
        else:
            print(f"✅ Импортировано из Excel (mode=force): обновлено {updated_count}, добавлено {inserted_count}")
            changed_fgrps_list = None
        
        return changed_fgrps_list if mode == "fill" else None
        
    except Exception as e:
        print(f"❌ Ошибка при импорте матрицы из Excel: {e}")
        return None


def ensure_catid_by_fgrp(mode: str = "fill", changed_fgrps: list[int] | None = None):
    """Заполняет CatID по матрице FunctionGroup -> CatID из таблицы fgrp_category.
    
    Args:
        mode: "fill" - заполняет только пустые CatID (по умолчанию)
              "force" - перезаписывает все CatID (включая заполненные)
        changed_fgrps: Список Fgrp, значения которых изменились в матрице.
                      Для этих Fgrp применяется force (ко всем позициям),
                      для остальных - fill (только пустые).
                      Используется при редактировании матрицы.
    """
    ensure_fgrp_table_exists()
    
    conn = sqlite3.connect(DATABASE_PATH)
    cur = conn.cursor()
    
    updated_count = 0
    
    try:
        if changed_fgrps and len(changed_fgrps) > 0:
            # Валидация: убеждаемся, что все значения - целые числа
            valid_fgrps = []
            for fgrp in changed_fgrps:
                try:
                    fgrp_int = int(fgrp)
                    if fgrp_int not in valid_fgrps:
                        valid_fgrps.append(fgrp_int)
                except (ValueError, TypeError):
                    continue
            
            if not valid_fgrps:
                print("⚠️ Нет валидных Fgrp для применения force. Применяем стандартный режим fill.")
                changed_fgrps = None
            else:
                changed_fgrps = valid_fgrps
            
        if changed_fgrps and len(changed_fgrps) > 0:
            # Для изменённых Fgrp применяем force (ко всем позициям группы)
            empty_condition = get_empty_catid_condition()
            placeholders = ','.join(['?'] * len(changed_fgrps))
            
            # Force для изменённых Fgrp
            cur.execute(f"""
                UPDATE {PRICELIST_TABLE_NAME}
                SET CatID = (
                    SELECT CatID FROM {FGRP_TABLE_NAME} f
                    WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
                )
                WHERE FunctionGroup IN ({placeholders})
                  AND EXISTS (
                      SELECT 1 FROM {FGRP_TABLE_NAME} f
                      WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
                        AND f.CatID IS NOT NULL
                  )
            """, tuple(changed_fgrps))
            updated_changed = cur.rowcount
            
            # Fill для остальных Fgrp (только пустые)
            cur.execute(f"""
                UPDATE {PRICELIST_TABLE_NAME}
                SET CatID = (
                    SELECT CatID FROM {FGRP_TABLE_NAME} f
                    WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
                )
                WHERE FunctionGroup IS NOT NULL
                  AND FunctionGroup NOT IN ({placeholders})
                  AND EXISTS (
                      SELECT 1 FROM {FGRP_TABLE_NAME} f
                      WHERE f.Fgrp = {PRICELIST_TABLE_NAME}.FunctionGroup
                        AND f.CatID IS NOT NULL
                  )
                  AND {empty_condition}
            """, tuple(changed_fgrps))
            updated_unchanged = cur.rowcount
            
            updated_count = updated_changed + updated_unchanged
            
            if updated_changed > 0 or updated_unchanged > 0:
                print(f"🏷️ Применена матрица Fgrp -> CatID:")
                print(f"   - Изменённые Fgrp (force, ко всем): {updated_changed} поз. ({len(changed_fgrps)} Fgrp)")
                print(f"   - Остальные Fgrp (fill, только пустые): {updated_unchanged} поз.")
    except Exception as e:
        print(f"❌ Ошибка при применении матрицы с changed_fgrps: {e}")
        import traceback
        traceback.print_exc()
        # Продолжаем со стандартным режимом fill
        changed_fgrps = None
    
    if not (changed_fgrps and len(changed_fgrps) > 0):
        # Если нет changed_fgrps или произошла ошибка, применяем стандартный режим
        if mode == "fill":
            # Стандартный режим fill - только пустые
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
            updated_count = cur.rowcount
        else:  # force
            # Стандартный режим force - все позиции
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
        if changed_fgrps:
            print(f"✅ Всего обновлено: {updated_count} поз.")
        else:
            print(f"🏷️ Заполнено CatID по Fgrp матрице: {updated_count} поз. (mode={mode})")
    else:
        print(f"ℹ️ CatID по Fgrp матрице: нет позиций для обновления (mode={mode})")


# ============================================================================
# Редактирование матрицы Fgrp -> CatID
# ============================================================================

def export_fgrp_matrix_to_excel(output_path: str | None = None):
    """Экспортирует матрицу Fgrp -> CatID в Excel для редактирования.
    
    Сохраняет начальную матрицу в отдельный лист для сравнения при импорте.
    
    Args:
        output_path: Путь для сохранения Excel файла. Если None, создаётся автоматически в папке export.
    
    Returns:
        Путь к созданному файлу
    """
    ensure_folder_exists()
    ensure_fgrp_table_exists()
    
    if output_path is None:
        output_path = os.path.join(FOLDER_PATH, f"fgrp_matrix_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx")
    
    conn = sqlite3.connect(DATABASE_PATH)
    
    # Загружаем матрицу из базы
    df_matrix = pd.read_sql(f"""
        SELECT 
            f.Fgrp,
            f.CatID,
            COUNT(DISTINCT p.PartNumber) AS UsageCount,
            COUNT(DISTINCT CASE WHEN p.CatID IS NOT NULL AND TRIM(COALESCE(p.CatID, '')) != '' THEN p.PartNumber END) AS FilledCount
        FROM {FGRP_TABLE_NAME} f
        LEFT JOIN {PRICELIST_TABLE_NAME} p ON p.FunctionGroup = f.Fgrp
        GROUP BY f.Fgrp, f.CatID
        ORDER BY f.Fgrp
    """, conn)
    
    # Сохраняем начальную матрицу для сравнения (только Fgrp и CatID)
    df_initial = df_matrix[["Fgrp", "CatID"]].copy()
    
    conn.close()
    
    # Сохраняем в Excel
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_matrix.to_excel(writer, sheet_name="Fgrp_Matrix", index=False)
        
        # Сохраняем начальную матрицу в скрытый лист для сравнения
        df_initial.to_excel(writer, sheet_name="_Initial_Matrix", index=False)
        
        # Добавляем инструкции на отдельном листе
        df_instructions = pd.DataFrame({
            "Инструкция": [
                "1. Отредактируйте колонку CatID в листе Fgrp_Matrix",
                "2. Сохраните файл",
                "3. Импортируйте обратно через: import_fgrp_matrix_from_excel('путь_к_файлу', mode='fill')",
                "4. При применении mode='fill':",
                "   - Изменённые Fgrp применятся ко ВСЕМ позициям группы (force)",
                "   - Неизменённые Fgrp применятся только к пустым CatID (fill)",
                "",
                "Примечания:",
                "- Fgrp: FunctionGroup (не изменяйте)",
                "- CatID: категория для данного Fgrp (можно редактировать)",
                "- UsageCount: количество позиций с этим Fgrp в прайсе",
                "- FilledCount: количество позиций с заполненным CatID",
                "",
                "Режимы импорта:",
                "- mode='fill': обновляет только изменённые значения в матрице.",
                "              При применении к прайсу: изменённые Fgrp → ко всем позициям,",
                "              остальные → только к пустым CatID",
                "- mode='force': обновляет все значения в матрице без сравнения"
            ]
        })
        df_instructions.to_excel(writer, sheet_name="Инструкции", index=False)
        
        # Скрываем лист с начальной матрицей
        workbook = writer.book
        worksheet_initial = writer.sheets["_Initial_Matrix"]
        worksheet_initial.hide()
    
    print(f"📤 Матрица Fgrp -> CatID экспортирована: {output_path}")
    print(f"   Всего записей: {len(df_matrix)}")
    print(f"   Заполненных CatID: {df_matrix['CatID'].notna().sum()}")
    
    return output_path


def edit_fgrp_matrix_interactive():
    """Интерактивное редактирование матрицы Fgrp -> CatID через Excel.
    
    Функция:
    1. Экспортирует матрицу в Excel
    2. Открывает файл для редактирования
    3. Ждёт закрытия файла
    4. Импортирует изменения обратно
    5. Применяет изменения к прайсу (опционально)
    """
    import subprocess
    import time
    
    # Экспортируем матрицу
    excel_path = export_fgrp_matrix_to_excel()
    
    print(f"\n📝 Открыт файл для редактирования: {excel_path}")
    print("   Отредактируйте колонку CatID и сохраните файл.")
    print("   После сохранения и закрытия файла изменения будут импортированы.\n")
    
    # Открываем файл в Excel
    try:
        os.startfile(excel_path)
    except:
        print(f"⚠️ Не удалось открыть файл автоматически. Откройте вручную: {excel_path}")
    
    # Ждём подтверждения от пользователя
    input("Нажмите Enter после того, как сохраните и закроете Excel файл...")
    
    # Спрашиваем режим импорта
    print("\n❓ Режим импорта матрицы из Excel?")
    print("   'fill' - обновить только изменённые значения (сравнение с начальной матрицей)")
    print("   'force' - обновить все значения без сравнения")
    
    import_mode = input("Ваш выбор (fill/force, по умолчанию fill): ").strip().lower()
    if import_mode not in ["fill", "force"]:
        import_mode = "fill"
    
    # Импортируем изменения
    print(f"\n📥 Импортируем изменения из Excel (mode={import_mode})...")
    changed_fgrps = import_fgrp_matrix_from_excel(excel_path, mode=import_mode)
    
    # Спрашиваем, применять ли изменения к прайсу
    print("\n❓ Применить изменения к прайсу?")
    if import_mode == "fill" and changed_fgrps:
        print(f"   'fill' (по умолчанию) - изменённые Fgrp применятся ко ВСЕМ позициям группы,")
        print(f"                        остальные - только к пустым CatID")
        print(f"   'force' - перезаписать все CatID для всех Fgrp")
        print(f"   'skip' - пропустить применение")
    else:
        print("   'fill' - заполнить только пустые CatID")
        print("   'force' - перезаписать все CatID")
        print("   'skip' - пропустить применение")
    
    choice = input("Ваш выбор (fill/force/skip, по умолчанию fill): ").strip().lower()
    
    if choice == 'force':
        ensure_catid_by_fgrp(mode="force")
        print("✅ Изменения применены к прайсу (режим force - все позиции)")
    elif choice == 'skip':
        print("ℹ️ Изменения в матрице сохранены, но не применены к прайсу.")
        print("   Используйте ensure_catid_by_fgrp() для применения.")
    else:  # fill по умолчанию
        if import_mode == "fill" and changed_fgrps:
            # Применяем с учётом изменённых Fgrp
            ensure_catid_by_fgrp(mode="fill", changed_fgrps=changed_fgrps)
            print(f"✅ Изменения применены к прайсу:")
            print(f"   - Изменённые Fgrp ({len(changed_fgrps)} шт.) применены ко ВСЕМ позициям группы")
            print(f"   - Остальные Fgrp применены только к пустым CatID")
        else:
            ensure_catid_by_fgrp(mode="fill")
            print("✅ Изменения применены к прайсу (режим fill - только пустые)")


# ============================================================================
# Отчёты и анализ
# ============================================================================

def analyze_fgrp_catid_distribution(fgrp: int | None = None, export_to_excel: bool = True):
    """Анализирует распределение CatID внутри FunctionGroup и определяет причины различий.
    
    Для каждой Fgrp показывает:
    - Какие CatID там есть и сколько позиций с каждым CatID
    - По каким правилам они могли быть заполнены (UnitSort, ProductGroup, Fgrp matrix, Overrides)
    - Рекомендации по унификации
    
    Args:
        fgrp: Если указан, анализирует только этот FunctionGroup. Если None - анализирует все Fgrp с разными CatID.
        export_to_excel: Если True, экспортирует детальный отчёт в Excel.
    
    Returns:
        DataFrame с результатами анализа
    """
    ensure_folder_exists()
    ensure_fgrp_table_exists()
    
    conn = sqlite3.connect(DATABASE_PATH)
    empty_condition = get_empty_catid_condition()
    
    # Получаем распределение CatID по Fgrp
    query = f"""
        SELECT 
            FunctionGroup AS Fgrp,
            CatID,
            UnitSort,
            ProductGroup AS PG,
            COUNT(*) AS Count,
            COUNT(DISTINCT PartNumber) AS PartCount
        FROM {PRICELIST_TABLE_NAME}
        WHERE FunctionGroup IS NOT NULL
          AND NOT ({empty_condition})
    """
    
    if fgrp is not None:
        query += f" AND FunctionGroup = {fgrp}"
    
    query += """
        GROUP BY FunctionGroup, CatID, UnitSort, ProductGroup
        ORDER BY FunctionGroup, Count DESC
    """
    
    df_dist = pd.read_sql(query, conn)
    
    if df_dist.empty:
        print("ℹ️ Нет данных для анализа.")
        conn.close()
        return None
    
    # Загружаем матрицу Fgrp -> CatID
    df_matrix = pd.read_sql(f"SELECT Fgrp, CatID AS MatrixCatID FROM {FGRP_TABLE_NAME}", conn)
    matrix_dict = {}
    for _, row in df_matrix.iterrows():
        fgrp_val = int(row["Fgrp"])
        catid_val = str(row["MatrixCatID"]).strip() if pd.notna(row["MatrixCatID"]) else None
        matrix_dict[fgrp_val] = catid_val if catid_val and catid_val != "" else None
    
    # Загружаем overrides (если есть таблица)
    overrides_set = set()  # Можно расширить, если есть таблица overrides
    
    # Анализируем каждую Fgrp
    results = []
    
    for fgrp_val in df_dist["Fgrp"].unique():
        df_fgrp = df_dist[df_dist["Fgrp"] == fgrp_val].copy()
        unique_catids = df_fgrp["CatID"].unique()
        
        if len(unique_catids) <= 1:
            continue  # Пропускаем Fgrp с одним CatID
        
        # Определяем причины для каждого CatID
        for catid in unique_catids:
            df_catid = df_fgrp[df_fgrp["CatID"] == catid]
            total_count = df_catid["Count"].sum()
            part_count = df_catid["PartCount"].sum()
            
            # Анализируем причины
            reasons = []
            primary_criterion = None  # Основной критерий заполнения
            
            # 1. UnitSort (высший приоритет)
            sw_count = df_catid[df_catid["UnitSort"] == "SW"]["Count"].sum()
            hw_count = df_catid[df_catid["UnitSort"] == "HW"]["Count"].sum()
            
            if catid == "1000" and sw_count > 0:
                reason_text = f"UnitSort=SW ({sw_count} поз.)"
                reasons.append(reason_text)
                if not primary_criterion:
                    primary_criterion = "UnitSort (SW → 1000)"
            elif catid == "4" and hw_count > 0:
                reason_text = f"UnitSort=HW ({hw_count} поз.)"
                reasons.append(reason_text)
                if not primary_criterion:
                    primary_criterion = "UnitSort (HW → 4)"
            
            # 2. ProductGroup
            # Проверяем правила из CatID_lib и известные правила из пайплайна
            pg_reasons = []
            pg_hard_found = False
            pg_soft_found = False
            
            for pg_val in df_catid["PG"].unique():
                if pd.notna(pg_val):
                    pg_val_int = int(pg_val)
                    pg_count = df_catid[df_catid["PG"] == pg_val]["Count"].sum()
                    
                    # Проверяем правила из CatID_lib (PG_MAP)
                    expected_catid = PG_MAP.get(pg_val_int)
                    
                    if expected_catid == catid:
                        pg_reasons.append(f"PG={pg_val} ({pg_count} поз.)")
                        if not primary_criterion:
                            primary_criterion = f"ProductGroup (PG={pg_val})"
                        pg_hard_found = True
            
            if pg_reasons:
                reasons.extend(pg_reasons)
            
            # 3. Fgrp matrix
            matrix_catid = matrix_dict.get(fgrp_val)
            if matrix_catid == catid:
                reason_text = f"Fgrp matrix ({total_count} поз.)"
                reasons.append(reason_text)
                if not primary_criterion:
                    primary_criterion = "Fgrp Matrix"
            elif matrix_catid is not None and matrix_catid != catid:
                reasons.append(f"⚠️ НЕ соответствует матрице (матрица: {matrix_catid})")
            
            # 4. Overrides (если есть)
            # Можно добавить проверку таблицы overrides
            
            # 5. Прочие причины
            if not reasons:
                reasons.append("❓ Неизвестная причина (возможно Backfill или ручное заполнение)")
                primary_criterion = "Неизвестно (Backfill/Ручное)"
            
            # Если не определён основной критерий, но есть причины - берём первую
            if not primary_criterion and reasons:
                if "UnitSort" in reasons[0]:
                    primary_criterion = "UnitSort"
                elif "PG" in reasons[0]:
                    primary_criterion = "ProductGroup"
                elif "Fgrp matrix" in reasons[0]:
                    primary_criterion = "Fgrp Matrix"
                else:
                    primary_criterion = "Другое"
            
            results.append({
                "Fgrp": fgrp_val,
                "CatID": catid,
                "Count": total_count,
                "PartCount": part_count,
                "Primary_Criterion": primary_criterion if primary_criterion else "Неизвестно",
                "Reasons": " | ".join(reasons) if reasons else "Неизвестно",
                "MatrixCatID": matrix_catid if matrix_catid else "",
                "SW_Count": sw_count,
                "HW_Count": hw_count,
                "Unique_PG": ", ".join([str(int(pg)) for pg in df_catid["PG"].unique() if pd.notna(pg)])
            })
    
    conn.close()
    
    if not results:
        print("✅ Все FunctionGroup имеют единый CatID (нет различий для анализа).")
        return None
    
    df_results = pd.DataFrame(results)
    
    # Группируем по Fgrp для сводки
    df_summary = df_results.groupby("Fgrp").agg({
        "CatID": lambda x: ", ".join(x),
        "Count": "sum",
        "PartCount": "sum",
        "MatrixCatID": "first"
    }).reset_index()
    df_summary.columns = ["Fgrp", "CatIDs", "Total_Count", "Total_PartCount", "MatrixCatID"]
    df_summary["CatID_Count"] = df_summary["CatIDs"].apply(lambda x: len(x.split(", ")))
    
    # Выводим результаты
    print(f"\n📊 Анализ распределения CatID по FunctionGroup:")
    print(f"   Найдено {len(df_summary)} FunctionGroup с разными CatID")
    print(f"\n   Топ-10 FunctionGroup с наибольшим количеством разных CatID:")
    
    for _, row in df_summary.nlargest(10, "CatID_Count").iterrows():
        print(f"\n   Fgrp={row['Fgrp']}: {row['CatIDs']} ({row['CatID_Count']} разных CatID)")
        print(f"      Всего позиций: {row['Total_Count']}, Матрица: {row['MatrixCatID'] if row['MatrixCatID'] else 'не задана'}")
        
        # Показываем детали для этого Fgrp
        df_fgrp_detail = df_results[df_results["Fgrp"] == row["Fgrp"]]
        for _, detail_row in df_fgrp_detail.iterrows():
            print(f"      - CatID={detail_row['CatID']}: {detail_row['Count']} поз. ({detail_row['Reasons']})")
    
    # Экспортируем в Excel
    if export_to_excel:
        report_path = os.path.join(FOLDER_PATH, f"fgrp_catid_analysis_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx")
        
        with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
            # Сводка по Fgrp
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
            
            # Детальный анализ
            df_results.to_excel(writer, sheet_name="Details", index=False)
            
            # Инструкции
            df_instructions = pd.DataFrame({
                "Инструкция": [
                    "Этот отчёт показывает причины различий CatID внутри одной FunctionGroup.",
                    "",
                    "Лист Summary:",
                    "- Fgrp: FunctionGroup",
                    "- CatIDs: список всех CatID в этой Fgrp",
                    "- CatID_Count: количество разных CatID",
                    "- MatrixCatID: значение из матрицы Fgrp -> CatID",
                    "",
                    "Лист Details:",
                    "- Fgrp: FunctionGroup",
                    "- CatID: конкретное значение CatID",
                    "- Count: количество позиций с этим CatID",
                    "- Primary_Criterion: ОСНОВНОЙ критерий заполнения (UnitSort/ProductGroup/Fgrp Matrix/Неизвестно)",
                    "- Reasons: детальные причины заполнения (UnitSort, ProductGroup, Fgrp matrix и т.д.)",
                    "- MatrixCatID: значение из матрицы Fgrp -> CatID",
                    "- SW_Count: количество позиций с UnitSort='SW'",
                    "- HW_Count: количество позиций с UnitSort='HW'",
                    "- Unique_PG: список ProductGroup для этого CatID",
                    "",
                    "Возможные причины различий:",
                    "1. UnitSort: SW -> 1000, HW -> 4 (высший приоритет)",
                    "2. ProductGroup: правила из PG_MAP (шаг 2 пайплайна)",
                    "3. Fgrp matrix: значение из таблицы fgrp_category",
                    "4. Overrides: ручные переопределения по PartNumber",
                    "5. Backfill: наследование из предыдущего прайса",
                    "",
                    "Рекомендации:",
                    "- Если в Fgrp есть разные CatID из-за UnitSort (SW/HW) - это нормально",
                    "- Если различия из-за ProductGroup - проверьте правильность правил PG",
                    "- Если CatID не соответствует матрице - обновите матрицу или проверьте приоритеты",
                    "- Если причины неизвестны - возможно, это Backfill из старого прайса"
                ]
            })
            df_instructions.to_excel(writer, sheet_name="Инструкции", index=False)
        
        print(f"\n📄 Детальный отчёт сохранён: {report_path}")
    
    return df_results

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
    overrides_force: bool = False,
):
    """Единая точка входа для выравнивания CatID по текущему контракту.
    
    Выполняет полный пайплайн заполнения категорий CatID в следующем порядке:
    1. UnitSort (freeze): SW -> 1000, HW -> 4
    2. ProductGroup: правила из PG_MAP
    3. Fgrp matrix: заполнение по матрице FunctionGroup -> CatID
    4. Overrides: ручные переопределения по PartNumber
    5. Backfill: дозаполнение из предыдущего прайса
    6. Quality report: генерация отчёта качества
    
    Args:
        previous_excel_path: Путь к предыдущему прайсу для backfill (опционально)
        fgrp_matrix_excel_path: Путь к Excel с матрицей Fgrp->CatID (опционально)
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

    # 2) ProductGroup правила
    apply_pg_rules()
    print_catid_statistics("После шага 2 (PG)")

    # 3) Fgrp -> CatID (матрица)
    # a) синхронизируем новые Fgrp в таблицу
    # b) при необходимости импортируем правки матрицы из Excel
    sync_new_fgrp_to_table()
    if fgrp_matrix_excel_path:
        import_fgrp_matrix_from_excel(fgrp_matrix_excel_path)
    ensure_catid_by_fgrp(mode="fill")  # fill-only
    print_catid_statistics("После шага 3 (Fgrp matrix)")

    # 4) Overrides (PartNumber)
    apply_catid_overrides(force=overrides_force)
    print_catid_statistics("После шага 4 (Overrides)")

    # 5) Backfill (previous)
    if previous_excel_path:
        backfill_catid_from_previous_excel(previous_excel_path)
        print_catid_statistics("После шага 5 (Backfill)")

    # 6) Quality report
    export_catid_quality_report(previous_excel_path=previous_excel_path)

