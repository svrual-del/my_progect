# ============================================
# МОДУЛЬ РАБОТЫ С GOOGLE SHEETS
# ============================================

import os
import re
import random
import gspread
from datetime import datetime
from google.oauth2.service_account import Credentials

# Настройки
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "google-credentials.json")
SPREADSHEET_ID = "16NoTXUjutOw_anh_oSuufYEEfEu6FiHGm9OFnrSkdN8"

# Список менеджеров
MANAGERS = [
    "Котенко Екатерина Константиновна",
    "Дементьева Алина Владимировна",
    "Кононенко Михаил Валерьевич",
    "Величковская Алиса Леонидовна",
    "Асреп Жулдыз Ерланкызы",
    "Тохсеитова Диана Жорабековна",
    "Сариева Дана Агадилдақызы",
]

# Цвета для подсветки (RGB от 0 до 1)
COLOR_RED = {"red": 1.0, "green": 0.8, "blue": 0.8}      # Красный фон (ARG)
COLOR_GREEN = {"red": 0.8, "green": 1.0, "blue": 0.8}    # Зелёный фон (не 30000)
COLOR_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}    # Белый (обычный)

# Структура столбцов (добавлен столбец "Кабинет" первым)
COLUMNS = ["Кабинет", "Артикул", "Название товара", "Дата добавления", "Менеджер", "Отметка менеджера", "Дата исчезновения"]


def get_sheet():
    """Подключение к Google Sheets"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    return spreadsheet


def get_or_create_month_sheet(spreadsheet, month_name=None):
    """Получить или создать лист для текущего месяца"""
    if month_name is None:
        # Формат: "2026-02" для февраля 2026
        month_name = datetime.now().strftime("%Y-%m")

    # Ищем существующий лист
    try:
        sheet = spreadsheet.worksheet(month_name)
        print(f"  [OK] Найден лист: {month_name}")

        # Проверяем, есть ли столбец "Кабинет" (миграция старых листов)
        headers = sheet.row_values(1)
        if headers and headers[0] != "Кабинет":
            print(f"  [INFO] Миграция листа: добавление столбца 'Кабинет'...")
            # Вставляем новый столбец A
            sheet.insert_cols([["Кабинет"]], col=1)
            # Помечаем старые записи как "Sulpak" (для совместимости)
            all_data = sheet.get_all_values()
            if len(all_data) > 1:
                updates = [["Sulpak"] for _ in range(len(all_data) - 1)]
                sheet.update(f"A2:A{len(all_data)}", updates)
            print(f"  [OK] Миграция завершена")

    except gspread.WorksheetNotFound:
        # Создаём новый лист с заголовками
        sheet = spreadsheet.add_worksheet(title=month_name, rows=1000, cols=10)
        sheet.append_row(COLUMNS)
        # Форматируем заголовок (жирный)
        sheet.format("A1:G1", {"textFormat": {"bold": True}})
        print(f"  [OK] Создан новый лист: {month_name}")

    return sheet


def get_existing_skus(sheet):
    """Получить список уже добавленных артикулов со всех кабинетов"""
    try:
        # Получаем все значения из столбцов A (кабинет) и B (артикулы)
        all_data = sheet.get_all_values()
        # Создаём set кортежей (кабинет, артикул) для уникальности
        skus = set()
        for row in all_data[1:]:  # Пропускаем заголовок
            if len(row) >= 2:
                merchant = row[0]  # Столбец A — Кабинет
                sku = row[1]       # Столбец B — Артикул
                skus.add((merchant, sku))
        return skus
    except Exception as e:
        print(f"  [WARN] Ошибка получения артикулов: {e}")
        return set()


def get_all_skus(sheet):
    """Получить все артикулы без учёта кабинета (для проверки дубликатов между кабинетами)"""
    try:
        all_data = sheet.get_all_values()
        skus = set()
        for row in all_data[1:]:
            if len(row) >= 2:
                skus.add(row[1])  # Столбец B — Артикул
        return skus
    except Exception as e:
        print(f"  [WARN] Ошибка получения артикулов: {e}")
        return set()


def get_manager_loads(sheet):
    """Получить загрузку всех менеджеров одним запросом"""
    try:
        all_data = sheet.get_all_values()
        loads = {manager: 0 for manager in MANAGERS}

        for row in all_data[1:]:  # Пропускаем заголовок
            # Столбец E (индекс 4) — Менеджер
            # Столбец F (индекс 5) — Отметка менеджера (пусто = не завершено)
            if len(row) > 5 and row[4] in loads and not row[5]:
                loads[row[4]] += 1

        return loads
    except Exception as e:
        print(f"  [WARN] Ошибка подсчёта задач: {e}")
        return {manager: 0 for manager in MANAGERS}


def get_least_loaded_manager(loads):
    """Найти менеджера с минимальной загрузкой"""
    return min(loads, key=loads.get)


def is_arg_product(product_name):
    """Проверить, содержит ли название бренд ARG"""
    # Ищем слово ARG отдельно (не как часть другого слова)
    return bool(re.search(r'\bARG\b', product_name, re.IGNORECASE))


def is_not_30000(sku):
    """Проверить, что артикул НЕ начинается на 30000"""
    return not str(sku).startswith("30000")


def add_products_to_sheet(sheet, products, merchant_name="Sulpak"):
    """
    Добавить товары в таблицу
    products: список словарей {"sku": "123", "name": "Товар..."}
    merchant_name: название кабинета (Sulpak, ARG и т.д.)
    Возвращает количество добавленных
    """
    existing_skus = get_existing_skus(sheet)  # set of (merchant, sku)
    all_skus = get_all_skus(sheet)  # set of sku (для проверки дубликатов между кабинетами)
    today = datetime.now().strftime("%d.%m.%Y")

    # Получаем загрузку всех менеджеров одним запросом
    loads = get_manager_loads(sheet)

    added_count = 0
    skipped_duplicates = 0
    rows_to_add = []
    row_colors = []

    for product in products:
        sku = str(product["sku"])
        name = product["name"]

        # Пропускаем если уже есть в этом кабинете
        if (merchant_name, sku) in existing_skus:
            continue

        # Пропускаем если артикул уже есть в другом кабинете (чтобы не дублировать)
        if sku in all_skus:
            skipped_duplicates += 1
            continue

        # Случайный выбор менеджера с минимальной загрузкой
        # Сначала находим минимальную загрузку
        min_load = min(loads.values())
        # Берём всех менеджеров с минимальной загрузкой
        candidates = [m for m, load in loads.items() if load == min_load]
        # Случайный выбор из них
        manager = random.choice(candidates)
        loads[manager] += 1  # Увеличиваем локальный счётчик

        # Добавляем строку (с кабинетом в первом столбце)
        row = [merchant_name, sku, name, today, manager, "", ""]
        rows_to_add.append(row)

        # Определяем цвет
        if is_arg_product(name):
            row_colors.append(COLOR_RED)
        elif is_not_30000(sku):
            row_colors.append(COLOR_GREEN)
        else:
            row_colors.append(COLOR_WHITE)

        existing_skus.add((merchant_name, sku))  # Чтобы не дублировать в этой сессии
        all_skus.add(sku)
        added_count += 1

    # Пакетное добавление строк
    if rows_to_add:
        sheet.append_rows(rows_to_add)
        print(f"  [OK] Добавлено строк: {len(rows_to_add)}")

        # Применяем цвета
        # Находим номер первой добавленной строки
        all_values = sheet.get_all_values()
        start_row = len(all_values) - len(rows_to_add) + 1

        for i, color in enumerate(row_colors):
            row_num = start_row + i
            if color != COLOR_WHITE:
                sheet.format(f"A{row_num}:G{row_num}", {"backgroundColor": color})

        print(f"  [OK] Цвета применены")

    if skipped_duplicates > 0:
        print(f"  [INFO] Пропущено дубликатов (есть в другом кабинете): {skipped_duplicates}")

    return added_count


def check_disappeared_products(sheet, current_skus, merchant_name="Sulpak"):
    """
    Проверить исчезнувшие товары и поставить дату исчезновения
    current_skus: set артикулов из текущего файла "Без привязки"
    merchant_name: проверяем только для конкретного кабинета
    """
    today = datetime.now().strftime("%d.%m.%Y")
    current_skus_set = set(str(s) for s in current_skus)

    all_data = sheet.get_all_values()
    updated_count = 0

    for i, row in enumerate(all_data[1:], start=2):  # start=2 потому что строка 1 = заголовок
        if len(row) < 7:
            continue

        merchant = row[0]           # Столбец A — Кабинет
        sku = row[1]                # Столбец B — Артикул
        date_disappeared = row[6]   # Столбец G — Дата исчезновения

        # Проверяем только для нужного кабинета
        if merchant != merchant_name:
            continue

        # Если артикул исчез и дата ещё не проставлена
        if sku not in current_skus_set and not date_disappeared:
            sheet.update_cell(i, 7, today)  # Столбец G = 7
            updated_count += 1

    if updated_count:
        print(f"  [OK] Отмечено исчезнувших ({merchant_name}): {updated_count}")

    return updated_count


def process_products_file(excel_path, merchant_name="Sulpak"):
    """
    Основная функция: обработать файл "Без привязки"
    - Добавить новые товары в таблицу
    - Отметить исчезнувшие
    merchant_name: название кабинета (Sulpak, ARG и т.д.)
    """
    import pandas as pd

    print("\n" + "="*50)
    print(f"ОБРАБОТКА GOOGLE SHEETS ({merchant_name})")
    print("="*50)

    # Читаем Excel
    print(f"[1] Чтение файла: {excel_path}")
    df = pd.read_excel(excel_path)

    # Находим столбцы
    sku_col = None
    name_col = None
    for col in df.columns:
        if 'артикул' in col.lower():
            sku_col = col
        if 'название' in col.lower():
            name_col = col

    if not sku_col or not name_col:
        print(f"  [FAIL] Не найдены столбцы. Есть: {list(df.columns)}")
        return

    # Формируем список товаров
    products = []
    current_skus = set()
    for _, row in df.iterrows():
        sku = str(row[sku_col])
        name = str(row[name_col])
        products.append({"sku": sku, "name": name})
        current_skus.add(sku)

    print(f"[2] Всего товаров в файле: {len(products)}")

    # Подключаемся к Google Sheets
    print("[3] Подключение к Google Sheets...")
    spreadsheet = get_sheet()
    sheet = get_or_create_month_sheet(spreadsheet)

    # Добавляем новые товары
    print("[4] Добавление новых товаров...")
    added = add_products_to_sheet(sheet, products, merchant_name)
    print(f"    Новых добавлено: {added}")

    # Проверяем исчезнувшие (только для данного кабинета)
    print("[5] Проверка исчезнувших товаров...")
    disappeared = check_disappeared_products(sheet, current_skus, merchant_name)
    print(f"    Исчезло: {disappeared}")

    print(f"\n[OK] Google Sheets обработан для {merchant_name}!")
    return {"added": added, "disappeared": disappeared}


# Тест при прямом запуске
if __name__ == "__main__":
    print("Тест подключения к Google Sheets...")
    try:
        spreadsheet = get_sheet()
        print(f"[OK] Подключено к: {spreadsheet.title}")

        sheet = get_or_create_month_sheet(spreadsheet)
        print(f"[OK] Текущий лист: {sheet.title}")

        existing = get_existing_skus(sheet)
        print(f"[OK] Записей в листе: {len(existing)}")
    except Exception as e:
        print(f"[FAIL] Ошибка: {e}")
