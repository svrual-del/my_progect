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

# Список менеджеров — загружается из managers.txt (по одному на строку)
# Если файл не найден, используется список по умолчанию
MANAGERS_FILE = os.path.join(os.path.dirname(__file__), "managers.txt")

def load_managers():
    """Загрузка списка менеджеров из файла managers.txt"""
    if os.path.exists(MANAGERS_FILE):
        with open(MANAGERS_FILE, "r", encoding="utf-8") as f:
            managers = [line.strip() for line in f if line.strip() and not line.startswith("#")]
        if managers:
            return managers
    # Список по умолчанию, если файл не найден
    return [
        "Дементьева Алина Владимировна",
        "Кононенко Михаил Валерьевич",
        "Величковская Алиса Леонидовна",
        "Асреп Жулдыз Ерланкызы",
        "Тохсеитова Диана Жорабековна",
    ]

MANAGERS = load_managers()

# Цвета для подсветки (RGB от 0 до 1)
COLOR_RED = {"red": 1.0, "green": 0.8, "blue": 0.8}      # Красный фон (ARG)
COLOR_GREEN = {"red": 0.8, "green": 1.0, "blue": 0.8}    # Зелёный фон (не 30000)
COLOR_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}    # Белый (обычный)

# Структура столбцов (добавлен столбец "Кабинет" первым)
COLUMNS = ["Кабинет", "Артикул", "Название товара", "Дата добавления", "Менеджер", "Отметка менеджера", "Дата исчезновения", "Дней до решения"]


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
        sheet.format("A1:H1", {"textFormat": {"bold": True}})
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


def get_sku_rows(sheet):
    """Получить словарь артикулов с номерами строк и кабинетами: {sku: {"row": N, "merchant": "Sulpak"}}"""
    try:
        all_data = sheet.get_all_values()
        sku_map = {}
        for i, row in enumerate(all_data[1:], start=2):
            if len(row) >= 2:
                sku = row[1]       # Столбец B — Артикул
                merchant = row[0]  # Столбец A — Кабинет
                sku_map[sku] = {"row": i, "merchant": merchant}
        return sku_map
    except Exception as e:
        print(f"  [WARN] Ошибка получения артикулов: {e}")
        return {}


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
    Если артикул уже есть в другом кабинете — добавляем кабинет через "+"
    Возвращает количество добавленных
    """
    existing_skus = get_existing_skus(sheet)  # set of (merchant, sku)
    sku_rows = get_sku_rows(sheet)  # {sku: {"row": N, "merchant": "Sulpak"}}
    today = datetime.now().strftime("%d.%m.%Y")

    # Получаем загрузку всех менеджеров одним запросом
    loads = get_manager_loads(sheet)

    added_count = 0
    merged_count = 0
    rows_to_add = []
    row_colors = []

    for product in products:
        sku = str(product["sku"])
        name = product["name"]

        # Пропускаем если уже есть именно в этом кабинете
        if (merchant_name, sku) in existing_skus:
            continue

        # Если артикул есть в другом кабинете — обновляем столбец "Кабинет" через "+"
        if sku in sku_rows:
            info = sku_rows[sku]
            current_merchant = info["merchant"]
            row_num = info["row"]
            # Проверяем что наш кабинет ещё не в списке
            if merchant_name not in current_merchant:
                new_merchant = f"{current_merchant}+{merchant_name}"
                sheet.update_cell(row_num, 1, new_merchant)
                # Обновляем локальный кэш
                sku_rows[sku]["merchant"] = new_merchant
                merged_count += 1
            existing_skus.add((merchant_name, sku))
            continue

        # Случайный выбор менеджера с минимальной загрузкой
        min_load = min(loads.values())
        candidates = [m for m, load in loads.items() if load == min_load]
        manager = random.choice(candidates)
        loads[manager] += 1

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

        existing_skus.add((merchant_name, sku))
        sku_rows[sku] = {"row": -1, "merchant": merchant_name}  # -1 т.к. ещё не записана
        added_count += 1

    # Пакетное добавление строк
    if rows_to_add:
        sheet.append_rows(rows_to_add, value_input_option="USER_ENTERED")
        print(f"  [OK] Добавлено строк: {len(rows_to_add)}")

        # Применяем цвета
        all_values = sheet.get_all_values()
        start_row = len(all_values) - len(rows_to_add) + 1

        for i, color in enumerate(row_colors):
            row_num = start_row + i
            if color != COLOR_WHITE:
                sheet.format(f"A{row_num}:H{row_num}", {"backgroundColor": color})

        print(f"  [OK] Цвета применены")

    if merged_count > 0:
        print(f"  [INFO] Объединено кабинетов (артикул в нескольких): {merged_count}")

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


def setup_days_column(sheet):
    """
    Настройка столбца H (Дней до решения):
    - Если F (отметка) есть, а G (исчезновение) нет — "ОБМАН"
    - Если G заполнена — G - D (дата исчезновения - дата добавления)
    - Зелёный: 1-3 дня (норма)
    - Красный: > 3 дней (долго), ОБМАН
    """
    all_data = sheet.get_all_values()
    if len(all_data) <= 1:
        return  # Только заголовок

    # Проверяем, есть ли столбец H
    headers = all_data[0]
    if len(headers) < 8 or headers[7] != "Дней до решения":
        sheet.update_cell(1, 8, "Дней до решения")
        sheet.format("H1", {"textFormat": {"bold": True}})

    # Обновляем формулы для всех строк
    # Приоритет: ОБМАН (F заполнена, G пустая) > G-D (если G есть) > пусто
    updates = []
    for i in range(2, len(all_data) + 1):
        formula = f'=ЕСЛИ(И(F{i}<>"";G{i}="");"ОБМАН";ЕСЛИ(G{i}<>"";ЦЕЛОЕ(G{i}-D{i});""))'
        updates.append({"range": f"H{i}", "values": [[formula]]})

    if updates:
        sheet.batch_update(updates, value_input_option="USER_ENTERED")

    # Принудительно задаём числовой формат для столбца H (чтобы не показывало дату)
    try:
        sheet.format(f"H2:H{len(all_data)}", {"numberFormat": {"type": "NUMBER", "pattern": "0"}})
    except Exception:
        pass

    # Условное форматирование и валидация
    try:
        spreadsheet = sheet.spreadsheet

        requests = [
            # Красный жирный для "ОБМАН" (отметка менеджера есть, а товар не исчез)
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet.id,
                            "startColumnIndex": 7,  # H = индекс 7
                            "endColumnIndex": 8,
                            "startRowIndex": 1
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "TEXT_EQ",
                                "values": [{"userEnteredValue": "ОБМАН"}]
                            },
                            "format": {
                                "backgroundColor": {"red": 1.0, "green": 0.6, "blue": 0.6},
                                "textFormat": {"bold": True}
                            }
                        }
                    },
                    "index": 0
                }
            },
            # Красный для значений > 3
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet.id,
                            "startColumnIndex": 7,
                            "endColumnIndex": 8,
                            "startRowIndex": 1
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "NUMBER_GREATER",
                                "values": [{"userEnteredValue": "3"}]
                            },
                            "format": {
                                "backgroundColor": {"red": 1.0, "green": 0.8, "blue": 0.8}
                            }
                        }
                    },
                    "index": 1
                }
            },
            # Зелёный для значений 1-3
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet.id,
                            "startColumnIndex": 7,
                            "endColumnIndex": 8,
                            "startRowIndex": 1
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "NUMBER_BETWEEN",
                                "values": [
                                    {"userEnteredValue": "1"},
                                    {"userEnteredValue": "3"}
                                ]
                            },
                            "format": {
                                "backgroundColor": {"red": 0.8, "green": 1.0, "blue": 0.8}
                            }
                        }
                    },
                    "index": 2
                }
            },
            # Data Validation для столбца F (Отметка менеджера) — только дата
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet.id,
                        "startColumnIndex": 5,  # F = индекс 5
                        "endColumnIndex": 6,
                        "startRowIndex": 1
                    },
                    "rule": {
                        "condition": {
                            "type": "DATE_IS_VALID"
                        },
                        "strict": True,
                        "showCustomUi": True
                    }
                }
            }
        ]

        spreadsheet.batch_update({"requests": requests})
        print("  [OK] Столбец 'Дней до решения' настроен (G-D, зелёный 1-3, красный >3)")

    except Exception as e:
        print(f"  [WARN] Ошибка настройки условного форматирования: {e}")


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

    # Настраиваем столбец "Дней до решения" (только один раз за сессию)
    print("[6] Настройка столбца 'Дней до решения'...")
    setup_days_column(sheet)

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
