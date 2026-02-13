"""
Kaspi Reporter - ежедневный отчёт по нераспознанным товарам
Запуск вручную: python test_steps.py
Запуск по расписанию (9:00): python test_steps.py --schedule
"""

import asyncio
import os
import sys
import aiohttp
from datetime import datetime
from playwright.async_api import async_playwright
from config import (
    KASPI_LOGIN,
    KASPI_PASSWORD,
    KASPI_LOGIN_URL,
    CATEGORY_URLS,
    DOWNLOADS_PATH,
    TELEGRAM_BOT_TOKEN,
    TELEGRAM_CHAT_ID,
    HEADLESS,
    MERCHANTS
)

# Создаём папку для загрузок
os.makedirs(DOWNLOADS_PATH, exist_ok=True)


async def test_step1_login():
    """ЭТАП 1: Авторизация"""
    print("\n" + "="*50)
    print("ЭТАП 1: АВТОРИЗАЦИЯ")
    print("="*50)

    playwright = await async_playwright().start()
    browser = await playwright.chromium.launch(headless=HEADLESS)
    context = await browser.new_context(accept_downloads=True, locale='ru-RU')
    page = await context.new_page()
    page.set_default_timeout(60000)

    try:
        # Переход на страницу входа
        print(f"[1] Переход на {KASPI_LOGIN_URL}")
        await page.goto(KASPI_LOGIN_URL)
        await asyncio.sleep(3)
        print(f"    Текущий URL: {page.url}")

        # Клик на вкладку Email
        email_tab = await page.query_selector('a:has-text("Email"), button:has-text("Email"), [role="tab"]:has-text("Email")')
        if email_tab:
            await email_tab.click()
            await asyncio.sleep(1)
            print("[2] Выбрана вкладка Email")

        # Ввод email
        login_input = await page.query_selector('#user_email_field, input[name="username"], input[placeholder="Email"], input.text-field')
        if login_input:
            await login_input.fill(KASPI_LOGIN)
            print(f"[3] Email введён: {KASPI_LOGIN}")
        else:
            print("[!] ОШИБКА: Поле email не найдено!")
            await browser.close()
            return None, None, None

        # Кнопка "Продолжить"
        continue_btn = await page.query_selector('button:has-text("Продолжить"), button:has-text("Continue"), button[type="submit"]')
        if continue_btn:
            await continue_btn.click()
            print("[4] Нажата кнопка 'Продолжить'")
        await asyncio.sleep(3)

        # Ввод пароля
        password_input = await page.query_selector('input[type="password"], input[name="password"]')
        if password_input:
            await password_input.fill(KASPI_PASSWORD)
            print("[5] Пароль введён")

            login_btn = await page.query_selector('button:has-text("Войти"), button:has-text("Продолжить"), button[type="submit"]')
            if login_btn:
                await login_btn.click()
                print("[6] Нажата кнопка входа")

            await asyncio.sleep(5)
        else:
            print("[!] Поле пароля не найдено")

        print(f"[7] Текущий URL после входа: {page.url}")

        # Проверка успеха
        if "login" not in page.url.lower():
            print("\n[OK] ЭТАП 1 УСПЕШЕН: Авторизация прошла!")
            return browser, context, page
        else:
            print("\n[FAIL] ЭТАП 1 ПРОВАЛЕН: Остались на странице входа")
            await browser.close()
            return None, None, None

    except Exception as e:
        print(f"\n[FAIL] ОШИБКА: {e}")
        await browser.close()
        return None, None, None


async def switch_merchant(page, merchant_id, merchant_name):
    """Переключение на другого мерчанта через выпадающий список"""
    print("\n" + "="*50)
    print(f"ПЕРЕКЛЮЧЕНИЕ НА МЕРЧАНТА: {merchant_name} (ID: {merchant_id})")
    print("="*50)

    try:
        # Ищем кнопку переключения мерчанта в header
        # На скрине это элемент "ID - 30409770" с иконкой стрелки вниз
        # Ищем в верхней части страницы (y < 80)

        dropdown_button = None
        all_elements = await page.query_selector_all('*')

        for elem in all_elements:
            try:
                box = await elem.bounding_box()
                if not box or box['y'] > 80 or box['x'] < 800:  # Справа вверху
                    continue

                text = await elem.inner_text()
                # Ищем элемент с текстом "ID - " (текущий мерчант)
                if text and "ID -" in text and len(text) < 50:
                    # Это наш переключатель
                    dropdown_button = elem
                    current_text = text.strip()
                    print(f"[1] Найден переключатель: '{current_text}'")

                    # Проверяем, уже ли на нужном мерчанте
                    if f"ID - {merchant_id}" in current_text:
                        print(f"[OK] Уже на мерчанте {merchant_name}")
                        return True
                    break
            except:
                continue

        if not dropdown_button:
            print("[!] Переключатель мерчантов не найден")
            return False

        # Кликаем чтобы открыть список
        await dropdown_button.click()
        await asyncio.sleep(2)
        print("[2] Список открыт")

        # Ищем нужный вариант в выпадающем списке
        target_text = f"ID - {merchant_id}"

        # Ищем элемент списка с нужным текстом
        options = await page.query_selector_all('div, li, a, span')
        target_option = None

        for opt in options:
            try:
                text = await opt.inner_text()
                if text and target_text in text and len(text) < 50:
                    box = await opt.bounding_box()
                    # Элемент должен быть видимым (высота > 0) и в разумных координатах
                    if box and box['height'] > 10 and box['y'] > 50 and box['y'] < 300:
                        target_option = opt
                        print(f"    Найден вариант: '{text.strip()}'")
                        break
            except:
                continue

        if target_option:
            await target_option.click()
            await asyncio.sleep(5)
            print(f"[3] Переключаемся на {merchant_name}...")

            # Ждём перезагрузку страницы
            try:
                await page.wait_for_load_state('networkidle', timeout=15000)
            except:
                pass

            print(f"    Текущий URL: {page.url}")
            return True
        else:
            print(f"[!] Вариант '{target_text}' не найден")
            return False

    except Exception as e:
        err_msg = str(e).encode('ascii', errors='replace').decode('ascii')
        print(f"[FAIL] ОШИБКА переключения: {err_msg}")
        return False


async def navigate_and_download(page, category_key, step_label="", merchant_name=""):
    """Переход в раздел и скачивание Excel"""
    print("\n" + "="*50)
    print(f"СКАЧИВАНИЕ EXCEL: {step_label} ({merchant_name})")
    print("="*50)

    # Маппинг ключей категорий на текст вкладок на странице
    TAB_LABELS = {
        "без_привязки": "Без привязки",
        "требуют_доработок": "Требуют доработок",
        "на_проверке": "На проверке",
        "отклонены": "Отклонены",
    }

    try:
        url = CATEGORY_URLS[category_key]
        print(f"[1] Переход на {url}")

        # Сначала переходим на страницу нераспознанных товаров
        if "products/pending" not in page.url:
            target_hash = url.split("#")[1] if "#" in url else ""
            await page.evaluate(f'window.location.hash = "{target_hash}"')
            await asyncio.sleep(3)

        # Кликаем по нужной вкладке на странице
        tab_text = TAB_LABELS.get(category_key, step_label)
        print(f"[2] Клик по вкладке '{tab_text}'...")

        # Ищем вкладку по тексту (текст содержит число в скобках, напр. "Требуют доработок (2)")
        tab = await page.query_selector(f'a:has-text("{tab_text}"):visible')
        if tab:
            await tab.click()
            await asyncio.sleep(5)
            print(f"    Текущий URL: {page.url}")
        else:
            # Fallback: прямой переход по hash
            print(f"    [!] Вкладка '{tab_text}' не найдена, переходим по URL...")
            target_hash = url.split("#")[1] if "#" in url else ""
            await page.evaluate(f'window.location.hash = "{target_hash}"')
            await asyncio.sleep(5)
            print(f"    Текущий URL: {page.url}")

        # Проверяем что URL сменился на нужную категорию
        expected_path = url.split("#")[1] if "#" in url else ""
        if expected_path and expected_path not in page.url:
            print(f"    [!] URL не сменился! Пробуем page.goto...")
            await page.goto(url, timeout=60000)
            await asyncio.sleep(5)
            print(f"    Текущий URL после goto: {page.url}")

        # Ждём загрузки таблицы
        print("[3] Ожидание загрузки страницы...")
        await asyncio.sleep(3)

        # Ищем кнопку скачивания (только видимую!)
        print("[4] Поиск видимой кнопки скачивания...")

        selectors = [
            'button:has-text("Выгрузить в EXCEL"):visible',
            'button:has-text("Выгрузить"):visible',
            'button:has-text("EXCEL"):visible',
            'a:has-text("Выгрузить в EXCEL"):visible',
            'a:has-text("EXCEL"):visible',
            'button:has-text("Скачать"):visible',
            'a:has-text("Скачать"):visible',
        ]
        button = None
        for sel in selectors:
            button = await page.query_selector(sel)
            if button:
                btn_text = await button.inner_text()
                print(f"    Найдена кнопка: '{btn_text.strip()}' (селектор: {sel})")
                break

        if not button:
            # Ищем среди всех видимых кнопок
            buttons = await page.query_selector_all('button')
            print(f"    Найдено кнопок на странице: {len(buttons)}")
            for i, btn in enumerate(buttons):
                visible = await btn.is_visible()
                if not visible:
                    continue
                text = await btn.inner_text()
                if text.strip():
                    print(f"    Кнопка {i}: '{text.strip()}'")
                if 'excel' in text.lower() or 'выгрузить' in text.lower():
                    button = btn
                    break

        if not button:
            print("\n[FAIL] ЭТАП 2 ПРОВАЛЕН: Кнопка Excel не найдена")
            print("   Проверьте браузер - есть ли там товары и кнопка?")
            return None

        print("[5] Кнопка найдена, скачиваем...")

        # Скачиваем файл
        async with page.expect_download(timeout=60000) as download_info:
            await button.click()

        download = await download_info.value
        orig_name = download.suggested_filename or "kaspi_export.xlsx"
        # Добавляем timestamp и название мерчанта
        name_base, name_ext = os.path.splitext(orig_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"{name_base}_{merchant_name}_{timestamp}{name_ext}"
        file_path = os.path.join(DOWNLOADS_PATH, file_name)
        await download.save_as(file_path)

        print(f"    Имя файла: {file_name}")
        if "activeorders" in file_name.lower() or "order" in file_name.lower():
            print("    [!] ВНИМАНИЕ: Скачан файл заказов, а не товаров!")
            print("        Возможно, навигация на страницу товаров не сработала.")
        print(f"\n[OK] ЭТАП 2 УСПЕШЕН: Файл сохранён: {file_path}")
        return file_path

    except Exception as e:
        err_msg = str(e).encode('ascii', errors='replace').decode('ascii')
        print(f"\n[FAIL] ОШИБКА: {err_msg}")
        return None


async def test_step3_parse_excel(file_path):
    """ЭТАП 3: Обработка Excel файла и подсчёт артикулов"""
    print("\n" + "="*50)
    print("ЭТАП 3: ОБРАБОТКА EXCEL")
    print("="*50)

    try:
        import pandas as pd

        print(f"[1] Открываем файл: {file_path}")
        df = pd.read_excel(file_path)

        print(f"[2] Количество строк: {len(df)}")
        print(f"[3] Колонки: {list(df.columns)}")

        # Находим колонку с артикулами
        sku_col = None
        for col in df.columns:
            if 'артикул' in col.lower() or 'sku' in col.lower():
                sku_col = col
                break

        if not sku_col:
            print("[!] Колонка с артикулами не найдена!")
            return None

        print(f"[4] Колонка артикулов: '{sku_col}'")

        # Приводим артикулы к строкам
        skus = df[sku_col].astype(str)

        # Общее количество артикулов
        total_count = len(skus)

        # Артикулы начинающиеся на 30000
        skus_30000 = skus[skus.str.startswith("30000")]
        count_30000 = len(skus_30000)

        print(f"[5] Общее количество артикулов: {total_count}")
        print(f"[6] Артикулов на 30000*: {count_30000}")

        # Показываем первые 5 строк
        print("\n[7] Первые 5 товаров:")
        print("-" * 80)
        name_col = None
        for col in df.columns:
            if 'название' in col.lower() or 'наименование' in col.lower():
                name_col = col
                break
        for i, row in df.head(5).iterrows():
            sku = str(row[sku_col])
            name = str(row[name_col])[:50] if name_col else ""
            print(f"   {sku} | {name}")
        print("-" * 80)

        print(f"\n[OK] ЭТАП 3 УСПЕШЕН: Обработано {total_count} товаров")
        return {"total": total_count, "count_30000": count_30000}

    except Exception as e:
        print(f"\n[FAIL] ОШИБКА: {e}")
        return None


async def send_telegram(message, parse_mode=None):
    """Отправка сообщения в Telegram. Возвращает message_id или None"""
    print(f"  Telegram: {message[:100]}...")

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
    }
    if parse_mode:
        payload["parse_mode"] = parse_mode

    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload) as resp:
                result = await resp.json()
                if result.get("ok"):
                    message_id = result["result"]["message_id"]
                    print(f"  [OK] Отправлено в Telegram! (message_id: {message_id})")
                    return message_id
                else:
                    print(f"  [FAIL] Telegram API ошибка: {result}")
                    return None
    except Exception as e:
        print(f"  [FAIL] ОШИБКА отправки: {e}")
        return None


async def pin_telegram_message(message_id):
    """Закрепление сообщения в чате. Бот должен быть администратором с правом закреплять."""
    print(f"  Закрепление сообщения {message_id}...")

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/pinChatMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "message_id": message_id,
        "disable_notification": True  # Без уведомления о закреплении
    }

    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload) as resp:
                result = await resp.json()
                if result.get("ok"):
                    print("  [OK] Сообщение закреплено!")
                    return True
                else:
                    print(f"  [WARN] Не удалось закрепить: {result}")
                    return False
    except Exception as e:
        print(f"  [WARN] ОШИБКА закрепления: {e}")
        return False


async def send_telegram_file(file_path, caption=""):
    """Отправка файла в Telegram"""
    print(f"  Telegram файл: {file_path}")

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"

    try:
        async with aiohttp.ClientSession() as session:
            with open(file_path, 'rb') as f:
                data = aiohttp.FormData()
                data.add_field('chat_id', TELEGRAM_CHAT_ID)
                data.add_field('document', f, filename=os.path.basename(file_path))
                if caption:
                    data.add_field('caption', caption)

                async with session.post(url, data=data) as resp:
                    result = await resp.json()
                    if result.get("ok"):
                        print(f"  [OK] Файл отправлен: {os.path.basename(file_path)}")
                        return True
                    else:
                        print(f"  [FAIL] Telegram API ошибка: {result}")
                        return False
    except Exception as e:
        print(f"  [FAIL] ОШИБКА отправки файла: {e}")
        return False


async def process_category(page, category_key, step_label, merchant_name=""):
    """Обработка одной категории: скачивание и анализ. Возвращает (stats, file_path) или (None, None)."""
    # Скачивание Excel
    file_path = await navigate_and_download(page, category_key, step_label, merchant_name)
    if not file_path:
        print(f"\n[STOP] Не удалось скачать Excel для '{step_label}'")
        return None, None

    await asyncio.sleep(2)

    # Обработка Excel
    stats = await test_step3_parse_excel(file_path)
    if not stats:
        print(f"\n[STOP] Не удалось обработать Excel для '{step_label}'")
        return None, file_path

    return stats, file_path


def build_report_message(all_results):
    """Формирует сводное сообщение-таблицу для Telegram с несколькими мерчантами"""
    today = datetime.now().strftime("%d.%m.%Y")

    # Короткие названия для таблицы
    labels = [
        ("Без привязки", "без_привязки"),
        ("Треб.доработок", "требуют_доработок"),
        ("На проверке", "на_проверке"),
        ("Отклонены", "отклонены"),
    ]

    lines = []
    lines.append(f"Kaspi Report | {today}")
    lines.append("=" * 45)

    grand_total = 0
    grand_30000 = 0

    # Для каждого мерчанта
    for merchant_name, results in all_results.items():
        lines.append(f"\n{merchant_name}")
        lines.append("-" * 34)
        lines.append(f"{'Категория':<17}|{'Всего':>7}|{'30000':>7}")
        lines.append("-" * 34)

        sum_total = 0
        sum_30000 = 0

        for label, key in labels:
            stats = results.get(key)
            if stats:
                t = stats["total"]
                c = stats["count_30000"]
            else:
                t = 0
                c = 0
            sum_total += t
            sum_30000 += c
            lines.append(f"{label:<17}|{t:>7}|{c:>7}")

        lines.append("-" * 34)
        lines.append(f"{'Итого ' + merchant_name:<17}|{sum_total:>7}|{sum_30000:>7}")

        grand_total += sum_total
        grand_30000 += sum_30000

    # Общий итог
    lines.append("\n" + "=" * 45)
    lines.append(f"{'ВСЕГО':<17}|{grand_total:>7}|{grand_30000:>7}")

    table = "\n".join(lines)

    # Ссылка на Google Таблицу с задачами
    sheets_url = "https://docs.google.com/spreadsheets/d/16NoTXUjutOw_anh_oSuufYEEfEu6FiHGm9OFnrSkdN8"

    return f"<pre>{table}</pre>\n\n<a href=\"{sheets_url}\">Задачи контент-менеджерам</a>"


async def process_merchant(page, merchant, categories):
    """Обработка одного мерчанта: переключение и сбор данных по всем категориям"""
    merchant_id = merchant["id"]
    merchant_name = merchant["name"]

    # Переключаемся на мерчанта (для первого это необязательно, он по умолчанию)
    if MERCHANTS.index(merchant) > 0:
        switched = await switch_merchant(page, merchant_id, merchant_name)
        if not switched:
            print(f"[WARN] Не удалось переключиться на {merchant_name}, пробуем продолжить...")

    # Переходим на страницу нераспознанных товаров
    await page.goto(CATEGORY_URLS["без_привязки"], timeout=60000)
    await asyncio.sleep(3)

    # Собираем статистику и пути к файлам по всем категориям
    results = {}
    files = {}
    for category_key, step_label in categories:
        stats, file_path = await process_category(page, category_key, step_label, merchant_name)
        results[category_key] = stats
        files[category_key] = file_path
        await asyncio.sleep(2)

    return results, files


async def main():
    """Главная функция - запуск всех этапов"""
    print("\n" + "="*50)
    print("KASPI REPORTER")
    print("="*50)

    # ЭТАП 1: Авторизация
    browser, context, page = await test_step1_login()

    if not page:
        print("\n[STOP] Тестирование прервано на этапе 1")
        return

    await asyncio.sleep(2)

    # Список категорий для обработки
    categories = [
        ("без_привязки", "Без привязки"),
        ("требуют_доработок", "Требуют доработок"),
        ("на_проверке", "На проверке"),
        ("отклонены", "Отклонены"),
    ]

    # Словари для хранения данных по всем мерчантам
    all_results = {}  # {merchant_name: {category: stats}}
    all_files = {}    # {merchant_name: {category: file_path}}

    # Обрабатываем каждого мерчанта
    for merchant in MERCHANTS:
        merchant_name = merchant["name"]
        print(f"\n{'='*50}")
        print(f"ОБРАБОТКА МЕРЧАНТА: {merchant_name}")
        print(f"{'='*50}")

        results, files = await process_merchant(page, merchant, categories)
        all_results[merchant_name] = results
        all_files[merchant_name] = files

    # Формируем и отправляем сводный отчёт
    print("\n" + "="*50)
    print("ОТПРАВКА СВОДНОГО ОТЧЁТА В TELEGRAM")
    print("="*50)

    message = build_report_message(all_results)
    message_id = await send_telegram(message, parse_mode="HTML")

    # Закрепляем сообщение в группе (бот должен быть админом)
    if message_id:
        await pin_telegram_message(message_id)

    # Обработка Google Sheets (для "Без привязки" всех мерчантов)
    print("\n" + "="*50)
    print("ОБРАБОТКА GOOGLE SHEETS")
    print("="*50)

    for merchant in MERCHANTS:
        merchant_name = merchant["name"]
        files = all_files.get(merchant_name, {})
        bez_privyazki_file = files.get("без_привязки")

        if bez_privyazki_file and os.path.exists(bez_privyazki_file):
            try:
                from google_sheets import process_products_file
                print(f"\n[Google Sheets] Обработка {merchant_name}...")
                process_products_file(bez_privyazki_file, merchant_name=merchant_name)
            except Exception as e:
                err_msg = str(e).encode('ascii', errors='replace').decode('ascii')
                print(f"[WARN] Ошибка Google Sheets для {merchant_name}: {err_msg}")

    print("\n" + "="*50)
    print("ВСЕ МЕРЧАНТЫ ОБРАБОТАНЫ!")
    print("="*50)

    await asyncio.sleep(2)
    await browser.close()


async def run_scheduled():
    """Запуск по расписанию: каждый день в 9:00"""
    print("[SCHEDULER] Kaspi Reporter запущен в режиме расписания")
    print("[SCHEDULER] Отчёт будет отправляться каждый день в 09:00")

    while True:
        now = datetime.now()
        # Вычисляем время до следующего запуска в 9:00
        target = now.replace(hour=9, minute=0, second=0, microsecond=0)
        if now >= target:
            # Если уже позже 9:00 - ждём до завтра
            from datetime import timedelta
            target += timedelta(days=1)

        wait_seconds = (target - now).total_seconds()
        hours = int(wait_seconds // 3600)
        minutes = int((wait_seconds % 3600) // 60)
        print(f"[SCHEDULER] Следующий запуск: {target.strftime('%d.%m.%Y %H:%M')}"
              f" (через {hours}ч {minutes}мин)")

        await asyncio.sleep(wait_seconds)

        print(f"\n[SCHEDULER] === Запуск отчёта {datetime.now().strftime('%d.%m.%Y %H:%M')} ===")
        try:
            await main()
        except Exception as e:
            err_msg = str(e).encode('ascii', errors='replace').decode('ascii')
            print(f"[SCHEDULER] ОШИБКА: {err_msg}")
            # Отправляем ошибку в Telegram
            await send_telegram(f"Kaspi Reporter: ОШИБКА при запуске\n{err_msg}")


if __name__ == "__main__":
    if "--schedule" in sys.argv:
        asyncio.run(run_scheduled())
    else:
        asyncio.run(main())
