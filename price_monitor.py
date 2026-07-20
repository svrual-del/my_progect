"""
Kaspi Price Monitor — мониторинг истории загрузок прайс-листов (кабинет Sulpak).

Проверяет ПОСЛЕДНЮЮ строку на странице "История загрузок" → "Прайс-листы".
Если статус — "Ошибка загрузки файла", шлёт алерт в Telegram.
Повторный алерт по той же строке не отправляется (состояние в monitor_state.json).

Запуск:            py price_monitor.py
Проверка парсинга: py price_monitor.py --dry-run  (без Telegram и записи состояния)
"""

import asyncio
import json
import os
import sys

from test_steps import test_step1_login, switch_merchant, send_telegram

HISTORY_URL = "https://kaspi.kz/mc/#/history?tab=priceList&page=1"
STATE_FILE = "monitor_state.json"
MERCHANT_ID = "Sulpak"
MERCHANT_NAME = "Sulpak"


def load_state():
    """Читаем состояние (дата последней строки, по которой уже был алерт)"""
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


async def open_history(page):
    """Переход на страницу истории загрузок и ожидание загрузки SPA.

    После холодного старта SPA игнорирует hash и уводит на страницу заказов,
    поэтому при необходимости повторяем hash-навигацию.
    """
    print(f"[1] Переход на {HISTORY_URL}")
    for attempt in range(3):
        await page.goto(HISTORY_URL, timeout=60000)
        await asyncio.sleep(5)
        try:
            await page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        if "history" in page.url:
            return True
        print(f"    SPA увела на {page.url}, повторный переход ({attempt + 1})...")
    return "history" in page.url


async def ensure_sulpak(page):
    """Проверяем активный кабинет по header; переключаемся при необходимости.

    После логина по умолчанию активен Sulpak, поэтому обычно переключение не нужно.
    """
    try:
        el = await page.query_selector('a.navbar-link:has-text("ID -")')
        text = (await el.inner_text()).strip() if el else ""
    except Exception:
        text = ""

    if f"ID - {MERCHANT_ID}" in text:
        print(f"[OK] Активный кабинет: {text}")
        return True

    print(f"[!] Активный кабинет: '{text}', переключаемся на {MERCHANT_NAME}...")
    if not await switch_merchant(page, MERCHANT_ID, MERCHANT_NAME):
        return False
    return await open_history(page)


async def parse_latest_row(page):
    """Парсим первую (самую свежую) строку таблицы истории загрузок.

    Колонки: Название файла | Статус | Загружено предложений | Дата загрузки
    """
    # На других страницах SPA тоже есть таблицы (например, заказы) —
    # ищем именно таблицу истории по заголовку "Название файла"
    for table in await page.query_selector_all("table"):
        try:
            if not await table.is_visible():
                continue
            header = await table.query_selector("tr")
            header_text = (await header.inner_text()) if header else ""
            if "Название файла" not in header_text:
                continue
            for row in await table.query_selector_all("tr"):
                cells = await row.query_selector_all("td")
                if len(cells) < 4:
                    continue  # заголовок (th) или служебная строка
                texts = [(await c.inner_text()).strip() for c in cells]
                return {
                    "file": texts[0],
                    "status": texts[1],
                    "offers": texts[2],
                    "date": texts[3],
                }
        except Exception:
            continue
    return None


async def save_debug(page):
    """Сохраняем скриншот и HTML для диагностики, если таблица не найдена"""
    os.makedirs("downloads", exist_ok=True)
    try:
        await page.screenshot(path="downloads/history_debug.png", full_page=True)
        html = await page.content()
        with open("downloads/history_debug.html", "w", encoding="utf-8") as f:
            f.write(html)
        print("    Сохранено: downloads/history_debug.png, downloads/history_debug.html")
    except Exception as e:
        print(f"    Не удалось сохранить отладку: {e}")


def build_alert_message(row):
    return (
        "\U0001F6A8 <b>Ошибка загрузки прайс-листа!</b>\n\n"
        f"Кабинет: {MERCHANT_NAME}\n"
        f"Файл: {row['file']}\n"
        f"Статус: {row['status']}\n"
        f"Загружено предложений: {row['offers']}\n"
        f"Дата загрузки: {row['date']}"
    )


async def main(dry_run=False):
    print("=" * 50)
    print("KASPI PRICE MONITOR: история загрузок прайс-листов")
    print("=" * 50)

    browser, context, page = await test_step1_login()
    if not page:
        print("[FAIL] Авторизация не удалась")
        sys.exit(1)

    try:
        if not await open_history(page):
            print("[FAIL] Не удалось открыть страницу истории загрузок")
            await save_debug(page)
            sys.exit(1)

        # Убеждаемся, что мы в кабинете Sulpak
        if not await ensure_sulpak(page):
            print("[FAIL] Не удалось переключиться на Sulpak")
            sys.exit(1)

        row = await parse_latest_row(page)
        if not row:
            print("[FAIL] Не удалось найти строки таблицы истории загрузок")
            await save_debug(page)
            sys.exit(1)

        print(f"[2] Последняя загрузка: {row['file']} | {row['status']} | "
              f"{row['offers']} | {row['date']}")

        is_error = "ошибка" in row["status"].lower()

        if not is_error:
            print("[OK] Статус в норме, алерт не нужен")
            return

        state = load_state()
        if state.get("last_alerted_date") == row["date"]:
            print(f"[OK] По этой строке ({row['date']}) алерт уже был, пропускаем")
            return

        print("[!] Обнаружена ОШИБКА загрузки файла!")

        if dry_run:
            print("[DRY-RUN] Telegram не отправляем. Сообщение было бы таким:")
            print(build_alert_message(row).encode('ascii', errors='replace').decode('ascii'))
            return

        message_id = await send_telegram(build_alert_message(row), parse_mode="HTML")
        if message_id:
            save_state({"last_alerted_date": row["date"]})
            print("[OK] Алерт отправлен, состояние сохранено")
        else:
            print("[FAIL] Не удалось отправить алерт в Telegram")
            sys.exit(1)

    finally:
        await browser.close()


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv
    asyncio.run(main(dry_run=dry_run))
