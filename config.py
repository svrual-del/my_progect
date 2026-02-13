# ============================================
# КОНФИГУРАЦИЯ KASPI REPORTER
# ============================================
import os

# Авторизация на Kaspi.kz (из переменных окружения или значения по умолчанию)
KASPI_LOGIN = os.environ.get("KASPI_LOGIN", "Aleksey.Rumyancev@sulpak.kz")
KASPI_PASSWORD = os.environ.get("KASPI_PASSWORD", "resMy3002Kasp!")

# URL-адреса Kaspi Магазин (Merchant Center)
KASPI_LOGIN_URL = "https://idmc.shop.kaspi.kz/login"
KASPI_PRODUCTS_BASE_URL = "https://kaspi.kz/mc/#/products/pending"

# Список мерчантов для обработки
# id — текст в выпадающем списке после "ID - " (Sulpak, 30409770, 30382295)
# name — название для отчётов
# sheet_name — название листа в Excel
MERCHANTS = [
    {"id": "Sulpak", "name": "Sulpak", "sheet_name": "Sulpak"},
    {"id": "30409770", "name": "ARG", "sheet_name": "ARG"},
    {"id": "30382295", "name": "Motorola", "sheet_name": "Motorola"},
]

# URL для каждой категории нераспознанных товаров
CATEGORY_URLS = {
    "без_привязки": "https://kaspi.kz/mc/#/products/pending/CHECK/1",
    "требуют_доработок": "https://kaspi.kz/mc/#/products/pending/IMPORTED/1",
    "на_проверке": "https://kaspi.kz/mc/#/products/pending/PENDING/1",
    "отклонены": "https://kaspi.kz/mc/#/products/pending/TRASH/1"
}

# Telegram настройки (из переменных окружения или значения по умолчанию)
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8510234232:AAHRcMRMP87na4Ci9GjvIb8Sp9Bz_bkzE7Q")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "-1002158563605")

# Настройки браузера
# В CI/облаке автоматически headless, локально — с окном
HEADLESS = os.environ.get("CI", "") == "true"
TIMEOUT = 30000  # Таймаут ожидания элементов (мс)

# Путь для сохранения отчётов
REPORTS_PATH = "./reports"

# Путь для скачанных Excel-файлов от Kaspi
DOWNLOADS_PATH = "./downloads"

# Использовать встроенную выгрузку в Excel (рекомендуется)
USE_BUILTIN_EXCEL_EXPORT = True
