# Kaspi Reporter

Автоматический ежедневный отчёт по нераспознанным товарам из Kaspi Merchant Center.

## Архитектура

```
test_steps.py        — основной скрипт (авторизация, навигация, скачивание, отчёт)
google_sheets.py     — модуль работы с Google Sheets (добавление товаров, менеджеры)
config.py            — конфигурация (секреты из env или дефолты)
.github/workflows/report.yml — GitHub Actions cron (03:00 UTC = 08:00 Алматы)
```

## Ключевые моменты

### Навигация в Kaspi MC
- SPA с hash-роутингом (`#/products/pending/CHECK/1`)
- Переключение вкладок через клик по `<a>` с текстом категории
- Кнопка "Выгрузить в EXCEL" — это `<a>`, не `<button>`
- Всегда ищем только `:visible` элементы (SPA держит скрытые DOM-элементы)

### Категории нераспознанных товаров
| Ключ | URL-часть | Название |
|------|-----------|----------|
| без_привязки | CHECK | Без привязки |
| требуют_доработок | IMPORTED | Требуют доработок |
| на_проверке | PENDING | На проверке |
| отклонены | TRASH | Отклонены |

### Telegram
- Сводная таблица в `<pre>` с `parse_mode: HTML`
- Автозакрепление сообщения в группе (бот должен быть админом)
- 3 Excel-файла отправляются после таблицы (кроме "Без привязки")
- Ссылка на Google Таблицу с задачами
- Bot API через aiohttp, функции: `send_telegram()`, `pin_telegram_message()`, `send_telegram_file()`

### Google Sheets
- Spreadsheet ID: `16NoTXUjutOw_anh_oSuufYEEfEu6FiHGm9OFnrSkdN8`
- Ежемесячные листы (формат: `2026-02`)
- Товары из "Без привязки" добавляются с назначением менеджера
- Цвета: красный = ARG, зелёный = не 30000*
- Отслеживание исчезнувших товаров (столбец F)

### Секреты
В GitHub Secrets (Settings → Secrets → Actions):
- `KASPI_LOGIN`
- `KASPI_PASSWORD`
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID` — ID группы (формат: `-1002158563605`)
- `GOOGLE_CREDENTIALS` — JSON Service Account одной строкой

Локально — дефолтные значения в `config.py`.

## Известные проблемы

1. **GitHub Actions cron задержка** — до 30-60 минут, поэтому cron на 03:00 UTC (08:00 Алматы) чтобы реально пришло ближе к 9:00

2. **cp1251 на Windows** — не использовать emoji в print(), заменять на `[OK]`, `[FAIL]`

3. **Файл заблокирован** — Excel-файлы сохраняем с timestamp в имени

4. **Закрепление не работает** — бот должен быть админом группы с правом "Закреплять сообщения"

## Команды

```bash
# Локальный запуск (с окном браузера)
py test_steps.py

# Локальный запуск по расписанию (демон)
py test_steps.py --schedule

# Пуш изменений
git add -A && git commit -m "описание" && git push
```

## Контакты
- GitHub: svrual-del/my_progect
- Telegram bot: @KaspiReporterBot (ID: 8510234232)
- Telegram group: -1002158563605
- Google Sheets: https://docs.google.com/spreadsheets/d/16NoTXUjutOw_anh_oSuufYEEfEu6FiHGm9OFnrSkdN8
