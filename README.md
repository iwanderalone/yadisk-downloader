# Yandex Shared Disk Downloader & Telegram Bot

Компактный клиент и Telegram-бот для работы с **Общим Диском (Яндекс 360 для бизнеса)** по [документации REST API](https://yandex.ru/dev/disk-api/doc/ru/reference/content_shd).

---

## ⚡️ Принцип работы (TL;DR)

Скачивание происходит в **2 этапа**:
1. **Запрос временного URL**: Скрипт обращается к `GET https://cloud-api.yandex.net/v1/disk/virtual-disks/resources/download` с пути `vd:<vd_hash>:disk:/<path>` и OAuth-токеном в заголовке (`cloud_api:disk.read`).
2. **Скачивание файла**: Из ответа извлекается динамическая ссылка `href` (хранилище Яндекс), по которой выполняется потоковое скачивание файла.

---

## ⚙️ Конфигурация в `.env`

Создайте файл `.env`:
```env
TELEGRAM_BOT_TOKEN=your_telegram_bot_token
YANDEX_OAUTH_TOKEN=your_yandex_oauth_token
YANDEX_VD_HASH=your_yandex_shared_disk_hash
YANDEX_FILE_PATH=table.xlsx
```

*Если вы переименуете или поменяете файл на общем диске, достаточно просто изменить имя в `YANDEX_FILE_PATH` в файле `.env`!*

---

## 🐳 Развертывание в Docker (Debian / Linux) — *Рекомендуемый способ*

Для развертывания на Debian сервере в 1 команду:

1. Перенесите файлы на сервер и заполните `.env`.
2. Запустите контейнер через Docker Compose:
   ```bash
   docker compose up -d --build
   ```
3. Просмотр логов:
   ```bash
   docker compose logs -f
   ```

---

## 🧪 Запуск вручную без Docker

1. Установка зависимостей:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. Запуск Telegram-бота:
   ```bash
   python bot_telegram.py
   ```

3. Скачивание файла через CLI:
   ```bash
   python bot.py download "table.xlsx" -o "./table.xlsx"
   ```
