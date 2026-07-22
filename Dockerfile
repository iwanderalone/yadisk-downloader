# Используем легкий официальный образ Python на базе Debian Bookworm
FROM python:3.11-slim

# Настройка часового пояса и кодировки
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    TZ=Europe/Moscow

WORKDIR /app

# Установка системных зависимостей
RUN apt-get update && apt-get install -y --no-install-recommends \
    tzdata \
    && rm -rf /var/lib/apt/lists/*

# Копируем requirements и устанавливаем библиотеки Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем исходный код приложения
COPY yandex_shared_disk.py bot_telegram.py bot.py ./
COPY .env.example .env.example

# Копируем .env если он существует локально (или монтируем через docker-compose)
COPY .env* ./

# Команда по умолчанию — запуск Telegram-бота
CMD ["python", "bot_telegram.py"]
