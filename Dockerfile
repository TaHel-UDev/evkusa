FROM python:3.11-slim

WORKDIR /app

# Установка системных зависимостей (если потребуются дополнительные пакеты, добавьте их здесь)
# RUN apt-get update && apt-get install -y <пакеты> && rm -rf /var/lib/apt/lists/*

# Копирование файлов зависимостей
COPY requirements.txt .

# Установка Python зависимостей
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Копирование исходного кода
COPY . .

# Создание рабочей директории
RUN mkdir -p work

# Healthcheck для проверки работоспособности контейнера
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD python -c "import sys; sys.exit(0)"

# Запуск бота
CMD ["python", "ev_bot.py"]

