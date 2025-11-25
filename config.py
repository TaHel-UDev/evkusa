# config.py
import os
from pathlib import Path

# Токен бота - берется из переменной окружения или из файла
# Для Coolify используйте переменную окружения BOT_TOKEN
BOT_TOKEN = os.getenv("BOT_TOKEN", "ВАШ_TELEGRAM_BOT_TOKEN")

# Базовая директория проекта
# По умолчанию — папка, где лежит сам config.py
BASE_DIR = Path(__file__).resolve().parent
