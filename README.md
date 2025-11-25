# EVKusa Presentation Bot

Telegram-бот, который по Excel-файлу с мастер-меню и изображению фона генерирует PowerPoint-презентацию с блюдам и категориями.

## Возможности

- Команда `/evkusa` запускает сценарий подготовки презентации.
- Бот просит прислать:
  1. Excel-файл с мастер-меню (`.xlsx` или `.xlsm`)
  2. Изображение фона (любое изображение, будет использовано как фон слайдов)
- После получения обоих файлов бот:
  - показывает статус: «Идет подготовка презентации...✨»
  - генерирует презентацию `КП <Название_из_B3>.pptx`
  - отправляет сообщение «✨ Презентация готова!»
  - присылает готовый `.pptx` в чат
  - очищает рабочую папку `work/`

Без команды `/evkusa` бот файлы не ожидает и презентацию не формирует.

---

## Требования

- Python 3.8+ (рекомендовано)
- Telegram Bot API токен (от @BotFather)

**Для Coolify:** дополнительные зависимости не требуются, всё работает в Docker-контейнере.

**Для ручной установки:**
- `git`
- `virtualenv` (или модуль `venv`)
- `supervisor` (опционально, только для автозапуска на сервере без Docker)

---

## Установка

### Вариант 1: Запуск через Coolify (рекомендуется)

Coolify автоматически соберет Docker-образ и запустит бота.

#### Шаги:

1. **Подключите репозиторий в Coolify**
   - Добавьте новый ресурс (Resource) в Coolify
   - Выберите тип "Git Repository"
   - Укажите URL вашего репозитория: `https://github.com/Baslykden/Evkusa.git`
   - Выберите ветку (обычно `main`)

2. **Настройте переменные окружения**
   - В настройках приложения в Coolify найдите раздел "Environment Variables"
   - Добавьте переменную:
     - **Имя**: `BOT_TOKEN`
     - **Значение**: ваш токен от @BotFather (например: `1234567890:ABCdefGHIjklMNOpqrsTUVwxyz`)

3. **Запустите приложение**
   - Coolify автоматически соберет Docker-образ из `Dockerfile`
   - Бот запустится и будет работать в контейнере
   - Логи можно просматривать в интерфейсе Coolify

4. **Проверка работы**
   - Найдите бота в Telegram по имени, которое вы указали при создании (у BotFather)
   - Отправьте команду: `/evkusa`

**ГОТОВО!** Бот работает в Coolify.

---

### Вариант 2: Ручная установка (для локальной разработки или сервера без Coolify)

#### 1. Клонирование репозитория

```bash
cd /opt
git clone https://github.com/Baslykden/Evkusa.git
cd Evkusa
```

#### 2. Виртуальное окружение и зависимости

```bash
python3 -m venv venv
source venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt
```

#### 3. Настройка config.py

Откройте файл `config.py` и:

1. Вставьте токен бота:

```python
BOT_TOKEN = "ВАШ_TELEGRAM_BOT_TOKEN"
```

**Или** используйте переменную окружения:
```bash
export BOT_TOKEN="ВАШ_TELEGRAM_BOT_TOKEN"
```

#### 4. Создание папки для временных файлов

```bash
mkdir -p work
```

#### 5. Тестовый запуск бота вручную

Из корня проекта:
```bash
source venv/bin/activate
python ev_bot.py
```

Если бот запустился без ошибок, можно настроить автозапуск через supervisor (опционально).

#### 6. Запуск под supervisor (опционально, только для продакшн-сервера без Docker)

> **Примечание:** Если вы используете Coolify, этот шаг не нужен — Coolify сам управляет перезапусками.

**6.1. Установка supervisor (если еще не установлен)**

```bash
sudo apt-get update
sudo apt-get install supervisor
```

**6.2. Конфиг supervisor**

Создайте файл, например: `/etc/supervisor/conf.d/ev_bot.conf`:

```ini
[program:Evkusa_bot]
directory=/opt/Evkusa
command=/opt/Evkusa/venv/bin/python ev_bot.py
autostart=true
autorestart=true
stderr_logfile=/var/log/ev_bot.err.log
stdout_logfile=/var/log/ev_bot.out.log
user=root
stopsignal=TERM
```

**6.3. Применение настроек supervisor**

После создания конфига:
```bash
sudo supervisorctl reread
sudo supervisorctl update
sudo supervisorctl start Evkusa_bot
```

Проверить статус:
```bash
sudo supervisorctl status Evkusa_bot
```

#### 7. Старт Бота

Найдите бота в Telegram по имени, которое вы указали при создании (у BotFather).

Отправьте команду:
```
/evkusa
```

**ГОТОВО!**

