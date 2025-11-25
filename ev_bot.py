from pathlib import Path
import shutil  # ← добавили

from openpyxl import load_workbook
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

from config import BOT_TOKEN, BASE_DIR
from ev_pptx import build_presentation

WORK_DIR = BASE_DIR / "work"
WORK_DIR.mkdir(parents=True, exist_ok=True)

SESSIONS = {}


def cleanup_work_dir():
    if not WORK_DIR.exists():
        return
    for item in WORK_DIR.iterdir():
        try:
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()
        except Exception:
            pass


def sanitize_filename(name: str) -> str:
    name = (name or "").strip()
    if not name:
        name = "Банкет"
    bad = '\\/:*?"<>|'
    for ch in bad:
        name = name.replace(ch, "_")
    return name


async def cmd_evkusa(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id

    SESSIONS[chat_id] = {
        "excel": None,
        "bg": None,
        "msg_id": None,
    }

    text = (
        "👋Здравствуйте!\n"
        "✨Я готов подготовить презентацию для вашего мероприятия.\n\n"
        "<b><u>Пришлите, пожалуйста, 2 файла:</u></b>\n"
        "1️⃣ Excel файл с Мастер меню\n"
        "2️⃣ Изображение фона для презентаций: image.png\n"
    )

    sent = await update.message.reply_text(text, parse_mode="HTML")
    SESSIONS[chat_id]["msg_id"] = sent.message_id


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await cmd_evkusa(update, context)


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message:
        return

    chat_id = message.chat_id
    session = SESSIONS.get(chat_id)
    if not session:
        return

    doc = message.document
    if not doc:
        return

    filename = (doc.file_name or "").lower()
    chat_dir = WORK_DIR / str(chat_id)
    chat_dir.mkdir(parents=True, exist_ok=True)

    if filename.endswith((".xlsx", ".xlsm")):
        file = await doc.get_file()
        local_path = chat_dir / "menu.xlsx"
        await file.download_to_drive(local_path.as_posix())
        session["excel"] = local_path
    elif filename.endswith((".png", ".jpg", ".jpeg", ".bmp")):
        file = await doc.get_file()
        ext = Path(filename).suffix or ".png"
        local_path = chat_dir / f"background{ext}"
        await file.download_to_drive(local_path.as_posix())
        session["bg"] = local_path
    else:
        await message.reply_text("Я принимаю только Excel (.xlsx/.xlsm) и изображения.")
        return

    await maybe_run_generation(update, context, chat_id)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message:
        return

    chat_id = message.chat_id
    session = SESSIONS.get(chat_id)
    if not session:
        return

    if not message.photo:
        return

    photo = message.photo[-1]
    chat_dir = WORK_DIR / str(chat_id)
    chat_dir.mkdir(parents=True, exist_ok=True)

    file = await photo.get_file()
    local_path = chat_dir / "background.jpg"
    await file.download_to_drive(local_path.as_posix())
    session["bg"] = local_path

    await maybe_run_generation(update, context, chat_id)


async def maybe_run_generation(update: Update, context: ContextTypes.DEFAULT_TYPE, chat_id: int):
    session = SESSIONS.get(chat_id)
    if not session:
        return

    excel_path = session.get("excel")
    bg_path = session.get("bg")

    if not excel_path or not bg_path:
        return

    msg_id = session.get("msg_id")
    if msg_id:
        try:
            await context.bot.edit_message_text(
                chat_id=chat_id,
                message_id=msg_id,
                text="Идет подготовка презентации...✨",
            )
        except Exception:
            pass

    chat_dir = excel_path.parent

    try:
        wb = load_workbook(excel_path, data_only=True)
        sheet1 = wb.worksheets[0]
        raw = sheet1["B3"].value
        event_name = str(raw).strip() if raw else "Фуршет"
    except Exception:
        event_name = "Фуршет"

    file_name_base = "КП " + sanitize_filename(event_name)
    out_path = chat_dir / (file_name_base + ".pptx")

    try:
        build_presentation(excel_path, bg_path, out_path)
    except Exception:
        await update.message.reply_text("Не получилось собрать презентацию. Проверьте файлы и попробуйте ещё раз.")
        SESSIONS.pop(chat_id, None)
        return

    # удаляем сообщение "Идет подготовка презентации...✨"
    if msg_id:
        try:
            await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except Exception:
            pass

    await context.bot.send_message(
        chat_id=chat_id,
        text="✨ Презентация готова!\n\n<i>Если еще потребуется моя помощь, отправьте команду: /pp</i>",
        parse_mode="HTML",
    )

    with out_path.open("rb") as f:
        await context.bot.send_document(chat_id=chat_id, document=f, filename=out_path.name)

    # чистим всю папку work после отправки
    cleanup_work_dir()

    SESSIONS.pop(chat_id, None)


def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("pp", cmd_evkusa))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))

    app.run_polling()


if __name__ == "__main__":
    main()
