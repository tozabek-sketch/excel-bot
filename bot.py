"""
Rasm → Excel Telegram Bot
Qo'llab-quvvatlanadigan formatlar: HEIC, JPG, JPEG, PNG, BMP, WEBP, TIFF
"""

import os
import io
import logging
from pathlib import Path
from telegram import Update, Document
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    filters, ContextTypes
)
from PIL import Image
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# HEIC support (ixtiyoriy)
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    HEIC_SUPPORTED = False

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# .env yoki muhit o'zgaruvchisidan token olish
# ─────────────────────────────────────────────
TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8951279462:AAEDjR6dMqi0BFBt5sSuUxmIdczCSP31u90")

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff", ".tif", ".heic", ".heif"}

MAX_DIMENSION = 200          # Excel uchun max piksel o'lchami (resize)
CELL_SIZE_PX  = 6            # Har bir katak necha pikselni ifodalaydi
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB


# ─────────────────────────────────────────────
# Yordamchi funksiyalar
# ─────────────────────────────────────────────

def image_bytes_to_excel(img_bytes: bytes, filename: str = "image") -> bytes:
    """PIL Image → piksel-art Excel fayli (in-memory)"""
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")

    # Katta rasmlarni kichraytirish
    w, h = img.size
    if w > MAX_DIMENSION or h > MAX_DIMENSION:
        ratio = min(MAX_DIMENSION / w, MAX_DIMENSION / h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

    width, height = img.size
    pixels = img.load()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rasm"

    # Katak o'lchamlarini sozlash
    col_width  = CELL_SIZE_PX * 0.14   # openpyxl width units (taxminan)
    row_height = CELL_SIZE_PX * 0.75   # points

    for col in range(1, width + 1):
        ws.column_dimensions[get_column_letter(col)].width = col_width

    for row in range(1, height + 1):
        ws.row_dimensions[row].height = row_height

    # Piksellarni rang sifatida yozish
    for y in range(height):
        for x in range(width):
            r, g, b = pixels[x, y]
            hex_color = f"{r:02X}{g:02X}{b:02X}"
            cell = ws.cell(row=y + 1, column=x + 1)
            cell.fill = PatternFill(fill_type="solid", fgColor=hex_color)

    # Metadata varag'i
    meta = wb.create_sheet("Ma'lumot")
    meta["A1"] = "Fayl nomi"
    meta["B1"] = filename
    meta["A2"] = "O'lcham (px)"
    meta["B2"] = f"{width} × {height}"
    meta["A3"] = "Bot"
    meta["B3"] = "Rasm→Excel Bot"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# Handler'lar
# ─────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    heic_note = "✅ HEIC/HEIF" if HEIC_SUPPORTED else "⚠️ HEIC (kutubxona o'rnatilmagan)"
    await update.message.reply_text(
        "👋 *Rasm → Excel Botga xush kelibsiz!*\n\n"
        "Menga quyidagi formatdagi rasm yuboring:\n"
        f"📷 JPG / JPEG\n📷 PNG\n📷 {heic_note}\n"
        "📷 BMP / WEBP / TIFF\n\n"
        "Men uni *piksel-art Excel fayli* (.xlsx) sifatida qaytaraman!\n\n"
        "📌 Rasm 200×200 pikselgacha siqiladi (sifat saqlanadi).",
        parse_mode="Markdown"
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "ℹ️ *Yordam*\n\n"
        "1️⃣ Rasmni bot chatiga yuboring (foto yoki fayl sifatida)\n"
        "2️⃣ Bot rasmni qayta ishlaydi\n"
        "3️⃣ Excel (.xlsx) fayli qaytariladi\n\n"
        "⚡️ Katta rasmlar avtomatik kichraytiriladi.\n"
        "📏 Maksimal fayl hajmi: 20 MB",
        parse_mode="Markdown"
    )


async def process_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Telegram foto (siqilgan) → Excel"""
    msg = await update.message.reply_text("⏳ Qayta ishlanmoqda...")
    photo = update.message.photo[-1]  # Eng yuqori sifat
    file = await context.bot.get_file(photo.file_id)
    img_bytes = await file.download_as_bytearray()

    try:
        excel_bytes = image_bytes_to_excel(bytes(img_bytes), "telegram_photo")
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=io.BytesIO(excel_bytes),
            filename="rasm.xlsx",
            caption="✅ Tayyor! Excel faylini yuklab oling."
        )
    except Exception as e:
        logger.error(f"Photo error: {e}")
        await update.message.reply_text(f"❌ Xatolik: {e}")
    finally:
        await msg.delete()


async def process_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fayl (HEIC, PNG, JPG va h.k.) → Excel"""
    doc: Document = update.message.document
    filename = doc.file_name or "file"
    ext = Path(filename).suffix.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        await update.message.reply_text(
            f"⚠️ *{ext}* formati qo'llab-quvvatlanmaydi.\n"
            "Iltimos JPG, PNG, HEIC, BMP, WEBP yoki TIFF yuboring.",
            parse_mode="Markdown"
        )
        return

    if not HEIC_SUPPORTED and ext in {".heic", ".heif"}:
        await update.message.reply_text(
            "⚠️ HEIC formati hozir mavjud emas.\n"
            "Serverda `pillow-heif` kutubxonasi o'rnatilmagan.\n"
            "Rasmni JPG yoki PNG formatiga o'tkazib yuboring."
        )
        return

    if doc.file_size and doc.file_size > MAX_FILE_SIZE:
        await update.message.reply_text("❌ Fayl juda katta (maks. 20 MB).")
        return

    msg = await update.message.reply_text(f"⏳ *{filename}* qayta ishlanmoqda...", parse_mode="Markdown")

    try:
        file = await context.bot.get_file(doc.file_id)
        img_bytes = await file.download_as_bytearray()
        excel_bytes = image_bytes_to_excel(bytes(img_bytes), Path(filename).stem)
        out_name = Path(filename).stem + ".xlsx"

        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=io.BytesIO(excel_bytes),
            filename=out_name,
            caption=f"✅ *{filename}* → *{out_name}* tayyor!",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(f"Document error: {e}")
        await update.message.reply_text(f"❌ Xatolik: {e}")
    finally:
        await msg.delete()


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(MessageHandler(filters.PHOTO, process_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, process_document))

    logger.info("Bot ishga tushdi...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
