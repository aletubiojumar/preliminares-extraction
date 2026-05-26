"""Bot de Telegram con botón persistente para enviar el Excel de estadísticas."""

from __future__ import annotations

import logging
from pathlib import Path

from telegram import KeyboardButton, ReplyKeyboardMarkup, Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

TOKEN = "8525851840:AAGnVAKoN290n-C9Sb_eg6OKC2CLfAd4Pk4"
EXCEL_PATH = Path.home() / "preliminares-extraction" / "estadisticas_preliminares.xlsx"
BTN_EXCEL = "📊 Excel Estadísticas de Preliminares"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

_TECLADO = ReplyKeyboardMarkup(
    [[KeyboardButton(BTN_EXCEL)]],
    resize_keyboard=True,
    is_persistent=True,
)


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "¡Listo! Pulsa el botón para recibir el Excel.",
        reply_markup=_TECLADO,
    )


async def enviar_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not EXCEL_PATH.exists():
        await update.message.reply_text(
            "❌ El fichero no existe todavía. Ejecuta primero exportar_en_curso_epac.py."
        )
        return
    await update.message.reply_document(
        document=EXCEL_PATH,
        filename=EXCEL_PATH.name,
        caption="📊 Estadísticas de Preliminares",
    )


def main() -> None:
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Text([BTN_EXCEL]), enviar_excel))
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
