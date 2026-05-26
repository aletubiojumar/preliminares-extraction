"""Bot de Telegram con botón persistente para enviar el Excel de estadísticas."""

from __future__ import annotations

import asyncio
import logging
import subprocess
import sys
from pathlib import Path

from telegram import KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

TOKEN = "8525851840:AAGnVAKoN290n-C9Sb_eg6OKC2CLfAd4Pk4"
EXCEL_PATH = Path.home() / "preliminares-extraction" / "estadisticas_preliminares.xlsx"
SCRIPT_EXPORTAR = Path(__file__).resolve().parent / "exportar_en_curso_epac.py"
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

_generando = False


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "¡Listo! Pulsa el botón para recibir el Excel.",
        reply_markup=_TECLADO,
    )


async def enviar_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    global _generando

    if _generando:
        await update.message.reply_text("⏳ Ya hay una generación en curso, espera un momento.")
        return

    _generando = True

    await update.message.reply_text(
        "⏳ Generando Excel, espera un momento...",
        reply_markup=ReplyKeyboardRemove(),
    )

    resultado = await asyncio.to_thread(
        subprocess.run,
        [sys.executable, str(SCRIPT_EXPORTAR), "--headless"],
        capture_output=True,
        text=True,
    )

    if resultado.returncode != 0 or not EXCEL_PATH.exists():
        detalle = resultado.stderr[-500:].strip() if resultado.stderr else "Sin detalles."
        await update.message.reply_text(
            f"❌ Error al generar el Excel.\n<code>{detalle}</code>",
            parse_mode="HTML",
            reply_markup=_TECLADO,
        )
        _generando = False
        return

    await update.message.reply_document(
        document=EXCEL_PATH,
        filename=EXCEL_PATH.name,
        caption="📊 Estadísticas de Preliminares",
        reply_markup=_TECLADO,
    )

    _generando = False


def main() -> None:
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Text([BTN_EXCEL]), enviar_excel))
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
