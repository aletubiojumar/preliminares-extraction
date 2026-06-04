"""Bot de Telegram con botón persistente para enviar el Excel de estadísticas."""

from __future__ import annotations

import asyncio
import html
import logging
import os
import re
import subprocess
import sys
from pathlib import Path

from telegram import KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
EXCEL_PATH = Path("/app/output/estadisticas_preliminares.xlsx")
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


def _cargar_whitelist() -> set[int]:
    valor = os.getenv("WHITELIST_IDS", "").strip()
    if not valor:
        return set()
    ids: set[int] = set()
    for parte in re.split(r"[,\s]+", valor):
        try:
            ids.add(int(parte))
        except ValueError:
            pass
    return ids


_WHITELIST: set[int] = _cargar_whitelist()


def _autorizado(update: Update) -> bool:
    if not _WHITELIST:
        return True
    user = update.effective_user
    return user is not None and user.id in _WHITELIST


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _autorizado(update):
        await update.message.reply_text("No estás autorizado para usar este bot.")
        return
    await update.message.reply_text(
        "¡Listo! Pulsa el botón para recibir el Excel.",
        reply_markup=_TECLADO,
    )


async def enviar_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _autorizado(update):
        await update.message.reply_text("No estás autorizado para usar este bot.")
        return

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
        stderr=subprocess.PIPE,
        text=True,
    )

    if resultado.returncode != 0 or not EXCEL_PATH.exists():
        detalle = resultado.stderr[-500:].strip() if resultado.stderr else "Sin detalles."
        await update.message.reply_text(
            f"❌ Error al generar el Excel.\n<code>{html.escape(detalle)}</code>",
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
