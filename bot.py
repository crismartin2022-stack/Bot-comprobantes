import os
import json
import base64
import logging
from datetime import datetime, timezone, timedelta
from io import BytesIO

import httpx
import anthropic
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, CallbackQueryHandler
)

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO
)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN  = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_KEY   = os.environ["ANTHROPIC_API_KEY"]
ALLOWED_GROUP   = os.environ.get("ALLOWED_GROUP_ID", "")

# Tu ID personal para recibir resúmenes privados
ADMIN_ID        = 531707598

# Zona horaria Argentina (UTC-3)
ARG_TZ = timezone(timedelta(hours=-3))

claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

# ── Storage en memoria ────────────────────────────────────────────────────────
# { chat_id: { "semana_actual": "...", "registros": [], "registros_hoy": [] } }
store: dict = {}

def get_store(chat_id: int) -> dict:
    cid = str(chat_id)
    if cid not in store:
        store[cid] = {
            "semana_actual": semana_label(),
            "registros": [],
            "registros_hoy": [],
            "chat_id": chat_id,
        }
    return store[cid]

def semana_label() -> str:
    ahora = datetime.now(ARG_TZ)
    return f"Semana {ahora.strftime('%d/%m/%Y')}"

def now_arg() -> datetime:
    return datetime.now(ARG_TZ)

# ── Análisis con Claude ───────────────────────────────────────────────────────
SYSTEM_PROMPT = """Eres un asistente experto en análisis de comprobantes bancarios argentinos.
Analizá la imagen y respondé ÚNICAMENTE con un JSON válido sin backticks ni markdown.
El JSON debe tener exactamente estos campos:
{
  "fecha": "DD/MM/YYYY",
  "hora": "HH:MM o vacío",
  "tipo": "TRF / DEP / PAGO / otro",
  "monto": número sin símbolos (ej: 15000.00),
  "moneda": "ARS / USD / otro",
  "remitente": "nombre completo del que envía o vacío",
  "destinatario": "nombre completo del que recibe o vacío",
  "banco_origen": "banco origen o vacío",
  "banco_destino": "banco destino o vacío",
  "referencia": "número de referencia o vacío",
  "concepto": "concepto o vacío",
  "estado": "Exitoso / Pendiente / Rechazado",
  "cvu_ultimos4": "últimos 4 dígitos del CVU/CBU del receptor. Si no está visible dejá VACÍO (nunca inventar)",
  "notas": "cualquier dato relevante adicional"
}
IMPORTANTE sobre cvu_ultimos4: Buscá el CVU o CBU del destinatario/receptor.
Tomá SOLO los últimos 4 dígitos numéricos. Si no existe CVU/CBU visible, dejá "".
"""

async def analizar_imagen(image_bytes: bytes, mime: str) -> dict:
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    resp = claude.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64}},
                {"type": "text", "text": "Analizá este comprobante bancario argentino y extraé todos los datos, especialmente los últimos 4 dígitos del CVU/CBU del receptor."}
            ]
        }]
    )
    text = resp.content[0].text
    try:
        return json.loads(text.replace("```json", "").replace("```", "").strip())
    except Exception:
        return {"notas": text, "estado": "Error al parsear", "cvu_ultimos4": ""}

# ── Generador de Excel ────────────────────────────────────────────────────────
def generar_excel(registros: list, semana: str) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Comprobantes"

    header_fill = PatternFill("solid", fgColor="4B0082")
    error_fill  = PatternFill("solid", fgColor="FF4444")
    ok_fill     = PatternFill("solid", fgColor="1A5C2A")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    headers = [
        "#", "FECHA DE ENVIO", "TRF O DEPOSITO", "TITULAR DE LA CTA",
        "FECHA TICKET", "HORA TICKET", "CUENTA (CVU)", "MONTO",
        "Remitente", "Banco Origen", "Estado", "Notas"
    ]
    col_widths = [4, 22, 12, 30, 13, 11, 14, 14, 25, 18, 11, 28]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 20

    sin_cvu = 0
    for i, r in enumerate(registros, 2):
        cvu = (r.get("cvu_ultimos4") or "").strip()
        tiene_cvu = bool(cvu)
        if not tiene_cvu:
            sin_cvu += 1

        fila = [
            i - 1, semana,
            r.get("tipo", "TRF"),
            r.get("destinatario", ""),
            r.get("fecha", ""),
            r.get("hora", ""),
            cvu if tiene_cvu else "⚠️ SIN CVU",
            r.get("monto", ""),
            r.get("remitente", ""),
            r.get("banco_origen", ""),
            r.get("estado", ""),
            r.get("notas", ""),
        ]
        for col, val in enumerate(fila, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if col == 7:
                cell.fill = ok_fill if tiene_cvu else error_fill
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # Hoja resumen
    ws2 = wb.create_sheet("Resumen")
    ws2["A1"] = "Semana"
    ws2["B1"] = "Total comprobantes"
    ws2["C1"] = "Con CVU OK"
    ws2["D1"] = "Sin CVU ⚠️"
    ws2["E1"] = "Monto total ARS"
    for cell in ws2[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill

    total = sum(float(r.get("monto") or 0) for r in registros)
    con_cvu = len(registros) - sin_cvu
    ws2.append([semana, len(registros), con_cvu, sin_cvu, total])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Formateo de respuesta ─────────────────────────────────────────────────────
def formatear_resultado(r: dict, num: int) -> str:
    cvu = (r.get("cvu_ultimos4") or "").strip()
    cvu_line = f"🏦 CVU (últimos 4): `****{cvu}`" if cvu else "🔴 *CVU NO ENCONTRADO — revisar manualmente*"
    monto = r.get("monto")
    monto_fmt = f"${float(monto):,.0f}" if monto else "—"
    return (
        f"✅ *Comprobante #{num} procesado*\n"
        f"─────────────────────\n"
        f"📅 Fecha: {r.get('fecha', '—')}\n"
        f"⏰ Hora: {r.get('hora', '—')}\n"
        f"💸 Tipo: {r.get('tipo', '—')}\n"
        f"💰 Monto: {monto_fmt} {r.get('moneda', 'ARS')}\n"
        f"👤 Remitente: {r.get('remitente', '—')}\n"
        f"👤 Destinatario: {r.get('destinatario', '—')}\n"
        f"🏛 Banco origen: {r.get('banco_origen', '—')}\n"
        f"{cvu_line}\n"
        f"📋 Estado: {r.get('estado', '—')}\n"
    )

# ── Tareas programadas ────────────────────────────────────────────────────────
async def tarea_resumen_diario(app):
    """Todos los días a las 20:00 Argentina — manda resumen al admin por privado."""
    ahora = now_arg()
    fecha_hoy = ahora.strftime("%d/%m/%Y")

    # Juntar todos los registros de hoy de todos los grupos
    todos_hoy = []
    grupos = []
    for cid, datos in store.items():
        hoy = [r for r in datos.get("registros", []) if r.get("_fecha_carga", "").startswith(fecha_hoy)]
        if hoy:
            todos_hoy.extend(hoy)
            grupos.append(cid)

    if not todos_hoy:
        texto = (
            f"📊 *Resumen diario — {fecha_hoy}*\n"
            f"─────────────────────\n"
            f"📭 Sin comprobantes hoy."
        )
    else:
        total = sum(float(r.get("monto") or 0) for r in todos_hoy)
        sin_cvu = sum(1 for r in todos_hoy if not (r.get("cvu_ultimos4") or "").strip())
        con_cvu = len(todos_hoy) - sin_cvu

        # Detalle por comprobante
        detalle = ""
        for i, r in enumerate(todos_hoy, 1):
            cvu = (r.get("cvu_ultimos4") or "").strip()
            monto = r.get("monto")
            monto_fmt = f"${float(monto):,.0f}" if monto else "—"
            cvu_txt = f"****{cvu}" if cvu else "⚠️ SIN CVU"
            detalle += (
                f"\n*#{i}* {r.get('destinatario', '—')}\n"
                f"   💰 {monto_fmt} | 🏦 {cvu_txt} | ⏰ {r.get('hora', '—')}\n"
            )

        texto = (
            f"📊 *Resumen diario — {fecha_hoy}*\n"
            f"─────────────────────\n"
            f"📄 Comprobantes: {len(todos_hoy)}\n"
            f"✅ Con CVU OK: {con_cvu}\n"
            f"🔴 Sin CVU: {sin_cvu}\n"
            f"💰 Total ARS: ${total:,.0f}\n"
            f"─────────────────────"
            f"{detalle}"
        )
        if sin_cvu > 0:
            texto += f"\n⚠️ *Hay {sin_cvu} comprobante(s) sin CVU.*"

    try:
        await app.bot.send_message(
            chat_id=ADMIN_ID,
            text=texto,
            parse_mode="Markdown"
        )
        log.info(f"Resumen diario enviado a {ADMIN_ID}")
    except Exception as e:
        log.error(f"Error enviando resumen diario: {e}")

async def tarea_excel_semanal(app):
    """Todos los jueves a las 21:00 Argentina — manda Excel y reinicia semana."""
    ahora = now_arg()
    semana_cerrada = ""

    for cid, datos in store.items():
        registros = datos.get("registros", [])
        semana_cerrada = datos.get("semana_actual", semana_label())

        if registros:
            try:
                buf = generar_excel(registros, semana_cerrada)
                fecha = ahora.strftime("%Y%m%d")
                nombre = f"Comprobantes_semana_{fecha}.xlsx"

                sin_cvu = sum(1 for r in registros if not (r.get("cvu_ultimos4") or "").strip())
                total = sum(float(r.get("monto") or 0) for r in registros)

                caption = (
                    f"📊 *Excel Semanal Automático*\n"
                    f"📅 {semana_cerrada}\n"
                    f"📄 {len(registros)} comprobantes\n"
                    f"💰 Total: ${total:,.0f} ARS\n"
                    f"{'⚠️ ' + str(sin_cvu) + ' sin CVU' if sin_cvu else '✅ Todos con CVU'}"
                )

                # Mandar al grupo
                await app.bot.send_document(
                    chat_id=int(cid),
                    document=buf,
                    filename=nombre,
                    caption=caption,
                    parse_mode="Markdown"
                )

                # Mandar también al admin por privado
                buf.seek(0)
                await app.bot.send_document(
                    chat_id=ADMIN_ID,
                    document=buf,
                    filename=nombre,
                    caption=f"📎 Copia del Excel semanal\n{caption}",
                    parse_mode="Markdown"
                )

            except Exception as e:
                log.error(f"Error enviando Excel semanal al grupo {cid}: {e}")

        # Reiniciar semana
        store[cid] = {
            "semana_actual": semana_label(),
            "registros": [],
            "registros_hoy": [],
            "chat_id": int(cid),
        }

    # Notificar al admin que la semana fue reiniciada
    try:
        await app.bot.send_message(
            chat_id=ADMIN_ID,
            text=(
                f"🔄 *Nueva semana iniciada*\n"
                f"Semana cerrada: {semana_cerrada}\n"
                f"Nueva semana desde: {now_arg().strftime('%d/%m/%Y %H:%M')} hs"
            ),
            parse_mode="Markdown"
        )
    except Exception as e:
        log.error(f"Error notificando nueva semana: {e}")

    log.info("Excel semanal enviado y semana reiniciada")

# ── Handlers ──────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Bot de Comprobantes Agilpagos*\n\n"
        "Mandame imágenes de comprobantes y las analizo automáticamente.\n\n"
        "📌 *Comandos:*\n"
        "/resumen — ver comprobantes cargados\n"
        "/excel — generar Excel ahora\n"
        "/hoy — ver resumen de hoy\n"
        "/nueva\\_semana — iniciar nueva semana\n"
        "/borrar — borrar todos los registros\n\n"
        "⏰ *Automático:*\n"
        "📊 Resumen diario → 20:00 hs\n"
        "📎 Excel semanal → Jueves 21:00 hs",
        parse_mode="Markdown"
    )

async def cmd_hoy(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Resumen de comprobantes del día de hoy."""
    chat_id = update.effective_chat.id
    datos = get_store(chat_id)
    fecha_hoy = now_arg().strftime("%d/%m/%Y")
    hoy = [r for r in datos.get("registros", []) if r.get("_fecha_carga", "").startswith(fecha_hoy)]

    if not hoy:
        await update.message.reply_text(f"📭 Sin comprobantes hoy ({fecha_hoy}).")
        return

    total = sum(float(r.get("monto") or 0) for r in hoy)
    sin_cvu = sum(1 for r in hoy if not (r.get("cvu_ultimos4") or "").strip())

    detalle = ""
    for i, r in enumerate(hoy, 1):
        cvu = (r.get("cvu_ultimos4") or "").strip()
        monto = r.get("monto")
        monto_fmt = f"${float(monto):,.0f}" if monto else "—"
        cvu_txt = f"****{cvu}" if cvu else "⚠️ SIN CVU"
        detalle += f"\n*#{i}* {r.get('destinatario', '—')} | {monto_fmt} | {cvu_txt}"

    texto = (
        f"📅 *Hoy {fecha_hoy}*\n"
        f"📄 {len(hoy)} comprobantes | 💰 ${total:,.0f} ARS\n"
        f"─────────────────────"
        f"{detalle}"
    )
    await update.message.reply_text(texto, parse_mode="Markdown")

async def cmd_resumen(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    datos = get_store(chat_id)
    registros = datos["registros"]

    if not registros:
        await update.message.reply_text("📭 No hay comprobantes cargados todavía.")
        return

    total = sum(float(r.get("monto") or 0) for r in registros)
    sin_cvu = sum(1 for r in registros if not (r.get("cvu_ultimos4") or "").strip())
    con_cvu = len(registros) - sin_cvu

    texto = (
        f"📊 *Resumen — {datos['semana_actual']}*\n"
        f"─────────────────────\n"
        f"📄 Comprobantes: {len(registros)}\n"
        f"✅ Con CVU OK: {con_cvu}\n"
        f"🔴 Sin CVU: {sin_cvu}\n"
        f"💰 Total ARS: ${total:,.0f}\n"
    )
    if sin_cvu > 0:
        texto += f"\n⚠️ *{sin_cvu} comprobante(s) sin CVU. Revisá antes de exportar.*"

    await update.message.reply_text(texto, parse_mode="Markdown")

async def cmd_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    datos = get_store(chat_id)
    registros = datos["registros"]

    if not registros:
        await update.message.reply_text("📭 No hay comprobantes para exportar.")
        return

    msg = await update.message.reply_text("⏳ Generando Excel...")
    try:
        buf = generar_excel(registros, datos["semana_actual"])
        fecha = now_arg().strftime("%Y%m%d")
        nombre = f"Comprobantes_{fecha}.xlsx"
        await update.message.reply_document(
            document=buf,
            filename=nombre,
            caption=f"📊 Excel — {len(registros)} comprobantes\n{datos['semana_actual']}"
        )
        await msg.delete()
    except Exception as e:
        log.error(f"Error generando Excel: {e}")
        await msg.edit_text(f"❌ Error: {e}")

async def cmd_nueva_semana(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    kb = [[
        InlineKeyboardButton("✅ Sí, nueva semana", callback_data="nueva_semana_si"),
        InlineKeyboardButton("❌ Cancelar", callback_data="cancelar"),
    ]]
    await update.message.reply_text(
        "⚠️ ¿Querés iniciar una nueva semana?\nSe borrarán todos los comprobantes actuales.",
        reply_markup=InlineKeyboardMarkup(kb)
    )

async def cmd_borrar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    kb = [[
        InlineKeyboardButton("✅ Sí, borrar todo", callback_data="borrar_si"),
        InlineKeyboardButton("❌ Cancelar", callback_data="cancelar"),
    ]]
    await update.message.reply_text(
        "⚠️ ¿Borrar TODOS los comprobantes?",
        reply_markup=InlineKeyboardMarkup(kb)
    )

async def callback_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    chat_id = update.effective_chat.id

    if query.data == "nueva_semana_si":
        store[str(chat_id)] = {
            "semana_actual": semana_label(),
            "registros": [],
            "registros_hoy": [],
            "chat_id": chat_id,
        }
        await query.edit_message_text("✅ Nueva semana iniciada.")
    elif query.data == "borrar_si":
        datos = get_store(chat_id)
        datos["registros"] = []
        await query.edit_message_text("🗑 Registros borrados.")
    elif query.data == "cancelar":
        await query.edit_message_text("❌ Cancelado.")

async def handle_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    if ALLOWED_GROUP and str(chat_id) != ALLOWED_GROUP:
        return

    msg = await update.message.reply_text("🔍 Analizando comprobante con IA...")
    try:
        photo = update.message.photo[-1]
        file = await ctx.bot.get_file(photo.file_id)
        async with httpx.AsyncClient() as client:
            resp = await client.get(file.file_path)
            image_bytes = resp.content

        resultado = await analizar_imagen(image_bytes, "image/jpeg")
        datos = get_store(chat_id)
        num = len(datos["registros"]) + 1
        resultado["_num"] = num
        resultado["_fecha_carga"] = now_arg().strftime("%d/%m/%Y %H:%M")
        datos["registros"].append(resultado)

        texto = formatear_resultado(resultado, num)
        cvu = (resultado.get("cvu_ultimos4") or "").strip()

        kb = None
        if not cvu:
            kb = InlineKeyboardMarkup([[
                InlineKeyboardButton(
                    f"✏️ Ingresar CVU del comprobante #{num}",
                    callback_data=f"editar_cvu_{num}"
                )
            ]])

        await msg.edit_text(texto, parse_mode="Markdown", reply_markup=kb)

    except Exception as e:
        log.error(f"Error procesando imagen: {e}")
        await msg.edit_text(f"❌ Error al procesar la imagen: {e}")

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.mime_type or not doc.mime_type.startswith("image/"):
        return

    chat_id = update.effective_chat.id
    if ALLOWED_GROUP and str(chat_id) != ALLOWED_GROUP:
        return

    msg = await update.message.reply_text("🔍 Analizando comprobante (documento)...")
    try:
        file = await ctx.bot.get_file(doc.file_id)
        async with httpx.AsyncClient() as client:
            resp = await client.get(file.file_path)
            image_bytes = resp.content

        resultado = await analizar_imagen(image_bytes, doc.mime_type)
        datos = get_store(chat_id)
        num = len(datos["registros"]) + 1
        resultado["_num"] = num
        resultado["_fecha_carga"] = now_arg().strftime("%d/%m/%Y %H:%M")
        datos["registros"].append(resultado)

        texto = formatear_resultado(resultado, num)
        await msg.edit_text(texto, parse_mode="Markdown")
    except Exception as e:
        log.error(f"Error procesando documento: {e}")
        await msg.edit_text(f"❌ Error: {e}")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # Handlers
    app.add_handler(CommandHandler("start",        cmd_start))
    app.add_handler(CommandHandler("hoy",          cmd_hoy))
    app.add_handler(CommandHandler("resumen",      cmd_resumen))
    app.add_handler(CommandHandler("excel",        cmd_excel))
    app.add_handler(CommandHandler("nueva_semana", cmd_nueva_semana))
    app.add_handler(CommandHandler("borrar",       cmd_borrar))
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.IMAGE, handle_document))

    # Scheduler — zona horaria Argentina (UTC-3)
    scheduler = AsyncIOScheduler(timezone="America/Argentina/Buenos_Aires")

    # Resumen diario a las 20:00 Argentina
    scheduler.add_job(
        tarea_resumen_diario,
        CronTrigger(hour=20, minute=0, timezone="America/Argentina/Buenos_Aires"),
        args=[app],
        id="resumen_diario"
    )

    # Excel semanal todos los jueves a las 21:00 Argentina
    scheduler.add_job(
        tarea_excel_semanal,
        CronTrigger(day_of_week="thu", hour=21, minute=0, timezone="America/Argentina/Buenos_Aires"),
        args=[app],
        id="excel_semanal"
    )

    scheduler.start()
    log.info("🤖 Bot iniciado con tareas programadas:")
    log.info("   📊 Resumen diario → 20:00 hs Argentina")
    log.info("   📎 Excel semanal  → Jueves 21:00 hs Argentina")

    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
