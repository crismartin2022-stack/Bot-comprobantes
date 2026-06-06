import os
import json
import base64
import logging
import tempfile
from datetime import datetime
from io import BytesIO

import httpx
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ContextTypes, CallbackQueryHandler
)

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO
)
log = logging.getLogger(__name__)

# ── Config desde variables de entorno ────────────────────────────────────────
TELEGRAM_TOKEN  = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_KEY   = os.environ["ANTHROPIC_API_KEY"]
# Opcional: limitar a un grupo específico (dejar vacío para permitir todos)
ALLOWED_GROUP   = os.environ.get("ALLOWED_GROUP_ID", "")

claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

# ── Storage en memoria (persiste mientras el bot corre) ───────────────────────
# Estructura: { chat_id: { "semana_actual": "...", "registros": [...] } }
store: dict = {}

def get_store(chat_id: int) -> dict:
    cid = str(chat_id)
    if cid not in store:
        store[cid] = {
            "semana_actual": semana_label(),
            "registros": []
        }
    return store[cid]

def semana_label() -> str:
    hoy = datetime.now()
    num = (hoy.day - 1) // 7 + 1
    return f"{hoy.strftime('%d/%m/%Y')} — Semana {num}"

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
  "cvu_ultimos4": "últimos 4 dígitos del CVU/CBU del receptor — buscá el número de cuenta destino/CVU destino. Si no está visible dejá VACÍO (nunca inventar)",
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

    # Estilos
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

    # Header
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
            i - 1,
            semana,
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
            if col == 7:  # Columna CVU
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

# ── Formateo de respuesta del bot ─────────────────────────────────────────────
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

# ── Handlers ──────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Bot de Comprobantes Agilpagos*\n\n"
        "Mandame imágenes de comprobantes y las analizo automáticamente.\n\n"
        "📌 *Comandos:*\n"
        "/resumen — ver comprobantes cargados\n"
        "/excel — generar Excel semanal\n"
        "/nueva\\_semana — iniciar nueva semana\n"
        "/borrar — borrar todos los registros\n"
        "/ayuda — más información",
        parse_mode="Markdown"
    )

async def cmd_ayuda(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *Cómo usar el bot:*\n\n"
        "1️⃣ Enviá imágenes de comprobantes al grupo\n"
        "2️⃣ El bot analiza y extrae los datos automáticamente\n"
        "3️⃣ Si falta el CVU te avisa con 🔴\n"
        "4️⃣ Usá /excel para bajar el archivo Excel\n\n"
        "💡 El bot guarda todos los comprobantes de la semana.\n"
        "Al usar /nueva\\_semana limpia los registros para empezar de cero.",
        parse_mode="Markdown"
    )

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
        texto += f"\n⚠️ *Hay {sin_cvu} comprobante(s) sin CVU. Revisá antes de exportar.*"

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
        fecha = datetime.now().strftime("%Y%m%d")
        nombre = f"Comprobantes_{fecha}.xlsx"
        await update.message.reply_document(
            document=buf,
            filename=nombre,
            caption=f"📊 Excel generado — {len(registros)} comprobantes\n{datos['semana_actual']}"
        )
        await msg.delete()
    except Exception as e:
        log.error(f"Error generando Excel: {e}")
        await msg.edit_text(f"❌ Error al generar el Excel: {e}")

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
    chat_id = update.effective_chat.id
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
        store[str(chat_id)] = {"semana_actual": semana_label(), "registros": []}
        await query.edit_message_text("✅ Nueva semana iniciada. Podés empezar a cargar comprobantes.")
    elif query.data == "borrar_si":
        datos = get_store(chat_id)
        datos["registros"] = []
        await query.edit_message_text("🗑 Registros borrados.")
    elif query.data == "cancelar":
        await query.edit_message_text("❌ Cancelado.")

async def handle_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id

    # Verificar grupo permitido si está configurado
    if ALLOWED_GROUP and str(chat_id) != ALLOWED_GROUP:
        return

    msg = await update.message.reply_text("🔍 Analizando comprobante con IA...")

    try:
        # Obtener la imagen en mayor resolución
        photo = update.message.photo[-1]
        file = await ctx.bot.get_file(photo.file_id)

        async with httpx.AsyncClient() as client:
            resp = await client.get(file.file_path)
            image_bytes = resp.content

        # Analizar con Claude
        resultado = await analizar_imagen(image_bytes, "image/jpeg")

        # Guardar en store
        datos = get_store(chat_id)
        num = len(datos["registros"]) + 1
        resultado["_num"] = num
        resultado["_fecha_carga"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        datos["registros"].append(resultado)

        # Responder
        texto = formatear_resultado(resultado, num)
        cvu = (resultado.get("cvu_ultimos4") or "").strip()

        # Botón para corregir CVU si falta
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
    """Maneja imágenes enviadas como documento (sin compresión)."""
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
        resultado["_fecha_carga"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        datos["registros"].append(resultado)

        texto = formatear_resultado(resultado, num)
        await msg.edit_text(texto, parse_mode="Markdown")
    except Exception as e:
        log.error(f"Error procesando documento: {e}")
        await msg.edit_text(f"❌ Error: {e}")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start",        cmd_start))
    app.add_handler(CommandHandler("ayuda",        cmd_ayuda))
    app.add_handler(CommandHandler("resumen",      cmd_resumen))
    app.add_handler(CommandHandler("excel",        cmd_excel))
    app.add_handler(CommandHandler("nueva_semana", cmd_nueva_semana))
    app.add_handler(CommandHandler("borrar",       cmd_borrar))
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.IMAGE, handle_document))

    log.info("🤖 Bot iniciado — esperando comprobantes...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
