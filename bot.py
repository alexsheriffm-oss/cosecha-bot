import os
import json
import logging
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler, ContextTypes
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ── Configuración ──────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN   = os.environ["BOT_TOKEN"]
SHEET_ID    = os.environ["SHEET_ID"]
TZ          = ZoneInfo("America/La_Paz")

# ── Google Sheets ──────────────────────────────────────────────────────────────
def get_sheet():
    creds_json = os.environ["GOOGLE_CREDS_JSON"]
    creds_dict = json.loads(creds_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("Registros")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Registros", rows=5000, cols=6)
        ws.append_row(["user_id", "nombre", "fecha", "tipo", "hora", "horas_dia"])
    return ws

# ── Helpers ────────────────────────────────────────────────────────────────────
def now_lp():
    return datetime.now(TZ)

def fmt_date(dt): return dt.strftime("%d/%m/%Y")
def fmt_time(dt): return dt.strftime("%H:%M")
def fmt_dur(secs):
    h = int(secs // 3600)
    m = int((secs % 3600) // 60)
    return f"{h}h {m:02d}min"

def get_user_rows(ws, user_id, fecha=None):
    all_rows = ws.get_all_records()
    return [
        r for r in all_rows
        if str(r["user_id"]) == str(user_id)
        and (fecha is None or r["fecha"] == fecha)
    ]

def calc_hours_today(rows):
    """Calcula horas trabajadas hoy sumando pares entrada/salida."""
    entradas = [r for r in rows if r["tipo"] == "ENTRO"]
    salidas  = [r for r in rows if r["tipo"] == "SALGO"]
    total = 0.0
    for i, e in enumerate(entradas):
        t_e = datetime.strptime(e["hora"], "%H:%M").replace(tzinfo=TZ)
        if i < len(salidas):
            t_s = datetime.strptime(salidas[i]["hora"], "%H:%M").replace(tzinfo=TZ)
            total += (t_s - t_e).total_seconds()
        else:
            # Sigue trabajando
            t_now = now_lp().replace(second=0, microsecond=0)
            total += (t_now - t_e).total_seconds()
    return total

def is_working(rows):
    entradas = [r for r in rows if r["tipo"] == "ENTRO"]
    salidas  = [r for r in rows if r["tipo"] == "SALGO"]
    return len(entradas) > len(salidas)

def build_keyboard(working: bool):
    if working:
        btn = InlineKeyboardButton("🔴 SALGO", callback_data="salgo")
        return InlineKeyboardMarkup([[btn]])
    else:
        btn = InlineKeyboardButton("✅ ENTRO", callback_data="entro")
        return InlineKeyboardMarkup([[btn]])

def build_message(nombre, horas_seg, working, fecha_str):
    estado = "🟢 Trabajando ahora" if working else "⚪ Fuera de turno"
    horas_txt = fmt_dur(horas_seg) if horas_seg > 0 else "0h 00min"
    return (
        f"🌱 *COSECHA COLECTIVA*\n"
        f"━━━━━━━━━━━━━━━━\n"
        f"👤 {nombre}\n"
        f"📅 {fecha_str}\n\n"
        f"⏱ Hoy trabajaste: *{horas_txt}*\n"
        f"Estado: {estado}\n"
        f"━━━━━━━━━━━━━━━━"
    )

# ── Comando /start ─────────────────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user    = update.effective_user
    now     = now_lp()
    fecha   = fmt_date(now)
    ws      = get_sheet()
    rows    = get_user_rows(ws, user.id, fecha)
    horas   = calc_hours_today(rows)
    working = is_working(rows)
    nombre  = user.full_name

    msg = build_message(nombre, horas, working, fecha)
    kb  = build_keyboard(working)
    await update.message.reply_text(msg, parse_mode="Markdown", reply_markup=kb)

# ── Callbacks ENTRO / SALGO ────────────────────────────────────────────────────
async def button_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query  = update.callback_query
    await query.answer()
    user   = query.from_user
    accion = query.data          # "entro" o "salgo"
    now    = now_lp()
    fecha  = fmt_date(now)
    hora   = fmt_time(now)
    ws     = get_sheet()
    rows   = get_user_rows(ws, user.id, fecha)
    working = is_working(rows)

    # Validar coherencia
    if accion == "entro" and working:
        await query.answer("⚠️ Ya estás trabajando. Primero marca SALGO.", show_alert=True)
        return
    if accion == "salgo" and not working:
        await query.answer("⚠️ No tienes una entrada activa.", show_alert=True)
        return

    tipo = "ENTRO" if accion == "entro" else "SALGO"
    ws.append_row([str(user.id), user.full_name, fecha, tipo, hora, ""])

    # Recalcular
    rows    = get_user_rows(ws, user.id, fecha)
    horas   = calc_hours_today(rows)
    working = is_working(rows)

    # Actualizar horas_dia en la última fila de SALGO
    if tipo == "SALGO":
        all_data = ws.get_all_values()
        for i in range(len(all_data) - 1, 0, -1):
            if all_data[i][0] == str(user.id) and all_data[i][3] == "SALGO" and all_data[i][5] == "":
                ws.update_cell(i + 1, 6, fmt_dur(horas))
                break

    confirmacion = "✅ ¡Entrada registrada!" if tipo == "ENTRO" else "🔴 ¡Salida registrada!"
    msg = build_message(user.full_name, horas, working, fecha)
    kb  = build_keyboard(working)

    await query.edit_message_text(
        f"{confirmacion}\n\n{msg}",
        parse_mode="Markdown",
        reply_markup=kb
    )

# ── Comando /reporte ───────────────────────────────────────────────────────────
async def reporte(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Genera reporte Excel del mes. Uso: /reporte o /reporte 05/2026"""
    now = now_lp()
    if ctx.args:
        try:
            mes, anio = ctx.args[0].split("/")
            mes, anio = int(mes), int(anio)
        except:
            await update.message.reply_text("❌ Formato: /reporte MM/AAAA  (ej: /reporte 05/2026)")
            return
    else:
        mes, anio = now.month, now.year

    await update.message.reply_text(f"⏳ Generando reporte {mes:02d}/{anio}...")

    ws       = get_sheet()
    all_rows = ws.get_all_records()
    mes_str  = f"{mes:02d}"
    anio_str = str(anio)

    # Filtrar registros del mes
    filtrados = [
        r for r in all_rows
        if r["fecha"] and r["fecha"][3:5] == mes_str and r["fecha"][6:] == anio_str
    ]

    if not filtrados:
        await update.message.reply_text(f"📭 No hay registros para {mes:02d}/{anio}.")
        return

    # Calcular horas por usuario/día
    from collections import defaultdict
    resumen = defaultdict(lambda: defaultdict(float))  # {nombre: {fecha: segundos}}
    nombres = {}

    usuarios = {}
    for r in filtrados:
        uid = str(r["user_id"])
        if uid not in usuarios:
            usuarios[uid] = []
        usuarios[uid].append(r)
        nombres[uid] = r["nombre"]

    for uid, rows in usuarios.items():
        fechas = set(r["fecha"] for r in rows)
        for fecha in fechas:
            rows_dia = [r for r in rows if r["fecha"] == fecha]
            secs = calc_hours_today(rows_dia)
            resumen[nombres[uid]][fecha] += secs

    # ── Crear Excel ──────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws_xl = wb.active
    ws_xl.title = f"Reporte {mes:02d}-{anio}"

    verde_oscuro = "1B4332"
    verde_medio  = "40916C"
    verde_claro  = "D8F3DC"
    blanco       = "FFFFFF"
    gris         = "F8F9FA"

    # Título
    ws_xl.merge_cells("A1:E1")
    titulo = ws_xl["A1"]
    titulo.value = f"🌱 COSECHA COLECTIVA — Horas trabajadas {mes:02d}/{anio}"
    titulo.font      = Font(bold=True, size=14, color=blanco)
    titulo.fill      = PatternFill("solid", fgColor=verde_oscuro)
    titulo.alignment = Alignment(horizontal="center", vertical="center")
    ws_xl.row_dimensions[1].height = 30

    # Encabezados
    headers = ["Nombre", "Fecha", "Horas trabajadas", "En minutos", ""]
    for col, h in enumerate(headers, 1):
        cell = ws_xl.cell(row=2, column=col, value=h)
        cell.font      = Font(bold=True, color=blanco, size=11)
        cell.fill      = PatternFill("solid", fgColor=verde_medio)
        cell.alignment = Alignment(horizontal="center")

    # Datos
    fila = 3
    totales_persona = {}
    for nombre in sorted(resumen.keys()):
        total_persona = 0.0
        for fecha in sorted(resumen[nombre].keys()):
            secs = resumen[nombre][fecha]
            total_persona += secs
            h = int(secs // 3600)
            m = int((secs % 3600) // 60)
            mins = int(secs // 60)

            bg = blanco if fila % 2 == 0 else gris
            datos = [nombre, fecha, f"{h}h {m:02d}min", mins, ""]
            for col, val in enumerate(datos, 1):
                cell = ws_xl.cell(row=fila, column=col, value=val)
                cell.fill      = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="center" if col > 1 else "left")
            fila += 1

        totales_persona[nombre] = total_persona

        # Subtotal por persona
        h = int(total_persona // 3600)
        m = int((total_persona % 3600) // 60)
        ws_xl.merge_cells(f"A{fila}:B{fila}")
        cell_sub = ws_xl.cell(row=fila, column=1, value=f"TOTAL {nombre.upper()}")
        cell_sub.font      = Font(bold=True, color=blanco)
        cell_sub.fill      = PatternFill("solid", fgColor=verde_medio)
        cell_sub.alignment = Alignment(horizontal="center")
        cell_tot = ws_xl.cell(row=fila, column=3, value=f"{h}h {m:02d}min")
        cell_tot.font      = Font(bold=True, color=blanco)
        cell_tot.fill      = PatternFill("solid", fgColor=verde_medio)
        cell_tot.alignment = Alignment(horizontal="center")
        ws_xl.cell(row=fila, column=4, value=int(total_persona//60)).fill = PatternFill("solid", fgColor=verde_medio)
        fila += 2

    # Gran total
    gran_total = sum(totales_persona.values())
    h = int(gran_total // 3600)
    m = int((gran_total % 3600) // 60)
    ws_xl.merge_cells(f"A{fila}:B{fila}")
    cell_gt = ws_xl.cell(row=fila, column=1, value="TOTAL EQUIPO")
    cell_gt.font      = Font(bold=True, size=12, color=blanco)
    cell_gt.fill      = PatternFill("solid", fgColor=verde_oscuro)
    cell_gt.alignment = Alignment(horizontal="center")
    cell_gt2 = ws_xl.cell(row=fila, column=3, value=f"{h}h {m:02d}min")
    cell_gt2.font      = Font(bold=True, size=12, color=blanco)
    cell_gt2.fill      = PatternFill("solid", fgColor=verde_oscuro)
    cell_gt2.alignment = Alignment(horizontal="center")

    # Anchos de columna
    ws_xl.column_dimensions["A"].width = 28
    ws_xl.column_dimensions["B"].width = 14
    ws_xl.column_dimensions["C"].width = 18
    ws_xl.column_dimensions["D"].width = 12

    # Exportar a bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    nombre_archivo = f"Cosecha_Colectiva_Horas_{mes:02d}_{anio}.xlsx"
    await update.message.reply_document(
        document=buf,
        filename=nombre_archivo,
        caption=f"📊 Reporte de horas — {mes:02d}/{anio}\n🌱 Cosecha Colectiva"
    )

# ── Comando /hoy ───────────────────────────────────────────────────────────────
async def hoy(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user  = update.effective_user
    now   = now_lp()
    fecha = fmt_date(now)
    ws    = get_sheet()
    rows  = get_user_rows(ws, user.id, fecha)
    horas = calc_hours_today(rows)
    working = is_working(rows)

    msg = build_message(user.full_name, horas, working, fecha)
    kb  = build_keyboard(working)
    await update.message.reply_text(msg, parse_mode="Markdown", reply_markup=kb)

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("hoy",   hoy))
    app.add_handler(CommandHandler("reporte", reporte))
    app.add_handler(CallbackQueryHandler(button_handler))
    logger.info("🌱 Wason bot arrancando...")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
