"""
exporter.py
Generador de reportes en Excel (.xlsx) para Admisión / Jefatura Médica,
enlaces y mensajes directos a WhatsApp (+51), y comprobantes imprimibles (HTML/PDF)
de Citas Médicas para el C.S. Su Santidad Juan Pablo II.
"""

import io
import urllib.parse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from schedule_sync import get_day_ce_consultorios, get_month_ce_overview
from holidays_pe import get_month_info


def format_peru_whatsapp_number(raw_phone: str) -> str:
    """
    Normaliza el teléfono del paciente al formato internacional de Perú (51 + 9 dígitos)
    para enlaces oficiales de WhatsApp (wa.me).
    """
    digits = "".join(ch for ch in (raw_phone or "") if ch.isdigit())
    if not digits:
        return ""
    if len(digits) == 9 and digits.startswith("9"):
        return f"51{digits}"
    if len(digits) == 11 and digits.startswith("51"):
        return digits
    return digits


def generate_whatsapp_info(appointment: dict) -> dict:
    """
    Genera el texto oficial de la constancia de cita para WhatsApp y el enlace
    directo wa.me hacia el número de celular registrado por el paciente.
    """
    turno_label = (
        "MAÑANA (08:00 a 12:00 hrs)"
        if appointment["shift_period"] == "M"
        else "TARDE (14:00 a 18:00 hrs)"
    )
    consultorio_label = (
        "Consultorio de Medicina N° 1 (Mañana)"
        if appointment["shift_period"] == "M"
        else "Consultorio de Medicina N° 2 (Tarde)"
    )

    msg = (
        f"🏥 *C.S. SU SANTIDAD JUAN PABLO II*\n"
        f"*CONSTANCIA OFICIAL DE CITA MÉDICA*\n"
        f"──────────────────────\n"
        f"🎫 *Código de Ticket:* `{appointment['ticket_code']}`\n"
        f"📅 *Fecha de Atención:* *{appointment['appointment_date']}*\n"
        f"⏰ *Horario Exacto:* *{appointment['slot_time']} hrs*\n"
        f"🔢 *Cupo Asignado:* *Cupo #{appointment['slot_number']:02d} de 16* ({turno_label})\n"
        f"🚪 *Consultorio:* {consultorio_label}\n"
        f"👨‍⚕️ *Médico Cirujano:* {appointment['doctor_name']} (CMP: {appointment['doctor_cmp']})\n"
        f"──────────────────────\n"
        f"👤 *Paciente:* {appointment['patient_names']}\n"
        f"🪪 *DNI / Doc.:* {appointment['patient_dni']} (Edad: {appointment.get('patient_age') or '-'} años)\n"
        f"🛡️ *Seguro:* {appointment.get('insurance_type', 'SIS')}\n"
        f"🩺 *Motivo:* {appointment.get('reason') or 'Consulta de Medicina General'}\n"
        f"──────────────────────\n"
        f"📌 *INDICACIONES IMPORTANTES:*\n"
        f"1. Acudir a Admisión *20 minutos antes* de su horario ({appointment['slot_time']}).\n"
        f"2. Presentar su *DNI físico* y mostrar este mensaje o su ticket en ventanilla."
    )

    wa_phone = format_peru_whatsapp_number(appointment.get("patient_phone", ""))
    encoded_text = urllib.parse.quote(msg)

    if wa_phone:
        wa_url = f"https://wa.me/{wa_phone}?text={encoded_text}"
        display_phone = f"+{wa_phone[:2]} {wa_phone[2:]}" if wa_phone.startswith("51") else f"+{wa_phone}"
    else:
        wa_url = f"https://wa.me/?text={encoded_text}"
        display_phone = "Seleccionar contacto"

    return {
        "phone_normalized": wa_phone,
        "display_phone": display_phone,
        "message_text": msg,
        "wa_url": wa_url,
    }


def generate_ticket_html(appointment: dict) -> str:
    """Genera un comprobante HTML listo para imprimir, guardar en PDF o enviar por WhatsApp."""
    turno_label = (
        "MAÑANA (08:00 a 12:00 hrs)"
        if appointment["shift_period"] == "M"
        else "TARDE (14:00 a 18:00 hrs)"
    )
    consultorio_label = (
        "Consultorio de Medicina N° 1 (Turno Mañana)"
        if appointment["shift_period"] == "M"
        else "Consultorio de Medicina N° 2 (Turno Tarde)"
    )
    status_upper = appointment.get("status", "reservada").upper()
    wa_info = generate_whatsapp_info(appointment)

    qr_payload = urllib.parse.quote(
        f"CS-JP2|{appointment['ticket_code']}|DNI:{appointment['patient_dni']}|"
        f"FECHA:{appointment['appointment_date']}|HORA:{appointment['slot_time']}|"
        f"CUPO:{appointment['slot_number']}|MED:{appointment['doctor_cmp']}"
    )
    qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=130x130&data={qr_payload}"

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Ticket de Cita {appointment['ticket_code']} - C.S. Su Santidad Juan Pablo II</title>
<style>
    body {{
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background-color: #f1f5f9;
        padding: 25px;
        color: #0f172a;
    }}
    .ticket-box {{
        max-width: 620px;
        margin: 0 auto;
        background: #ffffff;
        border: 2px solid #0284c7;
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 10px 25px rgba(2, 132, 199, 0.14);
    }}
    .header {{
        background: linear-gradient(135deg, #0369a1 0%, #0f172a 100%);
        color: white;
        padding: 18px 22px;
        text-align: center;
    }}
    .header h2 {{
        margin: 0;
        font-size: 19px;
        letter-spacing: 0.5px;
    }}
    .header p {{
        margin: 4px 0 0 0;
        font-size: 12.5px;
        opacity: 0.92;
    }}
    .ticket-code-banner {{
        background-color: #e0f2fe;
        border-bottom: 1px dashed #7dd3fc;
        padding: 12px 22px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .code-val {{
        font-size: 19px;
        font-weight: 800;
        color: #0369a1;
    }}
    .status-badge {{
        background-color: #dcfce7;
        color: #166534;
        padding: 5px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
        border: 1px solid #86efac;
    }}
    .content {{
        padding: 20px 24px;
    }}
    .top-grid {{
        display: flex;
        gap: 16px;
        align-items: center;
        margin-bottom: 14px;
    }}
    .highlight-slot {{
        flex: 1;
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 12px 16px;
        text-align: center;
    }}
    .highlight-slot .big-time {{
        font-size: 19px;
        font-weight: 800;
        color: #0f172a;
        margin-top: 3px;
    }}
    .qr-box {{
        text-align: center;
        padding: 8px;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        background: #ffffff;
    }}
    .row {{
        display: flex;
        justify-content: space-between;
        padding: 8px 0;
        border-bottom: 1px solid #f1f5f9;
        font-size: 14px;
    }}
    .row strong {{
        color: #334155;
    }}
    .actions {{
        display: flex;
        gap: 10px;
        margin-top: 16px;
    }}
    .btn {{
        flex: 1;
        display: inline-block;
        padding: 12px;
        text-align: center;
        text-decoration: none;
        border: none;
        border-radius: 8px;
        font-size: 14px;
        font-weight: bold;
        cursor: pointer;
    }}
    .btn-print {{
        background-color: #0284c7;
        color: white;
    }}
    .btn-wa {{
        background-color: #16a34a;
        color: white;
    }}
    .footer {{
        background-color: #f8fafc;
        padding: 14px 22px;
        font-size: 12px;
        color: #475569;
        border-top: 1px solid #e2e8f0;
    }}
    @media print {{
        body {{ background: white; padding: 0; }}
        .actions {{ display: none; }}
    }}
</style>
</head>
<body>
    <div class="ticket-box">
        <div class="header">
            <h2>🏥 C.S. SU SANTIDAD JUAN PABLO II</h2>
            <p>SERVICIO DE MEDICINA GENERAL - CONSULTORIO EXTERNO (CE)</p>
        </div>
        <div class="ticket-code-banner">
            <div>
                <span style="font-size:12px; color:#475569;">TICKET OFICIAL DE CITA:</span><br>
                <span class="code-val">{appointment['ticket_code']}</span>
            </div>
            <span class="status-badge">{status_upper}</span>
        </div>
        <div class="content">
            <div class="top-grid">
                <div class="highlight-slot">
                    <div style="font-size:11.5px; color:#64748b; text-transform:uppercase; font-weight:700;">
                        Fecha y Hora Exacta de Atención
                    </div>
                    <div class="big-time">
                        📅 {appointment['appointment_date']}<br>⏰ {appointment['slot_time']} hrs
                    </div>
                    <div style="font-size:13.5px; color:#0369a1; font-weight:700; margin-top:5px;">
                        CUPO N° {appointment['slot_number']:02d} de 16 — TURNO {turno_label}
                    </div>
                </div>
                <div class="qr-box">
                    <img src="{qr_url}" width="110" height="110" alt="QR Cita" />
                    <div style="font-size:10px; color:#64748b; margin-top:2px;">Verificación QR</div>
                </div>
            </div>
            <div class="row">
                <strong>Consultorio:</strong>
                <span>{consultorio_label}</span>
            </div>
            <div class="row">
                <strong>Médico Cirujano:</strong>
                <span>{appointment['doctor_name']} (CMP: {appointment['doctor_cmp']})</span>
            </div>
            <div class="row">
                <strong>Paciente:</strong>
                <span>{appointment['patient_names']}</span>
            </div>
            <div class="row">
                <strong>DNI / Documento:</strong>
                <span>{appointment['patient_dni']} (Edad: {appointment.get('patient_age') or '-'} años)</span>
            </div>
            <div class="row">
                <strong>Seguro / Financiamiento:</strong>
                <span>{appointment.get('insurance_type', 'SIS')}</span>
            </div>
            <div class="row">
                <strong>Celular / WhatsApp:</strong>
                <span>{appointment.get('patient_phone') or '-'}</span>
            </div>
            <div class="row">
                <strong>Motivo de Consulta:</strong>
                <span>{appointment.get('reason') or 'Consulta de Medicina General'}</span>
            </div>
            <div class="actions">
                <button class="btn btn-print" onclick="window.print()">🖨️ Imprimir / Guardar en PDF</button>
                <a class="btn btn-wa" href="{wa_info['wa_url']}" target="_blank">📲 Enviar a mi WhatsApp ({wa_info['display_phone']})</a>
            </div>
        </div>
        <div class="footer">
            <strong>📌 Indicaciones para el Paciente:</strong><br>
            1. Acuda a Admisión del C.S. Su Santidad Juan Pablo II <strong>20 minutos antes</strong> de su horario ({appointment['slot_time']}).<br>
            2. Presente su <strong>DNI físico</strong> y muestre este comprobante o el mensaje de WhatsApp en ventanilla.<br>
            3. Programación sincronizada con el Rol Asistencial Oficial (<a href="https://turnos-ipress.streamlit.app/" target="_blank">turnos-ipress.streamlit.app</a>).
        </div>
    </div>
</body>
</html>"""


def generate_daily_padron_excel(year: int, month: int, day: int) -> io.BytesIO:
    """
    Genera el Padrón Diario en Excel (.xlsx) con los 16 cupos del Consultorio de Medicina Mañana
    y los 16 cupos del Consultorio de Medicina Tarde, más una hoja de resumen mensual.
    """
    day_data = get_day_ce_consultorios(year, month, day)
    month_info = get_month_info(year, month)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Padron_{day:02d}_{month:02d}_{year}"

    title_fill = PatternFill("solid", fgColor="0F172A")
    sub_fill = PatternFill("solid", fgColor="0369A1")
    morning_hdr_fill = PatternFill("solid", fgColor="1E40AF")
    afternoon_hdr_fill = PatternFill("solid", fgColor="065F46")
    table_hdr_fill = PatternFill("solid", fgColor="334155")
    free_fill = PatternFill("solid", fgColor="F8FAFC")
    booked_fill = PatternFill("solid", fgColor="EFF6FF")

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_tbl_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_regular = Font(name="Calibri", size=10, color="0F172A")
    font_muted = Font(name="Calibri", size=9, italic=True, color="64748B")

    thin_side = Side(border_style="thin", color="CBD5E1")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws.merge_cells("A1:J1")
    ws["A1"] = "CENTRO DE SALUD SU SANTIDAD JUAN PABLO II - PADRÓN DIARIO DE CITAS DE MEDICINA (CE)"
    ws["A1"].fill = title_fill
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:J2")
    day_name = day_data["day_meta"]["weekday_name"]
    ws["A2"] = (
        f"Fecha de Atención: {day_name} {day:02d} de {month_info['month_name']} de {year} ({day_data['date_iso']}) | "
        f"Sincronizado con Rol Asistencial turnos-ipress.streamlit.app"
    )
    ws["A2"].fill = sub_fill
    ws["A2"].font = font_sub
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    current_row = 4

    for shift_key, hdr_fill in [("morning", morning_hdr_fill), ("afternoon", afternoon_hdr_fill)]:
        s_info = day_data[shift_key]
        doc = s_info["titular_doctor"]
        doc_desc = (
            f"{doc['doctor_name']} (CMP: {doc['doctor_cmp']} - {doc['condition']})"
            if doc
            else "SIN MÉDICO PROGRAMADO EN CONSULTORIO DE MEDICINA (CE)"
        )

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        cell_banner = ws.cell(row=current_row, column=1)
        cell_banner.value = (
            f"{s_info['consultorio_title'].upper()} | MÉDICO: {doc_desc} | "
            f"OCUPADOS: {s_info['occupied_slots']}/16 | DISPONIBLES: {s_info['available_slots']}/16"
        )
        cell_banner.fill = hdr_fill
        cell_banner.font = font_sub
        cell_banner.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[current_row].height = 22
        current_row += 1

        headers = [
            "Cupo N°",
            "Horario (15 min)",
            "Código Ticket",
            "DNI / Doc.",
            "Apellidos y Nombres del Paciente",
            "Edad",
            "Seguro",
            "Teléfono",
            "Motivo de Consulta",
            "Estado / Firma",
        ]
        for col_idx, h_text in enumerate(headers, start=1):
            c = ws.cell(row=current_row, column=col_idx, value=h_text)
            c.fill = table_hdr_fill
            c.font = font_tbl_hdr
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_cell
        ws.row_dimensions[current_row].height = 20
        current_row += 1

        for slot in s_info["slots"]:
            app = slot["appointment"]
            row_vals = [
                f"Cupo #{slot['slot_number']:02d}",
                slot["label"],
                app["ticket_code"] if app else "DISPONIBLE",
                app["patient_dni"] if app else "-",
                app["patient_names"] if app else "(Cupo Libre)",
                app["patient_age"] if (app and app.get("patient_age")) else "-",
                app["insurance_type"] if app else "-",
                app["patient_phone"] if (app and app.get("patient_phone")) else "-",
                app["reason"] if app else "-",
                app["status"].upper() if app else "LIBRE",
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws.cell(row=current_row, column=col_idx, value=val)
                c.fill = booked_fill if app else free_fill
                c.font = font_bold if (col_idx in (1, 2, 3) and app) else (font_regular if app else font_muted)
                c.alignment = Alignment(
                    horizontal="left" if col_idx in (5, 9) else "center",
                    vertical="center",
                )
                c.border = border_cell
            ws.row_dimensions[current_row].height = 19
            current_row += 1

        current_row += 2

    col_widths = [11, 16, 20, 14, 34, 8, 12, 14, 28, 16]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws_month = wb.create_sheet(title=f"Rol_CE_{month_info['month_name']}_{year}")
    ws_month.merge_cells("A1:H1")
    ws_month["A1"] = (
        f"PROGRAMACIÓN MENSUAL DE CONSULTORIO EXTERNO DE MEDICINA (CE) - {month_info['month_name'].upper()} {year}"
    )
    ws_month["A1"].fill = title_fill
    ws_month["A1"].font = font_title
    ws_month["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_month.row_dimensions[1].height = 26

    m_headers = [
        "Fecha",
        "Día",
        "Médico CE Mañana (08:00-12:00)",
        "Citas Mañana",
        "Libres Mañana",
        "Médico CE Tarde (14:00-18:00)",
        "Citas Tarde",
        "Libres Tarde",
    ]
    for col_idx, h_text in enumerate(m_headers, start=1):
        c = ws_month.cell(row=3, column=col_idx, value=h_text)
        c.fill = table_hdr_fill
        c.font = font_tbl_hdr
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell

    overview = get_month_ce_overview(year, month)
    r_idx = 4
    for d_item in overview:
        m_doc = d_item["morning"]["titular_doctor"]
        t_doc = d_item["afternoon"]["titular_doctor"]
        vals = [
            d_item["date_iso"],
            d_item["day_meta"]["weekday_name"],
            f"{m_doc['doctor_name']} (CMP {m_doc['doctor_cmp']})" if m_doc else "Sin Turno CE",
            f"{d_item['morning']['occupied_slots']}/16",
            d_item["morning"]["available_slots"],
            f"{t_doc['doctor_name']} (CMP {t_doc['doctor_cmp']})" if t_doc else "Sin Turno CE",
            f"{d_item['afternoon']['occupied_slots']}/16",
            d_item["afternoon"]["available_slots"],
        ]
        for col_idx, val in enumerate(vals, start=1):
            c = ws_month.cell(row=r_idx, column=col_idx, value=val)
            c.font = font_regular
            c.alignment = Alignment(
                horizontal="left" if col_idx in (3, 6) else "center",
                vertical="center",
            )
            c.border = border_cell
        r_idx += 1

    m_widths = [14, 14, 38, 14, 14, 38, 14, 14]
    for idx, width in enumerate(m_widths, start=1):
        ws_month.column_dimensions[get_column_letter(idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
