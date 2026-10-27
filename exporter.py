"""
exporter.py
Generador de reportes y Hojas de Datos en Excel (.xlsx) y CSV (.csv) para el Responsable de Admisión,
enlaces y mensajes directos a WhatsApp (+51), y comprobantes imprimibles (HTML/PDF)
de Citas Médicas para el C.S. Su Santidad Juan Pablo II.
"""

import io
import os
import csv
import base64
import hashlib
import textwrap
import urllib.parse
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from schedule_sync import get_day_ce_consultorios, get_month_ce_overview
from holidays_pe import get_month_info


def _load_font(size: int, bold: bool = False):
    """Carga una fuente TrueType limpia compatible con Windows y Linux (Streamlit Cloud)."""
    candidates = (
        [
            "C:/Windows/Fonts/segoeuib.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
            "C:/Windows/Fonts/calibrib.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        ]
        if bold
        else [
            "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibri.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ]
    )
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                pass
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _build_qr_image(payload_text: str, target_size: int = 148) -> Image.Image:
    """
    Genera la imagen del código QR del ticket en alta nitidez.
    Usa la librería `qrcode` si está disponible; de lo contrario, genera una matriz QR
    determinística con patrones de posición estándar para que nunca falle.
    """
    try:
        import qrcode

        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=6,
            border=2,
        )
        qr.add_data(payload_text)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="#0f172a", back_color="#ffffff").convert("RGB")
        return qr_img.resize((target_size, target_size), Image.Resampling.NEAREST)
    except Exception:
        # Matriz QR de respaldo 25x25 con los 3 patrones finder + timing + hash del ticket
        modules = 25
        grid = [[0 for _ in range(modules)] for _ in range(modules)]
        reserved = [[False for _ in range(modules)] for _ in range(modules)]

        def place_finder(r0: int, c0: int):
            for r in range(-1, 8):
                for c in range(-1, 8):
                    rr, cc = r0 + r, c0 + c
                    if 0 <= rr < modules and 0 <= cc < modules:
                        reserved[rr][cc] = True
                        if 0 <= r <= 6 and 0 <= c <= 6:
                            if r in (0, 6) or c in (0, 6) or (2 <= r <= 4 and 2 <= c <= 4):
                                grid[rr][cc] = 1
                            else:
                                grid[rr][cc] = 0

        place_finder(0, 0)
        place_finder(0, modules - 7)
        place_finder(modules - 7, 0)

        for i in range(8, modules - 8):
            reserved[6][i] = True
            reserved[i][6] = True
            grid[6][i] = 1 if i % 2 == 0 else 0
            grid[i][6] = 1 if i % 2 == 0 else 0

        digest = hashlib.sha256(payload_text.encode("utf-8")).digest()
        digest_ext = digest + hashlib.sha256(digest).digest() + hashlib.md5(payload_text.encode("utf-8")).digest()
        bit_idx = 0
        for r in range(modules):
            for c in range(modules):
                if not reserved[r][c]:
                    byte_val = digest_ext[(bit_idx // 8) % len(digest_ext)]
                    bit_val = (byte_val >> (bit_idx % 8)) & 1
                    grid[r][c] = bit_val
                    bit_idx += 1

        quiet = 2
        total_m = modules + quiet * 2
        img = Image.new("RGB", (total_m, total_m), "#ffffff")
        px = img.load()
        for r in range(modules):
            for c in range(modules):
                if grid[r][c]:
                    px[c + quiet, r + quiet] = (15, 23, 42)
        return img.resize((target_size, target_size), Image.Resampling.NEAREST)


def generate_ticket_image_png(appointment: dict) -> bytes:
    """
    Genera el comprobante gráfico oficial de cita médica en formato imagen (.PNG)
    con el mismo diseño, estructura, código QR, colores y datos que el ticket de impresión,
    listo para enviarse como aviso gráfico por WhatsApp o guardarse en el celular/PC.
    """
    width = 860
    height = 1040
    img = Image.new("RGB", (width, height), "#f1f5f9")
    draw = ImageDraw.Draw(img)

    # Fuentes tipográficas escaladas
    f_hdr_title = _load_font(26, bold=True)
    f_hdr_sub = _load_font(16, bold=False)
    f_banner_lbl = _load_font(14, bold=True)
    f_banner_code = _load_font(26, bold=True)
    f_badge = _load_font(15, bold=True)
    f_slot_lbl = _load_font(14, bold=True)
    f_slot_big = _load_font(24, bold=True)
    f_slot_sub = _load_font(16, bold=True)
    f_qr_cap = _load_font(12, bold=False)
    f_row_lbl = _load_font(17, bold=True)
    f_row_val = _load_font(17, bold=False)
    f_row_val_bold = _load_font(17, bold=True)
    f_foot_title = _load_font(15, bold=True)
    f_foot_txt = _load_font(14, bold=False)

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
    doc_type_lbl = appointment.get("patient_doc_type") or "DNI"

    # Marco principal de la tarjeta
    card_x0, card_y0, card_x1, card_y1 = 28, 24, width - 28, height - 24
    draw.rounded_rectangle(
        [card_x0, card_y0, card_x1, card_y1],
        radius=18,
        fill="#ffffff",
        outline="#0284c7",
        width=3,
    )

    # 1. Encabezado Institucional (Azul Oscuro)
    hdr_h = 104
    draw.rounded_rectangle(
        [card_x0, card_y0, card_x1, card_y0 + hdr_h],
        radius=16,
        fill="#0369a1",
    )
    draw.rectangle(
        [card_x0, card_y0 + hdr_h - 18, card_x1, card_y0 + hdr_h],
        fill="#0369a1",
    )
    draw.text(
        (width // 2, card_y0 + 34),
        "C.S. SU SANTIDAD JUAN PABLO II",
        fill="#ffffff",
        font=f_hdr_title,
        anchor="mm",
    )
    draw.text(
        (width // 2, card_y0 + 72),
        "SERVICIO DE MEDICINA GENERAL - CONSULTORIO EXTERNO (CE)",
        fill="#e0f2fe",
        font=f_hdr_sub,
        anchor="mm",
    )

    # 2. Franja de Código de Ticket y Estado
    ban_y0 = card_y0 + hdr_h
    ban_y1 = ban_y0 + 78
    draw.rectangle([card_x0 + 3, ban_y0, card_x1 - 3, ban_y1], fill="#e0f2fe")
    # Línea punteada inferior
    for x_dash in range(card_x0 + 6, card_x1 - 10, 14):
        draw.line([(x_dash, ban_y1), (min(x_dash + 8, card_x1 - 6), ban_y1)], fill="#38bdf8", width=2)

    draw.text(
        (card_x0 + 28, ban_y0 + 16),
        "TICKET OFICIAL DE CITA:",
        fill="#475569",
        font=f_banner_lbl,
    )
    draw.text(
        (card_x0 + 28, ban_y0 + 38),
        str(appointment["ticket_code"]),
        fill="#0369a1",
        font=f_banner_code,
    )

    # Badge de Estado (RESERVADA)
    badge_w, badge_h = 150, 38
    badge_x0 = card_x1 - 28 - badge_w
    badge_y0 = ban_y0 + (78 - badge_h) // 2
    draw.rounded_rectangle(
        [badge_x0, badge_y0, badge_x0 + badge_w, badge_y0 + badge_h],
        radius=19,
        fill="#dcfce7",
        outline="#86efac",
        width=2,
    )
    draw.text(
        (badge_x0 + badge_w // 2, badge_y0 + badge_h // 2),
        status_upper,
        fill="#166534",
        font=f_badge,
        anchor="mm",
    )

    # 3. Bloque Superior Destacado: Fecha/Hora/Cupo + Código QR
    top_y0 = ban_y1 + 22
    top_h = 182
    qr_box_w = 180
    slot_x0 = card_x0 + 28
    slot_x1 = card_x1 - 28 - qr_box_w - 18

    draw.rounded_rectangle(
        [slot_x0, top_y0, slot_x1, top_y0 + top_h],
        radius=14,
        fill="#f8fafc",
        outline="#cbd5e1",
        width=2,
    )
    slot_cx = (slot_x0 + slot_x1) // 2
    draw.text(
        (slot_cx, top_y0 + 26),
        "FECHA Y HORA EXACTA DE ATENCIÓN",
        fill="#64748b",
        font=f_slot_lbl,
        anchor="mm",
    )
    draw.text(
        (slot_cx, top_y0 + 66),
        f"Fecha: {appointment['appointment_date']}",
        fill="#0f172a",
        font=f_slot_big,
        anchor="mm",
    )
    draw.text(
        (slot_cx, top_y0 + 104),
        f"Hora: {appointment['slot_time']} hrs",
        fill="#0f172a",
        font=f_slot_big,
        anchor="mm",
    )
    draw.text(
        (slot_cx, top_y0 + 148),
        f"CUPO N° {appointment['slot_number']:02d} de 16 — TURNO {turno_label}",
        fill="#0369a1",
        font=f_slot_sub,
        anchor="mm",
    )

    # Caja QR a la derecha
    qr_x0 = slot_x1 + 18
    qr_x1 = card_x1 - 28
    draw.rounded_rectangle(
        [qr_x0, top_y0, qr_x1, top_y0 + top_h],
        radius=14,
        fill="#ffffff",
        outline="#e2e8f0",
        width=2,
    )
    qr_raw_payload = (
        f"CS-JP2|{appointment['ticket_code']}|{doc_type_lbl}:{appointment['patient_dni']}|"
        f"FECHA:{appointment['appointment_date']}|HORA:{appointment['slot_time']}|"
        f"CUPO:{appointment['slot_number']}|MED:{appointment['doctor_cmp']}"
    )
    qr_img = _build_qr_image(qr_raw_payload, target_size=138)
    qr_paste_x = qr_x0 + (qr_box_w - 138) // 2
    qr_paste_y = top_y0 + 12
    img.paste(qr_img, (qr_paste_x, qr_paste_y))
    draw.text(
        (qr_x0 + qr_box_w // 2, top_y0 + 163),
        "Verificación QR",
        fill="#64748b",
        font=f_qr_cap,
        anchor="mm",
    )

    # 4. Filas Detalladas de la Cita y del Paciente
    rows_data = [
        ("Consultorio:", consultorio_label, False),
        ("Médico Cirujano:", f"{appointment['doctor_name']} (CMP: {appointment['doctor_cmp']})", True),
        ("Paciente:", str(appointment["patient_names"]), True),
        (
            f"Documento ({doc_type_lbl}):",
            f"{appointment['patient_dni']} (Edad: {appointment.get('patient_age') or '-'} años)",
            True,
        ),
        ("Seguro / Financiamiento:", str(appointment.get("insurance_type") or "SIS"), False),
        ("Celular / WhatsApp:", str(appointment.get("patient_phone") or "-"), False),
        ("Motivo de Consulta:", str(appointment.get("reason") or "Consulta de Medicina General"), False),
    ]

    curr_y = top_y0 + top_h + 20
    row_x0 = card_x0 + 28
    row_x1 = card_x1 - 28
    row_h = 56

    for label_txt, val_txt, is_bold_val in rows_data:
        draw.text(
            (row_x0, curr_y + row_h // 2),
            label_txt,
            fill="#334155",
            font=f_row_lbl,
            anchor="lm",
        )
        # Truncar suavemente si es muy largo para que no se superponga
        val_clean = val_txt if len(val_txt) <= 48 else (val_txt[:45] + "...")
        draw.text(
            (row_x1, curr_y + row_h // 2),
            val_clean,
            fill="#0f172a",
            font=f_row_val_bold if is_bold_val else f_row_val,
            anchor="rm",
        )
        draw.line([(row_x0, curr_y + row_h), (row_x1, curr_y + row_h)], fill="#e2e8f0", width=1)
        curr_y += row_h

    # 5. Pie de Página con Indicaciones Oficiales
    foot_y0 = card_y1 - 155
    draw.rounded_rectangle(
        [card_x0 + 3, foot_y0, card_x1 - 3, card_y1 - 3],
        radius=15,
        fill="#f8fafc",
    )
    draw.line([(card_x0 + 3, foot_y0), (card_x1 - 3, foot_y0)], fill="#cbd5e1", width=2)

    draw.text(
        (row_x0, foot_y0 + 22),
        "Indicaciones para el Paciente:",
        fill="#0369a1",
        font=f_foot_title,
    )
    ind_lines = [
        f"1. Acuda a Admisión del C.S. Su Santidad Juan Pablo II 20 minutos antes de su horario ({appointment['slot_time']} hrs).",
        f"2. Presente su {doc_type_lbl} físico y muestre este aviso gráfico o su ticket impreso en ventanilla.",
        "3. Portal Oficial de Citas: citascsjp2.streamlit.app  |  Rol Asistencial: turnos-ipress.streamlit.app",
    ]
    iy = foot_y0 + 52
    for line in ind_lines:
        draw.text((row_x0, iy), line, fill="#334155", font=f_foot_txt)
        iy += 30

    # Reforzar borde exterior redondeado
    draw.rounded_rectangle(
        [card_x0, card_y0, card_x1, card_y1],
        radius=18,
        fill=None,
        outline="#0284c7",
        width=3,
    )

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def generate_whatsapp_graphic_share_html(appointment: dict, png_bytes: bytes | None = None) -> str:
    """
    Genera un componente interactivo HTML/JS para compartir el Ticket Gráfico (.PNG)
    directamente por WhatsApp:
    - En Celular (Android / iPhone): Abre el menú nativo de compartir con el archivo de imagen PNG
      adjunto (`navigator.share({ files: [pngFile] })`) para enviarlo directo por WhatsApp.
    - En PC / WhatsApp Web (o como segunda opción directa al número): Copia automáticamente la imagen
      gráfica del ticket al portapapeles (`ClipboardItem`), descarga la imagen PNG y abre el chat
      de WhatsApp del paciente para pegar (`Ctrl+V`) la tarjeta gráfica.
    """
    if png_bytes is None:
        png_bytes = generate_ticket_image_png(appointment)

    b64_png = base64.b64encode(png_bytes).decode("ascii")
    wa_info = generate_whatsapp_info(appointment)
    ticket_code = appointment["ticket_code"]
    display_phone = wa_info["display_phone"]
    wa_url = wa_info["wa_url"]
    caption_short = (
        f"🏥 C.S. SU SANTIDAD JUAN PABLO II — CONSTANCIA GRÁFICA DE CITA MÉDICA\n"
        f"🎫 Ticket: {ticket_code}\n"
        f"📅 Fecha: {appointment['appointment_date']} | ⏰ Hora: {appointment['slot_time']} hrs (Cupo #{appointment['slot_number']:02d})\n"
        f"👤 Paciente: {appointment['patient_names']}"
    )
    caption_js = caption_short.replace("\\", "\\\\").replace("`", "\\`").replace("\n", "\\n")

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<style>
    * {{ box-sizing: border-box; font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; }}
    body {{ margin: 0; padding: 2px 0; background: transparent; }}
    .wa-card {{
        display: flex;
        flex-direction: column;
        gap: 8px;
    }}
    .btn-wa-graphic {{
        width: 100%;
        padding: 13px 16px;
        border: none;
        border-radius: 10px;
        background: linear-gradient(135deg, #16a34a 0%, #15803d 100%);
        color: #ffffff;
        font-size: 15px;
        font-weight: 700;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.25);
        transition: transform 0.1s ease, opacity 0.15s ease;
    }}
    .btn-wa-graphic:active {{ transform: scale(0.99); opacity: 0.92; }}
    .btn-wa-direct {{
        width: 100%;
        padding: 10px 14px;
        border: 1.5px solid #16a34a;
        border-radius: 9px;
        background: #f0fdf4;
        color: #166534;
        font-size: 13.5px;
        font-weight: 700;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
    }}
    .status-note {{
        font-size: 12.5px;
        color: #166534;
        background: #dcfce7;
        border: 1px solid #86efac;
        border-radius: 8px;
        padding: 8px 12px;
        text-align: center;
        display: none;
        font-weight: 600;
    }}
</style>
</head>
<body>
<div class="wa-card">
    <button type="button" class="btn-wa-graphic" onclick="shareGraphicTicket()">
        🖼️📲 Enviar Aviso Gráfico por WhatsApp (Imagen del Ticket)
    </button>
    <button type="button" class="btn-wa-direct" onclick="copyImageAndOpenWhatsApp()">
        💬 Copiar Imagen Gráfica y Abrir Chat de WhatsApp ({display_phone})
    </button>
    <div id="waStatusBox" class="status-note"></div>
</div>

<script>
const B64_PNG = "{b64_png}";
const FILE_NAME = "Ticket_{ticket_code}.png";
const WA_URL = "{wa_url}";
const CAPTION_TEXT = `{caption_js}`;

function b64ToBlob(b64Data, contentType='image/png') {{
    const byteCharacters = atob(b64Data);
    const byteArrays = [];
    for (let offset = 0; offset < byteCharacters.length; offset += 512) {{
        const slice = byteCharacters.slice(offset, offset + 512);
        const byteNumbers = new Array(slice.length);
        for (let i = 0; i < slice.length; i++) {{
            byteNumbers[i] = slice.charCodeAt(i);
        }}
        byteArrays.push(new Uint8Array(byteNumbers));
    }}
    return new Blob(byteArrays, {{type: contentType}});
}}

function showNote(msg) {{
    const box = document.getElementById("waStatusBox");
    box.innerText = msg;
    box.style.display = "block";
}}

function triggerDownload(blob) {{
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = FILE_NAME;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 4000);
}}

async function shareGraphicTicket() {{
    const blob = b64ToBlob(B64_PNG, "image/png");
    const file = new File([blob], FILE_NAME, {{ type: "image/png" }});

    // 1. En Celular (Android / iPhone): Compartir directamente el archivo de imagen PNG a WhatsApp
    if (navigator.canShare && navigator.canShare({{ files: [file] }})) {{
        try {{
            await navigator.share({{
                files: [file],
                title: "Ticket de Cita " + "{ticket_code}",
                text: CAPTION_TEXT
            }});
            showNote("✅ Selecciona WhatsApp en tu pantalla para enviar el Ticket Gráfico.");
            return;
        }} catch (err) {{
            if (err && err.name === "AbortError") return;
        }}
    }}

    // 2. Respaldo automático en PC o navegadores sin Web Share de archivos
    await copyImageAndOpenWhatsApp();
}}

async function copyImageAndOpenWhatsApp() {{
    const blob = b64ToBlob(B64_PNG, "image/png");
    let copied = false;
    if (navigator.clipboard && window.ClipboardItem) {{
        try {{
            await navigator.clipboard.write([
                new ClipboardItem({{ "image/png": blob }})
            ]);
            copied = true;
        }} catch (e) {{
            copied = false;
        }}
    }}
    triggerDownload(blob);
    if (copied) {{
        showNote("✅ ¡Imagen del Ticket copiada al portapapeles y descargada! En WhatsApp solo presiona PEGAR (Ctrl+V) para enviar el aviso gráfico.");
    }} else {{
        showNote("✅ ¡Imagen del Ticket descargada! Adjunta la imagen descargada en la ventana de WhatsApp que se abrirá.");
    }}
    setTimeout(() => {{
        window.open(WA_URL, "_blank");
    }}, 450);
}}
</script>
</body>
</html>"""


def format_peru_whatsapp_number(raw_phone: str) -> str:
    """
    Normaliza el teléfono del paciente al formato internacional de Perú (51 + 9 dígitos)
    para enlaces oficiales de WhatsApp (wa.me).
    """
    digits = "".join(ch for ch in (raw_phone or "") if ch.isdigit())
    if not digits:
        return ""
    if len(digits) == 9:
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
    doc_type_lbl = appointment.get("patient_doc_type") or "DNI"

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
        f"🪪 *{doc_type_lbl}:* {appointment['patient_dni']} (Edad: {appointment.get('patient_age') or '-'} años)\n"
        f"📱 *Celular:* {appointment.get('patient_phone') or '-'}\n"
        f"🛡️ *Seguro:* {appointment.get('insurance_type', 'SIS')}\n"
        f"🩺 *Motivo:* {appointment.get('reason') or 'Consulta de Medicina General'}\n"
        f"──────────────────────\n"
        f"📌 *INDICACIONES IMPORTANTES:*\n"
        f"1. Acudir a Admisión *20 minutos antes* de su horario ({appointment['slot_time']}).\n"
        f"2. Presentar su *{doc_type_lbl} físico* y mostrar este mensaje o su ticket en ventanilla.\n"
        f"🌐 *Portal Web de Citas:* https://citascsjp2.streamlit.app/"
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
    doc_type_lbl = appointment.get("patient_doc_type") or "DNI"
    wa_info = generate_whatsapp_info(appointment)

    qr_raw_payload = (
        f"CS-JP2|{appointment['ticket_code']}|{doc_type_lbl}:{appointment['patient_dni']}|"
        f"FECHA:{appointment['appointment_date']}|HORA:{appointment['slot_time']}|"
        f"CUPO:{appointment['slot_number']}|MED:{appointment['doctor_cmp']}"
    )
    qr_pil = _build_qr_image(qr_raw_payload, target_size=130)
    qr_buf = io.BytesIO()
    qr_pil.save(qr_buf, format="PNG")
    qr_url = "data:image/png;base64," + base64.b64encode(qr_buf.getvalue()).decode("ascii")

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
                <strong>Documento ({doc_type_lbl}):</strong>
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
            2. Presente su <strong>{doc_type_lbl} físico</strong> y muestre este comprobante o el mensaje de WhatsApp en ventanilla.<br>
            3. Portal Oficial de Citas: <a href="https://citascsjp2.streamlit.app/" target="_blank"><strong>citascsjp2.streamlit.app</strong></a> | Rol Asistencial: <a href="https://turnos-ipress.streamlit.app/" target="_blank">turnos-ipress.streamlit.app</a>.
        </div>
    </div>
</body>
</html>"""


def build_admission_daily_records(year: int, month: int, day: int, only_booked: bool = False) -> list[dict]:
    """
    Construye la lista estructurada de registros del día para el Responsable de Admisión,
    incluyendo Día/Fecha, Consultorio (Mañana / Tarde), Cupo, Horario, Tipo y N° de Documento
    (conservando ceros iniciales), Paciente, Edad, Celular (9 dígitos), Seguro, Motivo, Médico y Estado.
    """
    day_data = get_day_ce_consultorios(year, month, day)
    date_iso = day_data["date_iso"]
    weekday_name = day_data["day_meta"]["weekday_name"]

    records = []
    for shift_key, cons_label, turno_corto in [
        ("morning", "Consultorio Medicina N° 1 - Mañana (08:00-12:00)", "MAÑANA"),
        ("afternoon", "Consultorio Medicina N° 2 - Tarde (14:00-18:00)", "TARDE"),
    ]:
        s_info = day_data[shift_key]
        doc = s_info["titular_doctor"]
        doc_name = doc["doctor_name"] if doc else "Sin médico CE programado"
        doc_cmp = doc["doctor_cmp"] if doc else "-"

        for slot in s_info["slots"]:
            app = slot["appointment"]
            if only_booked and app is None:
                continue

            doc_type_val = (app.get("patient_doc_type") or "DNI") if app else "-"
            records.append(
                {
                    "Fecha": date_iso,
                    "Día": weekday_name,
                    "Turno": turno_corto,
                    "Consultorio": cons_label,
                    "Cupo N°": f"{slot['slot_number']:02d}",
                    "Horario Atención": slot["label"],
                    "Código Ticket": app["ticket_code"] if app else "DISPONIBLE",
                    "Tipo Doc.": doc_type_val,
                    "N° Documento Identidad": str(app["patient_dni"]) if app else "-",
                    "Apellidos y Nombres del Paciente": app["patient_names"] if app else "(Cupo Libre)",
                    "Edad": app["patient_age"] if (app and app.get("patient_age") is not None) else "-",
                    "Celular (9 dígitos)": str(app["patient_phone"]) if (app and app.get("patient_phone")) else "-",
                    "Seguro / Financiamiento": app["insurance_type"] if app else "-",
                    "Motivo de Consulta": app["reason"] if app else "-",
                    "Médico Cirujano": app["doctor_name"] if app else doc_name,
                    "CMP Médico": app["doctor_cmp"] if app else doc_cmp,
                    "Estado Cita": app["status"].upper() if app else "LIBRE",
                }
            )
    return records


def generate_admission_datasheet_csv(year: int, month: int, day: int, only_booked: bool = False) -> bytes:
    """
    Genera la Hoja de Datos en formato CSV (con BOM UTF-8 para abrir directamente en Excel)
    para el Responsable de Admisión, preservando los ceros iniciales del DNI.
    """
    records = build_admission_daily_records(year, month, day, only_booked=only_booked)
    output = io.StringIO()
    headers = [
        "Fecha",
        "Día",
        "Turno",
        "Consultorio",
        "Cupo N°",
        "Horario Atención",
        "Código Ticket",
        "Tipo Doc.",
        "N° Documento Identidad",
        "Apellidos y Nombres del Paciente",
        "Edad",
        "Celular (9 dígitos)",
        "Seguro / Financiamiento",
        "Motivo de Consulta",
        "Médico Cirujano",
        "CMP Médico",
        "Estado Cita",
    ]
    writer = csv.DictWriter(output, fieldnames=headers, delimiter=";")
    writer.writeheader()
    for r in records:
        row_copy = dict(r)
        # Forzar formato texto en CSV para Excel si empieza con 0
        if row_copy["N° Documento Identidad"] != "-":
            row_copy["N° Documento Identidad"] = f'="{row_copy["N° Documento Identidad"]}"'
        writer.writerow(row_copy)
    return ("\ufeff" + output.getvalue()).encode("utf-8")


def generate_daily_padron_excel(year: int, month: int, day: int) -> io.BytesIO:
    """
    Genera el Libro de Excel (.xlsx) para el Responsable de Admisión con 3 hojas:
    1. 'Hoja_Datos_Admision': Tabla plana estructurada con Día, Consultorio, Tipo y N° de Documento
       (con formato texto '@' para conservar ceros iniciales), Paciente, Edad, Celular, Seguro, Médico y Estado.
    2. 'Padron_Consultorios': Padrón imprimible separado por Consultorio Mañana (16 cupos) y Tarde (16 cupos).
    3. 'Rol_Mensual_CE': Resumen mensual de Consultorios de Medicina.
    """
    day_data = get_day_ce_consultorios(year, month, day)
    month_info = get_month_info(year, month)
    records_all = build_admission_daily_records(year, month, day, only_booked=False)

    wb = Workbook()

    # Estilos comunes
    title_fill = PatternFill("solid", fgColor="0F172A")
    sub_fill = PatternFill("solid", fgColor="0369A1")
    morning_hdr_fill = PatternFill("solid", fgColor="1E40AF")
    afternoon_hdr_fill = PatternFill("solid", fgColor="065F46")
    table_hdr_fill = PatternFill("solid", fgColor="334155")
    free_fill = PatternFill("solid", fgColor="F8FAFC")
    booked_fill = PatternFill("solid", fgColor="EFF6FF")

    font_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
    font_tbl_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_regular = Font(name="Calibri", size=10, color="0F172A")
    font_muted = Font(name="Calibri", size=9, italic=True, color="64748B")

    thin_side = Side(border_style="thin", color="CBD5E1")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # =========================================================================
    # HOJA 1: HOJA DE DATOS DE ADMISIÓN (LISTA DEL DÍA, CONSULTORIO, DOCUMENTO)
    # =========================================================================
    ws_data = wb.active
    ws_data.title = f"Hoja_Datos_{day:02d}_{month:02d}_{year}"

    ws_data.merge_cells("A1:Q1")
    ws_data["A1"] = (
        "C.S. SU SANTIDAD JUAN PABLO II — HOJA DE DATOS DE ADMISIÓN (LISTA DIARIA POR CONSULTORIO DE MEDICINA)"
    )
    ws_data["A1"].fill = title_fill
    ws_data["A1"].font = font_title
    ws_data["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_data.row_dimensions[1].height = 26

    ws_data.merge_cells("A2:Q2")
    day_name = day_data["day_meta"]["weekday_name"]
    ws_data["A2"] = (
        f"Fecha: {day_name} {day:02d}/{month:02d}/{year} ({day_data['date_iso']}) | "
        f"Citas Reservadas: {day_data['total_occupied_day']}/32 | Cupos Libres: {day_data['total_available_day']}/32"
    )
    ws_data["A2"].fill = sub_fill
    ws_data["A2"].font = font_sub
    ws_data["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_data.row_dimensions[2].height = 20

    data_headers = [
        "Fecha",
        "Día",
        "Turno",
        "Consultorio",
        "Cupo N°",
        "Horario Atención",
        "Código Ticket",
        "Tipo Doc.",
        "N° Documento Identidad",
        "Apellidos y Nombres del Paciente",
        "Edad",
        "Celular (9 dígitos)",
        "Seguro / Financiamiento",
        "Motivo de Consulta",
        "Médico Cirujano",
        "CMP Médico",
        "Estado Cita",
    ]
    for col_idx, h_text in enumerate(data_headers, start=1):
        c = ws_data.cell(row=4, column=col_idx, value=h_text)
        c.fill = table_hdr_fill
        c.font = font_tbl_hdr
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_cell
    ws_data.row_dimensions[4].height = 21

    for idx_r, rec in enumerate(records_all, start=5):
        is_booked = rec["Estado Cita"] != "LIBRE"
        for col_idx, col_key in enumerate(data_headers, start=1):
            val = rec[col_key]
            c = ws_data.cell(row=idx_r, column=col_idx, value=str(val) if col_idx in (5, 9, 12, 16) else val)
            if col_idx in (5, 9, 12, 16):
                c.number_format = "@"  # Preservar ceros a la izquierda en DNI (ej. 08123456)
            c.fill = booked_fill if is_booked else free_fill
            c.font = font_bold if (is_booked and col_idx in (7, 8, 9, 10)) else (font_regular if is_booked else font_muted)
            c.alignment = Alignment(
                horizontal="left" if col_idx in (4, 10, 14, 15) else "center",
                vertical="center",
            )
            c.border = border_cell
        ws_data.row_dimensions[idx_r].height = 19

    data_widths = [12, 12, 10, 38, 10, 16, 20, 10, 20, 34, 8, 16, 18, 28, 32, 12, 14]
    for idx, width in enumerate(data_widths, start=1):
        ws_data.column_dimensions[get_column_letter(idx)].width = width

    # =========================================================================
    # HOJA 2: PADRÓN IMPRIMIBLE POR CONSULTORIO (MAÑANA Y TARDE)
    # =========================================================================
    ws = wb.create_sheet(title=f"Padron_{day:02d}_{month:02d}_{year}")
    ws.merge_cells("A1:K1")
    ws["A1"] = "CENTRO DE SALUD SU SANTIDAD JUAN PABLO II - PADRÓN DIARIO DE CITAS DE MEDICINA (CE)"
    ws["A1"].fill = title_fill
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:K2")
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

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
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
            "Tipo Doc.",
            "N° Documento",
            "Apellidos y Nombres del Paciente",
            "Edad",
            "Seguro",
            "Celular (9 díg.)",
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
                (app.get("patient_doc_type") or "DNI") if app else "-",
                str(app["patient_dni"]) if app else "-",
                app["patient_names"] if app else "(Cupo Libre)",
                app["patient_age"] if (app and app.get("patient_age") is not None) else "-",
                app["insurance_type"] if app else "-",
                str(app["patient_phone"]) if (app and app.get("patient_phone")) else "-",
                app["reason"] if app else "-",
                app["status"].upper() if app else "LIBRE",
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws.cell(row=current_row, column=col_idx, value=val)
                if col_idx in (5, 9):
                    c.number_format = "@"
                c.fill = booked_fill if app else free_fill
                c.font = font_bold if (col_idx in (1, 2, 3, 5) and app) else (font_regular if app else font_muted)
                c.alignment = Alignment(
                    horizontal="left" if col_idx in (6, 10) else "center",
                    vertical="center",
                )
                c.border = border_cell
            ws.row_dimensions[current_row].height = 19
            current_row += 1

        current_row += 2

    col_widths = [11, 16, 20, 10, 16, 34, 8, 12, 15, 28, 16]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    # =========================================================================
    # HOJA 3: ROL MENSUAL DE CONSULTORIO EXTERNO DE MEDICINA (CE)
    # =========================================================================
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
