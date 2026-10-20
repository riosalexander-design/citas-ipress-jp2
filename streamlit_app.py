"""
streamlit_app.py
Aplicativo Web en Streamlit para Sacar Citas por Internet - Servicio de Medicina General
Centro de Salud Su Santidad Juan Pablo II.
Sincronizado con la programación asistencial de https://turnos-ipress.streamlit.app/
- Consultorio de Medicina Mañana (08:00 a 12:00): 16 cupos (cada 15 minutos)
- Consultorio de Medicina Tarde  (14:00 a 18:00): 16 cupos (cada 15 minutos)
Diseño 100% Nativo para Móvil (Celular) y PC — Sin etiquetas HTML crudas ni bloques de código.
"""

import os
import tempfile
import importlib
import urllib.parse
from datetime import date
import pandas as pd
import streamlit as st

import database
import holidays_pe
import schedule_sync
import exporter

importlib.reload(database)
importlib.reload(holidays_pe)
importlib.reload(schedule_sync)
importlib.reload(exporter)

from database import init_db, get_connection, hash_pin, log_audit
from holidays_pe import get_month_info
from schedule_sync import (
    get_peru_today,
    sync_from_external_turnos_db,
    ensure_month_schedule,
    get_day_ce_consultorios,
    get_month_ce_overview,
    validate_booking_data,
    book_appointment,
    cancel_appointment,
    update_appointment_status,
    search_appointments,
    set_consultorio_override,
    clear_consultorio_override,
)
from exporter import (
    generate_ticket_html,
    generate_whatsapp_info,
    build_admission_daily_records,
    generate_admission_datasheet_csv,
    generate_daily_padron_excel,
)

# Inicializar Base de Datos
init_db()

# Sincronización automática inicial con la BD de turnos-ipress si existe
if "initial_sync_done" not in st.session_state:
    sync_res = sync_from_external_turnos_db()
    st.session_state["initial_sync_done"] = True
    st.session_state["last_sync_info"] = sync_res

st.set_page_config(
    page_title="C.S. Juan Pablo II - Citas Médicas",
    page_icon="🏥",
    layout="centered",
    initial_sidebar_state="collapsed",
)

if "staff_user" not in st.session_state:
    st.session_state["staff_user"] = None
if "last_booked_ticket" not in st.session_state:
    st.session_state["last_booked_ticket"] = None
if "show_ticket_modal" not in st.session_state:
    st.session_state["show_ticket_modal"] = False
if "pending_booking" not in st.session_state:
    st.session_state["pending_booking"] = None
if "show_preconfirm_modal" not in st.session_state:
    st.session_state["show_preconfirm_modal"] = False
if "booking_error_banner" not in st.session_state:
    st.session_state["booking_error_banner"] = None


# -------------------------------------------------------------------
# VENTANA MODAL 1: CONFIRMACIÓN PREVIA EN GRANDE ANTES DE SACAR LA CITA
# (Condición c: 100% nativa y adaptable a celular sin etiquetas HTML)
# -------------------------------------------------------------------
@st.dialog("🔎 Confirma tus Datos antes de Sacar la Cita", width="large")
def show_pre_booking_confirmation_dialog(p_data: dict):
    st.warning("⚠️ **Por favor verifica que todos tus datos estén correctos:**")
    doc_label = (
        "DNI (8 dígitos)"
        if p_data["patient_doc_type"] == "DNI"
        else "Carné de Extranjería (CE)"
    )

    with st.container(border=True):
        st.markdown(f"### 🪪 {doc_label}: **{p_data['patient_dni']}**")
        st.markdown(f"### 👤 Paciente: **{p_data['patient_names']}**")
        st.markdown(
            f"#### 🎂 Edad: **{p_data['patient_age']} años** &nbsp;|&nbsp; 🛡️ Seguro: **{p_data['insurance_type']}**"
        )
        st.markdown(f"#### 📱 Celular (WhatsApp): **{p_data['patient_phone']}**")
        st.divider()
        st.markdown(
            f"### 📅 Fecha de la Cita: **{p_data['weekday_name']} {p_data['date_iso']}**"
        )
        st.markdown(f"#### 🏥 Consultorio: **{p_data['consultorio_title']}**")
        st.markdown(
            f"### ⏰ Horario: **{p_data['slot_time']} hrs** (Cupo N° {p_data['slot_number']:02d} de 16)"
        )
        st.markdown(
            f"#### 👨‍⚕️ Médico: **{p_data['doctor_name']}** (CMP {p_data['doctor_cmp']})"
        )
        st.markdown(f"#### 🩺 Motivo: **{p_data['reason']}**")

    st.markdown("#### ❓ ¿Los datos mostrados arriba son correctos?")

    if st.button(
        "✅ SÍ, SON CORRECTOS — CONFIRMAR MI CITA",
        use_container_width=True,
        type="primary",
    ):
        ok, msg, ticket_data = book_appointment(
            year=p_data["year"],
            month=p_data["month"],
            day=p_data["day"],
            shift_period=p_data["shift_period"],
            slot_number=p_data["slot_number"],
            patient_doc_type=p_data["patient_doc_type"],
            patient_dni=p_data["patient_dni"],
            patient_names=p_data["patient_names"],
            patient_age=p_data["patient_age"],
            patient_phone=p_data["patient_phone"],
            insurance_type=p_data["insurance_type"],
            reason=p_data["reason"],
            booked_by="Portal Web Paciente (Confirmado)",
            check_past_date=True,
        )
        st.session_state["pending_booking"] = None
        st.session_state["show_preconfirm_modal"] = False
        if ok:
            st.session_state["booking_error_banner"] = None
            st.session_state["last_booked_ticket"] = ticket_data
            st.session_state["show_ticket_modal"] = True
            st.rerun()
        else:
            st.session_state["booking_error_banner"] = msg
            st.rerun()

    if st.button(
        "✏️ NO, VOLVER PARA CORREGIR MIS DATOS",
        use_container_width=True,
        type="secondary",
    ):
        st.session_state["show_preconfirm_modal"] = False
        st.session_state["pending_booking"] = None
        st.rerun()


# -------------------------------------------------------------------
# VENTANA MODAL 2: TICKET OFICIAL CON QR, DESCARGA Y ENVÍO A WHATSAPP
# (100% nativo para celular y PC)
# -------------------------------------------------------------------
@st.dialog("🎫 ¡Cita Confirmada! Tu Constancia de Atención", width="large")
def show_appointment_voucher_dialog(tk: dict):
    t_nom = "MAÑANA (08:00 a 12:00 hrs)" if tk["shift_period"] == "M" else "TARDE (14:00 a 18:00 hrs)"
    doc_type_lbl = tk.get("patient_doc_type") or "DNI"
    wa_info = generate_whatsapp_info(tk)
    html_voucher = generate_ticket_html(tk)

    qr_payload = urllib.parse.quote(
        f"CS-JP2|{tk['ticket_code']}|{doc_type_lbl}:{tk['patient_dni']}|"
        f"FECHA:{tk['appointment_date']}|HORA:{tk['slot_time']}|CUPO:{tk['slot_number']}"
    )
    qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=150x150&data={qr_payload}"

    st.success(f"✅ **¡Tu cita fue reservada con éxito!** N° de Ticket: **{tk['ticket_code']}**")

    with st.container(border=True):
        st.markdown("#### 🏥 C.S. Su Santidad Juan Pablo II — Medicina General")
        st.markdown(f"### 📅 Fecha: **{tk['appointment_date']}** | ⏰ Hora: **{tk['slot_time']} hrs**")
        st.markdown(f"**🔢 Cupo N° {tk['slot_number']:02d} de 16** — Turno {t_nom}")
        st.markdown(f"**👨‍⚕️ Médico:** {tk['doctor_name']} (CMP {tk['doctor_cmp']})")
        st.markdown(
            f"**👤 Paciente:** {tk['patient_names']} ({doc_type_lbl}: {tk['patient_dni']} — {tk.get('patient_age') or '-'} años)"
        )
        st.markdown(
            f"**🛡️ Seguro:** {tk.get('insurance_type', 'SIS')} | **📱 Celular:** {wa_info['display_phone']}"
        )
        st.markdown(f"**🩺 Motivo:** {tk.get('reason') or 'Consulta de Medicina General'}")
        st.image(qr_url, caption=f"Código QR {tk['ticket_code']}", width=130)

    st.markdown("#### 👇 Guarda o envía tu constancia de cita:")
    st.link_button(
        f"📲 1. Enviar Constancia a mi WhatsApp ({wa_info['display_phone']})",
        url=wa_info["wa_url"],
        use_container_width=True,
        type="primary",
    )
    st.download_button(
        "📥 2. Descargar Ticket Oficial (Imprimible / PDF)",
        data=html_voucher.encode("utf-8"),
        file_name=f"Ticket_{tk['ticket_code']}.html",
        mime="text/html",
        use_container_width=True,
        type="secondary",
    )

    st.caption(
        f"📌 Acude a Admisión del C.S. Su Santidad Juan Pablo II 20 minutos antes de tu horario ({tk['slot_time']} hrs) portando tu {doc_type_lbl} físico."
    )


def get_all_doctors():
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT id, name, condition, cmp, role FROM doctors WHERE is_active = 1 ORDER BY condition DESC, name ASC"
        ).fetchall()
        return [dict(r) for r in rows]


doctors_all = get_all_doctors()
doctors_only = [d for d in doctors_all if d["role"] == "doctor"]

meses_nombres = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Setiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

today = get_peru_today()
default_year = today.year if today.year in (2026, 2027) else 2026
default_month = today.month

# -------------------------------------------------------------------
# SIDEBAR: PERIODO Y ACCESO DE PERSONAL
# -------------------------------------------------------------------
with st.sidebar:
    st.title("🏥 C.S. Juan Pablo II")
    st.caption("Servicio de Medicina General")
    st.divider()

    st.subheader("📅 Mes de Programación")
    selected_year = st.selectbox("Año", [2026, 2027], index=0 if default_year == 2026 else 1)
    selected_month = st.selectbox(
        "Mes",
        list(meses_nombres.keys()),
        index=default_month - 1,
        format_func=lambda m: meses_nombres[m],
    )

    ensure_month_schedule(selected_year, selected_month)

    st.divider()
    st.subheader("🔐 Acceso Personal de Salud")
    if st.session_state["staff_user"] is None:
        staff_options = {"ADMIN": "Jefatura / Admisión (ADMIN)"}
        for d in doctors_only:
            staff_options[d["cmp"]] = f"{d['name']} (CMP {d['cmp']})"

        sel_cmp = st.selectbox(
            "Perfil de personal:",
            list(staff_options.keys()),
            format_func=lambda k: staff_options[k],
        )
        pin_input = st.text_input("Contraseña / PIN:", type="password", placeholder="Ingresa tu PIN...")
        if st.button("Ingresar", use_container_width=True, type="primary"):
            with get_connection() as conn:
                u_row = conn.execute(
                    "SELECT id, name, condition, cmp, pin_hash, role FROM doctors WHERE cmp = ?",
                    (sel_cmp,),
                ).fetchone()
                if u_row and u_row["pin_hash"] == hash_pin(pin_input):
                    st.session_state["staff_user"] = dict(u_row)
                    log_audit(u_row["name"], "Inicio de Sesión Personal", f"CMP {sel_cmp}")
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
    else:
        su = st.session_state["staff_user"]
        st.success(f"👤 **{su['name']}**")
        if st.button("Cerrar Sesión", use_container_width=True):
            st.session_state["staff_user"] = None
            st.rerun()

# Disparar ventana emergente de verificación previa (Condición c) o de Ticket final
if st.session_state.get("show_preconfirm_modal") and st.session_state.get("pending_booking"):
    st.session_state["show_preconfirm_modal"] = False
    show_pre_booking_confirmation_dialog(st.session_state["pending_booking"])
elif st.session_state.get("show_ticket_modal") and st.session_state.get("last_booked_ticket"):
    st.session_state["show_ticket_modal"] = False
    show_appointment_voucher_dialog(st.session_state["last_booked_ticket"])

# -------------------------------------------------------------------
# ENCABEZADO PRINCIPAL NATIVO (SIN HTML CRUDO)
# -------------------------------------------------------------------
st.markdown("### 🏥 C.S. Su Santidad Juan Pablo II — Citas de Medicina")
st.caption("🌅 Turno Mañana: 08:00 a 12:00 hrs (16 cupos)  •  🌇 Turno Tarde: 14:00 a 18:00 hrs (16 cupos)")

month_info = get_month_info(selected_year, selected_month)
days_list = month_info["days"]

# -------------------------------------------------------------------
# PESTAÑAS PRINCIPALES (CORTAS PARA CELULAR)
# -------------------------------------------------------------------
tab_reservar, tab_consultar, tab_rol, tab_admision = st.tabs(
    [
        "📅 Sacar Cita",
        "🔍 Mi Cita",
        "🗓️ Rol Médico",
        "🔐 Admisión",
    ]
)

# ===================================================================
# PESTAÑA 1: SACAR CITA POR INTERNET (100% NATIVO PARA CELULAR)
# ===================================================================
with tab_reservar:
    # Mostrar aviso de error si ocurrió alguno al validar o confirmar
    if st.session_state.get("booking_error_banner"):
        err_txt = st.session_state["booking_error_banner"]
        st.error(err_txt)

    # Mostrar tarjeta con botón de descarga y WhatsApp si acaba de reservar
    if st.session_state["last_booked_ticket"] is not None:
        tk = st.session_state["last_booked_ticket"]
        doc_type_lbl = tk.get("patient_doc_type") or "DNI"
        wa_info = generate_whatsapp_info(tk)
        html_voucher = generate_ticket_html(tk)
        t_nom = "MAÑANA" if tk["shift_period"] == "M" else "TARDE"
        with st.container(border=True):
            st.success(f"🎫 **Cita Reservada — Ticket N° {tk['ticket_code']}**")
            st.markdown(
                f"📅 **Fecha:** {tk['appointment_date']} | ⏰ **Hora:** **{tk['slot_time']} hrs** (Cupo N° {tk['slot_number']:02d} - {t_nom})\n\n"
                f"👤 **Paciente:** {tk['patient_names']} ({doc_type_lbl}: {tk['patient_dni']}) • 👨‍⚕️ **Médico:** {tk['doctor_name']}"
            )
            st.link_button(
                f"📲 Enviar Constancia a mi WhatsApp ({wa_info['display_phone']})",
                url=wa_info["wa_url"],
                use_container_width=True,
                type="primary",
            )
            st.download_button(
                "📥 Descargar Ticket (.html / PDF)",
                data=html_voucher.encode("utf-8"),
                file_name=f"Ticket_{tk['ticket_code']}.html",
                mime="text/html",
                key="banner_dl_btn",
                use_container_width=True,
            )

    working_days_meta = [d for d in days_list if not d["is_sunday"]]
    default_day_idx = 0
    if selected_year == today.year and selected_month == today.month:
        for idx, dm in enumerate(working_days_meta):
            if dm["day"] >= today.day:
                default_day_idx = idx
                break

    # ---------------------------------------------------------------
    # PASO 1: SELECCIONAR FECHA DE LA CITA (NATIVO Y GRANDE)
    # ---------------------------------------------------------------
    with st.container(border=True):
        st.markdown("#### 👉 PASO 1: Elige la Fecha de tu Cita Médica")
        selected_day_num = st.selectbox(
            f"Selecciona el día de atención ({meses_nombres[selected_month]} {selected_year}):",
            [d["day"] for d in working_days_meta],
            index=default_day_idx,
            format_func=lambda dnum: next(
                (
                    f"📅 {dm['weekday_name'].upper()} {dm['day']:02d}/{selected_month:02d}/{selected_year}"
                    + (
                        " — ⛔ FECHA PASADA"
                        if date(selected_year, selected_month, dm["day"]) < today
                        else (
                            " — 🟢 HOY"
                            if date(selected_year, selected_month, dm["day"]) == today
                            else " — 🟢 DISPONIBLE"
                        )
                    )
                    + (f" ({dm['holiday_name']})" if dm["is_holiday"] else "")
                    for dm in working_days_meta
                    if dm["day"] == dnum
                ),
                str(dnum),
            ),
        )

    selected_date_obj = date(selected_year, selected_month, selected_day_num)
    is_past_date = selected_date_obj < today

    day_consultorios = get_day_ce_consultorios(selected_year, selected_month, selected_day_num)
    m_info = day_consultorios["morning"]
    t_info = day_consultorios["afternoon"]
    sel_day_meta = day_consultorios["day_meta"]

    if is_past_date:
        st.error(
            f"⛔ **FECHA PASADA NO DISPONIBLE:** La fecha escogida (**{sel_day_meta['weekday_name']} {selected_day_num:02d}/{selected_month:02d}/{selected_year}**) "
            f"ya pasó (hoy es **{today.strftime('%d/%m/%Y')}**). No se permite generar citas en fechas pasadas. "
            "Por favor elige hoy o una fecha futura en el **Paso 1**."
        )
    else:
        available_shifts = []
        if m_info["is_enabled"] and m_info["available_slots"] > 0:
            available_shifts.append("M")
        if t_info["is_enabled"] and t_info["available_slots"] > 0:
            available_shifts.append("T")

        if not available_shifts:
            st.warning(
                "No quedan cupos disponibles en Consultorio de Medicina para esta fecha. "
                "Por favor elige otro día en el Paso 1."
            )
        else:
            # ---------------------------------------------------------------
            # PASO 2: ELEGIR TURNO (MAÑANA O TARDE)
            # ---------------------------------------------------------------
            with st.container(border=True):
                st.markdown("#### 🏥 PASO 2: Elige el Turno (Mañana o Tarde)")
                chosen_shift = st.radio(
                    "Selecciona el turno en el que deseas atenderte:",
                    available_shifts,
                    format_func=lambda sp: (
                        f"🌅 MAÑANA (08:00 a 12:00) — Dr(a). {m_info['titular_doctor']['doctor_name']} ({m_info['available_slots']} cupos libres)"
                        if sp == "M"
                        else f"🌇 TARDE (14:00 a 18:00) — Dr(a). {t_info['titular_doctor']['doctor_name']} ({t_info['available_slots']} cupos libres)"
                    ),
                    horizontal=False,
                )

                active_shift_info = m_info if chosen_shift == "M" else t_info
                doc_act = active_shift_info["titular_doctor"]
                st.info(
                    f"👨‍⚕️ **Médico Programado:** {doc_act['doctor_name']} (CMP {doc_act['doctor_cmp']}) — "
                    f"**✅ {active_shift_info['available_slots']} de 16 cupos libres**"
                )

            # ---------------------------------------------------------------
            # PASO 3: ELEGIR HORARIO Y LLENAR DATOS (1 SOLA COLUMNA PARA CELULAR)
            # ---------------------------------------------------------------
            free_slots_list = [s for s in active_shift_info["slots"] if s["is_available"]]
            with st.form("booking_patient_form", clear_on_submit=False):
                st.markdown(
                    f"#### 📝 PASO 3: Horario y Datos del Paciente ({sel_day_meta['weekday_name']} {selected_day_num:02d}/{selected_month:02d}/{selected_year})"
                )
                st.caption("⚠️ Todos los campos con (*) son obligatorios y deben llenarse por completo.")

                selected_slot_num = st.selectbox(
                    "⏰ 1. Elige tu Horario de Cita (cada 15 min) (*):",
                    [s["slot_number"] for s in free_slots_list],
                    format_func=lambda snum: next(
                        (
                            f"⏰ {s['label']} hrs — Cupo N° {s['slot_number']:02d}"
                            for s in free_slots_list
                            if s["slot_number"] == snum
                        ),
                        str(snum),
                    ),
                )

                inp_doc_type = st.selectbox(
                    "🪪 2. Tipo de Documento de Identidad (*):",
                    ["DNI", "CE"],
                    format_func=lambda x: "DNI (8 dígitos numéricos)" if x == "DNI" else "Carné de Extranjería (CE)",
                )

                inp_dni = st.text_input(
                    "🔢 3. Número de Documento (Solo números) (*):",
                    max_chars=12,
                    placeholder="Ejemplo DNI: 08451236",
                )

                inp_names = st.text_input(
                    "👤 4. Apellidos y Nombres Completos del Paciente (*):",
                    placeholder="Ejemplo: PEREZ GOMEZ, JUAN CARLOS",
                )

                inp_age = st.number_input(
                    "🎂 5. Edad del Paciente (en años) (*):",
                    min_value=0,
                    max_value=120,
                    value=None,
                    step=1,
                    placeholder="Ejemplo: 35",
                )

                inp_phone = st.text_input(
                    "📱 6. Número de Celular (9 dígitos obligatorio) (*):",
                    max_chars=9,
                    placeholder="Ejemplo: 987654321",
                )

                inp_insurance = st.selectbox(
                    "🛡️ 7. Tipo de Seguro / Financiamiento (*):",
                    [
                        "-- Selecciona tu Seguro --",
                        "SIS (Seguro Integral de Salud)",
                        "Pagante / Particular",
                        "Exonerado / Otro",
                    ],
                )

                inp_reason = st.text_input(
                    "🩺 8. Motivo de Consulta Médica (*):",
                    placeholder="Ejemplo: Control médico, dolor de cabeza, presión alta...",
                )

                submit_verify = st.form_submit_button(
                    "✅ CONTINUAR Y VERIFICAR MIS DATOS",
                    use_container_width=True,
                    type="primary",
                )

                if submit_verify:
                    ins_val = (
                        ""
                        if inp_insurance.startswith("--")
                        else inp_insurance.split(" ")[0]
                    )
                    ok_v, msg_v, v_payload = validate_booking_data(
                        year=selected_year,
                        month=selected_month,
                        day=selected_day_num,
                        shift_period=chosen_shift,
                        slot_number=selected_slot_num,
                        patient_doc_type=inp_doc_type,
                        patient_dni=inp_dni,
                        patient_names=inp_names,
                        patient_age=inp_age,
                        patient_phone=inp_phone,
                        insurance_type=ins_val,
                        reason=inp_reason,
                        check_past_date=True,
                    )
                    if ok_v and v_payload:
                        st.session_state["booking_error_banner"] = None
                        st.session_state["pending_booking"] = v_payload
                        st.session_state["show_preconfirm_modal"] = True
                        st.rerun()
                    else:
                        st.session_state["booking_error_banner"] = msg_v
                        st.rerun()

# ===================================================================
# PESTAÑA 2: CONSULTAR / CANCELAR MI CITA (ADAPTADO A CELULAR)
# ===================================================================
with tab_consultar:
    st.subheader("🔍 Consulta tu Cita, Descarga tu Ticket o Envíalo a WhatsApp")
    is_staff_logged = st.session_state["staff_user"] is not None
    if is_staff_logged:
        st.markdown(
            "Ingresa el **DNI / Carné de Extranjería**, **Código de Ticket** o **Apellidos** "
            "para verificar el horario, descargar el comprobante o enviarlo por WhatsApp."
        )
    else:
        st.markdown(
            "Ingresa tu número exacto de **DNI / Carné de Extranjería** o tu **Código de Ticket** "
            "para ver tu cita, enviarla a tu WhatsApp o cancelarla."
        )

    query_str = st.text_input(
        "Ingresa tu N° de Documento (DNI / CE)" + (" o Apellidos:" if is_staff_logged else " o Código de Ticket:"),
        placeholder="Ejemplo: 08451236",
    )
    btn_search = st.button("🔎 Buscar mi Cita", use_container_width=True, type="primary")

    if query_str.strip() or btn_search:
        if not query_str.strip():
            st.warning("Por favor ingresa tu número de DNI/CE o código de ticket para buscar.")
        else:
            results = search_appointments(query_str, allow_name_like=is_staff_logged)
            if not results:
                st.info("No se encontraron citas registradas con el dato ingresado.")
            else:
                st.markdown(f"Se encontraron **{len(results)}** cita(s):")
                for app_item in results:
                    t_label = (
                        "Turno Mañana (08:00 - 12:00)"
                        if app_item["shift_period"] == "M"
                        else "Turno Tarde (14:00 - 18:00)"
                    )
                    status_icon = {
                        "reservada": "🟢 RESERVADA",
                        "atendida": "🔵 ATENDIDA",
                        "no_asistio": "🟠 NO ASISTIÓ",
                        "cancelada": "⚪ CANCELADA",
                    }.get(app_item["status"], app_item["status"].upper())

                    doc_type_lbl = app_item.get("patient_doc_type") or "DNI"
                    wa_item = generate_whatsapp_info(app_item)
                    voucher_html = generate_ticket_html(app_item)

                    with st.container(border=True):
                        st.markdown(f"#### 🎫 Ticket: **{app_item['ticket_code']}** ({status_icon})")
                        st.markdown(
                            f"👤 **Paciente:** {app_item['patient_names']} ({doc_type_lbl}: {app_item['patient_dni']})"
                        )
                        st.markdown(
                            f"📅 **Fecha:** {app_item['appointment_date']} — **{t_label}**"
                        )
                        st.markdown(
                            f"⏰ **Horario:** **{app_item['slot_time']} hrs** (Cupo N° {app_item['slot_number']:02d} de 16)"
                        )
                        st.markdown(
                            f"👨‍⚕️ **Médico:** {app_item['doctor_name']} (CMP {app_item['doctor_cmp']})"
                        )
                        st.link_button(
                            f"📲 Enviar a WhatsApp ({wa_item['display_phone']})",
                            url=wa_item["wa_url"],
                            use_container_width=True,
                            type="primary",
                        )
                        st.download_button(
                            "📥 Descargar Ticket (.html / PDF)",
                            data=voucher_html.encode("utf-8"),
                            file_name=f"Ticket_{app_item['ticket_code']}.html",
                            mime="text/html",
                            key=f"dl_{app_item['id']}",
                            use_container_width=True,
                        )
                        if app_item["status"] == "reservada":
                            if st.button(
                                "🗑️ Cancelar esta Cita",
                                key=f"canc_{app_item['id']}",
                                use_container_width=True,
                            ):
                                ok_c, msg_c = cancel_appointment(
                                    app_item["ticket_code"],
                                    patient_dni=app_item["patient_dni"],
                                    cancelled_by=f"Paciente {doc_type_lbl} {app_item['patient_dni']}",
                                )
                                if ok_c:
                                    st.success(msg_c)
                                    st.rerun()
                                else:
                                    st.error(msg_c)

# ===================================================================
# PESTAÑA 3: ROL MENSUAL DE MEDICINA
# ===================================================================
with tab_rol:
    st.subheader(
        f"🗓️ Rol de Médicos en Consultorio — {meses_nombres[selected_month]} {selected_year}"
    )
    st.caption(
        "Programación oficial de médicos del C.S. Su Santidad Juan Pablo II en los turnos Mañana (08:00 a 12:00) y Tarde (14:00 a 18:00)."
    )

    month_overview = get_month_ce_overview(selected_year, selected_month)
    table_rows = []
    for item in month_overview:
        dm = item["day_meta"]
        if dm["is_sunday"]:
            m_doc_txt = "🔒 Domingo"
            t_doc_txt = "🔒 Domingo"
        else:
            m_doc = item["morning"]["titular_doctor"]
            t_doc = item["afternoon"]["titular_doctor"]
            m_doc_txt = (
                f"{m_doc['doctor_name']} (CMP {m_doc['doctor_cmp']})"
                if m_doc
                else "⚠️ Sin turno Mañana"
            )
            t_doc_txt = (
                f"{t_doc['doctor_name']} (CMP {t_doc['doctor_cmp']})"
                if t_doc
                else "⚠️ Sin turno Tarde"
            )

        table_rows.append(
            {
                "Fecha": item["date_iso"],
                "Día": dm["weekday_name"] + (f" ({dm['holiday_name']})" if dm["is_holiday"] else ""),
                "🌅 Médico Mañana (08:00-12:00)": m_doc_txt,
                "Libres Mañana": f"{item['morning']['available_slots']}/16" if item["morning"]["is_enabled"] else "0/16",
                "🌇 Médico Tarde (14:00-18:00)": t_doc_txt,
                "Libres Tarde": f"{item['afternoon']['available_slots']}/16" if item["afternoon"]["is_enabled"] else "0/16",
            }
        )

    st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True, height=440)

# ===================================================================
# PESTAÑA 4: PANEL DE ADMISIÓN Y DESCARGA DE HOJA DE DATOS (Condición d)
# EXCLUSIVO PARA PERSONAL DE SALUD CON CONTRASEÑA
# ===================================================================
with tab_admision:
    st.subheader("🏥 Panel de Admisión y Personal de Salud")

    if st.session_state["staff_user"] is None:
        st.warning(
            "🔒 **Acceso Restringido:** La lista de pacientes y la **descarga de citas en Excel o CSV** "
            "están disponibles **únicamente para el Personal de Salud con contraseña**."
        )
        with st.container(border=True):
            st.markdown("#### 🔐 Ingreso para Personal de Salud / Admisión")
            staff_login_options = {"ADMIN": "Jefatura / Responsable de Admisión (ADMIN)"}
            for d in doctors_only:
                staff_login_options[d["cmp"]] = f"{d['name']} (CMP {d['cmp']})"

            tab_sel_cmp = st.selectbox(
                "👤 Selecciona tu Perfil:",
                list(staff_login_options.keys()),
                format_func=lambda k: staff_login_options[k],
                key="tab_adm_cmp_select",
            )
            tab_pin_input = st.text_input(
                "🔑 Contraseña / PIN de Acceso:",
                type="password",
                placeholder="Ingresa tu contraseña...",
                key="tab_adm_pin_input",
            )
            if st.button(
                "🔓 Ingresar y Habilitar Descargas",
                use_container_width=True,
                type="primary",
                key="tab_adm_login_btn",
            ):
                with get_connection() as conn:
                    u_row = conn.execute(
                        "SELECT id, name, condition, cmp, pin_hash, role FROM doctors WHERE cmp = ?",
                        (tab_sel_cmp,),
                    ).fetchone()
                    if u_row and u_row["pin_hash"] == hash_pin(tab_pin_input):
                        st.session_state["staff_user"] = dict(u_row)
                        log_audit(u_row["name"], "Inicio de Sesión Personal (Pestaña Admisión)", f"CMP {tab_sel_cmp}")
                        st.rerun()
                    else:
                        st.error("❌ Contraseña incorrecta. Acceso denegado.")
    else:
        staff = st.session_state["staff_user"]
        st.success(f"🔓 **Sesión activa:** {staff['name']} ({staff['cmp']})")
        if st.button("🔒 Cerrar Sesión de Personal", key="tab_adm_logout_btn"):
            st.session_state["staff_user"] = None
            st.rerun()

        st.divider()
        adm_day_num = st.selectbox(
            "📅 Seleccionar Día para Descargar o Consultar:",
            [d["day"] for d in days_list],
            index=selected_day_num - 1 if (1 <= selected_day_num <= len(days_list)) else 0,
            format_func=lambda dnum: (
                f"{days_list[dnum - 1]['weekday_name']} {dnum:02d}/{selected_month:02d}/{selected_year}"
            ),
            key="adm_day_selector",
        )
        only_booked_filter = st.selectbox(
            "📋 Filtrar registros en la Hoja de Datos:",
            [False, True],
            format_func=lambda b: "Todos los 32 Cupos (Mañana y Tarde)" if not b else "Solo Pacientes Citados del Día",
        )

        excel_bytes = generate_daily_padron_excel(selected_year, selected_month, adm_day_num)
        st.download_button(
            "📊 Descargar Hoja de Datos Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Hoja_Datos_Admision_CS_JuanPabloII_{selected_year}_{selected_month:02d}_{adm_day_num:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        csv_bytes = generate_admission_datasheet_csv(
            selected_year, selected_month, adm_day_num, only_booked=only_booked_filter
        )
        st.download_button(
            "📄 Descargar Hoja de Datos (.csv)",
            data=csv_bytes,
            file_name=f"Lista_Admision_Pacientes_{selected_year}_{selected_month:02d}_{adm_day_num:02d}.csv",
            mime="text/csv",
            use_container_width=True,
        )

        adm_records = build_admission_daily_records(
            selected_year, selected_month, adm_day_num, only_booked=only_booked_filter
        )
        st.markdown("#### 📋 Lista del Día para Admisión")
        if not adm_records:
            st.info("No hay pacientes citados aún para la fecha seleccionada.")
        else:
            st.dataframe(pd.DataFrame(adm_records), use_container_width=True, hide_index=True, height=340)

        adm_day_data = get_day_ce_consultorios(selected_year, selected_month, adm_day_num)

        with st.expander("✅ Control de Asistencia y Reenvío a WhatsApp"):
            booked_today = []
            for sk in ("morning", "afternoon"):
                for sl in adm_day_data[sk]["slots"]:
                    if sl["appointment"]:
                        booked_today.append(sl["appointment"])

            if not booked_today:
                st.caption("No hay citas reservadas en esta fecha para actualizar estado.")
            else:
                sel_app_id = st.selectbox(
                    "Seleccionar Cita del Día:",
                    [a["id"] for a in booked_today],
                    format_func=lambda aid: next(
                        (
                            f"[{a['shift_period']} Cupo #{a['slot_number']:02d} {a['slot_time']}] "
                            f"{a['patient_names']} ({a.get('patient_doc_type', 'DNI')} {a['patient_dni']}) — {a['status'].upper()}"
                            for a in booked_today
                            if a["id"] == aid
                        ),
                        str(aid),
                    ),
                )
                selected_app_obj = next((a for a in booked_today if a["id"] == sel_app_id), booked_today[0])
                wa_adm = generate_whatsapp_info(selected_app_obj)

                new_st = st.selectbox(
                    "Nuevo Estado:",
                    ["reservada", "atendida", "no_asistio", "cancelada"],
                    format_func=lambda s: {
                        "reservada": "🟢 Reservada",
                        "atendida": "🔵 Atendida en Consultorio",
                        "no_asistio": "🟠 No Asistió (Falta)",
                        "cancelada": "⚪ Cancelada (Liberar Cupo)",
                    }[s],
                )
                if st.button("Guardar Estado", use_container_width=True, type="primary"):
                    ok_u, msg_u = update_appointment_status(sel_app_id, new_st, staff["name"])
                    if ok_u:
                        st.success(msg_u)
                        st.rerun()
                    else:
                        st.error(msg_u)
                st.link_button(
                    f"📲 Enviar WhatsApp al Paciente ({wa_adm['display_phone']})",
                    url=wa_adm["wa_url"],
                    use_container_width=True,
                )

        with st.expander("👨‍⚕️ Asignación / Contingencia de Consultorio"):
            ov_shift = st.selectbox(
                "Turno:",
                ["M", "T"],
                format_func=lambda s: "🌅 Mañana (08:00-12:00)" if s == "M" else "🌇 Tarde (14:00-18:00)",
            )
            ov_doc_id = st.selectbox(
                "Médico Cirujano Titular:",
                [d["id"] for d in doctors_only],
                format_func=lambda did: next(
                    (f"{d['name']} (CMP {d['cmp']})" for d in doctors_only if d["id"] == did),
                    str(did),
                ),
            )
            ov_enabled = st.checkbox("Consultorio Habilitado (16 cupos)", value=True)
            ov_notes = st.text_input("Observación:", placeholder="Ej. Cobertura / Reemplazo")

            if st.button("💾 Aplicar Asignación en esta Fecha", use_container_width=True, type="primary"):
                ok_o, msg_o = set_consultorio_override(
                    selected_year,
                    selected_month,
                    adm_day_num,
                    ov_shift,
                    ov_doc_id if ov_enabled else None,
                    ov_enabled,
                    ov_notes,
                    staff["name"],
                )
                st.success(msg_o)
                st.rerun()
            if st.button("🔄 Restaurar Según Rol Original de Turnos-IPRESS", use_container_width=True):
                clear_consultorio_override(
                    selected_year, selected_month, adm_day_num, ov_shift, staff["name"]
                )
                st.success("Restaurado al rol automático de Turnos-IPRESS.")
                st.rerun()

        with st.expander("🔄 Importar Base de Datos de Turnos-IPRESS (.db)"):
            uploaded_db = st.file_uploader("Subir archivo ipress_turnos.db", type=["db", "sqlite", "sqlite3"])
            if uploaded_db is not None:
                if st.button("📥 Importar Turnos desde Archivo Subido", type="primary"):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".db") as tmp_f:
                        tmp_f.write(uploaded_db.read())
                        tmp_path = tmp_f.name
                    res_up = sync_from_external_turnos_db(tmp_path)
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
                    if res_up["synced"]:
                        st.success(res_up["message"])
                        st.rerun()
                    else:
                        st.error(res_up["message"])
