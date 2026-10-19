"""
streamlit_app.py
Aplicativo Web en Streamlit para Sacar Citas por Internet - Servicio de Medicina General
Centro de Salud Su Santidad Juan Pablo II.
Sincronizado con la programación asistencial de https://turnos-ipress.streamlit.app/
- Consultorio de Medicina Mañana (08:00 a 12:00): 16 cupos (cada 15 minutos)
- Consultorio de Medicina Tarde  (14:00 a 18:00): 16 cupos (cada 15 minutos)
"""

import os
import tempfile
import importlib
from datetime import datetime, date
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

from database import init_db, get_connection, hash_pin, log_audit, find_external_turnos_db
from holidays_pe import get_month_info
from schedule_sync import (
    MAX_SLOTS_PER_SHIFT,
    MORNING_SLOTS,
    AFTERNOON_SLOTS,
    sync_from_external_turnos_db,
    ensure_month_schedule,
    get_day_ce_consultorios,
    get_month_ce_overview,
    book_appointment,
    cancel_appointment,
    update_appointment_status,
    search_appointments,
    set_consultorio_override,
    clear_consultorio_override,
)
from exporter import generate_ticket_html, generate_daily_padron_excel

# Inicializar Base de Datos
init_db()

# Sincronización automática inicial con la BD de turnos-ipress si existe en el equipo
if "initial_sync_done" not in st.session_state:
    sync_res = sync_from_external_turnos_db()
    st.session_state["initial_sync_done"] = True
    st.session_state["last_sync_info"] = sync_res

st.set_page_config(
    page_title="C.S. Su Santidad Juan Pablo II - Citas de Medicina en Línea",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS institucionales
st.markdown(
    """
<style>
    .hero-banner {
        background: linear-gradient(135deg, #0369a1 0%, #0f172a 100%);
        color: white;
        padding: 20px 26px;
        border-radius: 14px;
        margin-bottom: 16px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.14);
    }
    .hero-banner h1 {
        margin: 0;
        font-size: 25px;
        font-weight: 800;
        color: #ffffff;
    }
    .hero-banner p {
        margin: 6px 0 0 0;
        font-size: 14px;
        color: #e0f2fe;
    }
    .consultorio-card {
        border-radius: 12px;
        padding: 16px 18px;
        border: 1px solid #cbd5e1;
        background-color: #ffffff;
        margin-bottom: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
    }
    .consultorio-morning {
        border-left: 6px solid #0284c7;
        background: linear-gradient(180deg, #f0f9ff 0%, #ffffff 100%);
    }
    .consultorio-afternoon {
        border-left: 6px solid #059669;
        background: linear-gradient(180deg, #ecfdf5 0%, #ffffff 100%);
    }
    .consultorio-disabled {
        border-left: 6px solid #94a3b8;
        background-color: #f8fafc;
        color: #64748b;
    }
    .slot-pill-free {
        background-color: #dcfce7;
        color: #166534;
        border: 1px solid #86efac;
        border-radius: 8px;
        padding: 7px 6px;
        text-align: center;
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .slot-pill-busy {
        background-color: #fee2e2;
        color: #991b1b;
        border: 1px solid #fca5a5;
        border-radius: 8px;
        padding: 7px 6px;
        text-align: center;
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .ticket-card {
        background-color: #f0f9ff;
        border: 2px solid #0284c7;
        border-radius: 12px;
        padding: 18px 22px;
        margin-top: 12px;
    }
    .badge-shift {
        display: inline-block;
        padding: 2px 6px;
        border-radius: 6px;
        font-size: 11px;
        font-weight: bold;
    }
    .badge-M-CE { background-color: #dbeafe; color: #1e40af; border: 1px solid #93c5fd; }
    .badge-T-CE { background-color: #d1fae5; color: #065f46; border: 1px solid #6ee7b7; }
    .badge-12h  { background-color: #fef3c7; color: #92400e; border: 1px solid #fcd34d; }
    .badge-DES  { background-color: #f3e8ff; color: #6b21a8; border: 1px solid #d8b4fe; }
    .badge-other { background-color: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; }
</style>
""",
    unsafe_allow_html=True,
)

if "staff_user" not in st.session_state:
    st.session_state["staff_user"] = None
if "last_booked_ticket" not in st.session_state:
    st.session_state["last_booked_ticket"] = None


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

# -------------------------------------------------------------------
# SIDEBAR: PERIODO, ESTADO DE SINCRONIZACIÓN Y ACCESO DE PERSONAL
# -------------------------------------------------------------------
with st.sidebar:
    st.title("🏥 C.S. Juan Pablo II")
    st.caption("Servicio de Medicina General — Citas por Internet")
    st.markdown("---")

    st.subheader("📅 Periodo de Programación")
    today = date.today()
    default_year = today.year if today.year in (2026, 2027) else 2026
    default_month = today.month

    col_y, col_m = st.columns(2)
    with col_y:
        selected_year = st.selectbox("Año", [2026, 2027], index=0 if default_year == 2026 else 1)
    with col_m:
        selected_month = st.selectbox(
            "Mes",
            list(meses_nombres.keys()),
            index=default_month - 1,
            format_func=lambda m: meses_nombres[m],
        )

    ensure_month_schedule(selected_year, selected_month)

    st.markdown("---")
    st.subheader("🔄 Enlace con Turnos-IPRESS")
    st.markdown(
        "Las citas se programan según el rol oficial de **[Turnos-IPRESS Streamlit](https://turnos-ipress.streamlit.app/)** "
        "(Servicio **`CE`: Consultorio Externo de Medicina**)."
    )
    if st.button("🔄 Sincronizar Rol Médico Ahora", use_container_width=True):
        res = sync_from_external_turnos_db()
        ensure_month_schedule(selected_year, selected_month)
        st.session_state["last_sync_info"] = res
        st.success(res["message"])
        st.rerun()

    sync_info = st.session_state.get("last_sync_info", {})
    if sync_info.get("synced"):
        st.caption(f"✅ {sync_info.get('message')}")
    else:
        st.caption("ℹ️ Rol mensual activo según reglas oficiales de Turnos-IPRESS.")

    st.markdown("---")
    st.subheader("🔐 Acceso Admisión / Médicos")
    if st.session_state["staff_user"] is None:
        staff_options = {"ADMIN": "Jefatura / Admisión (ADMIN)"}
        for d in doctors_only:
            staff_options[d["cmp"]] = f"{d['name']} (CMP {d['cmp']})"

        sel_cmp = st.selectbox(
            "Perfil de personal:",
            list(staff_options.keys()),
            format_func=lambda k: staff_options[k],
        )
        pin_input = st.text_input("PIN de acceso:", type="password", placeholder="Ingresa tu PIN...")
        if st.button("Ingresar a Admisión", use_container_width=True):
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
                    st.error("PIN incorrecto.")
    else:
        su = st.session_state["staff_user"]
        st.success(f"👤 **{su['name']}**")
        st.caption(f"CMP/Usuario: `{su['cmp']}` | Rol: `{su['role'].upper()}`")
        if st.button("Cerrar Sesión", use_container_width=True):
            st.session_state["staff_user"] = None
            st.rerun()

# -------------------------------------------------------------------
# BANNER PRINCIPAL
# -------------------------------------------------------------------
st.markdown(
    f"""
<div class="hero-banner">
    <h1>🏥 Centro de Salud Su Santidad Juan Pablo II — Citas Médicas en Línea</h1>
    <p>
        <strong>Servicio de Medicina General (Consultorio Externo - CE)</strong> &nbsp;|&nbsp;
        🌅 <strong>Consultorio Mañana:</strong> 16 cupos (08:00 a 12:00, cada 15 min) &nbsp;|&nbsp;
        🌇 <strong>Consultorio Tarde:</strong> 16 cupos (14:00 a 18:00, cada 15 min) &nbsp;|&nbsp;
        🔗 Programación sincronizada con <a href="https://turnos-ipress.streamlit.app/" target="_blank" style="color:#7dd3fc; font-weight:bold;">turnos-ipress.streamlit.app</a>
    </p>
</div>
""",
    unsafe_allow_html=True,
)

month_info = get_month_info(selected_year, selected_month)
days_list = month_info["days"]

# -------------------------------------------------------------------
# PESTAÑAS PRINCIPALES
# -------------------------------------------------------------------
tab_reservar, tab_consultar, tab_rol, tab_admision = st.tabs(
    [
        "📅 1. Sacar Cita por Internet",
        "🔍 2. Consultar / Cancelar mi Cita",
        "🗓️ 3. Rol Mensual de Medicina (Turnos-IPRESS)",
        "🏥 4. Panel de Admisión y Padrón Diario",
    ]
)

# ===================================================================
# PESTAÑA 1: SACAR CITA POR INTERNET (PORTAL DEL PACIENTE)
# ===================================================================
with tab_reservar:
    # Mostrar ticket recién generado si existe
    if st.session_state["last_booked_ticket"] is not None:
        tk = st.session_state["last_booked_ticket"]
        t_nom = "MAÑANA (08:00 - 12:00)" if tk["shift_period"] == "M" else "TARDE (14:00 - 18:00)"
        st.success(f"✅ **¡Tu cita fue registrada exitosamente!** Código de Ticket Oficial: **`{tk['ticket_code']}`**")
        with st.container(border=True):
            c_tk1, c_tk2, c_tk3 = st.columns([2, 2, 1.3])
            with c_tk1:
                st.markdown(f"**🎫 Ticket:** `{tk['ticket_code']}`")
                st.markdown(f"**👤 Paciente:** {tk['patient_names']} (DNI: `{tk['patient_dni']}`)")
                st.markdown(f"**🛡️ Seguro:** {tk['insurance_type']} | **📞 Teléfono:** {tk.get('patient_phone') or '-'}")
            with c_tk2:
                st.markdown(f"**📅 Fecha de Atención:** `{tk['appointment_date']}`")
                st.markdown(f"**⏰ Horario Asignado:** **`{tk['slot_time']}` (Cupo #{tk['slot_number']} de 16 - Turno {t_nom})**")
                st.markdown(f"**👨‍⚕️ Médico Cirujano:** {tk['doctor_name']} (CMP: `{tk['doctor_cmp']}`)")
            with c_tk3:
                html_voucher = generate_ticket_html(tk)
                st.download_button(
                    "📥 Descargar Ticket Imprimible (.html)",
                    data=html_voucher.encode("utf-8"),
                    file_name=f"Ticket_{tk['ticket_code']}.html",
                    mime="text/html",
                    use_container_width=True,
                    type="primary",
                )
                if st.button("✖️ Cerrar Aviso", use_container_width=True):
                    st.session_state["last_booked_ticket"] = None
                    st.rerun()

    st.subheader(f"📌 Paso 1: Selecciona la Fecha de Atención ({meses_nombres[selected_month]} {selected_year})")

    # Determinar día predeterminado (hoy si pertenece al mes seleccionado y no es domingo, o el primer día hábil)
    working_days_meta = [d for d in days_list if not d["is_sunday"]]
    default_day_idx = 0
    if selected_year == today.year and selected_month == today.month:
        for idx, dm in enumerate(working_days_meta):
            if dm["day"] >= today.day:
                default_day_idx = idx
                break

    col_sel_day, col_summary_day = st.columns([1.4, 2.6])
    with col_sel_day:
        selected_day_num = st.selectbox(
            "Elige el día de tu cita médica (Lunes a Sábado):",
            [d["day"] for d in working_days_meta],
            index=default_day_idx,
            format_func=lambda dnum: next(
                (
                    f"{dm['weekday_name']} {dm['day']:02d}/{selected_month:02d}/{selected_year}"
                    + (f" — 🎉 {dm['holiday_name']}" if dm["is_holiday"] else "")
                    for dm in working_days_meta
                    if dm["day"] == dnum
                ),
                str(dnum),
            ),
        )

    day_consultorios = get_day_ce_consultorios(selected_year, selected_month, selected_day_num)
    m_info = day_consultorios["morning"]
    t_info = day_consultorios["afternoon"]

    with col_summary_day:
        k1, k2, k3 = st.columns(3)
        k1.metric(
            "🌅 Cupos Libres Mañana (08:00-12:00)",
            f"{m_info['available_slots']} / 16",
            delta="Consultorio Activo" if m_info["is_enabled"] else "Sin turno CE",
            delta_color="normal" if m_info["available_slots"] > 0 else "off",
        )
        k2.metric(
            "🌇 Cupos Libres Tarde (14:00-18:00)",
            f"{t_info['available_slots']} / 16",
            delta="Consultorio Activo" if t_info["is_enabled"] else "Sin turno CE",
            delta_color="normal" if t_info["available_slots"] > 0 else "off",
        )
        k3.metric(
            "🏥 Total Cupos Disponibles en el Día",
            f"{day_consultorios['total_available_day']} / 32",
            delta=f"{day_consultorios['total_occupied_day']} reservados",
        )

    st.markdown("#### 👨‍⚕️ Médicos Programados en Consultorio de Medicina (`CE`) para este día")
    col_m_card, col_t_card = st.columns(2)

    with col_m_card:
        if m_info["is_enabled"] and m_info["titular_doctor"]:
            doc_m = m_info["titular_doctor"]
            extra_m = (
                f"<br><small style='color:#475569;'>También en rol CE Mañana: "
                + ", ".join(d["doctor_name"] for d in m_info["all_ce_doctors"][1:])
                + "</small>"
                if len(m_info["all_ce_doctors"]) > 1
                else ""
            )
            st.markdown(
                f"""
                <div class="consultorio-card consultorio-morning">
                    <div style="font-size:13px; font-weight:700; color:#0284c7; text-transform:uppercase;">
                        🌅 Consultorio de Medicina — Turno Mañana (08:00 a 12:00 hrs)
                    </div>
                    <div style="font-size:17px; font-weight:800; color:#0f172a; margin-top:4px;">
                        👨‍⚕️ {doc_m['doctor_name']}
                    </div>
                    <div style="font-size:13px; color:#334155; margin-top:2px;">
                        CMP: <strong>{doc_m['doctor_cmp']}</strong> &nbsp;|&nbsp; Condición: <strong>{doc_m['condition']}</strong> &nbsp;|&nbsp;
                        Turno en Rol: <code>{doc_m['raw_shift']} ({doc_m['raw_service']})</code>
                        {extra_m}
                    </div>
                    <div style="margin-top:8px; font-size:13px; font-weight:700; color:#166534;">
                        ✅ Cupos Disponibles: {m_info['available_slots']} de 16 (Intervalos de 15 min: 08:00 - 12:00)
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.progress(m_info["occupied_slots"] / 16.0, text=f"Ocupación Mañana: {m_info['occupied_slots']}/16 cupos reservados")
        else:
            st.markdown(
                """
                <div class="consultorio-card consultorio-disabled">
                    <div style="font-size:13px; font-weight:700; text-transform:uppercase;">
                        🌅 Consultorio de Medicina — Turno Mañana (08:00 a 12:00 hrs)
                    </div>
                    <div style="font-size:15px; font-weight:700; margin-top:6px;">
                        ⚠️ Sin programación de Consultorio Externo (CE) en la Mañana según el rol de Turnos-IPRESS
                    </div>
                    <div style="font-size:12.5px; margin-top:4px;">
                        En esta fecha los médicos del turno mañana están asignados a Tópico de Urgencias (TU) o Telesalud (TS).
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with col_t_card:
        if t_info["is_enabled"] and t_info["titular_doctor"]:
            doc_t = t_info["titular_doctor"]
            extra_t = (
                f"<br><small style='color:#475569;'>También en rol CE Tarde: "
                + ", ".join(d["doctor_name"] for d in t_info["all_ce_doctors"][1:])
                + "</small>"
                if len(t_info["all_ce_doctors"]) > 1
                else ""
            )
            st.markdown(
                f"""
                <div class="consultorio-card consultorio-afternoon">
                    <div style="font-size:13px; font-weight:700; color:#059669; text-transform:uppercase;">
                        🌇 Consultorio de Medicina — Turno Tarde (14:00 a 18:00 hrs)
                    </div>
                    <div style="font-size:17px; font-weight:800; color:#0f172a; margin-top:4px;">
                        👩‍⚕️ {doc_t['doctor_name']}
                    </div>
                    <div style="font-size:13px; color:#334155; margin-top:2px;">
                        CMP: <strong>{doc_t['doctor_cmp']}</strong> &nbsp;|&nbsp; Condición: <strong>{doc_t['condition']}</strong> &nbsp;|&nbsp;
                        Turno en Rol: <code>{doc_t['raw_shift']} ({doc_t['raw_service']})</code>
                        {extra_t}
                    </div>
                    <div style="margin-top:8px; font-size:13px; font-weight:700; color:#166534;">
                        ✅ Cupos Disponibles: {t_info['available_slots']} de 16 (Intervalos de 15 min: 14:00 - 18:00)
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.progress(t_info["occupied_slots"] / 16.0, text=f"Ocupación Tarde: {t_info['occupied_slots']}/16 cupos reservados")
        else:
            st.markdown(
                """
                <div class="consultorio-card consultorio-disabled">
                    <div style="font-size:13px; font-weight:700; text-transform:uppercase;">
                        🌇 Consultorio de Medicina — Turno Tarde (14:00 a 18:00 hrs)
                    </div>
                    <div style="font-size:15px; font-weight:700; margin-top:6px;">
                        ⚠️ Sin programación de Consultorio Externo (CE) en la Tarde según el rol de Turnos-IPRESS
                    </div>
                    <div style="font-size:12.5px; margin-top:4px;">
                        En esta fecha no hay médico asignado al servicio CE en el Turno Tarde. Por favor elige el Turno Mañana u otra fecha.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("---")
    st.subheader("📝 Paso 2: Elige tu Turno, Horario (cada 15 min) y Registra tus Datos")

    # Determinar turnos habilitados con cupos libres
    available_shifts = []
    if m_info["is_enabled"] and m_info["available_slots"] > 0:
        available_shifts.append("M")
    if t_info["is_enabled"] and t_info["available_slots"] > 0:
        available_shifts.append("T")

    if not available_shifts:
        st.warning(
            "No quedan cupos disponibles en Consultorio de Medicina (CE) para la fecha seleccionada. "
            "Por favor selecciona otro día en el Paso 1."
        )
    else:
        col_grid, col_form = st.columns([1.35, 1.65])

        with col_grid:
            chosen_shift = st.radio(
                "Selecciona el Consultorio / Turno de Atención:",
                available_shifts,
                format_func=lambda sp: (
                    f"🌅 Turno Mañana (08:00 - 12:00) — {m_info['titular_doctor']['doctor_name']} ({m_info['available_slots']} libres)"
                    if sp == "M"
                    else f"🌇 Turno Tarde (14:00 - 18:00) — {t_info['titular_doctor']['doctor_name']} ({t_info['available_slots']} libres)"
                ),
                horizontal=False,
            )

            active_shift_info = m_info if chosen_shift == "M" else t_info
            st.markdown(
                f"** Tablero de 16 Cupos ({active_shift_info['hours_range']} — cada 15 min):**"
            )

            # Mostrar grilla 4x4 de los 16 cupos
            grid_cols = st.columns(4)
            for idx, sl in enumerate(active_shift_info["slots"]):
                with grid_cols[idx % 4]:
                    if sl["is_available"]:
                        st.markdown(
                            f"""<div class="slot-pill-free">
                                Cupo #{sl['slot_number']:02d}<br>{sl['label']}<br>🟢 LIBRE
                            </div>""",
                            unsafe_allow_html=True,
                        )
                    else:
                        st.markdown(
                            f"""<div class="slot-pill-busy">
                                Cupo #{sl['slot_number']:02d}<br>{sl['label']}<br>🔴 OCUPADO
                            </div>""",
                            unsafe_allow_html=True,
                        )

        with col_form:
            free_slots_list = [s for s in active_shift_info["slots"] if s["is_available"]]
            with st.form("booking_patient_form", clear_on_submit=False):
                st.markdown(
                    f"#### 🎫 Formulario de Reserva — Turno {active_shift_info['shift_name']} "
                    f"({day_consultorios['date_iso']})"
                )
                st.caption(
                    f"Médico Asignado: **{active_shift_info['titular_doctor']['doctor_name']}** "
                    f"(CMP: {active_shift_info['titular_doctor']['doctor_cmp']})"
                )

                selected_slot_num = st.selectbox(
                    "⏰ Selecciona tu Cupo y Horario (bloques de 15 minutos):",
                    [s["slot_number"] for s in free_slots_list],
                    format_func=lambda snum: next(
                        (
                            f"Cupo #{s['slot_number']:02d} — Horario: {s['label']} hrs"
                            for s in free_slots_list
                            if s["slot_number"] == snum
                        ),
                        str(snum),
                    ),
                )

                f_col1, f_col2 = st.columns([1.2, 1])
                with f_col1:
                    inp_dni = st.text_input(
                        "🪪 DNI / Documento de Identidad (*):",
                        max_chars=12,
                        placeholder="Ej. 45892134",
                    )
                with f_col2:
                    inp_age = st.number_input("🎂 Edad (años):", min_value=0, max_value=120, value=30, step=1)

                inp_names = st.text_input(
                    "👤 Apellidos y Nombres Completos del Paciente (*):",
                    placeholder="Ej. PEREZ GOMEZ, JUAN CARLOS",
                )

                f_col3, f_col4 = st.columns(2)
                with f_col3:
                    inp_insurance = st.selectbox(
                        "🛡️ Tipo de Seguro / Financiamiento:",
                        ["SIS (Seguro Integral de Salud)", "Pagante / Particular", "Exonerado / Otro"],
                    )
                with f_col4:
                    inp_phone = st.text_input(
                        "📱 Teléfono / Celular (WhatsApp):",
                        max_chars=15,
                        placeholder="Ej. 987654321",
                    )

                inp_reason = st.text_input(
                    "🩺 Motivo de Consulta Médica:",
                    placeholder="Ej. Control médico general, cefalea, hipertensión, evaluación...",
                )

                submit_book = st.form_submit_button(
                    "✅ Confirmar y Generar Ticket de Cita",
                    use_container_width=True,
                    type="primary",
                )

                if submit_book:
                    ok, msg, ticket_data = book_appointment(
                        year=selected_year,
                        month=selected_month,
                        day=selected_day_num,
                        shift_period=chosen_shift,
                        slot_number=selected_slot_num,
                        patient_dni=inp_dni,
                        patient_names=inp_names,
                        patient_age=int(inp_age),
                        patient_phone=inp_phone,
                        insurance_type=inp_insurance.split(" ")[0],
                        reason=inp_reason,
                        booked_by="Portal Web Paciente",
                    )
                    if ok:
                        st.session_state["last_booked_ticket"] = ticket_data
                        st.rerun()
                    else:
                        st.error(msg)

# ===================================================================
# PESTAÑA 2: CONSULTAR / CANCELAR MI CITA
# ===================================================================
with tab_consultar:
    st.subheader("🔍 Consulta el Estado de tu Cita o Descarga tu Comprobante")
    st.markdown(
        "Ingresa tu número de **DNI**, **Código de Ticket** (ej. `JP2-20261019-M01`) o **Apellidos** "
        "para verificar tu horario, descargar tu ticket imprimible o cancelar tu cita si no podrás asistir."
    )

    c_search1, c_search2 = st.columns([3, 1])
    with c_search1:
        query_str = st.text_input(
            "Buscar por DNI, Código de Ticket o Apellidos:",
            placeholder="Ingresa tu DNI (ej. 45892134) o Ticket (ej. JP2-20261019-M01)...",
        )
    with c_search2:
        st.write("")
        st.write("")
        btn_search = st.button("🔎 Buscar Cita", use_container_width=True, type="primary")

    if query_str.strip():
        results = search_appointments(query_str)
        if not results:
            st.info("No se encontraron citas registradas con el dato ingresado.")
        else:
            st.markdown(f"Se encontraron **{len(results)}** cita(s) asociada(s):")
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

                with st.container(border=True):
                    r1, r2, r3 = st.columns([2, 2, 1.3])
                    with r1:
                        st.markdown(f"**🎫 Ticket:** `{app_item['ticket_code']}` — **{status_icon}**")
                        st.markdown(f"**👤 Paciente:** {app_item['patient_names']} (DNI: `{app_item['patient_dni']}`)")
                        st.markdown(f"**🛡️ Seguro:** {app_item['insurance_type']} | **Motivo:** {app_item.get('reason') or '-'}")
                    with r2:
                        st.markdown(f"**📅 Fecha:** `{app_item['appointment_date']}` ({t_label})")
                        st.markdown(f"**⏰ Horario:** **`{app_item['slot_time']}` — Cupo #{app_item['slot_number']:02d} de 16**")
                        st.markdown(f"**👨‍⚕️ Médico:** {app_item['doctor_name']} (CMP: `{app_item['doctor_cmp']}`)")
                    with r3:
                        voucher_html = generate_ticket_html(app_item)
                        st.download_button(
                            "📥 Descargar Ticket",
                            data=voucher_html.encode("utf-8"),
                            file_name=f"Ticket_{app_item['ticket_code']}.html",
                            mime="text/html",
                            key=f"dl_{app_item['id']}",
                            use_container_width=True,
                        )
                        if app_item["status"] == "reservada":
                            if st.button(
                                "🗑️ Cancelar Cita",
                                key=f"canc_{app_item['id']}",
                                use_container_width=True,
                            ):
                                ok_c, msg_c = cancel_appointment(
                                    app_item["ticket_code"],
                                    patient_dni=app_item["patient_dni"],
                                    cancelled_by=f"Paciente DNI {app_item['patient_dni']}",
                                )
                                if ok_c:
                                    st.success(msg_c)
                                    st.rerun()
                                else:
                                    st.error(msg_c)

# ===================================================================
# PESTAÑA 3: ROL MENSUAL DE MEDICINA (SINCRONIZADO CON TURNOS-IPRESS)
# ===================================================================
with tab_rol:
    st.subheader(
        f"🗓️ Programación Mensual de Consultorios de Medicina (`CE`) — {meses_nombres[selected_month]} {selected_year}"
    )
    st.markdown(
        "Esta tabla muestra día por día qué médico del **C.S. Su Santidad Juan Pablo II** atiende el "
        "**Consultorio de Medicina en la Mañana (08:00 a 12:00, 16 cupos)** y en la **Tarde (14:00 a 18:00, 16 cupos)**, "
        "de acuerdo con la programación de **[turnos-ipress.streamlit.app](https://turnos-ipress.streamlit.app/)**."
    )

    month_overview = get_month_ce_overview(selected_year, selected_month)
    table_rows = []
    for item in month_overview:
        dm = item["day_meta"]
        if dm["is_sunday"]:
            m_doc_txt = "🔒 Domingo (Sin consulta externa)"
            t_doc_txt = "🔒 Domingo (Sin consulta externa)"
        else:
            m_doc = item["morning"]["titular_doctor"]
            t_doc = item["afternoon"]["titular_doctor"]
            m_doc_txt = (
                f"{m_doc['doctor_name']} (CMP {m_doc['doctor_cmp']})"
                if m_doc
                else "⚠️ Sin médico en CE Mañana"
            )
            t_doc_txt = (
                f"{t_doc['doctor_name']} (CMP {t_doc['doctor_cmp']})"
                if t_doc
                else "⚠️ Sin médico en CE Tarde"
            )

        table_rows.append(
            {
                "Fecha": item["date_iso"],
                "Día": dm["weekday_name"] + (f" ({dm['holiday_name']})" if dm["is_holiday"] else ""),
                "🌅 Médico CE Mañana (08:00 - 12:00)": m_doc_txt,
                "Cupos Libres Mañana": f"{item['morning']['available_slots']} / 16" if item["morning"]["is_enabled"] else "0 / 16",
                "🌇 Médico CE Tarde (14:00 - 18:00)": t_doc_txt,
                "Cupos Libres Tarde": f"{item['afternoon']['available_slots']} / 16" if item["afternoon"]["is_enabled"] else "0 / 16",
            }
        )

    st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True, height=460)

    with st.expander("🔎 Ver Matriz Completa de Turnos Asistenciales (TU, CE, TS, 12h, DES) importada de Turnos-IPRESS"):
        with get_connection() as conn:
            all_shifts_m = conn.execute(
                """
                SELECT day, shift_period, service, doctor_id
                FROM shifts
                WHERE year = ? AND month = ?
                """,
                (selected_year, selected_month),
            ).fetchall()

        shift_lookup = {}
        for s in all_shifts_m:
            shift_lookup[(s["doctor_id"], s["day"])] = f"{s['shift_period']}:{s['service']}"

        matrix_rows = []
        for doc in doctors_only:
            row_dict = {
                "Médico Cirujano": doc["name"],
                "Condición": doc["condition"],
                "CMP": doc["cmp"],
            }
            for dm in days_list:
                dnum = dm["day"]
                col_key = f"{dm['weekday_short']} {dnum:02d}"
                val = shift_lookup.get((doc["id"], dnum), "-")
                row_dict[col_key] = val
            matrix_rows.append(row_dict)

        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, hide_index=True)

# ===================================================================
# PESTAÑA 4: PANEL DE ADMISIÓN, PADRÓN DIARIO Y CONTROL DE CONSULTORIOS
# ===================================================================
with tab_admision:
    st.subheader("🏥 Panel de Admisión, Padrón Diario de 16 Cupos y Gestión de Consultorios")

    col_adm_day, col_adm_export = st.columns([2, 1.2])
    with col_adm_day:
        adm_day_num = st.selectbox(
            "Seleccionar Día para ver el Padrón de Citas (16 cupos Mañana + 16 cupos Tarde):",
            [d["day"] for d in days_list],
            index=selected_day_num - 1 if (1 <= selected_day_num <= len(days_list)) else 0,
            format_func=lambda dnum: (
                f"{days_list[dnum - 1]['weekday_name']} {dnum:02d}/{selected_month:02d}/{selected_year}"
            ),
            key="adm_day_selector",
        )
    with col_adm_export:
        st.write("")
        st.write("")
        excel_bytes = generate_daily_padron_excel(selected_year, selected_month, adm_day_num)
        st.download_button(
            "📊 Descargar Padrón Diario y Rol en Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Padron_Citas_Medicina_CS_JuanPabloII_{selected_year}_{selected_month:02d}_{adm_day_num:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    adm_day_data = get_day_ce_consultorios(selected_year, selected_month, adm_day_num)

    col_pad_m, col_pad_t = st.columns(2)

    for col_widget, shift_key, icon_str in [
        (col_pad_m, "morning", "🌅"),
        (col_pad_t, "afternoon", "🌇"),
    ]:
        with col_widget:
            s_data = adm_day_data[shift_key]
            doc_t = s_data["titular_doctor"]
            st.markdown(f"#### {icon_str} {s_data['consultorio_title']}")
            if doc_t:
                st.caption(
                    f"**Médico Responsable:** {doc_t['doctor_name']} (CMP: `{doc_t['doctor_cmp']}`) | "
                    f"**Ocupados:** `{s_data['occupied_slots']}/16` | **Libres:** `{s_data['available_slots']}/16`"
                )
            else:
                st.caption("⚠️ Sin médico programado en Consultorio de Medicina (CE) para este turno.")

            padron_rows = []
            for sl in s_data["slots"]:
                ap = sl["appointment"]
                padron_rows.append(
                    {
                        "Cupo": f"#{sl['slot_number']:02d}",
                        "Horario": sl["label"],
                        "Estado": ap["status"].upper() if ap else "🟢 LIBRE",
                        "Ticket": ap["ticket_code"] if ap else "-",
                        "DNI": ap["patient_dni"] if ap else "-",
                        "Paciente": ap["patient_names"] if ap else "(Disponible)",
                        "Edad": ap["patient_age"] if (ap and ap.get("patient_age")) else "-",
                        "Seguro": ap["insurance_type"] if ap else "-",
                        "Teléfono": ap["patient_phone"] if (ap and ap.get("patient_phone")) else "-",
                    }
                )
            st.dataframe(pd.DataFrame(padron_rows), use_container_width=True, hide_index=True, height=580)

    st.markdown("---")
    if st.session_state["staff_user"] is None:
        st.info(
            "🔒 **Para registrar asistencia de pacientes ('Atendida' / 'No Asistió'), asignar médico titular en días con doble turno CE o subir una base de datos de Turnos-IPRESS**, inicia sesión en el panel izquierdo (**Acceso Admisión / Médicos**, PIN Admin: `admin2026`)."
        )
    else:
        staff = st.session_state["staff_user"]
        st.success(f"🔓 Controles Administrativos habilitados para: **{staff['name']}**")

        adm_sub1, adm_sub2, adm_sub3 = st.tabs(
            [
                "✅ Control de Asistencia de Citas del Día",
                "👨‍⚕️ Asignación / Contingencia de Consultorio CE",
                "🔄 Importar Base de Datos de Turnos-IPRESS (.db)",
            ]
        )

        with adm_sub1:
            booked_today = []
            for sk in ("morning", "afternoon"):
                for sl in adm_day_data[sk]["slots"]:
                    if sl["appointment"]:
                        booked_today.append(sl["appointment"])

            if not booked_today:
                st.caption("No hay citas reservadas en esta fecha para actualizar estado.")
            else:
                c_up1, c_up2, c_up3 = st.columns([2.2, 1.2, 1])
                with c_up1:
                    sel_app_id = st.selectbox(
                        "Seleccionar Cita del Día:",
                        [a["id"] for a in booked_today],
                        format_func=lambda aid: next(
                            (
                                f"[{a['shift_period']} Cupo #{a['slot_number']:02d} {a['slot_time']}] "
                                f"{a['patient_names']} (DNI {a['patient_dni']}) — Estado: {a['status'].upper()}"
                                for a in booked_today
                                if a["id"] == aid
                            ),
                            str(aid),
                        ),
                    )
                with c_up2:
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
                with c_up3:
                    st.write("")
                    st.write("")
                    if st.button("Guardar Estado", use_container_width=True, type="primary"):
                        ok_u, msg_u = update_appointment_status(sel_app_id, new_st, staff["name"])
                        if ok_u:
                            st.success(msg_u)
                            st.rerun()
                        else:
                            st.error(msg_u)

        with adm_sub2:
            st.markdown(
                "Si en un día el rol de `turnos-ipress` tiene más de un médico en `CE` o hubo una contingencia/reemplazo, "
                "aquí puedes definir qué médico figura como titular del **Consultorio de Medicina Mañana (16 cupos)** o **Tarde (16 cupos)**."
            )
            ov_col1, ov_col2, ov_col3, ov_col4 = st.columns([1, 1.6, 1, 1.2])
            with ov_col1:
                ov_shift = st.selectbox(
                    "Turno:",
                    ["M", "T"],
                    format_func=lambda s: "🌅 Mañana (08:00-12:00)" if s == "M" else "🌇 Tarde (14:00-18:00)",
                )
            with ov_col2:
                ov_doc_id = st.selectbox(
                    "Médico Cirujano Titular:",
                    [d["id"] for d in doctors_only],
                    format_func=lambda did: next(
                        (f"{d['name']} (CMP {d['cmp']})" for d in doctors_only if d["id"] == did),
                        str(did),
                    ),
                )
            with ov_col3:
                ov_enabled = st.checkbox("Consultorio Habilitado (16 cupos)", value=True)
            with ov_col4:
                ov_notes = st.text_input("Observación:", placeholder="Ej. Cobertura / Reemplazo")

            b_col1, b_col2 = st.columns(2)
            with b_col1:
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
            with b_col2:
                if st.button("🔄 Restaurar Según Rol Original de Turnos-IPRESS", use_container_width=True):
                    clear_consultorio_override(
                        selected_year, selected_month, adm_day_num, ov_shift, staff["name"]
                    )
                    st.success("Restaurado al rol automático de Turnos-IPRESS.")
                    st.rerun()

        with adm_sub3:
            st.markdown(
                "Puedes sincronizar directamente desde el archivo local `ipress_turnos.db` o subir el archivo `ipress_turnos.db` "
                "descargado desde el proyecto **Turnos-IPRESS**:"
            )
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
