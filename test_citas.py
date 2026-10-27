"""
test_citas.py
Pruebas automatizadas del Sistema de Citas por Internet - Servicio de Medicina
del C.S. Su Santidad Juan Pablo II.
"""

from database import init_db, get_connection
from schedule_sync import (
    MORNING_SLOTS,
    AFTERNOON_SLOTS,
    sync_from_external_turnos_db,
    get_day_ce_consultorios,
    validate_booking_data,
    book_appointment,
    cancel_appointment,
    search_appointments,
)
from exporter import (
    generate_ticket_html,
    generate_ticket_image_png,
    generate_whatsapp_graphic_share_html,
    generate_whatsapp_info,
    generate_daily_padron_excel,
    generate_admission_datasheet_csv,
)


def run_tests():
    init_db()
    sync_res = sync_from_external_turnos_db()
    print("Sync Result:", sync_res["message"])

    # Limpiar datos previos de prueba si existieran
    with get_connection() as conn:
        conn.execute("DELETE FROM appointments WHERE patient_dni IN ('08123456', '001598745', '44556677')")
        conn.commit()

    assert len(MORNING_SLOTS) == 16, "Mañana debe tener 16 cupos"
    assert len(AFTERNOON_SLOTS) == 16, "Tarde debe tener 16 cupos"

    # a) Probar rechazo de DNI con letras o con menos de 8 dígitos
    ok_bad_dni, msg_bad_dni, _ = validate_booking_data(
        2026, 11, 16, "M", 1, "DNI", "1234567", "RAMIREZ SOTO, CARLOS", 40, "987654321", "SIS", "Control"
    )
    assert ok_bad_dni is False and "8 dígitos" in msg_bad_dni

    ok_alpha_dni, _, _ = validate_booking_data(
        2026, 11, 16, "M", 1, "DNI", "0812345A", "RAMIREZ SOTO, CARLOS", 40, "987654321", "SIS", "Control"
    )
    assert ok_alpha_dni is False

    # b) Probar rechazo de celular que no tenga exactamente 9 dígitos
    ok_bad_phone, msg_bad_phone, _ = validate_booking_data(
        2026, 11, 16, "M", 1, "DNI", "08123456", "RAMIREZ SOTO, CARLOS", 40, "98765432", "SIS", "Control"
    )
    assert ok_bad_phone is False and "9 dígitos" in msg_bad_phone

    # Probar rechazo cuando la fecha es pasada (ej. 05/01/2026)
    ok_past, msg_past, _ = validate_booking_data(
        2026, 1, 5, "M", 1, "DNI", "08123456", "RAMIREZ SOTO, CARLOS", 40, "987654321", "SIS", "Control médico"
    )
    assert ok_past is False and "FECHA PASADA" in msg_past

    # Probar rechazo cuando falta llenar algún campo obligatorio (ej. Edad vacía o Motivo vacío)
    ok_inc, msg_inc, _ = validate_booking_data(
        2026, 11, 16, "M", 1, "DNI", "08123456", "RAMIREZ SOTO, CARLOS", None, "987654321", "-- Selecciona tu Seguro --", ""
    )
    assert ok_inc is False and "CAMPOS INCOMPLETOS" in msg_inc

    # c) Probar validación previa (para ventana en grande) con DNI que empieza con 0 ('08123456')
    ok_pre, _, pre_payload = validate_booking_data(
        2026, 11, 16, "M", 1, "DNI", "08123456", "RAMIREZ SOTO, CARLOS", 40, "999888777", "SIS", "Control preventivo"
    )
    assert ok_pre is True
    assert pre_payload["patient_dni"] == "08123456"
    assert pre_payload["patient_doc_type"] == "DNI"

    # Confirmar la cita del Lunes 16/11/2026
    ok, msg, ticket = book_appointment(
        year=2026,
        month=11,
        day=16,
        shift_period="M",
        slot_number=1,
        patient_doc_type="DNI",
        patient_dni="08123456",
        patient_names="RAMIREZ SOTO, CARLOS",
        patient_age=40,
        patient_phone="999888777",
        insurance_type="SIS",
        reason="Control preventivo",
        booked_by="Test Suite",
    )
    assert ok is True, f"Error al reservar: {msg}"
    assert ticket["patient_dni"] == "08123456"
    print("Cita reservada con 0 inicial:", ticket["ticket_code"], "DNI:", ticket["patient_dni"])

    # e.1) Probar restricción de 1 cita por semana con el MISMO DNI en otro día de la misma semana (Miércoles 18/11/2026)
    ok_wk_dni, msg_wk_dni, _ = book_appointment(
        year=2026,
        month=11,
        day=18,
        shift_period="M",
        slot_number=3,
        patient_doc_type="DNI",
        patient_dni="08123456",
        patient_names="RAMIREZ SOTO, CARLOS",
        patient_age=40,
        patient_phone="999888777",
        insurance_type="SIS",
        reason="Segunda consulta en la semana",
        booked_by="Test Suite",
    )
    assert ok_wk_dni is False, "No debe permitir 2 citas en la misma semana con el mismo documento"
    assert "Servicio de Admisión de manera presencial" in msg_wk_dni
    print("Bloqueo semanal por DNI verificado correctamente.")

    # e.2) Probar restricción de 1 cita por semana por COINCIDENCIA EN EL NOMBRE (misma semana, Viernes 20/11/2026)
    ok_wk_name, msg_wk_name, _ = book_appointment(
        year=2026,
        month=11,
        day=20,
        shift_period="T",
        slot_number=2,
        patient_doc_type="DNI",
        patient_dni="44556677",
        patient_names="CARLOS RAMIREZ SOTO",
        patient_age=40,
        patient_phone="999888777",
        insurance_type="SIS",
        reason="Intento con otro doc pero mismo nombre",
        booked_by="Test Suite",
    )
    assert ok_wk_name is False, "No debe permitir 2 citas en la misma semana cuando coincide el nombre del paciente"
    assert "Servicio de Admisión de manera presencial" in msg_wk_name
    print("Bloqueo semanal por coincidencia de nombre verificado correctamente.")

    # Probar Carné de Extranjería (CE) numérico en otra persona
    ok_ce, msg_ce, ticket_ce = book_appointment(
        year=2026,
        month=11,
        day=16,
        shift_period="T",
        slot_number=2,
        patient_doc_type="CE",
        patient_dni="001598745",
        patient_names="GONZALEZ PEREZ, MARIA ELENA",
        patient_age=29,
        patient_phone="912345678",
        insurance_type="SIS",
        reason="Control general",
        booked_by="Test Suite",
    )
    assert ok_ce is True, f"Error al reservar con CE: {msg_ce}"
    assert ticket_ce["patient_doc_type"] == "CE"

    # d) Verificar Hoja de Datos de Admisión en Excel (.xlsx), CSV (.csv) y Ticket Gráfico (.PNG)
    wa_info = generate_whatsapp_info(ticket)
    assert wa_info["phone_normalized"] == "51999888777"
    html_str = generate_ticket_html(ticket)
    assert "08123456" in html_str

    png_bytes = generate_ticket_image_png(ticket)
    assert png_bytes.startswith(b"\x89PNG"), "Debe generar una imagen PNG válida"
    assert len(png_bytes) > 5000, "El ticket gráfico PNG debe tener contenido visual completo"
    wa_share_comp = generate_whatsapp_graphic_share_html(ticket, png_bytes=png_bytes)
    assert "shareGraphicTicket" in wa_share_comp
    print("Generacion de Ticket Grafico (.PNG) para WhatsApp verificada:", len(png_bytes), "bytes")

    excel_io = generate_daily_padron_excel(2026, 11, 16)
    assert len(excel_io.getvalue()) > 1500
    csv_bytes = generate_admission_datasheet_csv(2026, 11, 16, only_booked=True)
    csv_text = csv_bytes.decode("utf-8")
    assert "08123456" in csv_text and "001598745" in csv_text
    print("Exportacion de Hoja de Datos de Admision (.xlsx y .csv) verificada.")

    # Limpiar citas de prueba
    with get_connection() as conn:
        conn.execute("DELETE FROM appointments WHERE patient_dni IN ('08123456', '001598745', '44556677')")
        conn.execute("DELETE FROM audit_logs WHERE user_name = 'Test Suite'")
        conn.commit()

    print("[OK] TODAS LAS CONDICIONES (a, b, c, d, e) PASARON EXITOSAMENTE.")


if __name__ == "__main__":
    run_tests()
