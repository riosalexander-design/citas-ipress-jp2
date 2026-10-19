"""
test_citas.py
Pruebas automatizadas del Sistema de Citas por Internet - Servicio de Medicina
del C.S. Su Santidad Juan Pablo II.
"""

from database import init_db
from schedule_sync import (
    MORNING_SLOTS,
    AFTERNOON_SLOTS,
    sync_from_external_turnos_db,
    get_day_ce_consultorios,
    book_appointment,
    cancel_appointment,
    search_appointments,
)
from exporter import generate_ticket_html, generate_daily_padron_excel


def run_tests():
    init_db()
    sync_res = sync_from_external_turnos_db()
    print("Sync Result:", sync_res["message"])

    assert len(MORNING_SLOTS) == 16, "Mañana debe tener 16 cupos"
    assert MORNING_SLOTS[0]["label"] == "08:00 - 08:15"
    assert MORNING_SLOTS[-1]["label"] == "11:45 - 12:00"

    assert len(AFTERNOON_SLOTS) == 16, "Tarde debe tener 16 cupos"
    assert AFTERNOON_SLOTS[0]["label"] == "14:00 - 14:15"
    assert AFTERNOON_SLOTS[-1]["label"] == "17:45 - 18:00"

    # Verificar el día Lunes 19 de Octubre de 2026 (tiene CE en Mañana y CE en Tarde)
    day_info = get_day_ce_consultorios(2026, 10, 19)
    assert day_info["morning"]["is_enabled"] is True, "Lunes 19/10/2026 debe tener CE Mañana habilitado"
    assert day_info["afternoon"]["is_enabled"] is True, "Lunes 19/10/2026 debe tener CE Tarde habilitado"
    print(
        f"19/10/2026 -> Mañana: {day_info['morning']['titular_doctor']['doctor_name']} | "
        f"Tarde: {day_info['afternoon']['titular_doctor']['doctor_name']}"
    )

    # Probar reserva de cita en Mañana Cupo #1 (08:00 - 08:15)
    ok, msg, ticket = book_appointment(
        year=2026,
        month=10,
        day=19,
        shift_period="M",
        slot_number=1,
        patient_dni="99887766",
        patient_names="PACIENTE DE PRUEBA UNO",
        patient_age=35,
        patient_phone="999888777",
        insurance_type="SIS",
        reason="Control preventivo",
        booked_by="Test Suite",
    )
    assert ok is True, f"Error al reservar: {msg}"
    assert ticket["slot_time"] == "08:00 - 08:15"
    print("Cita reservada:", ticket["ticket_code"])

    # Verificar bloqueo de doble cita el mismo día para el mismo DNI
    ok_dup, msg_dup, _ = book_appointment(
        year=2026,
        month=10,
        day=19,
        shift_period="T",
        slot_number=2,
        patient_dni="99887766",
        patient_names="PACIENTE DE PRUEBA UNO",
        patient_age=35,
        patient_phone="999888777",
        insurance_type="SIS",
        reason="Otra consulta",
        booked_by="Test Suite",
    )
    assert ok_dup is False, "No debe permitir dos citas el mismo día para el mismo DNI"

    # Verificar generación de Ticket HTML y Excel
    html_str = generate_ticket_html(ticket)
    assert "C.S. SU SANTIDAD JUAN PABLO II" in html_str
    excel_io = generate_daily_padron_excel(2026, 10, 19)
    assert len(excel_io.getvalue()) > 1000

    # Buscar y luego cancelar la cita de prueba para dejar todos los cupos limpios
    found = search_appointments("99887766")
    assert len(found) >= 1

    ok_c, msg_c = cancel_appointment(ticket["ticket_code"], "99887766", "Test Suite")
    assert ok_c is True, f"Error al cancelar: {msg_c}"
    print("Cancelacion exitosa:", msg_c)

    # Limpiar registro de prueba para dejar la base de datos impecable
    from database import get_connection
    with get_connection() as conn:
        conn.execute("DELETE FROM appointments WHERE patient_dni = '99887766'")
        conn.execute("DELETE FROM audit_logs WHERE user_name = 'Test Suite'")
        conn.commit()

    print("[OK] TODAS LAS PRUEBAS PASARON EXITOSAMENTE.")


if __name__ == "__main__":
    run_tests()
