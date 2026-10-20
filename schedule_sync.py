"""
schedule_sync.py
Motor de sincronización con la programación médica de https://turnos-ipress.streamlit.app/
y gestión de los 16 cupos diarios (Mañana 08:00-12:00 y Tarde 14:00-18:00, cada 15 min)
para el Consultorio Externo de Medicina (CE) del C.S. Su Santidad Juan Pablo II.
"""

import os
import json
import sqlite3
import unicodedata
from datetime import datetime, date, timedelta, timezone
from database import get_connection, find_external_turnos_db, log_audit
from holidays_pe import get_month_info

PERU_TZ = timezone(timedelta(hours=-5))


def get_peru_today() -> date:
    """Retorna la fecha actual en hora oficial del Perú (UTC-5)."""
    return datetime.now(PERU_TZ).date()


MAX_SLOTS_PER_SHIFT = 16

# 16 cupos escalonados cada 15 minutos en Turno Mañana (08:00 a 12:00)
MORNING_SLOTS = [
    {"slot_number": 1, "start": "08:00", "end": "08:15", "label": "08:00 - 08:15"},
    {"slot_number": 2, "start": "08:15", "end": "08:30", "label": "08:15 - 08:30"},
    {"slot_number": 3, "start": "08:30", "end": "08:45", "label": "08:30 - 08:45"},
    {"slot_number": 4, "start": "08:45", "end": "09:00", "label": "08:45 - 09:00"},
    {"slot_number": 5, "start": "09:00", "end": "09:15", "label": "09:00 - 09:15"},
    {"slot_number": 6, "start": "09:15", "end": "09:30", "label": "09:15 - 09:30"},
    {"slot_number": 7, "start": "09:30", "end": "09:45", "label": "09:30 - 09:45"},
    {"slot_number": 8, "start": "09:45", "end": "10:00", "label": "09:45 - 10:00"},
    {"slot_number": 9, "start": "10:00", "end": "10:15", "label": "10:00 - 10:15"},
    {"slot_number": 10, "start": "10:15", "end": "10:30", "label": "10:15 - 10:30"},
    {"slot_number": 11, "start": "10:30", "end": "10:45", "label": "10:30 - 10:45"},
    {"slot_number": 12, "start": "10:45", "end": "11:00", "label": "10:45 - 11:00"},
    {"slot_number": 13, "start": "11:00", "end": "11:15", "label": "11:00 - 11:15"},
    {"slot_number": 14, "start": "11:15", "end": "11:30", "label": "11:15 - 11:30"},
    {"slot_number": 15, "start": "11:30", "end": "11:45", "label": "11:30 - 11:45"},
    {"slot_number": 16, "start": "11:45", "end": "12:00", "label": "11:45 - 12:00"},
]

# 16 cupos escalonados cada 15 minutos en Turno Tarde (14:00 a 18:00)
AFTERNOON_SLOTS = [
    {"slot_number": 1, "start": "14:00", "end": "14:15", "label": "14:00 - 14:15"},
    {"slot_number": 2, "start": "14:15", "end": "14:30", "label": "14:15 - 14:30"},
    {"slot_number": 3, "start": "14:30", "end": "14:45", "label": "14:30 - 14:45"},
    {"slot_number": 4, "start": "14:45", "end": "15:00", "label": "14:45 - 15:00"},
    {"slot_number": 5, "start": "15:00", "end": "15:15", "label": "15:00 - 15:15"},
    {"slot_number": 6, "start": "15:15", "end": "15:30", "label": "15:15 - 15:30"},
    {"slot_number": 7, "start": "15:30", "end": "15:45", "label": "15:30 - 15:45"},
    {"slot_number": 8, "start": "15:45", "end": "16:00", "label": "15:45 - 16:00"},
    {"slot_number": 9, "start": "16:00", "end": "16:15", "label": "16:00 - 16:15"},
    {"slot_number": 10, "start": "16:15", "end": "16:30", "label": "16:15 - 16:30"},
    {"slot_number": 11, "start": "16:30", "end": "16:45", "label": "16:30 - 16:45"},
    {"slot_number": 12, "start": "16:45", "end": "17:00", "label": "16:45 - 17:00"},
    {"slot_number": 13, "start": "17:00", "end": "17:15", "label": "17:00 - 17:15"},
    {"slot_number": 14, "start": "17:15", "end": "17:30", "label": "17:15 - 17:30"},
    {"slot_number": 15, "start": "17:30", "end": "17:45", "label": "17:30 - 17:45"},
    {"slot_number": 16, "start": "17:45", "end": "18:00", "label": "17:45 - 18:00"},
]

SERVICE_LABELS = {
    "TU": "Tópico de Urgencias (TU)",
    "CE": "Consultorio de Medicina (CE)",
    "TS": "Telesalud/TBC/MCI (TS)",
}


def get_slots_definition(shift_period: str) -> list[dict]:
    """Retorna la lista de los 16 horarios escalonados cada 15 min según el turno ('M' o 'T')."""
    return MORNING_SLOTS if shift_period == "M" else AFTERNOON_SLOTS


SEED_JSON_PATH = os.path.join(os.path.dirname(__file__), "turnos_ipress_seed.json")


def sync_from_external_turnos_db(source_db_path: str = None) -> dict:
    """
    Sincroniza los médicos y los turnos desde la base de datos de turnos-ipress (ipress_turnos.db)
    o desde el archivo semilla JSON (turnos_ipress_seed.json) en la nube (Streamlit Cloud).
    """
    ext_path = source_db_path or find_external_turnos_db()

    ext_docs = []
    ext_shifts = []
    ext_id_to_cmp = {}
    source_label = ""

    if ext_path and os.path.exists(ext_path):
        try:
            ext_conn = sqlite3.connect(ext_path, timeout=15.0)
            ext_conn.row_factory = sqlite3.Row
            ext_docs = [
                dict(r)
                for r in ext_conn.execute(
                    "SELECT id, name, condition, cmp, pin_hash, role, is_active FROM doctors"
                ).fetchall()
            ]
            ext_shifts = [
                dict(r)
                for r in ext_conn.execute(
                    "SELECT year, month, day, shift_period, service, doctor_id, is_proposal FROM shifts"
                ).fetchall()
            ]
            ext_id_to_cmp = {d["id"]: d["cmp"] for d in ext_docs}
            ext_conn.close()
            source_label = ext_path

            # Guardar semilla JSON actualizada para despliegues en Streamlit Cloud
            if source_db_path is None:
                try:
                    with open(SEED_JSON_PATH, "w", encoding="utf-8") as jf:
                        json.dump({"doctors": ext_docs, "shifts": ext_shifts}, jf, ensure_ascii=False, indent=2)
                except OSError:
                    pass
        except Exception as exc:
            return {
                "synced": False,
                "source": ext_path,
                "shifts_synced": 0,
                "message": f"Aviso al sincronizar con BD externa: {exc}",
            }
    elif os.path.exists(SEED_JSON_PATH):
        try:
            with open(SEED_JSON_PATH, "r", encoding="utf-8") as jf:
                seed_data = json.load(jf)
            ext_docs = seed_data.get("doctors", [])
            ext_shifts = seed_data.get("shifts", [])
            ext_id_to_cmp = {d["id"]: d["cmp"] for d in ext_docs}
            source_label = SEED_JSON_PATH
        except Exception as exc:
            return {
                "synced": False,
                "source": SEED_JSON_PATH,
                "shifts_synced": 0,
                "message": f"Aviso al cargar semilla JSON: {exc}",
            }
    else:
        return {
            "synced": False,
            "source": "Motor Interno Turnos-IPRESS (Standalone Cloud)",
            "shifts_synced": 0,
            "message": "Usando programación nativa compatible con turnos-ipress.streamlit.app.",
        }

    try:
        with get_connection() as conn:
            # 1. Sincronizar médicos por CMP
            for d in ext_docs:
                row = conn.execute("SELECT id FROM doctors WHERE cmp = ?", (d["cmp"],)).fetchone()
                if not row:
                    conn.execute(
                        """
                        INSERT INTO doctors (name, condition, cmp, pin_hash, role, is_active)
                        VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (d["name"], d["condition"], d["cmp"], d["pin_hash"], d["role"], d["is_active"]),
                    )
                else:
                    conn.execute(
                        """
                        UPDATE doctors
                        SET name = ?, condition = ?, pin_hash = ?, role = ?, is_active = ?
                        WHERE cmp = ?
                        """,
                        (d["name"], d["condition"], d["pin_hash"], d["role"], d["is_active"], d["cmp"]),
                    )

            local_docs = conn.execute("SELECT id, cmp FROM doctors").fetchall()
            cmp_to_local_id = {r["cmp"]: r["id"] for r in local_docs}

            # 2. Sincronizar turnos agrupados por año/mes presentes en ext_shifts
            ym_pairs = {(s["year"], s["month"]) for s in ext_shifts}
            for y, m in ym_pairs:
                conn.execute("DELETE FROM shifts WHERE year = ? AND month = ?", (y, m))

            synced_count = 0
            for s in ext_shifts:
                cmp_val = ext_id_to_cmp.get(s["doctor_id"])
                local_doc_id = cmp_to_local_id.get(cmp_val)
                if local_doc_id:
                    conn.execute(
                        """
                        INSERT OR REPLACE INTO shifts
                        (year, month, day, shift_period, service, doctor_id, is_proposal, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                        """,
                        (
                            s["year"],
                            s["month"],
                            s["day"],
                            s["shift_period"],
                            s["service"],
                            local_doc_id,
                            s["is_proposal"],
                        ),
                    )
                    synced_count += 1

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                """
                INSERT OR REPLACE INTO sync_metadata (key, value, updated_at)
                VALUES ('last_sync', ?, CURRENT_TIMESTAMP)
                """,
                (f"Sincronizado con {os.path.basename(source_label)} ({synced_count} turnos) el {now_str}",),
            )
            conn.commit()

        return {
            "synced": True,
            "source": source_label,
            "shifts_synced": synced_count,
            "message": f"Sincronizado exitosamente con turnos-ipress ({synced_count} registros de turnos).",
        }
    except Exception as exc:
        return {
            "synced": False,
            "source": source_label,
            "shifts_synced": 0,
            "message": f"Aviso al sincronizar: {exc}",
        }


def _pick_needed_morning_service(grid_day: dict, preferred_order=("TU", "CE", "TS")) -> str:
    covered = set()
    for sh in grid_day.values():
        if sh["period"] == "M":
            covered.add(sh["service"])
        elif sh["period"] == "12h":
            covered.add(sh["service"].split("+")[0] if "+" in sh["service"] else sh["service"])
    for srv in preferred_order:
        if srv not in covered:
            return srv
    return preferred_order[0]


def _pick_needed_afternoon_service(grid_day: dict, preferred_order=("TU", "CE")) -> str:
    covered = set()
    for sh in grid_day.values():
        if sh["period"] == "T":
            covered.add(sh["service"])
        elif sh["period"] == "12h":
            parts = sh["service"].split("+")
            covered.add(parts[1] if len(parts) > 1 else parts[0])
    for srv in preferred_order:
        if srv not in covered:
            return srv
    return preferred_order[0]


def ensure_month_schedule(year: int, month: int):
    """
    Verifica si el mes (year, month) tiene turnos cargados en la base de datos.
    Si no los tiene, ejecuta el motor exacto de turnos-ipress.streamlit.app
    (auto_assign_missing_shifts_day10) para generar el rol mensual oficial.
    """
    with get_connection() as conn:
        cnt = conn.execute(
            "SELECT COUNT(*) as c FROM shifts WHERE year = ? AND month = ?", (year, month)
        ).fetchone()["c"]
        if cnt > 0:
            return

    month_info = get_month_info(year, month)
    quota = month_info["monthly_quota"]
    num_days = month_info["num_days"]
    days_meta = {d["day"]: d for d in month_info["days"]}

    with get_connection() as conn:
        doctors = conn.execute(
            """
            SELECT id, name, condition, cmp
            FROM doctors
            WHERE is_active = 1 AND role = 'doctor'
            ORDER BY condition DESC, name ASC
            """
        ).fetchall()

        doc_dict = {d["id"]: dict(d) for d in doctors}
        solano_id = next((d["id"] for d in doctors if "Solano" in d["name"]), None)
        caldas_id = next((d["id"] for d in doctors if "Caldas" in d["name"]), None)
        laureano_id = next((d["id"] for d in doctors if "Laureano" in d["name"]), None)
        sanchez_id = next((d["id"] for d in doctors if "Sanchez" in d["name"]), None)
        chunga_id = next((d["id"] for d in doctors if "Chunga" in d["name"]), None)

        grid = {d: {} for d in range(1, num_days + 1)}
        doc_shifts_count = {doc_id: 0 for doc_id in doc_dict}
        doc_12h_count = {doc_id: 0 for doc_id in doc_dict}

        non_holiday_working_days = [
            d["day"] for d in month_info["days"] if not d["is_sunday"] and not d["is_holiday"]
        ]
        holiday_working_days = [
            d["day"] for d in month_info["days"] if not d["is_sunday"] and d["is_holiday"]
        ]
        ordered_working_days = non_holiday_working_days + holiday_working_days

        def assign_12h(did: int, day_num: int, srv: str):
            grid[day_num][did] = {"period": "12h", "service": srv, "is_proposal": 0}
            doc_shifts_count[did] += 2
            doc_12h_count[did] += 1
            if day_num + 1 <= num_days and not days_meta[day_num + 1]["is_sunday"]:
                if did not in grid[day_num + 1]:
                    grid[day_num + 1][did] = {"period": "DES", "service": "DES", "is_proposal": 0}

        def can_assign_12h(did: int, day_num: int) -> bool:
            if doc_shifts_count[did] + 2 > quota:
                return False
            if did in grid[day_num]:
                return False
            if day_num > 1 and grid[day_num - 1].get(did, {}).get("period") == "12h":
                return False
            if day_num + 1 <= num_days and not days_meta[day_num + 1]["is_sunday"]:
                if did in grid[day_num + 1]:
                    return False
            return True

        def can_assign_6h(did: int, day_num: int) -> bool:
            if doc_shifts_count[did] + 1 > quota:
                return False
            if did in grid[day_num]:
                return False
            if day_num > 1 and grid[day_num - 1].get(did, {}).get("period") == "12h":
                return False
            return True

        # 1. Dr. Solano Ricaldi
        if solano_id:
            wednesdays = [d["day"] for d in month_info["days"] if d["weekday_index"] == 2]
            for w in wednesdays:
                if doc_12h_count[solano_id] < 5 and can_assign_12h(solano_id, w):
                    assign_12h(solano_id, w, "TS+TU")

            candidate_12h_days = [
                d["day"]
                for d in month_info["days"]
                if d["weekday_index"] in (4, 0, 1, 5) and not d["is_holiday"]
            ] + [
                d["day"]
                for d in month_info["days"]
                if d["weekday_index"] in (4, 0, 1, 5) and d["is_holiday"]
            ]
            for cd in candidate_12h_days:
                if doc_12h_count[solano_id] >= 5:
                    break
                if can_assign_12h(solano_id, cd):
                    m_srv = _pick_needed_morning_service(grid[cd], ("TS", "TU", "CE"))
                    t_srv = _pick_needed_afternoon_service(grid[cd], ("TU", "CE"))
                    assign_12h(solano_id, cd, f"{m_srv}+{t_srv}")

        # 2. Dra. Sanchez Quintanilla
        if sanchez_id:
            mwf_days = [d for d in ordered_working_days if days_meta[d]["weekday_index"] in (0, 2, 4)]
            for d in mwf_days:
                if can_assign_12h(sanchez_id, d):
                    assign_12h(sanchez_id, d, "TU+CE")
            for d in mwf_days + ordered_working_days:
                if doc_shifts_count[sanchez_id] >= quota:
                    break
                if can_assign_6h(sanchez_id, d):
                    srv = _pick_needed_morning_service(grid[d], ("TU", "CE", "TS"))
                    grid[d][sanchez_id] = {"period": "M", "service": srv, "is_proposal": 0}
                    doc_shifts_count[sanchez_id] += 1

        # 3. Dra. Chunga Torres
        if chunga_id:
            mwf_days = [d for d in ordered_working_days if days_meta[d]["weekday_index"] in (0, 2, 4)]
            for d in mwf_days:
                if can_assign_12h(chunga_id, d):
                    srv_12 = "CE+CE" if days_meta[d]["weekday_index"] == 2 else "CE+TU"
                    assign_12h(chunga_id, d, srv_12)
            for d in mwf_days + ordered_working_days:
                if doc_shifts_count[chunga_id] >= quota:
                    break
                if can_assign_6h(chunga_id, d):
                    srv = _pick_needed_morning_service(grid[d], ("CE", "TU", "TS"))
                    grid[d][chunga_id] = {"period": "M", "service": srv, "is_proposal": 0}
                    doc_shifts_count[chunga_id] += 1

        # 4. Dra. Caldas Pacsi
        if caldas_id:
            saturdays = [d for d in ordered_working_days if days_meta[d]["weekday_index"] == 5]
            for sat in saturdays:
                if can_assign_12h(caldas_id, sat):
                    assign_12h(caldas_id, sat, "TS+TU")

            caldas_morning_days = [
                d for d in ordered_working_days if days_meta[d]["weekday_index"] in (0, 2, 3, 4)
            ]
            for d in caldas_morning_days + ordered_working_days:
                if doc_shifts_count[caldas_id] >= quota:
                    break
                if can_assign_6h(caldas_id, d):
                    srv = _pick_needed_morning_service(grid[d], ("TS", "TU", "CE"))
                    grid[d][caldas_id] = {"period": "M", "service": srv, "is_proposal": 0}
                    doc_shifts_count[caldas_id] += 1

        # 5. Dra. Laureano Varillas
        if laureano_id:
            for d in ordered_working_days:
                if doc_shifts_count[laureano_id] >= quota:
                    break
                if not can_assign_6h(laureano_id, d):
                    continue
                w_idx = days_meta[d]["weekday_index"]
                if w_idx == 3:
                    srv = _pick_needed_afternoon_service(grid[d], ("CE", "TU"))
                    grid[d][laureano_id] = {"period": "T", "service": srv, "is_proposal": 0}
                    doc_shifts_count[laureano_id] += 1
                elif w_idx in (0, 1, 2, 4, 5):
                    srv = _pick_needed_morning_service(grid[d], ("CE", "TU", "TS"))
                    grid[d][laureano_id] = {"period": "M", "service": srv, "is_proposal": 0}
                    doc_shifts_count[laureano_id] += 1

        # 6. Completar turnos de 6h de Dr. Solano
        if solano_id:
            solano_fill_days = sorted(
                ordered_working_days,
                key=lambda d: (days_meta[d]["is_holiday"], len(grid[d]), d),
            )
            for d in solano_fill_days:
                if doc_shifts_count[solano_id] >= quota:
                    break
                if not can_assign_6h(solano_id, d):
                    continue
                t_covered = {
                    (sh["service"].split("+")[1] if "+" in sh["service"] else sh["service"])
                    for sh in grid[d].values()
                    if sh["period"] in ("T", "12h")
                }
                m_covered = {
                    (sh["service"].split("+")[0] if "+" in sh["service"] else sh["service"])
                    for sh in grid[d].values()
                    if sh["period"] in ("M", "12h")
                }
                if len(m_covered) < 3 and len(m_covered) <= len(t_covered):
                    srv = _pick_needed_morning_service(grid[d], ("TU", "CE", "TS"))
                    grid[d][solano_id] = {"period": "M", "service": srv, "is_proposal": 0}
                else:
                    srv = _pick_needed_afternoon_service(grid[d], ("TU", "CE"))
                    grid[d][solano_id] = {"period": "T", "service": srv, "is_proposal": 0}
                doc_shifts_count[solano_id] += 1

        for d in range(1, num_days + 1):
            for did, sh in grid[d].items():
                conn.execute(
                    """
                    INSERT OR REPLACE INTO shifts
                    (year, month, day, shift_period, service, doctor_id, is_proposal, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    """,
                    (year, month, d, sh["period"], sh["service"], did, sh.get("is_proposal", 0)),
                )
        conn.commit()


def get_day_ce_consultorios(year: int, month: int, day: int) -> dict:
    """
    Determina la programación del Consultorio de Medicina (CE) en Turno Mañana ('M')
    y Turno Tarde ('T') para una fecha específica, junto con sus 16 cupos cada uno.
    """
    ensure_month_schedule(year, month)
    date_iso = f"{year:04d}-{month:02d}-{day:02d}"
    month_info = get_month_info(year, month)
    day_meta = month_info["days"][day - 1]

    with get_connection() as conn:
        # Obtener todos los turnos del día
        shift_rows = conn.execute(
            """
            SELECT s.shift_period, s.service, s.doctor_id, d.name, d.cmp, d.condition
            FROM shifts s
            JOIN doctors d ON s.doctor_id = d.id
            WHERE s.year = ? AND s.month = ? AND s.day = ?
            ORDER BY CASE WHEN s.shift_period IN ('M', 'T') THEN 0 ELSE 1 END, d.name ASC
            """,
            (year, month, day),
        ).fetchall()

        # Verificar si existen overrides administrativos para CE Mañana o Tarde
        overrides = conn.execute(
            """
            SELECT o.shift_period, o.doctor_id, o.is_enabled, o.notes, d.name, d.cmp, d.condition
            FROM clinic_schedule_overrides o
            LEFT JOIN doctors d ON o.doctor_id = d.id
            WHERE o.year = ? AND o.month = ? AND o.day = ?
            """,
            (year, month, day),
        ).fetchall()
        override_map = {r["shift_period"]: dict(r) for r in overrides}

        # Clasificar médicos programados en CE en Mañana y en Tarde
        morning_ce_docs = []
        afternoon_ce_docs = []

        for r in shift_rows:
            sp = r["shift_period"]
            srv = r["service"]
            doc_obj = {
                "doctor_id": r["doctor_id"],
                "doctor_name": r["name"],
                "doctor_cmp": r["cmp"],
                "condition": r["condition"],
                "raw_shift": sp,
                "raw_service": srv,
            }
            if sp == "M" and srv == "CE":
                morning_ce_docs.append(doc_obj)
            elif sp == "T" and srv == "CE":
                afternoon_ce_docs.append(doc_obj)
            elif sp == "12h":
                parts = srv.split("+") if "+" in srv else [srv, srv]
                m_srv = parts[0]
                t_srv = parts[1] if len(parts) > 1 else parts[0]
                if m_srv == "CE":
                    morning_ce_docs.append(doc_obj)
                if t_srv == "CE":
                    afternoon_ce_docs.append(doc_obj)

        # Obtener citas activas (reservada o atendida o no_asistio) en esa fecha
        app_rows = conn.execute(
            """
            SELECT *
            FROM appointments
            WHERE appointment_date = ? AND status != 'cancelada'
            ORDER BY slot_number ASC
            """,
            (date_iso,),
        ).fetchall()

        apps_by_shift_slot = {"M": {}, "T": {}}
        for a in app_rows:
            sp = a["shift_period"]
            if sp in apps_by_shift_slot:
                apps_by_shift_slot[sp][a["slot_number"]] = dict(a)

        def build_shift_info(period_code: str, ce_docs: list[dict]) -> dict:
            ov = override_map.get(period_code)
            slots_def = get_slots_definition(period_code)
            shift_label = (
                "Consultorio de Medicina - Turno Mañana (08:00 a 12:00)"
                if period_code == "M"
                else "Consultorio de Medicina - Turno Tarde (14:00 a 18:00)"
            )
            hours_range = "08:00 - 12:00" if period_code == "M" else "14:00 - 18:00"

            # Determinar si está habilitado y quién es el médico titular
            if ov is not None:
                is_active = bool(ov["is_enabled"]) and (ov["doctor_id"] is not None)
                titular_doc = (
                    {
                        "doctor_id": ov["doctor_id"],
                        "doctor_name": ov["name"],
                        "doctor_cmp": ov["cmp"],
                        "condition": ov["condition"],
                        "raw_shift": period_code,
                        "raw_service": "CE (Asignación Admisión)",
                    }
                    if is_active
                    else None
                )
            else:
                if day_meta["is_sunday"] or len(ce_docs) == 0:
                    is_active = False
                    titular_doc = None
                else:
                    is_active = True
                    titular_doc = ce_docs[0]

            slots_detail = []
            occupied_count = 0
            for sdef in slots_def:
                snum = sdef["slot_number"]
                booked = apps_by_shift_slot[period_code].get(snum)
                if booked is not None:
                    occupied_count += 1
                slots_detail.append(
                    {
                        "slot_number": snum,
                        "start": sdef["start"],
                        "end": sdef["end"],
                        "label": sdef["label"],
                        "is_available": is_active and (booked is None),
                        "appointment": booked,
                    }
                )

            available_count = (MAX_SLOTS_PER_SHIFT - occupied_count) if is_active else 0

            return {
                "shift_period": period_code,
                "shift_name": "Mañana" if period_code == "M" else "Tarde",
                "consultorio_title": shift_label,
                "hours_range": hours_range,
                "is_enabled": is_active,
                "titular_doctor": titular_doc,
                "all_ce_doctors": ce_docs,
                "max_slots": MAX_SLOTS_PER_SHIFT,
                "occupied_slots": occupied_count,
                "available_slots": max(0, available_count),
                "slots": slots_detail,
                "override_notes": ov.get("notes") if ov else None,
            }

        morning_info = build_shift_info("M", morning_ce_docs)
        afternoon_info = build_shift_info("T", afternoon_ce_docs)

        return {
            "year": year,
            "month": month,
            "day": day,
            "date_iso": date_iso,
            "day_meta": day_meta,
            "morning": morning_info,
            "afternoon": afternoon_info,
            "total_available_day": morning_info["available_slots"] + afternoon_info["available_slots"],
            "total_occupied_day": morning_info["occupied_slots"] + afternoon_info["occupied_slots"],
        }


def get_month_ce_overview(year: int, month: int) -> list[dict]:
    """
    Retorna el resumen de todos los días del mes indicando para cada día:
    - Médico programado en Consultorio de Medicina Mañana y cupos disponibles / 16
    - Médico programado en Consultorio de Medicina Tarde y cupos disponibles / 16
    """
    ensure_month_schedule(year, month)
    month_info = get_month_info(year, month)
    overview = []
    for d in range(1, month_info["num_days"] + 1):
        overview.append(get_day_ce_consultorios(year, month, d))
    return overview


def _normalize_name_tokens(raw_name: str) -> list[str]:
    """Normaliza un nombre eliminando tildes y signos de puntuación para comparar coincidencia."""
    if not raw_name:
        return []
    nfkd = unicodedata.normalize("NFKD", raw_name.upper())
    ascii_str = "".join(c for c in nfkd if not unicodedata.combining(c))
    cleaned = "".join(c if c.isalnum() or c.isspace() else " " for c in ascii_str)
    stopwords = {"DE", "DEL", "LA", "LAS", "LOS", "Y"}
    return [tok for tok in cleaned.split() if len(tok) > 1 and tok not in stopwords]


def _names_match(name1: str, name2: str) -> bool:
    """
    Determina si dos cadenas de Apellidos y Nombres corresponden a la misma persona
    (independientemente del orden 'APELLIDOS, NOMBRES' o 'NOMBRES APELLIDOS' y tildes).
    """
    t1 = set(_normalize_name_tokens(name1))
    t2 = set(_normalize_name_tokens(name2))
    if not t1 or not t2:
        return False
    if t1 == t2:
        return True
    overlap = len(t1.intersection(t2))
    min_len = min(len(t1), len(t2))
    if min_len >= 2 and overlap >= min_len:
        return True
    if len(t1) >= 3 and len(t2) >= 3 and overlap >= 3:
        return True
    return False


def validate_booking_data(
    year: int,
    month: int,
    day: int,
    shift_period: str,
    slot_number: int,
    patient_doc_type: str,
    patient_dni: str,
    patient_names: str,
    patient_age: int | None,
    patient_phone: str,
    insurance_type: str,
    reason: str,
    check_past_date: bool = True,
) -> tuple[bool, str, dict | None]:
    """
    Valida estrictamente todos los datos del paciente y la disponibilidad del cupo ANTES
    de confirmar o grabar la cita:
    - Prohibido generar citas en fechas pasadas (anteriores a la fecha actual).
    - Todos los campos solicitados al paciente deben estar llenados por completo.
    a) Documento de identidad: DNI o Carné de Extranjería (CE), SOLO NÚMEROS.
       - Si es DNI: siempre exactamente 8 dígitos numéricos (el primero puede ser 0).
       - Si es Carné de Extranjería (CE): solo números (entre 8 y 12 dígitos).
    b) Número de celular: obligatorio de exactamente 9 dígitos numéricos.
    e) Solamente una (01) cita a la semana por documento de identidad y verificación de
       coincidencia en el nombre; de existir, indica acudir presencialmente al Servicio de Admisión.
    """
    # 0. Validar que la fecha seleccionada NO sea una fecha pasada
    target_dt = date(year, month, day)
    today_pe = get_peru_today()
    if check_past_date and target_dt < today_pe:
        return (
            False,
            f"⛔ FECHA PASADA NO PERMITIDA: La fecha escogida ({day:02d}/{month:02d}/{year}) ya pasó (fecha actual: {today_pe.strftime('%d/%m/%Y')}). "
            f"No se permite generar citas en fechas pasadas. Por favor selecciona el día de hoy o una fecha futura en el Paso 1.",
            None,
        )

    doc_type_raw = (patient_doc_type or "").strip()
    dni_clean = (patient_dni or "").strip()
    names_clean = (patient_names or "").strip().upper()
    phone_clean = (patient_phone or "").strip()
    insurance_clean = (insurance_type or "").strip()
    reason_clean = (reason or "").strip()

    # Verificar que TODOS los campos solicitados al paciente hayan sido llenados por completo
    missing_fields = []
    if not doc_type_raw or doc_type_raw.startswith("--"):
        missing_fields.append("Tipo de Documento (DNI / CE)")
    if not dni_clean:
        missing_fields.append("N° de Documento de Identidad")
    if patient_age is None or str(patient_age).strip() == "":
        missing_fields.append("Edad del Paciente")
    if not names_clean:
        missing_fields.append("Apellidos y Nombres Completos")
    if not insurance_clean or insurance_clean.startswith("--"):
        missing_fields.append("Tipo de Seguro / Financiamiento")
    if not phone_clean:
        missing_fields.append("Número de Celular (9 dígitos)")
    if not reason_clean:
        missing_fields.append("Motivo de Consulta Médica")

    if missing_fields:
        return (
            False,
            "⚠️ CAMPOS INCOMPLETOS: No se puede generar la cita porque todos los campos solicitados al paciente deben ser llenados por completo. "
            f"Falta completar: {', '.join(missing_fields)}.",
            None,
        )

    doc_type_clean = "CE" if ("CE" in doc_type_raw.upper() or "EXTRANJER" in doc_type_raw.upper()) else "DNI"

    # Validar edad numérica válida
    try:
        age_int = int(patient_age)
        if age_int < 0 or age_int > 120:
            return False, "La Edad del paciente debe estar entre 0 y 120 años.", None
    except (ValueError, TypeError):
        return False, "Por favor ingresa una Edad válida en años.", None

    # Validación a) Documento de identidad: solo números
    if not dni_clean.isdigit():
        return (
            False,
            f"El número de {doc_type_clean} debe contener ÚNICAMENTE números (sin letras, espacios ni guiones).",
            None,
        )

    if doc_type_clean == "DNI":
        if len(dni_clean) != 8:
            return (
                False,
                f"El DNI debe tener siempre exactamente 8 dígitos numéricos (ingresaste {len(dni_clean)} dígitos). Recuerda que el primer dígito puede ser 0.",
                None,
            )
    else:
        if len(dni_clean) < 8 or len(dni_clean) > 12:
            return (
                False,
                f"El Carné de Extranjería (CE) debe contener solo números (entre 8 y 12 dígitos; ingresaste {len(dni_clean)}).",
                None,
            )

    if len(names_clean) < 5 or len(_normalize_name_tokens(names_clean)) < 2:
        return (
            False,
            "Por favor ingresa los Apellidos y Nombres completos del paciente (mínimo un apellido y un nombre).",
            None,
        )

    # Validación b) Número celular: 9 dígitos obligatorios, solo números
    if not phone_clean.isdigit() or len(phone_clean) != 9:
        return (
            False,
            f"El número de celular es obligatorio y debe tener exactamente 9 dígitos numéricos (ingresaste '{phone_clean}').",
            None,
        )

    # Validar motivo de consulta completo
    if len(reason_clean) < 3:
        return (
            False,
            "Por favor describe el Motivo de Consulta Médica completo (mínimo 3 caracteres).",
            None,
        )

    if slot_number < 1 or slot_number > MAX_SLOTS_PER_SHIFT:
        return False, f"El número de cupo debe estar entre 1 y {MAX_SLOTS_PER_SHIFT}.", None

    day_data = get_day_ce_consultorios(year, month, day)
    shift_data = day_data["morning"] if shift_period == "M" else day_data["afternoon"]

    if not shift_data["is_enabled"] or not shift_data["titular_doctor"]:
        return (
            False,
            f"El Consultorio de Medicina en el Turno {shift_data['shift_name']} no tiene programación habilitada en esta fecha.",
            None,
        )

    slot_obj = next((s for s in shift_data["slots"] if s["slot_number"] == slot_number), None)
    if not slot_obj or not slot_obj["is_available"]:
        return (
            False,
            f"El Cupo #{slot_number:02d} ({slot_obj['label'] if slot_obj else ''}) ya se encuentra ocupado. Por favor elige otro horario.",
            None,
        )

    date_iso = day_data["date_iso"]
    doc = shift_data["titular_doctor"]

    # Calcular el rango de la semana calendario (Lunes a Domingo) de la fecha solicitada
    week_start = target_dt - timedelta(days=target_dt.weekday())  # Lunes
    week_end = week_start + timedelta(days=6)                    # Domingo
    week_start_iso = week_start.isoformat()
    week_end_iso = week_end.isoformat()

    with get_connection() as conn:
        # 1. Verificar consistencia histórica de identidad (mismo Documento pero distinto Nombre)
        prev_doc_row = conn.execute(
            """
            SELECT patient_names, patient_doc_type, patient_dni
            FROM appointments
            WHERE patient_dni = ? AND status != 'cancelada'
            ORDER BY id DESC LIMIT 1
            """,
            (dni_clean,),
        ).fetchone()

        if prev_doc_row and not _names_match(prev_doc_row["patient_names"], names_clean):
            return (
                False,
                f"⚠️ DISCREPANCIA DE IDENTIDAD: El documento {doc_type_clean} {dni_clean} ya se encuentra registrado en el sistema "
                f"bajo el nombre '{prev_doc_row['patient_names']}', el cual no coincide con '{names_clean}'. "
                f"Por favor verifique sus datos o solicite su cita de manera presencial a través del Servicio de Admisión en el Centro de Salud Su Santidad Juan Pablo II.",
                None,
            )

        # 2. Verificar restricción de MÁXIMO 1 CITA A LA SEMANA por Documento de Identidad o Coincidencia en el Nombre
        week_rows = conn.execute(
            """
            SELECT ticket_code, appointment_date, shift_period, slot_number, slot_time,
                   patient_doc_type, patient_dni, patient_names
            FROM appointments
            WHERE appointment_date BETWEEN ? AND ?
              AND status IN ('reservada', 'atendida')
            ORDER BY appointment_date ASC
            """,
            (week_start_iso, week_end_iso),
        ).fetchall()

        for w_row in week_rows:
            same_doc = (w_row["patient_dni"] == dni_clean)
            same_name = _names_match(w_row["patient_names"], names_clean)
            if same_doc or same_name:
                turno_txt = "Mañana" if w_row["shift_period"] == "M" else "Tarde"
                doc_lbl = w_row["patient_doc_type"] if "patient_doc_type" in w_row.keys() and w_row["patient_doc_type"] else doc_type_clean
                motivo_coincidencia = (
                    "Documento de Identidad y Nombre"
                    if (same_doc and same_name)
                    else ("Documento de Identidad" if same_doc else "Coincidencia en el Nombre del Paciente")
                )
                return (
                    False,
                    f"⚠️ RESTRICCIÓN DE 1 CITA POR SEMANA ({motivo_coincidencia}): "
                    f"El paciente {w_row['patient_names']} ({doc_lbl}: {w_row['patient_dni']}) ya registra una cita programada en esta semana "
                    f"(del {week_start_iso} al {week_end_iso}) para el día {w_row['appointment_date']} en el Turno {turno_txt} — "
                    f"Cupo #{w_row['slot_number']:02d} ({w_row['slot_time']} hrs), Ticket: {w_row['ticket_code']}. "
                    f"Por normativa del establecimiento solo se puede reservar UNA (01) cita a la semana por internet. "
                    f"De requerir una atención adicional, debe solicitar la cita a través del Servicio de Admisión de manera presencial en el Centro de Salud Su Santidad Juan Pablo II.",
                    None,
                )

    return (
        True,
        "Datos validados correctamente.",
        {
            "year": year,
            "month": month,
            "day": day,
            "date_iso": date_iso,
            "weekday_name": day_data["day_meta"]["weekday_name"],
            "shift_period": shift_period,
            "shift_name": shift_data["shift_name"],
            "consultorio_title": shift_data["consultorio_title"],
            "slot_number": slot_number,
            "slot_time": slot_obj["label"],
            "doctor_id": doc["doctor_id"],
            "doctor_name": doc["doctor_name"],
            "doctor_cmp": doc["doctor_cmp"],
            "patient_doc_type": doc_type_clean,
            "patient_dni": dni_clean,
            "patient_names": names_clean,
            "patient_age": age_int,
            "patient_phone": phone_clean,
            "insurance_type": insurance_clean,
            "reason": reason_clean,
        },
    )


def book_appointment(
    year: int,
    month: int,
    day: int,
    shift_period: str,
    slot_number: int,
    patient_dni: str,
    patient_names: str,
    patient_age: int | None,
    patient_phone: str,
    insurance_type: str,
    reason: str,
    booked_by: str = "Portal Web Paciente",
    patient_doc_type: str = "DNI",
    check_past_date: bool = True,
) -> tuple[bool, str, dict | None]:
    """
    Registra una nueva cita médica en uno de los 16 cupos del Consultorio de Medicina (Mañana o Tarde).
    Valida que la fecha no sea pasada, que todos los campos estén completos, documento (DNI 8 dígitos
    con soporte para 0 inicial o CE numérico), celular de 9 dígitos, disponibilidad, médico programado
    en CE y máximo 1 cita por semana.
    """
    ok_val, msg_val, vdata = validate_booking_data(
        year=year,
        month=month,
        day=day,
        shift_period=shift_period,
        slot_number=slot_number,
        patient_doc_type=patient_doc_type,
        patient_dni=patient_dni,
        patient_names=patient_names,
        patient_age=patient_age,
        patient_phone=patient_phone,
        insurance_type=insurance_type,
        reason=reason,
        check_past_date=check_past_date,
    )
    if not ok_val or not vdata:
        return False, msg_val, None

    date_iso = vdata["date_iso"]

    with get_connection() as conn:
        existing_slot = conn.execute(
            """
            SELECT id FROM appointments
            WHERE appointment_date = ? AND shift_period = ? AND slot_number = ? AND status != 'cancelada'
            """,
            (date_iso, shift_period, slot_number),
        ).fetchone()

        if existing_slot:
            return False, "El cupo seleccionado acaba de ser reservado por otro usuario.", None

        base_code = f"JP2-{year:04d}{month:02d}{day:02d}-{shift_period}{slot_number:02d}"
        ticket_code = base_code
        suffix = 1
        while conn.execute("SELECT 1 FROM appointments WHERE ticket_code = ?", (ticket_code,)).fetchone():
            suffix += 1
            ticket_code = f"{base_code}-{suffix}"

        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO appointments (
                ticket_code, year, month, day, appointment_date,
                shift_period, slot_number, slot_time,
                doctor_id, doctor_name, doctor_cmp,
                patient_doc_type, patient_dni, patient_names, patient_age, patient_phone,
                insurance_type, reason, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'reservada')
            """,
            (
                ticket_code,
                year,
                month,
                day,
                date_iso,
                shift_period,
                slot_number,
                vdata["slot_time"],
                vdata["doctor_id"],
                vdata["doctor_name"],
                vdata["doctor_cmp"],
                vdata["patient_doc_type"],
                vdata["patient_dni"],
                vdata["patient_names"],
                vdata["patient_age"],
                vdata["patient_phone"],
                vdata["insurance_type"],
                vdata["reason"],
            ),
        )
        new_id = cursor.lastrowid
        log_audit(
            booked_by,
            "Reserva de Cita Médica",
            f"Ticket {ticket_code} | Fecha {date_iso} ({shift_period} Cupo #{slot_number:02d} {vdata['slot_time']}) | "
            f"Paciente: {vdata['patient_names']} ({vdata['patient_doc_type']} {vdata['patient_dni']}) | Médico: {vdata['doctor_name']}",
            conn=conn,
        )
        conn.commit()

        row = conn.execute("SELECT * FROM appointments WHERE id = ?", (new_id,)).fetchone()
        return True, f"¡Cita reservada con éxito! Código de Ticket: {ticket_code}", dict(row)


def cancel_appointment(ticket_code: str, patient_dni: str = None, cancelled_by: str = "Paciente") -> tuple[bool, str]:
    """Cancela una cita reservada liberando el cupo inmediatamente."""
    code_clean = ticket_code.strip().upper()
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM appointments WHERE UPPER(ticket_code) = ?", (code_clean,)
        ).fetchone()
        if not row:
            return False, "No se encontró ninguna cita con el código de ticket indicado."

        if patient_dni is not None and row["patient_dni"].strip() != patient_dni.strip():
            return False, "El DNI ingresado no coincide con el titular del ticket de cita."

        if row["status"] == "cancelada":
            return False, "Esta cita ya se encontraba cancelada previamente."

        if row["status"] == "atendida":
            return False, "No se puede cancelar una cita que ya fue atendida en consultorio."

        conn.execute(
            """
            UPDATE appointments
            SET status = 'cancelada', updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (row["id"],),
        )
        log_audit(
            cancelled_by,
            "Cancelación de Cita",
            f"Ticket {row['ticket_code']} liberado ({row['appointment_date']} Turno {row['shift_period']} Cupo #{row['slot_number']})",
            conn=conn,
        )
        conn.commit()
        return True, f"La cita {row['ticket_code']} ha sido cancelada y el Cupo #{row['slot_number']} quedó libre."


def update_appointment_status(appointment_id: int, new_status: str, updated_by: str) -> tuple[bool, str]:
    """Actualiza el estado de asistencia de una cita ('reservada', 'atendida', 'no_asistio', 'cancelada')."""
    valid_statuses = ("reservada", "atendida", "no_asistio", "cancelada")
    if new_status not in valid_statuses:
        return False, "Estado inválido."

    with get_connection() as conn:
        row = conn.execute("SELECT * FROM appointments WHERE id = ?", (appointment_id,)).fetchone()
        if not row:
            return False, "Cita no encontrada."

        conn.execute(
            """
            UPDATE appointments
            SET status = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (new_status, appointment_id),
        )
        log_audit(
            updated_by,
            "Cambio de Estado de Cita",
            f"Ticket {row['ticket_code']} ({row['patient_names']}) -> {new_status.upper()}",
            conn=conn,
        )
        conn.commit()
        return True, f"Estado de la cita {row['ticket_code']} actualizado a '{new_status.upper()}'."


def search_appointments(query_dni_or_ticket: str, allow_name_like: bool = True) -> list[dict]:
    """Busca citas por DNI/CE exacto o Código de Ticket exacto (y por nombre parcial si es personal autorizado)."""
    q = query_dni_or_ticket.strip().upper()
    if not q:
        return []
    with get_connection() as conn:
        if allow_name_like:
            rows = conn.execute(
                """
                SELECT *
                FROM appointments
                WHERE UPPER(patient_dni) = ? OR UPPER(ticket_code) = ? OR UPPER(patient_names) LIKE ?
                ORDER BY appointment_date DESC, shift_period ASC, slot_number ASC
                """,
                (q, q, f"%{q}%"),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT *
                FROM appointments
                WHERE UPPER(patient_dni) = ? OR UPPER(ticket_code) = ?
                ORDER BY appointment_date DESC, shift_period ASC, slot_number ASC
                """,
                (q, q),
            ).fetchall()
        return [dict(r) for r in rows]


def set_consultorio_override(
    year: int,
    month: int,
    day: int,
    shift_period: str,
    doctor_id: int | None,
    is_enabled: bool,
    notes: str,
    updated_by: str,
) -> tuple[bool, str]:
    """Permite a Jefatura/Admisión definir o cambiar el médico titular del Consultorio CE Mañana o Tarde."""
    with get_connection() as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO clinic_schedule_overrides
            (year, month, day, shift_period, doctor_id, is_enabled, notes, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """,
            (year, month, day, shift_period, doctor_id, 1 if is_enabled else 0, notes, updated_by),
        )
        # Si ya había citas reservadas en ese día y turno y cambió el médico titular, actualizar doctor_id/name
        if is_enabled and doctor_id is not None:
            doc = conn.execute("SELECT id, name, cmp FROM doctors WHERE id = ?", (doctor_id,)).fetchone()
            if doc:
                date_iso = f"{year:04d}-{month:02d}-{day:02d}"
                conn.execute(
                    """
                    UPDATE appointments
                    SET doctor_id = ?, doctor_name = ?, doctor_cmp = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE appointment_date = ? AND shift_period = ? AND status != 'cancelada'
                    """,
                    (doc["id"], doc["name"], doc["cmp"], date_iso, shift_period),
                )
        log_audit(
            updated_by,
            "Ajuste de Consultorio CE",
            f"Fecha {year:04d}-{month:02d}-{day:02d} Turno {shift_period}: habilitado={is_enabled}, doctor_id={doctor_id} ({notes})",
            conn=conn,
        )
        conn.commit()
    return True, "Configuración del consultorio actualizada correctamente."


def clear_consultorio_override(year: int, month: int, day: int, shift_period: str, updated_by: str):
    """Restaura la asignación automática original proveniente de turnos-ipress."""
    with get_connection() as conn:
        conn.execute(
            """
            DELETE FROM clinic_schedule_overrides
            WHERE year = ? AND month = ? AND day = ? AND shift_period = ?
            """,
            (year, month, day, shift_period),
        )
        log_audit(
            updated_by,
            "Restaurar Rol Original Turnos-IPRESS",
            f"Fecha {year:04d}-{month:02d}-{day:02d} Turno {shift_period} restaurado al rol de turnos-ipress",
            conn=conn,
        )
        conn.commit()
