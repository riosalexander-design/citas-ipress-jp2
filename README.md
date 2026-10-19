# 🏥 Sistema de Citas por Internet — Servicio de Medicina
## Centro de Salud Su Santidad Juan Pablo II

Aplicativo web en **Streamlit** para la reserva, consulta y gestión de citas médicas por internet del **Servicio de Medicina General (Consultorio Externo - `CE`)** del **C.S. Su Santidad Juan Pablo II**, sincronizado con la programación asistencial publicada en **[turnos-ipress.streamlit.app](https://turnos-ipress.streamlit.app/)**.

---

## 🌟 Características Principales

### 1. Consultorios y Horarios Escalonados (16 Cupos por Turno)
- **🌅 Consultorio de Medicina — Turno Mañana (`08:00 a 12:00 hrs`):**
  - **16 cupos diarios** escalonados cada **15 minutos** (`08:00 - 08:15` hasta `11:45 - 12:00`).
- **🌇 Consultorio de Medicina — Turno Tarde (`14:00 a 18:00 hrs`):**
  - **16 cupos diarios** escalonados cada **15 minutos** (`14:00 - 14:15` hasta `17:45 - 18:00`).
- **Capacidad Total:** Hasta **32 cupos diarios** cuando ambos turnos cuentan con médico programado en `CE`.

### 2. Sincronización Automática con Turnos-IPRESS
- Sincroniza automáticamente la nómina médica y el rol mensual del servicio **`CE` (Consultorio de Medicina)** tanto en turnos de 6 horas (`M` / `T`) como en guardias de 12 horas (`12h`: `CE+TU`, `TU+CE`, `CE+CE`).
- **Médicos Cirujanos del Establecimiento:**
  1. **Caldas Pacsi, Lily Katherine** *(CMP: 71201 - CAS)*
  2. **Chunga Torres, Tuly Magnolia** *(CMP: 68432 - CAS)*
  3. **Laureano Varillas, Elizabeth Miriam** *(CMP: 75319 - CAS)*
  4. **Sanchez Quintanilla, Cinthya** *(CMP: 80145 - CAS)*
  5. **Solano Ricaldi, Edward Jhonny** *(CMP: 54982 - Nombrado)*

### 3. Módulos del Sistema
1. **📅 Sacar Cita por Internet (Portal del Paciente):**
   - Selección de fecha y visualización en tiempo real de los médicos programados en Mañana y Tarde.
   - Tablero visual de los **16 horarios de 15 minutos** (`LIBRE` / `OCUPADO`).
   - Validación anti-duplicidad (máximo 1 cita activa por DNI en el mismo día).
   - Emisión inmediata de **Ticket Oficial de Cita (`JP2-YYYYMMDD-M01`)** descargable e imprimible en HTML.
2. **🔍 Consultar o Cancelar mi Cita:**
   - Búsqueda por **DNI**, **Código de Ticket** o **Apellidos**, reimpresión de comprobante y cancelación en línea con liberación inmediata del cupo.
3. **🗓️ Rol Mensual de Medicina (Sincronizado con Turnos-IPRESS):**
   - Calendario mensual completo con los médicos asignados a `CE Mañana` y `CE Tarde` y disponibilidad por día.
4. **🏥 Panel de Admisión y Jefatura Médica (`ADMIN` / PIN `admin2026`):**
   - Padrón diario de los 16 cupos de Mañana y 16 cupos de Tarde.
   - Exportación institucional a **Excel (`.xlsx`)** lista para impresión y firma médica.
   - Control de asistencia (`Reservada`, `Atendida`, `No Asistió`, `Cancelada`) y gestión de contingencias.

---

## 🚀 Cómo Ejecutar Localmente

### Opción 1: Doble clic en Windows
Haz doble clic sobre `run_streamlit.bat`.

### Opción 2: Terminal PowerShell
```powershell
cd "c:\Users\Alexander\OneDrive\Apps\app_citas_jp2"
streamlit run streamlit_app.py
```
