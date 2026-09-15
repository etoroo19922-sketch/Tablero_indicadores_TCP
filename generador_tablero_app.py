"""
GENERADOR DE TABLERO DE INDICADORES - TRIPULACION POR BASE  (v2)
====================================================================
App web (Streamlit) que guía al usuario paso a paso:
  1) Elegir la base a procesar
  2) Cargar cada archivo requerido, EN ORDEN, con validación inmediata
  3) Generar y descargar el tablero de indicadores en Excel

NOVEDADES v2:
  - Excluye los DH (registros PFTR/RFTR en el tx time) de todos los cálculos
  - Clasificación por continente/región en vez de solo Europa/Internacional/Nacional:
    Norteamérica (USA+Canadá), Centroamérica (incl. México), Caribe, Suramérica,
    Europa, Otro Continente (preparado para AUH - Abu Dhabi - desde octubre), Nacional
  - Resumen operativo: total de legs de la base y total de Block Hours de la base
  - Supervisores: distribución de legs con 1/2/3+ supervisores, y % de cobertura
    de SUPINT y de SUPNAL por separado, todo por continente
  - Especialistas: cobertura cruzada por continente Y por flota (787/330/320)
  - Viáticos y Block Time: incluye el tripulante con el valor más bajo y más alto
    de cada categoría (solo ESTADO = LINEA)
  - Cada sección trae una explicación breve de cómo se calcula y de dónde sale
    la información

Autor: generado con Claude a partir del proceso ya validado manualmente.
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict, Counter
from datetime import datetime
import io

st.set_page_config(page_title="Generador de Tablero por Base", layout="centered")

# =========================================================================
# CONSTANTES DE NEGOCIO
# =========================================================================
AEROPUERTOS_COLOMBIA = {
    "ADZ", "AXM", "BAQ", "BGA", "BOG", "CLO", "CTG", "CUC", "EJA", "IBE", "IPI",
    "LET", "MDE", "MTR", "NVA", "PEI", "PPN", "PSO", "RCH", "SMR", "VUP",
}
AEROPUERTOS_EUROPA = {"BCN", "CDG", "LHR", "MAD"}
AEROPUERTOS_NORTEAMERICA = {"BOS", "DFW", "FLL", "IAD", "JFK", "MCO", "MIA", "ORD", "TPA", "YUL", "YYZ"}
AEROPUERTOS_CENTROAMERICA = {"GUA", "SAL", "PTY", "MEX", "MTY", "CUN"}  # incluye Mexico, por decision del usuario
AEROPUERTOS_CARIBE = {"PUJ", "SDQ", "SJU", "AUA", "CUR", "HAV"}
AEROPUERTOS_SURAMERICA = {
    "ASU", "EZE", "AEP", "GIG", "GRU", "GYE", "UIO", "LIM", "LPB", "MVD",
    "SCL", "VVI", "GEO", "COR", "BEL", "MAO", "CUZ", "VLN", "MAR",
}
AEROPUERTOS_OTRO_CONTINENTE = {"AUH"}  # Abu Dhabi (Emiratos Arabes Unidos) - entra en octubre 2026

REGIONES_ORDEN = ["NORTEAMERICA", "CENTROAMERICA", "CARIBE", "SURAMERICA", "EUROPA", "OTRO_CONTINENTE", "NACIONAL"]
REGIONES_LABEL = {
    "NORTEAMERICA": "Norteamérica (USA y Canadá)",
    "CENTROAMERICA": "Centroamérica (incl. México)",
    "CARIBE": "Caribe",
    "SURAMERICA": "Suramérica",
    "EUROPA": "Europa",
    "OTRO_CONTINENTE": "Otro Continente (ej. AUH)",
    "NACIONAL": "Nacional",
}

CODIGOS_ESPECIALES = {"ILLP", "VAC", "LUS", "OFI", "PSI", "XMTR", "MAT", "XMAT", "OFFICE_M"}
CATEGORIAS_ORDEN = ["SUPINT", "SUPINTN", "INTNAL", "SUPNAL", "AUXNAL"]  # de mas antiguo a mas reciente
BASES_VALIDAS = ["BOG", "MDE", "CLO", "BAQ", "CTG", "PEI", "BGA"]


def region_aeropuerto(aeropuerto):
    """Clasifica un aeropuerto en su region/continente. Devuelve 'SIN_CLASIFICAR'
    si no esta en ninguna lista conocida, para poder detectarlo y avisar."""
    a = aeropuerto.strip().upper()
    if a in AEROPUERTOS_COLOMBIA:
        return "NACIONAL"
    if a in AEROPUERTOS_EUROPA:
        return "EUROPA"
    if a in AEROPUERTOS_NORTEAMERICA:
        return "NORTEAMERICA"
    if a in AEROPUERTOS_CENTROAMERICA:
        return "CENTROAMERICA"
    if a in AEROPUERTOS_CARIBE:
        return "CARIBE"
    if a in AEROPUERTOS_SURAMERICA:
        return "SURAMERICA"
    if a in AEROPUERTOS_OTRO_CONTINENTE:
        return "OTRO_CONTINENTE"
    return "SIN_CLASIFICAR"


def fleet_group(equip):
    equip = str(equip).strip().upper()
    if equip in ("788", "78N"):
        return "B787"
    if equip == "320":
        return "A320"
    if equip in ("333", "332"):
        return "A330"
    if equip == "AT7":
        return "ATR72"
    if equip == "7M8":
        return "737MAX"
    return "OTRO"


def hhmm_a_horas(txt):
    if pd.isna(txt) or ":" not in str(txt):
        return None
    h, m = str(txt).split(":")
    try:
        return int(h) + int(m) / 60
    except ValueError:
        return None


# =========================================================================
# PASO 0: elegir base
# =========================================================================
st.title("📊 Generador de Tablero de Indicadores — Tripulación por Base")
st.markdown(
    "Esta herramienta procesa el tx time, el reporte KPI, ESTADOS, la planta, "
    "la lista de vuelos/flota y la lista de especialistas, y genera el tablero "
    "de indicadores para la base que elijas. **Los DH (deadhead) se excluyen "
    "automáticamente de todos los cálculos.** Sube los archivos en el orden pedido; "
    "cada uno se valida antes de dejarte avanzar."
)

with st.expander("ℹ️ ¿Cómo se calcula cada indicador? (léelo antes de empezar)"):
    st.markdown("""
**Resumen operativo (Block Hours y legs por flota):** las Block Hours son las horas
de bloque de todos los tripulantes en estado LINEA, tomadas del reporte KPI. Los legs
se cuentan de forma única (fecha + vuelo + aeropuerto) — si un vuelo tiene 7
tripulantes asignados, sigue contando como **1 solo leg**, no como 7. Se desglosan
por flota (avión) cruzando el número de vuelo con la Flight List, y se excluyen
siempre los DH.

**Continentes:** cada vuelo se clasifica usando su **origen y destino reales**, tomados
de la Flight List (columnas Dept Sta / Arvl Sta) — no solo del campo del tx time. Esto
es importante porque el campo del tx time (PLEG) registra el **aeropuerto de origen**
del leg, así que un vuelo que sale de BOG hacia el exterior aparece como "BOG" en el
tx time y se clasificaría mal como Nacional si no se cruzara con la ruta real.

**1. Supervisores por vuelo:** por cada leg único, se cuenta cuántos tripulantes de
categoría SUPINT, SUPINTN o SUPNAL están asignados. Se agrupan los legs según si
tienen 1, 2, 3 o más supervisores asignados (cuenta combinada de las tres categorías).
El % de cobertura de SUPINT es (legs con al menos un SUPINT o SUPINTN) / (total de
legs) de ese continente. El % de SUPNAL **solo cuenta los legs donde NO hay ya un
supervisor internacional** — es decir, mide cuánto está cubriendo el SUPNAL como
respaldo cuando no hay SUPINT/SUPINTN, no cuántos legs tienen un SUPNAL en general.

**2. Vuelos por categoría:** cuenta cuántas asignaciones tripulante-leg hay por cada
categoría y continente — viene directo del tx time, cruzado con la categoría de cada
tripulante (columna CATEGORIA de la planta).

**3. Especialistas:** por cada leg único, se determina su continente (según el
aeropuerto de destino) y su flota (según el número de vuelo, cruzado con la Flight
List). El % de cobertura es (legs con al menos un especialista asignado) / (total de
legs) para cada combinación continente × flota.

**4. Viáticos (Allowances):** promedio, target, mínimo y máximo tomados del reporte
KPI (columnas ALLOWANCES y TARGET ALLOWANCES), filtrando solo tripulantes con
ESTADO = LINEA (con la validación de estado aplicada — ver más abajo). Se identifica
por nombre a quien tiene el valor más bajo y más alto de cada categoría.

**5. Block Time:** mismo tratamiento que Viáticos, pero con las columnas BLOCK HOURS
y TARGET PER DIEM del reporte KPI (convertidas de HH:MM a horas decimales).

**3.1 Análisis Europa (incluye AUH):** para los vuelos a Europa y Abu Dhabi (AUH),
SOLO SUPINT y SUPINTN cuentan como supervisor válido — el SUPNAL en Europa NO cuenta.
Muestra: cuántos legs a Europa/AUH lleva cada categoría, cuántos están cubiertos por
supervisor internacional vs. sin cobertura, y un listado de los vuelos (fecha + número)
que no tienen SUPINT/SUPINTN, con las categorías que sí llevaron.

**Validación de estado:** revisa el tx time en busca de tripulantes marcados
"LINEA" en el archivo de ESTADOS que en realidad tengan días con código ILLP, VAC,
LUS, OFI, PSI, XMTR, MAT, XMAT u OFFICE_M — si los tiene, se usa ese código real en
vez de LINEA, y por lo tanto queda excluido de los cálculos de Viáticos y Block Time.
    """)

base = st.selectbox("1️⃣ Elige la BASE a procesar", BASES_VALIDAS)
st.divider()

# =========================================================================
# PASO 1: tx time (.nlc)
# =========================================================================
st.subheader("2️⃣ TXT (exportación JEP → NLC)")
st.caption("El archivo de texto que se genera al exportar la programación desde JEP a formato NLC. Debe ser el archivo con extensión **.nlc**, sin abrir ni modificar en Excel.")
nlc_file = st.file_uploader("Cargar archivo .nlc", type=["nlc", "txt"], key="nlc")

lineas_nlc = None
if nlc_file is not None:
    try:
        contenido = nlc_file.read().decode("utf-8", errors="replace")
        lineas_nlc = contenido.splitlines()
    except Exception as e:
        st.error(f"❌ No pude leer el archivo como texto. Detalle: {e}")
        st.stop()

    if len(lineas_nlc) == 0:
        st.error("❌ El archivo está vacío. Verifica que hayas subido el archivo correcto.")
        st.stop()
    if not lineas_nlc[0].startswith("FHDR"):
        st.error(
            "❌ La primera línea del archivo no empieza con 'FHDR'. Esto no parece un tx time "
            "(.nlc) válido — revisa que sea el archivo exportado del sistema de programación, "
            "sin abrir/editar en Excel (eso corrompe el formato)."
        )
        st.stop()
    tiene_roster = any(l.startswith(("RABS", "RPRG", "RROL", "RFTR")) for l in lineas_nlc)
    if not tiene_roster:
        st.error(
            "❌ El archivo no trae la sección de asignación de tripulantes (líneas RABS/RPRG/RROL/RFTR). "
            "Puede que el archivo esté incompleto o truncado — vuelve a exportarlo."
        )
        st.stop()

    # ---- Extraer periodo (FHDR) y nombre del archivo, para trazabilidad ----
    fhdr_parts = lineas_nlc[0].split("|")
    fecha_inicio_txt = fhdr_parts[1] if len(fhdr_parts) > 1 else "?"
    fecha_fin_txt = fhdr_parts[2] if len(fhdr_parts) > 2 else "?"

    def formatear_fecha_yyyymmdd(s):
        if len(s) == 8 and s.isdigit():
            return f"{s[6:8]}/{s[4:6]}/{s[0:4]}"
        return s

    periodo_txt = f"{formatear_fecha_yyyymmdd(fecha_inicio_txt)} — {formatear_fecha_yyyymmdd(fecha_fin_txt)}"
    nombre_archivo_nlc = nlc_file.name

    st.success(f"✅ Tx time válido — {len(lineas_nlc):,} líneas leídas.")
    st.info(
        f"📄 **Archivo:** `{nombre_archivo_nlc}`  \n"
        f"📅 **Periodo programado (según encabezado FHDR):** {periodo_txt}  \n\n"
        f"⚠️ Confirma que este es el tx time del **periodo y escenario correcto** antes de continuar — "
        f"si comparas el resultado contra un reporte de asignaciones u otro archivo, ambos deben "
        f"corresponder a la misma fecha de generación."
    )

if lineas_nlc is None:
    st.info("⬆️ Sube el tx time para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 2: KPI
# =========================================================================
st.subheader("3️⃣ Reporte KPI")
st.caption("Se genera junto con el **CTF del mismo escenario** a revisar. Se descarga en Excel y **se procesa sin ninguna modificación** antes de cargarlo aquí. Debe tener la hoja 'JCR KPIs Report'.")
kpi_file = st.file_uploader("Cargar reporte KPI", type=["xls", "xlsx"], key="kpi")

kpi_df = None
if kpi_file is not None:
    try:
        xls = pd.ExcelFile(kpi_file)
    except Exception as e:
        st.error(f"❌ No pude abrir el archivo como Excel. Detalle: {e}")
        st.stop()

    if "JCR KPIs Report" not in xls.sheet_names:
        st.error(
            f"❌ El archivo no tiene una hoja llamada 'JCR KPIs Report' (hojas encontradas: "
            f"{', '.join(xls.sheet_names)}). Verifica que sea el reporte KPI correcto y no lo hayas renombrado."
        )
        st.stop()

    kpi_df = xls.parse("JCR KPIs Report", header=0)
    columnas_requeridas = ["CREW ID", "CREW MEMBER NAME", "HOMEBASE", "CATEGORY",
                            "ALLOWANCES", "TARGET ALLOWANCES", "BLOCK HOURS", "TARGET PER DIEM", "# EUROPEOS"]
    faltantes = [c for c in columnas_requeridas if c not in kpi_df.columns]
    if faltantes:
        st.error(
            f"❌ Al archivo KPI le faltan estas columnas: {', '.join(faltantes)}. "
            "No edites los encabezados del reporte original."
        )
        st.stop()

    kpi_df["CREW ID"] = kpi_df["CREW ID"].astype(str).str.strip()
    n_base = (kpi_df["HOMEBASE"] == base).sum()
    if n_base == 0:
        st.error(
            f"❌ No hay ningún tripulante con HOMEBASE = '{base}' en este archivo. "
            "Verifica que elegiste la base correcta arriba, o que este es el archivo KPI de esta base."
        )
        st.stop()

    st.success(f"✅ KPI válido — {len(kpi_df):,} filas totales, {n_base:,} de la base {base}.")

if kpi_df is None:
    st.info("⬆️ Sube el reporte KPI para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 3: ESTADOS
# =========================================================================
st.subheader("4️⃣ ESTADOS (.xlsx)")
st.caption("Archivo con el **estado de cada tripulante para el mes que se va a analizar**. Debe mantener siempre el mismo formato: columna B = Crew ID, columna G = ESTADO, con encabezados en la fila 1.")
estados_file = st.file_uploader("Cargar archivo de ESTADOS", type=["xlsx"], key="estados")

estado_map = None
if estados_file is not None:
    try:
        wb_e = openpyxl.load_workbook(estados_file, data_only=True)
        ws_e = wb_e[wb_e.sheetnames[0]]
    except Exception as e:
        st.error(f"❌ No pude abrir el archivo. Detalle: {e}")
        st.stop()

    header_row = [c.value for c in ws_e[1]]
    if len(header_row) < 7:
        st.error(
            "❌ El archivo debe tener al menos 7 columnas (A a G), con el Crew ID en la columna B "
            "y el ESTADO en la columna G. Revisa que no se hayan borrado columnas."
        )
        st.stop()

    estado_map = {}
    for row in ws_e.iter_rows(min_row=2, values_only=True):
        if row[1]:
            estado_map[str(row[1]).strip()] = row[6]

    if len(estado_map) == 0:
        st.error("❌ No encontré ningún Crew ID en la columna B. Verifica el archivo.")
        st.stop()

    n_linea = sum(1 for v in estado_map.values() if v == "LINEA")
    st.success(f"✅ ESTADOS válido — {len(estado_map):,} tripulantes, {n_linea:,} en estado LINEA.")

if estado_map is None:
    st.info("⬆️ Sube el archivo de ESTADOS para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 4: Planta / RPT_TRIPULACIÓN
# =========================================================================
st.subheader("5️⃣ Planta activa (Fly Up)")
st.caption("Reporte de planta activa descargado directamente de **Fly Up**, sin ninguna modificación. Columnas: CREW CODE, CATEGORIA, APELLIDO, NOMBRE, HOME BASE.")
planta_file = st.file_uploader("Cargar archivo de planta", type=["xlsx"], key="planta")

crew_info = None
if planta_file is not None:
    try:
        wb_p = openpyxl.load_workbook(planta_file, data_only=True)
        ws_p = wb_p[wb_p.sheetnames[0]]
    except Exception as e:
        st.error(f"❌ No pude abrir el archivo. Detalle: {e}")
        st.stop()

    headers = [c.value for c in ws_p[1]]
    requeridas = ["CREW CODE", "CATEGORIA", "APELLIDO", "NOMBRE", "HOME BASE"]
    faltantes = [c for c in requeridas if c not in headers]
    if faltantes:
        st.error(f"❌ A la planta le faltan estas columnas: {', '.join(faltantes)}.")
        st.stop()

    idx = {h: i for i, h in enumerate(headers)}
    crew_info = {}
    for row in ws_p.iter_rows(min_row=2, values_only=True):
        code = row[idx["CREW CODE"]]
        if code is None:
            continue
        cid = str(code).strip().zfill(8)
        crew_info[cid] = {
            "categoria": row[idx["CATEGORIA"]],
            "apellido": row[idx["APELLIDO"]],
            "nombre": row[idx["NOMBRE"]],
            "home_base": row[idx["HOME BASE"]],
        }

    n_base_planta = sum(1 for v in crew_info.values() if v["home_base"] == base)
    if n_base_planta == 0:
        st.warning(
            f"⚠️ No encontré tripulantes con HOME BASE = '{base}' en la planta. "
            "No es necesariamente un error (se usa el HOMEBASE del KPI como fuente principal), "
            "pero revisa si es el archivo correcto."
        )
    st.success(f"✅ Planta válida — {len(crew_info):,} tripulantes cargados.")

if crew_info is None:
    st.info("⬆️ Sube el archivo de planta para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 5: Lista de vuelos / Flight List
# =========================================================================
st.subheader("6️⃣ Flight List (SSIM — Itinerarios)")
st.caption("El archivo de Flight List que llega por correo del equipo de **Itinerarios** cuando comparten el SSIM. Debe tener la hoja 'LT' con columnas Equip, Flt Num, Dept Sta, Arvl Sta.")
flota_file = st.file_uploader("Cargar Flight List", type=["xlsx"], key="flota")

flt_flota = None
flt_ruta = None
if flota_file is not None:
    try:
        wb_f = openpyxl.load_workbook(flota_file, data_only=True)
    except Exception as e:
        st.error(f"❌ No pude abrir el archivo. Detalle: {e}")
        st.stop()

    if "LT" not in wb_f.sheetnames:
        st.error(
            f"❌ El archivo no tiene una hoja llamada 'LT' (hojas encontradas: {', '.join(wb_f.sheetnames)})."
        )
        st.stop()
    ws_f = wb_f["LT"]
    headers = [c.value for c in ws_f[1]]
    columnas_flt_requeridas = ["Equip", "Flt Num", "Dept Sta", "Arvl Sta"]
    faltantes_flt = [c for c in columnas_flt_requeridas if c not in headers]
    if faltantes_flt:
        st.error(f"❌ A la hoja 'LT' le faltan estas columnas: {', '.join(faltantes_flt)}.")
        st.stop()

    idx_equip = headers.index("Equip")
    idx_flt = headers.index("Flt Num")
    idx_dept = headers.index("Dept Sta")
    idx_arvl = headers.index("Arvl Sta")
    flt_equip_counts = defaultdict(Counter)
    flt_ruta = {}
    for row in ws_f.iter_rows(min_row=2, values_only=True):
        flt = row[idx_flt]
        equip = row[idx_equip]
        if flt is None:
            continue
        flt = int(flt)
        flt_equip_counts[flt][equip] += 1
        if flt not in flt_ruta:
            flt_ruta[flt] = (row[idx_dept], row[idx_arvl])

    flt_flota = {}
    for flt, counter in flt_equip_counts.items():
        equip_top = counter.most_common(1)[0][0]
        flt_flota[flt] = fleet_group(equip_top)

    if len(flt_flota) == 0:
        st.error("❌ No encontré ningún número de vuelo válido en la hoja 'LT'.")
        st.stop()
    st.success(
        f"✅ Flight List válida — {len(flt_flota):,} números de vuelo mapeados a flota y ruta "
        f"(origen/destino real, para clasificar correctamente por continente)."
    )

if flt_flota is None:
    st.info("⬆️ Sube la Flight List para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 6: Especialistas
# =========================================================================
st.subheader("7️⃣ Especialistas (MINT)")
st.caption("Listado descargado de **MINT**, incluyendo únicamente las categorías que aplican como especialistas para el roster. Columna A = Crew ID.")
esp_file = st.file_uploader("Cargar lista de especialistas", type=["xlsx"], key="especialistas")

especialista_set = None
if esp_file is not None:
    try:
        wb_esp = openpyxl.load_workbook(esp_file, data_only=True)
        ws_esp = wb_esp[wb_esp.sheetnames[0]]
    except Exception as e:
        st.error(f"❌ No pude abrir el archivo. Detalle: {e}")
        st.stop()

    especialista_set = set()
    for row in ws_esp.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            especialista_set.add(str(row[0]).strip().zfill(8))

    if len(especialista_set) == 0:
        st.error("❌ No encontré ningún Crew ID en la columna A. Verifica el archivo.")
        st.stop()
    st.success(f"✅ Especialistas válido — {len(especialista_set):,} tripulantes especialistas.")

if especialista_set is None:
    st.info("⬆️ Sube la lista de especialistas para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 7.5: Deteccion y clasificacion manual de aeropuertos desconocidos
# =========================================================================
if "aeropuertos_manuales" not in st.session_state:
    st.session_state["aeropuertos_manuales"] = {}

def region_aeropuerto_con_manual(aeropuerto):
    """Igual que region_aeropuerto, pero primero revisa si el usuario ya
    clasifico manualmente este codigo en la sesion actual."""
    if aeropuerto is None:
        return "SIN_CLASIFICAR"
    a = aeropuerto.strip().upper()
    manual = st.session_state["aeropuertos_manuales"].get(a)
    if manual:
        return manual
    return region_aeropuerto(a)

@st.cache_data(show_spinner=False)
def parsear_tx_time(lineas_nlc_tuple):
    """Parsea el tx time una sola vez (cacheado): separa la seccion de patrones
    de vuelo (PPRG/PLEG) de la seccion de roster (RABS/RPRG/RROL/RFTR), excluyendo
    siempre los DH (PFTR en patrones, RFTR en roster)."""
    lineas_nlc = list(lineas_nlc_tuple)
    boundary = None
    for i, l in enumerate(lineas_nlc):
        if l.startswith(("RABS", "RPRG", "RROL", "RFTR")):
            boundary = i
            break

    nl_flights = defaultdict(list)
    cur_nl = None
    for l in lineas_nlc[:boundary]:
        if not l.strip():
            continue
        p = l.split("|")
        if p[0] == "PPRG" and len(p) > 14:
            cur_nl = p[14]
        elif p[0] == "PLEG" and cur_nl and len(p) > 3:
            nl_flights[cur_nl].append((p[1], p[2].strip(), p[3].strip()))

    crew_legs = defaultdict(list)
    crew_code_days = defaultdict(Counter)
    cur_crew = None
    for l in lineas_nlc[boundary:]:
        if not l.strip():
            continue
        q = l.split("|")
        tipo = q[0]
        if tipo == "RABS":
            cur_crew = q[1]
            if len(q) > 5 and q[5] in CODIGOS_ESPECIALES:
                crew_code_days[cur_crew][q[5]] += 1
        elif tipo == "RPRG":
            cur_crew = q[1]
            if len(q) > 3:
                for ev in nl_flights.get(q[3], []):
                    crew_legs[cur_crew].append(ev)
        elif tipo == "RFTR":
            cur_crew = q[1]
        elif tipo == "RROL":
            cur_crew = q[1]
            if len(q) > 5 and q[2] == "L":
                crew_legs[cur_crew].append((q[3], q[4].strip(), q[5].strip()))

    validacion_estado = {}
    for cid, counter in crew_code_days.items():
        mejor_codigo, dias = counter.most_common(1)[0]
        validacion_estado[cid] = (mejor_codigo, dias)

    return dict(crew_legs), validacion_estado

crew_legs, validacion_estado = parsear_tx_time(tuple(lineas_nlc))
kpi_homebase_map = dict(zip(kpi_df["CREW ID"], kpi_df["HOMEBASE"]))

# Pre-escaneo: solo aeropuertos/vuelos que SI aplican a la base elegida
codigos_en_uso = set()
for crew_id, legs in crew_legs.items():
    info = crew_info.get(crew_id)
    if info is None:
        continue
    hb = kpi_homebase_map.get(crew_id, info["home_base"])
    if hb != base:
        continue
    for fecha, vuelo, aeropuerto in legs:
        codigos_en_uso.add(aeropuerto.strip().upper())
        try:
            num_vuelo = int(vuelo.split()[1])
        except (IndexError, ValueError):
            num_vuelo = None
        ruta = flt_ruta.get(num_vuelo)
        if ruta:
            if ruta[0]:
                codigos_en_uso.add(str(ruta[0]).strip().upper())
            if ruta[1]:
                codigos_en_uso.add(str(ruta[1]).strip().upper())

codigos_desconocidos = sorted(
    c for c in codigos_en_uso
    if region_aeropuerto_con_manual(c) == "SIN_CLASIFICAR"
)

if codigos_desconocidos:
    st.subheader("7.5️⃣ Aeropuertos sin clasificar")
    st.warning(
        f"⚠️ Encontré {len(codigos_desconocidos)} código(s) de aeropuerto que no están en ninguna "
        "lista de continente. Clasifícalos abajo antes de generar el tablero — así no se pierde "
        "ningún leg de la clasificación por continente."
    )
    with st.form("form_aeropuertos_desconocidos"):
        opciones_continente = ["NORTEAMERICA", "CENTROAMERICA", "CARIBE", "SURAMERICA",
                                "EUROPA", "OTRO_CONTINENTE", "NACIONAL"]
        labels_continente = {
            "NORTEAMERICA": "Norteamérica (USA/Canadá)", "CENTROAMERICA": "Centroamérica (incl. México)",
            "CARIBE": "Caribe", "SURAMERICA": "Suramérica", "EUROPA": "Europa",
            "OTRO_CONTINENTE": "Otro Continente", "NACIONAL": "Nacional (Colombia)",
        }
        respuestas = {}
        for codigo in codigos_desconocidos:
            col1, col2 = st.columns([1, 2])
            with col1:
                st.text_input("Código", value=codigo, disabled=True, key=f"cod_{codigo}")
            with col2:
                pais = st.text_input(f"País de {codigo}", key=f"pais_{codigo}")
                continente = st.selectbox(
                    f"Continente/región de {codigo}", opciones_continente,
                    format_func=lambda x: labels_continente[x], key=f"cont_{codigo}"
                )
                respuestas[codigo] = continente
        enviado = st.form_submit_button("✅ Confirmar clasificación")
        if enviado:
            for codigo, continente in respuestas.items():
                st.session_state["aeropuertos_manuales"][codigo] = continente
            st.success("Clasificación guardada. Ya puedes generar el tablero abajo.")
            st.rerun()
    st.stop()

st.divider()

# =========================================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================================
st.subheader("8️⃣ Generar tablero")

if st.button("🚀 Generar tablero de indicadores", type="primary"):
    with st.spinner("Calculando indicadores..."):


        leg_map = {}  # (fecha,vuelo,aeropuerto) -> {region, flota, cat_counts:Counter, esp:bool}
        cat_legs = Counter()  # (categoria, region) -> n asignaciones tripulante-leg
        aeropuertos_sin_clasificar = set()
        vuelos_sin_ruta = set()

        def region_de_vuelo(num_vuelo, aeropuerto_txtime):
            """Determina la region real del vuelo usando el origen/destino verdadero
            de la Flight List (Dept Sta / Arvl Sta), en vez de confiar solo en el
            campo del tx time (que en PLEG resulta ser el AEROPUERTO DE ORIGEN del
            leg, no el destino -- por eso un vuelo que SALE de BOG hacia el exterior
            aparece con "BOG" en el tx time y se clasificaria mal como Nacional si
            se usara ese campo solo)."""
            ruta = flt_ruta.get(num_vuelo) if num_vuelo is not None else None
            if ruta is None:
                vuelos_sin_ruta.add(num_vuelo)
                # sin dato de ruta: usamos el campo del tx time como respaldo
                return region_aeropuerto_con_manual(aeropuerto_txtime)
            dept, arvl = ruta
            r_dept = region_aeropuerto_con_manual(dept) if dept else "SIN_CLASIFICAR"
            r_arvl = region_aeropuerto_con_manual(arvl) if arvl else "SIN_CLASIFICAR"
            if r_dept not in ("NACIONAL", "SIN_CLASIFICAR"):
                return r_dept
            if r_arvl not in ("NACIONAL", "SIN_CLASIFICAR"):
                return r_arvl
            if "SIN_CLASIFICAR" in (r_dept, r_arvl):
                aeropuertos_sin_clasificar.add(dept if r_dept == "SIN_CLASIFICAR" else arvl)
                return "SIN_CLASIFICAR"
            return "NACIONAL"

        for crew_id, legs in crew_legs.items():
            info = crew_info.get(crew_id)
            if info is None:
                continue
            hb = kpi_homebase_map.get(crew_id, info["home_base"])
            if hb != base:
                continue
            categoria = info["categoria"]
            es_esp = crew_id in especialista_set

            for fecha, vuelo, aeropuerto in legs:
                try:
                    num_vuelo = int(vuelo.split()[1])
                except (IndexError, ValueError):
                    num_vuelo = None

                region = region_de_vuelo(num_vuelo, aeropuerto)

                cat_legs[(categoria, region)] += 1

                leg_key = (fecha, vuelo, aeropuerto)
                if leg_key not in leg_map:
                    flota = flt_flota.get(num_vuelo, "OTRO")
                    leg_map[leg_key] = {"region": region, "flota": flota, "cat_counts": Counter(), "esp": False}
                leg_map[leg_key]["cat_counts"][categoria] += 1
                if es_esp:
                    leg_map[leg_key]["esp"] = True

        # ---- Resumen operativo: total legs y block hours de la base ----
        total_legs_base = len(leg_map)
        # Legs unicos por flota (avion) — cada leg (fecha+vuelo+aeropuerto) cuenta 1 sola vez
        # sin importar cuantos tripulantes tenga asignados, y sin contar los DH.
        legs_por_flota = Counter(d["flota"] for d in leg_map.values())

        # ---- Seccion 1: supervisores (distribucion 1/2/3+ y cobertura SUPINT/SUPNAL) por region ----
        sup_dist = defaultdict(Counter)      # region -> {1:n legs con 1 sup, 2:.., '3+':..}
        sup_total_legs = Counter()           # region -> total legs
        sup_con_supint = Counter()           # region -> legs con >=1 SUPINT/SUPINTN
        sup_con_supnal = Counter()           # region -> legs con >=1 SUPNAL
        sup_sin_ninguno = Counter()          # region -> legs sin ningun supervisor

        for d in leg_map.values():
            region = d["region"]
            cats = d["cat_counts"]
            n_sup = cats.get("SUPINT", 0) + cats.get("SUPINTN", 0) + cats.get("SUPNAL", 0)
            sup_total_legs[region] += 1
            if n_sup == 0:
                sup_sin_ninguno[region] += 1
            elif n_sup == 1:
                sup_dist[region][1] += 1
            elif n_sup == 2:
                sup_dist[region][2] += 1
            else:
                sup_dist[region]["3+"] += 1
            if cats.get("SUPINT", 0) + cats.get("SUPINTN", 0) > 0:
                sup_con_supint[region] += 1
            elif cats.get("SUPNAL", 0) > 0:
                # El SUPNAL solo cuenta como cobertura cuando NO hay ya un
                # supervisor internacional (SUPINT/SUPINTN) en ese leg - es un
                # respaldo, no una coincidencia con el supervisor internacional.
                sup_con_supnal[region] += 1

        # ---- Seccion 3: especialistas por region x flota ----
        esp_tot_rf = Counter()   # (region, flota) -> total legs
        esp_cov_rf = Counter()   # (region, flota) -> legs con especialista
        for d in leg_map.values():
            key = (d["region"], d["flota"])
            esp_tot_rf[key] += 1
            if d["esp"]:
                esp_cov_rf[key] += 1

        # ---- Seccion 3.1: ANALISIS EUROPA (incluye AUH) ----
        # Regla especial: en Europa SOLO SUPINT/SUPINTN cuentan como supervisor valido.
        # Un vuelo a Europa que solo lleve SUPNAL se considera SIN supervisor.
        # AUH (OTRO_CONTINENTE) se suma junto con Europa en TODO este analisis 3.1.
        REGIONES_EUROPA_AMPLIADA = {"EUROPA", "OTRO_CONTINENTE"}  # AUH esta en OTRO_CONTINENTE
        eur_total_legs = 0
        eur_con_supint = 0        # legs con >=1 SUPINT o SUPINTN
        eur_solo_intnal = 0       # legs SIN SUPINT/SUPINTN pero con INTNAL
        eur_solo_supnal = 0       # legs SIN SUPINT/SUPINTN, solo SUPNAL (cuenta como sin supervisor)
        eur_sin_ninguno = 0       # legs sin SUPINT/SUPINTN/INTNAL/SUPNAL
        # tabla A (nueva): legs con 2 o MAS de una misma categoria, sobre el total de legs
        eur_legs_2plus_categoria = Counter()   # categoria -> nro de legs que llevan 2+ de esa categoria
        # tabla especialistas Europa+AUH por flota
        eur_esp_tot_flota = Counter()   # flota -> total legs
        eur_esp_cov_flota = Counter()   # flota -> legs con especialista
        # listado de vuelos a Europa SIN SUPINT/SUPINTN
        eur_vuelos_sin_supint = []           # (fecha, vuelo, "cat1, cat2, ...")

        for (fecha, vuelo, aeropuerto), d in leg_map.items():
            if d["region"] not in REGIONES_EUROPA_AMPLIADA:
                continue
            cats = d["cat_counts"]
            eur_total_legs += 1
            tiene_supint = (cats.get("SUPINT", 0) + cats.get("SUPINTN", 0)) > 0
            tiene_intnal = cats.get("INTNAL", 0) > 0
            tiene_supnal = cats.get("SUPNAL", 0) > 0

            for cat in CATEGORIAS_ORDEN:
                if cats.get(cat, 0) >= 2:
                    eur_legs_2plus_categoria[cat] += 1

            # especialistas por flota (Europa+AUH)
            eur_esp_tot_flota[d["flota"]] += 1
            if d["esp"]:
                eur_esp_cov_flota[d["flota"]] += 1

            if tiene_supint:
                eur_con_supint += 1
            else:
                cats_presentes = ", ".join(f"{c}({cats[c]})" for c in CATEGORIAS_ORDEN if cats.get(c, 0) > 0)
                eur_vuelos_sin_supint.append((fecha, vuelo.strip(), cats_presentes))
                if tiene_intnal:
                    eur_solo_intnal += 1
                elif tiene_supnal:
                    eur_solo_supnal += 1
                else:
                    eur_sin_ninguno += 1

        def _sort_key_vuelo(item):
            fecha, vuelo, _ = item
            try:
                num = int(vuelo.split()[1])
            except (IndexError, ValueError):
                num = 999999
            return (fecha, num)
        eur_vuelos_sin_supint.sort(key=_sort_key_vuelo)

        # ---- Seccion 4 y 5: viaticos / block time (LINEA con validacion), con min/max nombrado ----
        kpi_base = kpi_df[kpi_df["HOMEBASE"] == base].copy()

        def estado_corregido(cid):
            if cid in validacion_estado:
                return validacion_estado[cid][0]
            return estado_map.get(cid, "")

        kpi_base["ESTADO_FINAL"] = kpi_base["CREW ID"].apply(estado_corregido)
        linea = kpi_base[kpi_base["ESTADO_FINAL"] == "LINEA"].copy()
        linea["BLOCK_DEC"] = linea["BLOCK HOURS"].apply(hhmm_a_horas)
        linea["TARGET_BLOCK_DEC"] = linea["TARGET PER DIEM"].apply(hhmm_a_horas)

        total_block_hours_base = linea["BLOCK_DEC"].sum()

        indicadores_viaticos = {}
        for cat in CATEGORIAS_ORDEN:
            g = linea[linea["CATEGORY"] == cat]
            if len(g) == 0:
                continue
            fila_min = g.loc[g["ALLOWANCES"].idxmin()]
            fila_max = g.loc[g["ALLOWANCES"].idxmax()]
            fila_bmin = g.loc[g["BLOCK_DEC"].idxmin()] if g["BLOCK_DEC"].notna().any() else None
            fila_bmax = g.loc[g["BLOCK_DEC"].idxmax()] if g["BLOCK_DEC"].notna().any() else None
            indicadores_viaticos[cat] = {
                "n": len(g),
                "prom_allow": g["ALLOWANCES"].mean(),
                "target_allow": g["TARGET ALLOWANCES"].mean(),
                "min_allow": fila_min["ALLOWANCES"], "min_allow_nombre": fila_min["CREW MEMBER NAME"],
                "max_allow": fila_max["ALLOWANCES"], "max_allow_nombre": fila_max["CREW MEMBER NAME"],
                "prom_block": g["BLOCK_DEC"].mean(),
                "target_block": g["TARGET_BLOCK_DEC"].mean(),
                "min_block": fila_bmin["BLOCK_DEC"] if fila_bmin is not None else None,
                "min_block_nombre": fila_bmin["CREW MEMBER NAME"] if fila_bmin is not None else "",
                "max_block": fila_bmax["BLOCK_DEC"] if fila_bmax is not None else None,
                "max_block_nombre": fila_bmax["CREW MEMBER NAME"] if fila_bmax is not None else "",
            }

        if aeropuertos_sin_clasificar:
            st.warning(
                "⚠️ Estos aeropuertos no están en ninguna lista de continente y no se contaron en "
                "ningún grupo regional (avísame para agregarlos): " + ", ".join(sorted(x for x in aeropuertos_sin_clasificar if x))
            )
        if vuelos_sin_ruta:
            st.warning(
                f"⚠️ {len(vuelos_sin_ruta):,} número(s) de vuelo no se encontraron en la Flight List "
                "(no se pudo determinar su ruta real), así que se clasificaron usando el campo del tx time "
                "como respaldo — esto puede ser menos preciso. Números: " +
                ", ".join(str(v) for v in sorted(vuelos_sin_ruta) if v is not None)
            )

        # =====================================================================
        # ESCRIBIR EXCEL DE SALIDA
        # =====================================================================
        wb_out = openpyxl.Workbook()
        ws = wb_out.active
        ws.title = "RESUMEN"

        ARIAL = "Arial"
        TITLE_FONT = Font(name=ARIAL, bold=True, size=15, color="1F4E78")
        SECTION_FONT = Font(name=ARIAL, bold=True, size=12, color="FFFFFF")
        SECTION_FILL = PatternFill("solid", fgColor="1F4E78")
        SUBHEAD_FONT = Font(name=ARIAL, bold=True, size=10, color="FFFFFF")
        SUBHEAD_FILL = PatternFill("solid", fgColor="2E75B6")
        BODY_FONT = Font(name=ARIAL, size=10)
        LABEL_FONT = Font(name=ARIAL, bold=True, size=10)
        NEUTRAL_FILL = PatternFill("solid", fgColor="F2F2F2")
        NOTE_FONT = Font(name=ARIAL, italic=True, size=8, color="808080")

        def sc(coord, value, font=BODY_FONT, fill=None, numfmt=None):
            cell = ws[coord]
            cell.value = value
            cell.font = font
            if fill:
                cell.fill = fill
            if numfmt:
                cell.number_format = numfmt
            return cell

        row = 2
        sc(f"B{row}", f"TABLERO DE INDICADORES — BASE {base}", TITLE_FONT)
        row += 1
        sc(f"B{row}", f"Generado {datetime.now().strftime('%Y-%m-%d %H:%M')} — DH excluidos de todos los cálculos", NOTE_FONT)
        row += 1
        sc(f"B{row}", f"Archivo tx time usado: {nombre_archivo_nlc}  |  Periodo programado (FHDR): {periodo_txt}", NOTE_FONT)
        row += 2

        # ---- Resumen operativo ----
        sc(f"B{row}", "RESUMEN OPERATIVO DE LA BASE", SECTION_FONT, SECTION_FILL)
        row += 1
        sc(f"B{row}", "Total Block Hours de la base (LINEA)", LABEL_FONT)
        sc(f"C{row}", round(total_block_hours_base, 2) if pd.notna(total_block_hours_base) else None)
        row += 2

        sc(f"B{row}", "LEGS ÚNICOS POR AVIÓN (FLOTA) — cada leg cuenta 1 sola vez, sin DH", LABEL_FONT)
        row += 1
        sc(f"B{row}", "FLOTA", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"C{row}", "TOTAL LEGS", SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for flota in ["B787", "A330", "A320", "ATR72", "737MAX", "OTRO"]:
            n = legs_por_flota.get(flota, 0)
            if n == 0:
                continue
            sc(f"B{row}", flota)
            sc(f"C{row}", n)
            row += 1
        sc(f"B{row}", "TOTAL BASE (todas las flotas, sin DH)", LABEL_FONT, NEUTRAL_FILL)
        sc(f"C{row}", total_legs_base, LABEL_FONT, NEUTRAL_FILL)
        row += 2

        # ---- Seccion 1: Supervisores ----
        sc(f"B{row}", "1. SUPERVISORES POR VUELO — distribución y cobertura por continente", SECTION_FONT, SECTION_FILL)
        row += 1
        headers1 = ["CONTINENTE", "TOTAL LEGS", "LEGS 1 SUP.", "LEGS 2 SUP.", "LEGS 3+ SUP.",
                    "LEGS SIN SUP.", "% COBERTURA SUPINT", "% COBERTURA SUPNAL (respaldo)"]
        for i, h in enumerate(headers1):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for region in REGIONES_ORDEN:
            tot = sup_total_legs.get(region, 0)
            if tot == 0:
                continue
            sc(f"B{row}", REGIONES_LABEL[region])
            sc(f"C{row}", tot)
            sc(f"D{row}", sup_dist[region].get(1, 0))
            sc(f"E{row}", sup_dist[region].get(2, 0))
            sc(f"F{row}", sup_dist[region].get("3+", 0))
            sc(f"G{row}", sup_sin_ninguno.get(region, 0))
            sc(f"H{row}", sup_con_supint.get(region, 0) / tot, numfmt="0.0%")
            sc(f"I{row}", sup_con_supnal.get(region, 0) / tot, numfmt="0.0%")
            row += 1
        # Fila de TOTAL (suma de todos los continentes)
        tot_g = sum(sup_total_legs.values())
        d1_g = sum(sup_dist[r].get(1, 0) for r in REGIONES_ORDEN)
        d2_g = sum(sup_dist[r].get(2, 0) for r in REGIONES_ORDEN)
        d3_g = sum(sup_dist[r].get("3+", 0) for r in REGIONES_ORDEN)
        sin_g = sum(sup_sin_ninguno.values())
        supint_g = sum(sup_con_supint.values())
        supnal_g = sum(sup_con_supnal.values())
        sc(f"B{row}", "TOTAL (todos los continentes)", LABEL_FONT, NEUTRAL_FILL)
        sc(f"C{row}", tot_g, LABEL_FONT, NEUTRAL_FILL)
        sc(f"D{row}", d1_g, LABEL_FONT, NEUTRAL_FILL)
        sc(f"E{row}", d2_g, LABEL_FONT, NEUTRAL_FILL)
        sc(f"F{row}", d3_g, LABEL_FONT, NEUTRAL_FILL)
        sc(f"G{row}", sin_g, LABEL_FONT, NEUTRAL_FILL)
        sc(f"H{row}", (supint_g / tot_g) if tot_g else None, LABEL_FONT, NEUTRAL_FILL, "0.0%")
        sc(f"I{row}", (supnal_g / tot_g) if tot_g else None, LABEL_FONT, NEUTRAL_FILL, "0.0%")
        row += 1
        sc(f"B{row}", "Nota: %SUPNAL solo cuenta los legs donde NO hay ya un supervisor internacional "
                      "(SUPINT/SUPINTN) — es decir, SUPNAL cubre como respaldo. Por eso %SUPINT + %SUPNAL + "
                      "(LEGS SIN SUP./TOTAL) = 100% en cada continente.", NOTE_FONT)
        row += 1
        sc(f"B{row}", "IMPORTANTE (Europa): en vuelos a Europa el SUPNAL NO cuenta como supervisor. Solo "
                      "SUPINT/SUPINTN son supervisor válido. Ver el detalle en la Sección 3.1.",
           Font(name=ARIAL, italic=True, bold=True, size=8, color="C00000"))
        row += 2

        # ---- Seccion 2: Vuelos por categoria x continente ----
        sc(f"B{row}", "2. VUELOS POR CATEGORÍA (cantidad de asignaciones tripulante-leg)", SECTION_FONT, SECTION_FILL)
        row += 1
        headers2 = ["CATEGORIA"] + [REGIONES_LABEL[r] for r in REGIONES_ORDEN]
        for i, h in enumerate(headers2):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for cat in CATEGORIAS_ORDEN:
            sc(f"B{row}", cat)
            for i, region in enumerate(REGIONES_ORDEN):
                sc(f"{get_column_letter(3+i)}{row}", cat_legs.get((cat, region), 0))
            row += 1
        row += 2

        # ---- Seccion 3: Especialistas por continente x flota ----
        sc(f"B{row}", "3. ESPECIALISTAS — % de cobertura por continente y flota", SECTION_FONT, SECTION_FILL)
        row += 1
        flotas_orden = ["B787", "A330", "A320"]
        headers3 = ["CONTINENTE"] + [f"{f} - LEGS" for f in flotas_orden] + [f"{f} - % COB." for f in flotas_orden]
        for i, h in enumerate(headers3):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for region in REGIONES_ORDEN:
            fila_vals = [esp_tot_rf.get((region, f), 0) for f in flotas_orden]
            if sum(fila_vals) == 0:
                continue
            sc(f"B{row}", REGIONES_LABEL[region])
            for i, f in enumerate(flotas_orden):
                tot = esp_tot_rf.get((region, f), 0)
                sc(f"{get_column_letter(3+i)}{row}", tot)
            for i, f in enumerate(flotas_orden):
                tot = esp_tot_rf.get((region, f), 0)
                cov = esp_cov_rf.get((region, f), 0)
                sc(f"{get_column_letter(3+len(flotas_orden)+i)}{row}", (cov/tot) if tot else None, numfmt="0.0%")
            row += 1
        row += 2

        # ---- Seccion 3.1: ANALISIS EUROPA (incluye AUH) ----
        sc(f"B{row}", "3.1  ANÁLISIS EUROPA (incluye AUH) — cobertura de supervisión internacional", SECTION_FONT, SECTION_FILL)
        row += 1
        sc(f"B{row}", "Regla especial: en Europa/AUH SOLO cuentan como supervisor válido SUPINT y SUPINTN. "
                      "Un vuelo que solo lleve SUPNAL se considera SIN supervisor internacional.", NOTE_FONT)
        row += 2

        # Tabla A: legs con 2 o MAS de una misma categoria (sobre el total de legs a Europa/AUH)
        sc(f"B{row}", "A) De los legs a Europa/AUH, ¿cuántos llevan 2 o más tripulantes de una misma categoría?", LABEL_FONT)
        row += 1
        sc(f"B{row}", f"(Total de legs a Europa/AUH, sin DH: {eur_total_legs})", NOTE_FONT)
        row += 1
        sc(f"B{row}", "CATEGORIA", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"C{row}", "LEGS CON 2+ DE ESA CATEGORÍA", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"D{row}", "% SOBRE TOTAL LEGS EUROPA/AUH", SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for cat in CATEGORIAS_ORDEN:
            n = eur_legs_2plus_categoria.get(cat, 0)
            sc(f"B{row}", cat)
            sc(f"C{row}", n)
            sc(f"D{row}", (n / eur_total_legs) if eur_total_legs else None, numfmt="0.0%")
            row += 1
        sc(f"B{row}", "Nota: cada leg puede llevar 2+ de varias categorías, por eso estas filas NO se suman "
                      "entre sí (no son excluyentes). El denominador de cada % es el total de legs a Europa/AUH.", NOTE_FONT)
        row += 2

        # Tabla A2: ESPECIALISTAS en Europa+AUH por flota
        sc(f"B{row}", "A2) Especialistas en los legs a Europa/AUH, por flota", LABEL_FONT)
        row += 1
        sc(f"B{row}", "FLOTA", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"C{row}", "TOTAL LEGS", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"D{row}", "LEGS CON ESPECIALISTA", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"E{row}", "% COBERTURA", SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for flota in ["B787", "A330", "A320", "OTRO"]:
            t = eur_esp_tot_flota.get(flota, 0)
            if t == 0:
                continue
            c = eur_esp_cov_flota.get(flota, 0)
            sc(f"B{row}", flota)
            sc(f"C{row}", t)
            sc(f"D{row}", c)
            sc(f"E{row}", (c / t) if t else None, numfmt="0.0%")
            row += 1
        sc(f"B{row}", "TOTAL EUROPA/AUH", LABEL_FONT, NEUTRAL_FILL)
        sc(f"C{row}", sum(eur_esp_tot_flota.values()), LABEL_FONT, NEUTRAL_FILL)
        sc(f"D{row}", sum(eur_esp_cov_flota.values()), LABEL_FONT, NEUTRAL_FILL)
        _tot_eur_esp = sum(eur_esp_tot_flota.values())
        sc(f"E{row}", (sum(eur_esp_cov_flota.values()) / _tot_eur_esp) if _tot_eur_esp else None,
           LABEL_FONT, NEUTRAL_FILL, "0.0%")
        row += 2

        # Tabla B: resumen de cobertura de supervision internacional
        sc(f"B{row}", "B) Cobertura de supervisión internacional (SUPINT/SUPINTN) de los legs a Europa/AUH", LABEL_FONT)
        row += 1
        sc(f"B{row}", "CONDICIÓN", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"C{row}", "LEGS", SUBHEAD_FONT, SUBHEAD_FILL)
        sc(f"D{row}", "% DEL TOTAL", SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        filas_b = [
            ("CON SUPINT / SUPINTN (cubiertos)", eur_con_supint),
            ("SIN SUPINT/SUPINTN, pero con INTNAL", eur_solo_intnal),
            ("SIN SUPINT/SUPINTN, solo SUPNAL (= sin supervisor)", eur_solo_supnal),
            ("SIN SUPINT/SUPINTN/INTNAL/SUPNAL", eur_sin_ninguno),
        ]
        for etiqueta, n in filas_b:
            sc(f"B{row}", etiqueta)
            sc(f"C{row}", n)
            sc(f"D{row}", (n / eur_total_legs) if eur_total_legs else None, numfmt="0.0%")
            row += 1
        sc(f"B{row}", "TOTAL LEGS A EUROPA/AUH", LABEL_FONT, NEUTRAL_FILL)
        sc(f"C{row}", eur_total_legs, LABEL_FONT, NEUTRAL_FILL)
        sc(f"D{row}", 1.0 if eur_total_legs else None, LABEL_FONT, NEUTRAL_FILL, "0.0%")
        row += 1
        legs_sin_supervisor_eur = eur_solo_intnal + eur_solo_supnal + eur_sin_ninguno
        sc(f"B{row}", f"⚠️ LEGS A EUROPA/AUH SIN SUPERVISOR INTERNACIONAL (SUPINT/SUPINTN): {legs_sin_supervisor_eur}",
           Font(name=ARIAL, bold=True, size=10, color="C00000"))
        row += 2

        # Tabla C: listado de vuelos a Europa SIN SUPINT/SUPINTN
        sc(f"B{row}", "C) Listado de vuelos a Europa/AUH SIN SUPINT/SUPINTN (fecha, vuelo y categorías que llevó)", LABEL_FONT)
        row += 1
        if eur_vuelos_sin_supint:
            sc(f"B{row}", "FECHA", SUBHEAD_FONT, SUBHEAD_FILL)
            sc(f"C{row}", "VUELO", SUBHEAD_FONT, SUBHEAD_FILL)
            sc(f"D{row}", "CATEGORÍAS QUE LLEVÓ (cantidad)", SUBHEAD_FONT, SUBHEAD_FILL)
            row += 1
            for fecha, vuelo, cats_txt in eur_vuelos_sin_supint:
                # formatear fecha yyyymmdd -> dd/mm/yyyy
                fecha_fmt = f"{fecha[6:8]}/{fecha[4:6]}/{fecha[0:4]}" if len(fecha) == 8 and fecha.isdigit() else fecha
                sc(f"B{row}", fecha_fmt)
                sc(f"C{row}", vuelo)
                sc(f"D{row}", cats_txt)
                row += 1
        else:
            sc(f"B{row}", "✅ Todos los vuelos a Europa/AUH tienen al menos un SUPINT o SUPINTN.",
               Font(name=ARIAL, bold=True, size=10, color="006100"))
            row += 1
        row += 2

        # ---- Seccion 4: Viaticos ----
        sc(f"B{row}", "4. VIÁTICOS (ALLOWANCES) — solo ESTADO = LINEA (con validación aplicada)", SECTION_FONT, SECTION_FILL)
        row += 1
        headers4 = ["CATEGORIA", "N", "PROMEDIO ($)", "TARGET PROM ($)",
                    "MÍNIMO ($)", "TRIPULANTE (MENOR)", "MÁXIMO ($)", "TRIPULANTE (MAYOR)"]
        for i, h in enumerate(headers4):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for cat in CATEGORIAS_ORDEN:
            if cat not in indicadores_viaticos:
                continue
            d = indicadores_viaticos[cat]
            sc(f"B{row}", cat)
            sc(f"C{row}", d["n"])
            sc(f"D{row}", d["prom_allow"], numfmt="$#,##0.00")
            sc(f"E{row}", d["target_allow"], numfmt="$#,##0.00")
            sc(f"F{row}", d["min_allow"], numfmt="$#,##0.00")
            sc(f"G{row}", d["min_allow_nombre"])
            sc(f"H{row}", d["max_allow"], numfmt="$#,##0.00")
            sc(f"I{row}", d["max_allow_nombre"])
            row += 1
        row += 2

        # ---- Seccion 5: Block time ----
        sc(f"B{row}", "5. BLOCK TIME (horas) — solo ESTADO = LINEA", SECTION_FONT, SECTION_FILL)
        row += 1
        headers5 = ["CATEGORIA", "N", "PROMEDIO (h)", "TARGET PROM (h)",
                    "MÍNIMO (h)", "TRIPULANTE (MENOR)", "MÁXIMO (h)", "TRIPULANTE (MAYOR)"]
        for i, h in enumerate(headers5):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for cat in CATEGORIAS_ORDEN:
            if cat not in indicadores_viaticos:
                continue
            d = indicadores_viaticos[cat]
            sc(f"B{row}", cat)
            sc(f"C{row}", d["n"])
            sc(f"D{row}", d["prom_block"], numfmt="0.00")
            sc(f"E{row}", d["target_block"], numfmt="0.00")
            sc(f"F{row}", d["min_block"], numfmt="0.00")
            sc(f"G{row}", d["min_block_nombre"])
            sc(f"H{row}", d["max_block"], numfmt="0.00")
            sc(f"I{row}", d["max_block_nombre"])
            row += 1
        row += 2

        # ---- Seccion Validacion ----
        sc(f"B{row}", "VALIDACIÓN DE ESTADO (tripulantes LINEA con código especial detectado en el tx time)", SECTION_FONT, SECTION_FILL)
        row += 1
        for i, h in enumerate(["CREW ID", "CÓDIGO DETECTADO", "DÍAS"]):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for cid, (codigo, dias) in validacion_estado.items():
            if estado_map.get(cid) == "LINEA":
                sc(f"B{row}", cid)
                sc(f"C{row}", codigo)
                sc(f"D{row}", dias)
                row += 1

        for col, w in zip("BCDEFGHI", [34, 14, 18, 16, 12, 24, 12, 24]):
            ws.column_dimensions[col].width = w

        buffer = io.BytesIO()
        wb_out.save(buffer)
        buffer.seek(0)

    st.success("✅ ¡Tablero generado!")
    st.download_button(
        label="⬇️ Descargar tablero en Excel",
        data=buffer,
        file_name=f"TABLERO_INDICADORES_{base}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
