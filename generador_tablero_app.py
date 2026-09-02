"""
GENERADOR DE TABLERO DE INDICADORES - TRIPULACION POR BASE
============================================================
App web (Streamlit) que guía al usuario paso a paso:
  1) Elegir la base a procesar
  2) Cargar cada archivo requerido, EN ORDEN, con validación inmediata
     (si algo falta o está mal, explica exactamente qué corregir)
  3) Generar y descargar el tablero de indicadores en Excel

CÓMO EJECUTARLA (para el equipo de TI / quien la despliegue):
  1) Instalar dependencias:  pip install streamlit pandas openpyxl xlrd
  2) Ejecutar:               streamlit run generador_tablero_app.py
  3) Se abre en el navegador (local: http://localhost:8501). Para que la use
     todo el equipo, se despliega en un servidor interno o en Streamlit
     Community Cloud / similar, y se comparte el link.

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
# CONSTANTES DE NEGOCIO (mismas reglas ya validadas en el proceso manual)
# =========================================================================
AEROPUERTOS_COLOMBIA = {
    "ADZ", "AXM", "BAQ", "BGA", "BOG", "CLO", "CTG", "CUC", "IBE", "IPI",
    "LET", "MDE", "MTR", "NVA", "PEI", "PPN", "PSO", "RCH", "SMR", "VUP",
}
AEROPUERTOS_EUROPA = {"BCN", "CDG", "LHR", "MAD"}
CODIGOS_ESPECIALES = {"ILLP", "VAC", "LUS", "OFI", "PSI", "XMTR", "MAT", "XMAT", "OFFICE_M"}
CATEGORIAS_ORDEN = ["SUPINT", "SUPINTN", "INTNAL", "SUPNAL", "AUXNAL"]  # de mas antiguo a mas reciente
BASES_VALIDAS = ["BOG", "MDE", "CLO", "BAQ", "CTG", "PEI", "BGA"]


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
    "de indicadores para la base que elijas. Sube los archivos **en el orden pedido**; "
    "cada uno se valida antes de dejarte avanzar."
)

base = st.selectbox("1️⃣ Elige la BASE a procesar", BASES_VALIDAS)
st.divider()

# =========================================================================
# PASO 1: tx time (.nlc)
# =========================================================================
st.subheader("2️⃣ Tx time (.nlc)")
st.caption("El archivo de texto de programación de tripulantes, ej. JCR_AV_CCALL_8_..._estoro.nlc")
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
    st.success(f"✅ Tx time válido — {len(lineas_nlc):,} líneas leídas.")

if lineas_nlc is None:
    st.info("⬆️ Sube el tx time para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 2: KPI
# =========================================================================
st.subheader("3️⃣ Reporte KPI (.xls / .xlsx)")
st.caption("El reporte con ALLOWANCES, TARGET ALLOWANCES, BLOCK HOURS, HOMEBASE, etc. (hoja 'JCR KPIs Report')")
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
st.caption("Archivo con el estado de cada tripulante — columna B = Crew ID, columna G = ESTADO")
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
st.subheader("5️⃣ Planta / RPT_TRIPULACIÓN (.xlsx)")
st.caption("Columnas: CREW CODE, CATEGORIA, APELLIDO, NOMBRE, HOME BASE")
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
st.subheader("6️⃣ Lista de vuelos / Flight List (.xlsx)")
st.caption("Debe tener una hoja llamada 'LT' con columnas de Equip y Flt Num")
flota_file = st.file_uploader("Cargar Flight List", type=["xlsx"], key="flota")

flt_flota = None
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
    if "Equip" not in headers or "Flt Num" not in headers:
        st.error("❌ La hoja 'LT' debe tener columnas llamadas 'Equip' y 'Flt Num'.")
        st.stop()

    idx_equip = headers.index("Equip")
    idx_flt = headers.index("Flt Num")
    flt_equip_counts = defaultdict(Counter)
    for row in ws_f.iter_rows(min_row=2, values_only=True):
        flt = row[idx_flt]
        equip = row[idx_equip]
        if flt is None:
            continue
        flt_equip_counts[int(flt)][equip] += 1

    flt_flota = {}
    for flt, counter in flt_equip_counts.items():
        equip_top = counter.most_common(1)[0][0]
        flt_flota[flt] = fleet_group(equip_top)

    if len(flt_flota) == 0:
        st.error("❌ No encontré ningún número de vuelo válido en la hoja 'LT'.")
        st.stop()
    st.success(f"✅ Flight List válida — {len(flt_flota):,} números de vuelo mapeados a flota.")

if flt_flota is None:
    st.info("⬆️ Sube la Flight List para continuar.")
    st.stop()
st.divider()

# =========================================================================
# PASO 6: Especialistas
# =========================================================================
st.subheader("7️⃣ Especialistas (.xlsx)")
st.caption("Lista simple de Crew ID de tripulantes especialistas (columna A)")
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
# PROCESAMIENTO PRINCIPAL
# =========================================================================
st.subheader("8️⃣ Generar tablero")

if st.button("🚀 Generar tablero de indicadores", type="primary"):
    with st.spinner("Parseando el tx time y calculando indicadores... esto puede tardar 1-2 minutos"):

        # ---- Parsear el tx time ----
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
            elif p[0] in ("PLEG", "PFTR") and cur_nl and len(p) > 3:
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
                if len(q) > 5:
                    crew_legs[cur_crew].append((q[3], q[4].strip(), q[5].strip()))
            elif tipo == "RROL":
                cur_crew = q[1]
                if len(q) > 5 and q[2] == "L":
                    crew_legs[cur_crew].append((q[3], q[4].strip(), q[5].strip()))

        validacion_estado = {}
        for cid, counter in crew_code_days.items():
            mejor_codigo, dias = counter.most_common(1)[0]
            validacion_estado[cid] = (mejor_codigo, dias)

        # ---- Mapa de homebase preferente: KPI, con respaldo en planta ----
        kpi_homebase_map = dict(zip(kpi_df["CREW ID"], kpi_df["HOMEBASE"]))

        def clasif_final(clasif, europa):
            return "NACIONAL" if clasif == "NACIONAL" else ("EUROPA" if europa == "EUROPA" else "AMERICA")

        # ---- Recorrer legs y construir indicadores ----
        leg_map = {}  # (fecha,vuelo,aeropuerto) -> {tipo, flota, cats:set, esp:bool}
        cat_legs = Counter()  # (categoria, tipo) -> n

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
                if aeropuerto in AEROPUERTOS_COLOMBIA:
                    clasif = "NACIONAL"
                elif aeropuerto in AEROPUERTOS_EUROPA:
                    clasif = "EUROPA"
                else:
                    clasif = "AMERICA"

                cat_legs[(categoria, clasif)] += 1

                leg_key = (fecha, vuelo, aeropuerto)
                if leg_key not in leg_map:
                    try:
                        num_vuelo = int(vuelo.split()[1])
                    except (IndexError, ValueError):
                        num_vuelo = None
                    flota = flt_flota.get(num_vuelo, "OTRO")
                    leg_map[leg_key] = {"tipo": clasif, "flota": flota, "cats": set(), "esp": False}
                leg_map[leg_key]["cats"].add(categoria)
                if es_esp:
                    leg_map[leg_key]["esp"] = True

        # ---- Seccion 1 y 3 desde leg_map ----
        leg_totals = Counter()
        leg_sin_sup = Counter()
        esp_tot_d = Counter()
        esp_cov_d = Counter()
        esp_tot_f = Counter()
        esp_cov_f = Counter()

        for d in leg_map.values():
            tipo = d["tipo"]
            leg_totals[tipo] += 1
            tiene_sup_intl = bool(d["cats"] & {"SUPINT", "SUPINTN"})
            tiene_sup_nal = "SUPNAL" in d["cats"]
            if not tiene_sup_intl and not tiene_sup_nal:
                leg_sin_sup[tipo] += 1

            esp_tot_d[tipo] += 1
            if d["esp"]:
                esp_cov_d[tipo] += 1
            esp_tot_f[d["flota"]] += 1
            if d["esp"]:
                esp_cov_f[d["flota"]] += 1

        # ---- Seccion 4 y 5: viaticos / block time (LINEA con validacion) ----
        kpi_base = kpi_df[kpi_df["HOMEBASE"] == base].copy()

        def estado_corregido(cid):
            if cid in validacion_estado:
                return validacion_estado[cid][0]
            return estado_map.get(cid, "")

        kpi_base["ESTADO_FINAL"] = kpi_base["CREW ID"].apply(estado_corregido)
        linea = kpi_base[kpi_base["ESTADO_FINAL"] == "LINEA"].copy()
        linea["BLOCK_DEC"] = linea["BLOCK HOURS"].apply(hhmm_a_horas)
        linea["TARGET_BLOCK_DEC"] = linea["TARGET PER DIEM"].apply(hhmm_a_horas)

        indicadores_viaticos = {}
        for cat in CATEGORIAS_ORDEN:
            g = linea[linea["CATEGORY"] == cat]
            if len(g) == 0:
                continue
            indicadores_viaticos[cat] = {
                "n": len(g),
                "prom_allow": g["ALLOWANCES"].mean(),
                "target_allow": g["TARGET ALLOWANCES"].mean(),
                "min_allow": g["ALLOWANCES"].min(),
                "max_allow": g["ALLOWANCES"].max(),
                "prom_block": g["BLOCK_DEC"].mean(),
                "target_block": g["TARGET_BLOCK_DEC"].mean(),
            }

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
        sc(f"B{row}", f"Generado {datetime.now().strftime('%Y-%m-%d %H:%M')}", Font(name=ARIAL, italic=True, size=9, color="808080"))
        row += 2

        # Seccion 1
        sc(f"B{row}", "1. SUPERVISORES POR VUELO — % de legs SIN supervisor", SECTION_FONT, SECTION_FILL)
        row += 1
        for i, h in enumerate(["TIPO DE DESTINO", "TOTAL LEGS", "% SIN SUPERVISOR"]):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        tot_g, sin_g = 0, 0
        for tipo in ["AMERICA", "EUROPA", "NACIONAL"]:
            tot = leg_totals.get(tipo, 0)
            sin_ = leg_sin_sup.get(tipo, 0)
            tot_g += tot
            sin_g += sin_
            sc(f"B{row}", tipo)
            sc(f"C{row}", tot)
            sc(f"D{row}", (sin_ / tot) if tot else None, numfmt="0.0%")
            row += 1
        sc(f"B{row}", "TOTAL", LABEL_FONT, NEUTRAL_FILL)
        sc(f"C{row}", tot_g, LABEL_FONT, NEUTRAL_FILL)
        sc(f"D{row}", (sin_g / tot_g) if tot_g else None, LABEL_FONT, NEUTRAL_FILL, "0.0%")
        row += 3

        # Seccion 2
        sc(f"B{row}", "2. VUELOS POR CATEGORÍA (cantidad de legs)", SECTION_FONT, SECTION_FILL)
        row += 1
        for i, h in enumerate(["CATEGORIA", "EUROPA", "AMERICA", "NACIONAL"]):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for cat in CATEGORIAS_ORDEN:
            sc(f"B{row}", cat)
            sc(f"C{row}", cat_legs.get((cat, "EUROPA"), 0))
            sc(f"D{row}", cat_legs.get((cat, "AMERICA"), 0))
            sc(f"E{row}", cat_legs.get((cat, "NACIONAL"), 0))
            row += 1
        row += 2

        # Seccion 3
        sc(f"B{row}", "3. ESPECIALISTAS — % de cobertura", SECTION_FONT, SECTION_FILL)
        row += 1
        for i, h in enumerate(["TIPO / FLOTA", "TOTAL LEGS", "% COBERTURA"]):
            sc(f"{get_column_letter(2+i)}{row}", h, SUBHEAD_FONT, SUBHEAD_FILL)
        row += 1
        for tipo in ["AMERICA", "EUROPA", "NACIONAL"]:
            tot = esp_tot_d.get(tipo, 0)
            cov = esp_cov_d.get(tipo, 0)
            sc(f"B{row}", tipo)
            sc(f"C{row}", tot)
            sc(f"D{row}", (cov / tot) if tot else None, numfmt="0.0%")
            row += 1
        for flota in ["B787", "A320", "A330", "ATR72", "737MAX"]:
            tot = esp_tot_f.get(flota, 0)
            if tot == 0:
                continue
            cov = esp_cov_f.get(flota, 0)
            sc(f"B{row}", flota)
            sc(f"C{row}", tot)
            sc(f"D{row}", cov / tot, numfmt="0.0%")
            row += 1
        row += 2

        # Seccion 4
        sc(f"B{row}", "4. VIÁTICOS (ALLOWANCES) — solo ESTADO = LINEA (con validación aplicada)", SECTION_FONT, SECTION_FILL)
        row += 1
        for i, h in enumerate(["CATEGORIA", "N", "PROMEDIO ($)", "TARGET PROM ($)", "MINIMO ($)", "MAXIMO ($)"]):
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
            sc(f"G{row}", d["max_allow"], numfmt="$#,##0.00")
            row += 1
        row += 2

        # Seccion 5
        sc(f"B{row}", "5. BLOCK TIME (horas) — solo ESTADO = LINEA", SECTION_FONT, SECTION_FILL)
        row += 1
        for i, h in enumerate(["CATEGORIA", "N", "PROMEDIO (h)", "TARGET PROM (h)"]):
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
            row += 1
        row += 2

        # Seccion Validacion
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

        for col, w in zip("BCDEFG", [30, 14, 18, 16, 14, 14]):
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
