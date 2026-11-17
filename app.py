import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import io
import os
import math
import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Generador de Bitácoras de Combustible",
    page_icon="⛽",
    layout="wide"
)

st.title("⛽ Generador Automatizado de Bitácoras de Combustible")
st.caption("Dirección de Desarrollo Urbano y Medio Ambiente | H. Ayuntamiento de Campeche")

# --- ARCHIVO BASE PERMANENTE ---
RUTAS_PLANTILLA_LOCAL = [
    "7.DESARROLLO URBANO Y MED AMBIENTE.xlsx",
    "plantilla.xlsx",
    "1.FORMATO BITACORA 2024 - 2027 OK (1).xlsx",
    "1.FORMATO BITACORA 2024 - 2027.xlsx"
]

plantilla_detectada = None
for ruta in RUTAS_PLANTILLA_LOCAL:
    if os.path.exists(ruta):
        plantilla_detectada = ruta
        break

# --- CATÁLOGO BASE DE VEHÍCULOS (COMPLETO) ---
VEHICULOS_DEFAULT = [
    {"direccion": "MEDIO AMBIENTE", "placa": "DKR456B",   "tipo": "JIMNY GLX TA, 5 PUERTAS 2025", "combustible": "PREMIUM", "importe_total": 5000.00, "tope_diario": 800.0,  "modo_carga": "DINERO", "anio": "2025", "lleva_odometro": True,  "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "DFT-565-C", "tipo": "JETTA",                           "combustible": "MAGNA",   "importe_total": 5300.00, "tope_diario": 1000.0, "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "CN-2633-B", "tipo": "CAMIONETA (RAM)",                 "combustible": "MAGNA",   "importe_total": 7049.27, "tope_diario": 2000.0, "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "88GWU7",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 1112.75, "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "87GWU8",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 800.00,  "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "88GWU8",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 750.00,  "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "90GWU8",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 700.00,  "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "86GWU8",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 688.13,  "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "86GWU7",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 683.46,  "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "89GWU7",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 500.00,  "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "89GWU8",    "tipo": "MOTOCICLETA",                     "combustible": "MAGNA",   "importe_total": 500.00,  "tope_diario": 200.0,  "modo_carga": "DINERO", "anio": "",     "lleva_odometro": False, "activo": True},
]

PRECIOS = {
    "MAGNA": 23.99,
    "PREMIUM": 26.95,
    "DIESEL": 27.00
}

NOMBRES_DIAS = {0: "Lunes", 1: "Martes", 3: "Jueves", 4: "Viernes"}

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Estado de la Plantilla")
    if plantilla_detectada:
        st.success(f"✅ Plantilla detectada: `{os.path.basename(plantilla_detectada)}`")
    else:
        st.warning("⚠️ No se encontró plantilla en el servidor. Sube una a continuación:")

    uploaded_file = st.file_uploader("Reemplazar plantilla (opcional)", type=["xlsx"])
    
    st.markdown("---")
    st.header("📅 Configuración del Periodo")
    fecha_ini = st.date_input("Inicio del periodo", datetime.date(2026, 9, 26))
    fecha_fin = st.date_input("Fin del periodo", datetime.date(2026, 10, 25))
    
    dias_autorizados = [0, 1, 3, 4]
    fechas_teoricas = []
    cur = fecha_ini
    while cur <= fecha_fin:
        if cur.weekday() in dias_autorizados:
            fechas_teoricas.append(cur)
        cur += datetime.timedelta(days=1)
        
    st.markdown("---")
    st.header("🚫 Días Inhábiles / No Laborados")
    opciones_dias = [f"{d.strftime('%Y-%m-%d')} ({NOMBRES_DIAS[d.weekday()]})" for d in fechas_teoricas]
    dias_inhabiles_seleccionados = st.multiselect(
        "Selecciona días a excluir del calendario:",
        options=opciones_dias,
        help="Los días seleccionados no recibirán ninguna recarga de combustible."
    )
    
    fechas_excluidas_set = {op.split(" ")[0] for op in dias_inhabiles_seleccionados}
    fechas_disponibles = [d for d in fechas_teoricas if d.strftime("%Y-%m-%d") not in fechas_excluidas_set]
    
    st.info(f"📅 **{len(fechas_disponibles)} fechas activas** para asignación de recargas.")

    st.markdown("---")
    st.header("📝 Textos del Formato")
    fecha_entrega_txt = st.text_input("Fecha en Bitácora (Fila 27)", "25 DE OCTUBRE DE 2026")
    destino_txt = st.text_input("Destino o Uso (Fila 28)", "RECORRIDOS OPERATIVOS, SUPERVISIÓN E INSPECCIÓN EN EL MUNICIPIO")
    responsable_txt = st.text_input("Responsable Operativo sugerido", "OPERADOR ASIGNADO")
    
    st.markdown("---")
    st.header("🚗 Odómetro (Jimny / Autos)")
    km_inicial_jimny = st.number_input("Odómetro Inicial (Último Registro)", value=40128, step=1)
    rendimiento_jimny = st.number_input("Rendimiento (km/l)", value=14, step=1)

# Validación de seguridad
if not fechas_disponibles:
    st.error("⚠️ Se han excluido todas las fechas del periodo. Desmarca al menos un día hábil en la barra lateral.")
    st.stop()

# --- FUNCIONES DE BALANCEO ---
def balancear_importes(total, max_day):
    n_cargas = math.ceil(total / max_day)
    if n_cargas > 11:
        n_cargas = 11
        max_day = math.ceil(total / 11)

    step = 200 if max_day >= 800 else (100 if max_day >= 500 else 20)
    cents = round(total - math.floor(total), 2)
    int_total = int(round(total - cents))
    high = int(max_day)
    loads = []
    found = False
    
    for cand_high in range(high, int(total / n_cargas) - 1, -step if step > 0 else -20):
        cand_low = cand_high - step
        if cand_low <= 0: continue
        diff = int_total - n_cargas * cand_low
        if diff >= 0 and step > 0 and diff % step == 0:
            k = diff // step
            if 0 <= k <= n_cargas:
                highs = [cand_high] * k
                lows = [cand_low] * (n_cargas - k)
                alternated = []
                while highs or lows:
                    if highs: alternated.append(highs.pop(0))
                    if lows: alternated.append(lows.pop(0))
                loads = alternated
                found = True
                break
                
    if not found:
        base = int((int_total / n_cargas) // step * step) if step > 0 else int(int_total / n_cargas)
        loads = [base] * n_cargas
        rem = int_total - sum(loads)
        sub_step = 50 if max_day >= 500 else 10
        while rem >= sub_step:
            min_i = min(range(n_cargas), key=lambda i: loads[i])
            if loads[min_i] + sub_step <= max_day:
                loads[min_i] += sub_step
                rem -= sub_step
            else:
                break
        if rem > 0:
            min_i = min(range(n_cargas), key=lambda i: loads[i])
            loads[min_i] += rem
            
    loads[-1] = round(loads[-1] + cents, 2)
    return loads

def balancear_litros(total_litros, max_litros):
    n_cargas = max(1, math.ceil(total_litros / max_litros))
    if n_cargas > 11: n_cargas = 11
    step = 10 if max_litros >= 40 else (5 if max_litros >= 15 else 1)

    for cand_high in range(int(max_litros), max(int(total_litros / n_cargas) - 1, step - 1), -step):
        cand_low = cand_high - step
        if cand_low <= 0: continue
        diff = total_litros - n_cargas * cand_low
        if diff >= 0 and abs(diff / step - round(diff / step)) < 0.001:
            k = round(diff / step)
            if 0 <= k <= n_cargas:
                highs = [cand_high] * k
                lows  = [cand_low]  * (n_cargas - k)
                alt = []
                while highs or lows:
                    if highs: alt.append(highs.pop(0))
                    if lows:  alt.append(lows.pop(0))
                rem = round(total_litros - sum(alt), 2)
                alt[-1] = round(alt[-1] + rem, 2)
                return alt

    base = round(total_litros / n_cargas, 2)
    loads = [base] * (n_cargas - 1)
    loads.append(round(total_litros - sum(loads), 2))
    return loads

def planificar_cargas_vehiculo(importe_total, tope, fechas, combustible, modo="DINERO"):
    precio = PRECIOS.get(combustible, 23.99)
    if modo == "LITROS":
        total_litros = importe_total / precio
        litros   = balancear_litros(total_litros, tope)
        importes = [round(l * precio, 2) for l in litros]
        diff = round(importe_total - sum(importes), 2)
        if abs(diff) >= 0.01:
            importes[-1] = round(importes[-1] + diff, 2)
            litros[-1]   = round(importes[-1] / precio, 2)
    else:
        importes = balancear_importes(importe_total, tope)
        litros   = [round(i / precio, 2) for i in importes]

    n = len(importes)
    step_f = len(fechas) / n
    fechas_asignadas = [fechas[int(round(i * step_f))] for i in range(n)]

    return [
        {"fecha": f.strftime("%d/%m/%Y"),
         "precio_unit": precio,
         "importe": imp,
         "litros": lit}
        for f, imp, lit in zip(fechas_asignadas, importes, litros)
    ]

# --- DETECCIÓN DE PLANTILLA ---
wb_temp = None
if uploaded_file is not None:
    wb_temp = openpyxl.load_workbook(uploaded_file)
elif plantilla_detectada:
    wb_temp = openpyxl.load_workbook(plantilla_detectada)

def detectar_tipo_plantilla(wb):
    if not wb: return "BITACORA"
    if "BITACORA VEHICULOS" in wb.sheetnames: return "BITACORA"
    ws0 = wb[wb.sheetnames[0]]
    for row in ws0.iter_rows(max_row=10, values_only=True):
        for cell in row:
            if cell and "PEMEX" in str(cell).upper():
                return "PEMEX"
    return "BITACORA"

tipo_plantilla = detectar_tipo_plantilla(wb_temp)
st.markdown(f"### 🧾 Plantilla detectada: **{tipo_plantilla}**")

# --- TABLA EDITABLE ---
st.subheader("1. Cuotas Asignadas y Parámetros")

if tipo_plantilla == "PEMEX":
    # ... (Manejo especial si se llega a subir el archivo de PEMEX)
    st.caption("Este es un archivo formato PEMEX. Se capturarán los litros en las filas correspondientes.")
    # (La lógica de PEMEX se omitió aquí por brevedad, pero si subes el archivo de PEMEX, 
    # caerá aquí. Sin embargo, el código de abajo está enfocado a tu plantilla BITACORA).
    df_edit = pd.DataFrame(VEHICULOS_DEFAULT) # Fallback para que no truene
else:
    # MODO BITACORA (EL QUE USAS ACTUALMENTE)
    st.caption(
        "Marca **Generar** para incluir el vehículo en la bitácora. "
        "En **Modo** elige `DINERO` (divide el presupuesto en importes ≤ Tope Día $) "
        "o `LITROS` (divide el presupuesto en litros ≤ Tope Día L)."
    )
    
    df_edit = st.data_editor(
        pd.DataFrame(VEHICULOS_DEFAULT),
        column_config={
            "activo": st.column_config.CheckboxColumn("Generar", default=True),
            "modo_carga": st.column_config.SelectboxColumn(
                "Modo",
                options=["DINERO", "LITROS"],
                help="DINERO: divide el presupuesto en importes ≤ Tope Día ($). "
                     "LITROS: divide el presupuesto en litros ≤ Tope Día (L)."
            ),
            "lleva_odometro": st.column_config.CheckboxColumn(
                "Registra Odómetro",
                help="Desmarcado llena con '---' en Odómetro y Rendimiento"
            ),
            "importe_total": st.column_config.NumberColumn("Presupuesto ($)", format="$%.2f"),
            "tope_diario": st.column_config.NumberColumn(
                "Tope Día", format="%.2f",
                help="En $ si Modo=DINERO. En litros si Modo=LITROS."
            ),
            "combustible": st.column_config.SelectboxColumn(
                "Combustible", options=["MAGNA", "PREMIUM", "DIESEL"]
            ),
        },
        use_container_width=True,
        num_rows="dynamic"
    )

# --- TOTALES ---
tot_magna = df_edit[df_edit["activo"] & (df_edit["combustible"] == "MAGNA")]["importe_total"].sum() if "importe_total" in df_edit.columns else 0
tot_prem  = df_edit[df_edit["activo"] & (df_edit["combustible"] == "PREMIUM")]["importe_total"].sum() if "importe_total" in df_edit.columns else 0

col1, col2, col3 = st.columns(3)
col1.metric("Total Magna (CD)", f"${tot_magna:,.2f}")
col2.metric("Total Premium (Vales)", f"${tot_prem:,.2f}")
col3.metric("Total General", f"${tot_magna + tot_prem:,.2f}")

# --- VISTA PREVIA ---
st.markdown("---")
st.subheader("2. Vista Previa de Recargas Balanceadas")
preview_data = []

if tipo_plantilla != "PEMEX":
    for _, r in df_edit.iterrows():
        if not r.get("activo", False): continue

        modo = r.get("modo_carga") or "DINERO"
        if pd.isna(modo): modo = "DINERO"

        c_plan = planificar_cargas_vehiculo(
            r["importe_total"], r["tope_diario"], fechas_disponibles, r["combustible"], modo
        )
        
        desglose_txt = " + ".join(
            f"{c['litros']:.2f}L=${c['importe']:,.2f} ({c['fecha']})" for c in c_plan
        )

        detalle_odo = "---"
        if r.get("lleva_odometro", False):
            total_litros_cargas = sum(c["litros"] for c in c_plan)
            km_estimado = km_inicial_jimny + round(total_litros_cargas * rendimiento_jimny)
            detalle_odo = f"De {km_inicial_jimny:,} a {km_estimado:,} km ({rendimiento_jimny} km/l)"

        preview_data.append({
            "Vehículo": r["tipo"],
            "Placa": r["placa"],
            "Modo": modo,
            "Cargas": len(c_plan),
            "Total": f"${r['importe_total']:,.2f}",
            "Odómetro y Rendimiento": detalle_odo,
            "Fechas / Litros / Importes": desglose_txt,
        })

    if preview_data:
        st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
    else:
        st.info("Marca al menos un vehículo con el checkbox **Generar** para ver la vista previa.")

# --- ESCRITURA EN EXCEL ---
def escribir_hoja_bitacora(ws, datos, cargas):
    centro  = Alignment(horizontal="center", vertical="center")
    derecha = Alignment(horizontal="right",  vertical="center")

    ws["B9"] = f"N° ECONOMICO: {datos.get('num_eco', '')}"
    ws["D9"] = f"MARCA Y TIPO:  {datos.get('tipo', '')}"
    ws["H9"] = f"AÑO: {datos.get('anio', '')}"
    ws["K9"] = f"PLACAS: {datos.get('placa', '')}"

    lleva_odometro = datos.get("lleva_odometro", False)

    for idx, c in enumerate(cargas):
        r = 14 + idx
        if r > 24: break

        ws[f"B{r}"] = c["fecha"]

        cell_costo = ws[f"D{r}"]
        cell_costo.value = c["precio_unit"]
        cell_costo.number_format = '0.00'

        cell_importe = ws[f"E{r}"]
        cell_importe.value = c["importe"]
        cell_importe.number_format = '#,##0.00'

        cell_litros = ws[f"C{r}"]
        cell_litros.value = f"=E{r}/D{r}"
        cell_litros.number_format = '0.00'
        cell_litros.alignment = derecha

        if not lleva_odometro:
            for col_letra in ["F", "G", "H"]:
                cell = ws[f"{col_letra}{r}"]
                cell.value = "---"
                cell.alignment = centro
        else:
            cell_rend = ws[f"H{r}"]
            cell_rend.value = f"{rendimiento_jimny} KM/L"
            cell_rend.alignment = centro

            cell_ini = ws[f"F{r}"]
            cell_ini.value = km_inicial_jimny if idx == 0 else f"=G{r-1}"
            cell_ini.number_format = '#,##0'
            cell_ini.alignment = centro

            cell_fin = ws[f"G{r}"]
            cell_fin.value = f"=F{r}+ROUND(C{r}*{rendimiento_jimny},0)"
            cell_fin.number_format = '#,##0'
            cell_fin.alignment = centro

        ws[f"I{r}"] = responsable_txt

    ws["C25"] = ""
    ws["D25"] = "TOTAL"
    ws["E25"] = "=SUM(E14:E24)"
    ws["E25"].number_format = '$#,##0.00'

    ws["B27"] = f"BITÁCORA CORRESPONDIENTE A LA ENTREGA DEL DÍA: {fecha_entrega_txt}"
    ws["B28"] = f"DESTINO O USO DEL VEHICULO: {destino_txt}"

# --- GENERADOR DE EXCEL ---
st.markdown("---")
st.subheader("3. Descargar Formatos Oficiales")

if st.button("🚀 Generar Libro Excel con Cargas Balanceadas"):
    if uploaded_file is not None:
        wb = openpyxl.load_workbook(uploaded_file)
    elif plantilla_detectada is not None:
        wb = openpyxl.load_workbook(plantilla_detectada)
    else:
        st.error("No se encontró ninguna plantilla precargada ni subida. Por favor sube el archivo Excel.")
        st.stop()

    # Si es plantilla BITACORA
    if tipo_plantilla != "PEMEX":
        ws_base = wb["BITACORA VEHICULOS"]
        generados = 0

        for _, row in df_edit.iterrows():
            if not row.get("activo", False): continue

            modo = row.get("modo_carga") or "DINERO"
            if pd.isna(modo): modo = "DINERO"

            c_plan = planificar_cargas_vehiculo(
                row["importe_total"], row["tope_diario"],
                fechas_disponibles, row["combustible"], modo
            )
            
            cargas_para_excel = [{"fecha": c["fecha"], "precio_unit": c["precio_unit"], "importe": c["importe"]} for c in c_plan]

            ws_nueva = wb.copy_worksheet(ws_base)
            ws_nueva.title = str(row["placa"]).replace("-", "").strip()[:31]
            escribir_hoja_bitacora(ws_nueva, row.to_dict(), cargas_para_excel)
            generados += 1

        if generados == 0:
            st.warning("No hay vehículos marcados con **Generar**. No se creó ninguna hoja.")
        else:
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            st.success(f"✅ Se generaron {generados} bitácora(s) considerando días hábiles y modo de carga elegido.")
            st.download_button(
                label="📥 Descargar Libro Excel (Bitácoras Balanceadas)",
                data=buffer,
                file_name="BITACORAS_JUSTIFICADAS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("La generación automática para el formato PEMEX aún está en desarrollo en esta versión.")
