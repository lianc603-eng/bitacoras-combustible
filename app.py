import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import io
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

# --- CATÁLOGO BASE DE VEHÍCULOS ---
VEHICULOS_DEFAULT = [
    {"direccion": "MEDIO AMBIENTE", "placa": "DKR456B", "tipo": "JIMNY GLX TA, 5 PUERTAS 2025", "combustible": "PREMIUM", "importe_total": 5000.00, "tope_diario": 800.0, "anio": "2025", "lleva_odometro": True, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "DFT-565-C", "tipo": "JETTA", "combustible": "MAGNA", "importe_total": 5300.00, "tope_diario": 1000.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "CN-2633-B", "tipo": "CAMIONETA (RAM)", "combustible": "MAGNA", "importe_total": 7049.27, "tope_diario": 2000.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "88GWU7", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 1112.75, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "87GWU8", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 800.00, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "88GWU8", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 750.00, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "90GWU8", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 700.00, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "86GWU8", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 688.13, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "86GWU7", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 683.46, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "89GWU7", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 500.00, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
    {"direccion": "MEDIO AMBIENTE", "placa": "89GWU8", "tipo": "MOTOCICLETA", "combustible": "MAGNA", "importe_total": 500.00, "tope_diario": 200.0, "anio": "", "lleva_odometro": False, "activo": True},
]

PRECIOS = {
    "MAGNA": 23.99,
    "PREMIUM": 26.95,
    "DIESEL": 27.00
}

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración del Periodo")
    fecha_ini = st.date_input("Inicio del periodo", datetime.date(2026, 9, 26))
    fecha_fin = st.date_input("Fin del periodo", datetime.date(2026, 10, 25))
    
    fecha_entrega_txt = st.text_input("Fecha en Bitácora (Fila 27)", "25 DE OCTUBRE DE 2026")
    destino_txt = st.text_input("Destino o Uso (Fila 28)", "RECORRIDOS OPERATIVOS, SUPERVISIÓN E INSPECCIÓN EN EL MUNICIPIO")
    responsable_txt = st.text_input("Responsable Operativo sugerido", "OPERADOR ASIGNADO")
    
    st.markdown("---")
    uploaded_file = st.file_uploader("Subir Plantilla Excel (.xlsx)", type=["xlsx"])

# Días autorizados: Lunes (0), Martes (1), Jueves (3), Viernes (4)
dias_autorizados = [0, 1, 3, 4]
fechas_disponibles = []
cur = fecha_ini
while cur <= fecha_fin:
    if cur.weekday() in dias_autorizados:
        fechas_disponibles.append(cur)
    cur += datetime.timedelta(days=1)

with st.sidebar:
    st.info(f"📅 **{len(fechas_disponibles)} fechas autorizadas** (L, M, J, V)")

# --- TABLA EDITABLE ---
st.subheader("1. Cuotas Asignadas y Parámetros")
df_edit = st.data_editor(
    pd.DataFrame(VEHICULOS_DEFAULT),
    column_config={
        "activo": st.column_config.CheckboxColumn("Generar", default=True),
        "lleva_odometro": st.column_config.CheckboxColumn("Registra Odómetro", help="Desmarcado llena con '---' en Odómetro y Rendimiento"),
        "importe_total": st.column_config.NumberColumn("Presupuesto ($)", format="$%.2f"),
        "tope_diario": st.column_config.NumberColumn("Tope Día ($)", format="$%.2f"),
        "combustible": st.column_config.SelectboxColumn("Combustible", options=["MAGNA", "PREMIUM", "DIESEL"])
    },
    use_container_width=True,
    num_rows="dynamic"
)

tot_magna = df_edit[df_edit["activo"] & (df_edit["combustible"] == "MAGNA")]["importe_total"].sum()
tot_prem = df_edit[df_edit["activo"] & (df_edit["combustible"] == "PREMIUM")]["importe_total"].sum()
col1, col2, col3 = st.columns(3)
col1.metric("Total Magna (CD)", f"${tot_magna:,.2f}")
col2.metric("Total Premium (Vales)", f"${tot_prem:,.2f}")
col3.metric("Total General", f"${tot_magna + tot_prem:,.2f}")

# --- BALANCEO DE CARGAS ---
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
        if cand_low <= 0:
            continue
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

def planificar_cargas_vehiculo(total_monto, tope_dia, fechas, combustible):
    precio_unit = PRECIOS.get(combustible, 23.99)
    importes = balancear_importes(total_monto, tope_dia)
    n = len(importes)
    
    step_f = len(fechas) / n
    fechas_asignadas = [fechas[int(round(i * step_f))] for i in range(n)]
    
    cargas = []
    for f, imp in zip(fechas_asignadas, importes):
        cargas.append({
            "fecha": f.strftime("%d/%m/%Y"),
            "precio_unit": precio_unit,
            "importe": imp
        })
    return cargas

# --- VISTA PREVIA ---
st.markdown("---")
st.subheader("2. Vista Previa de Recargas Balanceadas")
preview_data = []
for _, r in df_edit.iterrows():
    if not r["activo"]: continue
    c_plan = planificar_cargas_vehiculo(r["importe_total"], r["tope_diario"], fechas_disponibles, r["combustible"])
    desglose_txt = " + ".join([f"${c['importe']:,.2f}" for c in c_plan])
    preview_data.append({
        "Vehículo": r["tipo"],
        "Placa": r["placa"],
        "Cargas": len(c_plan),
        "Total": f"${r['importe_total']:,.2f}",
        "Odómetro/Rendimiento": "Con datos" if r.get("lleva_odometro", False) else "---",
        "Desglose de recargas ($)": desglose_txt
    })
st.dataframe(pd.DataFrame(preview_data), use_container_width=True)

# --- ESCRITURA EN EXCEL (CON FÓRMULA DE LITROS: =E/D) ---
def escribir_hoja_bitacora(ws, datos, cargas):
    centro = Alignment(horizontal="center", vertical="center")
    derecha = Alignment(horizontal="right", vertical="center")
    
    # 1. ENCABEZADO (Fila 9)
    num_eco = datos.get('num_eco', '')
    ws["B9"] = f"N° ECONOMICO: {num_eco}" if num_eco else "N° ECONOMICO: "
    ws["D9"] = f"MARCA Y TIPO:  {datos.get('tipo', '')}"
    ws["H9"] = f"AÑO: {datos.get('anio', '')}"
    ws["K9"] = f"PLACAS: {datos.get('placa', '')}"
    
    lleva_odometro = datos.get("lleva_odometro", False)
    
    # 2. RENGLONES DE CARGA (Filas 14 a 24)
    for idx, c in enumerate(cargas):
        r = 14 + idx
        if r > 24: break
        
        # B: Fecha
        ws[f"B{r}"] = c["fecha"]
        
        # D: Costo de Litro
        cell_costo = ws[f"D{r}"]
        cell_costo.value = c["precio_unit"]
        cell_costo.number_format = '0.00'
        
        # E: Importe ($)
        cell_importe = ws[f"E{r}"]
        cell_importe.value = c["importe"]
        cell_importe.number_format = '#,##0.00'
        
        # C: LITROS = Importe / Costo de Litro (=E{r}/D{r})
        cell_litros = ws[f"C{r}"]
        cell_litros.value = f"=E{r}/D{r}"
        cell_litros.number_format = '0.00'
        cell_litros.alignment = derecha
        
        # F, G, H: Odómetros y Rendimiento
        if not lleva_odometro:
            for col_letra in ["F", "G", "H"]:
                cell = ws[f"{col_letra}{r}"]
                cell.value = "---"
                cell.alignment = centro
        else:
            ws[f"F{r}"] = ""
            ws[f"G{r}"] = ""
            ws[f"H{r}"] = ""
            
        # I: Responsable Operativo
        ws[f"I{r}"] = responsable_txt
        
    # 3. TOTALES (Fila 25)
    ws["C25"] = "=SUM(C14:C24)"
    ws["C25"].number_format = '0.000'
    ws["D25"] = "TOTAL"
    ws["E25"] = "=SUM(E14:E24)"
    ws["E25"].number_format = '$#,##0.00'
    
    # 4. PIE DE PÁGINA
    ws["B27"] = f"BITÁCORA CORRESPONDIENTE A LA ENTREGA DEL DÍA: {fecha_entrega_txt}"
    ws["B28"] = f"DESTINO O USO DEL VEHICULO: {destino_txt}"

# --- GENERADOR DE EXCEL ---
st.markdown("---")
st.subheader("3. Descargar Formatos Oficiales")

if st.button("🚀 Generar Libro Excel con Cargas Balanceadas"):
    if uploaded_file is not None:
        wb = openpyxl.load_workbook(uploaded_file)
    else:
        try:
            wb = openpyxl.load_workbook("1.FORMATO BITACORA 2024 - 2027 OK (1).xlsx")
        except FileNotFoundError:
            try:
                wb = openpyxl.load_workbook("1.FORMATO BITACORA 2024 - 2027.xlsx")
            except FileNotFoundError:
                st.error("Por favor sube la plantilla de Excel en la barra lateral izquierda.")
                st.stop()
            
    ws_base = wb["BITACORA VEHICULOS"]
    
    for _, row in df_edit.iterrows():
        if not row["activo"]: continue
        
        cargas = planificar_cargas_vehiculo(
            row["importe_total"],
            row["tope_diario"],
            fechas_disponibles,
            row["combustible"]
        )
        
        ws_nueva = wb.copy_worksheet(ws_base)
        ws_nueva.title = str(row["placa"]).replace("-", "").strip()[:31]
        escribir_hoja_bitacora(ws_nueva, row.to_dict(), cargas)
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    st.success("✅ ¡Bitácoras generadas con fórmulas dinámicas en Litros e Importes!")
    st.download_button(
        label="📥 Descargar Libro Excel (Bitácoras Balanceadas)",
        data=buffer,
        file_name="BITACORAS_JUSTIFICADAS_OCTUBRE_2026.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
