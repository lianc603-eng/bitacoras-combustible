import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import io, os, math, datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Bitácoras de Combustible", page_icon="⛽", layout="wide")
st.title("⛽ Generador Automatizado de Bitácoras de Combustible")
st.caption("Dirección de Desarrollo Urbano y Medio Ambiente | H. Ayuntamiento de Campeche")

RUTAS_PLANTILLA_LOCAL = [
    "7.DESARROLLO URBANO Y MED AMBIENTE.xlsx",
    "plantilla.xlsx",
    "1.FORMATO BITACORA 2024 - 2027 OK (1).xlsx",
    "1.FORMATO BITACORA 2024 - 2027.xlsx"
]
plantilla_detectada = next((r for r in RUTAS_PLANTILLA_LOCAL if os.path.exists(r)), None)

# --- LISTA EXACTA DE LAS 13 PLACAS ACTIVAS (VERDES) ---
PLACAS_ACTIVAS = {
    "85GWU7", "86GWU7", "86GWU8", "87GWU8", "DFT565C", "88GWU7", 
    "88GWU8", "89GWU7", "89GWU8", "90GWU8", "91GWU7", "CN2633B", "93GWU7"
}

# --- CATÁLOGO BASE (Formato BITÁCORA) ---
VEHICULOS_DEFAULT = [
    {"direccion": "MEDIO AMBIENTE", "placa": "DKR456B",   "tipo": "JIMNY GLX TA, 5 PUERTAS 2025", "combustible": "PREMIUM", "importe_total": 5000.00, "tope_diario": 800.0,  "modo_carga": "DINERO", "anio": "2025", "lleva_odometro": True,  "activo": True},
]

PRECIOS = {"MAGNA": 23.99, "PREMIUM": 26.95, "DIESEL": 27.00}
NOMBRES_DIAS = {0: "Lunes", 1: "Martes", 3: "Jueves", 4: "Viernes"}

# --- FUNCIONES LECTURA PEMEX ---
def detectar_tipo_plantilla(wb):
    if not wb: return "BITACORA"
    if "BITACORA VEHICULOS" in wb.sheetnames: return "BITACORA"
    ws0 = wb[wb.sheetnames[0]]
    for row in ws0.iter_rows(max_row=10, values_only=True):
        for cell in row:
            if cell and "PEMEX" in str(cell).upper(): return "PEMEX"
    if len(wb.sheetnames) > 3: return "PEMEX"
    return "BITACORA"

def extraer_info_hoja(ws):
    info = {"serie": "", "tipo": "", "anio": "", "placa": ""}
    for r in range(1, 30):
        v = ws.cell(row=r, column=1).value
        if v and "NUMERO DE SERIE" in str(v).upper():
            fr = r + 1
            info["serie"] = ws.cell(row=fr, column=1).value or ""
            info["tipo"]  = ws.cell(row=fr, column=4).value or "MOTOCICLETA"
            info["anio"]  = ws.cell(row=fr, column=7).value or ""
            info["placa"] = ws.cell(row=fr, column=8).value or ""
            break
    return info

def obtener_max_litros(tipo_str):
    tipo_upper = str(tipo_str).upper()
    if "MOTOCICLETA" in tipo_upper: return 8.33
    elif "JETTA" in tipo_upper or "VOLKSWAGEN" in tipo_upper: return 41.0
    elif "RAM" in tipo_upper or "CAMIONETA" in tipo_upper: return 83.36
    return 15.0

def extraer_vehiculos_pemex(wb):
    vehiculos = []
    for sname in wb.sheetnames:
        if sname.upper().endswith(" R"): continue
        
        # Normalizamos el nombre (quitamos guiones y mayúsculas) para buscar en la lista
        sname_clean = sname.replace("-", "").replace(" ", "").upper()
        es_activo = sname_clean in PLACAS_ACTIVAS
        
        info = extraer_info_hoja(wb[sname])
        max_litros = obtener_max_litros(info["tipo"])
        
        vehiculos.append({
            "hoja": sname, "placa": sname, "tipo": info["tipo"],
            "total_litros": 0.0, "max_litros_carga": max_litros, 
            "activo": es_activo, # <--- AQUÍ SE MARCAN LOS 13
            "combustible": "MAGNA", "serie": info["serie"], "anio": str(info["anio"])
        })
        
    # Reordenamos las columnas para que se vean bien en la interfaz
    df = pd.DataFrame(vehiculos)
    if not df.empty:
        df = df[["hoja", "placa", "tipo", "total_litros", "max_litros_carga", "activo", "combustible"]]
    return df

def encontrar_filas_datos_pemex(ws):
    fila_header = None
    for r in range(1, 40):
        v = ws.cell(row=r, column=1).value
        if v and "11 FECHA DE SUMINISTRO" in str(v).upper():
            fila_header = r; break
    if fila_header is None: return None, None
    fila_ini = fila_header + 2
    fila_total = None
    for r in range(fila_ini, fila_ini + 60):
        v = ws.cell(row=r, column=1).value
        if v and "17 TOTAL" in str(v).upper():
            fila_total = r; break
    fila_fin = (fila_total - 1) if fila_total else fila_ini + 30
    return fila_ini, fila_fin

# --- FUNCIONES BALANCEO ---
def balancear_importes(total, max_day):
    n_cargas = math.ceil(total / max_day)
    if n_cargas > 11: n_cargas, max_day = 11, math.ceil(total / 11)
    step = 200 if max_day >= 800 else (100 if max_day >= 500 else 20)
    cents = round(total - math.floor(total), 2)
    int_total = int(round(total - cents))
    for cand_high in range(int(max_day), int(total / n_cargas) - 1, -step):
        cand_low = cand_high - step
        if cand_low <= 0: continue
        diff = int_total - n_cargas * cand_low
        if diff >= 0 and step > 0 and diff % step == 0:
            k = diff // step
            if 0 <= k <= n_cargas:
                highs = [cand_high]*k; lows = [cand_low]*(n_cargas-k)
                alt = []
                while highs or lows:
                    if highs: alt.append(highs.pop(0))
                    if lows:  alt.append(lows.pop(0))
                alt[-1] = round(alt[-1] + cents, 2)
                return alt
    base = int((int_total / n_cargas) // step * step)
    loads = [base]*n_cargas
    loads[-1] = round(loads[-1] + (int_total - sum(loads)) + cents, 2)
    return loads

def balancear_litros(total_litros, max_litros, dias_disponibles):
    if total_litros <= 0: return []
    
    min_cargas = math.ceil(total_litros / max_litros)
    n_cargas = max(min_cargas, min(dias_disponibles, math.ceil(total_litros)))
    
    step = 10 if max_litros >= 40 else (5 if max_litros >= 15 else 1)
    total_int = int(total_litros)
    cents = round(total_litros - total_int, 2)
    
    for cand_high in range(int(max_litros), max(int(total_litros / n_cargas) - 1, step - 1), -step):
        cand_low = cand_high - step
        if cand_low <= 0: continue
        
        diff = total_int - n_cargas * cand_low
        if diff >= 0 and diff % step == 0:
            k = diff // step
            if 0 <= k <= n_cargas:
                highs = [cand_high] * k
                lows = [cand_low] * (n_cargas - k)
                alternated = []
                while highs or lows:
                    if highs: alternated.append(highs.pop(0))
                    if lows: alternated.append(lows.pop(0))
                
                alternated[-1] = round(alternated[-1] + cents, 2)
                return alternated

    base = round(total_litros / n_cargas, 2)
    loads = [base] * (n_cargas - 1)
    loads.append(round(total_litros - sum(loads), 2))
    return loads

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Estado de la Plantilla")
    if plantilla_detectada:
        st.success(f"✅ Plantilla detectada: `{os.path.basename(plantilla_detectada)}`")
    else:
        st.warning("⚠️ Sube una plantilla:")
    uploaded_file = st.file_uploader("Reemplazar plantilla (opcional)", type=["xlsx"])
    
    st.markdown("---")
    st.header("📅 Configuración del Periodo")
    fecha_ini = st.date_input("Inicio", datetime.date(2026, 9, 26))
    fecha_fin = st.date_input("Fin", datetime.date(2026, 10, 25))
    
    dias_autorizados = [0, 1, 3, 4]
    fechas_teoricas = []
    cur = fecha_ini
    while cur <= fecha_fin:
        if cur.weekday() in dias_autorizados: fechas_teoricas.append(cur)
        cur += datetime.timedelta(days=1)
        
    st.markdown("---")
    st.header("🚫 Días Inhábiles")
    opciones_dias = [f"{d.strftime('%Y-%m-%d')} ({NOMBRES_DIAS[d.weekday()]})" for d in fechas_teoricas]
    dias_inhabiles = st.multiselect("Excluir:", options=opciones_dias)
    fechas_excluidas = {op.split(" ")[0] for op in dias_inhabiles}
    fechas_disponibles = [d for d in fechas_teoricas if d.strftime("%Y-%m-%d") not in fechas_excluidas]
    st.info(f"📅 **{len(fechas_disponibles)} fechas activas** para distribuir cargas")

    st.markdown("---")
    st.header("📝 Textos del Formato")
    fecha_entrega_txt = st.text_input("Fecha en Bitácora", "25 DE OCTUBRE DE 2026")
    destino_txt = st.text_input("Destino o Uso", "RECORRIDOS OPERATIVOS")
    responsable_txt = st.text_input("Responsable Operativo", "OPERADOR ASIGNADO")
    
    st.markdown("---")
    st.header("🚗 Odómetro")
    km_inicial_jimny = st.number_input("Odómetro Inicial", value=40128, step=1)
    rendimiento_jimny = st.number_input("Rendimiento (km/l)", value=14, step=1)

if not fechas_disponibles:
    st.error("⚠️ Excluiste todas las fechas. Desmarca al menos una.")
    st.stop()

# --- DETECCIÓN DE PLANTILLA ---
wb_temp = None
if uploaded_file: wb_temp = openpyxl.load_workbook(uploaded_file)
elif plantilla_detectada: wb_temp = openpyxl.load_workbook(plantilla_detectada)

tipo_plantilla = detectar_tipo_plantilla(wb_temp)
st.markdown(f"### 🧾 Plantilla detectada: **{tipo_plantilla}**")
st.subheader("1. Cuotas Asignadas y Parámetros")

# --- INTERFAZ DE EDICIÓN ---
if tipo_plantilla == "PEMEX":
    st.caption("✅ Los 13 vehículos operativos (verdes en tu Excel) ya vienen seleccionados. Escribe los litros totales y el sistema los repartirá equitativamente entre los días del mes.")
    df_base = extraer_vehiculos_pemex(wb_temp)
    
    df_edit = st.data_editor(
        df_base,
        column_config={
            "hoja": st.column_config.TextColumn("Hoja Excel", disabled=True),
            "placa": st.column_config.TextColumn("Placa", disabled=True),
            "tipo": st.column_config.TextColumn("Marca y Tipo", disabled=True),
            "total_litros": st.column_config.NumberColumn("Litros Totales (Mes)", format="%.2f L", min_value=0.0),
            "max_litros_carga": st.column_config.NumberColumn("Tope Máx (L)", format="%.2f L", min_value=1.0),
            "activo": st.column_config.CheckboxColumn("¿Generar?", default=True),
            "combustible": st.column_config.SelectboxColumn("Combustible", options=["MAGNA","PREMIUM","DIESEL"]),
        },
        use_container_width=True, hide_index=True
    )
    c1, c2 = st.columns(2)
    c1.metric("Vehículos Activos", int(df_edit["activo"].sum()))
    c2.metric("Total Litros a Repartir", f"{df_edit[df_edit['activo']]['total_litros'].sum():,.2f} L")

else: # BITACORA
    st.caption("Marca **Generar** para incluir. Si el modo es `LITROS`, el 'Tope Día' se interpreta como la capacidad máxima por carga.")
    df_edit = st.data_editor(
        pd.DataFrame(VEHICULOS_DEFAULT),
        column_config={
            "activo": st.column_config.CheckboxColumn("Generar", default=True),
            "modo_carga": st.column_config.SelectboxColumn("Modo", options=["DINERO","LITROS"]),
            "lleva_odometro": st.column_config.CheckboxColumn("Odómetro"),
            "importe_total": st.column_config.NumberColumn("Presupuesto ($)", format="$%.2f"),
            "tope_diario": st.column_config.NumberColumn("Tope / Max Litros", format="%.2f"),
            "combustible": st.column_config.SelectboxColumn("Combustible", options=["MAGNA","PREMIUM","DIESEL"]),
        },
        use_container_width=True, hide_index=True
    )
    tot_m = df_edit[df_edit["activo"] & (df_edit["combustible"]=="MAGNA")]["importe_total"].sum()
    tot_p = df_edit[df_edit["activo"] & (df_edit["combustible"]=="PREMIUM")]["importe_total"].sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("Magna", f"${tot_m:,.2f}")
    c2.metric("Premium", f"${tot_p:,.2f}")
    c3.metric("Total", f"${tot_m+tot_p:,.2f}")

# --- VISTA PREVIA ---
st.markdown("---")
st.subheader("2. Vista Previa")

if tipo_plantilla == "PEMEX":
    preview = []
    for _, r in df_edit.iterrows():
        if not r.get("activo") or r["total_litros"] <= 0: continue
        
        cargas = balancear_litros(r["total_litros"], r["max_litros_carga"], len(fechas_disponibles))
        
        preview.append({
            "Placa": r["placa"], "Tipo": r["tipo"],
            "Cargas": len(cargas), "Total L": f"{sum(cargas):.2f}",
            "Desglose de Litros": " + ".join(f"{c:.2f}L" for c in cargas),
        })
    if preview: st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

else:
    preview = []
    for _, r in df_edit.iterrows():
        if not r.get("activo"): continue
        modo = r.get("modo_carga") or "DINERO"
        precio = PRECIOS.get(r["combustible"], 23.99)
        if modo == "LITROS":
            litros = balancear_litros(r["importe_total"]/precio, r["tope_diario"], len(fechas_disponibles))
            importes = [round(l*precio,2) for l in litros]
        else:
            importes = balancear_importes(r["importe_total"], r["tope_diario"])
            litros = [round(i/precio,2) for i in importes]
        preview.append({
            "Vehículo": r["tipo"], "Placa": r["placa"], "Modo": modo,
            "Cargas": len(importes),
            "Desglose": " + ".join(f"{l:.2f}L=${i:,.0f}" for l,i in zip(litros,importes)),
        })
    if preview: st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

# --- ESCRITURA ---
def escribir_bitacora(ws, datos, cargas):
    centro, derecha = Alignment(horizontal="center", vertical="center"), Alignment(horizontal="right", vertical="center")
    ws["B9"] = f"N° ECONOMICO: {datos.get('num_eco','')}"; ws["D9"] = f"MARCA Y TIPO:  {datos.get('tipo','')}"
    ws["H9"] = f"AÑO: {datos.get('anio','')}"; ws["K9"] = f"PLACAS: {datos.get('placa','')}"
    for idx, c in enumerate(cargas):
        r = 14 + idx
        if r > 24: break
        ws[f"B{r}"] = c["fecha"]
        ws[f"D{r}"] = c["precio_unit"]; ws[f"D{r}"].number_format = '0.00'
        ws[f"E{r}"] = c["importe"]; ws[f"E{r}"].number_format = '#,##0.00'
        ws[f"C{r}"] = f"=E{r}/D{r}"; ws[f"C{r}"].number_format = '0.00'; ws[f"C{r}"].alignment = derecha
        if not datos.get("lleva_odometro"):
            for col in ["F","G","H"]: ws[f"{col}{r}"] = "---"; ws[f"{col}{r}"].alignment = centro
        else:
            ws[f"H{r}"] = f"{rendimiento_jimny} KM/L"; ws[f"H{r}"].alignment = centro
            ws[f"F{r}"] = km_inicial_jimny if idx==0 else f"=G{r-1}"
            ws[f"F{r}"].number_format='#,##0'; ws[f"F{r}"].alignment=centro
            ws[f"G{r}"] = f"=F{r}+ROUND(C{r}*{rendimiento_jimny},0)"
            ws[f"G{r}"].number_format='#,##0'; ws[f"G{r}"].alignment=centro
        ws[f"I{r}"] = responsable_txt
    ws["D25"]="TOTAL"; ws["E25"]="=SUM(E14:E24)"; ws["E25"].number_format='$#,##0.00'
    ws["B27"]=f"BITÁCORA CORRESPONDIENTE A LA ENTREGA DEL DÍA: {fecha_entrega_txt}"
    ws["B28"]=f"DESTINO O USO DEL VEHICULO: {destino_txt}"

def escribir_pemex(ws, datos, fechas, litros):
    fila_ini, fila_fin = encontrar_filas_datos_pemex(ws)
    if not fila_ini: return False
    n = min(len(fechas), len(litros), fila_fin - fila_ini + 1)
    km = km_inicial_jimny if datos.get("lleva_odometro") else None
    
    for i in range(n):
        r = fila_ini + i
        ws.cell(row=r, column=1).value = fechas[i].strftime("%Y-%m-%d")
        ws.cell(row=r, column=5).value = round(float(litros[i]), 2)
        ws.cell(row=r, column=5).number_format = '0.00'
        ws.cell(row=r, column=7).value = "N/F" if km is None else km
        ws.cell(row=r, column=4).value = datos.get("combustible", "MAGNA")
        
        if km is not None: km += round(litros[i] * rendimiento_jimny)
        
    for r in range(fila_ini + n, fila_fin + 1):
        ws.cell(row=r, column=1).value = None
        ws.cell(row=r, column=5).value = None
        ws.cell(row=r, column=7).value = None
    return True

# --- GENERACIÓN ---
st.markdown("---")
st.subheader("3. Descargar Formatos")
if st.button("🚀 Generar Libro Excel"):
    if wb_temp is None:
        st.error("Sube una plantilla primero."); st.stop()
    
    wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.getvalue())) if uploaded_file else openpyxl.load_workbook(plantilla_detectada)
    generados = 0

    if tipo_plantilla == "PEMEX":
        for _, row in df_edit.iterrows():
            if not row.get("activo") or row["total_litros"] <= 0: continue
            ws = wb[row["hoja"]]
            
            cargas = balancear_litros(row["total_litros"], row["max_litros_carga"], len(fechas_disponibles))
            if escribir_pemex(ws, row.to_dict(), fechas_disponibles, cargas):
                generados += 1
    else:
        ws_base = wb["BITACORA VEHICULOS"]
        for _, row in df_edit.iterrows():
            if not row.get("activo"): continue
            modo = row.get("modo_carga") or "DINERO"
            precio = PRECIOS.get(row["combustible"], 23.99)
            if modo == "LITROS":
                litros = balancear_litros(row["importe_total"]/precio, row["tope_diario"], len(fechas_disponibles))
                importes = [round(l*precio,2) for l in litros]
            else:
                importes = balancear_importes(row["importe_total"], row["tope_diario"])
                litros = [round(i/precio,2) for i in importes]
            step_f = len(fechas_disponibles)/len(importes)
            fechas = [fechas_disponibles[int(round(i*step_f))] for i in range(len(importes))]
            cargas = [{"fecha": f.strftime("%d/%m/%Y"), "precio_unit": precio, "importe": imp} for f, imp in zip(fechas, importes)]
            
            ws_nueva = wb.copy_worksheet(ws_base)
            ws_nueva.title = str(row["placa"]).replace("-","").strip()[:31]
            escribir_bitacora(ws_nueva, row.to_dict(), cargas)
            generados += 1

    if generados > 0:
        buffer = io.BytesIO(); wb.save(buffer); buffer.seek(0)
        st.success(f"✅ {generados} hoja(s) generadas con distribución lógica.")
        st.download_button("📥 Descargar Excel", data=buffer, file_name=f"BITACORAS_{tipo_plantilla}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("No hay vehículos seleccionados o no pusiste litros totales.")
