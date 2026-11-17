import streamlit as st
import pandas as pd
import openpyxl
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

PLACAS_ACTIVAS = [
    "85GWU7", "86GWU7", "86GWU8", "87GWU8", "DFT-565-C", 
    "88GWU7", "88GWU8", "89GWU7", "89GWU8", "90GWU8", 
    "91GWU7", "CN-2633-B", "93GWU7"
]

MAPEO_PLACAS = {
    "85GWU7": "85GWU7", "86GWU7": "86GWU7", "43GMP6": "86GWU8", "87GWU8": "87GWU8",
    "DFT-565-C": "DFT-565-C", "88GWU7": "88GWU7", "88GWU8": "88GWU8", "49GMP6": "89GWU7",
    "50GMP6": "89GWU8", "89GWU8": "89GWU8", "38GMP6": "90GWU8", "30GMP6": "91GWU7",
    "91GWU7": "91GWU7", "CN-2633-B": "CN-2633-B", "32GMP6": "93GWU7", "93GWU7": "93GWU7",
    "30GMY2": "30GMY2"
}

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
        placa_actual = MAPEO_PLACAS.get(sname, sname)
        es_activo = placa_actual in PLACAS_ACTIVAS
        info = extraer_info_hoja(wb[sname])
        max_litros = obtener_max_litros(info["tipo"])
        
        vehiculos.append({
            "hoja": sname, "placa": placa_actual, "tipo": info["tipo"],
            "total_litros": 0.0, "max_litros_carga": max_litros, 
            "activo": es_activo, "combustible": "MAGNA"
        })
        
    df = pd.DataFrame(vehiculos)
    if not df.empty:
        df = df[["placa", "tipo", "total_litros", "max_litros_carga", "activo", "combustible", "hoja"]]
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

# --- FUNCIONES DE REPARTO Y BALANCEO ---
def repartir_bolsa_global(df, bolsa_total, tope_mensual_camioneta=260.0, tope_mensual_auto=200.0):
    """
    Reparte la bolsa total de forma equilibrada usando pesos suavizados (raíz cuadrada)
    y respetando un tope máximo mensual para que ningún vehículo acapare demasiado.
    """
    df_res = df.copy()
    activos_mask = df_res["activo"] == True
    if not activos_mask.any() or bolsa_total <= 0:
        df_res["total_litros"] = 0.0
        return df_res

    # 1. Definir el tope mensual máximo permitido por vehículo según su tipo
    topes_mensuales = pd.Series(index=df_res[activos_mask].index, dtype=float)
    for idx, row in df_res[activos_mask].iterrows():
        tipo_u = str(row["tipo"]).upper()
        if "RAM" in tipo_u or "CAMIONETA" in tipo_u:
            topes_mensuales[idx] = tope_mensual_camioneta
        elif "JETTA" in tipo_u or "VOLKSWAGEN" in tipo_u or "AUTOMOV" in tipo_u:
            topes_mensuales[idx] = tope_mensual_auto
        else:
            topes_mensuales[idx] = 9999.0 

    # 2. Pesos suavizados (raíz cuadrada de la capacidad del tanque)
    pesos = df_res.loc[activos_mask, "max_litros_carga"].apply(math.sqrt)

    asignaciones = pd.Series(0.0, index=df_res[activos_mask].index)
    remanente = float(bolsa_total)
    disponibles = asignaciones.index.tolist()

    # 3. Reparto iterativo respetando topes mensuales
    while remanente > 0.01 and disponibles:
        suma_pesos = pesos.loc[disponibles].sum()
        if suma_pesos <= 0:
            break

        nuevos_disponibles = []
        repartido_en_ronda = 0.0

        for idx in disponibles:
            cuota = (pesos[idx] / suma_pesos) * remanente
            espacio = topes_mensuales[idx] - asignaciones[idx]

            if cuota >= espacio:
                asignaciones[idx] = topes_mensuales[idx]
                repartido_en_ronda += espacio
            else:
                asignaciones[idx] += cuota
                repartido_en_ronda += cuota
                nuevos_disponibles.append(idx)

        remanente = round(remanente - repartido_en_ronda, 4)
        disponibles = nuevos_disponibles

    # Si sobró algo porque todos llegaron a su tope, se reparte parejo entre todos los activos
    if remanente > 0.01:
        extra = remanente / len(asignaciones)
        asignaciones += extra

    asignaciones = asignaciones.round(2)

    # Ajuste final de centavos para cuadrar exactamente con la bolsa_total
    diferencia = round(bolsa_total - asignaciones.sum(), 2)
    motos_idx = [i for i in asignaciones.index if "MOTOCICLETA" in str(df_res.loc[i, "tipo"]).upper()]
    idx_ajuste = motos_idx[-1] if motos_idx else asignaciones.index[-1]
    asignaciones.loc[idx_ajuste] = round(asignaciones.loc[idx_ajuste] + diferencia, 2)

    df_res["total_litros"] = 0.0
    df_res.loc[activos_mask, "total_litros"] = asignaciones
    return df_res

def balancear_litros(total_litros, max_litros):
    """Divide el total de litros en cargas alternadas (ej. 8L, 7L, 8L...) que no superen el max_litros."""
    if total_litros <= 0: return []
    n_cargas = math.ceil(total_litros / max_litros)
    
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

def asignar_fechas_alternadas(fechas_disponibles, n_cargas):
    """
    Esparce las cargas evitando días seguidos.
    Prioriza días alternados (ej. Lunes y Jueves, o Martes y Viernes) antes de usar días consecutivos.
    """
    if n_cargas <= 0 or not fechas_disponibles:
        return []
    if n_cargas >= len(fechas_disponibles):
        return fechas_disponibles[:n_cargas]

    pares = fechas_disponibles[0::2]
    impares = fechas_disponibles[1::2]

    if n_cargas <= len(pares):
        step = len(pares) / n_cargas
        return sorted([pares[min(int(round(i * step)), len(pares) - 1)] for i in range(n_cargas)])
    else:
        faltantes = n_cargas - len(pares)
        step = len(impares) / faltantes
        extra = [impares[min(int(round(i * step)), len(impares) - 1)] for i in range(faltantes)]
        return sorted(pares + extra)

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
    st.info(f"📅 **{len(fechas_disponibles)} fechas activas**")

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
    if "df_vehiculos" not in st.session_state or st.session_state.get("plantilla_id") != (uploaded_file.name if uploaded_file else plantilla_detectada):
        st.session_state.df_vehiculos = extraer_vehiculos_pemex(wb_temp)
        st.session_state.plantilla_id = uploaded_file.name if uploaded_file else plantilla_detectada

    # --- REPARTO AUTOMÁTICO CON TOPE PARA CAMIONETA ---
    with st.container(border=True):
        st.markdown("#### 🎯 Reparto Automático del Total a Distribuir")
        col_tot1, col_tot2, col_tot3, col_tot4 = st.columns([2, 1.5, 1.5, 1.5])
        
        bolsa_litros = col_tot1.number_input(
            "Total de Litros a Repartir:",
            min_value=0.0,
            value=1256.0,
            step=50.0,
            format="%.2f"
        )
        tope_cam = col_tot2.number_input(
            "Tope Máx Camioneta (Mes):",
            min_value=50.0,
            value=260.0,
            step=20.0,
            format="%.2f"
        )
        tope_auto = col_tot3.number_input(
            "Tope Máx Auto/Jetta (Mes):",
            min_value=50.0,
            value=200.0,
            step=20.0,
            format="%.2f"
        )
        col_tot4.write("")
        col_tot4.write("")
        if col_tot4.button("⚡ Distribuir", use_container_width=True):
            st.session_state.df_vehiculos = repartir_bolsa_global(
                st.session_state.df_vehiculos, 
                bolsa_litros,
                tope_mensual_camioneta=tope_cam,
                tope_mensual_auto=tope_auto
            )
            st.rerun()

    st.caption("También puedes ajustar manualmente cualquier celda en la tabla de abajo o marcar/desmarcar vehículos:")

    # TABLA EDITABLE
    df_edit = st.data_editor(
        st.session_state.df_vehiculos,
        column_config={
            "placa": st.column_config.TextColumn("Placa Actual", disabled=True),
            "tipo": st.column_config.TextColumn("Marca y Tipo", disabled=True),
            "total_litros": st.column_config.NumberColumn(
                "Litros Totales (Mes)", 
                help="Haz clic y escribe los litros asignados.",
                format="%.2f L", 
                min_value=0.0,
                step=10.0
            ),
            "max_litros_carga": st.column_config.NumberColumn("Tope Máx (L)", format="%.2f L", disabled=True),
            "activo": st.column_config.CheckboxColumn("¿Generar?", default=True),
            "combustible": st.column_config.SelectboxColumn("Combustible", options=["MAGNA","PREMIUM","DIESEL"]),
            "hoja": st.column_config.TextColumn("Hoja Excel", disabled=True),
        },
        use_container_width=True, hide_index=True, key="editor_pemex"
    )
    st.session_state.df_vehiculos = df_edit

    c1, c2 = st.columns(2)
    c1.metric("Vehículos Activos", int(df_edit["activo"].sum()))
    c2.metric("Total Litros a Repartir", f"{df_edit[df_edit['activo']]['total_litros'].sum():,.2f} L")

# --- VISTA PREVIA ---
st.markdown("---")
st.subheader("2. Vista Previa")

if tipo_plantilla == "PEMEX":
    preview = []
    for _, r in df_edit.iterrows():
        if not r.get("activo") or r["total_litros"] <= 0: continue
        
        cargas = balancear_litros(r["total_litros"], r["max_litros_carga"])
        fechas_asignadas = asignar_fechas_alternadas(fechas_disponibles, len(cargas))
        
        desglose_str = " | ".join([f"{f.strftime('%d/%b')} ({NOMBRES_DIAS[f.weekday()][:3]}): {c:.2f}L" for f, c in zip(fechas_asignadas, cargas)])
        
        preview.append({
            "Placa": r["placa"], "Tipo": r["tipo"],
            "Cargas": len(cargas), "Total L": f"{sum(cargas):.2f}",
            "Fechas y Litros Asignados (Alternados)": desglose_str,
        })
    if preview: st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)
    else: st.info("Escribe el total de litros arriba y presiona 'Distribuir', o edita la tabla.")

# --- ESCRITURA ---
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
            
            cargas = balancear_litros(row["total_litros"], row["max_litros_carga"])
            fechas_asignadas = asignar_fechas_alternadas(fechas_disponibles, len(cargas))
            
            if escribir_pemex(ws, row.to_dict(), fechas_asignadas, cargas):
                generados += 1

    if generados > 0:
        buffer = io.BytesIO(); wb.save(buffer); buffer.seek(0)
        st.success(f"✅ {generados} hoja(s) generadas. Las cargas han sido esparcidas a lo largo del mes.")
        st.download_button("📥 Descargar Excel", data=buffer, file_name=f"BITACORAS_{tipo_plantilla}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("No hay vehículos seleccionados o no pusiste litros totales.")
