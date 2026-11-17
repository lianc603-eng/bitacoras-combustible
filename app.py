import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import io, os, math, datetime

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador de Bitácoras de Combustible", page_icon="⛽", layout="wide")
st.title("⛽ Generador Automatizado de Bitácoras de Combustible")
st.caption("Dirección de Desarrollo Urbano y Medio Ambiente | H. Ayuntamiento de Campeche")

RUTAS_PLANTILLA_LOCAL = [
    "7.DESARROLLO URBANO Y MED AMBIENTE.xlsx",
    "plantilla.xlsx",
    "1.FORMATO BITACORA 2024 - 2027 OK (1).xlsx",
    "1.FORMATO BITACORA 2024 - 2027.xlsx",
]
plantilla_detectada = next((r for r in RUTAS_PLANTILLA_LOCAL if os.path.exists(r)), None)

# Catálogo para el formato BITACORA VEHICULOS (no se usa en PEMEX)
VEHICULOS_DEFAULT = [
    {"direccion":"MEDIO AMBIENTE","placa":"DKR456B","tipo":"JIMNY GLX TA, 5 PUERTAS 2025","combustible":"PREMIUM","importe_total":5000.00,"tope_diario":800.0,"modo_carga":"DINERO","anio":"2025","lleva_odometro":True, "activo":True},
    {"direccion":"MEDIO AMBIENTE","placa":"DFT-565-C","tipo":"JETTA","combustible":"MAGNA","importe_total":5300.00,"tope_diario":1000.0,"modo_carga":"DINERO","anio":"","lleva_odometro":False,"activo":True},
    {"direccion":"MEDIO AMBIENTE","placa":"CN-2633-B","tipo":"CAMIONETA (RAM)","combustible":"MAGNA","importe_total":7049.27,"tope_diario":2000.0,"modo_carga":"DINERO","anio":"","lleva_odometro":False,"activo":True},
]

PRECIOS = {"MAGNA": 23.99, "PREMIUM": 26.95, "DIESEL": 27.00}
NOMBRES_DIAS = {0: "Lunes", 1: "Martes", 3: "Jueves", 4: "Viernes"}

# ------------------------------------------------------------------
# DETECCIÓN Y EXTRACCIÓN DE PLANTILLAS
# ------------------------------------------------------------------
def detectar_tipo_plantilla(wb):
    """Devuelve 'BITACORA' o 'PEMEX'."""
    if "BITACORA VEHICULOS" in wb.sheetnames:
        return "BITACORA"
    ws0 = wb[wb.sheetnames[0]]
    for row in ws0.iter_rows(max_row=10, values_only=True):
        for cell in row:
            if cell and "PEMEX" in str(cell).upper():
                return "PEMEX"
    if len(wb.sheetnames) > 3:
        return "PEMEX"
    return "BITACORA"


def extraer_info_hoja(ws):
    """Extrae serie, tipo, año, placa de una hoja PEMEX."""
    info = {"serie": "", "tipo": "", "anio": "", "placa": ""}
    fila_serie = None
    for r in range(1, 30):
        v = ws.cell(row=r, column=1).value
        if v and "NUMERO DE SERIE" in str(v).upper():
            fila_serie = r
            break
    if fila_serie is None:
        return info
    fr = fila_serie + 1
    info["serie"] = ws.cell(row=fr, column=1).value or ""
    info["tipo"]  = ws.cell(row=fr, column=4).value or "MOTOCICLETA"
    info["anio"]  = ws.cell(row=fr, column=7).value or ""
    info["placa"] = ws.cell(row=fr, column=8).value or ""
    return info


def extraer_vehiculos_pemex(wb):
    """Genera la lista de vehículos a partir de las hojas del archivo."""
    vehiculos = []
    for sname in wb.sheetnames:
        if sname.upper().endswith(" R"):
            continue  # hojas "R" son reverso, no llevan datos
        info = extraer_info_hoja(wb[sname])
        vehiculos.append({
            "hoja": sname,
            "placa": sname,
            "serie": info["serie"],
            "tipo": info["tipo"],
            "anio": str(info["anio"]) if info["anio"] else "",
            "combustible": "MAGNA",
            "total_litros": 800.0,
            "max_litros_carga": 15.0,
            "lleva_odometro": False,
            "activo": True,
        })
    return vehiculos


def encontrar_filas_datos_pemex(ws):
    """Localiza fila inicial y final de captura de litros."""
    fila_header = None
    for r in range(1, 40):
        v = ws.cell(row=r, column=1).value
        if v and "11 FECHA DE SUMINISTRO" in str(v).upper():
            fila_header = r
            break
    if fila_header is None:
        return None, None
    fila_ini = fila_header + 2
    fila_total = None
    for r in range(fila_ini, fila_ini + 60):
        v = ws.cell(row=r, column=1).value
        if v and "17 TOTAL" in str(v).upper():
            fila_total = r
            break
    fila_fin = (fila_total - 1) if fila_total else fila_ini + 25
    return fila_ini, fila_fin

# ------------------------------------------------------------------
# BALANCEO
# ------------------------------------------------------------------
def balancear_importes(total, max_day):
    n = math.ceil(total / max_day)
    if n > 11: n, max_day = 11, math.ceil(total / 11)
    step = 200 if max_day >= 800 else (100 if max_day >= 500 else 20)
    cents = round(total - math.floor(total), 2)
    int_total = int(round(total - cents))
    for cand_high in range(int(max_day), int(total / n) - 1, -step):
        cand_low = cand_high - step
        if cand_low <= 0: continue
        diff = int_total - n * cand_low
        if diff >= 0 and step > 0 and diff % step == 0:
            k = diff // step
            if 0 <= k <= n:
                highs = [cand_high]*k; lows = [cand_low]*(n-k)
                alt = []
                while highs or lows:
                    if highs: alt.append(highs.pop(0))
                    if lows:  alt.append(lows.pop(0))
                alt[-1] = round(alt[-1] + cents, 2)
                return alt
    base = int((int_total / n) // step * step)
    loads = [base]*n
    rem = int_total - sum(loads)
    loads[-1] = round(loads[-1] + rem + cents, 2)
    return loads


def balancear_litros(total, max_load):
    """Reparte total litros en cargas ≤ max_load."""
    n = max(1, math.ceil(total / max_load))
    step = 10 if max_load >= 40 else (5 if max_load >= 15 else 1)
    for cand_high in range(int(max_load), max(int(total/n)-1, step-1), -step):
        cand_low = cand_high - step
        if cand_low <= 0: continue
        diff = total - n * cand_low
        if diff >= 0 and abs(diff/step - round(diff/step)) < 0.001:
            k = round(diff / step)
            if 0 <= k <= n:
                highs = [cand_high]*k; lows = [cand_low]*(n-k)
                alt = []
                while highs or lows:
                    if highs: alt.append(highs.pop(0))
                    if lows:  alt.append(lows.pop(0))
                rem = round(total - sum(alt), 2)
                alt[-1] = round(alt[-1] + rem, 2)
                return alt
    base = round(total / n, 2)
    loads = [base]*(n-1)
    loads.append(round(total - sum(loads), 2))
    return loads

# ------------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Estado de la Plantilla")
    if plantilla_detectada:
        st.success(f"✅ Plantilla: `{os.path.basename(plantilla_detectada)}`")
    else:
        st.warning("⚠️ Sin plantilla. Sube una:")
    uploaded_file = st.file_uploader("Reemplazar plantilla (opcional)", type=["xlsx"])

    st.markdown("---")
    st.header("📅 Configuración del Periodo")
    fecha_ini = st.date_input("Inicio", datetime.date(2026, 9, 26))
    fecha_fin = st.date_input("Fin",   datetime.date(2026, 10, 25))

    dias_autorizados = [0, 1, 3, 4]
    fechas_teoricas = []
    cur = fecha_ini
    while cur <= fecha_fin:
        if cur.weekday() in dias_autorizados:
            fechas_teoricas.append(cur)
        cur += datetime.timedelta(days=1)

    st.markdown("---")
    st.header("🚫 Días Inhábiles")
    opciones_dias = [f"{d.strftime('%Y-%m-%d')} ({NOMBRES_DIAS[d.weekday()]})" for d in fechas_teoricas]
    excluidos = st.multiselect("Excluir del calendario:", options=opciones_dias)
    fechas_excluidas = {op.split(" ")[0] for op in excluidos}
    fechas_disponibles = [d for d in fechas_teoricas if d.strftime("%Y-%m-%d") not in fechas_excluidas]
    st.info(f"📅 **{len(fechas_disponibles)} fechas activas**")

    st.markdown("---")
    st.header("🚗 Odómetro (para quienes lo llevan)")
    km_inicial_jimny = st.number_input("Odómetro inicial", value=40128, step=1)
    rendimiento_jimny = st.number_input("Rendimiento (km/l)", value=14, step=1)

if not fechas_disponibles:
    st.error("⚠️ Excluiste todas las fechas. Desmarca al menos una en la barra lateral.")
    st.stop()

# ------------------------------------------------------------------
# CARGAR PLANTILLA Y DETECTAR TIPO
# ------------------------------------------------------------------
wb_temp = None
if uploaded_file is not None:
    wb_temp = openpyxl.load_workbook(uploaded_file)
elif plantilla_detectada:
    wb_temp = openpyxl.load_workbook(plantilla_detectada)

tipo_plantilla = detectar_tipo_plantilla(wb_temp) if wb_temp else "BITACORA"
st.markdown(f"### 🧾 Plantilla detectada: **{tipo_plantilla}**")

# ------------------------------------------------------------------
# TABLA EDITABLE (según el tipo)
# ------------------------------------------------------------------
st.subheader("1. Cuotas Asignadas y Parámetros")

if tipo_plantilla == "PEMEX":
    st.caption(
        "En este formato **solo se capturan litros** (col. E de cada hoja). "
        "Cada hoja ya corresponde a un vehículo. "
        "Marca **Generar** para incluirla en el archivo final."
    )
    df_base = pd.DataFrame(extraer_vehiculos_pemex(wb_temp))
    df_edit = st.data_editor(
        df_base,
        column_config={
            "hoja": st.column_config.TextColumn("Hoja", disabled=True),
            "activo": st.column_config.CheckboxColumn("Generar", default=True),
            "placa": st.column_config.TextColumn("Placa"),
            "tipo": st.column_config.TextColumn("Marca y Tipo"),
            "total_litros": st.column_config.NumberColumn("Total Litros", format="%.2f L"),
            "max_litros_carga": st.column_config.NumberColumn("Máx. Litros/Carga", format="%.2f L"),
            "lleva_odometro": st.column_config.CheckboxColumn("Registra Odómetro"),
            "combustible": st.column_config.SelectboxColumn("Combustible", options=["MAGNA","PREMIUM","DIESEL"]),
        },
        use_container_width=True,
        num_rows="dynamic",
    )
    col1, col2 = st.columns(2)
    col1.metric("Vehículos activos", int(df_edit["activo"].sum()))
    col2.metric("Total litros a repartir", f"{df_edit[df_edit['activo']]['total_litros'].sum():,.2f} L")

else:  # BITACORA
    st.caption(
        "Marca **Generar** para incluir el vehículo. "
        "En **Modo** elige `DINERO` (divide $) o `LITROS` (divide litros)."
    )
    df_edit = st.data_editor(
        pd.DataFrame(VEHICULOS_DEFAULT),
        column_config={
            "activo": st.column_config.CheckboxColumn("Generar", default=True),
            "modo_carga": st.column_config.SelectboxColumn("Modo", options=["DINERO","LITROS"]),
            "lleva_odometro": st.column_config.CheckboxColumn("Registra Odómetro"),
            "importe_total": st.column_config.NumberColumn("Presupuesto ($)", format="$%.2f"),
            "tope_diario": st.column_config.NumberColumn("Tope Día", format="%.2f"),
            "combustible": st.column_config.SelectboxColumn("Combustible", options=["MAGNA","PREMIUM","DIESEL"]),
        },
        use_container_width=True, num_rows="dynamic",
    )
    tot_m = df_edit[df_edit["activo"] & (df_edit["combustible"]=="MAGNA")]["importe_total"].sum()
    tot_p = df_edit[df_edit["activo"] & (df_edit["combustible"]=="PREMIUM")]["importe_total"].sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Magna", f"${tot_m:,.2f}")
    c2.metric("Total Premium", f"${tot_p:,.2f}")
    c3.metric("Total General", f"${tot_m+tot_p:,.2f}")

# ------------------------------------------------------------------
# VISTA PREVIA
# ------------------------------------------------------------------
st.markdown("---")
st.subheader("2. Vista Previa")

if tipo_plantilla == "PEMEX":
    preview = []
    for _, r in df_edit.iterrows():
        if not r.get("activo"): continue
        fila_ini, fila_fin = encontrar_filas_datos_pemex(wb_temp[r["hoja"]])
        if fila_ini is None:
            preview.append({"Hoja": r["hoja"], "Placa": r["placa"], "Estado": "❌ Estructura no reconocida"})
            continue
        n_filas = fila_fin - fila_ini + 1
        max_real = max(r["max_litros_carga"], r["total_litros"] / n_filas)
        cargas = balancear_litros(r["total_litros"], max_real)
        cargas = cargas[:n_filas]
        preview.append({
            "Hoja": r["hoja"],
            "Placa": r["placa"],
            "Filas disponibles": n_filas,
            "Cargas": len(cargas),
            "Total L": f"{sum(cargas):.2f}",
            "Desglose": " + ".join(f"{c:.1f}L" for c in cargas),
        })
    if preview:
        st.dataframe(pd.DataFrame(preview), use_container_width=True)

else:  # BITACORA
    preview = []
    for _, r in df_edit.iterrows():
        if not r.get("activo"): continue
        modo = r.get("modo_carga") or "DINERO"
        precio = PRECIOS.get(r["combustible"], 23.99)
        if modo == "LITROS":
            litros = balancear_litros(r["importe_total"]/precio, r["tope_diario"])
            importes = [round(l*precio,2) for l in litros]
        else:
            importes = balancear_importes(r["importe_total"], r["tope_diario"])
            litros = [round(i/precio,2) for i in importes]
        preview.append({
            "Vehículo": r["tipo"], "Placa": r["placa"], "Modo": modo,
            "Cargas": len(importes),
            "Desglose": " + ".join(f"{l:.1f}L=${i:,.0f}" for l,i in zip(litros,importes)),
        })
    if preview:
        st.dataframe(pd.DataFrame(preview), use_container_width=True)

# ------------------------------------------------------------------
# ESCRITURA (dos helpers)
# ------------------------------------------------------------------
def escribir_hoja_bitacora(ws, datos, cargas):
    centro  = Alignment(horizontal="center", vertical="center")
    derecha = Alignment(horizontal="right",  vertical="center")
    ws["B9"] = f"N° ECONOMICO: {datos.get('num_eco','')}"
    ws["D9"] = f"MARCA Y TIPO:  {datos.get('tipo','')}"
    ws["H9"] = f"AÑO: {datos.get('anio','')}"
    ws["K9"] = f"PLACAS: {datos.get('placa','')}"
    lleva = datos.get("lleva_odometro", False)
    for idx, c in enumerate(cargas):
        r = 14 + idx
        if r > 24: break
        ws[f"B{r}"] = c["fecha"]
        ws[f"D{r}"] = c["precio_unit"]; ws[f"D{r}"].number_format = '0.00'
        ws[f"E{r}"] = c["importe"];     ws[f"E{r}"].number_format = '#,##0.00'
        ws[f"C{r}"] = f"=E{r}/D{r}";    ws[f"C{r}"].number_format = '0.00'; ws[f"C{r}"].alignment = derecha
        if not lleva:
            for col in ["F","G","H"]:
                ws[f"{col}{r}"] = "---"; ws[f"{col}{r}"].alignment = centro
        else:
            ws[f"H{r}"] = f"{rendimiento_jimny} KM/L"; ws[f"H{r}"].alignment = centro
            ws[f"F{r}"] = km_inicial_jimny if idx==0 else f"=G{r-1}"
            ws[f"F{r}"].number_format='#,##0'; ws[f"F{r}"].alignment=centro
            ws[f"G{r}"] = f"=F{r}+ROUND(C{r}*{rendimiento_jimny},0)"
            ws[f"G{r}"].number_format='#,##0'; ws[f"G{r}"].alignment=centro
        ws[f"I{r}"] = datos.get("responsable", "OPERADOR ASIGNADO")
    ws["D25"]="TOTAL"; ws["E25"]="=SUM(E14:E24)"; ws["E25"].number_format='$#,##0.00'
    ws["B27"]=f"BITÁCORA CORRESPONDIENTE A LA ENTREGA DEL DÍA: {datos.get('fecha_txt','')}"
    ws["B28"]=f"DESTINO O USO DEL VEHICULO: {datos.get('destino_txt','')}"


def escribir_hoja_pemex(ws, datos, fechas, litros):
    """Llena una hoja del formato PEMEX: fechas en A, litros en E, odómetro en G."""
    fila_ini, fila_fin = encontrar_filas_datos_pemex(ws)
    if fila_ini is None:
        return False
    n_filas = fila_fin - fila_ini + 1
    n = min(len(fechas), len(litros), n_filas)

    km = km_inicial_jimny if datos.get("lleva_odometro") else None
    for i in range(n):
        r = fila_ini + i
        # Fecha (A)
        f = fechas[i]
        ws.cell(row=r, column=1).value = f.strftime("%Y-%m-%d") if hasattr(f, "strftime") else f
        # Litros (E)
        celda = ws.cell(row=r, column=5)
        celda.value = round(float(litros[i]), 2)
        celda.number_format = '0.00'
        # Odómetro (G)
        if km is not None:
            ws.cell(row=r, column=7).value = km
            km += round(litros[i] * rendimiento_jimny)
        else:
            ws.cell(row=r, column=7).value = "N/F"
        # Combustible (D)
        ws.cell(row=r, column=4).value = datos.get("combustible", "MAGNA")

    # Limpiar filas sobrantes
    for r in range(fila_ini + n, fila_fin + 1):
        ws.cell(row=r, column=1).value = None
        ws.cell(row=r, column=5).value = None
        ws.cell(row=r, column=7).value = None
    return True

# ------------------------------------------------------------------
# GENERACIÓN
# ------------------------------------------------------------------
st.markdown("---")
st.subheader("3. Descargar Formatos Oficiales")

if st.button("🚀 Generar Libro Excel"):
    if wb_temp is None:
        st.error("No hay plantilla. Sube el archivo Excel.")
        st.stop()

    wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.getvalue())) if uploaded_file else openpyxl.load_workbook(plantilla_detectada)
    generados = 0

    if tipo_plantilla == "PEMEX":
        for _, row in df_edit.iterrows():
            if not row.get("activo"): continue
            hoja = row["hoja"]
            if hoja not in wb.sheetnames:
                st.warning(f"⚠️ La hoja `{hoja}` ya no está en el archivo.")
                continue
            ws = wb[hoja]
            fila_ini, fila_fin = encontrar_filas_datos_pemex(ws)
            if fila_ini is None:
                st.warning(f"⚠️ No se reconoció la estructura en `{hoja}`.")
                continue
            n_filas = fila_fin - fila_ini + 1
            max_real = max(row["max_litros_carga"], row["total_litros"] / n_filas)
            cargas = balancear_litros(row["total_litros"], max_real)[:n_filas]
            fechas = fechas_disponibles[:len(cargas)]
            ok = escribir_hoja_pemex(ws, row.to_dict(), fechas, cargas)
            if ok:
                generados += 1

    else:  # BITACORA
        ws_base = wb["BITACORA VEHICULOS"]
        for _, row in df_edit.iterrows():
            if not row.get("activo"): continue
            modo = row.get("modo_carga") or "DINERO"
            precio = PRECIOS.get(row["combustible"], 23.99)
            if modo == "LITROS":
                litros = balancear_litros(row["importe_total"]/precio, row["tope_diario"])
                importes = [round(l*precio,2) for l in litros]
            else:
                importes = balancear_importes(row["importe_total"], row["tope_diario"])
                litros = [round(i/precio,2) for i in importes]
            step_f = len(fechas_disponibles)/len(importes)
            fechas = [fechas_disponibles[int(round(i*step_f))] for i in range(len(importes))]
            cargas = [{"fecha": f.strftime("%d/%m/%Y"), "precio_unit": precio, "importe": imp}
                      for f, imp in zip(fechas, importes)]
            ws_nueva = wb.copy_worksheet(ws_base)
            ws_nueva.title = str(row["placa"]).replace("-","").strip()[:31]
            datos = row.to_dict()
            datos["fecha_txt"] = fecha_entrega_txt
            datos["destino_txt"] = destino_txt
            datos["responsable"] = responsable_txt
            escribir_hoja_bitacora(ws_nueva, datos, cargas)
            generados += 1

    if generados == 0:
        st.warning("No se generó ninguna hoja (revisa los checkboxes y las estructuras).")
    else:
        buffer = io.BytesIO(); wb.save(buffer); buffer.seek(0)
        st.success(f"✅ {generados} hoja(s) generadas.")
        st.download_button("📥 Descargar Excel", data=buffer,
                           file_name=f"BITACORAS_{tipo_plantilla}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
