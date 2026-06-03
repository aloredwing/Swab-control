
import io
import re
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Control SWAB",
    page_icon="🛢️",
    layout="wide",
)

# ============================================================
# LISTAS FIJAS DE POZOS CONVERTIDOS
# Extraídas del archivo Producción por Swab.xlsx:
# Swab 2024 (Dia), Swab 2025 (Dia), Swab 2026 (Dia)
# Producción básica se calcula como todo pozo que no esté en estas listas.
# ============================================================

POZOS_CONVERTIDOS_2024 = [
    "EA11934D", "EA 1868", "EA 8291", "EA 8561", "EA 8547", "EA11873D",
    "EA11054", "EA 8569", "AA11122D", "EA11204D", "EA 1625", "EA 8532",
    "EA11703D", "EA11591D", "AA 6316", "EA 8593", "EA 8506", "AA11976D",
    "AA 5804", "EA 8209", "EA 8773", "AA11958D", "AA11214D", "EA11968D",
    "EA11221D", "EA 5997", "EA 8523", "EA 8783", "EA 5718", "EA 8317",
    "AA 8309D", "AA 6089", "EA11669D", "EA 1763", "EA 8961D", "EA11708D",
    "AA11408D", "EA 8901", "EA 7991", "EA11033", "EA11931D", "EA 8716",
    "EA11748D", "EA 9472", "EA11194", "EA11656D", "AA11342D", "EA 8687",
    "EA11431D",
]

POZOS_CONVERTIDOS_2025 = [
    "AA11062D", "EA11049D", "EA 8204", "EA 8043", "AA11118D", "EA11918D",
    "AA11702D", "AA11396D", "EA11929D", "EA 8544", "EA 2324", "EA 9241",
    "EA11224D", "EA11884D", "AA 8149D", "AA11363D", "AA 6274", "EA 5998",
    "EA 8088", "EA 8902", "EA 8724", "EA11813D", "EA 8759", "EA 7516",
    "AA 5931", "AA11801D",
]

POZOS_CONVERTIDOS_2026 = [
    "EA 9836", "EA 8417", "EA 9538", "EA 5811", "EA 8672", "EA11444",
    "AA11172D", "EA 8586", "AA   47", "AA11762", "EA11609D", "AA11144D",
    "EA 9041", "EA11388D", "AA11481D", "AA11327D", "AA  106", "EA8819",
    "EA8816D", "EA8662D",
]


def normalizar_pozo(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).upper().replace("\xa0", " ")
    texto = re.sub(r"\s+", "", texto)
    return texto.strip()


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).upper().replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


CONV_2024_N = {normalizar_pozo(p) for p in POZOS_CONVERTIDOS_2024}
CONV_2025_N = {normalizar_pozo(p) for p in POZOS_CONVERTIDOS_2025}
CONV_2026_N = {normalizar_pozo(p) for p in POZOS_CONVERTIDOS_2026}
CONVERTIDOS_TODOS_N = CONV_2024_N | CONV_2025_N | CONV_2026_N


def clasificar_grupo(pozo_norm):
    if pozo_norm in CONV_2024_N:
        return "Convertidos 2024"
    if pozo_norm in CONV_2025_N:
        return "Convertidos 2025"
    if pozo_norm in CONV_2026_N:
        return "Convertidos 2026"
    return "Producción básica"


def limpiar_columna(columna):
    texto = str(columna).strip().upper().replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def leer_archivo_swab(archivo):
    nombre = archivo.name.lower()

    if nombre.endswith(".csv"):
        archivo.seek(0)
        df = pd.read_csv(archivo)
        hoja_usada = "CSV"
    else:
        archivo.seek(0)
        xls = pd.ExcelFile(archivo)
        hojas = xls.sheet_names

        hoja_usada = None

        # En tu archivo real, la producción está en Datos de Swab.
        # La primera pestaña es Baterias Antiguas y no contiene PRCR ni PRAG.
        for h in hojas:
            if h.strip().lower() == "datos de swab":
                hoja_usada = h
                break

        # Si otro archivo no trae ese nombre, busca la hoja que tenga las columnas reales.
        if hoja_usada is None:
            for h in hojas:
                archivo.seek(0)
                preview = pd.read_excel(archivo, sheet_name=h, nrows=3)
                cols = {limpiar_columna(c) for c in preview.columns}
                if {"FECHA", "COD_POZ", "COD_BAT", "PRCR", "PRAG"}.issubset(cols):
                    hoja_usada = h
                    break

        if hoja_usada is None:
            hoja_usada = hojas[0]

        archivo.seek(0)
        df = pd.read_excel(archivo, sheet_name=hoja_usada)

    df.columns = [limpiar_columna(c) for c in df.columns]
    return df, hoja_usada


def preparar_data(df):
    requeridas = ["FECHA", "COD_POZ", "COD_BAT", "PRCR", "PRAG"]
    faltantes = [c for c in requeridas if c not in df.columns]

    if faltantes:
        return None, faltantes

    data = df.copy()
    data["FECHA"] = pd.to_datetime(data["FECHA"], errors="coerce")
    data["POZO"] = data["COD_POZ"].apply(limpiar_texto)
    data["POZO_NORM"] = data["COD_POZ"].apply(normalizar_pozo)
    data["BATERIA"] = data["COD_BAT"].apply(limpiar_texto)

    data["PRCR"] = pd.to_numeric(data["PRCR"], errors="coerce").fillna(0.0)
    data["PRAG"] = pd.to_numeric(data["PRAG"], errors="coerce").fillna(0.0)

    data = data.dropna(subset=["FECHA"])
    data = data[data["POZO_NORM"] != ""].copy()

    data["GRUPO"] = data["POZO_NORM"].apply(clasificar_grupo)
    data["MES"] = data["FECHA"].dt.to_period("M").astype(str)

    return data, []


def estado_activo_por_mes(data_base, pozos_norm, mes_corte):
    inicio = pd.Timestamp(mes_corte + "-01")
    fin = inicio + pd.offsets.MonthEnd(0)

    base_mes = data_base[
        (data_base["FECHA"] >= inicio)
        & (data_base["FECHA"] <= fin)
        & (data_base["POZO_NORM"].isin(pozos_norm))
    ].copy()

    resumen_mes = (
        base_mes.groupby(["POZO_NORM", "POZO"], as_index=False)
        .agg(
            PRCR_MES=("PRCR", "sum"),
            PRAG_MES=("PRAG", "sum"),
            INTERV_MES=("POZO", "size"),
        )
    )

    pozos_df = (
        data_base[data_base["POZO_NORM"].isin(pozos_norm)]
        .groupby("POZO_NORM", as_index=False)
        .agg(
            POZO=("POZO", "first"),
            BATERIA=("BATERIA", lambda x: x.mode().iloc[0] if not x.mode().empty else ""),
            GRUPO=("GRUPO", "first"),
        )
    )

    resumen = pozos_df.merge(resumen_mes, on=["POZO_NORM", "POZO"], how="left")
    resumen["PRCR_MES"] = resumen["PRCR_MES"].fillna(0.0)
    resumen["PRAG_MES"] = resumen["PRAG_MES"].fillna(0.0)
    resumen["INTERV_MES"] = resumen["INTERV_MES"].fillna(0).astype(int)
    resumen["ESTADO"] = np.where(resumen["PRCR_MES"] > 0, "Activo", "Inactivo")

    resumen = resumen.sort_values(["ESTADO", "GRUPO", "BATERIA", "POZO"])
    return resumen, inicio, fin


def resumen_diario(data, segmentar_por, variable):
    if segmentar_por == "Total general":
        diario = data.groupby("FECHA", as_index=False).agg(
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum"),
            INTERVENCIONES=("POZO", "size"),
            POZOS=("POZO_NORM", "nunique"),
        )
        diario["SEGMENTO"] = "Total general"
    else:
        col = {
            "Grupo de conversión": "GRUPO",
            "Batería": "BATERIA",
            "Pozo": "POZO",
            "Estado activo/inactivo": "ESTADO",
        }[segmentar_por]

        diario = data.groupby(["FECHA", col], as_index=False).agg(
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum"),
            INTERVENCIONES=("POZO", "size"),
            POZOS=("POZO_NORM", "nunique"),
        )
        diario = diario.rename(columns={col: "SEGMENTO"})

    diario = diario.sort_values(["FECHA", "SEGMENTO"])
    diario["PROM_MOVIL_7D"] = (
        diario.groupby("SEGMENTO")[variable]
        .transform(lambda s: s.rolling(7, min_periods=1).mean())
    )
    return diario


def resumen_mensual(data, segmentar_por):
    if segmentar_por == "Total general":
        mensual = data.groupby("MES", as_index=False).agg(
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum"),
            INTERVENCIONES=("POZO", "size"),
            POZOS=("POZO_NORM", "nunique"),
        )
        mensual["SEGMENTO"] = "Total general"
    else:
        col = {
            "Grupo de conversión": "GRUPO",
            "Batería": "BATERIA",
            "Pozo": "POZO",
            "Estado activo/inactivo": "ESTADO",
        }[segmentar_por]
        mensual = data.groupby(["MES", col], as_index=False).agg(
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum"),
            INTERVENCIONES=("POZO", "size"),
            POZOS=("POZO_NORM", "nunique"),
        )
        mensual = mensual.rename(columns={col: "SEGMENTO"})
    return mensual.sort_values(["MES", "SEGMENTO"])


def calcular_caidas(data):
    mensual_pozo = (
        data.groupby(["POZO_NORM", "POZO", "GRUPO", "BATERIA", "MES"], as_index=False)
        .agg(PRCR=("PRCR", "sum"), PRAG=("PRAG", "sum"), INTERVENCIONES=("POZO", "size"))
        .sort_values(["POZO_NORM", "MES"])
    )
    mensual_pozo["PRCR_MES_ANT"] = mensual_pozo.groupby("POZO_NORM")["PRCR"].shift(1)
    mensual_pozo["CAIDA_PRCR"] = mensual_pozo["PRCR"] - mensual_pozo["PRCR_MES_ANT"]
    mensual_pozo["CAIDA_PCT"] = np.where(
        mensual_pozo["PRCR_MES_ANT"] > 0,
        mensual_pozo["CAIDA_PRCR"] / mensual_pozo["PRCR_MES_ANT"] * 100,
        np.nan,
    )
    caidas = mensual_pozo.dropna(subset=["PRCR_MES_ANT"]).copy()
    caidas = caidas.sort_values("CAIDA_PRCR")
    return caidas


def descargar_excel(resultados):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre, df in resultados.items():
            safe = nombre[:31]
            df.to_excel(writer, index=False, sheet_name=safe)
    output.seek(0)
    return output.getvalue()


st.title("🛢️ Control SWAB en Streamlit")
st.caption("Versión corregida para tu archivo: usa Datos de Swab, COD_POZ, COD_BAT, PRCR y PRAG. Ejecuta recién cuando presionas el botón.")

with st.expander("Qué hace esta versión", expanded=True):
    st.write(
        """
        Esta app ya no pide mapear columnas manualmente. En tu Excel la primera pestaña es Baterias Antiguas,
        pero la producción real está en Datos de Swab. Por eso esta versión usa automáticamente esa hoja.

        PRCR se toma como producción de petróleo. PRAG se toma como producción de agua.
        Los pozos convertidos 2024, 2025 y 2026 están escritos dentro del código.
        Producción básica se calcula como los pozos que no están en esas listas.

        El estado activo o inactivo se calcula por mes calendario. Si seleccionas mayo,
        evalúa del 01 de mayo al último día de mayo. No usa treinta días hacia atrás.
        """
    )

archivo = st.sidebar.file_uploader(
    "Carga archivo SWAB",
    type=["xlsx", "xls", "csv"],
)

if archivo is None:
    st.info("Carga tu archivo Producción por Swab.xlsx para iniciar.")
    st.stop()

try:
    df_raw, hoja_usada = leer_archivo_swab(archivo)
except Exception as e:
    st.error("No se pudo leer el archivo. Verifica que sea Excel o CSV válido.")
    st.exception(e)
    st.stop()

data, faltantes = preparar_data(df_raw)

if faltantes:
    st.error("Faltan columnas obligatorias o no fueron reconocidas: " + ", ".join(faltantes))
    st.write("Hoja usada:", hoja_usada)
    st.write("Columnas encontradas:")
    st.write(list(df_raw.columns))
    st.dataframe(df_raw.head(20), use_container_width=True)
    st.stop()

st.success(f"Archivo leído correctamente. Hoja usada: {hoja_usada}. Registros: {len(data):,}. Pozos: {data['POZO_NORM'].nunique():,}.")

conteo_grupo = data.groupby("GRUPO")["POZO_NORM"].nunique().reset_index(name="N_POZOS")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Convertidos 2024 en lista", len(CONV_2024_N))
c2.metric("Convertidos 2025 en lista", len(CONV_2025_N))
c3.metric("Convertidos 2026 en lista", len(CONV_2026_N))
c4.metric("Pozos en producción", data["POZO_NORM"].nunique())

st.sidebar.subheader("Filtros de análisis")

grupos_disponibles = [
    "Convertidos 2024",
    "Convertidos 2025",
    "Convertidos 2026",
    "Producción básica",
]
grupos_sel = st.sidebar.multiselect(
    "Grupos de pozos",
    options=grupos_disponibles,
    default=grupos_disponibles,
)

data_grupo = data[data["GRUPO"].isin(grupos_sel)].copy()

baterias_disponibles = sorted([b for b in data_grupo["BATERIA"].dropna().unique() if b != ""])
usar_todas_baterias = st.sidebar.checkbox("Usar todas las baterías", value=True)

if usar_todas_baterias:
    baterias_sel = baterias_disponibles
else:
    baterias_sel = st.sidebar.multiselect(
        "Baterías",
        options=baterias_disponibles,
        default=baterias_disponibles[:10],
    )

data_pre = data_grupo[data_grupo["BATERIA"].isin(baterias_sel)].copy()

pozos_disponibles_df = (
    data_pre.groupby(["POZO_NORM", "POZO"], as_index=False)
    .agg(PRCR_TOTAL=("PRCR", "sum"))
    .sort_values(["POZO"])
)
pozos_disponibles = pozos_disponibles_df["POZO"].tolist()
pozo_display_to_norm = dict(zip(pozos_disponibles_df["POZO"], pozos_disponibles_df["POZO_NORM"]))

usar_todos_pozos = st.sidebar.checkbox("Usar todos los pozos filtrados", value=True)

if usar_todos_pozos:
    pozos_sel_display = pozos_disponibles
else:
    pozos_sel_display = st.sidebar.multiselect(
        "Pozos a analizar",
        options=pozos_disponibles,
        default=pozos_disponibles[:20],
    )

pozos_sel_norm = [pozo_display_to_norm[p] for p in pozos_sel_display if p in pozo_display_to_norm]

fecha_min = data_pre["FECHA"].min().date()
fecha_max = data_pre["FECHA"].max().date()

rango = st.sidebar.date_input(
    "Rango de fechas para análisis",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max,
)

meses_disponibles = sorted(data_pre["MES"].dropna().unique().tolist())
mes_default = meses_disponibles[-1] if meses_disponibles else ""
mes_corte = st.sidebar.selectbox(
    "Mes calendario para activo/inactivo",
    options=meses_disponibles,
    index=meses_disponibles.index(mes_default) if mes_default in meses_disponibles else 0,
)

segmentar_por = st.sidebar.selectbox(
    "Segmentar producción diaria por",
    options=[
        "Total general",
        "Grupo de conversión",
        "Batería",
        "Pozo",
        "Estado activo/inactivo",
    ],
)

variable_grafica = st.sidebar.selectbox(
    "Variable principal",
    options=["PRCR", "PRAG"],
    format_func=lambda x: "PRCR petróleo" if x == "PRCR" else "PRAG agua",
)

tipo_grafica = st.sidebar.radio(
    "Tipo de gráfica diaria",
    options=["Línea", "Barras"],
    horizontal=True,
)

st.sidebar.markdown("---")
ejecutar = st.sidebar.button("Ejecutar análisis", type="primary", use_container_width=True)

st.subheader("Selección antes de ejecutar")
prev1, prev2, prev3 = st.columns(3)
prev1.metric("Pozos seleccionados", len(pozos_sel_norm))
prev2.metric("Baterías seleccionadas", len(baterias_sel))
prev3.metric("Grupos seleccionados", len(grupos_sel))

if not ejecutar:
    st.info("Selecciona grupos, baterías, pozos, rango de fechas y mes de corte. Luego presiona Ejecutar análisis en el panel lateral.")
    st.dataframe(conteo_grupo, use_container_width=True)
    st.stop()

if len(pozos_sel_norm) == 0:
    st.error("No hay pozos seleccionados. Selecciona al menos un pozo o activa Usar todos los pozos filtrados.")
    st.stop()

if isinstance(rango, tuple) or isinstance(rango, list):
    if len(rango) != 2:
        st.error("Selecciona fecha inicial y fecha final.")
        st.stop()
    fecha_ini = pd.Timestamp(rango[0])
    fecha_fin = pd.Timestamp(rango[1])
else:
    fecha_ini = pd.Timestamp(rango)
    fecha_fin = pd.Timestamp(rango)

if fecha_ini > fecha_fin:
    st.error("La fecha inicial no puede ser mayor que la fecha final.")
    st.stop()

estado_mes, inicio_mes, fin_mes = estado_activo_por_mes(data, pozos_sel_norm, mes_corte)
map_estado = dict(zip(estado_mes["POZO_NORM"], estado_mes["ESTADO"]))

data_filtrada = data[
    (data["GRUPO"].isin(grupos_sel))
    & (data["BATERIA"].isin(baterias_sel))
    & (data["POZO_NORM"].isin(pozos_sel_norm))
    & (data["FECHA"] >= fecha_ini)
    & (data["FECHA"] <= fecha_fin)
].copy()

data_filtrada["ESTADO"] = data_filtrada["POZO_NORM"].map(map_estado).fillna("Inactivo")

if data_filtrada.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()

diario = resumen_diario(data_filtrada, segmentar_por, variable_grafica)
mensual = resumen_mensual(data_filtrada, segmentar_por)
caidas = calcular_caidas(data_filtrada)

total_prcr = data_filtrada["PRCR"].sum()
total_prag = data_filtrada["PRAG"].sum()
total_interv = len(data_filtrada)
pozos_con_prcr = data_filtrada.loc[data_filtrada["PRCR"] > 0, "POZO_NORM"].nunique()

k1, k2, k3, k4 = st.columns(4)
k1.metric("PRCR total", f"{total_prcr:,.2f}")
k2.metric("PRAG total", f"{total_prag:,.2f}")
k3.metric("Intervenciones", f"{total_interv:,}")
k4.metric("Pozos con PRCR", f"{pozos_con_prcr:,}")

st.caption(f"Estado activo/inactivo calculado con corte calendario: {inicio_mes.date()} al {fin_mes.date()}.")

tab_diario, tab_estado, tab_mensual, tab_interv, tab_caidas, tab_pozo, tab_datos = st.tabs(
    [
        "Producción diaria",
        "Activos e inactivos",
        "Producción mensual",
        "Intervenciones",
        "Caídas",
        "Detalle por pozo",
        "Datos y descarga",
    ]
)

with tab_diario:
    st.subheader("Producción diaria SWAB")
    st.write("Esta vista resume la producción diaria de los pozos seleccionados.")

    if segmentar_por == "Pozo":
        top_n = st.slider("Mostrar top N pozos en la gráfica", min_value=5, max_value=60, value=20, step=5)
        top_segmentos = (
            diario.groupby("SEGMENTO")[variable_grafica]
            .sum()
            .sort_values(ascending=False)
            .head(top_n)
            .index
            .tolist()
        )
        diario_plot = diario[diario["SEGMENTO"].isin(top_segmentos)].copy()
    else:
        diario_plot = diario.copy()

    if tipo_grafica == "Línea":
        fig = px.line(
            diario_plot,
            x="FECHA",
            y=variable_grafica,
            color="SEGMENTO",
            markers=False,
            title=f"Producción diaria {variable_grafica}",
        )
    else:
        fig = px.bar(
            diario_plot,
            x="FECHA",
            y=variable_grafica,
            color="SEGMENTO",
            title=f"Producción diaria {variable_grafica}",
        )

    fig.update_layout(height=520, hovermode="x unified")
    st.plotly_chart(fig, use_container_width=True)

    fig_pm = px.line(
        diario_plot,
        x="FECHA",
        y="PROM_MOVIL_7D",
        color="SEGMENTO",
        title=f"Promedio móvil 7 días de {variable_grafica}",
    )
    fig_pm.update_layout(height=420, hovermode="x unified")
    st.plotly_chart(fig_pm, use_container_width=True)

    st.dataframe(diario, use_container_width=True, height=350)

with tab_estado:
    st.subheader("Pozos activos e inactivos")
    st.write("Activo significa que el pozo tuvo PRCR mayor que cero dentro del mes calendario seleccionado.")

    col_a, col_i = st.columns(2)
    n_activos = int((estado_mes["ESTADO"] == "Activo").sum())
    n_inactivos = int((estado_mes["ESTADO"] == "Inactivo").sum())
    col_a.metric("Activos", n_activos)
    col_i.metric("Inactivos", n_inactivos)

    estado_filtro = st.radio(
        "Ver estado",
        options=["Todos", "Activo", "Inactivo"],
        horizontal=True,
    )
    estado_show = estado_mes.copy()
    if estado_filtro != "Todos":
        estado_show = estado_show[estado_show["ESTADO"] == estado_filtro]

    st.dataframe(
        estado_show[["POZO", "BATERIA", "GRUPO", "PRCR_MES", "PRAG_MES", "INTERV_MES", "ESTADO"]],
        use_container_width=True,
        height=500,
    )

    fig_estado = px.bar(
        estado_mes.groupby(["GRUPO", "ESTADO"], as_index=False)["POZO"].count(),
        x="GRUPO",
        y="POZO",
        color="ESTADO",
        barmode="group",
        title="Cantidad de pozos activos e inactivos por grupo",
        labels={"POZO": "N° pozos"},
    )
    fig_estado.update_layout(height=420)
    st.plotly_chart(fig_estado, use_container_width=True)

with tab_mensual:
    st.subheader("Producción mensual")
    fig_m = px.bar(
        mensual,
        x="MES",
        y=variable_grafica,
        color="SEGMENTO",
        title=f"Producción mensual {variable_grafica}",
    )
    fig_m.update_layout(height=520)
    st.plotly_chart(fig_m, use_container_width=True)

    st.dataframe(mensual, use_container_width=True, height=400)

with tab_interv:
    st.subheader("Intervenciones por mes")
    fig_int = px.bar(
        mensual,
        x="MES",
        y="INTERVENCIONES",
        color="SEGMENTO",
        title="Intervenciones mensuales",
    )
    fig_int.update_layout(height=520)
    st.plotly_chart(fig_int, use_container_width=True)

    st.dataframe(mensual[["MES", "SEGMENTO", "INTERVENCIONES", "POZOS"]], use_container_width=True, height=400)

with tab_caidas:
    st.subheader("Caídas de PRCR")
    st.write("Compara cada mes contra el mes anterior para cada pozo seleccionado.")

    solo_caidas = caidas[caidas["CAIDA_PRCR"] < 0].copy()
    top_caidas = solo_caidas.head(50)

    st.dataframe(
        top_caidas[["POZO", "BATERIA", "GRUPO", "MES", "PRCR_MES_ANT", "PRCR", "CAIDA_PRCR", "CAIDA_PCT", "INTERVENCIONES"]],
        use_container_width=True,
        height=450,
    )

    if not top_caidas.empty:
        fig_c = px.bar(
            top_caidas.head(25).sort_values("CAIDA_PRCR"),
            x="CAIDA_PRCR",
            y="POZO",
            color="GRUPO",
            orientation="h",
            title="Mayores caídas de PRCR por pozo",
        )
        fig_c.update_layout(height=600)
        st.plotly_chart(fig_c, use_container_width=True)

with tab_pozo:
    st.subheader("Detalle por pozo")
    pozos_en_data = sorted(data_filtrada["POZO"].dropna().unique().tolist())
    pozo_det = st.selectbox("Selecciona un pozo", options=pozos_en_data)

    dpozo = data_filtrada[data_filtrada["POZO"] == pozo_det].copy()
    dpozo_diario = dpozo.groupby("FECHA", as_index=False).agg(
        PRCR=("PRCR", "sum"),
        PRAG=("PRAG", "sum"),
        INTERVENCIONES=("POZO", "size"),
    )

    p1, p2, p3 = st.columns(3)
    p1.metric("PRCR pozo", f"{dpozo['PRCR'].sum():,.2f}")
    p2.metric("PRAG pozo", f"{dpozo['PRAG'].sum():,.2f}")
    p3.metric("Intervenciones", f"{len(dpozo):,}")

    fig_p = px.line(
        dpozo_diario,
        x="FECHA",
        y=["PRCR", "PRAG"],
        title=f"Producción diaria del pozo {pozo_det}",
    )
    fig_p.update_layout(height=500, hovermode="x unified")
    st.plotly_chart(fig_p, use_container_width=True)

    st.dataframe(dpozo.sort_values("FECHA"), use_container_width=True, height=400)

with tab_datos:
    st.subheader("Datos y descarga")
    st.write("Vista de datos filtrados y archivos de salida.")

    st.dataframe(data_filtrada.head(1000), use_container_width=True, height=420)

    resultados = {
        "Produccion_diaria": diario,
        "Activos_inactivos": estado_mes,
        "Produccion_mensual": mensual,
        "Caidas_PRCR": caidas,
        "Datos_filtrados": data_filtrada,
    }

    excel_bytes = descargar_excel(resultados)
    st.download_button(
        "Descargar reporte Excel",
        data=excel_bytes,
        file_name="reporte_swab.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    csv_bytes = diario.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Descargar producción diaria CSV",
        data=csv_bytes,
        file_name="produccion_diaria_swab.csv",
        mime="text/csv",
        use_container_width=True,
    )
