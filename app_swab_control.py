import re
from io import BytesIO
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Control de Producción por Swab",
    layout="wide"
)

st.title("Control de Producción por Swab")
st.caption(
    "Carga el Excel de Producción por Swab para analizar producción, intervenciones "
    "y pozos convertidos por periodo."
)


def normalizar_pozo(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).upper().strip()
    texto = re.sub(r"\s+", " ", texto)

    return texto


def leer_hoja_con_header(path_or_buffer, sheet_name):
    crudo = pd.read_excel(path_or_buffer, sheet_name=sheet_name, header=None)

    fila_header = None

    for i in range(min(10, len(crudo))):
        valores = crudo.iloc[i].astype(str).str.upper().tolist()

        if "FECHA" in valores:
            fila_header = i
            break

    if fila_header is None:
        fila_header = 0

    df = pd.read_excel(path_or_buffer, sheet_name=sheet_name, header=fila_header)
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    return df


def obtener_pozos_convertidos(path_or_buffer, sheet_name, anio):
    try:
        df = leer_hoja_con_header(path_or_buffer, sheet_name)
    except Exception:
        return []

    columnas = list(df.columns)

    pozos = []

    for col in columnas:
        col_str = str(col).strip()

        if col_str.upper() in ["FECHA", "N° INTERV", "N° INTERV.", "N° INTEV", "TOTAL", "NAN"]:
            continue

        if col_str == "" or col_str.lower() == "nan":
            continue

        if col_str.upper().startswith("UNNAMED"):
            continue

        serie = pd.to_numeric(df[col], errors="coerce")

        if serie.notna().sum() == 0:
            continue

        pozo = normalizar_pozo(col_str)

        if pozo:
            pozos.append(pozo)

    return sorted(list(dict.fromkeys(pozos)))


def leer_datos_swab(path_or_buffer):
    df = leer_hoja_con_header(path_or_buffer, "Datos de Swab")

    columnas_upper = {str(c).strip().upper(): c for c in df.columns}

    col_fecha = columnas_upper.get("FECHA")
    col_pozo = columnas_upper.get("COD_POZ")
    col_bateria = columnas_upper.get("COD_BAT")
    col_unidad = columnas_upper.get("UNIDAD")
    col_prcr = columnas_upper.get("PRCR")
    col_prag = columnas_upper.get("PRAG")
    col_corridas = columnas_upper.get("CORRIDAS")

    if col_fecha is None or col_pozo is None:
        raise ValueError("No se encontró FECHA o COD_POZ en la hoja Datos de Swab.")

    data = pd.DataFrame()

    data["fecha"] = pd.to_datetime(df[col_fecha], errors="coerce")
    data["pozo"] = df[col_pozo].apply(normalizar_pozo)

    if col_bateria:
        data["bateria"] = df[col_bateria].astype(str).str.strip()
    else:
        data["bateria"] = ""

    if col_unidad:
        data["unidad"] = df[col_unidad].astype(str).str.strip()
    else:
        data["unidad"] = ""

    if col_prcr:
        data["produccion"] = pd.to_numeric(df[col_prcr], errors="coerce").fillna(0)
    else:
        data["produccion"] = 0

    if col_prag:
        data["agua"] = pd.to_numeric(df[col_prag], errors="coerce").fillna(0)
    else:
        data["agua"] = 0

    if col_corridas:
        data["corridas"] = pd.to_numeric(df[col_corridas], errors="coerce").fillna(0)
    else:
        data["corridas"] = 0

    data = data.dropna(subset=["fecha"])
    data = data[data["pozo"] != ""].copy()

    data["anio"] = data["fecha"].dt.year
    data["mes"] = data["fecha"].dt.to_period("M").dt.to_timestamp()
    data["intervenciones"] = 1

    return data


def clasificar_pozos(data, pozos_2024, pozos_2025, pozos_2026):
    set_2024 = set(pozos_2024)
    set_2025 = set(pozos_2025)
    set_2026 = set(pozos_2026)

    def clasificar(pozo):
        if pozo in set_2024:
            return "Convertidos 2024"

        if pozo in set_2025:
            return "Convertidos 2025"

        if pozo in set_2026:
            return "Convertidos 2026"

        return "Producción básica"

    data["grupo"] = data["pozo"].apply(clasificar)

    return data


def formato_numero(valor, decimales=2):
    try:
        return f"{float(valor):,.{decimales}f}"
    except Exception:
        return "0"


def generar_documento_html(
    fecha_inicio,
    fecha_fin,
    grupos,
    pozos,
    resumen_general,
    resumen_pozo,
    interv_mes
):
    grupos_txt = ", ".join(grupos)
    pozos_txt = "Todos los pozos seleccionados" if len(pozos) > 20 else ", ".join(pozos)

    filas_pozo = ""

    for _, row in resumen_pozo.head(50).iterrows():
        filas_pozo += f"""
        <tr>
            <td>{row['pozo']}</td>
            <td>{row['grupo']}</td>
            <td style="text-align:right">{formato_numero(row['produccion_total'], 2)}</td>
            <td style="text-align:right">{int(row['intervenciones_total'])}</td>
            <td style="text-align:right">{formato_numero(row['produccion_por_intervencion'], 2)}</td>
            <td>{row['primera_fecha'].date()}</td>
            <td>{row['ultima_fecha'].date()}</td>
        </tr>
        """

    filas_mes = ""

    for _, row in interv_mes.iterrows():
        filas_mes += f"""
        <tr>
            <td>{row['mes'].strftime('%Y-%m')}</td>
            <td>{row['grupo']}</td>
            <td style="text-align:right">{int(row['intervenciones'])}</td>
            <td style="text-align:right">{formato_numero(row['produccion'], 2)}</td>
        </tr>
        """

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>Resumen de Producción por Swab</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 28px;
                color: #1f2937;
            }}
            h1 {{
                color: #0f172a;
            }}
            h2 {{
                color: #1e3a8a;
                margin-top: 26px;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin-top: 10px;
                font-size: 12px;
            }}
            th {{
                background: #1e3a8a;
                color: white;
                padding: 7px;
                border: 1px solid #d1d5db;
            }}
            td {{
                padding: 6px;
                border: 1px solid #d1d5db;
            }}
            .kpi {{
                display: inline-block;
                width: 23%;
                background: #eff6ff;
                border: 1px solid #bfdbfe;
                padding: 12px;
                margin: 4px;
                border-radius: 8px;
            }}
            .kpi b {{
                font-size: 18px;
                color: #1e40af;
            }}
        </style>
    </head>
    <body>
        <h1>Resumen de Producción por Swab</h1>

        <p><b>Periodo analizado:</b> {fecha_inicio} al {fecha_fin}</p>
        <p><b>Grupos seleccionados:</b> {grupos_txt}</p>
        <p><b>Pozos:</b> {pozos_txt}</p>

        <h2>Resumen general</h2>

        <div class="kpi">
            Producción total<br>
            <b>{formato_numero(resumen_general['produccion_total'], 2)}</b>
        </div>

        <div class="kpi">
            Intervenciones<br>
            <b>{int(resumen_general['intervenciones_total'])}</b>
        </div>

        <div class="kpi">
            Pozos activos<br>
            <b>{int(resumen_general['pozos_activos'])}</b>
        </div>

        <div class="kpi">
            Producción por intervención<br>
            <b>{formato_numero(resumen_general['produccion_por_intervencion'], 2)}</b>
        </div>

        <h2>Resumen por pozo</h2>
        <table>
            <tr>
                <th>Pozo</th>
                <th>Grupo</th>
                <th>Producción total</th>
                <th>Intervenciones</th>
                <th>Producción por intervención</th>
                <th>Primera fecha</th>
                <th>Última fecha</th>
            </tr>
            {filas_pozo}
        </table>

        <h2>Intervenciones por mes</h2>
        <table>
            <tr>
                <th>Mes</th>
                <th>Grupo</th>
                <th>Intervenciones</th>
                <th>Producción</th>
            </tr>
            {filas_mes}
        </table>
    </body>
    </html>
    """

    return html


with st.sidebar:
    st.header("Archivo")

    archivo = st.file_uploader(
        "Sube el Excel de Producción por Swab",
        type=["xlsx", "xls"]
    )


if archivo is None:
    st.info("Sube el archivo Excel para empezar.")
    st.stop()


bytes_excel = archivo.getvalue()
buffer = BytesIO(bytes_excel)


try:
    pozos_2024 = obtener_pozos_convertidos(BytesIO(bytes_excel), "Swab 2024 (Dia)", 2024)
    pozos_2025 = obtener_pozos_convertidos(BytesIO(bytes_excel), "Swab 2025 (Dia)", 2025)
    pozos_2026 = obtener_pozos_convertidos(BytesIO(bytes_excel), "Swab 2026 (Dia)", 2026)

    data = leer_datos_swab(BytesIO(bytes_excel))
    data = clasificar_pozos(data, pozos_2024, pozos_2025, pozos_2026)

except Exception as e:
    st.error(f"No pude procesar el archivo: {e}")
    st.stop()


todos_los_pozos = sorted(data["pozo"].unique())

mapa_grupos = {
    "Convertidos 2024": pozos_2024,
    "Convertidos 2025": pozos_2025,
    "Convertidos 2026": pozos_2026,
    "Producción básica": sorted(data.loc[data["grupo"] == "Producción básica", "pozo"].unique()),
    "Todos": todos_los_pozos
}


st.sidebar.header("Filtros")

fecha_min = data["fecha"].min().date()
fecha_max = data["fecha"].max().date()

rango_fechas = st.sidebar.date_input(
    "Rango de fechas",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max
)

if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
    fecha_inicio = pd.to_datetime(rango_fechas[0])
    fecha_fin = pd.to_datetime(rango_fechas[1])
else:
    fecha_inicio = pd.to_datetime(fecha_min)
    fecha_fin = pd.to_datetime(fecha_max)


grupos_disponibles = [
    "Convertidos 2024",
    "Convertidos 2025",
    "Convertidos 2026",
    "Producción básica",
    "Todos"
]

grupos_seleccionados = st.sidebar.multiselect(
    "Tipo de pozos",
    grupos_disponibles,
    default=["Todos"]
)

if not grupos_seleccionados:
    grupos_seleccionados = ["Todos"]


if "Todos" in grupos_seleccionados:
    pozos_base = todos_los_pozos
    grupos_efectivos = ["Todos"]
else:
    pozos_base = []

    for grupo in grupos_seleccionados:
        pozos_base.extend(mapa_grupos.get(grupo, []))

    pozos_base = sorted(list(dict.fromkeys(pozos_base)))
    grupos_efectivos = grupos_seleccionados


pozos_seleccionados = st.sidebar.multiselect(
    "Pozos",
    pozos_base,
    default=pozos_base
)

if not pozos_seleccionados:
    st.warning("Selecciona al menos un pozo.")
    st.stop()


data_filtrada = data[
    (data["fecha"] >= fecha_inicio) &
    (data["fecha"] <= fecha_fin) &
    (data["pozo"].isin(pozos_seleccionados))
].copy()


if data_filtrada.empty:
    st.warning("No hay datos para el rango de fechas y pozos seleccionados.")
    st.stop()


st.subheader("Resumen general del periodo")

produccion_total = data_filtrada["produccion"].sum()
intervenciones_total = data_filtrada["intervenciones"].sum()
corridas_total = data_filtrada["corridas"].sum()
pozos_activos = data_filtrada["pozo"].nunique()
prod_por_interv = produccion_total / intervenciones_total if intervenciones_total > 0 else 0

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Producción total", f"{produccion_total:,.2f}")
k2.metric("Intervenciones", f"{intervenciones_total:,.0f}")
k3.metric("Corridas", f"{corridas_total:,.0f}")
k4.metric("Pozos activos", f"{pozos_activos:,.0f}")
k5.metric("Prod. por intervención", f"{prod_por_interv:,.2f}")


resumen_general = {
    "produccion_total": produccion_total,
    "intervenciones_total": intervenciones_total,
    "corridas_total": corridas_total,
    "pozos_activos": pozos_activos,
    "produccion_por_intervencion": prod_por_interv
}


tab_resumen, tab_pozo, tab_mes, tab_detalle, tab_doc = st.tabs(
    [
        "Resumen por pozo",
        "Producción por pozo",
        "Intervenciones por mes",
        "Detalle",
        "Documento resumen"
    ]
)


with tab_resumen:
    st.subheader("Resumen por pozo")

    resumen_pozo = (
        data_filtrada
        .groupby(["pozo", "grupo"], as_index=False)
        .agg(
            produccion_total=("produccion", "sum"),
            agua_total=("agua", "sum"),
            intervenciones_total=("intervenciones", "sum"),
            corridas_total=("corridas", "sum"),
            primera_fecha=("fecha", "min"),
            ultima_fecha=("fecha", "max"),
            dias_con_intervencion=("fecha", lambda x: x.dt.date.nunique())
        )
    )

    resumen_pozo["produccion_por_intervencion"] = np.where(
        resumen_pozo["intervenciones_total"] > 0,
        resumen_pozo["produccion_total"] / resumen_pozo["intervenciones_total"],
        0
    )

    resumen_pozo = resumen_pozo.sort_values(
        ["produccion_total", "intervenciones_total"],
        ascending=[False, False]
    )

    st.dataframe(
        resumen_pozo.round(3),
        use_container_width=True
    )

    csv_resumen = resumen_pozo.to_csv(index=False).encode("utf-8")

    st.download_button(
        "Descargar resumen por pozo CSV",
        data=csv_resumen,
        file_name="resumen_por_pozo_swab.csv",
        mime="text/csv"
    )


with tab_pozo:
    st.subheader("Producción por pozo en el periodo")

    prod_pozo = (
        data_filtrada
        .groupby("pozo", as_index=False)
        .agg(produccion=("produccion", "sum"))
        .sort_values("produccion", ascending=False)
    )

    fig_prod = px.bar(
        prod_pozo.head(40),
        x="pozo",
        y="produccion",
        title="Producción total por pozo",
        labels={
            "pozo": "Pozo",
            "produccion": "Producción"
        }
    )

    fig_prod.update_layout(
        xaxis_tickangle=-45,
        height=520
    )

    st.plotly_chart(fig_prod, use_container_width=True)

    st.subheader("Producción diaria por pozo seleccionado")

    prod_diaria = (
        data_filtrada
        .groupby(["fecha", "pozo"], as_index=False)
        .agg(produccion=("produccion", "sum"))
    )

    fig_linea = px.line(
        prod_diaria,
        x="fecha",
        y="produccion",
        color="pozo",
        title="Producción diaria por pozo",
        labels={
            "fecha": "Fecha",
            "produccion": "Producción",
            "pozo": "Pozo"
        }
    )

    fig_linea.update_layout(height=560)

    st.plotly_chart(fig_linea, use_container_width=True)


with tab_mes:
    st.subheader("Intervenciones por mes")

    interv_mes = (
        data_filtrada
        .groupby(["mes", "grupo"], as_index=False)
        .agg(
            intervenciones=("intervenciones", "sum"),
            produccion=("produccion", "sum"),
            pozos=("pozo", "nunique")
        )
    )

    fig_interv = px.bar(
        interv_mes,
        x="mes",
        y="intervenciones",
        color="grupo",
        barmode="group",
        title="Intervenciones por mes y grupo",
        labels={
            "mes": "Mes",
            "intervenciones": "Intervenciones",
            "grupo": "Grupo"
        }
    )

    fig_interv.update_layout(height=540)

    st.plotly_chart(fig_interv, use_container_width=True)

    st.dataframe(
        interv_mes.round(3),
        use_container_width=True
    )

    st.subheader("Intervenciones mensuales por pozo")

    interv_mes_pozo = (
        data_filtrada
        .groupby(["mes", "pozo", "grupo"], as_index=False)
        .agg(
            intervenciones=("intervenciones", "sum"),
            produccion=("produccion", "sum"),
            corridas=("corridas", "sum")
        )
    )

    st.dataframe(
        interv_mes_pozo.round(3),
        use_container_width=True
    )

    csv_interv = interv_mes_pozo.to_csv(index=False).encode("utf-8")

    st.download_button(
        "Descargar intervenciones mensuales por pozo CSV",
        data=csv_interv,
        file_name="intervenciones_mensuales_por_pozo.csv",
        mime="text/csv"
    )


with tab_detalle:
    st.subheader("Detalle de registros filtrados")

    columnas_detalle = [
        "fecha",
        "mes",
        "pozo",
        "grupo",
        "bateria",
        "unidad",
        "produccion",
        "agua",
        "intervenciones",
        "corridas"
    ]

    st.dataframe(
        data_filtrada[columnas_detalle].sort_values(["fecha", "pozo"]),
        use_container_width=True
    )

    csv_detalle = data_filtrada[columnas_detalle].to_csv(index=False).encode("utf-8")

    st.download_button(
        "Descargar detalle filtrado CSV",
        data=csv_detalle,
        file_name="detalle_swab_filtrado.csv",
        mime="text/csv"
    )


with tab_doc:
    st.subheader("Documento resumen descargable")

    resumen_pozo_doc = (
        data_filtrada
        .groupby(["pozo", "grupo"], as_index=False)
        .agg(
            produccion_total=("produccion", "sum"),
            intervenciones_total=("intervenciones", "sum"),
            primera_fecha=("fecha", "min"),
            ultima_fecha=("fecha", "max")
        )
    )

    resumen_pozo_doc["produccion_por_intervencion"] = np.where(
        resumen_pozo_doc["intervenciones_total"] > 0,
        resumen_pozo_doc["produccion_total"] / resumen_pozo_doc["intervenciones_total"],
        0
    )

    resumen_pozo_doc = resumen_pozo_doc.sort_values(
        "produccion_total",
        ascending=False
    )

    interv_mes_doc = (
        data_filtrada
        .groupby(["mes", "grupo"], as_index=False)
        .agg(
            intervenciones=("intervenciones", "sum"),
            produccion=("produccion", "sum")
        )
    )

    html = generar_documento_html(
        fecha_inicio.date(),
        fecha_fin.date(),
        grupos_efectivos,
        pozos_seleccionados,
        resumen_general,
        resumen_pozo_doc,
        interv_mes_doc
    )

    st.markdown(
        """
        Este documento resume el periodo seleccionado, los grupos elegidos,
        la producción total, las intervenciones, el ranking por pozo y las intervenciones por mes.
        """
    )

    st.download_button(
        "Descargar documento resumen HTML",
        data=html.encode("utf-8"),
        file_name="resumen_produccion_swab.html",
        mime="text/html"
    )

    st.components.v1.html(html, height=700, scrolling=True)
