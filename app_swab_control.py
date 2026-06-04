import io
import re

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt


st.set_page_config(
    page_title="Dashboard SWAB Lote X",
    page_icon="🛢️",
    layout="wide"
)


# ============================================================
# POZOS CONVERTIDOS
# Estos pozos quedan fijos en el código.
# Todo pozo que no esté en estas listas se considera BÁSICA.
# ============================================================

CONVERTIDOS = {
    2024: [
        "EA11934D", "EA 1868", "EA 8291", "EA 8561", "EA 8547",
        "EA11873D", "EA11054", "EA 8569", "AA11122D", "EA11204D",
        "EA 1625", "EA 8532", "EA11703D", "EA11591D", "AA 6316",
        "EA 8593", "EA 8506", "AA11976D", "AA 5804", "EA 8209",
        "EA 8773", "AA11958D", "AA11214D", "EA11968D", "EA11221D",
        "EA 5997", "EA 8523", "EA 8783", "EA 5718", "EA 8317",
        "AA 8309D", "AA 6089", "EA11669D", "EA 1763", "EA 8961D",
        "EA11708D", "AA11408D", "EA 8901", "EA 7991", "EA11033",
        "EA11931D", "EA 8716", "EA11748D", "EA 9472", "EA11194",
        "EA11656D", "AA11342D", "EA 8687", "EA11431D"
    ],
    2025: [
        "AA11062D", "EA11049D", "EA 8204", "EA 8043", "AA11118D",
        "EA11918D", "AA11702D", "AA11396D", "EA11929D", "EA 8544",
        "EA 2324", "EA 9241", "EA11224D", "EA11884D", "AA 8149D",
        "AA11363D", "AA 6274", "EA 5998", "EA 8088", "EA 8902",
        "EA 8724", "EA11813D", "EA 8759", "EA 7516", "AA 5931",
        "AA11801D"
    ],
    2026: [
        "EA 9836", "EA 8417", "EA 9538", "EA 5811", "EA 8672",
        "EA11444", "AA11172D", "EA 8586", "AA   47", "AA11762",
        "EA11609D", "AA11144D", "EA 9041", "EA11388D", "AA11481D",
        "AA11327D", "AA  106", "EA8819", "EA8816D", "EA8662D"
    ]
}


MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre"
}


# ============================================================
# FUNCIONES DE LIMPIEZA
# ============================================================

def limpiar_pozo(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = re.sub(r"\s+", "", texto)
    return texto


def mostrar_pozo(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_columna(columna):
    texto = str(columna).strip().upper()
    texto = texto.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    texto = re.sub(r"\s+", "_", texto)
    return texto


def convertir_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def clasificar_pozo(pozo_key):
    for anio, pozos in CONVERTIDOS.items():
        lista_key = [limpiar_pozo(p) for p in pozos]
        if pozo_key in lista_key:
            return f"Convertido {anio}"
    return "Básica"


def obtener_anio_conversion(pozo_key):
    for anio, pozos in CONVERTIDOS.items():
        lista_key = [limpiar_pozo(p) for p in pozos]
        if pozo_key in lista_key:
            return anio
    return 0


def normalizar_tipo_swab(valor):
    texto = limpiar_texto(valor)

    if texto in ["TBG", "TB", "TS", "TUBING", "TUBING SWAB"]:
        return "TS"

    if texto in ["CS", "CSG", "CASING", "CASING SWAB"]:
        return "CS"

    if "TBG" in texto or "TUB" in texto or texto == "TS":
        return "TS"

    if "CS" in texto or "CAS" in texto:
        return "CS"

    if texto == "":
        return "SIN TIPO"

    return texto


def periodo_texto(anio, mes):
    if mes == 0:
        return f"Todo el año {anio}"
    return f"{MESES[mes]} {anio}"


# ============================================================
# CARGA DE EXCEL
# ============================================================

@st.cache_data(show_spinner=False)
def cargar_excel_swab(bytes_excel):
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))

    if "Datos de Swab" in xls.sheet_names:
        hoja = "Datos de Swab"
    else:
        hoja = xls.sheet_names[0]

    df = pd.read_excel(io.BytesIO(bytes_excel), sheet_name=hoja)
    df.columns = [normalizar_columna(c) for c in df.columns]

    requeridas = ["FECHA", "COD_POZ", "COD_BAT", "UNIDAD", "TSER", "PRCR", "PRAG"]
    faltantes = [c for c in requeridas if c not in df.columns]

    if faltantes:
        raise ValueError(
            "Faltan estas columnas obligatorias: "
            + ", ".join(faltantes)
            + ". Revisa que el Excel tenga FECHA, COD_POZ, COD_BAT, UNIDAD, TSER, PRCR y PRAG."
        )

    df = df[requeridas].copy()

    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df = df[df["FECHA"].notna()].copy()

    df["POZO"] = df["COD_POZ"].apply(mostrar_pozo)
    df["POZO_KEY"] = df["COD_POZ"].apply(limpiar_pozo)
    df["BATERIA"] = df["COD_BAT"].apply(limpiar_texto)
    df["UNIDAD"] = df["UNIDAD"].apply(limpiar_texto)
    df["TIPO_SWAB"] = df["TSER"].apply(normalizar_tipo_swab)
    df["PRCR"] = convertir_numero(df["PRCR"])
    df["PRAG"] = convertir_numero(df["PRAG"])

    df = df[df["POZO_KEY"] != ""].copy()

    df["ANIO"] = df["FECHA"].dt.year
    df["MES"] = df["FECHA"].dt.month
    df["MES_NOMBRE"] = df["MES"].map(MESES)

    df["CLASIFICACION"] = df["POZO_KEY"].apply(clasificar_pozo)
    df["ANIO_CONVERSION"] = df["POZO_KEY"].apply(obtener_anio_conversion)

    return df, hoja


def construir_universo(df):
    universo = (
        df.sort_values("FECHA")
        .groupby("POZO_KEY", as_index=False)
        .agg(
            POZO=("POZO", "last"),
            BATERIA=("BATERIA", "last"),
            CLASIFICACION=("CLASIFICACION", "last"),
            ANIO_CONVERSION=("ANIO_CONVERSION", "last"),
            ULTIMA_FECHA_HISTORICA=("FECHA", "max")
        )
    )

    return universo


# ============================================================
# CÁLCULOS
# ============================================================

def aplicar_filtros_universo(universo, baterias_sel, clases_sel):
    data = universo.copy()

    if baterias_sel:
        data = data[data["BATERIA"].isin(baterias_sel)]

    if clases_sel:
        data = data[data["CLASIFICACION"].isin(clases_sel)]

    return data


def filtrar_movimientos(df, anio, mes, baterias_sel, tipos_sel, clases_sel):
    data = df[df["ANIO"] == anio].copy()

    if mes != 0:
        data = data[data["MES"] == mes].copy()

    if baterias_sel:
        data = data[data["BATERIA"].isin(baterias_sel)]

    if tipos_sel:
        data = data[data["TIPO_SWAB"].isin(tipos_sel)]

    if clases_sel:
        data = data[data["CLASIFICACION"].isin(clases_sel)]

    return data


def resumir_pozos(data_periodo, universo_filtrado):
    if data_periodo.empty:
        salida = universo_filtrado.copy()
        salida["INTERVENCIONES"] = 0
        salida["PRCR"] = 0.0
        salida["PRAG"] = 0.0
        salida["OIL_POR_INTERV"] = 0.0
        salida["AGUA_POR_INTERV"] = 0.0
        salida["TIPO_SWAB"] = ""
        salida["UNIDADES"] = ""
        salida["PRIMERA_FECHA"] = pd.NaT
        salida["ULTIMA_FECHA"] = pd.NaT
        salida["ESTADO"] = "No intervenido"
        return salida

    resumen = (
        data_periodo
        .groupby("POZO_KEY", as_index=False)
        .agg(
            POZO_REAL=("POZO", "last"),
            BATERIA_REAL=("BATERIA", "last"),
            INTERVENCIONES=("FECHA", "count"),
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum"),
            TIPO_SWAB=("TIPO_SWAB", lambda x: ", ".join(sorted(set([v for v in x if v])))),
            UNIDADES=("UNIDAD", lambda x: ", ".join(sorted(set([v for v in x if v])))),
            PRIMERA_FECHA=("FECHA", "min"),
            ULTIMA_FECHA=("FECHA", "max")
        )
    )

    salida = universo_filtrado.merge(resumen, on="POZO_KEY", how="left")

    salida["POZO"] = salida["POZO"].fillna(salida["POZO_REAL"])
    salida["BATERIA"] = salida["BATERIA"].fillna(salida["BATERIA_REAL"])

    salida["INTERVENCIONES"] = salida["INTERVENCIONES"].fillna(0).astype(int)
    salida["PRCR"] = salida["PRCR"].fillna(0)
    salida["PRAG"] = salida["PRAG"].fillna(0)

    salida["OIL_POR_INTERV"] = np.where(
        salida["INTERVENCIONES"] > 0,
        salida["PRCR"] / salida["INTERVENCIONES"],
        0
    )

    salida["AGUA_POR_INTERV"] = np.where(
        salida["INTERVENCIONES"] > 0,
        salida["PRAG"] / salida["INTERVENCIONES"],
        0
    )

    salida["TIPO_SWAB"] = salida["TIPO_SWAB"].fillna("")
    salida["UNIDADES"] = salida["UNIDADES"].fillna("")
    salida["ESTADO"] = np.where(salida["INTERVENCIONES"] > 0, "Intervenido", "No intervenido")

    salida = salida.drop(columns=[c for c in ["POZO_REAL", "BATERIA_REAL"] if c in salida.columns])

    return salida


def resumen_baterias(resumen_pozos):
    salida = (
        resumen_pozos
        .groupby("BATERIA", as_index=False)
        .agg(
            POZOS_TOTAL=("POZO_KEY", "nunique"),
            POZOS_INTERVENIDOS=("ESTADO", lambda x: (x == "Intervenido").sum()),
            POZOS_NO_INTERVENIDOS=("ESTADO", lambda x: (x == "No intervenido").sum()),
            INTERVENCIONES=("INTERVENCIONES", "sum"),
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum")
        )
    )

    salida["OIL_POR_INTERV"] = np.where(
        salida["INTERVENCIONES"] > 0,
        salida["PRCR"] / salida["INTERVENCIONES"],
        0
    )

    return salida.sort_values("PRCR", ascending=False)


def resumen_tipo_swab(data_periodo):
    if data_periodo.empty:
        return pd.DataFrame(columns=["TIPO_SWAB", "POZOS", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"])

    salida = (
        data_periodo
        .groupby("TIPO_SWAB", as_index=False)
        .agg(
            POZOS=("POZO_KEY", "nunique"),
            INTERVENCIONES=("FECHA", "count"),
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum")
        )
    )

    salida["OIL_POR_INTERV"] = np.where(
        salida["INTERVENCIONES"] > 0,
        salida["PRCR"] / salida["INTERVENCIONES"],
        0
    )

    return salida.sort_values("PRCR", ascending=False)


def resumen_clasificacion(resumen_pozos):
    salida = (
        resumen_pozos
        .groupby("CLASIFICACION", as_index=False)
        .agg(
            POZOS_TOTAL=("POZO_KEY", "nunique"),
            POZOS_INTERVENIDOS=("ESTADO", lambda x: (x == "Intervenido").sum()),
            POZOS_NO_INTERVENIDOS=("ESTADO", lambda x: (x == "No intervenido").sum()),
            INTERVENCIONES=("INTERVENCIONES", "sum"),
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum")
        )
    )

    salida["OIL_POR_INTERV"] = np.where(
        salida["INTERVENCIONES"] > 0,
        salida["PRCR"] / salida["INTERVENCIONES"],
        0
    )

    return salida.sort_values("PRCR", ascending=False)


def tendencia_mensual(df, anio, baterias_sel, tipos_sel, clases_sel):
    data = df[df["ANIO"] == anio].copy()

    if baterias_sel:
        data = data[data["BATERIA"].isin(baterias_sel)]

    if tipos_sel:
        data = data[data["TIPO_SWAB"].isin(tipos_sel)]

    if clases_sel:
        data = data[data["CLASIFICACION"].isin(clases_sel)]

    if data.empty:
        return pd.DataFrame(columns=["MES", "MES_NOMBRE", "POZOS", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"])

    salida = (
        data
        .groupby(["MES", "MES_NOMBRE"], as_index=False)
        .agg(
            POZOS=("POZO_KEY", "nunique"),
            INTERVENCIONES=("FECHA", "count"),
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum")
        )
        .sort_values("MES")
    )

    salida["OIL_POR_INTERV"] = np.where(
        salida["INTERVENCIONES"] > 0,
        salida["PRCR"] / salida["INTERVENCIONES"],
        0
    )

    return salida


def formatear_tabla(df):
    salida = df.copy()

    for col in salida.columns:
        if pd.api.types.is_datetime64_any_dtype(salida[col]):
            salida[col] = salida[col].dt.strftime("%Y-%m-%d")

    for col in salida.select_dtypes(include=["number"]).columns:
        salida[col] = salida[col].round(2)

    return salida


# ============================================================
# EXPORTACIÓN
# ============================================================

def crear_excel_descarga(tablas):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for nombre, tabla in tablas.items():
            hoja = nombre[:31]
            tabla_out = formatear_tabla(tabla)
            tabla_out.to_excel(writer, sheet_name=hoja, index=False)

            workbook = writer.book
            worksheet = writer.sheets[hoja]

            formato_header = workbook.add_format({
                "bold": True,
                "bg_color": "#D9EAF7",
                "border": 1
            })

            for col_num, value in enumerate(tabla_out.columns):
                worksheet.write(0, col_num, value, formato_header)
                worksheet.set_column(col_num, col_num, 18)

    buffer.seek(0)
    return buffer.getvalue()


def agregar_titulo(slide, titulo, subtitulo):
    box = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(12.4), Inches(0.5))
    tf = box.text_frame
    tf.text = titulo
    tf.paragraphs[0].font.size = Pt(22)
    tf.paragraphs[0].font.bold = True

    box2 = slide.shapes.add_textbox(Inches(0.4), Inches(0.75), Inches(12.4), Inches(0.35))
    tf2 = box2.text_frame
    tf2.text = subtitulo
    tf2.paragraphs[0].font.size = Pt(11)


def agregar_tabla(slide, df, x, y, w, h, font_size=7):
    if df.empty:
        box = slide.shapes.add_textbox(x, y, w, h)
        box.text_frame.text = "Sin datos"
        return

    tabla = formatear_tabla(df).head(15)
    rows = len(tabla) + 1
    cols = len(tabla.columns)

    table_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    table = table_shape.table

    for j, col in enumerate(tabla.columns):
        cell = table.cell(0, j)
        cell.text = str(col)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(font_size)

    for i, (_, row) in enumerate(tabla.iterrows(), start=1):
        for j, col in enumerate(tabla.columns):
            cell = table.cell(i, j)
            cell.text = "" if pd.isna(row[col]) else str(row[col])
            cell.text_frame.paragraphs[0].font.size = Pt(font_size)


def agregar_barras(slide, titulo, categorias, valores, x, y, w, h, nombre_serie):
    categorias = list(categorias)
    valores = [0 if pd.isna(v) else float(v) for v in valores]

    if len(categorias) == 0:
        return

    chart_data = CategoryChartData()
    chart_data.categories = categorias
    chart_data.add_series(nombre_serie, valores)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        x, y, w, h,
        chart_data
    ).chart

    chart.has_legend = False
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = titulo


def agregar_linea(slide, titulo, categorias, serie_prcr, serie_prag, x, y, w, h):
    categorias = list(categorias)

    if len(categorias) == 0:
        return

    chart_data = CategoryChartData()
    chart_data.categories = categorias
    chart_data.add_series("PRCR", [0 if pd.isna(v) else float(v) for v in serie_prcr])
    chart_data.add_series("PRAG", [0 if pd.isna(v) else float(v) for v in serie_prag])

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        x, y, w, h,
        chart_data
    ).chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = titulo


def crear_ppt(periodo, kpis, resumen_pozos, resumen_bateria, resumen_tipo, resumen_clase, tendencia):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank = prs.slide_layouts[6]

    # Slide 1
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Dashboard operativo SWAB Lote X", periodo)

    kpi_df = pd.DataFrame({
        "Indicador": [
            "Pozos universo",
            "Pozos intervenidos",
            "Pozos no intervenidos",
            "Intervenciones",
            "PRCR",
            "PRAG",
            "Oil por intervención"
        ],
        "Valor": [
            kpis["pozos_universo"],
            kpis["pozos_intervenidos"],
            kpis["pozos_no_intervenidos"],
            kpis["intervenciones"],
            round(kpis["prcr"], 2),
            round(kpis["prag"], 2),
            round(kpis["oil_interv"], 2)
        ]
    })

    agregar_tabla(slide, kpi_df, Inches(0.5), Inches(1.25), Inches(5.2), Inches(3.0), font_size=10)

    estado_df = pd.DataFrame({
        "Estado": ["Intervenido", "No intervenido"],
        "Pozos": [kpis["pozos_intervenidos"], kpis["pozos_no_intervenidos"]]
    })

    agregar_barras(
        slide,
        "Pozos intervenidos vs no intervenidos",
        estado_df["Estado"],
        estado_df["Pozos"],
        Inches(6.2),
        Inches(1.2),
        Inches(6.5),
        Inches(4.8),
        "Pozos"
    )

    # Slide 2
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Clasificación de pozos", periodo)

    clase_plot = resumen_clase.sort_values("PRCR", ascending=False)
    agregar_barras(
        slide,
        "PRCR por clasificación",
        clase_plot["CLASIFICACION"],
        clase_plot["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(6.4),
        Inches(5.4),
        "PRCR"
    )

    agregar_tabla(
        slide,
        resumen_clase[["CLASIFICACION", "POZOS_TOTAL", "POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES", "PRCR", "PRAG"]],
        Inches(7.1),
        Inches(1.2),
        Inches(5.7),
        Inches(5.4),
        font_size=7
    )

    # Slide 3
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Top pozos por producción PRCR", periodo)

    top_pozos = resumen_pozos[resumen_pozos["INTERVENCIONES"] > 0].sort_values("PRCR", ascending=False).head(15)

    agregar_barras(
        slide,
        "Top 15 pozos por PRCR",
        top_pozos["POZO"],
        top_pozos["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(7.0),
        Inches(5.4),
        "PRCR"
    )

    agregar_tabla(
        slide,
        top_pozos[["POZO", "BATERIA", "CLASIFICACION", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"]],
        Inches(7.7),
        Inches(1.2),
        Inches(5.2),
        Inches(5.4),
        font_size=7
    )

    # Slide 4
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Top baterías", periodo)

    top_bat = resumen_bateria.sort_values("PRCR", ascending=False).head(15)

    agregar_barras(
        slide,
        "Top 15 baterías por PRCR",
        top_bat["BATERIA"],
        top_bat["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(7.0),
        Inches(5.4),
        "PRCR"
    )

    agregar_tabla(
        slide,
        top_bat[["BATERIA", "POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES", "PRCR", "PRAG"]],
        Inches(7.7),
        Inches(1.2),
        Inches(5.2),
        Inches(5.4),
        font_size=7
    )

    # Slide 5
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "TS y CS", periodo)

    agregar_barras(
        slide,
        "PRCR por tipo de swab",
        resumen_tipo["TIPO_SWAB"],
        resumen_tipo["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(6.5),
        Inches(5.4),
        "PRCR"
    )

    agregar_tabla(
        slide,
        resumen_tipo[["TIPO_SWAB", "POZOS", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"]],
        Inches(7.2),
        Inches(1.2),
        Inches(5.6),
        Inches(4.8),
        font_size=8
    )

    # Slide 6
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Tendencia mensual", periodo)

    if not tendencia.empty:
        agregar_linea(
            slide,
            "PRCR y PRAG por mes",
            tendencia["MES_NOMBRE"],
            tendencia["PRCR"],
            tendencia["PRAG"],
            Inches(0.7),
            Inches(1.2),
            Inches(12.0),
            Inches(5.5)
        )

    # Slide 7
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Pozos no intervenidos", periodo)

    no_interv = resumen_pozos[resumen_pozos["ESTADO"] == "No intervenido"].copy()
    no_interv = no_interv.sort_values(["CLASIFICACION", "BATERIA", "POZO"]).head(20)

    agregar_tabla(
        slide,
        no_interv[["POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION", "ESTADO", "ULTIMA_FECHA_HISTORICA"]],
        Inches(0.5),
        Inches(1.2),
        Inches(12.3),
        Inches(5.8),
        font_size=8
    )

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

st.title("🛢️ Dashboard operativo SWAB Lote X")
st.caption("Sube solo el Excel principal. El sistema clasificará convertidos 2024, 2025, 2026 y el resto como Básica.")

archivo = st.file_uploader(
    "Sube tu Excel principal, por ejemplo Data Swab Python.xlsx",
    type=["xlsx"]
)

if archivo is None:
    st.info("Sube el Excel para habilitar los filtros.")
    st.stop()

bytes_excel = archivo.getvalue()

try:
    df, hoja_usada = cargar_excel_swab(bytes_excel)
except Exception as e:
    st.error(f"No se pudo cargar el Excel: {e}")
    st.stop()

universo = construir_universo(df)

fecha_min = df["FECHA"].min().date()
fecha_max = df["FECHA"].max().date()

anios = sorted(df["ANIO"].dropna().unique().astype(int).tolist())
baterias = sorted([x for x in universo["BATERIA"].dropna().unique().tolist() if x != ""])
tipos = sorted([x for x in df["TIPO_SWAB"].dropna().unique().tolist() if x != ""])
clases = ["Básica", "Convertido 2024", "Convertido 2025", "Convertido 2026"]

with st.sidebar:
    st.header("Filtros")

    anio_sel = st.selectbox(
        "Año",
        anios,
        index=len(anios) - 1
    )

    meses_opciones = [0] + list(range(1, 13))
    mes_default = fecha_max.month if anio_sel == fecha_max.year else 0

    mes_sel = st.selectbox(
        "Mes",
        meses_opciones,
        index=meses_opciones.index(mes_default),
        format_func=lambda x: "Todo el año" if x == 0 else MESES[x]
    )

    baterias_sel = st.multiselect(
        "Batería",
        baterias,
        default=[]
    )

    tipos_sel = st.multiselect(
        "Tipo de swab",
        tipos,
        default=[]
    )

    clases_sel = st.multiselect(
        "Tipo de pozo",
        clases,
        default=[]
    )

    top_n = st.slider(
        "Top para gráficos",
        min_value=5,
        max_value=50,
        value=20,
        step=5
    )

    ejecutar = st.button("Ejecutar análisis", type="primary")

if not ejecutar and "swab_resultados" not in st.session_state:
    st.warning("Selecciona los filtros en la izquierda y luego presiona Ejecutar análisis.")
    st.stop()

if ejecutar:
    universo_filtrado = aplicar_filtros_universo(universo, baterias_sel, clases_sel)
    data_periodo = filtrar_movimientos(df, anio_sel, mes_sel, baterias_sel, tipos_sel, clases_sel)

    resumen_pozos = resumir_pozos(data_periodo, universo_filtrado)
    resumen_bateria = resumen_baterias(resumen_pozos)
    resumen_tipo = resumen_tipo_swab(data_periodo)
    resumen_clase = resumen_clasificacion(resumen_pozos)
    tendencia = tendencia_mensual(df, anio_sel, baterias_sel, tipos_sel, clases_sel)

    st.session_state["swab_resultados"] = {
        "anio": anio_sel,
        "mes": mes_sel,
        "baterias_sel": baterias_sel,
        "tipos_sel": tipos_sel,
        "clases_sel": clases_sel,
        "data_periodo": data_periodo,
        "resumen_pozos": resumen_pozos,
        "resumen_bateria": resumen_bateria,
        "resumen_tipo": resumen_tipo,
        "resumen_clase": resumen_clase,
        "tendencia": tendencia
    }

res = st.session_state["swab_resultados"]

anio_sel = res["anio"]
mes_sel = res["mes"]
data_periodo = res["data_periodo"]
resumen_pozos = res["resumen_pozos"]
resumen_bateria = res["resumen_bateria"]
resumen_tipo = res["resumen_tipo"]
resumen_clase = res["resumen_clase"]
tendencia = res["tendencia"]

periodo = periodo_texto(anio_sel, mes_sel)

pozos_universo = resumen_pozos["POZO_KEY"].nunique()
pozos_intervenidos = int((resumen_pozos["ESTADO"] == "Intervenido").sum())
pozos_no_intervenidos = int((resumen_pozos["ESTADO"] == "No intervenido").sum())
intervenciones = int(resumen_pozos["INTERVENCIONES"].sum())
prcr_total = float(resumen_pozos["PRCR"].sum())
prag_total = float(resumen_pozos["PRAG"].sum())
oil_interv = prcr_total / intervenciones if intervenciones > 0 else 0

kpis = {
    "pozos_universo": pozos_universo,
    "pozos_intervenidos": pozos_intervenidos,
    "pozos_no_intervenidos": pozos_no_intervenidos,
    "intervenciones": intervenciones,
    "prcr": prcr_total,
    "prag": prag_total,
    "oil_interv": oil_interv
}

st.subheader(f"Estatus operativo: {periodo}")

k1, k2, k3, k4, k5, k6 = st.columns(6)

with k1:
    st.metric("Pozos universo", f"{pozos_universo:,}")

with k2:
    st.metric("Pozos intervenidos", f"{pozos_intervenidos:,}")

with k3:
    st.metric("Pozos no intervenidos", f"{pozos_no_intervenidos:,}")

with k4:
    st.metric("Intervenciones", f"{intervenciones:,}")

with k5:
    st.metric("PRCR", f"{prcr_total:,.2f}")

with k6:
    st.metric("Oil / Interv.", f"{oil_interv:,.2f}")

st.caption(
    f"Hoja usada: {hoja_usada}. Rango de datos cargado: {fecha_min} al {fecha_max}. "
    f"PRCR = petróleo recuperado. PRAG = agua recuperada."
)

st.divider()

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Estatus y listados",
    "Performance por pozo",
    "Baterías",
    "TS y CS",
    "Tendencia mensual",
    "Descargas"
])


columnas_pozos = [
    "ESTADO",
    "POZO",
    "BATERIA",
    "CLASIFICACION",
    "ANIO_CONVERSION",
    "TIPO_SWAB",
    "UNIDADES",
    "INTERVENCIONES",
    "PRCR",
    "PRAG",
    "OIL_POR_INTERV",
    "AGUA_POR_INTERV",
    "PRIMERA_FECHA",
    "ULTIMA_FECHA",
    "ULTIMA_FECHA_HISTORICA"
]

columnas_pozos = [c for c in columnas_pozos if c in resumen_pozos.columns]


with tab1:
    st.subheader("Pozos intervenidos y no intervenidos")

    estado_sel = st.radio(
        "Estado",
        ["Todos", "Intervenido", "No intervenido"],
        horizontal=True
    )

    tabla = resumen_pozos.copy()

    if estado_sel != "Todos":
        tabla = tabla[tabla["ESTADO"] == estado_sel]

    tabla = tabla.sort_values(["ESTADO", "CLASIFICACION", "BATERIA", "POZO"])

    st.dataframe(
        formatear_tabla(tabla[columnas_pozos]),
        use_container_width=True,
        hide_index=True
    )

    col_a, col_b = st.columns(2)

    with col_a:
        estado_df = pd.DataFrame({
            "Estado": ["Intervenido", "No intervenido"],
            "Pozos": [pozos_intervenidos, pozos_no_intervenidos]
        })

        fig_estado = px.bar(
            estado_df,
            x="Estado",
            y="Pozos",
            text="Pozos",
            title="Pozos intervenidos vs no intervenidos"
        )
        st.plotly_chart(fig_estado, use_container_width=True)

    with col_b:
        fig_clase = px.bar(
            resumen_clase,
            x="CLASIFICACION",
            y="PRCR",
            text="PRCR",
            title="PRCR por clasificación"
        )
        st.plotly_chart(fig_clase, use_container_width=True)

    st.write("Resumen por clasificación")
    st.dataframe(
        formatear_tabla(resumen_clase),
        use_container_width=True,
        hide_index=True
    )


with tab2:
    st.subheader("Performance por pozo")

    intervenidos = resumen_pozos[resumen_pozos["INTERVENCIONES"] > 0].copy()

    if intervenidos.empty:
        st.info("No hay pozos intervenidos para el periodo seleccionado.")
    else:
        top_prcr = intervenidos.sort_values("PRCR", ascending=False).head(top_n)
        fig_top = px.bar(
            top_prcr,
            x="POZO",
            y="PRCR",
            color="CLASIFICACION",
            hover_data=["BATERIA", "INTERVENCIONES", "PRAG", "OIL_POR_INTERV", "TIPO_SWAB"],
            title=f"Top {top_n} pozos por PRCR"
        )
        st.plotly_chart(fig_top, use_container_width=True)

        top_oil_interv = intervenidos.sort_values("OIL_POR_INTERV", ascending=False).head(top_n)
        fig_oil = px.bar(
            top_oil_interv,
            x="POZO",
            y="OIL_POR_INTERV",
            color="CLASIFICACION",
            hover_data=["BATERIA", "INTERVENCIONES", "PRCR", "PRAG", "TIPO_SWAB"],
            title=f"Top {top_n} pozos por oil/intervención"
        )
        st.plotly_chart(fig_oil, use_container_width=True)

        st.dataframe(
            formatear_tabla(intervenidos[columnas_pozos].sort_values("PRCR", ascending=False)),
            use_container_width=True,
            hide_index=True
        )


with tab3:
    st.subheader("Análisis por batería")

    if resumen_bateria.empty:
        st.info("No hay información de batería para el filtro seleccionado.")
    else:
        c1, c2 = st.columns(2)

        with c1:
            top_bat_prcr = resumen_bateria.sort_values("PRCR", ascending=False).head(top_n)
            fig_bat_prcr = px.bar(
                top_bat_prcr,
                x="BATERIA",
                y="PRCR",
                hover_data=["POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES", "PRAG"],
                title=f"Top {top_n} baterías por PRCR"
            )
            st.plotly_chart(fig_bat_prcr, use_container_width=True)

        with c2:
            top_bat_interv = resumen_bateria.sort_values("INTERVENCIONES", ascending=False).head(top_n)
            fig_bat_interv = px.bar(
                top_bat_interv,
                x="BATERIA",
                y="INTERVENCIONES",
                hover_data=["POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "PRCR", "PRAG"],
                title=f"Top {top_n} baterías por intervenciones"
            )
            st.plotly_chart(fig_bat_interv, use_container_width=True)

        st.dataframe(
            formatear_tabla(resumen_bateria),
            use_container_width=True,
            hide_index=True
        )


with tab4:
    st.subheader("Análisis por TS y CS")

    if resumen_tipo.empty:
        st.info("No hay información por tipo de swab para el filtro seleccionado.")
    else:
        c1, c2 = st.columns(2)

        with c1:
            fig_tipo_prcr = px.bar(
                resumen_tipo,
                x="TIPO_SWAB",
                y="PRCR",
                text="PRCR",
                hover_data=["POZOS", "INTERVENCIONES", "PRAG", "OIL_POR_INTERV"],
                title="PRCR por tipo de swab"
            )
            st.plotly_chart(fig_tipo_prcr, use_container_width=True)

        with c2:
            fig_tipo_interv = px.bar(
                resumen_tipo,
                x="TIPO_SWAB",
                y="INTERVENCIONES",
                text="INTERVENCIONES",
                hover_data=["POZOS", "PRCR", "PRAG", "OIL_POR_INTERV"],
                title="Intervenciones por tipo de swab"
            )
            st.plotly_chart(fig_tipo_interv, use_container_width=True)

        st.dataframe(
            formatear_tabla(resumen_tipo),
            use_container_width=True,
            hide_index=True
        )


with tab5:
    st.subheader(f"Tendencia mensual {anio_sel}")

    if tendencia.empty:
        st.info("No hay tendencia mensual para los filtros seleccionados.")
    else:
        orden_meses = [MESES[m] for m in sorted(tendencia["MES"].unique())]

        fig_tend = px.line(
            tendencia,
            x="MES_NOMBRE",
            y=["PRCR", "PRAG"],
            markers=True,
            category_orders={"MES_NOMBRE": orden_meses},
            title=f"Tendencia mensual PRCR y PRAG en {anio_sel}"
        )
        st.plotly_chart(fig_tend, use_container_width=True)

        fig_interv = px.bar(
            tendencia,
            x="MES_NOMBRE",
            y="INTERVENCIONES",
            text="INTERVENCIONES",
            category_orders={"MES_NOMBRE": orden_meses},
            hover_data=["POZOS", "PRCR", "PRAG", "OIL_POR_INTERV"],
            title=f"Intervenciones mensuales en {anio_sel}"
        )
        st.plotly_chart(fig_interv, use_container_width=True)

        st.dataframe(
            formatear_tabla(tendencia),
            use_container_width=True,
            hide_index=True
        )


with tab6:
    st.subheader("Descargas")

    tablas = {
        "Resumen pozos": resumen_pozos[columnas_pozos],
        "Intervenidos": resumen_pozos[resumen_pozos["ESTADO"] == "Intervenido"][columnas_pozos],
        "No intervenidos": resumen_pozos[resumen_pozos["ESTADO"] == "No intervenido"][columnas_pozos],
        "Resumen baterias": resumen_bateria,
        "Resumen TS CS": resumen_tipo,
        "Resumen clasificacion": resumen_clase,
        "Tendencia mensual": tendencia
    }

    excel_bytes = crear_excel_descarga(tablas)

    st.download_button(
        "Descargar toda la información en Excel",
        data=excel_bytes,
        file_name=f"resultado_swab_{anio_sel}_{mes_sel}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    ppt_bytes = crear_ppt(
        periodo=periodo,
        kpis=kpis,
        resumen_pozos=resumen_pozos,
        resumen_bateria=resumen_bateria,
        resumen_tipo=resumen_tipo,
        resumen_clase=resumen_clase,
        tendencia=tendencia
    )

    st.download_button(
        "Descargar PPT editable",
        data=ppt_bytes,
        file_name=f"dashboard_swab_{anio_sel}_{mes_sel}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    st.subheader("Validación")

    validacion = pd.DataFrame({
        "Concepto": [
            "Hoja usada",
            "Fecha mínima",
            "Fecha máxima",
            "Registros cargados",
            "Pozos históricos",
            "Baterías históricas",
            "Convertidos 2024 en data",
            "Convertidos 2025 en data",
            "Convertidos 2026 en data",
            "Pozos básicos en data"
        ],
        "Valor": [
            hoja_usada,
            str(fecha_min),
            str(fecha_max),
            f"{len(df):,}",
            f"{df['POZO_KEY'].nunique():,}",
            f"{df['BATERIA'].nunique():,}",
            f"{(universo['CLASIFICACION'] == 'Convertido 2024').sum():,}",
            f"{(universo['CLASIFICACION'] == 'Convertido 2025').sum():,}",
            f"{(universo['CLASIFICACION'] == 'Convertido 2026').sum():,}",
            f"{(universo['CLASIFICACION'] == 'Básica').sum():,}"
        ]
    })

    st.dataframe(validacion, use_container_width=True, hide_index=True)

    st.write("Muestra de datos limpios")
    st.dataframe(
        formatear_tabla(df[["FECHA", "POZO", "BATERIA", "UNIDAD", "TIPO_SWAB", "PRCR", "PRAG", "CLASIFICACION"]].head(50)),
        use_container_width=True,
        hide_index=True
    )
