import io
import re
import calendar

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt


st.set_page_config(
    page_title="SWAB Lote X - Servicio y Potencial",
    page_icon="🛢️",
    layout="wide"
)


# ============================================================
# POZOS CONVERTIDOS
# Todo pozo fuera de estas listas se clasifica como Básica.
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

# ============================================================
# POZOS CANDIDATOS PARA ATA
# Estos 88 pozos se excluyen de todo el análisis SWAB.
# No se consideran en producción, potencial, gráficos, tablas ni descargas.
# ============================================================

CANDIDATOS_ATA = [
    'AA37', 'AA54', 'AA76', 'AA112',
    'AA1577', 'AA1598', 'AA1599', 'AA1633',
    'AA1661', 'AA1847', 'AA1930', 'AA5631',
    'AA5707', 'AA5861', 'AA5926', 'AA5971',
    'AA6192', 'AA6338', 'AA6342', 'AA6372',
    'AA6423', 'AA6454', 'AA6517', 'AA6646',
    'AA6762', 'AA7201', 'AA9154', 'AA9329',
    'AA9364', 'AA10013', 'EA216', 'EA264',
    'EA364', 'EA440', 'EA741', 'EA771',
    'EA876', 'EA888', 'EA987', 'EA1054',
    'EA1081', 'EA1161', 'EA1167', 'EA1233',
    'EA1302', 'EA1506', 'EA1511', 'EA1513',
    'EA1581', 'EA1630', 'EA1885', 'EA2067',
    'EA2249', 'EA2254', 'EA2256', 'EA2304',
    'EA2372', 'EA2389', 'EA2403', 'EA5682D',
    'EA5694', 'EA5739', 'EA5766', 'EA5868',
    'EA5874', 'EA5914', 'EA5921', 'EA5957',
    'EA6130', 'EA6237', 'EA6918', 'EA7027',
    'EA7158', 'EA8574', 'EA9242', 'EA9251',
    'EA9287', 'EA9409', 'EA9417', 'EA9491',
    'EA9668', 'EA9752', 'EA9779', 'EA11128',
    'PB47', 'PB232', 'PE171', 'PT4-3',
]


MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

PLOT_TEMPLATE = "plotly_white"

# Etiquetas visibles para tablas, gráficos, Excel y PPT.
# Internamente se mantiene PRCR y PRAG porque así vienen en el Excel.
LABELS_COLUMNAS = {
    "PRCR": "Producción de petróleo",
    "PRAG": "Producción de agua",
    "PRCR_BASE": "Producción de petróleo base",
    "PRCR_ACTUAL": "Producción de petróleo actual",
    "PRCR_BASE_NO_REALIZADO": "Producción de petróleo base no realizada",
    "INTERV_MES_ANTERIOR_OBJETIVO": "Intervenciones mes anterior",
    "PRCR_MES_ANTERIOR_OBJETIVO": "Producción de petróleo mes anterior",
    "PRAG_MES_ANTERIOR_OBJETIVO": "Producción de agua mes anterior",
    "VAR_PETROLEO_VS_MES_ANTERIOR": "Variación petróleo vs mes anterior",
    "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR": "Producción de petróleo dejada vs mes anterior",
    "INTERVENCIONES_DEJADAS_MES_ANTERIOR": "Intervenciones dejadas vs mes anterior",
    "PRAG_BASE": "Producción de agua base",
    "PRAG_ACTUAL": "Producción de agua actual",
    "PRAG_BASE_NO_REALIZADO": "Producción de agua base no realizada",
    "VAR_PRCR": "Variación producción de petróleo",
    "VAR_PRAG": "Variación producción de agua",
    "PRCR_ULTIMO_MES": "Producción de petróleo último mes",
    "PRAG_ULTIMO_MES": "Producción de agua último mes",
    "ULTIMA_FECHA_CON_PRCR": "Última fecha con producción de petróleo",
    "OIL_POR_INTERV": "Petróleo por intervención",
    "OIL_INTERV_BASE": "Petróleo por intervención base",
    "OIL_INTERV_ACTUAL": "Petróleo por intervención actual",
    "OIL_POR_INTERV_ULTIMO_MES": "Petróleo por intervención último mes",
    "POTENCIAL_BOPD": "Potencial estimado bopd",
    "POTENCIAL_ULTIMO_MES_BOPD": "Potencial último mes activo bopd",
    "MES_ANTERIOR_ULTIMO_ACTIVO": "Mes anterior al último mes activo",
    "DIAS_MES_ANTERIOR": "Días mes anterior",
    "PRCR_MES_ANTERIOR": "Producción de petróleo mes anterior",
    "PRAG_MES_ANTERIOR": "Producción de agua mes anterior",
    "INTERV_MES_ANTERIOR": "Intervenciones mes anterior",
    "POTENCIAL_MES_ANTERIOR_BOPD": "Potencial mes anterior bopd",
    "POTENCIAL_PROMEDIO_2_MESES_BOPD": "Potencial promedio 2 meses bopd",
    "INTERV_BASE": "Intervenciones base promedio",
    "INTERV_ACTUAL": "Intervenciones mes objetivo",
    "INTERV_BASE_NO_REALIZADAS": "Intervenciones base no realizadas",
    "PRIORIDAD_REVISION": "Prioridad de revisión",
    "ESTADO_DESPLAZAMIENTO": "Estado de desplazamiento",
    "ANIO_CONVERSION": "Año conversión",
    "ULTIMO_MES_ACTIVO": "Último mes activo",
    "DIAS_MES": "Días del mes",
    "INTERV_ULTIMO_MES": "Intervenciones último mes",
    "ULTIMA_FECHA_HISTORICA": "Última fecha histórica",
    "POZOS_INTERVENIDOS": "Pozos intervenidos",
    "POZOS_NO_INTERVENIDOS": "Pozos no intervenidos",
    "POZOS_TOTAL": "Pozos total",
    "POZOS_AFECTADOS": "Pozos afectados",
    "POZOS_DEJADOS": "Pozos dejados",
    "POZOS_REDUCIDOS": "Pozos reducidos",
    "TIPO_SWAB": "Tipo de swab",
    "CLASIFICACION": "Clasificación",
    "BATERIA": "Batería",
    "POZO": "Pozo",
    "UNIDADES": "Unidades",
    "INTERVENCIONES": "Intervenciones",
    "MES_NOMBRE": "Mes"
}


def etiqueta(campo):
    return LABELS_COLUMNAS.get(campo, campo)


# ============================================================
# UTILIDADES
# ============================================================

def limpiar_pozo(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = re.sub(r"\s+", "", texto)
    return texto


def candidatos_ata_keys():
    return {limpiar_pozo(pozo) for pozo in CANDIDATOS_ATA}


def excluir_candidatos_ata(df):
    """
    Excluye del análisis los 88 pozos candidatos para ATA.
    Retorna la data filtrada y la data excluida para validación.
    """
    keys_ata = candidatos_ata_keys()
    excluidos = df[df["POZO_KEY"].isin(keys_ata)].copy()
    filtrado = df[~df["POZO_KEY"].isin(keys_ata)].copy()
    return filtrado, excluidos


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
    reemplazos = {
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "á": "A", "é": "E", "í": "I", "ó": "O", "ú": "U"
    }
    for k, v in reemplazos.items():
        texto = texto.replace(k, v)
    texto = re.sub(r"\s+", "_", texto)
    return texto


def convertir_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def clasificar_pozo(pozo_key):
    for anio, pozos in CONVERTIDOS.items():
        if pozo_key in [limpiar_pozo(p) for p in pozos]:
            return f"Convertido {anio}"
    return "Básica"


def obtener_anio_conversion(pozo_key):
    for anio, pozos in CONVERTIDOS.items():
        if pozo_key in [limpiar_pozo(p) for p in pozos]:
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


def periodo_mes_texto(anio, mes):
    return f"{MESES[int(mes)]} {int(anio)}"


def primer_dia_mes(anio, mes):
    return pd.Timestamp(int(anio), int(mes), 1)


def ultimo_dia_mes(anio, mes):
    return primer_dia_mes(anio, mes) + pd.offsets.MonthEnd(0)


def dias_calendario_mes(anio, mes):
    return calendar.monthrange(int(anio), int(mes))[1]


def mes_anterior(anio, mes, n=1):
    fecha = pd.Timestamp(int(anio), int(mes), 1) - pd.DateOffset(months=n)
    return int(fecha.year), int(fecha.month)


def lista_meses_previos(anio, mes, n_meses):
    meses = []
    for i in range(1, int(n_meses) + 1):
        a, m = mes_anterior(anio, mes, i)
        meses.append((a, m))
    return list(reversed(meses))


def formatear_tabla(df):
    salida = df.copy()

    for col in salida.columns:
        if pd.api.types.is_datetime64_any_dtype(salida[col]):
            salida[col] = salida[col].dt.strftime("%Y-%m-%d")

    for col in salida.select_dtypes(include=["number"]).columns:
        salida[col] = salida[col].round(2)

    salida = salida.rename(columns=LABELS_COLUMNAS)
    return salida


def convertir_excel(tablas):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for nombre, tabla in tablas.items():
            hoja = nombre[:31]
            tabla_out = formatear_tabla(tabla)
            tabla_out.to_excel(writer, sheet_name=hoja, index=False)

            workbook = writer.book
            worksheet = writer.sheets[hoja]
            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#D9EAF7",
                "border": 1
            })

            for col_num, value in enumerate(tabla_out.columns):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 18)

    buffer.seek(0)
    return buffer.getvalue()


def es_serie_numerica(valores):
    """
    Valida si una secuencia de valores es numérica.
    Se usa para aplicar formato de 2 decimales solo cuando corresponde.
    """
    if valores is None:
        return False

    try:
        serie = pd.Series(list(valores)).dropna()
    except Exception:
        return False

    if serie.empty:
        return False

    convertida = pd.to_numeric(serie, errors="coerce")
    return convertida.notna().mean() >= 0.80


def aplicar_formato_2_decimales_fig(fig):
    """
    Aplica 2 decimales a textos, ejes y hover de los gráficos.
    No cambia los cálculos, solo la visualización.
    """
    hay_y_numerico = False
    hay_x_numerico = False

    for trace in fig.data:
        x_numerico = es_serie_numerica(getattr(trace, "x", None))
        y_numerico = es_serie_numerica(getattr(trace, "y", None))

        hay_x_numerico = hay_x_numerico or x_numerico
        hay_y_numerico = hay_y_numerico or y_numerico

        # Barras con etiquetas encima.
        if getattr(trace, "type", "") == "bar" and getattr(trace, "text", None) is not None:
            if es_serie_numerica(trace.text):
                trace.texttemplate = "%{text:,.2f}"
                trace.textposition = "outside"
                trace.cliponaxis = False

        # Pie o donut.
        if getattr(trace, "type", "") == "pie":
            trace.texttemplate = "%{percent:.2%}"
            trace.hovertemplate = "%{label}<br>Valor: %{value:,.2f}<br>Porcentaje: %{percent:.2%}<extra></extra>"

        # Hover de ejes principales.
        ht = getattr(trace, "hovertemplate", None)
        if ht:
            if y_numerico:
                ht = ht.replace("%{y}", "%{y:,.2f}")
            if x_numerico:
                ht = ht.replace("%{x}", "%{x:,.2f}")

            # Hover de customdata. Plotly Express guarda ahí columnas adicionales.
            customdata = getattr(trace, "customdata", None)
            if customdata is not None:
                try:
                    matriz = np.array(customdata)
                    if matriz.ndim == 2:
                        for i in range(matriz.shape[1]):
                            if es_serie_numerica(matriz[:, i]):
                                ht = ht.replace(f"%{{customdata[{i}]}}", f"%{{customdata[{i}]:,.2f}}")
                except Exception:
                    pass

            trace.hovertemplate = ht

    if hay_y_numerico:
        fig.update_yaxes(tickformat=",.2f")

    if hay_x_numerico:
        fig.update_xaxes(tickformat=",.2f")

    return fig


def aplicar_layout_fig(fig, titulo=None, altura=480):
    fig.update_layout(
        template=PLOT_TEMPLATE,
        height=altura,
        title=titulo,
        margin=dict(l=20, r=20, t=65, b=40),
        legend_title_text="",
        hovermode="closest"
    )

    # Cambia nombres técnicos de columnas por etiquetas gerenciales.
    for trace in fig.data:
        if getattr(trace, "name", None) in LABELS_COLUMNAS:
            trace.name = LABELS_COLUMNAS[trace.name]
            trace.legendgroup = trace.name
        if getattr(trace, "hovertemplate", None):
            ht = trace.hovertemplate
            for raw, nice in LABELS_COLUMNAS.items():
                ht = ht.replace(raw, nice)
            trace.hovertemplate = ht

    for axis_name in ["xaxis", "yaxis"]:
        axis = getattr(fig.layout, axis_name, None)
        if axis and axis.title and axis.title.text in LABELS_COLUMNAS:
            axis.title.text = LABELS_COLUMNAS[axis.title.text]

    fig = aplicar_formato_2_decimales_fig(fig)

    return fig


# ============================================================
# CARGA DE DATA
# ============================================================

@st.cache_data(show_spinner=False)
def cargar_excel_swab(bytes_excel):
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))
    hoja = "Datos de Swab" if "Datos de Swab" in xls.sheet_names else xls.sheet_names[0]

    df = pd.read_excel(io.BytesIO(bytes_excel), sheet_name=hoja)
    df.columns = [normalizar_columna(c) for c in df.columns]

    requeridas = ["FECHA", "COD_POZ", "COD_BAT", "UNIDAD", "TSER", "PRCR", "PRAG"]
    faltantes = [c for c in requeridas if c not in df.columns]

    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias: "
            + ", ".join(faltantes)
            + ". El Excel debe tener FECHA, COD_POZ, COD_BAT, UNIDAD, TSER, PRCR y PRAG. PRCR se usa como producción de petróleo y PRAG como producción de agua."
        )

    df = df[requeridas].copy()
    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df = df[df["FECHA"].notna()].copy()

    df["POZO"] = df["COD_POZ"].apply(mostrar_pozo)
    df["POZO_KEY"] = df["COD_POZ"].apply(limpiar_pozo)
    df["BATERIA"] = df["COD_BAT"].apply(limpiar_texto)
    df["UNIDAD"] = df["UNIDAD"].apply(limpiar_texto)
    df["TIPO_SWAB"] = df["TSER"].apply(normalizar_tipo_swab)
    df["PRCR"] = convertir_numero(df["PRCR"])  # petróleo recuperado
    df["PRAG"] = convertir_numero(df["PRAG"])  # agua recuperada

    df = df[df["POZO_KEY"] != ""].copy()

    df["ANIO"] = df["FECHA"].dt.year
    df["MES"] = df["FECHA"].dt.month
    df["MES_NOMBRE"] = df["MES"].map(MESES)
    df["PERIODO_MES"] = df["FECHA"].dt.to_period("M").dt.to_timestamp()

    df["CLASIFICACION"] = df["POZO_KEY"].apply(clasificar_pozo)
    df["ANIO_CONVERSION"] = df["POZO_KEY"].apply(obtener_anio_conversion)

    return df, hoja


def construir_convertidos_df():
    registros = []

    for anio, pozos in CONVERTIDOS.items():
        for pozo in pozos:
            key = limpiar_pozo(pozo)
            if key:
                registros.append({
                    "POZO_KEY": key,
                    "POZO_CONVERTIDO": mostrar_pozo(pozo),
                    "CLASIFICACION_CONVERTIDO": f"Convertido {anio}",
                    "ANIO_CONVERSION_CONVERTIDO": anio
                })

    return pd.DataFrame(registros).drop_duplicates("POZO_KEY")


def construir_universo(df):
    historico = (
        df.sort_values("FECHA")
        .groupby("POZO_KEY", as_index=False)
        .agg(
            POZO_HIST=("POZO", "last"),
            BATERIA_HIST=("BATERIA", "last"),
            ULTIMA_FECHA_HISTORICA=("FECHA", "max")
        )
    )

    convertidos = construir_convertidos_df()
    universo = historico.merge(convertidos, on="POZO_KEY", how="outer")

    universo["POZO"] = universo["POZO_HIST"]
    universo.loc[
        universo["POZO"].isna() | (universo["POZO"] == ""),
        "POZO"
    ] = universo["POZO_CONVERTIDO"]

    universo["BATERIA"] = universo["BATERIA_HIST"].fillna("SIN BATERIA")
    universo["CLASIFICACION"] = universo["CLASIFICACION_CONVERTIDO"].fillna("Básica")
    universo["ANIO_CONVERSION"] = universo["ANIO_CONVERSION_CONVERTIDO"].fillna(0).astype(int)

    universo = universo[[
        "POZO_KEY", "POZO", "BATERIA", "CLASIFICACION",
        "ANIO_CONVERSION", "ULTIMA_FECHA_HISTORICA"
    ]].copy()

    return universo[universo["POZO_KEY"] != ""].drop_duplicates("POZO_KEY")


# ============================================================
# FILTROS Y RESÚMENES BASE
# ============================================================

def filtrar_rango_fechas(df, fecha_inicio, fecha_fin):
    inicio = pd.to_datetime(fecha_inicio)
    fin = pd.to_datetime(fecha_fin)
    return df[(df["FECHA"] >= inicio) & (df["FECHA"] <= fin)].copy()


def aplicar_filtros_universo(universo, baterias_sel, clases_sel):
    data = universo.copy()
    if baterias_sel:
        data = data[data["BATERIA"].isin(baterias_sel)]
    if clases_sel:
        data = data[data["CLASIFICACION"].isin(clases_sel)]
    return data


def filtrar_movimientos_mes(df, anio, mes, baterias_sel=None, tipos_sel=None, clases_sel=None):
    data = df[(df["ANIO"] == int(anio)) & (df["MES"] == int(mes))].copy()
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
    salida["OIL_POR_INTERV"] = np.where(salida["INTERVENCIONES"] > 0, salida["PRCR"] / salida["INTERVENCIONES"], 0)
    return salida.sort_values("PRCR", ascending=False)


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
    salida["OIL_POR_INTERV"] = np.where(salida["INTERVENCIONES"] > 0, salida["PRCR"] / salida["INTERVENCIONES"], 0)
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
    salida["OIL_POR_INTERV"] = np.where(salida["INTERVENCIONES"] > 0, salida["PRCR"] / salida["INTERVENCIONES"], 0)
    return salida.sort_values("PRCR", ascending=False)


def tendencia_mensual(df, anio, baterias_sel=None, tipos_sel=None, clases_sel=None):
    data = df[df["ANIO"] == int(anio)].copy()
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
    salida["OIL_POR_INTERV"] = np.where(salida["INTERVENCIONES"] > 0, salida["PRCR"] / salida["INTERVENCIONES"], 0)
    return salida


# ============================================================
# SERVICIO Y DESPLAZAMIENTO
# ============================================================

def resumen_pozo_periodo(data, sufijo, divisor_meses=1):
    if data.empty:
        return pd.DataFrame(columns=[
            "POZO_KEY",
            f"INTERV_{sufijo}", f"PRCR_{sufijo}", f"PRAG_{sufijo}",
            f"OIL_INTERV_{sufijo}", f"ULTIMA_FECHA_{sufijo}"
        ])

    res = (
        data
        .groupby("POZO_KEY", as_index=False)
        .agg(
            INTERV=("FECHA", "count"),
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum"),
            ULTIMA_FECHA=("FECHA", "max")
        )
    )

    divisor = max(float(divisor_meses), 1.0)
    res[f"INTERV_{sufijo}"] = res["INTERV"] / divisor
    res[f"PRCR_{sufijo}"] = res["PRCR"] / divisor
    res[f"PRAG_{sufijo}"] = res["PRAG"] / divisor
    res[f"OIL_INTERV_{sufijo}"] = np.where(
        res[f"INTERV_{sufijo}"] > 0,
        res[f"PRCR_{sufijo}"] / res[f"INTERV_{sufijo}"],
        0
    )
    res[f"ULTIMA_FECHA_{sufijo}"] = res["ULTIMA_FECHA"]

    return res[[
        "POZO_KEY",
        f"INTERV_{sufijo}", f"PRCR_{sufijo}", f"PRAG_{sufijo}",
        f"OIL_INTERV_{sufijo}", f"ULTIMA_FECHA_{sufijo}"
    ]]


def calcular_servicio_y_desplazamiento(df, universo, anio_objetivo, mes_objetivo, meses_base, baterias_sel, tipos_sel, clases_desplazadas):
    actual_ini = primer_dia_mes(anio_objetivo, mes_objetivo)
    actual_fin = ultimo_dia_mes(anio_objetivo, mes_objetivo)
    actual = df[(df["FECHA"] >= actual_ini) & (df["FECHA"] <= actual_fin)].copy()

    meses_previos = lista_meses_previos(anio_objetivo, mes_objetivo, meses_base)
    partes_base = []
    for a, m in meses_previos:
        ini = primer_dia_mes(a, m)
        fin = ultimo_dia_mes(a, m)
        partes_base.append(df[(df["FECHA"] >= ini) & (df["FECHA"] <= fin)].copy())
    base = pd.concat(partes_base, ignore_index=True) if partes_base else df.iloc[0:0].copy()

    # Mes anterior directo al mes objetivo.
    # Esta comparación es independiente del promedio de N meses anteriores.
    anio_mes_ant, mes_mes_ant = mes_anterior(anio_objetivo, mes_objetivo, 1)
    mes_ant_ini = primer_dia_mes(anio_mes_ant, mes_mes_ant)
    mes_ant_fin = ultimo_dia_mes(anio_mes_ant, mes_mes_ant)
    mes_anterior_directo = df[(df["FECHA"] >= mes_ant_ini) & (df["FECHA"] <= mes_ant_fin)].copy()

    if baterias_sel:
        actual = actual[actual["BATERIA"].isin(baterias_sel)]
        base = base[base["BATERIA"].isin(baterias_sel)]
        mes_anterior_directo = mes_anterior_directo[mes_anterior_directo["BATERIA"].isin(baterias_sel)]
    if tipos_sel:
        actual = actual[actual["TIPO_SWAB"].isin(tipos_sel)]
        base = base[base["TIPO_SWAB"].isin(tipos_sel)]
        mes_anterior_directo = mes_anterior_directo[mes_anterior_directo["TIPO_SWAB"].isin(tipos_sel)]

    universo_eval = universo.copy()
    if baterias_sel:
        universo_eval = universo_eval[universo_eval["BATERIA"].isin(baterias_sel)]
    if clases_desplazadas:
        universo_eval = universo_eval[universo_eval["CLASIFICACION"].isin(clases_desplazadas)]

    res_actual = resumen_pozo_periodo(actual, "ACTUAL", divisor_meses=1)
    res_base = resumen_pozo_periodo(base, "BASE", divisor_meses=meses_base)
    res_mes_anterior = resumen_pozo_periodo(mes_anterior_directo, "MES_ANTERIOR_OBJETIVO", divisor_meses=1)

    tabla = universo_eval.merge(res_base, on="POZO_KEY", how="left")
    tabla = tabla.merge(res_actual, on="POZO_KEY", how="left")
    tabla = tabla.merge(res_mes_anterior, on="POZO_KEY", how="left")

    for col in [
        "INTERV_BASE", "PRCR_BASE", "PRAG_BASE", "OIL_INTERV_BASE",
        "INTERV_ACTUAL", "PRCR_ACTUAL", "PRAG_ACTUAL", "OIL_INTERV_ACTUAL",
        "INTERV_MES_ANTERIOR_OBJETIVO", "PRCR_MES_ANTERIOR_OBJETIVO",
        "PRAG_MES_ANTERIOR_OBJETIVO", "OIL_INTERV_MES_ANTERIOR_OBJETIVO"
    ]:
        if col not in tabla.columns:
            tabla[col] = 0
        tabla[col] = tabla[col].fillna(0)

    tabla["VAR_INTERV"] = tabla["INTERV_ACTUAL"] - tabla["INTERV_BASE"]
    tabla["VAR_PRCR"] = tabla["PRCR_ACTUAL"] - tabla["PRCR_BASE"]
    tabla["VAR_PRAG"] = tabla["PRAG_ACTUAL"] - tabla["PRAG_BASE"]

    tabla["INTERV_BASE_NO_REALIZADAS"] = np.where(tabla["VAR_INTERV"] < 0, -tabla["VAR_INTERV"], 0)
    tabla["PRCR_BASE_NO_REALIZADO"] = np.where(tabla["VAR_PRCR"] < 0, -tabla["VAR_PRCR"], 0)
    tabla["PRAG_BASE_NO_REALIZADO"] = np.where(tabla["VAR_PRAG"] < 0, -tabla["VAR_PRAG"], 0)

    # Comparación directa contra el mes anterior al mes objetivo.
    # Responde: cuánto petróleo se dejó de producir respecto al mes anterior.
    tabla["VAR_PETROLEO_VS_MES_ANTERIOR"] = tabla["PRCR_ACTUAL"] - tabla["PRCR_MES_ANTERIOR_OBJETIVO"]
    tabla["PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR"] = np.where(
        tabla["VAR_PETROLEO_VS_MES_ANTERIOR"] < 0,
        -tabla["VAR_PETROLEO_VS_MES_ANTERIOR"],
        0
    )
    tabla["INTERVENCIONES_DEJADAS_MES_ANTERIOR"] = np.where(
        tabla["INTERV_ACTUAL"] < tabla["INTERV_MES_ANTERIOR_OBJETIVO"],
        tabla["INTERV_MES_ANTERIOR_OBJETIVO"] - tabla["INTERV_ACTUAL"],
        0
    )

    def estado(row):
        if row["INTERV_BASE"] > 0 and row["INTERV_ACTUAL"] == 0:
            return "DEJADO DE HACER"
        if row["INTERV_BASE"] > 0 and row["INTERV_ACTUAL"] > 0 and row["INTERV_ACTUAL"] < row["INTERV_BASE"]:
            return "REDUCIDO"
        if row["INTERV_BASE"] == 0 and row["INTERV_ACTUAL"] > 0:
            return "NUEVO / RETOMADO"
        if row["INTERV_BASE"] > 0 and row["INTERV_ACTUAL"] >= row["INTERV_BASE"]:
            return "MANTENIDO / AUMENTADO"
        return "SIN ACTIVIDAD"

    tabla["ESTADO_DESPLAZAMIENTO"] = tabla.apply(estado, axis=1)

    candidatos = tabla[tabla["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])].copy()
    q75 = candidatos["PRCR_BASE_NO_REALIZADO"].quantile(0.75) if not candidatos.empty else 0
    q50 = candidatos["PRCR_BASE_NO_REALIZADO"].quantile(0.50) if not candidatos.empty else 0

    def prioridad(row):
        if row["ESTADO_DESPLAZAMIENTO"] not in ["DEJADO DE HACER", "REDUCIDO"]:
            return ""
        if row["PRCR_BASE_NO_REALIZADO"] >= q75 and row["PRCR_BASE_NO_REALIZADO"] > 0:
            return "ALTA"
        if row["PRCR_BASE_NO_REALIZADO"] >= q50 and row["PRCR_BASE_NO_REALIZADO"] > 0:
            return "MEDIA"
        return "BAJA"

    tabla["PRIORIDAD_REVISION"] = tabla.apply(prioridad, axis=1)

    orden = {
        "DEJADO DE HACER": 1,
        "REDUCIDO": 2,
        "NUEVO / RETOMADO": 3,
        "MANTENIDO / AUMENTADO": 4,
        "SIN ACTIVIDAD": 5
    }
    tabla["ORDEN"] = tabla["ESTADO_DESPLAZAMIENTO"].map(orden).fillna(9)
    tabla = tabla.sort_values(
        ["ORDEN", "PRCR_BASE_NO_REALIZADO", "INTERV_BASE_NO_REALIZADAS"],
        ascending=[True, False, False]
    ).drop(columns="ORDEN")

    periodos = {
        "actual_ini": actual_ini,
        "actual_fin": actual_fin,
        "base_meses": meses_previos,
        "meses_base": meses_base,
        "mes_anterior_directo": (anio_mes_ant, mes_mes_ant)
    }

    return tabla, periodos, actual, base

def resumen_desplazamiento_por_bateria(tabla):
    if tabla.empty:
        return pd.DataFrame()

    data = tabla.copy()
    data["ES_DEJADO"] = data["ESTADO_DESPLAZAMIENTO"].eq("DEJADO DE HACER")
    data["ES_REDUCIDO"] = data["ESTADO_DESPLAZAMIENTO"].eq("REDUCIDO")
    salida = (
        data
        .groupby("BATERIA", as_index=False)
        .agg(
            POZOS_DEJADOS=("ES_DEJADO", "sum"),
            POZOS_REDUCIDOS=("ES_REDUCIDO", "sum"),
            INTERV_BASE_NO_REALIZADAS=("INTERV_BASE_NO_REALIZADAS", "sum"),
            PRCR_BASE_NO_REALIZADO=("PRCR_BASE_NO_REALIZADO", "sum"),
            PRAG_BASE_NO_REALIZADO=("PRAG_BASE_NO_REALIZADO", "sum")
        )
        .sort_values("PRCR_BASE_NO_REALIZADO", ascending=False)
    )
    return salida


def resumen_impacto_convertidos_2026(actual, base, meses_base):
    conv_actual = actual[actual["CLASIFICACION"] == "Convertido 2026"].copy()
    conv_base = base[base["CLASIFICACION"] == "Convertido 2026"].copy()

    return {
        "POZOS_ACTUAL": conv_actual["POZO_KEY"].nunique() if not conv_actual.empty else 0,
        "INTERV_ACTUAL": len(conv_actual),
        "PRCR_ACTUAL": conv_actual["PRCR"].sum() if not conv_actual.empty else 0,
        "PRAG_ACTUAL": conv_actual["PRAG"].sum() if not conv_actual.empty else 0,
        "POZOS_BASE_PROM": conv_base["POZO_KEY"].nunique() if not conv_base.empty else 0,
        "INTERV_BASE_PROM": len(conv_base) / max(float(meses_base), 1.0),
        "PRCR_BASE_PROM": conv_base["PRCR"].sum() / max(float(meses_base), 1.0) if not conv_base.empty else 0,
        "PRAG_BASE_PROM": conv_base["PRAG"].sum() / max(float(meses_base), 1.0) if not conv_base.empty else 0
    }



def construir_listado_convertidos(universo, data_mes_convertidos):
    """
    Lista fija de pozos convertidos 2024, 2025 y 2026.
    Incluye el estado del mes objetivo solo como referencia operativa.
    """
    universo_conv = universo[universo["ANIO_CONVERSION"].isin([2024, 2025, 2026])].copy()

    if data_mes_convertidos.empty:
        resumen_mes = pd.DataFrame(columns=[
            "POZO_KEY", "INTERVENCIONES_MES", "PRODUCCION_PETROLEO_MES", "PRODUCCION_AGUA_MES",
            "TIPO_SWAB_MES", "ULTIMA_FECHA_MES"
        ])
    else:
        resumen_mes = (
            data_mes_convertidos
            .groupby("POZO_KEY", as_index=False)
            .agg(
                INTERVENCIONES_MES=("FECHA", "count"),
                PRODUCCION_PETROLEO_MES=("PRCR", "sum"),
                PRODUCCION_AGUA_MES=("PRAG", "sum"),
                TIPO_SWAB_MES=("TIPO_SWAB", lambda x: ", ".join(sorted(set([v for v in x if v])))),
                ULTIMA_FECHA_MES=("FECHA", "max")
            )
        )

    salida = universo_conv.merge(resumen_mes, on="POZO_KEY", how="left")

    salida["INTERVENCIONES_MES"] = salida["INTERVENCIONES_MES"].fillna(0).astype(int)
    salida["PRODUCCION_PETROLEO_MES"] = salida["PRODUCCION_PETROLEO_MES"].fillna(0)
    salida["PRODUCCION_AGUA_MES"] = salida["PRODUCCION_AGUA_MES"].fillna(0)
    salida["TIPO_SWAB_MES"] = salida["TIPO_SWAB_MES"].fillna("")
    salida["ESTADO_MES"] = np.where(salida["INTERVENCIONES_MES"] > 0, "Intervenido", "No intervenido")

    salida = salida.rename(columns={
        "ANIO_CONVERSION": "AÑO_CONVERSION",
        "CLASIFICACION": "TIPO_CONVERTIDO",
        "ULTIMA_FECHA_HISTORICA": "ULTIMA_FECHA_HISTORICA"
    })

    columnas = [
        "AÑO_CONVERSION", "TIPO_CONVERTIDO", "POZO", "BATERIA",
        "ESTADO_MES", "INTERVENCIONES_MES", "PRODUCCION_PETROLEO_MES",
        "PRODUCCION_AGUA_MES", "TIPO_SWAB_MES", "ULTIMA_FECHA_MES",
        "ULTIMA_FECHA_HISTORICA"
    ]

    salida = salida[columnas].sort_values(["AÑO_CONVERSION", "POZO"])
    return salida

# ============================================================
# POTENCIAL MENSUAL POR POZO
# ============================================================

def calcular_potencial_ultimo_mes_activo(df, universo, baterias_sel=None, clases_sel=None, tipos_sel=None):
    """
    Calcula el potencial de cada pozo con dos referencias:
    1. Último mes activo con producción de petróleo > 0.
    2. Mes calendario anterior al último mes activo.

    Fórmula:
    Potencial bopd = producción de petróleo del mes / días calendario del mes.
    """
    data = df.copy()
    if baterias_sel:
        data = data[data["BATERIA"].isin(baterias_sel)]
    if clases_sel:
        data = data[data["CLASIFICACION"].isin(clases_sel)]
    if tipos_sel:
        data = data[data["TIPO_SWAB"].isin(tipos_sel)]

    data_prod = data[data["PRCR"] > 0].copy()

    columnas_salida = [
        "POZO_KEY", "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION",
        "ULTIMO_MES_ACTIVO", "DIAS_MES", "PRCR_ULTIMO_MES", "PRAG_ULTIMO_MES",
        "INTERV_ULTIMO_MES", "OIL_POR_INTERV_ULTIMO_MES",
        "POTENCIAL_ULTIMO_MES_BOPD", "POTENCIAL_BOPD", "ULTIMA_FECHA_CON_PRCR",
        "MES_ANTERIOR_ULTIMO_ACTIVO", "DIAS_MES_ANTERIOR", "PRCR_MES_ANTERIOR",
        "PRAG_MES_ANTERIOR", "INTERV_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD"
    ]

    if data_prod.empty:
        salida = universo.copy()
        for col in columnas_salida:
            if col not in salida.columns:
                salida[col] = 0
        salida["ULTIMO_MES_ACTIVO"] = "Sin producción de petróleo > 0 en rango"
        salida["MES_ANTERIOR_ULTIMO_ACTIVO"] = "Sin referencia"
        return salida[columnas_salida]

    data_prod["ANIO_MES"] = data_prod["FECHA"].dt.to_period("M")

    mensual = (
        data_prod
        .groupby(["POZO_KEY", "ANIO_MES"], as_index=False)
        .agg(
            PRCR_MES=("PRCR", "sum"),
            PRAG_MES=("PRAG", "sum"),
            INTERV_MES=("FECHA", "count"),
            ULTIMA_FECHA_CON_PRCR=("FECHA", "max")
        )
    )

    idx = mensual.groupby("POZO_KEY")["ANIO_MES"].idxmax()
    ultimo = mensual.loc[idx].copy()

    ultimo["ANIO"] = ultimo["ANIO_MES"].dt.year
    ultimo["MES"] = ultimo["ANIO_MES"].dt.month
    ultimo["DIAS_MES"] = ultimo.apply(lambda r: dias_calendario_mes(r["ANIO"], r["MES"]), axis=1)
    ultimo["POTENCIAL_ULTIMO_MES_BOPD"] = ultimo["PRCR_MES"] / ultimo["DIAS_MES"]
    ultimo["POTENCIAL_BOPD"] = ultimo["POTENCIAL_ULTIMO_MES_BOPD"]
    ultimo["OIL_POR_INTERV_ULTIMO_MES"] = np.where(
        ultimo["INTERV_MES"] > 0,
        ultimo["PRCR_MES"] / ultimo["INTERV_MES"],
        0
    )
    ultimo["ULTIMO_MES_ACTIVO"] = ultimo.apply(lambda r: periodo_mes_texto(r["ANIO"], r["MES"]), axis=1)

    ultimo["ANIO_MES_ANTERIOR"] = ultimo["ANIO_MES"] - 1

    mensual_anterior = mensual.rename(columns={
        "ANIO_MES": "ANIO_MES_ANTERIOR",
        "PRCR_MES": "PRCR_MES_ANTERIOR",
        "PRAG_MES": "PRAG_MES_ANTERIOR",
        "INTERV_MES": "INTERV_MES_ANTERIOR",
        "ULTIMA_FECHA_CON_PRCR": "ULTIMA_FECHA_CON_PRCR_MES_ANTERIOR"
    })

    ultimo = ultimo.merge(
        mensual_anterior[[
            "POZO_KEY", "ANIO_MES_ANTERIOR", "PRCR_MES_ANTERIOR",
            "PRAG_MES_ANTERIOR", "INTERV_MES_ANTERIOR"
        ]],
        on=["POZO_KEY", "ANIO_MES_ANTERIOR"],
        how="left"
    )

    ultimo["ANIO_ANTERIOR"] = ultimo["ANIO_MES_ANTERIOR"].dt.year
    ultimo["MES_ANTERIOR"] = ultimo["ANIO_MES_ANTERIOR"].dt.month
    ultimo["DIAS_MES_ANTERIOR"] = ultimo.apply(
        lambda r: dias_calendario_mes(r["ANIO_ANTERIOR"], r["MES_ANTERIOR"]), axis=1
    )
    ultimo["MES_ANTERIOR_ULTIMO_ACTIVO"] = ultimo.apply(
        lambda r: periodo_mes_texto(r["ANIO_ANTERIOR"], r["MES_ANTERIOR"]), axis=1
    )

    for col in ["PRCR_MES_ANTERIOR", "PRAG_MES_ANTERIOR", "INTERV_MES_ANTERIOR"]:
        ultimo[col] = ultimo[col].fillna(0)

    ultimo["POTENCIAL_MES_ANTERIOR_BOPD"] = ultimo["PRCR_MES_ANTERIOR"] / ultimo["DIAS_MES_ANTERIOR"]
    ultimo["POTENCIAL_PROMEDIO_2_MESES_BOPD"] = (
        ultimo["PRCR_MES"] + ultimo["PRCR_MES_ANTERIOR"]
    ) / (
        ultimo["DIAS_MES"] + ultimo["DIAS_MES_ANTERIOR"]
    )

    salida = universo.merge(ultimo, on="POZO_KEY", how="left")
    salida = salida.rename(columns={
        "PRCR_MES": "PRCR_ULTIMO_MES",
        "PRAG_MES": "PRAG_ULTIMO_MES",
        "INTERV_MES": "INTERV_ULTIMO_MES"
    })

    salida = salida[columnas_salida].copy()

    numericas = [
        "DIAS_MES", "PRCR_ULTIMO_MES", "PRAG_ULTIMO_MES", "INTERV_ULTIMO_MES",
        "OIL_POR_INTERV_ULTIMO_MES", "POTENCIAL_ULTIMO_MES_BOPD", "POTENCIAL_BOPD",
        "DIAS_MES_ANTERIOR", "PRCR_MES_ANTERIOR", "PRAG_MES_ANTERIOR",
        "INTERV_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD"
    ]
    for col in numericas:
        salida[col] = salida[col].fillna(0)

    salida["ULTIMO_MES_ACTIVO"] = salida["ULTIMO_MES_ACTIVO"].fillna("Sin producción de petróleo > 0 en rango")
    salida["MES_ANTERIOR_ULTIMO_ACTIVO"] = salida["MES_ANTERIOR_ULTIMO_ACTIVO"].fillna("Sin referencia")

    return salida.sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False)


# ============================================================
# PPT EDITABLE
# ============================================================

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

    shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    table = shape.table

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
    if not categorias:
        return
    chart_data = CategoryChartData()
    chart_data.categories = categorias
    chart_data.add_series(nombre_serie, valores)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, chart_data).chart
    chart.has_legend = False
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = titulo


def agregar_linea(slide, titulo, categorias, serie_prcr, serie_prag, x, y, w, h):
    categorias = list(categorias)
    if not categorias:
        return
    chart_data = CategoryChartData()
    chart_data.categories = categorias
    chart_data.add_series("Producción de petróleo", [0 if pd.isna(v) else float(v) for v in serie_prcr])
    chart_data.add_series("Producción de agua", [0 if pd.isna(v) else float(v) for v in serie_prag])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, w, h, chart_data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = titulo


def crear_ppt(periodo, kpis, resumen_clase, resumen_bateria, resumen_tipo, tendencia, tabla_desplazamiento, potencial):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Dashboard SWAB Lote X", periodo)
    kpi_df = pd.DataFrame({
        "Indicador": ["Pozos universo", "Pozos intervenidos", "Pozos no intervenidos", "Intervenciones", "Producción de petróleo", "Producción de agua", "Oil por intervención"],
        "Valor": [kpis["pozos_universo"], kpis["pozos_intervenidos"], kpis["pozos_no_intervenidos"], kpis["intervenciones"], round(kpis["prcr"], 2), round(kpis["prag"], 2), round(kpis["oil_interv"], 2)]
    })
    agregar_tabla(slide, kpi_df, Inches(0.5), Inches(1.25), Inches(5.2), Inches(3.0), font_size=10)
    estado_df = pd.DataFrame({"Estado": ["Intervenido", "No intervenido"], "Pozos": [kpis["pozos_intervenidos"], kpis["pozos_no_intervenidos"]]})
    agregar_barras(slide, "Pozos intervenidos vs no intervenidos", estado_df["Estado"], estado_df["Pozos"], Inches(6.2), Inches(1.2), Inches(6.5), Inches(4.8), "Pozos")

    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Producción por clasificación", periodo)
    clase_plot = resumen_clase.sort_values("PRCR", ascending=False)
    agregar_barras(slide, "Producción de petróleo por clasificación", clase_plot["CLASIFICACION"], clase_plot["PRCR"], Inches(0.5), Inches(1.2), Inches(6.4), Inches(5.4), "PRCR")
    agregar_tabla(slide, resumen_clase, Inches(7.1), Inches(1.2), Inches(5.7), Inches(5.4), font_size=7)

    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Baterías", periodo)
    top_bat = resumen_bateria.sort_values("PRCR", ascending=False).head(15)
    agregar_barras(slide, "Top baterías por PRCR", top_bat["BATERIA"], top_bat["PRCR"], Inches(0.5), Inches(1.2), Inches(7.0), Inches(5.4), "PRCR")
    agregar_tabla(slide, top_bat, Inches(7.7), Inches(1.2), Inches(5.2), Inches(5.4), font_size=7)

    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "TS y CS", periodo)
    agregar_barras(slide, "Producción de petróleo por tipo de swab", resumen_tipo["TIPO_SWAB"], resumen_tipo["PRCR"], Inches(0.5), Inches(1.2), Inches(6.5), Inches(5.4), "PRCR")
    agregar_tabla(slide, resumen_tipo, Inches(7.2), Inches(1.2), Inches(5.6), Inches(4.8), font_size=8)

    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Tendencia mensual", periodo)
    if not tendencia.empty:
        agregar_linea(slide, "Producción de petróleo y Producción de agua por mes", tendencia["MES_NOMBRE"], tendencia["PRCR"], tendencia["PRAG"], Inches(0.7), Inches(1.2), Inches(12.0), Inches(5.5))

    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Pozos dejados o reducidos", periodo)
    dejados = tabla_desplazamiento[tabla_desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])].sort_values("PRCR_BASE_NO_REALIZADO", ascending=False).head(20)
    cols_dej = ["ESTADO_DESPLAZAMIENTO", "POZO", "BATERIA", "CLASIFICACION", "INTERV_BASE", "INTERV_ACTUAL", "PRCR_BASE_NO_REALIZADO", "PRIORIDAD_REVISION"]
    cols_dej = [c for c in cols_dej if c in dejados.columns]
    agregar_tabla(slide, dejados[cols_dej], Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.8), font_size=8)

    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Potencial mensual por pozo", periodo)
    pot = potencial.sort_values("POTENCIAL_BOPD", ascending=False).head(15)
    agregar_barras(slide, "Top pozos por potencial estimado BOPD", pot["POZO"], pot["POTENCIAL_BOPD"], Inches(0.5), Inches(1.2), Inches(7.0), Inches(5.4), "BOPD")
    cols_pot = ["POZO", "BATERIA", "CLASIFICACION", "ULTIMO_MES_ACTIVO", "PRCR_ULTIMO_MES", "POTENCIAL_BOPD"]
    agregar_tabla(slide, pot[cols_pot], Inches(7.7), Inches(1.2), Inches(5.2), Inches(5.4), font_size=7)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

st.title("🛢️ SWAB Lote X - Servicio, desplazamiento y potencial")
st.caption("Producción de petróleo = PRCR. Producción de agua = PRAG. Carga un solo Excel con la hoja Datos de Swab.")

archivo = st.file_uploader("Sube el Excel principal", type=["xlsx"])

if archivo is None:
    st.info("Sube el Excel para habilitar el análisis.")
    st.stop()

bytes_excel = archivo.getvalue()

try:
    df_raw_completo, hoja_usada = cargar_excel_swab(bytes_excel)
except Exception as e:
    st.error(f"No se pudo cargar el Excel: {e}")
    st.stop()

# Se excluyen del análisis los 88 pozos candidatos para ATA.
df_raw, df_ata_excluidos = excluir_candidatos_ata(df_raw_completo)
pozos_ata_excluidos_en_data = df_ata_excluidos["POZO_KEY"].nunique() if not df_ata_excluidos.empty else 0
registros_ata_excluidos = len(df_ata_excluidos)

if df_raw.empty:
    st.error("Después de excluir los pozos candidatos para ATA, no quedaron datos para analizar.")
    st.stop()

fecha_min = df_raw["FECHA"].min().date()
fecha_max = df_raw["FECHA"].max().date()

with st.sidebar:
    st.header("1. Rango de análisis")
    fecha_inicio_analisis = st.date_input("Desde", value=fecha_min, min_value=fecha_min, max_value=fecha_max)
    fecha_fin_analisis = st.date_input("Hasta", value=fecha_max, min_value=fecha_min, max_value=fecha_max)

    st.header("2. Módulo principal")
    modulo = st.radio(
        "Selecciona con qué trabajar",
        ["Análisis de pozos dejados", "Potencial mensual de pozos", "Vista completa"],
        index=2
    )

df = filtrar_rango_fechas(df_raw, fecha_inicio_analisis, fecha_fin_analisis)

if df.empty:
    st.error("El rango seleccionado no tiene datos.")
    st.stop()

universo = construir_universo(df)
anios = sorted(df["ANIO"].dropna().unique().astype(int).tolist())
baterias = sorted([x for x in universo["BATERIA"].dropna().unique().tolist() if x != ""])
tipos = sorted([x for x in df["TIPO_SWAB"].dropna().unique().tolist() if x != ""])
clases = ["Básica", "Convertido 2024", "Convertido 2025", "Convertido 2026"]

ultimo_mes_disponible = df["PERIODO_MES"].max()
anio_default = int(ultimo_mes_disponible.year)
mes_default = int(ultimo_mes_disponible.month)

with st.sidebar:
    st.header("3. Periodo objetivo")
    anio_objetivo = st.selectbox("Año objetivo", anios, index=anios.index(anio_default) if anio_default in anios else len(anios) - 1)
    meses_disponibles = sorted(df[df["ANIO"] == anio_objetivo]["MES"].unique().astype(int).tolist())
    mes_objetivo = st.selectbox(
        "Mes objetivo",
        meses_disponibles,
        index=meses_disponibles.index(mes_default) if mes_default in meses_disponibles else len(meses_disponibles) - 1,
        format_func=lambda x: MESES[int(x)]
    )
    meses_base = st.slider("Comparar contra N meses anteriores", min_value=1, max_value=15, value=3, step=1)

    st.header("4. Filtros operativos")
    baterias_sel = st.multiselect("Batería", baterias, default=[])
    tipos_sel = st.multiselect("Tipo de swab", tipos, default=[])
    clases_sel = st.multiselect("Tipo de pozo para vistas generales", clases, default=[])
    clases_desplazadas = st.multiselect(
        "Pozos a evaluar como desplazados",
        clases,
        default=["Básica", "Convertido 2024", "Convertido 2025"]
    )
    top_n = st.slider("Top para gráficos", min_value=5, max_value=50, value=20, step=5)
    ejecutar = st.button("Ejecutar análisis", type="primary")

if not ejecutar and "swab_servicio_resultados" not in st.session_state:
    st.warning("Configura el rango, el mes objetivo y presiona Ejecutar análisis.")
    st.stop()

if ejecutar:
    universo_general = aplicar_filtros_universo(universo, baterias_sel, clases_sel)
    data_mes = filtrar_movimientos_mes(df, anio_objetivo, mes_objetivo, baterias_sel, tipos_sel, clases_sel)

    resumen_pozos = resumir_pozos(data_mes, universo_general)
    res_clase = resumen_clasificacion(resumen_pozos)
    res_bateria = resumen_baterias(resumen_pozos)
    res_tipo = resumen_tipo_swab(data_mes)
    tend = tendencia_mensual(df, anio_objetivo, baterias_sel, tipos_sel, clases_sel)

    desplazamiento, periodos_desplazamiento, actual_mes, base_periodo = calcular_servicio_y_desplazamiento(
        df=df,
        universo=universo,
        anio_objetivo=anio_objetivo,
        mes_objetivo=mes_objetivo,
        meses_base=meses_base,
        baterias_sel=baterias_sel,
        tipos_sel=tipos_sel,
        clases_desplazadas=clases_desplazadas
    )

    potencial = calcular_potencial_ultimo_mes_activo(
        df=df,
        universo=universo,
        baterias_sel=baterias_sel,
        clases_sel=clases_sel if clases_sel else None,
        tipos_sel=tipos_sel
    )

    columnas_potencial_desplazamiento = [
        "POZO_KEY", "ULTIMO_MES_ACTIVO", "PRCR_ULTIMO_MES",
        "POTENCIAL_ULTIMO_MES_BOPD", "MES_ANTERIOR_ULTIMO_ACTIVO",
        "PRCR_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD"
    ]
    desplazamiento = desplazamiento.merge(
        potencial[columnas_potencial_desplazamiento],
        on="POZO_KEY",
        how="left"
    )

    for col in [
        "PRCR_ULTIMO_MES", "POTENCIAL_ULTIMO_MES_BOPD",
        "PRCR_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD"
    ]:
        desplazamiento[col] = desplazamiento[col].fillna(0)

    desplazamiento["ULTIMO_MES_ACTIVO"] = desplazamiento["ULTIMO_MES_ACTIVO"].fillna("Sin producción de petróleo > 0 en rango")
    desplazamiento["MES_ANTERIOR_ULTIMO_ACTIVO"] = desplazamiento["MES_ANTERIOR_ULTIMO_ACTIVO"].fillna("Sin referencia")

    impacto_2026 = resumen_impacto_convertidos_2026(actual_mes, base_periodo, meses_base)

    data_mes_convertidos = filtrar_movimientos_mes(
        df,
        anio_objetivo,
        mes_objetivo,
        baterias_sel,
        tipos_sel,
        ["Convertido 2024", "Convertido 2025", "Convertido 2026"]
    )
    listado_convertidos = construir_listado_convertidos(universo, data_mes_convertidos)

    st.session_state["swab_servicio_resultados"] = {
        "anio_objetivo": anio_objetivo,
        "mes_objetivo": mes_objetivo,
        "meses_base": meses_base,
        "modulo": modulo,
        "data_mes": data_mes,
        "resumen_pozos": resumen_pozos,
        "res_clase": res_clase,
        "res_bateria": res_bateria,
        "res_tipo": res_tipo,
        "tendencia": tend,
        "desplazamiento": desplazamiento,
        "periodos_desplazamiento": periodos_desplazamiento,
        "potencial": potencial,
        "impacto_2026": impacto_2026,
        "listado_convertidos": listado_convertidos,
        "filtros": {
            "baterias_sel": baterias_sel,
            "tipos_sel": tipos_sel,
            "clases_sel": clases_sel,
            "clases_desplazadas": clases_desplazadas
        }
    }

res = st.session_state["swab_servicio_resultados"]
anio_objetivo = res["anio_objetivo"]
mes_objetivo = res["mes_objetivo"]
meses_base = res["meses_base"]
data_mes = res["data_mes"]
resumen_pozos = res["resumen_pozos"]
res_clase = res["res_clase"]
res_bateria = res["res_bateria"]
res_tipo = res["res_tipo"]
tendencia = res["tendencia"]
desplazamiento = res["desplazamiento"]
periodos_desplazamiento = res["periodos_desplazamiento"]
potencial = res["potencial"]
impacto_2026 = res["impacto_2026"]
listado_convertidos = res.get("listado_convertidos", pd.DataFrame())

periodo = periodo_mes_texto(anio_objetivo, mes_objetivo)
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

dejados = desplazamiento[desplazamiento["ESTADO_DESPLAZAMIENTO"] == "DEJADO DE HACER"].copy()
reducidos = desplazamiento[desplazamiento["ESTADO_DESPLAZAMIENTO"] == "REDUCIDO"].copy()
res_desplazamiento_bateria = resumen_desplazamiento_por_bateria(desplazamiento)

st.subheader(f"Resumen ejecutivo: {periodo}")

k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1:
    st.metric("Pozos universo", f"{pozos_universo:,}")
with k2:
    st.metric("Intervenidos", f"{pozos_intervenidos:,}")
with k3:
    st.metric("No intervenidos", f"{pozos_no_intervenidos:,}")
with k4:
    st.metric("Intervenciones", f"{intervenciones:,}")
with k5:
    st.metric("Producción de petróleo", f"{prcr_total:,.2f}")
with k6:
    st.metric("Producción de agua", f"{prag_total:,.2f}")

st.caption(
    f"Hoja usada: {hoja_usada}. Rango analizado: {fecha_inicio_analisis} al {fecha_fin_analisis}. "
    f"Línea base: promedio mensual de {meses_base} mes(es) anteriores. "
    f"Se excluyeron {pozos_ata_excluidos_en_data} pozos candidatos para ATA, equivalentes a {registros_ata_excluidos:,} registros."
)

st.divider()

tabs = st.tabs([
    "Análisis de pozos dejados",
    "Impacto convertidos 2026",
    "Listado convertidos",
    "Potencial mensual",
    "Pozos y baterías",
    "TS y CS",
    "Descargas"
])


with tabs[0]:
    st.subheader("Análisis de pozos dejados de hacer o reducidos")
    st.caption(
        "Este bloque cruza los pozos dejados o reducidos con su potencial. "
        "El potencial se calcula con la producción de petróleo mensual dividida entre los días calendario del mes."
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        st.metric("Pozos dejados", f"{len(dejados):,}")
    with c2:
        st.metric("Pozos reducidos", f"{len(reducidos):,}")
    with c3:
        st.metric("Interv. base no realizadas", f"{desplazamiento['INTERV_BASE_NO_REALIZADAS'].sum():,.2f}")
    with c4:
        st.metric("Petróleo base no realizado", f"{desplazamiento['PRCR_BASE_NO_REALIZADO'].sum():,.2f}")
    with c5:
        st.metric("Petróleo dejado vs mes anterior", f"{desplazamiento['PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR'].sum():,.2f}")
    with c6:
        st.metric("Potencial promedio 2 meses", f"{desplazamiento['POTENCIAL_PROMEDIO_2_MESES_BOPD'].sum():,.2f} bopd")

    meses_base_txt = ", ".join([periodo_mes_texto(a, m) for a, m in periodos_desplazamiento["base_meses"]])
    st.info(
        f"Periodo objetivo: {periodo}. Base comparativa: promedio mensual de {meses_base} mes(es) anteriores: {meses_base_txt}. "
        "El potencial se muestra con dos referencias: último mes activo y mes anterior al último mes activo."
    )

    estado_sel = st.radio(
        "Estado de desplazamiento",
        ["DEJADO DE HACER", "REDUCIDO", "NUEVO / RETOMADO", "MANTENIDO / AUMENTADO", "SIN ACTIVIDAD", "TODOS"],
        horizontal=True
    )

    tabla_desp = desplazamiento.copy()
    if estado_sel != "TODOS":
        tabla_desp = tabla_desp[tabla_desp["ESTADO_DESPLAZAMIENTO"] == estado_sel]

    cols_desp = [
        "ESTADO_DESPLAZAMIENTO", "PRIORIDAD_REVISION", "POZO", "BATERIA",
        "CLASIFICACION", "ANIO_CONVERSION",
        "INTERV_BASE", "INTERV_ACTUAL", "INTERV_BASE_NO_REALIZADAS",
        "PRCR_BASE", "PRCR_ACTUAL", "PRCR_BASE_NO_REALIZADO",
        "PRCR_MES_ANTERIOR_OBJETIVO", "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR",
        "VAR_PETROLEO_VS_MES_ANTERIOR", "INTERV_MES_ANTERIOR_OBJETIVO",
        "INTERVENCIONES_DEJADAS_MES_ANTERIOR",
        "ULTIMO_MES_ACTIVO", "PRCR_ULTIMO_MES", "POTENCIAL_ULTIMO_MES_BOPD",
        "MES_ANTERIOR_ULTIMO_ACTIVO", "PRCR_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD",
        "PRAG_BASE", "PRAG_ACTUAL", "PRAG_BASE_NO_REALIZADO",
        "OIL_INTERV_BASE", "OIL_INTERV_ACTUAL", "ULTIMA_FECHA_HISTORICA"
    ]
    cols_desp = [c for c in cols_desp if c in tabla_desp.columns]
    st.dataframe(formatear_tabla(tabla_desp[cols_desp]), use_container_width=True, hide_index=True)

    data_graf_pozos = desplazamiento[
        desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])
    ].copy()
    data_graf_pozos = data_graf_pozos.sort_values("PRCR_BASE_NO_REALIZADO", ascending=False).head(top_n)

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        if not data_graf_pozos.empty:
            fig = px.bar(
                data_graf_pozos,
                x="POZO",
                y="PRCR_BASE_NO_REALIZADO",
                color="ESTADO_DESPLAZAMIENTO",
                hover_data=[
                    "BATERIA", "CLASIFICACION", "INTERV_BASE", "INTERV_ACTUAL",
                    "PRCR_BASE", "PRCR_ACTUAL", "PRCR_MES_ANTERIOR_OBJETIVO",
                    "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR", "ULTIMO_MES_ACTIVO",
                    "POTENCIAL_ULTIMO_MES_BOPD", "POTENCIAL_MES_ANTERIOR_BOPD",
                    "POTENCIAL_PROMEDIO_2_MESES_BOPD"
                ],
                text="PRCR_BASE_NO_REALIZADO"
            )
            fig.update_layout(uniformtext_minsize=8, uniformtext_mode="hide")
            fig = aplicar_layout_fig(fig, f"Top {top_n} pozos dejados o reducidos por producción de petróleo base no realizada", 560)
            st.plotly_chart(fig, use_container_width=True)

    with col_g2:
        data_graf_potencial = desplazamiento[
            desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])
        ].sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False).head(top_n)

        if not data_graf_potencial.empty:
            fig_pot_desp = px.bar(
                data_graf_potencial,
                x="POZO",
                y=["POTENCIAL_ULTIMO_MES_BOPD", "POTENCIAL_MES_ANTERIOR_BOPD"],
                barmode="group",
                hover_data=["BATERIA", "CLASIFICACION", "ULTIMO_MES_ACTIVO", "MES_ANTERIOR_ULTIMO_ACTIVO"]
            )
            fig_pot_desp = aplicar_layout_fig(fig_pot_desp, f"Top {top_n} pozos dejados o reducidos por potencial", 560)
            st.plotly_chart(fig_pot_desp, use_container_width=True)

    col_g3, col_g4 = st.columns(2)
    with col_g3:
        bubble = desplazamiento[desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])].copy()
        if not bubble.empty:
            fig_bubble = px.scatter(
                bubble,
                x="POTENCIAL_PROMEDIO_2_MESES_BOPD",
                y="PRCR_BASE_NO_REALIZADO",
                size="INTERV_BASE_NO_REALIZADAS",
                color="ESTADO_DESPLAZAMIENTO",
                hover_name="POZO",
                hover_data=[
                    "BATERIA", "CLASIFICACION", "PRIORIDAD_REVISION", "INTERV_BASE",
                    "INTERV_ACTUAL", "ULTIMO_MES_ACTIVO", "MES_ANTERIOR_ULTIMO_ACTIVO"
                ]
            )
            fig_bubble = aplicar_layout_fig(fig_bubble, "Mapa de criticidad: potencial vs producción de petróleo no realizada", 560)
            st.plotly_chart(fig_bubble, use_container_width=True)

    with col_g4:
        if not res_desplazamiento_bateria.empty:
            fig_bat = px.bar(
                res_desplazamiento_bateria.head(top_n),
                x="BATERIA",
                y="PRCR_BASE_NO_REALIZADO",
                hover_data=["POZOS_DEJADOS", "POZOS_REDUCIDOS", "INTERV_BASE_NO_REALIZADAS"],
                text="PRCR_BASE_NO_REALIZADO"
            )
            fig_bat = aplicar_layout_fig(fig_bat, "Baterías con mayor producción de petróleo base no realizada", 560)
            st.plotly_chart(fig_bat, use_container_width=True)

    data_graficos = {
        "Graf top pozos petroleo": data_graf_pozos,
        "Graf top potencial": data_graf_potencial,
        "Graf burbuja criticidad": desplazamiento[
            desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])
        ],
        "Graf bateria petroleo": res_desplazamiento_bateria
    }

    st.download_button(
        "Descargar data de las gráficas de pozos dejados",
        data=convertir_excel(data_graficos),
        file_name=f"data_graficas_pozos_dejados_{anio_objetivo}_{mes_objetivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


with tabs[1]:
    st.subheader("Impacto de priorizar los 20 pozos convertidos 2026")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Conv. 2026 intervenidos", f"{impacto_2026['POZOS_ACTUAL']:,} de 20")
    with c2:
        st.metric("Interv. conv. 2026", f"{impacto_2026['INTERV_ACTUAL']:,}")
    with c3:
        st.metric("Petróleo conv. 2026", f"{impacto_2026['PRCR_ACTUAL']:,.2f}")
    with c4:
        st.metric("Agua conv. 2026", f"{impacto_2026['PRAG_ACTUAL']:,.2f}")

    col_i1, col_i2 = st.columns(2)
    with col_i1:
        fig_clase = px.bar(
            res_clase,
            x="CLASIFICACION",
            y=["PRCR", "PRAG"],
            barmode="group",
            hover_data=["POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES"]
        )
        fig_clase = aplicar_layout_fig(fig_clase, "Producción de petróleo y Producción de agua por clasificación", 520)
        st.plotly_chart(fig_clase, use_container_width=True)

    with col_i2:
        fig_interv = px.bar(
            res_clase,
            x="CLASIFICACION",
            y="INTERVENCIONES",
            text="INTERVENCIONES",
            hover_data=["POZOS_TOTAL", "POZOS_INTERVENIDOS", "PRCR", "PRAG"]
        )
        fig_interv = aplicar_layout_fig(fig_interv, "Intervenciones por clasificación", 520)
        st.plotly_chart(fig_interv, use_container_width=True)

    afectados = desplazamiento[desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])].copy()
    if not afectados.empty:
        resumen_afectados = (
            afectados
            .groupby("CLASIFICACION", as_index=False)
            .agg(
                POZOS_AFECTADOS=("POZO_KEY", "nunique"),
                INTERV_BASE_NO_REALIZADAS=("INTERV_BASE_NO_REALIZADAS", "sum"),
                PRCR_BASE_NO_REALIZADO=("PRCR_BASE_NO_REALIZADO", "sum"),
                PRAG_BASE_NO_REALIZADO=("PRAG_BASE_NO_REALIZADO", "sum")
            )
            .sort_values("PRCR_BASE_NO_REALIZADO", ascending=False)
        )
        st.write("Pozos posiblemente desplazados por clasificación")
        st.dataframe(formatear_tabla(resumen_afectados), use_container_width=True, hide_index=True)


with tabs[2]:
    st.subheader("Listado de pozos convertidos 2024, 2025 y 2026")
    st.caption(
        "Listado fijo de pozos convertidos cargados en el código. "
        "La columna Estado del mes muestra si el pozo tuvo intervención en el mes objetivo seleccionado."
    )

    if listado_convertidos.empty:
        st.info("No se encontró listado de pozos convertidos.")
    else:
        anios_convertidos = sorted(listado_convertidos["AÑO_CONVERSION"].dropna().unique().astype(int).tolist())
        anios_sel_convertidos = st.multiselect(
            "Filtrar por año de conversión",
            anios_convertidos,
            default=anios_convertidos
        )

        tabla_conv = listado_convertidos.copy()
        if anios_sel_convertidos:
            tabla_conv = tabla_conv[tabla_conv["AÑO_CONVERSION"].isin(anios_sel_convertidos)]

        resumen_conv = (
            tabla_conv
            .groupby(["AÑO_CONVERSION", "TIPO_CONVERTIDO"], as_index=False)
            .agg(
                POZOS=("POZO", "nunique"),
                INTERVENIDOS_MES=("ESTADO_MES", lambda x: (x == "Intervenido").sum()),
                NO_INTERVENIDOS_MES=("ESTADO_MES", lambda x: (x == "No intervenido").sum()),
                INTERVENCIONES_MES=("INTERVENCIONES_MES", "sum"),
                PRODUCCION_PETROLEO_MES=("PRODUCCION_PETROLEO_MES", "sum"),
                PRODUCCION_AGUA_MES=("PRODUCCION_AGUA_MES", "sum")
            )
            .sort_values("AÑO_CONVERSION")
        )

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Pozos convertidos listados", f"{tabla_conv['POZO'].nunique():,}")
        with c2:
            st.metric("Intervenidos en el mes", f"{(tabla_conv['ESTADO_MES'] == 'Intervenido').sum():,}")
        with c3:
            st.metric("No intervenidos en el mes", f"{(tabla_conv['ESTADO_MES'] == 'No intervenido').sum():,}")
        with c4:
            st.metric("Producción de petróleo mes", f"{tabla_conv['PRODUCCION_PETROLEO_MES'].sum():,.2f}")

        st.write("Resumen por año de conversión")
        st.dataframe(formatear_tabla(resumen_conv), use_container_width=True, hide_index=True)

        st.write("Listado de pozos convertidos")
        st.dataframe(formatear_tabla(tabla_conv), use_container_width=True, hide_index=True)

        st.download_button(
            "Descargar listado de convertidos en Excel",
            data=convertir_excel({"Listado convertidos": tabla_conv, "Resumen convertidos": resumen_conv}),
            file_name=f"listado_convertidos_{anio_objetivo}_{mes_objetivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


with tabs[3]:
    st.subheader("Potencial mensual por pozo")
    st.caption(
        "Criterio: último mes activo con producción de petróleo > 0. "
        "También se calcula el mes anterior al último mes activo para validar si el último mes fue poco representativo."
    )

    cols_pot = [
        "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION",
        "ULTIMO_MES_ACTIVO", "DIAS_MES", "PRCR_ULTIMO_MES",
        "POTENCIAL_ULTIMO_MES_BOPD", "MES_ANTERIOR_ULTIMO_ACTIVO",
        "DIAS_MES_ANTERIOR", "PRCR_MES_ANTERIOR",
        "POTENCIAL_MES_ANTERIOR_BOPD", "POTENCIAL_PROMEDIO_2_MESES_BOPD",
        "PRAG_ULTIMO_MES", "PRAG_MES_ANTERIOR", "INTERV_ULTIMO_MES",
        "INTERV_MES_ANTERIOR", "ULTIMA_FECHA_CON_PRCR"
    ]
    cols_pot = [c for c in cols_pot if c in potencial.columns]
    st.dataframe(formatear_tabla(potencial[cols_pot]), use_container_width=True, hide_index=True)

    top_pot = potencial.sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False).head(top_n)
    if not top_pot.empty:
        fig_pot = px.bar(
            top_pot,
            x="POZO",
            y=["POTENCIAL_ULTIMO_MES_BOPD", "POTENCIAL_MES_ANTERIOR_BOPD"],
            barmode="group",
            hover_data=["BATERIA", "CLASIFICACION", "ULTIMO_MES_ACTIVO", "MES_ANTERIOR_ULTIMO_ACTIVO"]
        )
        fig_pot = aplicar_layout_fig(fig_pot, f"Top {top_n} pozos por potencial estimado", 560)
        st.plotly_chart(fig_pot, use_container_width=True)

    st.write("Cruce: pozos dejados o reducidos con mayor potencial")
    cruce = desplazamiento[
        desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])
    ].sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False)

    cols_cruce = [
        "ESTADO_DESPLAZAMIENTO", "PRIORIDAD_REVISION", "POZO", "BATERIA",
        "CLASIFICACION", "INTERV_BASE", "INTERV_ACTUAL", "PRCR_BASE_NO_REALIZADO",
        "ULTIMO_MES_ACTIVO", "POTENCIAL_ULTIMO_MES_BOPD",
        "MES_ANTERIOR_ULTIMO_ACTIVO", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD"
    ]
    cols_cruce = [c for c in cols_cruce if c in cruce.columns]
    st.dataframe(formatear_tabla(cruce[cols_cruce]), use_container_width=True, hide_index=True)

    if not cruce.empty:
        fig_cruce = px.scatter(
            cruce.head(80),
            x="POTENCIAL_PROMEDIO_2_MESES_BOPD",
            y="PRCR_BASE_NO_REALIZADO",
            size="INTERV_BASE_NO_REALIZADAS",
            color="ESTADO_DESPLAZAMIENTO",
            hover_name="POZO",
            hover_data=[
                "BATERIA", "CLASIFICACION", "ULTIMO_MES_ACTIVO",
                "MES_ANTERIOR_ULTIMO_ACTIVO", "POTENCIAL_ULTIMO_MES_BOPD",
                "POTENCIAL_MES_ANTERIOR_BOPD"
            ]
        )
        fig_cruce = aplicar_layout_fig(fig_cruce, "Pozos dejados/reducidos: potencial promedio 2 meses vs producción de petróleo base no realizada", 580)
        st.plotly_chart(fig_cruce, use_container_width=True)

    st.download_button(
        "Descargar data de gráficas de potencial",
        data=convertir_excel({
            "Top potencial pozos": top_pot,
            "Cruce dejados potencial": cruce
        }),
        file_name=f"data_graficas_potencial_{anio_objetivo}_{mes_objetivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


with tabs[4]:
    st.subheader("Pozos y baterías")
    estado_pozo = st.radio("Estado", ["Todos", "Intervenido", "No intervenido"], horizontal=True)
    tabla_pozos = resumen_pozos.copy()
    if estado_pozo != "Todos":
        tabla_pozos = tabla_pozos[tabla_pozos["ESTADO"] == estado_pozo]

    cols_pozos = [
        "ESTADO", "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION",
        "TIPO_SWAB", "UNIDADES", "INTERVENCIONES", "PRCR", "PRAG",
        "OIL_POR_INTERV", "AGUA_POR_INTERV", "PRIMERA_FECHA",
        "ULTIMA_FECHA", "ULTIMA_FECHA_HISTORICA"
    ]
    cols_pozos = [c for c in cols_pozos if c in tabla_pozos.columns]
    st.dataframe(formatear_tabla(tabla_pozos[cols_pozos].sort_values(["ESTADO", "CLASIFICACION", "BATERIA", "POZO"])), use_container_width=True, hide_index=True)

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if not res_bateria.empty:
            fig_bat = px.bar(
                res_bateria.head(top_n),
                x="BATERIA",
                y="PRCR",
                hover_data=["POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES", "PRAG"],
                text="PRCR"
            )
            fig_bat = aplicar_layout_fig(fig_bat, f"Top {top_n} baterías por Producción de petróleo", 520)
            st.plotly_chart(fig_bat, use_container_width=True)
    with col_b2:
        if not res_bateria.empty:
            fig_bat2 = px.scatter(
                res_bateria.head(top_n),
                x="INTERVENCIONES",
                y="PRCR",
                size="POZOS_INTERVENIDOS",
                hover_name="BATERIA",
                hover_data=["PRAG", "POZOS_NO_INTERVENIDOS", "OIL_POR_INTERV"]
            )
            fig_bat2 = aplicar_layout_fig(fig_bat2, "Baterías: intervenciones vs PRCR", 520)
            st.plotly_chart(fig_bat2, use_container_width=True)

    if not tendencia.empty:
        orden_meses = [MESES[m] for m in sorted(tendencia["MES"].unique())]
        fig_tend = px.line(
            tendencia,
            x="MES_NOMBRE",
            y=["PRCR", "PRAG"],
            markers=True,
            category_orders={"MES_NOMBRE": orden_meses}
        )
        fig_tend = aplicar_layout_fig(fig_tend, f"Tendencia mensual Producción de petróleo y Producción de agua en {anio_objetivo}", 520)
        st.plotly_chart(fig_tend, use_container_width=True)


with tabs[5]:
    st.subheader("TS y CS")
    if res_tipo.empty:
        st.info("No hay información por tipo de swab.")
    else:
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            fig_tipo = px.bar(
                res_tipo,
                x="TIPO_SWAB",
                y=["PRCR", "PRAG"],
                barmode="group",
                hover_data=["POZOS", "INTERVENCIONES", "OIL_POR_INTERV"]
            )
            fig_tipo = aplicar_layout_fig(fig_tipo, "Producción de petróleo y Producción de agua por tipo de swab", 520)
            st.plotly_chart(fig_tipo, use_container_width=True)
        with col_t2:
            fig_pie = px.pie(
                res_tipo,
                names="TIPO_SWAB",
                values="INTERVENCIONES",
                hole=0.45
            )
            fig_pie = aplicar_layout_fig(fig_pie, "Distribución de intervenciones por tipo de swab", 520)
            st.plotly_chart(fig_pie, use_container_width=True)

        st.dataframe(formatear_tabla(res_tipo), use_container_width=True, hide_index=True)


with tabs[6]:
    st.subheader("Descargas")
    cols_pozos = [
        "ESTADO", "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION",
        "TIPO_SWAB", "UNIDADES", "INTERVENCIONES", "PRCR", "PRAG",
        "OIL_POR_INTERV", "AGUA_POR_INTERV", "PRIMERA_FECHA",
        "ULTIMA_FECHA", "ULTIMA_FECHA_HISTORICA"
    ]
    cols_pozos = [c for c in cols_pozos if c in resumen_pozos.columns]

    cols_desp = [
        "ESTADO_DESPLAZAMIENTO", "PRIORIDAD_REVISION", "POZO", "BATERIA",
        "CLASIFICACION", "ANIO_CONVERSION",
        "INTERV_BASE", "INTERV_ACTUAL", "INTERV_BASE_NO_REALIZADAS",
        "PRCR_BASE", "PRCR_ACTUAL", "PRCR_BASE_NO_REALIZADO",
        "PRCR_MES_ANTERIOR_OBJETIVO", "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR",
        "VAR_PETROLEO_VS_MES_ANTERIOR", "INTERV_MES_ANTERIOR_OBJETIVO",
        "INTERVENCIONES_DEJADAS_MES_ANTERIOR",
        "ULTIMO_MES_ACTIVO", "PRCR_ULTIMO_MES", "POTENCIAL_ULTIMO_MES_BOPD",
        "MES_ANTERIOR_ULTIMO_ACTIVO", "PRCR_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD",
        "PRAG_BASE", "PRAG_ACTUAL", "PRAG_BASE_NO_REALIZADO",
        "OIL_INTERV_BASE", "OIL_INTERV_ACTUAL", "ULTIMA_FECHA_HISTORICA"
    ]
    cols_desp = [c for c in cols_desp if c in desplazamiento.columns]

    cols_pot = [
        "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION",
        "ULTIMO_MES_ACTIVO", "DIAS_MES", "PRCR_ULTIMO_MES",
        "POTENCIAL_ULTIMO_MES_BOPD", "MES_ANTERIOR_ULTIMO_ACTIVO",
        "DIAS_MES_ANTERIOR", "PRCR_MES_ANTERIOR",
        "POTENCIAL_MES_ANTERIOR_BOPD", "POTENCIAL_PROMEDIO_2_MESES_BOPD",
        "PRAG_ULTIMO_MES", "PRAG_MES_ANTERIOR", "INTERV_ULTIMO_MES",
        "INTERV_MES_ANTERIOR", "ULTIMA_FECHA_CON_PRCR"
    ]

    tablas = {
        "Resumen pozos": resumen_pozos[cols_pozos],
        "Listado convertidos": listado_convertidos,
        "Analisis pozos dejados": desplazamiento[cols_desp],
        "Pozos ATA excluidos": df_ata_excluidos,
        "Resumen desplaz bateria": res_desplazamiento_bateria,
        "Potencial pozos": potencial[cols_pot],
        "Data graf dejados": desplazamiento[desplazamiento["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])],
        "Data graf potencial": potencial.sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False).head(top_n),
        "Resumen clasificacion": res_clase,
        "Resumen baterias": res_bateria,
        "Resumen TS CS": res_tipo,
        "Tendencia mensual": tendencia
    }

    excel_bytes = convertir_excel(tablas)
    st.download_button(
        "Descargar toda la información en Excel",
        data=excel_bytes,
        file_name=f"swab_servicio_potencial_{anio_objetivo}_{mes_objetivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    ppt_bytes = crear_ppt(
        periodo=periodo,
        kpis=kpis,
        resumen_clase=res_clase,
        resumen_bateria=res_bateria,
        resumen_tipo=res_tipo,
        tendencia=tendencia,
        tabla_desplazamiento=desplazamiento,
        potencial=potencial
    )

    st.download_button(
        "Descargar PPT editable",
        data=ppt_bytes,
        file_name=f"dashboard_swab_servicio_potencial_{anio_objetivo}_{mes_objetivo}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    st.subheader("Validación")
    validacion = pd.DataFrame({
        "Concepto": [
            "Hoja usada", "Rango mínimo del Excel", "Rango máximo del Excel",
            "Rango seleccionado desde", "Rango seleccionado hasta", "Registros analizados",
            "Pozos históricos en rango", "Baterías en rango", "Convertidos 2026 fijos",
            "Pozos candidatos ATA excluidos", "Registros ATA excluidos", "PRCR", "PRAG"
        ],
        "Valor": [
            hoja_usada, str(fecha_min), str(fecha_max), str(fecha_inicio_analisis), str(fecha_fin_analisis),
            f"{len(df):,}", f"{df['POZO_KEY'].nunique():,}", f"{df['BATERIA'].nunique():,}", "20",
            f"{pozos_ata_excluidos_en_data:,} de 88", f"{registros_ata_excluidos:,}",
            "Petróleo recuperado", "Agua recuperada"
        ]
    })
    st.dataframe(validacion, use_container_width=True, hide_index=True)
