import io
import re
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN


st.set_page_config(
    page_title="Dashboard SWAB Lote X",
    page_icon="🛢️",
    layout="wide"
)


# ============================================================
# POZOS CONVERTIDOS CONOCIDOS
# Si el Excel cargado no trae las hojas Swab 2024, 2025 o 2026,
# el sistema usa esta lista como respaldo.
# ============================================================

CONVERTIDOS_FALLBACK = {
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


def normalizar_columna(col):
    texto = str(col).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def encontrar_hoja(sheet_names, objetivo):
    objetivo = objetivo.strip().lower()
    for hoja in sheet_names:
        if str(hoja).strip().lower() == objetivo:
            return hoja
    return None


def convertir_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def normalizar_tipo_swab(valor):
    texto = limpiar_texto(valor)

    if texto in ["CS", "CSG", "CASING", "CASING SWAB"]:
        return "CS"

    if texto in ["TS", "TBG", "T BG", "TUBING", "TUBING SWAB", "TB"]:
        return "TS"

    if "CS" in texto or "CAS" in texto:
        return "CS"

    if "TB" in texto or "TBG" in texto or "TUB" in texto or texto == "TS":
        return "TS"

    return texto if texto else "SIN TIPO"


def periodo_descripcion(anio, mes):
    if mes == 0:
        return f"Todo el año {anio}"
    return f"{MESES[mes]} {anio}"


def detectar_fila_encabezado_convertidos(bytes_excel, hoja):
    preview = pd.read_excel(
        io.BytesIO(bytes_excel),
        sheet_name=hoja,
        header=None,
        nrows=10
    )

    for idx, row in preview.iterrows():
        valores = [str(v).strip().lower() for v in row.tolist()]
        if "fecha" in valores and "total" in valores:
            return idx

    return 1


# ============================================================
# CARGA DE DATOS
# ============================================================

@st.cache_data(show_spinner=False)
def cargar_datos_swab(bytes_excel):
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))
    hoja = encontrar_hoja(xls.sheet_names, "Datos de Swab")

    if hoja is None:
        raise ValueError("No encontré la hoja 'Datos de Swab'.")

    columnas_necesarias = ["FECHA", "COD_POZ", "COD_BAT", "UNIDAD", "TSER", "PRCR", "PRAG"]
    df = pd.read_excel(
        io.BytesIO(bytes_excel),
        sheet_name=hoja,
        usecols=lambda c: str(c).strip().upper() in columnas_necesarias
    )

    df.columns = [str(c).strip().upper() for c in df.columns]

    faltantes = [c for c in columnas_necesarias if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en Datos de Swab: {faltantes}")

    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df = df[df["FECHA"].notna()].copy()

    df["POZO"] = df["COD_POZ"].apply(mostrar_pozo)
    df["POZO_KEY"] = df["COD_POZ"].apply(limpiar_pozo)
    df["BATERIA"] = df["COD_BAT"].apply(limpiar_texto)
    df["UNIDAD"] = df["UNIDAD"].apply(limpiar_texto)
    df["TSER_ORIGINAL"] = df["TSER"].apply(limpiar_texto)
    df["TIPO_SWAB"] = df["TSER"].apply(normalizar_tipo_swab)
    df["PRCR"] = convertir_numero(df["PRCR"])
    df["PRAG"] = convertir_numero(df["PRAG"])

    df = df[df["POZO_KEY"] != ""].copy()
    df["ANIO"] = df["FECHA"].dt.year
    df["MES"] = df["FECHA"].dt.month
    df["MES_NOMBRE"] = df["MES"].map(MESES)
    df["PERIODO_MES"] = df["FECHA"].dt.to_period("M").dt.to_timestamp()

    return df


@st.cache_data(show_spinner=False)
def extraer_convertidos_desde_excel(bytes_excel):
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))
    registros = []

    for anio in [2024, 2025, 2026]:
        nombre_hoja = f"Swab {anio} (Dia)"
        hoja = encontrar_hoja(xls.sheet_names, nombre_hoja)

        if hoja is None:
            continue

        try:
            header_row = detectar_fila_encabezado_convertidos(bytes_excel, hoja)
            temp = pd.read_excel(io.BytesIO(bytes_excel), sheet_name=hoja, header=header_row, nrows=2)

            columnas = [str(c).strip() for c in temp.columns]
            columnas_pozo = columnas[3:]

            for pozo in columnas_pozo:
                if pozo == "" or pozo.lower().startswith("unnamed"):
                    continue

                registros.append({
                    "POZO": mostrar_pozo(pozo),
                    "POZO_KEY": limpiar_pozo(pozo),
                    "ANIO_CONVERSION": anio,
                    "CLASIFICACION": f"Convertido {anio}",
                    "FUENTE_CONVERSION": f"Hoja {nombre_hoja}"
                })
        except Exception:
            continue

    if not registros:
        return pd.DataFrame(columns=["POZO", "POZO_KEY", "ANIO_CONVERSION", "CLASIFICACION", "FUENTE_CONVERSION"])

    salida = pd.DataFrame(registros)
    salida = salida[salida["POZO_KEY"] != ""].drop_duplicates("POZO_KEY")
    return salida


def convertir_fallback_a_df():
    registros = []

    for anio, pozos in CONVERTIDOS_FALLBACK.items():
        for pozo in pozos:
            registros.append({
                "POZO": mostrar_pozo(pozo),
                "POZO_KEY": limpiar_pozo(pozo),
                "ANIO_CONVERSION": anio,
                "CLASIFICACION": f"Convertido {anio}",
                "FUENTE_CONVERSION": "Lista fija de respaldo"
            })

    salida = pd.DataFrame(registros)
    salida = salida.drop_duplicates("POZO_KEY")
    return salida


@st.cache_data(show_spinner=False)
def cargar_base_pozos(bytes_excel):
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))

    posibles_hojas = ["Swab Básica", "Swab (Batería)", "Swab 2024", "Estado Oct 2024"]

    bases = []

    for nombre in posibles_hojas:
        hoja = encontrar_hoja(xls.sheet_names, nombre)

        if hoja is None:
            continue

        try:
            temp = pd.read_excel(io.BytesIO(bytes_excel), sheet_name=hoja)
            temp.columns = [normalizar_columna(c) for c in temp.columns]

            col_pozo = None
            for posible in ["Pozo", "POZO", "P O Z O", "*PFORMACION"]:
                if posible in temp.columns:
                    col_pozo = posible
                    break

            if col_pozo is None:
                continue

            col_bat = None
            for posible in ["Battery", "BATERIA", "BATERÍA", "*BATERIA"]:
                if posible in temp.columns:
                    col_bat = posible
                    break

            base = pd.DataFrame()
            base["POZO"] = temp[col_pozo].apply(mostrar_pozo)
            base["POZO_KEY"] = temp[col_pozo].apply(limpiar_pozo)

            if col_bat is not None:
                base["BATERIA_BASE"] = temp[col_bat].apply(limpiar_texto)
            else:
                base["BATERIA_BASE"] = ""

            if "Potential (bopd)" in temp.columns:
                base["POTENCIAL_BOPD"] = convertir_numero(temp["Potential (bopd)"])
            elif "Potential (bopd)." in temp.columns:
                base["POTENCIAL_BOPD"] = convertir_numero(temp["Potential (bopd)."])
            else:
                base["POTENCIAL_BOPD"] = 0

            base = base[base["POZO_KEY"] != ""].copy()
            bases.append(base)

        except Exception:
            continue

    if not bases:
        return pd.DataFrame(columns=["POZO", "POZO_KEY", "BATERIA_BASE", "POTENCIAL_BOPD"])

    salida = pd.concat(bases, ignore_index=True)
    salida = salida.drop_duplicates("POZO_KEY")
    return salida


def combinar_convertidos(convertidos_excel, usar_respaldo=True):
    respaldo = convertir_fallback_a_df()

    if convertidos_excel.empty:
        return respaldo.copy() if usar_respaldo else convertidos_excel.copy()

    combinado = pd.concat([convertidos_excel, respaldo], ignore_index=True)
    combinado["PRIORIDAD"] = combinado["FUENTE_CONVERSION"].apply(
        lambda x: 1 if str(x).startswith("Hoja") else 2
    )
    combinado = combinado.sort_values(["POZO_KEY", "PRIORIDAD"])
    combinado = combinado.drop_duplicates("POZO_KEY", keep="first")
    combinado = combinado.drop(columns=["PRIORIDAD"])

    return combinado


def construir_universo(df, convertidos, base_pozos):
    historico = (
        df.sort_values("FECHA")
        .groupby("POZO_KEY", as_index=False)
        .agg(
            POZO_HIST=("POZO", "last"),
            BATERIA_HIST=("BATERIA", "last"),
            ULTIMA_FECHA_HIST=("FECHA", "max")
        )
    )

    universo = historico.copy()
    universo = universo.rename(columns={
        "POZO_HIST": "POZO",
        "BATERIA_HIST": "BATERIA"
    })

    if not base_pozos.empty:
        universo = universo.merge(
            base_pozos[["POZO_KEY", "POZO", "BATERIA_BASE", "POTENCIAL_BOPD"]],
            on="POZO_KEY",
            how="outer",
            suffixes=("", "_BASE")
        )

        universo["POZO_FINAL"] = universo["POZO"]
        universo.loc[universo["POZO_FINAL"].isna() | (universo["POZO_FINAL"] == ""), "POZO_FINAL"] = universo["POZO_BASE"]

        universo["BATERIA_FINAL"] = universo["BATERIA"]
        universo.loc[
            universo["BATERIA_FINAL"].isna() | (universo["BATERIA_FINAL"] == ""),
            "BATERIA_FINAL"
        ] = universo["BATERIA_BASE"]

        universo["POZO"] = universo["POZO_FINAL"]
        universo["BATERIA"] = universo["BATERIA_FINAL"]

        universo = universo.drop(columns=[
            c for c in ["POZO_BASE", "POZO_FINAL", "BATERIA_BASE", "BATERIA_FINAL"]
            if c in universo.columns
        ])
    else:
        universo["POTENCIAL_BOPD"] = 0

    universo = universo.merge(
        convertidos[["POZO_KEY", "ANIO_CONVERSION", "CLASIFICACION", "FUENTE_CONVERSION"]],
        on="POZO_KEY",
        how="left"
    )

    universo["ANIO_CONVERSION"] = universo["ANIO_CONVERSION"].fillna(0).astype(int)
    universo["CLASIFICACION"] = universo["CLASIFICACION"].fillna("Básica")
    universo["FUENTE_CONVERSION"] = universo["FUENTE_CONVERSION"].fillna("No convertido")

    universo["POZO"] = universo["POZO"].fillna("")
    universo["BATERIA"] = universo["BATERIA"].fillna("")
    universo["POTENCIAL_BOPD"] = universo["POTENCIAL_BOPD"].fillna(0)

    universo = universo[universo["POZO_KEY"] != ""].drop_duplicates("POZO_KEY")

    return universo


# ============================================================
# ANÁLISIS
# ============================================================

def aplicar_filtros_base(universo, baterias_sel, clases_sel):
    salida = universo.copy()

    if baterias_sel:
        salida = salida[salida["BATERIA"].isin(baterias_sel)]

    if clases_sel:
        salida = salida[salida["CLASIFICACION"].isin(clases_sel)]

    return salida


def filtrar_periodo(df, anio, mes, baterias_sel, tipos_sel, clases_sel, universo):
    data = df[df["ANIO"] == anio].copy()

    if mes != 0:
        data = data[data["MES"] == mes].copy()

    if baterias_sel:
        data = data[data["BATERIA"].isin(baterias_sel)]

    if tipos_sel:
        data = data[data["TIPO_SWAB"].isin(tipos_sel)]

    mapa_clase = universo[["POZO_KEY", "CLASIFICACION"]].drop_duplicates("POZO_KEY")
    data = data.merge(mapa_clase, on="POZO_KEY", how="left")
    data["CLASIFICACION"] = data["CLASIFICACION"].fillna("Básica")

    if clases_sel:
        data = data[data["CLASIFICACION"].isin(clases_sel)]

    return data


def resumir_periodo(data_periodo, universo_filtrado):
    if data_periodo.empty:
        resumen = universo_filtrado.copy()
        resumen["INTERVENCIONES"] = 0
        resumen["PRCR"] = 0.0
        resumen["PRAG"] = 0.0
        resumen["OIL_POR_INTERV"] = 0.0
        resumen["AGUA_POR_INTERV"] = 0.0
        resumen["TIPO_SWAB"] = ""
        resumen["UNIDADES"] = ""
        resumen["PRIMERA_FECHA"] = pd.NaT
        resumen["ULTIMA_FECHA"] = pd.NaT
        resumen["ESTADO_MES"] = "No intervenido"
        return resumen

    agg = (
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

    resumen = universo_filtrado.merge(agg, on="POZO_KEY", how="left")

    resumen["POZO"] = resumen["POZO"].fillna(resumen["POZO_REAL"])
    resumen["BATERIA"] = resumen["BATERIA"].fillna(resumen["BATERIA_REAL"])

    for col in ["INTERVENCIONES", "PRCR", "PRAG"]:
        resumen[col] = resumen[col].fillna(0)

    resumen["INTERVENCIONES"] = resumen["INTERVENCIONES"].astype(int)
    resumen["OIL_POR_INTERV"] = np.where(
        resumen["INTERVENCIONES"] > 0,
        resumen["PRCR"] / resumen["INTERVENCIONES"],
        0
    )
    resumen["AGUA_POR_INTERV"] = np.where(
        resumen["INTERVENCIONES"] > 0,
        resumen["PRAG"] / resumen["INTERVENCIONES"],
        0
    )
    resumen["TIPO_SWAB"] = resumen["TIPO_SWAB"].fillna("")
    resumen["UNIDADES"] = resumen["UNIDADES"].fillna("")
    resumen["ESTADO_MES"] = np.where(resumen["INTERVENCIONES"] > 0, "Intervenido", "No intervenido")

    columnas_drop = [c for c in ["POZO_REAL", "BATERIA_REAL"] if c in resumen.columns]
    resumen = resumen.drop(columns=columnas_drop)

    return resumen


def resumen_por_bateria(data_periodo, universo_filtrado):
    universo_bat = (
        universo_filtrado
        .groupby("BATERIA", as_index=False)
        .agg(POZOS_UNIVERSO=("POZO_KEY", "nunique"))
    )

    if data_periodo.empty:
        universo_bat["POZOS_INTERVENIDOS"] = 0
        universo_bat["INTERVENCIONES"] = 0
        universo_bat["PRCR"] = 0.0
        universo_bat["PRAG"] = 0.0
        universo_bat["POZOS_NO_INTERVENIDOS"] = universo_bat["POZOS_UNIVERSO"]
        return universo_bat

    prod = (
        data_periodo
        .groupby("BATERIA", as_index=False)
        .agg(
            POZOS_INTERVENIDOS=("POZO_KEY", "nunique"),
            INTERVENCIONES=("FECHA", "count"),
            PRCR=("PRCR", "sum"),
            PRAG=("PRAG", "sum")
        )
    )

    salida = universo_bat.merge(prod, on="BATERIA", how="left")

    for col in ["POZOS_INTERVENIDOS", "INTERVENCIONES", "PRCR", "PRAG"]:
        salida[col] = salida[col].fillna(0)

    salida["POZOS_NO_INTERVENIDOS"] = salida["POZOS_UNIVERSO"] - salida["POZOS_INTERVENIDOS"]
    salida["OIL_POR_INTERV"] = np.where(
        salida["INTERVENCIONES"] > 0,
        salida["PRCR"] / salida["INTERVENCIONES"],
        0
    )

    return salida.sort_values("PRCR", ascending=False)


def resumen_por_tipo(data_periodo):
    if data_periodo.empty:
        return pd.DataFrame(columns=["TIPO_SWAB", "INTERVENCIONES", "POZOS", "PRCR", "PRAG", "OIL_POR_INTERV"])

    salida = (
        data_periodo
        .groupby("TIPO_SWAB", as_index=False)
        .agg(
            INTERVENCIONES=("FECHA", "count"),
            POZOS=("POZO_KEY", "nunique"),
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


def resumen_por_clase(resumen_pozos):
    salida = (
        resumen_pozos
        .groupby("CLASIFICACION", as_index=False)
        .agg(
            POZOS=("POZO_KEY", "nunique"),
            POZOS_INTERVENIDOS=("ESTADO_MES", lambda x: (x == "Intervenido").sum()),
            POZOS_NO_INTERVENIDOS=("ESTADO_MES", lambda x: (x == "No intervenido").sum()),
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


def tendencia_mensual(df, anio, universo, baterias_sel, clases_sel, tipos_sel):
    data = df[df["ANIO"] == anio].copy()

    if baterias_sel:
        data = data[data["BATERIA"].isin(baterias_sel)]

    if tipos_sel:
        data = data[data["TIPO_SWAB"].isin(tipos_sel)]

    mapa_clase = universo[["POZO_KEY", "CLASIFICACION"]].drop_duplicates("POZO_KEY")
    data = data.merge(mapa_clase, on="POZO_KEY", how="left")
    data["CLASIFICACION"] = data["CLASIFICACION"].fillna("Básica")

    if clases_sel:
        data = data[data["CLASIFICACION"].isin(clases_sel)]

    if data.empty:
        return pd.DataFrame(columns=["MES", "MES_NOMBRE", "INTERVENCIONES", "POZOS", "PRCR", "PRAG", "OIL_POR_INTERV"])

    salida = (
        data
        .groupby(["MES", "MES_NOMBRE"], as_index=False)
        .agg(
            INTERVENCIONES=("FECHA", "count"),
            POZOS=("POZO_KEY", "nunique"),
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


def top_no_intervenidos(resumen_pozos):
    salida = resumen_pozos[resumen_pozos["ESTADO_MES"] == "No intervenido"].copy()
    salida = salida.sort_values(["CLASIFICACION", "BATERIA", "POZO"])
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
# DESCARGAS
# ============================================================

def crear_excel_resultados(tablas):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for nombre, tabla in tablas.items():
            nombre_hoja = nombre[:31]
            formatear_tabla(tabla).to_excel(writer, sheet_name=nombre_hoja, index=False)

            workbook = writer.book
            worksheet = writer.sheets[nombre_hoja]

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#D9EAF7",
                "border": 1
            })

            for col_num, value in enumerate(formatear_tabla(tabla).columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 16)

    buffer.seek(0)
    return buffer.getvalue()


def agregar_titulo(slide, titulo, subtitulo=None):
    title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(12.5), Inches(0.5))
    tf = title_box.text_frame
    tf.text = titulo
    tf.paragraphs[0].font.size = Pt(22)
    tf.paragraphs[0].font.bold = True

    if subtitulo:
        sub_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.75), Inches(12.5), Inches(0.35))
        tf2 = sub_box.text_frame
        tf2.text = subtitulo
        tf2.paragraphs[0].font.size = Pt(11)


def agregar_tabla(slide, df, x, y, w, h, font_size=8):
    if df.empty:
        box = slide.shapes.add_textbox(x, y, w, h)
        box.text_frame.text = "Sin datos para mostrar"
        return

    df_show = formatear_tabla(df).head(12)
    rows = len(df_show) + 1
    cols = len(df_show.columns)

    table_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    table = table_shape.table

    for j, col in enumerate(df_show.columns):
        cell = table.cell(0, j)
        cell.text = str(col)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(font_size)

    for i, (_, row) in enumerate(df_show.iterrows(), start=1):
        for j, col in enumerate(df_show.columns):
            cell = table.cell(i, j)
            cell.text = "" if pd.isna(row[col]) else str(row[col])
            cell.text_frame.paragraphs[0].font.size = Pt(font_size)


def agregar_grafico_barras(slide, titulo, categorias, valores, x, y, w, h, nombre_serie="Valor"):
    categorias = list(categorias)
    valores = [0 if pd.isna(v) else float(v) for v in valores]

    if not categorias:
        return

    chart_data = CategoryChartData()
    chart_data.categories = categorias
    chart_data.add_series(nombre_serie, valores)

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        x,
        y,
        w,
        h,
        chart_data
    ).chart

    chart.has_legend = False
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = titulo
    chart.value_axis.has_major_gridlines = True


def agregar_grafico_linea(slide, titulo, categorias, serie_1, serie_2, x, y, w, h):
    categorias = list(categorias)

    if not categorias:
        return

    chart_data = CategoryChartData()
    chart_data.categories = categorias
    chart_data.add_series("PRCR", [0 if pd.isna(v) else float(v) for v in serie_1])
    chart_data.add_series("PRAG", [0 if pd.isna(v) else float(v) for v in serie_2])

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        x,
        y,
        w,
        h,
        chart_data
    ).chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.chart_title.has_text_frame = True
    chart.chart_title.text_frame.text = titulo


def crear_ppt_resultados(periodo_txt, kpis, resumen_pozos, resumen_bateria, resumen_tipo, resumen_clase, tendencia):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank = prs.slide_layouts[6]

    # Slide 1
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Dashboard SWAB Lote X", periodo_txt)

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

    agregar_tabla(slide, kpi_df, Inches(0.5), Inches(1.3), Inches(5.1), Inches(3.0), font_size=11)

    estado_df = pd.DataFrame({
        "Estado": ["Intervenido", "No intervenido"],
        "Pozos": [kpis["pozos_intervenidos"], kpis["pozos_no_intervenidos"]]
    })

    agregar_grafico_barras(
        slide,
        "Pozos intervenidos vs no intervenidos",
        estado_df["Estado"],
        estado_df["Pozos"],
        Inches(6.1),
        Inches(1.2),
        Inches(6.6),
        Inches(4.8),
        "Pozos"
    )

    # Slide 2
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Producción por tipo de pozo", periodo_txt)

    clase_plot = resumen_clase.sort_values("PRCR", ascending=False).head(8)
    agregar_grafico_barras(
        slide,
        "PRCR por clasificación",
        clase_plot["CLASIFICACION"],
        clase_plot["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(6.1),
        Inches(5.4),
        "PRCR"
    )

    agregar_tabla(
        slide,
        resumen_clase[["CLASIFICACION", "POZOS", "POZOS_INTERVENIDOS", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"]],
        Inches(6.8),
        Inches(1.2),
        Inches(6.1),
        Inches(5.4),
        font_size=7
    )

    # Slide 3
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Top pozos por producción de petróleo", periodo_txt)

    top_pozos = resumen_pozos[resumen_pozos["INTERVENCIONES"] > 0].sort_values("PRCR", ascending=False).head(15)
    agregar_grafico_barras(
        slide,
        "Top 15 pozos por PRCR",
        top_pozos["POZO"],
        top_pozos["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(7.0),
        Inches(5.5),
        "PRCR"
    )

    agregar_tabla(
        slide,
        top_pozos[["POZO", "BATERIA", "CLASIFICACION", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"]],
        Inches(7.7),
        Inches(1.2),
        Inches(5.2),
        Inches(5.5),
        font_size=7
    )

    # Slide 4
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Top baterías por producción", periodo_txt)

    top_bat = resumen_bateria.sort_values("PRCR", ascending=False).head(15)
    agregar_grafico_barras(
        slide,
        "Top 15 baterías por PRCR",
        top_bat["BATERIA"],
        top_bat["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(7.0),
        Inches(5.5),
        "PRCR"
    )

    agregar_tabla(
        slide,
        top_bat[["BATERIA", "POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"]],
        Inches(7.7),
        Inches(1.2),
        Inches(5.2),
        Inches(5.5),
        font_size=7
    )

    # Slide 5
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Producción por tipo de swab", periodo_txt)

    tipo_plot = resumen_tipo.sort_values("PRCR", ascending=False)
    agregar_grafico_barras(
        slide,
        "PRCR por TS o CS",
        tipo_plot["TIPO_SWAB"],
        tipo_plot["PRCR"],
        Inches(0.5),
        Inches(1.2),
        Inches(6.2),
        Inches(5.4),
        "PRCR"
    )

    agregar_tabla(
        slide,
        resumen_tipo[["TIPO_SWAB", "POZOS", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"]],
        Inches(7.0),
        Inches(1.2),
        Inches(5.8),
        Inches(4.5),
        font_size=9
    )

    # Slide 6
    slide = prs.slides.add_slide(blank)
    agregar_titulo(slide, "Tendencia mensual del año seleccionado", periodo_txt)

    if not tendencia.empty:
        agregar_grafico_linea(
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
    agregar_titulo(slide, "Pozos no intervenidos", periodo_txt)

    no_interv = resumen_pozos[resumen_pozos["ESTADO_MES"] == "No intervenido"].copy()
    no_interv = no_interv.sort_values(["CLASIFICACION", "BATERIA", "POZO"]).head(20)

    agregar_tabla(
        slide,
        no_interv[["POZO", "BATERIA", "CLASIFICACION", "POTENCIAL_BOPD", "ESTADO_MES"]],
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
# APP
# ============================================================

st.title("🛢️ Dashboard operativo SWAB Lote X")
st.caption("Carga tu Excel, selecciona año y mes, presiona ejecutar y revisa pozos intervenidos, no intervenidos, PRCR, PRAG, TS, CS y descargas.")

col_up1, col_up2 = st.columns(2)

with col_up1:
    archivo_principal = st.file_uploader(
        "Excel principal con hoja Datos de Swab",
        type=["xlsx"],
        key="archivo_principal"
    )

with col_up2:
    archivo_base = st.file_uploader(
        "Excel adicional opcional con convertidos y base, por ejemplo Producción por Swab.xlsx",
        type=["xlsx"],
        key="archivo_base"
    )

if archivo_principal is None:
    st.info("Primero sube el Excel principal. Puede ser Data Swab Python.xlsx o Producción por Swab.xlsx.")
    st.stop()

bytes_principal = archivo_principal.getvalue()
bytes_base = archivo_base.getvalue() if archivo_base is not None else None

try:
    df = cargar_datos_swab(bytes_principal)

    convertidos_principal = extraer_convertidos_desde_excel(bytes_principal)

    if bytes_base is not None:
        convertidos_base = extraer_convertidos_desde_excel(bytes_base)
        base_pozos = cargar_base_pozos(bytes_base)
        convertidos_excel = pd.concat([convertidos_principal, convertidos_base], ignore_index=True)
        convertidos_excel = convertidos_excel.drop_duplicates("POZO_KEY")
    else:
        base_pozos = cargar_base_pozos(bytes_principal)
        convertidos_excel = convertidos_principal

    convertidos = combinar_convertidos(convertidos_excel, usar_respaldo=True)
    universo = construir_universo(df, convertidos, base_pozos)

except Exception as e:
    st.error(f"No se pudo procesar el Excel: {e}")
    st.stop()

fecha_min = df["FECHA"].min().date()
fecha_max = df["FECHA"].max().date()

anios_disponibles = sorted(df["ANIO"].dropna().unique().astype(int).tolist())
baterias_disponibles = sorted([b for b in universo["BATERIA"].dropna().unique().tolist() if b != ""])
tipos_disponibles = sorted([t for t in df["TIPO_SWAB"].dropna().unique().tolist() if t != ""])
clases_disponibles = ["Básica", "Convertido 2024", "Convertido 2025", "Convertido 2026"]

with st.sidebar:
    st.header("Filtros")

    anio_sel = st.selectbox(
        "Año",
        anios_disponibles,
        index=len(anios_disponibles) - 1
    )

    mes_opciones = [0] + list(range(1, 13))
    mes_sel = st.selectbox(
        "Mes",
        mes_opciones,
        format_func=lambda x: "Todo el año" if x == 0 else MESES[x],
        index=mes_opciones.index(fecha_max.month) if anio_sel == fecha_max.year else 0
    )

    baterias_sel = st.multiselect(
        "Batería",
        baterias_disponibles,
        default=[]
    )

    tipos_sel = st.multiselect(
        "Tipo de swab",
        tipos_disponibles,
        default=[]
    )

    clases_sel = st.multiselect(
        "Tipo de pozo",
        clases_disponibles,
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

if not ejecutar and "resultados_swab" not in st.session_state:
    st.warning("Selecciona los filtros en la izquierda y presiona el botón Ejecutar análisis.")
    st.stop()

if ejecutar:
    universo_filtrado = aplicar_filtros_base(universo, baterias_sel, clases_sel)
    data_periodo = filtrar_periodo(df, anio_sel, mes_sel, baterias_sel, tipos_sel, clases_sel, universo)
    resumen_pozos = resumir_periodo(data_periodo, universo_filtrado)
    resumen_bateria = resumen_por_bateria(data_periodo, universo_filtrado)
    resumen_tipo = resumen_por_tipo(data_periodo)
    resumen_clase = resumen_por_clase(resumen_pozos)
    tendencia = tendencia_mensual(df, anio_sel, universo, baterias_sel, clases_sel, tipos_sel)

    st.session_state["resultados_swab"] = {
        "anio": anio_sel,
        "mes": mes_sel,
        "universo_filtrado": universo_filtrado,
        "data_periodo": data_periodo,
        "resumen_pozos": resumen_pozos,
        "resumen_bateria": resumen_bateria,
        "resumen_tipo": resumen_tipo,
        "resumen_clase": resumen_clase,
        "tendencia": tendencia,
        "filtros": {
            "baterias": baterias_sel,
            "tipos": tipos_sel,
            "clases": clases_sel
        }
    }

res = st.session_state["resultados_swab"]

anio_sel = res["anio"]
mes_sel = res["mes"]
data_periodo = res["data_periodo"]
resumen_pozos = res["resumen_pozos"]
resumen_bateria = res["resumen_bateria"]
resumen_tipo = res["resumen_tipo"]
resumen_clase = res["resumen_clase"]
tendencia = res["tendencia"]
universo_filtrado = res["universo_filtrado"]

periodo_txt = periodo_descripcion(anio_sel, mes_sel)

pozos_universo = int(resumen_pozos["POZO_KEY"].nunique())
pozos_intervenidos = int((resumen_pozos["ESTADO_MES"] == "Intervenido").sum())
pozos_no_intervenidos = int((resumen_pozos["ESTADO_MES"] == "No intervenido").sum())
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

st.subheader(f"Estatus operativo: {periodo_txt}")

c1, c2, c3, c4, c5, c6 = st.columns(6)

with c1:
    st.metric("Pozos universo", f"{pozos_universo:,}")

with c2:
    st.metric("Pozos intervenidos", f"{pozos_intervenidos:,}")

with c3:
    st.metric("Pozos no intervenidos", f"{pozos_no_intervenidos:,}")

with c4:
    st.metric("Intervenciones", f"{intervenciones:,}")

with c5:
    st.metric("PRCR", f"{prcr_total:,.2f}")

with c6:
    st.metric("Oil / Interv.", f"{oil_interv:,.2f}")

st.caption(
    f"Rango de datos cargado: {fecha_min} al {fecha_max}. "
    f"El análisis considera PRCR como petróleo recuperado y PRAG como agua recuperada."
)

st.divider()

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Estatus y listados",
    "Performance por pozo",
    "Baterías",
    "TS y CS",
    "Tendencia mensual",
    "Descargas y validación"
])


with tab1:
    st.subheader("Listado completo de pozos intervenidos y no intervenidos")

    filtro_estado = st.radio(
        "Ver",
        ["Todos", "Solo intervenidos", "Solo no intervenidos"],
        horizontal=True
    )

    tabla_estado = resumen_pozos.copy()

    if filtro_estado == "Solo intervenidos":
        tabla_estado = tabla_estado[tabla_estado["ESTADO_MES"] == "Intervenido"]
    elif filtro_estado == "Solo no intervenidos":
        tabla_estado = tabla_estado[tabla_estado["ESTADO_MES"] == "No intervenido"]

    columnas_listado = [
        "ESTADO_MES",
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
        "POTENCIAL_BOPD",
        "PRIMERA_FECHA",
        "ULTIMA_FECHA"
    ]

    columnas_listado = [c for c in columnas_listado if c in tabla_estado.columns]

    st.dataframe(
        formatear_tabla(tabla_estado[columnas_listado].sort_values(["ESTADO_MES", "CLASIFICACION", "BATERIA", "POZO"])),
        use_container_width=True,
        hide_index=True
    )

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Pozos por estado")
        estado_plot = pd.DataFrame({
            "Estado": ["Intervenido", "No intervenido"],
            "Pozos": [pozos_intervenidos, pozos_no_intervenidos]
        })
        fig_estado = px.bar(estado_plot, x="Estado", y="Pozos", text="Pozos")
        st.plotly_chart(fig_estado, use_container_width=True)

    with col_b:
        st.subheader("Resumen por tipo de pozo")
        st.dataframe(
            formatear_tabla(resumen_clase),
            use_container_width=True,
            hide_index=True
        )


with tab2:
    st.subheader("Performance por pozo")

    top_pozos = resumen_pozos[resumen_pozos["INTERVENCIONES"] > 0].sort_values("PRCR", ascending=False).head(top_n)

    if top_pozos.empty:
        st.info("No hay pozos intervenidos para el periodo seleccionado.")
    else:
        fig_top_pozos = px.bar(
            top_pozos,
            x="POZO",
            y="PRCR",
            color="CLASIFICACION",
            hover_data=["BATERIA", "INTERVENCIONES", "PRAG", "OIL_POR_INTERV"],
            title=f"Top {top_n} pozos por PRCR"
        )
        st.plotly_chart(fig_top_pozos, use_container_width=True)

        fig_oil_interv = px.bar(
            top_pozos.sort_values("OIL_POR_INTERV", ascending=False),
            x="POZO",
            y="OIL_POR_INTERV",
            color="CLASIFICACION",
            hover_data=["BATERIA", "INTERVENCIONES", "PRCR", "PRAG"],
            title=f"Top {top_n} pozos por oil/intervención"
        )
        st.plotly_chart(fig_oil_interv, use_container_width=True)

        st.dataframe(
            formatear_tabla(top_pozos[columnas_listado]),
            use_container_width=True,
            hide_index=True
        )


with tab3:
    st.subheader("Análisis por batería")

    if resumen_bateria.empty:
        st.info("No hay datos de baterías para el filtro seleccionado.")
    else:
        col_b1, col_b2 = st.columns(2)

        with col_b1:
            top_bat_prcr = resumen_bateria.sort_values("PRCR", ascending=False).head(top_n)
            fig_bat_prcr = px.bar(
                top_bat_prcr,
                x="BATERIA",
                y="PRCR",
                hover_data=["POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES", "PRAG"],
                title=f"Top {top_n} baterías por PRCR"
            )
            st.plotly_chart(fig_bat_prcr, use_container_width=True)

        with col_b2:
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
        st.info("No hay datos por tipo de swab para el filtro seleccionado.")
    else:
        col_t1, col_t2 = st.columns(2)

        with col_t1:
            fig_tipo_prcr = px.bar(
                resumen_tipo,
                x="TIPO_SWAB",
                y="PRCR",
                hover_data=["POZOS", "INTERVENCIONES", "PRAG", "OIL_POR_INTERV"],
                title="PRCR por tipo de swab"
            )
            st.plotly_chart(fig_tipo_prcr, use_container_width=True)

        with col_t2:
            fig_tipo_interv = px.bar(
                resumen_tipo,
                x="TIPO_SWAB",
                y="INTERVENCIONES",
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
        st.info("No hay tendencia mensual para el filtro seleccionado.")
    else:
        fig_tend = px.line(
            tendencia,
            x="MES_NOMBRE",
            y=["PRCR", "PRAG"],
            markers=True,
            title=f"Tendencia mensual PRCR y PRAG en {anio_sel}"
        )
        st.plotly_chart(fig_tend, use_container_width=True)

        fig_interv_mes = px.bar(
            tendencia,
            x="MES_NOMBRE",
            y="INTERVENCIONES",
            hover_data=["POZOS", "PRCR", "PRAG", "OIL_POR_INTERV"],
            title=f"Intervenciones mensuales en {anio_sel}"
        )
        st.plotly_chart(fig_interv_mes, use_container_width=True)

        st.dataframe(
            formatear_tabla(tendencia),
            use_container_width=True,
            hide_index=True
        )


with tab6:
    st.subheader("Descargar resultados")

    tablas_descarga = {
        "Resumen pozos": resumen_pozos[columnas_listado],
        "Intervenidos": resumen_pozos[resumen_pozos["ESTADO_MES"] == "Intervenido"][columnas_listado],
        "No intervenidos": resumen_pozos[resumen_pozos["ESTADO_MES"] == "No intervenido"][columnas_listado],
        "Resumen baterias": resumen_bateria,
        "Resumen TS CS": resumen_tipo,
        "Resumen clasificacion": resumen_clase,
        "Tendencia mensual": tendencia
    }

    excel_bytes = crear_excel_resultados(tablas_descarga)

    st.download_button(
        "Descargar toda la información en Excel",
        data=excel_bytes,
        file_name=f"resultado_swab_{anio_sel}_{mes_sel}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    ppt_bytes = crear_ppt_resultados(
        periodo_txt=periodo_txt,
        kpis=kpis,
        resumen_pozos=resumen_pozos,
        resumen_bateria=resumen_bateria,
        resumen_tipo=resumen_tipo,
        resumen_clase=resumen_clase,
        tendencia=tendencia
    )

    st.download_button(
        "Descargar presentación PPT editable",
        data=ppt_bytes,
        file_name=f"dashboard_swab_{anio_sel}_{mes_sel}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    st.subheader("Validación del archivo cargado")

    val_df = pd.DataFrame({
        "Concepto": [
            "Fecha mínima",
            "Fecha máxima",
            "Registros cargados",
            "Pozos históricos",
            "Baterías históricas",
            "Pozos convertidos 2024",
            "Pozos convertidos 2025",
            "Pozos convertidos 2026",
            "Pozos universo del análisis"
        ],
        "Valor": [
            str(fecha_min),
            str(fecha_max),
            f"{len(df):,}",
            f"{df['POZO_KEY'].nunique():,}",
            f"{df['BATERIA'].nunique():,}",
            f"{(convertidos['ANIO_CONVERSION'] == 2024).sum():,}",
            f"{(convertidos['ANIO_CONVERSION'] == 2025).sum():,}",
            f"{(convertidos['ANIO_CONVERSION'] == 2026).sum():,}",
            f"{pozos_universo:,}"
        ]
    })

    st.dataframe(val_df, use_container_width=True, hide_index=True)

    st.write("Muestra de datos limpios")
    muestra_cols = ["FECHA", "POZO", "BATERIA", "UNIDAD", "TIPO_SWAB", "PRCR", "PRAG", "CLASIFICACION"]
    muestra = df.merge(universo[["POZO_KEY", "CLASIFICACION"]], on="POZO_KEY", how="left")
    st.dataframe(
        formatear_tabla(muestra[muestra_cols].head(50)),
        use_container_width=True,
        hide_index=True
    )
