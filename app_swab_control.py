import io
import re
import calendar

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="SWAB Lote X - Servicio y Potencial",
    page_icon="🛢️",
    layout="wide"
)

# ============================================================
# CONFIGURACIÓN FIJA
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

CANDIDATOS_ATA = [
    "AA37", "AA54", "AA76", "AA112", "AA1577", "AA1598", "AA1599", "AA1633",
    "AA1661", "AA1847", "AA1930", "AA5631", "AA5707", "AA5861", "AA5926",
    "AA5971", "AA6192", "AA6338", "AA6342", "AA6372", "AA6423", "AA6454",
    "AA6517", "AA6646", "AA6762", "AA7201", "AA9154", "AA9329", "AA9364",
    "AA10013", "EA216", "EA264", "EA364", "EA440", "EA741", "EA771", "EA876",
    "EA888", "EA987", "EA1054", "EA1081", "EA1161", "EA1167", "EA1233",
    "EA1302", "EA1506", "EA1511", "EA1513", "EA1581", "EA1630", "EA1885",
    "EA2067", "EA2249", "EA2254", "EA2256", "EA2304", "EA2372", "EA2389",
    "EA2403", "EA5682D", "EA5694", "EA5739", "EA5766", "EA5868", "EA5874",
    "EA5914", "EA5921", "EA5957", "EA6130", "EA6237", "EA6918", "EA7027",
    "EA7158", "EA8574", "EA9242", "EA9251", "EA9287", "EA9409", "EA9417",
    "EA9491", "EA9668", "EA9752", "EA9779", "EA11128", "PB47", "PB232",
    "PE171", "PT4-3"
]

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

LABELS = {
    "POZO": "Pozo",
    "BATERIA": "Batería",
    "CLASIFICACION": "Clasificación",
    "ANIO_CONVERSION": "Año conversión",
    "ESTADO_DESPLAZAMIENTO": "Estado de desplazamiento",
    "PRIORIDAD_REVISION": "Prioridad de revisión",
    "INTERV_BASE": "Intervenciones base promedio",
    "INTERV_ACTUAL": "Intervenciones mes objetivo",
    "INTERV_BASE_NO_REALIZADAS": "Intervenciones base no realizadas",
    "PRCR_BASE": "Producción de petróleo base",
    "PRCR_ACTUAL": "Producción de petróleo actual",
    "PRCR_BASE_NO_REALIZADO": "Producción de petróleo base no realizada",
    "PRAG_BASE": "Producción de agua base",
    "PRAG_ACTUAL": "Producción de agua actual",
    "PRAG_BASE_NO_REALIZADO": "Producción de agua base no realizada",
    "PRCR_MES_ANTERIOR_OBJETIVO": "Producción de petróleo mes anterior al objetivo",
    "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR": "Producción de petróleo dejada vs mes anterior",
    "VAR_PETROLEO_VS_MES_ANTERIOR": "Variación petróleo vs mes anterior",
    "INTERV_MES_ANTERIOR_OBJETIVO": "Intervenciones mes anterior al objetivo",
    "INTERVENCIONES_DEJADAS_MES_ANTERIOR": "Intervenciones dejadas vs mes anterior",
    "OIL_INTERV_BASE": "Petróleo por intervención base",
    "OIL_INTERV_ACTUAL": "Petróleo por intervención actual",
    "ULTIMO_MES_ACTIVO": "Último mes activo",
    "DIAS_MES": "Días del mes",
    "PRCR_ULTIMO_MES": "Producción de petróleo último mes activo",
    "PRAG_ULTIMO_MES": "Producción de agua último mes activo",
    "INTERV_ULTIMO_MES": "Intervenciones último mes activo",
    "POTENCIAL_ULTIMO_MES_BOPD": "Potencial último mes activo bopd",
    "MES_ANTERIOR_ULTIMO_ACTIVO": "Mes anterior al último activo",
    "DIAS_MES_ANTERIOR": "Días mes anterior al último activo",
    "PRCR_MES_ANTERIOR": "Producción de petróleo mes anterior al último activo",
    "PRAG_MES_ANTERIOR": "Producción de agua mes anterior al último activo",
    "INTERV_MES_ANTERIOR": "Intervenciones mes anterior al último activo",
    "POTENCIAL_MES_ANTERIOR_BOPD": "Potencial mes anterior bopd",
    "POTENCIAL_PROMEDIO_2_MESES_BOPD": "Potencial promedio 2 meses bopd",
    "PRCR": "Producción de petróleo",
    "PRAG": "Producción de agua",
    "INTERVENCIONES": "Intervenciones",
    "TIPO_SWAB": "Tipo de swab",
    "POZOS": "Pozos",
    "POZOS_TOTAL": "Pozos total",
    "POZOS_INTERVENIDOS": "Pozos intervenidos",
    "POZOS_NO_INTERVENIDOS": "Pozos no intervenidos",
    "OIL_POR_INTERV": "Petróleo por intervención",
    "AGUA_POR_INTERV": "Agua por intervención",
    "MES_NOMBRE": "Mes",
    "ULTIMA_FECHA_HISTORICA": "Última fecha histórica",
    "ULTIMA_FECHA_CON_PRCR": "Última fecha con producción de petróleo",
    "ESTADO": "Estado"
}

# ============================================================
# FUNCIONES DE LIMPIEZA Y FORMATO
# ============================================================

def limpiar_pozo(valor):
    if pd.isna(valor):
        return ""
    return re.sub(r"\s+", "", str(valor).strip().upper())


def mostrar_pozo(valor):
    if pd.isna(valor):
        return ""
    return re.sub(r"\s+", " ", str(valor).strip().upper())


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return re.sub(r"\s+", " ", str(valor).strip().upper())


def normalizar_columna(columna):
    texto = str(columna).strip().upper()
    reemplazos = {"Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U"}
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    return re.sub(r"\s+", "_", texto)


def convertir_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def keys_ata():
    return {limpiar_pozo(p) for p in CANDIDATOS_ATA}


def clasificar_pozo(pozo_key):
    for anio, lista in CONVERTIDOS.items():
        if pozo_key in {limpiar_pozo(p) for p in lista}:
            return f"Convertido {anio}"
    return "Básica"


def obtener_anio_conversion(pozo_key):
    for anio, lista in CONVERTIDOS.items():
        if pozo_key in {limpiar_pozo(p) for p in lista}:
            return anio
    return 0


def normalizar_tipo_swab(valor):
    texto = limpiar_texto(valor)
    if texto in ["TBG", "TB", "TS", "TUBING", "TUBING SWAB"]:
        return "TS"
    if texto in ["CS", "CSG", "CASING", "CASING SWAB"]:
        return "CS"
    if "TBG" in texto or "TUB" in texto:
        return "TS"
    if "CS" in texto or "CAS" in texto:
        return "CS"
    if texto == "":
        return "SIN TIPO"
    return texto


def periodo_texto(anio, mes):
    return f"{MESES[int(mes)]} {int(anio)}"


def primer_dia_mes(anio, mes):
    return pd.Timestamp(int(anio), int(mes), 1)


def ultimo_dia_mes(anio, mes):
    return primer_dia_mes(anio, mes) + pd.offsets.MonthEnd(0)


def dias_mes(anio, mes):
    return calendar.monthrange(int(anio), int(mes))[1]


def mes_anterior(anio, mes, n=1):
    fecha = pd.Timestamp(int(anio), int(mes), 1) - pd.DateOffset(months=int(n))
    return int(fecha.year), int(fecha.month)


def lista_meses_previos(anio, mes, n):
    salida = []
    for i in range(int(n), 0, -1):
        salida.append(mes_anterior(anio, mes, i))
    return salida


def ultimo_mes_completo(df):
    fecha_max = df["FECHA"].max()
    fin_mes = fecha_max + pd.offsets.MonthEnd(0)
    if fecha_max.date() < fin_mes.date():
        fecha_ref = fecha_max - pd.DateOffset(months=1)
    else:
        fecha_ref = fecha_max
    return int(fecha_ref.year), int(fecha_ref.month)


def nombres_unicos(columnas):
    conteo = {}
    salida = []
    for col in columnas:
        nombre = str(col)
        if nombre not in conteo:
            conteo[nombre] = 0
            salida.append(nombre)
        else:
            conteo[nombre] += 1
            salida.append(f"{nombre} ({conteo[nombre] + 1})")
    return salida


def vista_tabla(df):
    salida = df.copy()
    salida = salida.rename(columns={c: LABELS.get(c, c) for c in salida.columns})
    salida.columns = nombres_unicos(salida.columns)

    for col in salida.columns:
        if pd.api.types.is_datetime64_any_dtype(salida[col]):
            salida[col] = salida[col].dt.strftime("%Y-%m-%d")
        elif pd.api.types.is_period_dtype(salida[col]):
            salida[col] = salida[col].astype(str)
        elif salida[col].dtype == "object":
            salida[col] = salida[col].fillna("").astype(str)
        elif pd.api.types.is_numeric_dtype(salida[col]):
            salida[col] = salida[col].fillna(0).round(2)

    return salida


def df_excel(df):
    return vista_tabla(df)


def convertir_excel(tablas):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for nombre, tabla in tablas.items():
            if tabla is None:
                continue
            hoja = str(nombre)[:31]
            out = df_excel(tabla)
            out.to_excel(writer, sheet_name=hoja, index=False)
            workbook = writer.book
            worksheet = writer.sheets[hoja]
            header = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
            for j, col in enumerate(out.columns):
                worksheet.write(0, j, col, header)
                worksheet.set_column(j, j, 22)
    buffer.seek(0)
    return buffer.getvalue()


def aplicar_layout(fig, titulo, altura=520):
    fig.update_layout(
        template="plotly_white",
        title=titulo,
        height=altura,
        margin=dict(l=20, r=20, t=70, b=50),
        legend_title_text=""
    )
    fig.update_yaxes(tickformat=",.2f")

    for tr in fig.data:
        if getattr(tr, "type", "") == "bar" and getattr(tr, "text", None) is not None:
            tr.texttemplate = "%{text:,.2f}"
            tr.textposition = "outside"
            tr.cliponaxis = False

        if getattr(tr, "hovertemplate", None):
            ht = tr.hovertemplate
            for bruto, bonito in LABELS.items():
                ht = ht.replace(bruto, bonito)
            ht = ht.replace("%{y}", "%{y:,.2f}").replace("%{x}", "%{x:,.2f}")
            tr.hovertemplate = ht

        if getattr(tr, "name", None) in LABELS:
            tr.name = LABELS[tr.name]

    if fig.layout.xaxis.title.text in LABELS:
        fig.layout.xaxis.title.text = LABELS[fig.layout.xaxis.title.text]
    if fig.layout.yaxis.title.text in LABELS:
        fig.layout.yaxis.title.text = LABELS[fig.layout.yaxis.title.text]

    return fig

# ============================================================
# CARGA DE DATOS
# ============================================================

def cargar_datos(bytes_excel):
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))
    hoja = "Datos de Swab" if "Datos de Swab" in xls.sheet_names else xls.sheet_names[0]

    df = pd.read_excel(io.BytesIO(bytes_excel), sheet_name=hoja)
    df.columns = [normalizar_columna(c) for c in df.columns]

    requeridas = ["FECHA", "COD_POZ", "COD_BAT", "UNIDAD", "TSER", "PRCR", "PRAG"]
    faltan = [c for c in requeridas if c not in df.columns]
    if faltan:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(faltan))

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
    df["PERIODO_MES"] = df["FECHA"].dt.to_period("M")
    df["CLASIFICACION"] = df["POZO_KEY"].apply(clasificar_pozo)
    df["ANIO_CONVERSION"] = df["POZO_KEY"].apply(obtener_anio_conversion)

    ata = df[df["POZO_KEY"].isin(keys_ata())].copy()
    swab = df[~df["POZO_KEY"].isin(keys_ata())].copy()

    return swab, ata, hoja


def construir_convertidos():
    registros = []
    for anio, lista in CONVERTIDOS.items():
        for pozo in lista:
            key = limpiar_pozo(pozo)
            if key in keys_ata():
                continue
            registros.append({
                "POZO_KEY": key,
                "POZO_CONVERTIDO": mostrar_pozo(pozo),
                "CLASIFICACION_CONVERTIDO": f"Convertido {anio}",
                "ANIO_CONVERSION_CONVERTIDO": anio
            })
    return pd.DataFrame(registros).drop_duplicates("POZO_KEY")


def construir_universo(df):
    hist = (
        df.sort_values("FECHA")
        .groupby("POZO_KEY", as_index=False)
        .agg(
            POZO_HIST=("POZO", "last"),
            BATERIA_HIST=("BATERIA", "last"),
            ULTIMA_FECHA_HISTORICA=("FECHA", "max")
        )
    )
    conv = construir_convertidos()
    universo = hist.merge(conv, on="POZO_KEY", how="outer")
    universo["POZO"] = universo["POZO_HIST"].fillna(universo["POZO_CONVERTIDO"])
    universo["BATERIA"] = universo["BATERIA_HIST"].fillna("SIN BATERIA")
    universo["CLASIFICACION"] = universo["CLASIFICACION_CONVERTIDO"].fillna("Básica")
    universo["ANIO_CONVERSION"] = universo["ANIO_CONVERSION_CONVERTIDO"].fillna(0).astype(int)
    return universo[["POZO_KEY", "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION", "ULTIMA_FECHA_HISTORICA"]].drop_duplicates("POZO_KEY")

# ============================================================
# CÁLCULOS
# ============================================================

def filtrar_fechas(df, inicio, fin):
    return df[(df["FECHA"] >= pd.to_datetime(inicio)) & (df["FECHA"] <= pd.to_datetime(fin))].copy()


def filtrar_mes(df, anio, mes, baterias=None, tipos=None, clases=None):
    data = df[(df["ANIO"] == int(anio)) & (df["MES"] == int(mes))].copy()
    if baterias:
        data = data[data["BATERIA"].isin(baterias)]
    if tipos:
        data = data[data["TIPO_SWAB"].isin(tipos)]
    if clases:
        data = data[data["CLASIFICACION"].isin(clases)]
    return data


def filtrar_universo(universo, baterias=None, clases=None):
    data = universo.copy()
    if baterias:
        data = data[data["BATERIA"].isin(baterias)]
    if clases:
        data = data[data["CLASIFICACION"].isin(clases)]
    return data


def resumen_pozos_mes(data_mes, universo):
    if data_mes.empty:
        out = universo.copy()
        out["INTERVENCIONES"] = 0
        out["PRCR"] = 0.0
        out["PRAG"] = 0.0
        out["OIL_POR_INTERV"] = 0.0
        out["AGUA_POR_INTERV"] = 0.0
        out["TIPO_SWAB"] = ""
        out["UNIDADES"] = ""
        out["PRIMERA_FECHA"] = pd.NaT
        out["ULTIMA_FECHA"] = pd.NaT
        out["ESTADO"] = "No intervenido"
        return out

    res = data_mes.groupby("POZO_KEY", as_index=False).agg(
        POZO_REAL=("POZO", "last"),
        BATERIA_REAL=("BATERIA", "last"),
        INTERVENCIONES=("FECHA", "count"),
        PRCR=("PRCR", "sum"),
        PRAG=("PRAG", "sum"),
        TIPO_SWAB=("TIPO_SWAB", lambda x: ", ".join(sorted(set(x)))),
        UNIDADES=("UNIDAD", lambda x: ", ".join(sorted(set(x)))),
        PRIMERA_FECHA=("FECHA", "min"),
        ULTIMA_FECHA=("FECHA", "max")
    )

    out = universo.merge(res, on="POZO_KEY", how="left")
    out["POZO"] = out["POZO"].fillna(out["POZO_REAL"])
    out["BATERIA"] = out["BATERIA"].fillna(out["BATERIA_REAL"])
    out["INTERVENCIONES"] = out["INTERVENCIONES"].fillna(0).astype(int)
    out["PRCR"] = out["PRCR"].fillna(0)
    out["PRAG"] = out["PRAG"].fillna(0)
    out["OIL_POR_INTERV"] = np.where(out["INTERVENCIONES"] > 0, out["PRCR"] / out["INTERVENCIONES"], 0)
    out["AGUA_POR_INTERV"] = np.where(out["INTERVENCIONES"] > 0, out["PRAG"] / out["INTERVENCIONES"], 0)
    out["TIPO_SWAB"] = out["TIPO_SWAB"].fillna("")
    out["UNIDADES"] = out["UNIDADES"].fillna("")
    out["ESTADO"] = np.where(out["INTERVENCIONES"] > 0, "Intervenido", "No intervenido")
    return out.drop(columns=[c for c in ["POZO_REAL", "BATERIA_REAL"] if c in out.columns])


def resumen_clase(resumen):
    out = resumen.groupby("CLASIFICACION", as_index=False).agg(
        POZOS_TOTAL=("POZO_KEY", "nunique"),
        POZOS_INTERVENIDOS=("ESTADO", lambda x: (x == "Intervenido").sum()),
        POZOS_NO_INTERVENIDOS=("ESTADO", lambda x: (x == "No intervenido").sum()),
        INTERVENCIONES=("INTERVENCIONES", "sum"),
        PRCR=("PRCR", "sum"),
        PRAG=("PRAG", "sum")
    )
    out["OIL_POR_INTERV"] = np.where(out["INTERVENCIONES"] > 0, out["PRCR"] / out["INTERVENCIONES"], 0)
    return out.sort_values("PRCR", ascending=False)


def resumen_bateria(resumen):
    out = resumen.groupby("BATERIA", as_index=False).agg(
        POZOS_TOTAL=("POZO_KEY", "nunique"),
        POZOS_INTERVENIDOS=("ESTADO", lambda x: (x == "Intervenido").sum()),
        POZOS_NO_INTERVENIDOS=("ESTADO", lambda x: (x == "No intervenido").sum()),
        INTERVENCIONES=("INTERVENCIONES", "sum"),
        PRCR=("PRCR", "sum"),
        PRAG=("PRAG", "sum")
    )
    out["OIL_POR_INTERV"] = np.where(out["INTERVENCIONES"] > 0, out["PRCR"] / out["INTERVENCIONES"], 0)
    return out.sort_values("PRCR", ascending=False)


def resumen_tipo(data_mes):
    if data_mes.empty:
        return pd.DataFrame(columns=["TIPO_SWAB", "POZOS", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV"])
    out = data_mes.groupby("TIPO_SWAB", as_index=False).agg(
        POZOS=("POZO_KEY", "nunique"),
        INTERVENCIONES=("FECHA", "count"),
        PRCR=("PRCR", "sum"),
        PRAG=("PRAG", "sum")
    )
    out["OIL_POR_INTERV"] = np.where(out["INTERVENCIONES"] > 0, out["PRCR"] / out["INTERVENCIONES"], 0)
    return out.sort_values("PRCR", ascending=False)


def tendencia(df, anio, baterias=None, tipos=None, clases=None):
    data = df[df["ANIO"] == int(anio)].copy()
    if baterias:
        data = data[data["BATERIA"].isin(baterias)]
    if tipos:
        data = data[data["TIPO_SWAB"].isin(tipos)]
    if clases:
        data = data[data["CLASIFICACION"].isin(clases)]
    if data.empty:
        return pd.DataFrame(columns=["MES", "MES_NOMBRE", "POZOS", "INTERVENCIONES", "PRCR", "PRAG"])
    return data.groupby(["MES", "MES_NOMBRE"], as_index=False).agg(
        POZOS=("POZO_KEY", "nunique"),
        INTERVENCIONES=("FECHA", "count"),
        PRCR=("PRCR", "sum"),
        PRAG=("PRAG", "sum")
    ).sort_values("MES")


def resumen_periodo_pozo(data, sufijo, divisor=1):
    if data.empty:
        return pd.DataFrame(columns=[
            "POZO_KEY", f"INTERV_{sufijo}", f"PRCR_{sufijo}", f"PRAG_{sufijo}",
            f"OIL_INTERV_{sufijo}", f"ULTIMA_FECHA_{sufijo}"
        ])
    res = data.groupby("POZO_KEY", as_index=False).agg(
        INTERV=("FECHA", "count"),
        PRCR=("PRCR", "sum"),
        PRAG=("PRAG", "sum"),
        ULTIMA_FECHA=("FECHA", "max")
    )
    divisor = max(float(divisor), 1.0)
    res[f"INTERV_{sufijo}"] = res["INTERV"] / divisor
    res[f"PRCR_{sufijo}"] = res["PRCR"] / divisor
    res[f"PRAG_{sufijo}"] = res["PRAG"] / divisor
    res[f"OIL_INTERV_{sufijo}"] = np.where(res[f"INTERV_{sufijo}"] > 0, res[f"PRCR_{sufijo}"] / res[f"INTERV_{sufijo}"], 0)
    res[f"ULTIMA_FECHA_{sufijo}"] = res["ULTIMA_FECHA"]
    return res[["POZO_KEY", f"INTERV_{sufijo}", f"PRCR_{sufijo}", f"PRAG_{sufijo}", f"OIL_INTERV_{sufijo}", f"ULTIMA_FECHA_{sufijo}"]]


def calcular_potencial(df, universo, baterias=None, tipos=None, clases=None):
    data = df.copy()
    if baterias:
        data = data[data["BATERIA"].isin(baterias)]
    if tipos:
        data = data[data["TIPO_SWAB"].isin(tipos)]
    if clases:
        data = data[data["CLASIFICACION"].isin(clases)]
    prod = data[data["PRCR"] > 0].copy()
    if prod.empty:
        out = universo.copy()
        out["ULTIMO_MES_ACTIVO"] = "Sin producción de petróleo"
        out["POTENCIAL_ULTIMO_MES_BOPD"] = 0.0
        out["POTENCIAL_MES_ANTERIOR_BOPD"] = 0.0
        out["POTENCIAL_PROMEDIO_2_MESES_BOPD"] = 0.0
        return out

    prod["ANIO_MES"] = prod["FECHA"].dt.to_period("M")
    mensual = prod.groupby(["POZO_KEY", "ANIO_MES"], as_index=False).agg(
        PRCR_ULTIMO_MES=("PRCR", "sum"),
        PRAG_ULTIMO_MES=("PRAG", "sum"),
        INTERV_ULTIMO_MES=("FECHA", "count"),
        ULTIMA_FECHA_CON_PRCR=("FECHA", "max")
    )
    idx = mensual.groupby("POZO_KEY")["ANIO_MES"].idxmax()
    ultimo = mensual.loc[idx].copy()
    ultimo["ANIO_ULT"] = ultimo["ANIO_MES"].dt.year
    ultimo["MES_ULT"] = ultimo["ANIO_MES"].dt.month
    ultimo["DIAS_MES"] = ultimo.apply(lambda r: dias_mes(r["ANIO_ULT"], r["MES_ULT"]), axis=1)
    ultimo["ULTIMO_MES_ACTIVO"] = ultimo.apply(lambda r: periodo_texto(r["ANIO_ULT"], r["MES_ULT"]), axis=1)
    ultimo["POTENCIAL_ULTIMO_MES_BOPD"] = ultimo["PRCR_ULTIMO_MES"] / ultimo["DIAS_MES"]

    previos = []
    for _, row in ultimo.iterrows():
        a_prev, m_prev = mes_anterior(row["ANIO_ULT"], row["MES_ULT"], 1)
        periodo_prev = pd.Period(f"{a_prev}-{m_prev:02d}", freq="M")
        parte = mensual[(mensual["POZO_KEY"] == row["POZO_KEY"]) & (mensual["ANIO_MES"] == periodo_prev)]
        if parte.empty:
            prcr_prev, prag_prev, interv_prev = 0.0, 0.0, 0
        else:
            prcr_prev = float(parte["PRCR_ULTIMO_MES"].iloc[0])
            prag_prev = float(parte["PRAG_ULTIMO_MES"].iloc[0])
            interv_prev = int(parte["INTERV_ULTIMO_MES"].iloc[0])
        dias_prev = dias_mes(a_prev, m_prev)
        previos.append({
            "POZO_KEY": row["POZO_KEY"],
            "MES_ANTERIOR_ULTIMO_ACTIVO": periodo_texto(a_prev, m_prev),
            "DIAS_MES_ANTERIOR": dias_prev,
            "PRCR_MES_ANTERIOR": prcr_prev,
            "PRAG_MES_ANTERIOR": prag_prev,
            "INTERV_MES_ANTERIOR": interv_prev,
            "POTENCIAL_MES_ANTERIOR_BOPD": prcr_prev / dias_prev if dias_prev > 0 else 0
        })
    prev = pd.DataFrame(previos)
    out = universo.merge(ultimo, on="POZO_KEY", how="left").merge(prev, on="POZO_KEY", how="left")
    for c in ["DIAS_MES", "PRCR_ULTIMO_MES", "PRAG_ULTIMO_MES", "INTERV_ULTIMO_MES", "POTENCIAL_ULTIMO_MES_BOPD", "DIAS_MES_ANTERIOR", "PRCR_MES_ANTERIOR", "PRAG_MES_ANTERIOR", "INTERV_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD"]:
        if c not in out.columns:
            out[c] = 0
        out[c] = out[c].fillna(0)
    out["ULTIMO_MES_ACTIVO"] = out["ULTIMO_MES_ACTIVO"].fillna("Sin producción de petróleo")
    out["MES_ANTERIOR_ULTIMO_ACTIVO"] = out["MES_ANTERIOR_ULTIMO_ACTIVO"].fillna("")
    out["POTENCIAL_PROMEDIO_2_MESES_BOPD"] = np.where(
        out["PRCR_MES_ANTERIOR"] > 0,
        (out["POTENCIAL_ULTIMO_MES_BOPD"] + out["POTENCIAL_MES_ANTERIOR_BOPD"]) / 2,
        out["POTENCIAL_ULTIMO_MES_BOPD"]
    )
    return out.sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False)


def calcular_desplazamiento(df, universo, potencial, anio, mes, meses_base, baterias, tipos, clases_desplazadas):
    actual = df[(df["FECHA"] >= primer_dia_mes(anio, mes)) & (df["FECHA"] <= ultimo_dia_mes(anio, mes))].copy()
    meses_previos = lista_meses_previos(anio, mes, meses_base)
    partes = [df[(df["FECHA"] >= primer_dia_mes(a, m)) & (df["FECHA"] <= ultimo_dia_mes(a, m))].copy() for a, m in meses_previos]
    base = pd.concat(partes, ignore_index=True) if partes else df.iloc[0:0].copy()
    a_ant, m_ant = mes_anterior(anio, mes, 1)
    mes_ant = df[(df["FECHA"] >= primer_dia_mes(a_ant, m_ant)) & (df["FECHA"] <= ultimo_dia_mes(a_ant, m_ant))].copy()

    if baterias:
        actual = actual[actual["BATERIA"].isin(baterias)]
        base = base[base["BATERIA"].isin(baterias)]
        mes_ant = mes_ant[mes_ant["BATERIA"].isin(baterias)]
    if tipos:
        actual = actual[actual["TIPO_SWAB"].isin(tipos)]
        base = base[base["TIPO_SWAB"].isin(tipos)]
        mes_ant = mes_ant[mes_ant["TIPO_SWAB"].isin(tipos)]

    universo_eval = universo.copy()
    if baterias:
        universo_eval = universo_eval[universo_eval["BATERIA"].isin(baterias)]
    if clases_desplazadas:
        universo_eval = universo_eval[universo_eval["CLASIFICACION"].isin(clases_desplazadas)]

    tabla = universo_eval.merge(resumen_periodo_pozo(base, "BASE", divisor=meses_base), on="POZO_KEY", how="left")
    tabla = tabla.merge(resumen_periodo_pozo(actual, "ACTUAL", divisor=1), on="POZO_KEY", how="left")
    tabla = tabla.merge(resumen_periodo_pozo(mes_ant, "MES_ANTERIOR_OBJETIVO", divisor=1), on="POZO_KEY", how="left")

    num_cols = [
        "INTERV_BASE", "PRCR_BASE", "PRAG_BASE", "OIL_INTERV_BASE",
        "INTERV_ACTUAL", "PRCR_ACTUAL", "PRAG_ACTUAL", "OIL_INTERV_ACTUAL",
        "INTERV_MES_ANTERIOR_OBJETIVO", "PRCR_MES_ANTERIOR_OBJETIVO", "PRAG_MES_ANTERIOR_OBJETIVO", "OIL_INTERV_MES_ANTERIOR_OBJETIVO"
    ]
    for c in num_cols:
        if c not in tabla.columns:
            tabla[c] = 0
        tabla[c] = tabla[c].fillna(0)

    tabla["VAR_INTERV"] = tabla["INTERV_ACTUAL"] - tabla["INTERV_BASE"]
    tabla["VAR_PRCR"] = tabla["PRCR_ACTUAL"] - tabla["PRCR_BASE"]
    tabla["VAR_PRAG"] = tabla["PRAG_ACTUAL"] - tabla["PRAG_BASE"]
    tabla["INTERV_BASE_NO_REALIZADAS"] = np.where(tabla["VAR_INTERV"] < 0, -tabla["VAR_INTERV"], 0)
    tabla["PRCR_BASE_NO_REALIZADO"] = np.where(tabla["VAR_PRCR"] < 0, -tabla["VAR_PRCR"], 0)
    tabla["PRAG_BASE_NO_REALIZADO"] = np.where(tabla["VAR_PRAG"] < 0, -tabla["VAR_PRAG"], 0)
    tabla["VAR_PETROLEO_VS_MES_ANTERIOR"] = tabla["PRCR_ACTUAL"] - tabla["PRCR_MES_ANTERIOR_OBJETIVO"]
    tabla["PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR"] = np.where(tabla["VAR_PETROLEO_VS_MES_ANTERIOR"] < 0, -tabla["VAR_PETROLEO_VS_MES_ANTERIOR"], 0)
    tabla["INTERVENCIONES_DEJADAS_MES_ANTERIOR"] = np.where(tabla["INTERV_ACTUAL"] < tabla["INTERV_MES_ANTERIOR_OBJETIVO"], tabla["INTERV_MES_ANTERIOR_OBJETIVO"] - tabla["INTERV_ACTUAL"], 0)

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
    candidatos = tabla[tabla["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])]
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

    pot_cols = [
        "POZO_KEY", "ULTIMO_MES_ACTIVO", "DIAS_MES", "PRCR_ULTIMO_MES", "PRAG_ULTIMO_MES", "INTERV_ULTIMO_MES",
        "POTENCIAL_ULTIMO_MES_BOPD", "MES_ANTERIOR_ULTIMO_ACTIVO", "DIAS_MES_ANTERIOR", "PRCR_MES_ANTERIOR",
        "PRAG_MES_ANTERIOR", "INTERV_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD", "POTENCIAL_PROMEDIO_2_MESES_BOPD"
    ]
    tabla = tabla.merge(potencial[pot_cols], on="POZO_KEY", how="left")
    orden = {"DEJADO DE HACER": 1, "REDUCIDO": 2, "NUEVO / RETOMADO": 3, "MANTENIDO / AUMENTADO": 4, "SIN ACTIVIDAD": 5}
    tabla["ORDEN"] = tabla["ESTADO_DESPLAZAMIENTO"].map(orden).fillna(9)
    tabla = tabla.sort_values(["ORDEN", "PRCR_BASE_NO_REALIZADO"], ascending=[True, False]).drop(columns="ORDEN")
    return tabla, actual, base, meses_previos


def listado_convertidos(universo, resumen_mes):
    conv = universo[universo["CLASIFICACION"].str.startswith("Convertido")].copy()
    cols = ["POZO_KEY", "ESTADO", "INTERVENCIONES", "PRCR", "PRAG", "TIPO_SWAB", "UNIDADES", "ULTIMA_FECHA"]
    disp = resumen_mes[[c for c in cols if c in resumen_mes.columns]].copy()
    out = conv.merge(disp, on="POZO_KEY", how="left")
    out["ESTADO"] = out["ESTADO"].fillna("No intervenido")
    for c in ["INTERVENCIONES", "PRCR", "PRAG"]:
        out[c] = out[c].fillna(0)
    out["TIPO_SWAB"] = out["TIPO_SWAB"].fillna("")
    out["UNIDADES"] = out["UNIDADES"].fillna("")
    return out.sort_values(["ANIO_CONVERSION", "POZO"])

# ============================================================
# INTERFAZ
# ============================================================

st.title("🛢️ SWAB Lote X - Servicio y Potencial")
st.caption("Análisis de pozos dejados, impacto de convertidos 2026, potencial mensual y exclusión de candidatos ATA.")

archivo = st.file_uploader("Sube el Excel principal con la hoja Datos de Swab", type=["xlsx"])

if archivo is None:
    st.info("Sube el Excel para iniciar el análisis.")
    st.stop()

try:
    df_total, df_ata, hoja = cargar_datos(archivo.getvalue())
except Exception as e:
    st.error(f"No se pudo cargar el Excel: {e}")
    st.stop()

fecha_min = df_total["FECHA"].min().date()
fecha_max = df_total["FECHA"].max().date()

with st.sidebar:
    st.header("1. Rango de análisis")
    fecha_inicio = st.date_input("Desde", value=fecha_min, min_value=fecha_min, max_value=fecha_max)
    fecha_fin = st.date_input("Hasta", value=fecha_max, min_value=fecha_min, max_value=fecha_max)

    st.header("2. Módulo principal")
    modulo = st.radio(
        "Selecciona con qué trabajar",
        ["Análisis de pozos dejados", "Potencial mensual de pozos", "Vista completa"],
        index=2
    )

df = filtrar_fechas(df_total, fecha_inicio, fecha_fin)
if df.empty:
    st.error("El rango seleccionado no tiene datos.")
    st.stop()

universo = construir_universo(df)
anios = sorted(df["ANIO"].unique().astype(int).tolist())
anio_def, mes_def = ultimo_mes_completo(df)
baterias = sorted([x for x in universo["BATERIA"].dropna().unique().tolist() if x != ""])
tipos = sorted([x for x in df["TIPO_SWAB"].dropna().unique().tolist() if x != ""])
clases = ["Básica", "Convertido 2024", "Convertido 2025", "Convertido 2026"]

with st.sidebar:
    st.header("3. Periodo objetivo")
    anio_obj = st.selectbox("Año objetivo", anios, index=anios.index(anio_def) if anio_def in anios else len(anios) - 1)
    meses_disp = sorted(df[df["ANIO"] == anio_obj]["MES"].unique().astype(int).tolist())
    mes_obj = st.selectbox(
        "Mes objetivo",
        meses_disp,
        index=meses_disp.index(mes_def) if mes_def in meses_disp else len(meses_disp) - 1,
        format_func=lambda x: MESES[int(x)]
    )
    meses_base = st.slider("Comparar contra N meses anteriores", min_value=1, max_value=15, value=3, step=1)

    st.header("4. Filtros operativos")
    baterias_sel = st.multiselect("Batería", baterias, default=[])
    tipos_sel = st.multiselect("Tipo de swab", tipos, default=[])
    clases_sel = st.multiselect("Tipo de pozo para vistas generales", clases, default=[])
    clases_desplazadas = st.multiselect("Pozos a evaluar como desplazados", clases, default=["Básica", "Convertido 2024", "Convertido 2025"])
    top_n = st.slider("Top para gráficos", min_value=5, max_value=50, value=20, step=5)
    ejecutar = st.button("Ejecutar análisis", type="primary")

if not ejecutar and "resultado_swab" not in st.session_state:
    st.warning("Configura los filtros y presiona Ejecutar análisis.")
    st.stop()

if ejecutar:
    universo_general = filtrar_universo(universo, baterias_sel, clases_sel)
    data_mes = filtrar_mes(df, anio_obj, mes_obj, baterias_sel, tipos_sel, clases_sel)
    res_pozos = resumen_pozos_mes(data_mes, universo_general)
    res_clase = resumen_clase(res_pozos)
    res_bat = resumen_bateria(res_pozos)
    res_tipo = resumen_tipo(data_mes)
    res_tend = tendencia(df, anio_obj, baterias_sel, tipos_sel, clases_sel)
    pot = calcular_potencial(df, universo, baterias_sel, tipos_sel, clases_sel if clases_sel else None)
    desp, actual, base, meses_previos = calcular_desplazamiento(df, universo, pot, anio_obj, mes_obj, meses_base, baterias_sel, tipos_sel, clases_desplazadas)
    list_conv = listado_convertidos(universo, res_pozos)
    st.session_state["resultado_swab"] = {
        "anio_obj": anio_obj,
        "mes_obj": mes_obj,
        "meses_base": meses_base,
        "res_pozos": res_pozos,
        "res_clase": res_clase,
        "res_bat": res_bat,
        "res_tipo": res_tipo,
        "res_tend": res_tend,
        "pot": pot,
        "desp": desp,
        "actual": actual,
        "base": base,
        "meses_previos": meses_previos,
        "list_conv": list_conv
    }

r = st.session_state["resultado_swab"]
anio_obj = r["anio_obj"]
mes_obj = r["mes_obj"]
meses_base = r["meses_base"]
res_pozos = r["res_pozos"]
res_clase = r["res_clase"]
res_bat = r["res_bat"]
res_tipo = r["res_tipo"]
res_tend = r["res_tend"]
pot = r["pot"]
desp = r["desp"]
actual = r["actual"]
base = r["base"]
meses_previos = r["meses_previos"]
list_conv = r["list_conv"]
periodo = periodo_texto(anio_obj, mes_obj)

fecha_max_mes = ultimo_dia_mes(anio_obj, mes_obj).date()
if fecha_max < fecha_max_mes and anio_obj == fecha_max.year and mes_obj == fecha_max.month:
    st.warning(f"El mes {periodo} está incompleto en la data. La información llega hasta {fecha_max}.")

pozos_universo = res_pozos["POZO_KEY"].nunique()
pozos_interv = int((res_pozos["ESTADO"] == "Intervenido").sum())
pozos_no = int((res_pozos["ESTADO"] == "No intervenido").sum())
interv = int(res_pozos["INTERVENCIONES"].sum())
petroleo = float(res_pozos["PRCR"].sum())
agua = float(res_pozos["PRAG"].sum())

st.subheader(f"Resumen ejecutivo: {periodo}")

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Pozos universo", f"{pozos_universo:,}")
c2.metric("Intervenidos", f"{pozos_interv:,}")
c3.metric("No intervenidos", f"{pozos_no:,}")
c4.metric("Intervenciones", f"{interv:,}")
c5.metric("Producción de petróleo", f"{petroleo:,.2f}")
c6.metric("Producción de agua", f"{agua:,.2f}")

st.caption(
    f"Hoja usada: {hoja}. Rango analizado: {fecha_inicio} al {fecha_fin}. "
    f"Se excluyen candidatos ATA: {df_ata['POZO_KEY'].nunique()} pozos encontrados en la data."
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
    dejados = desp[desp["ESTADO_DESPLAZAMIENTO"] == "DEJADO DE HACER"]
    reducidos = desp[desp["ESTADO_DESPLAZAMIENTO"] == "REDUCIDO"]
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Pozos dejados", f"{len(dejados):,}")
    k2.metric("Pozos reducidos", f"{len(reducidos):,}")
    k3.metric("Interv. base no realizadas", f"{desp['INTERV_BASE_NO_REALIZADAS'].sum():,.2f}")
    k4.metric("Petróleo base no realizado", f"{desp['PRCR_BASE_NO_REALIZADO'].sum():,.2f}")
    k5.metric("Petróleo dejado vs mes anterior", f"{desp['PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR'].sum():,.2f}")
    k6.metric("Potencial promedio 2 meses", f"{desp['POTENCIAL_PROMEDIO_2_MESES_BOPD'].sum():,.2f} bopd")
    base_txt = ", ".join([periodo_texto(a, m) for a, m in meses_previos])
    st.info(f"Periodo objetivo: {periodo}. Base comparativa: promedio mensual de {meses_base} mes(es) anteriores: {base_txt}.")
    estado = st.radio("Estado de desplazamiento", ["DEJADO DE HACER", "REDUCIDO", "NUEVO / RETOMADO", "MANTENIDO / AUMENTADO", "SIN ACTIVIDAD", "TODOS"], horizontal=True)
    tabla = desp.copy()
    if estado != "TODOS":
        tabla = tabla[tabla["ESTADO_DESPLAZAMIENTO"] == estado]
    cols_desp = [
        "ESTADO_DESPLAZAMIENTO", "PRIORIDAD_REVISION", "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION",
        "INTERV_BASE", "INTERV_ACTUAL", "INTERV_BASE_NO_REALIZADAS", "PRCR_BASE", "PRCR_ACTUAL", "PRCR_BASE_NO_REALIZADO",
        "PRCR_MES_ANTERIOR_OBJETIVO", "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR", "VAR_PETROLEO_VS_MES_ANTERIOR",
        "INTERV_MES_ANTERIOR_OBJETIVO", "INTERVENCIONES_DEJADAS_MES_ANTERIOR", "ULTIMO_MES_ACTIVO", "PRCR_ULTIMO_MES",
        "POTENCIAL_ULTIMO_MES_BOPD", "MES_ANTERIOR_ULTIMO_ACTIVO", "PRCR_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD",
        "POTENCIAL_PROMEDIO_2_MESES_BOPD", "PRAG_BASE", "PRAG_ACTUAL", "PRAG_BASE_NO_REALIZADO", "OIL_INTERV_BASE",
        "OIL_INTERV_ACTUAL", "ULTIMA_FECHA_HISTORICA"
    ]
    cols_desp = [c for c in cols_desp if c in tabla.columns]
    st.dataframe(vista_tabla(tabla[cols_desp]), use_container_width=True, hide_index=True)

    top_dej = desp[desp["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])].sort_values("PRCR_BASE_NO_REALIZADO", ascending=False).head(top_n)
    if not top_dej.empty:
        fig1 = px.bar(top_dej, x="POZO", y="PRCR_BASE_NO_REALIZADO", color="ESTADO_DESPLAZAMIENTO", text="PRCR_BASE_NO_REALIZADO", hover_data=["BATERIA", "CLASIFICACION", "INTERV_BASE", "INTERV_ACTUAL", "POTENCIAL_PROMEDIO_2_MESES_BOPD"])
        st.plotly_chart(aplicar_layout(fig1, f"Top {top_n} pozos con petróleo base no realizado", 560), use_container_width=True)
        fig2 = px.scatter(top_dej, x="POTENCIAL_PROMEDIO_2_MESES_BOPD", y="PRCR_BASE_NO_REALIZADO", size="INTERV_BASE_NO_REALIZADAS", color="CLASIFICACION", hover_name="POZO", hover_data=["BATERIA", "ESTADO_DESPLAZAMIENTO", "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR"])
        st.plotly_chart(aplicar_layout(fig2, "Potencial promedio 2 meses vs petróleo base no realizado", 560), use_container_width=True)
    st.download_button("Descargar data de gráficas de pozos dejados", data=convertir_excel({"Top dejados reducidos": top_dej, "Tabla desplazamiento": desp[cols_desp]}), file_name=f"data_graficas_pozos_dejados_{anio_obj}_{mes_obj}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tabs[1]:
    st.subheader("Impacto de priorizar convertidos 2026")
    conv_actual = actual[actual["CLASIFICACION"] == "Convertido 2026"].copy()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Conv. 2026 intervenidos", f"{conv_actual['POZO_KEY'].nunique():,} de 20")
    c2.metric("Interv. conv. 2026", f"{len(conv_actual):,}")
    c3.metric("Producción de petróleo conv. 2026", f"{conv_actual['PRCR'].sum():,.2f}")
    c4.metric("Producción de agua conv. 2026", f"{conv_actual['PRAG'].sum():,.2f}")
    fig = px.bar(res_clase, x="CLASIFICACION", y=["PRCR", "PRAG"], barmode="group", text_auto=".2f")
    st.plotly_chart(aplicar_layout(fig, "Producción de petróleo y agua por clasificación", 520), use_container_width=True)
    afectados = desp[desp["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])].copy()
    if not afectados.empty:
        res_afect = afectados.groupby("CLASIFICACION", as_index=False).agg(
            POZOS_AFECTADOS=("POZO_KEY", "nunique"),
            INTERV_BASE_NO_REALIZADAS=("INTERV_BASE_NO_REALIZADAS", "sum"),
            PRCR_BASE_NO_REALIZADO=("PRCR_BASE_NO_REALIZADO", "sum"),
            PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR=("PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR", "sum"),
            POTENCIAL_PROMEDIO_2_MESES_BOPD=("POTENCIAL_PROMEDIO_2_MESES_BOPD", "sum")
        ).sort_values("PRCR_BASE_NO_REALIZADO", ascending=False)
        st.write("Pozos posiblemente desplazados por clasificación")
        st.dataframe(vista_tabla(res_afect), use_container_width=True, hide_index=True)

with tabs[2]:
    st.subheader("Listado de pozos convertidos 2024, 2025 y 2026")
    filtro_anio = st.radio("Año de conversión", ["Todos", "2024", "2025", "2026"], horizontal=True)
    tabla_conv = list_conv.copy()
    if filtro_anio != "Todos":
        tabla_conv = tabla_conv[tabla_conv["ANIO_CONVERSION"] == int(filtro_anio)]
    cols_conv = ["ANIO_CONVERSION", "CLASIFICACION", "POZO", "BATERIA", "ESTADO", "INTERVENCIONES", "PRCR", "PRAG", "TIPO_SWAB", "ULTIMA_FECHA", "ULTIMA_FECHA_HISTORICA"]
    cols_conv = [c for c in cols_conv if c in tabla_conv.columns]
    st.dataframe(vista_tabla(tabla_conv[cols_conv]), use_container_width=True, hide_index=True)

with tabs[3]:
    st.subheader("Potencial mensual por pozo")
    st.caption("Criterio: último mes activo con producción de petróleo mayor a cero. Potencial = producción de petróleo mensual dividida entre días calendario del mes. También se muestra el mes anterior para validar representatividad.")
    cols_pot = ["POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION", "ULTIMO_MES_ACTIVO", "DIAS_MES", "PRCR_ULTIMO_MES", "POTENCIAL_ULTIMO_MES_BOPD", "MES_ANTERIOR_ULTIMO_ACTIVO", "PRCR_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD", "POTENCIAL_PROMEDIO_2_MESES_BOPD", "PRAG_ULTIMO_MES", "PRAG_MES_ANTERIOR", "INTERV_ULTIMO_MES", "INTERV_MES_ANTERIOR", "ULTIMA_FECHA_CON_PRCR"]
    cols_pot = [c for c in cols_pot if c in pot.columns]
    st.dataframe(vista_tabla(pot[cols_pot]), use_container_width=True, hide_index=True)
    top_pot = pot.sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False).head(top_n)
    if not top_pot.empty:
        fig = px.bar(top_pot, x="POZO", y=["POTENCIAL_ULTIMO_MES_BOPD", "POTENCIAL_MES_ANTERIOR_BOPD"], barmode="group", hover_data=["BATERIA", "CLASIFICACION", "ULTIMO_MES_ACTIVO", "MES_ANTERIOR_ULTIMO_ACTIVO"])
        st.plotly_chart(aplicar_layout(fig, f"Top {top_n} pozos por potencial", 560), use_container_width=True)

with tabs[4]:
    st.subheader("Pozos y baterías")
    estado_pozo = st.radio("Estado", ["Todos", "Intervenido", "No intervenido"], horizontal=True)
    tabla_p = res_pozos.copy()
    if estado_pozo != "Todos":
        tabla_p = tabla_p[tabla_p["ESTADO"] == estado_pozo]
    cols = ["ESTADO", "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION", "TIPO_SWAB", "UNIDADES", "INTERVENCIONES", "PRCR", "PRAG", "OIL_POR_INTERV", "AGUA_POR_INTERV", "ULTIMA_FECHA", "ULTIMA_FECHA_HISTORICA"]
    cols = [c for c in cols if c in tabla_p.columns]
    st.dataframe(vista_tabla(tabla_p[cols]), use_container_width=True, hide_index=True)
    if not res_bat.empty:
        fig = px.bar(res_bat.head(top_n), x="BATERIA", y="PRCR", text="PRCR", hover_data=["POZOS_INTERVENIDOS", "POZOS_NO_INTERVENIDOS", "INTERVENCIONES", "PRAG"])
        st.plotly_chart(aplicar_layout(fig, f"Top {top_n} baterías por producción de petróleo", 520), use_container_width=True)
    if not res_tend.empty:
        orden = [MESES[m] for m in sorted(res_tend["MES"].unique())]
        fig2 = px.line(res_tend, x="MES_NOMBRE", y=["PRCR", "PRAG"], markers=True, category_orders={"MES_NOMBRE": orden})
        st.plotly_chart(aplicar_layout(fig2, f"Tendencia mensual {anio_obj}", 520), use_container_width=True)

with tabs[5]:
    st.subheader("TS y CS")
    if res_tipo.empty:
        st.info("No hay información por tipo de swab.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(res_tipo, x="TIPO_SWAB", y=["PRCR", "PRAG"], barmode="group", text_auto=".2f")
            st.plotly_chart(aplicar_layout(fig, "Producción de petróleo y agua por tipo de swab", 520), use_container_width=True)
        with col2:
            figp = px.pie(res_tipo, names="TIPO_SWAB", values="INTERVENCIONES", hole=0.45)
            figp.update_traces(texttemplate="%{percent:.2%}", hovertemplate="%{label}<br>Intervenciones: %{value:,.2f}<br>%{percent:.2%}<extra></extra>")
            st.plotly_chart(aplicar_layout(figp, "Distribución de intervenciones por tipo de swab", 520), use_container_width=True)
        st.dataframe(vista_tabla(res_tipo), use_container_width=True, hide_index=True)

with tabs[6]:
    st.subheader("Descargas")
    cols_desp_export = [c for c in [
        "ESTADO_DESPLAZAMIENTO", "PRIORIDAD_REVISION", "POZO", "BATERIA", "CLASIFICACION", "ANIO_CONVERSION", "INTERV_BASE", "INTERV_ACTUAL", "INTERV_BASE_NO_REALIZADAS", "PRCR_BASE", "PRCR_ACTUAL", "PRCR_BASE_NO_REALIZADO", "PRCR_MES_ANTERIOR_OBJETIVO", "PRODUCCION_PETROLEO_DEJADA_MES_ANTERIOR", "VAR_PETROLEO_VS_MES_ANTERIOR", "INTERV_MES_ANTERIOR_OBJETIVO", "INTERVENCIONES_DEJADAS_MES_ANTERIOR", "ULTIMO_MES_ACTIVO", "PRCR_ULTIMO_MES", "POTENCIAL_ULTIMO_MES_BOPD", "MES_ANTERIOR_ULTIMO_ACTIVO", "PRCR_MES_ANTERIOR", "POTENCIAL_MES_ANTERIOR_BOPD", "POTENCIAL_PROMEDIO_2_MESES_BOPD", "PRAG_BASE", "PRAG_ACTUAL", "PRAG_BASE_NO_REALIZADO", "OIL_INTERV_BASE", "OIL_INTERV_ACTUAL", "ULTIMA_FECHA_HISTORICA"
    ] if c in desp.columns]
    tablas = {
        "Resumen pozos": res_pozos,
        "Analisis pozos dejados": desp[cols_desp_export],
        "Listado convertidos": list_conv,
        "Potencial pozos": pot,
        "Resumen clasificacion": res_clase,
        "Resumen baterias": res_bat,
        "Resumen TS CS": res_tipo,
        "Tendencia mensual": res_tend,
        "Pozos ATA excluidos": df_ata,
        "Data graf dejados": desp[desp["ESTADO_DESPLAZAMIENTO"].isin(["DEJADO DE HACER", "REDUCIDO"])].sort_values("PRCR_BASE_NO_REALIZADO", ascending=False).head(top_n),
        "Data graf potencial": pot.sort_values("POTENCIAL_PROMEDIO_2_MESES_BOPD", ascending=False).head(top_n)
    }
    st.download_button("Descargar toda la información en Excel", data=convertir_excel(tablas), file_name=f"swab_lote_x_{anio_obj}_{mes_obj}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    validacion = pd.DataFrame({
        "Concepto": ["Hoja usada", "Fecha mínima", "Fecha máxima", "Rango desde", "Rango hasta", "Registros analizados", "Pozos SWAB analizados", "Pozos ATA excluidos", "Registros ATA excluidos", "Convertidos 2026 fijos", "PRCR", "PRAG"],
        "Valor": [hoja, str(fecha_min), str(fecha_max), str(fecha_inicio), str(fecha_fin), f"{len(df):,}", f"{df['POZO_KEY'].nunique():,}", f"{df_ata['POZO_KEY'].nunique():,} de 88", f"{len(df_ata):,}", "20", "Producción de petróleo", "Producción de agua"]
    })
    st.write("Validación")
    st.dataframe(vista_tabla(validacion), use_container_width=True, hide_index=True)
