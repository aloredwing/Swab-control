import io
import re
from datetime import date

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Control de Producción por Swab",
    page_icon="🛢️",
    layout="wide"
)


# ============================================================
# FUNCIONES BASE
# ============================================================

def limpiar_pozo(valor):
    """
    Normaliza el nombre del pozo para poder cruzar hojas.
    Ejemplo:
    'AA   47' -> 'AA47'
    'EA 8586' -> 'EA8586'
    """
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = re.sub(r"\s+", "", texto)
    return texto


def encontrar_hoja(lista_hojas, nombre_buscado):
    """
    Busca una hoja aunque tenga diferencias menores de mayúsculas o espacios.
    """
    objetivo = nombre_buscado.strip().lower()
    for hoja in lista_hojas:
        if hoja.strip().lower() == objetivo:
            return hoja
    return None


def convertir_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def cargar_excel(bytes_excel):
    """
    Carga el Excel en memoria.
    No se usa st.cache_data aquí porque pd.ExcelFile no es serializable.
    """
    return pd.ExcelFile(io.BytesIO(bytes_excel))


@st.cache_data(show_spinner=False)
def cargar_datos_swab(bytes_excel):
    """
    Lee la hoja Datos de Swab.
    Esta es la base real de intervenciones.
    """
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))
    hoja = encontrar_hoja(xls.sheet_names, "Datos de Swab")

    if hoja is None:
        raise ValueError("No encontré la hoja 'Datos de Swab'.")

    df = pd.read_excel(xls, sheet_name=hoja)
    df.columns = [str(c).strip().upper() for c in df.columns]

    columnas_necesarias = ["FECHA", "COD_POZ", "COD_BAT", "UNIDAD", "PRCR"]
    faltantes = [c for c in columnas_necesarias if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en 'Datos de Swab': {faltantes}")

    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df = df[df["FECHA"].notna()].copy()

    df["POZO"] = df["COD_POZ"].astype(str).str.strip()
    df["POZO_KEY"] = df["POZO"].apply(limpiar_pozo)
    df["BATERIA"] = df["COD_BAT"].astype(str).str.strip()
    df["UNIDAD"] = df["UNIDAD"].astype(str).str.strip()
    df["PRCR"] = convertir_numero(df["PRCR"])

    if "PRAG" in df.columns:
        df["PRAG"] = convertir_numero(df["PRAG"])
    else:
        df["PRAG"] = 0

    if "TSER" in df.columns:
        df["TSER"] = df["TSER"].astype(str).str.strip()
    else:
        df["TSER"] = ""

    return df


@st.cache_data(show_spinner=False)
def cargar_swab_basica(bytes_excel):
    """
    Lee la hoja Swab Básica, que sirve como padrón de pozos.
    """
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))
    hoja = encontrar_hoja(xls.sheet_names, "Swab Básica")

    if hoja is None:
        return pd.DataFrame()

    base = pd.read_excel(xls, sheet_name=hoja)
    base.columns = [str(c).strip() for c in base.columns]

    if "Pozo" not in base.columns:
        return pd.DataFrame()

    base["POZO"] = base["Pozo"].astype(str).str.strip()
    base["POZO_KEY"] = base["POZO"].apply(limpiar_pozo)

    if "Battery" in base.columns:
        base["BATERIA_BASE"] = base["Battery"]
    else:
        base["BATERIA_BASE"] = ""

    if "Potential (bopd)" in base.columns:
        base["POTENCIAL_BOPD"] = convertir_numero(base["Potential (bopd)"])
    else:
        base["POTENCIAL_BOPD"] = 0

    columnas = ["POZO_KEY", "POZO", "BATERIA_BASE", "POTENCIAL_BOPD"]
    if "Resumen" in base.columns:
        base["RESUMEN_BASE"] = base["Resumen"]
        columnas.append("RESUMEN_BASE")

    base = base[columnas].drop_duplicates("POZO_KEY")
    base = base[base["POZO_KEY"] != ""].copy()

    return base


@st.cache_data(show_spinner=False)
def cargar_convertidos(bytes_excel, anio):
    """
    Lee la hoja Swab 2024 (Dia), Swab 2025 (Dia) o Swab 2026 (Dia).
    En esas hojas los pozos están como columnas desde la cuarta columna.
    """
    xls = pd.ExcelFile(io.BytesIO(bytes_excel))
    nombre_hoja = f"Swab {anio} (Dia)"
    hoja = encontrar_hoja(xls.sheet_names, nombre_hoja)

    if hoja is None:
        return pd.DataFrame()

    df = pd.read_excel(xls, sheet_name=hoja, header=1)

    if df.shape[1] <= 3:
        return pd.DataFrame()

    pozos = []
    columnas_pozo = list(df.columns[3:])

    for col in columnas_pozo:
        if pd.isna(col):
            continue

        pozo = str(col).strip()
        if pozo == "" or pozo.lower().startswith("unnamed"):
            continue

        valores = convertir_numero(df[col])
        total_hoja = float(valores.sum())
        dias_con_valor = int((valores > 0).sum())

        pozos.append({
            "AÑO_CONVERSION": anio,
            "POZO": pozo,
            "POZO_KEY": limpiar_pozo(pozo),
            "PRODUCCION_HOJA_CONVERTIDO": total_hoja,
            "DIAS_CON_PRODUCCION_HOJA": dias_con_valor
        })

    convertidos = pd.DataFrame(pozos)
    convertidos = convertidos[convertidos["POZO_KEY"] != ""].drop_duplicates("POZO_KEY")

    return convertidos


def performance_por_pozo_unidad(df, fecha_inicio, fecha_fin):
    """
    Calcula performance por pozo y por unidad.
    """
    data = df[(df["FECHA"] >= pd.to_datetime(fecha_inicio)) & (df["FECHA"] <= pd.to_datetime(fecha_fin))].copy()

    if data.empty:
        return pd.DataFrame()

    perf = (
        data
        .groupby(["POZO_KEY", "POZO", "BATERIA", "UNIDAD"], dropna=False)
        .agg(
            INTERVENCIONES=("FECHA", "count"),
            PRODUCCION_PRCR=("PRCR", "sum"),
            AGUA_PRAG=("PRAG", "sum"),
            PRIMERA_INTERVENCION=("FECHA", "min"),
            ULTIMA_INTERVENCION=("FECHA", "max")
        )
        .reset_index()
    )

    perf["OIL_POR_INTERV"] = perf["PRODUCCION_PRCR"] / perf["INTERVENCIONES"].replace(0, pd.NA)
    perf["OIL_POR_INTERV"] = perf["OIL_POR_INTERV"].fillna(0)

    return perf


def performance_por_pozo_total(df, fecha_inicio, fecha_fin):
    """
    Calcula performance total por pozo, sin separar por unidad.
    """
    data = df[(df["FECHA"] >= pd.to_datetime(fecha_inicio)) & (df["FECHA"] <= pd.to_datetime(fecha_fin))].copy()

    if data.empty:
        return pd.DataFrame()

    perf = (
        data
        .groupby(["POZO_KEY", "POZO"], dropna=False)
        .agg(
            BATERIA=("BATERIA", "last"),
            UNIDADES=("UNIDAD", lambda x: ", ".join(sorted(set([v for v in x.astype(str) if v and v != "nan"])))),
            INTERVENCIONES=("FECHA", "count"),
            PRODUCCION_PRCR=("PRCR", "sum"),
            AGUA_PRAG=("PRAG", "sum"),
            PRIMERA_INTERVENCION=("FECHA", "min"),
            ULTIMA_INTERVENCION=("FECHA", "max")
        )
        .reset_index()
    )

    perf["OIL_POR_INTERV"] = perf["PRODUCCION_PRCR"] / perf["INTERVENCIONES"].replace(0, pd.NA)
    perf["OIL_POR_INTERV"] = perf["OIL_POR_INTERV"].fillna(0)

    return perf


def construir_no_explotados(df, base_pozos, fecha_corte):
    """
    Lista pozos con última intervención y días sin explotar/intervenir.
    """
    ultima = (
        df
        .groupby("POZO_KEY", dropna=False)
        .agg(
            POZO_REAL=("POZO", "last"),
            BATERIA_REAL=("BATERIA", "last"),
            ULTIMA_INTERVENCION=("FECHA", "max"),
            INTERVENCIONES_TOTAL=("FECHA", "count"),
            PRODUCCION_TOTAL_PRCR=("PRCR", "sum")
        )
        .reset_index()
    )

    if base_pozos.empty:
        universo = ultima[["POZO_KEY", "POZO_REAL", "BATERIA_REAL"]].copy()
        universo = universo.rename(columns={"POZO_REAL": "POZO", "BATERIA_REAL": "BATERIA_BASE"})
        universo["POTENCIAL_BOPD"] = 0
    else:
        universo = base_pozos.copy()

    salida = universo.merge(ultima, on="POZO_KEY", how="left")

    salida["POZO_FINAL"] = salida["POZO"]
    salida.loc[salida["POZO_FINAL"].isna(), "POZO_FINAL"] = salida["POZO_REAL"]

    salida["BATERIA_FINAL"] = salida["BATERIA_BASE"]
    salida.loc[salida["BATERIA_FINAL"].isna() | (salida["BATERIA_FINAL"].astype(str).str.strip() == ""), "BATERIA_FINAL"] = salida["BATERIA_REAL"]

    salida["INTERVENCIONES_TOTAL"] = salida["INTERVENCIONES_TOTAL"].fillna(0).astype(int)
    salida["PRODUCCION_TOTAL_PRCR"] = salida["PRODUCCION_TOTAL_PRCR"].fillna(0)

    corte = pd.to_datetime(fecha_corte)

    salida["DIAS_SIN_INTERVENCION"] = (corte - salida["ULTIMA_INTERVENCION"]).dt.days
    salida.loc[salida["ULTIMA_INTERVENCION"].isna(), "DIAS_SIN_INTERVENCION"] = 99999

    def estado(row):
        if pd.isna(row["ULTIMA_INTERVENCION"]):
            return "Nunca intervenido en la base"
        return "Con última intervención"

    salida["ESTADO"] = salida.apply(estado, axis=1)

    def rango_dias(dias):
        if dias >= 99999:
            return "Nunca"
        if dias <= 7:
            return "0 a 7 días"
        if dias <= 30:
            return "8 a 30 días"
        if dias <= 60:
            return "31 a 60 días"
        if dias <= 90:
            return "61 a 90 días"
        if dias <= 120:
            return "91 a 120 días"
        if dias <= 180:
            return "121 a 180 días"
        return "Más de 180 días"

    salida["RANGO_SIN_INTERVENCION"] = salida["DIAS_SIN_INTERVENCION"].apply(rango_dias)

    columnas_finales = [
        "POZO_KEY",
        "POZO_FINAL",
        "BATERIA_FINAL",
        "POTENCIAL_BOPD",
        "ULTIMA_INTERVENCION",
        "DIAS_SIN_INTERVENCION",
        "RANGO_SIN_INTERVENCION",
        "INTERVENCIONES_TOTAL",
        "PRODUCCION_TOTAL_PRCR",
        "ESTADO"
    ]

    if "RESUMEN_BASE" in salida.columns:
        columnas_finales.append("RESUMEN_BASE")

    salida = salida[columnas_finales].rename(columns={
        "POZO_FINAL": "POZO",
        "BATERIA_FINAL": "BATERIA"
    })

    return salida


def formato_tabla(df):
    """
    Ordena y redondea campos numéricos para mostrar mejor.
    """
    salida = df.copy()

    for col in salida.columns:
        if "FECHA" in col or "INTERVENCION" in col and "DIAS" not in col:
            try:
                salida[col] = pd.to_datetime(salida[col]).dt.strftime("%Y-%m-%d")
            except Exception:
                pass

    columnas_redondeo = [
        "PRODUCCION_PRCR",
        "AGUA_PRAG",
        "OIL_POR_INTERV",
        "PRODUCCION_HOJA_CONVERTIDO",
        "POTENCIAL_BOPD",
        "PRODUCCION_TOTAL_PRCR"
    ]

    for col in columnas_redondeo:
        if col in salida.columns:
            salida[col] = pd.to_numeric(salida[col], errors="coerce").round(2)

    return salida


def descargar_csv(df):
    return df.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# INTERFAZ
# ============================================================

st.title("🛢️ Control de pozos por Swab")
st.caption("App para revisar pozos convertidos por año, performance por unidad y pozos sin intervención reciente.")

archivo = st.file_uploader("Sube el Excel de Producción por Swab", type=["xlsx"])

if archivo is None:
    st.info("Sube el archivo Excel para comenzar.")
    st.stop()

bytes_excel = archivo.getvalue()

try:
    xls = cargar_excel(bytes_excel)
    datos = cargar_datos_swab(bytes_excel)
    base_pozos = cargar_swab_basica(bytes_excel)
except Exception as e:
    st.error(f"No se pudo cargar el Excel: {e}")
    st.stop()

fecha_min = datos["FECHA"].min().date()
fecha_max = datos["FECHA"].max().date()

with st.sidebar:
    st.header("Filtros")

    anios_disponibles = []
    for anio in [2024, 2025, 2026]:
        if encontrar_hoja(xls.sheet_names, f"Swab {anio} (Dia)") is not None:
            anios_disponibles.append(anio)

    if not anios_disponibles:
        st.error("No encontré hojas tipo 'Swab 2024 (Dia)', 'Swab 2025 (Dia)' o 'Swab 2026 (Dia)'.")
        st.stop()

    anio_sel = st.selectbox("Año de pozos convertidos", anios_disponibles, index=len(anios_disponibles)-1)

    modo_fecha = st.radio(
        "Periodo para calcular performance",
        ["Año seleccionado", "Últimos 30 días", "Últimos 90 días", "Rango manual"]
    )

    fecha_corte = st.date_input("Fecha de corte", value=fecha_max, min_value=fecha_min, max_value=fecha_max)

    if modo_fecha == "Año seleccionado":
        fecha_inicio = date(anio_sel, 1, 1)
        fecha_fin = min(date(anio_sel, 12, 31), fecha_corte)
    elif modo_fecha == "Últimos 30 días":
        fecha_fin = fecha_corte
        fecha_inicio = fecha_fin - pd.Timedelta(days=30)
    elif modo_fecha == "Últimos 90 días":
        fecha_fin = fecha_corte
        fecha_inicio = fecha_fin - pd.Timedelta(days=90)
    else:
        col_a, col_b = st.columns(2)
        with col_a:
            fecha_inicio = st.date_input("Desde", value=max(fecha_min, date(anio_sel, 1, 1)), min_value=fecha_min, max_value=fecha_max)
        with col_b:
            fecha_fin = st.date_input("Hasta", value=fecha_corte, min_value=fecha_min, max_value=fecha_max)

    umbral_dias = st.selectbox(
        "Pozos sin explotar/intervenir desde",
        [7, 30, 60, 90, 120, 150, 180, 365],
        format_func=lambda x: "1 semana" if x == 7 else f"{int(x/30)} mes(es)" if x < 365 else "1 año"
    )

    ordenar_por = st.selectbox(
        "Ordenar performance por",
        ["PRODUCCION_PRCR", "OIL_POR_INTERV", "INTERVENCIONES"],
        index=0
    )

convertidos = cargar_convertidos(bytes_excel, anio_sel)
perf_unidad = performance_por_pozo_unidad(datos, fecha_inicio, fecha_fin)
perf_total = performance_por_pozo_total(datos, fecha_inicio, fecha_fin)

if convertidos.empty:
    st.warning(f"No encontré pozos convertidos para {anio_sel}.")
    st.stop()

tabla_convertidos_unidad = convertidos.merge(
    perf_unidad,
    on="POZO_KEY",
    how="left",
    suffixes=("_CONVERTIDO", "")
)

tabla_convertidos_unidad["INTERVENCIONES"] = tabla_convertidos_unidad["INTERVENCIONES"].fillna(0).astype(int)
tabla_convertidos_unidad["PRODUCCION_PRCR"] = tabla_convertidos_unidad["PRODUCCION_PRCR"].fillna(0)
tabla_convertidos_unidad["AGUA_PRAG"] = tabla_convertidos_unidad["AGUA_PRAG"].fillna(0)
tabla_convertidos_unidad["OIL_POR_INTERV"] = tabla_convertidos_unidad["OIL_POR_INTERV"].fillna(0)

tabla_convertidos_unidad["POZO_MOSTRAR"] = tabla_convertidos_unidad["POZO_CONVERTIDO"]
tabla_convertidos_unidad.loc[
    tabla_convertidos_unidad["POZO_MOSTRAR"].isna(), "POZO_MOSTRAR"
] = tabla_convertidos_unidad["POZO"]

tabla_convertidos_total = convertidos.merge(
    perf_total,
    on="POZO_KEY",
    how="left",
    suffixes=("_CONVERTIDO", "")
)

tabla_convertidos_total["INTERVENCIONES"] = tabla_convertidos_total["INTERVENCIONES"].fillna(0).astype(int)
tabla_convertidos_total["PRODUCCION_PRCR"] = tabla_convertidos_total["PRODUCCION_PRCR"].fillna(0)
tabla_convertidos_total["OIL_POR_INTERV"] = tabla_convertidos_total["OIL_POR_INTERV"].fillna(0)

no_explotados = construir_no_explotados(datos, base_pozos, fecha_corte)

# ============================================================
# KPIs
# ============================================================

st.subheader(f"Resumen general al {fecha_corte}")

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

with kpi1:
    st.metric("Pozos convertidos", f"{convertidos['POZO_KEY'].nunique():,.0f}")

with kpi2:
    st.metric("Intervenciones", f"{tabla_convertidos_total['INTERVENCIONES'].sum():,.0f}")

with kpi3:
    st.metric("Producción PRCR", f"{tabla_convertidos_total['PRODUCCION_PRCR'].sum():,.2f}")

with kpi4:
    interv = tabla_convertidos_total["INTERVENCIONES"].sum()
    prod = tabla_convertidos_total["PRODUCCION_PRCR"].sum()
    prom = prod / interv if interv > 0 else 0
    st.metric("Oil / Interv.", f"{prom:,.2f}")

with kpi5:
    sin_prod = int((tabla_convertidos_total["INTERVENCIONES"] == 0).sum())
    st.metric("Convertidos sin intervención", f"{sin_prod:,.0f}")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs([
    "Convertidos y performance",
    "Pozos sin explotar",
    "Ranking y alertas",
    "Validación de datos"
])


# ============================================================
# TAB 1
# ============================================================

with tab1:
    st.subheader(f"Pozos convertidos en {anio_sel} con performance por unidad")
    st.caption(f"Periodo evaluado: {fecha_inicio} al {fecha_fin}")

    columnas_mostrar = [
        "AÑO_CONVERSION",
        "POZO_MOSTRAR",
        "BATERIA",
        "UNIDAD",
        "INTERVENCIONES",
        "PRODUCCION_PRCR",
        "AGUA_PRAG",
        "OIL_POR_INTERV",
        "PRIMERA_INTERVENCION",
        "ULTIMA_INTERVENCION",
        "PRODUCCION_HOJA_CONVERTIDO",
        "DIAS_CON_PRODUCCION_HOJA"
    ]

    columnas_mostrar = [c for c in columnas_mostrar if c in tabla_convertidos_unidad.columns]

    tabla1 = tabla_convertidos_unidad[columnas_mostrar].copy()
    tabla1 = tabla1.sort_values(ordenar_por, ascending=False)

    st.dataframe(
        formato_tabla(tabla1),
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "Descargar convertidos con performance por unidad",
        data=descargar_csv(formato_tabla(tabla1)),
        file_name=f"convertidos_{anio_sel}_performance_unidad.csv",
        mime="text/csv"
    )

    st.subheader("Gráfico de producción por pozo convertido")
    graf = (
        tabla_convertidos_total
        .groupby("POZO_CONVERTIDO", dropna=False)["PRODUCCION_PRCR"]
        .sum()
        .sort_values(ascending=False)
        .head(20)
    )
    st.bar_chart(graf)


# ============================================================
# TAB 2
# ============================================================

with tab2:
    st.subheader("Pozos sin explotar/intervenir recientemente")

    resumen = (
        no_explotados
        .groupby("RANGO_SIN_INTERVENCION")
        .agg(POZOS=("POZO", "count"))
        .reset_index()
    )

    orden_rangos = [
        "0 a 7 días",
        "8 a 30 días",
        "31 a 60 días",
        "61 a 90 días",
        "91 a 120 días",
        "121 a 180 días",
        "Más de 180 días",
        "Nunca"
    ]

    resumen["ORDEN"] = resumen["RANGO_SIN_INTERVENCION"].apply(
        lambda x: orden_rangos.index(x) if x in orden_rangos else 99
    )
    resumen = resumen.sort_values("ORDEN").drop(columns="ORDEN")

    c1, c2 = st.columns([1, 2])

    with c1:
        st.write("Resumen por antigüedad")
        st.dataframe(resumen, use_container_width=True, hide_index=True)

    with c2:
        st.write("Distribución")
        st.bar_chart(resumen.set_index("RANGO_SIN_INTERVENCION")["POZOS"])

    filtrado = no_explotados[
        (no_explotados["DIAS_SIN_INTERVENCION"] >= umbral_dias)
    ].copy()

    filtrado = filtrado.sort_values(
        ["POTENCIAL_BOPD", "DIAS_SIN_INTERVENCION"],
        ascending=[False, False]
    )

    st.caption(f"Mostrando pozos con {umbral_dias} días o más sin intervención. Fecha corte: {fecha_corte}")

    columnas_no_exp = [
        "POZO",
        "BATERIA",
        "POTENCIAL_BOPD",
        "ULTIMA_INTERVENCION",
        "DIAS_SIN_INTERVENCION",
        "RANGO_SIN_INTERVENCION",
        "INTERVENCIONES_TOTAL",
        "PRODUCCION_TOTAL_PRCR",
        "ESTADO"
    ]

    if "RESUMEN_BASE" in filtrado.columns:
        columnas_no_exp.append("RESUMEN_BASE")

    st.dataframe(
        formato_tabla(filtrado[columnas_no_exp]),
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "Descargar pozos sin explotar",
        data=descargar_csv(formato_tabla(filtrado[columnas_no_exp])),
        file_name=f"pozos_sin_intervencion_{umbral_dias}_dias.csv",
        mime="text/csv"
    )


# ============================================================
# TAB 3
# ============================================================

with tab3:
    st.subheader("Ranking y alertas operativas")

    col1, col2 = st.columns(2)

    with col1:
        st.write("Top 20 por producción PRCR")
        top_prod = tabla_convertidos_total.sort_values("PRODUCCION_PRCR", ascending=False).head(20)
        cols = ["POZO_CONVERTIDO", "BATERIA", "UNIDADES", "INTERVENCIONES", "PRODUCCION_PRCR", "OIL_POR_INTERV", "ULTIMA_INTERVENCION"]
        cols = [c for c in cols if c in top_prod.columns]
        st.dataframe(formato_tabla(top_prod[cols]), use_container_width=True, hide_index=True)

    with col2:
        st.write("Top 20 por oil/intervención")
        top_ratio = tabla_convertidos_total[
            tabla_convertidos_total["INTERVENCIONES"] > 0
        ].sort_values("OIL_POR_INTERV", ascending=False).head(20)
        cols = ["POZO_CONVERTIDO", "BATERIA", "UNIDADES", "INTERVENCIONES", "PRODUCCION_PRCR", "OIL_POR_INTERV", "ULTIMA_INTERVENCION"]
        cols = [c for c in cols if c in top_ratio.columns]
        st.dataframe(formato_tabla(top_ratio[cols]), use_container_width=True, hide_index=True)

    st.write("Convertidos sin intervención en el periodo seleccionado")
    sin_interv = tabla_convertidos_total[tabla_convertidos_total["INTERVENCIONES"] == 0].copy()
    cols = ["AÑO_CONVERSION", "POZO_CONVERTIDO", "PRODUCCION_HOJA_CONVERTIDO", "DIAS_CON_PRODUCCION_HOJA"]
    cols = [c for c in cols if c in sin_interv.columns]
    st.dataframe(formato_tabla(sin_interv[cols]), use_container_width=True, hide_index=True)

    st.write("Candidatos a revisar: alto potencial y varios días sin intervención")
    candidatos = no_explotados[
        (no_explotados["DIAS_SIN_INTERVENCION"] >= umbral_dias) &
        (no_explotados["POTENCIAL_BOPD"] > 0)
    ].sort_values(["POTENCIAL_BOPD", "DIAS_SIN_INTERVENCION"], ascending=[False, False]).head(30)

    cols = ["POZO", "BATERIA", "POTENCIAL_BOPD", "ULTIMA_INTERVENCION", "DIAS_SIN_INTERVENCION", "RANGO_SIN_INTERVENCION"]
    st.dataframe(formato_tabla(candidatos[cols]), use_container_width=True, hide_index=True)


# ============================================================
# TAB 4
# ============================================================

with tab4:
    st.subheader("Validación rápida de hojas y columnas")

    st.write("Hojas encontradas en el Excel:")
    st.dataframe(pd.DataFrame({"HOJA": xls.sheet_names}), use_container_width=True, hide_index=True)

    st.write("Columnas detectadas en Datos de Swab:")
    st.dataframe(pd.DataFrame({"COLUMNA": datos.columns}), use_container_width=True, hide_index=True)

    st.write("Rango de fechas detectado:")
    st.info(f"Desde {fecha_min} hasta {fecha_max}")

    st.write("Muestra de Datos de Swab:")
    muestra_cols = ["FECHA", "POZO", "BATERIA", "UNIDAD", "TSER", "PRCR", "PRAG"]
    muestra_cols = [c for c in muestra_cols if c in datos.columns]
    st.dataframe(formato_tabla(datos[muestra_cols].head(20)), use_container_width=True, hide_index=True)
