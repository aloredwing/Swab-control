import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Control SWAB", page_icon="🛢️", layout="wide")

# ============================================================
# LISTAS FIJAS DE POZOS CONVERTIDOS
# Estas listas quedan escritas dentro del código.
# Producción básica se calcula como todo pozo que no esté en estas listas.
# ============================================================

POZOS_CONVERTIDOS_2024 = [
    # Colocar aquí los 49 pozos convertidos 2024 si se requiere ajustar.
    # Ejemplo: "EA8204",
]

POZOS_CONVERTIDOS_2025 = [
    # Colocar aquí los 26 pozos convertidos 2025 si se requiere ajustar.
    # Ejemplo: "EA114049D",
]

POZOS_CONVERTIDOS_2026 = [
    # Pozos 2026 que se tenían identificados en la conversación previa.
    "EA8819",
    "EA8816D",
    "EA8662D",
    "EA11748D",
    "AA11396D",
    "EA11813D",
    "EA11968D",
    "EA6286D",
]

EXPECTED_COUNTS = {
    "Convertidos 2024": 49,
    "Convertidos 2025": 26,
    "Convertidos 2026": 20,
}

# ============================================================
# Utilidades
# ============================================================

def normalize_text(value: object) -> str:
    text = "" if value is None else str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def normalize_well(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", "", text)
    text = text.replace("-", "")
    return text


def find_column(columns: List[str], aliases: List[str]) -> Optional[str]:
    norm_cols = {c: normalize_text(c) for c in columns}
    norm_aliases = [normalize_text(a) for a in aliases]

    for col, norm in norm_cols.items():
        if norm in norm_aliases:
            return col

    for col, norm in norm_cols.items():
        for alias in norm_aliases:
            if alias and alias in norm:
                return col
    return None


def infer_well_column(df: pd.DataFrame) -> Optional[str]:
    columns = list(df.columns)
    direct = find_column(columns, ["pozo", "pozos", "well", "well_name", "nombre_pozo", "codigo_pozo", "codigo de pozo"])
    if direct:
        return direct

    best_col = None
    best_score = 0
    pattern = re.compile(r"\b[A-Z]{1,3}\s*\d{2,6}[A-Z]?\b", re.IGNORECASE)
    sample = df.head(5000)
    for col in columns:
        values = sample[col].astype(str)
        score = values.str.contains(pattern, na=False).sum()
        if score > best_score:
            best_score = int(score)
            best_col = col
    return best_col if best_score > 0 else None


def to_datetime_safe(series: pd.Series) -> pd.Series:
    result = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if result.notna().sum() == 0:
        result = pd.to_datetime(series, errors="coerce", dayfirst=False)
    return result


def to_numeric_safe(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)
    clean = (
        series.astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
    )
    return pd.to_numeric(clean, errors="coerce").fillna(0)


def read_first_sheet(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="latin1")
        df.columns = [str(c).strip() for c in df.columns]
        return df

    if name.endswith((".xlsx", ".xls")):
        data = uploaded_file.read()
        xls = pd.ExcelFile(BytesIO(data))
        first_sheet = xls.sheet_names[0]
        df = pd.read_excel(BytesIO(data), sheet_name=first_sheet)
        df.columns = [str(c).strip() for c in df.columns]
        return df

    raise ValueError("Formato no soportado")


def month_start_end(month_value: pd.Timestamp) -> Tuple[pd.Timestamp, pd.Timestamp]:
    start = pd.Timestamp(month_value).to_period("M").to_timestamp()
    end = start + pd.offsets.MonthEnd(0)
    return start, end


def list_to_set(values: List[str]) -> set:
    return {normalize_well(v) for v in values if normalize_well(v)}


SET_2024 = list_to_set(POZOS_CONVERTIDOS_2024)
SET_2025 = list_to_set(POZOS_CONVERTIDOS_2025)
SET_2026 = list_to_set(POZOS_CONVERTIDOS_2026)
SET_CONVERTIDOS = SET_2024 | SET_2025 | SET_2026


def assign_group(well_norm: str, year_value: object = None) -> str:
    if well_norm in SET_2024:
        return "Convertidos 2024"
    if well_norm in SET_2025:
        return "Convertidos 2025"
    if well_norm in SET_2026:
        return "Convertidos 2026"

    if year_value is not None and not pd.isna(year_value):
        year_text = str(year_value)
        if "2024" in year_text:
            return "Convertidos 2024"
        if "2025" in year_text:
            return "Convertidos 2025"
        if "2026" in year_text:
            return "Convertidos 2026"

    return "Producción básica"


def classify_drop(pct: float) -> str:
    if pd.isna(pct):
        return "Sin base"
    if pct >= 50:
        return "Crítica"
    if pct >= 30:
        return "Fuerte"
    if pct >= 15:
        return "Moderada"
    if pct > 0:
        return "Leve"
    return "Sin caída"


def prepare_dataframe(raw: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    columns = list(raw.columns)

    fecha_col = find_column(columns, ["fecha", "date", "dia", "día", "fecha_produccion", "fecha swab"])
    pozo_col = infer_well_column(raw)
    prcr_col = find_column(columns, ["prcr"])
    prag_col = find_column(columns, ["prag"])
    bateria_col = find_column(columns, ["bateria", "batería", "battery", "bat", "estacion", "estación"])
    gas_col = find_column(columns, ["gas", "qgas", "gas_mpc", "gas_mscf", "qg"])
    year_col = find_column(columns, ["anio", "año", "anio_conversion", "año_conversion", "conversion", "convertido"])

    missing = []
    if fecha_col is None:
        missing.append("FECHA")
    if pozo_col is None:
        missing.append("POZO")
    if prcr_col is None:
        missing.append("PRCR")
    if prag_col is None:
        missing.append("PRAG")

    if missing:
        st.error("Faltan columnas obligatorias o no fueron reconocidas: " + ", ".join(missing))
        st.write("Columnas encontradas en la hoja 1:", columns)
        st.dataframe(raw.head(20), use_container_width=True)
        st.stop()

    df = raw.copy()
    df[fecha_col] = to_datetime_safe(df[fecha_col])
    df = df[df[fecha_col].notna()].copy()

    df[pozo_col] = df[pozo_col].astype(str).str.strip()
    df["Pozo_norm"] = df[pozo_col].map(normalize_well)
    df = df[df["Pozo_norm"] != ""].copy()

    df[prcr_col] = to_numeric_safe(df[prcr_col])
    df[prag_col] = to_numeric_safe(df[prag_col])

    if gas_col:
        df[gas_col] = to_numeric_safe(df[gas_col])

    if bateria_col:
        df["Batería"] = df[bateria_col].astype(str).replace("nan", "Sin batería").fillna("Sin batería")
    else:
        df["Batería"] = "Sin batería"

    df["Intervenciones"] = 1
    df["Grupo"] = df.apply(
        lambda r: assign_group(r["Pozo_norm"], r.get(year_col) if year_col else None), axis=1
    )
    df["Mes"] = df[fecha_col].dt.to_period("M").dt.to_timestamp()
    df["Día"] = df[fecha_col].dt.floor("D")

    cols = {
        "fecha": fecha_col,
        "pozo": pozo_col,
        "prcr": prcr_col,
        "prag": prag_col,
        "gas": gas_col or "",
        "bateria_original": bateria_col or "",
        "year": year_col or "",
    }
    return df, cols


def build_status(df_scope: pd.DataFrame, selected_wells_norm: List[str], selected_month: pd.Timestamp, cols: Dict[str, str]) -> pd.DataFrame:
    start, end = month_start_end(selected_month)
    month_df = df_scope[(df_scope["Día"] >= start) & (df_scope["Día"] <= end)].copy()

    base = (
        df_scope[df_scope["Pozo_norm"].isin(selected_wells_norm)]
        .sort_values(cols["fecha"])
        .drop_duplicates("Pozo_norm", keep="last")
        [["Pozo_norm", cols["pozo"], "Batería", "Grupo"]]
        .copy()
    )

    prod_month = (
        month_df.groupby("Pozo_norm", dropna=False)
        .agg(
            PRCR_mes=(cols["prcr"], "sum"),
            PRAG_mes=(cols["prag"], "sum"),
            Intervenciones_mes=("Intervenciones", "sum"),
        )
        .reset_index()
    )

    last_prod = (
        df_scope[df_scope[cols["prcr"]] > 0]
        .groupby("Pozo_norm", dropna=False)[cols["fecha"]]
        .max()
        .reset_index(name="Última fecha con PRCR")
    )

    out = base.merge(prod_month, on="Pozo_norm", how="left").merge(last_prod, on="Pozo_norm", how="left")
    for c in ["PRCR_mes", "PRAG_mes", "Intervenciones_mes"]:
        out[c] = out[c].fillna(0)
    out["Estado"] = np.where(out["PRCR_mes"] > 0, "Activo", "Inactivo")
    out["Mes evaluado"] = start.strftime("%Y-%m")
    out["Días desde última producción"] = np.where(
        out["Última fecha con PRCR"].notna(),
        (end - out["Última fecha con PRCR"]).dt.days,
        np.nan,
    )
    return out.sort_values(["Estado", "PRCR_mes", cols["pozo"]], ascending=[True, False, True])


def daily_summary(df: pd.DataFrame, metric: str, segment: str) -> pd.DataFrame:
    if segment == "Total general":
        out = (
            df.groupby("Día", dropna=False)
            .agg(Producción=(metric, "sum"), Intervenciones=("Intervenciones", "sum"), Pozos=("Pozo_norm", "nunique"))
            .reset_index()
        )
        out["Segmento"] = "Total general"
        return out[["Día", "Segmento", "Producción", "Intervenciones", "Pozos"]]

    out = (
        df.groupby(["Día", segment], dropna=False)
        .agg(Producción=(metric, "sum"), Intervenciones=("Intervenciones", "sum"), Pozos=("Pozo_norm", "nunique"))
        .reset_index()
        .rename(columns={segment: "Segmento"})
    )
    out["Segmento"] = out["Segmento"].astype(str)
    return out[["Día", "Segmento", "Producción", "Intervenciones", "Pozos"]]


def month_drop(df_scope: pd.DataFrame, selected_month: pd.Timestamp, cols: Dict[str, str]) -> pd.DataFrame:
    start, end = month_start_end(selected_month)
    prev_start = start - pd.DateOffset(months=1)
    prev_start, prev_end = month_start_end(prev_start)

    curr = df_scope[(df_scope["Día"] >= start) & (df_scope["Día"] <= end)].copy()
    prev = df_scope[(df_scope["Día"] >= prev_start) & (df_scope["Día"] <= prev_end)].copy()

    keys = ["Pozo_norm", cols["pozo"], "Batería", "Grupo"]
    base = df_scope[keys].drop_duplicates("Pozo_norm")
    curr_sum = curr.groupby("Pozo_norm", dropna=False)[cols["prcr"]].sum().reset_index(name="PRCR mes actual")
    prev_sum = prev.groupby("Pozo_norm", dropna=False)[cols["prcr"]].sum().reset_index(name="PRCR mes anterior")

    out = base.merge(prev_sum, on="Pozo_norm", how="left").merge(curr_sum, on="Pozo_norm", how="left")
    out[["PRCR mes anterior", "PRCR mes actual"]] = out[["PRCR mes anterior", "PRCR mes actual"]].fillna(0)
    out["Caída PRCR"] = out["PRCR mes anterior"] - out["PRCR mes actual"]
    out["Caída %"] = np.where(out["PRCR mes anterior"] > 0, out["Caída PRCR"] / out["PRCR mes anterior"] * 100, np.nan)
    out["Clasificación"] = out["Caída %"].apply(classify_drop)
    return out.sort_values(["Caída PRCR", "Caída %"], ascending=[False, False])


def make_excel_report(tables: Dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, table in tables.items():
            safe_name = name[:31]
            table.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


def fmt(value: float, decimals: int = 2) -> str:
    if pd.isna(value):
        value = 0
    return f"{value:,.{decimals}f}"

# ============================================================
# Interfaz
# ============================================================

st.title("🛢️ Control SWAB en Streamlit")
st.caption("Versión corregida: trabaja con la hoja 1, usa PRCR para petróleo y PRAG para agua, y ejecuta recién cuando presionas el botón.")

with st.expander("Configuración fija de esta versión", expanded=True):
    st.markdown(
        """
Esta app usa solo la primera hoja del archivo cargado. No pide seleccionar hojas ni mapear columnas manualmente.

Columnas esperadas: FECHA, POZO, PRCR y PRAG. Si el nombre de pozo no se llama exactamente POZO, la app lo intenta detectar por los códigos de pozo.

Producción básica se calcula como todos los pozos que no están en las listas fijas de convertidos 2024, 2025 y 2026.

Activo significa que el pozo tuvo PRCR mayor que cero dentro del mes calendario seleccionado. El corte es del día 1 al último día del mes, no treinta días hacia atrás.
        """
    )

uploaded_file = st.sidebar.file_uploader("Carga archivo SWAB", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    st.info("Carga tu archivo SWAB para iniciar.")
    st.stop()

try:
    raw_df = read_first_sheet(uploaded_file)
except Exception as exc:
    st.error(f"No se pudo leer el archivo: {exc}")
    st.stop()

if raw_df.empty:
    st.error("La primera hoja está vacía.")
    st.stop()

df, cols = prepare_dataframe(raw_df)

if df.empty:
    st.error("No quedaron datos válidos después de limpiar fechas y pozos.")
    st.stop()

st.sidebar.header("Columnas usadas")
st.sidebar.write(f"Fecha: {cols['fecha']}")
st.sidebar.write(f"Pozo: {cols['pozo']}")
st.sidebar.write(f"Petróleo: {cols['prcr']}")
st.sidebar.write(f"Agua: {cols['prag']}")
st.sidebar.write(f"Batería: {'Batería' if cols['bateria_original'] else 'Sin batería'}")

# Aviso de listas incompletas, útil para no clasificar mal.
counts_fixed = {
    "Convertidos 2024": len(SET_2024),
    "Convertidos 2025": len(SET_2025),
    "Convertidos 2026": len(SET_2026),
}
missing_counts = [f"{k}: cargados {counts_fixed[k]} de {EXPECTED_COUNTS[k]}" for k in EXPECTED_COUNTS if counts_fixed[k] != EXPECTED_COUNTS[k]]
if missing_counts:
    st.warning(
        "Revisa el bloque de listas fijas al inicio del app.py. "
        + " | ".join(missing_counts)
        + ". La clasificación de básica depende de que esas listas estén completas."
    )

st.sidebar.header("Filtros antes de ejecutar")
min_date = df["Día"].min().date()
max_date = df["Día"].max().date()
date_range = st.sidebar.date_input("Rango de fechas para gráficas", value=(min_date, max_date), min_value=min_date, max_value=max_date)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date = pd.to_datetime(date_range[0])
    end_date = pd.to_datetime(date_range[1])
else:
    start_date = pd.to_datetime(min_date)
    end_date = pd.to_datetime(max_date)

all_groups = ["Convertidos 2024", "Convertidos 2025", "Convertidos 2026", "Producción básica"]
available_groups = [g for g in all_groups if g in df["Grupo"].unique()]
selected_groups = st.sidebar.multiselect("Grupos de pozos", available_groups, default=available_groups)

baterias = sorted(df["Batería"].astype(str).unique())
selected_baterias = st.sidebar.multiselect("Baterías", baterias, default=baterias)

scope_no_date = df[(df["Grupo"].isin(selected_groups)) & (df["Batería"].astype(str).isin(selected_baterias))].copy()

st.sidebar.markdown("### Pozos por grupo")
for group in selected_groups:
    n = scope_no_date.loc[scope_no_date["Grupo"] == group, "Pozo_norm"].nunique()
    st.sidebar.write(f"{group}: {n}")

well_options_df = (
    scope_no_date.sort_values(cols["pozo"])
    .drop_duplicates("Pozo_norm")
    [["Pozo_norm", cols["pozo"], "Grupo", "Batería"]]
)
well_labels = [f"{r[cols['pozo']]} | {r['Grupo']} | {r['Batería']}" for _, r in well_options_df.iterrows()]
well_norm_by_label = dict(zip(well_labels, well_options_df["Pozo_norm"]))
selected_labels = st.sidebar.multiselect("Pozos a analizar", well_labels, default=well_labels)
selected_wells_norm = [well_norm_by_label[label] for label in selected_labels]

available_months = sorted(df["Mes"].dropna().unique())
month_labels = [pd.Timestamp(m).strftime("%Y-%m") for m in available_months]
default_month = month_labels[-1] if month_labels else ""
selected_month_label = st.sidebar.selectbox("Mes calendario para activo/inactivo", month_labels, index=month_labels.index(default_month) if default_month in month_labels else 0)
selected_month = pd.to_datetime(selected_month_label + "-01")

st.sidebar.markdown("---")
run = st.sidebar.button("Ejecutar análisis", type="primary", use_container_width=True)

if not run:
    st.subheader("Archivo cargado correctamente")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Filas leídas", f"{len(df):,}")
    c2.metric("Pozos detectados", f"{df['Pozo_norm'].nunique():,}")
    c3.metric("Fecha mínima", str(min_date))
    c4.metric("Fecha máxima", str(max_date))

    st.markdown("### Vista previa de la hoja 1")
    st.dataframe(raw_df.head(30), use_container_width=True)
    st.info("Selecciona grupos, baterías, pozos y mes de corte. Luego presiona Ejecutar análisis en el panel izquierdo.")
    st.stop()

if not selected_wells_norm:
    st.error("Selecciona al menos un pozo para analizar.")
    st.stop()

filtered_no_date = scope_no_date[scope_no_date["Pozo_norm"].isin(selected_wells_norm)].copy()
filtered = filtered_no_date[(filtered_no_date["Día"] >= start_date) & (filtered_no_date["Día"] <= end_date)].copy()

if filtered.empty:
    st.warning("No hay datos en el rango seleccionado.")
    st.stop()

status_df = build_status(filtered_no_date, selected_wells_norm, selected_month, cols)
status_map = status_df[["Pozo_norm", "Estado"]].drop_duplicates("Pozo_norm")
filtered = filtered.merge(status_map, on="Pozo_norm", how="left")
filtered["Estado"] = filtered["Estado"].fillna("Sin estado")

st.subheader("Resumen general")
active = int((status_df["Estado"] == "Activo").sum()) if not status_df.empty else 0
inactive = int((status_df["Estado"] == "Inactivo").sum()) if not status_df.empty else 0
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Pozos seleccionados", f"{len(selected_wells_norm)}")
k2.metric("Activos", f"{active}")
k3.metric("Inactivos", f"{inactive}")
k4.metric("PRCR total", fmt(filtered[cols["prcr"]].sum(), 2))
k5.metric("PRAG total", fmt(filtered[cols["prag"]].sum(), 2))
k6.metric("Intervenciones", fmt(filtered["Intervenciones"].sum(), 0))

m_start, m_end = month_start_end(selected_month)
st.caption(f"Estado activo/inactivo evaluado con PRCR desde {m_start.date()} hasta {m_end.date()}.")

tab_diario, tab_estado, tab_mensual, tab_interv, tab_caidas, tab_pozo, tab_descarga = st.tabs(
    [
        "Producción diaria",
        "Activos e inactivos",
        "Mensual",
        "Intervenciones",
        "Caídas",
        "Detalle por pozo",
        "Descarga",
    ]
)

with tab_diario:
    st.subheader("Producción diaria SWAB")
    metric = st.selectbox("Variable", [cols["prcr"], cols["prag"]], index=0, key="metric_diario")
    segment_options = ["Total general", "Grupo", "Batería", cols["pozo"], "Estado"]
    segment = st.selectbox("Segmentar por", segment_options, index=0)
    chart_type = st.selectbox("Tipo de gráfica", ["Barras", "Línea"], index=0)

    ds = daily_summary(filtered, metric, segment)
    total_day = ds.groupby("Día")["Producción"].sum().reset_index()
    days_prod = int((total_day["Producción"] > 0).sum())
    avg_prod = total_day["Producción"].mean() if not total_day.empty else 0
    wells_prod = int(filtered.loc[filtered[metric] > 0, "Pozo_norm"].nunique())

    d1, d2, d3, d4 = st.columns(4)
    d1.metric(f"Total {metric}", fmt(ds["Producción"].sum(), 2))
    d2.metric("Días con producción", f"{days_prod}")
    d3.metric("Promedio diario", fmt(avg_prod, 2))
    d4.metric("Pozos con producción", f"{wells_prod}")

    plot_df = ds.copy()
    if segment != "Total general":
        top_segments = plot_df.groupby("Segmento")["Producción"].sum().sort_values(ascending=False).head(20).index
        plot_df = plot_df[plot_df["Segmento"].isin(top_segments)]

    if chart_type == "Barras":
        fig = px.bar(plot_df, x="Día", y="Producción", color="Segmento" if segment != "Total general" else None, hover_data=["Intervenciones", "Pozos"])
    else:
        fig = px.line(plot_df, x="Día", y="Producción", color="Segmento" if segment != "Total general" else None, markers=True, hover_data=["Intervenciones", "Pozos"])
    fig.update_layout(height=520, title=f"Producción diaria {metric}", hovermode="x unified")
    st.plotly_chart(fig, use_container_width=True)

    total_day["Promedio móvil 7 días"] = total_day["Producción"].rolling(7, min_periods=1).mean()
    fig_ma = go.Figure()
    fig_ma.add_bar(x=total_day["Día"], y=total_day["Producción"], name="Producción diaria")
    fig_ma.add_scatter(x=total_day["Día"], y=total_day["Promedio móvil 7 días"], mode="lines", name="Promedio móvil 7 días")
    fig_ma.update_layout(height=420, title=f"Tendencia diaria {metric}", hovermode="x unified")
    st.plotly_chart(fig_ma, use_container_width=True)

    st.dataframe(ds, use_container_width=True)

with tab_estado:
    st.subheader("Pozos activos e inactivos")
    st.markdown(f"Activo: PRCR mayor que cero en el mes {selected_month_label}.")
    f_estado = st.multiselect("Filtrar estado", ["Activo", "Inactivo"], default=["Activo", "Inactivo"])
    view_status = status_df[status_df["Estado"].isin(f_estado)].copy()

    c1, c2 = st.columns([1, 2])
    with c1:
        fig_status = px.pie(status_df, names="Estado", title="Distribución")
        st.plotly_chart(fig_status, use_container_width=True)
    with c2:
        st.dataframe(view_status, use_container_width=True)

with tab_mensual:
    st.subheader("Producción mensual")
    metric_m = st.selectbox("Variable mensual", [cols["prcr"], cols["prag"]], index=0, key="metric_mensual")
    monthly = filtered.groupby(["Mes", "Grupo", "Batería"], dropna=False)[metric_m].sum().reset_index()
    fig_m = px.bar(monthly, x="Mes", y=metric_m, color="Grupo", facet_col="Batería" if monthly["Batería"].nunique() <= 4 else None, title=f"{metric_m} mensual")
    fig_m.update_layout(height=560)
    st.plotly_chart(fig_m, use_container_width=True)

    monthly_avg = filtered.groupby("Mes", dropna=False)[metric_m].sum().reset_index()
    fig_line = px.line(monthly_avg, x="Mes", y=metric_m, markers=True, title=f"Tendencia mensual {metric_m}")
    fig_line.update_layout(height=420)
    st.plotly_chart(fig_line, use_container_width=True)
    st.dataframe(monthly, use_container_width=True)

with tab_interv:
    st.subheader("Intervenciones por mes")
    interv = filtered.groupby(["Mes", "Grupo", "Batería"], dropna=False)["Intervenciones"].sum().reset_index()
    fig_i = px.bar(interv, x="Mes", y="Intervenciones", color="Grupo", title="Intervenciones mensuales")
    fig_i.update_layout(height=500)
    st.plotly_chart(fig_i, use_container_width=True)

    interv_well = filtered.groupby(["Mes", cols["pozo"], "Grupo", "Batería"], dropna=False)["Intervenciones"].sum().reset_index().sort_values(["Mes", "Intervenciones"], ascending=[True, False])
    st.dataframe(interv_well, use_container_width=True)

with tab_caidas:
    st.subheader("Caídas de PRCR por mes calendario")
    top_n = st.slider("Top de caídas", 5, 50, 20)
    drops = month_drop(filtered_no_date, selected_month, cols)
    st.dataframe(drops.head(top_n), use_container_width=True)
    fig_d = px.bar(
        drops.head(top_n),
        x=cols["pozo"],
        y="Caída PRCR",
        color="Clasificación",
        hover_data=["Grupo", "Batería", "PRCR mes anterior", "PRCR mes actual", "Caída %"],
        title=f"Top {top_n} caídas PRCR: {selected_month_label} vs mes anterior",
    )
    fig_d.update_layout(height=500)
    st.plotly_chart(fig_d, use_container_width=True)

with tab_pozo:
    st.subheader("Detalle por pozo")
    wells_detail = sorted(filtered[cols["pozo"]].astype(str).unique())
    pozo_detail = st.selectbox("Pozo", wells_detail)
    one = filtered[filtered[cols["pozo"]].astype(str) == str(pozo_detail)].copy().sort_values("Día")

    daily_one = one.groupby("Día", dropna=False).agg(PRCR=(cols["prcr"], "sum"), PRAG=(cols["prag"], "sum"), Intervenciones=("Intervenciones", "sum")).reset_index()
    fig_p = go.Figure()
    fig_p.add_trace(go.Scatter(x=daily_one["Día"], y=daily_one["PRCR"], mode="lines+markers", name="PRCR"))
    fig_p.add_trace(go.Scatter(x=daily_one["Día"], y=daily_one["PRAG"], mode="lines+markers", name="PRAG"))
    fig_p.add_trace(go.Bar(x=daily_one["Día"], y=daily_one["Intervenciones"], name="Intervenciones", opacity=0.35, yaxis="y2"))
    fig_p.update_layout(
        title=f"Detalle diario del pozo {pozo_detail}",
        height=520,
        yaxis=dict(title="Producción"),
        yaxis2=dict(title="Intervenciones", overlaying="y", side="right"),
        hovermode="x unified",
    )
    st.plotly_chart(fig_p, use_container_width=True)
    st.dataframe(one, use_container_width=True)

with tab_descarga:
    st.subheader("Tablas y descarga")
    resumen_pozo = (
        filtered.groupby([cols["pozo"], "Grupo", "Batería"], dropna=False)
        .agg(PRCR=(cols["prcr"], "sum"), PRAG=(cols["prag"], "sum"), Intervenciones=("Intervenciones", "sum"), Días=("Día", "nunique"))
        .reset_index()
        .sort_values("PRCR", ascending=False)
    )
    resumen_mensual_pozo = (
        filtered.groupby(["Mes", cols["pozo"], "Grupo", "Batería"], dropna=False)
        .agg(PRCR=(cols["prcr"], "sum"), PRAG=(cols["prag"], "sum"), Intervenciones=("Intervenciones", "sum"))
        .reset_index()
        .sort_values(["Mes", cols["pozo"]])
    )
    diario_total = daily_summary(filtered, cols["prcr"], "Total general")
    drops_report = month_drop(filtered_no_date, selected_month, cols)

    st.markdown("### Resumen por pozo")
    st.dataframe(resumen_pozo, use_container_width=True)
    st.markdown("### Resumen mensual por pozo")
    st.dataframe(resumen_mensual_pozo, use_container_width=True)

    report = make_excel_report(
        {
            "datos_filtrados": filtered,
            "activos_inactivos": status_df,
            "resumen_pozo": resumen_pozo,
            "mensual_pozo": resumen_mensual_pozo,
            "produccion_diaria": diario_total,
            "caidas_prcr": drops_report,
        }
    )
    st.download_button(
        "Descargar reporte Excel",
        data=report,
        file_name="reporte_swab_streamlit.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.caption("La clasificación básica se obtiene por descarte: todo pozo que no esté en las listas fijas de convertidos queda como producción básica.")
