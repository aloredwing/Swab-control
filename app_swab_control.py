import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional, Set, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


st.set_page_config(
    page_title="Análisis SWAB",
    page_icon="🛢️",
    layout="wide",
)


# ============================================================
# Utilidades generales
# ============================================================


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def find_column(columns: List[str], aliases: List[str]) -> Optional[str]:
    normalized_map = {col: normalize_text(col) for col in columns}
    normalized_aliases = [normalize_text(a) for a in aliases]

    for col, norm_col in normalized_map.items():
        if norm_col in normalized_aliases:
            return col

    for col, norm_col in normalized_map.items():
        for alias in normalized_aliases:
            if alias and alias in norm_col:
                return col
    return None


def default_index(options: List[str], value: Optional[str]) -> int:
    if value in options:
        return options.index(value)
    return 0


REQUIRED_PLACEHOLDER = "Selecciona columna"


def required_column_selectbox(label: str, columns: List[str], detected: Optional[str]) -> str:
    """Selectbox para columnas obligatorias.

    Si no se detecta una columna, no fuerza la primera columna del Excel,
    porque eso puede hacer que FECHA quede también como POZO.
    """
    options = [REQUIRED_PLACEHOLDER] + columns
    index = default_index(options, detected) if detected else 0
    return st.sidebar.selectbox(label, options, index=index)


def validate_required_mapping(fecha_col: str, pozo_col: str, columns: List[str], preview_df: pd.DataFrame) -> None:
    missing = []
    if fecha_col == REQUIRED_PLACEHOLDER:
        missing.append("Columna fecha")
    if pozo_col == REQUIRED_PLACEHOLDER:
        missing.append("Columna pozo")

    if missing:
        st.warning("Selecciona las columnas obligatorias antes de continuar: " + ", ".join(missing))
        with st.expander("Ver columnas detectadas en el archivo", expanded=True):
            st.write(list(columns))
            st.dataframe(preview_df.head(10), use_container_width=True)
        st.stop()

    if fecha_col == pozo_col:
        st.error(
            "Mapeo incorrecto: la columna de fecha y la columna de pozo no pueden ser la misma. "
            "En tu captura, POZO quedó seleccionado como FECHA. Cambia 'Columna pozo' por la columna real del pozo."
        )
        with st.expander("Vista rápida para escoger la columna correcta", expanded=True):
            st.write("Columnas disponibles:", list(columns))
            st.dataframe(preview_df.head(15), use_container_width=True)
        st.stop()


def to_datetime_safe(series: pd.Series) -> pd.Series:
    result = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if result.isna().mean() > 0.5:
        result = pd.to_datetime(series, errors="coerce", dayfirst=False)
    return result


def to_numeric_safe(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    clean = (
        series.astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
    )
    return pd.to_numeric(clean, errors="coerce")


def normalize_well(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", "", text)
    return text


def split_wells(text: str) -> List[str]:
    if not text:
        return []
    parts = re.split(r"[,;\n\t]+", text)
    return [normalize_well(p) for p in parts if normalize_well(p)]


def read_uploaded_file(uploaded_file) -> Tuple[Dict[str, pd.DataFrame], str]:
    """Lee CSV o Excel. Si es Excel retorna todas las hojas para poder detectar listas 2024, 2025 y 2026."""
    if uploaded_file is None:
        return {}, ""

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="latin1")
        return {"Datos CSV": df}, "Datos CSV"

    if filename.endswith((".xlsx", ".xls")):
        data = uploaded_file.read()
        xls = pd.ExcelFile(BytesIO(data))
        sheets: Dict[str, pd.DataFrame] = {}
        for sheet in xls.sheet_names:
            try:
                temp = pd.read_excel(BytesIO(data), sheet_name=sheet)
                temp.columns = [str(c).strip() for c in temp.columns]
                sheets[sheet] = temp
            except Exception:
                pass
        preferred = pick_default_data_sheet(list(sheets.keys()))
        return sheets, preferred

    st.error("Formato no soportado. Carga un archivo .xlsx, .xls o .csv")
    return {}, ""


def pick_default_data_sheet(sheet_names: List[str]) -> str:
    if not sheet_names:
        return ""
    priority = ["datos_de_swab", "datos_swab", "produccion", "produccion_swab", "data"]
    normalized = {sheet: normalize_text(sheet) for sheet in sheet_names}
    for key in priority:
        for sheet, norm in normalized.items():
            if key in norm:
                return sheet
    return sheet_names[0]


def pick_default_conversion_sheet(sheet_names: List[str], year: int) -> str:
    normalized = {sheet: normalize_text(sheet) for sheet in sheet_names}
    candidates = []
    for sheet, norm in normalized.items():
        if str(year) in norm and "estado" not in norm:
            score = 0
            if "swab" in norm:
                score += 3
            if "dia" in norm:
                score += 2
            if "convert" in norm:
                score += 2
            candidates.append((score, sheet))
    if not candidates:
        return "No usar hoja"
    candidates.sort(reverse=True)
    return candidates[0][1]


def pick_default_basic_sheet(sheet_names: List[str]) -> str:
    normalized = {sheet: normalize_text(sheet) for sheet in sheet_names}
    for sheet, norm in normalized.items():
        if "basica" in norm or "basico" in norm:
            return sheet
    for sheet, norm in normalized.items():
        if "estado" in norm and "2024" in norm:
            return sheet
    return "No usar hoja"


def extract_wells_from_sheet(sheet_df: Optional[pd.DataFrame]) -> Set[str]:
    if sheet_df is None or sheet_df.empty:
        return set()
    columns = [str(c).strip() for c in sheet_df.columns]
    well_col = find_column(columns, ["pozo", "well", "nombre pozo", "codigo pozo", "id pozo"])
    if well_col is None:
        # Si no reconoce la columna, usa la primera columna con más valores tipo texto.
        best_col = None
        best_count = -1
        for col in columns:
            count = sheet_df[col].astype(str).str.contains(r"[A-Z]{1,3}\s*\d+", case=False, regex=True, na=False).sum()
            if count > best_count:
                best_col = col
                best_count = count
        well_col = best_col
    if well_col is None:
        return set()
    return set(sheet_df[well_col].map(normalize_well).dropna().loc[lambda s: s != ""].unique())


def build_group_for_well(
    well: str,
    year_value: object,
    year_col_enabled: bool,
    wells_2024: Set[str],
    wells_2025: Set[str],
    wells_2026: Set[str],
) -> str:
    well_norm = normalize_well(well)

    if year_col_enabled:
        year = pd.to_numeric(pd.Series([year_value]), errors="coerce").iloc[0]
        if pd.notna(year):
            year = int(year)
            if year in [2024, 2025, 2026]:
                return f"Convertidos {year}"

    if well_norm in wells_2024:
        return "Convertidos 2024"
    if well_norm in wells_2025:
        return "Convertidos 2025"
    if well_norm in wells_2026:
        return "Convertidos 2026"
    return "Producción básica"


def classify_drop(row):
    pct = row.get("Caída %", 0)
    if pd.isna(pct):
        return "Sin dato"
    if pct >= 50:
        return "Crítica"
    if pct >= 30:
        return "Fuerte"
    if pct >= 15:
        return "Moderada"
    if pct > 0:
        return "Leve"
    return "Sin caída"


def get_monthly(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    temp = df.copy()
    temp["Mes"] = temp[date_col].dt.to_period("M").dt.to_timestamp()
    return temp


def month_limits(month_start: pd.Timestamp) -> Tuple[pd.Timestamp, pd.Timestamp]:
    start = pd.Timestamp(month_start).to_period("M").to_timestamp()
    end = start + pd.offsets.MonthEnd(0)
    return start, end


def period_drop_by_well_calendar_month(
    df: pd.DataFrame,
    date_col: str,
    well_col: str,
    value_col: str,
    battery_col: str,
    current_month: pd.Timestamp,
) -> pd.DataFrame:
    if df.empty or value_col == "No aplica":
        return pd.DataFrame()

    current_start, current_end = month_limits(current_month)
    previous_start = current_start - pd.DateOffset(months=1)
    previous_start, previous_end = month_limits(previous_start)

    current = df[(df[date_col] >= current_start) & (df[date_col] <= current_end)].copy()
    previous = df[(df[date_col] >= previous_start) & (df[date_col] <= previous_end)].copy()

    group_cols = [well_col, battery_col, "Grupo conversión"]
    curr_sum = current.groupby(group_cols, dropna=False)[value_col].sum().reset_index(name="Producción mes de corte")
    prev_sum = previous.groupby(group_cols, dropna=False)[value_col].sum().reset_index(name="Producción mes anterior")

    base_keys = df[group_cols].drop_duplicates()
    result = base_keys.merge(prev_sum, on=group_cols, how="left").merge(curr_sum, on=group_cols, how="left")
    result[["Producción mes anterior", "Producción mes de corte"]] = result[
        ["Producción mes anterior", "Producción mes de corte"]
    ].fillna(0)
    result["Caída producción"] = result["Producción mes anterior"] - result["Producción mes de corte"]
    result["Caída %"] = np.where(
        result["Producción mes anterior"] > 0,
        result["Caída producción"] / result["Producción mes anterior"] * 100,
        np.nan,
    )
    result["Clasificación"] = result.apply(classify_drop, axis=1)
    return result.sort_values(["Caída producción", "Caída %"], ascending=[False, False])


def period_drop_by_well_days(
    df: pd.DataFrame,
    date_col: str,
    well_col: str,
    value_col: str,
    battery_col: str,
    days: int,
) -> pd.DataFrame:
    if df.empty or value_col == "No aplica":
        return pd.DataFrame()

    max_date = df[date_col].max()
    if pd.isna(max_date):
        return pd.DataFrame()

    current_start = max_date - pd.Timedelta(days=days - 1)
    previous_end = current_start - pd.Timedelta(days=1)
    previous_start = previous_end - pd.Timedelta(days=days - 1)

    current = df[(df[date_col] >= current_start) & (df[date_col] <= max_date)].copy()
    previous = df[(df[date_col] >= previous_start) & (df[date_col] <= previous_end)].copy()

    group_cols = [well_col, battery_col, "Grupo conversión"]
    curr_sum = current.groupby(group_cols, dropna=False)[value_col].sum().reset_index(name="Periodo actual")
    prev_sum = previous.groupby(group_cols, dropna=False)[value_col].sum().reset_index(name="Periodo anterior")

    result = pd.merge(prev_sum, curr_sum, on=group_cols, how="outer").fillna(0)
    result["Caída"] = result["Periodo anterior"] - result["Periodo actual"]
    result["Caída %"] = np.where(
        result["Periodo anterior"] > 0,
        result["Caída"] / result["Periodo anterior"] * 100,
        np.nan,
    )
    result["Clasificación"] = result.apply(classify_drop, axis=1)
    return result.sort_values(["Caída", "Caída %"], ascending=[False, False])


def build_active_inactive_table(
    df_scope: pd.DataFrame,
    selected_wells: List[str],
    date_col: str,
    well_col: str,
    battery_col: str,
    status_value_col: str,
    month_start_selected: pd.Timestamp,
) -> pd.DataFrame:
    if df_scope.empty:
        return pd.DataFrame()

    m_start, m_end = month_limits(month_start_selected)
    selected_set = set(normalize_well(w) for w in selected_wells)
    temp = df_scope[df_scope["Pozo_norm"].isin(selected_set)].copy()

    if temp.empty:
        return pd.DataFrame()

    month_df = temp[(temp[date_col] >= m_start) & (temp[date_col] <= m_end)].copy()

    base = temp[[well_col, "Pozo_norm", battery_col, "Grupo conversión"]].drop_duplicates("Pozo_norm").copy()

    if status_value_col != "Intervenciones_calc":
        monthly_prod = (
            month_df.groupby("Pozo_norm", dropna=False)[status_value_col]
            .sum()
            .reset_index(name="Producción mes de corte")
        )
        total_prod = (
            temp.groupby("Pozo_norm", dropna=False)[status_value_col]
            .sum()
            .reset_index(name="Producción total filtrada")
        )
        last_prod_date = (
            temp[temp[status_value_col] > 0]
            .groupby("Pozo_norm", dropna=False)[date_col]
            .max()
            .reset_index(name="Última fecha con producción")
        )
    else:
        monthly_prod = (
            month_df.groupby("Pozo_norm", dropna=False)["Intervenciones_calc"]
            .sum()
            .reset_index(name="Producción mes de corte")
        )
        total_prod = (
            temp.groupby("Pozo_norm", dropna=False)["Intervenciones_calc"]
            .sum()
            .reset_index(name="Producción total filtrada")
        )
        last_prod_date = (
            temp[temp["Intervenciones_calc"] > 0]
            .groupby("Pozo_norm", dropna=False)[date_col]
            .max()
            .reset_index(name="Última fecha con producción")
        )

    result = base.merge(monthly_prod, on="Pozo_norm", how="left")
    result = result.merge(total_prod, on="Pozo_norm", how="left")
    result = result.merge(last_prod_date, on="Pozo_norm", how="left")
    result[["Producción mes de corte", "Producción total filtrada"]] = result[
        ["Producción mes de corte", "Producción total filtrada"]
    ].fillna(0)
    result["Estado"] = np.where(result["Producción mes de corte"] > 0, "Activo", "Inactivo")
    result["Mes evaluado"] = m_start.strftime("%Y-%m")
    result["Días desde última producción"] = np.where(
        result["Última fecha con producción"].notna(),
        (m_end - result["Última fecha con producción"]).dt.days,
        np.nan,
    )
    result = result.sort_values(["Estado", "Producción mes de corte", well_col], ascending=[True, False, True])
    return result


def format_number(value, decimals=2):
    if pd.isna(value):
        return "0"
    return f"{value:,.{decimals}f}"


def prepare_daily_data(
    filtered: pd.DataFrame,
    status_df: pd.DataFrame,
    date_col: str,
) -> pd.DataFrame:
    temp = filtered.copy()
    temp["Día"] = temp[date_col].dt.floor("D")
    if status_df is not None and not status_df.empty:
        status_map = status_df[["Pozo_norm", "Estado"]].drop_duplicates("Pozo_norm")
        temp = temp.merge(status_map, on="Pozo_norm", how="left")
        temp["Estado"] = temp["Estado"].fillna("Sin estado")
    else:
        temp["Estado"] = "Sin estado"
    return temp


def build_daily_summary(
    daily_df: pd.DataFrame,
    metric_col: str,
    segment_col: str,
) -> pd.DataFrame:
    if daily_df.empty:
        return pd.DataFrame()

    if segment_col == "Total general":
        grouped = (
            daily_df.groupby("Día", dropna=False)
            .agg(
                Producción=(metric_col, "sum"),
                Intervenciones=("Intervenciones_calc", "sum"),
                Pozos_con_dato=("Pozo_norm", "nunique"),
            )
            .reset_index()
        )
        grouped["Segmento"] = "Total general"
        return grouped[["Día", "Segmento", "Producción", "Intervenciones", "Pozos_con_dato"]]

    grouped = (
        daily_df.groupby(["Día", segment_col], dropna=False)
        .agg(
            Producción=(metric_col, "sum"),
            Intervenciones=("Intervenciones_calc", "sum"),
            Pozos_con_dato=("Pozo_norm", "nunique"),
        )
        .reset_index()
        .rename(columns={segment_col: "Segmento"})
    )
    grouped["Segmento"] = grouped["Segmento"].astype(str)
    return grouped[["Día", "Segmento", "Producción", "Intervenciones", "Pozos_con_dato"]]


def top_segments_by_total(daily_summary: pd.DataFrame, top_n: int) -> List[str]:
    if daily_summary.empty or "Segmento" not in daily_summary.columns:
        return []
    totals = (
        daily_summary.groupby("Segmento", dropna=False)["Producción"]
        .sum()
        .sort_values(ascending=False)
        .head(top_n)
    )
    return list(totals.index.astype(str))


def make_daily_heatmap(
    daily_df: pd.DataFrame,
    date_col: str,
    pozo_col: str,
    metric_col: str,
    top_n_heatmap: int,
) -> pd.DataFrame:
    if daily_df.empty:
        return pd.DataFrame()
    well_totals = (
        daily_df.groupby(pozo_col, dropna=False)[metric_col]
        .sum()
        .sort_values(ascending=False)
        .head(top_n_heatmap)
    )
    wells = list(well_totals.index.astype(str))
    temp = daily_df[daily_df[pozo_col].astype(str).isin(wells)].copy()
    temp["Día"] = temp[date_col].dt.floor("D")
    pivot = (
        temp.pivot_table(index=pozo_col, columns="Día", values=metric_col, aggfunc="sum", fill_value=0)
        .reindex(wells)
    )
    return pivot


def make_excel_report(
    filtered: pd.DataFrame,
    status_df: pd.DataFrame,
    summary_well: pd.DataFrame,
    summary_month_well: pd.DataFrame,
    monthly_drop: pd.DataFrame,
    daily_summary: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Datos filtrados")
        status_df.to_excel(writer, index=False, sheet_name="Activos inactivos")
        summary_well.to_excel(writer, index=False, sheet_name="Resumen pozo")
        summary_month_well.to_excel(writer, index=False, sheet_name="Mensual pozo")
        if monthly_drop is not None and not monthly_drop.empty:
            monthly_drop.to_excel(writer, index=False, sheet_name="Caida mensual")
        if daily_summary is not None and not daily_summary.empty:
            daily_summary.to_excel(writer, index=False, sheet_name="Produccion diaria")
    return output.getvalue()


# ============================================================
# Interfaz principal
# ============================================================

st.title("🛢️ Análisis SWAB en Streamlit")
st.caption("Carga tu Excel o CSV, selecciona los pozos y ejecuta el análisis.")

with st.expander("Qué hace esta versión", expanded=True):
    st.markdown(
        """
Esta versión permite cargar el archivo, escoger los grupos de pozos convertidos 2024, 2025, 2026 o producción básica, seleccionar pozos específicos y recién ejecutar el análisis.

El estado activo o inactivo se calcula por mes calendario. Es decir, si seleccionas mayo, el corte será del 01 de mayo al último día de mayo. No usa 30 días hacia atrás.

Criterio usado: si el pozo no produjo en el mes calendario seleccionado, queda como inactivo. La variable para definir producción la puedes escoger en el panel lateral.

Además, al ejecutar el análisis se abre primero una vista de producción diaria SWAB de todos los pozos seleccionados. Esa vista puede segmentarse por batería, grupo de conversión, pozo o estado activo/inactivo.
        """
    )

uploaded_file = st.sidebar.file_uploader(
    "Carga archivo SWAB",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is None:
    st.info("Carga un archivo .xlsx, .xls o .csv para iniciar.")
    st.stop()

sheets, default_sheet = read_uploaded_file(uploaded_file)

if not sheets:
    st.error("No se pudo leer información válida del archivo.")
    st.stop()

sheet_names = list(sheets.keys())

st.sidebar.header("Hojas del archivo")
main_sheet = st.sidebar.selectbox(
    "Hoja principal de producción",
    sheet_names,
    index=default_index(sheet_names, default_sheet),
)

raw_df = sheets[main_sheet].copy()
raw_df.columns = [str(c).strip() for c in raw_df.columns]
columns = list(raw_df.columns)

if raw_df.empty:
    st.error("La hoja seleccionada está vacía.")
    st.stop()

# Detección automática de columnas principales
fecha_auto = find_column(columns, ["fecha", "date", "dia", "día", "periodo", "fecha produccion", "fecha swab"])
pozo_auto = find_column(
    columns,
    [
        "pozo", "pozos", "well", "wells", "nombre pozo", "codigo pozo", "id pozo",
        "pozo swab", "codigo", "codigo de pozo", "well name", "wellname", "uwi"
    ],
)
bateria_auto = find_column(columns, ["bateria", "batería", "battery", "estacion", "estación", "planta"])
gas_auto = find_column(columns, ["gas", "qgas", "q gas", "gas mscf", "gas mpcd", "gas mmscf", "produccion gas", "producción gas", "prod gas"])
petroleo_auto = find_column(columns, ["prcr", "petroleo", "petróleo", "oil", "condensado", "cond", "bopd", "bls petroleo", "bls petróleo", "prod petroleo", "produccion"])
agua_auto = find_column(columns, ["agua", "water", "bwpd", "bls agua", "agua bbl"])
interv_auto = find_column(columns, ["intervencion", "intervención", "intervenciones", "swab", "servicio", "evento", "trabajo"])
year_auto = find_column(columns, ["anio conversion", "año conversion", "year conversion", "conversion", "convertido"])

st.sidebar.header("Mapeo de columnas")
options_required = columns
options_optional = ["No aplica"] + columns

fecha_col = required_column_selectbox("Columna fecha", options_required, fecha_auto)
pozo_col = required_column_selectbox("Columna pozo", options_required, pozo_auto)
validate_required_mapping(fecha_col, pozo_col, columns, raw_df)

bateria_col = st.sidebar.selectbox("Columna batería", options_optional, index=default_index(options_optional, bateria_auto))
gas_col = st.sidebar.selectbox("Columna gas", options_optional, index=default_index(options_optional, gas_auto))
petroleo_col = st.sidebar.selectbox("Columna petróleo, condensado o PRCR", options_optional, index=default_index(options_optional, petroleo_auto))
agua_col = st.sidebar.selectbox("Columna agua", options_optional, index=default_index(options_optional, agua_auto))
interv_col = st.sidebar.selectbox("Columna intervenciones", options_optional, index=default_index(options_optional, interv_auto))
year_col = st.sidebar.selectbox("Columna año de conversión real", options_optional, index=default_index(options_optional, year_auto))

# Detección de listas de pozos convertidos desde hojas del Excel
st.sidebar.header("Listas de pozos por grupo")
st.sidebar.caption("La app intenta leer automáticamente las hojas de convertidos. También puedes pegarlos manualmente.")

sheet_options = ["No usar hoja"] + sheet_names
sheet_2024 = st.sidebar.selectbox(
    "Hoja pozos convertidos 2024",
    sheet_options,
    index=default_index(sheet_options, pick_default_conversion_sheet(sheet_names, 2024)),
)
sheet_2025 = st.sidebar.selectbox(
    "Hoja pozos convertidos 2025",
    sheet_options,
    index=default_index(sheet_options, pick_default_conversion_sheet(sheet_names, 2025)),
)
sheet_2026 = st.sidebar.selectbox(
    "Hoja pozos convertidos 2026",
    sheet_options,
    index=default_index(sheet_options, pick_default_conversion_sheet(sheet_names, 2026)),
)
sheet_basic = st.sidebar.selectbox(
    "Hoja padrón producción básica",
    sheet_options,
    index=default_index(sheet_options, pick_default_basic_sheet(sheet_names)),
)

wells_2024_auto = extract_wells_from_sheet(sheets.get(sheet_2024)) if sheet_2024 != "No usar hoja" else set()
wells_2025_auto = extract_wells_from_sheet(sheets.get(sheet_2025)) if sheet_2025 != "No usar hoja" else set()
wells_2026_auto = extract_wells_from_sheet(sheets.get(sheet_2026)) if sheet_2026 != "No usar hoja" else set()
wells_basic_auto = extract_wells_from_sheet(sheets.get(sheet_basic)) if sheet_basic != "No usar hoja" else set()

with st.sidebar.expander("Pegar pozos manualmente"):
    manual_2024 = st.text_area("Pozos convertidos 2024", value="", height=90)
    manual_2025 = st.text_area("Pozos convertidos 2025", value="", height=90)
    manual_2026 = st.text_area("Pozos convertidos 2026", value="", height=90)

wells_2024 = wells_2024_auto.union(split_wells(manual_2024))
wells_2025 = wells_2025_auto.union(split_wells(manual_2025))
wells_2026 = wells_2026_auto.union(split_wells(manual_2026))
converted_all = wells_2024.union(wells_2025).union(wells_2026)

# Preparación de datos
df = raw_df.copy()
df[fecha_col] = to_datetime_safe(df[fecha_col])
df = df[df[fecha_col].notna()].copy()

if df.empty:
    st.error("No se encontraron fechas válidas. Revisa la columna fecha.")
    st.stop()

for col in [gas_col, petroleo_col, agua_col, interv_col]:
    if col and col != "No aplica":
        df[col] = to_numeric_safe(df[col])

if bateria_col == "No aplica":
    df["Batería"] = "Sin batería"
    bateria_use = "Batería"
else:
    bateria_use = bateria_col
    df[bateria_use] = df[bateria_use].astype(str).replace("nan", "Sin batería").fillna("Sin batería")

df[pozo_col] = df[pozo_col].astype(str).str.strip()
df["Pozo_norm"] = df[pozo_col].map(normalize_well)
df = df[df["Pozo_norm"] != ""].copy()

use_year_col = year_col != "No aplica"
df["Grupo conversión"] = df.apply(
    lambda row: build_group_for_well(
        row[pozo_col],
        row.get(year_col) if use_year_col else None,
        use_year_col,
        wells_2024,
        wells_2025,
        wells_2026,
    ),
    axis=1,
)

# Si existe padrón básico, conserva esos pozos como universo básico aunque no estén en la hoja principal.
# La producción básica dentro del análisis se mantiene como todo lo no convertido.

st.sidebar.header("Criterios")
each_row_is_intervention = st.sidebar.checkbox("Cada fila representa una intervención", value=True)
resample_monthly = st.sidebar.checkbox("Graficar detalle por pozo en mensual", value=True)
top_n = st.sidebar.slider("Top de caídas a mostrar", min_value=5, max_value=50, value=15, step=1)

if each_row_is_intervention:
    df["Intervenciones_calc"] = 1
else:
    if interv_col and interv_col != "No aplica":
        df["Intervenciones_calc"] = df[interv_col].fillna(0)
    else:
        df["Intervenciones_calc"] = 1

# Variables numéricas disponibles
production_candidates = []
if gas_col != "No aplica":
    production_candidates.append(gas_col)
if petroleo_col != "No aplica":
    production_candidates.append(petroleo_col)
if agua_col != "No aplica":
    production_candidates.append(agua_col)
if not production_candidates:
    production_candidates = ["Intervenciones_calc"]

status_value_col = st.sidebar.selectbox(
    "Variable para definir activo/inactivo",
    production_candidates,
    index=0,
    help="Si la suma de esta variable en el mes calendario elegido es mayor que cero, el pozo queda activo.",
)

# Filtros principales
st.sidebar.header("Filtros")
# Refuerzo defensivo: conserva fecha como datetime aunque el usuario cambie el mapeo.
df[fecha_col] = to_datetime_safe(df[fecha_col])
df = df[df[fecha_col].notna()].copy()
if df.empty:
    st.error("No quedaron fechas válidas después del mapeo. Revisa la columna fecha.")
    st.stop()
min_timestamp = df[fecha_col].min()
max_timestamp = df[fecha_col].max()
if pd.isna(min_timestamp) or pd.isna(max_timestamp):
    st.error("No se pudo determinar el rango de fechas. Revisa la columna fecha.")
    st.stop()
min_date = min_timestamp.date()
max_date = max_timestamp.date()
date_range = st.sidebar.date_input("Rango de fecha para gráficas", value=(min_date, max_date), min_value=min_date, max_value=max_date)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
else:
    start_date, end_date = pd.to_datetime(min_date), pd.to_datetime(max_date)

baterias = sorted(df[bateria_use].dropna().astype(str).unique())
selected_baterias = st.sidebar.multiselect("Baterías", baterias, default=baterias)

group_order = ["Convertidos 2024", "Convertidos 2025", "Convertidos 2026", "Producción básica"]
available_groups = [g for g in group_order if g in df["Grupo conversión"].unique()]
if not available_groups:
    available_groups = ["Producción básica"]
selected_groups = st.sidebar.multiselect("Grupo de pozos", available_groups, default=available_groups)

# Universo de pozos según grupo y batería, pero sin aplicar rango de fecha para que no se pierdan los inactivos.
df_scope_no_date = df[
    (df[bateria_use].astype(str).isin(selected_baterias))
    & (df["Grupo conversión"].isin(selected_groups))
].copy()

pozos_by_group = {
    group: sorted(df_scope_no_date.loc[df_scope_no_date["Grupo conversión"] == group, pozo_col].astype(str).unique())
    for group in selected_groups
}

st.sidebar.markdown("### Selección de pozos")
with st.sidebar.expander("Ver cantidad por grupo", expanded=True):
    for group in selected_groups:
        st.write(f"{group}: {len(pozos_by_group.get(group, []))} pozos")

pozos_all = sorted(df_scope_no_date[pozo_col].dropna().astype(str).unique())
selected_pozos = st.sidebar.multiselect("Pozos a analizar", pozos_all, default=pozos_all)

available_months = sorted(df[fecha_col].dt.to_period("M").dt.to_timestamp().unique())
month_labels = [pd.Timestamp(m).strftime("%Y-%m") for m in available_months]
default_month_label = pd.Timestamp(available_months[-1]).strftime("%Y-%m") if available_months else ""
month_label = st.sidebar.selectbox(
    "Mes calendario para estado activo/inactivo",
    month_labels,
    index=month_labels.index(default_month_label) if default_month_label in month_labels else 0,
)
selected_month = pd.to_datetime(month_label + "-01")

run_analysis = st.sidebar.button("Ejecutar análisis", type="primary", use_container_width=True)

if not run_analysis:
    st.info("Carga el archivo, selecciona grupos y pozos en el panel lateral, luego presiona **Ejecutar análisis**.")

    st.subheader("Listas detectadas")
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Convertidos 2024", len(wells_2024))
    col_b.metric("Convertidos 2025", len(wells_2025))
    col_c.metric("Convertidos 2026", len(wells_2026))
    col_d.metric("Básica detectada", len(wells_basic_auto))

    with st.expander("Ver pozos detectados por grupo"):
        for title, wells in {
            "Convertidos 2024": wells_2024,
            "Convertidos 2025": wells_2025,
            "Convertidos 2026": wells_2026,
            "Padrón producción básica": wells_basic_auto,
        }.items():
            st.markdown(f"**{title}**")
            st.write(", ".join(sorted(wells)) if wells else "Sin lista detectada")
    st.stop()

# Aplicar filtros finales
filtered_no_date = df_scope_no_date[df_scope_no_date[pozo_col].astype(str).isin(selected_pozos)].copy()
filtered = filtered_no_date[(filtered_no_date[fecha_col] >= start_date) & (filtered_no_date[fecha_col] <= end_date)].copy()

if filtered.empty:
    st.warning("No hay datos con los filtros seleccionados para el rango de fecha de gráficas.")
    st.stop()

# Estado activo e inactivo con mes calendario
status_df = build_active_inactive_table(
    df_scope=filtered_no_date,
    selected_wells=selected_pozos,
    date_col=fecha_col,
    well_col=pozo_col,
    battery_col=bateria_use,
    status_value_col=status_value_col,
    month_start_selected=selected_month,
)

# KPIs
st.subheader("Resumen general")
active_count = int((status_df["Estado"] == "Activo").sum()) if not status_df.empty else 0
inactive_count = int((status_df["Estado"] == "Inactivo").sum()) if not status_df.empty else 0

kpi_cols = st.columns(6)
kpi_cols[0].metric("Pozos seleccionados", f"{len(selected_pozos)}")
kpi_cols[1].metric("Pozos activos", f"{active_count}")
kpi_cols[2].metric("Pozos inactivos", f"{inactive_count}")
kpi_cols[3].metric("Baterías", f"{filtered[bateria_use].nunique()}")
kpi_cols[4].metric("Intervenciones", format_number(filtered["Intervenciones_calc"].sum(), 0))
kpi_cols[5].metric(f"Producción total {status_value_col}", format_number(filtered[status_value_col].sum(), 2))

st.caption(
    f"Estado activo/inactivo evaluado con corte mensual {selected_month.strftime('%Y-%m')}: "
    f"desde {month_limits(selected_month)[0].date()} hasta {month_limits(selected_month)[1].date()}."
)

# Tabs principales
daily_summary_for_report = pd.DataFrame()

tab_diario, tab_estado, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Producción diaria SWAB",
    "Activos e inactivos",
    "Producción mensual",
    "Intervenciones",
    "Caídas",
    "Detalle por pozo",
    "Tablas y descarga",
])

with tab_diario:
    st.subheader("Producción diaria SWAB de los pozos seleccionados")
    st.caption("Esta vista toma todos los pozos seleccionados en el panel lateral y permite segmentar la producción diaria.")

    daily_metric = st.selectbox(
        "Variable diaria a analizar",
        production_candidates,
        index=0,
        key="daily_metric",
    )

    segment_options = [
        "Total general",
        bateria_use,
        "Grupo conversión",
        pozo_col,
        "Estado",
    ]
    segment_options = list(dict.fromkeys(segment_options))

    cseg1, cseg2, cseg3 = st.columns([1.2, 1, 1])
    segment_choice = cseg1.selectbox(
        "Segmentar por",
        segment_options,
        index=0,
        key="daily_segment",
    )
    chart_kind = cseg2.selectbox(
        "Tipo de gráfica",
        ["Barras", "Línea"],
        index=0,
        key="daily_chart_kind",
    )
    top_daily_segments = cseg3.slider(
        "Top segmentos",
        min_value=5,
        max_value=50,
        value=15,
        step=1,
        key="top_daily_segments",
    )

    daily_df = prepare_daily_data(filtered, status_df, fecha_col)
    daily_summary = build_daily_summary(daily_df, daily_metric, segment_choice)
    daily_summary_for_report = daily_summary.copy()

    if daily_summary.empty:
        st.warning("No hay información diaria con los filtros seleccionados.")
    else:
        total_daily_prod = daily_summary["Producción"].sum()
        days_with_activity = int((daily_summary.groupby("Día")["Producción"].sum() > 0).sum())
        avg_daily_prod = daily_summary.groupby("Día")["Producción"].sum().mean()
        wells_with_activity = int(daily_df.loc[daily_df[daily_metric] > 0, "Pozo_norm"].nunique())

        kd1, kd2, kd3, kd4 = st.columns(4)
        kd1.metric(f"Total {daily_metric}", format_number(total_daily_prod, 2))
        kd2.metric("Días con producción", f"{days_with_activity}")
        kd3.metric("Promedio diario", format_number(avg_daily_prod, 2))
        kd4.metric("Pozos con producción", f"{wells_with_activity}")

        plot_daily = daily_summary.copy()
        if segment_choice != "Total general":
            top_segments = top_segments_by_total(plot_daily, top_daily_segments)
            plot_daily = plot_daily[plot_daily["Segmento"].astype(str).isin(top_segments)].copy()
            st.caption(f"Se muestran los {len(top_segments)} segmentos con mayor producción acumulada para mantener la gráfica legible.")

        if chart_kind == "Barras":
            fig_daily = px.bar(
                plot_daily,
                x="Día",
                y="Producción",
                color="Segmento" if segment_choice != "Total general" else None,
                title=f"Producción diaria SWAB por {segment_choice}",
                hover_data=["Intervenciones", "Pozos_con_dato"],
            )
        else:
            fig_daily = px.line(
                plot_daily,
                x="Día",
                y="Producción",
                color="Segmento" if segment_choice != "Total general" else None,
                markers=True,
                title=f"Producción diaria SWAB por {segment_choice}",
                hover_data=["Intervenciones", "Pozos_con_dato"],
            )

        fig_daily.update_layout(
            height=540,
            xaxis_title="Día",
            yaxis_title=daily_metric,
            legend_title_text=segment_choice,
            hovermode="x unified",
        )
        st.plotly_chart(fig_daily, use_container_width=True)

        show_ma = st.checkbox("Mostrar promedio móvil de 7 días del total diario", value=True)
        if show_ma:
            total_by_day = (
                daily_summary.groupby("Día", dropna=False)["Producción"]
                .sum()
                .reset_index()
                .sort_values("Día")
            )
            total_by_day["Promedio móvil 7 días"] = total_by_day["Producción"].rolling(7, min_periods=1).mean()
            fig_ma = go.Figure()
            fig_ma.add_trace(go.Bar(x=total_by_day["Día"], y=total_by_day["Producción"], name="Producción diaria"))
            fig_ma.add_trace(go.Scatter(x=total_by_day["Día"], y=total_by_day["Promedio móvil 7 días"], mode="lines", name="Promedio móvil 7 días"))
            fig_ma.update_layout(
                title="Producción diaria total y tendencia de 7 días",
                height=430,
                xaxis_title="Día",
                yaxis_title=daily_metric,
                hovermode="x unified",
            )
            st.plotly_chart(fig_ma, use_container_width=True)

        if st.checkbox("Mostrar mapa diario por pozo", value=False):
            top_heatmap = st.slider("Cantidad de pozos en el mapa", 5, 40, 20, key="top_heatmap")
            pivot = make_daily_heatmap(daily_df, fecha_col, pozo_col, daily_metric, top_heatmap)
            if pivot.empty:
                st.info("No hay datos para el mapa diario por pozo.")
            else:
                fig_heat = px.imshow(
                    pivot,
                    aspect="auto",
                    labels=dict(x="Día", y="Pozo", color=daily_metric),
                    title=f"Mapa de producción diaria por pozo, top {top_heatmap}",
                )
                fig_heat.update_layout(height=max(420, 24 * len(pivot.index)))
                st.plotly_chart(fig_heat, use_container_width=True)

        st.markdown("### Tabla diaria segmentada")
        st.dataframe(daily_summary, use_container_width=True)

        csv_daily = daily_summary.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar producción diaria CSV",
            data=csv_daily,
            file_name="swab_produccion_diaria.csv",
            mime="text/csv",
        )

with tab_estado:
    st.subheader("Listado de pozos activos e inactivos")
    st.markdown(
        f"Criterio: pozo **activo** si la suma de **{status_value_col}** en el mes **{selected_month.strftime('%Y-%m')}** es mayor que cero."
    )

    if status_df.empty:
        st.warning("No se pudo generar la tabla de estado.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Activos", active_count)
        c2.metric("Inactivos", inactive_count)
        pct_inactive = inactive_count / max(active_count + inactive_count, 1) * 100
        c3.metric("Inactivos %", f"{pct_inactive:.1f}%")

        fig_status = px.pie(
            status_df,
            names="Estado",
            title="Distribución de pozos activos e inactivos",
        )
        st.plotly_chart(fig_status, use_container_width=True)

        status_filter = st.multiselect(
            "Filtrar estado",
            ["Activo", "Inactivo"],
            default=["Activo", "Inactivo"],
            key="status_filter_main",
        )
        status_view = status_df[status_df["Estado"].isin(status_filter)].copy()
        st.dataframe(status_view, use_container_width=True)

        csv_status = status_view.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar activos e inactivos CSV",
            data=csv_status,
            file_name="swab_activos_inactivos.csv",
            mime="text/csv",
        )

with tab1:
    st.subheader("Producción mensual por batería y grupo")
    monthly = get_monthly(filtered, fecha_col)

    numeric_cols = []
    if gas_col != "No aplica":
        numeric_cols.append(gas_col)
    if petroleo_col != "No aplica":
        numeric_cols.append(petroleo_col)
    if agua_col != "No aplica":
        numeric_cols.append(agua_col)

    if not numeric_cols:
        st.warning("No hay columnas de producción seleccionadas.")
    else:
        metric_choice = st.selectbox("Variable a graficar", numeric_cols, key="metric_monthly")
        grouped_month = (
            monthly.groupby(["Mes", bateria_use, "Grupo conversión"], dropna=False)[metric_choice]
            .sum()
            .reset_index()
        )

        fig = px.bar(
            grouped_month,
            x="Mes",
            y=metric_choice,
            color=bateria_use,
            facet_col="Grupo conversión" if grouped_month["Grupo conversión"].nunique() <= 4 else None,
            title=f"{metric_choice} mensual por batería",
        )
        fig.update_layout(height=540, legend_title_text="Batería")
        st.plotly_chart(fig, use_container_width=True)

        avg_month = grouped_month.groupby("Mes", dropna=False)[metric_choice].mean().reset_index(name="Promedio mensual")
        fig_avg = px.line(
            avg_month,
            x="Mes",
            y="Promedio mensual",
            markers=True,
            title=f"Promedio mensual de {metric_choice}",
        )
        fig_avg.update_layout(height=420)
        st.plotly_chart(fig_avg, use_container_width=True)

with tab2:
    st.subheader("Intervenciones por mes")
    monthly = get_monthly(filtered, fecha_col)
    interv_month = (
        monthly.groupby(["Mes", bateria_use], dropna=False)["Intervenciones_calc"]
        .sum()
        .reset_index()
    )

    fig_int = px.bar(
        interv_month,
        x="Mes",
        y="Intervenciones_calc",
        color=bateria_use,
        title="Cantidad de intervenciones por mes y batería",
    )
    fig_int.update_layout(height=500, yaxis_title="Intervenciones")
    st.plotly_chart(fig_int, use_container_width=True)

    interv_well = (
        monthly.groupby(["Mes", bateria_use, pozo_col, "Grupo conversión"], dropna=False)["Intervenciones_calc"]
        .sum()
        .reset_index()
        .sort_values(["Mes", bateria_use, "Intervenciones_calc"], ascending=[True, True, False])
    )
    st.dataframe(interv_well, use_container_width=True)

with tab3:
    st.subheader("Caídas de producción")

    drop_metric = st.selectbox("Variable para caídas", production_candidates, index=0, key="drop_metric")

    st.markdown("### Caída por mes calendario")
    monthly_drop = period_drop_by_well_calendar_month(
        filtered_no_date,
        fecha_col,
        pozo_col,
        drop_metric,
        bateria_use,
        selected_month,
    )

    if monthly_drop.empty:
        st.info("No hay datos suficientes para calcular caída mensual.")
    else:
        display_drop = monthly_drop.head(top_n).copy()
        st.dataframe(display_drop, use_container_width=True)
        fig_drop_month = px.bar(
            display_drop,
            x=pozo_col,
            y="Caída producción",
            color="Clasificación",
            hover_data=[bateria_use, "Grupo conversión", "Producción mes anterior", "Producción mes de corte", "Caída %"],
            title=f"Top {top_n} caídas, mes {selected_month.strftime('%Y-%m')} vs mes anterior",
        )
        fig_drop_month.update_layout(height=460, xaxis_title="Pozo", yaxis_title="Caída")
        st.plotly_chart(fig_drop_month, use_container_width=True)

    st.markdown("### Caídas por ventanas móviles")
    st.caption("Esta parte sí usa 7, 30 y 365 días hacia atrás. El estado activo/inactivo usa mes calendario.")
    for title, days in {"Últimos 7 días": 7, "Últimos 30 días": 30, "Últimos 365 días": 365}.items():
        with st.expander(title, expanded=False):
            drop_df = period_drop_by_well_days(filtered, fecha_col, pozo_col, drop_metric, bateria_use, days)
            if drop_df.empty:
                st.info("No hay datos suficientes para este cálculo.")
                continue
            st.dataframe(drop_df.head(top_n), use_container_width=True)

with tab4:
    st.subheader("Detalle por pozo")
    selected_well_detail = st.selectbox("Selecciona pozo para detalle", sorted(filtered[pozo_col].astype(str).unique()))
    well_df = filtered[filtered[pozo_col].astype(str) == str(selected_well_detail)].copy()
    well_df = well_df.sort_values(fecha_col)

    if resample_monthly:
        well_df["Mes"] = well_df[fecha_col].dt.to_period("M").dt.to_timestamp()
        agg_dict = {"Intervenciones_calc": "sum"}
        for col in production_candidates:
            if col != "Intervenciones_calc":
                agg_dict[col] = "sum"
        plot_df = well_df.groupby("Mes", dropna=False).agg(agg_dict).reset_index()
        x_col = "Mes"
    else:
        agg_dict = {"Intervenciones_calc": "sum"}
        for col in production_candidates:
            if col != "Intervenciones_calc":
                agg_dict[col] = "sum"
        plot_df = well_df.groupby(fecha_col, dropna=False).agg(agg_dict).reset_index()
        x_col = fecha_col

    fig_detail = make_subplots(specs=[[{"secondary_y": True}]])

    for col in production_candidates:
        if col == "Intervenciones_calc":
            continue
        fig_detail.add_trace(
            go.Scatter(x=plot_df[x_col], y=plot_df[col], mode="lines+markers", name=col),
            secondary_y=False,
        )

    fig_detail.add_trace(
        go.Bar(x=plot_df[x_col], y=plot_df["Intervenciones_calc"], name="Intervenciones", opacity=0.35),
        secondary_y=True,
    )

    fig_detail.update_layout(
        title=f"Producción e intervenciones del pozo {selected_well_detail}",
        height=540,
        hovermode="x unified",
    )
    fig_detail.update_yaxes(title_text="Producción", secondary_y=False)
    fig_detail.update_yaxes(title_text="Intervenciones", secondary_y=True)
    st.plotly_chart(fig_detail, use_container_width=True)

    st.dataframe(well_df, use_container_width=True)

with tab5:
    st.subheader("Tablas consolidadas y descarga")
    monthly = get_monthly(filtered, fecha_col)

    agg_dict = {"Intervenciones_calc": "sum"}
    for col in production_candidates:
        if col != "Intervenciones_calc":
            agg_dict[col] = "sum"

    summary_well = (
        monthly.groupby([bateria_use, pozo_col, "Grupo conversión"], dropna=False)
        .agg(agg_dict)
        .reset_index()
        .sort_values("Intervenciones_calc", ascending=False)
    )
    st.markdown("### Consolidado por pozo")
    st.dataframe(summary_well, use_container_width=True)

    summary_month_well = (
        monthly.groupby(["Mes", bateria_use, pozo_col, "Grupo conversión"], dropna=False)
        .agg(agg_dict)
        .reset_index()
        .sort_values(["Mes", bateria_use, pozo_col])
    )
    st.markdown("### Consolidado mensual por pozo")
    st.dataframe(summary_month_well, use_container_width=True)

    report = make_excel_report(
        filtered=filtered,
        status_df=status_df,
        summary_well=summary_well,
        summary_month_well=summary_month_well,
        monthly_drop=monthly_drop if "monthly_drop" in locals() else pd.DataFrame(),
        daily_summary=daily_summary_for_report,
    )
    st.download_button(
        "Descargar reporte Excel",
        data=report,
        file_name="reporte_swab.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("Nota: el estado activo/inactivo usa mes calendario completo; las ventanas móviles de caídas se muestran solo como análisis complementario.")
