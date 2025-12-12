import io
import re
import requests
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# =========================================================
# Configuración general
# =========================================================
st.set_page_config(
    page_title="Última Inspección por Patente – Aeropuerto",
    layout="wide"
)

# =========================================================
# URLs Google Sheets (exportación a Excel)
# =========================================================
PATENTES_XLSX_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1DSZA9DJkxWHBIOMyulW3638TF46Udq7tCWh489UAOAs/"
    "export?format=xlsx&gid=1266707607"
)

INSPECCIONES_XLSX_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1UR2GTh6l4nwmx4DdY6yX_BS7EF1l-LYHtzHQdkYx03c/"
    "export?format=xlsx"
)

# =========================================================
# Utilidades
# =========================================================
def normalize_plate(x: str) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def try_parse_date(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, errors="coerce", dayfirst=True)
    mask = dt.isna()
    if mask.any():
        numeric = pd.to_numeric(series[mask], errors="coerce")
        dt2 = pd.to_datetime(
            numeric, unit="D", origin="1899-12-30", errors="coerce"
        )
        dt.loc[mask] = dt2
    return dt

def load_excel_from_gsheets(url: str) -> pd.DataFrame:
    r = requests.get(url, timeout=30)
    ct = r.headers.get("content-type", "").lower()

    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}")
    if "text/html" in ct:
        raise RuntimeError("Google devolvió HTML (permisos/login)")

    return pd.read_excel(io.BytesIO(r.content))

def traffic_light(days, green_max, yellow_max):
    if pd.isna(days):
        return "⚫ Sin inspección"
    if days <= green_max:
        return "🟢 OK"
    if days <= yellow_max:
        return "🟡 Alerta"
    return "🔴 Crítico"

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="resultado")
    return output.getvalue()

# =========================================================
# Título
# =========================================================
st.title("🚗 Última inspección por patente – Aeropuerto")

# =========================================================
# Descargas rápidas
# =========================================================
st.subheader("⬇️ Descargas rápidas (Google Sheets → Excel)")

c1, c2 = st.columns(2)
with c1:
    st.link_button("📄 Descargar Base Patentes", PATENTES_XLSX_URL)
with c2:
    st.link_button("📝 Descargar Inspecciones", INSPECCIONES_XLSX_URL)

st.divider()

# =========================================================
# Fuente de datos
# =========================================================
st.subheader("📌 Fuente de datos")

source = st.radio(
    "¿Cómo deseas cargar los datos?",
    [
        "🌐 Usar versión online (recomendado)",
        "📤 Subir archivos manualmente",
    ],
)

df_pat = None
df_insp = None

if source.startswith("🌐"):
    with st.spinner("Descargando datos desde Google Sheets..."):
        try:
            df_pat = load_excel_from_gsheets(PATENTES_XLSX_URL)
            df_insp = load_excel_from_gsheets(INSPECCIONES_XLSX_URL)
            st.success("✅ Datos cargados correctamente desde Google Sheets")
        except Exception as e:
            st.error(
                f"No se pudieron cargar los datos online.\n\nDetalle: {e}"
            )
            st.stop()
else:
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        f_pat = st.file_uploader(
            "Sube Base Patentes Aeropuerto",
            type=["xlsx", "xls"]
        )
    with col_u2:
        f_insp = st.file_uploader(
            "Sube Inspecciones Aeropuerto",
            type=["xlsx", "xls"]
        )

    if not f_pat or not f_insp:
        st.info("👆 Sube ambos archivos para continuar.")
        st.stop()

    df_pat = pd.read_excel(f_pat)
    df_insp = pd.read_excel(f_insp)

# =========================================================
# Configuración semáforo
# =========================================================
with st.expander("⚙️ Configuración semáforo", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        green_max = st.number_input(
            "🟢 OK hasta (días)", min_value=0, value=7
        )
    with c2:
        yellow_max = st.number_input(
            "🟡 Alerta hasta (días)", min_value=0, value=30
        )
    if yellow_max < green_max:
        yellow_max = green_max

# =========================================================
# Validaciones mínimas
# =========================================================
if "REG PLATE" not in df_pat.columns:
    st.error("La base de Patentes debe tener la columna REG PLATE")
    st.stop()

required_insp = [
    "Fecha",
    "Patente del Vehículo",
    "Cumplimiento Exterior",
    "Cumplimiento Interior",
    "Cumplimiento Conductor",
]
missing = [c for c in required_insp if c not in df_insp.columns]
if missing:
    st.error(f"En Inspecciones faltan columnas: {missing}")
    st.stop()

# =========================================================
# Procesamiento
# =========================================================
df_pat = df_pat.copy()
df_insp = df_insp.copy()

df_pat["plate_norm"] = df_pat["REG PLATE"].apply(normalize_plate)
df_insp["plate_norm"] = df_insp["Patente del Vehículo"].apply(normalize_plate)
df_insp["Fecha_dt"] = try_parse_date(df_insp["Fecha"])

df_last = (
    df_insp[df_insp["Fecha_dt"].notna()]
    .sort_values(["plate_norm", "Fecha_dt"], ascending=[True, False])
    .groupby("plate_norm", as_index=False)
    .first()
)

df_last = df_last[
    [
        "plate_norm",
        "Fecha_dt",
        "Cumplimiento Exterior",
        "Cumplimiento Interior",
        "Cumplimiento Conductor",
    ]
].rename(columns={"Fecha_dt": "Última Fecha Inspección"})

df = df_pat.merge(df_last, on="plate_norm", how="left")

today = pd.Timestamp.today().normalize()
df["Inspeccionado"] = df["Última Fecha Inspección"].notna()
df["Días desde última inspección"] = (
    today - pd.to_datetime(df["Última Fecha Inspección"])
).dt.days

df["Semáforo"] = df["Días desde última inspección"].apply(
    lambda x: traffic_light(x, green_max, yellow_max)
)

# =========================================================
# KPIs
# =========================================================
k1, k2, k3 = st.columns(3)
k1.metric("Total Patentes", len(df))
k2.metric("Con inspección", int(df["Inspeccionado"].sum()))
k3.metric("Sin inspección", int((~df["Inspeccionado"]).sum()))

st.divider()

# =========================================================
# Tablas TOP
# =========================================================
left, right = st.columns(2)

with left:
    st.subheader("🧾 Top sin inspección")
    st.dataframe(
        df[df["Inspeccionado"] == False]
        .head(20)[["REG PLATE", "Semáforo"]],
        use_container_width=True,
        hide_index=True,
    )

with right:
    st.subheader("🕰️ Top inspecciones más antiguas")
    st.dataframe(
        df[df["Inspeccionado"] == True]
        .sort_values("Días desde última inspección", ascending=False)
        .head(20)[
            [
                "REG PLATE",
                "Última Fecha Inspección",
                "Días desde última inspección",
                "Semáforo",
                "Cumplimiento Exterior",
                "Cumplimiento Interior",
                "Cumplimiento Conductor",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

st.divider()

# =========================================================
# Gráfico distribución
# =========================================================
st.subheader("📊 Distribución de días desde última inspección")

vals = df.loc[df["Inspeccionado"], "Días desde última inspección"].dropna()
if not vals.empty:
    fig = plt.figure()
    plt.hist(vals, bins=20)
    plt.xlabel("Días")
    plt.ylabel("Cantidad de patentes")
    st.pyplot(fig)
else:
    st.info("No hay inspecciones con fecha válida.")

st.divider()

# =========================================================
# Tabla principal + descarga
# =========================================================
st.subheader("📋 Resultado completo")

df_view = df[
    [
        "REG PLATE",
        "Semáforo",
        "Última Fecha Inspección",
        "Días desde última inspección",
        "Cumplimiento Exterior",
        "Cumplimiento Interior",
        "Cumplimiento Conductor",
        "Inspeccionado",
    ]
]

st.dataframe(df_view, use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ Descargar resultado (Excel)",
    data=to_excel_bytes(df_view),
    file_name="ultima_inspeccion_por_patente.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
