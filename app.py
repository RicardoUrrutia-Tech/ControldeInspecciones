import io
import re
from pathlib import Path
from datetime import date

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt


# =========================================================
# Configuración general (solo 1 vez y al inicio)
# =========================================================
st.set_page_config(
    page_title="Reporte Inspecciones Vehiculares – Aeropuerto CL",
    layout="wide"
)

# =========================================================
# SSO / Acceso restringido (Google OAuth vía Streamlit)
# =========================================================
CORP_DOMAIN = "@cabify.com"

def require_login_and_domain():
    if "auth" not in st.secrets:
        st.error("No se encontró configuración [auth] en Secrets de Streamlit Cloud.")
        st.stop()

    if not getattr(st.user, "is_logged_in", False):
        st.title("🔐 Acceso restringido")
        st.write("Debes iniciar sesión con tu cuenta corporativa para usar esta aplicación.")
        st.button("Iniciar sesión con Google", on_click=st.login)
        st.stop()

    email = (getattr(st.user, "email", "") or "").strip().lower()
    if not email.endswith(CORP_DOMAIN):
        st.title("🔐 Acceso restringido")
        st.error(f"Debes ingresar con una cuenta corporativa ({CORP_DOMAIN}).")
        st.write(f"Sesión detectada: {email or '(sin email)'}")
        st.button("Cerrar sesión", on_click=st.logout)
        st.stop()

require_login_and_domain()

with st.sidebar:
    st.write("### Sesión")
    st.write(f"Conectado como: **{st.user.email}**")
    st.button("Cerrar sesión", on_click=st.logout)

# =========================================================
# Estilo Cabify (CSS)
# =========================================================
st.markdown("""
<style>
.stApp { background-color: #FAF8FE; }

h1, h2, h3 { color: #1F123F; }
h4, h5, h6 { color: #4A2B8D; }

section[data-testid="stSidebar"] { background-color: #1F123F; }
section[data-testid="stSidebar"] * { color: #FAF8FE !important; }

.stButton>button, .stDownloadButton>button {
    background-color: #7145D6;
    color: #FFFFFF;
    border-radius: 10px;
    border: none;
    padding: 0.55rem 0.9rem;
    font-weight: 700;
}
.stButton>button:hover, .stDownloadButton>button:hover {
    background-color: #5B34AC;
}

a[data-testid="stLinkButton"] {
    background-color: #8A6EE4 !important;
    color: #FFFFFF !important;
    border-radius: 10px !important;
    padding: 0.55rem 0.9rem !important;
    font-weight: 800 !important;
    border: none !important;
}

div[data-testid="stAlert"] { border-radius: 10px; }
hr { border-top: 1px solid #DFDAF8; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# Portada + Fecha de corte
# =========================================================
PORTADA_PATH = Path("assets/portada_aeropuerto.png")

if PORTADA_PATH.exists():
    st.image(str(PORTADA_PATH), use_container_width=True)

st.caption(f"📅 **Fecha de corte:** {date.today().strftime('%Y-%m-%d')}")
st.divider()

# =========================================================
# Links de descarga (Drive)
# =========================================================
PATENTES_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1DSZA9DJkxWHBIOMyulW3638TF46Udq7tCWh489UAOAs/"
    "export?format=xlsx&gid=1266707607"
)

INSPECCIONES_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1UR2GTh6l4nwmx4DdY6yX_BS7EF1l-LYHtzHQdkYx03c/"
    "export?format=xlsx"
)

# =========================================================
# Utilidades
# =========================================================
def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df

def normalize_plate(x: str) -> str:
    if pd.isna(x):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(x).upper().strip())

def try_parse_date(series: pd.Series) -> pd.Series:
    dt = pd.to_datetime(series, errors="coerce", dayfirst=True)
    mask = dt.isna()
    if mask.any():
        num = pd.to_numeric(series[mask], errors="coerce")
        dt.loc[mask] = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
    return dt

def compliance_label(x):
    if pd.isna(x):
        return pd.NA
    v = pd.to_numeric(str(x).replace("%", ""), errors="coerce")
    if pd.isna(v):
        return pd.NA
    return "Cumple" if v >= 100 else "No Cumple"

def traffic_light(days, g, y):
    if pd.isna(days):
        return "⚫ Sin inspección"
    if days <= g:
        return "🟢 OK"
    if days <= y:
        return "🟡 Alerta"
    return "🔴 Crítico"

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="resultado")
        ws = writer.sheets["resultado"]

        from openpyxl.styles import PatternFill, Font

        header_fill = PatternFill("solid", fgColor="7145D6")
        header_font = Font(color="FFFFFF", bold=True)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        ws.freeze_panes = "A2"
    return output.getvalue()

# =========================================================
# UI – Flujo principal (SIN CAMBIOS)
# =========================================================
st.subheader("Paso 1) Descarga los Excel")

c1, c2 = st.columns(2)
with c1:
    st.link_button("📄 Descargar Base Patentes (Excel)", PATENTES_EXPORT_URL)
with c2:
    st.link_button("📝 Descargar Inspecciones (Excel)", INSPECCIONES_EXPORT_URL)

st.divider()

st.subheader("Paso 2) Sube los archivos")

u1, u2 = st.columns(2)
with u1:
    f_pat = st.file_uploader("Base Patentes", type=["xlsx", "xls"])
with u2:
    f_insp = st.file_uploader("Inspecciones", type=["xlsx", "xls"])

if not f_pat or not f_insp:
    st.stop()

df_pat = normalize_headers(pd.read_excel(f_pat))
df_insp = normalize_headers(pd.read_excel(f_insp))

df_pat["plate_norm"] = df_pat["REG PLATE"].apply(normalize_plate)
df_insp["plate_norm"] = df_insp["Patente del Vehículo"].apply(normalize_plate)
df_insp["Fecha_dt"] = try_parse_date(df_insp["Fecha"])

df_last = (
    df_insp.sort_values("Fecha_dt", ascending=False)
    .groupby("plate_norm", as_index=False)
    .first()
)

for c in ["Cumplimiento Exterior", "Cumplimiento Interior", "Cumplimiento Conductor"]:
    df_last[c] = df_last[c].apply(compliance_label)

df = df_pat.merge(
    df_last[
        ["plate_norm", "Fecha_dt",
         "Cumplimiento Exterior", "Cumplimiento Interior", "Cumplimiento Conductor"]
    ].rename(columns={"Fecha_dt": "Última Fecha Inspección"}),
    on="plate_norm",
    how="left"
)

today = pd.Timestamp.today().normalize()
df["Días desde última inspección"] = (today - df["Última Fecha Inspección"]).dt.days
df["Semáforo"] = df["Días desde última inspección"].apply(lambda x: traffic_light(x, 7, 30))
df["Inspeccionado"] = df["Última Fecha Inspección"].notna()

st.divider()

st.subheader("📊 Distribución de días desde última inspección")

vals = df.loc[df["Inspeccionado"], "Días desde última inspección"].dropna()
fig = plt.figure(figsize=(6, 3))
plt.hist(vals, bins=20, color="#4A2B8D")
plt.xlabel("Días")
plt.ylabel("Cantidad")
st.pyplot(fig, use_container_width=True)

st.divider()

cols_view = [
    "REG PLATE",
    "Semáforo",
    "Última Fecha Inspección",
    "Días desde última inspección",
    "Cumplimiento Exterior",
    "Cumplimiento Interior",
    "Cumplimiento Conductor",
]

st.dataframe(df[cols_view], use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ Descargar resultado (Excel)",
    data=to_excel_bytes(df[cols_view]),
    file_name="ultima_inspeccion_por_patente.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

