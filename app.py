import io
import re
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
# Links de descarga (privado: descarga ocurre con sesión del usuario)
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
    """
    Normaliza encabezados para evitar errores por:
    - espacios dobles / al final
    - saltos de línea
    - NBSP (espacio no separable)
    """
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00a0", " ", regex=False)   # NBSP
        .str.replace(r"\s+", " ", regex=True)     # colapsa espacios/saltos de línea
        .str.strip()                              # quita espacios al inicio/fin
    )
    return df

def normalize_plate(x: str) -> str:
    """Normaliza patente: mayúsculas y solo A-Z0-9."""
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def try_parse_date(series: pd.Series) -> pd.Series:
    """
    Intenta parsear fechas:
    - texto datetime (dayfirst)
    - serial Excel
    """
    dt = pd.to_datetime(series, errors="coerce", dayfirst=True)
    mask = dt.isna()
    if mask.any():
        numeric = pd.to_numeric(series[mask], errors="coerce")
        dt2 = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
        dt.loc[mask] = dt2
    return dt

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
# UI
# =========================================================
st.title("🚗 Última inspección por patente – Aeropuerto")

st.caption(
    "Esta app mantiene los Google Sheets **privados**. "
    "La descarga se hace en tu navegador usando tu **sesión corporativa**, y luego subes los Excel aquí."
)

st.divider()

# =========================================================
# Asistente Paso 1: sesión corporativa + descargas
# =========================================================
st.subheader("Paso 1) Inicia sesión con tu cuenta corporativa y descarga los Excel")

st.info(
    "✅ Abre cada botón y descarga el Excel.\n"
    "Si te aparece **Access denied / Solicitar acceso**, inicia sesión con tu **cuenta corporativa** "
    "y vuelve a intentar.\n\n"
    "Luego vuelve a esta app para subir ambos archivos."
)

c1, c2 = st.columns(2)
with c1:
    st.link_button("📄 Descargar Base Patentes (Excel)", PATENTES_EXPORT_URL)
with c2:
    st.link_button("📝 Descargar Inspecciones (Excel)", INSPECCIONES_EXPORT_URL)

st.write("")
logged = st.checkbox("✅ Ya inicié sesión con mi cuenta corporativa (si era necesario)")

st.divider()

# =========================================================
# Paso 2: Subida de archivos
# =========================================================
st.subheader("Paso 2) Sube los dos archivos descargados")

if not logged:
    st.warning(
        "Marca el checkbox cuando hayas iniciado sesión con tu cuenta corporativa (si lo necesitabas). "
        "Luego sube ambos archivos."
    )

col_u1, col_u2 = st.columns(2)
with col_u1:
    f_pat = st.file_uploader("Sube **Base Patentes Aeropuerto** (Excel)", type=["xlsx", "xls"])
with col_u2:
    f_insp = st.file_uploader("Sube **Inspecciones Aeropuerto** (Excel)", type=["xlsx", "xls"])

if not f_pat or not f_insp:
    st.info("👆 Sube ambos archivos para continuar al Paso 3.")
    st.stop()

# Leer excels
df_pat = pd.read_excel(f_pat)
df_insp = pd.read_excel(f_insp)

# Normalizar encabezados (FIX clave)
df_pat = normalize_headers(df_pat)
df_insp = normalize_headers(df_insp)

# Validaciones (ya robustas gracias a normalize_headers)
if "REG PLATE" not in df_pat.columns:
    st.error(
        "❌ El archivo de Patentes no parece correcto: falta la columna **REG PLATE**.\n"
        "Descárgalo desde el botón corporativo y vuelve a subirlo."
    )
    with st.expander("Ver columnas detectadas (Patentes)"):
        st.write(list(df_pat.columns))
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
    st.error(
        "❌ El archivo de Inspecciones no parece correcto: faltan columnas clave:\n- "
        + "\n- ".join(missing)
        + "\n\nDescárgalo desde el botón corporativo y vuelve a subirlo."
    )
    with st.expander("Ver columnas detectadas (Inspecciones)"):
        st.write(list(df_insp.columns))
    st.stop()

st.success("✅ Archivos cargados correctamente.")
st.divider()

# =========================================================
# Paso 3: Configuración + Resultados
# =========================================================
st.subheader("Paso 3) Resultados")

with st.expander("⚙️ Configuración semáforo (días desde la última inspección)", expanded=True):
    a1, a2 = st.columns(2)
    with a1:
        green_max = st.number_input("🟢 OK hasta (días)", min_value=0, value=7, step=1)
    with a2:
        yellow_max = st.number_input("🟡 Alerta hasta (días)", min_value=0, value=30, step=1)

    if yellow_max < green_max:
        st.warning("El umbral 🟡 (Alerta) no puede ser menor que 🟢 (OK). Se ajustó automáticamente.")
        yellow_max = green_max

# Procesamiento
df_pat = df_pat.copy()
df_insp = df_insp.copy()

df_pat["plate_norm"] = df_pat["REG PLATE"].apply(normalize_plate)
df_insp["plate_norm"] = df_insp["Patente del Vehículo"].apply(normalize_plate)
df_insp["Fecha_dt"] = try_parse_date(df_insp["Fecha"])

# Última inspección por patente
df_last = (
    df_insp[df_insp["plate_norm"] != ""]
    .loc[df_insp["Fecha_dt"].notna()]
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
df["Días desde última inspección"] = (today - pd.to_datetime(df["Última Fecha Inspección"])).dt.days
df["Semáforo"] = df["Días desde última inspección"].apply(lambda x: traffic_light(x, green_max, yellow_max))

# KPIs
total_pat = int(df["REG PLATE"].notna().sum())
inspected = int(df["Inspeccionado"].sum())
never = total_pat - inspected

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Patentes", f"{total_pat:,}".replace(",", "."))
k2.metric("Con inspección", f"{inspected:,}".replace(",", "."))
k3.metric("Sin inspección", f"{never:,}".replace(",", "."))
k4.metric("Fecha de cálculo", today.strftime("%Y-%m-%d"))

st.divider()

# Top tablas
left, right = st.columns(2)

with left:
    st.subheader("🧾 Top sin inspección (primeras 20)")
    base_cols = [c for c in ["REG PLATE", "Flota", "Company", "Marca", "Modelo", "Color"] if c in df.columns]
    cols_never = base_cols + ["Semáforo"]
    st.dataframe(
        df[df["Inspeccionado"] == False][cols_never].head(20),
        use_container_width=True,
        hide_index=True
    )

with right:
    st.subheader("🕰️ Top inspecciones más antiguas (primeras 20)")
    cols_old = [
        "REG PLATE",
        "Última Fecha Inspección",
        "Días desde última inspección",
        "Semáforo",
        "Cumplimiento Exterior",
        "Cumplimiento Interior",
        "Cumplimiento Conductor",
    ]
    cols_old = [c for c in cols_old if c in df.columns]
    st.dataframe(
        df[df["Inspeccionado"] == True]
        .sort_values("Días desde última inspección", ascending=False)[cols_old]
        .head(20),
        use_container_width=True,
        hide_index=True
    )

st.divider()

# Gráfico distribución
st.subheader("📊 Distribución de días desde última inspección (solo inspeccionados)")

vals = df.loc[df["Inspeccionado"] == True, "Días desde última inspección"].dropna()
if vals.empty:
    st.info("No hay inspecciones con fecha válida para graficar.")
else:
    bins = st.slider("Número de bins (barras)", min_value=5, max_value=60, value=20, step=1)
    fig = plt.figure()
    plt.hist(vals, bins=bins)
    plt.xlabel("Días desde última inspección")
    plt.ylabel("Cantidad de patentes")
    st.pyplot(fig)

st.divider()

# Tabla principal filtrable
st.subheader("📋 Resultado completo (filtrable)")

f1, f2, f3, f4 = st.columns([1, 1, 1, 2])
with f1:
    only_never = st.checkbox("Solo SIN inspección", value=False)
with f2:
    only_inspected = st.checkbox("Solo CON inspección", value=False)
with f3:
    sem_filter = st.selectbox("Filtrar por semáforo", ["(Todos)", "⚫ Sin inspección", "🔴 Crítico", "🟡 Alerta", "🟢 OK"])
with f4:
    q = st.text_input("Buscar patente (ej: ABCD12)", value="").strip().upper()

df_show = df.copy()

if only_never and not only_inspected:
    df_show = df_show[df_show["Inspeccionado"] == False]
elif only_inspected and not only_never:
    df_show = df_show[df_show["Inspeccionado"] == True]

if sem_filter != "(Todos)":
    df_show = df_show[df_show["Semáforo"] == sem_filter]

if q:
    qn = normalize_plate(q)
    df_show = df_show[df_show["plate_norm"].str.contains(qn, na=False)]

cols_view = [
    "REG PLATE",
    "Semáforo",
    "Última Fecha Inspección",
    "Días desde última inspección",
    "Cumplimiento Exterior",
    "Cumplimiento Interior",
    "Cumplimiento Conductor",
    "Inspeccionado",
]
optional_base_cols = ["Marca", "Modelo", "Color", "Flota", "Company"]
for c in optional_base_cols:
    if c in df_show.columns and c not in cols_view:
        cols_view.insert(1, c)

cols_view = [c for c in cols_view if c in df_show.columns]

st.dataframe(df_show[cols_view], use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ Descargar resultado filtrado (Excel)",
    data=to_excel_bytes(df_show[cols_view]),
    file_name="ultima_inspeccion_por_patente.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# =========================================================
# Bloque final: solución de problemas
# =========================================================
with st.expander("🛠️ Solución de problemas (si no puedes descargar)"):
    st.markdown(
        "- Si al abrir los enlaces aparece **Solicitar acceso / Access denied**, inicia sesión con tu **cuenta corporativa**.\n"
        "- Si aun así no te deja, pide permisos al dueño del archivo (Drive corporativo).\n"
        "- Si subes un Excel y la app dice que faltan columnas, revisa que hayas descargado el archivo correcto desde los botones."
    )

