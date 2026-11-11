import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# =========================================================
# 1. CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Evolución de precios - Río Loa",
    page_icon="📊",
    layout="wide"
)

st.title("EVOLUCIÓN DE PRECIOS")
st.subheader("Planta Río Loa")

st.markdown("""
<style>
div[data-testid="stMetric"] {
    background: #F4F8FB;
    border: 1px solid #E1EAF0;
    padding: 15px;
    border-radius: 10px;
}
h1, h2, h3 {
    color: #183B56;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# 2. IDENTIFICAR ARCHIVO EXCEL
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

archivos = sorted([
    p for p in BASE_DIR.iterdir()
    if p.suffix.lower() in (".xlsx", ".xls")
    and "precios rio loa" in p.stem.lower()
    and not p.name.startswith("~$")
])

if not archivos:
    st.error(
        "No encontré el Excel 'PRECIOS RIO LOA' "
        "junto a app.py."
    )
    st.stop()

archivo = archivos[0]


# =========================================================
# 3. CARGAR Y LIMPIAR DATOS
# =========================================================

@st.cache_data(show_spinner=False)
def cargar_datos(ruta, mtime):

    df = pd.read_excel(
        ruta,
        sheet_name="Data"
    )

    def normalizar_columna(columna):
        return " ".join(
            str(columna).strip().upper().split()
        )

    df.columns = [
        normalizar_columna(c)
        for c in df.columns
    ]

    equivalencias = {
        "POSICION": "POSICIÓN",
        "FECHA CONTABILIZ.": "FECHA CONTABILIZACIÓN",
        "UN.MEDIDA DE ENTRADA": "UNIDAD",
        "UNIDAD DE MEDIDA": "UNIDAD",
        "DESCRIPCION MATERIAL": "DESCRIPCIÓN MATERIAL"
    }

    df = df.rename(columns=equivalencias)

    columnas_obligatorias = [
        "PEDIDO",
        "MATERIAL",
        "DESCRIPCIÓN MATERIAL",
        "CANTIDAD",
        "PRECIO",
        "MONEDA",
        "FECHA OC"
    ]

    faltantes = [
        c for c in columnas_obligatorias
        if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas: " + ", ".join(faltantes)
        )

    # Limpiar códigos de pedido y material
    for columna in ["PEDIDO", "MATERIAL"]:

        df[columna] = (
            df[columna]
            .astype("string")
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

        df.loc[
            df[columna].isin(["", "nan", "None", "<NA>"]),
            columna
        ] = pd.NA

    df["DESCRIPCIÓN MATERIAL"] = (
        df["DESCRIPCIÓN MATERIAL"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Convertir valores numéricos
    for columna in ["PRECIO", "CANTIDAD"]:

        df[columna] = pd.to_numeric(
            df[columna],
            errors="coerce"
        )

    # Normalizar moneda
    df["MONEDA"] = (
        df["MONEDA"]
        .astype("string")
        .str.upper()
        .str.strip()
    )

    # Convertir fechas
    for columna in [
        "FECHA OC",
        "FECHA CONTABILIZACIÓN"
    ]:

        if columna in df.columns:

            df[columna] = pd.to_datetime(
                df[columna],
                errors="coerce",
                dayfirst=True
            )

    # Completar datos faltantes dentro de la misma
    # combinación de Pedido + Material.
    # No se modifican las cantidades.

    for columna in [
        "PRECIO",
        "MONEDA",
        "FECHA OC"
    ]:

        df[columna] = (
            df.groupby(
                ["PEDIDO", "MATERIAL"],
                dropna=True
            )[columna]
            .transform(
                lambda serie: serie.ffill().bfill()
            )
        )

    # Crear columna de unidad si no existe
    if "UNIDAD" not in df.columns:
        df["UNIDAD"] = ""

    df["UNIDAD"] = (
        df["UNIDAD"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Filtrar registros válidos
    df = df.loc[
        df["PEDIDO"].notna()
        & df["MATERIAL"].notna()
        & df["FECHA OC"].notna()
        & df["PRECIO"].gt(0)
        & df["CANTIDAD"].gt(0)
        & df["MONEDA"].isin(["CLP", "USD"])
    ].copy()

    # Calcular precio unitario
    df["PRECIO UNITARIO"] = (
        df["PRECIO"] / df["CANTIDAD"]
    )

    # Crear etiqueta de materia prima
    df["ETIQUETA"] = (
        df["DESCRIPCIÓN MATERIAL"]
        + " ("
        + df["MATERIAL"].astype(str)
        + ")"
    )

    return df


try:

    datos = cargar_datos(
        str(archivo),
        archivo.stat().st_mtime_ns
    )

except Exception as exc:

    st.error(
        f"No fue posible cargar los datos: {exc}"
    )
    st.stop()


if datos.empty:
    st.warning("No hay registros válidos en el archivo.")
    st.stop()


# =========================================================
# 4. FUNCIONES AUXILIARES
# =========================================================

def precio_fmt(valor, moneda):

    if pd.isna(valor):
        return "-"

    if moneda == "CLP":

        return (
            "$"
            + f"{valor:,.0f}".replace(",", ".")
        )

    return f"US$ {valor:,.2f}"


def ranking_variaciones(df):

    resultados = []

    for material, grupo in df.groupby("MATERIAL"):

        grupo = grupo.sort_values(
            ["FECHA OC", "PEDIDO"],
            kind="stable"
        )

        if len(grupo) < 2:
            continue

        inicial = grupo["PRECIO UNITARIO"].iloc[0]
        final = grupo["PRECIO UNITARIO"].iloc[-1]

        if pd.notna(inicial) and inicial > 0:

            variacion = (
                (final / inicial) - 1
            ) * 100

            resultados.append({
                "MATERIAL": str(material),
                "ETIQUETA": grupo["ETIQUETA"].iloc[0],
                "VARIACIÓN": variacion
            })

    if not resultados:

        return pd.DataFrame(
            columns=[
                "MATERIAL",
                "ETIQUETA",
                "VARIACIÓN"
            ]
        )

    ranking = pd.DataFrame(resultados)

    ranking["ABS"] = ranking["VARIACIÓN"].abs()

    return (
        ranking
        .sort_values("ABS", ascending=False)
        .head(4)
    )


# =========================================================
# 5. FILTROS DE FECHA
# =========================================================

st.sidebar.header("Filtros")

min_fecha = datos["FECHA OC"].min().date()
max_fecha = datos["FECHA OC"].max().date()

fechas = st.sidebar.date_input(
    "Período de análisis",
    value=(min_fecha, max_fecha),
    min_value=min_fecha,
    max_value=max_fecha
)

if not isinstance(fechas, (tuple, list)) or len(fechas) != 2:

    st.info("Selecciona ambas fechas para continuar.")
    st.stop()

inicio = pd.Timestamp(fechas[0])

fin = (
    pd.Timestamp(fechas[1])
    + pd.Timedelta(days=1)
)

filtrado = datos.loc[
    (datos["FECHA OC"] >= inicio)
    & (datos["FECHA OC"] < fin)
].copy()

if filtrado.empty:

    st.warning(
        "No hay registros en el período seleccionado."
    )
    st.stop()


# =========================================================
# 6. SELECTOR DE MONEDA
# =========================================================

st.header("Moneda de análisis")

monedas = [
    moneda
    for moneda in ["CLP", "USD"]
    if moneda in filtrado["MONEDA"].unique()
]

if not monedas:
    st.warning("No hay datos disponibles en CLP o USD.")
    st.stop()

if (
    "moneda_rio" not in st.session_state
    or st.session_state["moneda_rio"] not in monedas
):

    st.session_state["moneda_rio"] = monedas[0]

moneda = st.selectbox(
    "Selecciona la moneda:",
    monedas,
    key="moneda_rio"
)

df = filtrado.loc[
    filtrado["MONEDA"] == moneda
].copy()


# =========================================================
# 7. ANÁLISIS GENERAL
# =========================================================

st.divider()

st.header("1. Análisis general")

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "Órdenes de compra",
    f"{df['PEDIDO'].nunique():,}".replace(",", ".")
)

c2.metric(
    "Materias primas",
    f"{df['MATERIAL'].nunique():,}".replace(",", ".")
)

c3.metric(
    "Precio unitario promedio",
    precio_fmt(
        df["PRECIO UNITARIO"].mean(),
        moneda
    )
)

c4.metric(
    "Precio unitario máximo",
    precio_fmt(
        df["PRECIO UNITARIO"].max(),
        moneda
    )
)


# =========================================================
# 8. ÚNICO GRÁFICO GENERAL: EVOLUCIÓN DE PRECIOS
# =========================================================

st.subheader(
    "Evolución de precios de las principales materias primas"
)

top = ranking_variaciones(df)

if top.empty:

    st.info(
        "No hay suficientes registros para comparar "
        "precios entre órdenes de compra."
    )

else:

    principales = df.loc[
        df["MATERIAL"].astype(str).isin(top["MATERIAL"])
    ].sort_values(
        ["FECHA OC", "PEDIDO"],
        kind="stable"
    ).copy()

    fig = px.line(
        principales,
        x="FECHA OC",
        y="PRECIO UNITARIO",
        color="ETIQUETA",
        markers=True,
        labels={
            "FECHA OC": "Fecha de creación de OC",
            "PRECIO UNITARIO": f"Precio unitario ({moneda})",
            "ETIQUETA": "Materia prima"
        },
        hover_data={
            "PEDIDO": True,
            "CANTIDAD": True,
            "UNIDAD": True
        }
    )

    fig.update_layout(
        height=480,
        hovermode="closest",
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(
            showgrid=True,
            gridcolor="#E9EFF3"
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor="#E9EFF3"
        )
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    st.caption(
        "Se muestran hasta cuatro materias primas con mayor "
        "variación absoluta entre su primer y último precio "
        "unitario del período. No se convierten monedas."
    )


# =========================================================
# 9. ANÁLISIS INDIVIDUAL
# =========================================================

st.divider()

st.header("2. Análisis individual de materias primas")

materiales = (
    df[
        ["MATERIAL", "DESCRIPCIÓN MATERIAL"]
    ]
    .drop_duplicates(subset=["MATERIAL"])
    .sort_values(
        ["DESCRIPCIÓN MATERIAL", "MATERIAL"]
    )
    .copy()
)

# Buscador
busqueda = st.text_input(
    "Buscar materia prima",
    placeholder="Buscar",
    key="buscador_material_rio"
).strip()


# =========================================================
# 10. FILTRAR MATERIAS PRIMAS
# =========================================================

coincidencias = materiales.copy()

if busqueda:

    mascara = (
        materiales["MATERIAL"]
        .astype(str)
        .str.contains(
            busqueda,
            case=False,
            regex=False,
            na=False
        )
        |
        materiales["DESCRIPCIÓN MATERIAL"]
        .str.contains(
            busqueda,
            case=False,
            regex=False,
            na=False
        )
    )

    coincidencias = materiales.loc[mascara].copy()


if coincidencias.empty:

    st.warning(
        "No se encontraron materias primas "
        "que coincidan con la búsqueda."
    )
    st.stop()


# =========================================================
# 11. SELECTOR DESPLEGABLE DE MATERIAS PRIMAS
# =========================================================

opciones = [
    f"{fila['DESCRIPCIÓN MATERIAL']} ({fila['MATERIAL']})"
    for _, fila in coincidencias.iterrows()
]

mapa = {
    f"{fila['DESCRIPCIÓN MATERIAL']} ({fila['MATERIAL']})":
    str(fila["MATERIAL"])
    for _, fila in coincidencias.iterrows()
}

# Reiniciar selección al cambiar la búsqueda
if (
    st.session_state.get("_busqueda_anterior_rio")
    != busqueda
):

    st.session_state["_busqueda_anterior_rio"] = busqueda

    st.session_state.pop(
        "selector_material_rio",
        None
    )

if (
    st.session_state.get("selector_material_rio")
    not in opciones
):

    st.session_state["selector_material_rio"] = opciones[0]


seleccion = st.selectbox(
    "Selecciona una materia prima:",
    opciones,
    key="selector_material_rio"
)

material = mapa[seleccion]

individual = df.loc[
    df["MATERIAL"].astype(str) == material
].sort_values(
    ["FECHA OC", "PEDIDO"],
    kind="stable"
).copy()


# =========================================================
# 12. INFORMACIÓN DEL MATERIAL SELECCIONADO
# =========================================================

st.subheader(
    individual["DESCRIPCIÓN MATERIAL"].iloc[0]
)

st.caption(
    f"Código SAP: {material}"
)

primero = individual["PRECIO UNITARIO"].iloc[0]

ultimo = individual["PRECIO UNITARIO"].iloc[-1]

variacion = (
    (ultimo / primero - 1) * 100
    if primero > 0 and len(individual) > 1
    else None
)


# =========================================================
# 13. INDICADORES INDIVIDUALES
# =========================================================

m1, m2, m3, m4 = st.columns(4)

m1.metric(
    "Precio inicial",
    precio_fmt(primero, moneda)
)

m2.metric(
    "Último precio",
    precio_fmt(ultimo, moneda)
)

m3.metric(
    "Variación del período",
    f"{variacion:+.2f}%"
    if variacion is not None
    else "-"
)

m4.metric(
    "Órdenes de compra",
    individual["PEDIDO"].nunique()
)


# =========================================================
# 14. GRÁFICO INDIVIDUAL
# =========================================================

st.markdown(
    "### Evolución histórica del precio unitario"
)

fig_ind = go.Figure()

fig_ind.add_trace(
    go.Scatter(
        x=individual["FECHA OC"],
        y=individual["PRECIO UNITARIO"],
        mode="lines+markers",
        name="Precio unitario",
        line=dict(
            width=3,
            color="#2878A8"
        ),
        marker=dict(size=8),
        customdata=individual[
            ["PEDIDO", "CANTIDAD", "UNIDAD"]
        ].fillna("").astype(str).to_numpy(),
        hovertemplate=(
            "<b>Fecha:</b> %{x|%d-%m-%Y}<br>"
            "<b>Precio unitario:</b> %{y:,.4f}<br>"
            "<b>OC:</b> %{customdata[0]}<br>"
            "<b>Cantidad:</b> %{customdata[1]} "
            "%{customdata[2]}<extra></extra>"
        )
    )
)

fig_ind.update_layout(
    height=450,
    xaxis_title="Fecha de creación de OC",
    yaxis_title=f"Precio unitario ({moneda})",
    hovermode="closest",
    plot_bgcolor="white",
    paper_bgcolor="white",
    xaxis=dict(
        showgrid=True,
        gridcolor="#E9EFF3"
    ),
    yaxis=dict(
        showgrid=True,
        gridcolor="#E9EFF3"
    )
)

st.plotly_chart(
    fig_ind,
    use_container_width=True
)


# =========================================================
# 15. HISTORIAL DE ÓRDENES DE COMPRA
# =========================================================

st.markdown(
    "### Historial de órdenes de compra"
)

columnas = [
    "PEDIDO",
    "POSICIÓN",
    "MATERIAL",
    "DESCRIPCIÓN MATERIAL",
    "CANTIDAD",
    "UNIDAD",
    "PRECIO",
    "MONEDA",
    "PRECIO UNITARIO",
    "FECHA OC",
    "FECHA CONTABILIZACIÓN"
]

columnas_disponibles = [
    columna
    for columna in columnas
    if columna in individual.columns
]

historial = individual[
    columnas_disponibles
].sort_values(
    "FECHA OC",
    ascending=False
).copy()

st.dataframe(
    historial,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# 16. DESCARGAR HISTORIAL
# =========================================================

csv = historial.to_csv(
    index=False,
    sep=";"
).encode("utf-8-sig")

st.download_button(
    "Descargar historial de órdenes de compra",
    data=csv,
    file_name=f"historial_{material}.csv",
    mime="text/csv"
)


# =========================================================
# 17. FUENTE DE INFORMACIÓN
# =========================================================

st.divider()

st.caption(
    "Fuente: registros SAP de órdenes de compra de Río Loa. "
    "Precio unitario = PRECIO / CANTIDAD. "
    "Los resultados dependen de la estructura y calidad del Excel."
)
