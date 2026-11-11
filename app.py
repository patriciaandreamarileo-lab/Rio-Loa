
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# =====================================================
# 1. CONFIGURACIÓN GENERAL
# =====================================================

st.set_page_config(
    page_title="Evolución de Precios | Río Loa",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Colores
AZUL = "#17365D"
CELESTE = "#3B82A0"
GRIS = "#64748B"
VERDE = "#25855A"
ROJO = "#C64B4B"

st.markdown("""
<style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    h1 {
        color: #17365D;
        font-weight: 800;
    }

    h2, h3 {
        color: #17365D;
    }

    [data-testid="stMetric"] {
        background-color: #F4F7FA;
        border: 1px solid #E1E7ED;
        padding: 15px;
        border-radius: 10px;
    }

    div[data-testid="stMetricValue"] {
        color: #17365D;
    }

    .subtitle {
        color: #64748B;
        font-size: 14px;
    }

    .section-title {
        color: #17365D;
        font-size: 19px;
        font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)


# =====================================================
# 2. CARGA Y PREPARACIÓN DE DATOS
# =====================================================

ARCHIVO = Path(__file__).parent / "precios rio loa.xlsx"


@st.cache_data
def cargar_datos():

    if not ARCHIVO.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo: {ARCHIVO.name}"
        )

    df = pd.read_excel(
        ARCHIVO,
        sheet_name="Data"
    )

    # Limpiar nombres de columnas
    df.columns = (
        df.columns.astype(str)
        .str.strip()
    )

    # Normalizar nombres alternativos
    df = df.rename(columns={
        "POSICION": "POSICIÓN",
        "Fecha OC": "FECHA OC",
        "Fecha contabiliz.": "FECHA CONTABILIZACIÓN",
        "Un.medida de entrada": "UNIDAD"
    })

    # Columnas obligatorias
    requeridas = [
        "Material",
        "Descripción material",
        "Pedido",
        "FECHA OC",
        "PRECIO",
        "CANTIDAD",
        "MONEDA"
    ]

    faltantes = [
        c for c in requeridas if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en el Excel: "
            + ", ".join(faltantes)
        )

    # Convertir fechas
    df["FECHA OC"] = pd.to_datetime(
        df["FECHA OC"],
        errors="coerce",
        dayfirst=True
    )

    # Convertir valores numéricos
    def convertir_numero(columna):

        if pd.api.types.is_numeric_dtype(columna):
            return pd.to_numeric(
                columna, errors="coerce"
            )

        serie = columna.astype(str).str.strip()

        serie = serie.str.replace(
            r"[$\s]", "", regex=True
        )

        # Formato chileno: 1.234,56
        serie = serie.str.replace(
            ".", "", regex=False
        ).str.replace(
            ",", ".", regex=False
        )

        return pd.to_numeric(
            serie, errors="coerce"
        )

    df["PRECIO"] = convertir_numero(df["PRECIO"])
    df["CANTIDAD"] = convertir_numero(df["CANTIDAD"])

    # Limpiar moneda y material
    df["MONEDA"] = (
        df["MONEDA"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["Material"] = (
        df["Material"]
        .astype(str)
        .str.strip()
    )

    df["Descripción material"] = (
        df["Descripción material"]
        .fillna("Sin descripción")
        .astype(str)
        .str.strip()
    )

    # Conservar cada posición como registro independiente
    if "POSICIÓN" not in df.columns:
        df["POSICIÓN"] = "Sin posición"

    df["POSICIÓN"] = (
        df["POSICIÓN"]
        .fillna("Sin posición")
        .astype(str)
    )

    # Calcular precio unitario
    df["PRECIO UNITARIO"] = np.where(
        df["CANTIDAD"] > 0,
        df["PRECIO"] / df["CANTIDAD"],
        np.nan
    )

    # Eliminar registros sin información válida
    df = df.dropna(
        subset=[
            "FECHA OC",
            "PRECIO UNITARIO",
            "PRECIO",
            "CANTIDAD"
        ]
    )

    df = df[
        (df["CANTIDAD"] > 0) &
        (df["PRECIO UNITARIO"] > 0) &
        (df["MONEDA"].isin(["CLP", "USD"]))
    ].copy()

    # Orden cronológico
    df = df.sort_values("FECHA OC")

    # Meses en español
    meses = {
        1: "Enero", 2: "Febrero", 3: "Marzo",
        4: "Abril", 5: "Mayo", 6: "Junio",
        7: "Julio", 8: "Agosto", 9: "Septiembre",
        10: "Octubre", 11: "Noviembre",
        12: "Diciembre"
    }

    df["MES"] = df["FECHA OC"].dt.month.map(meses)

    df["MES AÑO"] = (
        df["MES"] + " " +
        df["FECHA OC"].dt.year.astype(str)
    )

    return df


# =====================================================
# 3. FUNCIONES AUXILIARES
# =====================================================

def formato_precio(valor, moneda):

    if pd.isna(valor):
        return "-"

    if moneda == "CLP":
        return f"$ {valor:,.0f}".replace(",", ".")

    return f"US$ {valor:,.2f}"


def variacion_porcentual(inicial, final):

    if pd.isna(inicial) or inicial == 0:
        return np.nan

    return ((final - inicial) / inicial) * 100


# =====================================================
# 4. CARGAR INFORMACIÓN
# =====================================================

try:
    df = cargar_datos()

except Exception as e:
    st.error(f"Error al cargar los datos: {e}")
    st.stop()


# =====================================================
# 5. ENCABEZADO
# =====================================================

st.markdown(
    "<h1>EVOLUCIÓN DE PRECIOS</h1>",
    unsafe_allow_html=True
)

st.markdown(
    '<p class="subtitle">RÍO LOA | CALAMA | '
    'Análisis de precios de materias primas</p>',
    unsafe_allow_html=True
)

st.divider()


# =====================================================
# 6. FILTROS
# =====================================================

st.sidebar.title("Filtros de análisis")

st.sidebar.markdown("### Período de análisis")

fecha_min = df["FECHA OC"].min().date()
fecha_max = df["FECHA OC"].max().date()

fechas = st.sidebar.date_input(
    "Seleccionar fechas",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max
)

if not isinstance(fechas, (tuple, list)) or len(fechas) != 2:
    st.info("Selecciona una fecha inicial y una fecha final.")
    st.stop()

fecha_inicio, fecha_fin = fechas

if fecha_inicio > fecha_fin:
    st.error("La fecha inicial no puede ser posterior a la final.")
    st.stop()

st.sidebar.markdown("### Moneda")

monedas_disponibles = sorted(df["MONEDA"].unique())

monedas = st.sidebar.multiselect(
    "Seleccionar moneda",
    options=monedas_disponibles,
    default=monedas_disponibles
)

if not monedas:
    st.warning("Selecciona al menos una moneda.")
    st.stop()

# Aplicar filtros
filtrado = df[
    (df["FECHA OC"].dt.date >= fecha_inicio) &
    (df["FECHA OC"].dt.date <= fecha_fin) &
    (df["MONEDA"].isin(monedas))
].copy()

if filtrado.empty:
    st.warning("No existen registros para los filtros seleccionados.")
    st.stop()


# =====================================================
# 7. INDICADORES GENERALES
# =====================================================

st.markdown("### Resumen general")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "Registros de compra",
    f"{len(filtrado):,}".replace(",", ".")
)

col2.metric(
    "Materias primas",
    f"{filtrado['Material'].nunique():,}".replace(",", ".")
)

col3.metric(
    "Órdenes de compra",
    f"{filtrado['Pedido'].nunique():,}".replace(",", ".")
)

col4.metric(
    "Período",
    f"{(fecha_fin - fecha_inicio).days + 1} días"
)

st.divider()


# =====================================================
# 8. EVOLUCIÓN GENERAL POR MONEDA
# =====================================================

st.markdown("### Evolución general de precios")

st.caption(
    "Los precios unitarios se presentan separados por moneda. "
    "No se realiza conversión entre CLP y USD."
)

for moneda in monedas:

    datos_moneda = filtrado[
        filtrado["MONEDA"] == moneda
    ].copy()

    if datos_moneda.empty:
        continue

    st.markdown(f"#### Precios expresados en {moneda}")

    evolucion = (
        datos_moneda
        .sort_values("FECHA OC")
        .groupby(
            ["FECHA OC", "Material", "Descripción material"],
            as_index=False
        )
        .agg({
            "PRECIO UNITARIO": "mean"
        })
    )

    fig = px.scatter(
        evolucion,
        x="FECHA OC",
        y="PRECIO UNITARIO",
        color="Material",
        hover_data=[
            "Descripción material",
            "Material"
        ],
        title=f"Evolución de precios unitarios - {moneda}",
        labels={
            "FECHA OC": "Fecha de creación de OC",
            "PRECIO UNITARIO": f"Precio unitario ({moneda})",
            "Material": "Código de material"
        }
    )

    fig.update_traces(
        marker=dict(size=8)
    )

    fig.update_layout(
        template="plotly_white",
        height=480,
        hovermode="closest",
        legend_title="Material",
        xaxis=dict(
            tickformat="%b %Y",
            title="Fecha de OC"
        ),
        yaxis=dict(
            title=f"Precio unitario ({moneda})"
        )
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

st.divider()


# =====================================================
# 9. RANKING DE MATERIAS PRIMAS
# =====================================================

st.markdown("### Materias primas con mayores variaciones")

st.caption(
    "Ranking calculado comparando el primer y último precio "
    "unitario registrado de cada material dentro del período, "
    "por moneda."
)

ranking_resultados = []

for moneda in monedas:

    datos = filtrado[
        filtrado["MONEDA"] == moneda
    ]

    for material, grupo in datos.groupby("Material"):

        grupo = grupo.sort_values("FECHA OC")

        if len(grupo) < 2:
            continue

        primero = grupo.iloc[0]
        ultimo = grupo.iloc[-1]

        variacion = variacion_porcentual(
            primero["PRECIO UNITARIO"],
            ultimo["PRECIO UNITARIO"]
        )

        ranking_resultados.append({
            "Material": material,
            "Descripción": primero["Descripción material"],
            "Moneda": moneda,
            "Precio inicial": primero["PRECIO UNITARIO"],
            "Precio final": ultimo["PRECIO UNITARIO"],
            "Variación (%)": variacion,
            "Cambios registrados": len(grupo) - 1
        })

ranking = pd.DataFrame(ranking_resultados)

if not ranking.empty:

    ranking = ranking.sort_values(
        "Variación (%)",
        key=lambda x: x.abs(),
        ascending=False
    )

    st.dataframe(
        ranking,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Precio inicial": st.column_config.NumberColumn(
                format="%.4f"
            ),
            "Precio final": st.column_config.NumberColumn(
                format="%.4f"
            ),
            "Variación (%)": st.column_config.NumberColumn(
                format="%.2f%%"
            )
        }
    )

else:
    st.info(
        "No existen suficientes registros para calcular variaciones."
    )

st.divider()


# =====================================================
# 10. CATÁLOGO Y DETALLE DE MATERIAS PRIMAS
# =====================================================

st.markdown("## Análisis por materia prima")

st.caption(
    "Selecciona una materia prima para consultar "
    "su evolución individual."
)

materiales = (
    filtrado[
        ["Material", "Descripción material"]
    ]
    .drop_duplicates()
    .sort_values("Descripción material")
)

opciones = {
    f"{fila['Material']} - {fila['Descripción material']}":
    fila["Material"]
    for _, fila in materiales.iterrows()
}

seleccion = st.selectbox(
    "Seleccionar materia prima",
    options=list(opciones.keys())
)

material_seleccionado = opciones[seleccion]

detalle = filtrado[
    filtrado["Material"] == material_seleccionado
].copy()

detalle = detalle.sort_values("FECHA OC")


# =====================================================
# 11. INFORMACIÓN DEL MATERIAL
# =====================================================

st.markdown(f"### {seleccion}")

col1, col2, col3 = st.columns(3)

col1.metric(
    "Órdenes / posiciones",
    f"{len(detalle):,}".replace(",", ".")
)

col2.metric(
    "Cantidad total",
    f"{detalle['CANTIDAD'].sum():,.2f}"
)

col3.metric(
    "Monedas utilizadas",
    ", ".join(sorted(detalle["MONEDA"].unique()))
)


# =====================================================
# 12. GRÁFICO DE EVOLUCIÓN INDIVIDUAL
# =====================================================

st.markdown("### Evolución histórica del precio unitario")

for moneda in sorted(detalle["MONEDA"].unique()):

    datos = detalle[
        detalle["MONEDA"] == moneda
    ].sort_values("FECHA OC").copy()

    st.markdown(f"#### Evolución en {moneda}")

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=datos["FECHA OC"],
            y=datos["PRECIO UNITARIO"],
            mode="lines+markers",
            name="Precio unitario",
            line=dict(
                color=CELESTE,
                width=2
            ),
            marker=dict(size=9),
            customdata=np.column_stack([
                datos["Pedido"].astype(str),
                datos["POSICIÓN"].astype(str),
                datos["CANTIDAD"].astype(str)
            ]),
            hovertemplate=(
                "Fecha OC: %{x|%d/%m/%Y}<br>"
                "Precio unitario: %{y:,.4f}<br>"
                "Pedido: %{customdata[0]}<br>"
                "Posición: %{customdata[1]}<br>"
                "Cantidad: %{customdata[2]}"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        template="plotly_white",
        height=430,
        xaxis_title="Fecha de creación de OC",
        yaxis_title=f"Precio unitario ({moneda})",
        hovermode="closest"
    )

    fig.update_xaxes(
        tickformat="%b %Y"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    if len(datos) >= 2:

        inicial = datos.iloc[0]["PRECIO UNITARIO"]
        final = datos.iloc[-1]["PRECIO UNITARIO"]

        variacion = variacion_porcentual(
            inicial,
            final
        )

        a, b, c = st.columns(3)

        a.metric(
            "Primer precio registrado",
            formato_precio(inicial, moneda)
        )

        b.metric(
            "Último precio registrado",
            formato_precio(final, moneda)
        )

        c.metric(
            "Variación acumulada",
            f"{variacion:+.2f}%",
            delta=f"{variacion:+.2f}%"
        )


# =====================================================
# 13. TABLA DE ÓRDENES DE COMPRA
# =====================================================

st.markdown("### Historial de órdenes de compra")

columnas_mostrar = [
    "FECHA OC",
    "Pedido",
    "POSICIÓN",
    "Material",
    "Descripción material",
    "CANTIDAD",
    "PRECIO",
    "PRECIO UNITARIO",
    "MONEDA"
]

if "UNIDAD" in detalle.columns:
    columnas_mostrar.insert(6, "UNIDAD")

if "FECHA CONTABILIZACIÓN" in detalle.columns:
    columnas_mostrar.append("FECHA CONTABILIZACIÓN")

tabla = detalle[
    columnas_mostrar
].copy()

tabla = tabla.sort_values(
    "FECHA OC",
    ascending=False
)

st.dataframe(
    tabla,
    use_container_width=True,
    hide_index=True
)


# =====================================================
# 14. DESCARGA DE INFORMACIÓN
# =====================================================

st.markdown("### Exportar información")

csv = tabla.to_csv(
    index=False,
    sep=";",
    decimal=",",
    encoding="utf-8-sig"
)

st.download_button(
    label="Descargar historial en CSV",
    data=csv,
    file_name="evolucion_precios_rio_loa.csv",
    mime="text/csv"
)


# =====================================================
# 15. PIE DE PÁGINA
# =====================================================

st.divider()

st.caption(
    "Dashboard de análisis de precios | Río Loa | "
    "Fuente: registros de órdenes de compra."
)

st.caption(
    "El precio unitario corresponde a PRECIO / CANTIDAD. "
    "La variación porcentual compara el primer y último "
    "registro disponible de cada material dentro del período."
)

