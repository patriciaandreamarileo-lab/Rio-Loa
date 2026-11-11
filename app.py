
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path


# ==========================================================
# 1. CONFIGURACIÓN GENERAL
# ==========================================================

st.set_page_config(
    page_title="Evolución de Precios | Río Loa",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

AZUL = "#17365D"
CELESTE = "#3B82A0"


# ==========================================================
# 2. DISEÑO VISUAL
# ==========================================================

st.markdown("""
<style>

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}

h1, h2, h3 {
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

/* Fondo celeste para análisis individual */

.st-key-analisis_material {
    background-color: #F0F8FC;
    border: 1px solid #D8EAF3;
    border-radius: 14px;
    padding: 24px;
    margin-top: 20px;
    margin-bottom: 20px;
}

.st-key-analisis_material h2,
.st-key-analisis_material h3 {
    color: #17365D;
}

</style>
""", unsafe_allow_html=True)


# ==========================================================
# 3. UBICACIÓN DEL ARCHIVO
# ==========================================================

ARCHIVO = Path(__file__).parent / "precios rio loa.xlsx"


# ==========================================================
# 4. FUNCIONES DE LIMPIEZA
# ==========================================================

def convertir_numero(columna):

    if pd.api.types.is_numeric_dtype(columna):
        return pd.to_numeric(columna, errors="coerce")

    serie = columna.astype(str).str.strip()

    serie = serie.str.replace(
        r"[$\s]", "", regex=True
    )

    serie = serie.str.replace(
        ".", "", regex=False
    )

    serie = serie.str.replace(
        ",", ".", regex=False
    )

    return pd.to_numeric(
        serie,
        errors="coerce"
    )


def convertir_fecha(columna):

    return pd.to_datetime(
        columna,
        errors="coerce",
        dayfirst=True
    )


# ==========================================================
# 5. CARGA Y PREPARACIÓN DE DATOS
# ==========================================================

@st.cache_data
def cargar_datos():

    if not ARCHIVO.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo {ARCHIVO.name}"
        )

    df = pd.read_excel(
        ARCHIVO,
        sheet_name="Data"
    )

    df.columns = (
        df.columns.astype(str)
        .str.strip()
    )

    equivalencias = {
        "POSICION": "POSICIÓN",
        "Fecha OC": "FECHA OC",
        "Fecha contabiliz.": "FECHA CONTABILIZACIÓN",
        "Un.medida de entrada": "UNIDAD"
    }

    df = df.rename(columns=equivalencias)

    columnas_obligatorias = [
        "Material",
        "Descripción material",
        "Pedido",
        "FECHA OC",
        "PRECIO",
        "CANTIDAD",
        "MONEDA"
    ]

    faltantes = [
        columna
        for columna in columnas_obligatorias
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en el Excel: "
            + ", ".join(faltantes)
        )

    df["FECHA OC"] = convertir_fecha(
        df["FECHA OC"]
    )

    df["PRECIO"] = convertir_numero(
        df["PRECIO"]
    )

    df["CANTIDAD"] = convertir_numero(
        df["CANTIDAD"]
    )

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

    df["Pedido"] = (
        df["Pedido"]
        .astype(str)
        .str.strip()
    )

    if "POSICIÓN" not in df.columns:
        df["POSICIÓN"] = "Sin posición"

    df["POSICIÓN"] = (
        df["POSICIÓN"]
        .fillna("Sin posición")
        .astype(str)
    )

    # Precio unitario

    df["PRECIO UNITARIO"] = np.where(
        df["CANTIDAD"] > 0,
        df["PRECIO"] / df["CANTIDAD"],
        np.nan
    )

    # Eliminar registros inválidos

    df = df.dropna(
        subset=[
            "FECHA OC",
            "PRECIO",
            "CANTIDAD",
            "PRECIO UNITARIO"
        ]
    )

    df = df[
        (df["CANTIDAD"] > 0) &
        (df["PRECIO UNITARIO"] > 0) &
        (df["MONEDA"].isin(["CLP", "USD"]))
    ].copy()

    df = df.sort_values(
        ["Material", "FECHA OC", "Pedido", "POSICIÓN"]
    )

    # Meses en español

    meses = {
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

    df["MES"] = df["FECHA OC"].dt.month.map(meses)

    df["MES AÑO"] = (
        df["MES"] + " " +
        df["FECHA OC"].dt.year.astype(str)
    )

    return df


# ==========================================================
# 6. FUNCIONES AUXILIARES
# ==========================================================

def variacion_porcentual(inicial, final):

    if pd.isna(inicial) or inicial == 0:
        return np.nan

    return (
        (final - inicial) / inicial
    ) * 100


def formato_precio(valor, moneda):

    if pd.isna(valor):
        return "-"

    if moneda == "CLP":

        return (
            "$ " +
            f"{valor:,.0f}".replace(",", ".")
        )

    return f"US$ {valor:,.4f}"


# ==========================================================
# 7. CARGAR INFORMACIÓN
# ==========================================================

try:

    df = cargar_datos()

except Exception as error:

    st.error(
        f"Error al cargar los datos: {error}"
    )

    st.stop()


# ==========================================================
# 8. ENCABEZADO
# ==========================================================

st.markdown(
    "<h1>EVOLUCIÓN DE PRECIOS</h1>",
    unsafe_allow_html=True
)

st.markdown(
    "**RÍO LOA | CALAMA**"
)

st.caption(
    "Análisis de la evolución de precios de materias primas "
    "según órdenes de compra."
)

st.divider()


# ==========================================================
# 9. FILTROS DE ANÁLISIS
# ==========================================================

st.sidebar.title("Filtros de análisis")

st.sidebar.markdown(
    "### Período de análisis"
)

fecha_min = df["FECHA OC"].min().date()
fecha_max = df["FECHA OC"].max().date()

fechas = st.sidebar.date_input(
    "Seleccionar fechas",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max
)

if not isinstance(fechas, (tuple, list)) or len(fechas) != 2:

    st.info(
        "Selecciona una fecha inicial y una fecha final."
    )

    st.stop()

fecha_inicio, fecha_fin = fechas

if fecha_inicio > fecha_fin:

    st.error(
        "La fecha inicial no puede ser posterior a la final."
    )

    st.stop()


# Selector de moneda en la parte principal

st.markdown(
    "### Seleccionar moneda"
)

monedas_disponibles = sorted(
    df["MONEDA"].unique()
)

moneda_seleccionada = st.selectbox(
    "Moneda de análisis",
    options=monedas_disponibles,
    index=0,
    label_visibility="collapsed"
)


# Aplicar filtros

filtrado = df[
    (df["FECHA OC"].dt.date >= fecha_inicio) &
    (df["FECHA OC"].dt.date <= fecha_fin) &
    (df["MONEDA"] == moneda_seleccionada)
].copy()

if filtrado.empty:

    st.warning(
        "No existen registros para la moneda y período seleccionados."
    )

    st.stop()


# ==========================================================
# 10. RESUMEN GENERAL
# ==========================================================

st.markdown(
    "### Resumen general"
)

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
    "Moneda seleccionada",
    moneda_seleccionada
)

st.divider()


# ==========================================================
# 11. TOP 4 MATERIAS PRIMAS CON MAYORES VARIACIONES
# ==========================================================

st.markdown(
    "### Evolución de precios: Top 4 materias primas"
)

st.caption(
    "Se muestran las cuatro materias primas con mayores "
    "variaciones porcentuales, considerando aumentos "
    "y disminuciones de precios."
)

variaciones = []

for material, grupo in filtrado.groupby("Material"):

    grupo = grupo.sort_values("FECHA OC")

    if len(grupo) < 2:
        continue

    precio_inicial = grupo.iloc[0]["PRECIO UNITARIO"]

    precio_final = grupo.iloc[-1]["PRECIO UNITARIO"]

    variacion = variacion_porcentual(
        precio_inicial,
        precio_final
    )

    if pd.isna(variacion):
        continue

    variaciones.append({

        "Material": material,

        "Descripción": grupo.iloc[0]["Descripción material"],

        "Precio inicial": precio_inicial,

        "Precio final": precio_final,

        "Variación (%)": variacion,

        "Variación absoluta": abs(variacion)

    })


ranking_grafico = pd.DataFrame(variaciones)


if ranking_grafico.empty:

    st.info(
        "No existen suficientes registros para calcular variaciones."
    )

else:

    # Seleccionar máximo cuatro materias primas

    top_4 = (
        ranking_grafico
        .sort_values(
            "Variación absoluta",
            ascending=False
        )
        .head(4)
    )

    materiales_top = top_4["Material"].tolist()

    evolucion = filtrado[
        filtrado["Material"].isin(materiales_top)
    ].copy()

    evolucion = evolucion.sort_values(
        "FECHA OC"
    )

    # Gráfico de líneas

    fig = px.line(

        evolucion,

        x="FECHA OC",

        y="PRECIO UNITARIO",

        color="Descripción material",

        markers=True,

        hover_data={

            "Material": True,

            "Pedido": True,

            "POSICIÓN": True,

            "FECHA OC": True,

            "PRECIO UNITARIO": ":.4f"

        },

        labels={

            "FECHA OC": "Fecha de creación de OC",

            "PRECIO UNITARIO":
                f"Precio unitario ({moneda_seleccionada})",

            "Descripción material": "Materia prima"

        }

    )

    fig.update_traces(

        line=dict(width=2.5),

        marker=dict(size=7)

    )

    fig.update_layout(

        template="plotly_white",

        height=500,

        hovermode="x unified",

        legend_title="Materia prima",

        margin=dict(
            l=20,
            r=20,
            t=40,
            b=20
        ),

        xaxis=dict(
            title="Fecha de creación de OC",
            tickformat="%b %Y"
        ),

        yaxis=dict(
            title=f"Precio unitario ({moneda_seleccionada})"
        )

    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    with st.expander(
        "Ver las 4 materias primas seleccionadas"
    ):

        tabla_top = top_4[
            [
                "Material",
                "Descripción",
                "Precio inicial",
                "Precio final",
                "Variación (%)"
            ]
        ].copy()

        st.dataframe(

            tabla_top,

            use_container_width=True,

            hide_index=True,

            column_config={

                "Precio inicial":
                    st.column_config.NumberColumn(
                        format="%.4f"
                    ),

                "Precio final":
                    st.column_config.NumberColumn(
                        format="%.4f"
                    ),

                "Variación (%)":
                    st.column_config.NumberColumn(
                        format="%.2f%%"
                    )

            }

        )


# ==========================================================
# 12. FUNCIÓN DE ANÁLISIS INDIVIDUAL
# ==========================================================

def mostrar_analisis_individual(filtrado, moneda_seleccionada):

    st.markdown(
        "## Análisis por materia prima"
    )

    st.caption(
        "Selecciona una materia prima para consultar "
        "su evolución individual."
    )

    materiales = (

        filtrado[
            [
                "Material",
                "Descripción material"
            ]
        ]

        .drop_duplicates()

        .sort_values("Descripción material")

    )

    opciones = {

        f"{fila['Descripción material']} ({fila['Material']})":
        fila["Material"]

        for _, fila in materiales.iterrows()

    }

    seleccion = st.selectbox(

        "Seleccionar materia prima",

        options=list(opciones.keys()),

        key="selector_material"

    )

    material_seleccionado = opciones[seleccion]

    detalle = filtrado[
        filtrado["Material"] == material_seleccionado
    ].copy()

    detalle = detalle.sort_values(
        "FECHA OC"
    )

    # ------------------------------------------------------
    # INDICADORES DEL MATERIAL
    # ------------------------------------------------------

    st.markdown(
        f"### {seleccion}"
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Registros de compra",
        f"{len(detalle):,}".replace(",", ".")
    )

    col2.metric(
        "Cantidad total",
        f"{detalle['CANTIDAD'].sum():,.2f}"
    )

    col3.metric(
        "Órdenes de compra",
        f"{detalle['Pedido'].nunique():,}".replace(",", ".")
    )

    # ------------------------------------------------------
    # GRÁFICO INDIVIDUAL DE LÍNEAS
    # ------------------------------------------------------

    st.markdown(
        "### Evolución histórica del precio unitario"
    )

    fig_detalle = px.line(

        detalle,

        x="FECHA OC",

        y="PRECIO UNITARIO",

        markers=True,

        hover_data={

            "Pedido": True,

            "POSICIÓN": True,

            "CANTIDAD": True,

            "FECHA OC": True,

            "PRECIO UNITARIO": ":.4f"

        },

        labels={

            "FECHA OC": "Fecha de creación de OC",

            "PRECIO UNITARIO":
                f"Precio unitario ({moneda_seleccionada})"

        }

    )

    fig_detalle.update_traces(

        line=dict(
            color=CELESTE,
            width=2.5
        ),

        marker=dict(size=8)

    )

    fig_detalle.update_layout(

        template="plotly_white",

        height=430,

        hovermode="x unified",

        xaxis=dict(
            title="Fecha de creación de OC",
            tickformat="%b %Y"
        ),

        yaxis=dict(
            title=f"Precio unitario ({moneda_seleccionada})"
        )

    )

    st.plotly_chart(
        fig_detalle,
        use_container_width=True
    )

    # ------------------------------------------------------
    # VARIACIÓN ACUMULADA
    # ------------------------------------------------------

    if len(detalle) >= 2:

        inicial = detalle.iloc[0]["PRECIO UNITARIO"]

        final = detalle.iloc[-1]["PRECIO UNITARIO"]

        variacion = variacion_porcentual(
            inicial,
            final
        )

        st.markdown(
            "### Variación acumulada del precio"
        )

        a, b, c = st.columns(3)

        a.metric(
            "Primer precio registrado",
            formato_precio(
                inicial,
                moneda_seleccionada
            )
        )

        b.metric(
            "Último precio registrado",
            formato_precio(
                final,
                moneda_seleccionada
            )
        )

        c.metric(
            "Variación acumulada",
            f"{variacion:+.2f}%"
        )

    # ------------------------------------------------------
    # HISTORIAL DE ÓRDENES DE COMPRA
    # ------------------------------------------------------

    st.markdown(
        "### Historial de órdenes de compra"
    )

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

        columnas_mostrar.insert(
            6,
            "UNIDAD"
        )

    if "FECHA CONTABILIZACIÓN" in detalle.columns:

        columnas_mostrar.append(
            "FECHA CONTABILIZACIÓN"
        )

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

    # ------------------------------------------------------
    # DESCARGA CSV
    # ------------------------------------------------------

    st.markdown(
        "### Exportar información"
    )

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


# ==========================================================
# 13. SECCIÓN CON FONDO CELESTE
# ==========================================================

with st.container(key="analisis_material"):

    mostrar_analisis_individual(
        filtrado,
        moneda_seleccionada
    )


# ==========================================================
# 14. PIE DE PÁGINA
# ==========================================================

st.divider()

st.caption(
    "Dashboard de análisis de precios | Río Loa | "
    "Fuente: registros de órdenes de compra."
)

st.caption(
    "Precio unitario = PRECIO / CANTIDAD. "
    "La variación porcentual compara el primer y último "
    "precio unitario registrado dentro del período seleccionado."
)
