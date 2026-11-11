
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
    .stApp {
        background-color: #FFFFFF;
    }

    div[data-testid="stMetric"] {
        background-color: #F4F8FB;
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

BASE_DIR = Path(__file__).parent

archivos_excel = (
    list(BASE_DIR.glob("*.xlsx"))
    + list(BASE_DIR.glob("*.xls"))
)

archivo = None

for candidato in archivos_excel:
    if "precios rio loa" in candidato.name.lower():
        archivo = candidato
        break

if archivo is None:
    st.error(
        "No se encontró el archivo Excel. "
        "Debe estar en la misma carpeta que app.py "
        "y su nombre debe contener 'PRECIOS RIO LOA'."
    )
    st.stop()


# =========================================================
# 3. CARGAR Y LIMPIAR DATOS
# =========================================================

@st.cache_data
def cargar_datos(ruta_archivo, fecha_modificacion):

    df = pd.read_excel(
        ruta_archivo,
        sheet_name="Data"
    )

    df.columns = df.columns.astype(str).str.strip()

    # Normalizar nombres de columnas
    equivalencias = {
        "POSICION": "POSICIÓN",
        "Fecha OC": "FECHA OC",
        "Fecha contabiliz.": "FECHA CONTABILIZACIÓN",
        "Un.medida de entrada": "UNIDAD",
        "UNIDAD DE MEDIDA": "UNIDAD"
    }

    df = df.rename(columns=equivalencias)

    # Pasar columnas a mayúsculas
    df.columns = [
        columna.upper().strip()
        for columna in df.columns
    ]

    # Verificar columnas necesarias
    columnas_necesarias = [
        "PEDIDO",
        "MATERIAL",
        "DESCRIPCIÓN MATERIAL",
        "CANTIDAD",
        "PRECIO",
        "MONEDA",
        "FECHA OC"
    ]

    faltantes = [
        columna
        for columna in columnas_necesarias
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en el Excel: "
            + ", ".join(faltantes)
        )

    # Limpiar códigos
    for columna in ["PEDIDO", "MATERIAL"]:

        df[columna] = (
            df[columna]
            .astype("string")
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )

    # Limpiar descripción
    df["DESCRIPCIÓN MATERIAL"] = (
        df["DESCRIPCIÓN MATERIAL"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Convertir números
    for columna in ["CANTIDAD", "PRECIO"]:

        df[columna] = pd.to_numeric(
            df[columna],
            errors="coerce"
        )

    # Convertir fechas
    df["FECHA OC"] = pd.to_datetime(
        df["FECHA OC"],
        errors="coerce",
        dayfirst=True
    )

    if "FECHA CONTABILIZACIÓN" in df.columns:

        df["FECHA CONTABILIZACIÓN"] = pd.to_datetime(
            df["FECHA CONTABILIZACIÓN"],
            errors="coerce",
            dayfirst=True
        )

    # Recuperar información de posiciones de una misma
    # combinación Pedido + Material.
    # No se rellenan las cantidades.

    columnas_rellenar = [
        "PRECIO",
        "MONEDA",
        "FECHA OC"
    ]

    if "FECHA CONTABILIZACIÓN" in df.columns:
        columnas_rellenar.append(
            "FECHA CONTABILIZACIÓN"
        )

    for columna in columnas_rellenar:

        df[columna] = (
            df.groupby(
                ["PEDIDO", "MATERIAL"]
            )[columna]
            .transform(
                lambda x: x.ffill().bfill()
            )
        )

    # Normalizar moneda
    df["MONEDA"] = (
        df["MONEDA"]
        .astype("string")
        .str.upper()
        .str.strip()
    )

    # Crear unidad si no existe
    if "UNIDAD" not in df.columns:
        df["UNIDAD"] = ""

    df["UNIDAD"] = (
        df["UNIDAD"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Filtrar registros válidos
    df = df[
        df["PEDIDO"].notna()
        & df["MATERIAL"].notna()
        & df["FECHA OC"].notna()
        & df["CANTIDAD"].notna()
        & df["PRECIO"].notna()
        & (df["CANTIDAD"] > 0)
        & (df["PRECIO"] > 0)
        & df["MONEDA"].isin(["CLP", "USD"])
    ].copy()

    # Calcular precio unitario
    df["PRECIO UNITARIO"] = (
        df["PRECIO"] / df["CANTIDAD"]
    )

    # Crear mes y año
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
    df["AÑO"] = df["FECHA OC"].dt.year

    return df


try:

    fecha_modificacion = archivo.stat().st_mtime

    datos = cargar_datos(
        str(archivo),
        fecha_modificacion
    )

except Exception as error:

    st.error(f"Error al cargar el Excel: {error}")
    st.stop()


if datos.empty:
    st.warning("No existen registros válidos para analizar.")
    st.stop()


# =========================================================
# 4. FUNCIONES AUXILIARES
# =========================================================

def formato_precio(valor, moneda):

    if pd.isna(valor):
        return "-"

    if moneda == "CLP":
        return (
            "$" + f"{valor:,.0f}".replace(",", ".")
        )

    return f"US$ {valor:,.2f}"


def obtener_nombre_material(df, material):

    fila = df[
        df["MATERIAL"] == material
    ]

    if fila.empty:
        return material

    return fila.iloc[0]["DESCRIPCIÓN MATERIAL"]


def calcular_variacion(df):

    df = df.sort_values("FECHA OC")

    if len(df) < 2:
        return None

    precio_inicial = df.iloc[0]["PRECIO UNITARIO"]
    precio_final = df.iloc[-1]["PRECIO UNITARIO"]

    if pd.isna(precio_inicial) or precio_inicial == 0:
        return None

    return (
        (precio_final - precio_inicial)
        / precio_inicial
    ) * 100


# =========================================================
# 5. FILTROS DE FECHA
# =========================================================

st.sidebar.header("Filtros")

fecha_min = datos["FECHA OC"].min().date()
fecha_max = datos["FECHA OC"].max().date()

rango_fechas = st.sidebar.date_input(
    "Período de análisis",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max
)

if (
    isinstance(rango_fechas, (tuple, list))
    and len(rango_fechas) == 2
):

    fecha_inicio = pd.to_datetime(
        rango_fechas[0]
    )

    fecha_fin = (
        pd.to_datetime(rango_fechas[1])
        + pd.Timedelta(days=1)
        - pd.Timedelta(seconds=1)
    )

    datos_filtrados = datos[
        (datos["FECHA OC"] >= fecha_inicio)
        & (datos["FECHA OC"] <= fecha_fin)
    ].copy()

else:

    datos_filtrados = datos.copy()


if datos_filtrados.empty:
    st.warning("No existen datos para el período seleccionado.")
    st.stop()


# =========================================================
# 6. SELECTOR DE MONEDA
# =========================================================

st.divider()

st.header("Moneda de análisis")

monedas_disponibles = [
    moneda
    for moneda in ["CLP", "USD"]
    if moneda in datos_filtrados["MONEDA"].unique()
]

if not monedas_disponibles:
    st.warning("No hay datos disponibles en CLP o USD.")
    st.stop()

if (
    "moneda_seleccionada" not in st.session_state
    or st.session_state["moneda_seleccionada"]
    not in monedas_disponibles
):
    st.session_state["moneda_seleccionada"] = monedas_disponibles[0]

moneda = st.selectbox(
    "Selecciona la moneda:",
    options=monedas_disponibles,
    key="moneda_seleccionada"
)

datos_moneda = datos_filtrados[
    datos_filtrados["MONEDA"] == moneda
].copy()


if datos_moneda.empty:
    st.warning("No hay registros para la moneda seleccionada.")
    st.stop()


# =========================================================
# 7. ANÁLISIS GENERAL
# =========================================================

st.divider()

st.header("1. Análisis general")

col1, col2, col3, col4 = st.columns(4)

total_ocs = datos_moneda["PEDIDO"].nunique()

total_materiales = datos_moneda["MATERIAL"].nunique()

precio_promedio = datos_moneda["PRECIO UNITARIO"].mean()

precio_maximo = datos_moneda["PRECIO UNITARIO"].max()

col1.metric(
    "Órdenes de compra",
    f"{total_ocs:,}".replace(",", ".")
)

col2.metric(
    "Materias primas",
    f"{total_materiales:,}".replace(",", ".")
)

col3.metric(
    "Precio unitario promedio",
    formato_precio(precio_promedio, moneda)
)

col4.metric(
    "Precio unitario máximo",
    formato_precio(precio_maximo, moneda)
)


# =========================================================
# 8. TOP 4 MATERIAS PRIMAS CON MAYOR VARIACIÓN
# =========================================================

st.subheader(
    "Top 4 materias primas con mayor variación porcentual"
)

variaciones = []

for material in datos_moneda["MATERIAL"].unique():

    df_material = datos_moneda[
        datos_moneda["MATERIAL"] == material
    ].copy()

    variacion = calcular_variacion(df_material)

    if variacion is not None:

        variaciones.append({
            "MATERIAL": material,
            "DESCRIPCIÓN": obtener_nombre_material(
                datos_moneda,
                material
            ),
            "VARIACIÓN (%)": variacion
        })


if variaciones:

    df_variaciones = pd.DataFrame(variaciones)

    df_variaciones["VARIACIÓN ABSOLUTA"] = (
        df_variaciones["VARIACIÓN (%)"].abs()
    )

    # Seleccionar las cuatro mayores variaciones
    df_top = (
        df_variaciones
        .nlargest(4, "VARIACIÓN ABSOLUTA")
        .sort_values("VARIACIÓN (%)")
    )

    fig_variaciones = px.bar(
        df_top,
        x="VARIACIÓN (%)",
        y="DESCRIPCIÓN",
        orientation="h",
        text=df_top["VARIACIÓN (%)"].map(
            lambda valor: f"{valor:+.1f}%"
        ),
        title="Top 4 materias primas con mayor variación porcentual",
        labels={
            "VARIACIÓN (%)": "Variación porcentual",
            "DESCRIPCIÓN": "Materia prima"
        }
    )

    fig_variaciones.update_traces(
        marker_color="#3987B8",
        textposition="outside",
        cliponaxis=False
    )

    fig_variaciones.update_layout(
        height=380,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(
            title="Variación porcentual",
            ticksuffix="%",
            showgrid=True,
            gridcolor="#E9EFF3",
            zeroline=True,
            zerolinecolor="#AAB7C4"
        ),
        yaxis=dict(
            title="Materia prima",
            showgrid=False
        ),
        margin=dict(
            l=20,
            r=90,
            t=50,
            b=40
        )
    )

    st.plotly_chart(
        fig_variaciones,
        use_container_width=True
    )

    st.caption(
        "La variación porcentual compara el primer precio unitario "
        "registrado con el último precio unitario del período "
        "seleccionado. Se consideran las variaciones positivas y negativas."
    )

else:

    st.info(
        "No existen suficientes registros para calcular variaciones."
    )


# =========================================================
# 9. EVOLUCIÓN GENERAL DE PRECIOS
# =========================================================

st.subheader(
    "Evolución de precios de las principales materias primas"
)

if variaciones:

    materiales_top = df_top["MATERIAL"].tolist()

    datos_top = datos_moneda[
        datos_moneda["MATERIAL"].isin(materiales_top)
    ].copy()

    datos_top["MATERIA PRIMA"] = (
        datos_top["DESCRIPCIÓN MATERIAL"]
        + " ("
        + datos_top["MATERIAL"]
        + ")"
    )

    datos_top = datos_top.sort_values("FECHA OC")

    fig_evolucion = px.line(
        datos_top,
        x="FECHA OC",
        y="PRECIO UNITARIO",
        color="MATERIA PRIMA",
        markers=True,
        title="Evolución del precio unitario",
        labels={
            "FECHA OC": "Fecha de creación de OC",
            "PRECIO UNITARIO": f"Precio unitario ({moneda})",
            "MATERIA PRIMA": "Materia prima"
        }
    )

    fig_evolucion.update_layout(
        height=480,
        hovermode="x unified",
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
        fig_evolucion,
        use_container_width=True
    )


# =========================================================
# 10. ANÁLISIS INDIVIDUAL
# =========================================================

st.divider()

st.header("2. Análisis individual de materias primas")

# Preparar lista de materiales
materiales = (
    datos_moneda[
        ["MATERIAL", "DESCRIPCIÓN MATERIAL"]
    ]
    .drop_duplicates()
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
# 11. FILTRAR MATERIAS PRIMAS
# =========================================================

if busqueda:

    coincidencias = materiales[
        materiales["MATERIAL"].str.contains(
            busqueda,
            case=False,
            na=False,
            regex=False
        )
        |
        materiales["DESCRIPCIÓN MATERIAL"].str.contains(
            busqueda,
            case=False,
            na=False,
            regex=False
        )
    ].copy()

else:

    coincidencias = materiales.copy()


if coincidencias.empty:

    st.warning(
        "No se encontraron materias primas que coincidan con la búsqueda."
    )

else:

    # Crear opciones del selector
    opciones = [
        f"{fila['DESCRIPCIÓN MATERIAL']} ({fila['MATERIAL']})"
        for _, fila in coincidencias.iterrows()
    ]

    mapa_opciones = {
        f"{fila['DESCRIPCIÓN MATERIAL']} ({fila['MATERIAL']})":
        fila["MATERIAL"]
        for _, fila in coincidencias.iterrows()
    }

    # Clave diferente según la búsqueda.
    # Esto permite que el selector vuelva a la primera
    # coincidencia cada vez que se escribe otra búsqueda.
    clave_selector = (
        "selector_material_rio_"
        + busqueda.lower()
    )

    seleccion = st.selectbox(
        "Selecciona una materia prima:",
        options=opciones,
        index=0,
        key=clave_selector
    )

    material_seleccionado = mapa_opciones[seleccion]

    # Obtener datos del material
    df_material = datos_moneda[
        datos_moneda["MATERIAL"] == material_seleccionado
    ].copy()

    df_material = df_material.sort_values(
        "FECHA OC"
    )

    nombre_material = obtener_nombre_material(
        datos_moneda,
        material_seleccionado
    )

    st.markdown("---")

    st.subheader(nombre_material)

    st.caption(
        f"Código SAP: {material_seleccionado}"
    )


    # =====================================================
    # 12. INDICADORES INDIVIDUALES
    # =====================================================

    precio_inicial = df_material.iloc[0]["PRECIO UNITARIO"]

    precio_actual = df_material.iloc[-1]["PRECIO UNITARIO"]

    variacion_material = None

    if precio_inicial != 0:

        variacion_material = (
            (precio_actual - precio_inicial)
            / precio_inicial
        ) * 100

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Precio inicial",
        formato_precio(precio_inicial, moneda)
    )

    c2.metric(
        "Último precio",
        formato_precio(precio_actual, moneda)
    )

    if variacion_material is not None:

        c3.metric(
            "Variación del período",
            f"{variacion_material:+.2f}%"
        )

    else:

        c3.metric(
            "Variación del período",
            "-"
        )

    c4.metric(
        "Órdenes de compra",
        f"{df_material['PEDIDO'].nunique():,}".replace(",", ".")
    )


    # =====================================================
    # 13. GRÁFICO INDIVIDUAL
    # =====================================================

    st.markdown(
        "### Evolución histórica del precio unitario"
    )

    fig_individual = go.Figure()

    fig_individual.add_trace(
        go.Scatter(
            x=df_material["FECHA OC"],
            y=df_material["PRECIO UNITARIO"],
            mode="lines+markers",
            name="Precio unitario",
            line=dict(
                width=3,
                color="#2878A8"
            ),
            marker=dict(size=8),
            customdata=df_material[
                ["PEDIDO", "CANTIDAD", "UNIDAD"]
            ].fillna("").values,
            hovertemplate=(
                "<b>Fecha:</b> %{x|%d-%m-%Y}<br>"
                "<b>Precio unitario:</b> %{y:,.4f}<br>"
                "<b>OC:</b> %{customdata[0]}<br>"
                "<b>Cantidad:</b> %{customdata[1]} "
                "%{customdata[2]}<extra></extra>"
            )
        )
    )

    fig_individual.update_layout(
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
        fig_individual,
        use_container_width=True
    )


    # =====================================================
    # 14. HISTORIAL DE ÓRDENES DE COMPRA
    # =====================================================

    st.markdown(
        "### Historial de órdenes de compra"
    )

    columnas_historial = [
        "PEDIDO",
        "MATERIAL",
        "DESCRIPCIÓN MATERIAL",
        "CANTIDAD",
        "UNIDAD",
        "PRECIO",
        "MONEDA",
        "PRECIO UNITARIO",
        "FECHA OC"
    ]

    if "FECHA CONTABILIZACIÓN" in df_material.columns:

        columnas_historial.append(
            "FECHA CONTABILIZACIÓN"
        )

    columnas_historial = [
        columna
        for columna in columnas_historial
        if columna in df_material.columns
    ]

    historial = df_material[
        columnas_historial
    ].copy()

    historial = historial.sort_values(
        "FECHA OC",
        ascending=False
    )

    st.dataframe(
        historial,
        use_container_width=True,
        hide_index=True
    )

    # Descargar historial
    csv = historial.to_csv(
        index=False,
        sep=";",
        encoding="utf-8-sig"
    )

    st.download_button(
        label="Descargar historial de órdenes de compra",
        data=csv,
        file_name=(
            f"historial_{material_seleccionado}.csv"
        ),
        mime="text/csv"
    )


# =========================================================
# 15. FUENTE DE INFORMACIÓN
# =========================================================

st.divider()

st.caption(
    "Fuente: registros de órdenes de compra SAP de la planta Río Loa. "
    "El precio unitario se calcula dividiendo el precio por la cantidad. "
    "Los resultados dependen de los registros disponibles y del período seleccionado."
)
