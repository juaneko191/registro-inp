from pathlib import Path
from datetime import date
import io

import pandas as pd
import plotly.express as px
import streamlit as st

# ============================================================
# CONFIGURACION
# ============================================================
st.set_page_config(
    page_title="Dashboard INP",
    page_icon="📊",
    layout="wide"
)

# Carpeta donde se encuentra dashboard.py
CARPETA = Path(__file__).resolve().parent

# Excel ubicado en la misma carpeta
ARCHIVO_EXCEL = (
    CARPETA
    / "Reporte_SolicitudesInversionesNoProgramadas_PMI.xlsx"
)

ORDEN_NIVELES = ["GL-MD", "GL-MP", "GN", "GR"]
ORDEN_ESTADOS = [
    "Rechazado por DGPMI",
    "Pendiente de evaluación DGPMI",
    "Validado por DGPMI"
]

COLORES_ESTADO = {
    "Rechazado por DGPMI": "#ff0000",
    "Pendiente de evaluación DGPMI": "#0aa9df",
    "Validado por DGPMI": "#a9dfbf"
}


# ============================================================
# FUNCIONES
# ============================================================
def normalizar_columnas(columnas):
    return (
        columnas.astype(str)
        .str.strip()
        .str.replace("\n", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )


@st.cache_data
def cargar_datos(origen):
    # La cabecera esta en la fila 3 de Excel, por eso header=2.
    df = pd.read_excel(
        origen,
        sheet_name=0,
        header=2,
        engine="openpyxl"
    )

    df = df.dropna(axis=1, how="all").dropna(how="all").copy()
    df.columns = normalizar_columnas(df.columns)

    columnas_texto = [
        "ESTADO",
        "NIVEL DE GOBIERNO",
        "FUNCION",
        "ENTIDAD REGISTRADORA",
        "TIPO DE INVERSIÓN",
        "TIPO DE INP"
    ]

    for columna in columnas_texto:
        if columna in df.columns:
            df[columna] = df[columna].fillna("").astype(str).str.strip()

    for columna in ["FECHA SOLICITUD", "FECHA APROBACION DGPMI"]:
        if columna in df.columns:
            df[columna] = pd.to_datetime(
                df[columna],
                dayfirst=True,
                errors="coerce"
            )

    mapa_estados = {
        "Registrado por OPMI": "Pendiente de evaluación DGPMI",
        "Rechazado por DGPMI": "Rechazado por DGPMI",
        "Validado por DGPMI": "Validado por DGPMI"
    }

    df["ESTADO_RESUMEN"] = (
        df["ESTADO"].map(mapa_estados).fillna(df["ESTADO"])
    )
    df["FECHA"] = df["FECHA SOLICITUD"].dt.normalize()

    return df


def conservar_ultimo_registro(df):
    resultado = df.copy()

    if "CÓDIGO IDEA" in resultado.columns:
        resultado["ID_INVERSION"] = resultado["CÓDIGO ÚNICO"].fillna(
            resultado["CÓDIGO IDEA"]
        )
    else:
        resultado["ID_INVERSION"] = resultado["CÓDIGO ÚNICO"]

    # Evita agrupar todas las filas sin codigo como si fueran una sola.
    sin_id = resultado["ID_INVERSION"].isna()
    resultado.loc[sin_id, "ID_INVERSION"] = (
        "SIN_ID_" + resultado.index[sin_id].astype(str)
    )

    resultado = resultado.sort_values(
        "FECHA SOLICITUD",
        na_position="first"
    )
    resultado = resultado.drop_duplicates(
        subset="ID_INVERSION",
        keep="last"
    )

    return resultado


def crear_tabla_estado(df):
    tabla = pd.crosstab(
        df["NIVEL DE GOBIERNO"],
        df["ESTADO_RESUMEN"]
    )

    tabla = tabla.reindex(
        index=ORDEN_NIVELES,
        columns=ORDEN_ESTADOS,
        fill_value=0
    )
    tabla["Total"] = tabla.sum(axis=1)
    tabla.loc["Total"] = tabla.sum(axis=0)
    tabla.index.name = "Nivel de Gobierno"

    return tabla


def crear_tabla_porcentajes(tabla_estado):
    porcentajes = tabla_estado.div(
        tabla_estado["Total"].replace(0, pd.NA),
        axis=0
    ).mul(100)

    return porcentajes.fillna(0).round(0)


def crear_tabla_pendientes(df):
    pendientes = df[
        df["ESTADO_RESUMEN"] == "Pendiente de evaluación DGPMI"
    ].copy()

    tabla = pd.pivot_table(
        pendientes,
        index="FECHA",
        columns="NIVEL DE GOBIERNO",
        values="ESTADO_RESUMEN",
        aggfunc="count",
        fill_value=0
    )

    tabla = tabla.reindex(columns=ORDEN_NIVELES, fill_value=0)
    tabla["Total"] = tabla.sum(axis=1)
    tabla = tabla.sort_index()

    if not tabla.empty:
        tabla.loc["Total"] = tabla.sum(axis=0)

    tabla.index.name = "Fecha de solicitud"
    return tabla


def exportar_excel(tabla_estado, tabla_porcentaje, tabla_pendientes, detalle):
    salida = io.BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        tabla_estado.to_excel(writer, sheet_name="Estado_INP")
        tabla_porcentaje.to_excel(writer, sheet_name="Porcentajes")
        tabla_pendientes.to_excel(writer, sheet_name="Pendientes_Fecha")
        detalle.to_excel(writer, sheet_name="Detalle_Filtrado", index=False)

    salida.seek(0)
    return salida.getvalue()


# ============================================================
# ENCABEZADO
# ============================================================
st.title("📊 Registro de Inversiones No Programadas")
st.caption("Seguimiento del estado de las solicitudes de INP")


# ============================================================
# SELECCION DEL ARCHIVO
# ============================================================
st.sidebar.header("Fuente de información")

archivo_subido = st.sidebar.file_uploader(
    "Cargar reporte Excel",
    type=["xlsx"]
)

if archivo_subido is not None:
    origen = archivo_subido
    nombre_fuente = archivo_subido.name
else:
    origen = ARCHIVO_EXCEL
    nombre_fuente = str(ARCHIVO_EXCEL)

try:
    df_original = cargar_datos(origen)
except FileNotFoundError:
    st.error(
        "No se encontró el archivo Excel. Colóquelo en:\n\n"
        f"{ARCHIVO_EXCEL}"
    )
    st.stop()
except Exception as error:
    st.error(f"No se pudo leer el Excel: {error}")
    st.stop()

columnas_requeridas = [
    "CÓDIGO ÚNICO",
    "ESTADO",
    "NIVEL DE GOBIERNO",
    "FECHA SOLICITUD",
    "FUNCION",
    "ENTIDAD REGISTRADORA"
]

faltantes = [c for c in columnas_requeridas if c not in df_original.columns]

if faltantes:
    st.error("Faltan columnas requeridas: " + ", ".join(faltantes))
    st.write("Columnas encontradas:", df_original.columns.tolist())
    st.stop()

st.sidebar.success(f"Archivo cargado: {Path(nombre_fuente).name}")

# ============================================================
# OPCION DE DUPLICADOS
# ============================================================
st.sidebar.header("Criterio de conteo")

solo_ultimo = st.sidebar.checkbox(
    "Usar solo el último registro por inversión",
    value=False,
    help="Conserva el registro más reciente de cada Código Único."
)

df_base = (
    conservar_ultimo_registro(df_original)
    if solo_ultimo
    else df_original.copy()
)

# ============================================================
# FILTROS
# ============================================================
st.sidebar.header("Filtros")

niveles = sorted(df_base["NIVEL DE GOBIERNO"].dropna().unique())
estados = sorted(df_base["ESTADO_RESUMEN"].dropna().unique())
funciones = sorted(df_base["FUNCION"].dropna().unique())

filtro_nivel = st.sidebar.multiselect(
    "Nivel de gobierno",
    niveles,
    default=niveles
)
filtro_estado = st.sidebar.multiselect(
    "Estado",
    estados,
    default=estados
)
filtro_funcion = st.sidebar.multiselect(
    "Función",
    funciones,
    default=funciones
)

tipos = []
if "TIPO DE INVERSIÓN" in df_base.columns:
    tipos = sorted(df_base["TIPO DE INVERSIÓN"].dropna().unique())
    filtro_tipo = st.sidebar.multiselect(
        "Tipo de inversión",
        tipos,
        default=tipos
    )
else:
    filtro_tipo = []

fechas = df_base["FECHA SOLICITUD"].dropna()
rango = None

if not fechas.empty:
    fecha_min = fechas.min().date()
    fecha_max = fechas.max().date()
    rango = st.sidebar.date_input(
        "Rango de fecha de solicitud",
        value=(fecha_min, fecha_max),
        min_value=fecha_min,
        max_value=fecha_max
    )

# Aplicar filtros
mascara = (
    df_base["NIVEL DE GOBIERNO"].isin(filtro_nivel)
    & df_base["ESTADO_RESUMEN"].isin(filtro_estado)
    & df_base["FUNCION"].isin(filtro_funcion)
)

df = df_base.loc[mascara].copy()

if tipos:
    df = df[df["TIPO DE INVERSIÓN"].isin(filtro_tipo)]

if isinstance(rango, (tuple, list)) and len(rango) == 2:
    inicio = pd.Timestamp(rango[0])
    fin_exclusivo = pd.Timestamp(rango[1]) + pd.Timedelta(days=1)
    df = df[df["FECHA SOLICITUD"].between(
        inicio,
        fin_exclusivo,
        inclusive="left"
    )]


# ============================================================
# INDICADORES
# ============================================================
total = len(df)
pendientes = (df["ESTADO_RESUMEN"] == "Pendiente de evaluación DGPMI").sum()
rechazados = (df["ESTADO_RESUMEN"] == "Rechazado por DGPMI").sum()
validados = (df["ESTADO_RESUMEN"] == "Validado por DGPMI").sum()

costo_total = 0
if "COSTO ACTUALIZADO (S/)" in df.columns:
    costo_total = pd.to_numeric(
        df["COSTO ACTUALIZADO (S/)"],
        errors="coerce"
    ).sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total", f"{total:,}")
c2.metric("Pendientes", f"{pendientes:,}")
c3.metric("Rechazados", f"{rechazados:,}")
c4.metric("Validados", f"{validados:,}")
c5.metric("Costo actualizado", f"S/ {costo_total:,.0f}")


# ============================================================
# TABLAS
# ============================================================
tabla_estado = crear_tabla_estado(df)
tabla_porcentaje = crear_tabla_porcentajes(tabla_estado)
tabla_pendientes = crear_tabla_pendientes(df)

pestana1, pestana2, pestana3 = st.tabs([
    "Estado de INP",
    "Porcentajes",
    "Pendientes por fecha"
])

with pestana1:
    st.subheader("Estado por nivel de gobierno")
    st.dataframe(tabla_estado, width="stretch")

with pestana2:
    st.subheader("Porcentaje por nivel de gobierno")
    st.dataframe(
        tabla_porcentaje.style.format("{:.0f}%"),
        width="stretch"
    )

with pestana3:
    st.subheader("Pendientes por fecha de solicitud")
    mostrar_pendientes = tabla_pendientes.copy()
    nuevo_indice = []
    for valor in mostrar_pendientes.index:
        if isinstance(valor, pd.Timestamp):
            nuevo_indice.append(valor.strftime("%d/%m/%Y"))
        else:
            nuevo_indice.append(valor)
    mostrar_pendientes.index = nuevo_indice
    st.dataframe(mostrar_pendientes, width="stretch")


# ============================================================
# GRAFICOS
# ============================================================
st.divider()
g1, g2 = st.columns(2)

with g1:
    st.subheader("Estado por nivel de gobierno")
    datos_estado = (
        df.groupby(["NIVEL DE GOBIERNO", "ESTADO_RESUMEN"])
        .size()
        .reset_index(name="Cantidad")
    )
    fig_estado = px.bar(
        datos_estado,
        x="NIVEL DE GOBIERNO",
        y="Cantidad",
        color="ESTADO_RESUMEN",
        barmode="stack",
        text_auto=True,
        category_orders={"NIVEL DE GOBIERNO": ORDEN_NIVELES},
        color_discrete_map=COLORES_ESTADO,
        labels={
            "NIVEL DE GOBIERNO": "Nivel de gobierno",
            "ESTADO_RESUMEN": "Estado"
        }
    )
    fig_estado.update_layout(
        legend_title_text="Estado",
        yaxis_title="Número de registros"
    )
    st.plotly_chart(fig_estado, width="stretch")

with g2:
    st.subheader("Solicitudes por función")
    datos_funcion = (
        df["FUNCION"]
        .replace("", pd.NA)
        .dropna()
        .value_counts()
        .head(15)
        .rename_axis("Función")
        .reset_index(name="Cantidad")
        .sort_values("Cantidad")
    )
    fig_funcion = px.bar(
        datos_funcion,
        x="Cantidad",
        y="Función",
        orientation="h",
        text_auto=True,
        color="Cantidad",
        color_continuous_scale="Blues"
    )
    fig_funcion.update_layout(
        coloraxis_showscale=False,
        yaxis_title=""
    )
    st.plotly_chart(fig_funcion, width="stretch")

st.divider()
st.subheader("Evolución diaria de solicitudes")

evolucion = (
    df.dropna(subset=["FECHA"])
    .groupby(["FECHA", "ESTADO_RESUMEN"])
    .size()
    .reset_index(name="Cantidad")
)

fig_evolucion = px.line(
    evolucion,
    x="FECHA",
    y="Cantidad",
    color="ESTADO_RESUMEN",
    markers=True,
    color_discrete_map=COLORES_ESTADO,
    labels={
        "FECHA": "Fecha de solicitud",
        "ESTADO_RESUMEN": "Estado"
    }
)
fig_evolucion.update_layout(
    legend_title_text="Estado",
    yaxis_title="Número de registros"
)
st.plotly_chart(fig_evolucion, width="stretch")

st.divider()
g3, g4 = st.columns(2)

with g3:
    st.subheader("Antigüedad de pendientes")
    df_pend = df[
        df["ESTADO_RESUMEN"] == "Pendiente de evaluación DGPMI"
    ].copy()

    df_pend["DÍAS PENDIENTE"] = (
        pd.Timestamp(date.today())
        - df_pend["FECHA SOLICITUD"].dt.normalize()
    ).dt.days

    df_pend["RANGO"] = pd.cut(
        df_pend["DÍAS PENDIENTE"],
        bins=[-1, 3, 7, 15, float("inf")],
        labels=[
            "0 a 3 días",
            "4 a 7 días",
            "8 a 15 días",
            "Más de 15 días"
        ]
    )

    datos_antiguedad = (
        df_pend["RANGO"]
        .value_counts(sort=False)
        .rename_axis("Rango")
        .reset_index(name="Cantidad")
    )

    fig_antiguedad = px.bar(
        datos_antiguedad,
        x="Rango",
        y="Cantidad",
        text_auto=True,
        color="Rango",
        color_discrete_sequence=[
            "#a9dfbf", "#f9e79f", "#f5b041", "#e74c3c"
        ]
    )
    fig_antiguedad.update_layout(
        showlegend=False,
        yaxis_title="Número de pendientes"
    )
    st.plotly_chart(fig_antiguedad, width="stretch")

with g4:
    st.subheader("Entidades con más solicitudes")
    datos_entidad = (
        df["ENTIDAD REGISTRADORA"]
        .replace("", pd.NA)
        .dropna()
        .value_counts()
        .head(10)
        .rename_axis("Entidad")
        .reset_index(name="Cantidad")
        .sort_values("Cantidad")
    )

    fig_entidad = px.bar(
        datos_entidad,
        x="Cantidad",
        y="Entidad",
        orientation="h",
        text_auto=True,
        color="Cantidad",
        color_continuous_scale="Teal"
    )
    fig_entidad.update_layout(
        coloraxis_showscale=False,
        yaxis_title=""
    )
    st.plotly_chart(fig_entidad, width="stretch")


# ============================================================
# DETALLE Y DESCARGA
# ============================================================
st.divider()
st.subheader("Detalle de registros filtrados")

columnas_detalle = [
    "CÓDIGO ÚNICO",
    "CÓDIGO IDEA",
    "NOMBRE DE INVERSIÓN",
    "ENTIDAD REGISTRADORA",
    "TIPO DE INVERSIÓN",
    "FUNCION",
    "COSTO ACTUALIZADO (S/)",
    "FECHA SOLICITUD",
    "FECHA APROBACION DGPMI",
    "ESTADO_RESUMEN",
    "NIVEL DE GOBIERNO"
]
columnas_detalle = [c for c in columnas_detalle if c in df.columns]
detalle = df[columnas_detalle].copy()
detalle = detalle.rename(columns={"ESTADO_RESUMEN": "ESTADO DEL REGISTRO"})

st.dataframe(detalle, width="stretch", hide_index=True)

excel_descarga = exportar_excel(
    tabla_estado,
    tabla_porcentaje,
    tabla_pendientes,
    detalle
)

st.download_button(
    "📥 Descargar análisis en Excel",
    data=excel_descarga,
    file_name="Analisis_INP_Dashboard.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption(
    f"Registros originales: {len(df_original):,} | "
    f"Registros base: {len(df_base):,} | "
    f"Registros filtrados: {len(df):,}"
)
