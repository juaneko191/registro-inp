from pathlib import Path
from datetime import date
import io

import pandas as pd
import streamlit as st


# ============================================================
# RUTAS Y CONFIGURACIÓN GENERAL
# ============================================================

CARPETA = Path(__file__).resolve().parent

RUTA_FAVICON = CARPETA / "web-app-manifest-512x512.png"

ARCHIVO_EXCEL = (
    CARPETA
    / "Reporte_SolicitudesInversionesNoProgramadas_PMI.xlsx"
)

configuracion_pagina = {
    "page_title": "Dashboard INP",
    "layout": "wide"
}

if RUTA_FAVICON.exists():
    configuracion_pagina["page_icon"] = str(RUTA_FAVICON)

st.set_page_config(**configuracion_pagina)


# ============================================================
# CONSTANTES
# ============================================================

ORDEN_NIVELES = ["GL-MD", "GL-MP", "GN", "GR"]

ESTADO_RECHAZADO = "Rechazado por DGPMI"
ESTADO_PENDIENTE = "Pendiente de evaluación DGPMI"
ESTADO_VALIDADO = "Validado por DGPMI"

ORDEN_ESTADOS = [
    ESTADO_RECHAZADO,
    ESTADO_PENDIENTE,
    ESTADO_VALIDADO
]

COLORES_ESTADO = {
    ESTADO_RECHAZADO: "#f4a0a4",
    ESTADO_PENDIENTE: "#fff48a",
    ESTADO_VALIDADO: "#92d050",
    "Total": "#f2f2f2"
}


# ============================================================
# ESTILOS DEL DASHBOARD
# ============================================================

st.markdown(
    """
    <style>

    /* Reduce el espacio superior general */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    /* Título principal */
    .titulo-principal {
        font-size: 2.25rem;
        font-weight: 800;
        color: #1f2937;
        margin-bottom: 0.15rem;
    }

    .subtitulo-principal {
        color: #5f6368;
        font-size: 1rem;
        margin-bottom: 1.5rem;
    }

    /* Títulos grises de las cuatro secciones */
    .titulo-seccion {
        background-color: #d0d0d0;
        border: 1px solid #9ca3af;
        border-left: 7px solid #4b5563;
        border-radius: 4px;
        color: #000000;
        font-size: 2rem;
        font-weight: 800;
        line-height: 1.2;
        margin-top: 1.8rem;
        margin-bottom: 0.75rem;
        padding: 0.65rem 0.9rem;
    }

    /* Texto debajo del título de una sección */
    .descripcion-seccion {
        color: #5f6368;
        font-size: 0.92rem;
        margin-bottom: 0.8rem;
    }

    /* Contenedor visual de cada tabla */
    .contenedor-tabla {
        background-color: #ffffff;
        border: 1px solid #d1d5db;
        border-radius: 6px;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
        padding: 0.35rem;
        margin-bottom: 1rem;
    }

    /* Separador entre secciones */
    .separador-seccion {
        border-top: 1px solid #d1d5db;
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
    }

    /* Ajuste de los dataframes de Streamlit */
    div[data-testid="stDataFrame"] {
        border: 1px solid #c8c8c8;
        border-radius: 4px;
    }

    /* Botón de descarga */
    .stDownloadButton > button {
        background-color: #1f4e78;
        border: none;
        color: white;
        font-weight: 700;
    }

    .stDownloadButton > button:hover {
        background-color: #173b5c;
        color: white;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FUNCIONES DE APOYO
# ============================================================

def normalizar_columnas(columnas):
    """
    Elimina espacios adicionales y saltos de línea
    de los nombres de las columnas.
    """
    return (
        columnas.astype(str)
        .str.strip()
        .str.replace("\n", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )


def serie_texto_limpia(serie):
    """
    Convierte una serie a texto, reemplaza valores nulos
    y elimina espacios al inicio y al final.
    """
    return (
        serie
        .fillna("")
        .astype(str)
        .str.strip()
    )


def cargar_datos(origen):
    """
    Lee el archivo Excel. La cabecera se encuentra en la
    tercera fila del archivo, por eso se utiliza header=2.
    """
    df = pd.read_excel(
        origen,
        sheet_name=0,
        header=2,
        engine="openpyxl"
    )

    df = (
        df
        .dropna(axis=1, how="all")
        .dropna(how="all")
        .copy()
    )

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
            df[columna] = serie_texto_limpia(df[columna])

    columnas_fecha = [
        "FECHA SOLICITUD",
        "FECHA APROBACION DGPMI"
    ]

    for columna in columnas_fecha:
        if columna in df.columns:
            df[columna] = pd.to_datetime(
                df[columna],
                dayfirst=True,
                errors="coerce"
            )

    mapa_estados = {
        "Registrado por OPMI": ESTADO_PENDIENTE,
        ESTADO_RECHAZADO: ESTADO_RECHAZADO,
        ESTADO_VALIDADO: ESTADO_VALIDADO
    }

    df["ESTADO_RESUMEN"] = (
        df["ESTADO"]
        .map(mapa_estados)
        .fillna(df["ESTADO"])
    )

    # Fecha sin hora para las agrupaciones.
    df["FECHA"] = (
        df["FECHA SOLICITUD"]
        .dt.normalize()
    )

    return df


def conservar_ultimo_registro(df):
    """
    Conserva el registro más reciente de cada inversión.

    Primero utiliza Código Único. Si no existe Código Único,
    utiliza Código Idea. Si ambos están vacíos, genera un
    identificador temporal para no eliminar el registro.
    """
    resultado = df.copy()

    codigo_unico = resultado["CÓDIGO ÚNICO"].copy()

    if "CÓDIGO IDEA" in resultado.columns:
        resultado["ID_INVERSION"] = codigo_unico.fillna(
            resultado["CÓDIGO IDEA"]
        )
    else:
        resultado["ID_INVERSION"] = codigo_unico

    sin_id = resultado["ID_INVERSION"].isna()

    resultado.loc[sin_id, "ID_INVERSION"] = (
        "SIN_ID_"
        + resultado.index[sin_id].astype(str)
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


def agregar_fila_total(tabla):
    """
    Agrega una fila Total al final de la tabla.
    """
    resultado = tabla.copy()

    if "Total" in resultado.index:
        resultado = resultado.drop(index="Total")

    resultado.loc["Total"] = resultado.sum(axis=0)

    return resultado


def titulo_seccion(texto):
    """
    Muestra el encabezado gris utilizado en cada sección.
    """
    st.markdown(
        f'<div class="titulo-seccion">{texto}</div>',
        unsafe_allow_html=True
    )


def formatear_fechas_indice(tabla):
    """
    Convierte las fechas del índice al formato dd/mm/aaaa.
    Conserva el texto Total.
    """
    resultado = tabla.copy()

    nuevo_indice = []

    for valor in resultado.index:
        if isinstance(valor, pd.Timestamp):
            nuevo_indice.append(valor.strftime("%d/%m/%Y"))
        else:
            nuevo_indice.append(str(valor))

    resultado.index = nuevo_indice

    return resultado


def contar_por_nivel(df_conteo):
    """
    Cuenta registros por nivel de gobierno respetando
    el orden institucional definido.
    """
    return (
        df_conteo["NIVEL DE GOBIERNO"]
        .value_counts()
        .reindex(ORDEN_NIVELES, fill_value=0)
        .astype(int)
    )


# ============================================================
# FUNCIONES PARA LAS CUATRO SECCIONES
# ============================================================

def crear_tabla_estado(df):
    """
    Sección 1: Estado de solicitudes.

    Cuenta registros por nivel de gobierno y estado.
    """
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

    tabla = agregar_fila_total(tabla)

    tabla.index.name = "Nivel de gobierno"

    return tabla.astype(int)


def crear_tabla_solicitudes_atenciones(df, fecha_corte):
    """
    Sección 2: Solicitudes y atenciones.

    Del día:
    - Solicitudes: registros cuya fecha de solicitud
      coincide con la fecha de corte.
    - Atenciones: registros cuya fecha de aprobación
      coincide con la fecha de corte y cuyo estado es
      validado o rechazado.
    - Pendientes: solicitudes del día menos atenciones del día.

    Acumulado:
    - Solicitudes: todos los registros.
    - Atenciones validadas.
    - Atenciones rechazadas.
    - Pendientes de evaluación.
    """
    fecha_corte = pd.Timestamp(fecha_corte).normalize()

    estados_atendidos = [
        ESTADO_VALIDADO,
        ESTADO_RECHAZADO
    ]

    solicitudes_dia = df[
        df["FECHA SOLICITUD"]
        .dt.normalize()
        .eq(fecha_corte)
    ]

    atenciones_dia = df[
        df["FECHA APROBACION DGPMI"]
        .dt.normalize()
        .eq(fecha_corte)
        & df["ESTADO_RESUMEN"].isin(estados_atendidos)
    ]

    validados = df[
        df["ESTADO_RESUMEN"] == ESTADO_VALIDADO
    ]

    rechazados = df[
        df["ESTADO_RESUMEN"] == ESTADO_RECHAZADO
    ]

    pendientes = df[
        df["ESTADO_RESUMEN"] == ESTADO_PENDIENTE
    ]

    tabla = pd.DataFrame(
        index=ORDEN_NIVELES
    )

    tabla[("Del día (*)", "Solicitudes (a)")] = (
        contar_por_nivel(solicitudes_dia)
    )

    tabla[("Del día (*)", "Atenciones (b)")] = (
        contar_por_nivel(atenciones_dia)
    )

    tabla[("Del día (*)", "Pendientes (a)-(b)")] = (
        tabla[("Del día (*)", "Solicitudes (a)")]
        - tabla[("Del día (*)", "Atenciones (b)")]
    )

    tabla[("Acumulado", "Solicitudes (c)")] = (
        contar_por_nivel(df)
    )

    tabla[("Acumulado", "Atenciones-Validadas (d)")] = (
        contar_por_nivel(validados)
    )

    tabla[("Acumulado", "Atenciones-Rechazadas (e)")] = (
        contar_por_nivel(rechazados)
    )

    tabla[("Acumulado", "Pendientes (c)-(d+e)")] = (
        contar_por_nivel(pendientes)
    )

    tabla.columns = pd.MultiIndex.from_tuples(
        tabla.columns,
        names=["Estado", ""]
    )

    tabla = agregar_fila_total(tabla)

    tabla.index.name = "Nivel de gobierno"

    return tabla.astype(int)


def crear_tabla_atendidas(df):
    """
    Sección 3: Seguimiento de solicitudes atendidas.

    Considera como atendidos los registros validados
    y rechazados. Los agrupa por fecha de solicitud
    y nivel de gobierno.
    """
    atendidas = df[
        df["ESTADO_RESUMEN"].isin(
            [
                ESTADO_VALIDADO,
                ESTADO_RECHAZADO
            ]
        )
    ].copy()

    if atendidas.empty:
        tabla = pd.DataFrame(
            columns=ORDEN_NIVELES,
            dtype=int
        )
    else:
        tabla = pd.crosstab(
            atendidas["FECHA"],
            atendidas["NIVEL DE GOBIERNO"]
        )

    tabla = tabla.reindex(
        columns=ORDEN_NIVELES,
        fill_value=0
    )

    tabla = tabla.sort_index()

    tabla["Total de atendidos"] = tabla.sum(axis=1)

    tabla = agregar_fila_total(tabla)

    tabla.index.name = "Fecha de solicitud"

    return tabla.astype(int)


def crear_tabla_pendientes(df):
    """
    Sección 4: Pendientes de atención.

    Considera los registros que continúan pendientes
    de evaluación y los agrupa por fecha de solicitud
    y nivel de gobierno.
    """
    pendientes = df[
        df["ESTADO_RESUMEN"] == ESTADO_PENDIENTE
    ].copy()

    if pendientes.empty:
        tabla = pd.DataFrame(
            columns=ORDEN_NIVELES,
            dtype=int
        )
    else:
        tabla = pd.crosstab(
            pendientes["FECHA"],
            pendientes["NIVEL DE GOBIERNO"]
        )

    tabla = tabla.reindex(
        columns=ORDEN_NIVELES,
        fill_value=0
    )

    tabla = tabla.sort_index()

    tabla["Total de pendientes"] = tabla.sum(axis=1)

    tabla = agregar_fila_total(tabla)

    tabla.index.name = "Fecha de solicitud"

    return tabla.astype(int)


# ============================================================
# FORMATOS VISUALES
# ============================================================

def estilo_base_tabla(tabla):
    """
    Formato base para las tablas del dashboard.
    """
    return (
        tabla.style
        .format("{:,.0f}")
        .set_properties(
            **{
                "text-align": "center",
                "font-weight": "600",
                "border": "1px solid #8c8c8c"
            }
        )
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#f2f2f2"),
                        ("color", "#000000"),
                        ("font-weight", "700"),
                        ("text-align", "center"),
                        ("vertical-align", "middle"),
                        ("border", "1px solid #8c8c8c"),
                        ("white-space", "normal")
                    ]
                },
                {
                    "selector": "td",
                    "props": [
                        ("border", "1px solid #8c8c8c"),
                        ("text-align", "center")
                    ]
                }
            ]
        )
    )


def estilo_tabla_estado(tabla):
    """
    Aplica colores por estado a la primera tabla.
    """
    estilos = estilo_base_tabla(tabla)

    estilos = estilos.set_properties(
        subset=[ESTADO_RECHAZADO],
        **{
            "background-color": COLORES_ESTADO[ESTADO_RECHAZADO],
            "color": "#000000"
        }
    )

    estilos = estilos.set_properties(
        subset=[ESTADO_PENDIENTE],
        **{
            "background-color": COLORES_ESTADO[ESTADO_PENDIENTE],
            "color": "#000000"
        }
    )

    estilos = estilos.set_properties(
        subset=[ESTADO_VALIDADO],
        **{
            "background-color": COLORES_ESTADO[ESTADO_VALIDADO],
            "color": "#000000"
        }
    )

    estilos = estilos.set_properties(
        subset=["Total"],
        **{
            "background-color": COLORES_ESTADO["Total"],
            "color": "#000000"
        }
    )

    # Formato especial de la fila Total.
    estilos = estilos.set_properties(
        subset=pd.IndexSlice[["Total"], :],
        **{
            "font-weight": "800",
            "border-top": "2px solid #000000"
        }
    )

    return estilos


def estilo_tabla_general(tabla):
    """
    Aplica formato institucional a las otras tablas.
    """
    estilos = estilo_base_tabla(tabla)

    estilos = estilos.set_properties(
        subset=pd.IndexSlice[["Total"], :],
        **{
            "background-color": "#f2f2f2",
            "font-weight": "800",
            "border-top": "2px solid #000000"
        }
    )

    return estilos


# ============================================================
# EXPORTACIÓN A EXCEL
# ============================================================

def exportar_excel(
    tabla_estado,
    tabla_solicitudes,
    tabla_atendidas,
    tabla_pendientes,
    detalle
):
    """
    Genera un archivo Excel con las cuatro secciones
    y el detalle de registros filtrados.
    """
    salida = io.BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:

        tabla_estado.to_excel(
            writer,
            sheet_name="Estado_solicitudes"
        )

        tabla_solicitudes.to_excel(
            writer,
            sheet_name="Solicitudes_atenciones"
        )

        tabla_atendidas.to_excel(
            writer,
            sheet_name="Seguimiento_atendidas"
        )

        tabla_pendientes.to_excel(
            writer,
            sheet_name="Pendientes_atencion"
        )

        detalle.to_excel(
            writer,
            sheet_name="Detalle_filtrado",
            index=False
        )

        # Ajustes básicos de ancho de columnas.
        for nombre_hoja in writer.sheets:
            hoja = writer.sheets[nombre_hoja]

            for columna in hoja.columns:
                ancho_maximo = 0
                letra_columna = columna[0].column_letter

                for celda in columna:
                    valor = "" if celda.value is None else str(celda.value)
                    ancho_maximo = max(
                        ancho_maximo,
                        len(valor)
                    )

                hoja.column_dimensions[letra_columna].width = min(
                    ancho_maximo + 2,
                    45
                )

    salida.seek(0)

    return salida.getvalue()


# ============================================================
# ENCABEZADO PRINCIPAL
# ============================================================

st.markdown(
    '<div class="titulo-principal">'
    'Registro de Inversiones No Programadas'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitulo-principal">'
    'Seguimiento del estado, atención y pendientes de las '
    'solicitudes de Inversiones No Programadas'
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# CARGA DEL ARCHIVO
# ============================================================

try:
    df_original = cargar_datos(ARCHIVO_EXCEL)

except FileNotFoundError:
    st.error(
        "No se encontró el archivo Excel. El archivo debe "
        "estar en la misma carpeta que dashboard.py:\n\n"
        f"{ARCHIVO_EXCEL}"
    )
    st.stop()

except Exception as error:
    st.error(
        f"No se pudo leer el archivo Excel: {error}"
    )
    st.stop()


columnas_requeridas = [
    "CÓDIGO ÚNICO",
    "ESTADO",
    "NIVEL DE GOBIERNO",
    "FECHA SOLICITUD",
    "FECHA APROBACION DGPMI",
    "FUNCION",
    "ENTIDAD REGISTRADORA"
]

columnas_faltantes = [
    columna
    for columna in columnas_requeridas
    if columna not in df_original.columns
]

if columnas_faltantes:
    st.error(
        "Faltan las siguientes columnas requeridas: "
        + ", ".join(columnas_faltantes)
    )

    st.write(
        "Columnas encontradas:",
        df_original.columns.tolist()
    )

    st.stop()


# ============================================================
# BARRA LATERAL
# ============================================================

st.sidebar.header("Fuente de información")

st.sidebar.success(
    f"Archivo cargado: {ARCHIVO_EXCEL.name}"
)

st.sidebar.header("Criterio de conteo")

solo_ultimo = st.sidebar.checkbox(
    "Usar solo el último registro por inversión",
    value=False,
    help=(
        "Si se activa, conserva el registro más reciente "
        "de cada Código Único o Código Idea. "
        "Si se desactiva, contabiliza todos los registros "
        "del archivo."
    )
)

df_base = (
    conservar_ultimo_registro(df_original)
    if solo_ultimo
    else df_original.copy()
)


# ============================================================
# FECHA DE CORTE
# ============================================================

st.sidebar.header("Fecha de corte")

fechas_disponibles = pd.concat(
    [
        df_base["FECHA SOLICITUD"],
        df_base["FECHA APROBACION DGPMI"]
    ],
    ignore_index=True
).dropna()

if fechas_disponibles.empty:
    fecha_minima = date.today()
    fecha_maxima = date.today()
else:
    fecha_minima = fechas_disponibles.min().date()
    fecha_maxima = max(
        fechas_disponibles.max().date(),
        date.today()
    )

fecha_corte = st.sidebar.date_input(
    "Fecha para el cálculo del día (*)",
    value=date.today(),
    min_value=fecha_minima,
    max_value=fecha_maxima,
    help=(
        "Esta fecha se utiliza en la sección "
        "'Solicitudes y atenciones' para calcular "
        "las solicitudes y atenciones del día."
    )
)


# ============================================================
# FILTROS
# ============================================================

st.sidebar.header("Filtros")

niveles_disponibles = [
    nivel
    for nivel in ORDEN_NIVELES
    if nivel in df_base["NIVEL DE GOBIERNO"].unique()
]

otros_niveles = sorted(
    set(
        df_base["NIVEL DE GOBIERNO"]
        .replace("", pd.NA)
        .dropna()
        .unique()
    )
    - set(niveles_disponibles)
)

niveles_disponibles += otros_niveles

estados_disponibles = sorted(
    df_base["ESTADO_RESUMEN"]
    .replace("", pd.NA)
    .dropna()
    .unique()
)

funciones_disponibles = sorted(
    df_base["FUNCION"]
    .replace("", pd.NA)
    .dropna()
    .unique()
)

filtro_nivel = st.sidebar.multiselect(
    "Nivel de gobierno",
    options=niveles_disponibles,
    default=niveles_disponibles
)

filtro_estado = st.sidebar.multiselect(
    "Estado",
    options=estados_disponibles,
    default=estados_disponibles
)

filtro_funcion = st.sidebar.multiselect(
    "Función",
    options=funciones_disponibles,
    default=funciones_disponibles
)

tipos_disponibles = []

if "TIPO DE INVERSIÓN" in df_base.columns:
    tipos_disponibles = sorted(
        df_base["TIPO DE INVERSIÓN"]
        .replace("", pd.NA)
        .dropna()
        .unique()
    )

    filtro_tipo = st.sidebar.multiselect(
        "Tipo de inversión",
        options=tipos_disponibles,
        default=tipos_disponibles
    )
else:
    filtro_tipo = []


# Rango de fechas para todo el reporte.
fechas_solicitud = (
    df_base["FECHA SOLICITUD"]
    .dropna()
)

rango_fechas = None

if not fechas_solicitud.empty:
    fecha_solicitud_min = fechas_solicitud.min().date()
    fecha_solicitud_max = fechas_solicitud.max().date()

    rango_fechas = st.sidebar.date_input(
        "Rango de fecha de solicitud",
        value=(
            fecha_solicitud_min,
            fecha_solicitud_max
        ),
        min_value=fecha_solicitud_min,
        max_value=fecha_solicitud_max
    )


# ============================================================
# APLICACIÓN DE FILTROS
# ============================================================

mascara = (
    df_base["NIVEL DE GOBIERNO"].isin(filtro_nivel)
    & df_base["ESTADO_RESUMEN"].isin(filtro_estado)
    & df_base["FUNCION"].isin(filtro_funcion)
)

df = df_base.loc[mascara].copy()

if tipos_disponibles:
    df = df[
        df["TIPO DE INVERSIÓN"].isin(filtro_tipo)
    ]

if (
    isinstance(rango_fechas, (tuple, list))
    and len(rango_fechas) == 2
):
    fecha_inicio = pd.Timestamp(rango_fechas[0])

    fecha_fin_exclusiva = (
        pd.Timestamp(rango_fechas[1])
        + pd.Timedelta(days=1)
    )

    df = df[
        df["FECHA SOLICITUD"].between(
            fecha_inicio,
            fecha_fin_exclusiva,
            inclusive="left"
        )
    ]


# ============================================================
# CREACIÓN DE LAS CUATRO TABLAS
# ============================================================

tabla_estado = crear_tabla_estado(df)

tabla_solicitudes = crear_tabla_solicitudes_atenciones(
    df,
    fecha_corte
)

tabla_atendidas = crear_tabla_atendidas(df)

tabla_pendientes = crear_tabla_pendientes(df)


# ============================================================
# SECCIÓN 1: ESTADO DE SOLICITUDES
# ============================================================

titulo_seccion("Estado de solicitudes")

st.markdown(
    '<div class="descripcion-seccion">'
    'Distribución acumulada de solicitudes según el estado '
    'actual y el nivel de gobierno.'
    '</div>',
    unsafe_allow_html=True
)

st.dataframe(
    estilo_tabla_estado(tabla_estado),
    width="stretch",
)


# ============================================================
# SECCIÓN 2: SOLICITUDES Y ATENCIONES
# ============================================================

titulo_seccion("Solicitudes y atenciones")

st.markdown(
    f'<div class="descripcion-seccion">'
    f'(*) Las cifras del día corresponden al '
    f'{pd.Timestamp(fecha_corte).strftime("%d/%m/%Y")}. '
    f'Las atenciones incluyen solicitudes validadas '
    f'y rechazadas por la DGPMI.'
    f'</div>',
    unsafe_allow_html=True
)

st.dataframe(
    estilo_tabla_general(tabla_solicitudes),
    width="stretch",
)


# ============================================================
# SECCIÓN 3: SEGUIMIENTO DE SOLICITUDES ATENDIDAS
# ============================================================

titulo_seccion("Seguimiento de solicitudes atendidas")

st.markdown(
    '<div class="descripcion-seccion">'
    'Solicitudes validadas o rechazadas, agrupadas por '
    'fecha de solicitud y nivel de gobierno.'
    '</div>',
    unsafe_allow_html=True
)

tabla_atendidas_visual = formatear_fechas_indice(
    tabla_atendidas
)

altura_atendidas = min(
    max(
        220,
        38 * (len(tabla_atendidas_visual) + 2)
    ),
    650
)

st.dataframe(
    estilo_tabla_general(tabla_atendidas_visual),
    width="stretch",
    height=altura_atendidas
)


# ============================================================
# SECCIÓN 4: PENDIENTES DE ATENCIÓN
# ============================================================

titulo_seccion("Pendientes de atención")

st.markdown(
    '<div class="descripcion-seccion">'
    'Solicitudes pendientes de evaluación por la DGPMI, '
    'agrupadas por fecha de solicitud y nivel de gobierno.'
    '</div>',
    unsafe_allow_html=True
)

tabla_pendientes_visual = formatear_fechas_indice(
    tabla_pendientes
)

altura_pendientes = min(
    max(
        220,
        38 * (len(tabla_pendientes_visual) + 2)
    ),
    650
)

st.dataframe(
    estilo_tabla_general(tabla_pendientes_visual),
    width="stretch",
    height=altura_pendientes
)


# ============================================================
# DETALLE DE REGISTROS
# ============================================================

titulo_seccion("Detalle de registros")

columnas_detalle = [
    "TIPO DE INP",
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

columnas_detalle = [
    columna
    for columna in columnas_detalle
    if columna in df.columns
]

detalle = df[columnas_detalle].copy()

detalle = detalle.rename(
    columns={
        "ESTADO_RESUMEN": "ESTADO DEL REGISTRO"
    }
)

st.dataframe(
    detalle,
    width="stretch",
    hide_index=True,
    height=450
)

# ============================================================
# DESCARGA DEL REPORTE EN EXCEL
# ============================================================

def exportar_excel(
    tabla_estado,
    tabla_solicitudes,
    tabla_atendidas,
    tabla_pendientes,
    detalle
):
    """
    Genera un archivo Excel con las cuatro secciones
    del reporte y el detalle de registros filtrados.

    La tabla de solicitudes y atenciones se exporta
    sin combinar celdas para evitar errores asociados
    a encabezados multinivel.
    """

    salida = io.BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:

        # Sección 1: Estado de solicitudes
        tabla_estado.to_excel(
            writer,
            sheet_name="Estado_solicitudes"
        )

        # Sección 2: Solicitudes y atenciones
        tabla_solicitudes.to_excel(
            writer,
            sheet_name="Solicitudes_atenciones",
            merge_cells=False
        )

        # Sección 3: Seguimiento de atendidas
        tabla_atendidas.to_excel(
            writer,
            sheet_name="Seguimiento_atendidas"
        )

        # Sección 4: Pendientes de atención
        tabla_pendientes.to_excel(
            writer,
            sheet_name="Pendientes_atencion"
        )

        # Detalle de registros filtrados
        detalle.to_excel(
            writer,
            sheet_name="Detalle_filtrado",
            index=False
        )

    salida.seek(0)

    return salida.getvalue()


# ============================================================
# GENERAR EL ARCHIVO PARA DESCARGA
# ============================================================

# Inicializar la variable evita errores si la generación falla.
excel_descarga = None

try:
    excel_descarga = exportar_excel(
        tabla_estado=tabla_estado,
        tabla_solicitudes=tabla_solicitudes,
        tabla_atendidas=tabla_atendidas,
        tabla_pendientes=tabla_pendientes,
        detalle=detalle
    )

except Exception as error:
    st.error(
        "No se pudo generar el archivo Excel de descarga. "
        f"Detalle del error: {type(error).__name__}: {error}"
    )


# ============================================================
# BOTÓN DE DESCARGA
# ============================================================

if excel_descarga is not None:

    fecha_archivo = pd.Timestamp.today().strftime("%Y%m%d")

    nombre_archivo = (
        f"Reporte_Seguimiento_INP_{fecha_archivo}.xlsx"
    )

    st.download_button(
        label="📥 Descargar reporte completo en Excel",
        data=excel_descarga,
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        width="stretch"
    )
