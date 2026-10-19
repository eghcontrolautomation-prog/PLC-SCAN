import os
import re
import unicodedata
from pathlib import Path
from urllib.parse import urlparse
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

# Tesseract es opcional para que la aplicación pueda abrirse.
try:
    import pytesseract
    PYTESSERACT_DISPONIBLE = True
except ImportError:
    pytesseract = None
    PYTESSERACT_DISPONIBLE = False


# ==========================================================
# CONFIGURACIÓN
# ==========================================================

CARPETA_PROYECTO = Path(__file__).resolve().parent
ARCHIVO_EXCEL = CARPETA_PROYECTO / "base de datos 2.xlsx"

st.set_page_config(
    page_title="PLC SCAN",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================================
# DISEÑO
# ==========================================================


st.markdown(
    """
    <style>
    /* ========================================
       FONDO PRINCIPAL
    ======================================== */
    .stApp {
        background: linear-gradient(
            135deg,
            #061329,
            #0A2D50,
            #07182E
        );
        color: #F2F6FC;
    }

    /* ========================================
       BARRA LATERAL
    ======================================== */
    [data-testid="stSidebar"] {
        background: #071426;
        border-right: 1px solid #1C426B;
    }

    [data-testid="stSidebar"] * {
        color: #F2F6FC;
    }

    /* ========================================
       TÍTULOS Y TEXTO
    ======================================== */
    h1, h2, h3 {
        color: #66C7FF !important;
    }

    p, label, li {
        color: #E4ECF7;
    }

    /* ========================================
       MÉTRICAS
    ======================================== */
    div[data-testid="stMetric"] {
        background: rgba(17, 49, 82, 0.75);
        border: 1px solid #28618D;
        padding: 15px;
        border-radius: 12px;
    }

    /* ========================================
       BOTONES
    ======================================== */
    div.stButton > button {
        background-color: #087FC1;
        color: white;
        border: 1px solid #28618D;
        border-radius: 8px;
        font-weight: bold;
        transition: 0.3s;
    }

    div.stButton > button:hover {
        background-color: #119EEB;
        color: white;
        border-color: #66C7FF;
    }

    /* ========================================
       EXPANSORES
    ======================================== */
    div[data-testid="stExpander"] {
        border: 1px solid #28577E;
        border-radius: 10px;
        background: rgba(9, 35, 65, 0.30);
    }

    /* ========================================
       ENCABEZADO PLC SCAN
    ======================================== */
    .plc-header {
        position: relative;
        text-align: center;
        padding: 30px 22px;
        border: 1px solid rgba(102, 199, 255, 0.40);
        border-radius: 18px;
        background: linear-gradient(
            120deg,
            rgba(9, 35, 65, 0.96),
            rgba(14, 69, 109, 0.88)
        );
        margin-bottom: 25px;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.25);
    }

    .plc-header h1 {
        color: #66D5FF !important;
        font-size: 42px;
        font-weight: 900;
        letter-spacing: 4px;
        margin-bottom: 14px;
    }

    /* ========================================
       DESCRIPCIÓN DEL PROYECTO
    ======================================== */
    .plc-subtitle {
        color: #D4E9FA !important;
        font-size: 16px;
        line-height: 1.8;
        max-width: 800px;
        margin: 0 auto;
    }

    /* ========================================
       TARJETAS DE INFORMACIÓN
    ======================================== */
    .info-card {
        background: rgba(13, 42, 72, 0.82);
        border: 1px solid #245581;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 12px;
    }

    /* ========================================
       MARCA DE AGUA GRANDE EN EL CENTRO
    ======================================== */
    .stApp::before {
        content: "DJH";
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: max-content;
        font-size: clamp(120px, 22vw, 280px);
        font-weight: 900;
        letter-spacing: 12px;
        color: rgba(120, 190, 255, 0.07);
        pointer-events: none;
        user-select: none;
        z-index: 0;
        white-space: nowrap;
    }

    /* ========================================
       MARCA DE AGUA PEQUEÑA EN LA ESQUINA
    ======================================== */
    .stApp::after {
        content: "DJH";
        position: fixed;
        right: 22px;
        bottom: 15px;
        font-size: 22px;
        font-weight: 900;
        letter-spacing: 4px;
        color: rgba(160, 215, 255, 0.35);
        pointer-events: none;
        user-select: none;
        z-index: 999;
        white-space: nowrap;
    }

    /* ========================================
       AJUSTES PARA CELULARES
    ======================================== */
    @media (max-width: 768px) {
        .plc-header {
            padding: 22px 14px;
        }

        .plc-header h1 {
            font-size: 32px;
        }

        .plc-subtitle {
            font-size: 14px;
        }

        .stApp::before {
            font-size: 100px;
            letter-spacing: 5px;
        }

        .stApp::after {
            right: 12px;
            bottom: 10px;
            font-size: 17px;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ==========================================================
# FUNCIONES GENERALES
# ==========================================================

def normalizar(valor):
    """Normaliza texto para facilitar las búsquedas."""

    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()

    texto = unicodedata.normalize("NFD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    texto = re.sub(r"[^a-z0-9]+", " ", texto)

    return " ".join(texto.split())


def normalizar_ocr(texto):
    """Normaliza texto obtenido de una fotografía."""

    texto = str(texto).lower().strip()

    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    return re.sub(r"[^a-z0-9]", "", texto)


def valor_texto(valor):
    """Evita mostrar NaN y celdas vacías."""

    if pd.isna(valor):
        return "No disponible"

    texto = str(valor).strip()

    if texto.lower() in ("", "nan", "none", "null"):
        return "No disponible"

    return texto


def buscar_columna(df, nombres):
    """Localiza una columna usando posibles nombres."""

    columnas = {
        normalizar(columna): columna
        for columna in df.columns
    }

    for nombre in nombres:
        clave = normalizar(nombre)

        if clave in columnas:
            return columnas[clave]

    for nombre in nombres:
        clave = normalizar(nombre)

        for clave_columna, columna_original in columnas.items():

            if clave and clave in clave_columna:
                return columna_original

    return None


def obtener_valor(fila, nombres):
    """Obtiene el valor de una fila aunque cambie el encabezado."""

    columnas = {
        normalizar(columna): columna
        for columna in fila.index
    }

    for nombre in nombres:
        clave = normalizar(nombre)

        if clave in columnas:
            return valor_texto(fila[columnas[clave]])

    for nombre in nombres:
        clave = normalizar(nombre)

        for clave_columna, columna_original in columnas.items():

            if clave and clave in clave_columna:
                return valor_texto(fila[columna_original])

    return "No disponible"


def es_url(valor):
    """Comprueba si el texto es una dirección web válida."""

    if not valor or valor == "No disponible":
        return False

    try:
        resultado = urlparse(valor)

        return (
            resultado.scheme in ("http", "https")
            and bool(resultado.netloc)
        )

    except Exception:
        return False


def mostrar_manual(valor, texto="Abrir manual"):
    """Muestra el manual si existe una URL."""

    if valor == "No disponible":
        st.caption("No hay un manual registrado.")

    elif es_url(valor):
        st.markdown(f"[{texto}]({valor})")

    else:
        st.write(valor)


def buscar_coincidencias(df, consulta):
    """Busca coincidencias en las columnas principales."""

    consulta_normalizada = normalizar(consulta)

    if not consulta_normalizada:
        return df.copy()

    posibles_columnas = [
        buscar_columna(df, ["Marca", "Fabricante"]),
        buscar_columna(df, ["Modelo exacto", "Modelo"]),
        buscar_columna(df, ["Familia", "Familia PLC"]),
        buscar_columna(
            df,
            ["Serie / Subfamilia", "Serie", "Subfamilia"]
        ),
        buscar_columna(
            df,
            ["Clasificación", "Tipo de PLC", "Tipo"]
        ),
        buscar_columna(df, ["Arquitectura"]),
        buscar_columna(
            df,
            ["Interfaz / conector visible", "Interfaz"]
        ),
        buscar_columna(df, ["Software"]),
    ]

    columnas = [
        columna
        for columna in posibles_columnas
        if columna is not None
    ]

    if not columnas:
        return df.iloc[0:0].copy()

    mascara = pd.Series(False, index=df.index)

    for columna in columnas:

        valores = (
            df[columna]
            .fillna("")
            .astype(str)
            .map(normalizar)
        )

        mascara |= valores.str.contains(
            re.escape(consulta_normalizada),
            regex=True,
            na=False
        )

    return df[mascara].copy()


# ==========================================================
# CARGA DEL EXCEL
# ==========================================================

@st.cache_data
def cargar_base_datos(ruta_archivo, fecha_modificacion):
    """Carga los registros de la primera hoja del Excel."""

    return pd.read_excel(ruta_archivo)


if not ARCHIVO_EXCEL.exists():

    st.error(
        f"No se encontró el archivo: {ARCHIVO_EXCEL.name}"
    )

    st.info(
        "Coloca el archivo Excel junto a PLC SCAN.py y "
        "verifica que su nombre sea exactamente "
        "'base de datos 2.xlsx'."
    )

    st.stop()


try:

    df = cargar_base_datos(
        str(ARCHIVO_EXCEL),
        ARCHIVO_EXCEL.stat().st_mtime
    )

    df.columns = [
        str(columna).strip()
        for columna in df.columns
    ]

    df = df.dropna(how="all")

except Exception as error:

    st.error("No se pudo cargar la base de datos de Excel.")

    st.code(str(error))

    st.stop()

# ==========================================================
# ENCABEZADO
# ==========================================================

# ==========================================================
# ENCABEZADO PLC SCAN
# ==========================================================

st.markdown(
    """
    <div class="plc-header">
        <h1>⚙️ PLC SCAN</h1>
        <p class="plc-subtitle">
            Identificación inteligente de controladores PLC
            y consulta de información técnica para
            programación, diagnóstico y mantenimiento industrial.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==========================================================
# MENÚ LATERAL
# ==========================================================

st.sidebar.title("⚙️ PLC SCAN")

st.sidebar.caption("Automatización Industrial")

seccion = st.sidebar.radio(
    "Selecciona una sección:",
    [
        "Inicio",
        "Buscar PLC",
        "Ficha técnica",
        "Reconocimiento del PLC",
        "Comunicación industrial",
    ],
)

st.sidebar.divider()

st.sidebar.caption(
    "Consulta de modelos, especificaciones, pinout, "
    "software, manuales e interfaces."
)


# ==========================================================
# INICIO
# ==========================================================

if seccion == "Inicio":

    st.subheader("Panel principal")

    st.write(
        "PLC SCAN es una herramienta de consulta técnica "
        "diseñada para facilitar la identificación de "
        "controladores lógicos programables y la revisión "
        "de sus características."
    )

    columna_marca = buscar_columna(
        df,
        ["Marca", "Fabricante"]
    )

    columna_modelo = buscar_columna(
        df,
        ["Modelo exacto", "Modelo"]
    )

    total_marcas = (
        df[columna_marca].dropna().nunique()
        if columna_marca
        else 0
    )

    total_modelos = (
        df[columna_modelo].dropna().nunique()
        if columna_modelo
        else 0
    )

    col1, col2, col3 = st.columns(3)

    col1.metric("Registros disponibles", len(df))

    col2.metric("Marcas registradas", total_marcas)

    col3.metric("Modelos registrados", total_modelos)

    st.markdown("### Funciones disponibles")

    c1, c2 = st.columns(2)

    with c1:

        st.markdown(
            """
            <div class="info-card">
                <h4>🔎 Búsqueda de PLC</h4>
                Consulta equipos por marca, modelo, familia o serie.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <div class="info-card">
                <h4>📋 Ficha técnica</h4>
                Consulta especificaciones, pinout, software,
                cables y manuales.
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:

        st.markdown(
            """
            <div class="info-card">
                <h4>🧠 Reconocimiento visual</h4>
                Analiza fotografías para intentar leer
                las etiquetas de los equipos.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <div class="info-card">
                <h4>🔌 Comunicación industrial</h4>
                Consulta información sobre interfaces y redes.
            </div>
            """,
            unsafe_allow_html=True,
        )


# ==========================================================
# BASE DE DATOS
# ==========================================================

elif seccion == "Base de datos":

    st.subheader("📚 Base de datos de PLC")

    st.write(
        "Registros técnicos disponibles en el archivo Excel."
    )

    st.caption(
        f"Registros: {len(df)} | Columnas: {len(df.columns)}"
    )

    consulta = st.text_input(
        "Filtrar registros:",
        placeholder="Escribe una marca, modelo o característica..."
    )

    if consulta.strip():
        tabla = buscar_coincidencias(df, consulta)
    else:
        tabla = df.copy()

    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "Descargar resultados en CSV",
        data=tabla.to_csv(index=False).encode("utf-8-sig"),
        file_name="PLC_SCAN_resultados.csv",
        mime="text/csv",
    )


# ==========================================================
# BUSCAR PLC
# ==========================================================

elif seccion == "Buscar PLC":

    st.subheader("🔎 Buscar controlador")

    consulta = st.text_input(
        "Ingresa la marca, modelo, familia o serie:",
        placeholder="Ejemplo: Siemens S7-1200"
    )

    if consulta.strip():

        resultados = buscar_coincidencias(df, consulta)

        if resultados.empty:

            st.warning(
                "No se encontraron coincidencias "
                "en la base de datos."
            )

        else:

            st.success(
                f"Se encontraron {len(resultados)} coincidencia(s)."
            )

            for numero, (_, fila) in enumerate(
                resultados.iterrows()
            ):

                marca = obtener_valor(
                    fila,
                    ["Marca", "Fabricante"]
                )

                modelo = obtener_valor(
                    fila,
                    ["Modelo exacto", "Modelo"]
                )

                with st.expander(
                    f"{marca} | {modelo}",
                    expanded=(numero == 0)
                ):

                    st.write(f"**Marca:** {marca}")

                    st.write(f"**Modelo:** {modelo}")

                    st.write(
                        "**Familia:** "
                        + obtener_valor(
                            fila,
                            ["Familia", "Familia PLC"]
                        )
                    )

                    st.write(
                        "**Serie:** "
                        + obtener_valor(
                            fila,
                            ["Serie / Subfamilia", "Serie", "Subfamilia"]
                        )
                    )

                    st.write(
                        "**Clasificación:** "
                        + obtener_valor(
                            fila,
                            ["Clasificación", "Tipo de PLC", "Tipo"]
                        )
                    )

                    st.write(
                        "**Arquitectura:** "
                        + obtener_valor(fila, ["Arquitectura"])
                    )

                    st.write(
                        "**Pinout:** "
                        + obtener_valor(
                            fila,
                            ["Pin Out", "Pinout"]
                        )
                    )

                    st.write(
                        "**Interfaz visible:** "
                        + obtener_valor(
                            fila,
                            [
                                "Interfaz / conector visible",
                                "Interfaz",
                                "Interfaz / conector del pin out"
                            ]
                        )
                    )

                    st.write(
                        "**Cable / adaptador:** "
                        + obtener_valor(
                            fila,
                            ["Cable / adaptador", "Cable"]
                        )
                    )

                    st.write(
                        "**Software:** "
                        + obtener_valor(fila, ["Software"])
                    )

                    manual = obtener_valor(
                        fila,
                        [
                            "URL manual / documentación",
                            "Manual",
                            "Link manual"
                        ]
                    )

                    st.write("**Manual y documentación:**")

                    mostrar_manual(manual, "Abrir manual")


# ==========================================================
# FICHA TÉCNICA
# ==========================================================

elif seccion == "Ficha técnica":

    st.subheader("📋 Ficha técnica del PLC")

    columna_marca = buscar_columna(
        df,
        ["Marca", "Fabricante"]
    )

    columna_modelo = buscar_columna(
        df,
        ["Modelo exacto", "Modelo"]
    )

    if columna_marca is None or columna_modelo is None:

        st.error(
            "No se encontraron las columnas Marca y Modelo. "
            "Revisa los encabezados del Excel."
        )

    elif df.empty:

        st.warning("La base de datos no contiene registros.")

    else:

        opciones = {}

        for indice, fila_opcion in df.iterrows():

            marca_opcion = valor_texto(
                fila_opcion[columna_marca]
            )

            modelo_opcion = valor_texto(
                fila_opcion[columna_modelo]
            )

            etiqueta = f"{marca_opcion} | {modelo_opcion}"

            if etiqueta in opciones:
                etiqueta = f"{etiqueta} | Registro {indice}"

            opciones[etiqueta] = indice

        seleccion = st.selectbox(
            "Selecciona el equipo:",
            list(opciones.keys())
        )

        fila = df.loc[opciones[seleccion]]

        st.markdown(
            f"### {obtener_valor(fila, ['Marca', 'Fabricante'])} "
            f"— {obtener_valor(fila, ['Modelo exacto', 'Modelo'])}"
        )

        col1, col2 = st.columns(2)

        with col1:

            st.markdown("#### Identificación")

            st.write(
                "**Marca:** "
                + obtener_valor(fila, ["Marca", "Fabricante"])
            )

            st.write(
                "**Modelo:** "
                + obtener_valor(fila, ["Modelo exacto", "Modelo"])
            )

            st.write(
                "**Familia:** "
                + obtener_valor(fila, ["Familia", "Familia PLC"])
            )

            st.write(
                "**Serie / Subfamilia:** "
                + obtener_valor(
                    fila,
                    ["Serie / Subfamilia", "Serie", "Subfamilia"]
                )
            )

            st.write(
                "**Clasificación:** "
                + obtener_valor(
                    fila,
                    ["Clasificación", "Tipo de PLC", "Tipo"]
                )
            )

            st.write(
                "**Arquitectura:** "
                + obtener_valor(fila, ["Arquitectura"])
            )

        with col2:

            st.markdown("#### Conexión y programación")

            st.write(
                "**Pinout:** "
                + obtener_valor(fila, ["Pin Out", "Pinout"])
            )

            st.write(
                "**Interfaz / conector:** "
                + obtener_valor(
                    fila,
                    [
                        "Interfaz / conector del pin out",
                        "Interfaz / conector visible",
                        "Interfaz",
                    ]
                )
            )

            st.write(
                "**Cable / adaptador:** "
                + obtener_valor(
                    fila,
                    ["Cable / adaptador", "Cable"]
                )
            )

            st.write(
                "**Software:** "
                + obtener_valor(fila, ["Software"])
            )

            st.write(
                "**Estado del pinout:** "
                + obtener_valor(
                    fila,
                    ["Estado del pin out", "Estado del pinout"]
                )
            )

        st.divider()

        st.markdown("#### Manual y documentación")

        manual = obtener_valor(
            fila,
            [
                "URL manual / documentación",
                "Manual",
                "Link manual"
            ]
        )

        mostrar_manual(
            manual,
            "📘 Abrir manual del fabricante"
        )

        st.markdown("#### Notas para reconocimiento")

        st.write(
            obtener_valor(
                fila,
                ["Notas para reconocimiento", "Notas", "Observaciones"]
            )
        )

        with st.expander("Ver todos los datos del registro"):

            st.dataframe(
                fila.to_frame(name="Valor"),
                use_container_width=True
            )


# ==========================================================
# RECONOCIMIENTO DEL PLC
# ==========================================================

elif seccion == "Reconocimiento del PLC":

  
    st.subheader("🧠 Reconocimiento visual del PLC")

    st.write(
        "Carga una fotografía del PLC para mejorar la imagen, "
        "leer su etiqueta mediante EasyOCR y buscar el modelo "
        "en la base de datos."
    )

    st.info(
        "La identificación es preliminar. Se prioriza la coincidencia "
        "del modelo y no solamente la marca."
    )

    archivo_imagen = st.file_uploader(
        "Selecciona una fotografía del PLC:",
        type=["jpg", "jpeg", "png", "webp"],
        key="imagen_reconocimiento_plc"
    )

    if archivo_imagen is not None:

        try:
            imagen = Image.open(archivo_imagen).convert("RGB")

            st.image(
                imagen,
                caption="Fotografía original",
                use_container_width=True
            )

            if st.button("🔍 Analizar fotografía"):

                try:
                    import easyocr
                    import numpy as np

                    # Cargar EasyOCR una sola vez por sesión.
                    if "easyocr_reader_plc" not in st.session_state:
                        with st.spinner(
                            "Inicializando EasyOCR. "
                            "La primera ejecución puede tardar..."
                        ):
                            st.session_state[
                                "easyocr_reader_plc"
                            ] = easyocr.Reader(
                                ["en"],
                                gpu=False
                            )

                    reader = st.session_state[
                        "easyocr_reader_plc"
                    ]

                    with st.spinner(
                        "Mejorando la imagen y leyendo la etiqueta..."
                    ):

                        # 1. Ampliar imágenes pequeñas sin excederse.
                        ancho, alto = imagen.size
                        factor = max(
                            1,
                            min(3, 1800 // max(ancho, 1))
                        )

                        imagen_ampliada = imagen.resize(
                            (ancho * factor, alto * factor),
                            Image.Resampling.LANCZOS
                        )

                        # 2. Mejorar contraste y nitidez.
                        imagen_mejorada = ImageOps.autocontrast(
                            imagen_ampliada
                        )

                        imagen_mejorada = ImageEnhance.Contrast(
                            imagen_mejorada
                        ).enhance(1.8)

                        imagen_mejorada = ImageEnhance.Sharpness(
                            imagen_mejorada
                        ).enhance(2.0)

                        # 3. Crear versiones adicionales para OCR.
                        imagen_gris = ImageOps.grayscale(
                            imagen_mejorada
                        )

                        imagen_blanco_negro = imagen_gris.point(
                            lambda pixel: 255 if pixel > 145 else 0
                        )

                        versiones = [
                            ("Color mejorado", imagen_mejorada),
                            ("Escala de grises", imagen_gris),
                            ("Blanco y negro", imagen_blanco_negro),
                        ]

                        textos_detectados = []

                        # 4. Ejecutar EasyOCR sobre cada versión.
                        for nombre, version in versiones:

                            resultados = reader.readtext(
                                np.array(version),
                                detail=0,
                                paragraph=False
                            )

                            for resultado in resultados:
                                resultado = str(resultado).strip()

                                if (
                                    resultado
                                    and resultado not in textos_detectados
                                ):
                                    textos_detectados.append(resultado)

                    st.image(
                        imagen_mejorada,
                        caption="Imagen preparada para reconocimiento",
                        use_container_width=True
                    )

                    texto_detectado = "\n".join(
                        textos_detectados
                    )

                    st.markdown("#### Texto detectado")

                    st.text_area(
                        "Texto leído de la etiqueta:",
                        value=texto_detectado,
                        height=180,
                        key="texto_ocr_plc"
                    )

                    if not texto_detectado.strip():

                        st.warning(
                            "EasyOCR no pudo leer texto en esta fotografía. "
                            "Prueba con una imagen más cercana, enfocada "
                            "y con buena iluminación."
                        )

                    else:

                        st.markdown(
                            "#### Posibles coincidencias en la base de datos"
                        )

                        # Normalizar texto y eliminar separadores.
                        texto_normalizado = normalizar_ocr(
                            texto_detectado
                        )

                        texto_compacto = re.sub(
                            r"[^A-Z0-9]",
                            "",
                            texto_normalizado.upper()
                        )

                        # Comparar códigos de modelo, no solo la marca.
                        def comparar_modelo(modelo, texto):
                            modelo_normalizado = normalizar_ocr(
                                modelo
                            )

                            modelo_compacto = re.sub(
                                r"[^A-Z0-9]",
                                "",
                                modelo_normalizado.upper()
                            )

                            if len(modelo_compacto) < 4:
                                return 0.0

                            if modelo_compacto in texto:
                                return 100.0

                            # Comparar ventanas de tamaño parecido
                            # para tolerar errores pequeños del OCR.
                            longitud = len(modelo_compacto)
                            mejor_puntaje = 0.0

                            longitud_minima = max(
                                4,
                                longitud - 2
                            )

                            longitud_maxima = min(
                                len(texto),
                                longitud + 2
                            )

                            for tamano in range(
                                longitud_minima,
                                longitud_maxima + 1
                            ):
                                for inicio in range(
                                    0,
                                    len(texto) - tamano + 1
                                ):
                                    fragmento = texto[
                                        inicio:inicio + tamano
                                    ]

                                    puntaje = (
                                        SequenceMatcher(
                                            None,
                                            modelo_compacto,
                                            fragmento
                                        ).ratio() * 100
                                    )

                                    mejor_puntaje = max(
                                        mejor_puntaje,
                                        puntaje
                                    )

                            return mejor_puntaje

                        candidatos = []

                        for indice, fila in df.iterrows():

                            marca = obtener_valor(
                                fila,
                                ["Marca", "Fabricante"]
                            )

                            modelo = obtener_valor(
                                fila,
                                ["Modelo exacto", "Modelo"]
                            )

                            if not modelo:
                                continue

                            puntaje_modelo = comparar_modelo(
                                modelo,
                                texto_compacto
                            )

                            # Evitar resultados basados únicamente
                            # en una coincidencia de marca.
                            if puntaje_modelo >= 72:

                                coincidencias = [
                                    (
                                        "Modelo",
                                        modelo,
                                        round(puntaje_modelo, 1)
                                    )
                                ]

                                marca_normalizada = re.sub(
                                    r"[^A-Z0-9]",
                                    "",
                                    normalizar_ocr(marca).upper()
                                )

                                if (
                                    marca_normalizada
                                    and len(marca_normalizada) >= 3
                                    and marca_normalizada in texto_compacto
                                ):
                                    coincidencias.append(
                                        ("Marca", marca, 100.0)
                                    )

                                candidatos.append(
                                    (
                                        puntaje_modelo,
                                        marca,
                                        modelo,
                                        coincidencias,
                                        indice
                                    )
                                )

                        candidatos.sort(
                            key=lambda candidato: candidato[0],
                            reverse=True
                        )

                        if not candidatos:

                            st.warning(
                                "Se detectó texto, pero no se encontró "
                                "una coincidencia suficientemente clara "
                                "con los modelos de la base de datos. "
                                "Verifica el texto leído o prueba con "
                                "una fotografía más nítida."
                            )

                        else:

                            for numero, candidato in enumerate(
                                candidatos[:5],
                                start=1
                            ):

                                (
                                    puntaje,
                                    marca,
                                    modelo,
                                    coincidencias,
                                    indice
                                ) = candidato

                                with st.expander(
                                    f"{numero}. {marca} | {modelo}",
                                    expanded=(numero == 1)
                                ):

                                    st.write(
                                        f"**Puntuación orientativa del "
                                        f"modelo:** {puntaje:.1f}%"
                                    )

                                    st.write(f"**Marca:** {marca}")
                                    st.write(f"**Modelo:** {modelo}")

                                    for (
                                        nombre,
                                        valor,
                                        similitud
                                    ) in coincidencias:

                                        st.write(
                                            f"- **{nombre}:** {valor} "
                                            f"({similitud:.1f}%)"
                                        )

                                    if st.button(
                                        "Consultar ficha técnica",
                                        key=f"ficha_ocr_{indice}"
                                    ):
                                        st.session_state[
                                            "modelo_reconocido"
                                        ] = indice

                                        st.success(
                                            "Modelo seleccionado. "
                                            "Consulta la sección Ficha técnica."
                                        )

                            st.warning(
                                "La puntuación es una similitud orientativa, "
                                "no una probabilidad de identificación "
                                "correcta. Confirma siempre el código "
                                "en la etiqueta física."
                            )
                  
                except Exception as error:
                    st.error(
                        "No se pudo ejecutar el reconocimiento con EasyOCR."
                    )
                    st.exception(error)

        except Exception as error:
            st.error("No se pudo procesar la fotografía.")
            st.exception(error)


# ==========================================================
# COMUNICACIÓN INDUSTRIAL
# ==========================================================

elif seccion == "Comunicación industrial":

    st.subheader("🔌 Comunicación industrial")

    st.write(
        "Información introductoria sobre interfaces y redes "
        "utilizadas en automatización industrial."
    )

    with st.expander("RS-485"):

        st.write(
            "Interfaz eléctrica diferencial empleada para transmitir "
            "datos en ambientes industriales. Puede utilizarse con "
            "protocolos como Modbus RTU, según el equipo."
        )

        st.write(
            "**Verificar:** pinout, polaridad de las líneas, "
            "referencia eléctrica, velocidad y terminación cuando corresponda."
        )

    with st.expander("Ethernet industrial"):

        st.write(
            "Permite comunicación mediante redes basadas en Ethernet. "
            "La presencia de un conector RJ45 no garantiza que dos "
            "equipos utilicen el mismo protocolo."
        )

        st.write(
            "**Verificar:** protocolo compatible, parámetros de red, "
            "dirección IP y software del fabricante."
        )

    with st.expander("PROFIBUS"):

        st.write(
            "Red industrial utilizada para comunicar controladores "
            "y dispositivos de campo."
        )

        st.write(
            "**Verificar:** variante de PROFIBUS, conector, cableado, "
            "terminación y configuración de red."
        )

    with st.expander("USB"):

        st.write(
            "Puede utilizarse para programación, diagnóstico o "
            "configuración, dependiendo del modelo de PLC."
        )

        st.write(
            "**Verificar:** tipo de puerto, cable requerido, "
            "controladores y software de programación."
        )

    with st.expander("Pinout e interfaces"):

        st.write(
            "El pinout describe la asignación de señales de un conector. "
            "Debe verificarse en el manual del modelo exacto antes "
            "de realizar una conexión."
        )

        st.warning(
            "No selecciones un cable solamente por la forma del conector. "
            "Dos conectores físicamente iguales pueden tener asignaciones "
            "de pines diferentes."
        )


# ==========================================================
# PIE DE PÁGINA
# ==========================================================

st.divider()

st.markdown(
    """
    <div style="text-align:center; color:#AFC8E2; padding:10px;">
        PLC SCAN | Automatización Industrial
    </div>
    """,
    unsafe_allow_html=True,
)