
import streamlit as st
import pandas as pd
from pathlib import Path
import re
import unicodedata
from urllib.parse import urlparse
from PIL import Image
from functools import lru_cache

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="PLC SCAN | Automatización Industrial",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_EXCEL = BASE_DIR / "base_datos.xlsx.xlsx"
ARCHIVO_MODELO = BASE_DIR / "yolo26n.pt"

# ============================================================
# DISEÑO VISUAL
# ============================================================

st.markdown(
    """
    <style>
    .stApp {
        background:
            linear-gradient(
                135deg,
                rgba(5, 16, 39, 0.98),
                rgba(10, 38, 75, 0.96)
            );
        color: #F2F6FC;
    }

    [data-testid="stSidebar"] {
        background: #071426;
        border-right: 1px solid #1C426B;
    }

    [data-testid="stSidebar"] * {
        color: #F2F6FC;
    }

    h1, h2, h3 {
        color: #66C7FF !important;
    }

    p, label, li {
        color: #E4ECF7;
    }

    div[data-testid="stMetric"] {
        background: rgba(17, 49, 82, 0.75);
        border: 1px solid #28618D;
        padding: 15px;
        border-radius: 12px;
    }

    div.stButton > button {
        background-color: #087FC1;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: bold;
        padding: 0.55rem 1rem;
    }

    div.stButton > button:hover {
        background-color: #119EEB;
        color: white;
        border: 1px solid #7ED2FF;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #28577E;
        border-radius: 10px;
    }

    .djh-watermark {
        position: fixed;
        right: 25px;
        bottom: 12px;
        font-size: 45px;
        font-weight: 900;
        color: rgba(120, 190, 255, 0.10);
        letter-spacing: 5px;
        z-index: 0;
        pointer-events: none;
    }

    .plc-header {
        padding: 22px;
        border: 1px solid #245581;
        border-radius: 15px;
        background: linear-gradient(
            120deg,
            rgba(9, 35, 65, 0.95),
            rgba(14, 69, 109, 0.80)
        );
        margin-bottom: 20px;
    }

    .plc-subtitle {
        color: #B8D9F2;
        font-size: 16px;
    }

    .info-card {
        background: rgba(13, 42, 72, 0.82);
        border: 1px solid #245581;
        border-radius: 10px;
        padding: 14px;
        margin-bottom: 10px;
    }
    </style>

    <div class="djh-watermark">DJH</div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# FUNCIONES GENERALES
# ============================================================

def normalizar(valor):
    """Normaliza texto para comparar encabezados y búsquedas."""
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


def valor_texto(valor):
    """Convierte los valores vacíos de Excel en texto legible."""
    if pd.isna(valor):
        return "No disponible"

    texto = str(valor).strip()

    if not texto or texto.lower() in ("nan", "none", "null"):
        return "No disponible"

    return texto


def buscar_columna(df, posibles_nombres):
    """Encuentra una columna aunque tenga variaciones en su nombre."""
    columnas_normalizadas = {
        normalizar(columna): columna
        for columna in df.columns
    }

    for nombre in posibles_nombres:
        nombre_normalizado = normalizar(nombre)

        if nombre_normalizado in columnas_normalizadas:
            return columnas_normalizadas[nombre_normalizado]

    for nombre in posibles_nombres:
        nombre_normalizado = normalizar(nombre)

        for columna_normalizada, columna_original in (
            columnas_normalizadas.items()
        ):
            if (
                nombre_normalizado
                and nombre_normalizado in columna_normalizada
            ):
                return columna_original

    return None


def obtener_valor(fila, posibles_nombres):
    """Obtiene un dato de una fila usando distintos nombres posibles."""
    columnas_normalizadas = {
        normalizar(columna): columna
        for columna in fila.index
    }

    for nombre in posibles_nombres:
        clave = normalizar(nombre)

        if clave in columnas_normalizadas:
            return valor_texto(
                fila[columnas_normalizadas[clave]]
            )

    for nombre in posibles_nombres:
        clave = normalizar(nombre)

        for columna_normalizada, columna_original in (
            columnas_normalizadas.items()
        ):
            if clave and clave in columna_normalizada:
                return valor_texto(fila[columna_original])

    return "No disponible"


def es_url(valor):
    """Comprueba si un texto contiene una dirección web válida."""
    if not valor or valor == "No disponible":
        return False

    try:
        resultado = urlparse(valor)
        return resultado.scheme in ("http", "https") and bool(
            resultado.netloc
        )
    except Exception:
        return False


def mostrar_enlace_manual(valor, texto="Abrir recurso"):
    """Muestra un enlace si el dato de Excel es una URL."""
    if valor == "No disponible":
        st.caption("No hay un recurso registrado.")
    elif es_url(valor):
        st.markdown(f"[{texto}]({valor})")
    else:
        st.write(valor)


def buscar_coincidencias(df, consulta):
    """Busca por marca, modelo, familia, serie y otros datos."""
    consulta_normalizada = normalizar(consulta)

    if not consulta_normalizada:
        return df.copy()

    columnas_busqueda = [
        columna
        for columna in [
            buscar_columna(df, ["Marca", "Fabricante"]),
            buscar_columna(df, ["Modelo exacto", "Modelo"]),
            buscar_columna(df, ["Familia", "Familia PLC"]),
            buscar_columna(
                df,
                ["Serie / Subfamilia", "Serie", "Subfamilia"],
            ),
            buscar_columna(
                df,
                ["Clasificación", "Tipo de PLC", "Tipo"],
            ),
            buscar_columna(
                df,
                ["Arquitectura", "Interfaz", "Comunicación"],
            ),
        ]
        if columna is not None
    ]

    if not columnas_busqueda:
        return df.iloc[0:0].copy()

    mascara = pd.Series(False, index=df.index)

    for columna in columnas_busqueda:
        mascara |= (
            df[columna]
            .fillna("")
            .astype(str)
            .map(normalizar)
            .str.contains(
                re.escape(consulta_normalizada),
                regex=True,
                na=False,
            )
        )

    return df[mascara].copy()


# ============================================================
# CARGA DE BASE DE DATOS
# ============================================================

@st.cache_data
def cargar_base_datos(ruta_archivo, fecha_modificacion):
    """Carga el archivo Excel y conserva sus columnas originales."""
    return pd.read_excel(ruta_archivo)


if not ARCHIVO_EXCEL.exists():
    st.error(
        "No se encontró el archivo base_datos.xlsx. "
        "Colócalo junto al archivo PLC SCAN.py."
    )
    st.stop()

try:
    df = cargar_base_datos(
        str(ARCHIVO_EXCEL),
        ARCHIVO_EXCEL.stat().st_mtime,
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


# ============================================================
# ENCABEZADO
# ============================================================

st.markdown(
    """
    <div class="plc-header">
        <h1>⚙️ PLC SCAN</h1>
        <div class="plc-subtitle">
            Identificación técnica de controladores lógicos
            programables y consulta de interfaces de comunicación.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# MENÚ LATERAL
# ============================================================

st.sidebar.title("⚙️ PLC SCAN")
st.sidebar.caption("Automatización Industrial")

seccion = st.sidebar.radio(
    "Selecciona una sección:",
    [
        "Inicio",
        "Base de datos",
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

# ============================================================
# INICIO
# ============================================================

if seccion == "Inicio":

    st.subheader("Panel principal")

    st.write(
        "PLC SCAN es una herramienta de consulta técnica diseñada "
        "para facilitar la identificación de controladores lógicos "
        "programables y la revisión de sus características."
    )

    total_registros = len(df)

    columna_marca = buscar_columna(
        df,
        ["Marca", "Fabricante"],
    )

    if columna_marca:
        total_marcas = (
            df[columna_marca]
            .dropna()
            .astype(str)
            .str.strip()
            .nunique()
        )
    else:
        total_marcas = 0

    columna_modelo = buscar_columna(
        df,
        ["Modelo exacto", "Modelo"],
    )

    if columna_modelo:
        total_modelos = (
            df[columna_modelo]
            .dropna()
            .astype(str)
            .str.strip()
            .nunique()
        )
    else:
        total_modelos = 0

    col1, col2, col3 = st.columns(3)

    col1.metric("Registros disponibles", total_registros)
    col2.metric("Marcas registradas", total_marcas)
    col3.metric("Modelos registrados", total_modelos)

    st.markdown("### Funciones disponibles")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(
            """
            <div class="info-card">
                <h4>🔎 Búsqueda de PLC</h4>
                Consulta los equipos registrados por marca,
                modelo, familia o serie.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <div class="info-card">
                <h4>📋 Ficha técnica</h4>
                Consulta especificaciones, pinout,
                software, cables y manuales disponibles.
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            """
            <div class="info-card">
                <h4>🧠 Reconocimiento de imágenes</h4>
                Prueba el análisis visual de fotografías
                de equipos industriales.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            """
            <div class="info-card">
                <h4>🔌 Comunicación industrial</h4>
                Consulta información introductoria sobre
                las principales interfaces y redes industriales.
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# BASE DE DATOS
# ============================================================

elif seccion == "Base de datos":

    st.subheader("📚 Base de datos de PLC")

    st.write(
        "Registros técnicos disponibles en el archivo Excel."
    )

    st.caption(
        f"Total de registros: {len(df)} | "
        f"Total de columnas: {len(df.columns)}"
    )

    busqueda_tabla = st.text_input(
        "Filtrar registros:",
        placeholder="Escribe una marca, modelo o característica...",
        key="filtro_base_datos",
    )

    if busqueda_tabla.strip():
        tabla = buscar_coincidencias(df, busqueda_tabla)
    else:
        tabla = df.copy()

    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        label="Descargar resultados en CSV",
        data=tabla.to_csv(index=False).encode("utf-8-sig"),
        file_name="PLC_SCAN_resultados.csv",
        mime="text/csv",
    )


# ============================================================
# BUSCAR PLC
# ============================================================

elif seccion == "Buscar PLC":

    st.subheader("🔎 Buscar controlador")

    consulta = st.text_input(
        "Ingresa la marca, modelo, familia o serie:",
        placeholder="Ejemplo: Siemens, Allen-Bradley, KV-N60...",
    )

    if consulta.strip():

        resultados_busqueda = buscar_coincidencias(
            df,
            consulta,
        )

        if resultados_busqueda.empty:
            st.warning(
                "No se encontraron coincidencias en la base de datos."
            )
        else:
            st.success(
                f"Se encontraron {len(resultados_busqueda)} "
                "coincidencia(s)."
            )

            for indice, (_, fila) in enumerate(
                resultados_busqueda.iterrows()
            ):
                marca = obtener_valor(
                    fila,
                    ["Marca", "Fabricante"],
                )

                modelo = obtener_valor(
                    fila,
                    ["Modelo exacto", "Modelo"],
                )

                familia = obtener_valor(
                    fila,
                    ["Familia", "Familia PLC"],
                )

                with st.expander(
                    f"{marca} | {modelo}",
                    expanded=(indice == 0),
                ):
                    st.write(f"**Marca:** {marca}")
                    st.write(f"**Modelo:** {modelo}")
                    st.write(f"**Familia:** {familia}")

                    st.write(
                        "**Tipo de PLC:** "
                        + obtener_valor(
                            fila,
                            [
                                "Clasificación",
                                "Tipo de PLC",
                                "Tipo",
                            ],
                        )
                    )

                    st.write(
                        "**Serie:** "
                        + obtener_valor(
                            fila,
                            [
                                "Serie / Subfamilia",
                                "Serie",
                                "Subfamilia",
                            ],
                        )
                    )

                    st.write(
                        "**Pinout:** "
                        + obtener_valor(
                            fila,
                            [
                                "Pin Out",
                                "Pinout",
                                "Pin Out / Pinout",
                            ],
                        )
                    )

                    st.write(
                        "**Interfaz:** "
                        + obtener_valor(
                            fila,
                            [
                                "Interfaz",
                                "Interfaces",
                                "Interfaz de comunicación",
                            ],
                        )
                    )

                    st.write(
                        "**Cable:** "
                        + obtener_valor(
                            fila,
                            [
                                "Cable",
                                "Cable requerido",
                                "Cable de programación",
                            ],
                        )
                    )

                    st.write(
                        "**Software:** "
                        + obtener_valor(
                            fila,
                            [
                                "Software",
                                "Software de programación",
                            ],
                        )
                    )

                    manual = obtener_valor(
                        fila,
                        ["Manual", "Manual PDF", "Manual técnico"],
                    )

                    st.write("**Manual:**")
                    mostrar_enlace_manual(
                        manual,
                        "Abrir manual técnico",
                    )


# ============================================================
# FICHA TÉCNICA
# ============================================================

elif seccion == "Ficha técnica":

    st.subheader("📋 Ficha técnica del PLC")

    if df.empty:
        st.warning("La base de datos no contiene registros.")
    else:
        columna_marca = buscar_columna(
            df,
            ["Marca", "Fabricante"],
        )

        columna_modelo = buscar_columna(
            df,
            ["Modelo exacto", "Modelo"],
        )

        if columna_marca and columna_modelo:

            opciones = []

            for indice, fila in df.iterrows():
                marca = valor_texto(fila[columna_marca])
                modelo = valor_texto(fila[columna_modelo])

                opciones.append(
                    (f"{marca} | {modelo}", indice)
                )

            etiquetas = [
                etiqueta
                for etiqueta, _ in opciones
            ]

            seleccion = st.selectbox(
                "Selecciona el equipo:",
                etiquetas,
            )

            indice_seleccionado = dict(opciones)[seleccion]
            fila = df.loc[indice_seleccionado]

            marca = obtener_valor(
                fila,
                ["Marca", "Fabricante"],
            )

            modelo = obtener_valor(
                fila,
                ["Modelo exacto", "Modelo"],
            )

            st.markdown(
                f"### {marca} — {modelo}"
            )

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### Identificación")

                st.write(
                    "**Marca:** "
                    + marca
                )

                st.write(
                    "**Modelo:** "
                    + modelo
                )

                st.write(
                    "**Familia:** "
                    + obtener_valor(
                        fila,
                        ["Familia", "Familia PLC"],
                    )
                )

                st.write(
                    "**Serie / Subfamilia:** "
                    + obtener_valor(
                        fila,
                        [
                            "Serie / Subfamilia",
                            "Serie",
                            "Subfamilia",
                        ],
                    )
                )

                st.write(
                    "**Clasificación:** "
                    + obtener_valor(
                        fila,
                        [
                            "Clasificación",
                            "Tipo de PLC",
                            "Tipo",
                        ],
                    )
                )

                st.write(
                    "**Arquitectura:** "
                    + obtener_valor(
                        fila,
                        ["Arquitectura"],
                    )
                )

            with col2:
                st.markdown("#### Conexión y programación")

                st.write(
                    "**Pinout:** "
                    + obtener_valor(
                        fila,
                        [
                            "Pin Out",
                            "Pinout",
                            "Pin Out / Pinout",
                        ],
                    )
                )

                st.write(
                    "**Interfaz:** "
                    + obtener_valor(
                        fila,
                        [
                            "Interfaz",
                            "Interfaces",
                            "Interfaz de comunicación",
                        ],
                    )
                )

                st.write(
                    "**Cable requerido:** "
                    + obtener_valor(
                        fila,
                        [
                            "Cable",
                            "Cable requerido",
                            "Cable de programación",
                        ],
                    )
                )

                st.write(
                    "**Software:** "
                    + obtener_valor(
                        fila,
                        [
                            "Software",
                            "Software de programación",
                        ],
                    )
                )

                st.write(
                    "**Comunicación:** "
                    + obtener_valor(
                        fila,
                        [
                            "Comunicación",
                            "Protocolos",
                            "Red industrial",
                            "Interfaces de comunicación",
                        ],
                    )
                )

            st.divider()

            st.markdown("#### Manual y documentación")

            manual = obtener_valor(
                fila,
                [
                    "Manual",
                    "Manual PDF",
                    "Manual técnico",
                    "Link manual",
                ],
            )

            mostrar_enlace_manual(
                manual,
                "📘 Abrir manual del fabricante",
            )

            st.markdown("#### Información adicional")

            notas = obtener_valor(
                fila,
                [
                    "Notas",
                    "Observaciones",
                    "Información adicional",
                ],
            )

            st.write(notas)

            with st.expander(
                "Ver todos los datos originales del registro"
            ):
                st.dataframe(
                    fila.to_frame(name="Valor"),
                    use_container_width=True,
                )

        else:
            st.error(
                "No se identificaron las columnas de marca y modelo. "
                "Revisa los encabezados del archivo Excel."
            )


# ============================================================
# RECONOCIMIENTO DEL PLC
# ============================================================


elif seccion == "Reconocimiento del PLC":

    import re
    import unicodedata
    from difflib import SequenceMatcher
    from PIL import ImageEnhance, ImageFilter, ImageOps

    st.subheader("🧠 Reconocimiento visual del PLC")

    st.write(
        "Carga una fotografía del equipo para mejorar la imagen, "
        "leer el texto de la etiqueta y buscar posibles coincidencias "
        "en la base de datos de PLC SCAN."
    )

    st.info(
        "La identificación se basa en el texto visible de la etiqueta. "
        "Si la fotografía está demasiado borrosa, puede ser necesario "
        "tomar otra imagen o confirmar manualmente el modelo."
    )

    archivo_imagen = st.file_uploader(
        "Selecciona una fotografía del PLC:",
        type=["jpg", "jpeg", "png", "webp"],
        key="imagen_plc",
    )

    def normalizar_ocr(texto):
        texto = str(texto).lower().strip()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(
            caracter for caracter in texto
            if not unicodedata.combining(caracter)
        )
        return re.sub(r"[^a-z0-9]", "", texto)

    if archivo_imagen is not None:

        imagen = Image.open(archivo_imagen).convert("RGB")

        st.image(
            imagen,
            caption="Fotografía original",
            use_container_width=True,
        )

        if st.button(
            "🔍 Mejorar imagen y buscar PLC",
            key="analizar_plc",
        ):

            with st.spinner(
                "Mejorando la fotografía y leyendo la etiqueta..."
            ):

                try:
                    import pytesseract

                    # Aumentar el tamaño de la imagen
                    ancho, alto = imagen.size
                    factor = max(2, min(4, 1600 // max(ancho, 1)))

                    imagen_mejorada = imagen.resize(
                        (ancho * factor, alto * factor)
                    )

                    # Mejorar contraste y nitidez
                    imagen_mejorada = ImageOps.autocontrast(
                        imagen_mejorada
                    )

                    imagen_mejorada = ImageEnhance.Contrast(
                        imagen_mejorada
                    ).enhance(1.6)

                    imagen_mejorada = ImageEnhance.Sharpness(
                        imagen_mejorada
                    ).enhance(2.0)

                    imagen_mejorada = imagen_mejorada.filter(
                        ImageFilter.UnsharpMask(
                            radius=2,
                            percent=150,
                            threshold=3,
                        )
                    )

                    st.image(
                        imagen_mejorada,
                        caption="Fotografía mejorada para lectura",
                        use_container_width=True,
                    )

                    # Leer el texto de la etiqueta
                    texto_detectado = pytesseract.image_to_string(
                        imagen_mejorada,
                        config="--psm 11",
                    )

                    st.subheader("📝 Texto leído de la etiqueta")

                    if texto_detectado.strip():
                        st.text_area(
                            "Revisa el texto detectado:",
                            value=texto_detectado,
                            height=150,
                            key="texto_ocr_plc",
                        )
                    else:
                        st.warning(
                            "No se pudo leer texto suficiente. "
                            "Intenta tomar una fotografía más cercana, "
                            "con buena iluminación y enfocada en la etiqueta."
                        )

                    # Buscar coincidencias en el Excel
                    if texto_detectado.strip():

                        texto_ocr = normalizar_ocr(
                            texto_detectado
                        )

                        if "df" not in globals() or df.empty:
                            st.error(
                                "No se encontró una base de datos "
                                "disponible para realizar la búsqueda."
                            )

                        else:
                            candidatos = []

                            for _, fila in df.iterrows():

                                valores = {
                                    str(columna): str(valor)
                                    for columna, valor in fila.items()
                                    if valor is not None
                                    and str(valor).strip().lower()
                                    not in ("", "nan", "none")
                                }

                                # Buscar coincidencias por celda
                                mejor_puntaje = 0
                                coincidencias = []

                                for columna, valor in valores.items():

                                    valor_norm = normalizar_ocr(valor)

                                    if len(valor_norm) < 3:
                                        continue

                                    if valor_norm in texto_ocr:
                                        puntaje = 100
                                    else:
                                        palabras = re.findall(
                                            r"[A-Za-z0-9][A-Za-z0-9._/-]*",
                                            texto_detectado,
                                        )

                                        palabras_norm = [
                                            normalizar_ocr(p)
                                            for p in palabras
                                        ]

                                        puntaje = 0

                                        for palabra in palabras_norm:
                                            if len(palabra) < 3:
                                                continue

                                            similitud = (
                                                SequenceMatcher(
                                                    None,
                                                    valor_norm,
                                                    palabra,
                                                ).ratio() * 100
                                            )

                                            puntaje = max(
                                                puntaje,
                                                similitud,
                                            )

                                    if puntaje >= 75:
                                        coincidencias.append(
                                            (columna, valor, puntaje)
                                        )

                                    mejor_puntaje = max(
                                        mejor_puntaje,
                                        puntaje,
                                    )

                                if mejor_puntaje >= 65:
                                    candidatos.append(
                                        (
                                            mejor_puntaje,
                                            valores,
                                            coincidencias,
                                        )
                                    )

                            candidatos.sort(
                                key=lambda elemento: elemento[0],
                                reverse=True,
                            )

                            st.subheader(
                                "📋 Posibles coincidencias en la base de datos"
                            )

                            if candidatos:

                                for numero, (
                                    puntaje,
                                    valores,
                                    coincidencias,
                                ) in enumerate(candidatos[:5], start=1):

                                    with st.expander(
                                        f"Posible coincidencia {numero}",
                                        expanded=(numero == 1),
                                    ):

                                        # Mostrar los campos principales
                                        for columna, valor in valores.items():
                                            encabezado = normalizar_ocr(
                                                columna
                                            )

                                            if any(
                                                palabra in encabezado
                                                for palabra in [
                                                    "marca",
                                                    "fabricante",
                                                    "modelo",
                                                    "nombre",
                                                    "tipo",
                                                ]
                                            ):
                                                st.write(
                                                    f"**{columna}:** {valor}"
                                                )

                                        if coincidencias:
                                            st.write(
                                                "**Datos que coincidieron "
                                                "con la etiqueta:**"
                                            )

                                            for columna, valor, similitud in (
                                                coincidencias[:5]
                                            ):
                                                st.write(
                                                    f"- {columna}: {valor}"
                                                )

                                st.warning(
                                    "Estas son coincidencias preliminares, "
                                    "no una identificación confirmada. "
                                    "Verifica el modelo en la etiqueta física."
                                )

                            else:
                                st.warning(
                                    "No se encontraron coincidencias claras "
                                    "en el Excel. Revisa el texto detectado "
                                    "o intenta con una fotografía más nítida."
                                )

                except Exception as error:
                    st.error(
                        "No se pudo completar el reconocimiento."
                    )
                    st.code(str(error))
                    st.info(
                        "Comprueba que las dependencias de OCR estén "
                        "instaladas en la aplicación publicada."
                    )


# ============================================================
# COMUNICACIÓN INDUSTRIAL
# ============================================================

elif seccion == "Comunicación industrial":

    st.subheader("🔌 Comunicación industrial")

    st.write(
        "Consulta información introductoria sobre interfaces "
        "y redes utilizadas en automatización industrial."
    )

    with st.expander("RS-485"):
        st.write(
            "Interfaz eléctrica diferencial utilizada para "
            "transmisión de datos en entornos industriales. "
            "Puede emplearse con protocolos como Modbus RTU, "
            "según el equipo y su configuración."
        )

        st.write(
            "**Consideraciones:** verificar el pinout, la "
            "polaridad de las líneas, la referencia eléctrica, "
            "la velocidad de comunicación y la terminación "
            "cuando corresponda."
        )

    with st.expander("Ethernet industrial"):
        st.write(
            "Permite la comunicación mediante redes basadas "
            "en Ethernet. El conector físico por sí solo no "
            "determina el protocolo compatible."
        )

        st.write(
            "**Consideraciones:** comprobar el protocolo, "
            "la dirección IP, los parámetros de red y el "
            "software requerido por el fabricante."
        )

    with st.expander("PROFIBUS"):
        st.write(
            "Red de comunicación industrial utilizada para "
            "intercambio de datos entre controladores y "
            "dispositivos de campo."
        )

        st.write(
            "**Consideraciones:** identificar la variante "
            "de PROFIBUS, el conector, el cableado y la "
            "configuración de red existente."
        )

    with st.expander("USB"):
        st.write(
            "Puede utilizarse para programación, diagnóstico "
            "o configuración, dependiendo del modelo de PLC."
        )

        st.write(
            "**Consideraciones:** confirmar el tipo de puerto, "
            "el cable o interfaz requerido y los controladores "
            "de software necesarios."
        )

    with st.expander("Pinout e interfaces"):
        st.write(
            "El pinout describe la asignación de señales en "
            "un conector. Debe verificarse en el manual del "
            "modelo específico antes de conectar un equipo."
        )

        st.warning(
            "No selecciones un cable únicamente por la forma "
            "del conector. Dos conectores físicamente iguales "
            "pueden tener asignaciones de pines diferentes."
        )


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.divider()

st.markdown(
    """
    <div style="text-align:center; color:#AFC8E2; padding:10px;">
        PLC SCAN | Automatización Industrial | DJH
    </div>
    """,
    unsafe_allow_html=True,
)


