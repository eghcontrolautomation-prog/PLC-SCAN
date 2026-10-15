
import streamlit as st
import pandas as pd
from pathlib import Path
import re
import unicodedata
from functools import lru_cache
from ultralytics import YOLO
from PIL import Image

# ==========================================
# CONFIGURACION DE PLC SCAN
# ==========================================

st.set_page_config(
    page_title="PLC SCAN",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = Path(__file__).parent
ARCHIVO_EXCEL = BASE_DIR / "base_datos.xlsx"
ARCHIVO_MODELO = BASE_DIR / "yolo26n.pt"


# ==========================================
# CARGA DE BASE DE DATOS
# ==========================================

@st.cache_data
def cargar_base_datos(ruta, fecha_modificacion):
    return pd.read_excel(ruta)


# ==========================================
# CARGA DEL MODELO YOLO
# ==========================================

@st.cache_resource
def cargar_modelo_yolo():
    return YOLO(str(ARCHIVO_MODELO))


# ==========================================
# FUNCIONES AUXILIARES
# ==========================================

def normalizar(texto):
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )
    return "".join(
        caracter for caracter in texto
        if caracter.isalnum()
    )


def obtener_urls(texto):
    if not texto:
        return []

    return re.findall(
        r"https?://[^\s<>\"']+",
        str(texto)
    )


def buscar_columna(columnas, *nombres):
    columnas_normalizadas = {
        normalizar(columna): columna
        for columna in columnas
    }

    for nombre in nombres:
        clave = normalizar(nombre)

        if clave in columnas_normalizadas:
            return columnas_normalizadas[clave]

    return None


def obtener_valor(fila, *nombres):
    columna = buscar_columna(fila.index, *nombres)

    if columna is None:
        return None

    valor = fila[columna]

    if pd.isna(valor):
        return None

    valor = str(valor).strip()

    return valor if valor else None


def mostrar_enlace_manual(valor):
    if not valor:
        st.info("No hay un manual registrado para este PLC.")
        return

    urls = obtener_urls(valor)

    if urls:
        for numero, url in enumerate(dict.fromkeys(urls), start=1):
            st.link_button(
                f"Abrir manual o documentación {numero}",
                url
            )
    else:
        st.write(valor)


def buscar_coincidencias(datos, texto):
    if datos is None or datos.empty or not texto.strip():
        return datos.iloc[0:0] if datos is not None else pd.DataFrame()

    coincidencias = datos.astype(str).apply(
        lambda columna: columna.str.contains(
            texto.strip(),
            case=False,
            na=False,
            regex=False
        )
    ).any(axis=1)

    return datos.loc[coincidencias]


# ==========================================
# ESTILO VISUAL Y MARCA DE AGUA DJH
# ==========================================

st.markdown("""
<style>
.stApp {
    background: linear-gradient(
        135deg,
        #07111F,
        #12304B,
        #07111F
    );
    color: #F1F5F9;
}

[data-testid="stHeader"] {
    background: rgba(5, 16, 31, 0.85);
}

[data-testid="stSidebar"] {
    background: #071425;
    border-right: 1px solid #1E4970;
}

h1, h2, h3 {
    color: #55C5FF !important;
}

p, label, [data-testid="stMarkdownContainer"] {
    color: #E2EAF3;
}

div[data-baseweb="input"] input {
    background-color: #10243A;
    color: white;
    border: 1px solid #2C587B;
}

.stButton > button {
    background: #0879D9;
    color: white;
    border: 1px solid #36A9FF;
    border-radius: 8px;
    font-weight: 600;
}

.stButton > button:hover {
    background: #1598FF;
    color: white;
    border: 1px solid #8DD4FF;
}

[data-testid="stMetric"],
[data-testid="stExpander"] {
    background: rgba(12, 34, 56, 0.90);
    border: 1px solid #254967;
    border-radius: 12px;
}

hr {
    border-color: #254967;
}

.stApp::before {
    content: "DJH";
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    font-size: clamp(100px, 22vw, 320px);
    font-weight: 900;
    font-family: Arial, Helvetica, sans-serif;
    letter-spacing: 18px;
    color: rgba(100, 190, 255, 0.055);
    pointer-events: none;
    z-index: 0;
    white-space: nowrap;
}

[data-testid="stAppViewContainer"] {
    position: relative;
}
</style>
""", unsafe_allow_html=True)


# ==========================================
# ENCABEZADO PRINCIPAL
# ==========================================

st.title("⚙️ PLC SCAN")

st.subheader("Identificación y consulta técnica de PLC")

st.write(
    "Sistema de consulta de controladores lógicos "
    "programables para automatización industrial."
)

st.divider()


# ==========================================
# MENU LATERAL
# ==========================================

st.sidebar.title("PLC SCAN")
st.sidebar.caption("Panel de navegación")

seccion = st.sidebar.radio(
    "Selecciona una sección:",
    [
        "Inicio",
        "Base de datos",
        "Buscar PLC",
        "Ficha técnica",
        "Reconocimiento del PLC",
        "Comunicación industrial"
    ]
)


# ==========================================
# CARGAR EXCEL
# ==========================================

datos = None
error_excel = None

if ARCHIVO_EXCEL.exists():
    try:
        fecha_modificacion = ARCHIVO_EXCEL.stat().st_mtime

        datos = cargar_base_datos(
            str(ARCHIVO_EXCEL),
            fecha_modificacion
        )

        datos = datos.dropna(how="all").dropna(
            axis=1,
            how="all"
        )

        datos.columns = [
            str(columna).strip()
            for columna in datos.columns
        ]

    except Exception as error:
        error_excel = str(error)


# ==========================================
# INICIO
# ==========================================

if seccion == "Inicio":

    st.header("Panel principal")

    col1, col2 = st.columns(2)

    with col1:
        st.metric("Estado de la aplicación", "Activa")

    with col2:
        cantidad = len(datos) if datos is not None else 0
        st.metric("PLC registrados", cantidad)

    st.info(
        "Utiliza el menú lateral para consultar los PLC, "
        "buscar modelos o revisar la información disponible."
    )

    st.markdown("### Funciones del sistema")

    st.markdown("""
    - Consulta de PLC registrados en Excel.
    - Búsqueda por fabricante, modelo o características.
    - Visualización de fichas técnicas.
    - Consulta de interfaces y protocolos industriales.
    - Pruebas de reconocimiento de imágenes.
    """)


# ==========================================
# BASE DE DATOS
# ==========================================

elif seccion == "Base de datos":

    st.header("Base de datos de PLC")

    if datos is not None:

        st.success("Base de datos cargada correctamente.")

        col1, col2 = st.columns(2)

        with col1:
            st.metric("Registros", len(datos))

        with col2:
            st.metric("Columnas", len(datos.columns))

        st.dataframe(
            datos,
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "Descargar copia de los datos como CSV",
            data=datos.to_csv(index=False).encode("utf-8-sig"),
            file_name="base_datos_plc.csv",
            mime="text/csv"
        )

    elif error_excel:
        st.error("No fue posible leer el archivo Excel.")
        st.code(error_excel)

    else:
        st.warning(
            "No se encontró base_datos.xlsx. "
            "Coloca el archivo en la carpeta del proyecto."
        )


# ==========================================
# BUSCAR PLC
# ==========================================

elif seccion == "Buscar PLC":

    st.header("Buscar un PLC")

    if datos is not None and not datos.empty:

        busqueda = st.text_input(
            "Fabricante, modelo o característica",
            placeholder="Ejemplo: Siemens, S7-1200, Allen-Bradley"
        )

        if busqueda.strip():
            resultados_busqueda = buscar_coincidencias(
                datos,
                busqueda
            )
        else:
            resultados_busqueda = datos

        st.write(
            f"Resultados encontrados: {len(resultados_busqueda)}"
        )

        st.dataframe(
            resultados_busqueda,
            use_container_width=True,
            hide_index=True
        )

    elif error_excel:
        st.error("Error al leer el Excel.")
        st.code(error_excel)

    else:
        st.warning(
            "Agrega tu archivo base_datos.xlsx "
            "para comenzar a buscar PLC."
        )


# ==========================================
# FICHA TECNICA DEL PLC
# ==========================================

elif seccion == "Ficha técnica":

    st.header("Ficha técnica del PLC")

    st.write(
        "Consulta las características, conexiones, "
        "documentación y datos de identificación del equipo."
    )

    if datos is not None and not datos.empty:

        def etiqueta_fila(indice):
            fila = datos.loc[indice]

            marca = obtener_valor(fila, "Marca") or ""

            modelo = obtener_valor(
                fila,
                "Modelo exacto",
                "Modelo",
                "Nombre del PLC"
            ) or ""

            identificador = obtener_valor(fila, "ID") or ""

            partes = [
                parte for parte in
                [identificador, marca, modelo]
                if parte
            ]

            return " | ".join(partes) or f"Registro {indice}"

        registros = datos.index.tolist()

        seleccionado = st.selectbox(
            "Selecciona el PLC:",
            registros,
            format_func=etiqueta_fila,
            key="selector_ficha_plc"
        )

        ficha = datos.loc[seleccionado]

        marca = obtener_valor(ficha, "Marca", "Fabricante")
        modelo = obtener_valor(
            ficha,
            "Modelo exacto",
            "Modelo",
            "Nombre del PLC"
        )
        familia = obtener_valor(ficha, "Familia")
        serie = obtener_valor(
            ficha,
            "Serie / Subfamilia",
            "Serie",
            "Subfamilia"
        )
        clasificacion = obtener_valor(
            ficha,
            "Clasificación",
            "Tipo de PLC"
        )
        arquitectura = obtener_valor(ficha, "Arquitectura")

        st.divider()
        st.subheader(modelo or "Modelo no especificado")

        if marca:
            st.caption(f"Fabricante: {marca}")

        metricas = [
            ("Marca", marca),
            ("Familia", familia),
            ("Clasificación", clasificacion),
            ("Arquitectura", arquitectura)
        ]

        metricas_disponibles = [
            (titulo, valor)
            for titulo, valor in metricas
            if valor
        ]

        if metricas_disponibles:
            columnas = st.columns(
                min(len(metricas_disponibles), 4)
            )

            for i, (titulo, valor) in enumerate(
                metricas_disponibles
            ):
                with columnas[i]:
                    st.metric(titulo, valor)

        if serie:
            st.write(f"**Serie / subfamilia:** {serie}")

        st.divider()
        st.markdown("### 🔌 Conexión e interfaz")

        campos_conexion = [
            ("PIN OUT", ("PIN OUT", "Pinout", "Pin Out")),
            (
                "Interfaz del pinout",
                (
                    "INTERFAZ / CONECTOR DEL PIN OUT",
                    "Interfaz / conector del pinout"
                )
            ),
            ("Conector del pinout", ("CONECTOR DEL PIN OUT",)),
            ("Estado del pinout", ("ESTADO DEL PIN OUT",)),
            (
                "Interfaz visible",
                ("Interfaz / conector visible",)
            ),
            ("Cable o adaptador", ("Cable / adaptador",))
        ]

        for etiqueta, alias in campos_conexion:
            valor = obtener_valor(ficha, *alias)

            if valor:
                with st.container(border=True):
                    st.markdown(f"**{etiqueta}**")
                    st.write(valor)

        st.markdown("### 💻 Software y comunicación")

        campos_comunicacion = [
            columna for columna in datos.columns
            if any(
                palabra in normalizar(columna)
                for palabra in [
                    "software",
                    "comunicacion",
                    "protocolo",
                    "interfaz",
                    "puerto",
                    "red"
                ]
            )
        ]

        if campos_comunicacion:
            for columna in campos_comunicacion:
                valor = ficha[columna]

                if pd.notna(valor) and str(valor).strip():
                    st.markdown(f"**{columna}:**")
                    st.write(str(valor))
        else:
            st.info(
                "No hay campos de software o comunicación "
                "identificables en este registro."
            )

        st.markdown("### 📚 Manuales y documentación")

        manual = obtener_valor(
            ficha,
            "URL MANUAL / DOCUMENTACIÓN",
            "Manual",
            "Documentación",
            "URL Manual",
            "Enlace del manual"
        )

        mostrar_enlace_manual(manual)

        st.markdown("### 📷 Foto de referencia")

        foto = obtener_valor(
            ficha,
            "Foto de referencia",
            "Fotografía",
            "Foto"
        )

        if foto:
            urls_foto = obtener_urls(foto)

            if urls_foto:
                for numero, url in enumerate(
                    dict.fromkeys(urls_foto),
                    start=1
                ):
                    st.image(
                        url,
                        caption=f"Referencia {numero}",
                        width=350
                    )
            else:
                st.write(foto)
        else:
            st.info(
                "No hay fotografía de referencia registrada."
            )

    elif error_excel:
        st.error("No fue posible cargar la base de datos.")
        st.code(error_excel)

    else:
        st.warning(
            "Coloca base_datos.xlsx en la carpeta del proyecto."
        )


# ==========================================
# RECONOCIMIENTO VISUAL DEL PLC
# ==========================================

elif seccion == "Reconocimiento del PLC":

    st.header("🔎 Reconocimiento visual del PLC")

    st.write(
        "Carga una fotografía del PLC para visualizarla "
        "y realizar una prueba de análisis mediante YOLO. "
        "La identificación del modelo exacto todavía requiere "
        "un sistema de reconocimiento especializado."
    )

    st.divider()

    archivo_imagen = st.file_uploader(
        "📷 Selecciona una imagen del PLC",
        type=["jpg", "jpeg", "png", "webp"],
        key="cargar_foto_plc"
    )

    if archivo_imagen is not None:

        try:
            imagen = Image.open(archivo_imagen).convert("RGB")

            st.subheader("Fotografía del PLC")

            st.image(
                imagen,
                caption="Imagen cargada para reconocimiento",
                use_container_width=True
            )

            st.divider()

            if st.button(
                "🔍 ANALIZAR IMAGEN",
                use_container_width=True,
                key="analizar_plc"
            ):

                try:
                    with st.spinner(
                        "Analizando la imagen del PLC..."
                    ):
                        modelo_yolo = cargar_modelo_yolo()

                        resultados = modelo_yolo.predict(
                            source=imagen,
                            conf=0.25,
                            verbose=False
                        )

                        resultado = resultados[0]

st.subheader("📊 Resultados del análisis")

# Solo aceptar clases que identifiquen explícitamente un PLC
clases_plc = (
    "plc",
    "programmable logic controller",
    "controlador lógico programable",
)

coincidencias_plc = []

if resultado.boxes is not None:
    for caja in resultado.boxes:
        clase = int(caja.cls[0].item())
        confianza = float(caja.conf[0].item()) * 100
        nombre_clase = str(resultado.names[clase])

        if any(
            termino in nombre_clase.lower()
            for termino in clases_plc
        ):
            coincidencias_plc.append(
                (nombre_clase, confianza)
            )

if coincidencias_plc:
    st.success("Se detectó una clase compatible con PLC.")

    for nombre_clase, confianza in coincidencias_plc:
        st.write(f"**Clase detectada:** {nombre_clase}")
        st.write(f"**Confianza:** {confianza:.1f}%")

    st.info(
        "La detección no confirma por sí sola la marca "
        "ni el modelo exacto del equipo."
    )
else:
    st.warning(
        "No se pudo confirmar que el objeto sea un PLC."
    )

    st.write(
        "El modelo actual no reconoce una clase específica "
        "de PLC en esta imagen. No se mostrará un celular, "
        "refrigerador u otro objeto como identificación válida."
    )

    st.info(
        "Para identificar PLC reales necesitaremos integrar "
        "un modelo especializado y validarlo con fotografías "
        "de controladores industriales."
    )
                except Exception as error:
                    st.error(
                        "No fue posible analizar la imagen. "
                        "Revisa el modelo YOLO y las dependencias."
                    )
                    st.code(str(error))

        except Exception as error:
            st.error("No fue posible abrir la imagen.")
            st.code(str(error))

    else:
        st.info(
            "Selecciona una fotografía para comenzar. "
            "Después podrás visualizarla y analizarla."
        )


# ==========================================
# COMUNICACION INDUSTRIAL
# ==========================================

elif seccion == "Comunicación industrial":

    st.header("Comunicación industrial")

    st.write(
        "Consulta protocolos e interfaces utilizados "
        "por los controladores lógicos programables."
    )

    with st.expander("RS-485"):
        st.write(
            "Interfaz eléctrica diferencial utilizada en "
            "comunicaciones industriales. Puede emplearse con "
            "protocolos como Modbus RTU, según el equipo."
        )

    with st.expander("Ethernet industrial"):
        st.write(
            "Permite la comunicación mediante redes Ethernet. "
            "El protocolo y los parámetros dependen del fabricante "
            "y del modelo del PLC."
        )

    with st.expander("PROFIBUS"):
        st.write(
            "Red industrial utilizada para intercambiar datos "
            "entre controladores y dispositivos compatibles."
        )

    with st.expander("USB y puertos de programación"):
        st.write(
            "Algunos PLC utilizan USB u otras interfaces dedicadas "
            "para programación y diagnóstico. Se debe verificar "
            "el modelo y el cable requerido."
        )

    st.divider()
    st.markdown("### Datos de comunicación de tu base de datos")

    if datos is not None and not datos.empty:

        columnas_com = [
            columna for columna in datos.columns
            if any(
                palabra in normalizar(columna)
                for palabra in [
                    "comunicacion",
                    "protocolo",
                    "interfaz",
                    "red",
                    "puerto"
                ]
            )
        ]

        if columnas_com:
            st.dataframe(
                datos[columnas_com],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info(
                "Tu Excel todavía no contiene columnas "
                "identificables de comunicación o interfaz."
            )

    else:
        st.info(
            "Cuando cargues la base de datos, aquí aparecerán "
            "los datos de comunicación que contenga."
        )


# ==========================================
# PIE DE PAGINA
# ==========================================

st.divider()

st.caption(
    "PLC SCAN | Automatización Industrial | DJH"
)
