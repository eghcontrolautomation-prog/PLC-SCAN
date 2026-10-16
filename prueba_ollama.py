
import requests

OLLAMA_URL = "http://localhost:11434/api/generate"
MODELO = "qwen2.5:3b"


def consultar_ollama(pregunta):
    datos = {
        "model": MODELO,
        "prompt": pregunta,
        "stream": False
    }

    try:
        respuesta = requests.post(
            OLLAMA_URL,
            json=datos,
            timeout=120
        )

        respuesta.raise_for_status()
        resultado = respuesta.json()

        return resultado.get(
            "response",
            "No se recibió una respuesta del modelo."
        )

    except requests.exceptions.ConnectionError:
        return (
            "No se pudo conectar con Ollama. "
            "Verifica que esté ejecutándose."
        )

    except requests.exceptions.Timeout:
        return "La consulta tardó demasiado. Inténtalo nuevamente."

    except requests.exceptions.RequestException as error:
        return f"Error al consultar Ollama: {error}"


if __name__ == "__main__":
    pregunta = (
        "Explica qué información debemos consultar "
        "para identificar correctamente un PLC y su "
        "interfaz de programación. No inventes datos "
        "específicos de ningún fabricante."
    )

    print("\n=== PRUEBA DE PLC SCAN + OLLAMA ===\n")
    print(consultar_ollama(pregunta))
