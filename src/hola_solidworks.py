r"""Conecta Python con SolidWorks y muestra el documento activo.

Uso normal: .venv\Scripts\python.exe src\hola_solidworks.py
Para abrir SolidWorks: .venv\Scripts\python.exe src\hola_solidworks.py --iniciar
"""

import sys

import pywintypes

from com_utils import conectar, leer, texto_error_com


def describir_documento_activo(swapp):
    """Muestra el nombre, la ruta y el tipo del documento activo."""
    try:
        # ActiveDoc devuelve un objeto COM invocable; por eso no usamos leer().
        documento = swapp.ActiveDoc
    except (pywintypes.com_error, AttributeError) as error:
        print("No se pudo consultar el documento activo de SolidWorks.")
        print(f"Detalle COM: {texto_error_com(error)}")
        return 1

    if documento is None:
        print("SolidWorks está abierto pero no hay ningún documento activo.")
        return 0

    try:
        nombre = leer(documento, "GetTitle")
        ruta = leer(documento, "GetPathName")
        codigo_tipo = leer(documento, "GetType")
    except (pywintypes.com_error, AttributeError) as error:
        print("No se pudo leer la información del documento activo.")
        print(f"Detalle COM: {texto_error_com(error)}")
        return 1

    tipos_documento = {
        1: "pieza",
        2: "ensamblaje",
        3: "plano",
    }
    tipo = tipos_documento.get(codigo_tipo, f"desconocido (código {codigo_tipo})")

    print(f"Documento activo: {nombre}")
    print(f"Ruta completa: {ruta or '(documento todavía sin guardar)'}")
    print(f"Tipo: {tipo}")
    return 0


def main():
    """Coordina la conexión y la lectura de información de SolidWorks."""
    try:
        # Evita errores al imprimir tildes en consolas Windows con codepage antigua.
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        # Algunas salidas redirigidas no permiten cambiar la codificación.
        pass

    try:
        argumentos = sys.argv[1:]
        argumentos_desconocidos = [
            argumento for argumento in argumentos if argumento != "--iniciar"
        ]

        if argumentos_desconocidos:
            print(f"Argumento no reconocido: {argumentos_desconocidos[0]}")
            print("Uso: hola_solidworks.py [--iniciar]")
            return 2

        swapp = conectar(iniciar="--iniciar" in argumentos)
        if swapp is None:
            return 1

        try:
            # Usamos leer() porque late binding no trata siempre igual estos miembros COM.
            version = leer(swapp, "RevisionNumber")
        except (pywintypes.com_error, AttributeError) as error:
            print("No se pudo consultar la versión de SolidWorks.")
            print(f"Detalle COM: {texto_error_com(error)}")
            return 1

        print(f"Versión / build de SolidWorks: {version}")
        return describir_documento_activo(swapp)
    except Exception as error:
        print(
            "No se pudo conectar con SolidWorks o leer la información "
            "del documento activo."
        )
        print(f"Detalle del error: {type(error).__name__}: {error}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
