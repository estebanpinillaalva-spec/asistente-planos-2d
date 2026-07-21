r"""Conecta Python con SolidWorks y muestra el documento activo.

Uso normal: .venv\Scripts\python.exe src\hola_solidworks.py
Para abrir SolidWorks: .venv\Scripts\python.exe src\hola_solidworks.py --iniciar
"""

import sys

import pywintypes
import win32com.client


def texto_error_com(error):
    """Convierte un error COM en un texto corto y entendible."""
    if isinstance(error, AttributeError):
        return f"AttributeError: {error}"

    codigo = getattr(error, "hresult", None)
    descripcion = getattr(error, "strerror", None)

    # Algunos errores incluyen una descripción más concreta en excepinfo.
    informacion_extra = getattr(error, "excepinfo", None)
    if informacion_extra and len(informacion_extra) > 2 and informacion_extra[2]:
        descripcion = informacion_extra[2]

    if codigo is None and error.args:
        codigo = error.args[0]
    if not descripcion and len(error.args) > 1:
        descripcion = error.args[1]

    return f"código {codigo}: {descripcion or 'sin descripción disponible'}"


def leer(objeto, nombre):
    """Lee un miembro COM sin argumentos que devuelve un valor simple.

    En enlace tardío, pywin32 puede entregar cadenas, números o booleanos
    directamente o como un método sin argumentos que hay que llamar.

    No debe usarse con miembros que devuelven objetos COM, como ActiveDoc.
    Los envoltorios CDispatch de pywin32 son invocables y leer() los llamaría
    por error, lo que provoca DISP_E_MEMBERNOTFOUND.
    """
    valor = getattr(objeto, nombre)
    return valor() if callable(valor) else valor


def hacer_visible(swapp):
    """Intenta mostrar SolidWorks aunque Visible cambie con el enlace tardío."""
    try:
        # Primero probamos la forma normal de escribir una propiedad COM.
        swapp.Visible = True
        return
    except (AttributeError, TypeError):
        # Si pywin32 lo entrega como método, hay que llamarlo con el valor.
        pass
    except Exception as error:
        print("Aviso: no se pudo hacer visible SolidWorks, pero sigue iniciado.")
        print(f"Detalle del aviso: {type(error).__name__}: {error}")
        return

    try:
        visible = getattr(swapp, "Visible")
        visible(True)
    except Exception as error:
        # La visibilidad no es imprescindible: avisamos y continuamos.
        print("Aviso: no se pudo hacer visible SolidWorks, pero sigue iniciado.")
        print(f"Detalle del aviso: {type(error).__name__}: {error}")


def conectar(iniciar=False):
    """Se conecta a SolidWorks y solo lo inicia si se pidió expresamente."""
    try:
        # Primero reutilizamos SolidWorks si el usuario ya lo tiene abierto.
        return win32com.client.GetActiveObject("SldWorks.Application")
    except (pywintypes.com_error, AttributeError) as error:
        if not iniciar:
            print(
                "SolidWorks no está abierto. "
                "Ábrelo con una pieza y vuelve a ejecutar."
            )
            print(f"Detalle COM: {texto_error_com(error)}")
            return None

    print("Iniciando SolidWorks; puede tardar unos segundos...")
    try:
        # Dispatch crea la aplicación. Visible permite que el usuario la vea.
        swapp = win32com.client.Dispatch("SldWorks.Application")
        hacer_visible(swapp)
        return swapp
    except (pywintypes.com_error, AttributeError) as error:
        print("No se pudo iniciar SolidWorks.")
        print(f"Detalle COM: {texto_error_com(error)}")
        return None


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
