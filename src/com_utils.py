"""Ayudas sencillas para trabajar con SolidWorks mediante COM y pywin32."""

from pathlib import Path

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


def es_archivo_bloqueo(ruta):
    """Indica si una ruta es un archivo temporal de bloqueo de SolidWorks."""
    return Path(ruta).name.startswith("~$")


def normalizar_ruta(ruta):
    """Normaliza una ruta para poder compararla sin distinguir mayúsculas."""
    return str(Path(ruta).expanduser().resolve()).casefold()


def obtener_rutas_documentos_abiertos(swapp):
    """Devuelve las rutas normalizadas de los documentos abiertos por el usuario."""
    # GetDocuments devuelve objetos COM; por eso se accede directamente y sin leer().
    documentos = swapp.GetDocuments

    # Algunas versiones lo entregan como método. Un CDispatch también es invocable,
    # así que comprobamos _oleobj_ antes de llamar para no repetir aquel error.
    if callable(documentos) and not hasattr(documentos, "_oleobj_"):
        documentos = documentos()

    if not documentos:
        return set()

    # Normalmente llega una tupla, pero envolvemos un único documento por seguridad.
    if hasattr(documentos, "_oleobj_"):
        documentos = (documentos,)

    rutas = set()
    for documento in documentos:
        # GetPathName devuelve una cadena simple, así que aquí sí usamos leer().
        ruta = leer(documento, "GetPathName")
        if ruta:
            rutas.add(normalizar_ruta(ruta))

    return rutas
