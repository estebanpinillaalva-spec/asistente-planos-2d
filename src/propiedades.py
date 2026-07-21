"""Lee propiedades personalizadas de documentos abiertos de SolidWorks."""

import pythoncom
import win32com.client

from com_utils import leer


def nombre_configuracion_activa(documento):
    """Devuelve el nombre de la configuración activa del documento."""
    # ConfigurationManager y ActiveConfiguration devuelven objetos COM.
    # Por eso se accede directamente y no se usa leer() con ellos.
    administrador = documento.ConfigurationManager
    configuracion = administrador.ActiveConfiguration

    # Name sí devuelve un texto simple, así que aquí leer() es seguro.
    return str(leer(configuracion, "Name") or "")


def _texto(valor):
    """Convierte un valor COM vacío en una cadena fácil de procesar."""
    return "" if valor is None else str(valor)


def material_asignado(documento, configuracion):
    """Devuelve el material asignado a una configuración de una pieza."""
    # GetMaterialPropertyName2 devuelve el nombre del material y escribe, por
    # referencia, la ruta de la base de datos. Aquí solo necesitamos el nombre.
    base_datos = win32com.client.VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_BSTR, ""
    )

    try:
        material = documento.GetMaterialPropertyName2(
            configuracion, base_datos
        )
    except TypeError:
        # Con enlace temprano, pywin32 puede devolver juntos el resultado y
        # el parámetro de salida, sin aceptar un VARIANT por referencia.
        resultado = documento.GetMaterialPropertyName2(configuracion)
        if isinstance(resultado, tuple):
            material = resultado[0] if resultado else ""
        else:
            material = resultado

    return _texto(material)


def _leer_valor(administrador, nombre):
    """Lee el texto escrito y el valor resuelto de una propiedad."""
    # False pide a SolidWorks el valor actualizado, no el valor en caché.
    for nombre_metodo in ("Get6", "Get5", "Get4"):
        try:
            metodo = getattr(administrador, nombre_metodo)
        except AttributeError:
            continue

        valor = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_BSTR, ""
        )
        resuelto = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_BSTR, ""
        )
        fue_resuelto = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_BOOL, False
        )
        vinculado = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_BOOL, False
        )

        argumentos = [nombre, False, valor, resuelto]
        if nombre_metodo in ("Get6", "Get5"):
            argumentos.append(fue_resuelto)
        if nombre_metodo == "Get6":
            argumentos.append(vinculado)

        try:
            resultado = metodo(*argumentos)
        except TypeError:
            # Con enlace temprano, pywin32 puede devolver los parámetros de
            # salida juntos en una tupla en vez de aceptar VARIANT por referencia.
            resultado = metodo(nombre, False)
            cadenas = [dato for dato in resultado if isinstance(dato, str)]
            if len(cadenas) >= 2:
                return _texto(cadenas[0]), _texto(cadenas[1])
            raise

        return _texto(valor.value), _texto(resuelto.value)

    raise AttributeError(
        "ICustomPropertyManager no ofrece Get6, Get5 ni Get4."
    )


def _leer_origen(administrador):
    """Lee todas las propiedades de un administrador de propiedades COM."""
    # GetNames devuelve una lista simple; leer() cubre las dos formas de pywin32.
    nombres = leer(administrador, "GetNames")
    if not nombres:
        return {}
    if isinstance(nombres, str):
        nombres = [nombres]

    propiedades = {}
    for nombre in nombres:
        valor, resuelto = _leer_valor(administrador, str(nombre))
        propiedades[str(nombre)] = {
            "valor": valor,
            "resuelto": resuelto,
        }
    return propiedades


def leer_propiedades(documento, configuracion=None):
    """Lee propiedades del archivo y de una configuración del documento."""
    if configuracion is None:
        configuracion = nombre_configuracion_activa(documento)

    # Extension y CustomPropertyManager devuelven objetos COM. Se accede a
    # ambos directamente porque leer() llamaría por error a sus CDispatch.
    extension = documento.Extension
    propiedades_archivo = extension.CustomPropertyManager("")

    # Los planos no tienen una configuración de pieza. Una cadena vacía indica
    # que solo se debe consultar el bloque de propiedades del propio archivo.
    if configuracion:
        propiedades_configuracion = extension.CustomPropertyManager(
            configuracion
        )
        datos_configuracion = _leer_origen(propiedades_configuracion)
    else:
        datos_configuracion = {}

    return {
        "archivo": _leer_origen(propiedades_archivo),
        "configuracion": datos_configuracion,
    }
