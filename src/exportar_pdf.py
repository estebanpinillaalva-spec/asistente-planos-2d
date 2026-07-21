r"""Exporta planos de SolidWorks a PDF, PNG o ambos formatos.

Uso: .venv\Scripts\python.exe src\exportar_pdf.py [--formato pdf|png|ambos] "C:\ruta\plano.SLDDRW"
"""

import sys
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

from com_utils import (
    conectar,
    es_archivo_bloqueo,
    leer,
    normalizar_ruta,
    obtener_rutas_documentos_abiertos,
    texto_error_com,
)


# Valores numéricos definidos por la API de SolidWorks.
TIPO_DOCUMENTO_PLANO = 3
OPCION_ABRIR_SILENCIOSO = 1
VERSION_GUARDADO_ACTUAL = 0
OPCION_GUARDAR_SILENCIOSO = 1

# Constantes comprobadas en SolidWorks.Interop.swconst 31.2.1.1:
# swTiffScreenOrPrintCapture = 6 y swTiffPrintDPI = 9.
# SolidWorks reutiliza estas preferencias al guardar imágenes PNG.
PREFERENCIA_CAPTURA_TIFF = 6
PREFERENCIA_DPI_TIFF = 9
CAPTURA_DE_IMPRESION = 1
# A 300 DPI, incluso el lado corto de un A4 supera los 2000 píxeles.
DPI_PNG = 300

CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "notas" / "planos-pdf"


def mostrar_uso():
    """Muestra cómo indicar el formato y los planos que se quieren exportar."""
    print(
        "Uso: exportar_pdf.py [--formato pdf|png|ambos] "
        "<plano.SLDDRW o carpeta> [otra ruta ...]"
    )


def analizar_argumentos(argumentos):
    """Separa el formato de las rutas y comprueba los errores de uso."""
    formato = "pdf"
    entradas = []
    indice = 0

    while indice < len(argumentos):
        argumento = argumentos[indice]
        if argumento == "--formato":
            if indice + 1 >= len(argumentos):
                raise ValueError("Falta el valor de --formato.")
            formato = argumentos[indice + 1].lower()
            if formato not in ("pdf", "png", "ambos"):
                raise ValueError(
                    "Formato no válido. Usa pdf, png o ambos."
                )
            indice += 2
            continue

        if argumento.startswith("--"):
            raise ValueError(f"Opción no reconocida: {argumento}")

        entradas.append(argumento)
        indice += 1

    if not entradas:
        raise ValueError("Indica al menos un plano SLDDRW o una carpeta con planos.")

    return formato, entradas


def texto_error(error):
    """Devuelve un texto legible tanto para errores COM como para errores Python."""
    if isinstance(error, (pywintypes.com_error, AttributeError)):
        return texto_error_com(error)
    return f"{type(error).__name__}: {error}"


def configurar_resolucion_png(swapp):
    """Pide una captura de impresión a 300 DPI para obtener un PNG legible."""
    try:
        captura_configurada = swapp.SetUserPreferenceIntegerValue(
            PREFERENCIA_CAPTURA_TIFF,
            CAPTURA_DE_IMPRESION,
        )
        dpi_configurado = swapp.SetUserPreferenceIntegerValue(
            PREFERENCIA_DPI_TIFF,
            DPI_PNG,
        )
        if not captura_configurada or not dpi_configurado:
            raise RuntimeError("SolidWorks rechazó la preferencia de imagen")
        return True
    except Exception as error:
        # La resolución ayuda a leer las cotas, pero no debe impedir la exportación.
        print(
            "Aviso: no se pudo fijar la resolución PNG a 300 DPI; "
            "se exportará con la resolución por defecto de SolidWorks."
        )
        print(f"  Detalle: {texto_error(error)}")
        return False


def buscar_planos(entradas):
    """Recoge planos sueltos y planos dentro de carpetas, sin duplicados."""
    planos = set()

    for entrada in entradas:
        if entrada.is_file():
            if (
                entrada.suffix.lower() == ".slddrw"
                and not es_archivo_bloqueo(entrada)
            ):
                planos.add(entrada.resolve())
            else:
                print(f"Aviso: el archivo no es un plano SLDDRW válido: {entrada}")
            continue

        if entrada.is_dir():
            for ruta in entrada.rglob("*"):
                if (
                    ruta.is_file()
                    and ruta.suffix.lower() == ".slddrw"
                    and not es_archivo_bloqueo(ruta)
                ):
                    planos.add(ruta.resolve())
            continue

        print(f"Aviso: no existe el archivo o la carpeta: {entrada}")

    return sorted(planos, key=lambda ruta: str(ruta).lower())


def abrir_plano(swapp, ruta):
    """Abre un plano en silencio y devuelve documento, errores y avisos."""
    errores = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    avisos = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    documento = swapp.OpenDoc6(
        str(ruta),
        TIPO_DOCUMENTO_PLANO,
        OPCION_ABRIR_SILENCIOSO,
        "",
        errores,
        avisos,
    )

    return documento, int(errores.value), int(avisos.value)


def exportar_con_extension(documento, ruta_salida):
    """Exporta con IModelDocExtension.SaveAs y recoge errores y avisos."""
    # Extension devuelve un objeto COM; por eso no se usa leer().
    extension = documento.Extension
    errores = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    avisos = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    correcto = extension.SaveAs(
        str(ruta_salida),
        VERSION_GUARDADO_ACTUAL,
        OPCION_GUARDAR_SILENCIOSO,
        None,
        errores,
        avisos,
    )

    codigo_errores = int(errores.value)
    codigo_avisos = int(avisos.value)
    if not correcto:
        raise RuntimeError(
            "IModelDocExtension.SaveAs devolvió False; "
            f"errores={codigo_errores}, avisos={codigo_avisos}"
        )

    return codigo_errores, codigo_avisos


def exportar_documento(documento, ruta_salida):
    """Prueba la exportación recomendada y usa SaveAs antiguo como alternativa."""
    try:
        errores, avisos = exportar_con_extension(documento, ruta_salida)
        return "IModelDocExtension.SaveAs", errores, avisos
    except Exception as error_principal:
        try:
            correcto = documento.SaveAs(str(ruta_salida))
            if not correcto:
                raise RuntimeError("IModelDoc2.SaveAs devolvió False")
            return "IModelDoc2.SaveAs", 0, 0
        except Exception as error_alternativo:
            raise RuntimeError(
                "fallaron IModelDocExtension.SaveAs "
                f"({texto_error(error_principal)}) e IModelDoc2.SaveAs "
                f"({texto_error(error_alternativo)})"
            ) from error_alternativo


def procesar_plano(swapp, ruta, rutas_abiertas_antes, formato="pdf"):
    """Abre, exporta y cierra un plano sin propagar sus errores."""
    extensiones = ["pdf", "png"] if formato == "ambos" else [formato]
    resultado = {
        "plano": ruta,
        "salidas": [],
        "exportado": False,
        "error": "",
    }
    documento = None
    nombre_para_cerrar = ruta.name
    ya_estaba_abierto = normalizar_ruta(ruta) in rutas_abiertas_antes

    try:
        documento, errores_apertura, avisos_apertura = abrir_plano(swapp, ruta)
        if documento is None:
            raise RuntimeError(
                "OpenDoc6 no abrió el plano; "
                f"errores={errores_apertura}, avisos={avisos_apertura}"
            )

        # GetTitle devuelve una cadena simple, así que aquí sí usamos leer().
        nombre_para_cerrar = leer(documento, "GetTitle") or ruta.name

        notas = []
        if errores_apertura or avisos_apertura:
            notas.append(
                f"OpenDoc6: errores={errores_apertura}, avisos={avisos_apertura}"
            )

        # En modo "ambos", un formato puede fallar sin impedir probar el otro.
        for extension in extensiones:
            ruta_salida = CARPETA_SALIDA / f"{ruta.stem}.{extension}"
            try:
                via, errores_guardado, avisos_guardado = exportar_documento(
                    documento, ruta_salida
                )
                resultado["salidas"].append(
                    {"ruta": ruta_salida, "via": via}
                )
                if errores_guardado or avisos_guardado:
                    notas.append(
                        f"{extension.upper()} SaveAs: "
                        f"errores={errores_guardado}, avisos={avisos_guardado}"
                    )
            except Exception as error:
                notas.append(f"{extension.upper()}: {texto_error(error)}")

        resultado["exportado"] = len(resultado["salidas"]) == len(extensiones)
        resultado["error"] = "; ".join(notas)
    except (pywintypes.com_error, AttributeError) as error:
        resultado["error"] = texto_error_com(error)
    except Exception as error:
        resultado["error"] = texto_error(error)
    finally:
        # Los planos que ya estaban abiertos pertenecen al usuario y no se cierran.
        if documento is not None and not ya_estaba_abierto:
            try:
                swapp.CloseDoc(nombre_para_cerrar)
            except (pywintypes.com_error, AttributeError) as error:
                detalle = f"no se pudo cerrar: {texto_error_com(error)}"
                resultado["error"] = "; ".join(
                    texto for texto in [resultado["error"], detalle] if texto
                )
                resultado["exportado"] = False
            except Exception as error:
                detalle = f"no se pudo cerrar: {texto_error(error)}"
                resultado["error"] = "; ".join(
                    texto for texto in [resultado["error"], detalle] if texto
                )
                resultado["exportado"] = False

    return resultado


def imprimir_resultado(resultado):
    """Muestra el resultado de un plano en una sola línea principal."""
    for salida in resultado["salidas"]:
        print(
            f"EXPORTADO | {resultado['plano'].name} | "
            f"{salida['ruta']} | {salida['via']}"
        )

    if resultado["exportado"]:
        if resultado["error"]:
            print(f"  Aviso: {resultado['error']}")
        return

    print(f"ERROR | {resultado['plano'].name} | {resultado['error']}")


def main():
    """Coordina la búsqueda y exportación de todos los planos indicados."""
    try:
        # Evita fallos al imprimir tildes en consolas Windows antiguas.
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    try:
        formato, argumentos_entrada = analizar_argumentos(sys.argv[1:])
    except ValueError as error:
        mostrar_uso()
        print(f"Error: {error}")
        return 2

    try:
        entradas = [
            Path(argumento).expanduser().resolve()
            for argumento in argumentos_entrada
        ]
        swapp = conectar()
        if swapp is None:
            return 1

        # GetDocuments registra qué planos pertenecían ya al usuario.
        rutas_abiertas_antes = obtener_rutas_documentos_abiertos(swapp)
        planos = buscar_planos(entradas)
        CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

        if formato in ("png", "ambos"):
            configurar_resolucion_png(swapp)

        print(f"Planos SLDDRW encontrados: {len(planos)}")
        resultados = []
        for indice, ruta in enumerate(planos, start=1):
            print(f"[{indice}/{len(planos)}] {ruta}")
            resultado = procesar_plano(
                swapp, ruta, rutas_abiertas_antes, formato
            )
            resultados.append(resultado)
            imprimir_resultado(resultado)

        exportados = sum(resultado["exportado"] for resultado in resultados)
        fallidos = len(resultados) - exportados

        print("Resumen:")
        print(f"  Exportados: {exportados}")
        print(f"  Fallidos: {fallidos}")
        print(f"Carpeta de salida: {CARPETA_SALIDA}")
        return 0
    except Exception as error:
        print("No se pudo completar la exportación de planos.")
        print(f"Detalle del error: {texto_error(error)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
