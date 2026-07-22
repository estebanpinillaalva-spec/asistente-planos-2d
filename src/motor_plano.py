r"""Crea planos de SolidWorks con alzado e isométrica.

Uso: .venv\Scripts\python.exe src\motor_plano.py "C:\ruta\pieza.SLDPRT"
"""

import sys
import unicodedata
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


# Carpeta configurable donde están los formatos de hoja de la escuela.
CARPETA_PLANTILLAS = Path(
    r"C:\Users\esteb\Desktop\TFG DOCUMENTACION\Plantillas_TFG (1)"
)

# Medidas observadas de los formatos, en metros, usadas solo para elegir hoja
# y como alternativa. El .slddrt puede imponer otras medidas u orientación:
# la verdad será siempre lo que devuelva ISheet después de aplicarlo.
ANCHO_A4_M = 0.210  # A4 vertical observado
ALTO_A4_M = 0.297
ANCHO_A3_M = 0.420
ALTO_A3_M = 0.297  # A3 apaisado observado; pendiente de confirmar por API
ANCHO_A2_M = 0.594
ALTO_A2_M = 0.420  # A2 apaisado observado; pendiente de confirmar por API

# Se mantiene el orden A4, A3 y A2 para elegir primero el formato menor.
FORMATOS_HOJA = [
    {
        "nombre": "A4",
        "archivo": "ig_2017_a4_-_iso.slddrt",
        "ancho_m": ANCHO_A4_M,
        "alto_m": ALTO_A4_M,
    },
    {
        "nombre": "A3",
        "archivo": "ig_2017_a3_-_iso.slddrt",
        "ancho_m": ANCHO_A3_M,
        "alto_m": ALTO_A3_M,
    },
    {
        "nombre": "A2",
        "archivo": "ig_2017_a2_-_iso.slddrt",
        "ancho_m": ANCHO_A2_M,
        "alto_m": ALTO_A2_M,
    },
]

# Se respeta exactamente este orden de preferencia indicado en el plan.
ESCALAS = [
    {"texto": "1:1", "numerador": 1.0, "denominador": 1.0},
    {"texto": "1:2", "numerador": 1.0, "denominador": 2.0},
    {"texto": "1:5", "numerador": 1.0, "denominador": 5.0},
    {"texto": "1:10", "numerador": 1.0, "denominador": 10.0},
]

# Margen de encuadernación: el marco de la plantilla está más adentro
# por el lado izquierdo.
MARGEN_IZQUIERDO_MM = 30.0
MARGEN_DERECHO_MM = 20.0
# El marco superior también está más adentro en estas plantillas.
MARGEN_SUPERIOR_MM = 20.0
BANDA_CAJETIN_MM = 65.0
SEPARACION_VISTAS_MM = 12.0

# La isométrica se aproxima como un cuadrado. Este factor se podrá ajustar
# cuando comparemos la estimación con más planos generados.
FACTOR_TAMANO_ISOMETRICA = 0.85
TIPO_CUERPO_SOLIDO = 0

# Valores numéricos de SolidWorks.Interop.swconst 31.2.1.1 (SolidWorks 2023).
TIPO_DOCUMENTO_PIEZA = 1
OPCION_ABRIR_SILENCIOSO = 1
PAPEL_DE_USUARIO = 12  # swDwgPapersUserDefined, confirmado para 2023
PLANTILLA_PERSONALIZADA = 12  # swDwgTemplateCustom, confirmado para 2023
PREFERENCIA_PLANTILLA_PLANO = 10  # swDefaultTemplateDrawing
PREFERENCIA_AUTOESCALA_VISTAS = 86  # swAutomaticScaling3ViewDrawings
# Preferencias de documento confirmadas en swconst.tlb de SolidWorks 2023:
# swDetailingAutoInsertCenterMarksForHoles = 308
# swDetailingAutoInsertCenterLines = 190
# swDetailingConnectionLinesHolePatternsCenterMarks = 599
# swDetailingCenterMarkScaleByViewScale = 613
# swDetailingCenterMarkSize = 12
OPCION_DETALLADO_SIN_ESPECIFICAR = 0  # swDetailingNoOptionSpecified
PREFERENCIAS_BOOLEANAS_ANOTACIONES = [
    ("swDetailingAutoInsertCenterMarksForHoles", 308),
    ("swDetailingAutoInsertCenterLines", 190),
    ("swDetailingConnectionLinesHolePatternsCenterMarks", 599),
    ("swDetailingCenterMarkScaleByViewScale", 613),
]
PREFERENCIA_TAMANO_MARCA_CENTRO = 12
TAMANO_MARCA_CENTRO_M = 0.0015  # 1,5 mm sobre el plano.
VERSION_GUARDADO_ACTUAL = 0
OPCION_GUARDAR_SILENCIOSO = 1

CARPETA_SALIDA = (
    Path(__file__).resolve().parent.parent / "notas" / "planos-generados"
)


def _validar_caja(caja):
    """Convierte una caja COM en seis números y comprueba su formato."""
    if caja is None or len(caja) != 6:
        raise ValueError("la caja envolvente no contiene seis valores")
    return tuple(float(valor) for valor in caja)


def _caja_por_cuerpos(pieza):
    """Combina las cajas de los cuerpos si falla la caja general."""
    # Esta pequeña alternativa procede de clasificar_planas.py. GetBodies2
    # devuelve objetos COM, así que no se usa leer() en esa llamada.
    cuerpos = pieza.GetBodies2(TIPO_CUERPO_SOLIDO, False)
    if not cuerpos:
        raise ValueError("la pieza no contiene cuerpos sólidos")

    cajas = [_validar_caja(leer(cuerpo, "GetBodyBox")) for cuerpo in cuerpos]
    return (
        min(caja[0] for caja in cajas),
        min(caja[1] for caja in cajas),
        min(caja[2] for caja in cajas),
        max(caja[3] for caja in cajas),
        max(caja[4] for caja in cajas),
        max(caja[5] for caja in cajas),
    )


def _obtener_caja(pieza):
    """Obtiene la caja de pieza y usa las cajas de cuerpos como alternativa."""
    try:
        # True devuelve las coordenadas en las unidades internas: metros.
        return _validar_caja(pieza.GetPartBox(True))
    except Exception as error_principal:
        try:
            return _caja_por_cuerpos(pieza)
        except Exception as error_alternativo:
            raise RuntimeError(
                "no se pudo obtener la caja envolvente con GetPartBox ni "
                "con las cajas de los cuerpos"
            ) from error_alternativo


def dimensiones_pieza(pieza):
    """Devuelve espesor, dimensión media y largo de la pieza en milímetros."""
    xmin, ymin, zmin, xmax, ymax, zmax = _obtener_caja(pieza)
    dimensiones = sorted(
        [
            abs(xmax - xmin) * 1000.0,
            abs(ymax - ymin) * 1000.0,
            abs(zmax - zmin) * 1000.0,
        ]
    )
    if dimensiones[0] <= 0:
        raise ValueError("la caja envolvente tiene una dimensión igual a cero")
    return tuple(dimensiones)


def calcular_tamanos_vistas(largo_mm, medio_mm, escala):
    """Estima en milímetros cuánto ocupará cada vista sobre el papel."""
    factor_escala = escala["numerador"] / escala["denominador"]
    lado_isometrica = (
        FACTOR_TAMANO_ISOMETRICA * (largo_mm + medio_mm) * factor_escala
    )
    return {
        "alzado": {
            "ancho_mm": largo_mm * factor_escala,
            "alto_mm": medio_mm * factor_escala,
        },
        "isometrica": {
            "ancho_mm": lado_isometrica,
            "alto_mm": lado_isometrica,
        },
    }


def calcular_area_util(ancho_hoja_m, alto_hoja_m):
    """Calcula el rectángulo disponible tras descontar marco y cajetín."""
    ancho_hoja_mm = ancho_hoja_m * 1000.0
    alto_hoja_mm = alto_hoja_m * 1000.0
    ancho_util_mm = (
        ancho_hoja_mm - MARGEN_IZQUIERDO_MM - MARGEN_DERECHO_MM
    )
    alto_util_mm = alto_hoja_mm - BANDA_CAJETIN_MM - MARGEN_SUPERIOR_MM
    if ancho_util_mm <= 0 or alto_util_mm <= 0:
        raise ValueError(
            f"la hoja {ancho_hoja_mm:.1f} x {alto_hoja_mm:.1f} mm "
            "no deja un área útil positiva"
        )
    return {
        "ancho_mm": ancho_util_mm,
        "alto_mm": alto_util_mm,
        "izquierda_mm": MARGEN_IZQUIERDO_MM,
        "inferior_mm": BANDA_CAJETIN_MM,
        "derecha_mm": MARGEN_IZQUIERDO_MM + ancho_util_mm,
        "superior_mm": BANDA_CAJETIN_MM + alto_util_mm,
        "centro_x_mm": MARGEN_IZQUIERDO_MM + ancho_util_mm / 2.0,
        "centro_y_mm": BANDA_CAJETIN_MM + alto_util_mm / 2.0,
    }


def calcular_rectangulo_vista(coordenadas_mm, tamano):
    """Calcula los cuatro lados de una vista a partir de su centro."""
    centro_x_mm, centro_y_mm = coordenadas_mm
    mitad_ancho = tamano["ancho_mm"] / 2.0
    mitad_alto = tamano["alto_mm"] / 2.0
    return {
        "izquierda_mm": centro_x_mm - mitad_ancho,
        "inferior_mm": centro_y_mm - mitad_alto,
        "derecha_mm": centro_x_mm + mitad_ancho,
        "superior_mm": centro_y_mm + mitad_alto,
    }


def verificar_rectangulos_en_area(area, rectangulos):
    """Falla si alguna vista rebasa uno de los límites del área útil."""
    tolerancia_mm = 0.000001
    for nombre, rectangulo in rectangulos.items():
        nombre_visible = "isométrica" if nombre == "isometrica" else nombre
        excesos = []
        exceso_izquierda = area["izquierda_mm"] - rectangulo["izquierda_mm"]
        exceso_derecha = rectangulo["derecha_mm"] - area["derecha_mm"]
        exceso_inferior = area["inferior_mm"] - rectangulo["inferior_mm"]
        exceso_superior = rectangulo["superior_mm"] - area["superior_mm"]

        for lado, exceso in (
            ("por la izquierda", exceso_izquierda),
            ("por la derecha", exceso_derecha),
            ("por abajo", exceso_inferior),
            ("por arriba", exceso_superior),
        ):
            if exceso > tolerancia_mm:
                excesos.append(f"{lado} {exceso:.1f} mm")

        if excesos:
            raise ValueError(
                f"la vista {nombre_visible} se sale del área útil: "
                + ", ".join(excesos)
            )


def calcular_disposicion(ancho_hoja_m, alto_hoja_m, tamanos):
    """Coloca la isométrica arriba a la izquierda y el alzado donde quepa."""
    area = calcular_area_util(ancho_hoja_m, alto_hoja_m)
    alzado = tamanos["alzado"]
    isometrica = tamanos["isometrica"]

    # La isométrica siempre queda pegada a la esquina superior izquierda del
    # área útil. CreateDrawViewFromModelView3 recibe el centro de la vista.
    coordenadas_isometrica_mm = (
        area["izquierda_mm"] + isometrica["ancho_mm"] / 2.0,
        area["superior_mm"] - isometrica["alto_mm"] / 2.0,
    )
    borde_derecho_isometrica = (
        area["izquierda_mm"] + isometrica["ancho_mm"]
    )
    inicio_franja_derecha = borde_derecho_isometrica + SEPARACION_VISTAS_MM
    ancho_franja_derecha = area["derecha_mm"] - inicio_franja_derecha
    cabe_a_la_derecha = (
        alzado["ancho_mm"] <= ancho_franja_derecha
        and alzado["alto_mm"] <= area["alto_mm"]
    )

    borde_inferior_isometrica = (
        area["superior_mm"] - isometrica["alto_mm"]
    )
    fin_franja_inferior = borde_inferior_isometrica - SEPARACION_VISTAS_MM
    alto_franja_inferior = fin_franja_inferior - area["inferior_mm"]
    cabe_debajo = (
        alzado["ancho_mm"] <= area["ancho_mm"]
        and alzado["alto_mm"] <= alto_franja_inferior
    )

    if cabe_a_la_derecha:
        coordenadas_alzado_mm = (
            inicio_franja_derecha + ancho_franja_derecha / 2.0,
            area["centro_y_mm"],
        )
        tipo = "alzado a la derecha"
    elif cabe_debajo:
        coordenadas_alzado_mm = (
            area["centro_x_mm"],
            area["inferior_mm"] + alto_franja_inferior / 2.0,
        )
        tipo = "alzado debajo"
    else:
        raise ValueError(
            "las vistas no caben: "
            f"hoja={ancho_hoja_m * 1000.0:.1f} x "
            f"{alto_hoja_m * 1000.0:.1f} mm; "
            f"área útil={area['ancho_mm']:.1f} x {area['alto_mm']:.1f} mm; "
            f"alzado={alzado['ancho_mm']:.1f} x {alzado['alto_mm']:.1f} mm; "
            f"isométrica={isometrica['ancho_mm']:.1f} x "
            f"{isometrica['alto_mm']:.1f} mm; no cabe ni a la derecha "
            "ni debajo de la isométrica"
        )

    rectangulos = {
        "alzado": calcular_rectangulo_vista(coordenadas_alzado_mm, alzado),
        "isometrica": calcular_rectangulo_vista(
            coordenadas_isometrica_mm, isometrica
        ),
    }
    # Esta comprobación se repite también al recalcular con la hoja real,
    # justo antes de que el flujo llegue a insertar las vistas.
    verificar_rectangulos_en_area(area, rectangulos)

    return {
        "tipo": tipo,
        "area_util": area,
        "tamanos": tamanos,
        "rectangulos": rectangulos,
        "coordenadas_alzado_m": tuple(
            valor / 1000.0 for valor in coordenadas_alzado_mm
        ),
        "coordenadas_isometrica_m": tuple(
            valor / 1000.0 for valor in coordenadas_isometrica_mm
        ),
    }


def elegir_formato_y_escala(largo_mm, medio_mm):
    """Elige la primera hoja y escala donde quepan las dos vistas."""
    ultimo_error = None
    # Prima la legibilidad: se agranda la hoja antes de reducir la escala.
    for escala in ESCALAS:
        for formato in FORMATOS_HOJA:
            tamanos = calcular_tamanos_vistas(largo_mm, medio_mm, escala)
            try:
                disposicion = calcular_disposicion(
                    formato["ancho_m"], formato["alto_m"], tamanos
                )
                return formato, escala, disposicion
            except ValueError as error:
                ultimo_error = error

    raise ValueError(
        "la pieza no cabe ni en A2 con las escalas permitidas; "
        f"dimensiones principales={largo_mm:.1f} x {medio_mm:.1f} mm; "
        f"{ultimo_error}"
    )


def buscar_piezas(entradas):
    """Busca piezas SLDPRT sueltas o dentro de carpetas, sin duplicados."""
    piezas = set()

    for entrada in entradas:
        if entrada.is_file():
            if (
                entrada.suffix.lower() == ".sldprt"
                and not es_archivo_bloqueo(entrada)
            ):
                piezas.add(entrada.resolve())
            else:
                print(f"Aviso: no es una pieza SLDPRT válida: {entrada}")
            continue

        if entrada.is_dir():
            for ruta in entrada.rglob("*"):
                if (
                    ruta.is_file()
                    and ruta.suffix.lower() == ".sldprt"
                    and not es_archivo_bloqueo(ruta)
                ):
                    piezas.add(ruta.resolve())
            continue

        print(f"Aviso: no existe el archivo o la carpeta: {entrada}")

    return sorted(piezas, key=lambda ruta: str(ruta).casefold())


def validar_plantillas():
    """Comprueba que existen la carpeta y los tres formatos de la escuela."""
    if not CARPETA_PLANTILLAS.is_dir():
        print("No se encontró la carpeta de plantillas de la escuela:")
        print(f"  {CARPETA_PLANTILLAS}")
        return False

    faltantes = [
        formato["archivo"]
        for formato in FORMATOS_HOJA
        if not (CARPETA_PLANTILLAS / formato["archivo"]).is_file()
    ]
    if faltantes:
        print("Faltan formatos de hoja en la carpeta de plantillas:")
        for nombre in faltantes:
            print(f"  {nombre}")
        return False

    return True


def texto_error(error):
    """Convierte cualquier error esperado en un mensaje corto y legible."""
    if isinstance(error, (pywintypes.com_error, AttributeError)):
        return texto_error_com(error)
    return f"{type(error).__name__}: {error}"


def aplicar_preferencias_anotaciones(plano):
    """Activa marcas y ejes, guardando cada valor original que modifica."""
    cambios = []
    avisos = []
    try:
        # Extension devuelve un objeto COM: acceso directo, sin leer().
        extension = plano.Extension
    except Exception as error:
        avisos.append(
            "no se pudo acceder a IModelDoc2.Extension para configurar "
            f"marcas de centro y ejes: {texto_error(error)}"
        )
        return None, cambios, avisos

    for nombre, identificador in PREFERENCIAS_BOOLEANAS_ANOTACIONES:
        try:
            original = bool(
                extension.GetUserPreferenceToggle(
                    identificador, OPCION_DETALLADO_SIN_ESPECIFICAR
                )
            )
            correcto = extension.SetUserPreferenceToggle(
                identificador,
                OPCION_DETALLADO_SIN_ESPECIFICAR,
                True,
            )
            if correcto is False:
                raise RuntimeError("SolidWorks devolvió False")
            cambios.append(
                {
                    "tipo": "booleano",
                    "nombre": nombre,
                    "identificador": identificador,
                    "original": original,
                    "nuevo": True,
                }
            )
        except Exception as error:
            avisos.append(
                f"no se pudo aplicar {nombre} ({identificador}): "
                f"{texto_error(error)}"
            )

    try:
        original = float(
            extension.GetUserPreferenceDouble(
                PREFERENCIA_TAMANO_MARCA_CENTRO,
                OPCION_DETALLADO_SIN_ESPECIFICAR,
            )
        )
        correcto = extension.SetUserPreferenceDouble(
            PREFERENCIA_TAMANO_MARCA_CENTRO,
            OPCION_DETALLADO_SIN_ESPECIFICAR,
            TAMANO_MARCA_CENTRO_M,
        )
        if correcto is False:
            raise RuntimeError("SolidWorks devolvió False")
        cambios.append(
            {
                "tipo": "doble",
                "nombre": "swDetailingCenterMarkSize",
                "identificador": PREFERENCIA_TAMANO_MARCA_CENTRO,
                "original": original,
                "nuevo": TAMANO_MARCA_CENTRO_M,
            }
        )
    except Exception as error:
        avisos.append(
            "no se pudo aplicar swDetailingCenterMarkSize (12): "
            f"{texto_error(error)}"
        )

    return extension, cambios, avisos


def restaurar_preferencias_anotaciones(extension, cambios):
    """Restaura las preferencias del plano sin ocultar fallos individuales."""
    avisos = []
    if extension is None:
        return avisos

    for cambio in reversed(cambios):
        try:
            if cambio["tipo"] == "booleano":
                correcto = extension.SetUserPreferenceToggle(
                    cambio["identificador"],
                    OPCION_DETALLADO_SIN_ESPECIFICAR,
                    cambio["original"],
                )
            else:
                correcto = extension.SetUserPreferenceDouble(
                    cambio["identificador"],
                    OPCION_DETALLADO_SIN_ESPECIFICAR,
                    cambio["original"],
                )
            if correcto is False:
                raise RuntimeError("SolidWorks devolvió False")
        except Exception as error:
            avisos.append(
                f"no se pudo restaurar {cambio['nombre']} "
                f"({cambio['identificador']}): {texto_error(error)}"
            )
    return avisos


def abrir_pieza(swapp, ruta):
    """Abre una pieza en silencio y devuelve documento, errores y avisos."""
    errores = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    avisos = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    documento = swapp.OpenDoc6(
        str(ruta),
        TIPO_DOCUMENTO_PIEZA,
        OPCION_ABRIR_SILENCIOSO,
        "",
        errores,
        avisos,
    )
    return documento, int(errores.value), int(avisos.value)


def obtener_plantilla_dibujo(swapp):
    """Obtiene la plantilla de dibujo predeterminada de SolidWorks."""
    # El método recibe un argumento y devuelve texto simple; no necesita leer().
    ruta = swapp.GetUserPreferenceStringValue(PREFERENCIA_PLANTILLA_PLANO)
    if not ruta:
        raise RuntimeError(
            "SolidWorks no tiene configurada una plantilla de dibujo "
            "predeterminada"
        )

    plantilla = Path(str(ruta))
    if not plantilla.is_file():
        raise FileNotFoundError(
            f"no existe la plantilla de dibujo predeterminada: {plantilla}"
        )
    return plantilla


def configurar_hoja(plano, formato, escala):
    """Aplica formato, tamaño, escala y proyección europea a la hoja actual."""
    # GetSheetNames devuelve cadenas simples, por eso leer() es seguro.
    nombres = leer(plano, "GetSheetNames")
    if isinstance(nombres, str):
        nombre_hoja = nombres
    elif nombres:
        nombre_hoja = nombres[0]
    else:
        raise RuntimeError("el plano nuevo no contiene ninguna hoja")

    ruta_formato = CARPETA_PLANTILLAS / formato["archivo"]
    correcto = plano.SetupSheet5(
        nombre_hoja,
        PAPEL_DE_USUARIO,
        PLANTILLA_PERSONALIZADA,
        escala["numerador"],
        escala["denominador"],
        True,
        str(ruta_formato),
        formato["ancho_m"],
        formato["alto_m"],
        "",
        False,
    )
    if not correcto:
        raise RuntimeError("IDrawingDoc.SetupSheet5 devolvió False")

    # La API recomienda reconstruir después de cambiar el tipo de proyección.
    plano.ForceRebuild3(False)


def _normalizar_nombre_vista(nombre):
    """Quita asterisco, acentos y diferencias de mayúsculas para comparar."""
    texto = str(nombre).strip().lstrip("*")
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(letra for letra in texto if not unicodedata.combining(letra))
    return texto.casefold()


def seleccionar_nombre_vista(nombres_disponibles, tipo_vista):
    """Encuentra el nombre real del alzado o de la isométrica."""
    nombres = [str(nombre) for nombre in (nombres_disponibles or [])]
    if not nombres:
        return None

    # Los nombres de vista de SolidWorks se traducen según el idioma de la
    # instalación. Nunca deben escribirse a mano: hay que obtenerlos con
    # IModelDoc2.GetModelViewNames() y emparejarlos de forma flexible.
    # En esta instalación en español, el alzado se llama "*Frontal".
    aceptados = {
        "alzado": {"front", "frontal", "alzado"},
        "isometrica": {"isometric", "isometrica"},
    }
    for nombre in nombres:
        if _normalizar_nombre_vista(nombre) in aceptados[tipo_vista]:
            return nombre

    raise ValueError(
        f"no se encontró la vista {tipo_vista}; "
        f"nombres disponibles: {nombres}"
    )


def obtener_nombres_vistas_modelo(pieza):
    """Lee los nombres de vista reales de una pieza como textos sencillos."""
    nombres = leer(pieza, "GetModelViewNames")
    if not nombres:
        return []
    if isinstance(nombres, str):
        return [nombres]
    return [str(nombre) for nombre in nombres]


def activar_plano(swapp, nombre_plano):
    """Activa el plano con ActivateDoc3 y usa ActivateDoc2 como alternativa."""
    errores = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    try:
        # ActivateDoc3 funciona en SolidWorks 2023. True aplica la preferencia
        # de reconstrucción del usuario y hace que el tercer argumento se ignore.
        documento = swapp.ActivateDoc3(nombre_plano, True, 0, errores)
        metodo = "ISldWorks.ActivateDoc3"
    except (AttributeError, TypeError):
        # Esta vía queda para instalaciones cuyo enlace COM no exponga Doc3.
        errores = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, 0
        )
        documento = swapp.ActivateDoc2(nombre_plano, True, errores)
        metodo = "ISldWorks.ActivateDoc2"

    if documento is None:
        raise RuntimeError(
            f"{metodo} no activó el plano; código={int(errores.value)}"
        )
    return documento, metodo, int(errores.value)


def _miembro_com_sin_argumentos(objeto, nombre):
    """Obtiene un objeto COM sin llamar por error a un CDispatch ya resuelto."""
    miembro = getattr(objeto, nombre)
    if hasattr(miembro, "_oleobj_") or not callable(miembro):
        return miembro
    return miembro()


def leer_dimensiones_hoja(plano):
    """Lee el ancho y alto reales de la hoja actual, en metros."""
    # GetCurrentSheet devuelve un objeto ISheet: se obtiene directamente y no
    # se pasa por leer(), porque un CDispatch también puede ser invocable.
    hoja = _miembro_com_sin_argumentos(plano, "GetCurrentSheet")
    if hoja is None:
        raise RuntimeError("el plano no tiene una hoja activa")

    # En SolidWorks 2023, ISheet.GetProperties2 devuelve ocho valores:
    # [paperSize, templateIn, scale1, scale2, firstAngle, width, height,
    # sameCustomProp]. Por eso el ancho está en el índice 5 y el alto en el 6.
    propiedades = leer(hoja, "GetProperties2")
    if propiedades is None or len(propiedades) < 7:
        raise RuntimeError("ISheet.GetProperties2 no devolvió siete valores")

    ancho_m = float(propiedades[5])
    alto_m = float(propiedades[6])
    if ancho_m <= 0 or alto_m <= 0:
        raise ValueError("ISheet.GetProperties2 devolvió un tamaño no válido")
    return ancho_m, alto_m


def _documentos_abiertos(swapp):
    """Devuelve los documentos COM abiertos sin usar leer() con sus objetos."""
    documentos = _miembro_com_sin_argumentos(swapp, "GetDocuments")
    if not documentos:
        return []
    if hasattr(documentos, "_oleobj_"):
        return [documentos]
    return list(documentos)


def imprimir_diagnostico(
    swapp,
    pieza,
    plano,
    ruta_modelo,
    metodo_activacion,
    codigo_activacion,
    ancho_estimado_m,
    alto_estimado_m,
    ancho_m,
    alto_m,
    disposicion,
    tamano_leido_api,
    cambios_anotaciones,
):
    """Imprime los datos COM necesarios para diagnosticar las vistas."""
    nombres_vistas = obtener_nombres_vistas_modelo(pieza)
    documentos = _documentos_abiertos(swapp)

    # ActiveDoc devuelve un objeto COM y debe leerse directamente.
    activo = swapp.ActiveDoc
    titulo_activo = leer(activo, "GetTitle") if activo is not None else "(ninguno)"

    datos_documentos = []
    for documento in documentos:
        titulo = str(leer(documento, "GetTitle") or "(sin título)")
        ruta = str(leer(documento, "GetPathName") or "")
        datos_documentos.append((titulo, ruta))

    nombres_hojas = leer(plano, "GetSheetNames")
    if isinstance(nombres_hojas, str):
        nombres_hojas = [nombres_hojas]
    else:
        nombres_hojas = list(nombres_hojas or [])

    # GetCurrentSheet devuelve ISheet: acceso directo, sin leer().
    hoja_activa = _miembro_com_sin_argumentos(plano, "GetCurrentSheet")
    nombre_hoja_activa = (
        str(leer(hoja_activa, "GetName"))
        if hoja_activa is not None
        else "(ninguna)"
    )

    coincidencia_exacta = any(
        ruta_modelo == ruta_abierta for _, ruta_abierta in datos_documentos
    )

    print("  --- DIAGNÓSTICO COM ANTES DE CREAR VISTAS ---")
    print(f"  Nombres de vista reales: {nombres_vistas}")
    print(f"  Documento activo: {titulo_activo}")
    print(
        f"  Activación del plano: {metodo_activacion}; "
        f"código={codigo_activacion}"
    )
    print("  Documentos abiertos:")
    if datos_documentos:
        for titulo, ruta_abierta in datos_documentos:
            ruta_visible = ruta_abierta or "(sin ruta: documento nuevo)"
            print(f"    - {titulo} | {ruta_visible}")
    else:
        print("    (ninguno)")
    print(f"  Ruta enviada a CreateDrawViewFromModelView*: {ruta_modelo}")
    print(
        "  Coincide exactamente con una ruta abierta: "
        f"{'SÍ' if coincidencia_exacta else 'NO'}"
    )
    print(f"  Hojas del plano: {nombres_hojas}")
    print(f"  Hoja activa: {nombre_hoja_activa}")
    if tamano_leido_api:
        print(
            f"  Tamaño de hoja leído de la API: "
            f"ancho={ancho_m:.3f} m, alto={alto_m:.3f} m"
        )
    else:
        print("  Tamaño de hoja leído de la API: no disponible")

    tamano_distinto = (
        abs(ancho_m - ancho_estimado_m) > 0.000001
        or abs(alto_m - alto_estimado_m) > 0.000001
    )
    if tamano_leido_api and tamano_distinto:
        print(
            "AVISO: el tamaño leído no coincide con el supuesto "
            f"({ancho_estimado_m:.3f} x {alto_estimado_m:.3f} m)."
        )

    area = disposicion["area_util"]
    alzado = disposicion["tamanos"]["alzado"]
    isometrica = disposicion["tamanos"]["isometrica"]
    coordenadas_alzado = disposicion["coordenadas_alzado_m"]
    coordenadas_isometrica = disposicion["coordenadas_isometrica_m"]
    rectangulo_alzado = disposicion["rectangulos"]["alzado"]
    rectangulo_isometrica = disposicion["rectangulos"]["isometrica"]
    print(
        f"  Área útil: {area['ancho_mm']:.1f} x "
        f"{area['alto_mm']:.1f} mm"
    )
    print(
        f"  Centro del área útil: x={area['centro_x_mm']:.1f} mm, "
        f"y={area['centro_y_mm']:.1f} mm"
    )
    print(
        f"  Tamaño calculado del alzado: {alzado['ancho_mm']:.1f} x "
        f"{alzado['alto_mm']:.1f} mm"
    )
    print(
        f"  Tamaño calculado de la isométrica: "
        f"{isometrica['ancho_mm']:.1f} x {isometrica['alto_mm']:.1f} mm"
    )
    print(f"  Disposición elegida: {disposicion['tipo']}")
    print(
        "  Rectángulo del alzado: "
        f"izquierda={rectangulo_alzado['izquierda_mm']:.1f}, "
        f"inferior={rectangulo_alzado['inferior_mm']:.1f}, "
        f"derecha={rectangulo_alzado['derecha_mm']:.1f}, "
        f"superior={rectangulo_alzado['superior_mm']:.1f} mm"
    )
    print(
        "  Rectángulo de la isométrica: "
        f"izquierda={rectangulo_isometrica['izquierda_mm']:.1f}, "
        f"inferior={rectangulo_isometrica['inferior_mm']:.1f}, "
        f"derecha={rectangulo_isometrica['derecha_mm']:.1f}, "
        f"superior={rectangulo_isometrica['superior_mm']:.1f} mm"
    )
    print(
        "  Coordenadas finales: "
        f"alzado=(x={coordenadas_alzado[0]:.4f}, "
        f"y={coordenadas_alzado[1]:.4f}) m; "
        f"isométrica=(x={coordenadas_isometrica[0]:.4f}, "
        f"y={coordenadas_isometrica[1]:.4f}) m"
    )
    print("  Preferencias de anotación modificadas:")
    if cambios_anotaciones:
        for cambio in cambios_anotaciones:
            if cambio["tipo"] == "doble":
                original = f"{cambio['original'] * 1000.0:.3f} mm"
                nuevo = f"{cambio['nuevo'] * 1000.0:.3f} mm"
            else:
                original = str(cambio["original"])
                nuevo = str(cambio["nuevo"])
            print(
                f"    {cambio['nombre']} ({cambio['identificador']}): "
                f"original={original}; temporal={nuevo}"
            )
    else:
        print("    (ninguna; véanse los avisos anteriores)")
    print("  --- FIN DEL DIAGNÓSTICO COM ---")


def crear_vista_localizada(plano, pieza, ruta_modelo, tipo_vista, x, y):
    """Crea una vista usando su nombre real y dos versiones de la API."""
    nombres_disponibles = obtener_nombres_vistas_modelo(pieza)
    nombre_real = seleccionar_nombre_vista(nombres_disponibles, tipo_vista)

    if nombre_real is not None:
        candidatos = [nombre_real]
    elif tipo_vista == "alzado":
        candidatos = ["*Front", "*Alzado"]
    else:
        candidatos = ["*Isometric", "*Isométrica"]

    intentos = []
    for nombre in candidatos:
        for nombre_metodo in (
            "CreateDrawViewFromModelView3",
            "CreateDrawViewFromModelView2",
        ):
            try:
                # Ambos métodos devuelven IView: acceso COM directo, sin leer().
                metodo = getattr(plano, nombre_metodo)
                vista = metodo(str(ruta_modelo), nombre, x, y, 0.0)
                if vista is not None:
                    return vista, nombre, f"IDrawingDoc.{nombre_metodo}"
                intentos.append(
                    f"{nombre_metodo}({nombre}): no devolvió una vista"
                )
            except (pywintypes.com_error, AttributeError) as error:
                intentos.append(
                    f"{nombre_metodo}({nombre}): {texto_error_com(error)}"
                )
            except Exception as error:
                intentos.append(
                    f"{nombre_metodo}({nombre}): {texto_error(error)}"
                )

    raise RuntimeError("; ".join(intentos))


def guardar_plano(plano, ruta_salida):
    """Guarda el plano con la vía comprobada en SolidWorks 2023."""
    # IModelDocExtension.SaveAs falla en esta instalación con enlace tardío
    # por el argumento opcional de exportación. IModelDoc2.SaveAs ya funciona
    # en exportar_pdf.py y evita pasar un None crudo a COM.
    correcto = plano.SaveAs(str(ruta_salida))
    if not correcto:
        raise RuntimeError("IModelDoc2.SaveAs devolvió False")

    # Esta llamada sencilla no devuelve códigos separados de error y aviso.
    return 0, 0


def _registrar_fallo(resultado, mensaje):
    """Añade un detalle a un resultado y lo marca como fallido."""
    resultado["correcto"] = False
    resultado["error"] = "; ".join(
        texto for texto in (resultado["error"], mensaje) if texto
    )


def procesar_pieza(swapp, ruta, rutas_abiertas_antes, diagnostico=False):
    """Genera y guarda el plano de una pieza sin propagar sus errores."""
    resultado = {
        "archivo": ruta.name,
        "hoja": "-",
        "escala": "-",
        "vistas": 0,
        "nombres_vista": [],
        "metodos_vista": [],
        "correcto": False,
        "avisos": [],
        "error": "",
    }
    pieza = None
    plano = None
    nombre_pieza = ruta.name
    nombre_plano = ""
    autoescala_original = None
    autoescala_modificada = False
    extension_anotaciones = None
    cambios_anotaciones = []
    preferencias_anotaciones_pendientes = False
    pieza_ya_abierta = False
    ruta_salida = CARPETA_SALIDA / f"{ruta.stem}.SLDDRW"

    try:
        pieza_ya_abierta = normalizar_ruta(ruta) in rutas_abiertas_antes
        if normalizar_ruta(ruta_salida) in rutas_abiertas_antes:
            raise RuntimeError(
                f"el plano de destino ya está abierto: {ruta_salida}"
            )

        pieza, errores, avisos = abrir_pieza(swapp, ruta)
        if pieza is None:
            raise RuntimeError(
                f"OpenDoc6 no abrió la pieza; errores={errores}, avisos={avisos}"
            )
        nombre_pieza = leer(pieza, "GetTitle") or ruta.name
        # La ruta devuelta por el documento abierto conserva exactamente el
        # texto que SolidWorks reconoce. Es la que se enviará al crear vistas.
        ruta_modelo = str(leer(pieza, "GetPathName") or ruta)
        if errores or avisos:
            resultado["avisos"].append(
                f"OpenDoc6: errores={errores}, avisos={avisos}"
            )

        espesor, medio, largo = dimensiones_pieza(pieza)
        formato, escala, _ = elegir_formato_y_escala(largo, medio)
        resultado["hoja"] = formato["nombre"]
        resultado["escala"] = escala["texto"]

        plantilla_dibujo = obtener_plantilla_dibujo(swapp)
        # NewDocument devuelve IModelDoc2: se accede directamente, sin leer().
        plano = swapp.NewDocument(str(plantilla_dibujo), 0, 0.0, 0.0)
        if plano is None:
            raise RuntimeError("ISldWorks.NewDocument no creó el plano")
        nombre_plano = leer(plano, "GetTitle") or ""

        configurar_hoja(plano, formato, escala)

        ancho_estimado_m = formato["ancho_m"]
        alto_estimado_m = formato["alto_m"]
        tamano_leido_api = True
        try:
            # El formato .slddrt puede sustituir el tamaño solicitado. Se lee
            # la hoja ya configurada antes de calcular cualquier coordenada.
            ancho_m, alto_m = leer_dimensiones_hoja(plano)
        except Exception as error:
            tamano_leido_api = False
            ancho_m = ancho_estimado_m
            alto_m = alto_estimado_m
            aviso_tamano = (
                "no se pudo leer el tamaño real de la hoja; se usará el "
                f"estimado {ancho_m:.3f} x {alto_m:.3f} m: "
                f"{texto_error(error)}"
            )
            resultado["avisos"].append(aviso_tamano)
            print(f"AVISO: {aviso_tamano}")

        # La selección inicial usa las medidas declaradas. Aquí se repite el
        # cálculo con la medida real que impuso el formato .slddrt.
        tamanos_vistas = calcular_tamanos_vistas(largo, medio, escala)
        disposicion = calcular_disposicion(
            ancho_m,
            alto_m,
            tamanos_vistas,
        )
        coordenadas_alzado = disposicion["coordenadas_alzado_m"]
        coordenadas_isometrica = disposicion["coordenadas_isometrica_m"]

        # Estas opciones pertenecen al documento de plano. Se activan solo
        # durante la creación de las vistas y se restauran antes de guardar.
        (
            extension_anotaciones,
            cambios_anotaciones,
            avisos_preferencias,
        ) = aplicar_preferencias_anotaciones(plano)
        preferencias_anotaciones_pendientes = bool(cambios_anotaciones)
        for aviso in avisos_preferencias:
            resultado["avisos"].append(aviso)
            print(f"AVISO: {aviso}")

        # CreateDrawViewFromModelView* exige que el plano de destino sea el
        # documento activo. En SolidWorks 2023 se prueba primero ActivateDoc3.
        _, metodo_activacion, codigo_activacion = activar_plano(
            swapp, nombre_plano
        )
        if codigo_activacion:
            resultado["avisos"].append(
                f"{metodo_activacion}: código={codigo_activacion}"
            )

        if diagnostico:
            imprimir_diagnostico(
                swapp,
                pieza,
                plano,
                ruta_modelo,
                metodo_activacion,
                codigo_activacion,
                ancho_estimado_m,
                alto_estimado_m,
                ancho_m,
                alto_m,
                disposicion,
                tamano_leido_api,
                cambios_anotaciones,
            )

        # SolidWorks puede cambiar la escala al insertar la primera vista.
        # Desactivamos esa preferencia solo durante estas dos inserciones.
        autoescala_original = bool(
            swapp.GetUserPreferenceToggle(PREFERENCIA_AUTOESCALA_VISTAS)
        )
        # SetUserPreferenceToggle es un procedimiento COM: si funciona no
        # devuelve True, sino None. Un fallo real llega como excepción.
        swapp.SetUserPreferenceToggle(PREFERENCIA_AUTOESCALA_VISTAS, False)
        autoescala_modificada = True

        _, nombre_alzado, metodo_alzado = crear_vista_localizada(
            plano,
            pieza,
            ruta_modelo,
            "alzado",
            coordenadas_alzado[0],
            coordenadas_alzado[1],
        )
        resultado["vistas"] += 1
        resultado["nombres_vista"].append(nombre_alzado)
        resultado["metodos_vista"].append(metodo_alzado)
        if diagnostico:
            print(
                f"  Vista alzado creada con {metodo_alzado}: "
                f"{nombre_alzado}"
            )

        _, nombre_isometrica, metodo_isometrica = crear_vista_localizada(
            plano,
            pieza,
            ruta_modelo,
            "isometrica",
            coordenadas_isometrica[0],
            coordenadas_isometrica[1],
        )
        resultado["vistas"] += 1
        resultado["nombres_vista"].append(nombre_isometrica)
        resultado["metodos_vista"].append(metodo_isometrica)
        if diagnostico:
            print(
                f"  Vista isométrica creada con {metodo_isometrica}: "
                f"{nombre_isometrica}"
            )

        avisos_restauracion = restaurar_preferencias_anotaciones(
            extension_anotaciones, cambios_anotaciones
        )
        preferencias_anotaciones_pendientes = bool(avisos_restauracion)
        for aviso in avisos_restauracion:
            resultado["avisos"].append(aviso)
            print(f"AVISO: {aviso}")

        errores_guardado, avisos_guardado = guardar_plano(plano, ruta_salida)
        if errores_guardado or avisos_guardado:
            resultado["avisos"].append(
                "Guardado: "
                f"errores={errores_guardado}, avisos={avisos_guardado}"
            )
        resultado["correcto"] = True
    except (pywintypes.com_error, AttributeError) as error:
        _registrar_fallo(resultado, texto_error_com(error))
    except Exception as error:
        _registrar_fallo(resultado, texto_error(error))
    finally:
        # Si hubo un error antes de la restauración normal, se intenta aquí
        # antes de cerrar el plano. Cada fallo se informa, pero no oculta otro.
        if preferencias_anotaciones_pendientes:
            avisos_restauracion = restaurar_preferencias_anotaciones(
                extension_anotaciones, cambios_anotaciones
            )
            for aviso in avisos_restauracion:
                resultado["avisos"].append(aviso)
                print(f"AVISO: {aviso}")

        if autoescala_modificada:
            try:
                swapp.SetUserPreferenceToggle(
                    PREFERENCIA_AUTOESCALA_VISTAS, autoescala_original
                )
            except Exception as error:
                resultado["avisos"].append(
                    "no se pudo restaurar la escala automática de vistas: "
                    f"{texto_error(error)}"
                )

        if plano is not None:
            try:
                nombre_plano = leer(plano, "GetTitle") or nombre_plano
                swapp.CloseDoc(nombre_plano)
            except Exception as error:
                _registrar_fallo(
                    resultado, f"no se pudo cerrar el plano: {texto_error(error)}"
                )

        # Una pieza que ya estaba abierta pertenece al usuario y no se cierra.
        if pieza is not None and not pieza_ya_abierta:
            try:
                swapp.CloseDoc(nombre_pieza)
            except Exception as error:
                _registrar_fallo(
                    resultado, f"no se pudo cerrar la pieza: {texto_error(error)}"
                )

    return resultado


def imprimir_resultado(resultado):
    """Muestra una línea principal y, si hace falta, sus detalles legibles."""
    estado = "OK" if resultado["correcto"] else "FALLO"
    nombres = ",".join(resultado["nombres_vista"]) or "-"
    print(
        f"{estado} | {resultado['archivo']} | hoja={resultado['hoja']} | "
        f"escala={resultado['escala']} | vistas={resultado['vistas']} | "
        f"nombres={nombres}"
    )
    if resultado["error"]:
        print(f"  Error: {resultado['error']}")
    for aviso in resultado["avisos"]:
        print(f"  Aviso: {aviso}")


def mostrar_uso():
    """Explica las entradas aceptadas por el script."""
    print(
        "Uso: motor_plano.py [--diagnostico] <pieza.SLDPRT o carpeta> "
        "[otra pieza o carpeta ...]"
    )


def analizar_argumentos(argumentos):
    """Separa la opción de diagnóstico de las rutas de entrada."""
    diagnostico = False
    rutas = []
    for argumento in argumentos:
        if argumento == "--diagnostico":
            diagnostico = True
        else:
            rutas.append(argumento)
    return diagnostico, rutas


def main():
    """Coordina la generación de planos sin dejar tracebacks crudos."""
    try:
        # Evita UnicodeEncodeError en consolas Windows con páginas antiguas.
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    diagnostico, argumentos_ruta = analizar_argumentos(sys.argv[1:])
    if not argumentos_ruta:
        mostrar_uso()
        return 2

    try:
        if not validar_plantillas():
            return 1

        entradas = [
            Path(argumento).expanduser().resolve()
            for argumento in argumentos_ruta
        ]
        swapp = conectar()
        if swapp is None:
            return 1

        rutas_abiertas_antes = obtener_rutas_documentos_abiertos(swapp)
        piezas = buscar_piezas(entradas)
        CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
        print(f"Piezas encontradas: {len(piezas)}")

        resultados = []
        for ruta in piezas:
            resultado = procesar_pieza(
                swapp,
                ruta,
                rutas_abiertas_antes,
                diagnostico=diagnostico,
            )
            resultados.append(resultado)
            imprimir_resultado(resultado)

        generados = sum(resultado["correcto"] for resultado in resultados)
        fallidos = len(resultados) - generados
        print(f"Resumen: {generados} planos generados; {fallidos} fallidos.")
        print(f"Carpeta de salida: {CARPETA_SALIDA}")
        return 0
    except Exception as error:
        print("No se pudo completar la generación de planos.")
        print(f"Detalle del error: {texto_error(error)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
