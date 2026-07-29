r"""Crea planos de SolidWorks con alzado, isométrica y sección opcional.

Uso: .venv\Scripts\python.exe src\motor_plano.py "C:\ruta\pieza.SLDPRT"
"""

import sys
import unicodedata
from math import sqrt
from pathlib import Path
from statistics import median

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

# Medidas observadas de los formatos, en metros, usadas como estimación y
# alternativa. El .slddrt puede imponer otras medidas u orientación:
# la verdad será siempre lo que devuelva ISheet después de aplicarlo.
ANCHO_A4_M = 0.210  # A4 vertical observado
ALTO_A4_M = 0.297
ANCHO_A3_M = 0.420
ALTO_A3_M = 0.297  # A3 apaisado observado; pendiente de confirmar por API
ANCHO_A2_M = 0.594
ALTO_A2_M = 0.420  # A2 apaisado observado; pendiente de confirmar por API

# A3 es el formato normal del proyecto. A4 y A2 se conservan únicamente para
# poder forzarlos de forma puntual con la opción --formato.
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
FORMATOS_POR_NOMBRE = {
    formato["nombre"].casefold(): formato for formato in FORMATOS_HOJA
}
FORMATO_POR_DEFECTO = "a3"

# Primero se prueban las escalas de reducción, de más legible a menos.
ESCALAS_REDUCCION = [
    {"texto": "1:1", "numerador": 1.0, "denominador": 1.0},
    {"texto": "1:2", "numerador": 1.0, "denominador": 2.0},
    {"texto": "1:5", "numerador": 1.0, "denominador": 5.0},
    {"texto": "1:10", "numerador": 1.0, "denominador": 10.0},
    {"texto": "1:20", "numerador": 1.0, "denominador": 20.0},
]
ESCALAS_AMPLIACION = [
    {"texto": "2:1", "numerador": 2.0, "denominador": 1.0},
    {"texto": "5:1", "numerador": 5.0, "denominador": 1.0},
    {"texto": "10:1", "numerador": 10.0, "denominador": 1.0},
]
# La isométrica prueba de la mayor ampliación a la mayor reducción y conserva
# la primera escala que cabe. El alzado aplica más abajo su regla centrada en 1:1.
ESCALAS_NORMALIZADAS = list(reversed(ESCALAS_AMPLIACION)) + ESCALAS_REDUCCION

# Margen de encuadernación: el marco de la plantilla está más adentro
# por el lado izquierdo.
MARGEN_IZQUIERDO_MM = 30.0
MARGEN_DERECHO_MM = 20.0
# El marco superior también está más adentro en estas plantillas.
MARGEN_SUPERIOR_MM = 20.0
BANDA_CAJETIN_MM = 70.0
# El alzado se eleva ligeramente respecto al centro del espacio disponible.
DESPLAZAMIENTO_VERTICAL_ALZADO = 0.08  # 8 % del alto útil.

# Maqueta fija preparada para tres vistas: alzado a la izquierda e isométrica
# arriba a la derecha. La fila inferior derecha se reserva para la sección.
PROPORCION_COLUMNA_ALZADO = 0.55
PROPORCION_COLUMNA_DERECHA = 0.45
PROPORCION_FILA_DERECHA = 0.50
# 1:1 es la escala preferente. Solo se amplía cuando la vista ocupa menos del
# 25 % y se lee mal, o se reduce cuando supera el techo del 60 %.
OCUPACION_MINIMA_ALZADO = 0.25
# El 60 % es un techo, no un objetivo: el resto de la celda se reserva para
# cotas y para vistas todavía no dibujadas, como la sección y quizá otra más.
OCUPACION_MAXIMA_ALZADO = 0.60

# La proyección isométrica de una placa se aproxima como un rectángulo.
# El ancho usa aproximadamente cos(30°) y el alto sen(30°), con un pequeño
# margen por el espesor. Son factores ajustables al comparar planos reales.
FACTOR_ANCHO_ISOMETRICA = 0.87
FACTOR_ALTO_ISOMETRICA = 0.55
# La isométrica no lleva cotas, por eso puede ocupar más celda que el alzado.
OCUPACION_MAXIMA_ISOMETRICA = 0.80
# La línea de corte sobresale 5 mm por cada lado del alzado para que las
# flechas no queden pegadas a la geometría. Se podrá ajustar tras verla.
MARGEN_LINEA_CORTE_MM = 5.0
ALTURA_CORTE_POR_DEFECTO = 0.0
MODO_CORTE_POR_DEFECTO = "auto"
MODOS_CORTE_VALIDOS = ("auto", "horizontal", "vertical")
# Dos caras se consideran el mismo taladro si sus ejes y radios coinciden
# dentro de 0,1 mm. Es una tolerancia geométrica, no de fabricación.
TOLERANCIA_AGRUPAR_TALADROS_M = 0.0001
# La misma tolerancia agrupa los centros que forman una fila o columna.
TOLERANCIA_COLUMNA_TALADROS_M = 0.0001
# Se admite 1 mm numérico al comparar la proyección con el rectángulo
# calculado del alzado, usando únicamente la aritmética de caja y escala.
TOLERANCIA_PROYECCION_MM = 1.0
# El eje de un taladro pasante de placa debe coincidir prácticamente con el
# eje de espesor. Este coseno admite pequeñas imprecisiones numéricas.
ALINEACION_MINIMA_EJE_TALADRO = 0.999
# El cilindro central grande no es un agujero de fijación. Solo se excluye si
# su radio supera claramente (más de 3 veces) la mediana de los demás.
FACTOR_RADIO_AGUJERO_CENTRAL = 3.0
# CreateSectionViewAt5 trabaja en silencio. Con options=0 crea una sección
# completa, alineada y con la dirección predeterminada, sin cuadro de "flip".
OPCIONES_SECCION_NORMAL = 0
OPCIONES_SECCION_ESCALA_MODELO = 8  # swCreateSectionView_ScaleWithModel
# Un corte vertical se proyecta hacia un lado. NotAligned evita que
# SolidWorks ignore el centro de la celda inferior derecha.
OPCIONES_SECCION_NO_ALINEADA = 1  # swCreateSectionView_NotAligned
OPCIONES_SECCION_NO_ALINEADA_ESCALA_MODELO = 9  # NotAligned | ScaleWithModel
PROFUNDIDAD_SECCION_COMPLETA_M = 0.0
ETIQUETA_SECCION = "A"
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
# Esta es la única preferencia doble que modifica el motor. Como representa
# una longitud, SolidWorks espera metros aunque la opción de usuario use mm.
PREFERENCIA_TAMANO_MARCA_CENTRO = 12
TAMANO_MARCA_CENTRO_MM_POR_DEFECTO = 2.0
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


def _normalizar_eje(eje):
    """Normaliza un eje y fija su signo para poder comparar cilindros."""
    longitud = sqrt(sum(componente * componente for componente in eje))
    if longitud <= 0.0:
        raise ValueError("una superficie cilíndrica tiene eje nulo")
    eje = tuple(componente / longitud for componente in eje)
    for componente in eje:
        if abs(componente) > 0.000000001:
            if componente < 0.0:
                eje = tuple(-valor for valor in eje)
            break
    return eje


def _centro_canonico_eje(origen, eje, referencia):
    """Proyecta una referencia sobre el eje para obtener un punto único."""
    avance = sum(
        (referencia[i] - origen[i]) * eje[i] for i in range(3)
    )
    return tuple(origen[i] + avance * eje[i] for i in range(3))


def _son_el_mismo_cilindro(primero, segundo):
    """Compara eje, centro y radio con una tolerancia de 0,1 mm."""
    diferencia_eje = sqrt(
        sum(
            (primero["eje"][i] - segundo["eje"][i]) ** 2
            for i in range(3)
        )
    )
    diferencia_centro = sqrt(
        sum(
            (primero["centro"][i] - segundo["centro"][i]) ** 2
            for i in range(3)
        )
    )
    diferencia_radio = abs(primero["radio"] - segundo["radio"])
    return (
        diferencia_eje <= 0.000001
        and diferencia_centro <= TOLERANCIA_AGRUPAR_TALADROS_M
        and diferencia_radio <= TOLERANCIA_AGRUPAR_TALADROS_M
    )


def detectar_taladros(pieza):
    """Devuelve los cilindros distintos que representan taladros pequeños."""
    # GetBodies2, GetFaces y GetSurface devuelven objetos COM. Por eso se
    # accede a ellos directamente y nunca mediante leer().
    cuerpos = pieza.GetBodies2(TIPO_CUERPO_SOLIDO, False)
    if not cuerpos:
        return []
    caja = _obtener_caja(pieza)
    centro_pieza = (
        (caja[0] + caja[3]) / 2.0,
        (caja[1] + caja[4]) / 2.0,
        (caja[2] + caja[5]) / 2.0,
    )

    cilindros = []
    for cuerpo in cuerpos:
        caras = cuerpo.GetFaces()
        if not caras:
            continue
        for cara in caras:
            # GetSurface se comporta como PROPIEDAD en este late binding
            # (devuelve ya el ISurface). Ponerle () llamaría al objeto COM
            # devuelto y daría DISP_E_MEMBERNOTFOUND. Confirmado con sonda.
            superficie = cara.GetSurface
            if superficie is None or not leer(superficie, "IsCylinder"):
                continue
            # En una cara interior (agujero), la normal de la cara se opone a
            # la normal natural del cilindro. Así se descartan redondeos y
            # paredes cilíndricas exteriores antes de medir su eje.
            if not leer(cara, "FaceInSurfaceSense"):
                continue

            # SolidWorks 2023 devuelve exactamente:
            # [origenX, origenY, origenZ, ejeX, ejeY, ejeZ, radio].
            # El origen y el radio están expresados en metros.
            parametros = leer(superficie, "CylinderParams")
            if parametros is None or len(parametros) != 7:
                continue
            origen = tuple(float(valor) for valor in parametros[0:3])
            eje = _normalizar_eje(
                tuple(float(valor) for valor in parametros[3:6])
            )
            radio = abs(float(parametros[6]))
            if radio <= 0.0:
                continue

            cilindro = {
                # CylinderParams da un punto cualquiera del eje. El centro de
                # la caja se proyecta sobre él para situar el resultado en el
                # plano medio de la pieza; su proyección 2D no cambia.
                "centro": _centro_canonico_eje(
                    origen, eje, centro_pieza
                ),
                "eje": eje,
                "radio": radio,
            }
            if not any(
                _son_el_mismo_cilindro(cilindro, existente)
                for existente in cilindros
            ):
                cilindros.append(cilindro)

    if len(cilindros) > 1:
        radio_mediano = median(
            cilindro["radio"] for cilindro in cilindros
        )
        radio_mayor = max(cilindro["radio"] for cilindro in cilindros)
        if radio_mayor > FACTOR_RADIO_AGUJERO_CENTRAL * radio_mediano:
            cilindros = [
                cilindro
                for cilindro in cilindros
                if abs(cilindro["radio"] - radio_mayor)
                > TOLERANCIA_AGRUPAR_TALADROS_M
            ]

    return cilindros


def medir_pieza_y_elegir_alzado(pieza, caja=None):
    """Mide la pieza y elige la vista estándar que muestra su cara grande."""
    if caja is None:
        caja = _obtener_caja(pieza)
    xmin, ymin, zmin, xmax, ymax, zmax = caja
    extensiones = {
        "X": abs(xmax - xmin) * 1000.0,
        "Y": abs(ymax - ymin) * 1000.0,
        "Z": abs(zmax - zmin) * 1000.0,
    }
    dimensiones = sorted(extensiones.values())
    if dimensiones[0] <= 0:
        raise ValueError("la caja envolvente tiene una dimensión igual a cero")

    # En una placa, el eje de la extensión menor es perpendicular a su cara
    # grande. Por eso Z pide la vista frontal, Y la superior y X la derecha.
    # Se conserva además el orden de tamaños: las otras dos extensiones son
    # las que determinan cuánto ocupa el alzado sobre el papel.
    eje_espesor = min(extensiones, key=extensiones.get)
    vista_por_eje = {
        "Z": "frontal",
        "Y": "superior",
        "X": "derecha",
    }
    espesor, medio, largo = dimensiones
    return espesor, medio, largo, vista_por_eje[eje_espesor], eje_espesor


def dimensiones_pieza(pieza):
    """Devuelve espesor, dimensión media y largo de la pieza en milímetros."""
    espesor, medio, largo, _, _ = medir_pieza_y_elegir_alzado(pieza)
    return espesor, medio, largo


def calcular_tamanos_vistas(largo_mm, medio_mm, escala):
    """Estima en milímetros cuánto ocupará cada vista sobre el papel."""
    factor_escala = escala["numerador"] / escala["denominador"]
    ancho_isometrica = (
        FACTOR_ANCHO_ISOMETRICA * (largo_mm + medio_mm) * factor_escala
    )
    alto_isometrica = (
        FACTOR_ALTO_ISOMETRICA * (largo_mm + medio_mm) * factor_escala
    )
    return {
        "alzado": {
            "ancho_mm": largo_mm * factor_escala,
            "alto_mm": medio_mm * factor_escala,
        },
        "isometrica": {
            "ancho_mm": ancho_isometrica,
            "alto_mm": alto_isometrica,
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


def calcular_celdas(area):
    """Divide el área útil en la maqueta fija de alzado, iso y sección."""
    ancho_alzado = area["ancho_mm"] * PROPORCION_COLUMNA_ALZADO
    ancho_derecha = area["ancho_mm"] * PROPORCION_COLUMNA_DERECHA
    alto_fila = area["alto_mm"] * PROPORCION_FILA_DERECHA
    division_x = area["izquierda_mm"] + ancho_alzado
    division_y = area["inferior_mm"] + alto_fila

    return {
        "alzado": {
            "ancho_mm": ancho_alzado,
            "alto_mm": area["alto_mm"],
            "izquierda_mm": area["izquierda_mm"],
            "derecha_mm": division_x,
            "inferior_mm": area["inferior_mm"],
            "superior_mm": area["superior_mm"],
        },
        "isometrica": {
            "ancho_mm": ancho_derecha,
            "alto_mm": alto_fila,
            "izquierda_mm": division_x,
            "derecha_mm": area["derecha_mm"],
            "inferior_mm": division_y,
            "superior_mm": area["superior_mm"],
        },
        # Esta celda se deja vacía a propósito: está reservada para la futura
        # sección y no debe ocuparse con ninguna otra vista.
        "seccion_reservada": {
            "ancho_mm": ancho_derecha,
            "alto_mm": alto_fila,
            "izquierda_mm": division_x,
            "derecha_mm": area["derecha_mm"],
            "inferior_mm": area["inferior_mm"],
            "superior_mm": division_y,
        },
    }


def _centro_celda(celda):
    """Devuelve el centro de una celda en milímetros de hoja."""
    return (
        (celda["izquierda_mm"] + celda["derecha_mm"]) / 2.0,
        (celda["inferior_mm"] + celda["superior_mm"]) / 2.0,
    )


def calcular_disposicion(ancho_hoja_m, alto_hoja_m, tamanos):
    """Centra cada vista en su celda fija de la maqueta."""
    area = calcular_area_util(ancho_hoja_m, alto_hoja_m)
    celdas = calcular_celdas(area)
    centro_alzado = _centro_celda(celdas["alzado"])
    centro_isometrica = _centro_celda(celdas["isometrica"])
    centro_seccion = _centro_celda(celdas["seccion_reservada"])

    coordenadas_alzado_mm = (
        centro_alzado[0],
        centro_alzado[1]
        + DESPLAZAMIENTO_VERTICAL_ALZADO * area["alto_mm"],
    )
    coordenadas_isometrica_mm = centro_isometrica

    rectangulos = {
        "alzado": calcular_rectangulo_vista(
            coordenadas_alzado_mm, tamanos["alzado"]
        ),
        "isometrica": calcular_rectangulo_vista(
            coordenadas_isometrica_mm, tamanos["isometrica"]
        ),
    }

    return {
        "tipo": "maqueta fija: alzado izquierda, isométrica arriba derecha",
        "area_util": area,
        "celdas": celdas,
        "tamanos": tamanos,
        "rectangulos": rectangulos,
        "coordenadas_alzado_m": tuple(
            valor / 1000.0 for valor in coordenadas_alzado_mm
        ),
        "coordenadas_isometrica_m": tuple(
            valor / 1000.0 for valor in coordenadas_isometrica_mm
        ),
        "coordenadas_seccion_m": tuple(
            valor / 1000.0 for valor in centro_seccion
        ),
    }


def calcular_linea_corte(disposicion, escala_alzado, altura_corte):
    """Calcula la línea horizontal de corte sin consultar la geometría COM."""
    centro_x_m, centro_y_m = disposicion["coordenadas_alzado_m"]
    tamano_alzado = disposicion["tamanos"]["alzado"]
    semiancho_papel_m = tamano_alzado["ancho_mm"] / 2000.0
    semialto_papel_m = tamano_alzado["alto_mm"] / 2000.0
    margen_papel_m = MARGEN_LINEA_CORTE_MM / 1000.0

    y_hoja_m = centro_y_m + altura_corte * semialto_papel_m
    x1_hoja_m = centro_x_m - semiancho_papel_m - margen_papel_m
    x2_hoja_m = centro_x_m + semiancho_papel_m + margen_papel_m

    factor_escala = (
        escala_alzado["numerador"] / escala_alzado["denominador"]
    )
    # Al activar una vista de dibujo, CreateLine usa coordenadas locales de
    # esa vista, no coordenadas de hoja. Se elimina aquí el factor de escala.
    semiancho_modelo_m = semiancho_papel_m / factor_escala
    margen_modelo_m = margen_papel_m / factor_escala
    y_modelo_m = altura_corte * semialto_papel_m / factor_escala

    return {
        "orientacion": "horizontal",
        "hoja_m": (x1_hoja_m, y_hoja_m, x2_hoja_m, y_hoja_m),
        "vista_m": (
            -semiancho_modelo_m - margen_modelo_m,
            y_modelo_m,
            semiancho_modelo_m + margen_modelo_m,
            y_modelo_m,
        ),
    }


def proyectar_taladros(
    taladros, eje_espesor, caja, disposicion, escala_alzado
):
    """Proyecta centros del modelo a la hoja y devuelve avisos geométricos."""
    ejes_vista = {
        "Z": (0, 1),  # Frontal: horizontal X, vertical Y.
        "Y": (0, 2),  # Superior: horizontal X, vertical Z.
        "X": (1, 2),  # Derecha: horizontal Y, vertical Z.
    }
    indice_h, indice_v = ejes_vista[eje_espesor]
    indice_espesor = {"X": 0, "Y": 1, "Z": 2}[eje_espesor]
    centro_modelo = (
        (caja[0] + caja[3]) / 2.0,
        (caja[1] + caja[4]) / 2.0,
        (caja[2] + caja[5]) / 2.0,
    )
    h_centro = centro_modelo[indice_h]
    v_centro = centro_modelo[indice_v]
    x_centro, y_centro = disposicion["coordenadas_alzado_m"]
    factor = escala_alzado["numerador"] / escala_alzado["denominador"]
    rectangulo = disposicion["rectangulos"]["alzado"]

    proyectados = []
    avisos = []
    for taladro in taladros:
        if (
            abs(taladro["eje"][indice_espesor])
            < ALINEACION_MINIMA_EJE_TALADRO
        ):
            continue
        h = taladro["centro"][indice_h]
        v = taladro["centro"][indice_v]
        x_hoja = x_centro + (h - h_centro) * factor
        y_hoja = y_centro + (v - v_centro) * factor
        x_mm = x_hoja * 1000.0
        y_mm = y_hoja * 1000.0
        plausible = (
            rectangulo["izquierda_mm"] - TOLERANCIA_PROYECCION_MM
            <= x_mm
            <= rectangulo["derecha_mm"] + TOLERANCIA_PROYECCION_MM
            and rectangulo["inferior_mm"] - TOLERANCIA_PROYECCION_MM
            <= y_mm
            <= rectangulo["superior_mm"] + TOLERANCIA_PROYECCION_MM
        )
        proyectado = {
            **taladro,
            "h": h,
            "v": v,
            "x_hoja": x_hoja,
            "y_hoja": y_hoja,
            "plausible": plausible,
        }
        proyectados.append(proyectado)
        if not plausible:
            avisos.append(
                "la proyección de un taladro cae fuera del alzado: "
                f"centro_modelo={taladro['centro']}, h={h:.6f}, "
                f"v={v:.6f}, hoja=({x_hoja:.6f}, {y_hoja:.6f}) m, "
                f"rectángulo={rectangulo}"
            )

    return proyectados, avisos


def elegir_columna_derecha(taladros_proyectados):
    """Agrupa por coordenada horizontal y devuelve la columna más a la derecha."""
    validos = [
        taladro
        for taladro in taladros_proyectados
        if taladro["plausible"]
    ]
    if not validos:
        return None

    columnas = []
    for taladro in sorted(validos, key=lambda dato: dato["h"]):
        columna = next(
            (
                grupo
                for grupo in columnas
                if abs(taladro["h"] - grupo[0]["h"])
                <= TOLERANCIA_COLUMNA_TALADROS_M
            ),
            None,
        )
        if columna is None:
            columnas.append([taladro])
        else:
            columna.append(taladro)

    elegida = max(
        columnas,
        key=lambda grupo: sum(t["h"] for t in grupo) / len(grupo),
    )
    return {
        "taladros": elegida,
        "h": sum(t["h"] for t in elegida) / len(elegida),
        "x_hoja": (
            sum(t["x_hoja"] for t in elegida) / len(elegida)
        ),
    }


def calcular_linea_corte_vertical(
    disposicion, escala_alzado, columna_derecha=None
):
    """Calcula una línea vertical por una columna o por el centro del alzado."""
    centro_x_m, centro_y_m = disposicion["coordenadas_alzado_m"]
    tamano_alzado = disposicion["tamanos"]["alzado"]
    semialto_papel_m = tamano_alzado["alto_mm"] / 2000.0
    margen_papel_m = MARGEN_LINEA_CORTE_MM / 1000.0
    x_hoja_m = (
        columna_derecha["x_hoja"]
        if columna_derecha is not None
        else centro_x_m
    )
    y1_hoja_m = centro_y_m - semialto_papel_m - margen_papel_m
    y2_hoja_m = centro_y_m + semialto_papel_m + margen_papel_m

    factor = escala_alzado["numerador"] / escala_alzado["denominador"]
    x_modelo_m = (x_hoja_m - centro_x_m) / factor
    semialto_modelo_m = semialto_papel_m / factor
    margen_modelo_m = margen_papel_m / factor
    return {
        "orientacion": "vertical",
        "hoja_m": (x_hoja_m, y1_hoja_m, x_hoja_m, y2_hoja_m),
        "vista_m": (
            x_modelo_m,
            -semialto_modelo_m - margen_modelo_m,
            x_modelo_m,
            semialto_modelo_m + margen_modelo_m,
        ),
    }


def preparar_linea_corte(
    pieza,
    modo_corte,
    eje_espesor,
    caja,
    disposicion,
    escala_alzado,
    altura_corte,
):
    """Elige corte vertical por taladros o el horizontal de respaldo."""
    informacion = {
        "taladros": [],
        "proyectados": [],
        "columna_derecha": None,
    }
    avisos = []

    if modo_corte in ("auto", "vertical"):
        try:
            taladros = detectar_taladros(pieza)
            proyectados, avisos_proyeccion = proyectar_taladros(
                taladros,
                eje_espesor,
                caja,
                disposicion,
                escala_alzado,
            )
            columna = elegir_columna_derecha(proyectados)
            informacion.update(
                {
                    "taladros": taladros,
                    "proyectados": proyectados,
                    "columna_derecha": columna,
                }
            )
            avisos.extend(avisos_proyeccion)
        except Exception as error:
            avisos.append(
                "no se pudieron detectar o proyectar los taladros: "
                f"{texto_error(error)}"
            )

        columna = informacion["columna_derecha"]
        if columna is not None:
            return (
                calcular_linea_corte_vertical(
                    disposicion, escala_alzado, columna
                ),
                informacion,
                avisos,
            )
        if modo_corte == "vertical":
            avisos.append(
                "no se encontró una columna válida de taladros; el corte "
                "vertical pasará por el centro del alzado"
            )
            return (
                calcular_linea_corte_vertical(
                    disposicion, escala_alzado, None
                ),
                informacion,
                avisos,
            )

    # En modo auto, no encontrar taladros no pone en riesgo la sección:
    # se conserva el corte horizontal central comprobado en la ronda anterior.
    return (
        calcular_linea_corte(disposicion, escala_alzado, altura_corte),
        informacion,
        avisos,
    )


def _limites_seguros(celda):
    """Limita la isométrica al 80 % de su celda, centrada en ella."""
    margen_total = 1.0 - OCUPACION_MAXIMA_ISOMETRICA
    margen_x = celda["ancho_mm"] * margen_total / 2.0
    margen_y = celda["alto_mm"] * margen_total / 2.0
    return {
        "izquierda_mm": celda["izquierda_mm"] + margen_x,
        "derecha_mm": celda["derecha_mm"] - margen_x,
        "inferior_mm": celda["inferior_mm"] + margen_y,
        "superior_mm": celda["superior_mm"] - margen_y,
    }


def comprobar_cabida_vista(disposicion, nombre):
    """Comprueba con aritmética una vista contra su propia celda segura."""
    excesos = []
    rectangulo = disposicion["rectangulos"][nombre]
    limites = _limites_seguros(disposicion["celdas"][nombre])
    nombre_visible = "isométrica" if nombre == "isometrica" else nombre
    comparaciones = (
        ("izquierda", limites["izquierda_mm"] - rectangulo["izquierda_mm"]),
        ("derecha", rectangulo["derecha_mm"] - limites["derecha_mm"]),
        ("abajo", limites["inferior_mm"] - rectangulo["inferior_mm"]),
        ("arriba", rectangulo["superior_mm"] - limites["superior_mm"]),
    )
    for direccion, exceso in comparaciones:
        if exceso > 0.000001:
            excesos.append(
                f"{nombre_visible} se pasa {exceso:.1f} mm por {direccion}"
            )
    return not excesos, excesos


def _ocupacion_alzado(
    largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m, escala
):
    """Calcula el mayor uso porcentual del ancho o alto de la columna."""
    tamanos = calcular_tamanos_vistas(largo_mm, medio_mm, escala)
    area = calcular_area_util(ancho_hoja_m, alto_hoja_m)
    celda = calcular_celdas(area)["alzado"]
    ocupacion_ancho = tamanos["alzado"]["ancho_mm"] / celda["ancho_mm"]
    ocupacion_alto = tamanos["alzado"]["alto_mm"] / celda["alto_mm"]
    return max(ocupacion_ancho, ocupacion_alto)


def _elegir_escala_alzado(largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m):
    """Prefiere 1:1 y solo amplía o reduce si la ocupación lo exige."""
    escala_uno_a_uno = ESCALAS_REDUCCION[0]
    ocupacion_uno_a_uno = _ocupacion_alzado(
        largo_mm,
        medio_mm,
        ancho_hoja_m,
        alto_hoja_m,
        escala_uno_a_uno,
    )

    if ocupacion_uno_a_uno > OCUPACION_MAXIMA_ALZADO:
        ultimo_intento = None
        for escala in ESCALAS_REDUCCION[1:]:
            ocupacion = _ocupacion_alzado(
                largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m, escala
            )
            ultimo_intento = (escala, ocupacion)
            if ocupacion <= OCUPACION_MAXIMA_ALZADO:
                return escala, ocupacion, None

        escala, ocupacion = ultimo_intento
        exceso = (ocupacion - OCUPACION_MAXIMA_ALZADO) * 100.0
        aviso = (
            "el alzado supera el 60 % incluso a 1:20; se guardará "
            f"igualmente. Ocupación={ocupacion * 100.0:.1f} %, "
            f"exceso={exceso:.1f} puntos porcentuales"
        )
        return escala, ocupacion, aviso

    if ocupacion_uno_a_uno < OCUPACION_MINIMA_ALZADO:
        escala_elegida = escala_uno_a_uno
        ocupacion_elegida = ocupacion_uno_a_uno
        for escala in ESCALAS_AMPLIACION:
            ocupacion = _ocupacion_alzado(
                largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m, escala
            )
            if ocupacion <= OCUPACION_MAXIMA_ALZADO:
                escala_elegida = escala
                ocupacion_elegida = ocupacion
        return escala_elegida, ocupacion_elegida, None

    return escala_uno_a_uno, ocupacion_uno_a_uno, None


def _elegir_escala_para_vista(
    largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m, nombre_vista
):
    """Elige la mayor escala normalizada que cabe para una sola vista."""
    ultimo_intento = None
    for escala in ESCALAS_NORMALIZADAS:
        tamanos = calcular_tamanos_vistas(largo_mm, medio_mm, escala)
        disposicion = calcular_disposicion(
            ancho_hoja_m, alto_hoja_m, tamanos
        )
        cabe, excesos = comprobar_cabida_vista(disposicion, nombre_vista)
        ultimo_intento = (escala, excesos)
        if cabe:
            return escala, None

    escala, excesos = ultimo_intento
    nombre_visible = (
        "la isométrica" if nombre_vista == "isometrica" else "el alzado"
    )
    aviso = (
        f"{nombre_visible} no cabe ni a 1:20; se guardará igualmente. "
        + "; ".join(excesos)
    )
    return escala, aviso


def elegir_escalas_independientes(
    largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m
):
    """Elige por separado la escala del alzado y la de la isométrica."""
    # La escala del plano y del cajetín la manda el alzado. La isométrica es
    # ilustrativa y puede reducirse de forma independiente para usar su celda.
    escala_alzado, ocupacion_alzado, aviso_alzado = _elegir_escala_alzado(
        largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m
    )
    escala_isometrica, aviso_isometrica = _elegir_escala_para_vista(
        largo_mm, medio_mm, ancho_hoja_m, alto_hoja_m, "isometrica"
    )

    tamanos_alzado = calcular_tamanos_vistas(
        largo_mm, medio_mm, escala_alzado
    )
    tamanos_isometrica = calcular_tamanos_vistas(
        largo_mm, medio_mm, escala_isometrica
    )
    tamanos = {
        "alzado": tamanos_alzado["alzado"],
        "isometrica": tamanos_isometrica["isometrica"],
    }
    disposicion = calcular_disposicion(
        ancho_hoja_m, alto_hoja_m, tamanos
    )
    celda_isometrica = disposicion["celdas"]["isometrica"]
    ocupacion_isometrica = max(
        tamanos["isometrica"]["ancho_mm"]
        / celda_isometrica["ancho_mm"],
        tamanos["isometrica"]["alto_mm"]
        / celda_isometrica["alto_mm"],
    )
    avisos = [
        aviso
        for aviso in (aviso_alzado, aviso_isometrica)
        if aviso is not None
    ]
    return (
        escala_alzado,
        escala_isometrica,
        disposicion,
        ocupacion_alzado * 100.0,
        ocupacion_isometrica * 100.0,
        avisos,
    )


def obtener_formato(nombre):
    """Devuelve A3 por defecto o el formato forzado por línea de comandos."""
    try:
        return FORMATOS_POR_NOMBRE[str(nombre).casefold()]
    except KeyError as error:
        raise ValueError(
            "el valor de --formato debe ser a3, a4 o a2"
        ) from error


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


def validar_plantillas(formato):
    """Comprueba que existen la carpeta y el formato que se va a usar."""
    if not CARPETA_PLANTILLAS.is_dir():
        print("No se encontró la carpeta de plantillas de la escuela:")
        print(f"  {CARPETA_PLANTILLAS}")
        return False

    ruta_formato = CARPETA_PLANTILLAS / formato["archivo"]
    if not ruta_formato.is_file():
        print("Falta el formato de hoja seleccionado:")
        print(f"  {ruta_formato}")
        return False

    return True


def texto_error(error):
    """Convierte cualquier error esperado en un mensaje corto y legible."""
    if isinstance(error, (pywintypes.com_error, AttributeError)):
        return texto_error_com(error)
    return f"{type(error).__name__}: {error}"


def guardar_preferencias_y_desactivar_anotaciones(plano):
    """Guarda los originales y deja limpia la futura vista isométrica."""
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
                False,
            )
            if correcto is False:
                raise RuntimeError("SolidWorks devolvió False")
            cambios.append(
                {
                    "tipo": "booleano",
                    "nombre": nombre,
                    "identificador": identificador,
                    "original": original,
                    "nuevo": False,
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
        cambios.append(
            {
                "tipo": "doble",
                "nombre": "swDetailingCenterMarkSize",
                "identificador": PREFERENCIA_TAMANO_MARCA_CENTRO,
                "original": original,
                "nuevo": original,
            }
        )
    except Exception as error:
        avisos.append(
            "no se pudo aplicar swDetailingCenterMarkSize (12): "
            f"{texto_error(error)}"
        )

    return extension, cambios, avisos


def aplicar_preferencias_alzado(
    extension, cambios, marca_centro_mm, escalar_marcas
):
    """Activa marcas y ejes solo para crear el alzado."""
    avisos = []
    if extension is None:
        return avisos

    for cambio in cambios:
        try:
            if cambio["tipo"] == "booleano":
                if cambio["identificador"] == 613:
                    nuevo = escalar_marcas
                else:
                    nuevo = True
                correcto = extension.SetUserPreferenceToggle(
                    cambio["identificador"],
                    OPCION_DETALLADO_SIN_ESPECIFICAR,
                    nuevo,
                )
            else:
                # Las longitudes de la API de SolidWorks se expresan en
                # metros: 2 mm, por ejemplo, deben enviarse como 0.002 m.
                marca_centro_m = marca_centro_mm / 1000.0
                nuevo = marca_centro_m
                correcto = extension.SetUserPreferenceDouble(
                    cambio["identificador"],
                    OPCION_DETALLADO_SIN_ESPECIFICAR,
                    nuevo,
                )
            if correcto is False:
                raise RuntimeError("SolidWorks devolvió False")
            cambio["nuevo"] = nuevo
        except Exception as error:
            avisos.append(
                f"no se pudo aplicar {cambio['nombre']} "
                f"({cambio['identificador']}): {texto_error(error)}"
            )
    return avisos


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
    """Encuentra el nombre real de una vista estándar o de la isométrica."""
    nombres = [str(nombre) for nombre in (nombres_disponibles or [])]
    if not nombres:
        return None

    # Los nombres de vista de SolidWorks se traducen según el idioma de la
    # instalación. Nunca deben escribirse a mano: hay que obtenerlos con
    # IModelDoc2.GetModelViewNames() y emparejarlos de forma flexible.
    # Para mostrar la cara de una placa pueden hacer falta Frontal, Superior
    # o Derecha, según el eje en el que esté su espesor.
    aceptados = {
        "frontal": {"front", "frontal", "alzado"},
        "superior": {"top", "superior"},
        "derecha": {"right", "derecha"},
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

    print("  --- DIAGNÓSTICO COM ANTES DE CREAR EL ALZADO ---")
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

    candidatos_por_tipo = {
        "frontal": ["*Front", "*Frontal", "*Alzado"],
        "superior": ["*Top", "*Superior"],
        "derecha": ["*Right", "*Derecha"],
        "isometrica": ["*Isometric", "*Isométrica"],
    }
    candidatos = (
        [nombre_real]
        if nombre_real is not None
        else candidatos_por_tipo[tipo_vista]
    )

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


def aplicar_escala_isometrica(vista, escala):
    """Independiza la isométrica de la hoja y le asigna su propia escala."""
    aplicar_escala_independiente(vista, escala)


def aplicar_escala_independiente(vista, escala):
    """Asigna una escala propia a una vista que no cabe con la de la hoja."""
    # En la API 2023, UseSheetScale es 0 cuando la escala es independiente.
    # Se usa ScaleDecimal porque acepta un número simple y es más fiable con
    # enlace tardío que ScaleRatio, que obliga a transferir un array por COM.
    vista.UseSheetScale = 0
    vista.ScaleDecimal = escala["numerador"] / escala["denominador"]


def elegir_escala_seccion(largo_mm, espesor_mm, escala_alzado, celda):
    """Conserva la escala del alzado o busca una reducción que quepa."""
    factor_alzado = (
        escala_alzado["numerador"] / escala_alzado["denominador"]
    )
    candidatos = [
        escala
        for escala in ESCALAS_NORMALIZADAS
        if escala["numerador"] / escala["denominador"] <= factor_alzado
    ]
    ultimo = candidatos[-1]
    for escala in candidatos:
        factor = escala["numerador"] / escala["denominador"]
        ocupacion = max(
            largo_mm * factor / celda["ancho_mm"],
            espesor_mm * factor / celda["alto_mm"],
        )
        if ocupacion <= OCUPACION_MAXIMA_ISOMETRICA:
            return escala, escala["texto"] != escala_alzado["texto"], None

    factor = ultimo["numerador"] / ultimo["denominador"]
    ocupacion = max(
        largo_mm * factor / celda["ancho_mm"],
        espesor_mm * factor / celda["alto_mm"],
    )
    aviso = (
        "la sección no cabe en su celda ni a 1:20; se intentará crear "
        f"igualmente. Ocupación estimada={ocupacion * 100.0:.1f} %"
    )
    return ultimo, ultimo["texto"] != escala_alzado["texto"], aviso


def _eliminar_segmento_parcial(plano, segmento):
    """Borra la línea si la sección falló, para conservar el plano anterior."""
    if segmento is None:
        return
    try:
        plano.ClearSelection2(True)
        if segmento.Select2(False, 0):
            plano.EditDelete()
    except Exception:
        # Es una limpieza de último recurso: no debe ocultar el error original.
        pass


def crear_seccion_desde_linea(
    plano,
    vista_alzado,
    linea,
    punto_seccion,
    escala_seccion,
    usar_escala_independiente,
):
    """Crea una sección desde una línea ya calculada."""
    segmento = None
    nombre_vista_alzado = leer(vista_alzado, "GetName2")
    if not nombre_vista_alzado:
        raise RuntimeError("IView.GetName2 no devolvió el nombre del alzado")

    try:
        if not plano.ActivateView(nombre_vista_alzado):
            raise RuntimeError(
                f"IDrawingDoc.ActivateView no activó {nombre_vista_alzado}"
            )
        plano.ClearSelection2(True)

        # El ejemplo oficial de 2023 activa la vista y llama directamente a
        # CreateLine: no usa InsertSketch, porque la vista posee el croquis.
        # SketchManager devuelve un objeto COM y se accede directamente.
        sketch_manager = plano.SketchManager
        x1, y1, x2, y2 = linea["vista_m"]
        segmento = sketch_manager.CreateLine(x1, y1, 0.0, x2, y2, 0.0)
        if segmento is None:
            raise RuntimeError("ISketchManager.CreateLine no creó la línea")

        errores = []
        # Firma de enlace tardío comprobada en la ayuda 2023:
        # CreateSectionViewAt5(x, y, z, etiqueta, opciones,
        #                      componentes_excluidos, profundidad).
        # El método es silencioso. El corte vertical usa NotAligned para
        # respetar la celda reservada. Los fallbacks solo añaden ScaleWithModel;
        # nunca se activa ChangeDirection ("flip").
        opciones_posibles = (
            (
                OPCIONES_SECCION_NO_ALINEADA,
                OPCIONES_SECCION_NO_ALINEADA_ESCALA_MODELO,
            )
            if linea["orientacion"] == "vertical"
            else (
                OPCIONES_SECCION_NORMAL,
                OPCIONES_SECCION_ESCALA_MODELO,
            )
        )
        for opciones in opciones_posibles:
            try:
                plano.ClearSelection2(True)
                if not segmento.Select2(False, 0):
                    raise RuntimeError(
                        "ISketchSegment.Select2 no seleccionó la línea"
                    )
                vista_seccion = plano.CreateSectionViewAt5(
                    punto_seccion[0],
                    punto_seccion[1],
                    0.0,
                    ETIQUETA_SECCION,
                    opciones,
                    pythoncom.Empty,
                    PROFUNDIDAD_SECCION_COMPLETA_M,
                )
                if vista_seccion is None:
                    raise RuntimeError(
                        "CreateSectionViewAt5 no devolvió una vista"
                    )
                if usar_escala_independiente:
                    aplicar_escala_independiente(
                        vista_seccion, escala_seccion
                    )
                plano.ClearSelection2(True)
                return vista_seccion, linea, punto_seccion, opciones
            except Exception as error:
                errores.append(
                    f"opciones={opciones}: {texto_error(error)}"
                )

        raise RuntimeError("; ".join(errores))
    except Exception:
        _eliminar_segmento_parcial(plano, segmento)
        raise


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


def procesar_pieza(
    swapp,
    ruta,
    rutas_abiertas_antes,
    diagnostico=False,
    marca_centro_mm=TAMANO_MARCA_CENTRO_MM_POR_DEFECTO,
    escalar_marcas=True,
    formato=None,
    crear_seccion=False,
    altura_corte=ALTURA_CORTE_POR_DEFECTO,
    modo_corte="auto",
):
    """Genera y guarda el plano de una pieza sin propagar sus errores."""
    resultado = {
        "archivo": ruta.name,
        "hoja": "-",
        "escala": "-",
        "escala_isometrica": "-",
        "vistas": 0,
        "nombres_vista": [],
        "metodos_vista": [],
        "alzado_elegido": "-",
        "eje_espesor": "-",
        "ocupacion_alzado": 0.0,
        "ocupacion_isometrica": 0.0,
        "seccion": "pendiente" if crear_seccion else "omitida",
        "escala_seccion": "-",
        "marca_centro_mm": marca_centro_mm,
        "escalar_marcas": escalar_marcas,
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
        if formato is None:
            formato = obtener_formato(FORMATO_POR_DEFECTO)
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

        caja = _obtener_caja(pieza)
        (
            espesor,
            medio,
            largo,
            tipo_vista_alzado,
            eje_espesor,
        ) = medir_pieza_y_elegir_alzado(pieza, caja)
        resultado["eje_espesor"] = eje_espesor
        # La primera estimación permite aplicar ya una escala razonable al
        # crear la hoja. Después se repite con el tamaño real leído de la API.
        escala, _, _, _, _, _ = elegir_escalas_independientes(
            largo,
            medio,
            formato["ancho_m"],
            formato["alto_m"],
        )
        resultado["hoja"] = formato["nombre"]

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

        # La escala definitiva se decide con la medida real que impuso el
        # formato .slddrt. Si cambia, se vuelve a aplicar a la misma hoja.
        (
            escala_real,
            escala_isometrica,
            disposicion,
            ocupacion,
            ocupacion_isometrica,
            avisos_escala,
        ) = elegir_escalas_independientes(
            largo, medio, ancho_m, alto_m
        )
        if escala_real["texto"] != escala["texto"]:
            escala = escala_real
            configurar_hoja(plano, formato, escala)
        else:
            escala = escala_real
        resultado["escala"] = escala["texto"]
        resultado["escala_isometrica"] = escala_isometrica["texto"]
        resultado["ocupacion_alzado"] = ocupacion
        resultado["ocupacion_isometrica"] = ocupacion_isometrica
        for aviso_escala in avisos_escala:
            resultado["avisos"].append(aviso_escala)
            print(f"AVISO: {aviso_escala}")
        coordenadas_alzado = disposicion["coordenadas_alzado_m"]
        coordenadas_isometrica = disposicion["coordenadas_isometrica_m"]

        # La isométrica debe quedar limpia. Primero se guardan los valores del
        # documento y se desactivan las anotaciones automáticas.
        (
            extension_anotaciones,
            cambios_anotaciones,
            avisos_preferencias,
        ) = guardar_preferencias_y_desactivar_anotaciones(plano)
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

        # SolidWorks puede cambiar la escala al insertar la primera vista.
        # Desactivamos esa preferencia solo durante estas dos inserciones.
        autoescala_original = bool(
            swapp.GetUserPreferenceToggle(PREFERENCIA_AUTOESCALA_VISTAS)
        )
        # SetUserPreferenceToggle es un procedimiento COM: si funciona no
        # devuelve True, sino None. Un fallo real llega como excepción.
        swapp.SetUserPreferenceToggle(PREFERENCIA_AUTOESCALA_VISTAS, False)
        autoescala_modificada = True

        # La isométrica se crea primero, mientras marcas y ejes están
        # desactivados. Así queda limpia sin inspeccionar ni mover IView.
        vista_isometrica, nombre_isometrica, metodo_isometrica = (
            crear_vista_localizada(
                plano,
                pieza,
                ruta_modelo,
                "isometrica",
                coordenadas_isometrica[0],
                coordenadas_isometrica[1],
            )
        )
        aplicar_escala_isometrica(vista_isometrica, escala_isometrica)
        # Cambiar la escala modifica los gráficos. ForceRebuild3 ya está
        # comprobado en este motor y actualiza la vista antes de guardar.
        plano.ForceRebuild3(False)
        resultado["vistas"] += 1
        resultado["nombres_vista"].append(nombre_isometrica)
        resultado["metodos_vista"].append(metodo_isometrica)
        if diagnostico:
            print(
                f"  Vista isométrica creada con {metodo_isometrica}: "
                f"{nombre_isometrica}; escala independiente "
                f"{escala_isometrica['texto']}"
            )

        # Solo ahora se activan las marcas de centro y las líneas de eje.
        avisos_preferencias = aplicar_preferencias_alzado(
            extension_anotaciones,
            cambios_anotaciones,
            marca_centro_mm,
            escalar_marcas,
        )
        for aviso in avisos_preferencias:
            resultado["avisos"].append(aviso)
            print(f"AVISO: {aviso}")

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

        vista_alzado, nombre_alzado, metodo_alzado = crear_vista_localizada(
            plano,
            pieza,
            ruta_modelo,
            tipo_vista_alzado,
            coordenadas_alzado[0],
            coordenadas_alzado[1],
        )
        resultado["vistas"] += 1
        resultado["nombres_vista"].append(nombre_alzado)
        resultado["metodos_vista"].append(metodo_alzado)
        resultado["alzado_elegido"] = nombre_alzado
        if diagnostico:
            print(
                f"  Vista alzado creada con {metodo_alzado}: "
                f"{nombre_alzado}"
            )

        avisos_restauracion = restaurar_preferencias_anotaciones(
            extension_anotaciones, cambios_anotaciones
        )
        preferencias_anotaciones_pendientes = bool(avisos_restauracion)
        for aviso in avisos_restauracion:
            resultado["avisos"].append(aviso)
            print(f"AVISO: {aviso}")

        if crear_seccion:
            try:
                linea_corte, datos_corte, avisos_corte = (
                    preparar_linea_corte(
                        pieza,
                        modo_corte,
                        eje_espesor,
                        caja,
                        disposicion,
                        escala,
                        altura_corte,
                    )
                )
                # Un corte vertical muestra la dimensión media y el espesor;
                # uno horizontal muestra el largo y el espesor.
                dimension_seccion = (
                    medio
                    if linea_corte["orientacion"] == "vertical"
                    else largo
                )
                escala_seccion, escala_propia, aviso_escala_seccion = (
                    elegir_escala_seccion(
                        dimension_seccion,
                        espesor,
                        escala,
                        disposicion["celdas"]["seccion_reservada"],
                    )
                )
                resultado["escala_seccion"] = escala_seccion["texto"]
                if aviso_escala_seccion:
                    resultado["avisos"].append(aviso_escala_seccion)
                    print(f"AVISO: {aviso_escala_seccion}")

                for aviso_corte in avisos_corte:
                    resultado["avisos"].append(aviso_corte)
                    print(f"AVISO: {aviso_corte}")

                punto_seccion = disposicion["coordenadas_seccion_m"]
                if diagnostico:
                    print(
                        "  Taladros detectados: "
                        f"{len(datos_corte['taladros'])}"
                    )
                    for indice, taladro in enumerate(
                        datos_corte["proyectados"], start=1
                    ):
                        print(
                            f"    Taladro {indice}: "
                            f"centro modelo={taladro['centro']} m; "
                            "proyectado hoja="
                            f"({taladro['x_hoja']:.6f}, "
                            f"{taladro['y_hoja']:.6f}) m; "
                            f"radio={taladro['radio']:.6f} m"
                        )
                    columna = datos_corte["columna_derecha"]
                    if columna is None:
                        print("  Columna derecha: no encontrada")
                    else:
                        print(
                            "  Columna derecha elegida: "
                            f"h={columna['h']:.6f} m; "
                            f"x_hoja={columna['x_hoja']:.6f} m; "
                            f"taladros={len(columna['taladros'])}"
                        )
                    hoja = tuple(
                        valor * 1000.0
                        for valor in linea_corte["hoja_m"]
                    )
                    local = tuple(
                        valor * 1000.0
                        for valor in linea_corte["vista_m"]
                    )
                    print(
                        f"  Línea de corte final "
                        f"({linea_corte['orientacion']}): hoja={hoja} mm; "
                        f"vista local={local} mm"
                    )

                (
                    vista_seccion,
                    linea_corte,
                    punto_seccion,
                    opciones_seccion,
                ) = crear_seccion_desde_linea(
                    plano,
                    vista_alzado,
                    linea_corte,
                    punto_seccion,
                    escala_seccion,
                    escala_propia,
                )
                nombre_seccion = (
                    leer(vista_seccion, "GetName2") or "Sección A-A"
                )
                resultado["vistas"] += 1
                resultado["nombres_vista"].append(nombre_seccion)
                resultado["metodos_vista"].append(
                    f"CreateSectionViewAt5(opciones={opciones_seccion})"
                )
                resultado["seccion"] = "creada"
                plano.ForceRebuild3(False)
                if diagnostico:
                    print(
                        "  Sección creada con CreateSectionViewAt5: "
                        f"punto=({punto_seccion[0] * 1000.0:.1f}, "
                        f"{punto_seccion[1] * 1000.0:.1f}) mm; "
                        f"opciones={opciones_seccion}; "
                        f"escala={escala_seccion['texto']}"
                    )
            except Exception as error:
                aviso_seccion = (
                    "no se pudo crear la sección; se guardará el plano con "
                    f"alzado e isométrica: {texto_error(error)}"
                )
                resultado["avisos"].append(aviso_seccion)
                resultado["seccion"] = "aviso"
                print(f"AVISO: {aviso_seccion}")
                if diagnostico:
                    print(f"  Sección: AVISO: {texto_error(error)}")
        elif diagnostico:
            print("  Sección: omitida por --sin-seccion")

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
        f"escala-alzado={resultado['escala']} | "
        f"escala-isométrica={resultado['escala_isometrica']} | "
        f"ocupación-alzado={resultado['ocupacion_alzado']:.1f}% | "
        f"ocupación-isométrica="
        f"{resultado['ocupacion_isometrica']:.1f}% | "
        f"sección={resultado['seccion']} "
        f"(escala={resultado['escala_seccion']}) | "
        f"vistas={resultado['vistas']} | "
        f"nombres={nombres} | "
        f"alzado={resultado['alzado_elegido']} "
        f"(espesor en {resultado['eje_espesor']}) | "
        f"marca-centro={resultado['marca_centro_mm']:g} mm "
        f"({resultado['marca_centro_mm'] / 1000.0:g} m enviados a la API)"
    )
    if resultado["error"]:
        print(f"  Error: {resultado['error']}")
    for aviso in resultado["avisos"]:
        print(f"  Aviso: {aviso}")


def mostrar_uso():
    """Explica las entradas aceptadas por el script."""
    print(
        "Uso: motor_plano.py [--diagnostico] [--marca-centro <mm>] "
        "[--sin-escalar-marcas] [--formato a3|a4|a2] "
        "[--con-seccion] [--sin-seccion] [--altura-corte <fraccion>] "
        "[--corte auto|horizontal|vertical] "
        "<pieza.SLDPRT o carpeta> "
        "[otra pieza o carpeta ...]"
    )


def analizar_argumentos(argumentos):
    """Separa las opciones de vistas, formato y marcas de las rutas."""
    diagnostico = False
    marca_centro_mm = TAMANO_MARCA_CENTRO_MM_POR_DEFECTO
    escalar_marcas = True
    nombre_formato = FORMATO_POR_DEFECTO
    # La sección es opt-in por decisión de diseño: no toda placa la necesita.
    # La pide el usuario (y en el futuro el LLM) para cada pieza concreta.
    crear_seccion = False
    altura_corte = ALTURA_CORTE_POR_DEFECTO
    modo_corte = MODO_CORTE_POR_DEFECTO
    rutas = []
    indice = 0
    while indice < len(argumentos):
        argumento = argumentos[indice]
        if argumento == "--diagnostico":
            diagnostico = True
        elif argumento == "--con-seccion":
            crear_seccion = True
        elif argumento == "--sin-seccion":
            # Alias antiguo aceptado para no romper comandos guardados. Ya no
            # cambia nada porque la sección está desactivada por defecto.
            print(
                "AVISO: --sin-seccion ya no es necesario; la sección solo "
                "se crea cuando se indica --con-seccion."
            )
        elif argumento == "--sin-escalar-marcas":
            escalar_marcas = False
        elif argumento == "--marca-centro":
            indice += 1
            if indice >= len(argumentos):
                raise ValueError("falta el valor en milímetros de --marca-centro")
            texto_valor = argumentos[indice].replace(",", ".")
            try:
                marca_centro_mm = float(texto_valor)
            except ValueError as error:
                raise ValueError(
                    "el valor de --marca-centro debe ser un número decimal"
                ) from error
            if marca_centro_mm <= 0:
                raise ValueError(
                    "el valor de --marca-centro debe ser mayor que cero"
                )
        elif argumento == "--altura-corte":
            indice += 1
            if indice >= len(argumentos):
                raise ValueError(
                    "falta la fracción de semialtura de --altura-corte"
                )
            texto_valor = argumentos[indice].replace(",", ".")
            try:
                altura_corte = float(texto_valor)
            except ValueError as error:
                raise ValueError(
                    "el valor de --altura-corte debe ser un número decimal"
                ) from error
            if not -1.0 <= altura_corte <= 1.0:
                raise ValueError(
                    "--altura-corte debe estar entre -1 y 1"
                )
        elif argumento == "--corte":
            indice += 1
            if indice >= len(argumentos):
                raise ValueError(
                    "falta el valor de --corte: auto, horizontal o vertical"
                )
            modo_corte = argumentos[indice].casefold()
            if modo_corte not in MODOS_CORTE_VALIDOS:
                raise ValueError(
                    "--corte debe ser auto, horizontal o vertical"
                )
        elif argumento == "--formato":
            indice += 1
            if indice >= len(argumentos):
                raise ValueError("falta el valor de --formato: a3, a4 o a2")
            nombre_formato = argumentos[indice].casefold()
            obtener_formato(nombre_formato)
        elif argumento.startswith("--"):
            raise ValueError(f"opción desconocida: {argumento}")
        else:
            rutas.append(argumento)
        indice += 1
    return (
        diagnostico,
        rutas,
        marca_centro_mm,
        escalar_marcas,
        nombre_formato,
        crear_seccion,
        altura_corte,
        modo_corte,
    )


def main():
    """Coordina la generación de planos sin dejar tracebacks crudos."""
    try:
        # Evita UnicodeEncodeError en consolas Windows con páginas antiguas.
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    try:
        (
            diagnostico,
            argumentos_ruta,
            marca_centro_mm,
            escalar_marcas,
            nombre_formato,
            crear_seccion,
            altura_corte,
            modo_corte,
        ) = analizar_argumentos(sys.argv[1:])
    except ValueError as error:
        print(f"Error de argumentos: {error}")
        mostrar_uso()
        return 2
    if not argumentos_ruta:
        mostrar_uso()
        return 2

    try:
        formato = obtener_formato(nombre_formato)
        if not validar_plantillas(formato):
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
                marca_centro_mm=marca_centro_mm,
                escalar_marcas=escalar_marcas,
                formato=formato,
                crear_seccion=crear_seccion,
                altura_corte=altura_corte,
                modo_corte=modo_corte,
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
