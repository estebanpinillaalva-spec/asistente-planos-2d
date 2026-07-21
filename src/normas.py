"""Consulta las tablas normativas guardadas como datos del proyecto."""

from functools import lru_cache
from pathlib import Path

import yaml


RUTA_PROYECTO = Path(__file__).resolve().parent.parent
RUTA_AGUJEROS_PASO = RUTA_PROYECTO / "datos" / "iso273-agujeros-paso.yaml"


@lru_cache(maxsize=1)
def cargar_tabla_agujeros_paso():
    """Carga una sola vez la tabla ISO 273 desde el YAML del proyecto."""
    with RUTA_AGUJEROS_PASO.open(encoding="utf-8") as archivo:
        return yaml.safe_load(archivo)


def agujero_paso(metrica, serie=None):
    """Devuelve el diámetro y la justificación de un agujero de paso."""
    tabla = cargar_tabla_agujeros_paso()

    texto_metrica = str(metrica).strip().upper()
    if texto_metrica.startswith("M"):
        clave_metrica = texto_metrica
    else:
        clave_metrica = f"M{texto_metrica}"

    diametros = tabla["diametros"]
    if clave_metrica not in diametros:
        disponibles = ", ".join(diametros)
        raise ValueError(
            f"La métrica {clave_metrica} no está en la tabla ISO 273. "
            f"Métricas disponibles: {disponibles}."
        )

    if serie is None:
        serie_elegida = tabla["serie_por_defecto"]
    else:
        serie_elegida = str(serie).strip().lower()

    series_disponibles = diametros[clave_metrica]
    if serie_elegida not in series_disponibles:
        disponibles = ", ".join(series_disponibles)
        raise ValueError(
            f"La serie '{serie_elegida}' no es válida. "
            f"Series disponibles: {disponibles}."
        )

    diametro = float(series_disponibles[serie_elegida])
    grado = tabla["grado_tolerancia_por_defecto"]
    datos_verificacion = tabla["verificacion"][clave_metrica]
    estado_verificacion = datos_verificacion["estado"]
    advertencia = None

    if estado_verificacion == "discrepante":
        discrepancias = datos_verificacion["discrepancias"]
        if serie_elegida in discrepancias:
            diferencia = discrepancias[serie_elegida]
            advertencia = (
                f"ATENCION: las fuentes discrepan para {clave_metrica} "
                f"serie {serie_elegida} ({diferencia['adoptado']:g} adoptado, "
                f"{diferencia['fractory']:g} según Fractory). "
                "Verificar contra ISO 273 oficial antes de fabricar."
            )
        else:
            detalles = []
            for nombre_serie, diferencia in discrepancias.items():
                detalles.append(
                    f"{nombre_serie} ({diferencia['adoptado']:g} adoptado, "
                    f"{diferencia['fractory']:g} según Fractory)"
                )
            advertencia = (
                f"ATENCION: las fuentes discrepan para {clave_metrica} en "
                f"otras series: {', '.join(detalles)}. "
                "Verificar contra ISO 273 oficial antes de fabricar."
            )
    elif estado_verificacion == "sin_contraste":
        advertencia = (
            f"ATENCION: valor de {clave_metrica} publicado por una sola fuente, "
            "sin contraste independiente. Verificar contra ISO 273 oficial "
            "antes de fabricar."
        )
    elif estado_verificacion != "corroborado":
        raise ValueError(
            f"Estado de verificación desconocido para {clave_metrica}: "
            f"{estado_verificacion}."
        )

    justificacion = (
        f"Agujero de paso para {clave_metrica}, serie {serie_elegida} "
        f"({tabla['norma']}); grado {grado}"
    )
    if advertencia is not None:
        justificacion = f"{justificacion}. {advertencia}"

    return {
        "diametro_mm": diametro,
        "grado": grado,
        "cota": f"⌀{diametro:g} {grado}",
        "justificacion": justificacion,
        "verificacion": estado_verificacion,
        "advertencia": advertencia,
    }
