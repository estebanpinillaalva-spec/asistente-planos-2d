"""Comprueba la tabla normativa contra decisiones reales de los planos."""

import sys

from normas import agujero_paso


# Casos conocidos. Esta lista se puede ampliar al revisar más planos de Esteban.
CASOS_CONOCIDOS = [
    {
        "descripcion": "Placa motor KAM-0024: M8 media",
        "metrica": "M8",
        "serie": "media",
        "diametro_mm": 9.0,
        "cota": "⌀9 H13",
        "advertencia": None,
    },
    {
        "descripcion": "M6 media",
        "metrica": "M6",
        "serie": "media",
        "diametro_mm": 6.6,
    },
    {
        "descripcion": "M8 fina",
        "metrica": "M8",
        "serie": "fina",
        "diametro_mm": 8.4,
    },
    {
        "descripcion": "M12 media discrepante",
        "metrica": "M12",
        "serie": "media",
        "diametro_mm": 13.5,
        "advertencia_contiene": "Fractory",
        "advertencia_en_justificacion": True,
    },
    {
        "descripcion": "M14 media sin contraste",
        "metrica": "M14",
        "serie": "media",
        "diametro_mm": 15.5,
        "verificacion": "sin_contraste",
        "advertencia_no_nula": True,
    },
]

METRICA_INEXISTENTE = "M7"


def comprobar_caso(caso):
    """Comprueba un caso conocido y muestra un resultado legible."""
    try:
        resultado = agujero_paso(caso["metrica"], caso["serie"])
        diferencias = []

        if resultado["diametro_mm"] != caso["diametro_mm"]:
            diferencias.append(
                f"diámetro esperado {caso['diametro_mm']}, "
                f"obtenido {resultado['diametro_mm']}"
            )

        if "cota" in caso and resultado["cota"] != caso["cota"]:
            diferencias.append(
                f"cota esperada '{caso['cota']}', "
                f"obtenida '{resultado['cota']}'"
            )

        if (
            "advertencia" in caso
            and resultado["advertencia"] != caso["advertencia"]
        ):
            diferencias.append(
                f"advertencia esperada {caso['advertencia']!r}, "
                f"obtenida {resultado['advertencia']!r}"
            )

        if "advertencia_contiene" in caso:
            advertencia = resultado["advertencia"]
            texto_esperado = caso["advertencia_contiene"]
            if advertencia is None or texto_esperado not in advertencia:
                diferencias.append(
                    f"la advertencia debía mencionar '{texto_esperado}'"
                )

        if (
            caso.get("advertencia_no_nula")
            and resultado["advertencia"] is None
        ):
            diferencias.append("se esperaba una advertencia no nula")

        if (
            "verificacion" in caso
            and resultado["verificacion"] != caso["verificacion"]
        ):
            diferencias.append(
                f"verificación esperada '{caso['verificacion']}', "
                f"obtenida '{resultado['verificacion']}'"
            )

        if (
            caso.get("advertencia_en_justificacion")
            and resultado["advertencia"] not in resultado["justificacion"]
        ):
            diferencias.append("la justificación no contiene la advertencia")

        if diferencias:
            print(f"FALLO | {caso['descripcion']} | {'; '.join(diferencias)}")
            return False

        texto = f"{resultado['diametro_mm']:g} mm"
        if "cota" in caso:
            texto += f", {resultado['cota']}"
        if caso.get("advertencia") is None and "advertencia" in caso:
            texto += ", sin advertencia"
        if "verificacion" in caso:
            texto += f", {resultado['verificacion']}"
        if "advertencia_contiene" in caso or caso.get("advertencia_no_nula"):
            texto += ", advertencia presente"
        print(f"OK | {caso['descripcion']} -> {texto}")
        return True
    except Exception as error:
        print(
            f"FALLO | {caso['descripcion']} | "
            f"{type(error).__name__}: {error}"
        )
        return False


def comprobar_metrica_inexistente():
    """Comprueba que una métrica ausente produzca un mensaje útil."""
    try:
        agujero_paso(METRICA_INEXISTENTE)
    except ValueError as error:
        print(f"OK | {METRICA_INEXISTENTE} inexistente -> {error}")
        return True
    except Exception as error:
        print(
            f"FALLO | {METRICA_INEXISTENTE} inexistente | "
            f"{type(error).__name__}: {error}"
        )
        return False

    print(
        f"FALLO | {METRICA_INEXISTENTE} inexistente | "
        "la consulta no produjo ningún error"
    )
    return False


def main():
    """Ejecuta todas las comprobaciones y devuelve su código de salida."""
    try:
        # Evita fallos al imprimir tildes y el símbolo de diámetro en Windows.
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    resultados = [comprobar_caso(caso) for caso in CASOS_CONOCIDOS]
    resultados.append(comprobar_metrica_inexistente())

    correctos = sum(resultados)
    fallos = len(resultados) - correctos
    print(f"Resumen: {correctos} OK, {fallos} FALLO")

    return 0 if fallos == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
