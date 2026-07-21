r"""Clasifica piezas de SolidWorks como planas usando su caja envolvente.

Uso: .venv\Scripts\python.exe src\clasificar_planas.py "C:\ruta\carpeta"
"""

import csv
import sys
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

from com_utils import conectar, leer, texto_error_com


# Una placa de este banco de pruebas no debe superar este espesor.
ESPESOR_MAXIMO_MM = 15.0

# La dimensión intermedia debe ser, como mínimo, cuatro veces el espesor.
RATIO_MINIMO = 4.0

# Valores numéricos definidos por la API de SolidWorks.
TIPO_DOCUMENTO_PIEZA = 1
OPCION_ABRIR_SILENCIOSO = 1
TIPO_CUERPO_SOLIDO = 0

RUTA_CSV = Path(__file__).resolve().parent.parent / "notas" / "clasificacion-planas.csv"
CAMPOS_CSV = [
    "archivo",
    "ruta",
    "espesor_mm",
    "medio_mm",
    "largo_mm",
    "ratio",
    "clasificacion",
    "error",
]


def texto_error(error):
    """Devuelve un texto legible tanto para errores COM como para errores Python."""
    if isinstance(error, (pywintypes.com_error, AttributeError)):
        return texto_error_com(error)
    return f"{type(error).__name__}: {error}"


def buscar_piezas(carpetas):
    """Busca archivos SLDPRT de forma recursiva y sin duplicados."""
    piezas = set()

    for carpeta in carpetas:
        if not carpeta.is_dir():
            print(f"Aviso: no existe la carpeta o no es válida: {carpeta}")
            continue

        for ruta in carpeta.rglob("*"):
            if ruta.is_file() and ruta.suffix.lower() == ".sldprt":
                piezas.add(ruta.resolve())

    return sorted(piezas, key=lambda ruta: str(ruta).lower())


def abrir_pieza(swapp, ruta):
    """Abre una pieza en silencio y devuelve documento, errores y avisos."""
    # OpenDoc6 usa dos enteros por referencia para comunicar errores y avisos.
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


def validar_caja(caja):
    """Convierte una caja COM en seis números y comprueba su formato."""
    if caja is None or len(caja) != 6:
        raise ValueError("la caja envolvente no contiene seis valores")
    return tuple(float(valor) for valor in caja)


def caja_por_cuerpos(pieza):
    """Calcula una caja global combinando las cajas de los cuerpos sólidos."""
    # GetBodies2 devuelve objetos COM; por eso no se usa leer() en esta llamada.
    cuerpos = pieza.GetBodies2(TIPO_CUERPO_SOLIDO, False)
    if not cuerpos:
        raise ValueError("la pieza no contiene cuerpos sólidos")

    cajas = [validar_caja(leer(cuerpo, "GetBodyBox")) for cuerpo in cuerpos]
    return (
        min(caja[0] for caja in cajas),
        min(caja[1] for caja in cajas),
        min(caja[2] for caja in cajas),
        max(caja[3] for caja in cajas),
        max(caja[4] for caja in cajas),
        max(caja[5] for caja in cajas),
    )


def obtener_caja(pieza):
    """Obtiene la caja de la pieza y usa las cajas de cuerpos como alternativa."""
    try:
        # True conserva las unidades del sistema de SolidWorks: metros.
        return validar_caja(pieza.GetPartBox(True))
    except Exception as error_principal:
        try:
            return caja_por_cuerpos(pieza)
        except Exception as error_alternativo:
            raise RuntimeError(
                "fallaron GetPartBox "
                f"({texto_error(error_principal)}) y la alternativa por cuerpos "
                f"({texto_error(error_alternativo)})"
            ) from error_alternativo


def calcular_dimensiones(caja):
    """Convierte la caja en espesor, dimensión media y largo, en milímetros."""
    xmin, ymin, zmin, xmax, ymax, zmax = caja
    dimensiones = sorted(
        [
            abs(xmax - xmin) * 1000.0,
            abs(ymax - ymin) * 1000.0,
            abs(zmax - zmin) * 1000.0,
        ]
    )
    espesor, medio, largo = dimensiones

    if espesor <= 0:
        raise ValueError("la caja envolvente tiene espesor cero")

    return espesor, medio, largo


def clasificar_dimensiones(espesor, medio):
    """Aplica los dos umbrales del banco de pruebas."""
    ratio = medio / espesor
    es_plana = espesor <= ESPESOR_MAXIMO_MM and ratio >= RATIO_MINIMO
    return ratio, "PLANA" if es_plana else "NO PLANA"


def resultado_vacio(ruta):
    """Crea una fila vacía para una pieza antes de intentar medirla."""
    return {
        "archivo": ruta.name,
        "ruta": str(ruta),
        "espesor_mm": None,
        "medio_mm": None,
        "largo_mm": None,
        "ratio": None,
        "clasificacion": "ERROR",
        "error": "",
    }


def procesar_pieza(swapp, ruta):
    """Abre, mide, clasifica y cierra una pieza sin propagar sus errores."""
    resultado = resultado_vacio(ruta)
    documento = None
    nombre_para_cerrar = ruta.name

    try:
        documento, errores, avisos = abrir_pieza(swapp, ruta)
        if documento is None:
            raise RuntimeError(
                f"OpenDoc6 no abrió el documento; errores={errores}, avisos={avisos}"
            )

        # GetTitle devuelve una cadena, así que aquí sí es seguro usar leer().
        nombre_para_cerrar = leer(documento, "GetTitle") or ruta.name

        if errores or avisos:
            resultado["error"] = (
                f"OpenDoc6: errores={errores}, avisos={avisos}"
            )

        caja = obtener_caja(documento)
        espesor, medio, largo = calcular_dimensiones(caja)
        ratio, clasificacion = clasificar_dimensiones(espesor, medio)

        resultado.update(
            {
                "espesor_mm": espesor,
                "medio_mm": medio,
                "largo_mm": largo,
                "ratio": ratio,
                "clasificacion": clasificacion,
            }
        )
    except (pywintypes.com_error, AttributeError) as error:
        resultado["error"] = texto_error_com(error)
    except Exception as error:
        resultado["error"] = texto_error(error)
    finally:
        if documento is not None:
            try:
                swapp.CloseDoc(nombre_para_cerrar)
            except (pywintypes.com_error, AttributeError) as error:
                detalle = f"no se pudo cerrar: {texto_error_com(error)}"
                resultado["error"] = "; ".join(
                    texto for texto in [resultado["error"], detalle] if texto
                )
                resultado["clasificacion"] = "ERROR"
            except Exception as error:
                detalle = f"no se pudo cerrar: {texto_error(error)}"
                resultado["error"] = "; ".join(
                    texto for texto in [resultado["error"], detalle] if texto
                )
                resultado["clasificacion"] = "ERROR"

    return resultado


def ordenar_resultados(resultados):
    """Coloca primero las placas planas más finas y deja los errores al final."""
    orden_clasificacion = {"PLANA": 0, "NO PLANA": 1, "ERROR": 2}
    return sorted(
        resultados,
        key=lambda fila: (
            orden_clasificacion[fila["clasificacion"]],
            fila["espesor_mm"] if fila["espesor_mm"] is not None else float("inf"),
            fila["archivo"].lower(),
        ),
    )


def escribir_csv(resultados):
    """Escribe un CSV con BOM, punto y coma y números fáciles de leer en Excel."""
    RUTA_CSV.parent.mkdir(parents=True, exist_ok=True)

    with RUTA_CSV.open("w", encoding="utf-8-sig", newline="") as archivo_csv:
        escritor = csv.DictWriter(
            archivo_csv,
            fieldnames=CAMPOS_CSV,
            delimiter=";",
        )
        escritor.writeheader()

        for resultado in resultados:
            fila = resultado.copy()
            for campo in ("espesor_mm", "medio_mm", "largo_mm"):
                valor = fila[campo]
                fila[campo] = "" if valor is None else f"{valor:.3f}"
            valor_ratio = fila["ratio"]
            fila["ratio"] = "" if valor_ratio is None else f"{valor_ratio:.3f}"
            escritor.writerow(fila)


def imprimir_resultado(resultado):
    """Muestra una línea legible por cada pieza procesada."""
    if resultado["clasificacion"] == "ERROR":
        print(f"ERROR | {resultado['archivo']} | {resultado['error']}")
        return

    print(
        f"{resultado['clasificacion']} | "
        f"espesor={resultado['espesor_mm']:.3f} mm | "
        f"ratio={resultado['ratio']:.3f} | {resultado['archivo']}"
    )
    if resultado["error"]:
        print(f"  Aviso: {resultado['error']}")


def main():
    """Coordina la búsqueda, medición y escritura del informe."""
    try:
        # Evita fallos al imprimir tildes en consolas Windows antiguas.
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if len(sys.argv) < 2:
        print("Uso: clasificar_planas.py <carpeta> [otra carpeta ...]")
        print("Indica al menos una carpeta que contenga piezas SLDPRT.")
        return 2

    try:
        carpetas = [Path(argumento).expanduser().resolve() for argumento in sys.argv[1:]]
        swapp = conectar()
        if swapp is None:
            return 1

        piezas = buscar_piezas(carpetas)
        print(f"Piezas SLDPRT encontradas: {len(piezas)}")

        resultados = []
        for indice, ruta in enumerate(piezas, start=1):
            print(f"[{indice}/{len(piezas)}] {ruta}")
            resultado = procesar_pieza(swapp, ruta)
            resultados.append(resultado)

        resultados = ordenar_resultados(resultados)
        escribir_csv(resultados)

        print("Resultados ordenados:")
        for resultado in resultados:
            imprimir_resultado(resultado)

        planas = sum(fila["clasificacion"] == "PLANA" for fila in resultados)
        no_planas = sum(fila["clasificacion"] == "NO PLANA" for fila in resultados)
        fallidas = sum(fila["clasificacion"] == "ERROR" for fila in resultados)

        print("Resumen:")
        print(f"  Planas: {planas}")
        print(f"  No planas: {no_planas}")
        print(f"  Fallidas: {fallidas}")
        print(f"CSV: {RUTA_CSV}")
        return 0
    except Exception as error:
        print("No se pudo completar la clasificación de piezas.")
        print(f"Detalle del error: {texto_error(error)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
