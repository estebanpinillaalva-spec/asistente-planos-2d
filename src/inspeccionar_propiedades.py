r"""Inspecciona propiedades y materiales de piezas y planos de SolidWorks.

Uso: .venv\Scripts\python.exe src\inspeccionar_propiedades.py "C:\carpeta"
"""

import csv
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
from propiedades import (
    leer_propiedades,
    material_asignado,
    nombre_configuracion_activa,
)


TIPO_DOCUMENTO_PIEZA = 1
TIPO_DOCUMENTO_PLANO = 3
OPCION_ABRIR_SILENCIOSO = 1
TIPOS_ADMITIDOS = {
    ".sldprt": ("PIEZA", TIPO_DOCUMENTO_PIEZA),
    ".slddrw": ("PLANO", TIPO_DOCUMENTO_PLANO),
}
RUTA_CSV = (
    Path(__file__).resolve().parent.parent / "notas" / "propiedades-piezas.csv"
)
CAMPOS_CSV = [
    "archivo",
    "tipo",
    "origen",
    "propiedad",
    "valor",
    "valor_resuelto",
]


def texto_error(error):
    """Convierte errores COM y Python en mensajes cortos y legibles."""
    if isinstance(error, (pywintypes.com_error, AttributeError)):
        return texto_error_com(error)
    return f"{type(error).__name__}: {error}"


def buscar_documentos(entradas):
    """Recoge piezas y planos dentro de las rutas, sin duplicados."""
    documentos = set()

    for entrada in entradas:
        if entrada.is_file():
            if (
                entrada.suffix.lower() in TIPOS_ADMITIDOS
                and not es_archivo_bloqueo(entrada)
            ):
                documentos.add(entrada.resolve())
            else:
                print(
                    "Aviso: el archivo no es una pieza SLDPRT ni un plano "
                    f"SLDDRW válido: {entrada}"
                )
            continue

        if entrada.is_dir():
            for ruta in entrada.rglob("*"):
                if (
                    ruta.is_file()
                    and ruta.suffix.lower() in TIPOS_ADMITIDOS
                    and not es_archivo_bloqueo(ruta)
                ):
                    documentos.add(ruta.resolve())
            continue

        print(f"Aviso: no existe el archivo o la carpeta: {entrada}")

    return sorted(documentos, key=lambda ruta: str(ruta).lower())


def abrir_documento(swapp, ruta):
    """Abre una pieza o plano en silencio y devuelve errores y avisos."""
    errores = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    avisos = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

    _, tipo_documento = TIPOS_ADMITIDOS[ruta.suffix.lower()]
    documento = swapp.OpenDoc6(
        str(ruta),
        tipo_documento,
        OPCION_ABRIR_SILENCIOSO,
        "",
        errores,
        avisos,
    )
    return documento, int(errores.value), int(avisos.value)


def resultado_vacio(ruta):
    """Prepara un resultado vacío para una pieza o un plano."""
    tipo, _ = TIPOS_ADMITIDOS[ruta.suffix.lower()]
    return {
        "ruta": ruta,
        "archivo": ruta.name,
        "tipo": tipo,
        "configuracion": "",
        "material": "",
        "propiedades": {"archivo": {}, "configuracion": {}},
        "total_notas": 0,
        "notas": [],
        "avisos": [],
        "fallida": False,
        "error": "",
    }


def _miembro_com_sin_argumentos(objeto, nombre):
    """Obtiene un objeto o una colección COM sin llamar dos veces por error."""
    # No usamos leer(): un CDispatch también es invocable. Si pywin32 ya ha
    # devuelto el objeto COM, llamarlo provocaría DISP_E_MEMBERNOTFOUND.
    miembro = getattr(objeto, nombre)
    if hasattr(miembro, "_oleobj_") or not callable(miembro):
        return miembro
    return miembro()


def leer_notas_plano(documento, limite=15):
    """Cuenta las notas de la hoja activa y devuelve sus primeros textos."""
    total = 0
    textos = []

    # GetFirstView devuelve primero la vista de hoja, donde suele estar el
    # cajetín. GetNextView permite recorrer después las demás vistas.
    vista = _miembro_com_sin_argumentos(documento, "GetFirstView")
    while vista is not None:
        notas = _miembro_com_sin_argumentos(vista, "GetNotes")
        if notas:
            if hasattr(notas, "_oleobj_"):
                notas = (notas,)
            for nota in notas:
                total += 1
                if len(textos) < limite:
                    # GetText devuelve una cadena simple: leer() sí es seguro.
                    textos.append(str(leer(nota, "GetText") or ""))

        vista = _miembro_com_sin_argumentos(vista, "GetNextView")

    return total, textos


def procesar_documento(swapp, ruta, rutas_abiertas_antes):
    """Abre, inspecciona y cierra un documento sin propagar sus errores."""
    resultado = resultado_vacio(ruta)
    documento = None
    nombre_para_cerrar = ruta.name
    ya_estaba_abierta = normalizar_ruta(ruta) in rutas_abiertas_antes

    try:
        documento, errores, avisos = abrir_documento(swapp, ruta)
        if documento is None:
            raise RuntimeError(
                "OpenDoc6 no abrió el documento; "
                f"errores={errores}, avisos={avisos}"
            )

        # GetTitle devuelve texto simple, por eso aquí sí se usa leer().
        nombre_para_cerrar = leer(documento, "GetTitle") or ruta.name

        if resultado["tipo"] == "PIEZA":
            configuracion = nombre_configuracion_activa(documento)
            resultado["configuracion"] = configuracion
            resultado["propiedades"] = leer_propiedades(
                documento, configuracion
            )
            resultado["material"] = material_asignado(
                documento, configuracion
            )
        else:
            # Un plano no tiene configuración de pieza: sus propiedades están
            # en el bloque general del propio archivo.
            resultado["configuracion"] = "(no aplica)"
            resultado["propiedades"] = leer_propiedades(documento, "")
            try:
                total, notas = leer_notas_plano(documento)
                resultado["total_notas"] = total
                resultado["notas"] = notas
            except Exception as error:
                # Las notas son información adicional. Su fallo no debe ocultar
                # las propiedades personalizadas que sí se hayan podido leer.
                resultado["avisos"].append(
                    f"no se pudieron leer las notas: {texto_error(error)}"
                )

        if errores or avisos:
            resultado["avisos"].append(
                f"OpenDoc6: errores={errores}, avisos={avisos}"
            )
    except (pywintypes.com_error, AttributeError) as error:
        resultado["fallida"] = True
        resultado["error"] = texto_error_com(error)
    except Exception as error:
        resultado["fallida"] = True
        resultado["error"] = texto_error(error)
    finally:
        # Los documentos que ya estaban abiertos pertenecen al usuario.
        if documento is not None and not ya_estaba_abierta:
            try:
                swapp.CloseDoc(nombre_para_cerrar)
            except Exception as error:
                detalle = f"no se pudo cerrar: {texto_error(error)}"
                resultado["error"] = "; ".join(
                    texto for texto in [resultado["error"], detalle] if texto
                )
                resultado["fallida"] = True

    return resultado


def imprimir_documento(resultado):
    """Muestra la información encontrada en una pieza o un plano."""
    print(f"{resultado['tipo']} | {resultado['archivo']}")
    if resultado["fallida"]:
        print(f"  ERROR: {resultado['error']}")
        return

    print(f"  Configuración activa: {resultado['configuracion']}")
    if resultado["tipo"] == "PIEZA":
        material = resultado["material"] or "(sin material asignado)"
        print(f"  Material asignado: {material}")

    hay_propiedades = any(resultado["propiedades"].values())
    if not hay_propiedades:
        print("  No tiene propiedades personalizadas.")
    else:
        for origen in ("archivo", "configuracion"):
            print(f"  Propiedades de {origen}:")
            propiedades = resultado["propiedades"][origen]
            if not propiedades:
                print("    (ninguna)")
                continue

            for nombre in sorted(propiedades, key=str.casefold):
                datos = propiedades[nombre]
                texto = datos["valor"]
                if datos["resuelto"] != datos["valor"]:
                    texto += f" -> resuelto: {datos['resuelto']}"
                print(f"    {nombre}: {texto}")

    if resultado["tipo"] == "PLANO":
        print(f"  Notas de la hoja activa: {resultado['total_notas']}")
        for indice, nota in enumerate(resultado["notas"], start=1):
            nota_visible = nota.replace("\r", " ").replace("\n", " | ")
            print(f"    {indice}. {nota_visible}")
        restantes = resultado["total_notas"] - len(resultado["notas"])
        if restantes > 0:
            print(f"    ... y {restantes} notas más.")

    for aviso in resultado["avisos"]:
        print(f"  Aviso: {aviso}")


def crear_filas_csv(resultados):
    """Convierte propiedades y materiales en filas para el CSV."""
    filas = []
    for resultado in resultados:
        if resultado["tipo"] == "PIEZA":
            filas.append(
                {
                    "archivo": resultado["archivo"],
                    "tipo": resultado["tipo"],
                    "origen": "material",
                    "propiedad": "Material asignado",
                    "valor": resultado["material"],
                    "valor_resuelto": resultado["material"],
                }
            )

        for origen in ("archivo", "configuracion"):
            propiedades = resultado["propiedades"][origen]
            for nombre in sorted(propiedades, key=str.casefold):
                datos = propiedades[nombre]
                filas.append(
                    {
                        "archivo": resultado["archivo"],
                        "tipo": resultado["tipo"],
                        "origen": origen,
                        "propiedad": nombre,
                        "valor": datos["valor"],
                        "valor_resuelto": datos["resuelto"],
                    }
                )
    return filas


def escribir_csv(resultados):
    """Escribe todas las propiedades con formato cómodo para Excel."""
    RUTA_CSV.parent.mkdir(parents=True, exist_ok=True)
    with RUTA_CSV.open("w", encoding="utf-8-sig", newline="") as archivo:
        escritor = csv.DictWriter(
            archivo,
            fieldnames=CAMPOS_CSV,
            delimiter=";",
        )
        escritor.writeheader()
        escritor.writerows(crear_filas_csv(resultados))


def resumir_propiedades(resultados):
    """Cuenta en cuántos documentos aparece cada propiedad y si está vacía."""
    resumen = {}

    for resultado in resultados:
        propiedades_del_documento = {}
        for propiedades in resultado["propiedades"].values():
            for nombre, datos in propiedades.items():
                esta_vacia = not (
                    datos["valor"].strip() or datos["resuelto"].strip()
                )
                if nombre in propiedades_del_documento:
                    propiedades_del_documento[nombre] &= esta_vacia
                else:
                    propiedades_del_documento[nombre] = esta_vacia

        for nombre, esta_vacia in propiedades_del_documento.items():
            conteo = resumen.setdefault(
                nombre, {"documentos": 0, "vacias": 0}
            )
            conteo["documentos"] += 1
            conteo["vacias"] += int(esta_vacia)

    return resumen


def imprimir_resumen(resultados):
    """Muestra una tabla agregada de presencia y propiedades vacías."""
    resumen = resumir_propiedades(resultados)
    print("Resumen agregado de propiedades:")
    if not resumen:
        print("  No se encontró ninguna propiedad personalizada.")
    else:
        ancho = max(9, max(len(nombre) for nombre in resumen))
        print(
            f"  {'Propiedad':<{ancho}} | {'Documentos':>10} | "
            f"{'Vacías':>6}"
        )
        print(f"  {'-' * ancho}-+------------+-------")
        for nombre in sorted(resumen, key=str.casefold):
            conteo = resumen[nombre]
            print(
                f"  {nombre:<{ancho}} | {conteo['documentos']:>10} | "
                f"{conteo['vacias']:>6}"
            )

    piezas_validas = [
        resultado
        for resultado in resultados
        if resultado["tipo"] == "PIEZA" and not resultado["fallida"]
    ]
    con_material = sum(
        bool(resultado["material"].strip()) for resultado in piezas_validas
    )
    print(
        "Material asignado en piezas: "
        f"{con_material} con material; "
        f"{len(piezas_validas) - con_material} sin material."
    )

    fallidas = sum(resultado["fallida"] for resultado in resultados)
    print(
        f"Documentos inspeccionados: {len(resultados)}; "
        f"fallidos: {fallidas}"
    )


def main():
    """Coordina la inspección de piezas y planos y la escritura del CSV."""
    try:
        # Evita fallos al imprimir tildes en consolas Windows antiguas.
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if len(sys.argv) < 2:
        print(
            "Uso: inspeccionar_propiedades.py "
            "<pieza.SLDPRT, plano.SLDDRW o carpeta> [otra ruta ...]"
        )
        print(
            "Indica al menos una pieza SLDPRT, un plano SLDDRW o una carpeta."
        )
        return 2

    try:
        entradas = [
            Path(argumento).expanduser().resolve() for argumento in sys.argv[1:]
        ]
        swapp = conectar()
        if swapp is None:
            return 1

        rutas_abiertas_antes = obtener_rutas_documentos_abiertos(swapp)
        documentos = buscar_documentos(entradas)
        print(f"Piezas y planos encontrados: {len(documentos)}")

        resultados = []
        for indice, ruta in enumerate(documentos, start=1):
            print(f"[{indice}/{len(documentos)}] {ruta}")
            resultado = procesar_documento(
                swapp, ruta, rutas_abiertas_antes
            )
            resultados.append(resultado)
            imprimir_documento(resultado)

        escribir_csv(resultados)
        imprimir_resumen(resultados)
        print(f"CSV: {RUTA_CSV}")
        return 0
    except Exception as error:
        print("No se pudo completar la inspección de propiedades.")
        print(f"Detalle del error: {texto_error(error)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
