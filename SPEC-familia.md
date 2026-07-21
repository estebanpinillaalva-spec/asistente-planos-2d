# Mini-spec de la familia de piezas (tarea 0.3)

> Rellenar tras elegir la familia. Este documento es el conocimiento de dominio del proyecto:
> de aquí saldrá la base de reglas (YAML) que usa el motor. Sé concreto, no exhaustivo.

## 1. Familia elegida

- **Nombre:** **Placas y soportes planos taladrados** (chapa/plancha cortada plana, con taladros y contorno; espesor constante).
- **Por qué esta:**
  - Es la familia con más ejemplares reales disponibles (~15) y con **6 planos previos hechos por Esteban** que sirven de verdad de referencia.
  - **Vistas casi deterministas:** una vista principal + espesor. Simplifica mucho la Fase 1.
  - **Features enumerables:** contorno, taladros (⌀ y posición), avellanados, radios de esquina.
  - Salida directa a **DXF para corte láser**, que es lo facturable de la Fase 4.
  - Todas provienen de **trabajos universitarios de Esteban** (TFG del kart y prácticas): **no hay material confidencial**, así que los resultados son publicables en el portfolio.
- **Banco de pruebas — 8 piezas**, todas **medidas** con `src/clasificar_planas.py` (no elegidas por nombre). Espesores de 2 a 5 mm: familia muy homogénea.

  | Pieza | Espesor | Ratio | Origen |
  |---|---|---|---|
  | `placa rodamientos` | 2,0 mm | 55 | TFG (kart) |
  | `Chapa delantera` | 2,5 mm | 155 | TFG (kart) |
  | `chapa superior pata` | 3,0 mm | 20 | Trabajo universitario |
  | `tapa motor` | 4,0 mm | 150 | Trabajo universitario |
  | `chapa motor` | 5,0 mm | 29 | TFG (kart) |
  | `tapa patas` | 5,0 mm | 40 | Trabajo universitario |
  | `enlace bieletas` | 5,0 mm | 10 | TFG (kart) |
  | `soporte caliper` | 5,0 mm | 4,8 | TFG (freno) — la más justa |

- **Descartadas por medición o criterio de Esteban:** `soporte naylon`, `soporte dirección` (×3), `soporte mangueta` (chapa plegada en U), `soporte freno`, `pieza34 brida`, `chapa 1` (tiene pliegue) y `soporte transmisión` (34,7 mm de espesor).

> [!warning] El banco se queda en 8, no en 10-15
> Es **todo lo que hay** de placas planas propias: se barrieron 318 piezas del equipo y las 79 planas restantes eran de catálogo (carpetas de proveedor, espesores en pulgadas, duplicados). Consecuencia: el criterio de salida de la Fase 1 pasa de "≥9 de 10" a **"≥7 de 8"**. El banco crecerá con piezas reales que vayan apareciendo.
>
> **Decisión (2026-07-21): NO se rellena con piezas descargadas de internet.** Inflarían la métrica con "la placa de 4 agujeros" —el riesgo nº1 del estudio—, suelen venir como sólidos importados sin árbol de operaciones (y las reglas de acotación necesitan leer features) y carecen de propiedades personalizadas para el cajetín. Uso legítimo alternativo: un conjunto **aparte** de robustez, para probar si el motor sobrevive a geometría importada.
- **Planos de referencia** en `banco-de-pruebas/planos-referencia/` (14 `.SLDDRW` hechos por Esteban). Sirven para rellenar las secciones 2-6 de esta spec. Ojo: al copiarlos fuera de su carpeta original, SolidWorks puede pedir relocalizar el modelo asociado.

> [!note] Familia 2, para el futuro (Nivel 5, no antes)
> El **soporte de manguetas en U** es **chapa plegada**, no placa plana: exige vista de desarrollo, líneas y cotas de plegado y más vistas. Se deja fuera para no contaminar la métrica del 95%, pero es el candidato natural a segunda familia, porque comparte la salida a DXF para láser.

## 2. Vistas que exige el plano

| Vista | ¿Siempre? | Criterio |
|---|---|---|
| Alzado | | |
| Planta / Perfil | | |
| Cortes / Detalles | | (cuándo hace falta un corte, a qué escala un detalle) |

## 3. Cotas funcionales por tipo de feature

| Feature de la familia | Cotas que SIEMPRE lleva | Tolerancia típica |
|---|---|---|
| (p. ej. chavetero) | (ancho, profundidad, posición desde cara de referencia) | (p. ej. ancho P9) |
| (p. ej. alojamiento rodamiento) | (⌀, profundidad) | (H7) |
| | | |

## 4. Normas aplicables

- Tolerancias generales: (p. ej. ISO 2768-mK)
- Representación / acotación: (UNE 1027, ISO 128...)
- Acabados superficiales, si aplica:

## 5. Cajetín

- Plantilla de formato usada: (ruta al .drwdot / .slddrt)
- Campos que se rellenan desde propiedades personalizadas del modelo: (denominación, material, tratamiento, autor, fecha...)

## 6. Qué es un plano "MAL hecho" (checklist inverso)

> Esta lista alimenta el informe de cumplimiento de la Fase 3.

- [ ] (p. ej. chavetero sin cota de posición)
- [ ] (p. ej. ⌀ funcional sin tolerancia)
- [ ] (p. ej. campo de material vacío en cajetín)
- [ ] ...
