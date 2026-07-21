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
  - Son piezas del **TFG (kart) de Esteban, no del trabajo** → sin problema de confidencialidad, publicables en el portfolio.
- **Ejemplos en `banco-de-pruebas/`** (11 piezas copiadas el 2026-07-21):
  `Chapa delantera` · `chapa motor` · `placa rodamientos` · `Soporte direccion 3` · `soporte dirección 2` · `soporte dirección` · `soporte mangueta` · `soporte naylon` · `soporte transmisión` · `soporte caliper` · `soporte freno`
- **Planos de referencia** en `banco-de-pruebas/planos-referencia/` (14 `.SLDDRW` hechos por Esteban). Sirven para rellenar las secciones 2-6 de esta spec. Ojo: al copiarlos fuera de su carpeta original, SolidWorks puede pedir relocalizar el modelo asociado.

> [!warning] Por revisar por Esteban
> La selección de las 11 piezas se hizo por nombre, sin abrir la geometría. **Confirma que todas son realmente placas planas** y elimina del banco las que no lo sean (p. ej. si `soporte naylon` o `soporte transmisión` no son planas). Objetivo: 10-15 piezas homogéneas.

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
