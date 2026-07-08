# Mini-spec de la familia de piezas (tarea 0.3)

> Rellenar tras elegir la familia. Este documento es el conocimiento de dominio del proyecto:
> de aquí saldrá la base de reglas (YAML) que usa el motor. Sé concreto, no exhaustivo.

## 1. Familia elegida

- **Nombre:** (p. ej. "ejes con chavetero", "placas taladradas", "chapas plegadas para láser")
- **Por qué esta:** (la delineas a menudo, features enumerables, dolor real)
- **Ejemplos en `banco-de-pruebas/`:** (listar los archivos)

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
