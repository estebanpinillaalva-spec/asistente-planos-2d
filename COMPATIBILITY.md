# Compatibilidad

## Entorno pineado

- **Versión de SolidWorks:** (rellenar, p. ej. SolidWorks 2024 SP5)
- **Tipo de licencia y nota EULA (tarea 0.4):** (¿permite uso de API en proyectos personales? ¿restricciones comerciales?)
- **Windows:** 11 Pro
- **Python:** (versión) · **pywin32:** (versión)

> Regla: se desarrolla SOLO contra esta versión. La compatibilidad multi-versión es problema de una eventual fase producto, no del estudio.

## Registro de llamadas API usadas

> Anotar cada interfaz/método COM que use el motor. Si un Service Pack rompe algo, esta tabla dice dónde mirar.

| Interfaz.Método | Para qué se usa | Desde (fase/tarea) | Notas |
|---|---|---|---|
| (p. ej. `ISldWorks.ActiveDoc`) | (leer modelo activo) | (0.5) | |
