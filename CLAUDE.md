# Asistente IA de planos 2D para SolidWorks

## Qué es este proyecto

Proyecto de estudio (~5 meses, 8-10 h/semana) de Esteban, ingeniero industrial que se inicia en IA aplicada. Se construye un **asistente local que parte de un modelo 3D existente en SolidWorks y genera el plano 2D de fabricación**: vistas, acotación semiautomática, tolerancias, anotaciones, cajetín, BOM y exportación DXF/PDF. Se maneja conversacionalmente vía MCP.

Objetivo primario: **portfolio potente + aprender IA aplicada**. Negocio: solo opcionalidad futura (decisión en mes 6 con datos, ver ESTUDIO-Y-PLAN.md §Fase 5).

**NO se construye:** texto→modelo 3D, nada en la nube que reciba geometría, nada horizontal tipo "copiloto para todo SolidWorks".

## Principio rector de arquitectura (innegociable)

**El LLM nunca toca geometría ni emite llamadas crudas a la API.** Solo traduce intención del usuario → JSON estructurado (esquemas cerrados de parámetros) → un motor determinista en Python ejecuta contra la API COM de SolidWorks → las señales de verificación (rebuild OK, features sin acotar, errores COM legibles, checklist de norma) vuelven al LLM para autocorrección.

## Decisiones técnicas cerradas (no re-debatir)

- **Python + pywin32** para el motor (no C#/VBA). VBA solo como oráculo: grabar macro en SolidWorks → leer el VBA generado → traducir a pywin32.
- **MCP como interfaz** (cliente: Claude Desktop/Code). No se construye UI propia en los meses 1-4.
- **Partir de MCP open-source existente** como referencia de cómo resolver COM, no reinventar: `alisamsam/solidworks-mcp`, `eyfel/mcp-server-solidworks`.
- **Versión de SolidWorks pineada** — se desarrolla contra una sola versión; cada llamada API usada se anota en `COMPATIBILITY.md`.
- **Base de conocimiento normativo como datos** (YAML/JSON versionado), no como prompts: reglas UNE 1027 / ISO 128 / ISO 2768 de la familia de piezas. Es la aportación diferencial del proyecto.
- **Una sola familia de piezas** hasta dominarla (definida en `SPEC-familia.md`).

## Reglas de trabajo

1. **Regla del 95%:** no se sube de nivel de la escalera de alcance (ESTUDIO-Y-PLAN.md §4) hasta que el nivel actual funciona en ≥95% de los intentos sobre el banco de pruebas.
2. **Timebox COM:** ninguna pelea de infraestructura/interop dura más de 2 semanas; si se atasca, se rodea o se recorta el alcance.
3. **Entregable demoable cada 2-3 semanas** (script funcionando, vídeo corto). Si no hay demo, se ha sobre-alcanzado.
4. **Medir, no estimar:** fiabilidad y tiempos se cronometran contra el banco de pruebas (10-15 modelos reales en `banco-de-pruebas/`).
5. Esteban es novato en programación: explicar las decisiones de código, preferir soluciones simples y legibles sobre clever, y avanzar en pasos pequeños verificables.

## Herramientas y flujo de sesión

1. **Al iniciar cada sesión de este proyecto, invoca la skill `fable-mode`** (Esteban trabaja este proyecto con Opus + fable mode: rigor, verificación con evidencia, sin suposiciones).
2. **Generación de código sustancial → skill `codex-delegate`**: escribir un brief autocontenido, ejecutar `codex exec`, y hacer revisión frugal (resumen + diff stat + correr los criterios de aceptación). El objetivo es ahorrar tokens de Claude: Codex escribe, Claude especifica y verifica. Ediciones pequeñas (<~20 líneas), depuración con contexto de conversación y decisiones de diseño las hace Claude directamente.
3. **Esteban a veces usa Codex por su cuenta desde la extensión de VS Code** (no siempre): reanudar sesiones de Codex por session id, no con `--last` (podría ser la suya). Si está trabajando un archivo con Codex en ese momento, no tocarlo en paralelo.
4. **Claude es el revisor del código**: todo código sustancial (delegado a Codex o no) pasa `/code-review` sobre el diff antes de commitear; los hallazgos se corrigen (reanudando la sesión de Codex si fue suyo) antes del commit. A partir de la Fase 2, los hitos se integran vía rama + PR en GitHub y la revisión se hace sobre la PR (`/review`), dejando evidencia pública del proceso para el portfolio.

## Estado actual

- **Fase 0 (Validación y encuadre)** — recién iniciada. Ver plan completo por fases en `ESTUDIO-Y-PLAN.md` §7.
- **Repo:** https://github.com/estebanpinillaalva-spec/asistente-planos-2d — privado por ahora; pasar a público en Fase 4, cuando el portfolio esté presentable. Los modelos de `banco-de-pruebas/` están en `.gitignore` por **peso** (binarios CAD de ~2,4 MB que cambian enteros en cada guardado), no por confidencialidad: son piezas de trabajos universitarios de Esteban y **no hay material confidencial**, así que en Fase 4 se pueden publicar libremente capturas, DXF y PDF derivados de ellas.
- Próxima tarea: **0.1 — test de 48h** con un MCP de SolidWorks existente (instrucciones en `EMPEZAR-AQUI.md`).

## Mapa de documentos

| Archivo | Qué es |
|---|---|
| `ESTUDIO-Y-PLAN.md` | Estudio del reshape + plan de acción completo por fases. Documento maestro |
| `EMPEZAR-AQUI.md` | Arranque de sesión y primer encargo (test 48h) |
| `SPEC-familia.md` | Mini-spec de la familia de piezas (tarea 0.3, por rellenar) |
| `COMPATIBILITY.md` | Versión de SolidWorks y registro de llamadas API usadas |
| `banco-de-pruebas/` | 10-15 modelos reales de la familia para medir fiabilidad |
| `notas/test-48h.md` | Resultados del test de 48h (tarea 0.1, por rellenar) |
