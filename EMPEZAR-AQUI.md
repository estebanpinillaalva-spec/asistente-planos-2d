# Empezar aquí

## Cómo trabajar en este proyecto

1. Abre una terminal (o VS Code) **en esta carpeta**: `C:\Users\esteb\Desktop\asistente-planos-2d`
2. Arranca Claude Code desde aquí (`claude` en la terminal, o abre la carpeta en VS Code con la extensión). `CLAUDE.md` se carga solo como contexto del proyecto.
3. Dile a Claude en qué tarea del plan estás (p. ej. "vamos con la tarea 0.1") y trabajad contra `ESTUDIO-Y-PLAN.md`.

## Primer encargo: Tarea 0.1 — Test de 48 horas

**Objetivo:** saber dónde está el muro técnico antes de escribir una sola línea de código, usando un MCP de SolidWorks que ya existe.

**Prompt sugerido para la primera sesión:**

> Lee CLAUDE.md y ESTUDIO-Y-PLAN.md. Vamos a hacer la tarea 0.1 (test de 48h):
> ayúdame a instalar y configurar uno de estos servidores MCP de SolidWorks
> y conectarlo a Claude, y luego probamos a generar una pieza y un plano reales:
> - https://github.com/alisamsam/solidworks-mcp (22 herramientas, primera opción)
> - https://github.com/eyfel/mcp-server-solidworks (alternativa: SolidPilot)

**Requisitos previos:** SolidWorks instalado y abierto en esta máquina · Python instalado · Claude Code o Claude Desktop.

**Qué anotar durante la prueba** (en `notas/test-48h.md`):
- Qué funcionó a la primera y qué no.
- Dónde dolió la instalación / el COM.
- Qué le falta al MCP existente para ser útil de verdad (→ eso es exactamente nuestro proyecto).
- Si intentó generar un plano 2D: ¿qué hizo bien y qué hizo mal según tu criterio de ingeniero?

## Después del test (resto de Fase 0)

- **0.2** Elegir la familia de piezas y copiar 10-15 modelos reales a `banco-de-pruebas/`.
- **0.3** Rellenar `SPEC-familia.md` con tu criterio de qué es un plano "bueno" para esa familia.
- **0.4** Verificar el EULA de tu licencia de SolidWorks (uso de API / proyectos personales) y anotarlo en `COMPATIBILITY.md`.
- **0.5** Crear repo público en GitHub y conseguir el "hola mundo" COM: script Python que lee el nombre del modelo activo.

**⚠️ Punto de decisión de Fase 0:** si el test revela que COM es inviable en tu máquina/licencia, parar y reevaluar con Claude (plan B: Document Manager API, o FreeCAD como vía de aprendizaje). No seguir por inercia.
