# Compatibilidad

## Entorno pineado

- **Versión de SolidWorks:** **SOLIDWORKS 2023**, `RevisionNumber` = **31.2.1**. **Versión pineada del proyecto.** El año quedó **confirmado** (ya no inferido) por la ruta de instalación `C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2023\`.
- **Plantillas de dibujo de la familia:** `ig_2017_a2/a3/a4_-_iso.slddrt` (formatos de hoja de Ingeniería Gráfica, EPI Universidad de Vigo) en `Desktop\TFG DOCUMENTACION\Plantillas_TFG (1)`. Tres tamaños ⇒ el motor deberá elegir formato y escala según el tamaño de la pieza.
- **Tipo de licencia (tarea 0.4, 2026-07-21): EDUCATIVA.** Confirmado por Esteban en *? → Acerca de SOLIDWORKS*.
  - **Uso académico y personal: permitido.** Las fases 0-4 (estudio, aprendizaje y portfolio) no son uso comercial, así que el proyecto puede continuar tal cual.
  - **Uso comercial: prohibido.** Afecta a la [Fase 5](ESTUDIO-Y-PLAN.md) vías B (consultoría) y C (producto): exigirían una licencia comercial propia o un programa de partners. Es problema del mes 6, no de ahora.
  - **Uso de la API:** el condicionante es el *fin* (no comercial), no el mecanismo. Automatizar con pywin32 para aprender y para el portfolio encaja.
  - **Marcado de archivos:** las licencias educativas suelen estampar *"SOLIDWORKS Educational Product. For Instructional Use Only"* en los planos. En los PNG exportados el 2026-07-21 **no apareció**; conviene comprobarlo en un PDF impreso antes de grabar los vídeos del portfolio.
  - ⚠️ El texto vinculante es el `eula.rtf` instalado (`...\lang\spanish\eula.rtf`). No se ha leído íntegro: la extracción automática del RTF falló. **Releerlo antes de tomar cualquier decisión comercial.**
- **Windows:** 11 Pro
- **Python:** 3.14.4 (venv del proyecto en `.venv/`) · **pywin32:** 312 (wheel `cp314`)
- **COM registrado:** `SldWorks.Application` presente en `HKLM:\SOFTWARE\Classes` ✅

> Regla: se desarrolla SOLO contra esta versión. La compatibilidad multi-versión es problema de una eventual fase producto, no del estudio.

## Registro de llamadas API usadas

> Anotar cada interfaz/método COM que use el motor. Si un Service Pack rompe algo, esta tabla dice dónde mirar.

| Interfaz.Método | Para qué se usa | Desde (fase/tarea) | Notas |
|---|---|---|---|
| `ISldWorks.RevisionNumber` | Leer versión/build de SolidWorks | 0.5 | |
| `ISldWorks.ActiveDoc` | Obtener el documento activo | 0.5 | Devuelve `None` si no hay documento abierto |
| `ISldWorks.Visible` | Hacer visible SolidWorks al arrancarlo | 0.5 | ⚠️ **SIN VERIFICAR** — solo en la ruta `--iniciar`, que nunca se ha ejecutado |
| `IModelDoc2.GetTitle` | Nombre del documento activo | 0.5 | |
| `IModelDoc2.GetPathName` | Ruta completa del documento | 0.5 | Vacío si el documento no se ha guardado nunca |
| `IModelDoc2.GetType` | Tipo de documento | 0.5 | 1=pieza · 2=ensamblaje · 3=plano |
| `ISldWorks.OpenDoc6` | Abrir una pieza en silencio para medirla | 0.2 | Devuelve errores/warnings por referencia |
| `IPartDoc.GetPartBox` | Caja envolvente de la pieza | 0.2 | Devuelve 6 valores **en metros**: xmin,ymin,zmin,xmax,ymax,zmax |
| `IPartDoc.GetBodies2` | Sólidos de la pieza | 0.2 | Solo como plan B si `GetPartBox` falla |
| `IBody2.GetBodyBox` | Caja envolvente de un sólido | 0.2 | Plan B de `GetPartBox` |
| `ISldWorks.CloseDoc` | Cerrar el documento tras medirlo | 0.2 | Imprescindible para no acumular piezas abiertas |
| `ISldWorks.GetDocuments` | Listar documentos ya abiertos | 0.3 | Para **no cerrar** los que el usuario tenía abiertos |
| `IModelDoc2.Extension` | Acceso a `IModelDocExtension` | 1.4 | Devuelve objeto COM → acceso directo, no `leer()` |
| `IModelDocExtension.SaveAs` | Exportar plano a PDF/PNG | 1.4 | Vía preferente; formato deducido de la extensión |
| `IModelDoc2.SaveAs` | Exportar plano (plan B) | 1.4 | Es la que acabó funcionando en la 31.2.1 |
| `ISldWorks.SetUserPreferenceIntegerValue` | Fijar resolución de imagen | 1.4 | `swTiffPrintDPI` (=9) a 300 DPI |

> Conexión: `win32com.client.GetActiveObject("SldWorks.Application")` para engancharse a una instancia abierta; `Dispatch(...)` solo con `--iniciar`. Los errores se capturan como `pywintypes.com_error`.
