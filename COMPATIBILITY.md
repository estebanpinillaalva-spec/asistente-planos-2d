# Compatibilidad

## Entorno pineado

- **Versión de SolidWorks:** `RevisionNumber` = **31.2.1** (la serie 31 corresponde a SolidWorks 2023 — año comercial inferido, no leído de la API). **Esta es la versión pineada del proyecto.**
- **Tipo de licencia y nota EULA (tarea 0.4):** (¿permite uso de API en proyectos personales? ¿restricciones comerciales?)
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
