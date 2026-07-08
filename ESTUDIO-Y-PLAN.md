# Estudio de Reshape y Plan de Acción — Asistente IA para Planos 2D en SolidWorks

> Autor: Esteban (ingeniero industrial) · Fecha: julio 2026
> Origen: veredicto RESHAPE del consejo de evaluación (roast). Decisión tomada: ir a por planos 2D.

---

## 1. Resumen ejecutivo

**Qué se construye:** un asistente local que, partiendo de un modelo 3D existente en SolidWorks, genera y completa el plano 2D de fabricación: vistas, acotación semiautomática, tolerancias, anotaciones, cajetín, lista de materiales y exportaciones (DXF/PDF). Se maneja conversacionalmente ("hazme el plano de esta pieza según ISO 2768-mK, el taladro central es H7") pero **el LLM nunca toca geometría**: traduce intención a parámetros estructurados que ejecuta código determinista contra la API de SolidWorks.

**Qué NO se construye:** texto → modelo 3D (batalla perdida contra Zoo, Leo AI y la propia Dassault), nada en la nube que reciba geometría del cliente, y nada horizontal tipo "copiloto para todo SolidWorks".

**Por qué esta forma:** es el punto exacto donde coinciden las cinco voces del roast — técnicamente alcanzable hoy por una persona a tiempo parcial (Logician 7/10), con hueco de mercado real que Dassault deja descubierto (Researcher 6/10), donde el cliente abre la cartera (Buyer: 20-50 €/mes por esto, 0 € por texto→3D), donde vive tu ventaja de dominio (Expansionist 9/10), y con alcance recortable hasta que la fiabilidad sea aburrida (la única forma de esquivar el 3/10 del Contrarian).

**Objetivo del proyecto:** portfolio potente + aprendizaje de IA aplicada en ~5 meses a tiempo parcial, con puerta abierta (y criterios de decisión explícitos, §9) a consultoría o producto de nicho.

---

## 2. El reshape en detalle: de "asistente CAD" a "delineante IA"

### 2.1 Reformulación del producto

| | Idea original | Reshape |
|---|---|---|
| **Input** | Texto libre | Modelo 3D existente + intención conversacional |
| **Output** | Geometría 3D "desde cero" | Plano 2D normativo auditable (vistas, cotas, tolerancias, cajetín, BOM, DXF/PDF) |
| **Rol del LLM** | Autor de la geometría | Traductor de intención → JSON estructurado → código determinista |
| **Dónde corre** | Indefinido (¿nube?) | Local, sobre el SolidWorks de escritorio del usuario |
| **Dominio** | Cualquier pieza | UNA familia de piezas acotada al principio; se amplía solo cuando la fiabilidad es ≥95% |
| **Propuesta de valor** | "La IA diseña por ti" (nadie lo quiere) | "La IA te quita la delineación aburrida y te la deja lista para revisar cota a cota" |

### 2.2 Por qué planos 2D es el nicho correcto (análisis del dominio)

1. **Es el dolor real, no el espectacular.** Ningún ingeniero sufre modelando en 3D — croquizar un eje lleva 4 minutos. Lo que se odia es acotar según UNE 1027 / ISO 128, poner tolerancias ISO 2768, montar cajetines, sacar despieces y DXF. Es trabajo de horas, repetitivo, y facturado a precio de delineante.

2. **Es mucho más determinista que el 3D.** La geometría ya existe en el modelo; generar el plano es *aplicar reglas* (qué vistas, qué cotas son funcionales, qué norma aplica). Los LLMs son malos inventando geometría y buenos rellenando esquemas de parámetros y aplicando criterio sobre opciones enumeradas. El problema encaja con la herramienta.

3. **Tu ventaja competitiva vive exactamente aquí.** El 99% de la gente de IA que toca CAD no sabe qué es una cota funcional, por qué ⌀20 H7 no es "un agujero de 20", ni qué exige ISO 128 de una vista de corte. Tú sí. Esa asimetría no existe en el texto→3D (donde tu ventaja es cero frente a un equipo de ML).

4. **El hueco de mercado está documentado.** SOLIDWORKS 2026 mete IA vía AURA, pero exige 3DExperience Connected — fuera queda ~85% de la base instalada de escritorio. Los ~5 servidores MCP de SolidWorks en GitHub son wrappers genéricos de la API, ninguno entiende de delineación normativa. DriveWorks demostró hace décadas que la automatización de SolidWorks es un negocio (sin IA y a precio caro); la capa conversacional + criterio normativo es lo nuevo.

5. **El output es auditable, que es la condición del comprador.** El Buyer del roast fue tajante: un asistente que acierta el 90% le obliga a revisar el 100% igual, así que el ahorro tiende a cero *salvo que el output sea revisable cota a cota*. Un plano semiacotado con informe de cumplimiento normativo es exactamente eso: propuesta auditable, no caja negra.

### 2.3 A quién sirve (por orden de prioridad del proyecto)

1. **A ti mismo** (meses 1-4): tus propias piezas del trabajo. Eres usuario cero y validador.
2. **Empleadores / entrevistas** (mes 4+): el portfolio público es el producto principal de esta fase.
3. **Despachos y talleres pequeños** (mes 6+, opcional): los que pagan delineación por horas y no tienen DriveWorks ni 3DExperience. Solo si decides activar la vía negocio (§9).

---

## 3. Arquitectura técnica

### 3.1 Principio rector (del Logician, innegociable)

```
Usuario (lenguaje natural)
   │
   ▼
LLM  ──traduce──▶  JSON estructurado (esquema cerrado de parámetros)
   │                    │
   │                    ▼
   │             Motor determinista (Python) ──▶ API COM de SolidWorks
   │                    │
   │                    ▼
   └──corrige◀── Señales de verificación (rebuild OK, nº cotas,
                 features sin acotar, checklist de norma, errores COM)
```

El LLM **nunca** emite llamadas crudas a la API ni coordenadas. Rellena esquemas como:

```json
{
  "accion": "acotar_vista",
  "vista": "alzado",
  "norma_tolerancias": "ISO 2768-mK",
  "cotas_funcionales": [
    {"feature": "taladro_central", "tolerancia": "H7", "critica": true}
  ],
  "estilo": "UNE_1027"
}
```

### 3.2 Componentes

| Componente | Qué hace | Tecnología |
|---|---|---|
| **Motor de delineación** | Toda la lógica determinista: crear plano desde plantilla, insertar vistas, acotar, rellenar cajetín desde propiedades personalizadas, exportar | Python + `pywin32` (COM) contra la API de SolidWorks (`DrawingDoc`, `ModelDoc2`, `View`, `DisplayDimension`, autodimensioning, `SaveAs` para DXF/PDF) |
| **Servidor MCP** | Expone el motor como herramientas de ALTO nivel (`crear_plano`, `insertar_vistas`, `acotar_vista`, `verificar_norma`, `exportar_dxf`) — no un passthrough de la API | Python (SDK MCP), arrancando desde un servidor open-source existente como base de estudio |
| **Capa LLM** | Conversación, traducción de intención, decisión sobre opciones enumeradas, autocorrección con las señales de error | Claude vía MCP (Claude Desktop / Claude Code como cliente — no construyes UI en meses 1-4) |
| **Verificador** | Rebuild status, detección de features sin cota, recuento de cotas vs. esperadas, checklist ISO/UNE, log auditable de cada acción | Python, dentro del motor |
| **Base de conocimiento normativo** | Reglas UNE 1027 / ISO 128 / ISO 2768 codificadas como datos (no prompts): qué cotas exige cada tipo de feature, criterios de vistas, tolerancias generales | YAML/JSON versionado — **este es tu foso y tu aportación única** |

### 3.3 Decisiones técnicas ya tomadas (para no re-debatirlas cada semana)

- **Python, no C#/VBA.** Curva de aprendizaje menor, ecosistema MCP nativo, y tu objetivo es aprender IA aplicada, no .NET. VBA solo para espiar cómo hace algo el grabador de macros de SolidWorks (truco: grabar macro → leer el VBA generado → traducir a pywin32).
- **MCP como interfaz, no app propia.** Cero trabajo de UI, compatible con cualquier cliente LLM, y publicar "el primer MCP de SolidWorks especializado en planos normativos" es en sí el titular del portfolio.
- **Empezar desde un MCP existente, no desde cero.** `alisamsam/solidworks-mcp` (22 herramientas) o `eyfel/mcp-server-solidworks` como referencia de cómo resolver COM. El Contrarian tiene razón: la API COM te puede comer el 80% de las horas — no pagues ese peaje dos veces.
- **Versión de SolidWorks pineada.** Desarrollas y documentas contra UNA versión (la que tengas). Los Service Packs rompen macros; la compatibilidad multi-versión es problema de la fase negocio, no del estudio.
- **LLM por API en fase estudio; local/on-prem solo si hay fase producto.** Para portfolio con tus propias piezas no hay problema de confidencialidad. La línea roja del Buyer (nada de geometría de clientes a la nube) se resuelve en fase producto con modelo local o garantías contractuales — se documenta como requisito, no se implementa ahora.

---

## 4. Escalera de alcance (la defensa contra el abismo demo→útil)

Regla: **no se sube de nivel hasta que el actual funciona en el ~95% de los intentos sobre la familia de piezas elegida.** Cada nivel produce algo demoable en vídeo.

| Nivel | Qué hace | Riesgo técnico | Valor portfolio |
|---|---|---|---|
| **0** | Test de 48h: MCP existente + Claude sobre una pieza real tuya. Sin escribir código | Nulo | Mapa de dónde está el muro |
| **1** | Modelo 3D → esqueleto de plano: plantilla, 3 vistas estándar, escala, cajetín relleno desde propiedades | Bajo (todo determinista) | "Script que monta el plano en 30 segundos" |
| **2** | Acotación semiautomática: autodimension + reglas propias; el plano sale *semiacotado* para revisión humana | Medio | Primer vídeo con IA conversacional real |
| **3** | Criterio normativo: tolerancias por intención ("el alojamiento es H7"), ISO 2768 aplicada, **informe de cumplimiento de norma** | Medio-alto | El diferenciador. Nadie más tiene esto |
| **4** | Flujos batch: DXF para láser de todas las chapas, paquete PDF, BOM, familias de piezas | Bajo (repetición del motor) | "Ahorra una tarde entera" — lo facturable |

La **familia de piezas** se elige en la Fase 0 de entre tu trabajo real. Criterios: piezas que tú mismo delineas a menudo, con features enumerables (ejes con chaveteros, bridas, placas taladradas, chapas plegadas para láser…). Una sola familia. Ampliar familias es Nivel 5, no antes.

---

## 5. Riesgos y mitigaciones

| Riesgo (fuente del roast) | Probabilidad | Mitigación |
|---|---|---|
| **Infierno COM**: 80% del tiempo peleando con interop (Contrarian) | Alta | Partir de MCP open-source existente; grabadora de macros como oráculo; timebox de 2 semanas a cualquier pelea de infraestructura — si algo COM se atasca más, se rodea o se recorta |
| **Abismo demo→útil**: funciona con la placa de 4 agujeros, muere con piezas reales (Contrarian) | Alta | Escalera de alcance §4 + regla del 95% + familia única de piezas |
| **Muerte silenciosa por motivación** a las 6-8 semanas (Contrarian) | Media | Entregable demoable cada 2-3 semanas (§7); repo público desde el día 1 (presión social positiva); vídeos cortos como hitos |
| **Doble curva de aprendizaje** (Python + IA a la vez, siendo novato) | Media | Fase 1 es Python puro sin IA (una curva cada vez); Claude Code como tutor/pair-programmer durante todo el proyecto |
| **Trampa del 90%**: si hay que revisarlo todo, no ahorra nada (Buyer) | Media | El output es *propuesta auditable*: plano semiacotado + log de acciones + informe normativo. La revisión ES el flujo, no un fallo del flujo |
| **Licencia SolidWorks**: las educativas/de empresa prohíben uso comercial | — | Para el estudio/portfolio con licencia propia o de trabajo: verificar EULA (acción en Fase 0). Para fase negocio: licencia comercial propia o programa de partners — es problema del mes 6+, no de ahora |
| **Dassault te copia** (AURA baja al escritorio) | Baja a 12 meses | Es riesgo de la vía producto, no del estudio. Tu portfolio y aprendizaje sobreviven a cualquier movimiento de Dassault; y si pasa, valida que el nicho era real |
| **Service Pack rompe la API** (Buyer) | Media | Versión pineada (§3.3); anotar cada llamada API usada en un `COMPATIBILITY.md` |

---

## 6. Supuestos del plan

- Dedicación: **8-10 h/semana** (tiempo parcial real). Si tienes más, comprime el calendario; no amplíes el alcance.
- Tienes acceso a SolidWorks en una máquina Windows donde puedes instalar Python.
- Presupuesto: **< 30 €/mes** (suscripción/API de Claude). Todo lo demás es gratis (Python, GitHub, MCP SDK).
- Nivel de programación de partida: básico o nulo — el plan lo asume y por eso la Fase 1 es más larga de lo que parecería.

---

## 7. Plan de acción detallado

### Fase 0 — Validación y encuadre (Semanas 1-2, ~16-20 h)

**Objetivo:** saber dónde está el muro antes de construir nada, y fijar el encuadre del proyecto.

| # | Tarea | Resultado |
|---|---|---|
| 0.1 | **Test de 48 h**: instalar un MCP de SolidWorks existente (empezar por `alisamsam/solidworks-mcp`; alternativa `eyfel/mcp-server-solidworks`), conectarlo a Claude Desktop/Code, e intentar producir una pieza y un plano reales de tu trabajo | Documento de 1 página: qué funcionó, qué no, dónde duele COM, qué NO hace ningún MCP existente |
| 0.2 | Elegir **la familia de piezas** (criterios en §4) y reunir 10-15 modelos reales de esa familia como banco de pruebas | Carpeta `banco-de-pruebas/` con los modelos |
| 0.3 | Escribir la **mini-spec** (1-2 páginas): para esa familia, qué vistas, qué cotas son funcionales, qué norma aplica, qué es un plano "bueno". Esto es tu conocimiento de dominio volcado a papel — será la base de conocimiento del §3.2 | `SPEC-familia.md` |
| 0.4 | Verificar el **EULA de tu licencia** de SolidWorks respecto a uso de API y proyectos personales | Nota en el repo |
| 0.5 | Crear repo público en GitHub (README honesto: "proyecto de estudio, en construcción"), instalar Python + pywin32, y conseguir el "hola mundo" COM: script que abre SolidWorks y lee el nombre del modelo activo | Repo con primer commit funcional |

**Criterio de salida:** sabes qué hay que construir que no exista ya, tienes banco de pruebas, y Python habla con SolidWorks en tu máquina.

**⚠️ Punto de decisión:** si el test 48h revela que COM es inviable en tu máquina/licencia, parar y reevaluar (alternativa: automatización vía Document Manager API o incluso FreeCAD como plan B de aprendizaje). No seguir por inercia.

---

### Fase 1 — Motor determinista, sin IA (Semanas 3-6, ~36-40 h)

**Objetivo:** script Python que convierte modelo 3D → esqueleto de plano. Cero LLM. Una sola curva de aprendizaje: programar.

| # | Tarea |
|---|---|
| 1.1 | Función: abrir modelo, crear plano desde tu plantilla (formato, cajetín) |
| 1.2 | Función: insertar vistas estándar (alzado, planta, perfil) con escala correcta; vista isométrica opcional |
| 1.3 | Función: rellenar cajetín desde propiedades personalizadas del modelo (denominación, material, autor, fecha) |
| 1.4 | Función: exportar a PDF y DXF |
| 1.5 | Correr contra el banco de pruebas completo: **10-15 piezas, medir % de éxito y tiempo** |
| 1.6 | Truco de método permanente: para cada operación nueva, grabar macro en SolidWorks → leer el VBA → traducir a pywin32 con ayuda de Claude Code |

**Entregable demoable (semana 6):** vídeo de 60s — "de modelo a plano con cajetín y 3 vistas en 30 segundos, en 10 piezas distintas".
**Criterio de salida:** ≥ 9 de 10 piezas del banco pasan sin intervención manual.

---

### Fase 2 — Capa MCP + intención conversacional (Semanas 7-10, ~36-40 h)

**Objetivo:** Claude maneja el motor conversacionalmente. Aquí empieza la IA aplicada de verdad.

| # | Tarea |
|---|---|
| 2.1 | Aprender el SDK Python de MCP (tutorial oficial, 1 tarde) y envolver las funciones de Fase 1 como herramientas MCP de alto nivel |
| 2.2 | Diseñar los **esquemas JSON** de cada herramienta (el contrato LLM↔motor, §3.1) — dedicarle tiempo: es la decisión de diseño más importante del proyecto |
| 2.3 | Implementar el **bucle de verificación**: cada herramienta devuelve señales estructuradas (éxito, warnings, features sin procesar, errores COM legibles) para que Claude se autocorrija |
| 2.4 | Probar flujos conversacionales completos: "ábreme la pieza X y móntame el plano con alzado y planta a escala 1:2" |
| 2.5 | Log auditable: cada acción del asistente queda registrada (qué hizo, con qué parámetros, resultado) |

**Entregable demoable (semana 10):** vídeo — "le hablo a Claude y me monta el plano; cuando se equivoca, lee el error y se corrige solo".
**Criterio de salida:** 8 de 10 peticiones conversacionales sobre el banco de pruebas terminan en plano correcto sin tocar tú SolidWorks.

---

### Fase 3 — El diferenciador: acotación y criterio normativo (Semanas 11-16, ~50-60 h)

**Objetivo:** lo que nadie más tiene. La fase más difícil y la más valiosa. Aquí se gasta el presupuesto de riesgo.

| # | Tarea |
|---|---|
| 3.1 | Acotación semiautomática: combinar el autodimensioning de la API con **tus reglas** de la mini-spec (qué cotas exige cada feature de la familia) |
| 3.2 | Codificar la base de conocimiento normativo como datos (YAML): ISO 2768 (tolerancias generales), criterios UNE 1027/ISO 128 de la familia |
| 3.3 | Vocabulario de tolerancias por intención: "el alojamiento del rodamiento es H7" → cota con tolerancia en la vista correcta |
| 3.4 | **Informe de cumplimiento**: checklist automático — features sin acotar, cotas duplicadas, tolerancias faltantes en features funcionales, campos de cajetín vacíos. Sale como reporte auditable junto al plano |
| 3.5 | Medir contra el banco: % de cotas correctas propuestas, % de falsos positivos del checklist, **tiempo humano de "revisar y terminar" vs. delinear desde cero** (esta métrica es EL resultado del estudio) |

**Entregable demoable (semana 16):** el vídeo estrella del portfolio — plano semiacotado según norma + informe de cumplimiento, sobre una pieza real, con cronómetro.
**Criterio de salida:** el flujo asistido es ≥ 50% más rápido que delinear a mano en tu familia de piezas, medido en ≥ 5 piezas.
**⚠️ Válvula de escape:** si en la semana 14 la acotación automática no converge, recortar a "acotación 100% dirigida por conversación + checklist normativo" — sigue siendo portfolio de sobra.

---

### Fase 4 — Batch, pulido y publicación (Semanas 17-20, ~30-36 h)

**Objetivo:** convertir el proyecto técnico en activo de portfolio con audiencia.

| # | Tarea |
|---|---|
| 4.1 | Flujo batch: todas las chapas de un ensamblaje → DXF para láser + paquete PDF + BOM (es repetir el motor: barato y muy vistoso) |
| 4.2 | README profesional del repo: problema, arquitectura (diagrama del §3.1), demos en GIF, métricas medidas, limitaciones honestas |
| 4.3 | 3 vídeos cortos (los de las fases 1-3, regrabados con calidad) + 1 vídeo de 3 min contando el proyecto entero |
| 4.4 | **Contenido en español**: 2-3 posts (LinkedIn/blog) — "IA + CAD en español" es un nicho casi vacío con audiencia enorme de ingenieros (Expansionist). Documentar el viaje, no solo el resultado |
| 4.5 | Compartir en comunidades: r/SolidWorks, foros de SolidWorks, comunidad MCP, LinkedIn de ingeniería española |

**Entregable:** portfolio público completo y difundido.
**Criterio de salida:** el repo y los vídeos existen, y ≥ 3 conversaciones iniciadas con ingenieros/empresas que reaccionaron al contenido.

---

### Fase 5 — Puerta de decisión (Mes 6)

Con datos en la mano, decidir entre tres vías (no antes — decidir ahora sería especular):

| Vía | Señal que la activa | Siguiente paso |
|---|---|---|
| **A. Solo portfolio** | El proyecto gustó pero nadie preguntó "¿me lo instalas?" | Usarlo en entrevistas/CV; mantener el repo; pasar al siguiente proyecto de IA |
| **B. Consultoría/freelance** | Despachos o talleres preguntan por automatizaciones concretas | Ofrecer automatización de macros con IA por proyecto (hay demanda tipo Fiverr documentada); primer euro sin construir producto |
| **C. Producto de nicho** | ≥ 3 personas dicen "pagaría X €/mes por esto" espontáneamente | Entonces sí: requisitos de la línea roja (LLM local/on-prem o garantías), licencia comercial, pricing 20-50 €/mes, y un mini plan de negocio nuevo |

---

## 8. Métricas de éxito del estudio (definir "viable" por adelantado)

1. **Fiabilidad:** ≥ 95% de éxito del motor determinista y ≥ 80% del flujo conversacional completo, sobre el banco de pruebas.
2. **Ahorro medido:** flujo asistido ≥ 50% más rápido que delineación manual en ≥ 5 piezas reales (cronometrado, no estimado).
3. **Aprendizaje:** al final sabes explicar y has implementado: integración LLM-herramientas (MCP), esquemas estructurados, bucles de verificación/autocorrección, y evaluación con banco de pruebas. Eso es "IA aplicada" real, no un wrapper.
4. **Señal externa:** ≥ 3 conversaciones genuinas (no likes) generadas por el portfolio.
5. **Decisión informada:** en el mes 6 puedes elegir A/B/C del §7-Fase 5 con datos, no con fe.

Si se cumplen 1-3, el estudio es un éxito **aunque** la señal externa sea nula: el objetivo primario era portfolio + aprendizaje.

---

## 9. Presupuesto

| Concepto | Coste |
|---|---|
| Claude (Pro o API para MCP) | ~20-25 €/mes × 5 meses ≈ 100-125 € |
| Python, pywin32, MCP SDK, GitHub | 0 € |
| SolidWorks | Ya disponible (verificar EULA, tarea 0.4) |
| **Total del estudio** | **≈ 125 €** |

---

## 10. Primer paso (este fin de semana)

Tarea 0.1 — el test de 48 horas. Sin escribir una línea de código:

1. Instalar `alisamsam/solidworks-mcp` (o `eyfel/mcp-server-solidworks`) siguiendo su README.
2. Conectarlo a Claude Desktop o Claude Code.
3. Pedirle una pieza sencilla real de tu trabajo y su plano.
4. Anotar: qué funcionó, dónde falló, qué echaste de menos.

Ese documento de una página es el verdadero inicio del proyecto: te dice contra qué muro vas a construir y confirma (o desmiente) el plan entero por 0 €.
