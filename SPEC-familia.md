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

> [!warning] BORRADOR — pendiente de que Esteban lo corrija
> Las secciones 2-6 las redactó Claude **observando dos planos reales de Esteban** exportados a imagen: `placa motor` (KAM-0024, 10/14) y `placa retención soporte trasero` (KAM-0023, 09/14). No son criterios inventados: son los que él aplicó. Falta que confirme qué es regla y qué fue caso particular. Las dudas están al final, en §7.

## 2. Vistas que exige el plano

| Vista | ¿Siempre? | Criterio |
|---|---|---|
| **Alzado** (cara de la placa) | **Sí** | Muestra el contorno y todos los taladros. Es la vista principal y donde van las cotas de posición en el plano de la chapa |
| **Sección / corte** | **Sí, al menos una** | Se usa **corte (D-D, E-E), no vista lateral simple**. Ahí se acota el **espesor**. Una sección vertical para situar taladros en altura; una segunda sección horizontal si hay un agujero central que acotar (⌀50 en E-E) |
| **Isométrica** | **Sí** | Arriba a la derecha, **sin ninguna cota**. Solo para comprender la pieza |
| Planta / perfil estándar | **No** | No aparecen en ninguno de los dos planos: su función la cumplen los cortes |
| Detalles ampliados | No observados | (¿los usas cuando hay algo pequeño?) |

- **Escala:** 1:1 en la placa de 110 mm, 1:2 en la de ~190 mm. Se elige para que la pieza quepa con holgura en el formato.
- Los **ejes de simetría** se dibujan siempre en el alzado (líneas de centro que pasan por los taladros y por el centro de la placa).

## 3. Cotas funcionales por tipo de feature

| Feature | Cotas que SIEMPRE lleva | Dónde se acota | Tolerancia observada |
|---|---|---|---|
| **Contorno de la placa** | Alto total (110, 145) | Alzado | General (ISO 2768-mK) |
| **Grupo de taladros iguales** | **Un solo diámetro con multiplicador**: `4x ⌀8`, `4x ⌀9 H13` | Alzado, con directriz | `H13` cuando es de paso ajustado; sin tolerancia si es de paso libre |
| **Posición de taladros** | Distancia desde el borde de referencia (9, 12, 16) **y** distancia entre ejes (162, 92) | Alzado y sección | `±0,2` / `±0,5` explícitas cuando importa; si no, general |
| **Agujero central grande** | Diámetro (⌀50) | **En la sección**, no en el alzado | General |
| **Espesor** | Valor simple (2, 5) | **Siempre en la sección** | General |

> **Patrón importante:** el ancho total no siempre se acota directamente. En `placa motor` va en cadena `12 + 162 + 16` (borde → eje → eje → borde), es decir, se acota **respecto a los ejes de los taladros**, que es lo funcional para el montaje.

## 4. Normas aplicables

- **Tolerancias generales:** `UNE-EN ISO 2768-mK` — aparece en el cajetín, campo *Tol.Dim*. Coincide con la clase que ya documentamos en la biblioteca.
- **Acabado superficial general:** `Ra 6,3` en el cajetín. **Local:** `Ra 1,6` con símbolo y directriz sobre la cara que lo requiere.
- **Proyección:** primer diedro (europeo), con su símbolo en el cajetín.
- **Soldadura**, cuando la pieza va soldada: proceso `MIG/MAG UNE-EN ISO 13732` con símbolo de cordón (`a3`) y nota aclaratoria (*"soldada a chasis"*).
- **Material** habitual de la familia: `S235JR` (acero estructural de chapa) en los dos planos.

## 5. Cajetín

- **Plantilla:** formato de la **Escuela de Ingeniería Industrial, Universidad de Vigo** (con logo). Falta localizar el `.slddrt`.
- **Campos que rellena:**

  | Campo | Ejemplo | ¿De propiedad del modelo? |
  |---|---|---|
  | Dibujado (fecha + nombre) | 06/06/2026 · Esteban Pinilla Álvarez | Por confirmar |
  | Comprobado (fecha + nombre) | 12/06/2026 · Jacobo González Baldonedo | Por confirmar |
  | Acabado Sup | Ra 6,3 | Por confirmar |
  | Escala | 1:1 / 1:2 | Automático de la vista |
  | Material | S235JR | Sí, propiedad del material |
  | Tol.Dim | UNE-EN ISO 2768-mK | Por confirmar |
  | Proyección | Símbolo primer diedro | Fijo en la plantilla |
  | Plano (denominación) | Placa motor | Por confirmar |
  | Plano Nº | 10/14 | Manual, depende del conjunto |
  | Referencia | KAM-0024 | Código propio `KAM-00XX` |
  | Sustituye a / Sustituido por | (vacíos) | Manual |

## 6. Qué es un plano "MAL hecho" (checklist inverso)

> Alimenta el informe de cumplimiento de la Fase 3.

- [ ] Falta la **sección**: el espesor no está acotado en ninguna vista
- [ ] Un **taladro sin cota de posición** respecto a un borde o eje de referencia
- [ ] Grupo de taladros iguales **acotados uno a uno** en vez de con multiplicador `Nx ⌀`
- [ ] **Agujero de ajuste sin tolerancia** (debería llevar H13 o la que aplique)
- [ ] La **isométrica lleva cotas** (no debe llevar ninguna)
- [ ] Faltan los **ejes de simetría** en el alzado
- [ ] Cajetín con **Material** vacío
- [ ] Cajetín con **Tol.Dim** vacío (sin norma de tolerancias generales no se puede fabricar)
- [ ] Falta **Referencia** o **Plano Nº**
- [ ] Falta el **símbolo de proyección**
- [ ] Falta el **acabado superficial** general
- [ ] Pieza soldada **sin símbolo de soldadura** ni proceso indicado

## 7. Criterios confirmados por Esteban (2026-07-21)

### 7.1 Acotación: **siempre desde los bordes**
Regla cerrada. El origen de cotas es el **borde de la placa**, no el eje de otro taladro.

**Excepción confirmada:** se acota **entre ejes de taladros** cuando lo que hay que garantizar es la **posición relativa entre ellos** (el caso del `162` en `placa motor`, para que el patrón de taladros case con su contrapieza).

**Regla para el motor:** por defecto, todas las cotas de posición desde los bordes. La cota entre ejes se añade **solo** cuando el patrón de taladros tiene que casar con otra pieza, y en ese caso se **marca como propuesta a revisar**, porque el motor no puede saber por sí solo con qué se monta la placa.

### 7.2 Tolerancias: **las propone el motor, Esteban revisa**
Esteban no quiere teclear la tolerancia pieza a pieza. Quiere que el motor **deduzca la tolerancia a partir de la norma ISO y de la función de la pieza** (para qué va a usarse), la proponga, y él la revise y corrija.

Lo mismo aplica al `H13` de los taladros: hoy lo pone según el tornillo y el ajuste, pero es una decisión **derivable** de la función.

**Esto confirma y afina la tarea 3.3 del plan** ("vocabulario de tolerancias por intención"). Implica que la base de conocimiento normativo necesita **tres tablas**, no una:

| Tabla | Contenido | Norma |
|---|---|---|
| **Agujeros de paso** | tornillo `M6/M8/M10…` + serie (fina/media/basta) → diámetro y grado (p. ej. `H13`) | ISO 273 |
| **Ajustes funcionales** | alojamiento de rodamiento → `H7`; eje guía → `g6`; etc. | ISO 286 |
| **Tolerancias generales** | todo lo no acotado con tolerancia propia | ISO 2768-mK |

Y encima de ellas, un **diccionario de intención → ajuste**: frases como *"agujero de paso para M8"*, *"aloja un rodamiento"*, *"va soldada, no importa la precisión"* que el LLM traduce a una entrada de las tablas.

> [!danger] Este es el punto de mayor riesgo del proyecto
> Una tolerancia mal propuesta **cuesta dinero real** (pieza fabricada y tirada). Reglas de diseño innegociables:
> 1. La tolerancia sale **siempre marcada como propuesta**, nunca aplicada en silencio.
> 2. Va acompañada de **la razón** ("⌀9 H13 = agujero de paso para M8, serie media, ISO 273") para que revisarla cueste segundos.
> 3. Si la función de la pieza **no está clara, se cae a ISO 2768-mK general** y se avisa. Nunca se inventa un ajuste.

### 7.3 CONFIRMADO: los ⌀9 H13 son agujeros de paso para M8
Esteban lo confirma. Consecuencia de diseño importante: **el taladro es calculable, no memorizable**. Con saber *"tornillo M8, agujero de paso, serie media"* el motor deriva `⌀9 H13` de la tabla ISO 273.

Esto convierte la primera de las tres tablas en algo **pequeño, cerrado y verificable** (una entrada por métrica y serie), que es el mejor punto por donde empezar la base de conocimiento normativo: bajo riesgo, resultado inmediato y comprobable contra los planos que Esteban ya hizo.

> Pendiente: confirmar la tabla ISO 273 completa (fina / media / basta) para las métricas que use la familia — M4 a M12 probablemente basten.

### 7.4 Dudas que siguen abiertas
1. **Ancho total:** ¿debe aparecer siempre, o basta la cadena que lo define?
2. **Escala:** ¿regla fija (que quepa en A4/A3) o a ojo?
3. **Cortes:** ¿siempre sección, o a veces basta una vista lateral normal?
4. **Ra 1,6 local:** ¿en qué caras lo exiges y por qué?
