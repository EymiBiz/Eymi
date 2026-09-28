# Dashboard de Compras en Power BI: diseño del proyecto

Contexto: empresa industrial peruana (manufactura + distribución + servicios). Exportación del ERP con **proveedor, fecha, centro de costo, tipo de movimiento, cantidad y monto (S/)**, periodo 2023–2026.

---

## 0. Antes de los gráficos: datos y modelo

### 0.1 Qué validar en la exportación del ERP
| Tema | Por qué importa | Qué hacer |
|---|---|---|
| **IGV** | Si un monto incluye IGV y otro no, las comparaciones salen mal. | Trabaja **sin IGV** (valor neto) y escríbelo en el subtítulo: "Montos en S/ sin IGV". |
| **Tipo de movimiento** | Mezcla ingresos, devoluciones, anulaciones y ajustes. Si sumas todo, inflas el gasto. | Arma una tabla de tipos de movimiento con una columna `Signo` (+1 / −1 / 0). |
| **Compras en USD** | Muchas compras industriales se facturan en dólares. | Si el ERP ya convirtió a soles, anota qué tipo de cambio usó (el SBS del día o el de la factura). |
| **Nº de OC / documento** | Sin este número no puedes contar órdenes ni medir el ticket promedio. | Pide que la exportación lo incluya. |
| **Material / familia** | Sin categoría no hay análisis de *spend* por categoría, que es lo más valorado en procurement. | Pide el código de material o grupo de artículos. También puedes mapearlo tú a una clasificación estándar como UNSPSC. |
| **Fechas del ciclo** | Para calcular el **tiempo de atención** necesitas dos fechas, por ejemplo solicitud → OC → recepción. Una sola fecha no basta. | Pide la fecha de solicitud (SOLPED) o de OC, más la fecha de recepción. Si el ERP no las tiene, deja este KPI como "fase 2". |
| **2026 parcial** | El año todavía no termina y el mes en curso está incompleto. | Compara siempre YTD contra el mismo periodo del año anterior y calcula MoM con el **último mes cerrado**. |

### 0.2 Modelo estrella (buena práctica que los reclutadores revisan)
```
                 Dim_Calendario (2023-01-01 → 2026-12-31, marcada como tabla de fechas)
                        │
Dim_Proveedor ── Fact_Compras ── Dim_CentroCosto (→ Unidad de negocio: Manufactura / Distribución / Servicios)
                        │
                 Dim_TipoMovimiento (Descripción, Signo, EsDevolución)
```
- **Fact_Compras**: FechaKey, ProveedorKey, CentroCostoKey, TipoMovKey, NroOC, Cantidad, MontoPEN.
- **Dim_Proveedor**: RUC, razón social, rubro, ciudad/país (local vs importado), fecha de primera compra.
- **Dim_CentroCosto**: código, nombre, área y **unidad de negocio**. Esta es la dimensión que conecta el dashboard con la estructura de la empresa.
- Limpia los datos en **Power Query**: nombres de proveedor unificados por RUC, tipos de dato correctos y columnas que no usas eliminadas.
- Pon todas las medidas en una tabla `_Medidas`, organizadas en carpetas: Gasto, Proveedores, Tendencia, Operación.

### 0.3 Medidas DAX base
```DAX
Gasto Neto = SUMX ( Fact_Compras, Fact_Compras[MontoPEN] * RELATED ( Dim_TipoMovimiento[Signo] ) )

Gasto AA = CALCULATE ( [Gasto Neto], SAMEPERIODLASTYEAR ( Dim_Calendario[Fecha] ) )
Var % vs AA = DIVIDE ( [Gasto Neto] - [Gasto AA], [Gasto AA] )

Gasto Mes Anterior = CALCULATE ( [Gasto Neto], DATEADD ( Dim_Calendario[Fecha], -1, MONTH ) )
Var % MoM = DIVIDE ( [Gasto Neto] - [Gasto Mes Anterior], [Gasto Mes Anterior] )

Gasto YTD = TOTALYTD ( [Gasto Neto], Dim_Calendario[Fecha] )

Proveedores Activos = CALCULATE ( DISTINCTCOUNT ( Fact_Compras[ProveedorKey] ), Dim_TipoMovimiento[Signo] = 1 )

Nro OC = DISTINCTCOUNT ( Fact_Compras[NroOC] )
Ticket Promedio OC = DIVIDE ( [Gasto Neto], [Nro OC] )

Monto Devuelto = CALCULATE ( SUMX ( Fact_Compras, Fact_Compras[MontoPEN] ), Dim_TipoMovimiento[EsDevolucion] = TRUE () )
% Devoluciones = DIVIDE ( [Monto Devuelto], [Gasto Neto] + [Monto Devuelto] )

-- Pareto: % acumulado del proveedor en el ranking
% Acumulado Pareto =
VAR GastoProv = [Gasto Neto]
VAR Total = CALCULATE ( [Gasto Neto], ALLSELECTED ( Dim_Proveedor ) )
VAR Acum =
    SUMX (
        FILTER (
            ADDCOLUMNS ( ALLSELECTED ( Dim_Proveedor[Proveedor] ), "@g", [Gasto Neto] ),
            [@g] >= GastoProv
        ),
        [@g]
    )
RETURN DIVIDE ( Acum, Total )

Proveedores 80% =
COUNTROWS (
    FILTER ( ALLSELECTED ( Dim_Proveedor[Proveedor] ), [% Acumulado Pareto] <= 0.8 )
) + 1
```

---

## 1. Páginas del reporte y la decisión que responde cada una

| # | Página | Pregunta de negocio | Quién la usa | Decisión que habilita |
|---|---|---|---|---|
| 1 | **Resumen ejecutivo** | ¿Cuánto gastamos, cómo vamos frente al año pasado y dónde debo mirar? | Gerencia, jefe de compras | Aprobar o frenar gasto por unidad de negocio y fijar prioridades del mes. |
| 2 | **Análisis de proveedores** | ¿De quién dependemos y con quién conviene negociar o consolidar? | Jefe de compras, compradores | Negociar con proveedores Pareto A, buscar alternativas donde hay dependencia y dar de baja la cola larga. |
| 3 | **Tendencia y estacionalidad** | ¿El gasto crece, cae o es estacional? ¿Qué mes se salió de lo normal? | Compras, finanzas y presupuesto | Planificar compras anticipadas, ajustar el presupuesto e investigar picos anómalos. |
| 4 | **Detalle operativo / OCs** | ¿Qué transacciones explican las cifras y dónde hay fricción operativa? | Analista, auxiliar de compras, auditoría | Corregir registros, hacer seguimiento a devoluciones y OCs lentas, y detectar compras fraccionadas. |

Navegación: botones en la barra lateral y **drill-through** desde las páginas 1–3 hacia la página 4 (clic derecho sobre un proveedor → "Ver detalle").

---

## 2. KPIs por página (máximo 6–7) + 3. Tipo de gráfico y por qué

### Página 1: Resumen ejecutivo
| KPI | Visual | Por qué |
|---|---|---|
| Gasto neto YTD (S/) | **Tarjeta** (nueva card visual) con referencia al AA | Es el número principal y se lee en un segundo. |
| Var % vs mismo periodo AA | Dentro de la misma tarjeta, flecha y color semántico | El contexto va pegado al número. |
| Proveedores activos | Tarjeta | Muestra la amplitud de la base de proveedores. |
| % gasto en Top 10 proveedores | Tarjeta | Resume la concentración en un solo dato. |
| Ticket promedio por OC | Tarjeta | Indica la eficiencia del proceso (muchas OCs pequeñas generan costo administrativo). |
| Gasto mensual actual vs AA | **Gráfico de líneas** (dos series: actual en color primario, AA en gris) | La tendencia y la comparación se ven en un solo visual. |
| Gasto por unidad de negocio / centro de costo | **Barras horizontales** ordenadas | Las etiquetas largas se leen bien y el ranking queda claro. |

Evita en esta página las tablas y más de 6–7 visuales. El ejecutivo lee los números y luego hace drill.

### Página 2: Análisis de proveedores
| KPI | Visual | Por qué |
|---|---|---|
| Gasto por proveedor (Top 15) | **Barras horizontales** ordenadas de mayor a menor | Es la mejor forma de comparar magnitudes entre categorías con nombres largos. |
| Curva Pareto (% acumulado) | **Combo: columnas + línea** en eje secundario, con línea de referencia al 80% | Es el estándar de la industria para una clasificación ABC. |
| Nº de proveedores que concentran el 80% | Tarjeta | Traduce el Pareto a una cifra accionable ("12 de 340 proveedores = 80%"). |
| Participación del proveedor #1 (riesgo de dependencia) | Tarjeta con umbral (ej. >25% en naranja) | Señala el riesgo de abastecimiento. |
| Frecuencia vs monto por proveedor | **Scatter** (X = Nº de compras, Y = gasto, tamaño = nº de centros de costo) | Segmenta en estratégicos, transaccionales y cola larga. Es una lectura de analista, no de operativo. |
| Proveedores nuevos vs recurrentes | **Columnas apiladas** por año | Muestra si la base se renueva o se fragmenta. |
| Precio unitario promedio (Monto / Cantidad) | **Tabla con barras de datos** (tabla de apoyo) | Permite ver los valores exactos para negociar. |

### Página 3: Tendencia y estacionalidad
| KPI | Visual | Por qué |
|---|---|---|
| Gasto mensual 2023–2026 | **Línea** con promedio móvil de 3 meses | El promedio móvil suaviza el ruido y deja ver la tendencia real. |
| Var % MoM (último mes cerrado) | Tarjeta + **columnas** de variación con color condicional (+ / −) | Distingue de inmediato los meses que suben de los que bajan. |
| Var % YoY | Tarjeta | Aísla el efecto estacional. |
| Gasto acumulado YTD vs AA | **Línea acumulada** (dos series) | Muestra si el año va por encima o por debajo del anterior. |
| Estacionalidad (mes × año) | **Matriz tipo mapa de calor** (formato condicional de fondo, escala de un solo color) | Deja ver los patrones: cierres de año, paradas de planta, campañas. |
| Tendencia por unidad de negocio | **Small multiples** de línea (3 paneles) | Compara Manufactura, Distribución y Servicios sin mezclarlas en un gráfico ilegible. |

Usa una **cascada (waterfall)** si quieres explicar la variación anual por unidad de negocio o por proveedor ("¿qué explica los +S/ 1.2 M?"). Es muy efectiva en entrevistas.

### Página 4: Detalle operativo / órdenes de compra
| KPI | Visual | Por qué |
|---|---|---|
| Nº de OCs / movimientos | Tarjeta | Da el volumen operativo. |
| Tiempo de atención promedio (días solicitud → recepción) | Tarjeta + **histograma** (columnas por rango de días: 0–7, 8–15, 16–30, >30) | El promedio esconde los extremos y la distribución los muestra. |
| % pedidos atendidos dentro del plazo objetivo | Tarjeta con meta | Es el KPI de nivel de servicio del área. |
| Monto y % de devoluciones | Tarjeta + barras por proveedor | Da una señal de calidad del proveedor. |
| OCs bajo el umbral (ej. < S/ 500) | Tarjeta | Sirve para detectar fraccionamiento y oportunidades de consolidar. |
| Detalle transaccional | **Tabla** (no matriz) con formato condicional, ordenable y exportable | Es la única página donde una tabla grande tiene sentido. |

> Si el ERP no trae dos fechas del ciclo, reemplaza el tiempo de atención por **"días entre compras del mismo proveedor"** y documenta el KPI de lead time como mejora futura. Es mejor decirlo con honestidad que inventar el dato.

**Slicers globales** (sincronizados en todas las páginas): Año/Mes, Unidad de negocio, Centro de costo y Proveedor. Pon como máximo 3–4 visibles; el resto va en un panel desplegable o en el panel de filtros.

---

## 4. Diseño visual

### 4.1 Paleta (corporativa, sobria y accesible; la misma del tema JSON de la maqueta)
| Rol | Color | Uso |
|---|---|---|
| Primario | `#2F5478` (azul acero) | Serie principal y títulos de KPI. |
| Secundario | `#7FA3C4` (azul claro) | Segunda serie, clase B del Pareto. |
| Comparativo | `#B9BDC2` (gris) | Año anterior, benchmarks y la cola larga del Pareto. |
| Acento / alerta | `#D08A1E` (naranja) | **Solo** para lo que debe llamar la atención (proveedor sobre el umbral, mes anómalo). |
| Sube / baja | `#B8473D` / `#2F8062` | Solo en variaciones. En compras **más gasto no es "bueno"**: ▲ rojo = el gasto sube, ▼ verde = el gasto baja, igual en todo el reporte. |
| Unidades de negocio | `#3F6C97` Manufactura · `#3A8583` Distribución · `#7C6FA3` Servicios | Siempre el mismo color por unidad. |
| Fondo lienzo / tarjetas | `#EEF0F2` / `#FFFFFF` con borde `#DCDFE3` | Contenedores de visuales. |
| Texto | `#1D1F20` principal, `#5B6168` secundario | |

- Asigna un **color fijo por unidad de negocio** y mantenlo en todas las páginas (Manufactura, Distribución, Servicios).
- Guarda todo en un **archivo de tema JSON** (Vista → Temas → Personalizar → Guardar). Así demuestras estandarización.
- Para daltonismo, evita que rojo y verde sean la única diferencia. Agrega flechas ▲▼ o signos.

### 4.2 Jerarquía y layout (lienzo 1280×720, 16:9)
```
┌─────────────────────────────────────────────────────────────┐
│ Título de página · subtítulo "S/ sin IGV · Actualizado dd/mm"│ Slicers │
├─────────┬─────────┬─────────┬─────────┬─────────┬───────────┤
│  KPI 1  │  KPI 2  │  KPI 3  │  KPI 4  │  KPI 5  │           │  ← fila de tarjetas
├─────────┴─────────┴─────────┴──────────┬────────────────────┤
│  Visual principal (tendencia / Pareto) │ Visual secundario  │  ← 60 / 40
│                                        │ (ranking)          │
├────────────────────────────────────────┴────────────────────┤
│  Visual de apoyo / insight en texto                         │
└─────────────────────────────────────────────────────────────┘
```
- La lectura sigue un patrón **Z / F**: arriba a la izquierda va lo más importante y abajo el detalle.
- Usa una cuadrícula de 8 px, márgenes iguales (16–24 px) y visuales alineados. Revísalo con Formato → Alinear / Distribuir.
- Tipografía: Segoe UI (o DIN). Título de página 20–24, valor de KPI 28–32 en negrita, títulos de visual 12–14, ejes 9–10.
- **Títulos que dicen la conclusión**. Por ejemplo, "12 proveedores concentran el 80% del gasto" en lugar de "Gasto por proveedor". Se pueden hacer dinámicos con una medida de texto.
- Formato de números: `S/ 1.25 M`, `S/ 845 K`, porcentajes con 1 decimal. Nunca muestres `1254789.3456`.
- Agrega **tooltips personalizados** (páginas de tooltip) para dar detalle sin ocupar lienzo.

### 4.3 Qué evitar
- Más de **7–8 visuales por página** o más de 4 slicers visibles.
- **Gráficos circulares o de dona** con más de 3 categorías. Si los usas, que sea solo para mostrar la participación por unidad de negocio.
- **Velocímetros (gauges)**, efectos 3D, sombras fuertes, degradados e imágenes de fondo.
- Etiquetas de datos en todos los puntos: deja solo las del último valor o de los extremos.
- Líneas de cuadrícula marcadas, bordes gruesos y leyendas cuando puedes rotular directo en la serie.
- Más de 5–6 colores en un mismo gráfico. Agrupa la cola larga en "Otros".
- Tablas enormes en el resumen ejecutivo. Esas van en la página 4.
- Mezclar montos con IGV y sin IGV, o mostrar meses incompletos comparados contra meses completos.
- Usar el color naranja de alerta en elementos decorativos: pierde su función.

---

## 5. Para que funcione como portafolio
1. **Anonimiza**: cambia los nombres de proveedor por "Proveedor 001…" y escala los montos con un factor. Nunca publiques datos reales de la empresa.
2. Agrega una página inicial o un README con el problema, los datos, el modelo (captura de la vista de modelo) y **3 hallazgos con recomendación**. Por ejemplo: "consolidar 40 proveedores de cola larga en 5 → ahorro estimado X%".
3. Publica en *Publish to web* (con los datos anonimizados) o sube un video de 2 minutos recorriendo el reporte y enlázalo en LinkedIn.
4. En la entrevista, explica **decisiones**: por qué modelo estrella, por qué excluir IGV, cómo tratas las devoluciones y qué KPI no calculaste por limitación del dato.

---

## 6. Maqueta interactiva
`../maqueta/Dashboard_Compras.dc.html` es la maqueta de las 4 páginas en un lienzo de 1280×720, con datos ficticios, filtros sincronizados, drill-through a la página 4 y botón para descargar el tema JSON. Ábrela junto a `support.js`.
