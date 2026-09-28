# Dashboard de Compras · Power BI

Proyecto de portafolio: dashboard de compras para una empresa industrial peruana con tres unidades de negocio (Manufactura, Distribución y Servicios). Los datos vienen de una exportación de ERP con proveedor, fecha, centro de costo, tipo de movimiento, cantidad y monto en soles, periodo 2023–2026.

> Proveedores anonimizados y montos escalados. La maqueta usa datos ficticios.

## Contenido

| Ruta | Qué es |
|---|---|
| [`docs/DISENO_DASHBOARD.md`](docs/DISENO_DASHBOARD.md) | Diseño completo: validación de datos, modelo estrella, medidas DAX, páginas, KPIs, tipo de gráfico y guía visual. |
| [`maqueta/Dashboard_Compras.dc.html`](maqueta/Dashboard_Compras.dc.html) | Maqueta interactiva de las 4 páginas (lienzo 1280×720). Ábrela en el navegador junto a `support.js`; necesita conexión a internet. |

## Páginas del reporte

| Página | Pregunta de negocio | Decisión que habilita |
|---|---|---|
| 1. Resumen ejecutivo | ¿Cuánto gastamos y cómo vamos frente al año pasado? | Priorizar o frenar gasto por unidad de negocio |
| 2. Proveedores | ¿De quién dependemos? | Negociar con los proveedores clase A, reducir dependencia y consolidar la cola larga |
| 3. Tendencia | ¿El gasto crece o es estacional? ¿Qué mes fue anómalo? | Planificar compras anticipadas y ajustar el presupuesto |
| 4. Detalle de OCs | ¿Qué transacciones explican las cifras? | Seguir devoluciones y pedidos lentos, detectar compras fraccionadas |

## Decisiones de modelado

- **Modelo estrella:** `Fact_Compras` con las dimensiones Calendario, Proveedor, Centro de costo (con unidad de negocio) y Tipo de movimiento.
- **Montos sin IGV**, indicado en el subtítulo de cada página.
- **Devoluciones y anulaciones** se tratan con una columna `Signo` (+1 / −1 / 0) para no inflar el gasto.
- **2026 es parcial:** se compara acumulado contra el mismo periodo del año anterior, y la variación mensual usa el último mes cerrado.
- **Tiempo de atención** requiere dos fechas del ciclo (OC y recepción); si el ERP no las trae, queda como mejora de fase 2.
