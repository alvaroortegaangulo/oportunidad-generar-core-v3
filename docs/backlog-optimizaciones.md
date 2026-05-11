# Backlog de optimizaciones - modulo CORE/IPF

Fecha de analisis: 2026-05-05  
Proyecto: `oportunidad-generar-core-v2`  
Alcance revisado: `oportunidad-generar-core.xaml`, `oportunidad-generar-ipf.xaml`, `oportunidad-generar-core-lib/oportunidad-generar-core-at.xaml`, `oportunidad-generar-core-lib/oportunidad-generar-core-pc.xaml`, `contexto-y-backlog.md` y commits recientes.

## 1. Objetivo

Elevar la mantenibilidad del subproceso UiPath sin cambiar su contrato funcional ni complicar la integracion con el proceso padre. La prioridad es que el modulo pueda abrirse en Studio y entenderse visualmente: reglas de negocio visibles, `Invoke Code` pequenos y comentados, lecturas/escrituras Excel organizadas por mapas, y subworkflows con responsabilidades claras.

Este backlog no sustituye al backlog ya ejecutado en `contexto-y-backlog.md`; lo continua como una capa de optimizacion tecnica.

## 2. Diagnostico actual

- `Utilidad S04 - construir modelo desde DataTables PPO` esta bien formateado y comentado, pero sigue siendo monolitico: 589 lineas y 20.755 caracteres.
- `Utilidad S08 - preparar matriz intercompany y capacidad plantilla` esta en una sola linea: 5.240 caracteres, 78 sentencias aproximadas y sin comentarios internos.
- `Utilidad S08 - aplicar rojo dinamico y recalcular CORE` esta en una sola linea: 2.875 caracteres, 38 sentencias aproximadas y sin comentarios internos.
- `Utilidad S06 - preparar tablas Resources` esta duplicado en PC y AT: 10.198 caracteres por XAML, una sola linea y sin comentarios internos.
- `Utilidad S07 - preparar tablas Cost Planning` esta duplicado en PC y AT: 18.072 caracteres por XAML, una sola linea y sin comentarios internos.
- `oportunidad-generar-ipf.xaml` contiene un `Invoke Code` llamado `Utilidad S12 - validar listas y preparar valores IPF` con codigo vacio. Conviene eliminarlo o sustituirlo por actividades visibles.
- `03 Leer PPO con actividades Excel` usa 29 `Read Cell` y 4 `Read Range`. Funciona, pero es verboso y dificil de mantener ante cambios de celdas PPO.
- `Preparar valores Project Infor desde PPO SAP y Salesforce` contiene 45 asignaciones en un unico `Multiple Assign`. Es correcto funcionalmente, pero concentra demasiada decision OK/KO en una actividad.
- Los subworkflows PC y AT ya ayudan a separar escritura por tipo, pero hoy tambien contienen preparacion comun duplicada.

## 3. Principios de diseno

- Mantener el contrato publico actual de argumentos `in_`/`out_`. No agrupar SAP/Salesforce en JSON ni cambiar la forma de invocacion del proceso padre salvo decision explicita.
- Priorizar actividades UiPath para lectura, escritura, validacion, seleccion de plantilla y control de flujo.
- Permitir `Invoke Code` solo para funciones puras o utilidades tecnicas acotadas: parseo, normalizacion, prorrateo, construccion de `DataTable` o uso COM justificado.
- Ningun `Invoke Code` debe combinar mas de una responsabilidad de negocio.
- Todo `Invoke Code` que permanezca debe tener saltos de linea reales, cabecera `Objetivo / Entradas / Salidas / Regla / No hace` y comentarios breves en bloques no obvios.
- Reducir duplicacion entre PC y AT antes de optimizar microdetalles.
- No anadir dependencias nuevas. Mantener compatibilidad con UiPath Studio 23.10.4 y target `Windows`.
- Preservar la politica de datos funcionales faltantes: no abortar, escribir texto corto en rojo cuando aplique.

## 4. Arquitectura objetivo

Estructura recomendada, sin sobredisenar:

```text
oportunidad-generar-core.xaml
+-- 00 Preparar datos locales para pruebas y ejecucion
+-- 01 Validar entradas del modulo
+-- 02 Seleccionar y copiar plantilla CORE
+-- 03 Leer PPO
+-- 04 Construir modelo comun
+-- 04B Preparar reglas y capacidad de plantilla
+-- 05 Rellenar Project Infor
+-- 06 Delegar Resources y Cost Planning por tipo
+-- 08 Recalcular, guardar y registrar fin

oportunidad-generar-core-lib/
+-- core-common-leer-ppo.xaml                  (opcional si se decide extraer lectura)
+-- core-common-construir-modelo-ppo.xaml      (recomendado)
+-- core-common-preparar-resources.xaml        (recomendado)
+-- core-common-preparar-cost-planning.xaml    (recomendado)
+-- oportunidad-generar-core-pc.xaml           (solo escritura/rangos PC)
+-- oportunidad-generar-core-at.xaml           (solo escritura/rangos AT)
```

La fachada `oportunidad-generar-core.xaml` debe seguir siendo el punto de entrada estable. La extraccion a otros XAML tiene sentido solo cuando reduzca duplicacion real o deje un bloque visual mas claro. No conviene partir cada celda en un workflow distinto.

## 5. Backlog priorizado

### E01 - Estandarizar y controlar `Invoke Code`

**E01-T01 - Inventariar metricas de `Invoke Code` y fijar umbrales**

- Alcance: crear una revision estatica en `test/validar_negativos_core.ps1` o script nuevo que liste `DisplayName`, caracteres, lineas, numero aproximado de sentencias y presencia de comentarios.
- Criterio de aceptacion: el script identifica los bloques de una sola linea y falla si un `Invoke Code` supera el umbral acordado sin estar justificado.
- Prioridad: P0.

**E01-T02 - Formatear y documentar los `Invoke Code` S08 del XAML principal**

- Alcance: `Utilidad S08 - preparar matriz intercompany y capacidad plantilla` y `Utilidad S08 - aplicar rojo dinamico y recalcular CORE`.
- Criterio de aceptacion: ambos bloques tienen saltos de linea, cabecera, comentarios breves, manejo COM legible y no mezclan reglas de negocio con mecanica Excel mas de lo necesario.
- Prioridad: P0.

**E01-T03 - Formatear temporalmente S06/S07 en PC y AT antes de extraerlos**

- Alcance: `oportunidad-generar-core-at.xaml` y `oportunidad-generar-core-pc.xaml`.
- Criterio de aceptacion: si la extraccion comun no se acomete en el mismo sprint, los cuatro bloques S06/S07 quedan al menos legibles y comentados.
- Prioridad: P0.

**E01-T04 - Eliminar el `Invoke Code` vacio de IPF**

- Alcance: `Utilidad S12 - validar listas y preparar valores IPF`.
- Criterio de aceptacion: el XAML IPF no contiene actividades `Invoke Code` vacias ni residuos de implementacion.
- Prioridad: P1.

### E02 - Optimizar lectura PPO y construccion de modelo

**E02-T01 - Sustituir lecturas fijas dispersas por mapa de lectura PPO**

- Alcance: `03 Leer PPO con actividades Excel`.
- Propuesta: crear `dtMapaLecturaPPO` con columnas `Campo`, `Hoja`, `Celda`, `Tipo`, `Obligatorio`, `OrigenFuncional`. Leer mediante bucle o reducir a rangos amplios cuando sea mas robusto.
- Criterio de aceptacion: las 29 lecturas fijas quedan agrupadas en un mapa mantenible o se reducen mediante `Read Range` sobre bloques estables de `Hoja de datos`, `Facturacion y SAP`, `Presupuesto` y `Sintesis Precio`.
- Prioridad: P0.

**E02-T02 - Dividir `Utilidad S04 - construir modelo desde DataTables PPO`**

- Alcance: separar S04 en piezas pequenas.
- Propuesta de corte:
  - normalizacion y parseo comun;
  - cabecera PPO y fechas;
  - perfiles/costes/horas;
  - gastos/compras;
  - tarifas AT y meses.
- Criterio de aceptacion: ningun bloque resultante supera una pantalla razonable de Studio, cada uno declara entradas/salidas y se puede probar de forma aislada por datos de salida.
- Prioridad: P0.

**E02-T03 - Extraer construccion de modelo comun a `core-common-construir-modelo-ppo.xaml`**

- Alcance: mover la transformacion de rangos/celdas PPO a un subworkflow comun si el numero de argumentos de salida sigue siendo manejable.
- Criterio de aceptacion: la fachada mantiene visible el orden funcional y el subworkflow devuelve `dtCabeceraPPO`, `dtPerfilesPPO`, `dtCostesPorAnio`, `dtHorasPorAnio`, `dtGastos`, `dtCompras`, `dtTarifasAT` y `dtMesesPPO`.
- Prioridad: P1.
- Dependencia: E02-T02.

**E02-T04 - Documentar el mapa PPO como tabla tecnica**

- Alcance: `docs/guia-integracion-core.md` o seccion nueva en este backlog vivo.
- Criterio de aceptacion: cada campo CORE usado desde PPO indica hoja, celda/rango, destino de modelo y comportamiento si falta.
- Prioridad: P1.

### E03 - Simplificar `Project Infor`

**E03-T01 - Partir el `Multiple Assign` de Project Infor por bloques de negocio**

- Alcance: `Preparar valores Project Infor desde PPO SAP y Salesforce`.
- Propuesta: dividir en secuencias `05A Fechas e identificadores`, `05B Cliente y practica`, `05C Responsables`, `05D Totales economicos`.
- Criterio de aceptacion: no queda una unica actividad con 45 asignaciones; cada bloque agrupa datos relacionados y mantiene cerca la regla OK/KO.
- Prioridad: P0.

**E03-T02 - Centralizar la regla "valor o texto rojo"**

- Alcance: Project Infor y textos funcionales dinamicos.
- Propuesta: usar una tabla `dtErroresFuncionales` o una utilidad pequena para resolver `Valor`, `HayError`, `TextoError`, `CeldaDestino`.
- Criterio de aceptacion: los textos de error no estan duplicados de forma dispersa y cualquier nuevo dato funcional sigue el mismo patron.
- Prioridad: P1.

**E03-T03 - Evaluar escritura por mapa para celdas fijas de Project Infor**

- Alcance: `Write CellX` de `Project Infor`.
- Propuesta: tabla `dtMapaProjectInfor` con `Campo`, `Celda`, `Valor`, `AplicarRojo`; iterar si Studio permite rango/celda dinamica de forma estable.
- Criterio de aceptacion: si la escritura dinamica es fiable, reducir el bloque de 24 `Write CellX`; si no lo es, documentar la decision y mantener escritura explicita.
- Prioridad: P2.

### E04 - Eliminar duplicacion PC/AT en Resources y Cost Planning

**E04-T01 - Extraer preparacion comun de Resources**

- Alcance: duplicado `Utilidad S06 - preparar tablas Resources` en PC y AT.
- Propuesta: crear `oportunidad-generar-core-lib/core-common-preparar-resources.xaml`.
- Criterio de aceptacion: la normalizacion de compania, intercompany, coste base, tarifa AT y construccion de tablas se mantiene en un unico sitio; PC y AT solo escriben sus rangos especificos.
- Prioridad: P0.

**E04-T02 - Extraer preparacion comun de Cost Planning**

- Alcance: duplicado `Utilidad S07 - preparar tablas Cost Planning` en PC y AT.
- Propuesta: crear `oportunidad-generar-core-lib/core-common-preparar-cost-planning.xaml`.
- Criterio de aceptacion: prorrateos, gastos, compras, riesgos, garantia y textos funcionales se calculan una sola vez; PC y AT solo conservan diferencias de layout.
- Prioridad: P0.

**E04-T03 - Separar reglas reutilizables de intercompany y errores dinamicos**

- Alcance: funciones `NormalizarTexto`, `NormalizarCompania`, `BuscarIntercompanyEnMatriz`, `IntercompanyActual` y catalogo de errores.
- Criterio de aceptacion: la matriz visible sigue existiendo y la evaluacion tecnica se reutiliza sin duplicacion entre Resources y Cost Planning.
- Prioridad: P1.

**E04-T04 - Consolidar escrituras PC/AT por subworkflow**

- Alcance: `oportunidad-generar-core-pc.xaml` y `oportunidad-generar-core-at.xaml`.
- Propuesta: usar un unico `ExcelProcessScopeX` y un unico `ExcelApplicationCard` por subworkflow para escribir `Resources` y `Cost Planning`, si Studio 23.10 lo soporta sin perder claridad.
- Criterio de aceptacion: menos aperturas de Excel, mismo resultado en baseline y diagrama PC/AT mas compacto.
- Prioridad: P1.

### E05 - Uso de Excel, COM y rendimiento

**E05-T01 - Justificar o reemplazar COM en capacidad/recalculo**

- Alcance: S08 principal.
- Propuesta: intentar resolver validacion de hojas/capacidad con actividades Excel y reservar COM solo para `CalculateFullRebuild` o propiedades no disponibles en actividades UiPath.
- Criterio de aceptacion: cada uso COM restante tiene comentario de motivo, bloque `try/finally` legible y liberacion explicita.
- Prioridad: P0.

**E05-T02 - Reducir volumen de actividades Excel repetitivas**

- Alcance: lecturas PPO, escrituras Project Infor y escrituras de rangos PC/AT.
- Criterio de aceptacion: donde existan rangos rectangulares o mapas estables, se usan `Read Range`/`Write Range` en vez de muchas actividades unitarias.
- Prioridad: P1.

**E05-T03 - Medir impacto de rendimiento**

- Alcance: wrapper CORE PC y, si existe caso AT, wrapper AT.
- Criterio de aceptacion: registrar tiempo aproximado antes/despues para lectura PPO, escritura CORE y recalc final. No se exige optimizacion prematura, solo evidencia.
- Prioridad: P2.

### E06 - Calidad, pruebas y seguridad de cambios

**E06-T01 - Actualizar validadores tras cada refactor**

- Alcance: `test/validar_negativos_core.ps1`, `test/validar_integridad_core_pc.ps1`, baselines CORE/IPF.
- Criterio de aceptacion: los validadores cubren que no reaparecen `Invoke Code` monoliticos, que los textos rojos se conservan y que las celdas baseline no cambian sin decision explicita.
- Prioridad: P0.

**E06-T02 - Probar escenarios negativos estructurales y funcionales**

- Alcance: PPO inexistente, plantilla inexistente, tipo PC/AT invalido, hoja obligatoria ausente, exceso de perfiles/gastos/compras, datos SAP/SF incompletos.
- Criterio de aceptacion: los bloqueantes lanzan error claro y los funcionales se escriben en rojo sin abortar.
- Prioridad: P0.

**E06-T03 - Ejecutar Workflow Analyzer focalizado**

- Alcance: XAML principal, IPF, subworkflows PC/AT y wrappers de prueba.
- Criterio de aceptacion: sin errores tecnicos de parseo/carga. Las excepciones aceptadas deben estar documentadas, especialmente si se mantiene el contrato de mas de 20 argumentos.
- Prioridad: P0.

**E06-T04 - Mantener pruebas manuales Studio como cierre de sprint**

- Alcance: `test/test_generar_core_pc.xaml`, `test/test_generar_ipf.xaml`, `test/test_generar_ipf_datos_incompletos.xaml`, `test/ejemplo_generar_core_ipf.xaml`.
- Criterio de aceptacion: al cerrar una epica funcional, se genera Excel desde Studio y se ejecutan los validadores PowerShell sobre los ficheros reales generados.
- Prioridad: P1.

### E07 - Documentacion y handover

**E07-T01 - Actualizar guia de arquitectura del modulo**

- Alcance: `docs/guia-integracion-core.md`.
- Criterio de aceptacion: la guia refleja la arquitectura final, que hace cada XAML, que datos consume y que validaciones quedan para el padre.
- Prioridad: P1.

**E07-T02 - Crear tabla de decisiones tecnicas**

- Alcance: documento breve en `docs/` o seccion en `backlog-optimizaciones.md`.
- Criterio de aceptacion: quedan registradas decisiones como "mantener argumentos individuales", "mantener COM solo para recalculo", "usar subworkflows common", "no empaquetar".
- Prioridad: P2.

## 6. Orden recomendado de ejecucion

1. `O01` - E01-T01, E01-T02, E01-T04, E06-T01: control estatico y limpieza rapida de deuda evidente.
2. `O02` - E02-T01, E02-T02, E02-T04: lectura PPO y S04 menos monolitico.
3. `O03` - E03-T01, E03-T02: Project Infor mas legible sin tocar contrato.
4. `O04` - E04-T01, E04-T02, E04-T03: eliminar duplicacion fuerte PC/AT.
5. `O05` - E05-T01, E04-T04, E05-T02: uso de Excel/COM mas limpio y con menos aperturas.
6. `O06` - E06-T02, E06-T03, E06-T04, E07-T01: cierre de calidad y handover.

## 7. Fuera de alcance por ahora

- Cambiar el contrato publico del modulo CORE.
- Publicar paquete, configurar Orchestrator, assets o colas.
- Crear una libreria externa .NET o dependencia nueva.
- Rehacer la logica funcional ya validada de negocio salvo que una prueba demuestre error.
- Unificar CORE e IPF en un solo XAML. Deben seguir separados y coordinados por wrapper/proceso padre.

## 8. Avance E01 - Control de `Invoke Code`

- Se fija el umbral estatico de `Invoke Code` en `7000` caracteres, `220` lineas o `120` sentencias aproximadas. Cualquier bloque que lo supere debe estar justificado en `test/validar_negativos_core.ps1`.
- El validador falla si detecta un `Invoke Code` vacio, de una sola linea, sin cabecera `Objetivo / Entradas / Salidas / Regla / No hace` o sobredimensionado sin justificacion.
- Los bloques S08 del XAML principal quedan formateados con saltos reales, cabecera y comentarios internos. Tras E05, el uso COM ya no cubre la inspeccion estructural de plantilla; queda acotado al pintado dinamico final y `CalculateFullRebuild`.
- Los bloques S06/S07 de PC y AT quedaron formateados temporalmente con cabecera y comentarios durante E01; la duplicacion se resuelve en E04 mediante common workflows.
- El `Invoke Code` IPF `Utilidad S12 - validar listas y preparar valores IPF` no esta vacio en el estado actual del proyecto; se mantiene porque prepara valores funcionales de IPF, pero queda cubierto por el mismo control estatico de cabecera, metricas y bloque vacio.

## 9. Avance E02 - Lectura PPO y modelo comun

- `03 Leer PPO` deja de usar 29 `Read Cell` dispersos. La fachada lee cinco rangos estables: `Hoja de datos!B8:B29`, `Parametros!B18:J58`, `Presupuesto!B3:X53`, `Facturacion y SAP!A1:E10` y `Sintesis Precio!D12:P80`.
- Se crea `dtMapaLecturaPPO` con `Campo`, `Hoja`, `Celda`, `Tipo`, `Obligatorio` y `OrigenFuncional`. El mapa queda construido en `core-common-construir-modelo-ppo.xaml` y documentado en `docs/guia-integracion-core.md`.
- La construccion S04 se extrae a `oportunidad-generar-core-lib/core-common-construir-modelo-ppo.xaml`, manteniendo la fachada como orden funcional visible y devolviendo `dtCabeceraPPO`, `dtPerfilesPPO`, `dtCostesPorAnio`, `dtHorasPorAnio`, `dtGastos`, `dtCompras`, `dtTarifasAT` y `dtMesesPPO`.
- El antiguo `Invoke Code` monolitico `Utilidad S04 - construir modelo desde DataTables PPO` queda eliminado. La logica se divide en siete utilidades pequenas: mapa PPO, cabecera/fechas, totales fijos, tarifas AT, perfiles/costes/horas, gastos/compras y serializacion JSON.
- `test/validar_negativos_core.ps1` se actualiza para fallar si reaparecen `Read Cell` en CORE o la utilidad S04 monolitica, y ya no mantiene justificacion temporal para S04.
- Validacion estatica E02: XML correcto en XAML principal y subworkflow comun; `test/validar_negativos_core.ps1` OK con todos los nuevos S04 por debajo de `7000` caracteres, `220` lineas y `120` sentencias. UiPath Analyzer carga `core-common-construir-modelo-ppo.xaml` y `test/test_generar_core_pc.xaml` sin errores; la fachada principal mantiene solo deuda heredada de contrato/anidacion/nomenclatura fuera de E02.

## 10. Avance E03 - Simplificar `Project Infor`

- El antiguo `Multiple Assign` `Preparar valores Project Infor desde PPO SAP y Salesforce` queda eliminado. La preparacion se divide en cuatro bloques visibles: `05A Fechas e identificadores`, `05B Cliente y practica`, `05C Responsables` y `05D Totales economicos`.
- Se centraliza la politica "valor o texto rojo" en `dtMapaProjectInfor`, con columnas `Bloque`, `Campo`, `Celda`, `Valor`, `AplicarRojo`, `TextoError` y `OrigenFuncional`. `dtErroresFuncionales` filtra las filas con `AplicarRojo = True` para dejar trazable que celdas funcionales requieren fuente roja.
- La utilidad nueva `Utilidad S05 - construir mapa Project Infor` queda acotada y bajo umbral: 4.205 caracteres, 58 lineas, 37 sentencias y cabecera completa `Objetivo / Entradas / Salidas / Regla / No hace`.
- La escritura fija de Project Infor pasa de 24 `WriteCellX` explicitos y 19 `FormatRangeX` por celda a una iteracion sobre `dtMapaProjectInfor`: un `WriteCellX` dinamico y un `FormatRangeX` condicionado por `AplicarRojo`.
- Decision E03-T03: se acepta escritura por mapa para Project Infor porque las direcciones son celdas fijas de plantilla y Studio 23.10 soporta `ExcelCore.Sheet("Project Infor").Cell(filaProjectInfor("Celda").ToString())` de forma estable en el XAML. Si en una prueba Studio apareciera incompatibilidad runtime, la alternativa documentada es mantener el mapa y volver solo la escritura a actividades explicitas.
- `docs/guia-integracion-core.md` documenta el mapa Project Infor y `test/validar_negativos_core.ps1` falla si reaparece el `Multiple Assign` monolitico o si vuelven los `WriteCellX` fijos repetidos.
- Validacion estatica E03: XML correcto en XAML principal; `test/validar_negativos_core.ps1` OK con 16 bloqueantes, 19 casos Project Infor y 8 funcionales dinamicos. `analyze-file` del wrapper PC OK con solo `ST-ANA-009`; `analyze-file` de la fachada carga correctamente y mantiene incidencias aceptadas por contrato/diseno (`ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016`, `ST-NMG-009`, `ST-MRD-009` y `ST-ANA-009`), sin errores tecnicos de parseo/carga. Queda pendiente la ejecucion manual desde Studio del wrapper PC para confirmar runtime sobre Excel real, igual que en cierres funcionales anteriores.

## 11. Avance E04 - Eliminar duplicacion PC/AT en Resources y Cost Planning

- Se crean `oportunidad-generar-core-lib/core-common-preparar-resources.xaml` y `oportunidad-generar-core-lib/core-common-preparar-cost-planning.xaml`.
- `oportunidad-generar-core-pc.xaml` y `oportunidad-generar-core-at.xaml` dejan de contener `Invoke Code`: invocan los common workflows y conservan solo las actividades Excel de limpieza/escritura de rangos especificos de cada layout.
- La preparacion de `Resources` queda en un unico sitio: normalizacion de compania, matriz intercompany, codigos de empleado, coste base, tarifa AT y construccion de tablas PC/AT.
- La preparacion de `Cost Planning` queda en un unico sitio: prorrateos por anualidad, riesgos, garantia, gastos, compras y textos funcionales de proveedor/PO/Ariba.
- La regla tecnica de intercompany se reutiliza desde `Resources`: `Cost Planning` consume `dt_ResourcesATFijos` para reproducir codigo, sigla, intercompany y perfil en la rama AT, evitando duplicar `NormalizarTexto`, `NormalizarCompania`, `BuscarIntercompanyEnMatriz` e `IntercompanyActual`.
- `test/validar_negativos_core.ps1` se actualiza para fallar si PC/AT recuperan `Invoke Code`, si faltan los common workflows o si la logica tecnica de intercompany aparece mas de una vez. Los dos bloques comunes S06/S07 quedan justificados como deuda controlada no duplicada por superar el umbral estatico, pendiente de una posible division interna posterior.
- Validacion estatica E04: XML correcto en la fachada, PC, AT y los dos common workflows; `test/validar_negativos_core.ps1` OK con 6 XAML CORE revisados. `analyze-file` con rutas absolutas carga correctamente la fachada, PC, AT y los common workflows; el wrapper `test/test_generar_core_pc.xaml` analiza OK con solo `ST-ANA-009`. Quedan incidencias aceptadas por contrato/diseno (`ST-DBP-002`, `ST-NMG-002`, `ST-NMG-009`, `ST-NMG-011`, `ST-NMG-016`, `ST-MRD-009`, `ST-USG-020` y `ST-ANA-009`), sin errores tecnicos de parseo/carga. Queda pendiente ejecutar wrapper PC sobre Excel real para cierre runtime.

## 12. Avance E05 - Uso de Excel, COM y rendimiento

- La inspeccion estructural de la plantilla CORE deja de usar COM. La secuencia `04B Preparar reglas S08 de capacidad e intercompany` valida hojas obligatorias con `Read Range` y lee `Resources!A1:AF120` y `Cost Planning!A1:AF140`; la utilidad S08 calcula capacidad desde esos `DataTable`.
- El COM restante queda justificado en `Utilidad S08 - aplicar rojo dinamico y recalcular CORE`: se usa para recorrer textos funcionales escritos por `Write Range`, aplicar rojo dinamico y ejecutar `CalculateFullRebuild`/marcas de recalculo no expuestas por actividades UiPath 23.10. El bloque conserva `try/finally`, `Close`, `Quit`, `ReleaseComObject` y GC final.
- `oportunidad-generar-core-pc.xaml` y `oportunidad-generar-core-at.xaml` consolidan las escrituras de `Resources` y `Cost Planning`: preparan primero ambas tablas comunes y abren el CORE una sola vez por subworkflow mediante un unico `ExcelProcessScopeX`/`ExcelApplicationCard`, escribiendo todo por `Write Range` y guardando al final.
- Se anaden logs de rendimiento aproximado: lectura PPO (`03 Leer PPO por rangos estables`), escritura CORE (`05 Project Infor` + subworkflow PC/AT) y recalculo final (`08 Recalcular, guardar y registrar fin`).
- `test/validar_negativos_core.ps1` protege E05: falla si la capacidad vuelve a usar COM, si PC/AT tienen mas de una apertura Excel, o si desaparecen los logs de rendimiento. Validacion estatica: XML correcto en fachada/PC/AT y `test/validar_negativos_core.ps1` OK. Analyzer secuencial carga fachada, PC, AT y wrapper PC sin errores tecnicos; quedan solo incidencias aceptadas por contrato/diseno (`ST-DBP-002`, `ST-NMG-002`, `ST-NMG-009`, `ST-NMG-011`, `ST-NMG-016`, `ST-MRD-009` y `ST-ANA-009`), con wrapper PC en `exit=0`.

## 13. Avance E06 - Calidad, pruebas y seguridad de cambios

- `test/validar_negativos_core.ps1` queda reforzado para proteger E06: ademas de `Invoke Code`, mapas y COM, valida que existe el wrapper CORE de datos SAP/SF incompletos y que `test/baseline_core_pc.csv` conserva las celdas criticas de Project Infor, Resources y Cost Planning.
- Se crea `test/test_generar_core_datos_incompletos.xaml` para probar desde Studio el caso funcional negativo: PPO y plantilla validos, datos SAP/Salesforce incompletos, generacion sin aborto y textos cortos en rojo.
- `test/validar_integridad_core_pc.ps1` incorpora el parametro `-Scenario DatosIncompletos` y comprueba celdas rojas clave: WBS, cliente SAP, Sales Order, responsables, codigo de recurso, intercompany y datos de compras.
- Se crea `test/validar_calidad_e06.ps1` como cierre repetible: XML de XAML focalizados, negativos CORE/IPF, referencias IPF, Workflow Analyzer secuencial y validacion opcional de Excel generados. El parametro `-RequireGeneratedWorkbooks` fuerza que los ficheros reales existan en un cierre manual de sprint.
- Workflow Analyzer focalizado queda documentado con reglas aceptadas por contrato/diseno: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-009`, `ST-NMG-011`, `ST-NMG-016`, `ST-MRD-009`, `ST-USG-020` y `ST-ANA-009`. Cualquier regla nueva no aceptada o error de carga/parseo falla el cierre E06.
- Queda mantenida la naturaleza manual de E06-T04: la generacion runtime de CORE/IPF se hace desde UiPath Studio 23.10.4 y despues se ejecutan los validadores PowerShell sobre los Excel reales.

## 14. Avance E07 - Documentacion y handover

- `docs/guia-integracion-core.md` queda actualizada como guia de arquitectura final: fachada estable, workflows common, subworkflows PC/AT, IPF separado, orden visual del XAML principal, datos consumidos, validaciones internas y validaciones que quedan para el proceso padre.
- Se crea `docs/decisiones-tecnicas-core.md` como tabla de decisiones tecnicas cerradas. Registra, entre otras, mantener argumentos individuales, no empaquetar, no anadir dependencias, lectura PPO por mapa, common workflows, COM reservado al cierre de Excel, no abortar por datos funcionales faltantes y reglas aceptadas de Workflow Analyzer.
- `test/README.md` enlaza la guia de handover, la tabla de decisiones y los wrappers que debe revisar el integrador: CORE minimo, CORE con datos incompletos, IPF completo, IPF datos incompletos y CORE+IPF condicional.
- E07-T01 queda completada: la guia refleja que hace cada XAML, que datos consume, que produce y que validaciones quedan fuera del submodulo.
- E07-T02 queda completada: las decisiones tecnicas quedan registradas en `docs/decisiones-tecnicas-core.md` con motivo e impacto de mantenimiento.

## 15. Cierre del backlog de optimizaciones

Estado final: backlog de optimizaciones E01-E07 cerrado en modo handover tecnico.

Validacion de cierre ejecutada el 2026-05-05:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1 -SkipGeneratedWorkbooks
```

Resultado:

- XML OK en fachada CORE, IPF, subworkflows PC/AT, common workflows y wrappers focalizados.
- `test/validar_negativos_core.ps1` OK: contrato negativo CORE, mapas, `Invoke Code`, Project Infor, wrappers y baseline protegidos.
- `test/validar_negativos_ipf.ps1` OK: bloqueantes IPF, funcionales y wrappers S13 protegidos.
- `test/validar_integridad_ipf.ps1 -SkipGenerated` OK: plantilla IPF y ejemplo GAP validados como referencias.
- Workflow Analyzer focalizado OK sin errores tecnicos de carga/parseo; solo reglas aceptadas por contrato/diseno.
- Validaciones de Excel generado omitidas en este cierre automatico por `-SkipGeneratedWorkbooks`, ya que no habia workbooks runtime en `.local\test-output`.

Pendiente operativo no bloqueante del cierre documental: para una entrega funcional con evidencia Excel, ejecutar desde UiPath Studio 23.10.4 los wrappers `test/test_generar_core_pc.xaml`, `test/test_generar_core_datos_incompletos.xaml`, `test/test_generar_ipf.xaml`, `test/test_generar_ipf_datos_incompletos.xaml` y, si aplica integracion completa, `test/ejemplo_generar_core_ipf.xaml`; despues repetir `test/validar_calidad_e06.ps1 -RequireGeneratedWorkbooks`.
