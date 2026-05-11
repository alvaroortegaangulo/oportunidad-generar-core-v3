# Backlog de optimizacion de rendimiento CORE

Fecha de creacion: 2026-05-12  
Proyecto: `oportunidad-generar-core-main`  
Alcance: optimizacion iterativa de `Main.xaml` y workflows `lib` sin cambiar el resultado funcional de los ficheros CORE generados.

Este documento es el backlog operativo para que Codex implemente mejoras de rendimiento en conversaciones futuras, por epicas y tareas pequenas. Cada tarea debe poder ejecutarse de forma independiente, con validacion objetiva antes de pasar a la siguiente.

## Objetivo

Reducir el tiempo de ejecucion de `Main.xaml` y de la fachada `lib/oportunidad-generar-core.xaml` manteniendo equivalencia funcional estricta en los CORE generados.

La optimizacion debe atacar principalmente:

- exceso de aperturas, guardados y recalculos de Excel;
- actividades UiPath repetidas celda a celda;
- lecturas amplias de rangos cuando exista una alternativa equivalente;
- bloques que saturan el profile execution antes de terminar PC y AT.

## Regla principal de no cambio funcional

Cualquier cambio debe preservar el resultado observable de los CORE:

- mismos valores visibles en las hojas funcionales;
- mismas formulas relevantes y ausencia de `#REF!`;
- mismos formatos visibles relevantes, especialmente fuente roja en errores funcionales;
- mismas celdas marcadas en rojo dentro de los rangos actuales de revision;
- misma trazabilidad funcional en `Cost Summary`;
- mismas validaciones existentes;
- mismos argumentos publicos de entrada/salida;
- mismas rutas, nombres de ficheros y estructura de hojas.

No se acepta una optimizacion que cambie reglas de negocio, celdas destino, rangos funcionales, textos funcionales, criterios de marcado en rojo o comportamiento PC/AT.

La unica diferencia tolerable entre baseline y salida optimizada es un valor inherentemente variable, como timestamp de generacion, siempre que el comparador lo trate de forma explicita y documentada.

## Baseline tecnica actual

Profile analizado: `0393b8ac-5727-412a-bf8f-677d5564ab5e.uistat`.

Resumen del profile:

- `TotalWallMs`: `157242` ms.
- `TotalEvents`: `40000`.
- `ActivityExecutionLimitReached`: `True`.
- El profile se corta antes de terminar la primera generacion `CORE PC`.
- La ejecucion `CORE AT` no queda medida en el profile capturado porque el limite se alcanza antes.

Cuellos de botella principales detectados:

| Workflow | Bloque / actividad | Evidencia del profile | Riesgo de rendimiento |
| --- | --- | ---: | --- |
| `lib/oportunidad-generar-core.xaml` | `03 Leer PPO por rangos estables` | `58995 ms` | Apertura/lectura Excel costosa, posible fallback y actividad moderna sobre `.xlsm`. |
| `lib/oportunidad-generar-core.xaml` | `Leer rango para errores CORE` | `47463 ms` en 2 lecturas, maximo `44440 ms` | Lecturas muy amplias durante cierre final. |
| `lib/oportunidad-generar-core.xaml` | Revision celda a celda de errores CORE | `8463` eventos ya registrados antes de cortar profile | Explosion de actividades UiPath. |
| `lib/oportunidad-generar-core.xaml` | `Abrir Excel para Project Infor` | `13870 ms` | Apertura separada y 26 escrituras individuales. |
| `lib/oportunidad-generar-core.xaml` | `Leer estructura de plantilla CORE` | `10976 ms` | Lecturas amplias de plantilla. |
| `lib/oportunidad-generar-core-pc.xaml` | `Abrir Excel para escribir CORE PC` | `10137 ms` | Escritura con multiples `ClearRangeX` y `WriteRangeX`. |
| `lib/oportunidad-generar-core.xaml` | `Recalcular y guardar CORE con Excel COM` | No medido completo por corte del profile | Riesgo alto: reabre Excel y fuerza `CalculateFullRebuild`. |

Rangos actuales de revision de errores funcionales:

- `Project Infor!A1:J80`
- `Resources!A1:CA120`
- `Cost Planning!A1:CB140`

Estos rangos forman parte del contrato funcional actual. No deben reducirse salvo que negocio apruebe un cambio funcional, lo cual queda fuera de este backlog.

## Reglas para Codex en futuras conversaciones

- Implementar una sola tarea por conversacion salvo que el usuario pida explicitamente agrupar varias.
- Antes de modificar rendimiento, comprobar si existe baseline comparable de outputs.
- No tocar workflows no listados en la tarea salvo necesidad tecnica justificada.
- No cambiar contratos publicos de `Main.xaml` ni `lib/oportunidad-generar-core.xaml`.
- No cambiar datos de prueba SAP/Salesforce del `Main.xaml`.
- No cambiar plantillas CORE salvo que una tarea lo indique explicitamente; este backlog no lo requiere.
- Cualquier `Invoke Code` nuevo debe incluir cabecera de comentario con objetivo, entradas, salidas, regla y que no hace.
- Cualquier cambio debe cerrar con evidencia: comandos ejecutados, resultado, ficheros tocados y diferencias de rendimiento si aplica.

## Comandos base de validacion

Ejecutar desde `oportunidad-generar-core-main` salvo que se indique lo contrario.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -WorkbookPath .\data\output\CORE_AT_20251160543.xlsx -Kind AT -ExpectedStart '2026-01-01' -Months 48
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -WorkbookPath .\data\output\CORE_PC_20250256445.xlsx -Kind PC -ExpectedStart '2025-11-01' -Months 12
```

Si se empaqueta y ejecuta con UiRobot, usar la version de paquete que corresponda a la rama de trabajo y registrar tiempo total, profile generado y si `ActivityExecutionLimitReached` queda en `False`.

## Epica 0 - Baseline y comparador de equivalencia

Objetivo: crear una red de seguridad automatica antes de tocar rendimiento. Ninguna optimizacion relevante debe implementarse antes de que exista un comparador pre/post razonable.

### EP0-T01 - Definir contrato de equivalencia funcional

Contexto: el proyecto ya tiene validadores de calidad, pero no un comparador generico entre dos CORE. Antes de optimizar hay que documentar que se considera diferencia funcional.

Archivos objetivo:

- `docs/backlog-optimizacion-rendimiento-core.md`
- Opcional: nuevo documento corto en `docs/` si se quiere separar la especificacion del comparador.

Cambios permitidos:

- Precisar hojas, rangos y tolerancias.
- Definir exclusiones explicitas, como timestamp.
- Definir formato de reporte de diferencias.

Cambios prohibidos:

- Cambiar validadores existentes.
- Cambiar workflows.
- Relajar reglas de negocio para facilitar optimizaciones.

Criterios de aceptacion:

- Queda descrito que valores visibles, formulas, errores `#REF!` y fuente roja son parte del contrato.
- Queda descrito que cualquier exclusion debe estar codificada y justificada.

Comandos de validacion:

```powershell
git diff -- docs
```

Evidencia esperada:

- Resumen de reglas de equivalencia anadidas o confirmadas.
- Lista de exclusiones permitidas.

### EP0-T02 - Crear script comparador de outputs CORE

Contexto: se necesita un script que compare un workbook baseline contra un workbook candidato sin abrir Excel por interfaz. Debe servir para PC y AT.

Archivos objetivo:

- `test/comparar_outputs_core_equivalencia.ps1`
- Opcional: `docs/guia-integracion-core.md` solo si se anade una seccion de uso.

Cambios permitidos:

- Crear script PowerShell basado en OpenXML/ZIP XML y, si es necesario, COM solo para valores visibles cuando no haya alternativa fiable.
- Parametrizar rutas `-BaselinePath`, `-CandidatePath`, `-Kind`.
- Comparar hojas clave: `Project Infor`, `Resources`, `Cost Planning`, `Monthly View`, `Cost Summary`.
- Comparar errores `#REF!`.
- Comparar formulas presentes en celdas relevantes.
- Comparar celdas con fuente roja en los rangos de revision actuales.
- Ignorar diferencias de timestamp solo en celdas documentadas.

Cambios prohibidos:

- No modificar outputs durante la comparacion.
- No corregir automaticamente diferencias.
- No depender de rutas absolutas del entorno local.

Criterios de aceptacion:

- El script pasa comparando un fichero contra una copia identica.
- El script falla si se altera artificialmente un valor visible relevante.
- El script falla si se elimina una formula relevante.
- El script falla si se cambia una celda roja esperada en los rangos de revision.
- El script reporta diferencias con hoja, celda y tipo de diferencia.

Comandos de validacion:

```powershell
Copy-Item .\data\output\CORE_PC_20250256445.xlsx .\data\output\CORE_PC_20250256445_baseline_tmp.xlsx -Force
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\comparar_outputs_core_equivalencia.ps1 -BaselinePath .\data\output\CORE_PC_20250256445_baseline_tmp.xlsx -CandidatePath .\data\output\CORE_PC_20250256445.xlsx -Kind PC
```

```powershell
Copy-Item .\data\output\CORE_AT_20251160543.xlsx .\data\output\CORE_AT_20251160543_baseline_tmp.xlsx -Force
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\comparar_outputs_core_equivalencia.ps1 -BaselinePath .\data\output\CORE_AT_20251160543_baseline_tmp.xlsx -CandidatePath .\data\output\CORE_AT_20251160543.xlsx -Kind AT
```

Evidencia esperada:

- Salida OK contra copias identicas.
- Salida KO controlada ante una diferencia artificial.
- Explicacion de cualquier limitacion tecnica del comparador.

### EP0-T03 - Crear procedimiento de baseline pre/post

Contexto: cada epica de optimizacion necesitara generar outputs baseline, aplicar cambio, regenerar outputs candidatos y comparar.

Archivos objetivo:

- `docs/backlog-optimizacion-rendimiento-core.md`
- Opcional: `docs/guia-integracion-core.md`.

Cambios permitidos:

- Documentar carpeta sugerida para baselines temporales.
- Documentar comandos de copia y comparacion.
- Documentar convencion de nombres: `*_baseline.xlsx`, `*_candidate.xlsx`.

Cambios prohibidos:

- No versionar outputs temporales grandes salvo decision explicita.
- No cambiar `data/output` como fuente funcional de verdad sin indicarlo.

Criterios de aceptacion:

- Un implementador puede seguir los pasos sin decidir convenciones.
- El procedimiento cubre PC y AT.

Comandos de validacion:

```powershell
git diff -- docs
```

Evidencia esperada:

- Pasos documentados para generar y comparar baseline/candidato.

## Epica 1 - Escaneo final de errores en batch

Objetivo: optimizar el bloque mas costoso sin cambiar semantica. El resultado debe marcar exactamente las mismas celdas rojas que ahora.

### EP1-T01 - Aislar y caracterizar el bloque actual de revision de errores

Contexto: el bloque vive en `lib/oportunidad-generar-core.xaml`, dentro de `08 Recalcular, guardar y registrar fin`. Actualmente define textos, define rangos, lee cada rango y recorre celdas con actividades UiPath.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`
- `docs/backlog-optimizacion-rendimiento-core.md` si se documenta hallazgo adicional.

Cambios permitidos:

- Anadir logs temporales solo si son necesarios y se retiran antes del cierre.
- Documentar IDs de actividades y rangos.

Cambios prohibidos:

- No cambiar el comportamiento todavia.
- No reducir rangos.
- No cambiar textos de error.

Criterios de aceptacion:

- Queda identificada la lista exacta de textos de error vigente.
- Quedan identificados los tres rangos exactos.
- Queda claro donde se aplican `ReadRangeX`, loops y `FormatRangeX`.

Comandos de validacion:

```powershell
rg -n "Definir textos funcionales CORE|Definir rangos|Leer rango para errores CORE|Aplicar rojo funcional CORE" .\lib\oportunidad-generar-core.xaml
```

Evidencia esperada:

- Resumen con lineas/actividades localizadas.

### EP1-T02 - Implementar deteccion batch equivalente

Contexto: la optimizacion debe mantener el escaneo completo de los mismos rangos, pero ejecutarlo en bloque para no generar miles de eventos UiPath.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`

Cambios permitidos:

- Sustituir el `ReadRangeX` masivo y loops UiPath por un `Invoke Code` batch.
- Usar COM dentro de la sesion de cierre si es la via mas fiable para obtener valores visibles y aplicar formato.
- Mantener `Trim` y comparacion case-insensitive.
- Mantener los mismos rangos:
  - `Project Infor!A1:J80`
  - `Resources!A1:CA120`
  - `Cost Planning!A1:CB140`
- Mantener la misma lista de textos funcionales.

Cambios prohibidos:

- No limitar el marcado a celdas previamente conocidas.
- No cambiar ningun texto de error.
- No cambiar ninguna direccion de rango.
- No omitir celdas vacias por atajo si eso altera equivalencia.

Criterios de aceptacion:

- Comparador EP0 confirma mismas celdas rojas pre/post.
- Validadores existentes pasan.
- El profile muestra reduccion drastica de eventos en este bloque.
- `ActivityExecutionLimitReached` deja de activarse por el barrido de errores, o al menos se retrasa claramente si aun hay otros cuellos.

Comandos de validacion:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\comparar_outputs_core_equivalencia.ps1 -BaselinePath .\data\baseline\CORE_PC_20250256445.xlsx -CandidatePath .\data\output\CORE_PC_20250256445.xlsx -Kind PC
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\comparar_outputs_core_equivalencia.ps1 -BaselinePath .\data\baseline\CORE_AT_20251160543.xlsx -CandidatePath .\data\output\CORE_AT_20251160543.xlsx -Kind AT
```

Evidencia esperada:

- Diff de XAML.
- Resultado OK de comparador PC/AT.
- Extracto de profile antes/despues para el bloque.

### EP1-T03 - Eliminar actividades UiPath redundantes del barrido

Contexto: tras implementar el batch, deben retirarse actividades antiguas para evitar doble trabajo.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`

Cambios permitidos:

- Eliminar o desactivar el `ReadRangeX_1` antiguo y loops asociados si el batch ya es equivalente.
- Mantener logs utiles de cierre.

Cambios prohibidos:

- No dejar doble marcado si afecta rendimiento o formato.
- No dejar ramas muertas confusas en Studio.

Criterios de aceptacion:

- No quedan loops UiPath sobre filas/columnas de los rangos de revision.
- El XAML sigue abriendo correctamente en Studio.
- Comparador y validadores pasan.

Comandos de validacion:

```powershell
rg -n "Leer rango para errores CORE|Recorrer filas rango CORE|Recorrer columnas rango CORE|Aplicar rojo funcional CORE" .\lib\oportunidad-generar-core.xaml
```

Evidencia esperada:

- Confirmacion de que el recorrido celda a celda ya no existe como actividades UiPath.

## Epica 2 - Escritura batch de Project Infor

Objetivo: reducir 26 `WriteCellX` individuales sin cambiar ninguna celda, valor ni formato funcional.

### EP2-T01 - Mapear `dtMapaProjectInfor` y escrituras actuales

Contexto: `dtMapaProjectInfor` gobierna campo, celda, valor y si aplica rojo. Debe mantenerse como fuente logica aunque cambie la forma de escritura.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`

Cambios permitidos:

- Inspeccion y documentacion de columnas de `dtMapaProjectInfor`.
- Identificar todas las celdas destino y condiciones `AplicarRojo`.

Cambios prohibidos:

- No cambiar mapa ni datos.
- No cambiar escritura todavia.

Criterios de aceptacion:

- Lista de celdas destino actual identificada.
- Confirmacion de que son 26 filas/escrituras o el numero real vigente.

Comandos de validacion:

```powershell
rg -n "dtMapaProjectInfor|Escribir celda Project Infor desde mapa|Aplicar rojo Project Infor" .\lib\oportunidad-generar-core.xaml
```

Evidencia esperada:

- Resumen de mapa y rango de lineas.

### EP2-T02 - Reemplazar `WriteCellX` por escritura batch equivalente

Contexto: la escritura actual abre Excel para `Project Infor` y escribe cada celda individualmente. El objetivo es escribir los mismos valores en las mismas celdas de forma mas eficiente.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`

Cambios permitidos:

- Usar `Invoke Code` para iterar `dtMapaProjectInfor` dentro de una unica llamada.
- Escribir exactamente `filaProjectInfor("Valor")` en `filaProjectInfor("Celda")`.
- Aplicar rojo a las mismas celdas con `AplicarRojo=True`.
- Mantener `dtErroresFuncionales` y logs funcionales.

Cambios prohibidos:

- No cambiar nombres de campos.
- No cambiar celda destino.
- No cambiar texto escrito.
- No agrupar por rango si eso desplaza celdas o introduce blancos no previstos.

Criterios de aceptacion:

- Comparador confirma equivalencia de `Project Infor`.
- Validadores existentes pasan.
- Profile muestra reduccion del tiempo de `Abrir Excel para Project Infor` o de eventos internos.

Comandos de validacion:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\comparar_outputs_core_equivalencia.ps1 -BaselinePath .\data\baseline\CORE_PC_20250256445.xlsx -CandidatePath .\data\output\CORE_PC_20250256445.xlsx -Kind PC
```

Evidencia esperada:

- Reporte comparador OK.
- Tiempo antes/despues del bloque si hay profile disponible.

## Epica 3 - Consolidacion de sesiones Excel

Objetivo: reducir aperturas, guardados y reaperturas del mismo libro sin alterar orden logico ni recalculo final.

### EP3-T01 - Inventariar aperturas y guardados actuales

Contexto: el flujo abre Excel para lectura PPO, lectura de estructura de plantilla, `Project Infor`, escritura PC/AT, cierre final y recalculo COM.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`
- `lib/oportunidad-generar-core-pc.xaml`
- `lib/oportunidad-generar-core-at.xaml`

Cambios permitidos:

- Crear inventario en documentacion o comentario temporal.
- Medir duraciones por profile.

Cambios prohibidos:

- No consolidar todavia.
- No cambiar recalculo.

Criterios de aceptacion:

- Tabla de aperturas/guardados con workflow, actividad, proposito y coste aproximado.

Comandos de validacion:

```powershell
rg -n "ExcelProcessScopeX|ExcelApplicationCard|SaveExcelFileX|Recalcular y guardar CORE" .\lib -g "*.xaml"
```

Evidencia esperada:

- Inventario claro para decidir consolidacion.

### EP3-T02 - Consolidar cierre final y recalculo cuando sea seguro

Contexto: actualmente el cierre final guarda el CORE y luego otro bloque COM puede reabrirlo para recalcular. Si se puede mantener el mismo resultado sin reabrir, debe hacerse.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`

Cambios permitidos:

- Integrar recalculo en la misma sesion COM/batch usada para cierre final.
- Mantener `CalculateFullRebuild` o comportamiento equivalente.
- Mantener guardado final.

Cambios prohibidos:

- No eliminar recalculo final.
- No cambiar formulas.
- No dejar Excel abierto.
- No omitir liberacion COM.

Criterios de aceptacion:

- Comparador PC/AT OK.
- Validadores OK.
- No quedan procesos Excel huerfanos despues de ejecucion.
- Menor tiempo de cierre final.

Comandos de validacion:

```powershell
Get-Process Excel -ErrorAction SilentlyContinue
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
```

Evidencia esperada:

- Resultado de comparador.
- Confirmacion de no procesos Excel residuales.

### EP3-T03 - Evaluar consolidacion entre fachada y subworkflow PC/AT

Contexto: `oportunidad-generar-core-pc.xaml` y `oportunidad-generar-core-at.xaml` abren el CORE para escribir. Consolidar esta apertura con la fachada puede implicar cambios de contrato interno.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`
- `lib/oportunidad-generar-core-pc.xaml`
- `lib/oportunidad-generar-core-at.xaml`

Cambios permitidos:

- Proponer refactor interno si no cambia contrato publico.
- Mantener subworkflows PC/AT separados si la consolidacion aumenta riesgo.

Cambios prohibidos:

- No fusionar PC y AT en un unico workflow monolitico.
- No cambiar argumentos publicos de la fachada.
- No sacrificar revisabilidad en Studio sin justificacion.

Criterios de aceptacion:

- Decision documentada: consolidar, posponer o descartar.
- Si se implementa, comparador y validadores OK.

Comandos de validacion:

```powershell
git diff -- .\lib\oportunidad-generar-core.xaml .\lib\oportunidad-generar-core-pc.xaml .\lib\oportunidad-generar-core-at.xaml
```

Evidencia esperada:

- Decision tecnica con motivos.

## Epica 4 - Escrituras `Resources` y `Cost Planning`

Objetivo: optimizar pares `ClearRangeX + WriteRangeX` preservando exactamente el estado final de las hojas.

### EP4-T01 - Inventariar pares limpieza/escritura PC y AT

Contexto: PC y AT limpian rangos y luego escriben tablas. Solo se puede reemplazar un par si una escritura rectangular con blancos reproduce exactamente el resultado.

Archivos objetivo:

- `lib/oportunidad-generar-core-pc.xaml`
- `lib/oportunidad-generar-core-at.xaml`

Cambios permitidos:

- Crear tabla de pares: hoja, rango limpiado, rango escrito, DataTable fuente, dimensiones.

Cambios prohibidos:

- No modificar aun los rangos.

Criterios de aceptacion:

- Inventario completo de `ClearRangeX` y `WriteRangeX` de PC/AT.
- Identificacion de pares candidatos y no candidatos.

Comandos de validacion:

```powershell
rg -n "ClearRangeX|WriteRangeX" .\lib\oportunidad-generar-core-pc.xaml .\lib\oportunidad-generar-core-at.xaml
```

Evidencia esperada:

- Tabla de candidatos con riesgo bajo/medio/alto.

### EP4-T02 - Preparar DataTables con blancos para sobrescritura segura

Contexto: para reemplazar limpieza mas escritura, la tabla debe cubrir todo el rango que hoy queda limpio/escrito y contener blancos donde antes se limpiaba.

Archivos objetivo:

- `lib/core-common-preparar-resources.xaml`
- `lib/core-common-preparar-cost-planning.xaml`
- `lib/oportunidad-generar-core-pc.xaml`
- `lib/oportunidad-generar-core-at.xaml`

Cambios permitidos:

- Ajustar preparacion de DataTables para que tengan dimensiones exactas de destino cuando la tarea lo requiera.
- Usar blancos explicitos en celdas que deben quedar vacias.

Cambios prohibidos:

- No tocar formulas fuera de rangos actuales.
- No cambiar filas funcionales.
- No cambiar orden de recursos, gastos, compras, riesgos o garantia.

Criterios de aceptacion:

- Comparador confirma equivalencia en `Resources` y `Cost Planning`.
- No quedan residuos de ejecuciones previas si se reejecuta sobre un output existente.

Comandos de validacion:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
```

Evidencia esperada:

- Comparador OK tras reejecucion con outputs preexistentes.

### EP4-T03 - Sustituir pares seguros por escritura unica

Contexto: tras preparar tablas con blancos, reemplazar solo los pares demostrablemente equivalentes.

Archivos objetivo:

- `lib/oportunidad-generar-core-pc.xaml`
- `lib/oportunidad-generar-core-at.xaml`

Cambios permitidos:

- Eliminar `ClearRangeX` de pares seguros.
- Mantener `WriteRangeX` con tabla ampliada.

Cambios prohibidos:

- No eliminar limpiezas cuyo efecto no este cubierto por la escritura.
- No cambiar rango inicial de escritura.

Criterios de aceptacion:

- Menor numero de actividades Excel por PC/AT.
- Comparador y validadores OK.
- No hay residuos antiguos tras reejecucion.

Comandos de validacion:

```powershell
rg -n "ClearRangeX" .\lib\oportunidad-generar-core-pc.xaml .\lib\oportunidad-generar-core-at.xaml
```

Evidencia esperada:

- Lista de limpiezas eliminadas y razon de equivalencia.

## Epica 5 - Lectura de capacidades de plantilla

Objetivo: reducir lecturas amplias de estructura de plantilla sin cambiar la capacidad detectada ni la validacion estructural.

### EP5-T01 - Aislar algoritmo actual de capacidad

Contexto: la fachada mide capacidad de `Resources`, `Cost Planning` y `Monthly View` antes de escribir. Esta validacion es funcionalmente importante.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`

Cambios permitidos:

- Documentar inputs, marcadores, filas y columnas usados por el algoritmo.

Cambios prohibidos:

- No cambiar algoritmo aun.
- No cambiar capacidades esperadas.

Criterios de aceptacion:

- Queda documentado por que hoy se obtiene `Resources=60`, `Cost Planning=60`, `Monthly View=60`.

Comandos de validacion:

```powershell
rg -n "Leer estructura de plantilla CORE|capacidad|Monthly View|Resources|Cost Planning" .\lib\oportunidad-generar-core.xaml
```

Evidencia esperada:

- Resumen de algoritmo actual.

### EP5-T02 - Implementar lectura minima con fallback seguro

Contexto: se puede leer menos si la capacidad resultante es identica. Si no lo es, debe conservarse el metodo amplio.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`

Cambios permitidos:

- Leer solo marcadores/cabeceras necesarios.
- Usar fallback a lectura amplia si falta marcador o hay ambiguedad.
- Registrar en log si se usa fallback.

Cambios prohibidos:

- No asumir capacidad fija de 60.
- No omitir validacion de hojas obligatorias.
- No continuar si la plantilla no cumple estructura.

Criterios de aceptacion:

- Capacidades detectadas iguales a las actuales:
  - `Resources=60`
  - `Cost Planning=60`
  - `Monthly View=60`
- Comparador y validadores OK.
- Menor tiempo en bloque `Leer estructura de plantilla CORE`.

Comandos de validacion:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_plantillas_core_60_meses.ps1
```

Evidencia esperada:

- Logs o salida que confirmen capacidades.
- Profile antes/despues del bloque.

### EP5-T03 - Cachear capacidades durante ejecucion de `Main.xaml`

Contexto: `Main.xaml` genera PC y AT. Si una capacidad de plantilla ya fue medida para un fichero/tipo durante la ejecucion, puede reutilizarse siempre que la plantilla no cambie.

Archivos objetivo:

- `lib/oportunidad-generar-core.xaml`
- Posible helper interno si se decide extraer cache.

Cambios permitidos:

- Cache en memoria por ruta de plantilla, timestamp y longitud de fichero.
- Invalidar cache si cambia ruta, fecha o tamano.

Cambios prohibidos:

- No cachear entre ejecuciones de robot de forma persistente.
- No usar cache si no se puede demostrar que corresponde a la plantilla actual.

Criterios de aceptacion:

- PC y AT detectan capacidades correctas.
- No hay falsos positivos si se cambia una plantilla.

Comandos de validacion:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_plantillas_core_60_meses.ps1
```

Evidencia esperada:

- Explicacion de clave de cache.
- Resultado OK de validadores.

## Epica 6 - Profile final y documentacion de resultados

Objetivo: cerrar la optimizacion con evidencia objetiva de rendimiento y equivalencia.

### EP6-T01 - Ejecutar profile completo tras optimizaciones

Contexto: el profile inicial se corta a los 40000 eventos antes de terminar `CORE PC`. El objetivo minimo es que el profile capture PC y AT completos.

Archivos objetivo:

- Nuevo `.uistat` generado por UiPath Studio.
- `docs/backlog-optimizacion-rendimiento-core.md` o documento de resultados.

Cambios permitidos:

- Anadir resumen de mediciones antes/despues.

Cambios prohibidos:

- No ajustar el proceso para favorecer artificialmente el profile.
- No desactivar funcionalidad para medir mejor.

Criterios de aceptacion:

- `ActivityExecutionLimitReached=False`.
- Profile incluye generacion PC y AT completas.
- Tiempo total y top actividades documentados.

Comandos de validacion:

```powershell
Get-ChildItem -Filter *.uistat | Sort-Object LastWriteTime -Descending | Select-Object -First 3 Name,Length,LastWriteTime
```

Evidencia esperada:

- Ruta del nuevo profile.
- Tabla de tiempos antes/despues.

### EP6-T02 - Ejecutar suite de validacion final

Contexto: tras completar optimizaciones, hay que validar funcionalidad con scripts existentes y comparador.

Archivos objetivo:

- Outputs en `data/output`.
- Reporte en documentacion.

Cambios permitidos:

- Actualizar documentacion con resultados.

Cambios prohibidos:

- No modificar validadores para ocultar diferencias.

Criterios de aceptacion:

- `validar_calidad_e06.ps1`: OK.
- Validador AT 48 meses: OK.
- Validador PC 12 meses: OK.
- Comparador baseline/candidato: OK para PC y AT.

Comandos de validacion:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -WorkbookPath .\data\output\CORE_AT_20251160543.xlsx -Kind AT -ExpectedStart '2026-01-01' -Months 48
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -WorkbookPath .\data\output\CORE_PC_20250256445.xlsx -Kind PC -ExpectedStart '2025-11-01' -Months 12
```

Evidencia esperada:

- Salidas OK de todos los comandos.
- Comparador OK PC/AT.

### EP6-T03 - Documentar cierre tecnico

Contexto: el resultado debe quedar mantenible para futuros cambios.

Archivos objetivo:

- `docs/decisiones-tecnicas-core.md`
- `docs/guia-integracion-core.md`
- Este backlog, si se quiere marcar tareas completadas.

Cambios permitidos:

- Documentar decisiones finales de optimizacion.
- Documentar riesgos residuales y tareas descartadas.
- Documentar metricas antes/despues.

Cambios prohibidos:

- No declarar cerrada una optimizacion sin evidencia.
- No borrar contexto historico del profile inicial.

Criterios de aceptacion:

- Hay resumen claro de que se optimizo, que no se cambio y como se valido.
- Quedan instrucciones para repetir medicion.

Comandos de validacion:

```powershell
git diff -- docs
```

Evidencia esperada:

- Documentacion final revisable por otro mantenedor.

## Orden recomendado de implementacion

1. EP0-T01, EP0-T02, EP0-T03.
2. EP1-T01, EP1-T02, EP1-T03.
3. EP2-T01, EP2-T02.
4. EP3-T01, EP3-T02, EP3-T03.
5. EP4-T01, EP4-T02, EP4-T03.
6. EP5-T01, EP5-T02, EP5-T03.
7. EP6-T01, EP6-T02, EP6-T03.

No saltar EP0 salvo instruccion explicita del usuario. Sin comparador, las optimizaciones de rendimiento tienen demasiado riesgo de cambiar los CORE sin detectarlo.

## Definition of Done global

Una tarea se considera cerrada solo si:

- los cambios estan limitados al alcance de la tarea;
- se ha explicado cualquier decision tecnica relevante;
- se han ejecutado los comandos de validacion aplicables;
- los outputs CORE se mantienen equivalentes cuando la tarea toca comportamiento;
- no quedan procesos Excel residuales cuando se ejecuta runtime;
- se aporta evidencia suficiente para que otra conversacion pueda continuar desde el siguiente ID de tarea.

