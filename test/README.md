# Validacion local CORE e IPF

Este directorio contiene la base de validacion tecnica de `oportunidad-generar-core.xaml` y `oportunidad-generar-ipf.xaml`.

## Wrapper PC

`test_generar_core_pc.xaml` invoca `oportunidad-generar-core.xaml` con:

- PPO `ficheros-auxiliares\20250256445  PPO_NuevaTerminalPVR v2.xlsm`.
- Plantilla `ficheros-auxiliares\PMBox_plantilla CORE_PC_v1_04 (2).xlsx`.
- Salida `.local\test-output\CORE_PC_20250256445_test.xlsx`.
- Datos SAP/Salesforce de prueba controlados para que el resultado sea comparable.

Desde S02, el wrapper llama al módulo con el contrato público prefijado (`in_` / `out_`) y recoge `out_NumeroSFLeido`. El XAML principal copia esos argumentos a variables locales en `00 Preparar datos locales para pruebas y ejecución`.

Se debe ejecutar desde UiPath Studio 23.10.4 abriendo el proyecto y lanzando el workflow de test. No publica, no empaqueta y no modifica ficheros fuente.

`test_generar_core_datos_incompletos.xaml` cubre el escenario funcional negativo E06: usa PPO y plantilla validos, pero deja vacios los datos SAP/Salesforce que no deben bloquear la generacion. La salida esperada es `.local\test-output\CORE_PC_20250256445_datos_incompletos.xlsx`; el validador integral comprueba estructura y textos rojos clave con:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_core_pc.ps1 -CorePath .\.local\test-output\CORE_PC_20250256445_datos_incompletos.xlsx -Scenario DatosIncompletos -SkipBaseline
```

## Ejemplo de integracion

`ejemplo_invocacion_proceso_padre.xaml` simula una fila de oportunidad ya preparada por el proceso padre y muestra el patron de integracion con `Invoke Workflow File`. Usa variables de fila para mapear todos los argumentos `in_`, recoge `out_NumeroSFLeido` y deja la salida en `.local\handover-output`.

La guia de handover esta en `docs\guia-integracion-core.md` y las decisiones cerradas estan en `docs\decisiones-tecnicas-core.md`.

Desde S13, `ejemplo_generar_core_ipf.xaml` muestra el patron completo: invoca primero `oportunidad-generar-core.xaml` y despues `oportunidad-generar-ipf.xaml` si `GenerarIPF` esta activado. Deja ambos documentos en `.local\handover-core-ipf-output` y mantiene logs separados para CORE e IPF.

## Wrappers IPF S13

`test_generar_ipf.xaml` valida el contrato de `oportunidad-generar-ipf.xaml`. Desde S12 el submodulo IPF copia la plantilla recibida a `.local\test-output\IPF_20250256445_test.xlsx`, lee listas desde `Lists`, rellena `Invoice Request` con datos de ejemplo y devuelve `out_NumeroSFLeido`.

El resultado esperado incluye cabecera (`Request Type`, `Bill Type`, `Company Issuer`, cliente, Sold-to, PM, moneda), `Sales Order`, WBS, PO, summary/narrativa/importe si se aportan y comentarios. Los campos no confirmados de direccion, narrativa y comentarios de la plantilla se limpian para no arrastrar datos de ejemplo.

`test_generar_ipf_datos_incompletos.xaml` valida el caso negativo funcional: plantilla y ruta validas, pero SAP/Salesforce y narrativa sin aportar. El modulo debe generar el IPF sin abortar; los datos obligatorios quedan como texto corto rojo y los campos narrativos/importe/comentarios quedan vacios.

El mapa de celdas, listas y contrato IPF esta en `docs\guia-integracion-ipf.md`.

## Baseline

`baseline_core_pc.csv` lista celdas esperadas de `Project Infor`, `Cost Overview`, `Resources` y `Cost Planning`. Los valores `NO`/`INTERCO` se mantienen como literales oficiales de plantilla derivados de la matriz intercompany S08.

`baseline_ipf_referencias.csv` valida la plantilla IPF oficial y el ejemplo GAP. La plantilla recibida como "Plantilla IPF" contiene datos de ejemplo; por eso el baseline los documenta y el baseline generado comprueba que el modulo limpia las zonas no confirmadas.

`baseline_ipf_generado.csv` contiene los escenarios IPF `Completo` y `DatosIncompletos`.

Despues de ejecutar el wrapper, comparar el Excel generado con:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_baseline_core_pc.ps1
```

El script usa Excel COM en modo solo lectura y falla indicando hoja/celda si algun valor no coincide.

Desde S09, la validacion recomendada es ejecutar el script integral:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_core_pc.ps1
```

Ese script abre el CORE generado, fuerza recalculo, guarda y reabre el libro. Comprueba hojas obligatorias, ausencia de `Trazabilidad_RPA`, numero de hojas, dimensiones basicas, cuadros `NOTA`, recuento de formulas, errores de formula y baseline de celdas criticas.

Para IPF, se puede validar primero solo plantilla y ejemplo GAP:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_ipf.ps1 -SkipGenerated
```

Despues de ejecutar `test/test_generar_ipf.xaml` desde Studio:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_ipf.ps1
```

Despues de ejecutar `test/test_generar_ipf_datos_incompletos.xaml` desde Studio:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_ipf.ps1 -IpfPath .\.local\test-output\IPF_20250256445_datos_incompletos.xlsx -Scenario DatosIncompletos
```

## Casos negativos

La ejecucion runtime de negativos sigue requiriendo UiPath Studio o un proceso padre, porque `UiRobot` no ejecuta XAML sueltos en este proyecto. Para cubrir S09 desde PowerShell existe una validacion estatica del contrato negativo:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_negativos_core.ps1
```

El script comprueba que el XAML mantiene errores bloqueantes accionables para rutas, tipo, plantilla, hojas obligatorias, Excel, duracion y capacidad, y que los datos funcionales faltantes se escriben como textos cortos preparados para fuente roja. Desde E03 tambien protege que `Project Infor` siga dividido en `05A`-`05D`, con `dtMapaProjectInfor`/`dtErroresFuncionales` y escritura dinamica por mapa. Desde E04 valida que PC/AT no vuelvan a contener `Invoke Code`, que existan `core-common-preparar-resources.xaml` y `core-common-preparar-cost-planning.xaml`, y que Cost Planning reutilice la resolucion de Resources para intercompany/codigos. Desde E05 valida que la capacidad de plantilla se mida con actividades Excel/DataTables, que PC/AT tengan una unica apertura Excel consolidada y que existan logs de rendimiento para lectura PPO, escritura CORE y recalculo final.

Para IPF, S13 anade una validacion estatica equivalente:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_negativos_ipf.ps1
```

Este script comprueba errores bloqueantes de plantilla/ruta/carpeta, tratamiento de datos funcionales faltantes, defaults de listas, limpieza de narrativa no confirmada y que el wrapper CORE+IPF invoque CORE antes que IPF.

## Cierre E06

`validar_calidad_e06.ps1` orquesta el cierre focalizado de calidad: parsea los XAML CORE/IPF y wrappers, ejecuta validadores negativos CORE/IPF, valida referencias IPF, lanza Workflow Analyzer de forma secuencial y documenta las reglas aceptadas por contrato/diseno.

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
```

Si ya se han generado los Excel desde Studio, el mismo script ejecuta tambien los validadores integrales sobre los ficheros existentes. Para obligar a que esos Excel existan en un cierre manual de sprint:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1 -RequireGeneratedWorkbooks
```

## Cierre E07 / handover

Para entregar el modulo a integracion, revisar junto con el proceso padre:

- `docs\guia-integracion-core.md`: arquitectura final, contrato CORE, datos consumidos, validaciones internas y validaciones que quedan para el padre.
- `docs\decisiones-tecnicas-core.md`: decisiones cerradas y reglas de mantenimiento.
- `docs\guia-integracion-ipf.md`: contrato IPF si la oportunidad debe generar factura.
- `test\ejemplo_invocacion_proceso_padre.xaml`: ejemplo minimo CORE.
- `test\ejemplo_generar_core_ipf.xaml`: ejemplo completo CORE + IPF condicional.

El cierre documental no sustituye a la prueba runtime manual: ejecutar los wrappers desde UiPath Studio 23.10.4 y repetir `validar_calidad_e06.ps1 -RequireGeneratedWorkbooks` cuando existan los Excel generados.
