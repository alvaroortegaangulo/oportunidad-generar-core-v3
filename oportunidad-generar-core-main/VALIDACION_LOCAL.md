# Validación local de empaquetado

Validación realizada en el entorno de generación del ZIP. No sustituye la ejecución final en UiPath Studio/Windows con Excel instalado.

## Comprobaciones estáticas

- XAML parseados como XML: 15.
- Todas las referencias `WorkflowFileName="lib\..."` apuntan a ficheros existentes.
- No se incluyen directorios generados por UiPath Studio ni caches locales de compilacion.
- Referencias obsoletas detectadas en XAML/MD/JSON: 0.

## Lectura semántica comprobada en PPO

### `data/input/20250256445  PPO_NuevaTerminalPVR v2 1.xlsm`

- Tipo detectado: `PC`
- Opportunity: `20250256445`
- Cliente: `GAP`
- Descripción: `Smart Airport en la nueva terminal de Puerto Vallarta`
- Inicio: mes `11`, año `2025`
- Duración: `12` meses
- Presupuesto!X8: `583896.55`
- Sintesis Precio!D12: `583896.55`
- Importe elegido por la regla de mayor importe numérico: `583896.55`

### `data/input/20251160543_PPO_CEducación_AMS_SI_Lote 2_v01.xlsm`

- Tipo detectado: `AT`
- Opportunity: `20251160543`
- Cliente: `Consejería Educación JCyL`
- Descripción: `Lote 2: Núcleo Estructural Java`
- Inicio: mes `1`, año `2026`
- Duración: `48` meses
- Presupuesto!X8: `3478.4`
- Sintesis Precio!D12: `465000`
- Importe elegido por la regla de mayor importe numérico: `465000`

## Plantillas

### `data/templates/CORE_PC_template.xlsx`

- Hojas detectadas: `Project Infor, Cost Overview, Resources, Cost Planning, Monthly View, Cost Summary, Ayuda, aux Billing Plan`
- Dimensiones relevantes: `{'Resources': 'B2:BZ69', 'Cost Planning': 'B1:CA96'}`
- Capacidad mensual validada: `Resources F:BM = 60`, `Cost Planning G:BN = 60`, `Monthly View E:BL = 60`
- Formulas OpenXML: `3801`; validaciones/listas: `9`; celdas con estilo: `17939`
- `#REF!` en XML de plantilla: `0`

### `data/templates/CORE_AT_template.xlsx`

- Hojas detectadas: `Project Infor, Cost Overview, Resources, Cost Planning, Monthly View, Cost Summary, Ayuda, aux Billing Plan`
- Dimensiones relevantes: `{'Resources': 'B2:CA69', 'Cost Planning': 'B1:CB99'}`
- Capacidad mensual validada: `Resources G:BN = 60`, `Cost Planning H:BO = 60`, `Monthly View E:BL = 60`
- Formulas OpenXML: `4060`; validaciones/listas: `10`; celdas con estilo: `18347`
- `#REF!` en XML de plantilla: `0`

### `data/templates/IPF_template.xlsx`

- Hojas detectadas: `Invoice Request, UC Maint Contract+Inv Request, Billing template instructions, Lists, UC WBS codes`
- Dimensiones relevantes: `{}`
- `#REF!` en XML de plantilla: `0`

## Validacion especifica E04

- Se incorpora `test/validar_plantillas_core_60_meses.ps1` para validar de forma repetible las plantillas CORE extendidas sin abrir Excel.
- `lib/oportunidad-generar-core.xaml` mide capacidad mensual real desde `Resources`, `Cost Planning` y `Monthly View`; no delega ya en un limite fijo.
- `lib/core-common-preparar-resources.xaml` y `lib/core-common-preparar-cost-planning.xaml` reciben la capacidad mensual medida y fallan si la duracion PPO la supera.
- `lib/core-common-preparar-cost-planning.xaml` valida explicitamente que existe el indice del ultimo mes antes de imputar riesgos y garantia.
- `lib/oportunidad-generar-core-pc.xaml` y `lib/oportunidad-generar-core-at.xaml` validan capacidad mensual justo antes de escribir rangos Excel.

Comando ejecutado:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_plantillas_core_60_meses.ps1
```

Resultado: OK para CORE PC/AT con 60 meses utiles y sin `#REF!`.

## Validacion especifica E06

- Se incorpora `test/validar_calidad_e06.ps1` para proteger `Cost Planning` sin depender de COM ni de Excel abierto.
- El contrato estatico valida que `dtHorasPorAnio` conserva `ResourceKey`, que `Cost Planning` agrupa horas por clave robusta y que riesgos/garantia usan `duracionCostPlanning`, no 25 columnas fijas.
- El output PC conserva los importes funcionales de referencia: pedido `583896.55`, horas `13570`, coste recursos `367515.14`, riesgos `18375.76`, gastos `7400` y compras `40000`.
- El output AT conserva pedido `465000`, duracion `48` meses, tres recursos de negocio y dos filas por recurso (`Horas Reales` y `Horas/Jornadas Facturables`) con base inicial replicada para revision.
- El bloque de compras AT queda vacio cuando el PPO no trae compras reales; no se crean compras falsas desde riesgos/garantia.

Comando ejecutado:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
```

Resultado: OK para `ResourceKey`, prorrateo anual, AT real/facturable, riesgos en ultimo mes, cabeceras AT hasta diciembre 2029 y ausencia de compras falsas.

## Ejecucion runtime UiRobot E04

Paquete ejecutado: `oportunidad-generar-core-main.1.0.26-e04.nupkg`.

Comandos ejecutados:

```powershell
& 'C:\Program Files\UiPath\Studio\UiRobot.exe' execute --file '..\.local\packages\oportunidad-generar-core-main.1.0.26-e04.nupkg' --entry 'test\test_generar_core_at.xaml'
& 'C:\Program Files\UiPath\Studio\UiRobot.exe' execute --file '..\.local\packages\oportunidad-generar-core-main.1.0.26-e04.nupkg' --entry 'test\test_generar_core_pc.xaml'
```

Resultados:

- AT `20251160543`: capacidad detectada `Resources=60`, `Cost Planning=60`, `Monthly View=60`; duracion PPO `48/60`; ejecucion `00:05:33`; output validado y copiado a `data/output/CORE_AT_20251160543_test.xlsx`.
- PC `20250256445`: capacidad detectada `Resources=60`, `Cost Planning=60`, `Monthly View=60`; duracion PPO `12/60`; ejecucion `00:04:44`; output validado y copiado a `data/output/CORE_PC_20250256445_test.xlsx`.

Validaciones de salida:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -Kind AT -ExpectedStart '2026-01-01' -Months 48 -Path .\data\output\CORE_AT_20251160543_test.xlsx
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -Kind PC -ExpectedStart '2025-11-01' -Months 12 -Path .\data\output\CORE_PC_20250256445_test.xlsx
```

Resultado: OK en ambos outputs, cabeceras mensuales correctas y `#REF! = 0`.

## Ejecucion runtime UiRobot E06

Paquete ejecutado: `oportunidad-generar-core-main.1.0.29-e06.nupkg`.

Comandos ejecutados:

```powershell
& 'C:\Program Files\UiPath\Studio\UiRobot.exe' pack 'C:\Users\aortega\Documents\UiPath\oportunidad-generar-core-v3\oportunidad-generar-core-main\project.json' --output 'C:\Users\aortega\Documents\UiPath\oportunidad-generar-core-v3\.local\packages' -v 1.0.29-e06
& 'C:\Program Files\UiPath\Studio\UiRobot.exe' execute --file 'C:\Users\aortega\Documents\UiPath\oportunidad-generar-core-v3\.local\packages\oportunidad-generar-core-main.1.0.29-e06.nupkg' --entry 'test\test_generar_core_at.xaml'
& 'C:\Program Files\UiPath\Studio\UiRobot.exe' execute --file 'C:\Users\aortega\Documents\UiPath\oportunidad-generar-core-v3\.local\packages\oportunidad-generar-core-main.1.0.29-e06.nupkg' --entry 'test\test_generar_core_pc.xaml'
```

Resultados:

- AT `20251160543`: ejecucion `00:04:28`; output generado en cache NuGet y copiado a `data/output/CORE_AT_20251160543_test.xlsx`.
- PC `20250256445`: ejecucion `00:03:47`; output generado en cache NuGet y copiado a `data/output/CORE_PC_20250256445_test.xlsx`.

Validaciones tras copiar outputs:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -WorkbookPath .\data\output\CORE_AT_20251160543_test.xlsx -Kind AT -ExpectedStart '2026-01-01' -Months 48
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_salida_core_at_48_meses_e04.ps1 -WorkbookPath .\data\output\CORE_PC_20250256445_test.xlsx -Kind PC -ExpectedStart '2025-11-01' -Months 12
```

Resultado: OK. E06 validado con `ResourceKey`, prorrateo anual, AT real/facturable, riesgos en ultimo mes, 48 meses y ausencia de compras falsas.
