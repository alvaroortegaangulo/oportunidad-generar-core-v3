# oportunidad-generar-core-main

Proyecto UiPath limpio para validar y reutilizar los workflows de generación de CORE a partir de PPO. La carpeta está preparada para abrirse directamente en UiPath Studio desde `Main.xaml`.

## Qué genera el `Main.xaml`

El `Main.xaml` ejecuta dos invocaciones reales a `lib\oportunidad-generar-core.xaml`:

1. **CORE PC** desde `data\input\20250256445  PPO_NuevaTerminalPVR v2 1.xlsm`.
2. **CORE AT** desde `data\input\20251160543_PPO_CEducación_AMS_SI_Lote 2_v01.xlsm`.

Los resultados se escriben en:

- `data\output\CORE_PC_20250256445.xlsx`
- `data\output\CORE_AT_20251160543.xlsx`

Los datos que normalmente llegarían desde Salesforce y SAP están declarados en el propio `Main.xaml` como valores de prueba controlados. No se usan rutas absolutas ni recursos fuera de esta carpeta.

## Estructura

```text
oportunidad-generar-core-main\
  Main.xaml
  project.json
  README.md
  lib\
    oportunidad-generar-core.xaml
    oportunidad-generar-ipf.xaml
    core-common-construir-modelo-ppo.xaml
    core-common-preparar-resources.xaml
    core-common-preparar-cost-planning.xaml
    oportunidad-generar-core-pc.xaml
    oportunidad-generar-core-at.xaml
    core-recalcular-final.txt
  data\
    input\
      20250256445  PPO_NuevaTerminalPVR v2 1.xlsm
      20251160543_PPO_CEducación_AMS_SI_Lote 2_v01.xlsm
    templates\
      CORE_PC_template.xlsx
      CORE_AT_template.xlsx
      IPF_template.xlsx
    output\
  test\
    *.xaml
  docs\
    guia-integracion-core.md
    guia-integracion-ipf.md
    decisiones-tecnicas-core.md
    documento_funcional_core_ipf_uipath.pdf
```

## Workflows principales

### `lib\oportunidad-generar-core.xaml`

Fachada pública para generar un CORE PC o AT. Copia la plantilla correspondiente, lee el PPO, construye el modelo común, rellena `Project Infor`, invoca el workflow específico PC/AT y aplica el cierre/recalculo final.

Contrato de entrada más relevante:

- `in_RutaPPO`
- `in_RutaPlantillaCorePC`
- `in_RutaPlantillaCoreAT`
- `in_RutaCORE`
- `in_TipoProyecto` (`PC` o `AT`)
- campos SAP: fechas, WBS, cliente, sales order, responsables, compras
- campos Salesforce: account, practice, sales manager, ruta de proyecto, correo responsable

Salida:

- `out_NumeroSFLeido`: número de oportunidad detectado en el PPO.

### `lib\core-common-construir-modelo-ppo.xaml`

Construye el modelo de negocio común desde los rangos del PPO. La cabecera ya no depende de posiciones fijas como `B8:B29`; interpreta `Hoja de datos` por etiquetas normalizadas. Esto permite absorber variaciones esperables de la plantilla, por ejemplo `Company` frente a `Entidad Presentadora`, `Opportunity Code` frente a `Código Oferta/Proyecto`, o `PEP Type` frente a `Tipo de Contratación`.

También normaliza el importe total de venta del proyecto: la fuente prioritaria es el total semántico de `Sintesis Precio` (`Total Oferta sin IVA` / `Precio Total`). Esto cubre el PPO AT de 48 meses, donde el precio total fiable aparece en la síntesis y no debe confundirse con el Sales Order o número de pedido.

### `lib\core-common-preparar-resources.xaml`

Prepara las tablas rectangulares de `Resources` a partir de perfiles, meses y costes/tarifas. Recibe la capacidad mensual real medida desde la plantilla antes de escribir.

### `lib\core-common-preparar-cost-planning.xaml`

Prepara las tablas rectangulares de `Cost Planning`: horas, gastos, compras, riesgos y garantía. Riesgos y garantía se imputan al último mes real del proyecto y se valida que esa columna exista en la capacidad mensual detectada.

### `lib\oportunidad-generar-core-pc.xaml` y `lib\oportunidad-generar-core-at.xaml`

Escriben las hojas específicas de CORE PC y CORE AT. AT contempla tarifa de venta además de coste, y planificación real/facturable.

### `lib\oportunidad-generar-ipf.xaml`

Se incluye para reutilización en el proceso padre. El `Main.xaml` no ejecuta IPF en este test, porque el objetivo de esta carpeta es validar la generación de los dos CORE.

## Rutas relativas usadas por el test

Todas las rutas se construyen desde `Environment.CurrentDirectory`:

- Entrada: `data\input\...`
- Plantillas: `data\templates\...`
- Salida: `data\output\...`
- Workflows importables: `lib\...`

Al ejecutar desde UiPath Studio, abre el proyecto usando esta carpeta como raíz para que `Environment.CurrentDirectory` apunte a `oportunidad-generar-core-main`.

## Dependencias de entorno

- UiPath Studio compatible con Windows y los paquetes declarados en `project.json`.
- Microsoft Excel instalado. La lectura y escritura se hace con actividades modernas de Excel (`ExcelProcessScopeX`, `ExcelApplicationCard`, `ReadRangeX`, `WriteRangeX`, `WriteCellX`).
- Los PPO de entrada son `.xlsm`; las salidas CORE son `.xlsx`.

No se incluyen metadatos generados por UiPath Studio, caches de compilacion ni outputs previos. UiPath Studio debe recompilar desde los XAML incluidos.

## Plantillas CORE de 60 meses

Las plantillas `CORE_PC_template.xlsx` y `CORE_AT_template.xlsx` son las variantes extendidas incluidas en la referencia de trabajo y adoptadas tras validacion local. La fachada `lib\oportunidad-generar-core.xaml` mide la capacidad mensual real en `Resources`, `Cost Planning` y `Monthly View`; si un PPO supera esa capacidad, el robot falla antes de escribir con la duracion detectada y la capacidad de cada hoja.

Validacion estatica disponible:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_plantillas_core_60_meses.ps1
```

## Cómo ejecutar

1. Descomprime el ZIP.
2. Abre UiPath Studio.
3. Abre la carpeta `oportunidad-generar-core-main` como proyecto.
4. Ejecuta `Main.xaml`.
5. Revisa los dos ficheros generados en `data\output`.

Antes de reejecutar, borra los `.xlsx` anteriores de `data\output` si quieres validar que se han creado desde cero.

## Tests

La carpeta `test` contiene wrappers XAML de invocacion y validadores PowerShell. `validar_plantillas_core_60_meses.ps1` inspecciona las plantillas como OpenXML y comprueba hojas obligatorias, formulas, validaciones, ausencia de `#REF!` y capacidad mensual de 60 meses en `Resources`, `Cost Planning` y `Monthly View`. `validar_salida_core_at_48_meses_e04.ps1` valida outputs generados PC/AT comprobando cabeceras mensuales reales y ausencia de `#REF!`. `validar_calidad_e06.ps1` cierra la robustez de `Cost Planning`: identidad `ResourceKey`, prorrateo por anualidad, AT real/facturable, riesgos en ultimo mes, 48 meses y ausencia de compras falsas.

Validacion E06:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1
```

## Criterios funcionales aplicados

- No hay bifurcación rígida por “layout antiguo/nuevo”. La lectura de cabecera se basa en etiquetas normalizadas y sin acentos.
- La duración se deriva del PPO y se propaga a Resources y Cost Planning.
- Se mantiene la estrategia de pre-relleno: cuando faltan datos reales de Salesforce/SAP, el `Main.xaml` usa valores de prueba explícitos; no se inventan valores desde el PPO.
- Los campos incompletos que pertenecen a fuentes externas deben quedar trazables como datos de prueba en el ejemplo, no ocultos como si vinieran de producción.
- En PC, la duración del PPO gobierna el calendario CORE; si el PPO contiene anualidades fuera de calendario, se conserva la baseline completa y se cuantifica el pendiente en `Cost Summary!B18`.
- En AT, la venta facturable planificada se reconcilia contra el importe total de venta del proyecto; si se reescalan tarifas proporcionalmente, `Cost Summary!B18` lo deja auditado con prefijo `RPA_VALIDACION:`.
