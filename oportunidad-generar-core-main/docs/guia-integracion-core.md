# Guia de integracion CORE

`oportunidad-generar-core.xaml` es un submodulo importable. El proceso padre decide que oportunidad procesar, localiza el PPO, resuelve las plantillas oficiales PC/AT, calcula la ruta final del CORE y aporta los datos externos ya obtenidos de Salesforce/SAP u otro proceso.

El submodulo no usa assets, colas, Salesforce ni SAP. Solo copia la plantilla recibida, lee el PPO, rellena el CORE PC/AT y devuelve `out_NumeroSFLeido`.

## Arquitectura final

La fachada estable sigue siendo `oportunidad-generar-core.xaml`. El proceso padre solo debe invocar ese XAML y no necesita conocer los workflows internos salvo para despliegue de ficheros en la misma carpeta de proyecto.

| XAML | Responsabilidad | Consume | Produce / valida |
| --- | --- | --- | --- |
| `oportunidad-generar-core.xaml` | Orquestacion CORE: valida entradas tecnicas, copia plantilla, lee PPO por rangos, prepara Project Infor, delega PC/AT, pinta errores dinamicos y recalcula. | Argumentos `in_`, PPO, plantillas PC/AT y ruta final `in_RutaCORE`. | CORE generado, `out_NumeroSFLeido`, logs de inicio/rendimiento/fin/error y excepciones tecnicas relanzadas. |
| `lib/core-common-construir-modelo-ppo.xaml` | Interpretacion comun del PPO leido por rangos estables. | `Hoja de datos!A8:B35 (etiqueta + valor; interpretado por etiquetas)`, `Parametros!B18:J58`, `Presupuesto!B3:X53`, `Facturacion y SAP!A1:E10`, `Sintesis Precio!D12:P80`. | `dtMapaLecturaPPO`, `dtCabeceraPPO`, perfiles, costes, horas, gastos, compras, tarifas AT, meses, JSON trazable y tipo detectado. |
| `lib/core-common-preparar-resources.xaml` | Preparacion comun de `Resources` PC/AT. | Modelo PPO, matriz intercompany, WBS, fecha SAP, codigos de empleado y capacidad real de plantilla. | Tablas rectangulares PC/AT para `Write Range`, con errores funcionales preparados como texto corto. |
| `lib/core-common-preparar-cost-planning.xaml` | Preparacion comun de `Cost Planning` PC/AT. | Perfiles, horas, gastos, compras, meses, riesgos, garantia, datos SAP de compra y tabla AT de Resources. | Tablas PC/AT de horas, riesgos/garantia, gastos y compras. Valida capacidades de perfiles, gastos y compras. |
| `lib/oportunidad-generar-core-pc.xaml` | Escritura especifica del layout CORE PC. | Tablas comunes y ruta CORE copiada. | Limpieza/escritura consolidada de `Resources` y `Cost Planning` PC con una unica apertura Excel. |
| `lib/oportunidad-generar-core-at.xaml` | Escritura especifica del layout CORE AT. | Tablas comunes y ruta CORE copiada. | Limpieza/escritura consolidada de `Resources` y `Cost Planning` AT con una unica apertura Excel. |
| `oportunidad-generar-ipf.xaml` | Submodulo separado para IPF. | Plantilla IPF, ruta final y datos confirmados por el padre. | IPF generado. No forma parte del contrato CORE, pero `test/ejemplo_generar_core_ipf.xaml` muestra la integracion encadenada. |

Orden visual en la fachada CORE:

1. `00 Preparar datos locales para pruebas y ejecucion`.
2. `01 Validar entradas del modulo`.
3. `02 Seleccionar y copiar plantilla CORE`.
4. `03 Leer PPO por rangos estables`.
5. `04 Construir modelo comun`.
6. `04B Preparar reglas S08 de capacidad e intercompany`.
7. `05 Rellenar Project Infor`.
8. `06 Delegar Resources y Cost Planning por tipo`.
9. `08 Recalcular, guardar y registrar fin`.

Los `Invoke Code` que quedan son utilidades acotadas: construccion de tablas desde `DataTable`, reglas de normalizacion/prorrateo, mapa `Project Infor`, preparacion IPF y el cierre COM justificado para rojo dinamico y `CalculateFullRebuild`. Los XAML PC/AT no contienen `Invoke Code`.

## Ejemplo ejecutable

El workflow `test/ejemplo_invocacion_proceso_padre.xaml` simula una fila de oportunidad ya preparada por el proceso padre. La secuencia principal:

1. Asigna variables equivalentes a una fila de oportunidad.
2. Asegura la carpeta destino de ejemplo.
3. Invoca `oportunidad-generar-core.xaml` con `Invoke Workflow File`.
4. Recoge `out_NumeroSFLeido` y registra el fin OK.

Para probarlo, abrir el proyecto con UiPath Studio 23.10.4 y ejecutar `test/ejemplo_invocacion_proceso_padre.xaml`. La salida de ejemplo se escribe en `data\output`.

## Contrato de argumentos

Responsabilidades del proceso padre:

- Fase Salesforce/cierre: localizar el PPO, resolver `in_SF_RutaProyecto`, preparar `in_RutaCORE`, aportar cuenta, practica, sales manager y correo responsable si estan disponibles.
- SAP u otro proceso previo: aportar WBS/PEP, codigo cliente, Sales Order, PM, practice leader, datos de compras y codigos de empleado de recursos si existen.
- Plantillas: aportar siempre las rutas oficiales PC y AT, aunque solo se vaya a generar uno de los dos tipos.
- Valores vacios permitidos: fechas SAP, datos comerciales con fallback a PPO, datos trazables no funcionales y datos de compras/recursos cuando no existan. El modulo escribira texto rojo solo donde la plantilla necesite el dato.

| Argumento | Origen esperado en el padre | Puede ir vacio |
| --- | --- | --- |
| `in_RutaPPO` | Ruta absoluta del PPO descargado/localizado por el padre. | No |
| `in_RutaPlantillaCorePC` | Ruta absoluta de plantilla oficial CORE PC. | No |
| `in_RutaPlantillaCoreAT` | Ruta absoluta de plantilla oficial CORE AT. | No |
| `in_RutaCORE` | Ruta absoluta final, incluyendo nombre del fichero CORE. | No |
| `in_TipoProyecto` | Tipo decidido por el padre: `PC` o `AT`. | No |
| `in_SAP_FechaInicio` | Fecha real SAP si existe. Si falta, se usa PPO. | Si |
| `in_SAP_FechaFin` | Fecha fin SAP si existe. Si falta, se usa PPO. | Si |
| `in_SAP_CodigoPEP_WBS` | WBS/PEP desde SAP/Salesforce/proceso previo. | Si, queda texto rojo |
| `in_SAP_CodigoCliente` | Codigo cliente/Sold-to. | Si, queda texto rojo |
| `in_SAP_PedidoSalesOrder` | Pedido/Sales Order. | Si, queda texto rojo |
| `in_SAP_CodigoEmpleadoJefeProyecto` | Codigo empleado del PM. | Si, queda texto rojo |
| `in_SAP_NombreJefeProyecto` | Nombre del PM. | Si, queda texto rojo |
| `in_SAP_CodigoEmpleadoPracticeLeader` | Codigo empleado del practice leader. | Si, queda texto rojo |
| `in_SAP_NombrePracticeLeader` | Nombre del practice leader. | Si, queda texto rojo |
| `in_SF_CodigoEmpleadoSalesManager` | Codigo empleado del sales manager. | Si, queda texto rojo o fallback |
| `in_SF_NombreSalesManager` | Nombre del sales manager. Si falta, se usa Account Manager del PPO. | Si |
| `in_SF_AccountName` | Cuenta/cliente desde Salesforce. Si falta, se usa cliente del PPO. | Si |
| `in_SF_Practica` | Practica desde Salesforce. Si falta, se usa practica del PPO. | Si |
| `in_SF_RutaProyecto` | Trazabilidad del padre/ruta de oportunidad. | Si |
| `in_SF_CorreoResponsable` | Trazabilidad del responsable. | Si |
| `in_SAP_ProveedorCompra` | Proveedor de compras/subcontrataciones. | Si, queda texto rojo si hay compras |
| `in_SAP_PedidoCompra` | Pedido de compra. | Si, queda texto rojo si hay compras |
| `in_SAP_CodigoAriba` | Codigo Ariba. | Si, queda texto rojo si hay compras |
| `in_SAP_CodigosEmpleadoRecursosJson` | JSON opcional de sigla a codigo empleado de recurso. | Si |
| `out_NumeroSFLeido` | Salida con Opportunity Number leido del PPO. | N/A |

## Supuestos de integracion

- Studio obligatorio: 23.10.4.0.
- Target framework: Windows.
- Paquetes: los indicados en `project.json`; no anadir dependencias nuevas salvo decision explicita.
- Excel debe estar disponible en la maquina de ejecucion.
- La inspeccion de hojas/capacidad de la plantilla CORE se realiza con actividades `Read Range`; COM queda reservado al cierre final para pintado dinamico de errores y `CalculateFullRebuild`.
- La carpeta de `in_RutaCORE` debe existir antes de invocar el modulo.
- El padre mantiene el control transaccional del bucle de oportunidades. Si el submodulo relanza una excepcion tecnica, el padre decide si reintenta, marca KO o continua con la siguiente oportunidad.
- No se publica ni empaqueta como parte de este handover; el entregable es el XAML importable.

## Validaciones que quedan para el padre

El modulo CORE valida lo que necesita para no escribir un Excel corrupto o fuera de plantilla. El padre conserva las decisiones de negocio y de orquestacion que pertenecen al flujo de cierre.

| Validacion / decision del padre | Motivo | Comportamiento del CORE |
| --- | --- | --- |
| Seleccionar oportunidades y resolver la fila de trabajo. | CORE es submodulo importable, no proceso de extraccion. | No consulta Salesforce, SAP, colas ni assets. |
| Localizar PPO y plantillas oficiales. | Las rutas dependen del proceso de cierre y de los repositorios/document libraries del cliente. | Falla solo si las rutas recibidas no existen o la carpeta destino no existe. |
| Construir `in_RutaCORE` con naming y carpeta final. | El naming pertenece al proceso padre y puede variar por area/proyecto. | Copia exactamente a la ruta recibida; no decide nombres. |
| Mapear SAP/Salesforce a argumentos individuales. | El contrato se mantiene explicito para evitar estructuras opacas. | Usa cada dato si llega; si falta un dato funcional, escribe texto corto rojo cuando la plantilla lo requiere. |
| Decidir si tambien se genera IPF. | IPF es un submodulo separado y puede no aplicar a todas las oportunidades. | `test/ejemplo_generar_core_ipf.xaml` muestra CORE primero e IPF despues de forma condicional. |
| Gestionar reintentos, notificaciones y estado KO/OK. | El padre controla el bucle transaccional. | CORE registra contexto, relanza errores tecnicos y no captura decisiones de reintento. |
| Confirmar datos no disponibles en origen fiable: narrativa, hitos, facturacion emitida, condiciones contractuales. | No deben inventarse datos funcionales. | Deja vacio o escribe texto rojo segun el mapa; no genera auditoria externa. |

## Mapa tecnico PPO

La lectura del PPO se concentra en `03 Leer PPO por rangos estables` y la interpretacion en `lib/core-common-construir-modelo-ppo.xaml`. El mapa vivo `dtMapaLecturaPPO` tiene columnas `Campo`, `Hoja`, `Celda`, `Tipo`, `Obligatorio` y `OrigenFuncional`.

| Campo CORE/modelo | Origen PPO | Destino de modelo | Si falta |
| --- | --- | --- | --- |
| `company` | `Hoja de datos!B8` | `dtCabeceraPPO.company`, intercompany como apoyo | Queda vacio; intercompany puede usar WBS/SAP. |
| `unit` | `Hoja de datos!B9` | `dtCabeceraPPO.unit` | Queda vacio. |
| `location` | `Hoja de datos!B10` | `dtCabeceraPPO.location` | Queda vacio. |
| `serviceType` | `Hoja de datos!B11` | `dtCabeceraPPO.serviceType` | Queda vacio. |
| `pepType` | `Hoja de datos!B12` | `dtCabeceraPPO.pepType`, deteccion PC/AT | Si no indica AT, se interpreta como PC salvo tipo oportunidad AT. |
| `portfolio` | `Hoja de datos!B14` | `dtCabeceraPPO.portfolio` | Queda vacio. |
| `practice` | `Hoja de datos!B15` | `dtCabeceraPPO.practice`, `Project Infor` con fallback SF | Texto rojo `Practica no disponible` si tambien falta SF. |
| `subpractice` | `Hoja de datos!B16` | `dtCabeceraPPO.subpractice` | Queda vacio. |
| `opportunityType` | `Hoja de datos!B17` | `dtCabeceraPPO.opportunityType`, deteccion PC/AT | Si no indica AT, se interpreta como PC salvo PEP AT. |
| `accountManager` | `Hoja de datos!B18` | Fallback de sales manager | Texto rojo si tambien falta SF. |
| `applicant` | `Hoja de datos!B19` | Trazabilidad en modelo JSON | Queda vacio. |
| `subpracticeOrigin` | `Hoja de datos!B20` | Trazabilidad en modelo JSON | Queda vacio. |
| `opportunityCode` | `Hoja de datos!B22` | `NumeroSFExtraido`, `dtCabeceraPPO.opportunityCode` | Bloqueante: `No se ha podido leer Opportunity Number...`. |
| `customer` | `Hoja de datos!B23` | Fallback de cuenta/cliente | Texto rojo si tambien falta SF. |
| `description` | `Hoja de datos!B25` | `Project Infor` descripcion | Texto rojo `Descripcion no disponible en PPO`. |
| `proposalDelivery` | `Hoja de datos!B26` | Modelo JSON | Queda vacio. |
| `startMonth` / `startYear` | `Hoja de datos!B27:B28` | Calendario si no hay fecha PPO en SAP | Bloqueante si no hay fecha alternativa. |
| `duration` | `Hoja de datos!B29` | `dtMesesPPO`, capacidad plantilla | Bloqueante si no es valida. |
| `startDate` / `endDate` | `Facturacion y SAP!C4/E4` | `dtCabeceraPPO.startDate/endDate` | Se calcula desde mes/anio/duracion si falta. |
| `orderAmount` | `Presupuesto!X8` o `Sintesis Precio!D12` | `Project Infor C28`, JSON financiero | Texto rojo `Importe no disponible en PPO`. |
| `resourceCost` / `hoursEstimated` | `Presupuesto!K4/K5` | `Project Infor C29:C30`, JSON financiero | Texto rojo especifico de horas/coste. |
| `expenses` / `purchases` | `Presupuesto!K28/K34` | `Project Infor C33:C34`, JSON financiero | Texto rojo especifico de gastos/compras. |
| `risks` / `warranty` | `Presupuesto!K49/K53` | `Project Infor C31:C32`, Cost Planning ultimo mes | Texto rojo especifico de riesgos/garantia. |
| Perfiles/costes | `Parametros!B18:J58` | `dtPerfilesPPO`, `dtCostesPorAnio` | Coste no disponible se escribe en rojo en `Resources`. |
| Horas por anio | `Presupuesto!B3:X53` | `dtHorasPorAnio` | No se crea linea de horas si no hay importe. |
| Gastos/compras detalle | `Presupuesto!B3:X53` | `dtGastos`, `dtCompras` | Se omiten lineas sin descripcion o importe. |
| Tarifas AT | `Sintesis Precio!D12:P80` | `dtTarifasAT`, `Resources` AT | Texto rojo `Tarifa AT no disponible`. |

## Mapa Project Infor

La secuencia `05 Rellenar Project Infor` esta dividida en bloques visibles:

- `05A Fechas e identificadores`: fechas, WBS, descripcion, SLFC y Sales Order.
- `05B Cliente y practica`: practica, cliente SAP y cuenta con fallback Salesforce/PPO.
- `05C Responsables`: PM, practice leader y sales manager.
- `05D Totales economicos`: importe, horas, coste, riesgos, garantia, gastos y compras.

Despues de esos bloques, la utilidad acotada `Utilidad S05 - construir mapa Project Infor` crea `dtMapaProjectInfor` con `Bloque`, `Campo`, `Celda`, `Valor`, `AplicarRojo`, `TextoError` y `OrigenFuncional`. La escritura usa una unica iteracion sobre ese mapa: `WriteCellX` escribe la celda indicada y `FormatRangeX` aplica fuente roja solo cuando `AplicarRojo = True`.

La escritura dinamica de celda queda aceptada en Studio 23.10 para este caso porque el mapa usa direcciones fijas de plantilla (`C5`, `C8`, `C9`, `C10`, `C11`, `C12`, `C13`, `C16`, `D16`, `C18`, `C19`, `C22`, `D22`, `C23`, `D23`, `C24`, `D24`, `C28:C34`) y el validador estatico protege que no reaparezca el bloque monolitico ni los `WriteCellX` fijos repetidos.

## Resources y Cost Planning

La preparacion comun de `Resources` vive en `lib/core-common-preparar-resources.xaml`. Ese subworkflow normaliza compania, codigos de empleado, intercompany, coste base, tarifa AT y construye las tablas rectangulares PC/AT para escritura por `Write Range`.

La preparacion comun de `Cost Planning` vive en `lib/core-common-preparar-cost-planning.xaml`. Calcula prorrateos mensuales, riesgos, garantia, gastos, compras y textos funcionales de compras una sola vez. Para evitar duplicar reglas de intercompany, consume la tabla fija AT generada por Resources (`dt_ResourcesATFijos`) cuando necesita reproducir codigo, sigla, intercompany y perfil en Cost Planning AT.

Los subworkflows `oportunidad-generar-core-pc.xaml` y `oportunidad-generar-core-at.xaml` ya no contienen `Invoke Code`: invocan los common workflows y conservan solo la limpieza/escritura de rangos especificos de cada layout. Desde E05, cada subworkflow prepara primero `Resources` y `Cost Planning` y abre el CORE una sola vez para escribir ambos bloques por `Write Range`.

## Errores y logs

Errores bloqueantes: PPO inexistente, plantilla inexistente, carpeta destino inexistente, tipo distinto de `PC`/`AT`, hoja obligatoria ausente, libro corrupto, Excel no disponible, duracion invalida o plantilla sin capacidad.

Datos funcionales faltantes: no bloquean la generacion. En `Project Infor`, `dtErroresFuncionales` filtra las filas del mapa con `AplicarRojo = True`; en `Resources` y `Cost Planning`, el pintado dinamico final mantiene la misma politica de texto corto en rojo.

Logs del submodulo:

- Inicio: PPO, tipo y ruta CORE destino.
- Rendimiento E05: segundos aproximados para lectura PPO, escritura CORE y recalculo final.
- Fin OK: ruta CORE generada, `out_NumeroSFLeido` y tipo normalizado.
- Error: PPO, CORE y detalle de la excepcion, relanzada al proceso padre.

## Validacion recomendada

1. Ejecutar `test/test_generar_core_pc.xaml` desde Studio para generar `data\output\test\CORE_PC_20250256445_test.xlsx`.
2. Ejecutar `powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_core_pc.ps1`.
3. Ejecutar `powershell -ExecutionPolicy Bypass -File .\test\validar_negativos_core.ps1`.
4. Ejecutar `test/test_generar_core_datos_incompletos.xaml` desde Studio y validar el resultado con `powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_core_pc.ps1 -CorePath .\data\output\test\CORE_PC_20250256445_datos_incompletos.xlsx -Scenario DatosIncompletos -SkipBaseline`.
5. Ejecutar `powershell -ExecutionPolicy Bypass -File .\test\validar_calidad_e06.ps1` para cerrar XML, contratos negativos y Workflow Analyzer focalizado.
6. Abrir visualmente `oportunidad-generar-core.xaml` en Studio antes de integrarlo en el proyecto padre.

## Paquete de handover

Entregar juntos:

- `oportunidad-generar-core.xaml`.
- Carpeta `lib`.
- `oportunidad-generar-ipf.xaml` si el proceso padre va a generar IPF.
- Carpeta `docs`, especialmente esta guia, `docs/guia-integracion-ipf.md` y `docs/decisiones-tecnicas-core.md`.
- Carpeta `test` para wrappers, baselines y validadores PowerShell.

No entregar como requisito: paquete publicado, assets, colas, configuracion de Orchestrator ni dependencias nuevas.
