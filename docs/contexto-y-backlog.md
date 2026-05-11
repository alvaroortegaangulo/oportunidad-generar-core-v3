# Contexto y backlog - oportunidad-generar-core

Fecha de análisis: 2026-04-29  
Proyecto UiPath objetivo: `oportunidad-generar-core-v2`  
Módulo principal: `oportunidad-generar-core.xaml`  
Versión obligatoria de Studio: `23.10.4.0`  
Target framework: `Windows`  

## 1. Resumen ejecutivo

El objetivo real del módulo no es crear un proceso autónomo de extracción desde Salesforce o SAP. El objetivo es entregar un submódulo UiPath importable desde otro proceso, llamado `oportunidad-generar-core.xaml`, que reciba por argumentos el PPO, las plantillas CORE PC/AT, la ruta de salida y los datos externos ya resueltos por el proceso padre.

La automatización debe prerrellenar un CORE PC o AT desde el PPO, usando datos complementarios SAP/Salesforce cuando existan. Cuando un dato funcional no esté disponible, debe escribir en la celda correspondiente un texto corto de error en rojo, sin abortar la generación. El robot no debe inventar datos, periodificaciones finas, textos de facturación ni modificar formatos de plantilla. Debe respetar la plantilla oficial y entregar una base revisable por el jefe de proyecto.

El estado actual del XAML es un buen prototipo funcional, pero no está todavía al nivel de mantenibilidad esperado para producción. La deuda principal está en que dos bloques `Invoke Code` siguen concentrando demasiada lógica: lectura del PPO, construcción del modelo, escritura de CORE, reglas PC/AT, prorrateo, intercompany y uso COM de Excel. Además, el analizador de UiPath marca un error por exceso de argumentos.

## 2. Versiones y restricciones del proyecto

Versiones confirmadas en `project.json`:

| Elemento | Valor |
| --- | --- |
| UiPath Studio | `23.10.4.0` |
| Target framework | `Windows` |
| Lenguaje de expresiones | `VisualBasic` |
| Main | `oportunidad-generar-core.xaml` |
| `UiPath.Excel.Activities` | `[2.22.3]` |
| `UiPath.Mail.Activities` | `[1.21.1]` |
| `UiPath.System.Activities` | `[23.10.5]` |
| `UiPath.Testing.Activities` | `[23.10.1]` |
| `UiPath.UIAutomation.Activities` | `[23.10.7]` |

Restricciones clave:

- No añadir dependencias nuevas salvo decisión explícita.
- No usar assets ni colas dentro del submódulo.
- No generar JSON intermedios en disco ni auditoría Markdown desde el módulo final.
- La carpeta destino del CORE debe existir antes de invocar el módulo.
- El entregable operativo será el XAML final para que otro compañero lo integre en su proyecto; no se preparará publicación, paquete ni configuración específica de Orchestrator.
- No se debe abortar por datos funcionales faltantes para rellenar el CORE. Si un dato no existe o es inválido, se escribirá en la celda correspondiente un texto corto de error en rojo.
- El submódulo debe ser entendible por un desarrollador UiPath desde Studio, con secuencias, nombres y anotaciones en español.
- El uso de `Invoke Code` debe quedar limitado a utilidades o cálculos que no sean razonables de expresar con actividades.

## 3. Material revisado

### Documentos funcionales y conversación

| Fichero | Información relevante |
| --- | --- |
| `ficheros-auxiliares/historico-chat-gpt.txt` | Evolución completa del enfoque: de generación manual de un CORE, a proceso con assets, a arquitectura modular, y finalmente a un único submódulo importable. Los requisitos tardíos son los que gobiernan el diseño actual. |
| `ficheros-auxiliares/documento_funcional_core_ipf_uipath.pdf` | Funcional base: automatización de pre-relleno, PPO como fuente principal, SAP/Salesforce como fuentes complementarias, revisión humana obligatoria, reglas PC/AT, IPF parcial. |
| `ficheros-auxiliares/Reunion-explicacion-proceso.md` | Contexto de negocio explicado por Lluis/Jairo: PPO, CORE PC/AT, prorrateo plano, intercompany, datos SAP/SF, limitaciones de acceso, fase 1/fase 2 de cierre. |
| `ficheros-auxiliares/Informacion-ficheros-adjuntos.txt` | Listado de plantillas y ejemplos oficiales. Indica que existen variantes CORE PC 60/150, pero quedan fuera del alcance de este backlog. |

### Ficheros Excel relevantes

| Fichero | Uso |
| --- | --- |
| `ficheros-auxiliares/20250256445  PPO_NuevaTerminalPVR v2.xlsm` | PPO de ejemplo. Contiene hojas `Hoja de datos`, `Parametros`, `Presupuesto`, `Facturación y SAP`, `Sintesis Precio`, entre otras. |
| `ficheros-auxiliares/PMBox_plantilla CORE_PC_v1_04 (2).xlsx` | Plantilla oficial CORE PC. |
| `ficheros-auxiliares/PMBox_plantilla CORE_AT_v1_04.xlsx` | Plantilla oficial CORE AT. |
| `ficheros-auxiliares/ZEV-PCE0012-1001_CORE_GAP_NuevaTerminalPVR_202602_-v1 - EJEMPLOCOREFORMACION.xlsx` | CORE PC de referencia de negocio. |
| `ficheros-auxiliares/CORE_PC_NuevaTerminalPVR_prerrelleno.xlsx` | Resultado previo usado como referencia técnica, pero no como formato final porque contiene hoja de trazabilidad añadida. |
| `ficheros-auxiliares/Plantilla IPF.xlsx` y ejemplo IPF | Base para planificar un módulo IPF completo y robusto, separado del módulo CORE pero integrado en el mismo roadmap técnico. |

Datos leídos del PPO de ejemplo:

| Campo | Valor |
| --- | --- |
| Company | `Connectis-ICT` |
| Service type | `Fixed Price (services)` |
| PEP type | `PC` |
| Portfolio | `Smart Spaces` |
| Practice | `Software Development` |
| Opportunity type | `Expansion` |
| Opportunity / SLFC | `20250256445` |
| Inicio previsto | `11/2025` |
| Duración | `12` meses |
| Total oferta sin IVA | `583896.55` |

### Procesos UiPath de referencia

| Proceso | Relevancia para CORE |
| --- | --- |
| `sf-cierre-fase1-develop` | Proceso principal previo. Exporta informe Salesforce, genera Excel `Oportunidades_dd_MM_yyyy_HHmm.xlsx`, calcula rutas y crea items en cola `CierreOP`. |
| `sf-cierre-fase2-develop` | Complementario. Consume `CierreOP`, crea/usa carpetas WBS, descarga ficheros Salesforce y gestiona permisos. No extrae SAP estructurado. |
| `sf-apertura-fase1-feature-initial-import` | No es el proceso correcto para completar CORE de cierre. Sirve para apertura/generación inicial de artefactos de oportunidad. |
| `sf-forecast-fase1-develop` | No aporta datos finales de CORE. Es forecast comercial. |
| `core-from-ppo-uipath-modular-readable-invokecode` | Referencia de lógica ya desarrollada y modularizada, pero no es el entregable final porque el nuevo requisito exige un único XAML importable. |

## 4. Contrato funcional del submódulo

El proceso padre debe llamar a `oportunidad-generar-core.xaml` dentro de su bucle de oportunidades y pasar todos los datos ya preparados. El módulo no debe buscar en Orchestrator, no debe consultar Salesforce/SAP y no debe decidir carpetas de negocio.

Argumentos actuales del módulo:

| Argumento | Dirección | Uso |
| --- | --- | --- |
| `in_RutaPPO` | In | Ruta absoluta del PPO `.xlsm`. |
| `in_RutaPlantillaCorePC` | In | Ruta absoluta de plantilla CORE PC. |
| `in_RutaPlantillaCoreAT` | In | Ruta absoluta de plantilla CORE AT. |
| `in_RutaCORE` | In | Ruta absoluta final del CORE generado. |
| `in_TipoProyecto` | In | `PC` o `AT`. |
| `in_SAP_FechaInicio` | In | Fecha de inicio real si SAP/proceso previo la aporta. Si falta, se usa la fecha del PPO. |
| `in_SAP_FechaFin` | In | Fecha fin real si SAP/proceso previo la aporta. Si falta, se usa la fecha del PPO. |
| `in_SAP_CodigoPEP_WBS` | In | WBS/PEP. Si falta, escribir error corto en rojo en el CORE. |
| `in_SAP_CodigoCliente` | In | Código cliente/Sold-to. Si falta, escribir error corto en rojo en el CORE. |
| `in_SAP_PedidoSalesOrder` | In | Pedido/Sales Order. Si falta, escribir error corto en rojo en el CORE. |
| `in_SAP_CodigoEmpleadoJefeProyecto` | In | Código empleado PM. |
| `in_SAP_NombreJefeProyecto` | In | Nombre PM. |
| `in_SAP_CodigoEmpleadoPracticeLeader` | In | Código practice leader. |
| `in_SAP_NombrePracticeLeader` | In | Nombre practice leader. |
| `in_SF_CodigoEmpleadoSalesManager` | In | Código sales manager si llega de Salesforce/proceso previo. |
| `in_SF_NombreSalesManager` | In | Nombre sales manager. Si falta, se usa Account Manager del PPO. |
| `in_SF_AccountName` | In | Nombre de cuenta/cliente desde Salesforce. Si falta, se usa cliente del PPO. |
| `in_SF_Practica` | In | Práctica desde Salesforce. Si falta, se usa práctica del PPO. |
| `in_SF_RutaProyecto` | In | Trazabilidad para proceso padre. Actualmente no se usa de forma funcional en el CORE. |
| `in_SF_CorreoResponsable` | In | Trazabilidad para proceso padre. Actualmente no se usa de forma funcional en el CORE. |
| `in_SAP_ProveedorCompra` | In | Proveedor para compras/subcontrataciones. |
| `in_SAP_PedidoCompra` | In | PO compra. |
| `in_SAP_CodigoAriba` | In | Código Ariba. |
| `in_SAP_CodigosEmpleadoRecursosJson` | In | JSON opcional de sigla a código empleado de recurso. |
| `out_NumeroSFLeido` | Out | Opportunity Number leído del PPO. |

Riesgo detectado: este contrato tiene 25 argumentos y UiPath Workflow Analyzer marca `ST-DBP-002` porque supera el umbral de 20. Se confirma mantener argumentos individuales, porque el proceso padre debe mapear cada dato sin interpretar una estructura compuesta. No es obligatorio cumplir esta recomendación de forma estricta si empeora la claridad funcional, pero sí conviene usar prefijos `in_`, `out_` e `io_` y mantener nombres humanos en español. La prioridad será que el proceso padre entienda el contrato sin ambigüedad.

## 5. Reglas funcionales que deben preservarse

### 5.1 CORE PC

- Usar plantilla CORE PC oficial cuando `in_TipoProyecto = "PC"`.
- Rellenar `Project Infor` con cabecera y totales del PPO, completando SAP/Salesforce si el dato llega por argumento.
- Escribir recursos con una fila por perfil.
- Usar coste/hora del año de inicio o el primer coste disponible.
- Prorratear horas por anualidad entre los meses del proyecto.
- Prorratear gastos y compras cuando no exista distribución más fiable.
- Imputar riesgos y garantía en el último mes.
- Mantener facturación emitida/billing real vacío al generar el CORE inicial.

### 5.2 CORE AT

- Usar plantilla CORE AT oficial cuando `in_TipoProyecto = "AT"`.
- Recursos: dos filas por perfil, una para coste/hora y otra para tarifa de venta.
- Cost Planning: dos filas por perfil, horas reales y horas/jornadas facturables.
- Como base inicial, igualar horas facturables a horas reales prorrateadas.
- Intentar extraer tarifa desde `Sintesis Precio` por sigla, usando tarifa directa o `venta / horas`.
- No se exige PPO AT real para validar esta fase. La lógica AT debe inferirse de forma razonada desde la plantilla CORE AT, desde la estructura del PPO PC validado y desde los campos equivalentes del PPO, aplicando pruebas PC como base de integridad y revisando las diferencias AT por diseño.

### 5.3 Intercompany

El funcional y la transcripción son la fuente prioritaria para esta regla. La matriz definitiva operativa queda fijada a partir de lo explicado en la toma de requisitos: el intercompany depende primero de la compañía emisora inferida desde el WBS/código de proyecto y después de la entidad del recurso leída en el PPO, no de una comparación textual simple de campos sueltos.

Normalización de compañía del proyecto:

- WBS/código con prefijo `ZEV`, referencias `ZV`, `ICT`, `Connectis`, `Connectis-ICT` o equivalentes: compañía proyecto `ICT/Connectis`.
- WBS/código con prefijo `ZEZ`, referencias `EZZ`, `Global Rosetta`, `Rosetta` o equivalentes: compañía proyecto `Global Rosetta`.
- Si el WBS no permite inferir compañía, usar como apoyo los datos SAP/Salesforce aportados al módulo. Si aun así no se puede determinar, escribir error corto en rojo en la celda intercompany correspondiente: `Intercompany no determinable`.

Matriz de decisión:

| Compañía proyecto inferida | Entidad recurso normalizada | Valor intercompany |
| --- | --- | --- |
| `ICT/Connectis` | `ICT/Connectis` | No intercompany |
| `ICT/Connectis` | `Global Rosetta` | Intercompany |
| `Global Rosetta` | `Global Rosetta` | No intercompany |
| `Global Rosetta` | `ICT/Connectis` | Intercompany |
| No determinable | Cualquiera | Texto rojo `Intercompany no determinable` |
| Cualquiera | No determinable | Texto rojo `Entidad recurso no disponible` |

La implementación debe mantener esta matriz en una estructura visible y mantenible, preferiblemente una `DataTable` inicial o una secuencia claramente nombrada, para que negocio pueda revisarla sin abrir un bloque de código. Si se usa `Invoke Code`, solo puede encapsular la normalización o búsqueda de la matriz, nunca ocultar la regla completa.

### 5.4 Variantes de plantilla

Las plantillas CORE PC ampliadas de 60/150 líneas quedan fuera del alcance de este backlog. El módulo debe trabajar con la plantilla PC recibida por argumento y con la plantilla AT recibida por argumento.

El código no debe implementar selección automática de variantes ampliadas. Sí debe validar límites estructurales mínimos de la plantilla recibida:

- duración máxima: 25 meses;
- PC: capacidad real disponible en las filas útiles de `Resources` y `Cost Planning` de la plantilla recibida;
- AT: 15 perfiles por el modelo de dos filas.

Si el PPO supera la capacidad de la plantilla recibida, esto sí es un bloqueo técnico/estructural: no hay celda destino segura donde escribir. En ese caso el módulo debe lanzar un error claro indicando el número de perfiles detectado y la capacidad de la plantilla aportada.

### 5.5 Datos faltantes y escritura de errores en CORE

La ausencia de un dato funcional no debe detener la generación del CORE. La regla de diseño es escribir el dato correcto cuando exista y, cuando falte o no sea válido, escribir en la misma celda destino un texto corto de error con fuente roja.

Criterios:

- La lógica OK/KO debe estar agrupada cerca de la actividad que propaga cada dato al CORE. No crear una validación funcional previa que duplique el mapeo y luego otra secuencia separada para escribir.
- Formato recomendado del error: `[dato] no disponible`, `[dato] no disponible en PPO`, `[dato] no disponible en SAP` o `[dato] no disponible en Salesforce`, escogiendo el texto más corto que permita actuar.
- Aplicar color rojo solo a la celda afectada, sin añadir hojas de auditoría, comentarios extensos ni formatos invasivos.
- Mantener logs breves de advertencia cuando falten datos relevantes, pero la evidencia principal debe quedar en el CORE generado.
- Se permite crear una utilidad reutilizable, por ejemplo `Escribir dato o error en celda`, siempre que el uso visual siga mostrando qué dato se está propagando y a qué celda.
- Solo deben abortar condiciones que impiden generar un Excel válido: PPO inexistente, plantilla inexistente, ruta de salida no escribible, tipo PC/AT inválido, hoja obligatoria ausente, plantilla sin capacidad suficiente o libro corrupto.

## 6. Mapeo de datos principales

### PPO -> modelo interno

| Origen PPO | Dato |
| --- | --- |
| `Hoja de datos!B8` | Compañía |
| `Hoja de datos!B9` | Unidad |
| `Hoja de datos!B10` | Localización |
| `Hoja de datos!B11` | Tipo servicio |
| `Hoja de datos!B12` | Tipo PEP / PC-AT |
| `Hoja de datos!B14:B16` | Portfolio, práctica, subpráctica |
| `Hoja de datos!B17` | Tipo oportunidad |
| `Hoja de datos!B18:B20` | Comercial/solicitante/origen |
| `Hoja de datos!B22` | Opportunity Number |
| `Hoja de datos!B23` | Cliente |
| `Hoja de datos!B25` | Descripción |
| `Hoja de datos!B27:B29` | Mes inicio, año inicio, duración |
| `Facturación y SAP!C4/E4` | Fechas si están informadas |
| `Presupuesto!X8` o `Sintesis Precio!D12` | Importe pedido/oferta |
| `Presupuesto!K4/K5/K28/K34/K49/K53` | Coste recursos, horas, gastos, compras, riesgos, garantía |
| `Parametros!B:D` y costes por año | Perfil, entidad, sigla, coste/hora |
| `Presupuesto!B:D` y columnas anuales | Perfiles y horas por anualidad |
| `Sintesis Precio!D/N/O/P` | Tarifas AT por sigla |

### Datos externos -> CORE

| Dato destino | Argumento / fallback |
| --- | --- |
| WBS/PEP | `in_SAP_CodigoPEP_WBS` o texto rojo `WBS no disponible` |
| Código cliente | `in_SAP_CodigoCliente` o texto rojo `Cliente SAP no disponible` |
| Pedido/Sales Order | `in_SAP_PedidoSalesOrder` o texto rojo `Sales Order no disponible` |
| PM código/nombre | `in_SAP_CodigoEmpleadoJefeProyecto`, `in_SAP_NombreJefeProyecto` o texto rojo acotado |
| Practice leader código/nombre | `in_SAP_CodigoEmpleadoPracticeLeader`, `in_SAP_NombrePracticeLeader` o texto rojo acotado |
| Sales manager código/nombre | `in_SF_CodigoEmpleadoSalesManager`, `in_SF_NombreSalesManager` o PPO |
| Cliente/cuenta | `in_SF_AccountName` o PPO |
| Práctica | `in_SF_Practica` o PPO |
| Compras | `in_SAP_ProveedorCompra`, `in_SAP_PedidoCompra`, `in_SAP_CodigoAriba` o texto rojo acotado |

## 7. Estado técnico actual de `oportunidad-generar-core.xaml`

Validaciones realizadas:

| Validación | Resultado |
| --- | --- |
| XML/XAML bien formado | Correcto. |
| `Copy File` | Configurado con `Path="[RutaPlantillaSeleccionada]"`, `Destination="[rutaCORETrabajo]"`, `Overwrite="True"`, alineado con `Workflow-ejemplo.xaml`. |
| UiPath CLI instalada | Existe `C:\Program Files\UiPath\Studio\UiPath.Studio.CommandLine.exe`, versión `23.10.4`. |
| UiRobot instalado | Existe `C:\Program Files\UiPath\Studio\UiRobot.exe`, versión `23.10.4`. |
| Excel COM | Disponible en la máquina. |
| `analyze-file` sobre el XAML | Ejecuta, pero devuelve errores de Workflow Analyzer. |
| Ejecución directa de XAML con `UiRobot execute` | No es posible: UiRobot solo ejecuta procesos publicados o paquetes `.nupkg` .NET6, no XAML crudo. |

Bloques `Invoke Code` actuales:

| Bloque | Líneas aproximadas | Diagnóstico |
| --- | ---: | --- |
| Leer PPO y construir modelo de negocio | 637 | Sigue como bloque transitorio. Debe dividirse en S04: lectura por actividades Excel + pequeñas utilidades de normalización. |
| Rellenar CORE desde PPO y datos SAP/Salesforce | 770 | Sigue como bloque transitorio. Debe dividirse en S05-S07 por hoja y por rama PC/AT usando actividades Excel visibles. |

Tras S03, la validación del contrato de entrada y la selección/copia de plantilla ya se realizan con actividades visibles (`If`, `Throw`, `Assign`, `File Exists`, `Folder Exists`, `Copy File`). El primer `Invoke Code` queda eliminado.

Resultado de Workflow Analyzer sobre el XAML:

- `ST-DBP-002`: el workflow tiene más de 20 argumentos. Actualmente son 25.
- `ST-NMG-002`: los argumentos no siguen convención `in_` / `out_`.
- `ST-NMG-016`: nombres de algunos argumentos superan 30 caracteres.
- `ST-ANA-009`: información de recuento de actividades.

Resultado de análisis de proyecto completo:

- Falla porque `ficheros-auxiliares` está dentro del proyecto y `project.json` no la ignora.
- El analizador intenta cargar XAML de referencia con actividades PowerPoint/Word no presentes en las dependencias actuales.
- También arrastra advertencias/errores de procesos ajenos al módulo CORE.

Conclusión: como el entregable será el XAML final, el análisis de proyecto completo y el empaquetado quedan fuera del alcance ordinario. La validación relevante debe centrarse en el XAML, su apertura en Studio 23.10.4, la ejecución desde un wrapper local de prueba y la integridad del Excel generado.

### 7.1 Criterios de diseño visual y low-code para el módulo

El rediseño debe tratar el XAML como un artefacto que se revisa visualmente en Studio, no como un contenedor de C#. La prioridad es que un compañero pueda abrir el flujo y entender el negocio por la forma del diagrama.

Criterios obligatorios:

- La primera secuencia funcional del módulo debe llamarse `00 Preparar datos locales para pruebas y ejecución`.
- Dentro de esa secuencia debe existir un `Multiple Assign` inicial, por ejemplo `Inicializar variables locales desde argumentos`.
- Ese `Multiple Assign` copiará todos los argumentos `in_` a variables locales con nombres de negocio: `rutaPPOTrabajo`, `rutaCORETrabajo`, `tipoProyectoTrabajo`, `codigoPEPTrabajo`, etc.
- Para pruebas unitarias manuales, el desarrollador podrá sustituir temporalmente el valor asignado desde argumento por un valor hardcodeado sin tocar el resto del flujo.
- Después del `Multiple Assign`, el módulo debe trabajar siempre con variables locales, no directamente con argumentos.
- Las secciones principales deben estar numeradas y nombradas en español: `01 Validar entradas`, `02 Copiar plantilla`, `03 Leer PPO`, `04 Construir modelo`, `05 Rellenar Project Infor`, `06 Rellenar Resources`, `07 Rellenar Cost Planning`, `08 Recalcular y guardar`.
- Cada secuencia debe tener anotación funcional corta: qué hace, qué recibe y qué deja preparado.
- La bifurcación PC/AT debe ser visible mediante `If` o `Switch`, no escondida dentro de C#.
- Los bucles de perfiles, gastos y compras deben verse como `For Each Row` / `For Each` en Studio siempre que sea viable.
- Los mapeos de celdas fijas deben expresarse con `Write Cell` o secuencias equivalentes, agrupadas por bloque de negocio.
- Los `Invoke Code` se reservan para funciones puras y pequeñas: parsear número/fecha, calcular distribución mensual, resolver columna Excel, normalizar texto o evaluar una regla puntual.
- Ningún `Invoke Code` debe abrir libros, guardar ficheros, copiar plantillas, recorrer hojas completas o escribir grandes rangos si existe una actividad UiPath razonable para ello.
- Si queda un `Invoke Code`, su título debe explicar la intención de negocio y su código debe estar formateado con bloques, comentarios y menos de una pantalla de Studio siempre que sea posible.

Estándar de nombres, anotaciones y comentarios:

- `DisplayName` de secuencias: `NN Verbo + objeto + criterio`, por ejemplo `05 Rellenar Project Infor - identificadores y responsables`.
- `DisplayName` de actividades de escritura: `Escribir [dato de negocio] en [hoja!celda/rango]`, por ejemplo `Escribir WBS en Project Infor C16`.
- Anotación de secuencia, máximo 3 líneas: `Objetivo: ...`, `Entradas: ...`, `Salida: ...`.
- Comentarios de actividad solo cuando expliquen una decisión de negocio o una excepción técnica. No repetir lo que ya dice el nombre de la actividad.
- Para `Invoke Code`, usar cabecera fija en comentarios: `Objetivo`, `Entradas`, `Salidas`, `Regla de negocio`, `No hace`. La sección `No hace` debe dejar claro que no abre Excel ni escribe el CORE si es una utilidad pura.
- Las anotaciones deben ser estables y revisables por negocio. Evitar notas temporales, dudas personales o referencias a pruebas locales que no formen parte del diseño final.

Diseño visual recomendado:

```text
oportunidad-generar-core.xaml
└── Control general del submódulo CORE
    └── Generar CORE de oportunidad
        ├── 00 Preparar datos locales para pruebas y ejecución
        │   └── Multiple Assign - Inicializar variables locales desde argumentos
        ├── 01 Validar entradas del módulo
        ├── 02 Seleccionar y copiar plantilla CORE
        ├── 03 Leer PPO con actividades Excel
        ├── 04 Construir modelo de negocio en DataTables/variables
        ├── 05 Rellenar Project Infor
        ├── 06 Rellenar Resources
        │   ├── Si CORE PC
        │   └── Si CORE AT
        ├── 07 Rellenar Cost Planning
        │   ├── Si CORE PC
        │   └── Si CORE AT
        └── 08 Recalcular, guardar y registrar fin
```

Esta estructura permite alternar entre ejecución real y prueba unitaria local cambiando solo el `Multiple Assign` inicial. Es una concesión muy práctica: el contrato con el proceso padre queda estable, pero el desarrollador puede aislar el módulo sin montar todo el flujo previo.

## 8. Validación por Studio, CMD o PowerShell

Sí es posible validar parcialmente desde CMD/PowerShell con la instalación local de UiPath 23.10.4, pero no se puede ejecutar directamente el XAML suelto con `UiRobot`. Como el entregable será el XAML para integrarlo en otro proyecto, no forman parte del backlog las tareas de empaquetado ni publicación.

Comandos útiles:

```powershell
& "C:\Program Files\UiPath\Studio\UiPath.Studio.CommandLine.exe" --version
& "C:\Program Files\UiPath\Studio\UiPath.Studio.CommandLine.exe" analyze-file --workflow-file-path "C:\ruta\oportunidad-generar-core.xaml"
```

Plan de validación recomendado:

1. Validación XML: cargar el XAML como XML.
2. Workflow Analyzer por archivo: debe quedar sin errores bloqueantes. Las recomendaciones de estilo pueden documentarse si chocan con claridad funcional.
3. Apertura visual en UiPath Studio 23.10.4 para comprobar que las actividades se cargan con los paquetes del `project.json`.
4. Ejecución desde un wrapper local de prueba mediante `Invoke Workflow File`, cuando se necesite validar runtime.
5. Comparación de Excel generado contra un set de celdas esperado.

Para ejecución runtime se necesita un wrapper de test o un proceso padre que invoque el XAML con `Invoke Workflow File`. No se contempla publicar ni empaquetar como parte de este backlog.

## 9. Fuentes útiles de internet

Fuentes oficiales UiPath revisadas o útiles para esta tarea:

- Studio executables 2023.10: `UiPath.Studio.CommandLine.exe` y `UiRobot.exe` están documentados como ejecutables del entorno Studio/Robot.  
  https://docs.uipath.com/studio/standalone/2023.10/user-guide/studio-executables
- Workflow Analyzer: comandos `analyze` y `analyze-file`, configuración de reglas y salida JSON.  
  https://docs.uipath.com/studio/standalone/2023.10/user-guide/about-workflow-analyzer
- Invoke Code: actividad para ejecutar C#/VB.NET con argumentos; relevante para limitar su uso a piezas acotadas.  
  https://docs.uipath.com/activities/other/latest/workflow/invoke-code
- Argumentos en Studio 2023.10: uso con `Invoke Workflow File` y buenas prácticas de dirección `In`/`Out`.  
  https://docs.uipath.com/studio/standalone/2023.10/user-guide/using-arguments  
  https://docs.uipath.com/studio/standalone/2023.10/user-guide/managing-arguments
- Automatización Excel en StudioX/Studio 2023.10: referencia conceptual para mover lectura/escritura a actividades Excel visibles.  
  https://docs.uipath.com/studiox/standalone/2023.10/user-guide/excel-automation
- Actividad `Path Exists`: útil para reemplazar validaciones de rutas hoy metidas en `Invoke Code`.  
  https://docs.uipath.com/activities/other/latest/workflow/path-exists
- Automatización de ficheros: referencia de actividades como `Copy File`, `File Exists`, `Folder Exists`, etc.  
  https://docs.uipath.com/studiox/standalone/2023.10/user-guide/file-automation

## 10. Backlog ordenado

### E00 - Preparación de validación del XAML

Objetivo: preparar una base de prueba útil para el XAML final sin invertir esfuerzo en empaquetado, publicación ni configuración del proyecto del compañero.

1. `E00-T01` Crear workflow wrapper de test local.
   - Nombre recomendado: `test\test_generar_core_pc.xaml`.
   - Debe invocar `oportunidad-generar-core.xaml` con argumentos de ejemplo.
   - Debe usar rutas temporales controladas bajo `.local\test-output`.
   - Criterio de aceptación: ejecución local reproducible sin tocar ficheros fuente.

2. `E00-T02` Definir baseline de validación CORE PC.
   - Usar PPO `20250256445`.
   - Usar plantilla CORE PC oficial.
   - Crear lista de celdas esperadas en `Project Infor`, `Resources`, `Cost Planning`.
   - Criterio de aceptación: se puede comparar output vs baseline.

3. `E00-T03` Validar integridad técnica básica del XAML.
   - XML válido.
   - `analyze-file` orientativo sobre `oportunidad-generar-core.xaml`.
   - Apertura visual en Studio 23.10.4.
   - Criterio de aceptación: sin errores técnicos bloqueantes en el XAML. Las reglas de estilo de Workflow Analyzer pueden quedar documentadas si chocan con legibilidad humana.

### E01 - Contrato público del submódulo y modo de prueba

Objetivo: estabilizar cómo el proceso padre consume el módulo y permitir pruebas unitarias cómodas.

1. `E01-T01` Renombrar argumentos con prefijos `in_`, `out_`, `io_`.
   - Ejemplos: `in_RutaPPO`, `in_RutaCORE`, `in_TipoProyecto`, `out_NumeroSFLeido`.
   - Mantener nomenclatura humana en español.
   - No perseguir ciegamente nombres cortos si eso resta claridad.
   - Criterio de aceptación: cualquier desarrollador entiende el contrato de entrada al importar argumentos.

2. `E01-T02` Crear `Multiple Assign` inicial obligatorio.
   - DisplayName recomendado: `Inicializar variables locales desde argumentos`.
   - Secuencia contenedora: `00 Preparar datos locales para pruebas y ejecución`.
   - Copiar cada argumento `in_` a una variable local de trabajo.
   - Ejemplos: `in_RutaPPO` -> `rutaPPOTrabajo`, `in_TipoProyecto` -> `tipoProyectoTrabajo`.
   - Criterio de aceptación: para probar unitariamente basta con cambiar valores en este `Multiple Assign`, sin tocar el resto del flujo.

3. `E01-T03` Sustituir referencias directas a argumentos por variables locales.
   - El resto del módulo debe trabajar con `rutaPPOTrabajo`, `rutaCORETrabajo`, `codigoPEPTrabajo`, etc.
   - Los argumentos solo se usan en el `Multiple Assign` inicial y en outputs finales.
   - Criterio de aceptación: el módulo queda aislable y testeable visualmente.

4. `E01-T04` Consolidar contrato con argumentos individuales.
   - Mantener datos SAP/Salesforce como argumentos `in_` individuales.
   - Documentar `ST-DBP-002` como excepción aceptada por claridad de integración.
   - No agrupar datos externos en JSON/DataTable en el contrato público del módulo.
   - Criterio de aceptación: el compañero puede mapear cada dato desde el proceso padre sin interpretar estructuras compuestas.

5. `E01-T05` Definir contrato de errores.
   - Error bloqueante: PPO no existe, plantilla no existe, tipo proyecto inválido, carpeta destino no existe, hoja obligatoria ausente, libro corrupto, Excel no disponible o plantilla sin capacidad.
   - Dato funcional faltante: no aborta; se escribe texto corto en rojo en la celda correspondiente.
   - Criterio de aceptación: cada error bloqueante tiene mensaje accionable y cada ausencia funcional queda visible dentro del CORE.

### E02 - Diseño visual UiPath y minimización avanzada de `Invoke Code`

Objetivo: que el flujo sea legible en Studio y que el código quede como apoyo puntual, no como núcleo monolítico.

1. `E02-T01` Reemplazar `Validar contrato de entrada` por actividades.
   - Usar `Path Exists` / `File Exists` / `Folder Exists`.
   - Usar `If` + `Throw` para errores.
   - Usar `Assign` para normalizar `tipoProyectoTrabajo`.
   - Criterio de aceptación: primer `Invoke Code` eliminado.

2. `E02-T02` Diseñar el esqueleto visual por secuencias numeradas.
   - `01 Validar entradas del módulo`.
   - `02 Seleccionar y copiar plantilla CORE`.
   - `03 Leer PPO con actividades Excel`.
   - `04 Construir modelo de negocio`.
   - `05 Rellenar Project Infor`.
   - `06 Rellenar Resources`.
   - `07 Rellenar Cost Planning`.
   - `08 Recalcular y guardar`.
   - Criterio de aceptación: el flujo se entiende sin abrir C#.

3. `E02-T03` Leer PPO con actividades Excel.
   - `Use Excel File` / `Excel Process Scope` según lo que abra mejor `.xlsm` y ficheros etiquetados.
   - `Read Cell` para celdas fijas de cabecera.
   - `Read Range` para `Parametros`, `Presupuesto`, `Facturación y SAP`, `Sintesis Precio`.
   - Criterio de aceptación: el segundo `Invoke Code` deja de abrir Excel por COM.

4. `E02-T04` Construir DataTables de negocio.
   - `dtCabeceraPPO`.
   - `dtPerfilesPPO`.
   - `dtCostesPorAnio`.
   - `dtHorasPorAnio`.
   - `dtGastos`.
   - `dtCompras`.
   - `dtTarifasAT`.
   - Criterio de aceptación: los bucles de negocio se ven como `For Each Row`.

5. `E02-T05` Convertir mapeos fijos a actividades de escritura.
   - `Project Infor` debe rellenarse con `Write Cell` o secuencias de `Write Cell`.
   - Cada bloque debe tener título: fechas, identificadores, cliente, responsables, económicos.
   - En cada dato, resolver valor OK o texto de error rojo justo antes de escribir la celda.
   - Criterio de aceptación: un cambio de celda se hace en Studio sin tocar C#.

6. `E02-T06` Separar PC/AT con `Switch` o `If`.
   - Rama `CORE PC`.
   - Rama `CORE AT`.
   - Dentro de cada rama, `Resources` y `Cost Planning` deben tener subsecuencias propias.
   - Criterio de aceptación: la diferencia funcional PC/AT es visual.

7. `E02-T07` Encapsular solo utilidades pequeñas en `Invoke Code`.
   - Permitido: `ParseNumero`, `ParseFecha`, `ColumnaExcel`, `DistribucionMensual`, normalización de texto.
   - Permitido con cautela: cálculo de intercompany si se alimenta desde tabla externa.
   - Prohibido salvo bloqueo técnico: abrir libros, copiar ficheros, escribir hojas completas, generar outputs finales.
   - Criterio de aceptación: ningún `Invoke Code` concentra más de una responsabilidad.

8. `E02-T08` Añadir anotaciones visuales útiles.
   - Aplicar el estándar `Objetivo / Entradas / Salida` en secuencias principales.
   - Aplicar nombres de actividad orientados a negocio y celda destino.
   - Si queda `Invoke Code`, añadir cabecera `Objetivo / Entradas / Salidas / Regla de negocio / No hace`.
   - Cada secuencia debe explicar qué deja preparado para la siguiente.
   - Evitar comentarios decorativos.
   - Criterio de aceptación: una revisión por pares puede seguir el flujo como documento de negocio.

### E03 - Robustez funcional CORE PC/AT

Objetivo: asegurar que el CORE generado es correcto para casos reales dentro del alcance confirmado: CORE PC con plantilla recibida, CORE AT inferido de forma razonada y sin implementación de plantillas PC ampliadas 60/150.

1. `E03-T01` Validar mapa de celdas contra funcional y ejemplo.
   - Confirmar celdas `Project Infor`.
   - Confirmar rangos `Resources`.
   - Confirmar rangos `Cost Planning`.
   - Criterio de aceptación: mapeo aprobado o documentado con discrepancias.

2. `E03-T02` Acotar capacidad de la plantilla CORE recibida.
   - Dejar fuera de alcance la selección PC 60/150.
   - Detectar filas útiles de la plantilla recibida en `Resources` y `Cost Planning`.
   - Contar perfiles reales del PPO, no filas vacías.
   - Si el PPO supera la capacidad real de la plantilla recibida, lanzar error estructural claro.
   - Criterio de aceptación: el módulo nunca escribe fuera del área segura de la plantilla aportada.

3. `E03-T03` Inferir y robustecer PPO AT de forma científica.
   - Detectar AT por `in_TipoProyecto` del proceso padre y contrastar con `Hoja de datos!B12/B17`.
   - Leer tarifas desde `Sintesis Precio` por sigla, horas, tarifa directa y venta.
   - Si no hay tarifa directa, calcular `venta / horas`.
   - Si hay varias filas por sigla, calcular media ponderada por horas.
   - No exigir PPO AT real; inferir equivalencias desde PPO PC validado y plantilla AT.
   - Criterio de aceptación: la rama AT no depende de posiciones mágicas no documentadas y sus diferencias con PC están justificadas.

4. `E03-T04` Implementar regla intercompany según funcional y transcripción.
   - Prioridad 1: reglas del PDF funcional.
   - Prioridad 2: reglas explicadas en `Reunion-explicacion-proceso.md`.
   - Usar WBS/código proyecto, compañía emisora y entidad del recurso.
   - Implementar la matriz definitiva `ICT/Connectis` vs `Global Rosetta` indicada en la sección 5.3.
   - Criterio de aceptación: combinaciones `ZEV`/`ZEZ`, ICT/Connectis y Global Rosetta producen `Intercompany`, `No intercompany` o texto rojo acotado si falta información.

5. `E03-T05` Definir política de costes interanuales.
   - Versión inicial: coste del año de inicio, tal como se acordó pragmáticamente.
   - Evolución preparada: coste por anualidad cuando negocio lo active.
   - Criterio de aceptación: comportamiento explícito y testeado.

6. `E03-T06` Implementar propagación OK/KO por celda con formato rojo.
   - Para cada dato escrito en CORE, resolver valor válido o mensaje corto de error en el mismo bloque visual.
   - Aplicar fuente roja cuando se escriba un error funcional.
   - Evitar validaciones funcionales separadas que dupliquen la lógica de escritura.
   - Criterio de aceptación: el proceso no falla por falta de datos SAP/Salesforce/PPO no críticos y el CORE muestra exactamente dónde falta cada dato.

### E04 - Pruebas de integridad del Excel CORE

Objetivo: comprobar que el CORE resultante se abre, conserva estructura y contiene datos correctos.

1. `E04-T01` Validar estructura de hojas.
   - Deben existir `Project Infor`, `Cost Overview`, `Resources`, `Cost Planning`, `Monthly View`, `Cost Summary`.
   - No debe añadirse hoja `Trazabilidad_RPA`.
   - Criterio de aceptación: hoja a hoja OK.

2. `E04-T02` Validar conservación de formato.
   - Comparar número de hojas.
   - Comparar dimensiones básicas.
   - Verificar cuadros `NOTA`.
   - Criterio de aceptación: la plantilla conserva formato y fórmulas.

3. `E04-T03` Validar celdas críticas CORE PC.
   - `C8`, `C9`, `C10`, `C11`, `C12`, `C16`, `D16`, `C18`, `C19`.
   - `C28:C34`.
   - Primeras filas de `Resources`.
   - Primeras filas de `Cost Planning`.
   - Criterio de aceptación: valores esperados para el PPO de ejemplo.

4. `E04-T04` Validar recálculo.
   - Abrir con Excel en sesión robot.
   - Forzar cálculo si se usan fórmulas dependientes.
   - Guardar y reabrir.
   - Criterio de aceptación: sin errores de fórmula visibles y totales coherentes.

5. `E04-T05` Validar casos negativos.
   - PPO inexistente.
   - Plantilla inexistente.
   - Carpeta destino inexistente.
   - Tipo proyecto distinto de PC/AT.
   - Hoja obligatoria ausente.
   - Duración, WBS, cliente o responsables no informados.
   - Criterio de aceptación: los errores estructurales abortan con mensaje claro; los datos funcionales faltantes no abortan y aparecen en rojo en el CORE.

### E05 - Integración con proceso padre

Objetivo: que el compañero pueda invocar el módulo sin explicación adicional.

1. `E05-T01` Crear ejemplo de `Invoke Workflow File`.
   - Importar argumentos.
   - Pasar variables de una fila de oportunidad.
   - Recoger `out_NumeroSFLeido`.
   - Criterio de aceptación: workflow de ejemplo ejecutable.

2. `E05-T02` Documentar contrato de datos Salesforce/SAP.
   - Qué dato debe traer fase 1.
   - Qué dato debe traer SAP u otro proceso.
   - Qué dato puede quedar vacío.
   - Criterio de aceptación: el proceso padre sabe exactamente qué construir.

3. `E05-T03` Alinear logs.
   - Inicio con PPO/tipo/ruta.
   - Fin OK con `in_RutaCORE` y `out_NumeroSFLeido`.
   - Error con PPO/CORE/detalle.
   - Criterio de aceptación: trazabilidad suficiente en logs locales o en el entorno donde el proceso padre ejecute el módulo.

### E06 - Documentación de integración del XAML

Objetivo: dejar el XAML entendible y transferible para que un compañero lo incorpore a su propio proyecto, sin preparar paquete, publicación ni despliegue local específico.

1. `E06-T01` Preparar guía técnica mínima.
   - Propósito del módulo.
   - Argumentos `in_` y `out_`.
   - Ejemplo de invocación con `Invoke Workflow File`.
   - Limitaciones conocidas.
   - Criterio de aceptación: el compañero puede integrar el XAML sin reconstruir el contexto histórico.

2. `E06-T02` Documentar supuestos de integración.
   - Studio 23.10.4.
   - Paquetes indicados en `project.json`.
   - Plantillas oficiales aportadas por argumento.
   - `in_RutaCORE` ya contiene el nombre y ruta final definidos por el proceso padre.
   - Criterio de aceptación: queda claro qué depende del proceso padre y qué resuelve el submódulo.

3. `E06-T03` Registrar decisiones funcionales por iteración.
   - Mantener este documento como referencia viva.
   - Registrar cambios de alcance y descubrimientos relevantes.
   - No preparar paquete salvo que una tarea futura lo pida expresamente.

### E07 - Generación IPF completa y robusta

Objetivo: desarrollar la parte IPF con el mismo rigor que CORE PC/AT, pero en un módulo separado para no contaminar `oportunidad-generar-core.xaml`. Su alcance debe ceñirse a lo confirmado en el funcional y en la conversación de toma de requerimientos; no debe inventar narrativa, hitos ni datos comerciales no aportados.

1. `E07-T01` Analizar plantilla IPF y ejemplo real.
   - Revisar `Plantilla IPF.xlsx`.
   - Revisar `IPF - ICT - GAP - ZEV-PCE0012 - Despliegue Smart Airport nueva terminal-  FEB26 v1.xlsx`.
   - Identificar hojas, celdas obligatorias, listas validadas y fórmulas.
   - Criterio de aceptación: mapa de celdas IPF documentado.

2. `E07-T02` Definir contrato de `oportunidad-generar-ipf.xaml`.
   - `in_RutaPlantillaIPF`.
   - `in_RutaIPF`.
   - Argumentos individuales `in_` equivalentes a los datos IPF confirmados en funcional y conversación.
   - `out_NumeroSFLeido` o salida equivalente si aplica.
   - Criterio de aceptación: contrato compatible con el mismo proceso padre que genera CORE.

3. `E07-T03` Implementar `Multiple Assign` inicial para IPF.
   - Mismo patrón que CORE: argumentos a variables locales.
   - Facilitar pruebas unitarias con hardcodes temporales.
   - Criterio de aceptación: IPF se puede probar aislado.

4. `E07-T04` Rellenar cabecera IPF con actividades Excel.
   - Company Issuer.
   - Customer / Sold-to number.
   - WBS / código proyecto.
   - Jefe de proyecto.
   - Moneda, por defecto EUR salvo dato externo.
   - Sales Order / pedido.
   - Criterio de aceptación: cabecera completa con datos SAP/SF disponibles.

5. `E07-T05` Implementar narrativa y facturación de forma robusta.
   - Escribir únicamente campos narrativos o de facturación confirmados como disponibles por el funcional y la toma de requisitos.
   - Si el proceso padre no aporta un dato confirmado, dejar el campo vacío o escribir texto corto en rojo según corresponda a la plantilla.
   - No generar `Summary Text`, `Text Narrative`, hitos ni condiciones contractuales por inferencia.
   - Criterio de aceptación: el módulo respeta la restricción funcional de no inventar narrativa ni hitos.

6. `E07-T06` Validar IPF contra listas y formatos.
   - Comprobar listas desplegables relevantes.
   - Validar moneda.
   - Validar formato de fechas/importes.
   - Criterio de aceptación: el IPF se abre sin errores y queda listo para revisión.

7. `E07-T07` Crear tests IPF.
   - Baseline con plantilla vacía.
   - Baseline con ejemplo GAP.
   - Casos negativos: plantilla ausente, datos SAP incompletos, narrativa no aportada.
   - Criterio de aceptación: errores estructurales controlados y datos funcionales faltantes tratados con campo vacío o texto rojo, sin abortar.

8. `E07-T08` Integrar generación CORE + IPF en un wrapper de ejemplo.
   - El wrapper debe llamar primero a CORE y después a IPF si el caso lo requiere.
   - CORE e IPF no deben generar auditoría externa.
   - Criterio de aceptación: ambos documentos se generan en rutas controladas con logs claros.

## 11. Orden sugerido de ejecución por sprints

Cada sprint equivale a una tarea futura para Codex con GPT-5.5 y reasoning effort extra high. El objetivo es que cada sprint pueda completarse, validarse y dejar el repositorio en un estado mejor que el anterior.

| Sprint | Alcance recomendado | Épicas/tareas incluidas | Resultado esperado |
| --- | --- | --- | --- |
| `S01` | Preparar validación del XAML | `E00-T01`, `E00-T02`, `E00-T03` | Wrapper de test PC, baseline y validación técnica básica del XAML. |
| `S02` | Estabilizar contrato CORE y modo unitario | `E01-T01`, `E01-T02`, `E01-T03`, `E01-T05` | Argumentos con prefijos, `Multiple Assign` inicial y flujo testeable con hardcodes. |
| `S03` | Rediseñar esqueleto visual y eliminar validación en C# | `E02-T01`, `E02-T02`, `E02-T08` | Diagrama principal claro, numerado y entendible en Studio. |
| `S04` | Mover lectura PPO a actividades | `E02-T03`, `E02-T04` | PPO leído mediante actividades Excel y DataTables de negocio. |
| `S05` | Mover `Project Infor` y mapeos fijos a actividades | `E02-T05`, parte de `E03-T01` | Cabecera CORE mantenible visualmente con `Write Cell`. |
| `S06` | Implementar `Resources` visual PC/AT | `E02-T06`, parte de `E03-T02`, parte de `E03-T03` | Recursos PC/AT con ramas visibles y menor dependencia de C#. |
| `S07` | Implementar `Cost Planning` visual PC/AT | `E02-T07`, `E03-T05`, parte de `E03-T03` | Prorrateos controlados y planning visible por tipo de CORE. |
| `S08` | Robustecer intercompany y datos faltantes | `E03-T02`, `E03-T04`, `E03-T06` | Capacidad de plantilla recibida controlada, regla intercompany trazable y errores funcionales en rojo. |
| `S09` | Validación CORE end-to-end | `E04-T01`, `E04-T02`, `E04-T03`, `E04-T04`, `E04-T05` | CORE PC baseline validado e integridad Excel cubierta. |
| `S10` | Integración con proceso padre y handover del XAML | `E05-T01`, `E05-T02`, `E05-T03`, `E06-T01`, `E06-T02`, `E06-T03` | Ejemplo de invocación, documentación mínima y XAML listo para integrar. |
| `S11` | Análisis y contrato IPF | `E07-T01`, `E07-T02`, `E07-T03` | Mapa IPF, contrato del módulo y modo de prueba unitario. |
| `S12` | Implementación IPF base completa | `E07-T04`, `E07-T05`, `E07-T06` | IPF generado con cabecera, narrativa/datos estructurados si se aportan y validaciones de formato. |
| `S13` | Pruebas e integración CORE + IPF | `E07-T07`, `E07-T08` | Wrapper CORE+IPF, pruebas reproducibles y documentación final. |

## 12. Backlog vivo y actualización continua

Este backlog no debe tratarse como un contrato congelado. Cada sprint debe empezar revisando si los descubrimientos del sprint anterior cambian el orden, el alcance o la granularidad de las tareas restantes.

Reglas de mantenimiento:

- Mantener los identificadores `E##-T##` estables cuando una tarea conserve su intención.
- Si una tarea cambia de alcance, actualizar su descripción y dejar constancia en avances.
- Si aparece una tarea nueva, añadirla a la épica correspondiente o crear una épica nueva si cambia el mapa de trabajo.
- Si una tarea deja de aportar valor, marcarla como descartada explicando el motivo.
- Después de cada sprint, actualizar el orden sugerido de ejecución si hay una ruta mejor hacia el XAML final.
- Priorizar siempre el resultado funcional y la mantenibilidad visual en Studio sobre el cumplimiento mecánico del plan inicial.

## 13. Información de avances realizados en la ejecución del backlog

Codex debe actualizar esta sección al finalizar cada sprint para que las iteraciones posteriores no dependan de memoria conversacional. La actualización debe incluir qué se tocó, qué se validó, qué decisiones nuevas se tomaron y qué queda preparado para el siguiente sprint.

Formato recomendado:

| Sprint | Fecha | Tareas | Acciones realizadas | Decisiones / descubrimientos | Validaciones | Siguiente paso |
| --- | --- | --- | --- | --- | --- | --- |
| Pendiente | 2026-04-29 | Contexto inicial | Documento refinado con decisiones cerradas, matriz intercompany, política de datos faltantes, exclusión de PC ampliadas y handover por XAML. | El backlog queda vivo y se actualizará por sprint. | Revisión documental. | Iniciar `S01` cuando se empiece la ejecución técnica. |
| S01 | 2026-04-29 | `E00-T01`, `E00-T02`, `E00-T03` | Creado wrapper local `test/test_generar_core_pc.xaml` para invocar `oportunidad-generar-core.xaml` con PPO `20250256445`, plantillas oficiales y salida controlada en `.local/test-output`. Creado baseline `test/baseline_core_pc.csv` con 56 celdas esperadas de `Project Infor`, `Resources` y `Cost Planning`. Creado comparador `test/validar_baseline_core_pc.ps1` y guía `test/README.md`. | El baseline refleja el comportamiento técnico actual del XAML, incluyendo `NO`/`INTERCO` antes del sprint específico de intercompany. No se publica ni empaqueta; la ejecución runtime queda planteada desde Studio mediante el wrapper. El orden de sprints no cambia. | XML válido en `oportunidad-generar-core.xaml` y `test/test_generar_core_pc.xaml`. `analyze-file` del wrapper OK con solo `ST-ANA-009` informativo. `analyze-file` del XAML principal carga el archivo y mantiene las incidencias ya previstas: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016` y `ST-ANA-009`, sin error técnico nuevo de parseo/carga. Sintaxis PowerShell del comparador OK. | Iniciar `S02`: estabilizar contrato CORE con prefijos `in_`/`out_` y añadir `Multiple Assign` inicial para modo unitario. |
| S02 | 2026-04-29 | `E01-T01`, `E01-T02`, `E01-T03`, `E01-T05` | Renombrado el contrato público de `oportunidad-generar-core.xaml` a argumentos `in_` y salida `out_NumeroSFLeido`. Actualizado el wrapper `test/test_generar_core_pc.xaml` para invocar el nuevo contrato. Añadidas variables locales de trabajo y la secuencia `00 Preparar datos locales para pruebas y ejecución` con `Multiple Assign` inicial que copia todos los `in_`. Sustituidas las referencias directas a argumentos por variables locales en logs, validación, copia de plantilla, lectura PPO y relleno CORE. Documentado el contrato de errores bloqueantes frente a datos funcionales faltantes en el XAML y en la documentación viva. | Se mantiene la decisión de no agrupar argumentos individuales aunque persista `ST-DBP-002`. Se prioriza claridad de integración, por eso se conservan nombres públicos largos y con origen `SAP_`/`SF_`; `ST-NMG-002` y `ST-NMG-016` quedan como avisos aceptados. La lógica Excel/COM no se ha refactorizado en este sprint para no mezclar alcance con S03-S08. | XML válido en `oportunidad-generar-core.xaml` y `test/test_generar_core_pc.xaml`. `analyze-file` del wrapper OK con solo `ST-ANA-009` informativo. `analyze-file` del XAML principal carga el archivo y devuelve las incidencias esperadas/aceptadas: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016` y `ST-ANA-009`; sin error técnico nuevo de parseo/carga. Baseline CSV mantiene 56 filas y se actualizaron solo notas de contrato. | Iniciar `S03`: rediseño del esqueleto visual y eliminación progresiva de la validación en C# (`E02-T01`, `E02-T02`, `E02-T08`). |

| S03 | 2026-04-29 | `E02-T01`, `E02-T02`, `E02-T08` | Eliminado el `Invoke Code` de validación de entrada. La validación técnica queda con actividades visibles (`If`, `Throw`, `Assign`, `FileExistsX`, `FolderExistsX`) y la selección/copia de plantilla queda en `02 Seleccionar y copiar plantilla CORE`. Reorganizado el flujo principal con secuencias `01` a `08`: validación, selección/copia, lectura PPO, construcción de modelo, Project Infor, Resources, Cost Planning y recálculo/guardado/fin. Añadidas anotaciones `Objetivo / Entradas / Salida` y ramas visuales PC/AT en `Resources` y `Cost Planning`. Los dos `Invoke Code` restantes se renombran como bloques transitorios S04 y S05-S08 con cabecera `Objetivo / Entradas / Salidas / Regla de negocio / No hace`. | La extracción real de lectura PPO y escritura CORE queda para S04-S07; en S03 se deja visible el esqueleto sin cambiar el comportamiento funcional esperado. Las secuencias `05` a `07` son puntos visuales preparados y la escritura efectiva sigue dentro del bloque transitorio de `08` hasta los sprints específicos. | XML válido en `oportunidad-generar-core.xaml`. `analyze-file` del wrapper OK con solo `ST-ANA-009` informativo. `analyze-file` del XAML principal carga el archivo y devuelve las incidencias aceptadas: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016` y `ST-ANA-009` (`Actividades: 76, Ifs: 10`); sin errores nuevos de parseo/carga. | Iniciar `S04`: mover lectura PPO a actividades Excel y construir DataTables/variables de negocio (`E02-T03`, `E02-T04`). |
| S04 | 2026-04-29 | `E02-T03`, `E02-T04` | Sustituido el bloque transitorio `S04 - leer PPO y construir modelo` por lectura visible con actividades Excel: `Read Cell` para cabecera, fechas e importes clave; `Read Range` para `Parametros`, `Presupuesto`, `Facturación y SAP` y `Sintesis Precio`. Añadidas variables locales de celdas PPO, rangos raw y DataTables de negocio: `dtCabeceraPPO`, `dtPerfilesPPO`, `dtCostesPorAnio`, `dtHorasPorAnio`, `dtGastos`, `dtCompras`, `dtTarifasAT` y `dtMesesPPO`. La secuencia `04 Construir modelo de negocio` deja visibles `For Each Row` sobre perfiles, gastos y compras preparados. | Para mantener operativo el bloque de escritura pendiente de S05-S08, se conserva una utilidad C# de transformación desde celdas/rangos ya leídos a `ModeloPPOJson`; esta utilidad ya no abre Excel, no usa COM, no escribe el CORE y no crea ficheros. La lectura del PPO queda descargada a actividades UiPath y el modelo de negocio queda disponible en DataTables para los siguientes sprints. | XML válido en `oportunidad-generar-core.xaml`. `analyze-file` del XAML principal ejecutado con salida correcta y sin errores técnicos reportados. `analyze-file` del wrapper `test/test_generar_core_pc.xaml` OK: `No se encontraron errores`. Se limpiaron metadatos locales generados por el analizador para dejar modificado solo el XAML y este backlog. | Iniciar `S05`: mover `Project Infor` y mapeos fijos a actividades `Write Cell`, consumiendo `dtCabeceraPPO` y manteniendo la política OK/KO por celda. |
| S05 | 2026-04-29 | `E02-T05`, parte de `E03-T01` | Convertida la secuencia `05 Rellenar Project Infor` en mapeo visual con `Write Cell`: prepara valores desde `dtCabeceraPPO`, datos SAP/Salesforce y celdas PPO, y escribe `C5`, `C8:C13`, `C16`, `D16`, `C18:C19`, `C22:D24` y `C28:C34`. Cada dato funcional faltante se resuelve en el mismo bloque con texto corto y fuente roja sobre la celda afectada. El bloque C# transitorio deja de escribir `Project Infor`, se elimina la función `EscribirCabecera` y queda renombrado como `Transitorio S06-S08 - rellenar Resources Cost Planning y recalcular`. | Se usan actividades modernas de Excel para conservar escritura tipada y formato de fuente rojo. Para que Studio 23.10 cargue el XAML, las salidas de `Read Cell` clásicas se dejaron en formato XAML expandido con `OutArgument ui:GenericValue` y se retiraron atributos `ContinueOnError` no soportados por `ReadCell`/`ReadRange`. `Resources` y `Cost Planning` permanecen en el bloque transitorio hasta S06/S07. | XML válido en `oportunidad-generar-core.xaml`. `analyze-file` del wrapper `test/test_generar_core_pc.xaml` OK: `No se encontraron errores`. `analyze-file` del XAML principal carga correctamente y devuelve incidencias de analyzer no bloqueantes/aceptadas por contrato: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016`, `ST-NMG-009`, `ST-MRD-009` y `ST-ANA-009`; sin errores técnicos de parseo/carga. Se limpiaron metadatos locales generados por el analyzer. | Iniciar `S06`: implementar `Resources` visual PC/AT con ramas visibles y reducir dependencia del C# restante. |
| S06 | 2026-04-29 | `E02-T06`, parte de `E03-T02`, parte de `E03-T03` | Convertida la secuencia `06 Rellenar Resources` en escritura visible con actividades Excel. Se prepara una utilidad acotada `Utilidad S06 - preparar tablas Resources` que transforma `dtPerfilesPPO`, `dtCostesPorAnio`, `dtTarifasAT`, `dtMesesPPO` y el JSON opcional de codigos de empleado en tablas rectangulares. La rama AT limpia y escribe `B7:E36` y `G7:AF36` con dos filas por perfil; la rama PC limpia y escribe `B7:AD43` con una fila por perfil. El bloque transitorio queda renombrado a `Transitorio S07-S08 - rellenar Cost Planning y recalcular`, sin funcion ni llamada `EscribirRecursos`. | Se preserva la regla actual de intercompany `NO`/`INTERCO` hasta S08 para no mezclar alcance. En S06 se controla capacidad minima de Resources: AT hasta 15 perfiles y PC hasta 30 por compatibilidad con el Cost Planning pendiente. `Cost Planning` sigue en C# para S07; `Resources` ya no se escribe desde el bloque transitorio. | XML valido en `oportunidad-generar-core.xaml` y `test/test_generar_core_pc.xaml`. `analyze-file` del wrapper OK con solo `ST-ANA-009` informativo. `analyze-file` del XAML principal carga correctamente y devuelve incidencias aceptadas por contrato/diseno: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016`, `ST-NMG-009`, `ST-MRD-009` y `ST-ANA-009`; sin errores tecnicos de parseo/carga. Se limpiaron metadatos locales generados por el analyzer. | Iniciar `S07`: implementar `Cost Planning` visual PC/AT y retirar la escritura restante del bloque C# transitorio. |
| S07 | 2026-04-29 | `E02-T07`, `E03-T05`, parte de `E03-T03` | Convertida la secuencia `07 Rellenar Cost Planning` en escritura visible con actividades Excel. Se prepara la utilidad acotada `Utilidad S07 - preparar tablas Cost Planning`, que calcula prorrateos mensuales desde `dtHorasPorAnio`, `dtGastos`, `dtCompras` y `dtMesesPPO`, y genera tablas rectangulares PC/AT para horas, riesgos, garantia, gastos y compras. La rama AT limpia y escribe recursos `B10:F39`, horas `H10:AF39`, riesgos/garantia `H47:AF48`, gastos `B53:G60`/`H53:AF60` y compras `B68:G76`/`H68:AF76`. La rama PC limpia y escribe horas `G10:AE39`, riesgos/garantia `G47:AE48`, gastos `B53:D58`/`G53:AE58` y compras `B64:E73`/`G64:AE73`. La secuencia `08` queda reducida a vaciar facturacion real inicial en `Cost Overview`, guardar y recalcular; ya no escribe `Cost Planning`. | Se preserva la regla actual de intercompany `NO`/`INTERCO` hasta S08. La politica de costes interanuales queda explicita: `Resources` mantiene coste/hora del anio de inicio y `Cost Planning` prorratea horas/importes por anualidad entre los meses reales de cada anio. Se anade control estructural de capacidad de `Cost Planning`: PC 30 perfiles, 6 gastos y 10 compras; AT 15 perfiles, 8 gastos y 9 compras; duracion maxima 25 meses. | XML valido en `oportunidad-generar-core.xaml`. `analyze-file` del wrapper `test/test_generar_core_pc.xaml` OK con solo `ST-ANA-009` informativo. `analyze-file` del XAML principal carga correctamente y devuelve las incidencias aceptadas por contrato/diseno: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016`, `ST-NMG-009`, `ST-MRD-009` y `ST-ANA-009`; sin errores tecnicos de parseo/carga ni avisos nuevos de variables S07. Se limpiaron metadatos locales generados por el analyzer. | Iniciar `S08`: robustecer intercompany y datos faltantes con matriz trazable y errores funcionales en rojo. |
| S08 | 2026-04-29 | `E03-T02`, `E03-T04`, `E03-T06` | Añadida la secuencia `04B Preparar reglas S08 de capacidad e intercompany`, que crea `dtMatrizIntercompany` con compañía proyecto, entidad recurso, decisión de negocio y valor literal de plantilla. La capacidad real se mide sobre el CORE copiado antes de escribir: filas/perfiles de `Resources`, perfiles/gastos/compras de `Cost Planning` y se usan rangos dinámicos en las ramas PC/AT. Actualizadas las utilidades de `Resources` y `Cost Planning` para validar perfiles/gastos/compras contra la plantilla recibida, normalizar intercompany por WBS primero y entidad de recurso después, y escribir errores cortos cuando falten datos funcionales dinámicos. La utilidad final pasa a `Utilidad S08 - aplicar rojo dinámico y recalcular CORE`, coloreando en rojo los errores escritos por `Write Range`. | Aunque la matriz funcional usa `No intercompany` / `Intercompany`, se conserva en las celdas el literal oficial `NO` / `INTERCO` mediante la columna `ValorPlantilla`, porque las fórmulas de `Cost Planning` de la plantilla oficial calculan totales con esos textos. Los textos rojos dinámicos quedan acotados: `Intercompany no determinable`, `Entidad recurso no disponible`, `Codigo empleado no disponible`, `Coste no disponible`, `Tarifa AT no disponible`, `Proveedor no disponible`, `PO compra no disponible` y `Ariba no disponible`. | XML válido en `oportunidad-generar-core.xaml`. `analyze-file` del wrapper `test/test_generar_core_pc.xaml` OK con solo `ST-ANA-009` informativo. `analyze-file` del XAML principal carga correctamente y mantiene incidencias aceptadas por contrato/diseño: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016`, `ST-NMG-009`, `ST-MRD-009` y `ST-ANA-009` (`Actividades: 282, Ifs: 29`); sin errores nuevos de parseo/carga. Se limpiaron metadatos locales generados por el analyzer. | Iniciar `S09`: validación CORE end-to-end, actualizar baseline PC por los cambios de intercompany/datos faltantes y comprobar integridad del Excel generado. |
| S09 | 2026-04-29 | `E04-T01`, `E04-T02`, `E04-T03`, `E04-T04`, `E04-T05` | Añadida validacion explicita de hojas obligatorias en `Utilidad S08 - preparar matriz intercompany y capacidad plantilla`, con error accionable si falta `Project Infor`, `Cost Overview`, `Resources`, `Cost Planning`, `Monthly View` o `Cost Summary`. Creado `test/validar_integridad_core_pc.ps1` para abrir el CORE generado, forzar recalculo, guardar/reabrir, validar hojas, ausencia de `Trazabilidad_RPA`, numero de hojas, dimensiones, cuadros `NOTA`, recuento de formulas, errores de formula y baseline. Creado `test/validar_negativos_core.ps1` para comprobar estaticamente errores bloqueantes y textos funcionales en rojo. Actualizado `test/baseline_core_pc.csv` a 58 validaciones, incluyendo billing real vacio en `Cost Overview!G17/I17` y notas de intercompany S08. Actualizado `test/README.md` con el flujo S09. | La CLI local sigue sin ejecutar XAML sueltos; por tanto el CORE real debe generarse lanzando `test/test_generar_core_pc.xaml` desde Studio 23.10.4 y despues ejecutar `test/validar_integridad_core_pc.ps1`. Se mantiene la excepcion aceptada de Workflow Analyzer para argumentos individuales y nombres humanos. No se publica ni empaqueta. | XML valido en XAML principal y wrapper. Sintaxis PowerShell OK en los tres scripts de test. `test/validar_negativos_core.ps1` OK: 16 casos bloqueantes y 27 funcionales verificados estaticamente. Smoke test de `test/validar_integridad_core_pc.ps1 -SkipBaseline` OK sobre copia temporal de la plantilla. `analyze-file` del wrapper OK con solo `ST-ANA-009`. `analyze-file` del XAML principal carga correctamente y mantiene incidencias aceptadas: `ST-DBP-002`, `ST-NMG-002`, `ST-NMG-016`, `ST-NMG-009`, `ST-MRD-009` y `ST-ANA-009`; sin errores tecnicos nuevos. | Iniciar `S10`: preparar handover e integracion con proceso padre. Antes de cerrar una entrega funcional, ejecutar en Studio el wrapper PC y lanzar el validador integral sobre `.local/test-output/CORE_PC_20250256445_test.xlsx`. |
| S10 | 2026-04-29 | `E05-T01`, `E05-T02`, `E05-T03`, `E06-T01`, `E06-T02`, `E06-T03` | Creado `test/ejemplo_invocacion_proceso_padre.xaml` como workflow ejecutable de handover: simula una fila de oportunidad preparada por el padre, mapea todos los argumentos `in_`, invoca `oportunidad-generar-core.xaml` con `Invoke Workflow File` y recoge `out_NumeroSFLeido`. Creada `docs/guia-integracion-core.md` con proposito, contrato de datos Salesforce/SAP, supuestos de integracion, errores/logs y validacion recomendada. Actualizado `test/README.md` para enlazar el ejemplo y la guia. | Se confirma que el modulo no decide carpetas ni nombres: `in_RutaCORE` llega resuelto por el proceso padre. Los logs del XAML principal ya cumplen el patron exigido: inicio con PPO/tipo/ruta, fin OK con CORE/NumeroSF/tipo y error con PPO/CORE/detalle relanzado. El ejemplo de handover queda separado del wrapper de baseline para no mezclar pruebas comparativas con documentacion de integracion. | XML valido en `test/ejemplo_invocacion_proceso_padre.xaml`, `test/test_generar_core_pc.xaml` y `oportunidad-generar-core.xaml`. `analyze-file` del ejemplo de handover OK con solo `ST-ANA-009` informativo. `analyze-file` del wrapper OK con solo `ST-ANA-009`. `analyze-file` del XAML principal OK: `No se encontraron errores`. `test/validar_negativos_core.ps1` OK. | Iniciar `S11`: analisis y contrato IPF (`E07-T01`, `E07-T02`, `E07-T03`). Antes de entrega funcional, ejecutar desde Studio el wrapper PC y el validador integral sobre el CORE generado. |
| S11 | 2026-04-29 | `E07-T01`, `E07-T02`, `E07-T03` | Analizadas `Plantilla IPF.xlsx` y `IPF - ICT - GAP - ZEV-PCE0012 - Despliegue Smart Airport nueva terminal-  FEB26 v1.xlsx` con Excel COM: ambas tienen 5 hojas, la hoja operativa visible es `Invoice Request` y no hay diferencias de celdas, formulas ni validaciones entre plantilla y ejemplo en los rangos usados. Creada `docs/guia-integracion-ipf.md` con mapa de hojas, celdas, listas validadas y contrato S11. Creado `oportunidad-generar-ipf.xaml` como submodulo separado con argumentos `in_`/`out_`, `Multiple Assign` inicial, validacion tecnica de rutas y copia de plantilla a `in_RutaIPF`. Creado wrapper `test/test_generar_ipf.xaml` para probar IPF aislado en `.local\test-output`. Actualizado `test/README.md`. | El IPF queda separado de CORE para no contaminar `oportunidad-generar-core.xaml`. El contrato usa nombres compatibles con Workflow Analyzer (`in_IPFRequestType`, `in_SAPCodigoPEPWBS`, etc.) y se limita a datos confirmados: cabecera estable, WBS, cliente/Sold-to, jefe de proyecto, moneda, Sales Order/PO y campos narrativos o de importe solo si el padre los aporta de forma fiable. Se mantiene la restriccion de no inventar `Summary Text`, `Text Narrative`, hitos ni importes exactos. En S11 el modulo solo copia la plantilla para hacer el flujo ejecutable; la escritura de celdas queda para S12. | XML valido en `oportunidad-generar-ipf.xaml` y `test/test_generar_ipf.xaml`. `analyze-file` del submodulo IPF OK con solo `ST-ANA-009` informativo. `analyze-file` del wrapper IPF OK con solo `ST-ANA-009` informativo. Verificada sintaxis PowerShell de extraccion de estructura IPF y comparacion sin diferencias de celda/formula/validacion. | Iniciar `S12`: implementar escritura de cabecera IPF y campos confirmados con actividades Excel, validando listas (`Request Type`, `Bill Type`, `Company Issuer`, `Currency`) y manteniendo vacios o textos rojos para datos funcionales faltantes. |

| S12 | 2026-04-29 | `E07-T04`, `E07-T05`, `E07-T06` | Implementada la escritura base completa de `oportunidad-generar-ipf.xaml`: tras copiar la plantilla, el modulo lee las listas `Request Type`, `Bill Type`, `Currency` y `Company Issuer` desde `Lists`, normaliza/defaulta valores y rellena `Invoice Request` con actividades Excel. Se escriben `C3`, `C6`, `C8`, `C11`, `F11`, `F14`, `C16`, `F16`, `C26`, `C30`, `C32`, `F32`, `D46`, `D47` y `F47`; se limpian `C18:C24`, `C29:F38` y `F47:G57` para no arrastrar datos de ejemplo. Actualizada la guia IPF y el README de test. | Se mantiene la regla de no inventar narrativa, hitos ni importes: `Summary Text`, `Text Narrative`, importe neto y comentarios solo se escriben si los aporta el proceso padre. Los defaults validos son `Invoice`, `Fixed` y `EUR`; si llega un valor fuera de lista, se escribe el default y se marca en rojo. `Company Issuer` se acepta si esta en lista o se infiere por WBS (`ZEV`/Connectis/ICT, `ZEZ`/Rosetta); si no, texto rojo. `Customer`, `Sold-to`, `Requested by`, `Sales Order` y `WBS` faltantes quedan como texto corto rojo. | XML valido en `oportunidad-generar-ipf.xaml` y `test/test_generar_ipf.xaml`. `analyze-file` del wrapper IPF OK con solo `ST-ANA-009` informativo. `analyze-file` del submodulo IPF carga correctamente y devuelve `ST-MRD-009` por anidacion de actividades modernas Excel y `ST-ANA-009` informativo (`Actividades: 94, Ifs: 15`), sin errores tecnicos de parseo/carga. Se limpiaron metadatos locales generados por el analyzer. | Iniciar `S13`: crear pruebas IPF y wrapper de ejemplo CORE + IPF (`E07-T07`, `E07-T08`). Ejecutar desde Studio `test/test_generar_ipf.xaml` y revisar `.local/test-output/IPF_20250256445_test.xlsx`. |

| S13 | 2026-04-29 | `E07-T07`, `E07-T08` | Creados baselines IPF `test/baseline_ipf_referencias.csv` y `test/baseline_ipf_generado.csv`. Creado `test/validar_integridad_ipf.ps1` para validar plantilla IPF oficial, ejemplo GAP, estructura del IPF generado, hojas, visibilidad, dimensiones, formulas, validaciones de lista, ausencia de auditoria externa y celdas esperadas por escenario. Creado `test/validar_negativos_ipf.ps1` para comprobar estaticamente bloqueantes tecnicos, defaults de listas, textos rojos funcionales y limpieza de narrativa. Creado `test/test_generar_ipf_datos_incompletos.xaml` para probar datos SAP/SF incompletos sin abortar. Creado `test/ejemplo_generar_core_ipf.xaml` como wrapper de handover que invoca primero CORE y despues IPF de forma condicional. Actualizados `test/README.md` y `docs/guia-integracion-ipf.md`. | Descubrimiento S13: la plantilla IPF oficial no esta realmente vacia; contiene datos de ejemplo equivalentes al ejemplo GAP en los rangos revisados. Por eso el baseline de referencias documenta esos datos y el baseline generado valida que el modulo limpia direccion, narrativa y comentarios cuando no los aporta el padre. Se mantiene que CORE e IPF no generan auditoria externa ni Markdown. El wrapper CORE+IPF queda separado del wrapper CORE S10 para no mezclar handover basico con generacion IPF condicional. | XML valido en los nuevos XAML S13 y en el IPF principal. Sintaxis PowerShell OK en los nuevos scripts. `test/validar_integridad_ipf.ps1 -SkipGenerated` OK sobre plantilla y ejemplo GAP. `test/validar_negativos_ipf.ps1` OK: 7 bloqueantes, 10 funcionales y wrappers S13 revisados. `analyze-file` de `test/test_generar_ipf.xaml`, `test/test_generar_ipf_datos_incompletos.xaml` y `test/ejemplo_generar_core_ipf.xaml` OK con solo `ST-ANA-009` informativo. `analyze-file` de `oportunidad-generar-ipf.xaml` carga correctamente y mantiene `ST-MRD-009` aceptado por anidacion y `ST-ANA-009` informativo; sin errores tecnicos nuevos. | Backlog S13 completado. Antes de entrega funcional final, ejecutar desde Studio los wrappers CORE PC, IPF completo, IPF datos incompletos y CORE+IPF, y lanzar los validadores integrales sobre los Excel generados. |

## 14. Decisiones de diseño cerradas y pendientes

Decisiones cerradas:

- Mantener argumentos individuales para el contrato público del módulo. No agrupar datos SAP/Salesforce en JSON/DataTable.
- Priorizar nomenclatura humana en español, usando prefijos `in_`, `out_` e `io_`.
- No cumplir Workflow Analyzer a rajatabla si una regla reduce la claridad funcional.
- El módulo CORE final no debe generar auditoría externa ni Markdown.
- El módulo IPF tampoco debe generar auditoría externa.
- No publicar nada en Orchestrator ni preparar paquete dentro de este backlog.
- El entregable será el XAML final para que un compañero lo integre en su proyecto.
- El naming del CORE no se define en el submódulo: llega completamente resuelto en `in_RutaCORE`.
- La complejidad debe descargarse hacia actividades UiPath y estructura visual.
- El `Multiple Assign` inicial es obligatorio para permitir pruebas unitarias con valores hardcodeados.
- La matriz intercompany queda definida en la sección 5.3 a partir del PDF funcional y la transcripción.
- Las plantillas CORE PC ampliadas 60/150 quedan fuera del alcance.
- No se requiere PPO AT real para esta fase; AT se inferirá de forma razonada desde la plantilla AT y las pruebas PC.
- Para IPF, solo se automatiza lo confirmado en el funcional y la toma de requisitos. No se inventan narrativa, hitos ni datos contractuales.
- La ausencia de datos funcionales no aborta la generación del CORE/IPF; se escribe texto corto en rojo o se deja vacío cuando la plantilla lo pida.
- El backlog es un texto vivo y debe actualizarse con los avances de cada sprint.

Decisiones pendientes:

- Ajustar el backlog restante según los descubrimientos técnicos de cada sprint.
- Confirmar durante la implementación los rangos/celdas exactos que solo puedan verificarse al abrir las plantillas en Studio/Excel.
- Registrar cualquier excepción funcional nueva que aparezca al probar con PPOs adicionales.
