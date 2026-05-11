# Backlog tecnico - entrega `oportunidad-generar-core-main`

Fecha: 2026-05-11  
Proyecto destino: `oportunidad-generar-core-main`  
Objetivo: preparar un proyecto UiPath autocontenido, entregable en ZIP, que ejecute `Main.xaml` para generar CORE PC y CORE AT desde los dos PPO recibidos, reutilizando los workflows existentes y corrigiendo las variaciones reales detectadas en los Excel.

## 1. Contexto funcional que debe gobernar la implementacion

La automatizacion no sustituye la revision del jefe de proyecto. Su funcion es prerrellenar una base CORE/IPF fiable, usando PPO como fuente principal y datos externos SAP/Salesforce ya resueltos por el proceso padre.

Reglas que no se pueden relajar:

- El robot no consulta Salesforce ni SAP desde los subworkflows reutilizables. Recibe esos datos por argumentos.
- El robot no inventa datos contractuales, narrativas IPF, hitos, facturacion real ni periodificaciones finas.
- Los datos funcionales faltantes no bloquean la generacion si existe un Excel destino seguro: deben quedar como texto corto rojo en la celda afectada.
- Solo bloquean errores tecnicos o estructurales: PPO inexistente/corrupto, plantilla inexistente/corrupta, tipo PC/AT invalido, hoja obligatoria ausente, carpeta destino no escribible, duracion invalida o plantilla sin capacidad.
- La lectura del PPO debe ser profesional y dinamica. Queda prohibida una bifurcacion tipo "layout antiguo/layout nuevo"; hay que detectar etiquetas, secciones y capacidades.
- `Main.xaml` debe ser un ejemplo real de uso del contrato definitivo, no un parche con datos adaptados para que pasen dos ficheros concretos.

## 2. Hallazgos de analisis

### 2.1 Documentos de negocio

Fuentes revisadas:

- `docs/documento_funcional_core_ipf_uipath.pdf`
- `docs/Reunion-explicacion-proceso.md`
- `docs/contexto-y-backlog.md`
- `docs/backlog-optimizaciones.md`
- `docs/guia-integracion-core.md`
- `docs/guia-integracion-ipf.md`
- `docs/decisiones-tecnicas-core.md`

Conclusiones operativas:

- El CORE se genera a partir de PPO, mas datos ya preparados por fase previa/SAP/Salesforce.
- PC y AT comparten cabecera, recursos y planificacion, pero AT duplica datos de horas reales/facturables y necesita tarifa de venta.
- El PPO puede propagarse a SAP "tal cual", por lo que no corresponde al robot corregir calidad de dato salvo avisar/visibilizar.
- Riesgos y garantia se imputan inicialmente al ultimo mes.
- Compras, gastos y recursos se prorratean plano por anualidad cuando no hay periodificacion fina.
- Facturacion emitida/billing real debe quedar vacia en la creacion inicial.

### 2.2 Estado del proyecto destino

En `oportunidad-generar-core-main` existe una estructura parcial, pero:

- `Main.xaml` actual esta vacio: solo contiene `Main Sequence`.
- Hay metadatos locales que no deben ir en el ZIP: `.local`, `.objects`, `.project`, `.settings`, `.tmh`.
- Hay duplicados de input con nombres exactos y nombres normalizados. Para una entrega limpia debe quedar una politica unica.
- La carpeta `test` contiene wrappers XAML, pero no todos los scripts PowerShell y baselines del directorio raiz estan adaptados.
- `README.md` describe un Main funcional que todavia no existe en la carpeta actual.

El proyecto de referencia en `referencia-resultado-gpt-pro/oportunidad-generar-core-main` aporta ideas utiles:

- Rutas relativas bajo `data`, `lib` y `test`.
- `Main.xaml` con doble invocacion PC/AT.
- `WorkflowFileName="lib\..."`.
- Lectura de `Hoja de datos` por etiquetas.
- Plantillas CORE extendidas hasta 60 meses.

Pero no debe copiarse sin revision:

- No incorpora los validadores PowerShell completos adaptados al entregable.
- Usa inputs normalizados, no exactamente los nombres solicitados.
- Mantiene zonas de lectura de `Presupuesto` por filas fijas que pueden confundir riesgos/garantia del PPO AT con compras.
- En `core-common-preparar-cost-planning.xaml`, las tablas de riesgos/garantia siguen con 25 columnas aunque se permite duracion 60.
- Los intentos previos en `oportunidad-generar-core-v2` generaron salidas con importes incorrectos o no disponibles (`C28` vacio/rojo o `3478.4` para AT), por lo que deben tratarse como ejemplos de lo que no hay que repetir.

### 2.3 Diferencias reales entre PPO

PPO PC solicitado:

- Fichero exacto: `20250256445  PPO_NuevaTerminalPVR v2 1.xlsm`
- Estructura tipo PPO anterior: `Hoja de datos` con etiquetas inglesas en filas 8-29.
- Tipo detectado: `PC`.
- Opportunity: `20250256445`.
- Cliente PPO: `GAP`.
- Descripcion: `Smart Airport en la nueva terminal de Puerto Vallarta`.
- Inicio: noviembre 2025.
- Duracion: 12 meses.
- Importe oferta: `583896.55`.
- Perfiles en `Parametros`: 7.
- Lineas de recursos en `Presupuesto`: 9, con siglas repetidas por distribucion interanual.
- Gastos: `7400`.
- Compras: `40000`.
- Riesgo: `18375.757042`.
- Garantia: `0`.

PPO AT solicitado:

- Fichero exacto: `20251160543_PPO_CEducación_AMS_SI_Lote 2_v01.xlsm`
- Estructura tipo PPO posterior: `Hoja de datos` con etiquetas espanolas desplazadas.
- Tipo detectado: `AT`, desde `Tipo de Contratacion`.
- Opportunity: `20251160543`.
- Cliente PPO: `Consejeria Educacion JCyL`.
- Descripcion: `Lote 2: Nucleo Estructural Java`.
- Inicio: enero 2026.
- Duracion: 48 meses.
- `Presupuesto!X8` contiene `3478.4`, que no representa el total de pedido para CORE.
- `Sintesis Precio!D12/K14/K18` contiene `465000`, que es el total economico coherente.
- Perfiles/lineas principales: `Susana Matarranz` (`AP`), `Jorge Puertas` (`PR`) y `Sopore especializado` (`AP`).
- Hay siglas repetidas (`AP`) con perfiles/personas distintas. No se puede deduplicar solo por sigla.
- No hay gastos/compras reales informados.
- La zona de riesgos/garantia aparece desplazada en torno a filas 43-48 de `Presupuesto`, dentro del rango que una lectura fija podria tratar erroneamente como compras.

### 2.4 Plantillas

Las plantillas originales de `ficheros-auxiliares` tienen capacidad mensual aproximada de 25 meses:

- CORE PC original: `Resources` hasta `AD`, `Cost Planning` hasta `AE`, `Monthly View` hasta `AC`.
- CORE AT original: `Resources` hasta `AE`, `Cost Planning` hasta `AF`, `Monthly View` hasta `AC`.

El PPO AT dura 48 meses, por lo que la entrega necesita plantillas con capacidad ampliada o seleccion dinamica de variante. La referencia GPT Pro incluye plantillas extendidas:

- CORE PC extendida: cabeceras hasta 60 meses (`Resources` hasta `BM`, `Cost Planning` hasta `BN`, `Monthly View` hasta `BL`).
- CORE AT extendida: cabeceras hasta 60 meses (`Resources` hasta `BN`, `Cost Planning` hasta `BO`, `Monthly View` hasta `BL`).

Estas plantillas extendidas deben validarse, no asumirse correctas.

## 3. Principios tecnicos para la solucion

- Trabajar dentro de `oportunidad-generar-core-main` como proyecto final autocontenido.
- Mantener los workflows reutilizables en `lib`.
- Mantener `lib\oportunidad-generar-core.xaml` como fachada CORE.
- Mantener `lib\oportunidad-generar-ipf.xaml` separado.
- Usar `Main.xaml` solo como orquestador de ejemplo y prueba real: preparar rutas, datos SAP/SF de prueba e invocar el contrato publico.
- No meter reglas especificas de los dos PPO en `Main.xaml`.
- No crear dos parsers por version de PPO. Implementar lectura por etiquetas, cabeceras y secciones detectadas.
- No modificar formulas de salida mediante escrituras masivas que destruyan la plantilla.
- No entregar caches locales, metadatos generados ni ficheros temporales de UiPath.
- Los outputs en `data\output` deben generarse ejecutando `Main.xaml`; si se incluyen en el ZIP final, deben estar validados y ser reproducibles.

## 4. Backlog priorizado

### E00 - Preparar carpeta de entrega limpia

Objetivo: convertir `oportunidad-generar-core-main` en el proyecto real a entregar.

Tareas:

1. Limpiar del proyecto destino los directorios generados por UiPath que no deben versionarse ni empaquetarse:
   - `.local`
   - `.objects`
   - `.project`
   - `.settings`
   - `.tmh`
   - cualquier cache equivalente.

2. Definir estructura final:

```text
oportunidad-generar-core-main\
  Main.xaml
  project.json
  README.md
  VALIDACION_LOCAL.md
  lib\
  data\
    input\
    templates\
    output\
  test\
  docs\
```

3. En `data\input`, dejar solo los PPO que deben gobernar la prueba:
   - `20250256445  PPO_NuevaTerminalPVR v2 1.xlsm`
   - `20251160543_PPO_CEducación_AMS_SI_Lote 2_v01.xlsm`

4. Si se decide usar nombres normalizados para evitar problemas de encoding en XAML, documentar esa decision y mantener una sola copia por PPO. No deben coexistir duplicados que hagan no determinista `Directory.GetFiles`.

5. En `project.json`:
   - `main` debe ser `Main.xaml`.
   - `name` debe ser `oportunidad-generar-core-main`.
   - conservar `targetFramework: Windows`.
   - conservar dependencias necesarias para Excel moderno, System, Testing y UIAutomation solo si los XAML las requieren.
   - no anadir paquetes nuevos sin justificacion.

Criterios de aceptacion:

- Abrir la carpeta en UiPath Studio muestra `Main.xaml` como entry point.
- No hay rutas absolutas a la raiz antigua.
- No hay referencias a `ficheros-auxiliares` ni `.local` en XAML del proyecto entregable.
- `rg -n "C:\\|ficheros-auxiliares|oportunidad-generar-core-lib|\.local" oportunidad-generar-core-main` no devuelve referencias funcionales, salvo documentacion historica claramente marcada.

### E01 - Consolidar workflows reutilizables en `lib`

Objetivo: migrar la version mas avanzada y correcta de los workflows al layout `lib`.

Tareas:

1. Tomar como base los XAML de raiz y comparar con los de `referencia-resultado-gpt-pro`.

2. Portar al destino solo los cambios utiles de la referencia:
   - rutas `WorkflowFileName="lib\..."`;
   - lectura moderna `ReadRangeX` donde ya este validada;
   - cabecera PPO por etiquetas;
   - soporte mensual superior a 25 meses;
   - plantillas extendidas si superan la validacion de E04.

3. No portar cambios no justificados:
   - heuristicas opacas;
   - silencios de error;
   - salidas ya generadas sin trazabilidad;
   - templates reducidos o alterados sin validacion.

4. Validar XML de todos los XAML:
   - `Main.xaml`
   - `lib\oportunidad-generar-core.xaml`
   - `lib\oportunidad-generar-ipf.xaml`
   - `lib\core-common-construir-modelo-ppo.xaml`
   - `lib\core-common-preparar-resources.xaml`
   - `lib\core-common-preparar-cost-planning.xaml`
   - `lib\oportunidad-generar-core-pc.xaml`
   - `lib\oportunidad-generar-core-at.xaml`
   - wrappers de `test`.

Criterios de aceptacion:

- Todos los `Invoke Workflow File` dentro del entregable apuntan a `lib\...` o a workflows locales existentes.
- PC/AT no contienen `Invoke Code` de escritura duplicada si ya existen common workflows.
- `lib\oportunidad-generar-core.xaml` conserva el contrato publico de argumentos `in_`/`out_`.

### E02 - Robustecer lectura de `Hoja de datos` del PPO

Objetivo: soportar tanto etiquetas inglesas del PPO PC como etiquetas espanolas del PPO AT sin ramificar por layout.

Tareas:

1. Leer `Hoja de datos` con etiquetas y valores, no solo valores:
   - rango minimo recomendado: `A8:D35`;
   - conservar columna de etiqueta, valor y posible ayuda/comentario.

2. Implementar normalizacion unica:
   - trim;
   - lower case;
   - sin acentos;
   - eliminar signos no alfanumericos para comparar.

3. Crear mapa de sinonimos por campo de negocio:
   - `company`: `Company`, `Entidad Presentadora`, `Compania`, `Sociedad`;
   - `unit`: `Unidad Negocio`, `Business Unit`;
   - `location`: `Customer Location`, `Oficina`, `Localizacion`;
   - `serviceType`: `Project/Service Type`, `Tipo Servicio`;
   - `pepType`: `PEP Type`, `Tipo de Contratacion`;
   - `portfolio`: `Porfolio`, `Portfolio`, `Tipo Producto`;
   - `practice`: `Practice`, `Sector`;
   - `subpractice`: `Subpractice`, `Zona`;
   - `opportunityType`: `Opportunity Type`, `Tipo de Contratacion`;
   - `accountManager`: `Account Manager`, `Responsable Comercial`;
   - `applicant`: `Applicant`, `Responsable Oferta`;
   - `opportunityCode`: `Opportunity Code`, `Opportunity Number`, `Codigo Oferta/Proyecto`, `Codigo Oportunidad`;
   - `customer`: `Customer`, `Cliente`;
   - `description`: `Descripcion Oferta`, `Description`;
   - `proposalDelivery`: `Date of Proposal Delivery`, `Fecha prevista presentacion al Cliente`;
   - `startMonth`: `Service Start Month`, `Mes Inicio servicio`;
   - `startYear`: `Service Start Year`, `Ano Inicio servicio`;
   - `duration`: `Service Duration`, `Duracion prevista`.

4. Mantener fallback por fila solo como ultimo recurso y trazable, nunca como fuente principal.

5. En `dtCabeceraPPO`, guardar tambien el origen real usado:
   - ejemplo: `Hoja de datos: Tipo de Contratacion`;
   - no dejar origen ficticio `B12` cuando se leyo por etiqueta en `B15`.

6. La deteccion de tipo debe priorizar:
   - `in_TipoProyecto` del proceso padre;
   - `Tipo de Contratacion` / `PEP Type`;
   - `Tipo Servicio` solo como apoyo.

Criterios de aceptacion:

- PPO PC detecta `PC`, opportunity `20250256445`, cliente `GAP`, duracion `12`.
- PPO AT detecta `AT`, opportunity `20251160543`, cliente `Consejeria Educacion JCyL`, duracion `48`.
- No existe separacion `layoutAntiguo/layoutNuevo`.
- Si falta `Opportunity Number`, se puede inferir desde nombre de fichero solo con log/advertencia y mensaje claro; no debe ocultarse.

### E03 - Rehacer lectura semantica de `Presupuesto`, `Parametros` y `Sintesis Precio`

Objetivo: evitar que cambios de filas en PPO rompan el modelo de negocio.

Tareas:

1. `Parametros`:
   - detectar fila de cabecera por presencia de `Perfil`, `Entidad`, `Sigla`;
   - detectar columnas de coste por anos, aunque el texto sea `Coste hora\n2026`;
   - crear `dtPerfilesPPO` con identidad interna robusta;
   - no deduplicar solo por `Sigla`.

2. Identidad de recurso:
   - si varias lineas comparten sigla pero tienen distinto perfil/persona/entidad, conservarlas como recursos distintos;
   - si varias lineas son la misma sigla/perfil/entidad repetida por anualidad, agregar horas/importes donde corresponda;
   - incorporar una clave interna, por ejemplo `ResourceKey = Sigla + "|" + Perfil + "|" + Entidad + "|" + RowSource`;
   - mantener `Sigla` como campo visible, pero no como clave unica.

3. `Presupuesto`:
   - detectar cabecera de anos en vez de asumir filas fijas;
   - detectar bloques por etiquetas: recursos, gastos, compras, riesgos, garantia;
   - leer totales por bloque desde columna total detectada (`TOTAL`, normalmente K), no desde indices fijos;
   - para recursos, leer horas por anualidad y coste total por linea;
   - para gastos/compras, leer descripcion y anualidades solo dentro del bloque correcto;
   - para riesgos/garantia, leer sus propios bloques y no meterlos en `dtCompras`.

4. Corregir caso AT:
   - `Riesgos (Mano de Obra)` y `Garantia` aparecen desplazados en torno a filas 43-48;
   - esas filas no deben pasar a compras;
   - si riesgos/garantia son cero, deben quedar cero o vacio segun regla de plantilla, no como compra ficticia.

5. `Sintesis Precio`:
   - detectar `Total Oferta sin IVA` / `Precio Total` por etiqueta;
   - usar `Sintesis Precio!D12`, `K14` o `K18` como fuente de importe total solo cuando el contexto de etiquetas lo confirme;
   - para el PPO AT, el importe correcto esperado para CORE es `465000`, no `3478.4`;
   - evitar la regla permanente "usar el mayor numero" sin etiqueta semantica. Puede quedar como fallback con advertencia, pero no como fuente principal.

6. Tarifas AT:
   - intentar leer tarifa directa desde la columna de tarifa de `Sintesis Precio`;
   - si no hay tarifa y existe venta/horas fiable, calcular tarifa ponderada;
   - si no existe venta real o es `0`, escribir `Tarifa AT no disponible` en rojo en Resources AT; no inventar tarifa.

Criterios de aceptacion:

- PPO PC mantiene importes esperados:
  - pedido `583896.55`;
  - horas `13570` o valor justificado segun fuente final elegida;
  - coste recursos `367515.14084`;
  - riesgos `18375.757042`;
  - gastos `7400`;
  - compras `40000`.
- PPO AT usa pedido `465000`.
- PPO AT no genera compras a partir de filas de riesgos/garantia.
- PPO AT conserva tres recursos de negocio, incluyendo las dos lineas con sigla `AP`.

### E04 - Validar y formalizar plantillas CORE de 60 meses

Objetivo: que el AT de 48 meses pueda generarse sin deformar el CORE ni romper formulas.

Tareas:

1. Decidir explicitamente la fuente de las plantillas:
   - si existen plantillas oficiales ampliadas, incorporarlas en `data\templates`;
   - si solo existen plantillas estandar, derivar plantillas extendidas con proceso tecnico documentado y validarlas a fondo.

2. Validar las plantillas extendidas de la referencia GPT Pro antes de adoptarlas:
   - hojas obligatorias;
   - numero de formulas;
   - ausencia de `#REF!`;
   - rangos de resumen que cubren hasta el mes 60;
   - formulas de `Cost Summary`, `Monthly View`, `Cost Overview`, `Resources` y `Cost Planning`;
   - validaciones/listas y estilos de celdas.

3. Medir capacidad real desde la plantilla:
   - meses utiles por hoja;
   - filas utiles de `Resources`;
   - perfiles/gastos/compras utiles de `Cost Planning`;
   - no usar un numero fijo si la plantilla tiene otra capacidad.

4. Corregir `core-common-preparar-cost-planning.xaml`:
   - las tablas `out_dt_CostPlanningPCRiesgosGarantia` y `out_dt_CostPlanningATRiesgosGarantia` deben tener `duracionCostPlanning` columnas, no 25;
   - imputar riesgos/garantia en `duracionCostPlanning - 1`;
   - validar explicitamente que el indice del ultimo mes existe.

5. Corregir escrituras PC/AT:
   - rangos de limpieza y escritura deben acabar en la columna calculada por duracion real;
   - no escribir mas columnas que la plantilla;
   - no limpiar zonas fuera de la capacidad real.

Criterios de aceptacion:

- CORE AT de 48 meses escribe cabeceras mensuales hasta diciembre 2029.
- `Resources`, `Cost Planning` y `Monthly View` mantienen formulas y formatos.
- No aparecen `#REF!` ni formulas rotas tras recalculo.
- Si se prueba una duracion superior a la plantilla, el error indica duracion detectada y capacidad real.

### E05 - Robustecer Resources

Objetivo: que la hoja `Resources` represente correctamente perfiles, costes, tarifas e intercompany en PC y AT.

Tareas:

1. Revisar `core-common-preparar-resources.xaml` para que consuma la identidad robusta de E03.

2. Codigo empleado:
   - mantener JSON por sigla como compatibilidad;
   - permitir claves mas especificas para duplicados, por ejemplo `AP|Susana Matarranz`;
   - si no hay codigo especifico, usar sigla si procede;
   - si falta, escribir `Codigo empleado no disponible` en rojo.

3. Coste:
   - usar coste del ano de inicio si existe;
   - si no existe, usar primer coste disponible con advertencia;
   - si no hay coste, texto rojo `Coste no disponible`.

4. AT:
   - dos filas por recurso: coste/hora y tarifa venta;
   - no fusionar recursos distintos por compartir sigla;
   - si falta tarifa, texto rojo `Tarifa AT no disponible`.

5. Intercompany:
   - inferir compania proyecto desde WBS/PEP primero;
   - inferir entidad recurso desde PPO;
   - resolver con matriz parametrizada;
   - si no se puede resolver, `Intercompany no determinable` en rojo.

Criterios de aceptacion:

- PC mantiene recursos esperados y literales oficiales `NO` / `INTERCO`.
- AT genera filas para Susana, Jorge y soporte especializado.
- Las dos lineas `AP` del AT no se pierden ni se agregan indebidamente.

### E06 - Robustecer Cost Planning

Objetivo: que la planificacion refleje recursos, horas, gastos, compras, riesgos y garantia sin mezclar conceptos.

Tareas:

1. Consumir `dtHorasPorAnio` con clave de recurso robusta, no solo sigla.

2. Prorratear horas por anualidad:
   - distribuir cada anualidad entre los meses reales de ese ano dentro del proyecto;
   - no repartir horas de 2026 sobre meses de 2025 ni viceversa.

3. AT:
   - generar horas reales y facturables segun layout AT;
   - si no hay regla de facturable distinta, documentar que se replica la base inicial para revision del JP.

4. Gastos/compras:
   - cargar solo conceptos de los bloques correctos;
   - no crear compras si el PPO AT no las trae;
   - si hay compras y faltan proveedor/PO/Ariba de SAP, escribir textos rojos acotados.

5. Riesgos/garantia:
   - imputar al ultimo mes real del proyecto;
   - soportar 48 meses;
   - si vienen cero, respetar cero/vacio segun plantilla.

Criterios de aceptacion:

- PC conserva baseline funcional actual.
- AT no contiene compras falsas por filas de riesgo.
- AT no falla por duracion 48.
- Los totales de `Cost Summary` son coherentes tras recalculo.

### E07 - Construir `Main.xaml` definitivo

Objetivo: que `Main.xaml` sea el ejemplo real de proceso padre para los dos PPO.

Tareas:

1. Crear `Main.xaml` no vacio con secuencia clara:
   - `00 Preparar rutas`;
   - `01 Validar inputs locales`;
   - `02 Preparar datos SAP/Salesforce de prueba`;
   - `03 Generar CORE PC 20250256445`;
   - `04 Generar CORE AT 20251160543`;
   - `05 Registrar fin`.

2. Construir rutas relativas desde la raiz del proyecto:
   - `data\input`;
   - `data\templates`;
   - `data\output`;
   - `lib`.

3. Crear `data\output` si no existe.

4. Resolver PPOs de forma determinista:
   - opcion preferida: usar exactamente los nombres solicitados;
   - opcion alternativa: buscar por id (`20250256445*.xlsm`, `20251160543*.xlsm`) y fallar si hay 0 o mas de 1 resultado.

5. Invocar `lib\oportunidad-generar-core.xaml` dos veces:
   - PC con `in_TipoProyecto = "PC"`;
   - AT con `in_TipoProyecto = "AT"`.

6. Usar datos SAP/Salesforce de prueba, plausibles y declarados:
   - PC puede basarse en referencia funcional:
     - WBS `ZEV-PCE0012-1001`;
     - codigo cliente `461530`;
     - Sales Order `714603126`;
     - PM `3900331` / `LLUIS GUERRA GONZALEZ`;
     - Practice Leader `3659518` / `Carlos Alvarez Gonzalez`;
     - Sales Manager `458354` / `Joan MARCER`;
     - account `SERVICIOS A LA INFRAESTRUCTURA` o `GAP` segun celda destino acordada;
     - practice `Smart Integration` o fallback PPO si se decide mantener la practica del PPO.
   - AT debe usar datos de prueba coherentes, marcados como no productivos:
     - WBS con prefijo de compania valido para la matriz;
     - codigo cliente de prueba;
     - Sales Order de prueba;
     - PM/PL/Sales Manager de prueba;
     - account `Consejeria Educacion JCyL`;
     - practice/sector coherente con PPO.

7. Preparar JSON de codigos de empleado:
   - PC: `JP`, `CSS`, `CJSS`, `A`, `II`, `P`, `DG`;
   - AT: `AP`, `PR` y, si se implementa clave compuesta, `AP|Susana Matarranz`, `AP|Sopore especializado`.

8. Definir salidas:
   - `data\output\CORE_PC_20250256445.xlsx`;
   - `data\output\CORE_AT_20251160543.xlsx`.

9. `Main.xaml` no debe generar IPF salvo que se acuerde explicitamente. El IPF queda incluido y testeado como submodulo reutilizable.

Criterios de aceptacion:

- Ejecutar `Main.xaml` desde Studio genera ambos CORE.
- `out_NumeroSFLeido` de cada invocacion coincide con el PPO.
- Los datos hardcodeados son faciles de localizar y cambiar.
- Ningun dato hardcodeado corrige artificialmente fallos del parser PPO.

### E08 - Adaptar tests al proyecto entregable

Objetivo: llevar al ZIP pruebas utiles, no solo wrappers decorativos.

Tareas:

1. Copiar y adaptar desde `test` raiz:
   - `validar_baseline_core_pc.ps1`;
   - `validar_integridad_core_pc.ps1`;
   - `validar_negativos_core.ps1`;
   - `validar_integridad_ipf.ps1`;
   - `validar_negativos_ipf.ps1`;
   - `validar_calidad_e06.ps1` o nuevo `validar_calidad_entrega.ps1`;
   - baselines CSV necesarios.

2. Cambiar rutas por defecto:
   - de `.local\test-output` a `data\output\test` o a `data\output`;
   - de `ficheros-auxiliares` a `data\templates`;
   - de `oportunidad-generar-core-lib` a `lib`.

3. Mantener wrappers XAML adaptados:
   - `test_generar_core_pc.xaml`;
   - nuevo o actualizado `test_generar_core_at.xaml`;
   - `test_generar_core_datos_incompletos.xaml`;
   - `test_generar_ipf.xaml`;
   - `test_generar_ipf_datos_incompletos.xaml`;
   - `ejemplo_invocacion_proceso_padre.xaml`;
   - `ejemplo_generar_core_ipf.xaml`.

4. Crear baseline CORE AT:
   - celdas criticas de `Project Infor`;
   - primeras filas de `Resources`;
   - primeras filas de `Cost Planning`;
   - cabeceras mensuales hasta mes 48;
   - ausencia de compras falsas;
   - importe `465000`.

5. Actualizar baseline CORE PC si las plantillas extendidas cambian dimensiones o formulas, pero los valores de negocio deben permanecer coherentes.

6. `validar_calidad_entrega.ps1` debe:
   - parsear XAML como XML;
   - validar que todos los `WorkflowFileName` existen;
   - ejecutar validadores negativos;
   - opcionalmente ejecutar Workflow Analyzer si Studio CLI existe;
   - validar outputs si existen;
   - fallar si encuentra `.local`, `.objects`, `.project`, `.settings`, `.tmh`.

Criterios de aceptacion:

- Los tests se ejecutan desde la raiz de `oportunidad-generar-core-main`.
- No dependen de la raiz del repositorio.
- Hay cobertura de PC, AT, datos incompletos e IPF.
- Los validadores detectan los errores observados en intentos previos: importe AT `3478.4`, importe no disponible, duracion truncada a 25 meses y compras falsas desde riesgos.

### E09 - Documentacion de entrega

Objetivo: que el responsable pueda abrir, ejecutar y entender el proyecto.

Tareas:

1. Actualizar `README.md` con:
   - objetivo del proyecto;
   - estructura;
   - como abrir en UiPath Studio;
   - como ejecutar `Main.xaml`;
   - que ficheros genera;
   - que datos SAP/SF son de prueba;
   - como ejecutar tests;
   - que no consulta Salesforce/SAP.

2. Actualizar `VALIDACION_LOCAL.md` con:
   - checks estaticos realizados;
   - resumen de estructura de PPO;
   - resumen de capacidad de plantillas;
   - pruebas pendientes de Studio si no pueden ejecutarse desde terminal;
   - outputs generados y fecha de generacion.

3. Copiar docs relevantes:
   - `documento_funcional_core_ipf_uipath.pdf`;
   - `Reunion-explicacion-proceso.md`;
   - `guia-integracion-core.md`;
   - `guia-integracion-ipf.md`;
   - `decisiones-tecnicas-core.md`.

4. Anadir nota de datos de prueba:
   - los valores SAP/Salesforce en `Main.xaml` no son productivos;
   - estan ahi para probar el contrato y deben sustituirse por el proceso padre real.

Criterios de aceptacion:

- README no promete algo que `Main.xaml` no haga.
- Se indica claramente que los outputs se regeneran ejecutando `Main.xaml`.
- Se documenta por que el PPO AT necesita plantilla con capacidad de 48 meses.

### E10 - Validacion final y ZIP

Objetivo: cerrar el entregable con evidencias.

Tareas:

1. Validacion estatica:
   - XML parse OK para todos los XAML;
   - todas las referencias `WorkflowFileName` resuelven;
   - `project.json` valido;
   - no hay rutas absolutas ni dependencias a raiz.

2. Validacion UiPath:
   - abrir proyecto en UiPath Studio 23.10.4;
   - ejecutar Workflow Analyzer focalizado;
   - ejecutar `Main.xaml`;
   - ejecutar wrappers principales si procede.

3. Validacion Excel:
   - abrir ambos CORE generados;
   - recalcular;
   - validar hojas obligatorias;
   - validar ausencia de `#REF!`;
   - validar dimensiones/meses;
   - validar baselines PC/AT;
   - validar que `Cost Summary` y `Monthly View` no quedan truncados.

4. Preparar ZIP:
   - incluir `Main.xaml`, `project.json`, `lib`, `data`, `test`, `docs`, `README.md`, `VALIDACION_LOCAL.md`;
   - excluir caches y temporales;
   - incluir outputs finales en `data\output` solo despues de generarlos con el Main definitivo y validarlos.

Criterios de aceptacion:

- El ZIP se puede descomprimir y abrir directamente como proyecto UiPath.
- `Main.xaml` genera ambos CORE sin tocar rutas externas.
- Los validadores se pueden ejecutar dentro de la carpeta descomprimida.
- El contenido del ZIP no incluye metadatos locales ni duplicados de inputs.

## 5. Orden recomendado de ejecucion

1. E00: limpiar y fijar estructura de entrega.
2. E01: consolidar workflows en `lib`.
3. E02 y E03: robustecer parser PPO antes de tocar Main.
4. E04: validar/adoptar plantillas de 60 meses y corregir tablas de riesgos/garantia.
5. E05 y E06: cerrar Resources y Cost Planning con identidad robusta.
6. E07: construir `Main.xaml`.
7. E08: adaptar tests y baselines.
8. E09: documentar.
9. E10: ejecutar validacion final y preparar ZIP.

## 6. Anti-patrones a evitar

- Crear ramas `if layout antiguo` / `if layout nuevo`.
- Deduplicar recursos solo por sigla.
- Tratar filas de riesgos/garantia del PPO AT como compras.
- Usar `Presupuesto!X8` como importe total sin validar etiquetas.
- Usar siempre "el mayor importe" sin origen semantico.
- Aumentar duracion maxima a 60 pero dejar tablas internas de 25 columnas.
- Meter datos de prueba en workflows reutilizables.
- Hacer que `Main.xaml` compense fallos del parser.
- Entregar `.local`, `.objects`, `.project`, `.settings`, `.tmh`.
- Quitar validadores PowerShell para que la entrega parezca mas limpia.

## 7. Definicion de terminado

La tarea se considera completada cuando:

- `oportunidad-generar-core-main\Main.xaml` genera:
  - `data\output\CORE_PC_20250256445.xlsx`;
  - `data\output\CORE_AT_20251160543.xlsx`.
- Ambos CORE se abren y recalculan sin errores de formula.
- CORE PC conserva los valores funcionales de referencia.
- CORE AT refleja 48 meses, pedido `465000`, tres recursos de negocio y ausencia de compras falsas.
- Los datos SAP/Salesforce de prueba estan centralizados en `Main.xaml` y son trazables.
- Los subworkflows siguen siendo reutilizables por un proceso padre real.
- Los tests y validadores se ejecutan dentro del proyecto entregable.
- El ZIP final es autocontenido y no depende de la raiz del repositorio.
