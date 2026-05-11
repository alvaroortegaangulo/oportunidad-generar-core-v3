# Decisiones tecnicas CORE/IPF

Fecha de cierre: 2026-05-05  
Proyecto: `oportunidad-generar-core-v2`  
Alcance: handover de `oportunidad-generar-core.xaml`, libreria CORE, `oportunidad-generar-ipf.xaml`, wrappers y validadores.

Este registro captura las decisiones que debe respetar el proceso padre o cualquier mantenimiento posterior. No sustituye a `docs/guia-integracion-core.md`; la complementa con el motivo y la consecuencia tecnica de cada criterio.

| ID | Decision | Estado | Motivo | Impacto / validacion |
| --- | --- | --- | --- | --- |
| D01 | Mantener `oportunidad-generar-core.xaml` como fachada publica estable. | Cerrada | El padre debe invocar un unico submodulo CORE sin conocer la composicion interna. | Los subworkflows viven en `oportunidad-generar-core-lib`; el padre solo llama a CORE. |
| D02 | Mantener argumentos individuales `in_`/`out_` en CORE. | Cerrada | El mapeo SAP/Salesforce debe ser explicito para integradores UiPath y no depender de JSON opaco. | Se acepta `ST-DBP-002` por mas de 20 argumentos y nombres largos justificados. |
| D03 | No empaquetar ni publicar en Orchestrator dentro de este backlog. | Cerrada | El entregable operativo es XAML importable para integracion en otro proyecto. | No se crean assets, colas, paquetes ni configuracion de runtime. |
| D04 | No anadir dependencias nuevas. | Cerrada | Mantener compatibilidad con Studio 23.10.4 y target `Windows`. | `project.json` conserva los paquetes existentes. |
| D05 | Leer PPO por rangos estables y mapa tecnico. | Cerrada | Evita 29 `Read Cell` dispersos y concentra el mantenimiento de celdas. | `dtMapaLecturaPPO` documentado y protegido por `test/validar_negativos_core.ps1`. |
| D06 | Extraer construccion de modelo PPO a common workflow. | Cerrada | La fachada debe mostrar el orden funcional y no contener S04 monolitico. | `core-common-construir-modelo-ppo.xaml` devuelve cabecera, perfiles, costes, horas, gastos, compras, tarifas AT y meses. |
| D07 | Mantener `Project Infor` visible por bloques y escribir por mapa. | Cerrada | Agrupa reglas OK/KO cerca del dato sin repetir 24 escrituras fijas. | `dtMapaProjectInfor` y `dtErroresFuncionales` gobiernan valor, celda y rojo. |
| D08 | Extraer preparacion comun de `Resources` y `Cost Planning`. | Cerrada | Elimina duplicacion PC/AT sin mezclar layouts. | PC/AT no tienen `Invoke Code`; invocan `core-common-preparar-resources.xaml` y `core-common-preparar-cost-planning.xaml`. |
| D09 | Mantener subworkflows PC y AT separados para escritura. | Cerrada | Los layouts son distintos y deben ser revisables en Studio. | Cada subworkflow abre Excel una vez y escribe rangos especificos por `Write Range`. |
| D10 | Reutilizar la resolucion de intercompany desde `Resources` para AT en `Cost Planning`. | Cerrada | Evita duplicar normalizacion de compania, WBS, entidad y codigo de recurso. | `Cost Planning` consume `dt_ResourcesATFijos` cuando reproduce filas AT. |
| D11 | Reservar COM al cierre final de Excel. | Cerrada | UiPath 23.10 no expone de forma suficiente el pintado dinamico masivo y `CalculateFullRebuild`. | La inspeccion de capacidad ya usa actividades; COM queda en `Utilidad S08 - aplicar rojo dinamico y recalcular CORE` con `try/finally` y liberacion. |
| D12 | No abortar por datos funcionales faltantes. | Cerrada | El CORE/IPF debe quedar revisable aunque SAP/Salesforce no aporten todo. | Se escribe texto corto rojo o se deja vacio segun plantilla; los scripts negativos protegen textos esperados. |
| D13 | Abortan solo errores tecnicos o estructurales. | Cerrada | Sin ruta, plantilla, hoja o capacidad no existe destino seguro donde escribir. | CORE valida PPO, plantillas, carpeta, tipo PC/AT, hojas obligatorias, duracion y capacidad. IPF valida plantilla, ruta y carpeta. |
| D14 | Mantener CORE e IPF separados. | Cerrada | IPF no aplica siempre y tiene contrato/celdas propias. | `test/ejemplo_generar_core_ipf.xaml` muestra encadenado condicional sin fusionar XAML. |
| D15 | Automatizar IPF solo con datos confirmados. | Cerrada | No inventar narrativa, hitos, importes ni condiciones contractuales. | IPF limpia zonas de ejemplo y escribe narrativa/importe/comentarios solo si el padre los aporta. |
| D16 | Aceptar excepciones focalizadas de Workflow Analyzer. | Cerrada | Algunas reglas penalizan claridad del contrato o anidacion de actividades modernas. | `test/validar_calidad_e06.ps1` acepta solo `ST-ANA-009`, `ST-DBP-002`, `ST-MRD-009`, `ST-NMG-002`, `ST-NMG-009`, `ST-NMG-011`, `ST-NMG-016` y `ST-USG-020`. |
| D17 | Mantener validacion runtime manual desde Studio para XAML sueltos. | Cerrada | `UiRobot` local no ejecuta XAML crudos de este proyecto; Studio 23.10.4 es el camino fiable de prueba. | Wrappers en `test/` generan Excel reales y PowerShell valida integridad despues. |
| D18 | No crear auditoria externa ni Markdown desde los modulos. | Cerrada | La evidencia principal debe quedar en el Excel generado y en logs breves. | Los validadores comprueban ausencia de hojas externas como `Trazabilidad_RPA`. |

## Reglas para mantenimiento

- Cualquier cambio que toque contrato publico debe actualizar esta tabla, `docs/guia-integracion-core.md`, wrappers de `test/` y validadores.
- Cualquier `Invoke Code` nuevo debe ser pequeno, tener cabecera `Objetivo / Entradas / Salidas / Regla / No hace` y pasar `test/validar_negativos_core.ps1`.
- Si se cambia una celda de plantilla o un rango de escritura, actualizar primero el mapa tecnico y despues los baselines.
- Si negocio decide automatizar nuevos campos IPF, el dato debe venir de origen fiable del padre; no se infiere desde texto libre.
