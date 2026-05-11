# Guia de integracion IPF

`oportunidad-generar-ipf.xaml` es el submodulo separado para IPF. Desde S12 copia la plantilla a la ruta final, lee las listas de validacion desde la propia plantilla y rellena la hoja `Invoice Request` con los datos confirmados por el proceso padre.

El modulo IPF no usa assets, colas, Salesforce ni SAP. El proceso padre localiza la plantilla, decide la ruta final y aporta los datos ya resueltos.

## Analisis de plantilla

Ficheros revisados:

- `data\templates\IPF_template.xlsx`.
- Documento de referencia IPF original revisado durante el analisis funcional.
- `docs\documento_funcional_core_ipf_uipath.pdf`.
- `docs\Reunion-explicacion-proceso.md`.

El ejemplo GAP y la plantilla tienen hashes distintos, pero no presentan diferencias de celdas, formulas ni validaciones en los rangos usados. Por tanto, el ejemplo sirve como referencia de cumplimentacion de campos, no como libro con estructura distinta.

| Hoja | Visible | Rango usado | Uso |
| --- | --- | --- | --- |
| `Invoice Request` | Si | `B1:H62` | Hoja operativa para IPF de factura. |
| `UC Maint Contract+Inv Request` | No | `B1:J59` | Plantilla UC/mantenimiento fuera del alcance S11-S12 ordinario. |
| `Billing template instructions` | No | `B1:H58` | Instrucciones y ejemplo historico. |
| `Lists` | No | `A4:R43` | Listas de validacion. |
| `UC WBS codes` | No | `A1:D18` | Codigos UC, no aplican al caso GAP/CORE. |

## Mapa de celdas IPF

Hoja objetivo inicial: `Invoice Request`.

| Celda/rango | Campo | Regla S11/S12 |
| --- | --- | --- |
| `C3` | Request Type | Lista `Invoice`, `Credit Note`. Por defecto `Invoice`. |
| `C6` | Bill Type | Lista `Fixed`, `Variable`, `Adhoc`, `Contract`. Por defecto `Fixed`. |
| `C8` | Company Issuer | Lista `Lists!Q5:Q43`; valores relevantes: `Connectis ICT Services SAU`, `Global Rosetta`. Si no viene informado, S12 lo infiere por WBS (`ZEV`/Connectis/ICT o `ZEZ`/Rosetta); si no es determinable, texto rojo. |
| `C11` | Customer | `in_SFAccountName` o dato equivalente del padre. Si falta, texto rojo corto. |
| `F11` | Sold-to number | `in_SAPCodigoCliente`. Si falta, texto rojo corto. |
| `C14` | Contract number | No incluido en contrato S11 por no estar confirmado como dato disponible. Mantener vacio/formula de plantilla salvo decision futura. |
| `F14` | PO | `in_IPFPO` si el padre lo aporta. Si falta, dejar vacio salvo que negocio lo haga obligatorio. |
| `C16` | Requested by / jefe proyecto | `in_SAPNombreJefeProyecto` o responsable aportado por el padre. Si falta, texto rojo corto. |
| `F16` | Currency | Lista `GBP`, `EUR`, `USD $`, `CAN $`. Por defecto `EUR`. |
| `C18:C24` | Billing Address | No confirmado como origen fiable. S12 limpia el contenido de ejemplo y lo deja vacio. |
| `C26:F26` | Summary Text | Solo escribir `in_IPFSummaryText` si llega de origen fiable. No inventar texto. |
| `C28:F38` | Text Narrative / lineas | S12 limpia el ejemplo. Escribe `Pedido : [PO]` solo si llega PO, `in_IPFTextNarrative` en `C32` solo si llega texto exacto e `in_IPFImporteNeto` en `F32` solo si llega importe valido. No genera hitos por inferencia. |
| `D40` | Is VAT applicable? | Lista `Yes`, `No`. No automatizado en S11. |
| `D41` | VAT Rate | Lista `10%`, `15%`, `19%`, `20%`, `21%`, `0%`. No automatizado en S11. |
| `F41:F43` | Net/VAT/Gross | Formulas de plantilla. No sobrescribir. |
| `D46` | Sales Order | `in_SAPPedidoSalesOrder`. Si falta, texto rojo corto. |
| `D47` | WBS Code | `in_SAPCodigoPEPWBS`. Si falta, texto rojo corto. |
| `F47:G57` | Comments | S12 limpia el ejemplo y escribe `in_IPFComentariosFacturacion` si el padre aporta texto exacto. No inventa instrucciones de envio. |
| `D52` | WBS breakdown value | Formula/importe ligado al neto. S12 debe preservar formulas salvo dato confirmado. |
| `C61` | Please send to | Lista de destinatarios. Para Iberia existe `Iberiabilling @getronics.com`; no automatizado en S11. |

Listas relevantes detectadas:

| Lista | Rango | Valores |
| --- | --- | --- |
| Request Type | `Lists!B3:B5` | `Invoice`, `Credit Note` |
| Bill Type | `Lists!C3:C8` | `Fixed`, `Variable`, `Adhoc`, `Contract` |
| Currency | `Lists!D3:D7` | `GBP`, `EUR`, `USD $`, `CAN $` |
| VAT applicable | `Lists!F4:F6` | `Yes`, `No` |
| VAT rate | `Lists!H4:H9` | `10%`, `15%`, `19%`, `20%`, `21%`, `0%` |
| Billing mailbox | `Lists!L4:L10` | `Uk.Billing@Getronics.com`, `Iberiabilling @getronics.com`, `Belgium Billing`, `Germany Billing`, `Int. Billing`, `APAC billing`, `Other` |
| Company Issuer | `Lists!Q5:Q43` | Incluye `Connectis ICT Services SAU` y `Global Rosetta` |

## Contrato IPF

Para que el workflow nazca sin avisos de nomenclatura, los argumentos usan `in_` + CamelCase sin guiones bajos internos. El proceso padre debe mapearlos desde los mismos datos SAP/Salesforce que ya prepara para CORE.

| Argumento | Origen esperado | Puede ir vacio |
| --- | --- | --- |
| `in_RutaPlantillaIPF` | Ruta absoluta de plantilla IPF oficial. | No |
| `in_RutaIPF` | Ruta absoluta final del IPF generado. | No |
| `in_NumeroSF` | Identificador de oportunidad conocido por el proceso padre. | Si |
| `in_IPFRequestType` | Tipo de solicitud IPF. Default interno `Invoice`. | Si |
| `in_IPFBillType` | Tipo de facturacion. Default interno `Fixed`. | Si |
| `in_IPFCompanyIssuer` | Entidad emisora si el padre la conoce. | Si |
| `in_SFAccountName` | Cliente/cuenta desde Salesforce/proceso padre. | Si, queda texto rojo en S12 |
| `in_SAPCodigoCliente` | Sold-to number desde SAP/proceso previo. | Si, queda texto rojo en S12 |
| `in_SAPCodigoPEPWBS` | WBS/PEP. | Si, queda texto rojo en S12 |
| `in_SAPNombreJefeProyecto` | Jefe de proyecto o solicitante responsable. | Si |
| `in_IPFMoneda` | Moneda validada contra lista IPF. Default interno `EUR`. | Si |
| `in_SAPPedidoSalesOrder` | Sales Order/pedido SAP. | Si, queda texto rojo en S12 |
| `in_IPFPO` | PO/customer reference si existe. | Si |
| `in_IPFSummaryText` | Summary text exacto si viene de fuente fiable. | Si, dejar vacio si falta |
| `in_IPFTextNarrative` | Narrativa exacta si viene de fuente fiable. | Si, dejar vacio si falta |
| `in_IPFImporteNeto` | Importe neto exacto si viene de fuente fiable. | Si, dejar vacio si falta |
| `in_IPFComentariosFacturacion` | Comentarios exactos para facturacion si vienen informados. | Si |
| `out_NumeroSFLeido` | Salida de trazabilidad; devuelve `in_NumeroSF`. | N/A |

## Comportamiento S12

El modulo lee estas listas de la hoja `Lists` del IPF copiado: `Request Type`, `Bill Type`, `Currency` y `Company Issuer`. Si `Request Type`, `Bill Type` o `Currency` vienen vacios se usan los defaults funcionales `Invoice`, `Fixed` y `EUR`. Si vienen con un valor fuera de lista, se escribe el default valido y se marca la celda en rojo para revision.

Los datos obligatorios de cabecera que falten no abortan el proceso: `Customer`, `Sold-to`, `Company Issuer`, `Requested by`, `Sales Order` y `WBS` se escriben como texto corto en rojo. Los campos narrativos, comentarios e importe neto se dejan vacios si el padre no los aporta. Si el importe llega pero no se puede interpretar como numero, se escribe `Importe no valido` en rojo.

Para evitar datos ficticios de la plantilla de ejemplo, S12 limpia `C18:C24`, `C29:F38` y `F47:G57` antes de escribir los valores confirmados. No sobrescribe las formulas de `F41:F43` ni `D52`.

## Prueba local

Abrir el proyecto en UiPath Studio 23.10.4 y ejecutar `test/test_generar_ipf.xaml`.

El wrapper:

1. Crea `data\output\test`.
2. Invoca `oportunidad-generar-ipf.xaml`.
3. Genera `data\output\test\IPF_20250256445_test.xlsx` con cabecera, PO, narrativa/importe si se aportan y comentarios de prueba.
4. Recoge `out_NumeroSFLeido`.

Desde S13 hay dos pruebas IPF:

- `test/test_generar_ipf.xaml`: escenario completo con datos confirmados.
- `test/test_generar_ipf_datos_incompletos.xaml`: escenario funcional incompleto; debe generar sin abortar, con textos rojos en datos obligatorios y narrativa/importe/comentarios vacios.

La plantilla oficial revisada incluye datos de ejemplo, aunque funcionalmente se trate como plantilla base. El modulo limpia `C18:C24`, `C29:F38` y `F47:G57` para no arrastrar esos datos cuando el proceso padre no los aporta.

Validaciones PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_ipf.ps1 -SkipGenerated
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_ipf.ps1
powershell -ExecutionPolicy Bypass -File .\test\validar_integridad_ipf.ps1 -IpfPath .\data\output\test\IPF_20250256445_datos_incompletos.xlsx -Scenario DatosIncompletos
powershell -ExecutionPolicy Bypass -File .\test\validar_negativos_ipf.ps1
```

El wrapper `test/ejemplo_generar_core_ipf.xaml` muestra la integracion completa para el proceso padre: primero invoca CORE y despues IPF solo si el caso lo requiere.

## Restricciones funcionales

- No generar `Summary Text`, `Text Narrative`, hitos, importes exactos ni condiciones contractuales por inferencia.
- No generar auditoria externa ni Markdown desde el modulo IPF.
- Datos funcionales faltantes no bloquean. Se escriben como texto corto rojo o se dejan vacios segun corresponda a la plantilla.
- Errores bloqueantes: plantilla inexistente, ruta final no informada, carpeta destino inexistente o libro IPF no escribible/corrupto.
