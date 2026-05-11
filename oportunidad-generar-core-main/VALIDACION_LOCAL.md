# Validación local de empaquetado

Validación realizada en el entorno de generación del ZIP. No sustituye la ejecución final en UiPath Studio/Windows con Excel instalado.

## Comprobaciones estáticas

- XAML parseados como XML: 14.
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
- `#REF!` en XML de plantilla: `0`

### `data/templates/CORE_AT_template.xlsx`

- Hojas detectadas: `Project Infor, Cost Overview, Resources, Cost Planning, Monthly View, Cost Summary, Ayuda, aux Billing Plan`
- Dimensiones relevantes: `{'Resources': 'B2:CA69', 'Cost Planning': 'B1:CB99'}`
- `#REF!` en XML de plantilla: `0`

### `data/templates/IPF_template.xlsx`

- Hojas detectadas: `Invoice Request, UC Maint Contract+Inv Request, Billing template instructions, Lists, UC WBS codes`
- Dimensiones relevantes: `{}`
- `#REF!` en XML de plantilla: `0`

## Pendiente en entorno UiPath

Ejecutar `Main.xaml` desde UiPath Studio y verificar que se crean `data\output\CORE_PC_20250256445.xlsx` y `data\output\CORE_AT_20251160543.xlsx`. Este sandbox no dispone de UiPath Studio ni Excel COM, por lo que no se ha ejecutado el robot completo.
