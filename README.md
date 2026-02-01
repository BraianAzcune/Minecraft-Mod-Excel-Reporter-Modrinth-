# Minecraft Mod Excel Reporter (Modrinth)

Script en Python que lee un `instance.json` de ATLauncher y genera un Excel con los mods, usando **solo** metadata de Modrinth. El resultado incluye una tabla filtrable y un listado único de categorías.

## Requisitos

- Python 3.9+
- `openpyxl`

Instalar dependencia:

```powershell
pip install openpyxl
```

## Uso

```powershell
python minecraft_mods_report_to_excel.py recursos\instance.json
```

## Output

Se genera un Excel en la misma carpeta del `instance.json` con el nombre:

```
Mods {version}.xlsx
```

Ejemplo:

```
Mods 1.21.4.xlsx
```

## Columnas incluidas

- Mod Name
- Description
- Detail (modrinthProject.body)
- Category (categorías concatenadas con `;`)
- Links (link a Modrinth cuando existe)
- File Name
- Updated At

## Reglas

- **No se descartan mods** sin `modrinthProject`.
- Solo se usa metadata de `modrinthProject` (sin CurseForge ni otras fuentes).
- Si no hay `modrinthProject`:
  - Description: `No Modrinth source available (untrusted source)`
  - Detail: vacío
  - Links: `No Modrinth` (sin hipervínculo)
  - Updated At: vacío
  - Category: vacío

## Extras

- Se agrega un listado único de todas las categorías a la derecha de la tabla (por defecto en columna K).
- La altura de las filas es fija para evitar que `Detail` expanda la tabla.

