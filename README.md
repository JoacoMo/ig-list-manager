# IG List Manager

Programa de consola en Python para manejar una base de usuarios de Instagram: agregar, buscar, eliminar, comparar listas pegadas contra la base y exportar/importar Excel.

Está pensado para no perder datos nunca: guardado atómico, backups automáticos y frenos de seguridad antes de cualquier borrado grande.

## Requisitos

- Python 3.8 o superior
- [`openpyxl`](https://pypi.org/project/openpyxl/) (opcional, solo para Excel)

```bash
pip install -r requirements.txt
```

## Uso

```bash
python ig_list_manager.py
```

Aparece un menú numerado. En los modos de "pegar lista", pegá los usuarios en cualquier formato y cuando termines escribí `FIN` y dale Enter.

La primera vez, si no existe `usuarios_ig.txt`, arranca con la base vacía y lo crea al guardar.

## Opciones del menú

| # | Opción | Qué hace |
|---|--------|----------|
| 1 | Buscar un usuario | Te dice si está o no en la base. |
| 2 | Buscar varios | Varios usuarios en una línea; separa los que están de los que no. |
| 3 | Agregar un usuario | Lo agrega y guarda al instante. |
| 4 | Agregar varios | Varios usuarios en una sola línea. |
| 5 | Pegar lista para agregar | Pegás una lista larga (hasta `FIN`) y se suma a la base. |
| 6 | Eliminar usuario(s) | Si son varios, pide confirmación. |
| 7 | Ver toda la lista | Si son más de 500, ofrece exportar a Excel en vez de llenar la pantalla. |
| 8 | Desordenar lista temporal | Mezcla una lista pegada. No toca la base. |
| 9 | Salir | Guarda y cierra. |
| 10 | Filtrar Lista 1 | Lista 1 menos los que estén en la Lista 2. No toca la base. |
| 11 | Sumar 2 listas | Une dos listas sin repetidos. No toca la base. |
| 12 | Comparar base vs lista | Te devuelve la lista pegada sin los que ya tenés, respetando el orden. |
| 13 | Filtrar usuarios | Pegás N usuarios y te devuelve solo los que faltan en la base. |
| 14 | Exportar a Excel | Genera `usuarios_ig.xlsx`. |
| 15 | Importar desde archivo | Acepta `.xlsx`, `.csv`, `.json` y `.txt` (podés arrastrar el archivo a la consola). |
| 16 | Backups | Lista los últimos backups y te deja fusionar uno con la base (suma lo que falte, no borra). |
| 17 | Reparar base | Reescribe el archivo con un usuario limpio por línea. |

## Qué acepta al pegar

El texto se limpia solo, así que podés pegar casi cualquier cosa:

- usuarios separados por comas, espacios, tabs, comillas o saltos de línea,
- links: `https://www.instagram.com/pepito/` → `pepito`,
- `@usuario` → `usuario`.

Se descartan (y te avisa cuáles) los textos que no cumplen las reglas de Instagram: solo letras, números, `.` y `_`, hasta 30 caracteres. Todo se guarda en minúsculas.

## Cómo protege los datos

- **La base vive al lado del script.** No importa desde qué carpeta lo ejecutes, siempre usa el `usuarios_ig.txt` que está junto a `ig_list_manager.py`.
- **Guardado atómico.** Escribe primero un `.tmp` y recién al final reemplaza el archivo. Si se corta la luz a mitad de camino, la base anterior queda intacta.
- **Backup automático** en `backups/` antes de cada guardado (solo si hubo cambios). Se conservan los últimos 40.
- **Fusión con lo que hay en disco.** Antes de guardar relee el archivo y lo combina con lo que tiene en memoria, así dos ventanas abiertas no se pisan.
- **Frenos de seguridad.** Cancela cualquier guardado que deje la base vacía o que borre usuarios que no pediste eliminar, y pide confirmación si la base se fuera a achicar de golpe más de un 10 % (y más de 50 usuarios).

## Archivos

```
ig-list-manager/
├── ig_list_manager.py   # el programa
├── requirements.txt
├── README.md
├── usuarios_ig.txt      # la base (se crea sola, no se sube a git)
├── usuarios_ig.xlsx     # export de la opción 14 (no se sube a git)
└── backups/             # backups automáticos (no se suben a git)
```

La base y los backups tienen nombres de cuentas reales, por eso están en `.gitignore` y nunca se suben al repositorio.
