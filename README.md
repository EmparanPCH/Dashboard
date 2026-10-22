# Dashboard de Construcción BESS

Aplicación Streamlit para revisar el avance de construcción. El dashboard carga automáticamente `dashboard_data.xlsx`; las personas que lo visiten no necesitan seleccionar ni cargar un archivo.

## Archivos para publicar

- `Dashboard.py`
- `dashboard_data.xlsx.b64` (extracto codificado de la hoja de construcción)
- `requirements.txt`
- `preparar_datos.py`
- `.streamlit/config.toml`

El archivo original `.xlsm` no se publica. `preparar_datos.py` genera `dashboard_data.xlsx` y su copia `dashboard_data.xlsx.b64`, con solo valores de la hoja `DUNE BESS CONSTR.`, sin macros ni otras hojas. Como el repositorio y la app son públicos, los datos incluidos en ese extracto y mostrados por el dashboard serán visibles para cualquiera que acceda.

`.gitignore` excluye el libro original y el extracto Excel local, además del entorno virtual y los secretos. La copia codificada `.b64` se publica para que Streamlit Cloud pueda cargar los datos sin solicitar archivos a los visitantes.

## Actualizar los datos

Ejecuta este comando con la nueva versión del anexo:

```powershell
python preparar_datos.py "C:\ruta\al\Anexo.xlsx"
```

Actualiza `dashboard_data.xlsx.b64` en el repositorio con la nueva copia generada para actualizar la app.

## Despliegue en Streamlit Community Cloud

1. Sube los archivos indicados y `dashboard_data.xlsx.b64` a un repositorio de GitHub. No es necesario subir `dashboard_data.xlsx`.
2. En Streamlit Community Cloud, crea una aplicación desde ese repositorio.
3. Usa `Dashboard.py` como archivo principal y selecciona Python 3.14.
4. Despliega la aplicación y establece su acceso como público.
5. Comparte la URL `*.streamlit.app` que asigne el servicio.

