# Dashboard de Construcción BESS

Aplicación Streamlit para revisar el avance de construcción. Cada persona que use la aplicación debe cargar su propio archivo `.xlsm` o `.xlsx` desde el panel lateral.

## Archivos para publicar

- `Dashboard.py`
- `requirements.txt`
- `.streamlit/config.toml`

No publiques el archivo de avance del proyecto. `.gitignore` excluye archivos Excel, el entorno virtual y los secretos locales.

## Despliegue en Streamlit Community Cloud

1. Sube los archivos indicados a un repositorio de GitHub.
2. En Streamlit Community Cloud, crea una aplicación desde ese repositorio.
3. Usa `Dashboard.py` como archivo principal y selecciona Python 3.14.
4. Despliega la aplicación y establece su acceso como público.
5. Comparte la URL `*.streamlit.app` que asigne el servicio.

