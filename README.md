
# Valorador Automático de Currículum – UCCuyo (Streamlit)
Subí un CV (PDF o DOCX) y el sistema calcula automáticamente los puntajes por sección y la categoría.

## Archivos
- `app.py`: aplicación principal (Streamlit)
- `config_valorador.yaml`: configuración editable (topes, regex, calibración)
- `requirements.txt`, `runtime.txt`

## Despliegue
1) Subí estos archivos al repositorio en GitHub (raíz).
2) En Streamlit Cloud: New app → repo → branch → **app.py** → Deploy.
3) Si ves Python 3.13 en logs, asegurá `runtime.txt` con `python-3.11`.
