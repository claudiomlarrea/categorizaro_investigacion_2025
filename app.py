
import streamlit as st
import re, json, io
import pandas as pd
from docx import Document

st.set_page_config(page_title='Valorador de CV - UCCuyo', layout='wide')
st.title('Universidad Católica de Cuyo — Valorador de CV Docente')

@st.cache_data
def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

criteria = load_json('criteria.json')
patterns = load_json('patterns.json')

def extract_text(file):
    doc = Document(file)
    text = '\n'.join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += '\n' + ' | '.join(c.text for c in row.cells)
    return text

def match_count(pattern, text):
    return len(re.findall(pattern, text, re.IGNORECASE))

def clip(value, cap):
    return min(value, cap)

cv = st.file_uploader('Cargar CV (.docx)', type='docx')
if cv:
    texto = extract_text(cv)
    total = 0
    for seccion, cfg in criteria['sections'].items():
        st.subheader(seccion)
        datos = []
        subtotal = 0
        for item, icfg in cfg['items'].items():
            c = match_count(icfg['pattern'], texto)
            puntos = clip(c * icfg['unit_points'], icfg['max_points'])
            subtotal += puntos
            datos.append({'Ítem': item, 'Ocurrencias': c, 'Puntaje': puntos})
        df = pd.DataFrame(datos)
        st.dataframe(df, use_container_width=True)
        st.info(f'Subtotal {seccion}: {subtotal} / máx {cfg["max_points"]}')
        total += clip(subtotal, cfg['max_points'])
    st.metric('Puntaje total', total)
