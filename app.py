import streamlit as st
import re, json, io
from docx import Document
import pandas as pd

st.set_page_config(page_title='Valorador de CV - UCCuyo', layout='wide')
st.title('Universidad Católica de Cuyo — Valorador de CV Docente (PDF/DOCX)')

try:
    import pdfplumber
    HAVE_PDF = True
except Exception:
    HAVE_PDF = False

@st.cache_data
def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

criteria = load_json('criteria.json')
patterns = load_json('patterns.json')

def extract_text(file, ext):
    if ext == 'docx':
        doc = Document(file)
        text = '\n'.join(p.text for p in doc.paragraphs)
        for t in doc.tables:
            for row in t.rows:
                text += '\n' + ' | '.join(c.text for c in row.cells)
        return text
    elif ext == 'pdf':
        if not HAVE_PDF:
            st.error('Instalá pdfplumber para leer PDF: pip install pdfplumber')
            return ''
        import pdfplumber
        text = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text.append(page.extract_text() or '')
        return '\n'.join(text)
    return ''

def match_count(pattern, text):
    return len(re.findall(pattern, text, re.IGNORECASE)) if pattern else 0

def clip(v, cap): return min(v, cap)

cv = st.file_uploader('Cargar CV (.docx o .pdf)', type=['docx', 'pdf'])
if cv:
    ext = cv.name.split('.')[-1].lower()
    texto = extract_text(cv, ext)
    st.success(f'Archivo cargado: {cv.name}')
    total = 0
    for s, cfg in criteria['sections'].items():
        st.subheader(s)
        data = []
        subtotal = 0
        for item, icfg in cfg['items'].items():
            c = match_count(icfg['pattern'], texto)
            pts = clip(c * icfg['unit_points'], icfg['max_points'])
            subtotal += pts
            data.append({'Ítem': item, 'Ocurrencias': c, 'Puntaje': pts})
        df = pd.DataFrame(data)
        st.dataframe(df, use_container_width=True)
        subtotal = clip(subtotal, cfg['max_points'])
        st.info(f'Subtotal {s}: {subtotal} / máx {cfg['max_points']}')
        total += subtotal
    st.metric('Puntaje total', total)
