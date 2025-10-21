
import streamlit as st
import re, json, io
import pandas as pd
from docx import Document as DocxDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH

# PDF opcional
try:
    import pdfplumber
    HAVE_PDF = True
except Exception:
    HAVE_PDF = False

# === Rangos de categoría ===
CATEGORIA_RANGOS = [
    ("I – Investigador Superior",      1000, 2000),
    ("II – Investigador Principal",      500,  999),
    ("III – Investigador Independiente", 300,  499),
    ("IV – Investigador Adjunto",        100,  299),
    ("V – Investigador Asistente",         1,   99),
    ("VI – Becario de Iniciación",         0,    0),
]

def obtener_categoria(total):
    for nombre, minimo, maximo in CATEGORIA_RANGOS:
        if minimo <= total <= maximo:
            return nombre
    return "Sin categoría"

st.set_page_config(page_title="Valorador de CV - UCCuyo (DOCX/PDF)", layout="wide")
st.title("Universidad Católica de Cuyo — Valorador de CV Docente")
st.caption("Incluye exportación a Excel y Word + categoría automática según puntaje total.")

@st.cache_data
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

criteria = load_json("criteria.json")

def extract_text_docx(file):
    doc = DocxDocument(file)
    text = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            text += "\n" + " | ".join(c.text for c in row.cells)
    return text

def extract_text_pdf(file):
    if not HAVE_PDF:
        raise RuntimeError("Instalá pdfplumber: pip install pdfplumber")
    chunks = []
    with pdfplumber.open(file) as pdf:
        for p in pdf.pages:
            chunks.append(p.extract_text() or "")
    return "\n".join(chunks)

def match_count(pattern, text):
    return len(re.findall(pattern, text, re.IGNORECASE)) if pattern else 0

def clip(v, cap):
    return min(v, cap) if cap else v

uploaded = st.file_uploader("Cargar CV (.docx o .pdf)", type=["docx", "pdf"])
if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    try:
        raw_text = extract_text_docx(uploaded) if ext == "docx" else extract_text_pdf(uploaded)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.success(f"Archivo cargado: {uploaded.name}")
    with st.expander("Ver texto extraído (debug)"):
        st.text_area("Texto", raw_text, height=220)

    results = {}
    total = 0.0

    for section, cfg in criteria["sections"].items():
        st.markdown(f"### {section}")
        rows = []
        subtotal_raw = 0.0
        for item, icfg in cfg.get("items", {}).items():
            c = match_count(icfg.get("pattern", ""), raw_text)
            pts = clip(c * icfg.get("unit_points", 0), icfg.get("max_points", 0))
            rows.append({"Ítem": item, "Ocurrencias": c, "Puntaje (tope ítem)": pts, "Tope ítem": icfg.get("max_points", 0)})
            subtotal_raw += pts
        df = pd.DataFrame(rows)
        subtotal = clip(subtotal_raw, cfg.get("max_points", 0))
        st.dataframe(df, use_container_width=True)
        st.info(f"Subtotal {section}: {subtotal} / máx {cfg.get('max_points', 0)}")
        results[section] = {"df": df, "subtotal": subtotal}
        total += subtotal

    categoria = obtener_categoria(total)

    st.markdown("---")
    st.subheader("Puntaje total y categoría")
    st.metric("Total acumulado", f"{total:.1f}")
    st.metric("Categoría alcanzada", categoria)

    # Exportaciones
    st.markdown("---")
    st.subheader("Exportar resultados")

    # Excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        for sec, data in results.items():
            data["df"].to_excel(writer, sheet_name=sec[:31], index=False)
        resumen = pd.DataFrame({"Sección": list(results.keys()),
                                "Subtotal": [results[s]["subtotal"] for s in results]})
        resumen.loc[len(resumen)] = ["TOTAL", resumen["Subtotal"].sum()]
        resumen.loc[len(resumen)] = ["CATEGORÍA", categoria]
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
    st.download_button("Descargar Excel", data=out.getvalue(),
                       file_name="valoracion_cv.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

    # Word
    def export_word(results, total, categoria):
        doc = DocxDocument()
        p = doc.add_paragraph("Universidad Católica de Cuyo — Secretaría de Investigación")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Informe de valoración de CV").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("")
        doc.add_paragraph(f"Puntaje total: {total:.1f}")
        doc.add_paragraph(f"Categoría alcanzada: {categoria}")
        for sec, data in results.items():
            doc.add_heading(sec, level=2)
            df = data["df"]
            if df.empty:
                doc.add_paragraph("Sin ítems detectados.")
            else:
                tbl = doc.add_table(rows=1, cols=len(df.columns))
                hdr = tbl.rows[0].cells
                for i, c in enumerate(df.columns):
                    hdr[i].text = str(c)
                for _, row in df.iterrows():
                    cells = tbl.add_row().cells
                    for i, c in enumerate(df.columns):
                        cells[i].text = str(row[c])
            doc.add_paragraph(f"Subtotal sección: {data['subtotal']:.1f}")
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    st.download_button("Descargar informe Word",
                       data=export_word(results, total, categoria),
                       file_name="informe_valoracion_cv.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True)
else:
    st.info("Subí un archivo para iniciar la valoración.")
