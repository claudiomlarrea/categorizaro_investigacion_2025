
# -*- coding: utf-8 -*-
import io, os, re, sys
from datetime import datetime
from typing import Dict, Tuple

import streamlit as st
import pandas as pd
import numpy as np

# ---- Config ----
try:
    import yaml
except Exception as e:
    yaml = None

st.set_page_config(page_title="Valorador Automático de Currículum – UCCuyo", page_icon="📊", layout="wide")

def load_config(path="config_valorador.yaml"):
    if yaml is None:
        st.error("Falta la librería PyYAML. Agrega 'pyyaml' a requirements.txt")
        return None
    if not os.path.exists(path):
        st.warning("⚠️ No se encontró config_valorador.yaml. Se usarán valores por defecto.")
        return {
            "secciones_max": {"formacion":450,"cargos":350,"cyt":500,"producciones":500,"otros":200},
            "calibracion": {"intercept": -401.07,"form": -0.6246,"cargos": 10.0202,"cyt": 18.8741,"prod": 16.1052,"otros": 0.0},
            "detectores_topes": {
                "articulos_referato": 10,"articulos_sin_referato": 8,"libros": 4,
                "capitulos": 6,"eventos": 10,"cargos_docentes": 4,"cargos_gestion": 3,"premios": 5
            },
            "regex": {
                "doctorado": "(doctor(a|ado) en|ph.?d)",
                "maestria": "(maestr(ía|ia)|mag(í|i)ster)",
                "especializacion": "especializaci(ó|o)n",
                "diplomatura": "diplomatur(a|as)",
                "articulo_referato": "(revista|issn|scopus|wos|jcr|doi|q[1-4])",
                "articulo_sin_referato": "(art(í|i)culo|paper)",
                "libro": "(libro|isbn)",
                "capitulo": "(cap(í|i)tulo de libro)",
                "proyecto": "(proyecto|p.i.c.t|p.i.c.o|f.o.n.c.y.t|f.o.n.s.e.c.y.t)",
                "congreso": "(congreso|jornada|simposio|encuentro)",
                "beca": "(becario|beca doctoral|beca posdoctoral)",
                "premio": "(premio|menci(ó|o)n honor(í|i)fica)"
            }
        }
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

CFG = load_config()

SECCIONES_MAX = CFG["secciones_max"]
CAL_COEF = CFG["calibracion"]

CATEGORIA_RANGOS = [
    ("I – Investigador Superior", 1000, 2000),
    ("II – Investigador Principal", 500, 999),
    ("III – Investigador Independiente", 300, 499),
    ("IV – Investigador Adjunto", 100, 299),
    ("V – Investigador Asistente", 1, 99),
    ("VI – Becario de Iniciación", 0, 0),
]

# ---- Lectura de CV (PDF / DOCX) ----
def read_pdf_text(file) -> str:
    try:
        from pypdf import PdfReader
    except Exception:
        st.error("Falta 'pypdf' en requirements.txt para leer PDFs.")
        return ""
    try:
        reader = PdfReader(file)
        out = []
        for p in reader.pages:
            try:
                out.append(p.extract_text() or "")
            except Exception:
                pass
        return "\n".join(out)
    except Exception as e:
        st.exception(e)
        return ""

def read_docx_text(file) -> str:
    try:
        from docx import Document as DocxDocument
    except Exception:
        st.error("Falta 'python-docx' en requirements.txt para leer DOCX.")
        return ""
    try:
        doc = DocxDocument(file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        st.exception(e)
        return ""

def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s.lower()).strip()

# ---- Detectores con regex + topes ----
def parse_counts_from_text(text: str) -> Dict[str, int]:
    t = normalize(text)
    R = CFG["regex"]
    TOP = CFG["detectores_topes"]
    counts = {k:0 for k in [
        "doctorado","maestria","especializacion","diplomatura","posdoctorado",
        "cursos_posgrado","idiomas","estancias","grados_extra",
        "doc_titular","doc_asoc","doc_adj","doc_jtp","doc_posgrad",
        "gest_rector","gest_vice","gest_dec","gest_sec","gest_coord","oc",
        "dir_inv","dir_tes","bec","dir_p","codir_p","part_p","coord_lin",
        "tut","vinc","even","eg","ep","eprog","erev","einst","redes","prof",
        "art_ref","art_sin","libros","capitulos","doc_trab","pat_soft","procesos","serv_tec","informes",
        "redes2","org_ev","gest_ed","prem_int","prem_nac","menc"
    ]}

    # Formación (heurísticas básicas)
    counts["doctorado"] = len(re.findall(R["doctorado"], t))
    counts["maestria"] = len(re.findall(R["maestria"], t))
    counts["especializacion"] = len(re.findall(R["especializacion"], t))
    counts["diplomatura"] = len(re.findall(R["diplomatura"], t))
    counts["posdoctorado"] = len(re.findall(r"pos(doc|doctorado)", t))
    counts["cursos_posgrado"] = len(re.findall(r"(curso(s)? de (pos|post)grado|seminario de posgrado)", t))
    counts["idiomas"] = len(re.findall(r"(toefl|ielts|cambridge|dele|dalf|celu|b2|c1|c2)", t))
    counts["estancias"] = len(re.findall(r"(estancia|pasant(í|i)a|sabbatical)", t))
    multigrado = len(re.findall(r"(licenciado|ingeniero|abogado|m(é|e)dico|contador|arquitecto)", t))
    counts["grados_extra"] = 30 if multigrado >= 2 else 0

    # Docencia / Gestión (capadas)
    counts["doc_titular"] = min(len(re.findall(r"prof(\.|esor)? titular", t)), TOP["cargos_docentes"])
    counts["doc_asoc"]    = min(len(re.findall(r"prof(\.|esor)? asociado", t)), TOP["cargos_docentes"])
    counts["doc_adj"]     = min(len(re.findall(r"prof(\.|esor)? adjunto", t)), TOP["cargos_docentes"])
    counts["doc_jtp"]     = min(len(re.findall(r"(jtp|jefe de trabajos pr(á|a)cticos|ayudante)", t)), TOP["cargos_docentes"])
    counts["doc_posgrad"] = min(len(re.findall(r"(docencia|curso) de posgrado", t)), TOP["cargos_docentes"])

    counts["gest_rector"] = 1 if re.search(r"\brector\b", t) else 0
    counts["gest_vice"]   = 1 if re.search(r"\bvicerrector\b", t) else 0
    counts["gest_dec"]    = min(len(re.findall(r"(decano|director (de|del) (facultad|departamento|instituto))", t)), TOP["cargos_gestion"])
    counts["gest_sec"]    = min(len(re.findall(r"secretar(i|í)a (acad(é|e)mica|investigaci(ó|o)n|extensi(ó|o)n)", t)), TOP["cargos_gestion"])
    counts["gest_coord"]  = min(len(re.findall(r"(coordinador(a)?|responsable)", t)), TOP["cargos_gestion"])
    counts["oc"]          = min(len(re.findall(r"(comisi(ó|o)n|consejo|representante)", t)), TOP["cargos_gestion"])

    # CyT
    counts["dir_inv"] = len(re.findall(r"(direcci(ó|o)n|co-?direcci(ó|o)n) de (investigadores|becarios doctorales)", t))
    counts["dir_tes"] = len(re.findall(r"(direcci(ó|o)n|co-?direcci(ó|o)n) de (tesis|tesistas)", t))
    counts["bec"]     = len(re.findall(r"\bbecario(s)?\b", t))

    counts["dir_p"]   = len(re.findall(r"(dirigi(ó|o)|direcci(ó|o)n) (de )?proyecto(s)?", t))
    counts["codir_p"] = len(re.findall(r"co-?direcci(ó|o)n (de )?proyecto(s)?", t))
    counts["part_p"]  = len(re.findall(r"participaci(ó|o)n en proyecto(s)?", t))
    counts["coord_lin"]=len(re.findall(r"(coordinaci(ó|o)n) (de )?l(í|i)nea(s)? (interdisciplinaria(s)?)", t))

    counts["tut"]  = len(re.findall(r"tutor(í|i)a(s)? (de )?(pasant(í|i)as|pr(á|a)cticas)", t))
    counts["vinc"] = len(re.findall(r"(vinculaci(ó|o)n|transferencia tecnol(ó|o)gica)", t))
    counts["even"] = min(len(re.findall(R["congreso"], t)), TOP["eventos"])

    counts["eg"]   = len(re.findall(r"tribunal (de )?tesis (de )?grado", t))
    counts["ep"]   = len(re.findall(r"tribunal (de )?tesis (de )?posgrado", t))
    counts["eprog"]= len(re.findall(r"evaluaci(ó|o)n (de )?(programas|proyectos)", t))
    counts["erev"] = len(re.findall(r"(reviewer|evaluaci(ó|o)n) (de )?(art(í|i)culos|revistas|congresos)", t))
    counts["einst"]= len(re.findall(r"evaluaci(ó|o)n institucional|organismo evaluador", t))

    counts["redes"]= len(re.findall(r"(red(es)? acad(é|e)micas|comit(é|e)s|sociedad cient(í|i)fica)", t))
    counts["prof"] = len(re.findall(r"ejercicio profesional", t))

    # Producciones
    art_ref_hits = re.findall(R["articulo_referato"], t)
    art_sin_hits = re.findall(R["articulo_sin_referato"], t)
    counts["art_ref"] = min(len(art_ref_hits), TOP["articulos_referato"])
    # Evitar doble conteo
    counts["art_sin"] = min(max(len(art_sin_hits) - counts["art_ref"], 0), TOP["articulos_sin_referato"])

    counts["libros"]    = min(len(re.findall(R["libro"], t)), TOP["libros"])
    counts["capitulos"] = min(len(re.findall(R["capitulo"], t)), TOP["capitulos"])
    counts["doc_trab"]  = len(re.findall(r"(documento de trabajo|working paper|informe t(é|e)cnico)", t))

    counts["pat_soft"] = len(re.findall(r"(patente|modelo de utilidad|software registrado)", t))
    counts["procesos"] = len(re.findall(r"(proceso|innovaci(ó|o)n)", t))
    counts["serv_tec"] = len(re.findall(r"servicio(s)? tecn(ó|o)logico(s)?", t))
    counts["informes"] = len(re.findall(r"informe(s)? t(é|e)cnico(s)?", t))

    counts["redes2"] = len(re.findall(r"red(es)? (acad(é|e)micas|cient(í|i)ficas)", t))
    counts["org_ev"] = len(re.findall(r"(organizador|comit(é|e) organizador)", t))
    counts["gest_ed"]= len(re.findall(r"(editor|comit(é|e) editorial|gesti(ó|o)n editorial)", t))
    counts["prem_int"]=min(len(re.findall(r"premio(s)? (internacional(es)?)", t)), TOP["premios"])
    counts["prem_nac"]=min(len(re.findall(r"premio(s)? (nacional(es)?)", t)), TOP["premios"])
    counts["menc"]    =min(len(re.findall(r"menci(ó|o)n(es)? honor(í|i)fica(s)?", t)), TOP["premios"])

    return counts

# ---- Ponderación y topes por sección ----
def compute_scores(c: Dict[str,int]) -> Dict[str,int]:
    # Formación
    form_total = 0
    form_total += min(c["doctorado"] * 250, 375)
    form_total += min(c["maestria"] * 150, 225)
    form_total += min(c["especializacion"] * 70, 105)
    form_total += min(c["diplomatura"] * 50, 100)
    form_total += min(c["posdoctorado"] * 100, 100)
    form_total += min(c["cursos_posgrado"] * 5, 75)
    form_total += min(c["idiomas"] * 10, 50)
    form_total += min(c["estancias"] * 20, 60)
    form_total += min(c["grados_extra"], 30)
    form_total = min(int(form_total), SECCIONES_MAX["formacion"])

    # Cargos
    docencia = min(c["doc_titular"] * 30, 150) + min(c["doc_asoc"] * 25, 125) +                min(c["doc_adj"] * 20, 100) + min(c["doc_jtp"] * 10, 50) +                min(c["doc_posgrad"] * 20, 100)
    docencia = min(docencia, 300)
    gestion = (100 if c["gest_rector"] else 0) + (80 if c["gest_vice"] else 0) +               min(c["gest_dec"], 60) + min(c["gest_sec"], 60) + min(c["gest_coord"], 40)
    gestion = min(gestion, 200)
    otros_cargos = min(c["oc"] * 10, 50)
    cargos = min(docencia + gestion + otros_cargos, SECCIONES_MAX["cargos"])

    # CyT
    form_cyt = min(c["dir_inv"], 90) + min(c["dir_tes"], 50) + min(c["bec"] * 20, 40)
    form_cyt = min(form_cyt, 150)
    proyectos = min(c["dir_p"] * 50, 150) + min(c["codir_p"] * 30, 90) +                 min(c["part_p"] * 20, 60) + min(c["coord_lin"] * 20, 20)
    proyectos = min(proyectos, 150)
    extension = min(c["tut"] * 10, 20) + min(c["vinc"] * 15, 45) + min(c["even"], 100)
    extension = min(extension, 60)
    evaluacion = min(c["eg"] * 5, 20) + min(c["ep"] * 10, 30) + min(c["eprog"] * 10, 30) +                  min(c["erev"] * 10, 30) + min(c["einst"] * 10, 30)
    evaluacion = min(evaluacion, 100)
    otras_cyt = min(c["redes"] * 20, 60) + min(c["prof"] * 5, 20)
    otras_cyt = min(otras_cyt, 60)
    cyt = min(form_cyt + proyectos + extension + evaluacion + otras_cyt, SECCIONES_MAX["cyt"])

    # Producciones y servicios
    publicaciones = min(c["art_ref"] * 20, 140) + min(c["art_sin"] * 10, 80) +                     min(c["libros"] * 40, 80) + min(c["capitulos"] * 20, 60) +                     min(c["doc_trab"] * 10, 30)
    publicaciones = min(publicaciones, 300)
    desarrollos = min(c["pat_soft"] * 30, 60) + min(c["procesos"] * 20, 60)
    desarrollos = min(desarrollos, 100)
    servicios = min(c["serv_tec"] * 20, 40) + min(c["informes"] * 10, 20)
    servicios = min(servicios, 40)
    prod = min(publicaciones + desarrollos + servicios, SECCIONES_MAX["producciones"])

    # Otros
    redes_gestion = min(c["redes2"] * 10, 30) + min(c["org_ev"] * 20, 60) + min(c["gest_ed"], 60)
    redes_gestion = min(redes_gestion, 150)
    premios = min(c["prem_int"] * 50, 100) + min(c["prem_nac"] * 20, 100) + min(c["menc"] * 20, 100)
    premios = min(premios, 100)
    otros = min(redes_gestion + premios, SECCIONES_MAX["otros"])

    total = form_total + cargos + cyt + prod + otros
    return {"formacion":form_total, "cargos":cargos, "cyt":cyt, "prod":prod, "otros":otros, "total":total}

def determinar_categoria(total: int) -> str:
    for nombre, lo, hi in CATEGORIA_RANGOS:
        if lo <= total <= hi or (nombre == "I – Investigador Superior" and total > hi):
            return nombre
    return "VI – Becario de Iniciación"

def tag_categoria(cat: str) -> str:
    base = {"I": ("#065f46", "#ecfdf5"), "II": ("#064e3b", "#ecfdf5"),
            "III": ("#1e40af", "#eff6ff"), "IV": ("#7c2d12", "#fffbeb"),
            "V": ("#7f1d1d", "#fef2f2"), "VI": ("#334155", "#f1f5f9")}
    key = cat.split(" ")[0]
    fg, bg = base.get(key, ("#334155", "#f1f5f9"))
    return f"<span style='background:{bg}; color:{fg}; padding:6px 10px; border-radius:999px; font-weight:600;'>{cat}</span>"

# ---- UI ----
st.markdown(
    '''
    <div style="padding: 18px; border-radius: 14px; background: #0f172a; color: white; text-align:center;">
      <h2 style="margin: 0 0 6px 0;">Universidad Católica de Cuyo</h2>
      <div style="opacity:.9;">Secretaría de Investigación · <b>Valorador Automático de Currículum</b></div>
    </div>
    ''',
    unsafe_allow_html=True
)

colA, colB = st.columns([2,1])
with colA:
    uploaded = st.file_uploader("Subí el CV (PDF o Word .docx)", type=["pdf","docx"])
with colB:
    nombre = st.text_input("Nombre y Apellido", "")
    unidad = st.text_input("Unidad Académica / Instituto", "")
    fecha_eval = st.date_input("Fecha de evaluación", datetime.now())
    usar_calibracion = st.checkbox("Usar calibración empírica", value=True)

if not uploaded:
    st.info("Cargá un CV para valorar automáticamente.")
    st.stop()

# Extraer texto
try:
    if uploaded.name.lower().endswith(".pdf"):
        raw_text = read_pdf_text(uploaded)
    else:
        raw_text = read_docx_text(uploaded)
except Exception as e:
    st.error("No se pudo leer el archivo.")
    st.exception(e)
    st.stop()

if not raw_text.strip():
    st.warning("No se extrajo texto. Si el PDF es escaneado, se requiere OCR (no incluido).")
    st.stop()

with st.expander("Vista previa del texto extraído", expanded=False):
    st.text_area("Texto", raw_text[:20000], height=220)

# Detectar y puntuar
counts = parse_counts_from_text(raw_text)
scores = compute_scores(counts)
total_base = scores["total"]

# Calibración por secciones
def total_calibrado(form, cargos, cyt, prod, otros):
    t = (CAL_COEF["intercept"] +
         CAL_COEF["form"]   * form +
         CAL_COEF["cargos"] * cargos +
         CAL_COEF["cyt"]    * cyt +
         CAL_COEF["prod"]   * prod +
         CAL_COEF["otros"]  * otros)
    return max(0, min(2000, int(round(t))))

if usar_calibracion:
    total = total_calibrado(scores["formacion"], scores["cargos"], scores["cyt"], scores["prod"], scores["otros"])
else:
    total = total_base

categoria = determinar_categoria(total)

# Resultados
st.success("Valoración automática completada ✅")
st.markdown("---")
c1,c2,c3,c4 = st.columns(4)
c1.metric("Formación", scores["formacion"])
c2.metric("Cargos", scores["cargos"])
c3.metric("CyT", scores["cyt"])
c4.metric("Producciones/Servicios", scores["prod"])
c5,c6 = st.columns([1,2])
c5.metric("Otros antecedentes", scores["otros"])
c6.metric("TOTAL (base)", total_base)
st.markdown(f"**TOTAL {'calibrado' if usar_calibracion else 'sin calibrar'}:** {total}")
st.markdown(f"**Categoría alcanzada:** {tag_categoria(categoria)}", unsafe_allow_html=True)

# Tabla
resumen = pd.DataFrame([
    {"Sección":"Formación académica y complementaria","Puntaje":scores["formacion"],"Tope":SECCIONES_MAX["formacion"]},
    {"Sección":"Cargos (docencia, gestión y otros)","Puntaje":scores["cargos"],"Tope":SECCIONES_MAX["cargos"]},
    {"Sección":"Ciencia y Tecnología","Puntaje":scores["cyt"],"Tope":SECCIONES_MAX["cyt"]},
    {"Sección":"Producciones y servicios","Puntaje":scores["prod"],"Tope":SECCIONES_MAX["producciones"]},
    {"Sección":"Otros antecedentes","Puntaje":scores["otros"],"Tope":SECCIONES_MAX["otros"]},
    {"Sección":"TOTAL (base)","Puntaje":total_base,"Tope":2000},
    {"Sección":"TOTAL (resultado)","Puntaje":total,"Tope":2000},
])
st.dataframe(resumen, use_container_width=True)

# Exportables
colx1, colx2 = st.columns(2)
with colx1:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, index=False, sheet_name="Resultados")
        pd.DataFrame(sorted(counts.items()), columns=["Ítem","Conteo"]).to_excel(writer, index=False, sheet_name="Detectores")
    st.download_button("⬇️ Descargar Excel", out.getvalue(),
        file_name=f"Valorador_Investigadores_Auto_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with colx2:
    # Exportar Word con importación perezosa
    try:
        from docx import Document as DocxDocument
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        doc = DocxDocument()
        style = doc.styles["Normal"]; style.font.name = "Times New Roman"; style.font.size = Pt(11)
        h = doc.add_paragraph("Universidad Católica de Cuyo – Secretaría de Investigación")
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER; h.runs[0].bold = True
        doc.add_paragraph("Valorador Automático de Currículum – Categorización de Investigadores")
        info = doc.add_paragraph(f"Postulante: {nombre if nombre else '-'} | Unidad: {unidad if unidad else '-'} | Fecha: {fecha_eval.strftime('%Y-%m-%d')}")
        info.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.add_paragraph(f"Categoría alcanzada: {categoria}")
        table = doc.add_table(rows=1, cols=3); hdr = table.rows[0].cells
        hdr[0].text = "Sección"; hdr[1].text = "Puntaje"; hdr[2].text = "Tope"
        for _, row in resumen.iterrows():
            r = table.add_row().cells
            r[0].text = str(row["Sección"]); r[1].text = str(int(row["Puntaje"])); r[2].text = str(int(row["Tope"]))
        doc.add_paragraph("")
        doc.add_paragraph("Auditoría (detectores):")
        for k, v in sorted(counts.items()):
            doc.add_paragraph(f"• {k}: {v}")
        wb = io.BytesIO(); doc.save(wb)
        st.download_button("⬇️ Descargar Informe Word", wb.getvalue(),
            file_name=f"Valorador_Investigadores_Auto_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        st.warning("Para exportar Word, agrega 'python-docx' a requirements.txt")
        st.exception(e)
