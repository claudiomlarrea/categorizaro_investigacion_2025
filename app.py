# -*- coding: utf-8 -*-
import io
from datetime import datetime

import pandas as pd
import numpy as np
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# Configuración inicial (debe ser la primera llamada a Streamlit)
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Valorador de Currículum Docente – Categorización de Investigadores",
    page_icon="📊",
    layout="wide"
)

# ──────────────────────────────────────────────────────────────────────────────
# PARCHE: importación perezosa de python-docx
# ──────────────────────────────────────────────────────────────────────────────
Document = Pt = WD_ALIGN_PARAGRAPH = None
def _ensure_docx():
    """Carga python-docx solo cuando se necesita exportar a Word."""
    global Document, Pt, WD_ALIGN_PARAGRAPH
    if Document is None:
        from docx import Document  # type: ignore
        from docx.shared import Pt  # type: ignore
        from docx.enum.text import WD_ALIGN_PARAGRAPH  # type: ignore
    return Document, Pt, WD_ALIGN_PARAGRAPH

# ──────────────────────────────────────────────────────────────────────────────
# ENVOLTORIO: mostrar cualquier excepción en UI (evita pantalla negra)
# ──────────────────────────────────────────────────────────────────────────────
try:
    # ====== UI Header ======
    st.markdown(
        """
        <div style="padding: 14px 18px; border-radius: 14px; background: #0f172a; color: white;">
          <h2 style="margin: 0 0 6px 0;">Universidad Católica de Cuyo</h2>
          <div style="opacity: 0.85;">Secretaría de Investigación · Valorador de Currículum Docente</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.write("")

    # ====== Constantes ======
    CATEGORIA_RANGOS = [
        ("I – Investigador Superior", 1500, 2000),
        ("II – Investigador Principal", 1000, 1499),
        ("III – Investigador Independiente", 600, 999),
        ("IV – Investigador Adjunto", 300, 599),
        ("V – Investigador Asistente", 1, 299),
        ("VI – Becario de Iniciación", 0, 0),
    ]

    SECCIONES_MAX = {
        "Formación académica y complementaria": 450,
        "Cargos (docencia, gestión y otros)": 500,
        "Ciencia y Tecnología": 500,
        "Producciones y servicios": 350,
        "Otros antecedentes": 200,
    }

    def determinar_categoria(total: int) -> str:
        for nombre, lo, hi in CATEGORIA_RANGOS:
            if lo <= total <= hi or (nombre == "I – Investigador Superior" and total > hi):
                return nombre
        return "VI – Becario de Iniciación"

    def tag_categoria(cat: str) -> str:
        base = {
            "I": ("#065f46", "#ecfdf5"),
            "II": ("#064e3b", "#ecfdf5"),
            "III": ("#1e40af", "#eff6ff"),
            "IV": ("#7c2d12", "#fffbeb"),
            "V": ("#7f1d1d", "#fef2f2"),
            "VI": ("#334155", "#f1f5f9"),
        }
        key = cat.split(" ")[0]
        fg, bg = base.get(key, ("#334155", "#f1f5f9"))
        return f"<span style='background:{bg}; color:{fg}; padding:6px 10px; border-radius:999px; font-weight:600;'>{cat}</span>"

    # ====== Sidebar: Datos del postulante ======
    st.sidebar.header("Datos del postulante")
    col_p1, col_p2 = st.sidebar.columns(2)
    with col_p1:
        nombre = st.text_input("Nombre y Apellido", "")
    with col_p2:
        unidad = st.text_input("Unidad Académica / Instituto", "")

    col_p3, col_p4 = st.sidebar.columns(2)
    with col_p3:
        categoria_prev = st.selectbox(
            "Categoría previa (opcional)",
            ["(sin especificar)"] + [c[0] for c in CATEGORIA_RANGOS],
            index=0,
        )
    with col_p4:
        fecha_eval = st.date_input("Fecha de evaluación", datetime.now())

    st.sidebar.caption("Complete las secciones con puntajes. Los topes por sección se aplican automáticamente.")
    st.sidebar.write("")

    # ====== 1) Formación académica y complementaria (máx. 450) ======
    st.subheader("1) Formación académica y complementaria (máx. 450)")
    c1, c2, c3, c4 = st.columns(4)
    with c1: doctorado = st.number_input("Doctorado (250 c/u, máx. 375)", 0, value=0)
    with c2: maestria = st.number_input("Maestría (150 c/u, máx. 225)", 0, value=0)
    with c3: especializacion = st.number_input("Especialización (70 c/u, máx. 105)", 0, value=0)
    with c4: diplomatura = st.number_input("Diplomatura >200h (50 c/u, máx. 100)", 0, value=0)
    c5, c6, c7, c8 = st.columns(4)
    with c5: posdoctorado = st.number_input("Posdoctorado (100, máx. 100)", 0, value=0)
    with c6: cursos_posgrado = st.number_input("Cursos >40h (5 c/u, máx. 75)", 0, value=0)
    with c7: idiomas = st.number_input("Idiomas certificados (10 c/u, máx. 50)", 0, value=0)
    with c8: estancias = st.number_input("Estancias I+D (20 c/u, máx. 60)", 0, value=0)
    c9, _ = st.columns(2)
    with c9: grados_extra = st.number_input("Grados adicionales (hasta 30)", 0, value=0)

    formacion_total = min(
        int(
            min(doctorado * 250, 375)
            + min(maestria * 150, 225)
            + min(especializacion * 70, 105)
            + min(diplomatura * 50, 100)
            + min(posdoctorado * 100, 100)
            + min(cursos_posgrado * 5, 75)
            + min(idiomas * 10, 50)
            + min(estancias * 20, 60)
            + min(grados_extra, 30)
        ),
        SECCIONES_MAX["Formación académica y complementaria"],
    )
    st.info(f"Puntaje Formación: **{formacion_total}** / 450")

    # ====== 2) Cargos (docencia, gestión y otros) (máx. 500) ======
    st.subheader("2) Cargos – docencia, gestión y otros (máx. 500)")
    with st.expander("Docencia (hasta 300)", True):
        d1, d2, d3, d4, d5 = st.columns(5)
        with d1: doc_titular = st.number_input("Titular (30/año, máx. 150)", 0, value=0)
        with d2: doc_asoc = st.number_input("Asociado (25/año, máx. 125)", 0, value=0)
        with d3: doc_adj = st.number_input("Adjunto (20/año, máx. 100)", 0, value=0)
        with d4: doc_jtp = st.number_input("JTP/Ayud. (10/año, máx. 50)", 0, value=0)
        with d5: doc_posgrad = st.number_input("Posgrado (20/curso, máx. 100)", 0, value=0)
        docencia_total = (
            min(doc_titular * 30, 150)
            + min(doc_asoc * 25, 125)
            + min(doc_adj * 20, 100)
            + min(doc_jtp * 10, 50)
            + min(doc_posgrad * 20, 100)
        )
        docencia_total = min(docencia_total, 300)
        st.caption(f"Subtotal Docencia: {docencia_total}")

    with st.expander("Gestión (hasta 200)", True):
        g1, g2, g3, g4, g5 = st.columns(5)
        with g1: gest_rector = st.checkbox("Rector (100)")
        with g2: gest_vice = st.checkbox("Vicerrector/Directorio (80)")
        with g3: gest_dec = st.number_input("Decano/Director (hasta 60)", 0, value=0)
        with g4: gest_sec = st.number_input("Secretario Acad./Inv./Ext. (hasta 60)", 0, value=0)
        with g5: gest_coord = st.number_input("Coordinador/Responsable (hasta 40)", 0, value=0)
        gestion_total = (100 if gest_rector else 0) + (80 if gest_vice else 0) + min(gest_dec, 60) + min(gest_sec, 60) + min(gest_coord, 40)
        gestion_total = min(gestion_total, 200)
        st.caption(f"Subtotal Gestión: {gestion_total}")

    with st.expander("Otros cargos (hasta 75)", True):
        oc = st.number_input("Funciones especiales (10 c/u, máx. 50)", 0, value=0)
        otros_cargos_total = min(oc * 10, 50)
        st.caption(f"Subtotal Otros cargos: {otros_cargos_total}")

    cargos_total = min(docencia_total + gestion_total + otros_cargos_total, SECCIONES_MAX["Cargos (docencia, gestión y otros)"])
    st.info(f"Puntaje Cargos: **{cargos_total}** / 500")

    # ====== 3) Ciencia y Tecnología (máx. 500) ======
    st.subheader("3) Ciencia y Tecnología (máx. 500)")
    with st.expander("Formación de recursos humanos en CyT (hasta 150)", True):
        f1, f2, f3 = st.columns(3)
        with f1: dir_inv = st.number_input("Dir./Co-dir. investigadores (máx. 90)", 0, value=0)
        with f2: dir_tes = st.number_input("Dir./Co-dir. tesistas (máx. 50)", 0, value=0)
        with f3: bec = st.number_input("Formación de becarios (20 c/u, máx. 40)", 0, value=0)
        form_cyt_total = min(dir_inv, 90) + min(dir_tes, 50) + min(bec * 20, 40)
        form_cyt_total = min(form_cyt_total, 150)
        st.caption(f"Subtotal Formación CyT: {form_cyt_total}")

    with st.expander("Proyectos de I+D (hasta 150)", True):
        p1, p2, p3, p4 = st.columns(4)
        with p1: dir_p = st.number_input("Dirección (50 c/u, máx. 150)", 0, value=0)
        with p2: codir_p = st.number_input("Co-dirección (30 c/u, máx. 90)", 0, value=0)
        with p3: part_p = st.number_input("Participación (20 c/u, máx. 60)", 0, value=0)
        with p4: coord_lin = st.number_input("Coordinación líneas interdisc. (20, máx. 20)", 0, value=0)
        proyectos_total = min(dir_p * 50, 150) + min(codir_p * 30, 90) + min(part_p * 20, 60) + min(coord_lin * 20, 20)
        proyectos_total = min(proyectos_total, 150)
        st.caption(f"Subtotal Proyectos I+D: {proyectos_total}")

    with st.expander("Extensión (hasta 60)", True):
        e1, e2, e3 = st.columns(3)
        with e1: tut = st.number_input("Tutorías pasantías/prácticas (10 c/u, máx. 20)", 0, value=0)
        with e2: vinc = st.number_input("Vinculación/transferencia (15 c/u, máx. 45)", 0, value=0)
        with e3: even = st.number_input("Eventos científicos (hasta 100)", 0, value=0)
        extension_total = min(tut * 10, 20) + min(vinc * 15, 45) + min(even, 100)
        extension_total = min(extension_total, 60)
        st.caption(f"Subtotal Extensión: {extension_total}")

    with st.expander("Evaluación (hasta 100)", True):
        v1, v2, v3, v4, v5 = st.columns(5)
        with v1: eg = st.number_input("Tribunal grado (5 c/u, máx. 20)", 0, value=0)
        with v2: ep = st.number_input("Tribunal posgrado (10 c/u, máx. 30)", 0, value=0)
        with v3: eprog = st.number_input("Eval. programas/proyectos (10 c/u, máx. 30)", 0, value=0)
        with v4: erev = st.number_input("Eval. revistas/congresos (10 c/u, máx. 30)", 0, value=0)
        with v5: einst = st.number_input("Eval. institucional/organismos (10 c/u, máx. 30)", 0, value=0)
        evaluacion_total = min(eg * 5, 20) + min(ep * 10, 30) + min(eprog * 10, 30) + min(erev * 10, 30) + min(einst * 10, 30)
        evaluacion_total = min(evaluacion_total, 100)
        st.caption(f"Subtotal Evaluación: {evaluacion_total}")

    with st.expander("Otras actividades CyT (hasta 60)", True):
        o1, o2 = st.columns(2)
        with o1: redes = st.number_input("Redes/comités/eventos (20 c/u, máx. 60)", 0, value=0)
        with o2: prof = st.number_input("Ejercicio profesional (5/año, máx. 20)", 0, value=0)
        otras_cyt_total = min(redes * 20, 60) + min(prof * 5, 20)
        otras_cyt_total = min(otras_cyt_total, 60)
        st.caption(f"Subtotal Otras CyT: {otras_cyt_total}")

    cyt_total = min(form_cyt_total + proyectos_total + extension_total + evaluacion_total + otras_cyt_total, SECCIONES_MAX["Ciencia y Tecnología"])
    st.info(f"Puntaje CyT: **{cyt_total}** / 500")

    # ====== 4) Producciones y servicios (máx. 350) ======
    st.subheader("4) Producciones y servicios (máx. 350)")
    with st.expander("Publicaciones (hasta 300)", True):
        pb1, pb2, pb3, pb4, pb5 = st.columns(5)
        with pb1: art_ref = st.number_input("Artículos con referato (20 c/u, máx. 140)", 0, value=0)
        with pb2: art_sin = st.number_input("Artículos sin referato (10 c/u, máx. 80)", 0, value=0)
        with pb3: libros = st.number_input("Libros (40 c/u, máx. 80)", 0, value=0)
        with pb4: capitulos = st.number_input("Capítulos (20 c/u, máx. 60)", 0, value=0)
        with pb5: doc_trab = st.number_input("Docs. trabajo/técnicos (10 c/u, máx. 30)", 0, value=0)
        publicaciones_total = min(art_ref * 20, 140) + min(art_sin * 10, 80) + min(libros * 40, 80) + min(capitulos * 20, 60) + min(doc_trab * 10, 30)
        publicaciones_total = min(publicaciones_total, 300)
        st.caption(f"Subtotal Publicaciones: {publicaciones_total}")

    with st.expander("Desarrollos tecnológicos / organizacionales / socio-comunitarios (hasta 100)", True):
        d1, d2 = st.columns(2)
        with d1: pat_soft = st.number_input("Patentes/modelos/software (30 c/u, máx. 60)", 0, value=0)
        with d2: procesos = st.number_input("Procesos/innovación (20 c/u, máx. 60)", 0, value=0)
        desarrollos_total = min(pat_soft * 30, 60) + min(procesos * 20, 60)
        desarrollos_total = min(desarrollos_total, 100)
        st.caption(f"Subtotal Desarrollos: {desarrollos_total}")

    with st.expander("Servicios (hasta 40)", True):
        s1, s2 = st.columns(2)
        with s1: serv_tec = st.number_input("Servicios CyT (20 c/u, máx. 40)", 0, value=0)
        with s2: informes = st.number_input("Informes técnicos (10 c/u, máx. 20)", 0, value=0)
        servicios_total = min(serv_tec * 20, 40) + min(informes * 10, 20)
        servicios_total = min(servicios_total, 40)
        st.caption(f"Subtotal Servicios: {servicios_total}")

    prod_total = min(publicaciones_total + desarrollos_total + servicios_total, SECCIONES_MAX["Producciones y servicios"])
    st.info(f"Puntaje Producciones/Servicios: **{prod_total}** / 350")

    # ====== 5) Otros antecedentes (máx. 200) ======
    st.subheader("5) Otros antecedentes (máx. 200)")
    with st.expander("Redes, gestión editorial y eventos (hasta 150)", True):
        oa1, oa2, oa3 = st.columns(3)
        with oa1: redes2 = st.number_input("Participación en redes (10 c/u, máx. 30)", 0, value=0)
        with oa2: org_ev = st.number_input("Organización de eventos (20 c/u, máx. 60)", 0, value=0)
        with oa3: gest_ed = st.number_input("Gestión editorial (hasta 60)", 0, value=0)
        redes_gestion_total = min(redes2 * 10, 30) + min(org_ev * 20, 60) + min(gest_ed, 60)
        redes_gestion_total = min(redes_gestion_total, 150)
        st.caption(f"Subtotal Redes/Gestión/Eventos: {redes_gestion_total}")

    with st.expander("Premios y distinciones (hasta 100)", True):
        ob1, ob2, ob3 = st.columns(3)
        with ob1: prem_int = st.number_input("Premios internacionales (50 c/u, máx. 100)", 0, value=0)
        with ob2: prem_nac = st.number_input("Premios nacionales (20 c/u, máx. 100)", 0, value=0)
        with ob3: menc = st.number_input("Menciones/distinciones (20 c/u, máx. 100)", 0, value=0)
        premios_total = min(prem_int * 50, 100) + min(prem_nac * 20, 100) + min(menc * 20, 100)
        premios_total = min(premios_total, 100)
        st.caption(f"Subtotal Premios/Distinciones: {premios_total}")

    otros_total = min(redes_gestion_total + premios_total, SECCIONES_MAX["Otros antecedentes"])
    st.info(f"Puntaje Otros antecedentes: **{otros_total}** / 200")

    # ====== Totales y categoría ======
    total = formacion_total + cargos_total + cyt_total + prod_total + otros_total
    categoria = determinar_categoria(total)

    st.markdown("---")
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1: st.metric("Formación", formacion_total)
    with c2: st.metric("Cargos", cargos_total)
    with c3: st.metric("CyT", cyt_total)
    with c4: st.metric("Producciones/Servicios", prod_total)
    c5, c6 = st.columns([1, 3])
    with c5: st.metric("Otros antecedentes", otros_total)
    with c6: st.metric("TOTAL", total)

    st.markdown(f"**Categoría alcanzada:** {tag_categoria(categoria)}", unsafe_allow_html=True)

    # ====== Resumen + exportadores ======
    resumen = pd.DataFrame([
        {"Sección": "Formación académica y complementaria", "Puntaje": formacion_total, "Tope": SECCIONES_MAX["Formación académica y complementaria"]},
        {"Sección": "Cargos (docencia, gestión y otros)", "Puntaje": cargos_total, "Tope": SECCIONES_MAX["Cargos (docencia, gestión y otros)"]},
        {"Sección": "Ciencia y Tecnología", "Puntaje": cyt_total, "Tope": SECCIONES_MAX["Ciencia y Tecnología"]},
        {"Sección": "Producciones y servicios", "Puntaje": prod_total, "Tope": SECCIONES_MAX["Producciones y servicios"]},
        {"Sección": "Otros antecedentes", "Puntaje": otros_total, "Tope": SECCIONES_MAX["Otros antecedentes"]},
        {"Sección": "TOTAL", "Puntaje": total, "Tope": 2000},
    ])
    st.dataframe(resumen, use_container_width=True)

    colx1, colx2 = st.columns(2)
    with colx1:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            resumen.to_excel(writer, index=False, sheet_name="Resultados")
        st.download_button(
            "⬇️ Descargar Excel de resultados",
            out.getvalue(),
            file_name=f"Valorador_Investigadores_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with colx2:
        # Importación perezosa aquí
        Document, Pt, WD_ALIGN_PARAGRAPH = _ensure_docx()
        doc = Document()
        style = doc.styles["Normal"]
        style.font.name = "Times New Roman"
        style.font.size = Pt(11)
        h = doc.add_paragraph("Universidad Católica de Cuyo – Secretaría de Investigación")
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        h.runs[0].bold = True
        doc.add_paragraph("Valorador de Currículum Docente – Categorización de Investigadores")
        info = doc.add_paragraph(f"Postulante: {nombre if nombre else '-'} | Unidad: {unidad if unidad else '-'} | Fecha: {fecha_eval.strftime('%Y-%m-%d')}")
        info.alignment = WD_ALIGN_PARAGRAPH.LEFT
        table = doc.add_table(rows=1, cols=3)
        hdr = table.rows[0].cells
        hdr[0].text = "Sección"; hdr[1].text = "Puntaje"; hdr[2].text = "Tope"
        for _, row in resumen.iterrows():
            r = table.add_row().cells
            r[0].text = str(row["Sección"]); r[1].text = str(int(row["Puntaje"])); r[2].text = str(int(row["Tope"]))
        wb = io.BytesIO(); doc.save(wb)
        st.download_button(
            "⬇️ Descargar Informe Word",
            wb.getvalue(),
            file_name=f"Valorador_Investigadores_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.markdown(
        """
        <div style="margin-top:14px; padding:10px 12px; border:1px dashed #cbd5e1; border-radius:12px; background:#f8fafc;">
        <b>Notas metodológicas</b><br>
        • Se aplican topes por ítem y por sección conforme a la propuesta institucional.<br>
        • El cálculo de categoría utiliza los rangos: 1500–2000 (I), 1000–1499 (II), 600–999 (III), 300–599 (IV), 1–299 (V), 0 (VI).<br>
        • Versión inicial: ingreso manual de cantidades. Futura mejora: lectura automática de CV (SIGEVA-CONICET) y asignación asistida.<br>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ──────────────────────────────────────────────────────────────────────────────
except Exception as e:
    st.error("Se produjo un error al iniciar la app. Revisa el detalle debajo.")
    st.exception(e)
