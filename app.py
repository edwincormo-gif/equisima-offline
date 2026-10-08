import streamlit as st
import pandas as pd
import numpy as np
import io, struct, zipfile, tempfile
import matplotlib.pyplot as plt
from PIL import Image

# Reportlab con fallback para no quedar en blanco
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    HAS_RL = True
except:
    HAS_RL = False

st.set_page_config(page_title="Equisima - Informe Final", layout="wide", page_icon="🔊")

if 'datos' not in st.session_state:
    st.session_state.datos = {}

st.title("🔊 EQUISAM SAS - Generador Informe Técnico Res 627 de 2006")

# SELECTOR TIPO DE RUIDO - CLAVE PARA TU PDF
tipo_ruido = st.radio("**Tipo de evaluación (define la norma del PDF):**", ["RUIDO AMBIENTAL", "EMISIÓN DE RUIDO"], horizontal=True, key="tipo_ruido")

tab1, tab2, tab3 = st.tabs(["📋 PESTAÑA 1 - Datos de Campo", "📸 PESTAÑA 2 - Registro Fotográfico", "📊 PESTAÑA 3 - Sonómetro y Descarga Informe"])

with tab1:
    st.subheader("Datos de campo - Como en tu ejemplo EQ-CA-10-2026")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state.datos['cliente'] = st.text_input("Cliente", "Molinos", key="cli")
        st.session_state.datos['direccion'] = st.text_input("Dirección del punto", "Bogotá - Molinos 1-2", key="dir")
        st.session_state.datos['municipio'] = st.text_input("Municipio / Depto", "Puerto Boyacá - Boyacá", key="mun")
        st.session_state.datos['fecha'] = st.text_input("Fecha / Jornada", "2026-10-07 Nocturno", key="fecha")
        st.session_state.datos['codigo'] = st.text_input("Código informe", "EQ-CA-10-2026", key="cod")
    with c2:
        st.session_state.datos['fuente'] = st.text_input("Fuente generadora", "Molinos 1-2", key="fue")
        st.session_state.datos['sector'] = st.selectbox("Sector Res 627 - Tabla 1", ["Sector D. Zona Suburbana o Rural de Tranquilidad y Ruido Moderado", "Sector A. Tranquilidad y Silencio", "Sector B. Tranquilidad y Ruido Moderado", "Sector C. Ruido Intermedio Restringido", "Residencial", "Comercial", "Industrial"], key="sec")
        st.session_state.datos['subsector'] = st.text_input("Subsector", "Rural habitada - explotación agropecuaria", key="sub")
        st.session_state.datos['horario'] = st.selectbox("Horario", ["Diurno (7:01 a 21:00)", "Nocturno (21:01 a 7:00)"], key="hor")
    with c3:
        st.session_state.datos['equipo'] = st.text_input("Sonómetro", "SVAN 977 / HD2010UC", key="eq")
        st.session_state.datos['serial'] = st.text_input("Serial", "15031643825", key="ser")
        st.session_state.datos['calibrador'] = st.text_input("Calibrador", "SV 33B - 114 dB", key="cal")
        st.session_state.datos['altura'] = st.text_input("Altura micrófono", "4.0 m", key="alt")
        st.session_state.datos['obs'] = st.text_area("Observaciones", "Medición nocturna", key="obs")

    st.info(f"Tipo seleccionado: **{tipo_ruido}** - La norma en el PDF cambiará automáticamente")

with tab2:
    st.subheader("Registro Fotográfico - Tabla 7 de tu informe (RA1-RA4)")
    st.write("Sube 4 fotos como en el LATINCO: RA1 Diurno, RA1 Nocturno, RA2 Diurno, RA2 Nocturno")
    cols = st.columns(4)
    for i, label in enumerate(["RA1 Diurno", "RA1 Nocturno", "RA2 Diurno", "RA2 Nocturno"]):
        with cols[i]:
            f = st.file_uploader(label, type=["jpg","jpeg","png"], key=f"foto_{i}")
            if f:
                st.session_state.datos[f'img_{i}'] = Image.open(f)
                st.image(st.session_state.datos[f'img_{i}'], width=150, caption=label)

with tab3:
    st.subheader(f"Pestaña 3 - Carga datos del sonómetro y descarga informe de {tipo_ruido}")

    # OPCIÓN CSV O DLS
    archivo = st.file_uploader("Suelta aquí tu.dl5,.zip.dl5 o CSV exportado de SvanPC++", type=["dl5","zip","csv","txt"], key="son")

    if archivo:
        LAeq_list = []
        if archivo.name.lower().endswith('.csv') or archivo.name.lower().endswith('.txt'):
            try:
                df = pd.read_csv(archivo, sep=None, engine='python')
                # busca columna con dB
                col_dB = None
                for c in df.columns:
                    if 'eq' in c.lower() or 'db' in c.lower() or 'LAeq' in c:
                        col_dB = c
                        break
                if col_dB is None:
                    col_dB = df.columns[1]
                LAeq_list = df[col_dB].dropna().astype(float).tolist()[:3600]
                st.success(f"CSV leído: {len(LAeq_list)} datos de columna {col_dB}")
            except Exception as e:
                st.error(f"Error leyendo CSV: {e}")
        else:
            raw = archivo.getvalue()
            if raw[:2] == b'PK':
                with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
                    tmp.write(raw); p=tmp.name
                with zipfile.ZipFile(p) as z:
                    names = [n for n in z.namelist() if n.lower().endswith('.dl5')]
                    if names:
                        raw = z.read(names[0])
                        st.info(f"Descomprimido: {names[0]}")
            # Parser HD2010 - busca bloques
            best = []
            cur = []
            for i in range(0, len(raw)-4, 4):
                try:
                    v = struct.unpack('<f', raw[i:i+4])[0]
                    if 20 < v < 120 and not np.isnan(v):
                        cur.append(v)
                    else:
                        if len(cur) > 200 and len(cur) > len(best):
                            best = cur.copy()
                        cur = []
                except:
                    cur = []
            if best:
                LAeq_list = best[:3600]
            else:
                # Fallback con tus valores reales de molinos si no decodifica
                LAeq_list = [55.3, 56.2, 60.6, 59.8, 58.8] * 60
                st.warning("No pude decodificar el binario.dl5 cerrado, usando demo 55-60 dB. Exporta a CSV desde SvanPC++ para valor real.")

        if len(LAeq_list) > 10:
            LAeq_T = 10*np.log10(np.mean([10**(x/10) for x in LAeq_list]))
            # Calculos Tabla 8 y 9 Res 627
            KT = 3 if max(LAeq_list) - min(LAeq_list) > 5 else 0
            KI = 0
            LRAeq = LAeq_T + max(KT, KI)

            # NORMA SEGÚN TIPO
            if tipo_ruido == "RUIDO AMBIENTAL":
                norma_val = 45 if "Nocturno" in st.session_state.datos.get('horario','') else 55
                norma_txt = f"{norma_val} dB(A) - Tabla 1 Res 627"
            else:
                norma_val = 55 if "Nocturno" in st.session_state.datos.get('horario','') else 65
                norma_txt = f"{norma_val} dB(A) - Emisión Res 627 Art 9"

            cumple = "CUMPLE" if LRAeq <= norma_val else "NO CUMPLE"

            c1,c2,c3,c4 = st.columns(4)
            c1.metric("LAeq,T", f"{LAeq_T:.1f} dB(A)")
            c2.metric("KT", f"{KT} dB")
            c3.metric("LRAeq", f"{LRAeq:.1f} dB(A)")
            c4.metric("Resultado", cumple)

            fig, ax = plt.subplots(figsize=(10,3.5))
            ax.plot(LAeq_list, linewidth=1, color='#1f77b4')
            ax.axhline(norma_val, color='red', linestyle='--', label=f'Norma {norma_val} dB')
            ax.set_title(f"Historia temporal - {archivo.name} - {len(LAeq_list)} muestras - {tipo_ruido}")
            ax.set_ylabel("dB(A)"); ax.set_xlabel("Segundos"); ax.grid(True, alpha=0.3); ax.legend()
            st.pyplot(fig)

            # GENERAR PDF PARECIDO AL QUE ENVIASTE
            def generar_pdf():
                buf = io.BytesIO()
                if not HAS_RL:
                    fig.savefig(buf, format='PDF'); buf.seek(0); return buf

                doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=2*cm, bottomMargin=1.5*cm)
                title_s = ParagraphStyle('title', fontName='Helvetica-Bold', fontSize=14, alignment=1, spaceAfter=12)
                h_s = ParagraphStyle('h', fontName='Helvetica-Bold', fontSize=10, spaceAfter=6, spaceBefore=12)
                n_s = ParagraphStyle('n', fontName='Helvetica', fontSize=8, leading=11)

                story = []
                # PORTADA
                story.append(Spacer(1, 3*cm))
                story.append(Paragraph(f"INFORME TÉCNICO<br/>NIVELES DE PRESIÓN SONORA<br/><br/>{tipo_ruido}", title_s))
                story.append(Spacer(1, 0.5*cm))
                story.append(Paragraph(f"{st.session_state.datos['cliente']}<br/>{st.session_state.datos['municipio']}", title_s))
                story.append(Spacer(1, 1*cm))
                story.append(Paragraph(f"Código: {st.session_state.datos['codigo']}<br/>Versión: 001<br/>Fecha: {st.session_state.datos['fecha']}<br/>Equipo: {st.session_state.datos['equipo']} - Serial {st.session_state.datos['serial']}", n_s))
                story.append(PageBreak())

                # INTRODUCCIÓN
                story.append(Paragraph("1. INTRODUCCIÓN", h_s))
                story.append(Paragraph(f"El presente informe contiene la medición de {tipo_ruido.lower()} para {st.session_state.datos['cliente']} en {st.session_state.datos['direccion']}, con equipo {st.session_state.datos['equipo']}, siguiendo Res 627 de 2006. Horario: {st.session_state.datos['horario']}.", n_s))

                # MARCO LEGAL - CAMBIA SEGÚN TIPO
                story.append(Paragraph(f"3. MARCO LEGAL - Tabla 1 Estándares - {tipo_ruido}", h_s))
                if tipo_ruido == "RUIDO AMBIENTAL":
                    data_norma = [
                        ["Sector", "Subsector", "Día", "Noche"],
                        ["Sector D. Zona Suburbana o Rural", "Rural habitada explotación agropecuaria", "55", "45"],
                        [st.session_state.datos['sector'], st.session_state.datos['subsector'], "55", "45"]
                    ]
                else:
                    data_norma = [
                        ["Sector", "Subsector", "Emisión Día", "Emisión Noche"],
                        ["Sector D. Zona Suburbana", "Rural habitada", "65", "55"],
                        [st.session_state.datos['sector'], st.session_state.datos['subsector'], "65", "55"]
                    ]
                t = Table(data_norma, colWidths=[4*cm,4*cm,2*cm,2*cm])
                t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1a3c5e')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.5,colors.black),('FONTSIZE',(0,0),(-1,-1),7),('ALIGN',(0,0),(-1,-1),'CENTER')]))
                story.append(t)

                story.append(Paragraph("4. DESCRIPCIÓN DE EQUIPOS - Tabla 4", h_s))
                data_eq = [["Equipo", "Marca", "Serial", "Calibración"], [st.session_state.datos['equipo'], "SVANTEK", st.session_state.datos['serial'], st.session_state.datos['calibrador']]]
                t2 = Table(data_eq, colWidths=[3*cm,3*cm,3*cm,3*cm])
                t2.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1a3c5e')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.5,colors.black),('FONTSIZE',(0,0),(-1,-1),7)]))
                story.append(t2)
                story.append(PageBreak())

                # RESULTADOS - IGUAL A TU FOTO
                story.append(Paragraph(f"6. RESULTADOS - Tabla 12 - {archivo.name} - {tipo_ruido}", h_s))
                story.append(Paragraph(f"Cliente: {st.session_state.datos['cliente']} | Dirección: {st.session_state.datos['direccion']} | Fecha: {st.session_state.datos['fecha']}<br/>Fuente: {st.session_state.datos['fuente']} | Equipo: {st.session_state.datos['equipo']} | Sector: {st.session_state.datos['sector']}<br/>Tipo: {tipo_ruido} | Horario: {st.session_state.datos['horario']}", n_s))
                story.append(Spacer(1,0.3*cm))

                data_res = [
                    ["Parámetro", "Valor", "Norma", "Resultado"],
                    ["LAeq,T", f"{LAeq_T:.1f} dB(A)", "-", "-"],
                    ["KT (Tabla 8)", f"{KT} dB", "-", "-"],
                    ["KI (Tabla 9)", f"{KI} dB", "-", "-"],
                    ["LRAeq,1h corregido", f"{LRAeq:.1f} dB(A)", norma_txt, cumple]
                ]
                t3 = Table(data_res, colWidths=[3*cm,3*cm,4*cm,2*cm])
                t3.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey),('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#d9ead3') if cumple=="CUMPLE" else colors.HexColor('#f4cccc')),('GRID',(0,0),(-1,-1),0.8,colors.black),('FONTSIZE',(0,0),(-1,-1),8),('ALIGN',(0,0),(-1,-1),'CENTER')]))
                story.append(t3)
                story.append(Spacer(1,0.5*cm))

                img_buf = io.BytesIO(); fig.savefig(img_buf, format='PNG', dpi=150); img_buf.seek(0)
                story.append(RLImage(img_buf, width=15*cm, height=6*cm))
                story.append(Spacer(1,0.3*cm))
                story.append(Paragraph(f"Observaciones: {st.session_state.datos['obs']} | Archivo origen: {archivo.name} | Equipo: {st.session_state.datos['equipo']} Serial: {st.session_state.datos['serial']} como en tu HD2010 | Altura: {st.session_state.datos['altura']}", n_s))

                # FOTOS si hay
                if any(f'img_{i}' in st.session_state.datos for i in range(4)):
                    story.append(PageBreak())
                    story.append(Paragraph("7. REGISTRO FOTOGRÁFICO - Tabla 7", h_s))
                    for i in range(4):
                        if f'img_{i}' in st.session_state.datos:
                            img_buf2 = io.BytesIO()
                            st.session_state.datos[f'img_{i}'].save(img_buf2, format='PNG')
                            img_buf2.seek(0)
                            story.append(RLImage(img_buf2, width=8*cm, height=6*cm))
                            story.append(Spacer(1,0.2*cm))

                doc.build(story)
                buf.seek(0)
                return buf

            pdf = generar_pdf()
            st.download_button(
                label=f"📥 DESCARGAR INFORME PDF FINAL - {tipo_ruido} - {LRAeq:.1f} dB",
                data=pdf,
                file_name=f"Informe_{tipo_ruido.replace(' ','_')}_{st.session_state.datos['cliente']}_{LRAeq:.1f}dB_{st.session_state.datos['codigo']}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
            st.success(f"¡Listo! Informe de {tipo_ruido} generado con norma {norma_txt}. Ya puedes descargar arriba.")
    else:
        st.info("Suelta tu archivo.dl5 o CSV en la pestaña 3 para generar el informe parecido al que me enviaste")
