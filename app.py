import streamlit as st
import pandas as pd
import io, math, os
import matplotlib.pyplot as plt
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY

st.set_page_config(page_title="EQUISAM V14 MAPA + FOTOS", layout="wide")
st.title("EQUISAM V14 - RES 627 con Mapa + Registro Fotográfico")

# === SIDEBAR ===
st.sidebar.header("Tipo Res 627")
tipo = st.sidebar.selectbox("Tipo", ["RUIDO AMBIENTAL - Cap II", "EMISIÓN DE RUIDO - Cap I"])
es_emision = "EMISIÓN" in tipo
sector_sel = st.sidebar.selectbox("Sector", ["D. Zona Suburbana o Rural (55/45) - LATINCO","A. Tranquilidad (55/45)","B. Tranquilidad Moderado (65/50)","C. Intermedio (75/70)","C. Industrial (75/75)"])
diurno, nocturno = (55,45) if "55" in sector_sel else (75,75)
cliente = st.sidebar.text_input("Cliente", "LATINCO S.A.")
codigo = st.sidebar.text_input("Código", "EQ-CA-10-2026")
municipio = st.sidebar.text_input("Municipio", "Puerto Boyacá")
vereda = st.sidebar.text_input("Vereda", "La Pizarra Balastrera")

tab1, tab2, tab3 = st.tabs(["📍 Puntos + Coordenadas", "📸 Fotos RA1-RA4", "📄 Generar PDF"])

if "fotos" not in st.session_state:
    st.session_state.fotos = {}

with tab1:
    st.subheader("Tabla 6 - Ubicación puntos + Tabla 5 Meteo - Para Ilustración 1 Mapa")
    st.info("Pon LAT y LON reales. La app te genera el mapa automáticamente para el PDF")

    df_init = pd.DataFrame([
        ["RA1", 6.4321, -74.4321, "14/07/2026 10:00", "14/07/2026 11:00", 0.5, "N", 32, 68, "No", 61.3, 60.8, 55.0],
        ["RA2", 6.4330, -74.4310, "14/07/2026 11:15", "14/07/2026 12:15", 0.3, "NE", 33, 70, "No", 63.6, 63.1, 57.0],
        ["RA3", 6.4340, -74.4300, "14/07/2026 12:50", "14/07/2026 13:50", 0.4, "E", 32, 69, "No", 60.8, 62.6, 56.2],
        ["RA4", 6.4350, -74.4290, "14/07/2026 14:02", "14/07/2026 15:02", 0.2, "S", 31, 70, "No", 67.1, 64.7, 58.1],
    ], columns=["Punto","LAT","LON","Inicio","Fin","Vviento","Dir","Temp","Hum","Precip","LN","LRAeq","Lres"])

    df_puntos = st.data_editor(df_init, num_rows="dynamic", use_container_width=True, key="puntos")

    # PREVIEW MAPA
    if not df_puntos.empty:
        fig, ax = plt.subplots(figsize=(5,4))
        ax.scatter(df_puntos["LON"], df_puntos["LAT"], c='red', s=100, edgecolors='black', zorder=5)
        for _, r in df_puntos.iterrows():
            ax.text(r["LON"], r["LAT"]+0.0001, r["Punto"], fontsize=9, fontweight='bold')
        ax.set_xlabel("Longitud")
        ax.set_ylabel("Latitud")
        ax.set_title(f"Ilustración 1 - Ubicación puntos {municipio}")
        ax.grid(True, alpha=0.3)
        st.pyplot(fig)
        st.session_state.mapa_fig = fig

with tab2:
    st.subheader("Tabla 7 - Registro Fotográfico - Se pegan automático en PDF")
    if df_puntos.empty:
        st.warning("Primero agrega puntos en Pestaña 1")
    else:
        sel_punto = st.selectbox("Selecciona Punto para subir fotos", df_puntos["Punto"].tolist())
        files = st.file_uploader(f"Fotos para {sel_punto} (puedes subir 2: diurna y nocturna)", type=["jpg","png","jpeg"], accept_multiple_files=True, key=f"foto_{sel_punto}")
        if files:
            st.session_state.fotos[sel_punto] = files
            cols = st.columns(len(files))
            for i, f in enumerate(files):
                cols[i].image(f, caption=f"{sel_punto} - {f.name}", width=200)

        if st.session_state.fotos:
            st.success(f"Fotos cargadas: {list(st.session_state.fotos.keys())}")

with tab3:
    def generar_pdf_v14():
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=2*cm, bottomMargin=1.5*cm)
        s_title = ParagraphStyle('title', fontSize=13, alignment=TA_CENTER, fontName='Helvetica-Bold')
        s_h1 = ParagraphStyle('h1', fontSize=11, fontName='Helvetica-Bold', textColor=colors.HexColor('#0a2a5e'), spaceBefore=12, spaceAfter=6)
        s_n = ParagraphStyle('n', fontSize=9, leading=12, alignment=TA_JUSTIFY, fontName='Helvetica')
        s_small = ParagraphStyle('small', fontSize=7, fontName='Helvetica')

        def header(canvas, docu):
            canvas.saveState()
            canvas.setFont('Helvetica-Bold', 7)
            canvas.rect(1.5*cm, 26*cm, 18*cm, 1.2*cm, stroke=1, fill=0)
            canvas.drawString(2*cm, 26.6*cm, f"INFORME {tipo} - {cliente}")
            canvas.drawString(2*cm, 26.2*cm, f"Código: {codigo} R6-POE1-I V04 Pág: {docu.page}")
            canvas.restoreState()

        story = []
        # PORTADA
        story.append(Spacer(1,3*cm))
        story.append(Paragraph(f"INFORME TÉCNICO<br/>NIVELES DE PRESIÓN SONORA<br/><br/>{tipo}<br/><br/>{cliente}<br/>MUNICIPIO {municipio}<br/>Código: {codigo}", s_title))
        story.append(PageBreak())

        # TABLA 6 UBICACIÓN
        story.append(Paragraph("Tabla 6. Ubicación de los puntos de monitoreo", s_h1))
        rows6 = [["Punto","LAT","LON","Dirección"]] + [[r["Punto"], str(r["LAT"]), str(r["LON"]), vereda] for _, r in df_puntos.iterrows()]
        t6 = Table(rows6, colWidths=[2*cm,3*cm,3*cm,8*cm])
        t6.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0a2a5e')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.5,colors.black),('FONTSIZE',(0,0),(-1,-1),8)]))
        story.append(t6)
        story.append(Spacer(1,0.5*cm))

        # ILUSTRACIÓN 1 MAPA
        story.append(Paragraph("Ilustración 1 Ubicación de los puntos de monitoreo de ruido ambiental", s_h1))
        if "mapa_fig" in st.session_state:
            img_buf = io.BytesIO()
            st.session_state.mapa_fig.savefig(img_buf, format='PNG', dpi=150, bbox_inches='tight')
            img_buf.seek(0)
            story.append(RLImage(img_buf, width=14*cm, height=9*cm))
            story.append(Paragraph(f"Fuente: Equisam SAS - Coordenadas {municipio} {vereda}", s_small))
        story.append(PageBreak())

        # TABLA 7 FOTOS
        story.append(Paragraph("Tabla 7. Registro Fotográfico de Ruido Ambiental", s_h1))
        if st.session_state.fotos:
            for punto, fotos_list in st.session_state.fotos.items():
                story.append(Paragraph(f"{punto} - {vereda}", s_h1))
                img_row = []
                for f in fotos_list[:2]: # 2 por fila
                    try:
                        pil = Image.open(f)
                        # Redimensionar para PDF
                        pil.thumbnail((800,600))
                        buf_img = io.BytesIO()
                        pil.save(buf_img, format='PNG')
                        buf_img.seek(0)
                        img_row.append(RLImage(buf_img, width=7*cm, height=5*cm))
                    except:
                        pass
                if img_row:
                    # Tabla de 2 fotos
                    if len(img_row) == 1:
                        story.append(img_row[0])
                    else:
                        t_foto = Table([[img_row[0], img_row[1]]], colWidths=[8*cm,8*cm])
                        story.append(t_foto)
                story.append(Spacer(1,0.3*cm))
        else:
            story.append(Paragraph("Sin fotos cargadas - Suba fotos en Pestaña 2", s_n))

        story.append(PageBreak())

        # RESULTADOS
        story.append(Paragraph(f"Tabla 11. Resultados {tipo} - Diurno y Nocturno", s_h1))
        header_r = ["Punto","Inicio","Fin","LRAeq","Norma","Cumple?"]
        if es_emision: header_r.append("L Emisión")
        rows = [header_r]
        for _, r in df_puntos.iterrows():
            lra = float(r["LRAeq"])
            es_noct = "00:" in str(r["Inicio"]) or "21:" in str(r["Inicio"]) or "22:" in str(r["Inicio"]) or "23:" in str(r["Inicio"])
            norma_comp = nocturno if es_noct else diurno
            cumple = "CUMPLE" if lra <= norma_comp else "NO CUMPLE"
            row = [r["Punto"], str(r["Inicio"]), str(r["Fin"]), f"{lra} dB", f"{norma_comp} dB", cumple]
            if es_emision:
                lres = float(r["Lres"])
                try:
                    le = 10*math.log10(10**(lra/10)-10**(lres/10)) if lra>lres+3 else lra
                except: le=lra
                row.append(f"{le:.1f} dB")
            rows.append(row)

        t_r = Table(rows, colWidths=[2*cm,3*cm,3*cm,2*cm,2*cm,2*cm,2*cm][:len(header_r)])
        t_r.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0a2a5e')),('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('GRID',(0,0),(-1,-1),0.5,colors.black),('FONTSIZE',(0,0),(-1,-1),7),
            ('BACKGROUND',(0,1),(-1,-1),colors.HexColor('#e7f3ff'))
        ]))
        story.append(t_r)

        # Gráfica comparación
        story.append(Spacer(1,0.5*cm))
        fig2, ax2 = plt.subplots(figsize=(6,3))
        vals = df_puntos["LRAeq"].astype(float).tolist()
        labs = df_puntos["Punto"].tolist()
        ax2.bar(labs, vals, color='#1f77b4')
        ax2.axhline(diurno, color='red', linestyle='--', label=f'Diurno {diurno} dB')
        ax2.axhline(nocturno, color='orange', linestyle='--', label=f'Nocturno {nocturno} dB')
        ax2.set_ylabel('dB(A)'); ax2.legend(); plt.xticks(rotation=15)
        buf2 = io.BytesIO(); fig2.savefig(buf2, format='PNG', dpi=150); buf2.seek(0)
        story.append(RLImage(buf2, width=14*cm, height=5*cm))

        doc.build(story, onFirstPage=header, onLaterPages=header)
        buf.seek(0)
        return buf

    if st.button("📥 GENERAR PDF V14 CON MAPA + FOTOS", type="primary", use_container_width=True):
        if df_puntos.empty:
            st.error("Agrega puntos primero")
        else:
            pdf = generar_pdf_v14()
            st.balloons()
            st.success("¡PDF V14 con Mapa Ilustración 1 + Tabla 7 Fotos generado!")
            st.download_button("📥 DESCARGAR INFORME V14 - MAPA + FOTOS + RES627", pdf, f"{codigo}_V14_MAPA_FOTOS.pdf", "application/pdf", type="primary", use_container_width=True)
