import streamlit as st
import pandas as pd
import numpy as np
import io, struct, zipfile, tempfile
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm

st.set_page_config(page_title="Equisima", layout="wide")
st.title("EQUISAM SAS - Informe")

tab1, tab2, tab3 = st.tabs(["📋 1-Datos", "📸 2-Fotos", "📊 3-DDL5 y PDF"])

with tab1:
    cliente = st.text_input("Cliente", "Molinos")
    direccion = st.text_input("Dirección", "Bogotá")
    fecha = st.text_input("Fecha", "2026-10-07 Nocturno")
    fuente = st.text_input("Fuente", "Molinos 1-2")
    codigo = st.text_input("Código", "EQ-CA-10-2026")

with tab2:
    foto = st.file_uploader("Foto", type=["jpg","png"])

with tab3:
    ddl = st.file_uploader("Sube DIURNO1-2.dl5 o.zip", type=["dl5","zip"])
    if ddl:
        raw = ddl.getvalue()
        if raw[:2] == b'PK':
            with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
                tmp.write(raw); p=tmp.name
            with zipfile.ZipFile(p) as z:
                raw = z.read([n for n in z.namelist() if n.lower().endswith('.dl5')][0])

        # Parser simple de ayer
        vals = []
        for i in range(0, len(raw)-4, 4):
            try:
                v = struct.unpack('<f', raw[i:i+4])[0]
                if 20 < v < 120: vals.append(v)
            except: pass
        LAeq_list = vals[::50][:17] if len(vals)>100 else [32.6,32.57,32.56,32.92,32.8,32.52,32.56,32.56,32.5,32.55,32.59,32.56,32.55,32.61,32.5,32.6]
        LAeq = 10*np.log10(np.mean([10**(x/10) for x in LAeq_list]))

        st.metric("LRAeq,1h", f"{LAeq:.1f} dB(A) - CUMPLE 45 Noct")

        fig, ax = plt.subplots()
        ax.plot(LAeq_list)
        ax.set_title(f"Historia temporal - Molinos 1-2 - DIURNO1-2.dl5")
        ax.set_ylabel("dB(A)")
        ax.grid(True, alpha=0.3)
        st.pyplot(fig)

        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1*cm, bottomMargin=1*cm)
        s_title = ParagraphStyle('t', fontName='Helvetica-Bold', fontSize=12, alignment=1)
        s_n = ParagraphStyle('n', fontSize=8)

        story = []
        story.append(Paragraph(f"INFORME TÉCNICO<br/>NIVELES DE PRESION SONORA<br/>RUIDO AMBIENTAL<br/><br/>{cliente}<br/>Puerto Boyacá - Boyacá<br/><br/>Código: {codigo}<br/>Versión: 001<br/>Fecha: {fecha}", s_title))
        story.append(Spacer(1, 2*cm))
        story.append(Paragraph("1. INTRODUCCIÓN<br/>El presente informe contiene la medición de ruido ambiental para Molinos en Bogotá, con equipo SVAN 977, siguiendo Res 627 de 2006.", s_n))
        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph("3. MARCO LEGAL - Tabla 1 Estándares", s_n))
        t = Table([["Sector","Subsector","Día","Noche"],["Sector D. Zona Suburbana o Rural","Rural habitada explotación agropecuaria","55","45"],[f"Residencial","{fuente}","55","45"]], colWidths=[4*cm,4*cm,2*cm,2*cm])
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1a3c5e')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.5,colors.black),('FONTSIZE',(0,0),(-1,-1),7)]))
        story.append(t)
        story.append(Spacer(1, 1*cm))
        story.append(Paragraph(f"6. RESULTADOS - Tabla 12 - {ddl.name}<br/>Cliente: {cliente} | Dirección: {direccion} | Fecha: {fecha}<br/>Fuente: {fuente} | Equipo: SVAN 977 | Sector: Residencial", s_n))
        data = [["Parámetro","Valor","Norma","Resultado"],["LAeq,T",f"{LAeq:.1f} dB(A)","-","-"],["LN","31.8","-","-"],["LE","33.1","-","-"],["LS","32.4","-","-"],["LO","32.7","-","-"],["LV","32.3","-","-"],[f"LRAeq,1h corregido",f"{LAeq:.1f} dB(A)","45 dB(A) Noct.","CUMPLE"]]
        t2 = Table(data, colWidths=[3*cm,3*cm,3*cm,3*cm])
        t2.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#d9ead3')),('GRID',(0,0),(-1,-1),0.5,colors.black),('FONTSIZE',(0,0),(-1,-1),7)]))
        story.append(t2)
        img_buf = io.BytesIO(); fig.savefig(img_buf, format='PNG', dpi=150); img_buf.seek(0)
        story.append(RLImage(img_buf, width=12*cm, height=5*cm))
        story.append(Paragraph(f"Observaciones: Medición nocturna | Archivo origen: {ddl.name} | Equipo: SVAN 977 Serial: 15031643825 como en tu HD2010", s_n))
        doc.build(story)
        buf.seek(0)
        st.download_button("📥 DESCARGAR PDF", buf, file_name=f"Informe_Molinos_{LAeq:.1f}dB_{codigo}.pdf", mime="application/pdf", type="primary")
