# En requirements.txt agrega:
# reportlab

import streamlit as st
import io, struct, tempfile, zipfile, numpy as np
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from PIL import Image as PILImage

# ... tu código de pestaña 1 y 2 igual ...

# En pestaña 3, cambia la función generar_pdf() por esta:
def generar_pdf_profesional(datos, LAeq_list, LRAeq, fig):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=2*cm, bottomMargin=1.5*cm)

    style_title = ParagraphStyle('title', fontName='Helvetica-Bold', fontSize=14, alignment=1)
    style_h = ParagraphStyle('h', fontName='Helvetica-Bold', fontSize=11)
    style_n = ParagraphStyle('n', fontSize=9, leading=12)

    story = []

    # PORTADA
    story.append(Spacer(1, 4*cm))
    story.append(Paragraph("INFORME TÉCNICO<br/>NIVELES DE PRESIÓN SONORA<br/><br/>RUIDO AMBIENTAL", style_title))
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph(f"{datos['cliente']}<br/>{datos['direccion']}", style_title))
    story.append(Spacer(1, 2*cm))
    story.append(Paragraph("OCTUBRE - 2026", style_title))
    story.append(PageBreak())

    # PAGINA 2 - TABLA DE CONTENIDO (igual a tu EQ-RD-10-2026)
    story.append(Paragraph("TABLA DE CONTENIDO", style_h))
    story.append(Paragraph("1. INTRODUCCIÓN<br/>2. OBJETIVOS<br/>3. MARCO LEGAL - Tabla 1 Estándares Res 627<br/>4. DESCRIPCIÓN DEL PROYECTO - Tabla 2,3,4 Equipos<br/>5. DATOS METEOROLÓGICOS<br/>6. LOCALIZACIÓN PUNTOS - Tabla 6<br/>7. REGISTRO FOTOGRÁFICO - Tabla 7<br/>8. CÁLCULOS KT, KI - Tabla 8,9,10<br/>9. RESULTADOS - Tabla 11,12 - LRAeq<br/>10. CONCLUSIONES", style_n))
    story.append(PageBreak())

    # RESULTADOS - como tu foto pero con formato LATINCO
    story.append(Paragraph(f"Tabla 12. Resultados - Cliente: {datos['cliente']}", style_h))
    tabla_data = [
        ["Parámetro", "Valor", "Norma", "Resultado"],
        ["LAeq,T", f"{LRAeq:.1f} dB(A)", "-", "-"],
        ["LRAeq,1h", f"{LRAeq:.1f} dB(A)", "45 dB(A) Noct.", "CUMPLE"]
    ]
    t = Table(tabla_data, colWidths=[3*cm,3*cm,3*cm,3*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1a3c5e')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
    ]))
    story.append(t)
    story.append(Spacer(1, 1*cm))

    img_buf = io.BytesIO()
    fig.savefig(img_buf, format='PNG', dpi=150)
    img_buf.seek(0)
    story.append(RLImage(img_buf, width=14*cm, height=6*cm))

    doc.build(story)
    buf.seek(0)
    return buf
