import streamlit as st
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
import tempfile, os, zipfile, datetime
from fpdf import FPDF
from PIL import Image
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(layout="wide", page_title="EQUISIMA COMPLETA")
st.title("EQUISIMA - Completa")

if 'puntos' not in st.session_state: st.session_state.puntos=[]
if 'fotos' not in st.session_state: st.session_state.fotos=[]

def calcular_leq(vals):
    vals = np.array([v for v in vals if 20 < v < 140])
    if len(vals)==0: return 0
    return 10*np.log10(np.mean(10**(vals/10)))

def leer_hd2010(file):
    try:
        xls = pd.ExcelFile(file)
        hoja = xls.sheet_names[0]
        for n in xls.sheet_names:
            if "history" in n.lower() or "time" in n.lower() or "logger" in n.lower():
                hoja=n
                break
        df = pd.read_excel(file, sheet_name=hoja)
        col_db = None
        for col in df.columns:
            try:
                s = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(s)>10 and 25 < s.mean() < 110:
                    col_db = col
                    break
            except:
                pass
        if col_db is None:
            return None, None
        vals = pd.to_numeric(df[col_db], errors='coerce').dropna().tolist()
        return vals, col_db
    except:
        try:
            file.seek(0)
            content = file.read().decode('latin1', errors='ignore')
            lines = content.split('\n')
            vals=[]
            for i in range(1,len(lines)):
                parts = lines[i].split(';')
                if len(parts)>=2:
                    try:
                        x = parts[1].replace(',','.'); v=float(x)
                        if 20 < v < 140: vals.append(v)
                    except:
                        pass
            return vals, "CSV"
        except:
            return None, None

def norma_627(sector, periodo):
    tabla = {
        "A - Tranquilidad": {"diurno": 55, "nocturno": 45},
        "B - Residencial": {"diurno": 65, "nocturno": 50},
        "C - Industrial": {"diurno": 75, "nocturno": 70},
        "D - Centro": {"diurno": 70, "nocturno": 60},
    }
    return tabla.get(sector, {"diurno":65,"nocturno":50})[periodo]

tab1, tab2, tab3 = st.tabs(["DATOS CAMPO", "FOTOS PDF", "ANALISIS RES 0627"])

with tab1:
    c1,c2=st.columns(2)
    with c1:
        cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.", key="cli")
        depto=st.text_input("DEPARTAMENTO","TOLIMA", key="dep")
        muni=st.text_input("MUNICIPIO","ATACO", key="mun")
        responsable=st.text_input("Responsable","edwin cortes", key="resp")
    with c2:
        proyecto=st.text_input("PROYECTO","", key="proy")
    punto=st.text_input("PUNTO","Punto 1 nocturno")
    coord=st.text_input("COORD","3 35'11.30 N 75 23'27.72 W")
    desc=st.text_area("DESCRIPCION","En la entrada...")
    laeq=st.number_input("LAeq",0.0,140.0,65.0)
    foto=st.file_uploader("FOTO ESQUEMA",type=['jpg','jpeg','png'], key="foto1")
    vmax=st.text_input("Vmax","0.3")
    if st.button("AGREGAR PUNTO",type="primary",use_container_width=True):
        fp=None
        if foto:
            f=tempfile.NamedTemporaryFile(delete=False,suffix=".jpg"); f.write(foto.getbuffer()); f.close(); fp=f.name
        st.session_state.puntos.append([punto,coord,desc,laeq,fp])

with tab2:
    punto_foto = st.text_input("Punto foto", "Punto 1", key="pf_p")
    desc_foto = st.text_area("Desc foto", "Vista general...", key="pf_d")
    fotos_up = st.file_uploader("Sube fotos", type=['jpg','jpeg','png'], accept_multiple_files=True, key="pf_up")
    if st.button("AGREGAR FOTOS"):
        for f in fotos_up:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg"); tmp.write(f.getbuffer()); tmp.close()
            st.session_state.fotos.append({"punto": punto_foto, "desc": desc_foto, "path": tmp.name, "nombre": f.name})

with tab3:
    st.subheader("Analisis Sonometro - Res 0627 de 2006")
    sector = st.selectbox("Sector Res 0627", ["A - Tranquilidad", "B - Residencial", "C - Industrial", "D - Centro"])
    periodo = st.selectbox("Periodo", ["diurno", "nocturno"])
    col1, col2 = st.columns(2)
    with col1:
        f_total = st.file_uploader("TOTAL (fuente ON)", type=["xlsx","xls","csv"], key="tot")
    with col2:
        f_res = st.file_uploader("RESIDUAL (fuente OFF)", type=["xlsx","xls","csv"], key="res")

    if f_total and f_res:
        vals_total, _ = leer_hd2010(f_total)
        vals_res, _ = leer_hd2010(f_res)
        if vals_total and vals_res:
            leq_t = calcular_leq(vals_total)
            leq_r = calcular_leq(vals_res)
            try:
                emision = 10*np.log10(10**(leq_t/10) - 10**(leq_r/10))
            except:
                emision = leq_t
            limite = norma_627(sector, periodo)
            cumple = "CUMPLE" if emision <= limite else "NO CUMPLE"
            cumple_icon = "CUMPLE" if emision <= limite else "NO CUMPLE - Excede norma"

            c1,c2,c3,c4=st.columns(4)
            c1.metric("LAeq Total", f"{leq_t:.1f} dB")
            c2.metric("LAeq Residual", f"{leq_r:.1f} dB")
            c3.metric("LAeq Emision", f"{emision:.1f} dB")
            c4.metric(f"Limite {sector} {periodo}", f"{limite} dB")
            st.markdown(f"### {cumple_icon}")

            df_res = pd.DataFrame({
                "Concepto": ["LAeq Total", "LAeq Residual", "LAeq Emision (Fuente)", "Limite Norma", "Cumplimiento"],
                "Valor": [f"{leq_t:.1f} dB(A)", f"{leq_r:.1f} dB(A)", f"{emision:.1f} dB(A)", f"{limite} dB(A)", cumple]
            })
            st.table(df_res)

            fig, ax = plt.subplots(figsize=(8,3))
            ax.plot(vals_total[:600], label=f"Total {leq_t:.1f} dB", linewidth=1)
            if len(vals_res) >= 600:
                ax.plot(vals_res[:600], label=f"Residual {leq_r:.1f} dB", linewidth=1, alpha=0.7)
            ax.set_xlabel("Tiempo (s)"); ax.set_ylabel("dB(A)")
            ax.set_title(f"Perfil acustico - {muni} - {sector}")
            ax.legend(); ax.grid(True, alpha=0.3)
            graf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
            plt.tight_layout(); plt.savefig(graf_path, dpi=150); plt.close()

            def build_analisis_pdf():
                pdf = FPDF(orientation='P', unit='mm', format='A4')
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                pdf.set_font("Arial", 'B', 16)
                pdf.cell(0, 10, "ANALISIS DE RUIDO - RES 0627 DE 2006", ln=True, align='C')
                pdf.set_font("Arial", '', 10)
                pdf.cell(0, 6, f"Cliente: {cliente} | Municipio: {muni} - {depto}", ln=True, align='C')
                pdf.cell(0, 6, f"Fecha: {datetime.date.today()} | Sector: {sector} | Periodo: {periodo} | Resp: {responsable}", ln=True, align='C')
                pdf.ln(5)
                pdf.set_font("Arial", 'B', 11)
                pdf.cell(0, 8, "1. Resultados de Medicion", ln=True)
                pdf.set_font("Arial", '', 10)
                pdf.set_fill_color(200,200,200)
                pdf.cell(90, 8, "Concepto", border=1, fill=True, align='C')
                pdf.cell(90, 8, "Valor dB(A)", border=1, fill=True, align='C', ln=True)
                for _, row in df_res.iterrows():
                    concepto = str(row["Concepto"]).encode('latin-1','ignore').decode('latin-1')
                    valor = str(row["Valor"]).encode('latin-1','ignore').decode('latin-1')
                    pdf.cell(90, 7, concepto, border=1)
                    pdf.cell(90, 7, valor, border=1, ln=True)
                pdf.ln(5)
                pdf.set_font("Arial", 'B', 11)
                pdf.cell(0, 8, "2. Grafica Time History", ln=True)
                pdf.image(graf_path, x=10, y=pdf.get_y(), w=190)
                pdf.set_y(pdf.get_y()+65)
                pdf.ln(5)
                pdf.set_font("Arial", 'B', 11)
                pdf.cell(0, 8, "3. Calculo de Emision", ln=True)
                pdf.set_font("Arial", '', 9)
                txt = f"Formula: Le = 10*log10(10^(Lt/10) - 10^(Lr/10))\nLt = {leq_t:.1f} dB, Lr = {leq_r:.1f} dB => Le = {emision:.1f} dB(A)\nLimite para {sector} en {periodo}: {limite} dB(A)\nResultado: {cumple}\nDiferencia Lt-Lr: {leq_t-leq_r:.1f} dB."
                pdf.multi_cell(0, 5, txt)
                pdf.ln(5)
                pdf.set_font("Arial", 'B', 10)
                pdf.cell(0, 6, f"Responsable: {responsable}", ln=True)
                return BytesIO(pdf.output())

            st.download_button("DESCARGAR PDF ANALISIS CON GRAFICA", build_analisis_pdf(), f"Analisis_627_{muni}_{sector}.pdf", mime="application/pdf", type="primary", use_container_width=True)
        else:
            st.error("No pude leer los archivos")
