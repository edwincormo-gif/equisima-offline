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

st.set_page_config(layout="wide", page_title="EQUISIMA COMPLETA")
st.title("🟡 EQUISIMA - Completa")

if 'puntos' not in st.session_state: st.session_state.puntos=[]
if 'fotos' not in st.session_state: st.session_state.fotos=[]

tab1, tab2, tab3 = st.tabs(["📋 DATOS CAMPO", "📸 FOTOS PDF", "🔊 ANALISIS RES 0627"])

# ---------- FUNCIONES RES 627 ----------
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
                hoja=n; break
        df = pd.read_excel(file, sheet_name=hoja)
        # busca columna numerica con dB
        col_db = None
        for col in df.columns:
            try:
                s = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(s)>10 and 25 < s.mean() < 110:
                    col_db = col; break
            except: pass
        if col_db is None:
            # intenta csv
            return None, None
        vals = pd.to_numeric(df[col_db], errors='coerce').dropna().tolist()
        return vals, col_db
    except:
        # intenta CSV
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
                    except: pass
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

# ---------- TAB 1 ----------
with tab1:
    c1,c2=st.columns(2)
    with c1:
        cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.", key="cli")
        depto=st.text_input("DEPARTAMENTO","TOLIMA", key="dep")
        muni=st.text_input("MUNICIPIO","ATACO", key="mun")
        responsable=st.text_input("Responsable","edwin cortes", key="resp")
    with c2:
        proyecto=st.text_input("PROYECTO","", key="proy")
        marca_sono=st.text_input("Marca","HD2010UC/A", key="sono")

    punto=st.text_input("PUNTO","Punto 1 nocturno")
    coord=st.text_input("COORD","3°35'11.30\"N 75°23'27.72\"W")
    desc=st.text_area("DESCRIPCION","En la entrada...")
    barrido=st.number_input("BARRIDO dB",0,140,65)
    fecha=st.date_input("Fecha", key="fec"); hora=st.text_input("Hora","21:01")
    cal_i=st.text_input("Cal Ini","114"); cal_f=st.text_input("Cal Fin","114")
    mem=st.number_input("Memoria",1,99,1); laeq=st.number_input("LAeq",0.0,140.0,65.0)
    foto=st.file_uploader("FOTO ESQUEMA",type=['jpg','jpeg','png'], key="foto1")
    vmax=st.text_input("Vmax","0.3"); vdir=st.text_input("Dir","SW"); temp=st.text_input("Temp","29")
    hum=st.text_input("Hum","45"); precip=st.selectbox("Precip",["NO","SI"]); fuente=st.text_input("Fuente","mineria")
    tipo=st.text_input("Tipo","EMISION"); top=st.text_input("Tiempo","1"); obs=st.text_area("Obs","camino destapado...")

    if st.button("AGREGAR PUNTO",type="primary",use_container_width=True):
        fp=None
        if foto:
            f=tempfile.NamedTemporaryFile(delete=False,suffix=".jpg"); f.write(foto.getbuffer()); f.close(); fp=f.name
        st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs,fp])

    def build_excel():
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="Datos de Campo Emision"
        thin=Side(style='thin'); border=Border(left=thin,right=thin,top=thin,bottom=thin); no_border=Border(); bold=Font(bold=True,size=10)
        for k,w in {'A':22,'B':14,'C':16,'D':10,'E':10,'F':10,'G':8,'H':8,'I':8,'J':10,'K':14,'L':12,'M':12,'N':22}.items(): ws.column_dimensions[k].width=w
        ws['A1']="DATOS DE CAMPO EMISION DE RUIDO"; ws['A1'].font=Font(bold=True,size=12)
        r=10
        for p in st.session_state.puntos:
            ws.cell(row=r, column=1, value="PUNTO:").font=bold; ws.cell(row=r, column=2, value=p[0]); r+=1
            ws.cell(row=r, column=1, value="COORD:").font=bold; ws.cell(row=r, column=2, value=p[1]); r+=1
            ws.cell(row=r, column=1, value="LAeq:").font=bold; ws.cell(row=r, column=2, value=p[9]); ws.cell(row=r, column=3, value="Vmax:"); ws.cell(row=r, column=4, value=p[10]); r+=2
            foto_r=r
            if p[19] and os.path.exists(p[19]):
                try:
                    ws.merge_cells(start_row=foto_r, start_column=1, end_row=foto_r+5, end_column=14)
                    for rr in range(foto_r, foto_r+6): ws.row_dimensions[rr].height=70
                    img=XLImage(p[19]); img.width=850; img.height=380; ws.add_image(img, f'A{foto_r}')
                except: pass
            r=foto_r+7
        r+=2; ws.cell(row=r, column=1, value="Responsable:").font=Font(bold=True,size=11); ws.cell(row=r, column=3, value=responsable)
        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio
    if st.session_state.puntos:
        st.download_button("DESCARGAR EXCEL CAMPO", build_excel(), f"Campo_{muni}.xlsx", type="primary", use_container_width=True)

# ---------- TAB 2 ----------
with tab2:
    punto_foto = st.text_input("Punto foto", "Punto 1", key="pf_p")
    desc_foto = st.text_area("Desc foto", "Vista general...", key="pf_d")
    fotos_up = st.file_uploader("Sube fotos", type=['jpg','jpeg','png'], accept_multiple_files=True, key="pf_up")
    if st.button("AGREGAR FOTOS"):
        for f in fotos_up:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg"); tmp.write(f.getbuffer()); tmp.close()
            st.session_state.fotos.append({"punto": punto_foto, "desc": desc_foto, "path": tmp.name, "nombre": f.name})
    if st.session_state.fotos:
        def build_pdf():
            pdf = FPDF(); pdf.set_auto_page_break(auto=True, margin=15); pdf.add_page()
            pdf.set_font("Arial", 'B', 14); pdf.cell(0, 10, "REGISTRO FOTOGRAFICO", ln=True, align='C'); pdf.ln(5)
            for ft in st.session_state.fotos:
                y = pdf.get_y()
                if y > 180: pdf.add_page()
                try:
                    img = Image.open(ft['path']); img.thumbnail((800, 800))
                    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg"); img.save(tmp_pdf.name, "JPEG")
                    pdf.image(tmp_pdf.name, x=15, y=pdf.get_y(), w=180); pdf.set_y(pdf.get_y()+95)
                    pdf.set_font("Arial", 'B', 10); pdf.cell(0, 6, ft['punto'], ln=True)
                    pdf.set_font("Arial", '', 9); pdf.multi_cell(0, 5, ft['desc']); pdf.ln(5)
                except: pass
            return BytesIO(pdf.output())
        def build_zip():
            bio = BytesIO()
            with zipfile.ZipFile(bio, 'w') as z:
                for ft in st.session_state.fotos:
                    z.write(ft['path'], arcname=f"Fotos_{muni}/{ft['punto']}/{ft['nombre']}")
            bio.seek(0); return bio
        c1,c2=st.columns(2)
        c1.download_button("PDF FOTOS", build_pdf(), f"Fotos_{muni}.pdf", mime="application/pdf", type="primary", use_container_width=True)
        c2.download_button("ZIP CARPETA", build_zip(), f"Fotos_{muni}.zip", mime="application/zip", use_container_width=True)

# ---------- TAB 3 NUEVA - CON PDF CON GRAFICA ----------
with tab3:
    st.subheader("🔊 Análisis Sonómetro - Res 0627 de 2006")

    sector = st.selectbox("Sector Res 0627", ["A - Tranquilidad", "B - Residencial", "C - Industrial", "D - Centro"])
    periodo = st.selectbox("Periodo", ["diurno", "nocturno"])

    col1, col2 = st.columns(2)
    with col1:
        f_total = st.file_uploader("TOTAL (fuente ON) - xlsx o csv", type=["xlsx","xls","csv"], key="tot")
    with col2:
        f_res = st.file_uploader("RESIDUAL (fuente OFF) - xlsx o csv", type=["xlsx","xls","csv"], key="res")

    if f_total:
        vals_total, col_name = leer_hd2010(f_total)
        if vals_total:
            leq_total = calcular_leq(vals_total)
            st.metric(f"LAeq Total ({col_name})", f"{leq_total:.1f} dB(A)")
            st.line_chart(vals_total[:600])
        else:
            st.error("No pude leer dB. Mándame captura Time History")

    if f_total and f_res:
        vals_total, _ = leer_hd2010(f_total)
        vals_res, _ = leer_hd2010(f_res)
        leq_t = calcular_leq(vals_total)
        leq_r = calcular_leq(vals_res)
        try:
            emision = 10*np.log10(10**(leq_t/10) - 10**(leq_r/10))
        except:
            emision = leq_t

        limite = norma_627(sector, periodo)
        cumple = "CUMPLE" if emision <= limite else "NO CUMPLE - Excede norma"
        cumple_icon = "✅ CUMPLE" if emision <= limite else "❌ NO CUMPLE"

        st.divider()
        c1,c2,c3,c4=st.columns(4)
        c1.metric("LAeq Total", f"{leq_t:.1f} dB")
        c2.metric("LAeq Residual", f"{leq_r:.1f} dB")
        c3.metric("LAeq Emisión", f"{emision:.1f} dB")
        c4.metric(f"Límite {sector} {periodo}", f"{limite} dB")
        st.markdown(f"### {cumple_icon}")

        df_res = pd.DataFrame({
            "Concepto": ["LAeq Total", "LAeq Residual", "LAeq Emisión (Fuente)", "Límite Norma", "Cumplimiento"],
            "Valor": [f"{leq_t:.1f} dB(A)", f"{leq_r:.1f} dB(A)", f"{emision:.1f} dB(A)", f"{limite} dB(A)", cumple]
        })
        st.table(df_res)

        # ---- GENERAR GRAFICA PARA PDF ----
        import matplotlib.pyplot as plt
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
            # Encabezado
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 10, "ANALISIS DE RUIDO - RES 0627 DE 2006", ln=True, align='C')
            pdf.set_font("Arial", '', 10)
            pdf.cell(0, 6, f"Cliente: {cliente} | Proyecto: {proyecto} | Municipio: {muni} - {depto}", ln=True, align='C')
            pdf.cell(0, 6, f"Fecha analisis: {datetime.date.today()} | Sector: {sector} | Periodo: {periodo} | Responsable: {responsable}", ln=True, align='C')
            pdf.ln(5)
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(0, 8, "1. Resultados de Medicion", ln=True)
            pdf.set_font("Arial", '', 10)
            # Tabla
            pdf.set_fill_color(200,200,200)
            pdf.cell(90, 8, "Concepto", border=1, fill=True, align='C')
            pdf.cell(90, 8, "Valor dB(A)", border=1, fill=True, align='C', ln=True)
            for _, row in df_res.iterrows():
                pdf.cell(90, 7, row["Concepto"], border=1)
                pdf.cell(90, 7, row["Valor"], border=1, ln=True)
            pdf.ln(5)
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(0, 8, "2. Grafica Time History (primeros 10 minutos)", ln=True)
            pdf.image(graf_path, x=10, y=pdf.get_y(), w=190)
            pdf.set_y(pdf.get_y()+65)
            pdf.ln(5)
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(0, 8, "3. Calculo de Emision segun Res 0627", ln=True)
            pdf.set_font("Arial", '', 9)
            pdf.multi_cell(0, 5, f"Formula: L_emision = 10*log10(10^(Ltotal/10) - 10^(Lresidual/10))\nLtotal = {leq_t:.1f} dB, Lresidual = {leq_r:.1f} dB => L_emision = {emision:.1f} dB(A)\n\nLimite permisible para {sector} en periodo {periodo}: {limite} dB(A)\nResultado: {cumple_icon}\n\nObservacion: Si Ltotal - Lresidual < 3 dB, el aporte es despreciable segun norma. Diferencia medida: {leq_t-leq_r:.1f} dB.")
            pdf.ln(5)
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 6, f"Responsable de la Medicion: {responsable} ______________________", ln=True)
            return BytesIO(pdf.output())

        def build_analisis_excel():
            wb=openpyxl.Workbook(); ws=wb.active; ws.title="Analisis 627"
            ws['A1']="ANALISIS RES 0627"; ws['A1'].font=Font(bold=True,size=12)
            ws['A3']=f"Cliente: {cliente}"; ws['A4']=f"Municipio: {muni}"
            ws['A6']="Concepto"; ws['B6']="Valor dB(A)"
            for i,row in df_res.iterrows():
                ws.cell(row=7+i, column=1, value=row["Concepto"])
                ws.cell(row=7+i, column=2, value=row["Valor"])
            bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

        c1,c2 = st.columns(2)
        c1.download_button("📄 DESCARGAR PDF ANALISIS CON GRAFICA", build_analisis_pdf(), f"Analisis_627_{muni}_{sector}.pdf", mime="application/pdf", type="primary", use_container_width=True)
        c2.download_button("📊 DESCARGAR EXCEL ANALISIS", build_analisis_excel(), f"Analisis_627_{muni}.xlsx", use_container_width=True)
