import streamlit as st
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
import tempfile, os, zipfile
from fpdf import FPDF
from PIL import Image
import datetime

st.set_page_config(layout="wide", page_title="EQUISIMA COMPLETA")
st.title("🟡 EQUISIMA - Campo + Fotos Informe")

if 'puntos' not in st.session_state:
    st.session_state.puntos=[]
if 'fotos' not in st.session_state:
    st.session_state.fotos=[]

tab1, tab2 = st.tabs(["📋 DATOS DE CAMPO R2-POE37", "📸 INFORME FOTOGRAFICO"])

with tab1:
    c1,c2=st.columns(2)
    with c1:
        cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.", key="cli")
        depto=st.text_input("DEPARTAMENTO","TOLIMA", key="dep")
        muni=st.text_input("MUNICIPIO","ATACO", key="mun")
        responsable=st.text_input("Responsable Medicion","edwin cortes", key="resp")
    with c2:
        proyecto=st.text_input("PROYECTO","", key="proy")
        marca_sono=st.text_input("Marca Sonometro","HD2010UC/A", key="sono")

    punto=st.text_input("PUNTO","Punto 1 nocturno")
    coord=st.text_input("COORDENADAS","3°35'11.30\"N 75°23'27.72\"W")
    desc=st.text_area("DESCRIPCION","En la entrada...")
    barrido=st.number_input("BARRIDO dB",0,140,65)
    fecha=st.date_input("Fecha", key="fec")
    hora=st.text_input("Hora","21:01")
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
        thin=Side(style='thin'); border=Border(left=thin,right=thin,top=thin,bottom=thin); no_border=Border(); bold=Font(bold=True,size=10); bold8=Font(bold=True,size=8)
        for k,w in {'A':22,'B':14,'C':16,'D':10,'E':10,'F':10,'G':8,'H':8,'I':8,'J':10,'K':14,'L':12,'M':12,'N':22}.items(): ws.column_dimensions[k].width=w
        ws['A1']="DATOS DE CAMPO EMISION DE RUIDO"; ws['A1'].font=Font(bold=True,size=12)
        ws['A3']="Codigo: R2-POE37-EP"; ws['A3'].border=border; ws['F3']="Version: 04"; ws['F3'].border=border; ws['K3']="Fecha: 2024-05-20"; ws['K3'].border=border
        ws['A5']="CLIENTE:"; ws['A5'].font=bold; ws['A5'].border=border; ws['B5']=cliente; ws['B5'].border=border; ws.merge_cells('B5:E5')
        ws['A7']="DEPARTAMENTO :"; ws['A7'].font=bold; ws['A7'].border=border; ws['B7']=depto; ws['B7'].border=border; ws.merge_cells('B7:E7')
        ws['A8']="MUNICIPIO:"; ws['A8'].font=bold; ws['A8'].border=border; ws['B8']=muni; ws['B8'].border=border; ws.merge_cells('B8:E8')
        r=10
        for p in st.session_state.puntos:
            ws.cell(row=r, column=1, value="PUNTO DE MONITOREO:").font=bold; ws.cell(row=r, column=1).border=border
            ws.cell(row=r, column=2, value=p[0]).border=border; ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            ws.cell(row=r, column=6, value="COORDENADAS ORIGEN").font=bold; ws.cell(row=r, column=6).border=border
            ws.cell(row=r, column=7, value=p[1]).border=border; ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=14); r+=1
            ws.cell(row=r, column=1, value="DESCRIPCION:").font=bold; ws.cell(row=r, column=1).border=border; ws.cell(row=r, column=2, value=p[2]).border=border; ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14); r+=1
            ws.cell(row=r, column=1, value="BARRIDO (dB):").font=bold; ws.cell(row=r, column=1).border=border; ws.cell(row=r, column=2, value=p[3]).border=border; r+=2
            headers=["fecha","Hora","Calibracion","Memoria","LAeq","Vmax","Dir","Temp","Hum","Precip","Fuente","Tipo","Top","Obs"]
            for i,h in enumerate(headers, start=1): c=ws.cell(row=r, column=i, value=h); c.font=bold8; c.border=border; r+=0
            r+=1
            ws.cell(row=r, column=1, value=p[4]).border=border; ws.cell(row=r, column=2, value=p[5]).border=border; ws.cell(row=r, column=3, value=f"Ini {p[6]} Fin {p[7]}").border=border
            ws.cell(row=r, column=4, value=p[8]).border=border; ws.cell(row=r, column=5, value=p[9]).border=border; ws.cell(row=r, column=6, value=p[10]).border=border
            ws.cell(row=r, column=7, value=p[11]).border=border; ws.cell(row=r, column=8, value=p[12]).border=border; ws.cell(row=r, column=9, value=p[13]).border=border
            ws.cell(row=r, column=10, value=p[14]).border=border; ws.cell(row=r, column=11, value=p[15]).border=border; ws.cell(row=r, column=12, value=p[16]).border=border
            ws.cell(row=r, column=13, value=p[17]).border=border; ws.cell(row=r, column=14, value=p[18]).border=border; r+=2
            ws.cell(row=r, column=1, value="DIBUJE EL ESQUEMA").font=bold; r+=1
            foto_r=r; ws.merge_cells(start_row=foto_r, start_column=1, end_row=foto_r+5, end_column=14)
            for rr in range(foto_r, foto_r+6):
                ws.row_dimensions[rr].height=70
                for cc in range(1,15): ws.cell(row=rr, column=cc).border=no_border
            if p[19] and os.path.exists(p[19]):
                try: img=XLImage(p[19]); img.width=850; img.height=380; ws.add_image(img, f'A{foto_r}')
                except: pass
            r=foto_r+7
        r+=2
        ws.cell(row=r, column=1, value="Responsable de la Medicion:").font=Font(bold=True,size=11)
        ws.cell(row=r, column=6, value=responsable).font=Font(bold=True,size=11); ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=14)
        for cc in range(6,15): ws.cell(row=r, column=cc).border=Border(bottom=Side(style='thin'))
        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio
    if st.session_state.puntos:
        st.download_button("DESCARGAR EXCEL CAMPO", build_excel(), f"R2-POE37_{muni}.xlsx", type="primary", use_container_width=True)

with tab2:
    st.subheader("Fotos para informe")
    punto_foto = st.text_input("Punto", "Punto 1 nocturno", key="pf_p")
    desc_foto = st.text_area("Descripcion foto", "Vista general...", key="pf_d")
    fotos_up = st.file_uploader("Sube fotos", type=['jpg','jpeg','png'], accept_multiple_files=True, key="pf_up")

    if st.button("AGREGAR FOTOS", type="primary"):
        for f in fotos_up:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            tmp.write(f.getbuffer()); tmp.close()
            st.session_state.fotos.append({"punto": punto_foto, "desc": desc_foto, "path": tmp.name, "nombre": f.name})

    if st.session_state.fotos:
        st.write(f"{len(st.session_state.fotos)} fotos cargadas")
        for ft in st.session_state.fotos:
            st.write(f"- {ft['punto']} | {ft['nombre']}")

        # CORREGIDO - PDF sin encode
        def build_pdf():
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(0, 10, "REGISTRO FOTOGRAFICO", ln=True, align='C')
            pdf.set_font("Arial", '', 10)
            pdf.cell(0, 8, f"{cliente} - {muni} - {datetime.date.today()}", ln=True, align='C')
            pdf.ln(5)
            for ft in st.session_state.fotos:
                y = pdf.get_y()
                if y > 180:
                    pdf.add_page()
                    y = pdf.get_y()
                try:
                    img = Image.open(ft['path'])
                    img.thumbnail((800, 800))
                    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
                    img.save(tmp_pdf.name, "JPEG")
                    pdf.image(tmp_pdf.name, x=15, y=y, w=180)
                    pdf.set_y(y+95)
                    pdf.set_font("Arial", 'B', 10)
                    pdf.cell(0, 6, ft['punto'], ln=True)
                    pdf.set_font("Arial", '', 9)
                    pdf.multi_cell(0, 5, ft['desc'])
                    pdf.ln(5)
                except Exception as e:
                    pdf.cell(0, 10, f"Error {e}", ln=True)
            # CORRECCION CLAVE
            pdf_bytes = pdf.output()
            return BytesIO(pdf_bytes)

        def build_zip():
            bio = BytesIO()
            with zipfile.ZipFile(bio, 'w') as z:
                for ft in st.session_state.fotos:
                    carpeta = f"Fotos_{muni}/{ft['punto'].replace(' ','_')}/"
                    z.write(ft['path'], arcname=carpeta + ft['nombre'])
            bio.seek(0)
            return bio

        c1, c2 = st.columns(2)
        with c1:
            pdf_data = build_pdf()
            st.download_button("📄 PDF INFORME", pdf_data, f"Registro_{muni}.pdf", mime="application/pdf", type="primary", use_container_width=True)
        with c2:
            zip_data = build_zip()
            st.download_button("📁 ZIP CARPETA", zip_data, f"Fotos_{muni}.zip", mime="application/zip", use_container_width=True)
