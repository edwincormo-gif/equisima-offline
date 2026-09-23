import streamlit as st, pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Border, Side
from openpyxl.drawing.image import Image as XLImage
import tempfile, os

st.set_page_config(layout="wide")
st.title("🟡 EQUISIMA - V8 Corregido")

if 'puntos' not in st.session_state: st.session_state.puntos=[]
thin=Side(style='thin'); border=Border(left=thin,right=thin,top=thin,bottom=thin)
no_border=Border()

with st.sidebar:
    cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.")
    proyecto=st.text_input("PROYECTO","")
    depto=st.text_input("DEPARTAMENTO","TOLIMA")
    muni=st.text_input("MUNICIPIO","ATACO")
    resp=st.text_input("Responsable","edwin cortes")
    sono_m=st.text_input("Marca Sonómetro","HD2010UC/A")
    sono_s=st.text_input("Serie","15031643825")
    pisto=st.text_input("Pistófono","HD2020")

punto=st.text_input("PUNTO","Punto 1 nocturno")
coord=st.text_input("COORDENADAS ORIGEN","3°35'11.30\"N 75°23'27.72\"W")
desc=st.text_area("DESCRIPCIÓN","En la entrada...")
barrido=st.number_input("BARRIDO PERIMETRAL dB",0,140,65)
c1,c2,c3=st.columns(3)
fecha=c1.date_input("Fecha")
hora=c2.text_input("Hora","21:01")
cal_i=c3.text_input("Cal Ini","114")
cal_f=st.text_input("Cal Fin","114")
mem=st.number_input("Memoria",1,99,1)
laeq=st.number_input("LAeq",0.0,140.0,65.0)
foto=st.file_uploader("📸 Foto Earth",type=['jpg','png','jpeg'])
vmax=st.text_input("Vmax","0.3")
vdir=st.text_input("Dir","SW")
temp=st.text_input("Temp","29")
hum=st.text_input("Hum","45")
precip=st.selectbox("Precip",["NO","SI"])
fuente=st.text_input("Fuente","mineria")
tipo=st.text_input("Tipo","EMISION")
top=st.text_input("Tiempo","1")
obs=st.text_area("Obs","camino destapado...")

if st.button("➕ AGREGAR PUNTO",type="primary",use_container_width=True):
    fp=None
    if foto:
        with tempfile.NamedTemporaryFile(delete=False,suffix=".jpg") as t:
            t.write(foto.getbuffer()); fp=t.name
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs,fp])
    st.toast("Agregado!")

if st.session_state.puntos:
    def build():
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="Datos de Campo Emision"
        # ANCHOS EXACTOS
        widths={'A':2,'B':20,'C':12,'D':16,'E':10,'F':10,'G':10,'H':8,'I':8,'J':8,'K':10,'L':14,'M':12,'N':12,'O':20}
        for k,w in widths.items(): ws.column_dimensions[k].width=w

        # CABECERA FIJA CON RECUADROS
        ws['A1']="DATOS DE CAMPO EMISION DE RUIDO"; ws['A1'].font=Font(bold=True, size=11)
        ws['A3']="Codigo: R2-POE37-EP"; ws['A3'].border=border
        ws['F3']="Version: 04"; ws['F3'].border=border
        ws['K3']="Fecha: 2024-05-20"; ws['K3'].border=border

        ws['A5']="CLIENTE:"; ws['A5'].font=Font(bold=True); ws['A5'].border=border
        ws['B5']=cliente; ws['B5'].border=border; ws.merge_cells('B5:E5')
        ws['I5']="DATOS DEL EQUIPO UTILIZADO"; ws['I5'].font=Font(bold=True); ws['I5'].border=border; ws.merge_cells('I5:L5')

        ws['A6']="NOMBRE DEL PROYECTO:"; ws['A6'].font=Font(bold=True); ws['A6'].border=border
        ws['B6']=proyecto; ws['B6'].border=border; ws.merge_cells('B6:E6')
        ws['I6']="EQUIPO"; ws['I6'].font=Font(bold=True); ws['I6'].border=border
        ws['K6']="MARCA"; ws['K6'].font=Font(bold=True); ws['K6'].border=border

        ws['A7']="DEPARTAMENTO :"; ws['A7'].font=Font(bold=True); ws['A7'].border=border
        ws['B7']=depto; ws['B7'].border=border; ws.merge_cells('B7:E7')
        ws['I7']="SONÓMETRO"; ws['I7'].border=border; ws['K7']=sono_m; ws['K7'].border=border; ws.merge_cells('K7:L7')

        ws['A8']="MUNICIPIO:"; ws['A8'].font=Font(bold=True); ws['A8'].border=border
        ws['B8']=muni; ws['B8'].border=border; ws.merge_cells('B8:E8')
        ws['I8']="PISTÓFONO"; ws['I8'].border=border; ws['K8']=pisto; ws['K8'].border=border; ws.merge_cells('K8:L8')

        r=10
        for p in st.session_state.puntos:
            ws[f'A{r}']="PUNTO DE MONITOREO:"; ws[f'A{r}'].font=Font(bold=True); ws[f'A{r}'].border=border
            ws[f'B{r}']=p[0]; ws[f'B{r}'].border=border; ws.merge_cells(f'B{r}:E{r}')
            ws[f'F{r}']="COORDENADAS ORIGEN"; ws[f'F{r}'].font=Font(bold=True); ws[f'F{r}'].border=border
            ws[f'G{r}']=p[1]; ws[f'G{r}'].border=border; ws.merge_cells(f'G{r}:O{r}'); r+=1

            ws[f'A{r}']="DESCRIPCIÓN DEL PUNTO:"; ws[f'A{r}'].font=Font(bold=True); ws[f'A{r}'].border=border
            ws[f'B{r}']=p[2]; ws[f'B{r}'].border=border; ws.merge_cells(f'B{r}:O{r}'); r+=1

            ws[f'A{r}']="BARRIDO PERIMETRAL (dB):"; ws[f'A{r}'].font=Font(bold=True); ws[f'A{r}'].border=border
            ws[f'B{r}']=p[3]; ws[f'B{r}'].border=border; r+=1

            # TABLA - COLUMNAS SEPARADAS SIN JUNTARSE
            headers=["fecha","Hora","Calibración","Memoria","LAeq","Vmax","Dir","Temp","Hum","Precip","Fuente","Tipo","Top","Obs"]
            for i,h in enumerate(headers, start=1):
                ws.cell(row=r,column=i,value=h).font=Font(bold=True,size=8)
                ws.cell(row=r,column=i).border=border
            r+=1

            # DATOS - CADA UNO EN SU CELDA
            ws.cell(row=r,column=1,value=p[4]).border=border
            ws.cell(row=r,column=2,value=p[5]).border=border
            ws.cell(row=r,column=3,value=f"Ini {p[6]} Fin {p[7]}").border=border
            ws.cell(row=r,column=4,value=p[8]).border=border
            ws.cell(row=r,column=5,value=p[9]).border=border # LAeq SOLO
            ws.cell(row=r,column=6,value=p[10]).border=border # Vmax SOLO
            ws.cell(row=r,column=7,value=p[11]).border=border
            ws.cell(row=r,column=8,value=p[12]).border=border
            ws.cell(row=r,column=9,value=p[13]).border=border
            ws.cell(row=r,column=10,value=p[14]).border=border
            ws.cell(row=r,column=11,value=p[15]).border=border
            ws.cell(row=r,column=12,value=p[16]).border=border
            ws.cell(row=r,column=13,value=p[17]).border=border
            ws.cell(row=r,column=14,value=p[18]).border=border
            r+=2

            ws[f'A{r}']="DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"; ws[f'A{r}'].font=Font(bold=True); r+=1
            foto_r=r
            # SIN RECUADROS EN ESTA ZONA
            ws.merge_cells(f'A{foto_r}:O{foto_r+5}')
            for rr in range(foto_r, foto_r+6):
                ws.row_dimensions[rr].height=60
                for cc in range(1,16):
                    ws.cell(row=rr,column=cc).border=no_border
            if p[19] and os.path.exists(p[19]):
                try:
                    img=XLImage(p[19]); img.width=800; img.height=350
                    ws.add_image(img, f'A{foto_r}')
                except: pass
            r=foto_r+7

        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

    st.download_button("📥 DESCARGAR CORREGIDO (LAeq y Vmax separados)", build(), f"CAMPO_{muni}_CORREGIDO.xlsx", type="primary", use_container_width=True)
