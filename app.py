import streamlit as st, pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill
from openpyxl.drawing.image import Image as XLImage
import tempfile, os

st.set_page_config(layout="wide")
st.title("🟡 EQUISIMA - Sin recuadros en foto")

if 'puntos' not in st.session_state: st.session_state.puntos=[]
thin=Side(style='thin'); border_all=Border(left=thin,right=thin,top=thin,bottom=thin)
no_border=Border()

with st.sidebar:
    cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.")
    muni=st.text_input("MUNICIPIO","ATACO")

punto=st.text_input("PUNTO","Punto 1 nocturno")
coord=st.text_input("COORD","3°35'11.30\"N 75°23'27.72\"W")
desc=st.text_area("DESCRIPCIÓN","En la entrada...")
barrido=st.number_input("BARRIDO",0,140,65)
fecha=st.date_input("Fecha")
hora=st.text_input("Hora","21:01")
cal_i=st.text_input("Cal Ini","114")
cal_f=st.text_input("Cal Fin","114")
mem=st.number_input("Mem",1,99,1)
laeq=st.number_input("LAeq",0.0,140.0,65.0)
foto=st.file_uploader("📸 Foto Earth",type=['jpg','png','jpeg'])
vmax=st.text_input("Vmax","0.3"); vdir=st.text_input("Dir","SW"); temp=st.text_input("Temp","29"); hum=st.text_input("Hum","45")
precip=st.selectbox("Precip",["NO","SI"]); fuente=st.text_input("Fuente","mineria"); tipo=st.text_input("Tipo","EMISION"); top=st.text_input("Top","1")
obs=st.text_area("Obs","camino destapado...")

if st.button("➕ AGREGAR PUNTO",type="primary",use_container_width=True):
    fp=None
    if foto:
        with tempfile.NamedTemporaryFile(delete=False,suffix=".jpg") as t:
            t.write(foto.getbuffer()); fp=t.name
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs,fp])

if st.session_state.puntos:
    def build():
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="Datos de Campo Emision"
        for k,w in {'A':2,'B':22,'C':15,'D':18,'E':20,'F':18,'G':14,'H':14,'I':12,'J':12,'K':22,'L':20,'M':14,'N':14,'O':38}.items():
            ws.column_dimensions[k].width=w

        def celdab(r,c,v,b=False):
            cell=ws.cell(row=r,column=c,value=v)
            if b: cell.font=Font(bold=True,size=9)
            cell.border=border_all
            return cell

        ws.merge_cells('B2:O2'); ws['B2']="DATOS DE CAMPO EMISION DE RUIDO"; ws['B2'].font=Font(bold=True,size=12); ws['B2'].border=border_all
        ws['B4']="Codigo: R2-POE37-EP"; ws['B4'].border=border_all; ws['F4']="Version: 04"; ws['F4'].border=border_all; ws['K4']="Fecha: 2024-05-20"; ws['K4'].border=border_all
        r=6
        celdab(r,2,"CLIENTE:",True); ws.merge_cells(f'C{r}:H{r}'); ws[f'C{r}']=cliente; ws[f'C{r}'].border=border_all
        ws.merge_cells(f'I{r}:O{r}'); ws[f'I{r}']="DATOS DEL EQUIPO UTILIZADO"; ws[f'I{r}'].font=Font(bold=True); ws[f'I{r}'].border=border_all
        r=7
        celdab(r,2,"NOMBRE DEL PROYECTO:",True); ws.merge_cells(f'C{r}:H{r}'); ws[f'C{r}'].border=border_all
        ws.merge_cells(f'I{r}:J{r}'); celdab(r,9,"EQUIPO",True); ws.merge_cells(f'K{r}:L{r}'); celdab(r,11,"MARCA",True); ws.merge_cells(f'M{r}:O{r}'); celdab(r,13,"SERIE",True); r=8
        celdab(r,2,"DEPARTAMENTO :",True); ws.merge_cells(f'C{r}:H{r}'); ws[f'C{r}']="TOLIMA"; ws[f'C{r}'].border=border_all
        ws.merge_cells(f'I{r}:J{r}'); ws[f'I{r}']="SONÓMETRO"; ws[f'I{r}'].border=border_all; ws.merge_cells(f'K{r}:L{r}'); ws[f'K{r}']="HD2010UC/A"; ws[f'K{r}'].border=border_all; ws.merge_cells(f'M{r}:O{r}'); ws[f'M{r}']="15031643825"; ws[f'M{r}'].border=border_all; r=9
        celdab(r,2,"MUNICIPIO:",True); ws.merge_cells(f'C{r}:H{r}'); ws[f'C{r}']=muni; ws[f'C{r}'].border=border_all; ws.merge_cells(f'I{r}:J{r}'); ws[f'I{r}']="PISTÓFONO"; ws[f'I{r}'].border=border_all; r=12

        for p in st.session_state.puntos:
            celdab(r,2,"PUNTO DE MONITOREO:",True); ws.merge_cells(f'C{r}:E{r}'); ws[f'C{r}']=p[0]; ws[f'C{r}'].border=border_all
            celdab(r,6,"COORDENADAS ORIGEN NACIONAL:",True); ws.merge_cells(f'G{r}:O{r}'); ws[f'G{r}']=p[1]; ws[f'G{r}'].border=border_all; r+=1
            celdab(r,2,"DESCRIPCIÓN DEL PUNTO:",True); ws.merge_cells(f'C{r}:O{r}'); ws[f'C{r}']=p[2]; ws[f'C{r}'].border=border_all; r+=1
            celdab(r,2,"BARRIDO PERIMETRAL (dB):",True); ws[f'C{r}']=p[3]; ws[f'C{r}'].border=border_all; r+=1
            headers=["Fecha","Hora","Calibración","Memoria","LAeq","Vmax","Dir","Temp","Hum","Precip","Fuente","Tipo","Top","Obs"]
            for i,h in enumerate(headers, start=2):
                ws.cell(row=r,column=i,value=h).font=Font(bold=True,size=7); ws.cell(row=r,column=i).border=border_all
            r+=1
            vals=[p[4],p[5],f"Ini {p[6]} Fin {p[7]}",p[8],p[9],p[10],p[11],p[12],p[13],p[14],p[15],p[16],p[17],p[18]]
            for i,v in enumerate(vals, start=2):
                ws.cell(row=r,column=i,value=v).border=border_all
            r+=2

            # AREA FOTO SIN RECUADROS
            ws[f'B{r}']="DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"; ws[f'B{r}'].font=Font(bold=True); r+=1
            foto_r=r
            ws.merge_cells(f'B{foto_r}:O{foto_r+5}')
            for rr in range(foto_r, foto_r+6):
                ws.row_dimensions[rr].height=65
                for cc in range(1,16):
                    ws.cell(row=rr,column=cc).border=no_border
            if p[19] and os.path.exists(p[19]):
                img=XLImage(p[19]); img.width=900; img.height=400
                ws.add_image(img, f'B{foto_r}')
            r=foto_r+7

        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio
    st.download_button("📥 DESCARGAR SIN RECUADROS EN FOTO", build(), f"ORIGINAL_{muni}.xlsx", type="primary", use_container_width=True)
