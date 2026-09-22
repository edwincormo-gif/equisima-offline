import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
import tempfile, os

st.set_page_config(page_title="EQUISIMA ORIGINAL", layout="wide")
st.title("🟡 EQUISIMA - Original con recuadros")

if 'puntos' not in st.session_state: st.session_state.puntos=[]

# --- ESTILOS ORIGINALES ---
thin = Side(style='thin', color='000000')
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
bold8 = Font(bold=True, size=8)
bold10 = Font(bold=True, size=10)

with st.sidebar:
    cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
    proyecto = st.text_input("PROYECTO", "")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    muni = st.text_input("MUNICIPIO", "ATACO")
    resp = st.text_input("Responsable", "edwin cortes")
    st.divider()
    st.write("EQUIPOS (como original)")
    eq_sono = st.text_input("Sonómetro Marca", "HD2010UC/A")
    serie_sono = st.text_input("Serie Sonómetro", "15031643825")
    eq_pist = st.text_input("Pistófono", "HD2020")

st.subheader("Nuevo Punto")
punto = st.text_input("PUNTO", "Punto 1 nocturno")
coord = st.text_input("COORDENADAS ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W")
desc = st.text_area("DESCRIPCIÓN", "En la entrada o vía principal...")
barrido = st.number_input("BARRIDO dB", 0, 140, 65)
c1,c2,c3 = st.columns(3)
fecha = c1.date_input("Fecha")
hora = c2.text_input("Hora", "21:01")
cal_i = c3.text_input("Cal Ini", "114")
cal_f = st.text_input("Cal Final", "114")
mem = st.number_input("Memoria", 1, 99, 1)
laeq = st.number_input("LAeq", 0.0, 140.0, 65.0)
foto = st.file_uploader("📸 Foto Earth / Esquema", type=['jpg','png','jpeg'])

vmax = st.text_input("Vmax", "0.3")
vdir = st.text_input("Dir", "SW")
temp = st.text_input("Temp", "29")
hum = st.text_input("Hum", "45")
precip = st.selectbox("Precip", ["NO","SI"])
fuente = st.text_input("Fuente", "mineria")
tipo = st.text_input("Tipo", "EMISION")
top = st.text_input("Tiempo Op", "1")
obs = st.text_area("Observaciones", "camino destapado...")

if st.button("➕ AGREGAR PUNTO", type="primary", use_container_width=True):
    fpath=None
    if foto:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
            tmp.write(foto.getbuffer()); fpath=tmp.name
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs,fpath])

if st.session_state.puntos:
    st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)

    def build():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title="Datos de Campo Emision"
        # Anchos exactos
        for k,w in {'A':2,'B':22,'C':15,'D':18,'E':20,'F':18,'G':14,'H':14,'I':12,'J':12,'K':22,'L':20,'M':14,'N':14,'O':38}.items():
            ws.column_dimensions[k].width=w

        def celda(r,c,valor,bold=False,wrap=False,center=False):
            cell=ws.cell(row=r,column=c,value=valor)
            if bold: cell.font=Font(bold=True,size=9)
            cell.border=border_all
            if wrap: cell.alignment=Alignment(wrap_text=True, vertical='center')
            if center: cell.alignment=Alignment(horizontal='center', vertical='center', wrap_text=True)
            return cell

        # TITULO CON RECUADRO
        ws.merge_cells('B2:O2'); ws['B2']="DATOS DE CAMPO EMISION DE RUIDO"; ws['B2'].font=Font(bold=True,size=12); ws['B2'].alignment=Alignment(horizontal='center'); ws['B2'].border=border_all

        ws['B4']="Codigo: R2-POE37-EP"; ws['F4']="Version: 04"; ws['K4']="Fecha: 2024-05-20"
        for c in [2,6,11]: ws.cell(row=4,column=c).border=border_all

        # CLIENTE CON RECUADROS
        r=6
        celda(r,2,"CLIENTE:",True); celda(r,3,cliente); ws.merge_cells(f'C{r}:H{r}'); celda(r,9,"DATOS DEL EQUIPO UTILIZADO",True,center=True); ws.merge_cells(f'I{r}:O{r}')
        r=7
        celda(r,2,"NOMBRE DEL PROYECTO:",True); celda(r,3,proyecto); ws.merge_cells(f'C{r}:H{r}'); celda(r,9,"EQUIPO",True,center=True); celda(r,11,"MARCA",True,center=True); celda(r,13,"SERIE",True,center=True); ws.merge_cells(f'I{r}:J{r}'); ws.merge_cells(f'K{r}:L{r}'); ws.merge_cells(f'M{r}:O{r}')
        r=8
        celda(r,2,"DEPARTAMENTO :",True); celda(r,3,depto); ws.merge_cells(f'C{r}:H{r}'); celda(r,9,"SONÓMETRO"); celda(r,11,eq_sono); celda(r,13,serie_sono); ws.merge_cells(f'I{r}:J{r}'); ws.merge_cells(f'K{r}:L{r}'); ws.merge_cells(f'M{r}:O{r}')
        r=9
        celda(r,2,"MUNICIPIO:",True); celda(r,3,muni); ws.merge_cells(f'C{r}:H{r}'); celda(r,9,"PISTÓFONO"); celda(r,11,eq_pist); ws.merge_cells(f'I{r}:J{r}'); ws.merge_cells(f'K{r}:L{r}'); ws.merge_cells(f'M{r}:O{r}')
        r=10
        celda(r,2,"PLAN DE MUESTREO:",True); celda(r,3,"EMISIÓN RUIDO"); ws.merge_cells(f'C{r}:H{r}'); celda(r,9,"ESTACIÓN METEOROLÓGICA"); ws.merge_cells(f'I{r}:O{r}')

        r=12
        for p in st.session_state.puntos:
            # PUNTO CON RECUADROS
            celda(r,2,"PUNTO DE MONITOREO:",True); ws.merge_cells(f'C{r}:E{r}'); celda(r,3,p[0]); celda(r,6,"COORDENADAS ORIGEN NACIONAL:",True); ws.merge_cells(f'F{r}:I{r}'); celda(r,10,p[1]); ws.merge_cells(f'J{r}:O{r}'); r+=1
            celda(r,2,"DESCRIPCIÓN DEL PUNTO:",True); ws.merge_cells(f'C{r}:O{r}'); celda(r,3,p[2],wrap=True); ws.row_dimensions[r].height=40; r+=1
            celda(r,2,"BARRIDO PERIMETRAL (dB):",True); celda(r,3,p[3],center=True); r+=1

            # TABLA MEDICION CON RECUADROS
            headers=["Fecha","Hora","Calibración (dB)","Memoria","LAeq,T (dB)\nIn situ","Viento Max\n(m/s)","Dir Viento","Temp\n(°C)","Hum\n(%)","Precipitación","Fuente Ruido","Tipo Ruido","Tiempo Op","Observaciones"]
            for i,h in enumerate(headers, start=2):
                c=celda(r,i,h,True,True,True); c.font=Font(bold=True,size=7)
            ws.row_dimensions[r].height=35; r+=1
            vals=[p[4],p[5],f"Ini {p[6]}\nFin {p[7]}",p[8],p[9],p[10],p[11],p[12],p[13],p[14],p[15],p[16],p[17],p[18]]
            for i,v in enumerate(vals, start=2):
                celda(r,i,v,center=True,wrap=True)
            r+=2

            # ESQUEMA CON RECUADRO GRANDE PARA FOTO
            celda(r,2,"DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO",True); ws.merge_cells(f'B{r}:O{r}'); r+=1
            # Recuadro de 4 filas x 14 columnas para foto
            foto_r=r
            for rr in range(foto_r, foto_r+5):
                for cc in range(2,16):
                    ws.cell(row=rr,column=cc).border=border_all
                ws.row_dimensions[rr].height=60

            if p[19] and os.path.exists(p[19]):
                try:
                    img=XLImage(p[19]); img.width=650; img.height=280
                    ws.add_image(img, f'C{foto_r}')
                except: pass
            else:
                ws[f'C{foto_r}']="ESPACIO PARA FOTO EARTH"
            r=foto_r+6

        ws[f'B{r}']="Responsable:"; ws[f'C{r}']=resp
        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

    st.download_button("📥 DESCARGAR ORIGINAL CON RECUADROS Y FOTO", build(), f"R2-POE37-EP_{muni}_ORIGINAL.xlsx", type="primary", use_container_width=True)
