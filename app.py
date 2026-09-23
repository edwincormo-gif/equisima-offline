import streamlit as st
from io import BytesIO
import openpyxl
from openpyxl.styles import Border
from openpyxl.drawing.image import Image as XLImage
import tempfile, os

st.set_page_config(layout="wide")
st.title("EQUISIMA - V11 con plantilla")

if 'puntos' not in st.session_state:
    st.session_state.puntos=[]

with st.sidebar:
    cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.")
    depto=st.text_input("DEPARTAMENTO","TOLIMA")
    muni=st.text_input("MUNICIPIO","ATACO")

punto=st.text_input("PUNTO","Punto 1 nocturno")
coord=st.text_input("COORDENADAS","3°35'11.30\"N 75°23'27.72\"W")
desc=st.text_area("DESCRIPCION","En la entrada...")
barrido=st.number_input("BARRIDO dB",0,140,65)
fecha=st.date_input("Fecha")
hora=st.text_input("Hora","21:01")
cal_i=st.text_input("Cal Ini","114")
cal_f=st.text_input("Cal Fin","114")
mem=st.number_input("Memoria",1,99,1)
laeq=st.number_input("LAeq",0.0,140.0,65.0)
foto=st.file_uploader("Foto Earth",type=['jpg','png','jpeg'])

vmax=st.text_input("Vmax","0.3")
vdir=st.text_input("Dir","SW")
temp=st.text_input("Temp","29")
hum=st.text_input("Hum","45")
precip=st.selectbox("Precip",["NO","SI"])
fuente=st.text_input("Fuente","mineria")
tipo=st.text_input("Tipo","EMISION")
top=st.text_input("Tiempo","1")
obs=st.text_area("Obs","camino destapado...")

if st.button("AGREGAR PUNTO",type="primary",use_container_width=True):
    fp=None
    if foto:
        tmp=tempfile.NamedTemporaryFile(delete=False,suffix=".jpg")
        tmp.write(foto.getbuffer())
        tmp.close()
        fp=tmp.name
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs,fp])
    st.success("Agregado")

def build():
    if not os.path.exists("plantilla.xlsx"):
        st.error("No encuentro plantilla.xlsx - Súbela a la raiz del repo con ese nombre exacto")
        return None

    wb=openpyxl.load_workbook("plantilla.xlsx")
    ws=wb["Datos de Campo Emision"]

    # 1. DESCOMBINAR CABECERA PARA EVITAR EL ERROR AttributeError
    merges=list(ws.merged_cells.ranges)
    for mr in merges:
        if mr.min_row <= 12:
            ws.unmerge_cells(str(mr))

    # 2. ESCRIBIR CABECERA BUSCANDO TEXTO
    start_r=13
    for r in range(1,20):
        for c in range(1,10):
            v=ws.cell(row=r, column=c).value
            if not v:
                continue
            sv=str(v).upper()
            if "CLIENTE" in sv:
                ws.cell(row=r, column=c+1).value=cliente
            if "DEPARTAMENTO" in sv:
                ws.cell(row=r, column=c+1).value=depto
            if "MUNICIPIO" in sv:
                ws.cell(row=r, column=c+1).value=muni
            if "PUNTO DE MONITOREO" in sv:
                start_r=r

    r=start_r
    no_border=Border()

    for p in st.session_state.puntos:
        # Escribir punto - ya separado LAeq y Vmax
        ws.cell(row=r, column=2).value="PUNTO DE MONITOREO:"
        ws.cell(row=r, column=3).value=p[0]
        ws.cell(row=r, column=6).value="COORDENADAS ORIGEN"
        ws.cell(row=r, column=7).value=p[1]
        r+=1

        ws.cell(row=r, column=2).value="DESCRIPCION DEL PUNTO:"
        ws.cell(row=r, column=3).value=p[2]
        r+=1

        ws.cell(row=r, column=2).value="BARRIDO PERIMETRAL (dB):"
        ws.cell(row=r, column=3).value=p[3]
        r+=2

        ws.cell(row=r, column=1).value="fecha"
        ws.cell(row=r, column=2).value="Hora"
        ws.cell(row=r, column=3).value="Calibracion"
        ws.cell(row=r, column=4).value="Memoria"
        ws.cell(row=r, column=5).value="LAeq"
        ws.cell(row=r, column=6).value="Vmax"
        ws.cell(row=r, column=7).value="Dir"
        ws.cell(row=r, column=8).value="Temp"
        ws.cell(row=r, column=9).value="Hum"
        ws.cell(row=r, column=10).value="Precip"
        ws.cell(row=r, column=11).value="Fuente"
        ws.cell(row=r, column=12).value="Tipo"
        ws.cell(row=r, column=13).value="Top"
        ws.cell(row=r, column=14).value="Obs"
        r+=1

        ws.cell(row=r, column=1).value=p[4]
        ws.cell(row=r, column=2).value=p[5]
        ws.cell(row=r, column=3).value=f"Ini {p[6]} Fin {p[7]}"
        ws.cell(row=r, column=4).value=p[8]
        ws.cell(row=r, column=5).value=p[9]
        ws.cell(row=r, column=6).value=p[10]
        ws.cell(row=r, column=7).value=p[11]
        ws.cell(row=r, column=8).value=p[12]
        ws.cell(row=r, column=9).value=p[13]
        ws.cell(row=r, column=10).value=p[14]
        ws.cell(row=r, column=11).value=p[15]
        ws.cell(row=r, column=12).value=p[16]
        ws.cell(row=r, column=13).value=p[17]
        ws.cell(row=r, column=14).value=p[18]
        r+=2

        ws.cell(row=r, column=1).value="DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"
        r+=1
        foto_r=r
        for rr in range(foto_r, foto_r+6):
            ws.row_dimensions[rr].height=70
            for cc in range(1,15):
                ws.cell(row=rr, column=cc).border=no_border

        if p[19] and os.path.exists(p[19]):
            try:
                img=XLImage(p[19])
                img.width=900
                img.height=380
                ws.add_image(img, f'B{foto_r}')
            except:
                pass
        r=foto_r+7

    bio=BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

if st.session_state.puntos:
    st.divider()
    if st.button("GENERAR EXCEL CON PLANTILLA", type="primary"):
        excel_data=build()
        if excel_data:
            st.download_button("DESCARGAR EXCEL FINAL", excel_data, f"R2-POE37-EP_{muni}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
