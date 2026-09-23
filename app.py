import streamlit as st
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.drawing.image import Image as XLImage
import tempfile, os

st.set_page_config(layout="wide")
st.title("EQUISIMA - V12 FINAL SIN PLANTILLA")

if 'puntos' not in st.session_state:
    st.session_state.puntos=[]

# --- SIDEBAR ---
with st.sidebar:
    cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.")
    proyecto=st.text_input("PROYECTO","")
    depto=st.text_input("DEPARTAMENTO","TOLIMA")
    muni=st.text_input("MUNICIPIO","ATACO")
    marca_sono=st.text_input("Marca Sonometro","HD2010UC/A")
    serie_sono=st.text_input("Serie","15031643825")

# --- PUNTO ---
punto=st.text_input("PUNTO DE MONITOREO","Punto 1 nocturno")
coord=st.text_input("COORDENADAS ORIGEN","3°35'11.30\"N 75°23'27.72\"W")
desc=st.text_area("DESCRIPCION DEL PUNTO","En la entrada...")
barrido=st.number_input("BARRIDO PERIMETRAL dB",0,140,65)
c1,c2=st.columns(2)
fecha=c1.date_input("Fecha")
hora=c2.text_input("Hora","21:01")
cal_i=st.text_input("Cal Ini","114")
cal_f=st.text_input("Cal Fin","114")
mem=st.number_input("Memoria",1,99,1)
laeq=st.number_input("LAeq",0.0,140.0,65.0)
foto=st.file_uploader("FOTO DEL ESQUEMA",type=['jpg','jpeg','png'])

vmax=st.text_input("Vmax","0.3")
vdir=st.text_input("Dir","SW")
temp=st.text_input("Temp","29")
hum=st.text_input("Hum","45")
precip=st.selectbox("Precip",["NO","SI"])
fuente=st.text_input("Fuente","mineria")
tipo=st.text_input("Tipo","EMISION")
top=st.text_input("Tiempo","1")
obs=st.text_area("Observaciones","camino destapado...")

if st.button("AGREGAR PUNTO",type="primary",use_container_width=True):
    fp=None
    if foto:
        f=tempfile.NamedTemporaryFile(delete=False,suffix=".jpg")
        f.write(foto.getbuffer())
        f.close()
        fp=f.name
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs,fp])
    st.success("Punto agregado - ya puedes generar")

def build_excel():
    wb=openpyxl.Workbook()
    ws=wb.active
    ws.title="Datos de Campo Emision"

    thin=Side(style='thin')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    no_border=Border()
    bold=Font(bold=True,size=10)

    # ANCHOS EXACTOS COMO TU ORIGINAL
    ws.column_dimensions['A'].width=22
    ws.column_dimensions['B'].width=14
    ws.column_dimensions['C'].width=16
    ws.column_dimensions['D'].width=10
    ws.column_dimensions['E'].width=10
    ws.column_dimensions['F'].width=10
    ws.column_dimensions['G'].width=8
    ws.column_dimensions['H'].width=8
    ws.column_dimensions['I'].width=8
    ws.column_dimensions['J'].width=10
    ws.column_dimensions['K'].width=14
    ws.column_dimensions['L'].width=12
    ws.column_dimensions['M'].width=12
    ws.column_dimensions['N'].width=22

    # CABECERA
    ws['A1']="DATOS DE CAMPO EMISION DE RUIDO"
    ws['A1'].font=Font(bold=True,size=12)
    ws['A3']="Codigo: R2-POE37-EP"
    ws['A3'].border=border
    ws['F3']="Version: 04"
    ws['F3'].border=border
    ws['K3']="Fecha: 2024-05-20"
    ws['K3'].border=border

    ws['A5']="CLIENTE:"
    ws['A5'].font=bold
    ws['A5'].border=border
    ws['B5']=cliente
    ws['B5'].border=border
    ws.merge_cells('B5:E5')

    ws['I5']="DATOS DEL EQUIPO UTILIZADO"
    ws['I5'].font=bold
    ws['I5'].border=border

    ws['A6']="NOMBRE DEL PROYECTO:"
    ws['A6'].font=bold
    ws['A6'].border=border
    ws['B6']=proyecto
    ws['B6'].border=border
    ws.merge_cells('B6:E6')

    ws['I6']="EQUIPO"
    ws['I6'].font=bold
    ws['I6'].border=border
    ws['K6']="MARCA"
    ws['K6'].font=bold
    ws['K6'].border=border

    ws['A7']="DEPARTAMENTO :"
    ws['A7'].font=bold
    ws['A7'].border=border
    ws['B7']=depto
    ws['B7'].border=border
    ws.merge_cells('B7:E7')

    ws['I7']="SONOMETRO"
    ws['I7'].border=border
    ws['K7']=marca_sono
    ws['K7'].border=border

    ws['A8']="MUNICIPIO:"
    ws['A8'].font=bold
    ws['A8'].border=border
    ws['B8']=muni
    ws['B8'].border=border
    ws.merge_cells('B8:E8')

    ws['I8']="PISTOFONO"
    ws['I8'].border=border
    ws['K8']="HD2020"
    ws['K8'].border=border

    r=10
    for p in st.session_state.puntos:
        ws.cell(row=r, column=1, value="PUNTO DE MONITOREO:").font=bold
        ws.cell(row=r, column=1).border=border
        ws.cell(row=r, column=2, value=p[0]).border=border
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)

        ws.cell(row=r, column=6, value="COORDENADAS ORIGEN").font=bold
        ws.cell(row=r, column=6).border=border
        ws.cell(row=r, column=7, value=p[1]).border=border
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=14)
        r+=1

        ws.cell(row=r, column=1, value="DESCRIPCION DEL PUNTO:").font=bold
        ws.cell(row=r, column=1).border=border
        ws.cell(row=r, column=2, value=p[2]).border=border
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
        r+=1

        ws.cell(row=r, column=1, value="BARRIDO PERIMETRAL (dB):").font=bold
        ws.cell(row=r, column=1).border=border
        ws.cell(row=r, column=2, value=p[3]).border=border
        r+=2

        headers=["fecha","Hora","Calibracion","Memoria","LAeq","Vmax","Dir","Temp","Hum","Precip","Fuente","Tipo","Top","Obs"]
        for i,h in enumerate(headers, start=1):
            c=ws.cell(row=r, column=i, value=h)
            c.font=Font(bold=True,size=8)
            c.border=border
            c.alignment=Alignment(horizontal='center')
        r+=1

        # DATOS YA SEPARADOS CORRECTAMENTE
        ws.cell(row=r, column=1, value=p[4]).border=border
        ws.cell(row=r, column=2, value=p[5]).border=border
        ws.cell(row=r, column=3, value=f"Ini {p[6]} Fin {p[7]}").border=border
        ws.cell(row=r, column=4, value=p[8]).border=border
        ws.cell(row=r, column=5, value=p[9]).border=border # LAeq 65
        ws.cell(row=r, column=6, value=p[10]).border=border # Vmax 0.3 SEPARADO
        ws.cell(row=r, column=7, value=p[11]).border=border
        ws.cell(row=r, column=8, value=p[12]).border=border
        ws.cell(row=r, column=9, value=p[13]).border=border
        ws.cell(row=r, column=10, value=p[14]).border=border
        ws.cell(row=r, column=11, value=p[15]).border=border
        ws.cell(row=r, column=12, value=p[16]).border=border
        ws.cell(row=r, column=13, value=p[17]).border=border
        ws.cell(row=r, column=14, value=p[18]).border=border
        r+=2

        ws.cell(row=r, column=1, value="DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO").font=bold
        r+=1

        # FOTO SIN RECUADROS
        foto_r=r
        ws.merge_cells(start_row=foto_r, start_column=1, end_row=foto_r+5, end_column=14)
        for rr in range(foto_r, foto_r+6):
            ws.row_dimensions[rr].height=70
            for cc in range(1,15):
                ws.cell(row=rr, column=cc).border=no_border

        if p[19] and os.path.exists(p[19]):
            try:
                img=XLImage(p[19])
                img.width=850
                img.height=380
                ws.add_image(img, f'A{foto_r}')
            except:
                pass
        r=foto_r+7

    bio=BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

if st.session_state.puntos:
    excel=build_excel()
    st.download_button("DESCARGAR EXCEL FINAL SIN ERRORES", excel, f"R2-POE37-EP_{muni}_FINAL.xlsx", type="primary", use_container_width=True)
