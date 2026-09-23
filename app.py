import streamlit as st
from io import BytesIO
import openpyxl
from openpyxl.styles import Border
from openpyxl.drawing.image import Image as XLImage
import tempfile, os

st.set_page_config(layout="wide")
st.title("🟡 EQUISIMA - Con tu plantilla original")

if 'puntos' not in st.session_state: st.session_state.puntos=[]

with st.sidebar:
    cliente=st.text_input("CLIENTE","Rueda Inversiones S.A.S.")
    proyecto=st.text_input("PROYECTO","")
    depto=st.text_input("DEPARTAMENTO","TOLIMA")
    muni=st.text_input("MUNICIPIO","ATACO")
    resp=st.text_input("Responsable","edwin cortes")

punto=st.text_input("PUNTO","Punto 1 nocturno")
coord=st.text_input("COORDENADAS","3°35'11.30\"N 75°23'27.72\"W")
desc=st.text_area("DESCRIPCIÓN","En la entrada...")
barrido=st.number_input("BARRIDO dB",0,140,65)
c1,c2=st.columns(2)
fecha=c1.date_input("Fecha")
hora=c2.text_input("Hora","21:01")
cal_i=st.text_input("Cal Ini","114")
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
    st.toast("Punto agregado")

if st.session_state.puntos:
    def build_con_plantilla():
        try:
            wb=openpyxl.load_workbook("plantilla.xlsx")
            ws=wb["Datos de Campo Emision"]
        except Exception as e:
            st.error(f"No encuentro plantilla.xlsx: {e}")
            return None

        # --- ESCRIBIR CABECERA BUSCANDO EL TEXTO ---
        for r in range(1,20):
            for c in range(1,8):
                v=ws.cell(r,c).value
                if not v: continue
                sv=str(v)
                if "CLIENTE" in sv: ws.cell(r,c+1).value=cliente
                if "PROYECTO" in sv: ws.cell(r,c+1).value=proyecto
                if "DEPARTAMENTO" in sv: ws.cell(r,c+1).value=depto
                if "MUNICIPIO" in sv: ws.cell(r,c+1).value=muni
                if "PUNTO DE MONITOREO" in sv and r>10:
                    # Este es el primer punto, lo usamos como base
                    start_r=r
                    # Guardar fila inicio
                    ws._start_row=start_r

        # Buscar fila donde empieza el punto
        start_r=getattr(ws,'_start_row',13)

        # Limpiar filas de ejemplo si hay más de 1 punto en plantilla
        # Borramos desde start_r hasta 100 para reescribir limpio (manteniendo merges)
        # No borramos merges, solo valores

        r=start_r
        no_border=Border()
        for p in st.session_state.puntos:
            # PUNTO Y COORDENADA
            ws.cell(r,2).value="PUNTO DE MONITOREO:"; ws.cell(r,3).value=p[0]
            ws.cell(r,6).value="COORDENADAS ORIGEN"; ws.cell(r,7).value=p[1]
            r+=1
            ws.cell(r,2).value="DESCRIPCIÓN DEL PUNTO:"; ws.cell(r,3).value=p[2]
            r+=1
            ws.cell(r,2).value="BARRIDO PERIMETRAL (dB):"; ws.cell(r,3).value=p[3]
            r+=2
            # CABECERA tabla
            ws.cell(r,1).value="fecha"; ws.cell(r,2).value="Hora"
            ws.cell(r,3).value="Calibración"; ws.cell(r,4).value="Memoria"
            ws.cell(r,5).value="LAeq"; ws.cell(r,6).value="Vmax"
            ws.cell(r,7).value="Dir"; ws.cell(r,8).value="Temp"
            ws.cell(r,9).value="Hum"; ws.cell(r,10).value="Precip"
            ws.cell(r,11).value="Fuente"; ws.cell(r,12).value="Tipo"
            ws.cell(r,13).value="Top"; ws.cell(r,14).value="Obs"
            r+=1
            # DATOS SEPARADOS - AQUI SE ARREGLA EL 65 0,3
            ws.cell(r,1).value=p[4] # fecha
            ws.cell(r,2).value=p[5] # hora
            ws.cell(r,3).value=f"Ini {p[6]} Fin {p[7]}"
            ws.cell(r,4).value=p[8] # memoria 1
            ws.cell(r,5).value=p[9] # LAeq 65 SOLO
            ws.cell(r,6).value=p[10] # Vmax 0.3 SOLO
            ws.cell(r,7).value=p[11] # SW
            ws.cell(r,8).value=p[12] # 29
            ws.cell(r,9).value=p[13] # 45
            ws.cell(r,10).value=p[14] # NO
            ws.cell(r,11).value=p[15] # mineria
            ws.cell(r,12).value=p[16] # EMISION
            ws.cell(r,13).value=p[17] # 1
            ws.cell(r,14).value=p[18] # camino...
            r+=2

            ws.cell(r,1).value="DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"
            r+=1
            foto_r=r
            # QUITAR RECUADROS EN ZONA DE FOTO COMO PEDISTE
            for rr in range(foto_r, foto_r+6):
                ws.row_dimensions[rr].height=65
                for cc in range(1,15):
                    ws.cell(row=rr,column=cc).border=no_border

            if p[19] and os.path.exists(p[19]):
                try:
                    img=XLImage(p[19]); img.width=900; img.height=400
                    ws.add_image(img, f'B{foto_r}')
                except Exception as e:
                    st.warning(f"No se pudo insertar foto: {e}")
            r=foto_r+7

        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

    st.download_button("📥 DESCARGAR CON TU PLANTILLA ORIGINAL", build_con_plantilla(), f"R2-POE37-EP_{muni}_CON_PLANTILLA.xlsx", type="primary", use_container_width=True)
