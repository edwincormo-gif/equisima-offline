import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import tempfile, os

st.set_page_config(page_title="EQUISIMA CON FOTO", layout="wide")
st.title("🟡 EQUISIMA - Con Foto Earth")

if 'puntos' not in st.session_state: st.session_state.puntos=[]

with st.sidebar:
    cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
    proyecto = st.text_input("PROYECTO", "rd-eq15-2026")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    muni = st.text_input("MUNICIPIO", "ATACO")
    resp = st.text_input("Responsable", "edwin cortes")

st.subheader("Nuevo Punto + Foto Earth")
punto = st.text_input("PUNTO", "Punto 1 nocturno")
coord = st.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W")
desc = st.text_area("DESCRIPCIÓN", "En la entrada o vía principal...")
barrido = st.number_input("BARRIDO dB", 0, 140, 65)

c1,c2,c3 = st.columns(3)
fecha = c1.date_input("Fecha")
hora = c2.text_input("Hora", "21:01")
cal_i = c3.text_input("Cal Ini", "114")
cal_f = st.text_input("Cal Final", "114")
mem = st.number_input("Memoria", 1, 99, 1)
laeq = st.number_input("LAeq", 0.0, 140.0, 65.0)

foto = st.file_uploader("📸 FOTO DEL PUNTO (Earth, croquis)", type=['jpg','jpeg','png'], help="Sube la captura de Google Earth como la de tu foto")

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
    # Guardar foto en temp
    foto_path = None
    if foto:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
            tmp.write(foto.getbuffer())
            foto_path = tmp.name
    st.session_state.puntos.append({
        "punto":punto,"coord":coord,"desc":desc,"barrido":barrido,
        "fecha":str(fecha),"hora":hora,"cal_i":cal_i,"cal_f":cal_f,
        "mem":mem,"laeq":laeq,"vmax":vmax,"vdir":vdir,"temp":temp,"hum":hum,
        "precip":precip,"fuente":fuente,"tipo":tipo,"top":top,"obs":obs,
        "foto":foto_path
    })
    st.toast("Punto con foto agregado!")

if st.session_state.puntos:
    st.dataframe(pd.DataFrame([{k:v for k,v in p.items() if k!='foto'} for p in st.session_state.puntos]), use_container_width=True)

    def build_con_foto():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Datos de Campo Emision"
        for k,w in {'A':3,'B':28,'C':15,'D':18,'E':22,'F':18,'G':15,'H':15,'I':12,'J':12,'K':22,'L':20,'M':14,'N':14,'O':40}.items():
            ws.column_dimensions[k].width = w

        ws['F2'] = "DATOS DE CAMPO EMISION DE RUIDO"; ws['F2'].font = Font(bold=True, size=12)
        ws['B5'] = "Codigo: R2-POE37-EP"; ws['F5'] = "Version: 04"; ws['K5'] = "Fecha: 2024-05-20"
        ws['B7'] = "CLIENTE: "; ws['E7'] = cliente
        ws['B8'] = "NOMBRE DEL PROYECTO:"; ws['E8'] = proyecto
        ws['B9'] = "DEPARTAMENTO : "; ws['E9'] = depto
        ws['B10'] = "MUNICIPIO:"; ws['E10'] = muni
        ws['B11'] = "PLAN DE MUESTREO:"; ws['B12'] = "EMISIÓN RUIDO"
        ws['N7'] = "DATOS DEL EQUIPO UTILIZADO"; ws['N7'].font = Font(bold=True)

        r = 13
        for p in st.session_state.puntos:
            ws[f'B{r}'] = "PUNTO DE MONITOREO:"; ws[f'B{r}'].font = Font(bold=True); ws[f'D{r}'] = p['punto']
            ws[f'J{r}'] = "COORDENADAS ORIGEN NACIONAL:"; ws[f'J{r}'].font = Font(bold=True); ws[f'L{r}'] = p['coord']; r+=1
            ws[f'B{r}'] = "DESCRIPCIÓN DEL PUNTO DE MONITOREO:"; ws[f'B{r}'].font = Font(bold=True)
            ws[f'F{r}'] = p['desc']; ws[f'F{r}'].alignment = Alignment(wrap_text=True); r+=1
            ws[f'B{r}'] = "REGISTRO DEL BARRIDO PERIMETRAL (dB):"; ws[f'B{r}'].font = Font(bold=True); ws[f'F{r}'] = p['barrido']; r+=1

            headers = ["Fecha de Toma","Hora de Toma","Calibración (dB)","Memoria / No. Estudio","LAeq,T (dB)\nIn situ","Velocidad del Viento Máx.\n(m/s)","Dirección del Viento","Temperatura\n(°C)","Humedad Relativa\n(%)","Se evidencias precipitaciones","Fuente de Ruido / Equipo","*Tipo de Ruido","Tiempo de Operación","Observaciones"]
            for i,h in enumerate(headers, start=2):
                c = ws.cell(row=r, column=i, value=h); c.font = Font(bold=True, size=8); c.alignment = Alignment(wrap_text=True, horizontal='center')
            ws.row_dimensions[r].height = 35; r+=1

            ws[f'B{r}'] = p['fecha']; ws[f'C{r}'] = p['hora']
            ws[f'D{r}'] = f"Inicial {p['cal_i']}\nFinal {p['cal_f']}"
            ws[f'E{r}'] = p['mem']; ws[f'F{r}'] = p['laeq']; ws[f'G{r}'] = p['vmax']; ws[f'H{r}'] = p['vdir']
            ws[f'I{r}'] = p['temp']; ws[f'J{r}'] = p['hum']; ws[f'K{r}'] = p['precip']; ws[f'L{r}'] = p['fuente']
            ws[f'M{r}'] = p['tipo']; ws[f'N{r}'] = p['top']; ws[f'O{r}'] = p['obs']; r+=1

            ws[f'B{r}'] = "DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"; ws[f'B{r}'].font = Font(bold=True); r+=1
            ws[f'B{r}'] = "(Identifique obstáculos)"; r+=1

            # --- AQUÍ VA LA FOTO ---
            foto_row = r
            ws.row_dimensions[r].height = 180
            ws.row_dimensions[r+1].height = 180
            if p['foto'] and os.path.exists(p['foto']):
                try:
                    # redimensionar
                    img = XLImage(p['foto'])
                    img.width = 500
                    img.height = 280
                    ws.add_image(img, f'C{r}') # La pone desde columna C
                except Exception as e:
                    ws[f'C{r}'] = f"Foto: {e}"
            else:
                ws[f'C{r}'] = "ESPACIO PARA FOTO EARTH / CROQUIS - Sube la foto al agregar el punto"
            r+=4

        bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

    st.download_button("📥 DESCARGAR EXCEL CON FOTO IGUAL AL ORIGINAL", build_con_foto(), f"ORIGINAL_CON_FOTO_{muni}.xlsx", type="primary", use_container_width=True)
