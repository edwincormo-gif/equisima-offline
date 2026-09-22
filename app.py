import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="EQUISIMA CAMPO", page_icon="🟡", layout="wide")
st.title("🟡 EQUISIMA - Campo R2-POE37-EP V3")

def calc_csv(f):
    try:
        df = pd.read_csv(f, sep=';', encoding='latin-1', engine='python')
        col = df.columns[0]
        for c in df.columns:
            if 'LAeq' in str(c): col=c; break
        v = pd.to_numeric(df[col].astype(str).str.replace(',','.'), errors='coerce').dropna()
        v = v[(v>20)&(v<140)]
        return round(10*np.log10((10**(v/10)).mean()),1)
    except: return 0

if 'puntos' not in st.session_state: st.session_state.puntos=[]

with st.sidebar:
    cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
    proyecto = st.text_input("PROYECTO", "rd-eq15-2026")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    municipio = st.text_input("MUNICIPIO", "ATACO")
    responsable = st.text_input("Responsable", "edwin cortes")

punto = st.text_input("PUNTO", "Punto 1 nocturno")
coord = st.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W")
desc = st.text_area("DESCRIPCIÓN", "En la entrada o vía principal...")
obs = st.text_area("Observaciones", "camino destapado...")
barrido = st.number_input("BARRIDO dB", 0, 140, 65)
fecha = st.date_input("Fecha")
hora = st.text_input("Hora", "21:01")
cal_i = st.text_input("Cal Inicial", "114")
cal_f = st.text_input("Cal Final", "114")
mem = st.number_input("Memoria", 1, 99, 1)
laeq = st.number_input("LAeq", 0.0, 140.0, 65.0)
csv = st.file_uploader("CSV para calcular", type='csv')
if csv: laeq = calc_csv(csv)

vmax = st.text_input("Viento Max", "0.3")
vdir = st.text_input("Dir Viento", "SW")
temp = st.text_input("Temp", "29")
hum = st.text_input("Humedad", "45")
precip = st.selectbox("Precip?", ["NO","SI"])
fuente = st.text_input("Fuente", "mineria")
tipo = st.selectbox("Tipo", ["EMISION","Continuo"])
top = st.text_input("Tiempo", "1")

if st.button("AGREGAR PUNTO", type="primary", use_container_width=True):
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs])

if st.session_state.puntos:
    st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
    def build():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title="Datos de Campo Emision"
        for k,w in {'A':2,'B':22,'C':14,'D':16,'E':22,'F':16,'G':12,'H':14,'I':12,'J':12,'K':14,'L':16,'M':14,'N':14,'O':30}.items():
            ws.column_dimensions[k].width=w
        ws.merge_cells('F1:J1'); ws['F1']="DATOS DE CAMPO EMISION DE RUIDO"; ws['F1'].font=Font(bold=True,size=12)
        ws['B3']="Codigo: R2-POE37-EP"; ws['E3']="Version: 04"; ws['J3']="Fecha: 2024-05-20"
        ws['B5']="CLIENTE:"; ws['D5']=cliente
        ws['B6']="NOMBRE DEL PROYECTO:"; ws['D6']=proyecto
        ws['B7']="DEPARTAMENTO :"; ws['D7']=depto
        ws['B8']="MUNICIPIO:"; ws['D8']=municipio
        ws['J5']="DATOS DEL EQUIPO UTILIZADO"; ws['J5'].font=Font(bold=True)
        ws['J6']="EQUIPO"; ws['L6']="MARCA"; ws['N6']="SERIE"
        ws['J7']="SONÓMETRO"; ws['L7']="HD2010UC/A"
        ws['J8']="PISTÓFONO"; ws['L8']="HD2020"
        r=11
        for p in st.session_state.puntos:
            ws[f'B{r}']="PUNTO DE MONITOREO:"; ws[f'D{r}']=p[0]; ws[f'J{r}']="COORDENADAS:"; ws[f'L{r}']=p[1]; r+=1
            ws[f'B{r}']="DESCRIPCIÓN:"; ws.merge_cells(f'D{r}:O{r}'); ws[f'D{r}']=p[2]; ws[f'D{r}'].alignment=Alignment(wrap_text=True); r+=1
            ws[f'B{r}']="BARRIDO (dB):"; ws[f'D{r}']=p[3]; r+=2
            hdr=["Fecha","Hora","Cal","Mem","LAeq","Vmax","Dir","Temp","Hum","Precip","Fuente","Tipo","Top","Obs"]
            for i,h in enumerate(hdr, start=2): ws.cell(row=r, column=i, value=h).font=Font(bold=True,size=7)
            r+=1
            vals=[p[4],p[5],f"Ini {p[6]} Fin {p[7]}",p[8],p[9],p[10],p[11],p[12],p[13],p[14],p[15],p[16],p[17],p[18]]
            for i,v in enumerate(vals, start=2): ws.cell(row=r, column=i, value=v).alignment=Alignment(wrap_text=True, horizontal='center')
            r+=3
        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio
    st.download_button("DESCARGAR EXCEL", build(), "CAMPO.xlsx", use_container_width=True, type="primary")
