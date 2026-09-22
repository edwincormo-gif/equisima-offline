import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="EQUISIMA CAMPO", page_icon="🟡", layout="wide")
st.title("🟡 EQUISIMA - Campo R2-POE37-EP V3")
st.caption("Genera Excel idéntico a tu plantilla")

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
    st.header("Datos Generales")
    cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
    proyecto = st.text_input("NOMBRE DEL PROYECTO", "rd-eq15-2026")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    municipio = st.text_input("MUNICIPIO", "ATACO")
    responsable = st.text_input("Responsable", "edwin cortes")

st.subheader("Nuevo Punto")
c1,c2 = st.columns(2)
punto = c1.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
coord = c2.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W")
desc = st.text_area("DESCRIPCIÓN DEL PUNTO", "En la entrada o vía principal de acceso por el lado norte.", height=60)
obs = st.text_area("Observaciones / Camino", "camino destapado por donde entran y salen los camiones...", height=60)
c3,c4,c5 = st.columns(3)
barrido = c3.number_input("BARRIDO PERIMETRAL dB", 0, 140, 65)
fecha = c4.date_input("Fecha")
hora = c5.text_input("Hora", "21:01")
c6,c7,c8,c9 = st.columns(4)
cal_i = c6.text_input("Cal Inicial", "114")
cal_f = c7.text_input("Cal Final", "114")
mem = c8.number_input("Memoria", 1, 99, 1)
laeq = c9.number_input("LAeq In situ", 0.0, 140.0, 65.0)
csv = st.file_uploader("Si tienes CSV, súbelo y calcula LAeq auto", type='csv')
if csv: laeq = calc_csv(csv); st.success(f"LAeq: {laeq} dB")
c10,c11,c12,c13 = st.columns(4)
vmax = c10.text_input("Viento Max", "0.3")
vdir = c11.text_input("Dir Viento", "SW")
temp = c12.text_input("Temp", "29")
hum = c13.text_input("Humedad", "45")
c14,c15,c16,c17 = st.columns(4)
precip = c14.selectbox("Precip?", ["NO","SI"])
fuente = c15.text_input("Fuente", "mineria")
tipo = c16.selectbox("Tipo", ["EMISION","Continuo","Intermitente"])
top = c17.text_input("Tiempo Op", "1")

if st.button("➕ AGREGAR PUNTO", type="primary", use_container_width=True):
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs])
    st.toast("Punto agregado")

if st.session_state.puntos:
    st.dataframe(pd.DataFrame(st.session_state.puntos, columns=["Punto","Coord","Desc","Barrido","Fecha","Hora","CalI","CalF","Mem","LAeq","Vmax","Vdir","Temp","Hum","Prec","Fuente","Tipo","Top","Obs"]), use_container_width=True)

    def build():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title="Datos de Campo Emision"
        # Anchos como original
        ws.column_dimensions['A'].width=2
        ws.column_dimensions['B'].width=22
        ws.column_dimensions['C'].width=14
        ws.column_dimensions['D'].width=16
        ws.column_dimensions['E'].width=22
        ws.column_dimensions['F'].width=16
        ws.column_dimensions['G'].width=12
        ws.column_dimensions['H'].width=14
        ws.column_dimensions['I'].width=12
        ws.column_dimensions['J'].width=12
        ws.column_dimensions['K'].width=14
        ws.column_dimensions['L'].width=16
        ws.column_dimensions['M'].width=14
        ws.column_dimensions['N'].width=14
        ws.column_dimensions['O'].width=30

        ws.merge_cells('F1:J1'); ws['F1']="DATOS DE CAMPO EMISION DE RUIDO"; ws['F1'].font=Font(bold=True,size=12); ws['F1'].alignment=Alignment(horizontal='center')
        ws['B3']="Codigo: R2-POE37-EP"; ws['E3']="Version: 04"; ws['J3']="Fecha: 2024-05-20"
        ws['B5']="CLIENTE:"; ws['D5']=cliente
        ws['B6']="NOMBRE DEL PROYECTO:"; ws['D6']=proyecto
        ws['B7']="DEPARTAMENTO :"; ws['D7']=depto
        ws['B8']="MUNICIPIO:"; ws['D8']=municipio
        ws['J5']="DATOS DEL EQUIPO UTILIZADO"; ws['J5'].font=Font(bold=True)
        ws['J6']="EQUIPO"; ws['L6']="MARCA"; ws['N6']="SERIE"
        ws['J7']="SONÓMETRO"; ws['L7']="HD2010UC/A"; ws['N7']="15031643825"
        ws['J8']="PISTÓFONO"; ws['L8']="HD2020"

        r=11
        for p in st.session_state.puntos:
            ws[f'B{r}']="PUNTO DE MONITOREO:"; ws[f'B{r}'].font=Font(bold=True); ws[f'D{r}']=p[0]
            ws[f'J{r}']="COORDENADAS ORIGEN NACIONAL:"; ws[f'J{r}'].font=Font(bold=True); ws[f'L{r}']=p[1]; r+=1
            ws[f'B{r}']="DESCRIPCIÓN DEL PUNTO DE MONITOREO:"; ws[f'B{r}'].font=Font(bold=True); ws.merge_cells(f'D{r}:O{r}'); ws[f'D{r}']=p[2]; ws[f'D{r}'].alignment=Alignment(wrap_text=True); r+=1
            ws[f'B{r}']="REGISTRO DEL BARRIDO PERIMETRAL (dB):"; ws[f'B{r}'].font=Font(bold=True); ws[f'D{r}']=p[3]; r+=2

            hdr=["Fecha de Toma","Hora de Toma","Calibración (dB)","Memoria / No. Estudio","LAeq,T (dB)\nIn situ","Velocidad del Viento Máx.\n(m/s)","Dirección del Viento","Temperatura\n(°C)","Humedad Relativa\n(%)","Se evidencian precipitaciones","Fuente de Ruido","*Tipo de Ruido","Tiempo de Operación","Observaciones"]
            for i,h in enumerate(hdr, start=2):
                c=ws.cell(row=r, column=i, value=h); c.font=Font(bold=True,size=7); c.alignment=Alignment(wrap_text=True, horizontal='center', vertical='center')
            ws.row_dimensions[r].height=35; r+=1
            vals=[p[4],p[5],f"Inicial {p[6]}",p[8],p[9],p[10],p[11],p[12],p[13],p[14],p[15],p[16],p[17],p[18]]
            for i,v in enumerate(vals, start=2):
                c=ws.cell(row=r, column=i, value=v); c.alignment=Alignment(wrap_text=True, horizontal='center')
            r+=1
            ws.cell(row=r, column=4, value=f"Final {p[7]}"); r+=2
            ws[f'B{r}']="DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"; ws[f'B{r}'].font=Font(bold=True); r+=1
            ws[f'B{r}']=" (Identifique obstáculos)"; r+=3
            ws[f'B{r}']=f"Foto: {p[0]} - ver archivo adjunto si tomó foto"; r+=3

        ws[f'F{r}']="Responsable:"; ws[f'K{r}']=responsable
        bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

    st.download_button("📥 DESCARGAR EXCEL ORDENADO", build(), f"CAMPO_{municipio}_{datetime.now().strftime('%Y%m%d')}.xlsx", use_container_width=True, type="primary")

2. **Commit changes**

3. Recarga tu app y descarga de nuevo. Ahora sí te queda con columnas anchas, sin que se monten las letras.

Mándame captura del nuevo Excel.
