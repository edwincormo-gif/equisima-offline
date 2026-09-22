import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="EQUISIMA", page_icon="🟡", layout="wide")
st.title("🟡 EQUISIMA - Formato R2-POE37-EP")

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
    proyecto = st.text_input("PROYECTO", "rd-eq15-2026")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    muni = st.text_input("MUNICIPIO", "ATACO")
    resp = st.text_input("Responsable", "edwin cortes")

st.subheader("Agregar Punto de Campo")
punto = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
coord = st.text_input("COORDENADAS ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W")
desc = st.text_area("DESCRIPCIÓN DEL PUNTO", "En la entrada o vía principal de acceso por el lado norte.")
barrido = st.number_input("BARRIDO PERIMETRAL dB", 0, 140, 65)

c1,c2,c3 = st.columns(3)
fecha = c1.date_input("Fecha")
hora = c2.text_input("Hora", "21:01")
cal_i = c3.text_input("Cal Inicial", "114")

c4,c5,c6 = st.columns(3)
cal_f = c4.text_input("Cal Final", "114")
mem = c5.number_input("Memoria", 1, 99, 1)
laeq = c6.number_input("LAeq", 0.0, 140.0, 65.0)

fcsv = st.file_uploader("CSV TOTAL para calcular LAeq auto", type='csv')
if fcsv:
    laeq = calc_csv(fcsv)
    st.success(f"Calculado: {laeq} dB")

c7,c8,c9,c10 = st.columns(4)
vmax = c7.text_input("Viento Max", "0.3")
vdir = c8.text_input("Dir", "SW")
temp = c9.text_input("Temp °C", "29")
hum = c10.text_input("Hum %", "45")

c11,c12,c13,c14 = st.columns(4)
precip = c11.selectbox("Precipitación", ["NO","SI"])
fuente = c12.text_input("Fuente", "mineria")
tipo = c13.selectbox("Tipo", ["EMISION","Continuo"])
top = c14.text_input("Tiempo Op", "1")
obs = st.text_area("Observaciones / camino destapado...", "camino destapado por donde entran y salen los camiones...")

if st.button("➕ AGREGAR PUNTO", type="primary", use_container_width=True):
    st.session_state.puntos.append({
        "punto":punto,"coord":coord,"desc":desc,"barrido":barrido,
        "fecha":str(fecha),"hora":hora,"cal_i":cal_i,"cal_f":cal_f,
        "mem":mem,"laeq":laeq,"vmax":vmax,"vdir":vdir,"temp":temp,
        "hum":hum,"precip":precip,"fuente":fuente,"tipo":tipo,"top":top,"obs":obs
    })
    st.toast("Punto agregado!")

if st.session_state.puntos:
    st.divider()
    st.write(f"Llevas **{len(st.session_state.puntos)} puntos**")
    st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
    if st.button("Borrar último"): st.session_state.puntos.pop(); st.rerun()

    def build_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Datos de Campo Emision"

        # Anchos iguales a tu original
        widths = {'A':3,'B':28,'C':18,'D':18,'E':12,'F':18,'G':14,'H':14,'I':12,'J':12,'K':18,'L':20,'M':14,'N':12,'O':35}
        for k,v in widths.items(): ws.column_dimensions[k].width = v

        # Encabezado
        ws['B1'] = "DATOS DE CAMPO EMISION DE RUIDO"; ws['B1'].font = Font(bold=True, size=14)
        ws.merge_cells('B1:O1'); ws['B1'].alignment = Alignment(horizontal='center')

        ws['B3'] = "Codigo: R2-POE37-EP"; ws['F3'] = "Version: 04"; ws['K3'] = "Fecha: 2024-05-20"

        ws['B5'] = "CLIENTE:"; ws['B5'].font = Font(bold=True); ws['C5'] = cliente
        ws['B6'] = "NOMBRE DEL PROYECTO:"; ws['B6'].font = Font(bold=True); ws['C6'] = proyecto
        ws['B7'] = "DEPARTAMENTO :"; ws['B7'].font = Font(bold=True); ws['C7'] = depto
        ws['B8'] = "MUNICIPIO:"; ws['B8'].font = Font(bold=True); ws['C8'] = muni

        ws['K5'] = "DATOS DEL EQUIPO UTILIZADO"; ws['K5'].font = Font(bold=True)
        ws['K6'] = "EQUIPO"; ws['K6'].font = Font(bold=True); ws['M6'] = "MARCA"; ws['M6'].font = Font(bold=True); ws['O6'] = "SERIE"; ws['O6'].font = Font(bold=True)
        ws['K7'] = "SONÓMETRO"; ws['M7'] = "HD2010UC/A"; ws['O7'] = "15031643825"
        ws['K8'] = "PISTÓFONO"; ws['M8'] = "HD2020"

        r = 11
        for p in st.session_state.puntos:
            ws[f'B{r}'] = "PUNTO DE MONITOREO:"; ws[f'B{r}'].font = Font(bold=True); ws[f'C{r}'] = p['punto']
            ws[f'K{r}'] = "COORDENADAS ORIGEN NACIONAL:"; ws[f'K{r}'].font = Font(bold=True); ws[f'M{r}'] = p['coord']; r+=1

            ws[f'B{r}'] = "DESCRIPCIÓN DEL PUNTO DE MONITOREO:"; ws[f'B{r}'].font = Font(bold=True)
            ws[f'C{r}'] = p['desc']; ws.merge_cells(f'C{r}:O{r}'); ws[f'C{r}'].alignment = Alignment(wrap_text=True); r+=1

            ws[f'B{r}'] = "REGISTRO DEL BARRIDO PERIMETRAL (dB):"; ws[f'B{r}'].font = Font(bold=True); ws[f'C{r}'] = p['barrido']; r+=2

            # Tabla de medición
            headers = ["Fecha de Toma","Hora de Toma","Calibración (dB)","Memoria / No. Estudio","LAeq,T (dB)\nIn situ","Velocidad del Viento Máx.\n(m/s)","Dirección del Viento","Temperatura\n(°C)","Humedad Relativa\n(%)","Se evidencias precipitaciones","Fuente de Ruido / Equipo","*Tipo de Ruido","Tiempo de Operación","Obs"]
            for i,h in enumerate(headers, start=2):
                cell = ws.cell(row=r, column=i, value=h)
                cell.font = Font(bold=True, size=8); cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            ws.row_dimensions[r].height = 40; r+=1

            ws.cell(row=r, column=2, value=p['fecha']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=3, value=p['hora']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=4, value=f"Inicial {p['cal_i']}\nFinal {p['cal_f']}").alignment = Alignment(wrap_text=True, horizontal='center')
            ws.cell(row=r, column=5, value=p['mem']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=6, value=p['laeq']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=7, value=p['vmax']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=8, value=p['vdir']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=9, value=p['temp']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=10, value=p['hum']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=11, value=p['precip']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=12, value=p['fuente']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=13, value=p['tipo']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=14, value=p['top']).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=15, value=p['obs']).alignment = Alignment(wrap_text=True); r+=2

            ws[f'B{r}'] = "DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"; ws[f'B{r}'].font = Font(bold=True); r+=1
            ws[f'B{r}'] = "(Identifique obstáculos entre fuente y punto)"; r+=3

        ws[f'B{r+2}'] = "Responsable de la Medición:"; ws[f'C{r+2}'] = resp; ws[f'C{r+2}'].font = Font(bold=True)
        ws[f'B{r+4}'] = "Elaboró:"; ws[f'H{r+4}'] = "Revisó:"; ws[f'L{r+4}'] = "Aprobó: Gerente G."
        ws[f'B{r+6}'] = "DOCUMENTO CONTROLADO"

        bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

    st.download_button("📥 DESCARGAR EXCEL IGUAL AL ORIGINAL", build_excel(), f"R2-POE37-EP_{muni}_{datetime.now().strftime('%d%m%Y')}.xlsx", type="primary", use_container_width=True)
