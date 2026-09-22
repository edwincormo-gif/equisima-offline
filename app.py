import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="EQUISIMA ORIGINAL", layout="wide")
st.title("🟡 EQUISIMA - Formato Original R2-POE37-EP")

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
    proyecto = st.text_input("NOMBRE DEL PROYECTO", "")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    muni = st.text_input("MUNICIPIO", "ATACO")
    resp = st.text_input("Responsable", "edwin cortes")

punto = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
coord = st.text_input("COORDENADAS ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W")
desc = st.text_area("DESCRIPCIÓN DEL PUNTO", "En la entrada o vía principal de acceso por el lado norte.")
barrido = st.number_input("REGISTRO BARRIDO PERIMETRAL dB", 0, 140, 65)

c1,c2,c3 = st.columns(3)
fecha = c1.date_input("Fecha Toma")
hora = c2.text_input("Hora Toma", "21:01")
cal_i = c3.text_input("Cal Inicial", "114")
cal_f = st.text_input("Cal Final", "114")
mem = st.number_input("Memoria / No. Estudio", 1, 99, 1)
laeq = st.number_input("LAeq,T In situ", 0.0, 140.0, 65.0)

fcsv = st.file_uploader("CSV TOTAL (opcional)", type='csv')
if fcsv: laeq = calc_csv(fcsv); st.success(f"LAeq: {laeq}")

vmax = st.text_input("Viento Max m/s", "0.3")
vdir = st.text_input("Dirección Viento", "SW")
temp = st.text_input("Temperatura °C", "29")
hum = st.text_input("Humedad %", "45")
precip = st.selectbox("Precipitación", ["NO","SI"])
fuente = st.text_input("Fuente Ruido", "mineria")
tipo = st.text_input("Tipo Ruido", "EMISION")
top = st.text_input("Tiempo Operación", "1")
obs = st.text_area("Observaciones / camino", "camino destapado por donde entran y salen los camiones...")

if st.button("➕ AGREGAR PUNTO", type="primary", use_container_width=True):
    st.session_state.puntos.append([punto,coord,desc,barrido,str(fecha),hora,cal_i,cal_f,mem,laeq,vmax,vdir,temp,hum,precip,fuente,tipo,top,obs])

if st.session_state.puntos:
    st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)

    def build_original():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Datos de Campo Emision"

        # Anchos EXACTOS como tu original
        ws.column_dimensions['A'].width = 2
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 22
        ws.column_dimensions['F'].width = 18
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 15
        ws.column_dimensions['I'].width = 12
        ws.column_dimensions['J'].width = 12
        ws.column_dimensions['K'].width = 22
        ws.column_dimensions['L'].width = 20
        ws.column_dimensions['M'].width = 14
        ws.column_dimensions['N'].width = 14
        ws.column_dimensions['O'].width = 40

        # HEADER EXACTO
        ws['F2'] = "DATOS DE CAMPO EMISION DE RUIDO"; ws['F2'].font = Font(bold=True, size=12)
        ws['B5'] = "Codigo: R2-POE37-EP"; ws['F5'] = "Version: 04"; ws['K5'] = "Fecha: 2024-05-20"
        ws['B7'] = "CLIENTE: "; ws['E7'] = cliente
        ws['B8'] = "NOMBRE DEL PROYECTO:"; ws['E8'] = proyecto
        ws['B9'] = "DEPARTAMENTO : "; ws['E9'] = depto
        ws['B10'] = "MUNICIPIO:"; ws['E10'] = muni
        ws['B11'] = "PLAN DE MUESTREO:"
        ws['B12'] = "EMISIÓN RUIDO"

        ws['N7'] = "DATOS DEL EQUIPO UTILIZADO"; ws['N7'].font = Font(bold=True)
        ws['N8'] = "EQUIPO"; ws['N8'].font = Font(bold=True)
        ws['N9'] = "SONÓMETRO"; ws['N10'] = "PISTÓFONO"; ws['N11'] = "ESTACIÓN METEOROLÓGICA"

        r = 13
        for p in st.session_state.puntos:
            # BLOQUE EXACTO COMO TU ORIGINAL
            ws[f'B{r}'] = "PUNTO DE MONITOREO:"; ws[f'B{r}'].font = Font(bold=True)
            ws[f'D{r}'] = p[0]
            ws[f'J{r}'] = "COORDENADAS ORIGEN NACIONAL:"; ws[f'J{r}'].font = Font(bold=True)
            ws[f'L{r}'] = p[1]
            ws[f'O{r}'] = "ESCENARIO DE MEDICIÓN:"; r+=1

            ws[f'B{r}'] = "DESCRIPCIÓN DEL PUNTO DE MONITOREO:"; ws[f'B{r}'].font = Font(bold=True)
            ws[f'F{r}'] = p[2]; ws[f'F{r}'].alignment = Alignment(wrap_text=True)
            ws[f'O{r}'] = "¿Se realiza medición del Ruido Residual?"; r+=1

            ws[f'B{r}'] = "REGISTRO DEL BARRIDO PERIMETRAL (dB):"; ws[f'B{r}'].font = Font(bold=True)
            ws[f'F{r}'] = p[3]
            ws[f'O{r}'] = "¿Se realiza el barrido perimetral al límite perimetral o fachada?"; r+=1

            # CABECERA DE TABLA EXACTA
            ws[f'B{r}'] = "Fecha de Toma"; ws[f'C{r}'] = "Hora de Toma"; ws[f'D{r}'] = "Calibración (dB)"
            ws[f'E{r}'] = "Memoria / No. Estudio"; ws[f'F{r}'] = "LAeq,T (dB)\nIn situ"
            ws[f'G{r}'] = "Velocidad del Viento Máx. \n(m/s)"; ws[f'H{r}'] = "Dirección del Viento "
            ws[f'I{r}'] = "Temperatura\n(°C)"; ws[f'J{r}'] = "Humedad Relativa \n(%)"
            ws[f'K{r}'] = "Se evidencias precipacitones durante la medición (Si/No)"
            ws[f'L{r}'] = "Fuente de Ruido / Equipo"; ws[f'M{r}'] = "*Tipo de Ruido"; ws[f'N{r}'] = "Tiempo de Operación"
            ws[f'O{r}'] = p[18] # tu observación larga va aquí como en original
            for c in range(2,16):
                ws.cell(row=r, column=c).font = Font(bold=True, size=8)
                ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, horizontal='center')
            r+=1

            # DATOS
            ws[f'B{r}'] = p[4]; ws[f'C{r}'] = p[5]
            ws[f'D{r}'] = f"Inicial {p[6]}\nFinal {p[7]}"; ws[f'D{r}'].alignment = Alignment(wrap_text=True, horizontal='center')
            ws[f'E{r}'] = p[8]; ws[f'F{r}'] = p[9]; ws[f'G{r}'] = p[10]; ws[f'H{r}'] = p[11]
            ws[f'I{r}'] = p[12]; ws[f'J{r}'] = p[13]; ws[f'K{r}'] = p[14]; ws[f'L{r}'] = p[15]
            ws[f'M{r}'] = p[16]; ws[f'N{r}'] = p[17]
            for c in range(2,15):
                ws.cell(row=r, column=c).alignment = Alignment(horizontal='center')
            r+=2

            ws[f'B{r}'] = "*Tipo de Ruido: Continuo e Intermitente"; r+=1
            ws[f'B{r}'] = "DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO \n(Identifique además los obstáculos entre fuente de medición y el punto de monitoreo definido)"
            ws[f'B{r}'].font = Font(bold=True); r+=4

        ws[f'F{r}'] = "Responsable de la Medición:"; ws[f'K{r}'] = resp; r+=2
        ws[f'C{r}'] = "Nombre y firma "; ws[f'H{r}'] = "Nombre y firma "; ws[f'L{r}'] = "Nombre y firma "; r+=1
        ws[f'C{r}'] = "Elaboró: "; ws[f'H{r}'] = "Revisó: "; ws[f'L{r}'] = "Aprobó: Gerente G."; r+=1
        ws[f'C{r}'] = "Fecha: "; ws[f'H{r}'] = "Fecha: "; ws[f'L{r}'] = "Fecha: "; r+=1
        ws[f'B{r}'] = "DOCUMENTO CONTROLADO"

        bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

    st.download_button("📥 DESCARGAR EXCEL ORIGINAL", build_original(), f"{muni}_R2-POE37-EP_{datetime.now().strftime('%d%m%Y')}.xlsx", type="primary", use_container_width=True)
