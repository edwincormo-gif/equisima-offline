import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from copy import copy

st.set_page_config(page_title="EQUISIMA", page_icon="🟡", layout="wide")
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
    proyecto = st.text_input("PROYECTO", "rd-eq15-2026")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    muni = st.text_input("MUNICIPIO", "ATACO")
    resp = st.text_input("Responsable", "edwin cortes")

st.subheader("Nuevo Punto")
punto = st.text_input("Punto", "Punto 1 nocturno")
coord = st.text_input("Coordenadas", "3°35'11.30\"N 75°23'27.72\"W")
desc = st.text_area("Descripción punto", "En la entrada o vía principal...")
barrido = st.number_input("Barrido dB", 0, 140, 65)
col1,col2,col3 = st.columns(3)
fecha = col1.date_input("Fecha")
hora = col2.text_input("Hora", "21:01")
cal_i = col3.text_input("Cal Ini", "114")
cal_f = st.text_input("Cal Final", "114")
mem = st.number_input("Memoria", 1, 99, 1)
laeq = st.number_input("LAeq In situ", 0.0, 140.0, 65.0)
fcsv = st.file_uploader("CSV para auto-calcular LAeq", type='csv')
if fcsv: laeq = calc_csv(fcsv); st.info(f"LAeq calc: {laeq}")

vmax = st.text_input("Vmax m/s", "0.3")
vdir = st.text_input("Dir", "SW")
temp = st.text_input("Temp", "29")
hum = st.text_input("Hum", "45")
precip = st.selectbox("Precip", ["NO","SI"])
fuente = st.text_input("Fuente", "mineria")
tipo = st.selectbox("Tipo", ["EMISION","Continuo","Intermitente"])
top = st.text_input("Tiempo Op", "1")
obs = st.text_area("Observaciones", "camino destapado...")

if st.button("➕ AGREGAR PUNTO", type="primary", use_container_width=True):
    st.session_state.puntos.append({
        "punto":punto,"coord":coord,"desc":desc,"barrido":barrido,"fecha":fecha,"hora":hora,
        "cal_i":cal_i,"cal_f":cal_f,"mem":mem,"laeq":laeq,"vmax":vmax,"vdir":vdir,
        "temp":temp,"hum":hum,"precip":precip,"fuente":fuente,"tipo":tipo,"top":top,"obs":obs
    })

if st.session_state.puntos:
    st.table(pd.DataFrame(st.session_state.puntos))

    def generar():
        try:
            wb = openpyxl.load_workbook("plantilla.xlsx")
            ws = wb["Datos de Campo Emision"]
        except:
            st.error("Sube primero plantilla.xlsx al repo")
            return None

        # Datos cabecera fijos
        ws['E7'] = cliente
        ws['E8'] = proyecto
        ws['E9'] = depto
        ws['E10'] = muni

        # Limpiar datos viejos de plantilla desde fila 12 hacia abajo y reconstruir
        # Borramos desde fila 12 hasta 100
        for r in range(12, 200):
            for c in range(1, 16):
                if ws.cell(r,c).value and "PUNTO DE MONITOREO" in str(ws.cell(r,c).value):
                    # dejar para sobrescribir
                    pass

        # Vamos a escribir desde fila 12 en adelante usando el formato de tu plantilla
        r = 12
        for p in st.session_state.puntos:
            ws[f'B{r}'] = "PUNTO DE MONITOREO:"; ws[f'D{r}'] = p["punto"]
            ws[f'J{r}'] = "COORDENADAS ORIGEN NACIONAL:"; ws[f'L{r}'] = p["coord"]; r+=1
            ws[f'B{r}'] = "DESCRIPCIÓN DEL PUNTO DE MONITOREO:"; ws[f'F{r}'] = p["desc"]; r+=1
            ws[f'B{r}'] = "REGISTRO DEL BARRIDO PERIMETRAL (dB):"; ws[f'F{r}'] = p["barrido"]; r+=1
            ws[f'B{r}'] = "Fecha de Toma"; ws[f'C{r}'] = "Hora de Toma"; ws[f'D{r}'] = "Calibración (dB)"
            ws[f'E{r}'] = "Memoria / No. Estudio"; ws[f'F{r}'] = "LAeq,T (dB)\nIn situ"
            ws[f'G{r}'] = "Velocidad del Viento Máx.\n(m/s)"; ws[f'H{r}'] = "Dirección del Viento"
            ws[f'I{r}'] = "Temperatura\n(°C)"; ws[f'J{r}'] = "Humedad Relativa\n(%)"
            ws[f'K{r}'] = "Se evidencias precipacitones durante la medición (Si/No)"
            ws[f'L{r}'] = "Fuente de Ruido / Equipo"; ws[f'M{r}'] = "*Tipo de Ruido"
            ws[f'N{r}'] = "Tiempo de Operación"; ws[f'O{r}'] = "Observaciones extra"; r+=1

            ws[f'B{r}'] = p["fecha"]; ws[f'C{r}'] = p["hora"]
            ws[f'D{r}'] = f"Inicial {p['cal_i']}\nFinal {p['cal_f']}"
            ws[f'E{r}'] = p["mem"]; ws[f'F{r}'] = p["laeq"]; ws[f'G{r}'] = p["vmax"]
            ws[f'H{r}'] = p["vdir"]; ws[f'I{r}'] = p["temp"]; ws[f'J{r}'] = p["hum"]
            ws[f'K{r}'] = p["precip"]; ws[f'L{r}'] = p["fuente"]; ws[f'M{r}'] = p["tipo"]
            ws[f'N{r}'] = p["top"]; ws[f'O{r}'] = p["obs"]; r+=2
            ws[f'B{r}'] = "DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"; r+=4

        ws[f'F{r}'] = "Responsable de la Medición:"; ws[f'K{r}'] = resp
        bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

    bio = generar()
    if bio:
        st.download_button("📥 DESCARGAR EXCEL IGUAL AL ORIGINAL", bio, f"CAMPO_{muni}_{datetime.now().strftime('%d%m%Y')}.xlsx", type="primary", use_container_width=True)
