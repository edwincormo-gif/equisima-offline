import streamlit as st, pandas as pd, numpy as np
from datetime import datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="EQUISIMA CAMPO", page_icon="🟡", layout="wide")
st.title("🟡 EQUISIMA - Datos de Campo R2-POE37-EP")
st.caption("Offline | 3224523451")

def calc_csv(f):
    try:
        df = pd.read_csv(f, sep=';', encoding='latin-1', engine='python')
        col = df.columns[0]
        for c in df.columns:
            if 'LAeq' in str(c) or 'Leq' in str(c): col=c; break
        v = pd.to_numeric(df[col].astype(str).str.replace(',','.'), errors='coerce').dropna()
        v = v[(v>20)&(v<140)]
        return round(10*np.log10((10**(v/10)).mean()),1)
    except: return 0

if 'puntos' not in st.session_state: st.session_state.puntos=[]

with st.sidebar:
    st.header("Datos Generales")
    cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    municipio = st.text_input("MUNICIPIO", "ATACO")
    proyecto = st.text_input("PROYECTO", "")
    responsable = st.text_input("Responsable", "edwin cortes")

st.subheader("Nuevo Punto")
c1,c2,c3 = st.columns(3)
punto = c1.text_input("Punto", "Punto 1 nocturno")
coord = c2.text_input("Coordenadas", "3°35'11.30\"N 75°23'27.72\"W")
barrido = c3.number_input("Barrido dB", 0, 140, 65)

desc = st.text_area("Descripción del punto", "En la entrada o vía principal...")
obs = st.text_area("Observaciones / Fuente", "camino destapado...")

c4,c5,c6 = st.columns(3)
fecha = c4.date_input("Fecha toma")
hora = c5.text_input("Hora", "21:01")
cal_ini = c6.text_input("Calibración Inicial", "114")
cal_fin = st.text_input("Calibración Final", "114")
memoria = st.number_input("Memoria / Estudio", 1, 100, 1)

c7,c8,c9 = st.columns(3)
laeq = c7.number_input("LAeq,T In situ (o calcular con CSV)", 0.0, 140.0, 65.0)
vmax = c8.text_input("Viento Max m/s", "0.3")
vdir = c9.text_input("Dirección", "SW")

c10,c11,c12 = st.columns(3)
temp = c10.text_input("Temp °C", "29")
hum = c11.text_input("Humedad %", "45")
precip = c12.selectbox("Precipitación", ["NO","SI"])
fuente = st.text_input("Fuente Ruido", "mineria")
tipo = st.selectbox("Tipo Ruido", ["EMISION","Continuo","Intermitente"])
topera = st.text_input("Tiempo operación", "1")

st.divider()
a1,a2 = st.columns(2)
ft = a1.file_uploader("CSV TOTAL (opcional) calcula LAeq", type='csv')
fr = a2.file_uploader("CSV RESIDUAL (opcional)", type='csv')
if ft: laeq = calc_csv(ft); st.info(f"LAeq calculado: {laeq} dB")

foto = st.camera_input("Foto Croquis / Esquema")

if st.button("➕ AGREGAR PUNTO A LA LISTA", type="primary", use_container_width=True):
    st.session_state.puntos.append({
        "punto":punto,"coord":coord,"barrido":barrido,"desc":desc,"obs":obs,
        "fecha":str(fecha),"hora":hora,"cal_ini":cal_ini,"cal_fin":cal_fin,
        "memoria":memoria,"laeq":laeq,"vmax":vmax,"vdir":vdir,"temp":temp,
        "hum":hum,"precip":precip,"fuente":fuente,"tipo":tipo,"topera":topera
    })
    st.success(f"Punto {punto} guardado. Llevas {len(st.session_state.puntos)} puntos")

if st.session_state.puntos:
    st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
    if st.button("🗑️ Borrar último"): st.session_state.puntos.pop()

    # GENERAR EXCEL IGUAL A TU PLANTILLA
    def generar_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Datos de Campo Emision"
        # Encabezado
        ws['F1']="DATOS DE CAMPO EMISION DE RUIDO"; ws['F1'].font=Font(bold=True)
        ws['B5']="Codigo: R2-POE37-EP"; ws['F5']="Version: 04"; ws['K5']="Fecha: 2024-05-20"
        ws['B7']="CLIENTE:"; ws['E7']=cliente
        ws['B9']="DEPARTAMENTO :"; ws['E9']=depto
        ws['B10']="MUNICIPIO:"; ws['E10']=municipio
        ws['B8']="NOMBRE DEL PROYECTO:"; ws['E8']=proyecto
        ws['N7']="DATOS DEL EQUIPO UTILIZADO"
        ws['N8']="EQUIPO"; ws['Q8']="MARCA"; ws['S8']="SERIE"
        ws['N9']="SONÓMETRO"; ws['Q9']="HD2010UC/A"
        ws['N10']="PISTÓFONO"; ws['Q10']="HD2020"

        row = 13
        for p in st.session_state.puntos:
            ws[f'B{row}']="PUNTO DE MONITOREO:"; ws[f'D{row}']=p['punto']
            ws[f'J{row}']="COORDENADAS ORIGEN NACIONAL:"; ws[f'L{row}']=p['coord']; row+=1
            ws[f'B{row}']="DESCRIPCIÓN DEL PUNTO DE MONITOREO:"; ws[f'F{row}']=p['desc']; row+=1
            ws[f'B{row}']="REGISTRO DEL BARRIDO PERIMETRAL (dB):"; ws[f'F{row}']=p['barrido']; row+=1
            headers=["Fecha de Toma","Hora de Toma","Calibración (dB)","Memoria / No. Estudio","LAeq,T (dB)\nIn situ","Velocidad del Viento Máx.\n(m/s)","Dirección del Viento","Temperatura\n(°C)","Humedad Relativa\n(%)","Se evidencias precipacitones","Fuente de Ruido","*Tipo de Ruido","Tiempo de Operación","Observaciones"]
            for c,h in enumerate(headers, start=2): ws.cell(row=row, column=c, value=h).font=Font(bold=True,size=8)
            row+=1
            ws.cell(row=row, column=2, value=p['fecha']); ws.cell(row=row, column=3, value=p['hora'])
            ws.cell(row=row, column=4, value=f"Inicial {p['cal_ini']}"); ws.cell(row=row, column=5, value=p['memoria'])
            ws.cell(row=row, column=6, value=p['laeq']); ws.cell(row=row, column=7, value=p['vmax'])
            ws.cell(row=row, column=8, value=p['vdir']); ws.cell(row=row, column=9, value=p['temp'])
            ws.cell(row=row, column=10, value=p['hum']); ws.cell(row=row, column=11, value=p['precip'])
            ws.cell(row=row, column=12, value=p['fuente']); ws.cell(row=row, column=13, value=p['tipo'])
            ws.cell(row=row, column=14, value=p['topera']); ws.cell(row=row, column=15, value=p['obs'])
            row+=1
            ws.cell(row=row, column=4, value=f"Final {p['cal_fin']}"); row+=2
            ws[f'B{row}']="DIBUJE EL ESQUEMA DEL PUNTO DE MONITOREO"; ws[f'B{row}'].font=Font(bold=True); row+=3

        ws[f'F{row}']="Responsable de la Medición:"; ws[f'K{row}']=responsable
        row+=2; ws[f'C{row}']="Elaboró:"; ws[f'H{row}']="Revisó:"; ws[f'L{row}']="Aprobó: Gerente G."

        bio = BytesIO(); wb.save(bio); bio.seek(0); return bio

    excel = generar_excel()
    st.download_button("📥 DESCARGAR EXCEL FORMATO R2-POE37-EP", excel, f"CAMPO_{municipio}_{datetime.now().strftime('%d%m%Y')}.xlsx", use_container_width=True, type="primary")
