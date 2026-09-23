import streamlit as st
import pandas as pd
from datetime import datetime, time
import io

st.set_page_config(page_title="EQUISIMA - R2-POE37-EP", layout="wide")
st.title("📋 DATOS DE CAMPO EMISION DE RUIDO - R2-POE37-EP V04")

if "puntos" not in st.session_state:
    st.session_state.puntos = []
if "fotos" not in st.session_state:
    st.session_state.fotos = []

# --- ENCABEZADO ---
c1,c2,c3 = st.columns(3)
with c1:
    cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
    depto = st.text_input("DEPARTAMENTO", "TOLIMA")
    municipio = st.text_input("MUNICIPIO", "ATACO")
with c2:
    punto = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
    coord = st.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W")
    desc_punto = st.text_area("DESCRIPCIÓN DEL PUNTO", "En la entrada o vía principal")
with c3:
    son_marca = st.text_input("SONÓMETRO MARCA", "HD2010UC/A")
    son_serie = st.text_input("SONÓMETRO SERIE", "15031643825")

st.divider()
with st.form("form_campo", clear_on_submit=True):
    col1,col2,col3,col4 = st.columns(4)
    with col1:
        fecha = st.date_input("Fecha de Toma", datetime(2026,9,9))
        hora = st.time_input("Hora de Toma", time(21,1))
        calib = st.number_input("Calibración (dB)", 90.0, 120.0, 114.0)
        memoria = st.text_input("Memoria / No. Estudio", "1")
    with col2:
        laeq_insitu = st.number_input("LAeq,T (dB) In situ", 0.0, 140.0, 65.0)
        vel_viento = st.number_input("Velocidad Viento Máx (m/s)", 0.0, 20.0, 0.3)
        dir_viento = st.selectbox("Dirección Viento", ["N","NE","E","SE","S","SW","W","NW"])
        temp = st.number_input("Temperatura (°C)", -10.0, 60.0, 29.0)
    with col3:
        humedad = st.number_input("Humedad Relativa (%)", 0, 100, 45)
        precip = st.selectbox("¿Precipitaciones?", ["NO","SI"])
        fuente_ruido = st.text_input("Fuente Ruido", "mineria")
        tipo_ruido = st.selectbox("Tipo Ruido", ["EMISION","RESIDUAL"])
    with col4:
        tiempo_op = st.number_input("Tiempo Operación (h)", 0.0, 24.0, 1.0)
        obs_detalle = st.text_area("Detalle camino / maquinaria")
    
    if st.form_submit_button("📍 AGREGAR PUNTO COMPLETO"):
        nuevo = {
            "Fecha de Toma": fecha, "Hora de Toma": str(hora),
            "Calibración (dB)": calib, "LAeq,T (dB) In situ": laeq_insitu,
            "Velocidad Viento Máx. (m/s)": vel_viento,
            "Dirección del Viento": dir_viento,
            "Temperatura (°C)": temp, "Humedad Relativa (%)": humedad,
            "Precipitaciones": precip, "Fuente de Ruido": fuente_ruido,
            "Tipo de Ruido": tipo_ruido, "Tiempo Operación": tiempo_op,
            "Observaciones": obs_detalle, "COORDENADAS": coord
        }
        st.session_state.puntos.append(nuevo)
        st.rerun()

if st.session_state.puntos:
    df = pd.DataFrame(st.session_state.puntos)
    st.dataframe(df, use_container_width=True)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Datos de Campo Emision", index=False)
    st.download_button("📥 DESCARGAR EXCEL FORMATO R2-POE37-EP", output.getvalue(), file_name="Campo_Emision.xlsx")
