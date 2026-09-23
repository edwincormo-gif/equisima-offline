import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="EQUISIMA R2-POE37-EP V04", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

tab1, tab2, tab3 = st.tabs(["📋 DATOS DE CAMPO", "🗺️ PUNTOS GUARDADOS", "📸 FOTOS Y EXCEL"])

with tab1:
    st.subheader("Datos Generales - Igual a tu Excel original")
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA")
        municipio = st.text_input("MUNICIPIO", "ATACO")
    with c2:
        punto_mon = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
        coord = st.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W")
        desc = st.text_area("DESCRIPCIÓN DEL PUNTO", "En la entrada o vía principal")
    with c3:
        marca = st.text_input("SONÓMETRO MARCA", "HD2010UC/A")
        serie = st.text_input("SONÓMETRO SERIE", "15031643825")
        cert = st.text_input("Cert. Calibración", "2024-05-20")

    st.divider()
    st.markdown("**Agregar punto - Formato R2-POE37-EP**")

    col1,col2,col3,col4 = st.columns(4)
    with col1:
        fecha = st.date_input("Fecha de Toma", datetime(2026,9,9))
        hora = st.text_input("Hora de Toma", "21:01")
        calib = st.number_input("Calibración (dB)", 0.0, 140.0, 114.0)
        memoria = st.text_input("Memoria / No. Estudio", "1")
    with col2:
        laeq = st.number_input("LAeq,T (dB) In situ", 0.0, 140.0, 65.0)
        vel_viento = st.number_input("Velocidad Viento Máx (m/s)", 0.0, 20.0, 0.3)
        dir_viento = st.selectbox("Dirección del Viento", ["N","NE","E","SE","S","SW","W","NW","Calma"])
        temp = st.number_input("Temperatura (°C)", -10.0, 60.0, 29.0)
    with col3:
        humedad = st.number_input("Humedad Relativa (%)", 0, 100, 45)
        precip = st.selectbox("¿Precipitaciones?", ["NO","SI"])
        fuente = st.text_input("Fuente de Ruido / Equipo", "mineria")
        tipo_ruido = st.selectbox("Tipo de Ruido", ["EMISION","RESIDUAL"])
    with col4:
        tiempo_op = st.number_input("Tiempo Operación (h)", 0.0, 24.0, 1.0)
        barrido = st.text_input("Barrido Perimetral", "No")
        obs_det = st.text_area("Observaciones")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append([
            cliente, depto, municipio, punto_mon, coord, desc,
            str(fecha), hora, calib, memoria, laeq, vel_viento,
            dir_viento, temp, humedad, precip, fuente, tipo_ruido,
            tiempo_op, barrido, obs_det, marca, serie
        ])
        st.success(f"Punto agregado! Total: {len(st.session_state.puntos)}")
        st.balloons()

with tab2:
    if st.session_state.puntos:
        df = pd.DataFrame(st.session_state.puntos, columns=[
            "CLIENTE","DEPARTAMENTO","MUNICIPIO","PUNTO DE MONITOREO","COORDENADAS","DESCRIPCIÓN",
            "Fecha de Toma","Hora de Toma","Calibración (dB)","Memoria","LAeq,T (dB) In situ",
            "Vel Viento Máx (m/s)","Dirección Viento","Temperatura (°C)","Humedad (%)",
            "Precipitaciones","Fuente Ruido","Tipo Ruido","Tiempo Op","Barrido","Obs",
            "Marca Son","Serie Son"
        ])
        st.dataframe(df, use_container_width=True)
        if st.button("🗑️ Borrar todos"):
            st.session_state.puntos = []
            st.rerun()
    else:
        st.info("No hay puntos aún")

with tab3:
    st.subheader("Fotos de Campo")
    # FIX DEL ERROR: key diferente a session_state
    fotos_up = st.file_uploader("Sube fotos JPG/PNG", type=["jpg","jpeg","png"], accept_multiple_files=True, key="uploader_fotos_fix")
    if fotos_up:
        for f in fotos_up:
            st.image(f, width=200, caption=f.name)

    st.divider()
    st.subheader("Descargar EXACTO a tu plantilla")
    if st.session_state.puntos:
        # Creamos Excel idéntico a tu plantilla R2-POE37-EP
        wb = Workbook()
        ws = wb.active
        ws.title = "Datos de Campo Emision"

        # Encabezados iguales a tu Excel original
        headers = ["CLIENTE","DEPARTAMENTO","MUNICIPIO","PUNTO DE MONITOREO","COORDENADAS","DESCRIPCIÓN DEL PUNTO",
                   "Fecha de Toma","Hora de Toma","Calibración (dB)","Memoria / No. Estudio","LAeq,T (dB) In situ",
                   "Velocidad Viento Máx. (m/s)","Dirección del Viento","Temperatura (°C)","Humedad Relativa (%)",
                   "Precipitaciones","Fuente de Ruido","Tipo de Ruido","Tiempo de Operación","Barrido Perimetral",
                   "Observaciones","SONÓMETRO MARCA","SONÓMETRO SERIE"]
        ws.append(headers)
        for row in st.session_state.puntos:
            ws.append(row)

        # Ajuste ancho columnas
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18

        output = io.BytesIO()
        wb.save(output)

        st.download_button(
            "📥 DESCARGAR EXCEL FORMATO ORIGINAL R2-POE37-EP",
            output.getvalue(),
            file_name="R2-POE37-EP_V04_ATACO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    else:
        st.warning("Agrega al menos 1 punto para descargar")
