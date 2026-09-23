import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

st.set_page_config(page_title="EQUISIMA R2-POE37-EP", layout="wide")
st.title("EQUISIMA - Formato Campo Ruido")

if "puntos" not in st.session_state:
    st.session_state.puntos = []
if "fotos" not in st.session_state:
    st.session_state.fotos = []

tab1, tab2, tab3 = st.tabs(["📋 DATOS DE CAMPO", "🗺️ PUNTOS", "📸 FOTOS Y EXCEL"])

with tab1:
    st.subheader("Datos Generales")
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.", key="cli")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA", key="dep")
        muni = st.text_input("MUNICIPIO", "ATACO", key="mun")
    with c2:
        punto_nom = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno", key="pnom")
        coord = st.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W", key="coord")
    with c3:
        marca = st.text_input("SONOMETRO MARCA", "HD2010UC/A", key="marca")
        serie = st.text_input("SONOMETRO SERIE", "15031643825", key="serie")

    st.divider()
    st.markdown("### Datos del punto (OBLIGATORIOS)")
    fecha = st.date_input("Fecha de Toma", datetime(2026,9,9), key="fecha")
    hora = st.time_input("Hora de Toma", time(21,1), key="hora")
    
    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        calib = st.number_input("Calibración (dB)", 0.0, 140.0, 114.0, key="cal")
        laeq = st.number_input("LAeq,T (dB) In situ", 0.0, 140.0, 65.0, key="laeq2")
        temp = st.number_input("Temperatura (°C)", -20.0, 60.0, 29.0, key="temp2")
    with cc2:
        vel = st.number_input("Velocidad Viento Máx (m/s)", 0.0, 20.0, 0.3, key="vel2")
        dirv = st.selectbox("Dirección del Viento", ["N","NE","E","SE","S","SW","W","NW","Calma"], key="dir2")
        hum = st.number_input("Humedad Relativa (%)", 0, 100, 45, key="hum2")
    with cc3:
        prec = st.selectbox("¿Precipitaciones?", ["NO","SI"], key="prec2")
        fuente = st.text_input("Fuente de Ruido", "mineria", key="fuente")
        tipo = st.selectbox("Tipo de Ruido", ["EMISION","RESIDUAL"], key="tipo")
    
    obs = st.text_area("Observaciones / Descripción del punto", key="obs")

    # BOTON FUERA DE FORM - ESTE SI FUNCIONA EN CELULAR
    if st.button("📍 AGREGAR PUNTO AHORA", type="primary", use_container_width=True, key="btn_agregar_final"):
        st.session_state.puntos.append({
            "CLIENTE": cliente,
            "DEPARTAMENTO": depto,
            "MUNICIPIO": muni,
            "PUNTO": punto_nom,
            "COORDENADAS": coord,
            "Fecha": str(fecha),
            "Hora": str(hora),
            "Calibración (dB)": calib,
            "LAeq,T (dB) In situ": laeq,
            "Temperatura (°C)": temp,
            "Velocidad Viento (m/s)": vel,
            "Dirección Viento": dirv,
            "Humedad (%)": hum,
            "Precipitaciones": prec,
            "Fuente Ruido": fuente,
            "Tipo Ruido": tipo,
            "Observaciones": obs
        })
        st.success(f"Punto guardado! Total: {len(st.session_state.puntos)}")
        st.balloons()

with tab2:
    st.subheader(f"Puntos guardados: {len(st.session_state.puntos)}")
    if st.session_state.puntos:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df, use_container_width=True)
        if st.button("🗑️ Borrar todo"):
            st.session_state.puntos = []
            st.rerun()
    else:
        st.info("Aún no has agregado puntos. Ve a pestaña 1")

with tab3:
    st.subheader("Fotos y Descarga")
    fotos = st.file_uploader("Sube fotos", type=["jpg","jpeg","png"], accept_multiple_files=True, key="fotos")
    if fotos:
        for f in fotos:
            st.image(f, width=200)
    
    st.divider()
    if st.session_state.puntos:
        df_out = pd.DataFrame(st.session_state.puntos)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_out.to_excel(writer, sheet_name="Datos de Campo Emision", index=False)
        st.download_button("📥 DESCARGAR EXCEL", output.getvalue(), file_name="R2-POE37-EP_Ataco.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.warning("Agrega puntos para descargar Excel")
