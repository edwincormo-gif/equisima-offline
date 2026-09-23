import streamlit as st
import pandas as pd
from datetime import datetime, time
import io

st.set_page_config(page_title="EQUISIMA - R2-POE37-EP V04", layout="wide")
st.title("EQUISIMA - Medición de Ruido Ataco")

if "puntos" not in st.session_state:
    st.session_state.puntos = []
if "fotos" not in st.session_state:
    st.session_state.fotos = []

tab1, tab2, tab3 = st.tabs(["📋 DATOS DE CAMPO", "🗺️ PUNTOS Y MAPA", "📸 FOTOS Y DESCARGA"])

with tab1:
    st.subheader("Datos Generales - R2-POE37-EP V04 (2024-05-20)")
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S.A.S.")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA")
        municipio = st.text_input("MUNICIPIO", "ATACO")
    with c2:
        punto_mon = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
        coord = st.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W")
        desc_punto = st.text_area("DESCRIPCIÓN DEL PUNTO", "Entrada vía principal")
    with c3:
        son_marca = st.text_input("SONÓMETRO MARCA", "HD2010UC/A")
        son_serie = st.text_input("SONÓMETRO SERIE", "15031643825")
        cal_lab = st.text_input("Cert. Calibración", "2024")

    st.divider()
    st.markdown("**Datos de Campo por Punto (obligatorios Res. 0627)**")
    # FORM estable para celular
    with st.form("form_campo_completo", clear_on_submit=True):
        col1,col2,col3,col4 = st.columns(4)
        with col1:
            fecha = st.date_input("Fecha de Toma", datetime(2026,9,9))
            hora = st.time_input("Hora de Toma", time(21,1))
            calib_db = st.number_input("Calibración (dB)", 90.0, 120.0, 114.0)
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
            barrido = st.number_input("Barrido Perimetral (dB)", 0.0, 140.0, 0.0)
            obs = st.text_area("Observaciones / Maquinaria")
        
        if st.form_submit_button("📍 AGREGAR PUNTO", use_container_width=True):
            st.session_state.puntos.append({
                "CLIENTE": cliente, "DEPARTAMENTO": depto, "MUNICIPIO": municipio,
                "PUNTO DE MONITOREO": punto_mon, "COORDENADAS": coord,
                "Fecha de Toma": fecha, "Hora de Toma": str(hora),
                "Calibración (dB)": calib_db, "Memoria / No. Estudio": memoria,
                "LAeq,T (dB) In situ": laeq, "Velocidad Viento Máx. (m/s)": vel_viento,
                "Dirección del Viento": dir_viento, "Temperatura (°C)": temp,
                "Humedad Relativa (%)": humedad, "Precipitaciones": precip,
                "Fuente de Ruido": fuente, "Tipo de Ruido": tipo_ruido,
                "Tiempo Operación": tiempo_op, "Barrido Perimetral": barrido,
                "Observaciones": obs
            })
            st.success("Punto agregado")
            st.rerun()

with tab2:
    st.subheader("Puntos Guardados")
    if st.session_state.puntos:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df, use_container_width=True)
        st.map(pd.DataFrame([{"lat": 3.586472, "lon": -75.391033}])) # ejemplo
        if st.button("🗑️ Borrar todos los puntos"):
            st.session_state.puntos = []
            st.rerun()
    else:
        st.info("Aún no hay puntos agregados. Ve a la pestaña DATOS DE CAMPO")

with tab3:
    st.subheader("Fotos de Campo")
    with st.form("form_fotos", clear_on_submit=True):
        fotos_up = st.file_uploader("Sube fotos JPG/PNG", type=["jpg","jpeg","png"], accept_multiple_files=True)
        desc_foto = st.text_input("Descripción de la foto")
        if st.form_submit_button("📸 AGREGAR FOTOS", use_container_width=True):
            if fotos_up:
                for f in fotos_up:
                    st.session_state.fotos.append({"nombre": f.name, "desc": desc_foto, "file": f})
                st.success(f"{len(fotos_up)} fotos guardadas")
                st.rerun()
    
    if st.session_state.fotos:
        for foto in st.session_state.fotos:
            st.write(f"📷 {foto['nombre']} - {foto['desc']}")

    st.divider()
    st.subheader("Descargar Excel")
    if st.session_state.puntos:
        df_export = pd.DataFrame(st.session_state.puntos)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Hoja 1: Como tu plantilla original
            df_export.to_excel(writer, sheet_name="Datos de Campo Emision", index=False)
            # Hoja 2: Resumen para informe
            df_export[["PUNTO DE MONITOREO","LAeq,T (dB) In situ","Temperatura (°C)","Dirección del Viento","Velocidad Viento Máx. (m/s)"]].to_excel(writer, sheet_name="Resumen Puntos", index=False)
        
        st.download_button("📥 DESCARGAR EXCEL SEPARADO POR HOJAS", output.getvalue(), file_name="R2-POE37-EP_Ataco.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.warning("Agrega al menos 1 punto para poder descargar")
