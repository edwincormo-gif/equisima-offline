import streamlit as st
import io
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="EQUISIMA R2-POE37-EP V04 ORIGINAL", layout="wide")
st.title("EQUISIMA - Formato Original R2-POE37-EP")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

tab1, tab2 = st.tabs(["📋 DATOS DE CAMPO", "📥 DESCARGA ORIGINAL"])

with tab1:
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S A S")
        proyecto = st.text_input("NOMBRE DEL PROYECTO", "Ataco Tolima")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA")
        municipio = st.text_input("MUNICIPIO", "ATACO")
    with c2:
        plan_m = st.text_input("PLAN DE MUESTREO", "PM-001")
        punto_mon = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
        coord = st.text_input("COORDENADAS ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W")
    with c3:
        desc_punto = st.text_area("DESCRIPCIÓN DEL PUNTO", "En la entrada o vía principal")
        barrido = st.text_input("BARRIDO PERIMETRAL (dB)", "68")

    st.divider()
    col1,col2,col3,col4 = st.columns(4)
    with col1:
        fecha = st.date_input("Fecha de Toma", datetime(2026,9,9))
        hora = st.text_input("Hora de Toma", "21:01")
        calib = st.number_input("Calibración Inicial", 0.0, 140.0, 114.0)
    with col2:
        memoria = st.text_input("Memoria / No. Estudio", "1")
        laeq = st.number_input("LAeq,T In situ", 0.0, 140.0, 65.0)
        vel = st.number_input("Vel Viento Max (m/s)", 0.0, 20.0, 0.3)
    with col3:
        dirv = st.selectbox("Dirección Viento", ["N","NE","E","SE","S","SW","W","NW","Calma"])
        temp = st.number_input("Temperatura (°C)", -20.0, 60.0, 29.0)
        hum = st.number_input("Humedad (%)", 0, 100, 45)
    with col4:
        precip = st.selectbox("¿Precipitaciones?", ["NO","SI"])
        fuente = st.text_input("Fuente Ruido", "mineria")
        tipo = st.selectbox("Tipo Ruido", ["Continuo","Intermitente","Impulsivo"])

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "fecha": str(fecha), "hora": hora, "calib": calib, "memoria": memoria,
            "laeq": laeq, "vel": vel, "dirv": dirv, "temp": temp, "hum": hum,
            "precip": precip, "fuente": fuente, "tipo": tipo
        })
        st.success(f"Punto agregado! Total {len(st.session_state.puntos)}")

with tab2:
    st.info(f"Puntos guardados: {len(st.session_state.puntos)}")
    if st.session_state.puntos:
        try:
            wb = load_workbook("plantilla.xlsx")
            ws = wb["Datos de Campo Emision"]

            # Llenamos encabezado EXACTO donde va en tu plantilla
            ws["E8"].value = cliente
            ws["E9"].value = proyecto
            ws["E10"].value = depto
            ws["E11"].value = municipio
            ws["E12"].value = plan_m
            ws["D15"].value = punto_mon
            ws["L15"].value = coord
            ws["F16"].value = desc_punto
            ws["F17"].value = barrido

            # Llenamos datos desde fila 21 (que es donde tu plantilla tiene datos)
            fila_inicio = 21
            for i, p in enumerate(st.session_state.puntos):
                f = fila_inicio + i
                ws.cell(row=f, column=2).value = p["fecha"]  # B = Fecha
                ws.cell(row=f, column=3).value = p["hora"]   # C = Hora
                ws.cell(row=f, column=4).value = p["calib"]  # D = Calib
                ws.cell(row=f, column=5).value = p["memoria"]
                ws.cell(row=f, column=6).value = p["laeq"]
                ws.cell(row=f, column=7).value = p["vel"]
                ws.cell(row=f, column=8).value = p["dirv"]
                ws.cell(row=f, column=9).value = p["temp"]
                ws.cell(row=f, column=10).value = p["hum"]
                ws.cell(row=f, column=11).value = p["precip"]
                ws.cell(row=f, column=12).value = p["fuente"]
                ws.cell(row=f, column=13).value = p["tipo"]

            output = io.BytesIO()
            wb.save(output)

            st.download_button(
                "📥 DESCARGAR EXCEL IDÉNTICO A ORIGINAL",
                output.getvalue(),
                file_name=f"R2-POE37-EP_{municipio}_{punto_mon}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            st.success("Excel generado con tu formato original, con logos, colores y celdas combinadas")
        except FileNotFoundError:
            st.error("❌ No encuentro plantilla.xlsx en el repo. Súbela primero al lado de app.py")
    else:
        st.warning("Agrega puntos en pestaña 1")
