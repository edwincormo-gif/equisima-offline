import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="EQUISIMA R2-POE37-EP V04", layout="wide")
st.title("EQUISIMA - Formato Campo R2-POE37-EP V04")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

tab1, tab2, tab3 = st.tabs(["📋 DATOS DE CAMPO", "🗺️ PUNTOS GUARDADOS", "📸 FOTOS Y EXCEL"])

with tab1:
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S A S")
        proyecto = st.text_input("NOMBRE PROYECTO", "Ataco Tolima")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA")
        municipio = st.text_input("MUNICIPIO", "ATACO")
    with c2:
        plan_m = st.text_input("PLAN DE MUESTREO", "PM-001")
        punto_mon = st.text_input("PUNTO DE MONITOREO", "Punto 1 nocturno")
        coord = st.text_input("COORDENADAS", "3°35'11.30\"N 75°23'27.72\"W")
    with c3:
        desc_punto = st.text_area("DESCRIPCIÓN DEL PUNTO", "En la entrada o vía principal")
        barrido = st.text_input("Barrido Perimetral (dB)", "68")

    st.divider()
    col1,col2,col3,col4 = st.columns(4)
    with col1:
        fecha = st.date_input("Fecha", datetime(2026,9,9))
        hora = st.text_input("Hora", "21:01")
        calib = st.number_input("Calibración (dB)", 0.0, 140.0, 114.0)
    with col2:
        memoria = st.text_input("Memoria / No.", "1")
        laeq = st.number_input("LAeq,T (dB)", 0.0, 140.0, 65.0)
        vel = st.number_input("Vel Viento (m/s)", 0.0, 20.0, 0.3)
    with col3:
        dirv = st.selectbox("Dirección Viento", ["N","NE","E","SE","S","SW","W","NW","Calma"])
        temp = st.number_input("Temp (°C)", -20.0, 60.0, 29.0)
        hum = st.number_input("Humedad (%)", 0, 100, 45)
    with col4:
        precip = st.selectbox("Precipitaciones?", ["NO","SI"])
        fuente = st.text_input("Fuente Ruido", "mineria")
        tipo = st.selectbox("Tipo Ruido", ["EMISION","RESIDUAL"])

    obs = st.text_area("Observaciones")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "cliente": cliente, "proyecto": proyecto, "depto": depto, "muni": municipio,
            "plan": plan_m, "punto": punto_mon, "coord": coord, "desc": desc_punto,
            "fecha": str(fecha), "hora": hora, "calib": calib, "memoria": memoria,
            "laeq": laeq, "vel": vel, "dirv": dirv, "temp": temp, "hum": hum,
            "precip": precip, "fuente": fuente, "tipo": tipo, "barrido": barrido, "obs": obs
        })
        st.success(f"Punto guardado! Total: {len(st.session_state.puntos)}")
        st.balloons()

with tab2:
    if st.session_state.puntos:
        st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
        if st.button("🗑️ Borrar todo"):
            st.session_state.puntos = []
            st.rerun()
    else:
        st.info("Aún no has agregado puntos")

with tab3:
    fotos_up = st.file_uploader("Sube fotos", type=["jpg","jpeg","png"], accept_multiple_files=True, key="uploader_fotos_final_v3")
    if fotos_up:
        for f in fotos_up:
            st.image(f, width=200)

    st.divider()
    if not st.session_state.puntos:
        st.warning("Agrega puntos en pestaña 1")
    else:
        st.success(f"{len(st.session_state.puntos)} punto(s) listos")
        # BOTON QUE ARREGLA EL ERROR DE MERGEDCELL
        if st.button("📥 GENERAR EXCEL ORIGINAL", type="primary", use_container_width=True):
            try:
                wb = load_workbook("plantilla.xlsx")
                ws = wb["Datos de Campo Emision"]

                def escribir(celda_texto, valor):
                    # Si la celda está combinada, escribe en la celda madre
                    cell = ws[celda_texto]
                    if str(type(cell)) == "<class 'openpyxl.cell.cell.MergedCell'>":
                        for rango in ws.merged_cells.ranges:
                            if cell.coordinate in rango:
                                ws.cell(row=rango.min_row, column=rango.min_col).value = valor
                                return
                    else:
                        ws[celda_texto] = valor

                p0 = st.session_state.puntos[0]
                escribir("E8", p0["cliente"])
                escribir("E9", p0["proyecto"])
                escribir("E10", p0["depto"])
                escribir("E11", p0["muni"])
                escribir("E12", p0["plan"])
                escribir("D15", p0["punto"])
                escribir("L15", p0["coord"])
                escribir("F16", p0["desc"])
                escribir("F17", p0["barrido"])

                fila = 21
                for p in st.session_state.puntos:
                    ws.cell(row=fila, column=2).value = p["fecha"]
                    ws.cell(row=fila, column=3).value = p["hora"]
                    ws.cell(row=fila, column=4).value = p["calib"]
                    ws.cell(row=fila, column=5).value = p["memoria"]
                    ws.cell(row=fila, column=6).value = p["laeq"]
                    ws.cell(row=fila, column=7).value = p["vel"]
                    ws.cell(row=fila, column=8).value = p["dirv"]
                    ws.cell(row=fila, column=9).value = p["temp"]
                    ws.cell(row=fila, column=10).value = p["hum"]
                    ws.cell(row=fila, column=11).value = p["precip"]
                    ws.cell(row=fila, column=12).value = p["fuente"]
                    ws.cell(row=fila, column=13).value = p["tipo"]
                    ws.cell(row=fila, column=14).value = p["obs"]
                    fila += 1

                output = io.BytesIO()
                wb.save(output)

                st.download_button(
                    "📥 DESCARGAR AHORA EXCEL IDÉNTICO",
                    output.getvalue(),
                    file_name=f"R2-POE37-EP_{p0['muni']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                st.success("¡Excel generado con tu formato original!")
            except Exception as e:
                st.error(f"Error: {e}")
