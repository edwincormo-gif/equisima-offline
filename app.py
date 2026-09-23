import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="EQUISIMA RES 627 + FORMATO", layout="wide")
st.title("EQUISIMA - Formato R2-POE37-EP V04 + Res 627")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

# LIMITES RES 627 DE 2006
NORMA = {
    "Sector A. Tranquilidad y Silencio": {"Diurno": 55, "Nocturno": 45},
    "Sector B. Tranquilidad y Ruido Moderado": {"Diurno": 65, "Nocturno": 50},
    "Sector C. Ruido Intermedio Restringido": {"Diurno": 75, "Nocturno": 70},
    "Sector C. Industrial": {"Diurno": 75, "Nocturno": 75},
    "Sector C. Centro Ciudad / Comercial": {"Diurno": 70, "Nocturno": 55},
    "Sector D. Suburbana o Rural": {"Diurno": 55, "Nocturno": 45},
}

tab1, tab2, tab3, tab4 = st.tabs(["📋 DATOS PLANTILLA", "🗺️ PUNTOS", "⚖️ RES 627", "📸 DESCARGA FORMATO ORIGINAL"])

with tab1:
    st.subheader("Datos de la Plantilla Original")
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE:", "Rueda Inversiones S A S")
        proyecto = st.text_input("NOMBRE DEL PROYECTO:", "Ataco Tolima")
        depto = st.text_input("DEPARTAMENTO:", "TOLIMA")
        municipio = st.text_input("MUNICIPIO:", "ATACO")
        plan_m = st.text_input("PLAN DE MUESTREO:", "PM-001")
    with c2:
        punto_mon = st.text_input("PUNTO DE MONITOREO:", "Punto 1 nocturno")
        coord = st.text_input("COORDENADAS ORIGEN NACIONAL:", "3°35'11.30\"N 75°23'27.72\"W")
        desc_punto = st.text_area("DESCRIPCIÓN DEL PUNTO:", "En la entrada o vía principal")
        barrido = st.text_input("REGISTRO BARRIDO PERIMETRAL (dB):", "68")
        sector = st.selectbox("SECTOR USO DE SUELO RES 627", list(NORMA.keys()))
    with c3:
        periodo = st.selectbox("ESCENARIO MEDICIÓN", ["Diurno","Nocturno"], index=1)
        fecha = st.date_input("Fecha de Toma", datetime.now())
        hora = st.text_input("Hora de Toma", "21:01")
        memoria = st.text_input("Memoria / No. Estudio", "1")

    st.divider()
    col1,col2,col3,col4 = st.columns(4)
    with col1:
        calib_ini = st.text_input("Calibración Inicial dB", "114")
        laeq = st.number_input("LAeq,T (dB) In situ - TU MEDICIÓN", 0.0, 140.0, 65.0)
    with col2:
        vel = st.number_input("Vel Viento Max (m/s)", 0.0, 20.0, 0.3)
        dirv = st.text_input("Dirección Viento", "N")
    with col3:
        temp = st.number_input("Temperatura (°C)", -20.0, 60.0, 29.0)
        hum = st.number_input("Humedad Relativa (%)", 0, 100, 45)
    with col4:
        precip = st.selectbox("Precipitaciones Si/No", ["NO","SI"])
        fuente = st.text_input("Fuente Ruido / Equipo", "mineria")
        tipo_ruido = st.text_input("Tipo de Ruido", "Continuo")

    limite = NORMA[sector][periodo]
    cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"
    st.info(f"Límite {sector} - {periodo}: **{limite} dB** | Tu medición: **{laeq} dB** | Resultado: **{cumple}**")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "PROYECTO": proyecto, "DEPARTAMENTO": depto, "MUNICIPIO": municipio,
            "PLAN": plan_m, "PUNTO": punto_mon, "COORD": coord, "DESC": desc_punto, "BARRIDO": barrido,
            "SECTOR": sector, "PERIODO": periodo, "FECHA": str(fecha), "HORA": hora,
            "CALIB": calib_ini, "MEMORIA": memoria, "LAEQ": laeq, "LIMITE": limite, "CUMPLE": cumple,
            "VEL": vel, "DIR": dirv, "TEMP": temp, "HUM": hum, "PRECIP": precip, "FUENTE": fuente, "TIPO": tipo_ruido
        })
        st.success(f"Punto agregado: {cumple}")
        st.balloons()

with tab2:
    if st.session_state.puntos:
        st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
        if st.button("🗑️ Borrar todo"):
            st.session_state.puntos = []
            st.rerun()
    else:
        st.info("Agrega datos en pestaña 1")

with tab3:
    st.subheader("Comparación Res 627")
    if st.session_state.puntos:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df[["PUNTO","SECTOR","PERIODO","LAEQ","LIMITE","CUMPLE"]], use_container_width=True)
        nc = df[df["CUMPLE"] == "NO CUMPLE"]
        if not nc.empty:
            st.error(f"🚨 {len(nc)} puntos NO CUMPLEN")
        else:
            st.success("✅ Todos CUMPLEN")
    else:
        st.info("Sin datos aún")

with tab4:
    fotos_up = st.file_uploader("Sube fotos del punto (JPG/PNG)", type=["jpg","jpeg","png"], accept_multiple_files=True, key="fotos_final_res627")
    if fotos_up:
        for f in fotos_up:
            st.image(f, width=200, caption=f.name)

    st.divider()
    if not st.session_state.puntos:
        st.warning("Primero agrega puntos")
    else:
        if st.button("📥 GENERAR EXCEL EN FORMATO ORIGINAL", type="primary", use_container_width=True):
            try:
                # Esta es la función que arregla el error MergedCell
                wb = load_workbook("plantilla.xlsx")
                ws = wb["Datos de Campo Emision"]

                def set_val(r, c, v):
                    for mr in ws.merged_cells.ranges:
                        if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                            ws.cell(row=mr.min_row, column=mr.min_col).value = v
                            return
                    ws.cell(row=r, column=c).value = v

                p0 = st.session_state.puntos[0]
                set_val(8,5,p0["CLIENTE"]) # E8
                set_val(9,5,p0["PROYECTO"]) # E9
                set_val(10,5,p0["DEPARTAMENTO"])
                set_val(11,5,p0["MUNICIPIO"])
                set_val(12,5,p0["PLAN"])
                set_val(15,4,p0["PUNTO"]) # D15
                set_val(15,12,p0["COORD"]) # L15
                set_val(16,6,p0["DESC"])
                set_val(17,6,p0["BARRIDO"])

                # Datos de medición desde fila 21
                fila = 21
                for p in st.session_state.puntos:
                    ws.cell(row=fila, column=2).value = p["FECHA"]
                    ws.cell(row=fila, column=3).value = p["HORA"]
                    ws.cell(row=fila, column=4).value = p["CALIB"]
                    ws.cell(row=fila, column=5).value = p["MEMORIA"]
                    ws.cell(row=fila, column=6).value = p["LAEQ"]
                    ws.cell(row=fila, column=7).value = p["VEL"]
                    ws.cell(row=fila, column=8).value = p["DIR"]
                    ws.cell(row=fila, column=9).value = p["TEMP"]
                    ws.cell(row=fila, column=10).value = p["HUM"]
                    ws.cell(row=fila, column=11).value = p["PRECIP"]
                    ws.cell(row=fila, column=12).value = p["FUENTE"]
                    ws.cell(row=fila, column=13).value = p["TIPO"]
                    fila += 1

                out = io.BytesIO()
                wb.save(out)

                st.download_button(
                    "📥 DESCARGAR AHORA EN FORMATO PLANTILLA",
                    out.getvalue(),
                    file_name=f"R2-POE37-EP_{p0['MUNICIPIO']}_{p0['PUNTO']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                st.success("¡Listo! Ahora sí descarga en tu formato original y ya evalúa RES 627")

            except FileNotFoundError:
                st.error("No encuentro plantilla.xlsx. Súbela a la raíz del repo como hiciste en la captura")
                # Plan B: descarga simple si no está la plantilla
                df = pd.DataFrame(st.session_state.puntos)
                out2 = io.BytesIO()
                with pd.ExcelWriter(out2, engine='openpyxl') as w:
                    df.to_excel(w, index=False)
                st.download_button("📥 Descargar Excel simple (mientras subes plantilla)", out2.getvalue(), "res627.xlsx", use_container_width=True)
            except Exception as e:
                st.error(f"Error: {e}")
