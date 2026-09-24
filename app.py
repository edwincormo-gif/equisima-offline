import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="EQUISIMA V7 ORIGINAL COMPLETO", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04 COMPLETO + RES 627")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

NORMA = {
    "A. Tranquilidad y Silencio": {"Diurno": 55, "Nocturno": 45},
    "B. Tranquilidad y Ruido Moderado": {"Diurno": 65, "Nocturno": 50},
    "C. Ruido Intermedio Restringido": {"Diurno": 75, "Nocturno": 70},
    "C. Industrial": {"Diurno": 75, "Nocturno": 75},
    "C. Centro Ciudad / Comercial": {"Diurno": 70, "Nocturno": 55},
    "D. Rural": {"Diurno": 55, "Nocturno": 45},
}

tab1, tab2 = st.tabs(["📋 DATOS DE CAMPO", "📥 DESCARGA FORMATO ORIGINAL"])

with tab1:
    st.subheader("Datos Generales (como tu plantilla original)")
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE:", "Rueda Inversiones S A S")
        proyecto = st.text_input("NOMBRE PROYECTO:", "Ataco Tolima")
        municipio = st.text_input("MUNICIPIO:", "ATACO")
        depto = st.text_input("DEPARTAMENTO:", "TOLIMA")
    with c2:
        punto = st.text_input("PUNTO:", "Punto 1")
        coord_n = st.text_input("COORD N: ORIGEN NACIONAL", "3°35'11N 75°23'27W")
        coord_c12 = st.text_input("COORD C12: CTM12", "E 876543 N 912345")
        desc = st.text_area("DESCRIPCIÓN PUNTO", "En la entrada")
    with c3:
        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 2)
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 0)
        plan = st.text_input("PLAN MUESTREO", "PM-001")
        barrido = st.text_input("BARRIDO PERIMETRAL dB", "68")

    st.divider()
    st.subheader("Datos de Medición (todos los del formato original)")
    col1,col2,col3,col4 = st.columns(4)
    with col1:
        fecha = st.date_input("FECHA", datetime.now())
        hora = st.text_input("HORA", "21:01")
        calib = st.text_input("Calibración Inicial (dB)", "114.0")
        memoria = st.text_input("Memoria / No.", "1")
    with col2:
        laeq = st.number_input("LAeq,T (dB)", 0.0, 140.0, 65.0)
        vel = st.number_input("Vel Viento (m/s)", 0.0, 20.0, 0.3)
        dirv = st.text_input("Dir Viento", "N")
    with col3:
        temp = st.number_input("Temp (°C)", -20.0, 60.0, 29.0)
        hum = st.number_input("Humedad (%)", 0, 100, 45)
        precip = st.selectbox("Precipitaciones", ["NO","SI"])
    with col4:
        fuente = st.text_input("Fuente Ruido", "mineria")
        tipo = st.text_input("Tipo Ruido", "Continuo")
        altura = st.text_input("Altura Sonómetro m", "1.5")

    limite = NORMA[sector][periodo]
    cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"
    st.info(f"**RES 627:** Sector {sector} - {periodo} = Límite {limite} dB | Medido {laeq} dB | **{cumple}**")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "PROYECTO": proyecto, "MUNICIPIO": municipio, "DEPTO": depto,
            "PUNTO": punto, "COORD_N": coord_n, "COORD_C12": coord_c12, "DESC": desc, "PLAN": plan, "BARRIDO": barrido,
            "FECHA": str(fecha), "HORA": hora, "CALIB": calib, "MEMORIA": memoria, "LAEQ": laeq,
            "VEL": vel, "DIR": dirv, "TEMP": temp, "HUM": hum, "PRECIP": precip,
            "FUENTE": fuente, "TIPO": tipo, "ALTURA": altura,
            "SECTOR": sector, "PERIODO": periodo, "LIMITE": limite, "CUMPLE": cumple
        })
        st.success("Punto agregado")
        st.balloons()

with tab2:
    if not st.session_state.puntos:
        st.warning("Agrega puntos primero")
    else:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df, use_container_width=True)

        if st.button("📥 GENERAR EXCEL FORMATO ORIGINAL COMPLETO", type="primary", use_container_width=True):
            wb = Workbook()
            ws = wb.active
            ws.title = "Datos de Campo Emision"

            bold = Font(bold=True, size=11)
            bold12 = Font(bold=True, size=12)
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            left = Alignment(horizontal="left", vertical="center")

            # TITULO
            ws.merge_cells("A1:O1")
            ws["A1"] = "R2-POE37-EP V04"
            ws["A1"].font = bold12
            ws["A1"].alignment = center

            # BLOQUE SUPERIOR - IGUAL A TU FOTO
            ws["A3"] = "CLIENTE:"; ws["A3"].font = bold
            ws.merge_cells("B3:O3")
            ws["B3"] = df.iloc[0]["CLIENTE"]

            ws["A4"] = "MUNICIPIO:"; ws["A4"].font = bold
            ws["B4"] = df.iloc[0]["MUNICIPIO"]
            ws["C4"] = "DEPARTAMENTO"; ws["C4"].font = bold
            ws.merge_cells("D4:O4")
            ws["D4"] = df.iloc[0]["DEPTO"]

            ws["A5"] = "PUNTO:"; ws["A5"].font = bold
            ws.merge_cells("B5:O5")
            ws["B5"] = df.iloc[0]["PUNTO"]

            ws["A6"] = "COORD N:"; ws["A6"].font = bold
            ws.merge_cells("B6:O6")
            ws["B6"] = df.iloc[0]["COORD_N"]

            ws["A7"] = "COORD C12:"; ws["A7"].font = bold
            ws.merge_cells("B7:O7")
            ws["B7"] = df.iloc[0]["COORD_C12"]

            ws["A8"] = "PROYECTO:"; ws["A8"].font = bold
            ws.merge_cells("B8:O8")
            ws["B8"] = df.iloc[0]["PROYECTO"]

            ws["A9"] = "DESCRIPCIÓN:"; ws["A9"].font = bold
            ws.merge_cells("B9:O9")
            ws["B9"] = df.iloc[0]["DESC"]

            # EVALUACION RES 627
            ws.merge_cells("A11:O11")
            ws["A11"] = f"EVALUACIÓN RES 627 - SECTOR: {df.iloc[0]['SECTOR']} {df.iloc[0]['PERIODO']} - LIMITE {limite} dB"
            ws["A11"].font = bold
            ws["A11"].alignment = center

            # TABLA COMPLETA ORIGINAL
            headers = ["FECHA","HORA","CALIB","MEMORIA","LAEQ","VEL VIENTO","DIR VIENTO","TEMP","HUM","PRECIP","FUENTE","TIPO","BARRIDO","LIMITE","CUMPLE"]
            for i, h in enumerate(headers, start=1):
                c = ws.cell(row=12, column=i)
                c.value = h
                c.font = bold
                c.alignment = center
                c.border = border

            fila = 13
            for _, p in df.iterrows():
                vals = [p["FECHA"], p["HORA"], p["CALIB"], p["MEMORIA"], p["LAEQ"], p["VEL"], p["DIR"], p["TEMP"], p["HUM"], p["PRECIP"], p["FUENTE"], p["TIPO"], p["BARRIDO"], p["LIMITE"], p["CUMPLE"]]
                for col, v in enumerate(vals, start=1):
                    cell = ws.cell(row=fila, column=col, value=v)
                    cell.border = border
                    cell.alignment = center
                fila += 1

            # Anchos
            widths = [12,10,10,10,10,12,12,8,8,10,15,12,10,10,12]
            for i, w in enumerate(widths, start=1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

            out = io.BytesIO()
            wb.save(out)

            st.download_button(
                "📥 DESCARGAR EXCEL ORIGINAL COMPLETO",
                out.getvalue(),
                file_name=f"R2-POE37-EP_V04_{df.iloc[0]['MUNICIPIO']}_{df.iloc[0]['PUNTO']}_RES627.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            st.success("Listo: Formato idéntico a tu foto + todos los datos originales + RES 627")
