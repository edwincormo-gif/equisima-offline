import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="EQUISIMA V6 FORMATO", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04 - FORMATO ORIGINAL BONITO")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

NORMA = {
    "A. Tranquilidad y Silencio": {"Diurno": 55, "Nocturno": 45},
    "B. Tranquilidad y Ruido Moderado": {"Diurno": 65, "Nocturno": 50},
    "C. Ruido Intermedio Restringido": {"Diurno": 75, "Nocturno": 70},
    "C. Industrial": {"Diurno": 75, "Nocturno": 75},
    "C. Centro Ciudad": {"Diurno": 70, "Nocturno": 55},
}

tab1, tab2 = st.tabs(["📋 DATOS", "📥 DESCARGA FORMATO"])

with tab1:
    c1,c2 = st.columns(2)
    with c1:
        cliente = st.text_input("CLIENTE:", "Rueda Inversiones S A S")
        municipio = st.text_input("MUNICIPIO:", "ATACO")
        depto = st.text_input("DEPARTAMENTO:", "TOLIMA")
        punto = st.text_input("PUNTO DE MONITOREO:", "Punto 1")
        coord_n = st.text_input("COORD ORIGEN NACIONAL:", "3°35'11N 75°23'27W")
        coord_ctm = st.text_input("COORD CTM12:", "E 876543 N 912345")
    with c2:
        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 1)
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 1)
        laeq = st.number_input("LAeq,T (dB)", 0.0, 140.0, 65.0)
        fecha = st.date_input("FECHA", datetime.now())
        hora = st.text_input("HORA", "21:01")
        fuente = st.text_input("FUENTE", "mineria")

    limite = NORMA[sector][periodo]
    cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"
    st.info(f"RES 627: Límite {limite} dB | Medido {laeq} dB | **{cumple}**")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "MUNICIPIO": municipio, "DEPTO": depto, "PUNTO": punto,
            "COORD_NACIONAL": coord_n, "COORD_CTM12": coord_ctm,
            "SECTOR": sector, "PERIODO": periodo, "LAEQ": laeq, "LIMITE": limite, "CUMPLE": cumple,
            "FECHA": str(fecha), "HORA": hora, "FUENTE": fuente
        })
        st.success(f"Guardado - {cumple}")
        st.balloons()

with tab2:
    if not st.session_state.puntos:
        st.warning("Agrega puntos en pestaña 1")
    else:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df, use_container_width=True)

        if st.button("📥 GENERAR EXCEL EN FORMATO ORIGINAL BONITO", type="primary", use_container_width=True):
            wb = Workbook()
            ws = wb.active
            ws.title = "Datos de Campo Emision"

            # Estilos
            bold = Font(bold=True, size=11)
            bold12 = Font(bold=True, size=12)
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Alignment(horizontal="center", vertical="center")
            left = Alignment(horizontal="left", vertical="center")

            # TÍTULO
            ws.merge_cells("A1:F1")
            ws["A1"] = "R2-POE37-EP V04"
            ws["A1"].font = bold12
            ws["A1"].alignment = center

            # DATOS GENERALES - Como tu plantilla original
            ws["A3"] = "CLIENTE:"; ws["A3"].font = bold
            ws.merge_cells("B3:F3")
            ws["B3"] = df.iloc[0]["CLIENTE"]

            ws["A4"] = "MUNICIPIO:"; ws["A4"].font = bold
            ws.merge_cells("B4:D4")
            ws["B4"] = df.iloc[0]["MUNICIPIO"]
            ws["E4"] = "DEPARTAMENTO:"; ws["E4"].font = bold
            ws["F4"] = df.iloc[0]["DEPTO"]

            ws["A5"] = "PUNTO:"; ws["A5"].font = bold
            ws.merge_cells("B5:F5")
            ws["B5"] = df.iloc[0]["PUNTO"]

            ws["A6"] = "COORD N:"; ws["A6"].font = bold
            ws.merge_cells("B6:F6")
            ws["B6"] = df.iloc[0]["COORD_NACIONAL"]

            ws["A7"] = "COORD C12:"; ws["A7"].font = bold
            ws.merge_cells("B7:F7")
            ws["B7"] = df.iloc[0]["COORD_CTM12"]

            # TABLA
            ws["A9"] = "EVALUACIÓN RES 627 - SECTOR: " + df.iloc[0]["SECTOR"] + " " + df.iloc[0]["PERIODO"] + f" - LIMITE {limite} dB"
            ws["A9"].font = bold
            ws.merge_cells("A9:F9")

            headers = ["FECHA", "HORA", "LAEQ", "LIMITE", "CUMPLE", "FUENTE"]
            for i, h in enumerate(headers, start=1):
                c = ws.cell(row=10, column=i)
                c.value = h
                c.font = bold
                c.alignment = center
                c.border = border

            fila = 11
            for _, p in df.iterrows():
                ws.cell(row=fila, column=1, value=p["FECHA"]).border = border
                ws.cell(row=fila, column=2, value=p["HORA"]).border = border
                ws.cell(row=fila, column=3, value=p["LAEQ"]).border = border
                ws.cell(row=fila, column=4, value=p["LIMITE"]).border = border
                ws.cell(row=fila, column=5, value=p["CUMPLE"]).border = border
                ws.cell(row=fila, column=6, value=p["FUENTE"]).border = border
                ws.cell(row=fila, column=1).alignment = center
                ws.cell(row=fila, column=2).alignment = center
                ws.cell(row=fila, column=3).alignment = center
                ws.cell(row=fila, column=4).alignment = center
                ws.cell(row=fila, column=5).alignment = center
                fila += 1

            # Anchos
            ws.column_dimensions["A"].width = 18
            ws.column_dimensions["B"].width = 18
            ws.column_dimensions["C"].width = 12
            ws.column_dimensions["D"].width = 12
            ws.column_dimensions["E"].width = 15
            ws.column_dimensions["F"].width = 20

            out = io.BytesIO()
            wb.save(out)

            st.download_button(
                "📥 DESCARGAR EXCEL BONITO FORMATO ORIGINAL",
                out.getvalue(),
                file_name=f"R2-POE37-EP_V04_{df.iloc[0]['MUNICIPIO']}_{df.iloc[0]['PUNTO']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            st.success("¡Listo! Este es el formato bonito, ya no da error y trae las 2 coordenadas.")
