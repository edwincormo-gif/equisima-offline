import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="EQUISIMA V7.1 FIX", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04 COMPLETO + RES 627")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

NORMA = {
    "A. Tranquilidad y Silencio": {"Diurno": 55, "Nocturno": 45},
    "B. Tranquilidad y Ruido Moderado": {"Diurno": 65, "Nocturno": 50},
    "C. Ruido Intermedio Restringido": {"Diurno": 75, "Nocturno": 70},
    "C. Industrial": {"Diurno": 75, "Nocturno": 75},
    "C. Centro Ciudad": {"Diurno": 70, "Nocturno": 55},
}

tab1, tab2 = st.tabs(["📋 DATOS DE CAMPO", "📥 DESCARGA"])

with tab1:
    c1,c2 = st.columns(2)
    with c1:
        cliente = st.text_input("CLIENTE:", "Rueda Inversiones S A S")
        municipio = st.text_input("MUNICIPIO:", "ATACO")
        depto = st.text_input("DEPARTAMENTO:", "TOLIMA")
        punto = st.text_input("PUNTO:", "Punto 1")
        coord_n = st.text_input("COORD N:", "3°35'11N 75°23'27W")
        coord_c12 = st.text_input("COORD C12:", "E 876543 N 912345")
    with c2:
        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 2)
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 0)
        laeq = st.number_input("LAeq,T (dB)", 0.0, 140.0, 65.0)
        fecha = st.date_input("FECHA", datetime.now())
        hora = st.text_input("HORA", "21:01")
        fuente = st.text_input("FUENTE", "mineria")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        limite = NORMA[sector][periodo]
        cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"
        st.session_state.puntos.append({
            "CLIENTE": cliente, "MUNICIPIO": municipio, "DEPTO": depto, "PUNTO": punto,
            "COORD_N": coord_n, "COORD_C12": coord_c12,
            "FECHA": str(fecha), "HORA": hora, "LAEQ": laeq,
            "LIMITE": limite, "CUMPLE": cumple, "FUENTE": fuente,
            "SECTOR": sector, "PERIODO": periodo
        })
        st.success(f"Guardado - {cumple}")

with tab2:
    if not st.session_state.puntos:
        st.warning("Agrega puntos")
    else:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df, use_container_width=True)
        if st.button("📥 GENERAR EXCEL FORMATO ORIGINAL", type="primary", use_container_width=True):
            wb = Workbook()
            ws = wb.active
            ws.title = "Datos de Campo"
            bold = Font(bold=True, size=11)
            bold12 = Font(bold=True, size=12)
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Alignment(horizontal="center", vertical="center")

            ws.merge_cells("A1:F1")
            ws["A1"] = "R2-POE37-EP V04"
            ws["A1"].font = bold12
            ws["A1"].alignment = center

            ws["A3"] = "CLIENTE:"; ws["A3"].font = bold
            ws.merge_cells("B3:F3"); ws["B3"] = df.iloc[0]["CLIENTE"]

            ws["A4"] = "MUNICIPIO:"; ws["A4"].font = bold
            ws["B4"] = df.iloc[0]["MUNICIPIO"]
            ws["C4"] = "DEPARTAMENTO"; ws["C4"].font = bold
            ws["D4"] = df.iloc[0]["DEPTO"]

            ws["A5"] = "PUNTO:"; ws["A5"].font = bold
            ws.merge_cells("B5:F5"); ws["B5"] = df.iloc[0]["PUNTO"]

            ws["A6"] = "COORD N:"; ws["A6"].font = bold
            ws.merge_cells("B6:F6"); ws["B6"] = df.iloc[0]["COORD_N"]

            ws["A7"] = "COORD C12:"; ws["A7"].font = bold
            ws.merge_cells("B7:F7"); ws["B7"] = df.iloc[0]["COORD_C12"]

            ws.merge_cells("A9:F9")
            ws["A9"] = f"EVALUACIÓN RES 627 - SECTOR: {df.iloc[0]['SECTOR']} {df.iloc[0]['PERIODO']} - LIMITE {df.iloc[0]['LIMITE']} dB"
            ws["A9"].font = bold; ws["A9"].alignment = center

            headers = ["FECHA","HORA","LAEQ","LIMITE","CUMPLE","FUENTE"]
            for i, h in enumerate(headers, start=1):
                c = ws.cell(row=10, column=i, value=h)
                c.font = bold; c.alignment = center; c.border = border

            fila = 11
            for _, p in df.iterrows():
                for col, val in enumerate([p["FECHA"], p["HORA"], p["LAEQ"], p["LIMITE"], p["CUMPLE"], p["FUENTE"]], start=1):
                    cell = ws.cell(row=fila, column=col, value=val)
                    cell.border = border; cell.alignment = center
                fila += 1

            for i in range(1,7):
                ws.column_dimensions[get_column_letter(i)].width = 16

            out = io.BytesIO()
            wb.save(out)
            st.download_button("📥 DESCARGAR EXCEL", out.getvalue(), f"R2-POE37-EP_{df.iloc[0]['MUNICIPIO']}.xlsx", use_container_width=True, type="primary")
