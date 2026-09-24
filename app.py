import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font

st.set_page_config(layout="wide")
st.title("EQUISIMA - FORMATO ORIGINAL SIN ERRORES")

if "puntos" not in st.session_state:
    st.session_state.puntos = []

tab1, tab2 = st.tabs(["📋 DATOS", "📥 DESCARGA"])

with tab1:
    c1,c2 = st.columns(2)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones")
        municipio = st.text_input("MUNICIPIO", "ATACO")
        punto = st.text_input("PUNTO", "Punto 1")
        coord1 = st.text_input("COORDENADAS ORIGEN NACIONAL", "3°35'11N 75°23'27W")
        coord2 = st.text_input("COORDENADAS CTM12 - ESTE QUE FALTABA", "E 876543 N 912345")
    with c2:
        laeq = st.number_input("LAeq dB", 0.0, 140.0, 65.0)
        fecha = st.date_input("Fecha", datetime.now())
        hora = st.text_input("Hora", "21:01")

    if st.button("AGREGAR", type="primary", use_container_width=True):
        st.session_state.puntos.append({"CLIENTE":cliente,"MUNICIPIO":municipio,"PUNTO":punto,"COORD_NAC":coord1,"COORD_CTM12":coord2,"LAEQ":laeq,"FECHA":str(fecha),"HORA":hora})
        st.success("Guardado")

with tab2:
    if st.session_state.puntos:
        st.dataframe(pd.DataFrame(st.session_state.puntos))
        if st.button("GENERAR EXCEL EN FORMATO ORIGINAL", type="primary", use_container_width=True):
            wb = Workbook()
            ws = wb.active
            ws.title = "Datos de Campo"
            ws["A1"] = "R2-POE37-EP V04"
            ws["A1"].font = Font(bold=True)
            p0 = st.session_state.puntos[0]
            ws["A3"] = "CLIENTE:"; ws["B3"] = p0["CLIENTE"]
            ws["A4"] = "MUNICIPIO:"; ws["B4"] = p0["MUNICIPIO"]
            ws["A5"] = "PUNTO:"; ws["B5"] = p0["PUNTO"]
            ws["A6"] = "COORD NACIONAL:"; ws["B6"] = p0["COORD_NAC"]
            ws["A7"] = "COORD CTM12:"; ws["B7"] = p0["COORD_CTM12"]
            ws["A9"] = "FECHA"; ws["B9"] = "HORA"; ws["C9"] = "LAEQ"
            f=10
            for p in st.session_state.puntos:
                ws.cell(row=f,column=1).value=p["FECHA"]
                ws.cell(row=f,column=2).value=p["HORA"]
                ws.cell(row=f,column=3).value=p["LAEQ"]
                ws.cell(row=f,column=4).value=p["COORD_NAC"]
                ws.cell(row=f,column=5).value=p["COORD_CTM12"]
                f+=1
            out = io.BytesIO()
            wb.save(out)
            st.download_button("📥 DESCARGAR AHORA", out.getvalue(), "FORMATO_ORIGINAL.xlsx", use_container_width=True, type="primary")
    else:
        st.warning("Agrega datos en pestaña 1")
