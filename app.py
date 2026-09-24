import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage
import os

st.set_page_config(page_title="EQUISIMA V9 CON CROQUIS", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04 - CON CROQUIS FOTO")

if "puntos" not in st.session_state:
    st.session_state.puntos = []
if "foto_croquis" not in st.session_state:
    st.session_state.foto_croquis = None

NORMA = {
    "A. Tranquilidad y Silencio": {"Diurno": 55, "Nocturno": 45},
    "B. Tranquilidad y Ruido Moderado": {"Diurno": 65, "Nocturno": 50},
    "C. Ruido Intermedio Restringido": {"Diurno": 75, "Nocturno": 70},
    "C. Industrial": {"Diurno": 75, "Nocturno": 75},
    "C. Centro Ciudad": {"Diurno": 70, "Nocturno": 55},
}

tab1, tab2 = st.tabs(["📋 DATOS + CROQUIS", "📥 DESCARGA EXCEL"])

with tab1:
    c1,c2 = st.columns(2)
    with c1:
        cliente = st.text_input("CLIENTE:", "Rueda Inversiones S A S")
        municipio = st.text_input("MUNICIPIO:", "ATACO")
        depto = st.text_input("DEPARTAMENTO:", "TOLIMA")
        punto = st.text_input("PUNTO:", "Punto 1 nocturno")
        coord_n = st.text_input("COORD ORIGEN NACIONAL:", "3°35'11.30\"N 75°23'27.72\"W")
        coord_c12 = st.text_input("COORD CTM12:", "E 876543 N 912345")
    with c2:
        st.subheader("📸 DIBUJE EL PUNTO - CROQUIS")
        foto = st.file_uploader("Sube foto/croquis del punto (donde dice Dibuje el punto)", type=["jpg","jpeg","png"])
        if foto:
            st.image(foto, caption="Croquis a insertar en formato", width=300)
            st.session_state.foto_croquis = foto
        else:
            st.info("Aquí va el dibujo/croquis/foto del punto como en tu formato original")

        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 2)
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 1)
        laeq = st.number_input("LAeq,T dB", 0.0, 140.0, 65.0)
        fecha = st.date_input("FECHA", datetime.now())

    if st.button("📍 AGREGAR PUNTO CON CROQUIS", type="primary", use_container_width=True):
        limite = NORMA[sector][periodo]
        cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"
        st.session_state.puntos.append({
            "CLIENTE": cliente, "MUNICIPIO": municipio, "DEPTO": depto, "PUNTO": punto,
            "COORD_N": coord_n, "COORD_C12": coord_c12,
            "FECHA": str(fecha), "LAEQ": laeq, "LIMITE": limite, "CUMPLE": cumple,
            "SECTOR": sector, "PERIODO": periodo
        })
        st.success(f"Guardado - {cumple}")

with tab2:
    if not st.session_state.puntos:
        st.warning("Agrega puntos")
    else:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df, use_container_width=True)

        if st.button("📥 GENERAR EXCEL CON CROQUIS", type="primary", use_container_width=True):
            wb = Workbook()
            ws = wb.active
            ws.title = "Datos de Campo Emision"

            bold = Font(bold=True, size=11)
            bold12 = Font(bold=True, size=12)
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Alignment(horizontal="center", vertical="center", wrap_text=True)

            ws.merge_cells("A1:H1")
            ws["A1"] = "FORMATO R2-POE37-EP V04 - DATOS DE CAMPO EMISIÓN"
            ws["A1"].font = bold12; ws["A1"].alignment = center

            p0 = df.iloc[0]
            ws["A3"] = "CLIENTE:"; ws["A3"].font = bold
            ws.merge_cells("B3:H3"); ws["B3"] = p0["CLIENTE"]
            ws["A4"] = "MUNICIPIO:"; ws["A4"].font = bold; ws["B4"] = p0["MUNICIPIO"]
            ws["C4"] = "DEPTO:"; ws["C4"].font = bold; ws["D4"] = p0["DEPTO"]
            ws["A5"] = "PUNTO:"; ws["A5"].font = bold; ws.merge_cells("B5:H5"); ws["B5"] = p0["PUNTO"]
            ws["A6"] = "COORD N:"; ws["A6"].font = bold; ws.merge_cells("B6:H6"); ws["B6"] = p0["COORD_N"]
            ws["A7"] = "COORD C12:"; ws["A7"].font = bold; ws.merge_cells("B7:H7"); ws["B7"] = p0["COORD_C12"]

            # AQUI VA EL CROQUIS - COMO EN TU FORMATO ORIGINAL
            ws.merge_cells("A9:H9")
            ws["A9"] = "DIBUJE EL PUNTO DE MONITOREO / CROQUIS / FOTO:"
            ws["A9"].font = bold; ws["A9"].alignment = center

            # Reservar espacio para foto
            ws.merge_cells("A10:H25")
            ws["A10"].alignment = center
            ws.row_dimensions[10].height = 300
            for i in range(10, 26):
                ws.row_dimensions[i].height = 20

            # Insertar foto si existe
            if st.session_state.foto_croquis:
                try:
                    # Guardar temporal
                    pil_img = PILImage.open(st.session_state.foto_croquis)
                    # Redimensionar
                    pil_img.thumbnail((600, 400))
                    temp_path = "/tmp/croquis_temp.png"
                    pil_img.save(temp_path)

                    img = ExcelImage(temp_path)
                    img.anchor = "A10"
                    img.width = 400
                    img.height = 300
                    ws.add_image(img)
                except Exception as e:
                    ws["A10"] = f"CROQUIS: {st.session_state.foto_croquis.name} - Ver archivo adjunto"
            else:
                ws["A10"] = "ESPACIO PARA DIBUJO / CROQUIS DEL PUNTO"

            ws.merge_cells("A27:H27")
            ws["A27"] = f"EVALUACIÓN RES 627 - {p0['SECTOR']} {p0['PERIODO']} LIMITE {p0['LIMITE']} dB - {p0['CUMPLE']}"
            ws["A27"].font = bold; ws["A27"].alignment = center

            # Tabla
            headers = ["FECHA","PUNTO","LAEQ","LIMITE","CUMPLE","COORD_N","COORD_C12"]
            for i, h in enumerate(headers, start=1):
                c = ws.cell(row=28, column=i, value=h)
                c.font = bold; c.alignment = center; c.border = border

            fila = 29
            for _, p in df.iterrows():
                vals = [p["FECHA"], p["PUNTO"], p["LAEQ"], p["LIMITE"], p["CUMPLE"], p["COORD_N"], p["COORD_C12"]]
                for col, v in enumerate(vals, start=1):
                    cell = ws.cell(row=fila, column=col, value=v)
                    cell.border = border; cell.alignment = center
                fila += 1

            for i in range(1, 8):
                ws.column_dimensions[get_column_letter(i)].width = 18

            out = io.BytesIO()
            wb.save(out)

            st.download_button(
                "📥 DESCARGAR EXCEL CON CROQUIS/FOTO",
                out.getvalue(),
                file_name=f"R2-POE37-EP_V04_{p0['MUNICIPIO']}_{p0['PUNTO']}_CON_CROQUIS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            st.success("Listo! Excel con espacio DIBUJE EL PUNTO y tu foto incrustada")
