import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os

st.set_page_config(page_title="EQUISIMA V13.7 NO BORRA TITULO", layout="wide")
st.title("EQUISIMA - V13.7 FIX TITULO GRIS")

if "puntos" not in st.session_state:
    st.session_state.puntos = []
if "fotos_puntos" not in st.session_state:
    st.session_state.fotos_puntos = {}

NORMA = {
    "A. Tranquilidad y Silencio": {"Diurno": 55, "Nocturno": 45},
    "B. Tranquilidad y Ruido Moderado": {"Diurno": 65, "Nocturno": 50},
    "C. Ruido Intermedio Restringido": {"Diurno": 75, "Nocturno": 70},
    "C. Industrial": {"Diurno": 75, "Nocturno": 75},
    "C. Centro Ciudad": {"Diurno": 70, "Nocturno": 55},
}

def calcular_leq(vals):
    vals = np.array(vals)
    vals = vals[vals > 0]
    return 10*np.log10(np.mean(10**(vals/10)))

def crear_excel_con_plantilla_oficial(puntos, foto_mapa=None, plantilla_path="plantilla.xlsx", escenario=None):
    from openpyxl.cell.cell import MergedCell
    from openpyxl.utils import get_column_letter
    def safe_set(ws, r, c, val):
        try:
            cell = ws.cell(row=r, column=c)
            if isinstance(cell, MergedCell):
                for mr in ws.merged_cells.ranges:
                    if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                        ws.cell(row=mr.min_row, column=mr.min_col).value = val
                        return
            else:
                cell.value = val
        except:
            pass

    if os.path.exists(plantilla_path):
        wb = load_workbook(plantilla_path)
        ws = wb["Datos de Campo Emision"] if "Datos de Campo Emision" in wb.sheetnames else wb.active
        p0 = puntos[0] if puntos else {}

        # Encabezado
        safe_set(ws, 8, 5, p0.get("CLIENTE",""))
        safe_set(ws, 9, 5, p0.get("PROYECTO",""))
        safe_set(ws, 10, 5, p0.get("DEPTO",""))
        safe_set(ws, 11, 5, p0.get("MUNICIPIO",""))
        safe_set(ws, 12, 5, p0.get("FUENTE",""))

        # --- FIX DEFINITIVO NO TOCA TITULO GRIS B-E ---
        if len(puntos) >= 1:
            pt1 = puntos[0]
            safe_set(ws, 15, 4, pt1.get("PUNTO","")) # D15
            safe_set(ws, 15, 9, pt1.get("COORD_N","")) # I15
            safe_set(ws, 16, 6, pt1.get("DESC","")) # F16 - DESPUES del merge B16:E16
            safe_set(ws, 17, 6, pt1.get("BARRIDO_DB","")) # F17 - DESPUES del merge B17:E17
            safe_set(ws, 19, 2, str(pt1.get("FECHA","")))
            safe_set(ws, 19, 3, str(pt1.get("HORA","")))
            safe_set(ws, 19, 4, f"Ini {pt1.get('CALIB','114')}")
            safe_set(ws, 20, 4, f"Fin {pt1.get('CALIB','114')}")
            safe_set(ws, 19, 5, str(pt1.get("MEMORIA","")))
            safe_set(ws, 19, 6, pt1.get("LAEQ",""))
            safe_set(ws, 19, 7, pt1.get("VEL",""))
            safe_set(ws, 19, 8, pt1.get("DIR",""))
            safe_set(ws, 19, 9, pt1.get("TEMP",""))
            safe_set(ws, 19, 10, pt1.get("HUM",""))
            safe_set(ws, 19, 11, pt1.get("PRECIP",""))
            safe_set(ws, 19, 12, pt1.get("FUENTE",""))
            safe_set(ws, 19, 13, pt1.get("TIPO_RUIDO",""))
            safe_set(ws, 19, 14, pt1.get("TIEMPO_OP",""))
            safe_set(ws, 24, 2, pt1.get("ESQUEMA",""))

        if len(puntos) >= 2:
            pt2 = puntos[1]
            safe_set(ws, 27, 4, pt2.get("PUNTO","")) # D27
            safe_set(ws, 27, 9, pt2.get("COORD_N","")) # I27
            safe_set(ws, 28, 6, pt2.get("DESC","")) # F28
            safe_set(ws, 29, 6, pt2.get("BARRIDO_DB","")) # F29
            safe_set(ws, 31, 2, str(pt2.get("FECHA","")))
            safe_set(ws, 31, 3, str(pt2.get("HORA","")))
            safe_set(ws, 31, 4, f"Ini {pt2.get('CALIB','114')}")
            safe_set(ws, 32, 4, f"Fin {pt2.get('CALIB','114')}")
            safe_set(ws, 31, 5, str(pt2.get("MEMORIA","")))
            safe_set(ws, 31, 6, pt2.get("LAEQ",""))
            safe_set(ws, 31, 7, pt2.get("VEL",""))
            safe_set(ws, 31, 8, pt2.get("DIR",""))
            safe_set(ws, 31, 9, pt2.get("TEMP",""))
            safe_set(ws, 31, 10, pt2.get("HUM",""))
            safe_set(ws, 31, 11, pt2.get("PRECIP",""))
            safe_set(ws, 31, 12, pt2.get("FUENTE",""))
            safe_set(ws, 31, 13, pt2.get("TIPO_RUIDO",""))
            safe_set(ws, 31, 14, pt2.get("TIEMPO_OP",""))
            safe_set(ws, 36, 2, pt2.get("ESQUEMA",""))

        if escenario:
            q1 = f"Diurno: {'X' if escenario['diurno'] else ' '} Nocturno: {'X' if escenario['nocturno'] else ' '} Fuente Encendida: {'X' if escenario['encendida'] else ' '} Apagada: {'X' if escenario['apagada'] else ' '} "
            q2 = f"Sí: {'X' if escenario['residual']=='Sí' else ' '} No: {'X' if escenario['residual']=='No' else ' '} Just: {escenario['justificacion']}"
            q3 = f"Sí: {'X' if escenario['barrido']=='Sí' else ' '} No: {'X' if escenario['barrido']=='No' else ' '} "
            safe_set(ws, 15, 17, q1); safe_set(ws, 16, 17, q2); safe_set(ws, 17, 17, q3)
            safe_set(ws, 27, 17, q1); safe_set(ws, 28, 17, q2); safe_set(ws, 29, 17, q3)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out
    else:
        from openpyxl import Workbook
        wb = Workbook(); ws = wb.active; ws["A1"]="Sube plantilla.xlsx"
        out = io.BytesIO(); wb.save(out); out.seek(0); return out

tab1, tab2 = st.tabs(["📋 FORMATO", "📸 FOTOS"])

with tab1:
    if not os.path.exists("plantilla.xlsx"):
        st.warning("Sube plantilla.xlsx")
    else:
        st.success("✅ plantilla.xlsx OK - FIX TITULO GRIS")

    if st.button("🗑️ LIMPIAR TODO", use_container_width=True):
        st.session_state.puntos = []
        st.session_state.fotos_puntos = {}
        st.rerun()

    c1,c2 = st.columns(2)
    with c1:
        cliente = st.text_input("CLIENTE", "Molinos el yopal")
        municipio = st.text_input("MUNICIPIO", "YOPAL")
        depto = st.text_input("DEPARTAMENTO", "Casanare")
        punto = st.text_input("PUNTO No", "Punto 1")
        coord_n = st.text_input("COORD", "5°21'N 72°23'W")
        desc = st.text_area("DESCRIPCIÓN DEL PUNTO (va en F16, no borra titulo)", "ubicado al costado norte de la planta en porteria principal se evidencia alto flujo entrada y salida de vehiculos de carga", height=90)
        barrido_db = st.text_input("BARRIDO PERIMETRAL dB (va en F17)", "68.5 / 70.2 / 69.1")
        fuente = st.text_input("Fuente", "MOLIENDA")
    with c2:
        proyecto = st.text_input("PROYECTO", "Molinos el yopal")
        fecha = st.date_input("FECHA", datetime.now())
        hora = st.text_input("HORA INICIO", "17:00")
        hora_fin = st.text_input("HORA FIN", "17:15")
        calib = st.text_input("Calib", "114.0")
        memoria = st.text_input("Memoria", "1")
        laeq = st.number_input("LAeq", 0.0, 140.0, 73.0)
        vel = st.number_input("Viento", 0.0, 20.0, 0.3)
        dir_v = st.text_input("Dir", "N")
        temp = st.number_input("Temp", -10.0, 60.0, 32.0)
        hum = st.number_input("Hum", 0.0, 100.0, 68.0)
        precip = st.selectbox("Precip?", ["No","Sí"],0)
        tipo = st.selectbox("Tipo", ["Continuo","Intermitente"],0)
        sector = st.selectbox("SECTOR", list(NORMA.keys()),2)
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"],0)
        esquema_txt = st.text_area("ESQUEMA", "Norte: vía")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "PROYECTO": proyecto, "MUNICIPIO": municipio, "DEPTO": depto,
            "PUNTO": punto, "COORD_N": coord_n, "DESC": desc, "BARRIDO_DB": barrido_db,
            "FECHA": str(fecha), "HORA": f"{hora}-{hora_fin}", "CALIB": calib, "MEMORIA": memoria,
            "LAEQ": laeq, "VEL": vel, "DIR": dir_v, "TEMP": temp, "HUM": hum, "PRECIP": precip,
            "FUENTE": fuente, "TIPO_RUIDO": tipo, "TIEMPO_OP": "60 min", "ESQUEMA": esquema_txt
        })
        st.success(f"Agregado {punto}")

    if st.session_state.puntos:
        df = pd.DataFrame(st.session_state.puntos)
        st.dataframe(df[["PUNTO","DESC","BARRIDO_DB","LAEQ"]], use_container_width=True)
        if st.button("📥 GENERAR EXCEL", type="primary", use_container_width=True):
            escenario_data = {"diurno": periodo=="Diurno", "nocturno": periodo=="Nocturno", "encendida": True, "apagada": False, "residual": "Sí", "justificacion": "", "barrido": "Sí"}
            excel_file = crear_excel_con_plantilla_oficial(st.session_state.puntos, None, "plantilla.xlsx", escenario_data)
            st.download_button("📥 DESCARGAR", excel_file.getvalue(), f"R2_FIX_TITULO_{municipio}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

with tab2:
    st.write("Fotos por punto")
    if st.session_state.puntos:
        sel = st.selectbox("Punto", [p["PUNTO"] for p in st.session_state.puntos])
        fotos = st.file_uploader("Fotos", type=["jpg","png","jpeg"], accept_multiple_files=True, key=sel)
        if fotos:
            st.session_state.fotos_puntos[sel] = fotos
            st.success(f"{len(fotos)} fotos")
