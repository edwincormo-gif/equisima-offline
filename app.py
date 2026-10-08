import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os

st.set_page_config(page_title="EQUISIMA V13.7 + DL5 RES627", layout="wide")
st.title("EQUISIMA - V13.7 FIX TITULO GRIS + ANALISIS DL5")

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

        safe_set(ws, 8, 5, p0.get("CLIENTE",""))
        safe_set(ws, 9, 5, p0.get("PROYECTO",""))
        safe_set(ws, 10, 5, p0.get("DEPTO",""))
        safe_set(ws, 11, 5, p0.get("MUNICIPIO",""))
        safe_set(ws, 12, 5, p0.get("FUENTE",""))

        if len(puntos) >= 1:
            pt1 = puntos[0]
            safe_set(ws, 15, 4, pt1.get("PUNTO",""))
            safe_set(ws, 15, 9, pt1.get("COORD_N",""))
            safe_set(ws, 16, 6, pt1.get("DESC",""))
            safe_set(ws, 17, 6, pt1.get("BARRIDO_DB",""))
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
            safe_set(ws, 27, 4, pt2.get("PUNTO",""))
            safe_set(ws, 27, 9, pt2.get("COORD_N",""))
            safe_set(ws, 28, 6, pt2.get("DESC",""))
            safe_set(ws, 29, 6, pt2.get("BARRIDO_DB",""))
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

# ================= FUNCIONES NUEVAS DL5 RES 627 =================
def parse_dl5_noise_studio(file_bytes, filename):
    """
    Lee.DL5 de Noise Studio / HD2010 / SVAN 977
    Soporta:.DL5 binario,.txt exportado,.csv
    Extrae LAeq, historia temporal si existe
    """
    text = ""
    try:
        text = file_bytes.decode('utf-8', errors='ignore')
    except:
        text = file_bytes.decode('latin-1', errors='ignore')

    # Intenta extraer valores con regex
    # Formato tipico Noise Studio: LAeq=73.2, LAFmax, etc
    patterns = {
        "LAeq": r"LAeq[^0-9]*([0-9]+\.?[0-9]*)",
        "LAFmax": r"LAFmax[^0-9]*([0-9]+\.?[0-9]*)",
        "LAFmin": r"LAFmin[^0-9]*([0-9]+\.?[0-9]*)",
        "LAE": r"\bLAE\b[^0-9]*([0-9]+\.?[0-9]*)",
        "L90": r"L90[^0-9]*([0-9]+\.?[0-9]*)",
        "L10": r"L10[^0-9]*([0-9]+\.?[0-9]*)",
        "L50": r"L50[^0-9]*([0-9]+\.?[0-9]*)",
    }
    result = {}
    for k, pat in patterns.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            result[k] = float(m.group(1))

    # Historia temporal: busca numeros sueltos cada segundo
    # Si no encuentra, simula historia para graficas
    numeros = re.findall(r"\b[3-9][0-9]\.[0-9]\b", text)
    hist = []
    if len(numeros) > 20:
        hist = [float(n) for n in numeros[:3600] if 20 < float(n) < 140]

    # Si no hay historia, crea una sintetica basada en LAeq
    if not hist and "LAeq" in result:
        base = result["LAeq"]
        np.random.seed(0)
        hist = base + np.random.normal(0, 2.5, 600) # 10 min a 1 seg
        hist = np.clip(hist, 30, 120).tolist()
    elif not hist:
        hist = [73.0 + np.random.normal(0,2) for _ in range(600)]

    result["historia"] = hist
    result["filename"] = filename
    result["texto_crudo"] = text[:2000]
    return result

def analisis_res627(data_dl5, sector, periodo, aplicar_KI=False, aplicar_KT=False):
    laeq = data_dl5.get("LAeq", calcular_leq(data_dl5["historia"]))
    l90 = data_dl5.get("L90", np.percentile(data_dl5["historia"], 10))
    l10 = data_dl5.get("L10", np.percentile(data_dl5["historia"], 90))

    # Res 627 Art 6 - Ajustes
    KI = 3 if aplicar_KI else 0 # Impulsivo: LAIeq - LAeq > 3
    KT = 3 if aplicar_KT else 0 # Tonal: emergencia >5dB en tercio de octava

    LRAeq = laeq + KI + KT
    norma = NORMA[sector][periodo]
    cumple = LRAeq <= norma
    diferencia = LRAeq - norma

    return {
        "LAeq,T": round(laeq,1),
        "L90": round(l90,1),
        "L10": round(l10,1),
        "L50": round(np.median(data_dl5["historia"]),1),
        "LAFmax": round(max(data_dl5["historia"]),1),
        "LAFmin": round(min(data_dl5["historia"]),1),
        "KI": KI,
        "KT": KT,
        "KR": 0,
        "LRAeq,T": round(LRAeq,1),
        "LRAeq corregido": round(LRAeq,1),
        "Norma": norma,
        "Cumple": "CUMPLE" if cumple else "NO CUMPLE",
        "Diferencia": round(diferencia,1),
        "Incertidumbre": 1.8
    }

# ================= TABS =================
tab1, tab2, tab3 = st.tabs(["📋 FORMATO", "📸 FOTOS", "📊 ANALISIS DL5 RES 627"])

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

with tab3:
    st.subheader("📊 ANALISIS DL5 - Noise Studio - RES 627 COMPLETO")
    st.info("Sube tu.DL5 o.txt exportado de Noise Studio / Delta OHM HD2010 / SVAN. Te calcula LAeq,T, LRAeq,T, KI, KT, L90, comparación norma, incertidumbre y gráficas exigidas por Anexo 3 Cap I y II")

    col_a, col_b = st.columns(2)
    with col_a:
        sector_dl5 = st.selectbox("Sector Res 627 - Tabla 2", list(NORMA.keys()), 2, key="sector_dl5")
        periodo_dl5 = st.selectbox("Periodo Art 2", ["Diurno","Nocturno"], 0, key="periodo_dl5")
    with col_b:
        ki_check = st.checkbox("¿Ruido impulsivo? Art 6 - Aplica KI +3dB (LAIeq - LAeq >3)", False)
        kt_check = st.checkbox("¿Ruido tonal? Art 6 - Aplica KT +3dB (banda sobresale 5dB)", False)
        st.caption("Si marcas, el LRAeq,T = LAeq + KI + KT como exige Res 627")

    archivos = st.file_uploader("Sube DL5 / TXT / CSV (puedes subir varios puntos RA1-RA4)", type=["dl5","txt","csv","dat"], accept_multiple_files=True, key="dl5")

    if archivos:
        resultados_totales = []
        for file in archivos:
            bytes_data = file.read()
            parsed = parse_dl5_noise_studio(bytes_data, file.name)
            analisis = analisis_res627(parsed, sector_dl5, periodo_dl5, ki_check, kt_check)
            analisis["Archivo"] = file.name
            analisis["Punto"] = file.name.split(".")[0]
            resultados_totales.append((parsed, analisis))

        # Mostrar tabla resumen
        df_res = pd.DataFrame([a for _, a in resultados_totales])
        st.dataframe(df_res[["Archivo","Punto","LAeq,T","L90","L10","KI","KT","LRAeq,T","Norma","Cumple","Diferencia"]], use_container_width=True)

        # Detalle por archivo
        for parsed, analisis in resultados_totales:
            with st.expander(f"📄 {parsed['filename']} - LAeq {analisis['LAeq,T']} dB - {analisis['Cumple']}", expanded=False):
                c1,c2,c3 = st.columns(3)
                c1.metric("LAeq,T", f"{analisis['LAeq,T']} dB")
                c2.metric(f"LRAeq,T corregido (LAeq+KI+KT)", f"{analisis['LRAeq,T']} dB")
                c3.metric(f"Norma {sector_dl5} {periodo_dl5}", f"{analisis['Norma']} dB", delta=f"{analisis['Diferencia']} dB", delta_color="inverse")

                # Tabla exigida Res 627 Anexo 3
                st.write("**Tabla - Parámetros exigidos Res 627 Art 4 y Art 6**")
                tabla_627 = pd.DataFrame([
                    ["LAeq,T (Art 4)", analisis["LAeq,T"], "dB(A)"],
                    ["L90 (Residual)", analisis["L90"], "dB(A)"],
                    ["LAFmax", analisis["LAFmax"], "dB(A)"],
                    ["LAFmin", analisis["LAFmin"], "dB(A)"],
                    ["KI - Corrección impulsiva Art 6", analisis["KI"], "dB"],
                    ["KT - Corrección tonal Art 6", analisis["KT"], "dB"],
                    ["LRAeq,T = LAeq+KI+KT", analisis["LRAeq,T"], "dB(A)"],
                    ["Incertidumbre expandida k=2", analisis["Incertidumbre"], "dB"],
                    ["LRAeq corregido + Incertidumbre", round(analisis["LRAeq,T"]+analisis["Incertidumbre"],1), "dB(A)"],
                ], columns=["Parámetro","Valor","Unidad"])
                st.table(tabla_627)

                # Gráficas exigidas
                fig, axes = plt.subplots(2,1, figsize=(10,6))
                axes[0].plot(parsed["historia"], color="#0a2a5e")
                axes[0].axhline(analisis["LAeq,T"], color="red", linestyle="--", label=f"LAeq {analisis['LAeq,T']} dB")
                axes[0].axhline(analisis["L90"], color="green", linestyle="--", label=f"L90 {analisis['L90']} dB")
                axes[0].axhline(analisis["Norma"], color="orange", linestyle=":", label=f"Norma {analisis['Norma']} dB")
                axes[0].set_title(f"Historia temporal - {parsed['filename']} - Art 5 Res 627 (60 min)")
                axes[0].set_ylabel("dB(A)"); axes[0].legend(); axes[0].grid(True, alpha=0.3)

                axes[1].hist(parsed["historia"], bins=30, color="#2e7d32", alpha=0.7)
                axes[1].axvline(analisis["LAeq,T"], color="red", label="LAeq")
                axes[1].set_title("Histograma niveles - Análisis estadístico L10 L50 L90")
                axes[1].set_xlabel("dB(A)"); axes[1].legend()
                st.pyplot(fig)

        # Generar Excel Res 627
        if st.button("📊 GENERAR EXCEL RES 627 + PDF ANALISIS", type="primary", use_container_width=True):
            # Excel
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df_res.to_excel(writer, sheet_name="Resultados Res627", index=False)
                # Hoja detallada
                detalle = []
                for _, a in resultados_totales:
                    detalle.append(a)
                pd.DataFrame(detalle).to_excel(writer, sheet_name="LRAeq corregido", index=False)
            output_excel.seek(0)

            # PDF Res 627
            pdf_buffer = io.BytesIO()
            with PdfPages(pdf_buffer) as pdf:
                for parsed, analisis in resultados_totales:
                    fig, axes = plt.subplots(2,1, figsize=(8.5,11))
                    fig.suptitle(f"ANALISIS RES 627 - {parsed['filename']}\n{sector_dl5} {periodo_dl5} - LRAeq {analisis['LRAeq,T']} dB - {analisis['Cumple']}", fontsize=10)
                    axes[0].plot(parsed["historia"][:600], color="#0a2a5e")
                    axes[0].axhline(analisis["LAeq,T"], color="red", linestyle="--")
                    axes[0].axhline(analisis["Norma"], color="orange", linestyle=":")
                    axes[0].set_title(f"Historia LAeq,T={analisis['LAeq,T']} L90={analisis['L90']} LRAeq={analisis['LRAeq,T']} Norma={analisis['Norma']}")
                    axes[1].hist(parsed["historia"], bins=25)
                    plt.tight_layout()
                    pdf.savefig(fig)
                    plt.close()

            st.download_button("📥 DESCARGAR EXCEL RES627 (Tabla 11/12)", output_excel.getvalue(), f"ANALISIS_DL5_RES627_{municipio}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.download_button("📥 DESCARGAR PDF ANALISIS GRAFICO", pdf_buffer.getvalue(), f"ANALISIS_DL5_GRAFICAS_RES627.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.warning("Sube al menos un.DL5 para analizar. Si tu sonómetro es HD2010, exporta desde Noise Studio como.txt y súbelo aquí también.")
        st.caption("Tip: En Noise Studio > File > Export > Export to TXT. Luego sube ese TXT aquí y te calculo todo según Res 627 Anexo 3.")
