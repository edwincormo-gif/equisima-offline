import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os

st.set_page_config(page_title="EQUISIMA FINAL V13.3 - FIX MERGED", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04 + PLANTILLA TAL CUAL")

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

def leer_excel_sonometro(file):
    try:
        df = None
        nombre = file.name.lower()
        if nombre.endswith(".csv"):
            try:
                df = pd.read_csv(file, sep=None, engine='python', encoding='latin1')
            except:
                file.seek(0)
                try:
                    df = pd.read_csv(file, sep=";", encoding='latin1')
                except:
                    file.seek(0)
                    df = pd.read_csv(file, encoding='utf-8', sep=None, engine='python')
            hoja = "CSV"
        else:
            xls = pd.ExcelFile(file)
            hoja = xls.sheet_names[0]
            for n in xls.sheet_names:
                if "history" in n.lower() or "time" in n.lower() or "data" in n.lower() or "profile" in n.lower():
                    hoja = n
                    break
            df = pd.read_excel(file, sheet_name=hoja)
        mejor_col = None
        mejor_len = 0
        for col in df.columns:
            try:
                serie = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(serie) > 50 and 20 < serie.mean() < 130:
                    if len(serie) > mejor_len:
                        mejor_len = len(serie)
                        mejor_col = serie
            except:
                continue
        if mejor_col is None:
            for col in df.columns:
                try:
                    serie = df[col].astype(str).str.replace(',', '.').astype(float)
                    serie = pd.to_numeric(serie, errors='coerce').dropna()
                    if len(serie) > 50 and 20 < serie.mean() < 130:
                        mejor_col = serie
                        break
                except:
                    continue
        if mejor_col is not None:
            mejor_col = mejor_col.reset_index(drop=True)
            mejor_col.name = "dB"
        return mejor_col, hoja, df
    except Exception as e:
        return None, str(e), None

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
            try:
                ws[f"{get_column_letter(c)}{r}"] = val
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
        safe_set(ws, 15, 4, p0.get("PUNTO",""))
        safe_set(ws, 16, 2, p0.get("DESC",""))
        safe_set(ws, 25, 2, p0.get("ESQUEMA",""))

        if escenario:
            q15 = f"Diurno: {'X' if escenario['diurno'] else ' '} Nocturno: {'X' if escenario['nocturno'] else ' '} Fuente Encendida: {'X' if escenario['encendida'] else ' '} Fuente Apagada: {'X' if escenario['apagada'] else ' '} "
            q16 = f"Sí: {'X' if escenario['residual']=='Sí' else ' '} No: {'X' if escenario['residual']=='No' else ' '} En caso de no, justique: {escenario['justificacion']}"
            q17 = f"Sí: {'X' if escenario['barrido']=='Sí' else ' '} No: {'X' if escenario['barrido']=='No' else ' '} "
            safe_set(ws, 15, 17, q15)
            safe_set(ws, 16, 17, q16)
            safe_set(ws, 17, 17, q17)
            if len(puntos) > 1:
                safe_set(ws, 27, 17, q15)
                safe_set(ws, 28, 17, q16)
                safe_set(ws, 29, 17, q17)

        fila_inicio = 19
        for idx, pt in enumerate(puntos):
            fila = fila_inicio + idx
            if fila > 22: break
            safe_set(ws, fila, 2, str(pt.get("FECHA","")))
            safe_set(ws, fila, 3, str(pt.get("HORA","")))
            safe_set(ws, fila, 4, f"Ini {pt.get('CALIB','114')} Fin {pt.get('CALIB','114')}")
            safe_set(ws, fila, 5, str(pt.get("MEMORIA","")))
            safe_set(ws, fila, 6, pt.get("LAEQ",""))
            safe_set(ws, fila, 7, pt.get("VEL",""))
            safe_set(ws, fila, 8, pt.get("DIR",""))
            safe_set(ws, fila, 9, pt.get("TEMP",""))
            safe_set(ws, fila, 10, pt.get("HUM",""))
            safe_set(ws, fila, 11, pt.get("PRECIP",""))
            safe_set(ws, fila, 12, pt.get("FUENTE",""))
            safe_set(ws, fila, 13, pt.get("TIPO_RUIDO","Continuo"))
            safe_set(ws, fila, 14, pt.get("TIEMPO_OP","60 min"))

        if foto_mapa:
            try:
                pil_img = PILImage.open(foto_mapa)
                pil_img.thumbnail((900, 500))
                tmp = "/tmp/mapa_esquema.png"
                pil_img.save(tmp)
                img = ExcelImage(tmp)
                img.anchor = "B26"
                img.width = 800
                img.height = 300
                ws.add_image(img)
            except Exception as e:
                print(e)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out
    else:
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "SUBE plantilla.xlsx"
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out

tab1, tab2, tab3 = st.tabs(["📋 1. FORMATO ORIGINAL TAL CUAL", "📸 2. FOTOS", "📊 3. SONOMETRO PDF"])

with tab1:
    st.subheader("Pestaña 1 - Formato oficial R2-POE37-EP V04")
    if not os.path.exists("plantilla.xlsx"):
        st.warning("⚠️ Sube plantilla.xlsx al repo")
    else:
        st.success("✅ plantilla.xlsx encontrada")

    c1,c2 = st.columns(2)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S A S")
        municipio = st.text_input("MUNICIPIO", "ATACO")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA")
        punto = st.text_input("PUNTO No", "Punto 1 nocturno")
        coord_n = st.text_input("COORD ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W")
        desc = st.text_area("DESCRIPCION", "En la entrada o vía principal")
        fuente = st.text_input("Fuente", "mineria")
    with c2:
        proyecto = st.text_input("PROYECTO", "Ataco Tolima")
        fecha = st.date_input("FECHA", datetime.now())
        hora = st.text_input("HORA INICIO", "21:01")
        hora_fin = st.text_input("HORA FIN", "21:16")
        calib = st.text_input("Calib dB", "114.0")
        memoria = st.text_input("Memoria No", "1")
        laeq = st.number_input("LAeq,T", 0.0, 140.0, 65.0)
        vel = st.number_input("Vel Viento (m/s)", 0.0, 20.0, 0.3)
        dir_viento = st.text_input("Dir Viento", "N")
        temp = st.number_input("Temperatura (°C)", -10.0, 60.0, 29.0)
        hum = st.number_input("Humedad Relativa (%)", 0.0, 100.0, 75.0)
        precip = st.selectbox("¿Precipitaciones?", ["No", "Sí"], 0)
        tipo_ruido = st.selectbox("Tipo de Ruido", ["Continuo", "Intermitente"], 0)
        tiempo_op = st.text_input("Tiempo Operación", "60 min")
        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 2)
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 1)
        mapa = st.file_uploader("MAPA SATELITAL / ESQUEMA", type=["jpg","png","jpeg"])
        esquema_txt = st.text_area("Texto ESQUEMA PUNTO", "Norte: vía, Sur: casas a 15m")

    st.divider()
    st.markdown("### ✅ ESCENARIO DE MEDICIÓN")
    e1, e2, e3 = st.columns(3)
    with e1:
        esc_diurno = st.checkbox("Diurno", value=(periodo=="Diurno"))
        esc_nocturno = st.checkbox("Nocturno", value=(periodo=="Nocturno"))
    with e2:
        esc_encendida = st.checkbox("Fuente Encendida", value=True)
        esc_apagada = st.checkbox("Fuente Apagada", value=False)
    with e3:
        mide_residual = st.selectbox("¿Mide Ruido Residual?", ["Sí", "No"], 0)
        barrido = st.selectbox("¿Barrido perimetral?", ["Sí", "No"], 0)
    justificacion = st.text_input("Justificación si NO mide residual", "")

    limite = NORMA[sector][periodo]
    cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"
    st.info(f"Límite {limite} dB - {cumple} | Temp {temp}°C Hum {hum}%")

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "PROYECTO": proyecto, "MUNICIPIO": municipio, "DEPTO": depto,
            "PUNTO": punto, "COORD_N": coord_n, "DESC": desc,
            "FECHA": str(fecha), "HORA": f"{hora}-{hora_fin}", "CALIB": calib, "MEMORIA": memoria,
            "LAEQ": laeq, "VEL": vel, "DIR": dir_viento, "TEMP": temp, "HUM": hum, "PRECIP": precip,
            "FUENTE": fuente, "TIPO_RUIDO": tipo_ruido, "TIEMPO_OP": tiempo_op,
            "SECTOR": sector, "PERIODO": periodo, "LIMITE": limite, "CUMPLE": cumple,
            "ESQUEMA": esquema_txt
        })
        st.success(f"Punto {punto} agregado")

    if st.session_state.puntos:
        st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
        if st.button("📥 GENERAR EXCEL TAL CUAL + ESCENARIO + TEMP/HUM", type="primary", use_container_width=True):
            escenario_data = {
                "diurno": esc_diurno, "nocturno": esc_nocturno,
                "encendida": esc_encendida, "apagada": esc_apagada,
                "residual": mide_residual, "justificacion": justificacion, "barrido": barrido
            }
            excel_file = crear_excel_con_plantilla_oficial(st.session_state.puntos, mapa, "plantilla.xlsx", escenario_data)
            st.download_button("📥 DESCARGAR EXCEL V04 COMPLETO", excel_file.getvalue(), f"R2-POE37-EP_V04_{municipio}_{punto}_COMPLETO.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

with tab2:
    st.subheader("Pestaña 2 - Fotos")
    if not st.session_state.puntos:
        st.warning("Agrega puntos en pestaña 1")
    else:
        nombres = [p["PUNTO"] for p in st.session_state.puntos]
        sel = st.selectbox("Punto", nombres)
        fotos = st.file_uploader(f"Fotos {sel}", type=["jpg","png","jpeg"], accept_multiple_files=True, key=f"f_{sel}")
        if fotos:
            if sel not in st.session_state.fotos_puntos:
                st.session_state.fotos_puntos[sel] = []
            st.session_state.fotos_puntos[sel].extend(fotos)
            st.success(f"{len(fotos)} fotos")
        for pn, lista in st.session_state.fotos_puntos.items():
            st.write(f"**{pn}** {len(lista)} fotos")
            cols = st.columns(5)
            for i,f in enumerate(lista):
                cols[i%5].image(f, width=100)
        if st.session_state.fotos_puntos and st.button("📦 ZIP", type="primary", use_container_width=True):
            zb = io.BytesIO()
            with zipfile.ZipFile(zb, "w") as zf:
                for pn, lista in st.session_state.fotos_puntos.items():
                    carp = pn.replace(" ","_")
                    for idx,f in enumerate(lista):
                        zf.writestr(f"{carp}/foto_{idx+1}_{f.name}", f.getvalue())
            zb.seek(0)
            st.download_button("📥 DESCARGAR ZIP", zb.getvalue(), "FOTOS_PUNTO_A_PUNTO.zip", use_container_width=True, type="primary")

with tab3:
    st.subheader("Pestaña 3 - Sonómetro CSV + PDF")
    c1,c2 = st.columns(2)
    with c1:
        f_total = st.file_uploader("Excel TOTAL", type=["xlsx","xls","csv"], key="tot")
        sector3 = st.selectbox("Sector", list(NORMA.keys()), 2, key="sec3")
        periodo3 = st.selectbox("Periodo", ["Diurno","Nocturno"], 1, key="per3")
    with c2:
        f_res = st.file_uploader("Excel RESIDUAL", type=["xlsx","xls","csv"], key="res")
        limite3 = NORMA[sector3][periodo3]
        st.metric("Límite", f"{limite3} dB")
    if f_total:
        serie_total, hoja, df_raw = leer_excel_sonometro(f_total)
        if serie_total is not None:
            leq_total = calcular_leq(serie_total)
            st.success(f"TOTAL {hoja} - {len(serie_total)} datos - {leq_total:.1f} dB")
            st.line_chart(pd.DataFrame({"TOTAL dB": serie_total.values}))
            leq_res = None
            serie_res = None
            if f_res:
                serie_res, _, _ = leer_excel_sonometro(f_res)
                if serie_res is not None:
                    leq_res = calcular_leq(serie_res)
                    st.info(f"RESIDUAL {leq_res:.1f} dB")
                    st.line_chart(pd.DataFrame({"RESIDUAL dB": serie_res.values}))
            if leq_res:
                dif = leq_total - leq_res
                if dif < 3:
                    leq_corr = leq_total
                    cumple3 = f"No medible Dif {dif:.1f}dB"
                elif dif > 10:
                    leq_corr = leq_total
                    cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
                else:
                    leq_corr = 10*np.log10(10**(leq_total/10) - 10**(leq_res/10))
                    cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
            else:
                leq_corr = leq_total
                cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
            st.metric("Corregido", f"{leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'} - {cumple3}")
            if st.button("📄 GENERAR PDF INFORME", type="primary", use_container_width=True):
                pdf_buffer = io.BytesIO()
                with PdfPages(pdf_buffer) as pdf:
                    fig1, ax1 = plt.subplots(figsize=(8.5,11))
                    ax1.axis('off')
                    p0 = st.session_state.puntos[0] if st.session_state.puntos else {}
                    txt = f"""INFORME RUIDO RES 627 - EQUISIMA
Sector: {sector3} Periodo: {periodo3} Limite: {limite3} dB
LAeq Total: {leq_total:.1f} dB
LAeq Residual: {f'{leq_res:.1f} dB' if leq_res else 'No medido'}
LAeq Corregido: {leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'}
Resultado: {cumple3}
Temp: {p0.get('TEMP','')}C Hum: {p0.get('HUM','')}%
"""
                    ax1.text(0.05,0.95, txt, fontsize=11, va='top', fontfamily='monospace')
                    pdf.savefig(fig1); plt.close(fig1)
                    fig2, ax2 = plt.subplots(figsize=(10,5))
                    ax2.plot(serie_total.values, label="TOTAL", linewidth=1)
                    if serie_res is not None:
                        ax2.plot(serie_res.values, label="RESIDUAL", alpha=0.7, linewidth=1)
                    ax2.axhline(limite3, color='r', ls='--', label=f"Limite {limite3}")
                    ax2.set_title(f"Time History - {cumple3}")
                    ax2.set_ylabel("dB(A)"); ax2.set_xlabel("Muestras"); ax2.legend(); ax2.grid(alpha=0.3)
                    pdf.savefig(fig2); plt.close(fig2)
                pdf_buffer.seek(0)
                st.download_button("📥 DESCARGAR PDF FINAL", pdf_buffer.getvalue(), f"INFORME_RES627_{cumple3}.pdf", mime="application/pdf", use_container_width=True, type="primary")
