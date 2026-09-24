import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
from PIL import Image as PILImage
import fitz

st.set_page_config(page_title="EQUISIMA 3 PESTAÑAS COMPLETA", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04 + FOTOS + SONOMETRO PDF")

if "puntos" not in st.session_state:
    st.session_state.puntos = []
if "fotos_puntos" not in st.session_state:
    st.session_state.fotos_puntos = {}

NORMA = {
    "A. Tranquilidad y Silencio": {"Diurno": 55, "Nocturno": 45},
    "B. Tranquilidad y Ruido Moderado": {"Diurno": 65, "Nocturno": 50},
    "C. Ruido Intermedio Restringido": {"Diurno": 75, "Nocturno": 70},
    "C. Industrial": {"Diurno": 75, "Nocturno": 75},
    "C. Centro Ciudad / Comercial": {"Diurno": 70, "Nocturno": 55},
}

def calcular_leq(vals):
    vals = np.array(vals)
    vals = vals[vals > 0]
    return 10*np.log10(np.mean(10**(vals/10)))

def leer_excel_sonometro(file):
    try:
        xls = pd.ExcelFile(file)
        hoja = xls.sheet_names[0]
        for n in xls.sheet_names:
            if "history" in n.lower() or "time" in n.lower() or "data" in n.lower():
                hoja = n
                break
        df = pd.read_excel(file, sheet_name=hoja)
        # busca columna numerica de dB
        for col in df.columns:
            try:
                serie = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(serie) > 10 and serie.mean() > 20 and serie.mean() < 130:
                    return serie, hoja, df
            except:
                continue
        return None, hoja, df
    except Exception as e:
        return None, str(e), None

tab1, tab2, tab3 = st.tabs(["📋 1. FORMATO ORIGINAL EXCEL", "📸 2. FOTOS PUNTO A PUNTO", "📊 3. SONOMETRO + PDF INFORME"])

with tab1:
    st.subheader("Pestaña 1 - Descarga Excel con todos los datos del original")
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S A S", key="c1")
        municipio = st.text_input("MUNICIPIO", "ATACO", key="m1")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA", key="d1")
        punto = st.text_input("PUNTO", "Punto 1 nocturno", key="p1")
    with c2:
        coord_n = st.text_input("COORD ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W", key="cn1")
        coord_c12 = st.text_input("COORD CTM12", "E 876543 N 912345", key="cc1")
        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 2, key="sec1")
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 1, key="per1")
    with c3:
        fecha = st.date_input("FECHA", datetime.now(), key="f1")
        hora = st.text_input("HORA", "21:01", key="h1")
        laeq = st.number_input("LAeq,T dB", 0.0, 140.0, 65.0, key="laeq1")
        fuente = st.text_input("FUENTE", "mineria", key="fu1")

    c4,c5,c6,c7 = st.columns(4)
    with c4:
        calib = st.text_input("Calibración", "114.0")
        memoria = st.text_input("Memoria", "1")
    with c5:
        vel = st.number_input("Vel Viento", 0.0, 20.0, 0.3)
        dirv = st.text_input("Dir Viento", "N")
    with c6:
        temp = st.number_input("Temp", -20.0, 60.0, 29.0)
        hum = st.number_input("Humedad", 0, 100, 45)
    with c7:
        precip = st.selectbox("Precip", ["NO","SI"])
        altura = st.text_input("Altura", "1.5")

    limite = NORMA[sector][periodo]
    cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "MUNICIPIO": municipio, "DEPTO": depto, "PUNTO": punto,
            "COORD_N": coord_n, "COORD_C12": coord_c12, "SECTOR": sector, "PERIODO": periodo,
            "FECHA": str(fecha), "HORA": hora, "LAEQ": laeq, "LIMITE": limite, "CUMPLE": cumple,
            "CALIB": calib, "MEMORIA": memoria, "VEL": vel, "DIR": dirv, "TEMP": temp, "HUM": hum,
            "PRECIP": precip, "FUENTE": fuente, "ALTURA": altura
        })
        st.success("Guardado")

    if st.session_state.puntos:
        st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
        if st.button("📥 DESCARGAR EXCEL FORMATO ORIGINAL COMPLETO", type="primary"):
            wb = Workbook()
            ws = wb.active
            ws.title = "Datos de Campo"
            bold = Font(bold=True, size=11)
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Alignment(horizontal="center", vertical="center")
            ws.merge_cells("A1:Q1")
            ws["A1"] = "R2-POE37-EP V04 - FORMATO ORIGINAL COMPLETO"; ws["A1"].font = Font(bold=True, size=12)
            p0 = st.session_state.puntos[0]
            ws["A3"] = "CLIENTE:"; ws["A3"].font = bold; ws.merge_cells("B3:Q3"); ws["B3"] = p0["CLIENTE"]
            ws["A4"] = "MUNICIPIO:"; ws["A4"].font = bold; ws["B4"] = p0["MUNICIPIO"]; ws["C4"] = "DEPTO:"; ws["C4"].font = bold; ws["D4"] = p0["DEPTO"]
            ws["A5"] = "PUNTO:"; ws["A5"].font = bold; ws.merge_cells("B5:Q5"); ws["B5"] = p0["PUNTO"]
            ws["A6"] = "COORD N:"; ws["A6"].font = bold; ws.merge_cells("B6:Q6"); ws["B6"] = p0["COORD_N"]
            ws["A7"] = "COORD C12:"; ws["A7"].font = bold; ws.merge_cells("B7:Q7"); ws["B7"] = p0["COORD_C12"]
            headers = ["FECHA","HORA","CALIB","MEMORIA","LAEQ","VEL","DIR","TEMP","HUM","PRECIP","FUENTE","ALTURA","COORD_N","COORD_C12","LIMITE","CUMPLE","SECTOR"]
            for i,h in enumerate(headers,1):
                c = ws.cell(row=9, column=i, value=h); c.font = bold; c.border = border; c.alignment = center
            fila=10
            for p in st.session_state.puntos:
                vals = [p["FECHA"],p["HORA"],p["CALIB"],p["MEMORIA"],p["LAEQ"],p["VEL"],p["DIR"],p["TEMP"],p["HUM"],p["PRECIP"],p["FUENTE"],p["ALTURA"],p["COORD_N"],p["COORD_C12"],p["LIMITE"],p["CUMPLE"],p["SECTOR"]]
                for col,v in enumerate(vals,1):
                    cell = ws.cell(row=fila, column=col, value=v); cell.border = border; cell.alignment = center
                fila+=1
            for i in range(1,18):
                ws.column_dimensions[get_column_letter(i)].width = 14
            out = io.BytesIO(); wb.save(out)
            st.download_button("📥 DESCARGAR AHORA", out.getvalue(), f"FORMATO_ORIGINAL_{p0['MUNICIPIO']}.xlsx", use_container_width=True, type="primary")

with tab2:
    st.subheader("Pestaña 2 - Fotos punto a punto y carpeta")
    if not st.session_state.puntos:
        st.warning("Primero agrega puntos en pestaña 1")
    else:
        puntos_nombres = [p["PUNTO"] for p in st.session_state.puntos]
        punto_sel = st.selectbox("Selecciona punto para subir fotos", puntos_nombres)
        fotos = st.file_uploader(f"Fotos para {punto_sel}", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"foto_{punto_sel}")

        if fotos:
            if punto_sel not in st.session_state.fotos_puntos:
                st.session_state.fotos_puntos[punto_sel] = []
            for f in fotos:
                st.session_state.fotos_puntos[punto_sel].append(f)
            st.success(f"{len(fotos)} fotos agregadas a {punto_sel}")

        # Mostrar fotos organizadas
        for punto_nombre, lista_fotos in st.session_state.fotos_puntos.items():
            st.write(f"**{punto_nombre}** - {len(lista_fotos)} fotos")
            cols = st.columns(4)
            for idx, f in enumerate(lista_fotos):
                cols[idx % 4].image(f, width=150)

        if st.session_state.fotos_puntos:
            if st.button("📦 ARMAR CARPETA PUNTO A PUNTO Y DESCARGAR ZIP", type="primary", use_container_width=True):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for punto_nombre, lista_fotos in st.session_state.fotos_puntos.items():
                        carpeta = punto_nombre.replace(" ","_")
                        for i, f in enumerate(lista_fotos):
                            zf.writestr(f"{carpeta}/foto_{i+1}_{f.name}", f.getvalue())
                    # También mete el excel
                    df = pd.DataFrame(st.session_state.puntos)
                    excel_buf = io.BytesIO()
                    df.to_excel(excel_buf, index=False)
                    zf.writestr("00_Formato_Original/Listado_Puntos.xlsx", excel_buf.getvalue())

                st.download_button("📥 DESCARGAR CARPETA PUNTO A PUNTO ZIP", zip_buffer.getvalue(), "FOTOS_PUNTO_A_PUNTO.zip", use_container_width=True, type="primary")

with tab3:
    st.subheader("Pestaña 3 - Sube datos sonómetro, compara y PDF con gráficas")
    col1,col2 = st.columns(2)
    with col1:
        f_total = st.file_uploader("Excel TOTAL (HD2010)", type=["xlsx","xls"], key="tot")
        sector3 = st.selectbox("Sector", list(NORMA.keys()), 2, key="sec3")
        periodo3 = st.selectbox("Periodo", ["Diurno","Nocturno"], 1, key="per3")
    with col2:
        f_res = st.file_uploader("Excel RESIDUAL (opcional)", type=["xlsx","xls"], key="res")
        limite3 = NORMA[sector3][periodo3]
        st.metric("Límite Res 627", f"{limite3} dB")

    if f_total:
        serie_total, hoja, df_raw = leer_excel_sonometro(f_total)
        if serie_total is not None:
            leq_total = calcular_leq(serie_total)
            st.success(f"TOTAL leído: Hoja {hoja} - {len(serie_total)} datos - LAeq {leq_total:.1f} dB")
            st.line_chart(serie_total)

            leq_res = None
            if f_res:
                serie_res, _, _ = leer_excel_sonometro(f_res)
                if serie_res is not None:
                    leq_res = calcular_leq(serie_res)
                    st.info(f"RESIDUAL: {leq_res:.1f} dB")

            # Cálculo corrección
            if leq_res:
                if leq_total - leq_res < 3:
                    leq_corr = "No medible - Aporte bajo"
                    cumple3 = "No evaluable"
                elif leq_total - leq_res > 10:
                    leq_corr = leq_total
                    cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
                else:
                    leq_corr = 10*np.log10(10**(leq_total/10) - 10**(leq_res/10))
                    cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
            else:
                leq_corr = leq_total
                cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"

            st.metric("LAeq Corregido Emisión", f"{leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'}")
            st.metric("Resultado", cumple3)

            if st.button("📄 GENERAR PDF LISTO PARA INFORME", type="primary", use_container_width=True):
                # Graficas
                fig, ax = plt.subplots(figsize=(8,4))
                ax.plot(serie_total.values, label="TOTAL")
                if f_res and leq_res:
                    ax.plot(serie_res.values, label="RESIDUAL", alpha=0.7)
                ax.axhline(limite3, color="r", linestyle="--", label=f"Límite {limite3} dB")
                ax.set_title(f"Ruido - {sector3} {periodo3} - {cumple3}")
                ax.set_ylabel("dB(A)"); ax.set_xlabel("Tiempo (s)")
                ax.legend(); ax.grid(True)
                buf = io.BytesIO()
                fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
                plt.close(fig)
                buf.seek(0)

                # PDF con PyMuPDF
                doc = fitz.open()
                page = doc.new_page()
                page.insert_text((50,50), f"INFORME RUIDO - RES 627 - EQUISIMA", fontsize=16)
                page.insert_text((50,80), f"Sector: {sector3} - {periodo3} - Limite: {limite3} dB", fontsize=11)
                page.insert_text((50,100), f"LAeq Total: {leq_total:.1f} dB", fontsize=11)
                if leq_res:
                    page.insert_text((50,120), f"LAeq Residual: {leq_res:.1f} dB", fontsize=11)
                    page.insert_text((50,140), f"LAeq Corregido: {leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'}", fontsize=11)
                page.insert_text((50,160), f"Resultado: {cumple3}", fontsize=12)
                page.insert_text((50,180), f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}", fontsize=10)

                # Inserta grafica
                page2 = doc.new_page()
                # Guarda imagen temp
                img_path = "/tmp/graf.png"
                with open(img_path, "wb") as f:
                    f.write(buf.getvalue())
                rect = fitz.Rect(50, 50, 550, 350)
                page2.insert_image(rect, filename=img_path)
                page2.insert_text((50,400), f"Grafica Time History - {len(serie_total)} datos", fontsize=10)

                pdf_buf = io.BytesIO()
                doc.save(pdf_buf)
                doc.close()

                st.download_button("📥 DESCARGAR PDF INFORME CON GRÁFICAS", pdf_buf.getvalue(), f"INFORME_RUIDO_RES627_{cumple3}.pdf", mime="application/pdf", use_container_width=True, type="primary")
        else:
            st.error(f"No pude leer dB en esa hoja. Hoja detectada: {hoja}")
            if df_raw is not None:
                st.dataframe(df_raw.head(20))
