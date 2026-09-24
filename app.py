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
from matplotlib.backends.backend_pdf import PdfPages

st.set_page_config(page_title="EQUISIMA 3 PESTAÑAS V10", layout="wide")
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
    st.subheader("Pestaña 1 - Excel con todos los datos del original")
    c1,c2,c3 = st.columns(3)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S A S", key="c1")
        municipio = st.text_input("MUNICIPIO", "ATACO", key="m1")
        punto = st.text_input("PUNTO", "Punto 1 nocturno", key="p1")
    with c2:
        coord_n = st.text_input("COORD ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W", key="cn1")
        coord_c12 = st.text_input("COORD CTM12", "E 876543 N 912345", key="cc1")
        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 2, key="sec1")
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 1, key="per1")
    with c3:
        fecha = st.date_input("FECHA", datetime.now(), key="f1")
        laeq = st.number_input("LAeq,T dB", 0.0, 140.0, 65.0, key="laeq1")
        fuente = st.text_input("FUENTE", "mineria", key="fu1")
    c4,c5 = st.columns(2)
    with c4:
        calib = st.text_input("Calibración", "114.0")
        vel = st.number_input("Vel Viento", 0.0, 20.0, 0.3)
        temp = st.number_input("Temp", -20.0, 60.0, 29.0)
    with c5:
        memoria = st.text_input("Memoria", "1")
        hum = st.number_input("Humedad", 0, 100, 45)
        precip = st.selectbox("Precip", ["NO","SI"])

    limite = NORMA[sector][periodo]
    cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"

    if st.button("📍 AGREGAR PUNTO", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "MUNICIPIO": municipio, "PUNTO": punto,
            "COORD_N": coord_n, "COORD_C12": coord_c12, "SECTOR": sector, "PERIODO": periodo,
            "FECHA": str(fecha), "LAEQ": laeq, "LIMITE": limite, "CUMPLE": cumple,
            "CALIB": calib, "MEMORIA": memoria, "VEL": vel, "TEMP": temp, "HUM": hum,
            "PRECIP": precip, "FUENTE": fuente
        })
        st.success("Guardado")

    if st.session_state.puntos:
        st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
        if st.button("📥 DESCARGAR EXCEL FORMATO ORIGINAL", type="primary"):
            wb = Workbook()
            ws = wb.active
            bold = Font(bold=True, size=11)
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Alignment(horizontal="center", vertical="center")
            ws.merge_cells("A1:Q1")
            ws["A1"] = "R2-POE37-EP V04 - FORMATO ORIGINAL"; ws["A1"].font = Font(bold=True, size=12)
            p0 = st.session_state.puntos[0]
            ws["A3"] = "CLIENTE:"; ws["A3"].font = bold; ws.merge_cells("B3:Q3"); ws["B3"] = p0["CLIENTE"]
            ws["A4"] = "PUNTO:"; ws["A4"].font = bold; ws.merge_cells("B4:Q4"); ws["B4"] = p0["PUNTO"]
            ws["A5"] = "COORD N:"; ws["A5"].font = bold; ws.merge_cells("B5:Q5"); ws["B5"] = p0["COORD_N"]
            ws["A6"] = "COORD C12:"; ws["A6"].font = bold; ws.merge_cells("B6:Q6"); ws["B6"] = p0["COORD_C12"]
            headers = ["FECHA","PUNTO","CALIB","MEMORIA","LAEQ","VEL","TEMP","HUM","PRECIP","FUENTE","COORD_N","COORD_C12","LIMITE","CUMPLE","SECTOR"]
            for i,h in enumerate(headers,1):
                c = ws.cell(row=8, column=i, value=h); c.font = bold; c.border = border; c.alignment = center
            fila=9
            for p in st.session_state.puntos:
                vals = [p["FECHA"],p["PUNTO"],p["CALIB"],p["MEMORIA"],p["LAEQ"],p["VEL"],p["TEMP"],p["HUM"],p["PRECIP"],p["FUENTE"],p["COORD_N"],p["COORD_C12"],p["LIMITE"],p["CUMPLE"],p["SECTOR"]]
                for col,v in enumerate(vals,1):
                    cell = ws.cell(row=fila, column=col, value=v); cell.border = border; cell.alignment = center
                fila+=1
            for i in range(1,16):
                ws.column_dimensions[get_column_letter(i)].width = 14
            out = io.BytesIO(); wb.save(out)
            st.download_button("📥 DESCARGAR AHORA", out.getvalue(), f"FORMATO_ORIGINAL_{p0['MUNICIPIO']}.xlsx", use_container_width=True, type="primary")

with tab2:
    st.subheader("Pestaña 2 - Fotos punto a punto")
    if not st.session_state.puntos:
        st.warning("Agrega puntos en pestaña 1")
    else:
        puntos_nombres = [p["PUNTO"] for p in st.session_state.puntos]
        punto_sel = st.selectbox("Selecciona punto", puntos_nombres)
        fotos = st.file_uploader(f"Fotos para {punto_sel}", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"foto_{punto_sel}")
        if fotos:
            if punto_sel not in st.session_state.fotos_puntos:
                st.session_state.fotos_puntos[punto_sel] = []
            st.session_state.fotos_puntos[punto_sel].extend(fotos)
            st.success(f"{len(fotos)} fotos agregadas")

        for pn, lista in st.session_state.fotos_puntos.items():
            st.write(f"**{pn}** - {len(lista)} fotos")
            cols = st.columns(4)
            for idx, f in enumerate(lista):
                cols[idx % 4].image(f, width=150)

        if st.session_state.fotos_puntos and st.button("📦 ARMAR CARPETA ZIP PUNTO A PUNTO", type="primary", use_container_width=True):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for pn, lista in st.session_state.fotos_puntos.items():
                    carpeta = pn.replace(" ","_")
                    for i, f in enumerate(lista):
                        zf.writestr(f"{carpeta}/foto_{i+1}_{f.name}", f.getvalue())
            st.download_button("📥 DESCARGAR ZIP", zip_buffer.getvalue(), "FOTOS_PUNTO_A_PUNTO.zip", use_container_width=True, type="primary")

with tab3:
    st.subheader("Pestaña 3 - Sonómetro + PDF con gráficas")
    c1,c2 = st.columns(2)
    with c1:
        f_total = st.file_uploader("Excel TOTAL HD2010", type=["xlsx","xls"], key="tot")
        sector3 = st.selectbox("Sector", list(NORMA.keys()), 2, key="sec3")
        periodo3 = st.selectbox("Periodo", ["Diurno","Nocturno"], 1, key="per3")
    with c2:
        f_res = st.file_uploader("Excel RESIDUAL opcional", type=["xlsx","xls"], key="res")
        limite3 = NORMA[sector3][periodo3]
        st.metric("Límite Res 627", f"{limite3} dB")

    if f_total:
        serie_total, hoja, df_raw = leer_excel_sonometro(f_total)
        if serie_total is not None:
            leq_total = calcular_leq(serie_total)
            st.success(f"TOTAL: {hoja} - {len(serie_total)} datos - {leq_total:.1f} dB")
            st.line_chart(serie_total)

            leq_res = None
            if f_res:
                sr, _, _ = leer_excel_sonometro(f_res)
                if sr is not None:
                    leq_res = calcular_leq(sr)
                    st.info(f"RESIDUAL: {leq_res:.1f} dB")

            if leq_res:
                if leq_total - leq_res < 3:
                    leq_corr = "No medible"
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

            st.metric("LAeq Corregido", f"{leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'} - {cumple3}")

            if st.button("📄 GENERAR PDF LISTO PARA INFORME", type="primary", use_container_width=True):
                pdf_buffer = io.BytesIO()
                with PdfPages(pdf_buffer) as pdf:
                    # Página 1 - Datos
                    fig1, ax1 = plt.subplots(figsize=(8.5, 11))
                    ax1.axis('off')
                    texto = f"""
                    INFORME RUIDO - RESOLUCION 627
                    EQUISIMA S.A.S

                    Sector: {sector3}
                    Periodo: {periodo3}
                    Limite: {limite3} dB

                    LAeq Total: {leq_total:.1f} dB
                    LAeq Residual: {f'{leq_res:.1f} dB' if leq_res else 'No medido'}
                    LAeq Corregido Emision: {leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'}

                    Resultado: {cumple3}

                    Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}
                    Punto: {st.session_state.puntos[0]['PUNTO'] if st.session_state.puntos else 'N/A'}
                    Coord N: {st.session_state.puntos[0]['COORD_N'] if st.session_state.puntos else 'N/A'}
                    Coord C12: {st.session_state.puntos[0]['COORD_C12'] if st.session_state.puntos else 'N/A'}
                    """
                    ax1.text(0.1, 0.9, texto, fontsize=12, verticalalignment='top', fontfamily='monospace')
                    pdf.savefig(fig1)
                    plt.close(fig1)

                    # Página 2 - Gráfica
                    fig2, ax2 = plt.subplots(figsize=(8.5, 6))
                    ax2.plot(serie_total.values, label="TOTAL", linewidth=1)
                    ax2.axhline(limite3, color="r", linestyle="--", label=f"Límite {limite3} dB")
                    if leq_res:
                        ax2.axhline(leq_corr if not isinstance(leq_corr,str) else leq_total, color="g", linestyle="-.", label=f"Corregido {leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'}")
                    ax2.set_title(f"Time History - {cumple3} - {len(serie_total)} datos")
                    ax2.set_ylabel("dB(A)")
                    ax2.set_xlabel("Tiempo (s)")
                    ax2.legend()
                    ax2.grid(True, alpha=0.3)
                    pdf.savefig(fig2)
                    plt.close(fig2)

                st.download_button("📥 DESCARGAR PDF INFORME", pdf_buffer.getvalue(), f"INFORME_RES627_{cumple3}.pdf", mime="application/pdf", use_container_width=True, type="primary")
        else:
            st.error(f"No encontré columna dB. Hoja: {hoja}")
            if df_raw is not None:
                st.dataframe(df_raw.head(20))
