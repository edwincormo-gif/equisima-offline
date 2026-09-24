import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

st.set_page_config(page_title="EQUISIMA FINAL ORIGINAL", layout="wide")
st.title("EQUISIMA - R2-POE37-EP V04 ORIGINAL COMPLETO")

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
    vals = np.array(vals); vals = vals[vals>0]
    return 10*np.log10(np.mean(10**(vals/10)))

def leer_excel_sonometro(file):
    try:
        xls = pd.ExcelFile(file)
        hoja = xls.sheet_names[0]
        for n in xls.sheet_names:
            if "history" in n.lower() or "time" in n.lower() or "data" in n.lower():
                hoja = n; break
        df = pd.read_excel(file, sheet_name=hoja)
        for col in df.columns:
            try:
                serie = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(serie)>10 and 20 < serie.mean() < 130:
                    return serie, hoja, df
            except: continue
        return None, hoja, df
    except Exception as e:
        return None, str(e), None

def crear_excel_original(puntos, foto_mapa=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "DATOS DE CAMPO EMISION DE RUIDO"

    bold = Font(bold=True, size=9)
    bold10 = Font(bold=True, size=10)
    normal = Font(size=9)
    fill_gray = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
    fill_dark = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # ENCABEZADO COMO TU FOTO
    ws.merge_cells("A1:I1")
    ws["A1"] = "DATOS DE CAMPO EMISION DE RUIDO"
    ws["A1"].font = Font(bold=True, size=12); ws["A1"].alignment = center
    ws["A1"].fill = fill_gray

    ws["A2"] = "Código: R2-POE37-EP V04"; ws["A2"].font = bold
    ws["E2"] = f"Fecha: {datetime.now().strftime('%Y-%m-%d')}"; ws["E2"].font = bold
    ws["H2"] = "Página: 1 de 1"; ws["H2"].font = bold

    p = puntos[0] if puntos else {}

    # Fila 3 - Cliente
    ws["A3"] = "CLIENTE:"; ws["A3"].font = bold; ws["A3"].fill = fill_gray
    ws.merge_cells("B3:D3"); ws["B3"] = p.get("CLIENTE",""); ws["B3"].border = border
    ws["E3"] = "PROYECTO:"; ws["E3"].font = bold; ws["E3"].fill = fill_gray
    ws.merge_cells("F3:I3"); ws["F3"] = p.get("PROYECTO","Ataco Tolima"); ws["F3"].border = border

    ws["A4"] = "MUNICIPIO:"; ws["A4"].font = bold; ws["A4"].fill = fill_gray
    ws["B4"] = p.get("MUNICIPIO",""); ws["B4"].border = border
    ws["C4"] = "DEPARTAMENTO:"; ws["C4"].font = bold; ws["C4"].fill = fill_gray
    ws.merge_cells("D4:E4"); ws["D4"] = p.get("DEPTO","TOLIMA"); ws["D4"].border = border
    ws["F4"] = "FECHA:"; ws["F4"].font = bold; ws["F4"].fill = fill_gray
    ws.merge_cells("G4:I4"); ws["G4"] = p.get("FECHA",""); ws["G4"].border = border

    # Fila 5 - Punto y coordenadas - COMO EN TU FOTO
    ws.merge_cells("A5:I5")
    ws["A5"] = "DATOS DEL PUNTO"; ws["A5"].font = bold10; ws["A5"].fill = fill_gray; ws["A5"].alignment = center

    ws["A6"] = "PUNTO No:"; ws["A6"].font = bold; ws["A6"].fill = fill_gray
    ws.merge_cells("B6:C6"); ws["B6"] = p.get("PUNTO",""); ws["B6"].border = border
    ws["D6"] = "COORDENADAS ORIGEN NACIONAL:"; ws["D6"].font = bold; ws["D6"].fill = fill_gray
    ws.merge_cells("E6:I6"); ws["E6"] = p.get("COORD_N",""); ws["E6"].border = border

    ws["A7"] = "COORD CTM12:"; ws["A7"].font = bold; ws["A7"].fill = fill_gray
    ws.merge_cells("B7:I7"); ws["B7"] = p.get("COORD_C12",""); ws["B7"].border = border

    ws["A8"] = "DESCRIPCION DEL PUNTO DE MEDICION:"; ws["A8"].font = bold; ws["A8"].fill = fill_gray
    ws.merge_cells("B8:I9"); ws["B8"] = p.get("DESC","En la entrada o vía principal - zona minera"); ws["B8"].border = border; ws["B8"].alignment = left

    ws["A10"] = "IDENTIFICACION DE FUENTES SONORAS:"; ws["A10"].font = bold; ws["A10"].fill = fill_gray
    ws.merge_cells("B10:I10"); ws["B10"] = p.get("FUENTE","Minería - maquinaria"); ws["B10"].border = border

    # TABLA MEDICION - COMO EN TU FOTO
    ws.merge_cells("A11:I11")
    ws["A11"] = "DATOS DE LA MEDICION"; ws["A11"].font = bold10; ws["A11"].fill = fill_gray; ws["A11"].alignment = center

    headers = ["Punto","Hora Inicio","Hora Fin","Calib (dB)","Memoria No","LAeq,T (dB)","Vel Viento m/s","Dir Viento","Temp °C","Hum %","Precip","Fuente Ruido","Tipo Ruido","Altura (m)","Barrido perim (dB)","Limite Res 627","Cumple"]
    # Ajustamos a 9 columnas visibles como en foto, resto en segunda fila
    headers_foto = ["PUNTO","HORA","MEMORIA","CALIB","LAeq","VEL","DIR","TEMP","HUM","PRECIP","FUENTE","TIPO","ALTURA","BARRIDO","LIMITE","CUMPLE","OBS"]
    for i,h in enumerate(headers_foto, start=1):
        if i>9: break
        c = ws.cell(row=12, column=i, value=h); c.font = bold; c.fill = fill_gray; c.border = border; c.alignment = center

    # Datos
    fila = 13
    for pt in puntos:
        vals = [pt.get("PUNTO",""), pt.get("HORA",""), pt.get("MEMORIA","1"), pt.get("CALIB","114"), pt.get("LAEQ",""), pt.get("VEL",""), pt.get("DIR",""), pt.get("TEMP",""), pt.get("HUM","")]
        for col,v in enumerate(vals, start=1):
            if col>9: break
            cell = ws.cell(row=fila, column=col, value=v); cell.border = border; cell.alignment = center; cell.font = normal
        fila+=1

    # UBICACION GEOGRAFICA - CON MAPA
    fila_mapa = fila+1
    ws.merge_cells(f"A{fila_mapa}:I{fila_mapa}")
    ws[f"A{fila_mapa}"] = "UBICACION GEOGRAFICA DEL PUNTO DE MONITOREO"; ws[f"A{fila_mapa}"].font = bold10; ws[f"A{fila_mapa}"].fill = fill_gray; ws[f"A{fila_mapa}"].alignment = center

    ws.merge_cells(f"A{fila_mapa+1}:I{fila_mapa+6}")
    ws[f"A{fila_mapa+1}"] = "MAPA - Ver imagen satelital adjunta"
    ws[f"A{fila_mapa+1}"].alignment = center
    for r in range(fila_mapa+1, fila_mapa+7):
        ws.row_dimensions[r].height = 25

    if foto_mapa:
        try:
            pil_img = PILImage.open(foto_mapa)
            pil_img.thumbnail((800, 400))
            tmp = "/tmp/mapa.png"
            pil_img.save(tmp)
            img = ExcelImage(tmp)
            img.anchor = f"A{fila_mapa+1}"
            img.width = 700; img.height = 150
            ws.add_image(img)
        except: pass

    # RESPONSABLE
    fila_resp = fila_mapa+8
    ws.merge_cells(f"A{fila_resp}:I{fila_resp}")
    ws[f"A{fila_resp}"] = "Responsable de la medicion:"; ws[f"A{fila_resp}"].font = bold

    ws[f"A{fila_resp+1}"] = "Nombre:"; ws[f"A{fila_resp+1}"].font = bold
    ws.merge_cells(f"B{fila_resp+1}:D{fila_resp+1}"); ws[f"B{fila_resp+1}"].border = border
    ws[f"E{fila_resp+1}"] = "Firma:"; ws[f"E{fila_resp+1}"].font = bold
    ws.merge_cells(f"F{fila_resp+1}:I{fila_resp+1}"); ws[f"F{fila_resp+1}"].border = border

    for i in range(1,10):
        ws.column_dimensions[get_column_letter(i)].width = 13

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

tab1, tab2, tab3 = st.tabs(["📋 1. FORMATO ORIGINAL", "📸 2. FOTOS PUNTO A PUNTO", "📊 3. SONOMETRO PDF"])

with tab1:
    st.subheader("Pestaña 1 - Formato original calcado a tu foto")
    c1,c2 = st.columns(2)
    with c1:
        cliente = st.text_input("CLIENTE", "Rueda Inversiones S A S")
        municipio = st.text_input("MUNICIPIO", "ATACO")
        depto = st.text_input("DEPARTAMENTO", "TOLIMA")
        punto = st.text_input("PUNTO No", "Punto 1 nocturno")
        coord_n = st.text_input("COORDENADAS ORIGEN NACIONAL", "3°35'11.30\"N 75°23'27.72\"W")
        coord_c12 = st.text_input("COORD CTM12", "E 876543 N 912345")
        desc = st.text_area("DESCRIPCION DEL PUNTO", "En la entrada o vía principal, zona de minería")
    with c2:
        proyecto = st.text_input("PROYECTO", "Ataco Tolima")
        fecha = st.date_input("FECHA", datetime.now())
        hora = st.text_input("HORA INICIO", "21:01")
        hora_fin = st.text_input("HORA FIN", "21:16")
        calib = st.text_input("Calibración dB", "114.0")
        memoria = st.text_input("Memoria No", "1")
        laeq = st.number_input("LAeq,T", 0.0, 140.0, 65.0)
        vel = st.number_input("Vel Viento", 0.0, 20.0, 0.3)
        fuente = st.text_input("Fuente", "mineria - maquinaria")
        sector = st.selectbox("SECTOR RES 627", list(NORMA.keys()), 2)
        periodo = st.selectbox("PERIODO", ["Diurno","Nocturno"], 1)
        mapa = st.file_uploader("MAPA SATELITAL (como en tu foto)", type=["jpg","png","jpeg"])

    limite = NORMA[sector][periodo]
    cumple = "CUMPLE" if laeq <= limite else "NO CUMPLE"
    st.info(f"Límite {limite} dB - {cumple}")

    if st.button("📍 AGREGAR PUNTO AL FORMATO ORIGINAL", type="primary", use_container_width=True):
        st.session_state.puntos.append({
            "CLIENTE": cliente, "PROYECTO": proyecto, "MUNICIPIO": municipio, "DEPTO": depto,
            "PUNTO": punto, "COORD_N": coord_n, "COORD_C12": coord_c12, "DESC": desc,
            "FECHA": str(fecha), "HORA": f"{hora}-{hora_fin}", "CALIB": calib, "MEMORIA": memoria,
            "LAEQ": laeq, "VEL": vel, "FUENTE": fuente, "SECTOR": sector, "PERIODO": periodo,
            "LIMITE": limite, "CUMPLE": cumple, "DIR": "N", "TEMP": 29, "HUM": 45
        })
        st.success("Punto agregado")

    if st.session_state.puntos:
        st.dataframe(pd.DataFrame(st.session_state.puntos), use_container_width=True)
        if st.button("📥 GENERAR EXCEL ORIGINAL COMPLETO", type="primary", use_container_width=True):
            excel_file = crear_excel_original(st.session_state.puntos, mapa)
            st.download_button("📥 DESCARGAR EXCEL ORIGINAL CALCADO A TU FOTO", excel_file.getvalue(), f"R2-POE37-EP_V04_{municipio}_{punto}_ORIGINAL.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

with tab2:
    st.subheader("Pestaña 2 - Fotos punto a punto")
    if not st.session_state.puntos:
        st.warning("Primero agrega puntos en pestaña 1")
    else:
        nombres = [p["PUNTO"] for p in st.session_state.puntos]
        sel = st.selectbox("Punto", nombres)
        fotos = st.file_uploader(f"Fotos {sel}", type=["jpg","png","jpeg"], accept_multiple_files=True, key=f"f_{sel}")
        if fotos:
            if sel not in st.session_state.fotos_puntos:
                st.session_state.fotos_puntos[sel] = []
            st.session_state.fotos_puntos[sel].extend(fotos)
            st.success(f"{len(fotos)} fotos en {sel}")
        for pn, lista in st.session_state.fotos_puntos.items():
            st.write(f"**{pn}** {len(lista)} fotos")
            cols = st.columns(5)
            for i,f in enumerate(lista):
                cols[i%5].image(f, width=100)
        if st.session_state.fotos_puntos and st.button("📦 ZIP PUNTO A PUNTO", type="primary", use_container_width=True):
            zb = io.BytesIO()
            with zipfile.ZipFile(zb, "w") as zf:
                for pn, lista in st.session_state.fotos_puntos.items():
                    carp = pn.replace(" ","_")
                    for idx,f in enumerate(lista):
                        zf.writestr(f"{carp}/foto_{idx+1}_{f.name}", f.getvalue())
                df = pd.DataFrame(st.session_state.puntos)
                eb = io.BytesIO(); df.to_excel(eb, index=False)
                zf.writestr("FORMATO_ORIGINAL/listado.xlsx", eb.getvalue())
            zb.seek(0)
            st.download_button("📥 DESCARGAR ZIP", zb.getvalue(), "FOTOS_PUNTO_A_PUNTO.zip", use_container_width=True, type="primary")

with tab3:
    st.subheader("Pestaña 3 - Sonómetro PDF")
    c1,c2 = st.columns(2)
    with c1:
        f_total = st.file_uploader("Excel TOTAL", type=["xlsx","xls"], key="tot")
        sector3 = st.selectbox("Sector", list(NORMA.keys()), 2, key="sec3")
        periodo3 = st.selectbox("Periodo", ["Diurno","Nocturno"], 1, key="per3")
    with c2:
        f_res = st.file_uploader("Excel RESIDUAL", type=["xlsx","xls"], key="res")
        limite3 = NORMA[sector3][periodo3]
        st.metric("Límite", f"{limite3} dB")

    if f_total:
        serie_total, hoja, df_raw = leer_excel_sonometro(f_total)
        if serie_total is not None:
            leq_total = calcular_leq(serie_total)
            st.success(f"TOTAL {hoja} {len(serie_total)} datos {leq_total:.1f} dB")
            st.line_chart(serie_total)
            leq_res = None
            serie_res = None
            if f_res:
                serie_res, _, _ = leer_excel_sonometro(f_res)
                if serie_res is not None:
                    leq_res = calcular_leq(serie_res)
                    st.info(f"RESIDUAL {leq_res:.1f} dB")
            if leq_res:
                if leq_total - leq_res < 3:
                    leq_corr = leq_total; cumple3 = "No medible <3dB"
                elif leq_total - leq_res > 10:
                    leq_corr = leq_total; cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
                else:
                    leq_corr = 10*np.log10(10**(leq_total/10) - 10**(leq_res/10))
                    cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
            else:
                leq_corr = leq_total; cumple3 = "CUMPLE" if leq_corr <= limite3 else "NO CUMPLE"
            st.metric("Corregido", f"{leq_corr if isinstance(leq_corr,str) else f'{leq_corr:.1f} dB'} {cumple3}")

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
Punto: {p0.get('PUNTO','')}
Municipio: {p0.get('MUNICIPIO','')}
Coord N: {p0.get('COORD_N','')}
Coord C12: {p0.get('COORD_C12','')}
Fecha: {datetime.now()}
"""
                    ax1.text(0.05,0.95, txt, fontsize=11, va='top', fontfamily='monospace')
                    pdf.savefig(fig1); plt.close(fig1)
                    fig2, ax2 = plt.subplots(figsize=(10,5))
                    ax2.plot(serie_total.values, label="TOTAL")
                    if serie_res is not None:
                        ax2.plot(serie_res.values, label="RESIDUAL", alpha=0.6)
                    ax2.axhline(limite3, color='r', ls='--', label=f"Limite {limite3}")
                    ax2.set_title(f"Time History {cumple3}"); ax2.set_ylabel("dB(A)"); ax2.legend(); ax2.grid(alpha=0.3)
                    pdf.savefig(fig2); plt.close(fig2)
                pdf_buffer.seek(0)
                st.download_button("📥 DESCARGAR PDF FINAL", pdf_buffer.getvalue(), f"INFORME_RES627_{cumple3}.pdf", mime="application/pdf", use_container_width=True, type="primary")
        else:
            st.error(f"No encontré dB Hoja {hoja}")
            if df_raw is not None:
                st.dataframe(df_raw.head(20))
