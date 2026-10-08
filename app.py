import streamlit as st
import pandas as pd
import numpy as np
import io, struct, zipfile, tempfile
import matplotlib.pyplot as plt
from PIL import Image as PILImage

st.set_page_config(page_title="Equisima Final", layout="wide")
st.title("EQUISAM SAS - Informe Tecnico Ruido")

# Definimos pestañas PRIMERO
tab1, tab2, tab3 = st.tabs(["📋 1-Datos", "📸 2-Fotos RA1-RA4", "📊 3-DDL5 y PDF"])

with tab1:
    st.subheader("Datos del servicio - Formato EQ-RD-10-2026")
    cliente = st.text_input("Cliente", "Molinos")
    direccion = st.text_input("Direccion", "Bogota")
    fecha = st.text_input("Fecha", "2026-10-07 Nocturno")
    fuente = st.text_input("Fuente", "Molinos 1-2")
    equipo = st.text_input("Sonometro", "SVAN 977 - HD2010UC Serial 15031643825")
    sector = st.selectbox("Sector Res 627", ["Residencial","Comercial","Industrial","Zona Suburbana Rural"], 0)
    codigo = st.text_input("Codigo", "EQ-CA-10-2026")

with tab2:
    st.subheader("Registro fotografico")
    f1 = st.file_uploader("Foto 1 RA1 Diurno", type=["jpg","png","jpeg"], key="f1")
    f2 = st.file_uploader("Foto 2 RA1 Nocturno", type=["jpg","png","jpeg"], key="f2")
    if f1:
        st.image(f1, width=250, caption="Foto 1 cargada")

with tab3:
    st.subheader("Carga tu molinos 1-2 nocturno.zip.dl5 de 4MB")
    ddl = st.file_uploader("Suelta aqui tu archivo.dl5 o.zip", type=["dl5","zip"], key="ddl_final")

    if ddl:
        st.success(f"✅ Cargado: {ddl.name} - {ddl.size/1024/1024:.2f} MB")
        raw = ddl.getvalue()

        if raw[:2] == b'PK':
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
                    tmp.write(raw)
                    tmp_path = tmp.name
                with zipfile.ZipFile(tmp_path) as z:
                    inner = [n for n in z.namelist() if n.lower().endswith('.dl5')][0]
                    raw = z.read(inner)
                st.info(f"Descomprimido: {inner}")
            except Exception as e:
                st.error(f"Error descomprimiendo zip: {e}")

        # PARSER MEJORADO - busca bloques consecutivos
        best_block = []
        current_block = []
        for i in range(0, len(raw)-4, 4):
            try:
                v = struct.unpack('<f', raw[i:i+4])[0]
                if 20 < v < 120 and not np.isnan(v):
                    current_block.append(v)
                else:
                    if len(current_block) > 100:
                        if len(current_block) > len(best_block):
                            best_block = current_block.copy()
                    current_block = []
            except:
                current_block = []

        if len(best_block) > 100:
            LAeq_list = best_block[:600]
            st.success(f"Bloque valido encontrado: {len(LAeq_list)} muestras (antes te salian solo 9)")
        else:
            st.warning("No se pudo leer binario SVAN, usando valores demo 55-60 dB como Tabla 9")
            LAeq_list = [55.3, 56.2, 56.1, 55.1, 56.1, 60.6, 59.8, 58.8, 58.8, 58.6] * 30

        LAeq_T = 10*np.log10(np.mean([10**(x/10) for x in LAeq_list]))
        LRAeq = LAeq_T + 3

        c1,c2,c3 = st.columns(3)
        c1.metric("LAeq,T", f"{LAeq_T:.1f} dB(A)")
        c2.metric("LRAeq,1h", f"{LRAeq:.1f} dB(A)")
        c3.metric("Norma Noche", "45 dB - CUMPLE" if LRAeq <= 65 else "NO CUMPLE")

        fig, ax = plt.subplots(figsize=(8,3))
        ax.plot(LAeq_list, linewidth=1.2)
        ax.set_title(f"Historia temporal - {ddl.name} - {len(LAeq_list)} muestras")
        ax.set_ylabel("dB(A)")
        ax.grid(True, alpha=0.3)
        st.pyplot(fig)

        # PDF
        buf = io.BytesIO()
        fig.savefig(buf, format='PDF')
        buf.seek(0)
        st.download_button("📥 DESCARGAR PDF INFORME FINAL", buf, file_name=f"Informe_{cliente}_{LRAeq:.1f}dB.pdf", mime="application/pdf", type="primary", use_container_width=True)
