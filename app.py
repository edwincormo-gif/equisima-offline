import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

st.set_page_config(page_title="EQUISIMA OFFLINE", page_icon="🟡", layout="wide")
st.title("🟡 EQUISIMA SAS - APP CAMPO OFFLINE")
st.caption("3224523451 | edwincormo@gmail.com")

def calcular(file):
    df = pd.read_csv(file, sep=';', encoding='latin-1', engine='python')
    col = df.columns[0]
    for c in df.columns:
        if 'LAeq' in str(c) or 'Leq' in str(c):
            col = c
            break
    v = pd.to_numeric(df[col].astype(str).str.replace(',','.'), errors='coerce').dropna()
    v = v[(v>20)&(v<140)]
    return round(10*np.log10((10**(v/10)).mean()),1)

if 'data' not in st.session_state:
    st.session_state.data = []

with st.form("form"):
    punto = st.text_input("Punto", "R-1 Puente Tierra")
    limite = st.selectbox("Limite dB", [75,55,65])
    macro = st.text_input("Macro", "Par Vial")
    c1,c2 = st.columns(2)
    ft = c1.file_uploader("CSV TOTAL a.csv", type='csv')
    fr = c2.file_uploader("CSV RESIDUAL a.csv", type='csv')
    foto = st.camera_input("Foto del punto")
    btn = st.form_submit_button("CALCULAR Y GUARDAR", type="primary", use_container_width=True)

if btn:
    if ft and fr:
        lt = calcular(ft)
        lr = calcular(fr)
        lra = lt if (lt-lr)>=10 else round(10*np.log10(10**(lt/10)-10**(lr/10)),1)
        estado = "CUMPLE" if lra<=limite else "NO CUMPLE"
        st.success(f"LT:{lt} dB | LR:{lr} dB | LRAeq:{lra} dB -> {estado}")
        st.session_state.data.append({"Punto":punto,"LRAeq":lra,"Estado":estado,"Fecha":datetime.now().strftime("%d/%m %H:%M")})
    else:
        st.error("Sube los 2 CSV")

if st.session_state.data:
    df = pd.DataFrame(st.session_state.data)
    st.dataframe(df, use_container_width=True)
    st.download_button("Descargar CSV", df.to_csv(index=False).encode('utf-8'), "informe.csv")
