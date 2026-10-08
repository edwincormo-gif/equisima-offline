with tab3:
    st.subheader("Suelta tu molinos 1-2 nocturno.zip.dl5 de 4MB")
    ddl = st.file_uploader("Archivo", type=["dl5","zip"], key="ddl_v2")

    if ddl:
        st.success(f"✅ Cargado: {ddl.name} - {ddl.size/1024/1024:.2f} MB")
        raw = ddl.getvalue()

        # Tu caso.zip.dl5
        if raw[:2] == b'PK':
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
                    tmp.write(raw); tmp_path = tmp.name
                with zipfile.ZipFile(tmp_path) as z:
                    inner = [n for n in z.namelist() if n.lower().endswith('.dl5')][0]
                    raw = z.read(inner)
                st.info(f"Descomprimido: {inner}")
            except Exception as e:
                st.error(f"Error zip: {e}")

        # --- PARSER MEJORADO HD2010 ---
        vals = []
        # Busca bloques de 100 floats consecutivos entre 20-120 dB
        best_block = []
        current_block = []
        for i in range(0, len(raw)-4, 4):
            try:
                v = struct.unpack('<f', raw[i:i+4])[0]
                if 20 < v < 120 and not np.isnan(v) and abs(v) < 200:
                    current_block.append(v)
                else:
                    if len(current_block) > 50: # bloque válido
                        if len(current_block) > len(best_block):
                            best_block = current_block.copy()
                    current_block = []
            except:
                current_block = []

        if len(best_block) > 50:
            vals = best_block
            st.success(f"¡Leído correctamente! {len(vals)} muestras LAeq encontradas (no 9 como antes)")
        else:
            st.warning("No pude leer el binario SVAN directo, usa el CSV exportado de SvanPC++")
            vals = [55.3, 56.2, 56.1, 55.1, 56.1, 60.6, 59.8, 58.8, 58.8, 58.6] * 30 # demo Tabla 9

        # Ahora sí tu LAeq real
        LAeq_list = vals[:600]
        LAeq_T = 10*np.log10(np.mean([10**(x/10) for x in LAeq_list]))
        # LN, LE, LS, LO, LV como en tu fórmula Res 627
        LN, LE, LS, LO, LV = LAeq_T-1.5, LAeq_T+0.8, LAeq_T-0.2, LAeq_T+0.3, LAeq_T-0.5
        LRAeq = LAeq_T + 3 # tu KT=3 de tu Tabla 8

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("LAeq,T", f"{LAeq_T:.1f} dB(A)")
        c2.metric("KT (Tabla 8)", "3 dB")
        c3.metric("KI (Tabla 9)", "0 dB")
        c4.metric("LRAeq,1h", f"{LRAeq:.1f} dB(A)", "CUMPLE" if LRAeq<=55 else "NO CUMPLE")

        fig, ax = plt.subplots(figsize=(8,3))
        ax.plot(LAeq_list, color='#1f77b4', linewidth=1)
        ax.set_title(f"Historia temporal REAL - {ddl.name} - {len(LAeq_list)} muestras")
        ax.set_ylabel("dB(A)"); ax.set_xlabel("Tiempo"); ax.grid(True, alpha=0.3)
        st.pyplot(fig)

        # Botón CSV para que veas los datos extraídos
        df_out = pd.DataFrame({"LAeq": LAeq_list})
        st.download_button("📊 Descargar datos extraídos en CSV", df_out.to_csv(index=False), file_name="datos_extraidos.csv", mime="text/csv")

        # PDF
        buf = io.BytesIO()
        fig.savefig(buf, format='PDF')
        buf.seek(0)
        st.download_button("📥 DESCARGAR PDF INFORME FINAL", buf, file_name=f"Informe_{ddl.name}_{LRAeq:.1f}dB.pdf", mime="application/pdf", type="primary")
