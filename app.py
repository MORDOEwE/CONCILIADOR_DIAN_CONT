import streamlit as st
import pandas as pd
import io
import engine  # Importamos tu lógica del otro archivo

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Conciliador DIAN Netsuite", page_icon="📊", layout="wide")

# ESTILOS CSS (Tus colores Morado Cabify)
st.markdown("""
    <style>
    .stApp { background-color: #F8F9FA; }
    div.stButton > button:first-child {
        background-color: #7145D6;
        color: white;
        border-radius: 8px;
        height: 3em;
        width: 100%;
        border: none;
        font-weight: bold;
    }
    div.stButton > button:first-child:hover {
        background-color: #5a37ab;
        color: white;
    }
    .stDownloadButton > button {
        background-color: #28a745;
        color: white;
        border-radius: 8px;
    }
    h1 { color: #7145D6; }
    </style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.title("💜 Conciliador Contable: DIAN vs Netsuite")
st.markdown("---")

# --- INPUTS (Columnas) ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Archivos DIAN / Gosocket")
    file_dian = st.file_uploader("Cargar Excel DIAN", type=["xlsx", "xls"])
    file_rec = st.file_uploader("Gosocket Recibidos (Opcional)", type=["xlsx", "xls"])

with col2:
    st.subheader("📚 Archivos Contables")
    file_cont = st.file_uploader("Cargar Contabilidad Unificada", type=["xlsx", "xls"])
    file_emi = st.file_uploader("Gosocket Emitidos (Opcional)", type=["xlsx", "xls"])

# --- LÓGICA DE PROCESO ---
st.markdown("###")
if st.button("INICIAR PROCESO DE CONCILIACIÓN"):
    if not file_dian or not file_cont:
        st.error("⚠️ Error: Debes cargar obligatoriamente el archivo DIAN y la Contabilidad.")
    else:
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # 1. LECTURA
            status_text.info("Leyendo archivo DIAN...")
            progress_bar.progress(10)
            df_dian_raw = engine.leer_dian(file_dian)
            df_dian_raw = engine.crear_llave_conciliacion(df_dian_raw)
            
            # Separar DIAN
            df_dian_gastos = engine.filtrar_dian_gastos(df_dian_raw)
            df_dian_ingresos = engine.filtrar_dian_ingresos(df_dian_raw)
            
            status_text.info("Leyendo Contabilidad (esto puede tardar un poco)...")
            progress_bar.progress(30)
            df_cont_full = engine.leer_contabilidad_completa(file_cont)
            
            if df_cont_full is None:
                st.error("Error leyendo contabilidad. Revisa el formato.")
                st.stop()
                
            # Segregación Contable
            df_cont_gastos = engine.filtrar_solo_gastos(df_cont_full)
            df_cont_ingresos = engine.filtrar_solo_ingresos(df_cont_full)
            df_cont_iva_desc = engine.filtrar_solo_iva_descontable(df_cont_full)
            df_cont_iva_gen = engine.filtrar_solo_iva_generado(df_cont_full)
            
            # Gosocket
            df_rec = engine.leer_gosocket(file_rec)
            df_emi = engine.leer_gosocket(file_emi)
            
            # 2. PROCESAMIENTO
            status_text.info("Realizando cruces de datos...")
            progress_bar.progress(50)
            
            # Gastos
            c_gas, sd_gas, sc_gas = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_gastos)
            # Ingresos
            c_ing, sd_ing, sc_ing = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_ingresos)
            # IVA
            c_iva_d, sd_iva_d, sc_iva_d = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_iva_desc)
            c_iva_g, sd_iva_g, sc_iva_g = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_iva_gen)
            
            # Gosocket Match
            cg_ing, sg_cont, sg_go = None, None, None
            if df_emi is not None:
                cg_ing, sg_cont, sg_go = engine.conciliar_ingresos_vs_gosocket(df_cont_ingresos, df_emi)

            # 3. GENERACIÓN EXCEL
            status_text.info("Generando archivo Excel con formato...")
            progress_bar.progress(80)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Recuperar nombres de columnas DIAN
                emisor_d = next((c for c in df_dian_raw.columns if 'nombre_emisor' in c), 'Emisor') 
                receptor_d = next((c for c in df_dian_raw.columns if 'nombre_receptor' in c), 'Receptor')
                total_d = next((c for c in df_dian_raw.columns if 'total_bruto' in c or 'total' in c), 'Total')
                iva_d = next((c for c in df_dian_raw.columns if 'iva' in c or 'impuesto' in c), None)

                # Reportes
                engine.procesar_reporte_cabify_generico(c_gas, sd_gas, sc_gas, writer, '1. Conciliacion Gastos', emisor_d, total_d, iva_d, False)
                engine.procesar_reporte_cabify_generico(c_ing, sd_ing, sc_ing, writer, '2. Conciliacion Ingresos', receptor_d, total_d, iva_d, False)
                
                if iva_d:
                    engine.procesar_reporte_cabify_generico(c_iva_d, sd_iva_d, sc_iva_d, writer, '3. IVA Descontable', emisor_d, iva_d, None, True)
                    engine.procesar_reporte_cabify_generico(c_iva_g, sd_iva_g, sc_iva_g, writer, '3.1 IVA Generado', receptor_d, iva_d, None, True)

                # Bases
                df_cont_full.to_excel(writer, sheet_name='Base Contable Depurada', index=False)
                engine.formatear_hoja_base(writer, 'Base Contable Depurada', df_cont_full)
                
                if df_emi is not None:
                    df_emi.to_excel(writer, sheet_name='Base Gosocket Emitidos', index=False)
                    engine.formatear_hoja_base(writer, 'Base Gosocket Emitidos', df_emi)
                
                df_dian_raw.to_excel(writer, sheet_name='Base DIAN', index=False)
                engine.formatear_hoja_base(writer, 'Base DIAN', df_dian_raw)

            progress_bar.progress(100)
            status_text.success("✅ ¡Proceso finalizado con éxito!")
            
            # 4. DOWNLOAD BUTTON
            st.download_button(
                label="📥 Descargar Reporte Final (Excel)",
                data=output.getvalue(),
                file_name="Reporte_Conciliacion_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocurrió un error inesperado: {e}")
            # st.exception(e) # Descomentar para ver el error completo en pantalla si lo necesitas