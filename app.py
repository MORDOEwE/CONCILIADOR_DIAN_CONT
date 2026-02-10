import streamlit as st
import pandas as pd
import io
import engine  # Tu archivo de lógica

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Conciliador Fiscal",
    page_icon="🟣",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- INYECCIÓN DE CSS REFINADO ---
st.markdown("""
    <style>
    /* 1. FONDO Y TIPOGRAFÍA */
    .stApp {
        background-color: #F8F9FA;
        font-family: 'Inter', -apple-system, sans-serif;
    }

    /* 2. ETIQUETAS DE INPUTS */
    label[data-testid="stWidgetLabel"] p {
        color: #262730 !important;
        font-weight: 500 !important;
        font-size: 0.95rem !important;
    }
    
    /* 3. BOTONES MÁS PEQUEÑOS Y ELEGANTES */
    div.stButton > button {
        background-color: #6D31ED;
        color: white;
        border: none;
        padding: 8px 24px; /* Tamaño reducido */
        font-size: 14px;
        font-weight: 500;
        border-radius: 6px;
        width: auto; /* No ocupa todo el ancho */
        transition: all 0.3s ease;
        border: 1px solid #6D31ED;
    }
    div.stButton > button:hover {
        background-color: #5521B0;
        border-color: #5521B0;
        color: white !important;
    }

    /* 4. SECCIONES Y CONTENEDORES */
    div[data-testid="stFileUploader"] {
        background-color: white;
        border-radius: 8px;
        border: 1px solid #E0E0E0;
    }

    .section-header {
        color: #1F2937;
        font-size: 1.1rem;
        font-weight: 700;
        margin-bottom: 15px;
        border-left: 4px solid #6D31ED;
        padding-left: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* Ocultar elementos innecesarios */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# --- ENCABEZADO CORPORATIVO ---
st.markdown("""
    <div style="
        background-color: #FFFFFF;
        padding: 20px 40px;
        border-bottom: 1px solid #E5E7EB;
        margin-bottom: 30px;
    ">
        <h2 style="color: #111827; margin:0; font-size: 1.8rem; font-weight: 800;">Conciliador Fiscal</h2>
        <p style="color: #6B7280; margin-top: 2px; font-size: 0.9rem;">
            Cruce de datos DIAN y ERP Netsuite
        </p>
    </div>
""", unsafe_allow_html=True)

# --- CONTENEDOR PRINCIPAL ---
main_container = st.container()

with main_container:
    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        st.markdown('<div class="section-header">Documentos Fiscales (DIAN)</div>', unsafe_allow_html=True)
        st.info("Sube el reporte oficial descargado de la DIAN.")
        file_dian = st.file_uploader("Archivo DIAN", type=["xlsx", "xls"], key="dian")
        
        st.markdown("---")
        st.caption("Complementario")
        file_rec = st.file_uploader("Gosocket Recibidos", type=["xlsx", "xls"], key="rec")

    with col2:
        st.markdown('<div class="section-header">Contabilidad Interna (Netsuite)</div>', unsafe_allow_html=True)
        st.warning("Sube el auxiliar contable unificado.")
        file_cont = st.file_uploader("Archivo Contabilidad", type=["xlsx", "xls"], key="cont")
        
        st.markdown("---")
        st.caption("Complementario")
        file_emi = st.file_uploader("Gosocket Emitidos", type=["xlsx", "xls"], key="emi")

# --- BOTÓN DE ACCIÓN ---
st.markdown("<br>", unsafe_allow_html=True)
# Centramos el botón pero manteniendo su tamaño pequeño
_, col_center, _ = st.columns([1, 0.5, 1])

with col_center:
    process_btn = st.button("EJECUTAR CONCILIACIÓN")

# --- LÓGICA DE PROCESAMIENTO ---
if process_btn:
    if not file_dian or not file_cont:
        st.error("Atención: Faltan archivos obligatorios (DIAN y Contabilidad).")
    else:
        status_box = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # 1. LECTURA
            status_box.markdown("Procesando datos de la DIAN...")
            progress_bar.progress(15)
            df_dian_raw = engine.leer_dian(file_dian)
            df_dian_raw = engine.crear_llave_conciliacion(df_dian_raw)
            
            df_dian_gastos = engine.filtrar_dian_gastos(df_dian_raw)
            df_dian_ingresos = engine.filtrar_dian_ingresos(df_dian_raw)
            
            status_box.markdown("Procesando contabilidad Netsuite...")
            progress_bar.progress(35)
            df_cont_full = engine.leer_contabilidad_completa(file_cont)
            
            if df_cont_full is None:
                st.error("Error leyendo el archivo contable. Verifica el formato.")
                st.stop()
                
            df_cont_gastos = engine.filtrar_solo_gastos(df_cont_full)
            df_cont_ingresos = engine.filtrar_solo_ingresos(df_cont_full)
            df_cont_iva_desc = engine.filtrar_solo_iva_descontable(df_cont_full)
            df_cont_iva_gen = engine.filtrar_solo_iva_generado(df_cont_full)
            
            df_rec = engine.leer_gosocket(file_rec)
            df_emi = engine.leer_gosocket(file_emi)
            
            # 2. PROCESAMIENTO
            status_box.markdown("Cruzando bases de datos...")
            progress_bar.progress(60)
            
            c_gas, sd_gas, sc_gas = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_gastos)
            c_ing, sd_ing, sc_ing = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_ingresos)
            c_iva_d, sd_iva_d, sc_iva_d = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_iva_desc)
            c_iva_g, sd_iva_g, sc_iva_g = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_iva_gen)
            
            cg_ing, sg_cont, sg_go = None, None, None
            if df_emi is not None:
                cg_ing, sg_cont, sg_go = engine.conciliar_ingresos_vs_gosocket(df_cont_ingresos, df_emi)

            # 3. GENERACIÓN EXCEL
            status_box.markdown("Generando reporte final...")
            progress_bar.progress(85)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Lógica de detección de columnas simplificada
                cols = df_dian_raw.columns
                emisor_d = next((c for c in cols if 'nombre_emisor' in c), 'Emisor') 
                receptor_d = next((c for c in cols if 'nombre_receptor' in c), 'Receptor')
                total_d = next((c for c in cols if 'total_bruto' in c or 'total' in c), 'Total')
                iva_d = next((c for c in cols if 'iva' in c or 'impuesto' in c), None)

                engine.procesar_reporte_cabify_generico(c_gas, sd_gas, sc_gas, writer, '1. Conciliacion Gastos', emisor_d, total_d, iva_d, False)
                engine.procesar_reporte_cabify_generico(c_ing, sd_ing, sc_ing, writer, '2. Conciliacion Ingresos', receptor_d, total_d, iva_d, False)
                
                if iva_d:
                    engine.procesar_reporte_cabify_generico(c_iva_d, sd_iva_d, sc_iva_d, writer, '3. IVA Descontable', emisor_d, iva_d, None, True)
                    engine.procesar_reporte_cabify_generico(c_iva_g, sd_iva_g, sc_iva_g, writer, '3.1 IVA Generado', receptor_d, iva_d, None, True)

                df_cont_full.to_excel(writer, sheet_name='Base Contable Depurada', index=False)
                engine.formatear_hoja_base(writer, 'Base Contable Depurada', df_cont_full)
                
                if df_emi is not None:
                    df_emi.to_excel(writer, sheet_name='Base Gosocket Emitidos', index=False)
                
                df_dian_raw.to_excel(writer, sheet_name='Base DIAN', index=False)

            progress_bar.progress(100)
            status_box.success("Proceso finalizado con éxito.")
            
            st.download_button(
                label="Descargar Reporte Final",
                data=output.getvalue(),
                file_name="Reporte_Conciliacion_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Se presentó un error en el proceso: {e}")
