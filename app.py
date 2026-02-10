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

# --- INYECCIÓN DE CSS (ESTILO VISUAL CORREGIDO) ---
st.markdown("""
    <style>
    /* 1. FONDO GENERAL */
    .stApp {
        background-color: #F4F6F9;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    /* --- CORRECCIÓN DE COLORES (SOLUCIÓN AL TEXTO INVISIBLE) --- */
    /* Forzar que todas las etiquetas de los inputs sean oscuras */
    label[data-testid="stWidgetLabel"] p {
        color: #31333F !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
    }
    
    /* Forzar que el texto dentro de las cajas de carga sea oscuro */
    div[data-testid="stFileUploader"] small {
        color: #555555 !important;
    }

    /* Forzar color oscuro en las alertas (info/warning) */
    div[data-baseweb="notification"] p, div[data-baseweb="notification"] {
        color: #31333F !important;
    }
    /* --------------------------------------------------------- */

    /* 2. ESTILO DEL BOTÓN PRINCIPAL */
    div.stButton > button {
        background: linear-gradient(90deg, #7145D6 0%, #5633A8 100%);
        color: white;
        border: none;
        padding: 15px 32px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        font-weight: bold;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 12px;
        width: 100%;
        box-shadow: 0 4px 14px 0 rgba(113, 69, 214, 0.39);
        transition: transform 0.2s ease-in-out;
    }
    div.stButton > button:hover {
        transform: scale(1.02);
        box-shadow: 0 6px 20px 0 rgba(113, 69, 214, 0.50);
        color: white !important;
    }

    /* 3. ESTILO DE LOS SUBIDORES DE ARCHIVO (File Uploader) */
    div[data-testid="stFileUploader"] {
        background-color: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        border: 1px solid #E0E0E0;
    }
    
    /* 4. TÍTULOS DE SECCIÓN */
    .section-header {
        color: #4A4A4A;
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 10px;
        border-left: 5px solid #7145D6;
        padding-left: 10px;
    }

    /* Ocultar menú de hamburguesa y footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# --- ENCABEZADO PERSONALIZADO ---
st.markdown("""
    <div style="
        background: linear-gradient(135deg, #7145D6 0%, #9C7FE4 100%);
        padding: 30px;
        border-radius: 0 0 20px 20px;
        margin-bottom: 30px;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    ">
        <h1 style="color: white; margin:0; font-size: 2.5rem;">🟣 Conciliador Fiscal</h1>
        <p style="color: #E0E0E0; margin-top: 5px; font-size: 1.1rem;">
            Cruce inteligente de datos: DIAN vs ERP Netsuite
        </p>
    </div>
""", unsafe_allow_html=True)

# --- CONTENEDOR PRINCIPAL ---
main_container = st.container()

with main_container:
    col1, col2 = st.columns([1, 1], gap="large")

    # --- COLUMNA IZQUIERDA: DIAN ---
    with col1:
        st.markdown('<div class="section-header">📂 Documentos Fiscales (DIAN)</div>', unsafe_allow_html=True)
        
        with st.container():
            st.info("Obligatorio: Sube aquí el reporte descargado de la DIAN.")
            file_dian = st.file_uploader("Cargar Excel DIAN", type=["xlsx", "xls"], key="dian")
            
            st.markdown("---")
            st.caption("Opcional: Si tienes facturas de Gosocket")
            file_rec = st.file_uploader("Gosocket Recibidos", type=["xlsx", "xls"], key="rec")

    # --- COLUMNA DERECHA: CONTABILIDAD ---
    with col2:
        st.markdown('<div class="section-header">📊 Documentos Internos (Netsuite)</div>', unsafe_allow_html=True)
        
        with st.container():
            st.warning("Obligatorio: Sube el auxiliar contable unificado.")
            file_cont = st.file_uploader("Cargar Contabilidad", type=["xlsx", "xls"], key="cont")
            
            st.markdown("---")
            st.caption("Opcional: Si emites facturación electrónica externa")
            file_emi = st.file_uploader("Gosocket Emitidos", type=["xlsx", "xls"], key="emi")

# --- BOTÓN DE ACCIÓN ---
st.markdown("###")
col_b1, col_b2, col_b3 = st.columns([1, 2, 1])

with col_b2:
    process_btn = st.button("EJECUTAR CONCILIACIÓN AUTOMÁTICA")

# --- LÓGICA DE PROCESAMIENTO ---
if process_btn:
    if not file_dian or not file_cont:
        st.error("⚠️  ¡Atención! Faltan archivos obligatorios. Por favor carga el **Excel de la DIAN** y la **Contabilidad**.")
    else:
        status_box = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # 1. LECTURA
            status_box.markdown("🔄 **Leyendo y normalizando datos de la DIAN...**")
            progress_bar.progress(15)
            df_dian_raw = engine.leer_dian(file_dian)
            df_dian_raw = engine.crear_llave_conciliacion(df_dian_raw)
            
            # Separar DIAN
            df_dian_gastos = engine.filtrar_dian_gastos(df_dian_raw)
            df_dian_ingresos = engine.filtrar_dian_ingresos(df_dian_raw)
            
            status_box.markdown("🔄 **Procesando contabilidad Netsuite...**")
            progress_bar.progress(35)
            df_cont_full = engine.leer_contabilidad_completa(file_cont)
            
            if df_cont_full is None:
                st.error("❌ Error leyendo el archivo contable. Verifica el formato.")
                st.stop()
                
            # Segregación
            df_cont_gastos = engine.filtrar_solo_gastos(df_cont_full)
            df_cont_ingresos = engine.filtrar_solo_ingresos(df_cont_full)
            df_cont_iva_desc = engine.filtrar_solo_iva_descontable(df_cont_full)
            df_cont_iva_gen = engine.filtrar_solo_iva_generado(df_cont_full)
            
            df_rec = engine.leer_gosocket(file_rec)
            df_emi = engine.leer_gosocket(file_emi)
            
            # 2. PROCESAMIENTO
            status_box.markdown("⚙️ **Cruzando bases de datos...**")
            progress_bar.progress(60)
            
            c_gas, sd_gas, sc_gas = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_gastos)
            c_ing, sd_ing, sc_ing = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_ingresos)
            c_iva_d, sd_iva_d, sc_iva_d = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_iva_desc)
            c_iva_g, sd_iva_g, sc_iva_g = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_iva_gen)
            
            cg_ing, sg_cont, sg_go = None, None, None
            if df_emi is not None:
                cg_ing, sg_cont, sg_go = engine.conciliar_ingresos_vs_gosocket(df_cont_ingresos, df_emi)

            # 3. GENERACIÓN EXCEL
            status_box.markdown("📝 **Escribiendo reporte final...**")
            progress_bar.progress(85)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                try:
                    emisor_d = next((c for c in df_dian_raw.columns if 'nombre_emisor' in c), 'Emisor') 
                    receptor_d = next((c for c in df_dian_raw.columns if 'nombre_receptor' in c), 'Receptor')
                    total_d = next((c for c in df_dian_raw.columns if 'total_bruto' in c or 'total' in c), 'Total')
                    iva_d = next((c for c in df_dian_raw.columns if 'iva' in c or 'impuesto' in c), None)
                except:
                    emisor_d, receptor_d, total_d, iva_d = 'Emisor', 'Receptor', 'Total', 'IVA'

                engine.procesar_reporte_cabify_generico(c_gas, sd_gas, sc_gas, writer, '1. Conciliacion Gastos', emisor_d, total_d, iva_d, False)
                engine.procesar_reporte_cabify_generico(c_ing, sd_ing, sc_ing, writer, '2. Conciliacion Ingresos', receptor_d, total_d, iva_d, False)
                
                if iva_d:
                    engine.procesar_reporte_cabify_generico(c_iva_d, sd_iva_d, sc_iva_d, writer, '3. IVA Descontable', emisor_d, iva_d, None, True)
                    engine.procesar_reporte_cabify_generico(c_iva_g, sd_iva_g, sc_iva_g, writer, '3.1 IVA Generado', receptor_d, iva_d, None, True)

                df_cont_full.to_excel(writer, sheet_name='Base Contable Depurada', index=False)
                engine.formatear_hoja_base(writer, 'Base Contable Depurada', df_cont_full)
                
                if df_emi is not None:
                    df_emi.to_excel(writer, sheet_name='Base Gosocket Emitidos', index=False)
                    engine.formatear_hoja_base(writer, 'Base Gosocket Emitidos', df_emi)
                
                df_dian_raw.to_excel(writer, sheet_name='Base DIAN', index=False)
                engine.formatear_hoja_base(writer, 'Base DIAN', df_dian_raw)

            progress_bar.progress(100)
            status_box.success("✅ ¡Reporte generado! Descárgalo abajo.")
            
            st.markdown("###")
            st.download_button(
                label="📥  DESCARGAR REPORTE EXCEL FINAL",
                data=output.getvalue(),
                file_name="Reporte_Conciliacion_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")

