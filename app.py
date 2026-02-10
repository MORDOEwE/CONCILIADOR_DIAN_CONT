import streamlit as st
import pandas as pd
import io
import engine  # Tu archivo de lógica

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Cabify | Conciliador Fiscal",
    page_icon="💜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- ESTILOS CABIFY BRANDING ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');

    /* Reset y Tipografía */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #F3F3F7;
    }

    /* Fondo de la App */
    .stApp {
        background-color: #F3F3F7;
    }

    /* Header Personalizado */
    .cabify-header {
        background-color: #7145D6;
        padding: 2rem 3rem;
        margin: -5rem -5rem 2rem -5rem;
        color: white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }

    /* Tarjetas de Contenido */
    .upload-card {
        background-color: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border: 1px solid #E5E7EB;
        height: 100%;
    }

    /* Botones Estilo Cabify */
    .stButton > button {
        background-color: #7145D6 !important;
        color: white !important;
        border-radius: 50px !important;
        padding: 0.6rem 2.5rem !important;
        font-weight: 600 !important;
        border: none !important;
        transition: all 0.2s ease-in-out !important;
        box-shadow: 0 4px 6px rgba(113, 69, 214, 0.2) !important;
        width: 100%;
    }

    .stButton > button:hover {
        background-color: #5A36AD !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(113, 69, 214, 0.3) !important;
    }

    /* Inputs de Archivo */
    section[data-testid="stFileUploadDropzone"] {
        border: 2px dashed #D1D5DB !important;
        border-radius: 10px !important;
        background-color: #FAFAFB !important;
    }

    /* Títulos de sección */
    .section-title {
        color: #1F1F3D;
        font-size: 1.25rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    hr {
        margin: 1.5rem 0 !important;
        opacity: 0.1;
    }
    
    /* Ocultar elementos de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# --- CABECERA ---
st.markdown("""
    <div class="cabify-header">
        <p style="font-weight: 800; letter-spacing: 1px; margin-bottom: 0;">CABIFY FINANCE</p>
        <h1 style="margin-top: 0; font-weight: 800; font-size: 2.5rem;">Conciliador Fiscal</h1>
        <p style="opacity: 0.9; font-size: 1.1rem;">Ecosistema de cruce de datos contables y fiscales.</p>
    </div>
""", unsafe_allow_html=True)

# --- LAYOUT PRINCIPAL ---
st.markdown("<br>", unsafe_allow_html=True)

# Usamos un contenedor principal para centrar el contenido y dar aire
main_layout = st.container()

with main_layout:
    col_left, col_right = st.columns(2, gap="large")

    with col_left:
        st.markdown(f"""
            <div class="upload-card">
                <div class="section-title">🏢 Fuentes Externas (DIAN)</div>
                <p style="color: #6B7280; font-size: 0.9rem;">Carga los reportes oficiales de la autoridad tributaria.</p>
        """, unsafe_allow_html=True)
        
        file_dian = st.file_uploader("Reporte Oficial DIAN", type=["xlsx", "xls"], key="dian_up")
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.caption("Opcional: Gosocket Recibidos")
        file_rec = st.file_uploader("Cargar archivos XML recibidos", type=["xlsx", "xls"], key="rec_up")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_right:
        st.markdown(f"""
            <div class="upload-card">
                <div class="section-title">⚙️ Sistema Interno (ERP)</div>
                <p style="color: #6B7280; font-size: 0.9rem;">Carga el auxiliar contable extraído de Netsuite.</p>
        """, unsafe_allow_html=True)
        
        file_cont = st.file_uploader("Auxiliar Contable Netsuite", type=["xlsx", "xls"], key="cont_up")
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.caption("Opcional: Gosocket Emitidos")
        file_emi = st.file_uploader("Cargar archivos XML emitidos", type=["xlsx", "xls"], key="emi_up")
        st.markdown("</div>", unsafe_allow_html=True)

# --- ÁREA DE ACCIÓN ---
st.markdown("<br><br>", unsafe_allow_html=True)
c1, c2, c3 = st.columns([1, 1.2, 1])

with c2:
    process_btn = st.button("🚀 INICIAR PROCESO DE CONCILIACIÓN")

# --- LÓGICA DE PROCESAMIENTO ---
if process_btn:
    if not file_dian or not file_cont:
        st.toast("⚠️ Error: Faltan archivos críticos", icon="❌")
        st.error("Por favor, asegúrate de cargar tanto el archivo de la DIAN como el de Contabilidad.")
    else:
        # Contenedor para el estado visual
        with st.status("Ejecutando algoritmos de cruce...", expanded=True) as status:
            try:
                # 1. LECTURA
                st.write("Analizando estructuras de datos...")
                df_dian_raw = engine.leer_dian(file_dian)
                df_dian_raw = engine.crear_llave_conciliacion(df_dian_raw)
                
                df_dian_gastos = engine.filtrar_dian_gastos(df_dian_raw)
                df_dian_ingresos = engine.filtrar_dian_ingresos(df_dian_raw)
                
                st.write("Procesando Netsuite ERP...")
                df_cont_full = engine.leer_contabilidad_completa(file_cont)
                
                if df_cont_full is None:
                    st.error("Estructura contable no reconocida.")
                    st.stop()
                    
                df_cont_gastos = engine.filtrar_solo_gastos(df_cont_full)
                df_cont_ingresos = engine.filtrar_solo_ingresos(df_cont_full)
                df_cont_iva_desc = engine.filtrar_solo_iva_descontable(df_cont_full)
                df_cont_iva_gen = engine.filtrar_solo_iva_generado(df_cont_full)
                
                # 2. PROCESAMIENTO
                st.write("Realizando Match universal...")
                c_gas, sd_gas, sc_gas = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_gastos)
                c_ing, sd_ing, sc_ing = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_ingresos)
                c_iva_d, sd_iva_d, sc_iva_d = engine.ejecutar_conciliacion_universal(df_dian_gastos, df_cont_iva_desc)
                c_iva_g, sd_iva_g, sc_iva_g = engine.ejecutar_conciliacion_universal(df_dian_ingresos, df_cont_iva_gen)
                
                # 3. GENERACIÓN EXCEL
                st.write("Compilando reporte final en Excel...")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
                    df_dian_raw.to_excel(writer, sheet_name='Base DIAN', index=False)

                status.update(label="✅ Conciliación completada", state="complete", expanded=False)
                
                st.balloons()
                
                # Botón de descarga destacado
                st.markdown("---")
                st.success("El reporte ha sido generado exitosamente.")
                st.download_button(
                    label="📥 DESCARGAR REPORTE CONCILIACIÓN",
                    data=output.getvalue(),
                    file_name=f"Cabify_Conciliacion_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Error crítico: {str(e)}")
                st.info("Por favor, verifica que los encabezados de las columnas coincidan con el formato esperado.")
