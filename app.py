#pip install streamlit pandas openpyxl reportlab matplotlib seaborn

import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import base64
from functools import partial
from PIL import Image

# ===========================================
# CONFIGURACIÓN INICIAL DE LA PÁGINA
# ===========================================
st.set_page_config(
    page_title="Generador de reportes Hella",
    page_icon="📊",
    layout="wide"
)

# ===========================================
# FUNCIONES CACHEADAS
# ===========================================
@st.cache_data
def cached_load_excel(uploaded_file):
    # Leer todas las hojas del Excel
    xls = pd.ExcelFile(uploaded_file)
    
    # Verificar si existe la hoja "Crosstab"
    if "Crosstab" not in xls.sheet_names:
        available_sheets = ", ".join(xls.sheet_names)
        st.error(f'No se encontró la hoja "Crosstab" en el archivo. Hojas disponibles: {available_sheets}')
        st.stop()  # Detiene la ejecución
    
    # Cargar solo la hoja "Crosstab"
    return pd.read_excel(uploaded_file, sheet_name="Crosstab")

@st.cache_data
def cached_apply_filters(_df, pais, cliente, categoria, col_pais, col_cliente, col_categoria):
    filtered = _df.copy()
    if pais != "Todos" and col_pais in _df.columns:
        filtered = filtered[_df[col_pais] == pais]
    if cliente != "Todos" and col_cliente in _df.columns:
        filtered = filtered[_df[col_cliente] == cliente]
    if categoria != "Todos" and col_categoria in _df.columns:
        filtered = filtered[_df[col_categoria] == categoria]
    return filtered

@st.cache_data
def cached_generate_comparison(_filtered_df, year, months):
    try:
        # 1. Filtrar datos para el año y meses seleccionados
        current_data = _filtered_df[
            (_filtered_df['T200 Year'] == year) & 
            (_filtered_df['T230 Month'].isin(months))]
        
        # 2. Obtener productos únicos del año actual
        productos_actuales = current_data['M010 Material'].unique()
        
        # 3. Filtrar datos para los mismos meses del año anterior
        previous_data = _filtered_df[
            (_filtered_df['T200 Year'] == year - 1) & 
            (_filtered_df['T230 Month'].isin(months)) &
            (_filtered_df['M010 Material'].isin(productos_actuales))]
        
        # 4. Agrupar y sumar cantidades
        ventas_actual = current_data.groupby('M010 Material')['Net Product Sales'].sum().reset_index()
        ventas_actual.columns = ['Producto', f'Ventas_{year}']
        
        ventas_anterior = previous_data.groupby('M010 Material')['Net Product Sales'].sum().reset_index()
        ventas_anterior.columns = ['Producto', f'Ventas_{year-1}']
        
        # 5. Combinar ambos dataframes
        comparativa = pd.merge(ventas_actual, ventas_anterior, on='Producto', how='left')
        
        # 6. Calcular diferencia y porcentaje
        comparativa['Diferencia'] = comparativa[f'Ventas_{year}'] - comparativa[f'Ventas_{year-1}'].fillna(0)
        comparativa['Cambio_%'] = (comparativa['Diferencia'] / comparativa[f'Ventas_{year-1}'].replace(0, np.nan)) * 100
        comparativa['Cambio_%'] = comparativa['Cambio_%'].round(1)
        
        # 7. Ordenar por ventas actuales
        comparativa = comparativa.sort_values(by=f'Ventas_{year}', ascending=False)
        
        # 8. Generar gráficas
        # Top 50
        top_50 = comparativa.head(50).sort_values(by=f'Ventas_{year}', ascending=True)
        fig_top, ax_top = plt.subplots(figsize=(12, 14))
        y_pos = np.arange(len(top_50))
        bar_width = 0.35
        
        ax_top.barh(y_pos - bar_width/2, top_50[f'Ventas_{year}'], height=bar_width, color='#1f77b4', label=f'{year}')
        ax_top.barh(y_pos + bar_width/2, top_50[f'Ventas_{year-1}'].fillna(0), height=bar_width, color='#ff7f0e', label=f'{year-1}')
        
        for i, (_, row) in enumerate(top_50.iterrows()):
            cambio = row['Cambio_%']
            if not np.isnan(cambio):
                color = 'green' if cambio >= 0 else 'red'
                ax_top.text(max(row[f'Ventas_{year}'], row[f'Ventas_{year-1}']) + (max(top_50[f'Ventas_{year}'])/20),
                          i, f"{cambio}%", ha='left', va='center', color=color, fontweight='bold')
        
        ax_top.set_yticks(y_pos)
        ax_top.set_yticklabels(top_50['Producto'])
        ax_top.set_xlabel('Cantidad Vendida (Acumulada) - Euros €')
        ax_top.set_title(f'Top 50 Productos - Mes/Meses {", ".join(map(str, months))}\n{year-1} vs {year}')
        ax_top.legend()
        plt.tight_layout()
        
        # Bottom 50
        bottom_50 = comparativa.tail(50).sort_values(by=f'Ventas_{year}', ascending=True)
        fig_bottom, ax_bottom = plt.subplots(figsize=(12, 14))
        y_pos = np.arange(len(bottom_50))
        
        ax_bottom.barh(y_pos - bar_width/2, bottom_50[f'Ventas_{year}'], height=bar_width, color='#1f77b4', label=f'{year}')
        ax_bottom.barh(y_pos + bar_width/2, bottom_50[f'Ventas_{year-1}'].fillna(0), height=bar_width, color='#ff7f0e', label=f'{year-1}')
        
        for i, (_, row) in enumerate(bottom_50.iterrows()):
            cambio = row['Cambio_%']
            if not np.isnan(cambio):
                color = 'green' if cambio >= 0 else 'red'
                ax_bottom.text(max(row[f'Ventas_{year}'], row[f'Ventas_{year-1}']) + (max(bottom_50[f'Ventas_{year}'])/20),
                             i, f"{cambio}%", ha='left', va='center', color=color, fontweight='bold')
        
        ax_bottom.set_yticks(y_pos)
        ax_bottom.set_yticklabels(bottom_50['Producto'])
        ax_bottom.set_xlabel('Cantidad Vendida (Acumulada) - Euros €')
        ax_bottom.set_title(f'Bottom 50 Productos - Mes/Meses {", ".join(map(str, months))}\n{year-1} vs {year}')
        ax_bottom.legend()
        plt.tight_layout()
        
        return {
            'comparativa': comparativa,
            'fig_top': fig_top,
            'fig_bottom': fig_bottom
        }
        
    except Exception as e:
        st.error(f"Error en cached_generate_comparison: {str(e)}")
        return None

# ===========================================
# INTERFAZ DE USUARIO
# ===========================================

st.markdown(
    """
    <style>
    .stButton>button {
        background-color: #003c70;  /* Azul Dodger */
        color: white;
        border: none;
        border-radius: 5px;
        padding: 10px 24px;
    }
    .stButton>button:hover {
        background-color: #0066CC;  /* Azul más oscuro al pasar el mouse */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Logo y título
col1, col2, col3 = st.columns([3, 3, 1])
with col2:
    try:
        image = Image.open("Hella-1.webp")
        st.image(image, width=200)
    except Exception as e:
        st.warning(f"No se pudo cargar la imagen del logo: {e}")

st.markdown("<h1 style='text-align: center;'>Generador de reportes de ventas de Hella</h1>", unsafe_allow_html=True)

# Instrucciones
st.markdown("""
### Instrucciones:
1. **Sube el archivo Excel** (formatos .xlsx o .xls)
2. **Selecciona los filtros** para personalizar el reporte
3. **Revisa los datos filtrados**
4. **Genera y descarga** el reporte en PDF
""")

# Carga de archivo
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Cargar datos (con cache)
        df = cached_load_excel(uploaded_file)
        st.success("✅ Archivo cargado correctamente!")
        st.subheader("Vista previa de los datos")
        st.dataframe(df.head())

        # Definir nombres de columnas
        COL_PAIS = "C421 Country (Ship-to Party)"
        COL_CLIENTE = "C100 Sold-to party (STAIRS)"
        COL_CATEGORIA = "M136 Product Subcategory"

        # Obtener opciones para filtros
        def get_unique_values(df, column_name, default_options=["Todos"]):
            if column_name in df.columns:
                return default_options + sorted(df[column_name].dropna().unique().tolist())
            else:
                st.warning(f"No se encontró la columna: '{column_name}'")
                return default_options

        paises = get_unique_values(df, COL_PAIS)
        clientes = get_unique_values(df, COL_CLIENTE)
        categorias = get_unique_values(df, COL_CATEGORIA)

        # Filtros
        col1, col2, col3 = st.columns(3)
        with col1:
            selected_pais = st.selectbox(f"País ({COL_PAIS}):", options=paises, index=0)
        with col2:
            selected_cliente = st.selectbox(f"Cliente ({COL_CLIENTE}):", options=clientes, index=0)
        with col3:
            selected_categoria = st.selectbox(f"Categoría ({COL_CATEGORIA}):", options=categorias, index=0)

        # Aplicar filtros (con cache)
        filtered_df = cached_apply_filters(
            df, selected_pais, selected_cliente, selected_categoria,
            COL_PAIS, COL_CLIENTE, COL_CATEGORIA
        )

        # Análisis comparativo - Versión mejorada
        if not filtered_df.empty and 'T200 Year' in filtered_df.columns and 'T230 Month' in filtered_df.columns:
            
            # 1. Años disponibles (con verificación)
            available_years = sorted(filtered_df['T200 Year'].unique(), reverse=True)
            if not available_years:
                st.warning("⚠️ No hay datos para los filtros aplicados")
                st.stop()  # Detiene la ejecución si no hay años
            
            # 2. Selector de año (siempre visible)
            col1, col2 = st.columns(2)
            with col1:
                selected_year = st.selectbox(
                    "Selecciona el año a analizar:", 
                    options=available_years,
                    index=0
                )
            
            # 3. Meses disponibles para el año seleccionado (con verificación)
            with col2:
                available_months = sorted(filtered_df[filtered_df['T200 Year'] == selected_year]['T230 Month'].unique())
                if not available_months:
                    st.warning(f"No hay meses disponibles para {selected_year}")
                else:
                    selected_months = st.multiselect(
                        "Selecciona los meses:",
                        options=available_months,
                        default=[available_months[0]],
                        format_func=lambda x: f"Mes {x}"
                    )
            
            # Resto del código de generación de comparativa...
            st.write("")
            col1, col2, col3 = st.columns([3, 3, 1])
            with col2:                                  
                generar = st.button("Generar Reporte", type="primary", disabled=not selected_months)

            if generar and selected_months:
                with st.spinner("Calculando comparativa..."):
                    results = cached_generate_comparison(filtered_df, selected_year, selected_months)
                    if results:
                        comparativa = results['comparativa']
                        fig_top = results['fig_top']
                        fig_bottom = results['fig_bottom']
                        
                        # Mostrar resultados
                        st.subheader(f"📈 Top 50 Productos - Cambio {selected_year} vs {selected_year-1}")
                        st.pyplot(fig_top)
                        
                        st.subheader(f"📉 Bottom 50 Productos - Cambio {selected_year} vs {selected_year-1}")
                        st.pyplot(fig_bottom)
                        
                        # Guardar en session state
                        st.session_state.update({
                            'comparativa': comparativa,
                            'fig_top': fig_top,
                            'fig_bottom': fig_bottom
                        })

                        # Mostrar tabla
                        st.subheader("📊 Tabla Comparativa")
                        display_df = comparativa.copy()
                        display_df[f'Ventas_{selected_year}'] = display_df[f'Ventas_{selected_year}'].apply(lambda x: f"{x:,.0f}")
                        display_df[f'Ventas_{selected_year-1}'] = display_df[f'Ventas_{selected_year-1}'].apply(lambda x: f"{x:,.0f}" if not pd.isna(x) else "N/A")
                        display_df['Diferencia'] = display_df['Diferencia'].apply(lambda x: f"{x:+,.0f}")
                        display_df['Cambio_%'] = display_df['Cambio_%'].apply(lambda x: f"{x:+.1f}%" if not pd.isna(x) else "N/A")
                        
                        st.dataframe(
                            display_df,
                            use_container_width=True,
                            height=min(len(display_df) * 35 + 35, 500)
                        )

        # Generación de PDF
        st.subheader("📑 Configuración del Reporte PDF")

        st.markdown("""
        <style>
            div[data-testid="stSelectbox"] > div {
                border: none !important;
                box-shadow: none !important;
            }
        </style>
        """, unsafe_allow_html=True)

        # Menú desplegable centrado
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            graph_selection = st.selectbox(
                "**Seleccione las gráficas a incluir en el PDF:**",
                options=["Ambas gráficas (Top 50 y Bottom 50)", 
                        "Solo Top 50 Productos", 
                        "Solo Bottom 50 Productos"],
                index=0
            )

        # Definir qué gráficas incluir
        include_top = graph_selection in ["Ambas gráficas (Top 50 y Bottom 50)", "Solo Top 50 Productos"]
        include_bottom = graph_selection in ["Ambas gráficas (Top 50 y Bottom 50)", "Solo Bottom 50 Productos"]

        # Botón para generar PDF
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("Generar PDF", type="primary", use_container_width=True):
                if 'comparativa' not in st.session_state:
                    st.error("Primero genera la comparativa")
                    st.stop()
                
                # Función para generar PDF
                def generate_pdf(filtered_df, comparativa_data, options, pais, cliente, categoria, fig_top=None, fig_bottom=None):
                    from reportlab.lib.utils import ImageReader
                    buffer = BytesIO()
                    c = canvas.Canvas(buffer, pagesize=letter)
                    width, height = letter

                    try:
                        logo_path = "Hella-1.webp"
                        c.drawImage(logo_path, (width-250)/2, height-250, width=250, height=125, preserveAspectRatio=True)
                    except:
                        pass

                    c.setFont("Helvetica-Bold", 20)
                    c.drawCentredString(width/2, height-270, "Reporte de ventas de Hella")
                    c.setFont("Helvetica", 12)
                    c.drawCentredString(width/2, height-300, f"Generado el: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")

                    # Espacio en blanco antes de los filtros
                    c.setFont("Helvetica-Bold", 14)
                    c.drawCentredString(width/2, height-340, "Filtros Aplicados:")

                    # Lista de filtros centrada
                    c.setFont("Helvetica", 12)
                    filtros = [
                        f"• País: {pais}",
                        f"• Cliente: {cliente}",
                        f"• Categoría: {categoria}",
                        f"• Periodo: Mes/Meses {', '.join(map(str, options['selected_months']))} del año {options['selected_year']}"
                    ]

                    # Calculamos el ancho máximo del texto para centrarlo
                    max_width = max([c.stringWidth(filtro, "Helvetica", 12) for filtro in filtros])
                    start_x = (width - max_width) / 2

                    for i, filtro in enumerate(filtros):
                        c.drawString(start_x, height-370-(i*25), filtro)
                    
                    if fig_top:
                        img_temp = BytesIO()
                        fig_top.savefig(img_temp, format='png', bbox_inches='tight', dpi=150)
                        img_temp.seek(0)
                        c.showPage()
                        c.drawImage(ImageReader(img_temp), (width-500)/2, height-750, width=500, height=700)
                    
                    if fig_bottom:
                        img_temp = BytesIO()
                        fig_bottom.savefig(img_temp, format='png', bbox_inches='tight', dpi=150)
                        img_temp.seek(0)
                        c.showPage()
                        c.drawImage(ImageReader(img_temp), (width-500)/2, height-750, width=500, height=700)
                    
                    c.save()
                    buffer.seek(0)
                    return buffer

                pdf_buffer = generate_pdf(
                    filtered_df=filtered_df,
                    comparativa_data=st.session_state['comparativa'],
                    options={
                        'selected_year': selected_year,
                        'selected_months': selected_months
                    },
                    pais=selected_pais,
                    cliente=selected_cliente,
                    categoria=selected_categoria,
                    fig_top=st.session_state.get('fig_top') if include_top else None,
                    fig_bottom=st.session_state.get('fig_bottom') if include_bottom else None
                )

                st.success("¡Reporte generado con éxito!")
                filename = f"Reporte_Ventas_Hella_{selected_pais}_{selected_cliente}_{selected_categoria}_{selected_year}_Meses_{'_'.join(map(str, selected_months))}.pdf"

                # Botón de descarga centrado
                col1, col2, col3 = st.columns([2, 3, 1])  # Columnas con proporción 1-2-1 (centro más ancho)
                with col2:
                    st.download_button(
                        label="Descargar Reporte PDF",
                        data=pdf_buffer,
                        file_name=filename,
                        mime="application/pdf",
                    )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
else:
    st.info("ℹ️ Por favor, sube un archivo Excel para comenzar.")