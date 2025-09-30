import pandas as pd
import streamlit as st
import os

# Función para leer el archivo Excel
def leer_datos(ruta_archivo):
    df = pd.read_excel(ruta_archivo)
    return df

# Función para mostrar resumen y detalle de un proyecto seleccionado
def mostrar_resumen_y_detalle(df, proyecto_sel):
    df_proy = df[df['Proyecto'] == proyecto_sel]
    suma_horas = df_proy['Horas'].sum()
    suma_jornadas = df_proy['Jornadas'].sum()
    suma_coste = df_proy['CosteEUR'].sum()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Suma de Horas", f"{suma_horas:.2f}")
    with col2:
        st.metric("Suma de Jornadas", f"{suma_jornadas:.2f}")
    with col3:
        st.metric("Suma de CosteEUR", f"{suma_coste:.2f}")
    
    st.subheader("Detalle del proyecto")
    st.dataframe(df_proy, use_container_width=True)
    
    # Resumen por empleado
    st.subheader("📊 Resumen por Empleado")
    resumen_empleados = df_proy.groupby('Nombre Completo').agg({
        'Horas': 'sum',
        'Jornadas': 'sum',
        'CosteEUR': 'sum'
    }).round(2).reset_index()
    
    # Ordenar por horas descendente
    resumen_empleados = resumen_empleados.sort_values('Horas', ascending=False)
    
    # Mostrar métricas por empleado
    for _, empleado in resumen_empleados.iterrows():
        with st.expander(f"👤 {empleado['Nombre Completo']}"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Horas", f"{empleado['Horas']:.2f}")
            with col2:
                st.metric("Jornadas", f"{empleado['Jornadas']:.2f}")
            with col3:
                st.metric("CosteEUR", f"{empleado['CosteEUR']:.2f}")
    
    # Tabla resumen completa
    st.subheader("📋 Tabla Resumen por Empleado")
    st.dataframe(resumen_empleados, use_container_width=True)

# Configuración de la página
st.set_page_config(page_title="Resumen Incurridos", layout="wide")
st.title("📊 Resumen de Incurridos")

# Sidebar para seleccionar archivo
st.sidebar.header("📁 Seleccionar Archivo")
archivo = st.sidebar.file_uploader("Selecciona el archivo Excel", type=['xlsx', 'xls'])

if archivo is not None:
    try:
        # Leer el archivo
        df = leer_datos(archivo)
        
        # Mostrar información general
        st.success(f"✅ Archivo cargado exitosamente. Total de registros: {len(df)}")
        
        # Obtener lista de proyectos
        proyectos = df['Proyecto'].dropna().unique()
        
        # Selector de proyecto
        st.header("🎯 Seleccionar Proyecto")
        proyecto_seleccionado = st.selectbox(
            "Elige un proyecto para ver su resumen:",
            proyectos,
            index=0
        )
        
        if proyecto_seleccionado:
            mostrar_resumen_y_detalle(df, proyecto_seleccionado)
            
    except Exception as e:
        st.error(f"❌ Error al leer el archivo: {str(e)}")
else:
    st.info("👆 Por favor, selecciona un archivo Excel en el panel lateral para comenzar.") 