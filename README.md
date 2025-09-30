# 📊 Resumen de Incurridos

Aplicación web desarrollada con Streamlit para analizar y resumir datos de proyectos de incurridos desde archivos Excel.

## 🚀 Cómo ejecutar la aplicación

Para ejecutar la aplicación, abre la terminal en la carpeta del proyecto y ejecuta:

```powershell
py -m streamlit run resumen_incurridos.py
```

La aplicación se abrirá automáticamente en tu navegador en la dirección `http://localhost:8501`.

## 📋 Funcionalidades

- **Selección de archivo**: Carga archivos Excel (.xlsx, .xls) desde tu computadora
- **Selección de proyecto**: Elige un proyecto específico para analizar
- **Métricas del proyecto**: 
  - Suma total de horas
  - Suma total de jornadas
  - Suma total de coste en EUR
- **Tabla de detalle**: Muestra todos los registros del proyecto seleccionado
- **Resumen por empleado**: 
  - Métricas individuales por empleado
  - Tabla resumen ordenada por horas

## 📁 Estructura de datos esperada

El archivo Excel debe contener las siguientes columnas:
- `Proyecto`
- `Nombre Completo`
- `Horas`
- `Jornadas`
- `FechaImputacion`
- `CosteEUR`
- Y otras columnas adicionales

## 🛠️ Requisitos

- Python 3.7+
- Streamlit
- Pandas
- OpenPyXL

## 📦 Instalación de dependencias

```powershell
py -m pip install streamlit pandas openpyxl
```
