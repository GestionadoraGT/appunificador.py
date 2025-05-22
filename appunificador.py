import pandas as pd
import streamlit as st
from datetime import datetime
import tempfile
import os

def convertir_a_xlsx_si_es_necesario(uploaded_file):
    if uploaded_file.name.endswith('.xlsx'):
        return uploaded_file

    df_dict = pd.read_excel(uploaded_file, sheet_name=None)

    temp_xlsx = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with pd.ExcelWriter(temp_xlsx.name, engine="openpyxl") as writer:
        for sheet, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    return temp_xlsx.name

def consolidar_excel(ruta_archivo):
    xls = pd.ExcelFile(ruta_archivo)
    nombres_hojas = xls.sheet_names

    df_base = pd.read_excel(xls, sheet_name=nombres_hojas[0], skiprows=14).dropna(how='all')
    columnas_base = df_base.columns

    columnas_a_eliminar = [3, 4, 7, 17, 19, 20, 21, 23, 24]
    columnas_existentes = [col for i, col in enumerate(df_base.columns) if i in columnas_a_eliminar]
    df_base.drop(columns=columnas_existentes, inplace=True)
    df_base = df_base.iloc[1:].reset_index(drop=True)

    dataframes = [df_base]

    for hoja in nombres_hojas[1:]:
        df = pd.read_excel(xls, sheet_name=hoja, skiprows=14).dropna(how='all')
        df.columns = columnas_base
        columnas_existentes = [col for i, col in enumerate(df.columns) if i in columnas_a_eliminar]
        df.drop(columns=columnas_existentes, inplace=True)
        df = df.iloc[1:].reset_index(drop=True)
        dataframes.append(df)

    df_final = pd.concat(dataframes, ignore_index=True)

    fecha_actual = datetime.today().strftime("%d-%m-%Y")
    nombre_archivo = f"CARTERA POR ESTADO {fecha_actual}.xlsx"

    output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(output.name, index=False)

    return output.name, nombre_archivo

# ================= STREAMLIT APP =================

st.set_page_config(page_title="Unificador de Excel", layout="centered")
st.title("📊 Unificador de Excel - CARTERA POR ESTADO")

st.write("Sube un archivo Excel (.xls o .xlsx) con múltiples hojas. Consolidaremos la información desde la fila 15.")

uploaded_file = st.file_uploader("Selecciona tu archivo Excel", type=["xls", "xlsx"])

if uploaded_file:
    with st.spinner("Procesando archivo..."):
        try:
            archivo_convertido = convertir_a_xlsx_si_es_necesario(uploaded_file)
            archivo_final, nombre_archivo = consolidar_excel(archivo_convertido)

            with open(archivo_final, "rb") as f:
                st.success("✅ Archivo procesado exitosamente.")
                st.download_button(
                    label="📥 Descargar archivo consolidado",
                    data=f,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {str(e)}")
