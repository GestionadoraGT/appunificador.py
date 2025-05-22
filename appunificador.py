import streamlit as st
import pandas as pd
import os
import tempfile
from datetime import datetime
from openpyxl import Workbook
import xlrd

def convertir_a_xlsx_si_es_necesario(ruta_archivo):
    if ruta_archivo.endswith('.xls'):
        libro_xls = xlrd.open_workbook(ruta_archivo)
        nuevo_archivo = ruta_archivo + 'x'
        libro_xlsx = Workbook()
        hoja_xlsx = libro_xlsx.active

        hoja_xls = libro_xls.sheet_by_index(0)
        for fila_idx in range(hoja_xls.nrows):
            fila = hoja_xls.row_values(fila_idx)
            hoja_xlsx.append(fila)

        libro_xlsx.save(nuevo_archivo)
        return nuevo_archivo
    return ruta_archivo

def consolidar_archivo(ruta_archivo):
    ruta_archivo = convertir_a_xlsx_si_es_necesario(ruta_archivo)
    xls = pd.ExcelFile(ruta_archivo)
    nombres_hojas = xls.sheet_names

    columnas_a_eliminar = ['D', 'E', 'H', 'R', 'T', 'U', 'V', 'X', 'Y']
    df_total = pd.DataFrame()

    for hoja in nombres_hojas:
        df = pd.read_excel(xls, sheet_name=hoja, header=None, skiprows=14)
        df = df.dropna(how='all')  # elimina filas completamente vacías
        df_total = pd.concat([df_total, df], ignore_index=True)

    # Eliminar columnas vacías específicas (por letras)
    columnas_indices = [ord(c) - ord('A') for c in columnas_a_eliminar]
    df_total.drop(df_total.columns[columnas_indices], axis=1, inplace=True, errors='ignore')

    # Eliminar las primeras dos filas del DataFrame consolidado
    df_total = df_total.iloc[1:].reset_index(drop=True)

    return df_total

def generar_nombre_archivo():
    fecha_actual = datetime.today().strftime("%d-%m-%Y")
    return f"CARTERA POR ESTADO {fecha_actual}.xlsx"

def main():
    st.set_page_config(page_title="Unificador de Carteras", layout="centered")
    st.title("📊 Convertidor y Unificador de Reportes de Cartera")

    st.write("Sube un archivo de Excel (`.xls` o `.xlsx`) para consolidar todas sus hojas desde la fila 15.")

    archivo_subido = st.file_uploader("📁 Selecciona el archivo", type=["xls", "xlsx"])

    if archivo_subido is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            tmp.write(archivo_subido.read())
            tmp_path = tmp.name

        st.info("⏳ Procesando el archivo...")
        df_final = consolidar_archivo(tmp_path)

        nombre_salida = generar_nombre_archivo()
        salida_path = os.path.join(tempfile.gettempdir(), nombre_salida)
        df_final.to_excel(salida_path, index=False)

        with open(salida_path, "rb") as f:
            st.success("✅ Consolidación completada.")
            st.download_button(
                label="📥 Descargar archivo consolidado",
                data=f,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
