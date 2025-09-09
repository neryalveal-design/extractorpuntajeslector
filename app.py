import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Extractor SIMCE", layout="centered")

st.title("📊 Extractor de Nóminas y Puntajes SIMCE")
st.write("""
Sube un archivo Excel con múltiples hojas (una por curso). Cada hoja debe contener los datos a partir de la **fila 11**, 
con la **columna C** como nombres (índice 2) y la **columna 167** como puntajes SIMCE (índice 166).
""")

uploaded_file = st.file_uploader("📂 Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        # Cargar archivo Excel
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        resultados = {}

        for sheet in sheet_names:
            df = excel_file.parse(sheet_name=sheet, header=None, skiprows=10)

            try:
                # Extraer columnas: nombre y puntaje
                df_nomina_simce = df.iloc[:, [2, 166]]  # C = index 2, nueva columna de puntajes = index 166
                df_nomina_simce.columns = ["Nombre", "Puntaje SIMCE"]

                # Filtrar solo filas con puntajes numéricos válidos (100-400)
                df_nomina_simce["Puntaje SIMCE"] = pd.to_numeric(df_nomina_simce["Puntaje SIMCE"], errors='coerce')
                df_nomina_simce = df_nomina_simce.dropna(subset=["Nombre", "Puntaje SIMCE"])
                df_nomina_simce = df_nomina_simce[df_nomina_simce["Puntaje SIMCE"].between(100, 400)]

                if not df_nomina_simce.empty:
                    resultados[sheet] = df_nomina_simce

            except Exception as e:
                st.warning(f"⚠️ No se pudo procesar la hoja '{sheet}': {e}")

        if resultados:
            # Crear archivo Excel en memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for sheet, df in resultados.items():
                    df.to_excel(writer, index=False, sheet_name=sheet)

            st.success("✅ Procesamiento completo. Descarga el archivo con las nóminas y puntajes por curso.")

            st.download_button(
                label="📥 Descargar Excel limpio",
                data=output.getvalue(),
                file_name="nóminas_y_puntajes_SIMCE.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se pudo procesar ninguna hoja del archivo.")

    except Exception as e:
        st.error(f"❌ Error al leer el archivo: {e}")
