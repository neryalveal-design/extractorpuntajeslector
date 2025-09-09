import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Extractor y Analizador SIMCE / PAES", layout="centered")

st.title("📊 Extractor y Analizador de Puntajes SIMCE / PAES")
st.write("""
Esta aplicación te permite:
- Extraer nóminas de estudiantes y sus puntajes desde archivos Excel complejos.
- Reutilizar nóminas ya procesadas.
- Realizar un análisis de rendimiento según los criterios **SIMCE** o **PAES**.
""")

# Selección del tipo de análisis
analisis_tipo = st.selectbox(
    "Selecciona el tipo de análisis:",
    ["SIMCE", "PAES"]
)

# Función para clasificar según criterios
def clasificar_rendimiento(puntaje, tipo):
    if tipo == "SIMCE":
        if puntaje <= 250:
            return "Insuficiente"
        elif puntaje <= 285:
            return "Intermedio"
        else:
            return "Adecuado"
    elif tipo == "PAES":
        if puntaje <= 599:
            return "Insuficiente"
        elif puntaje <= 799:
            return "Intermedio"
        else:
            return "Adecuado"
    return "Desconocido"

# Cargar archivo
uploaded_file = st.file_uploader("📂 Sube tu archivo Excel (.xlsx)", type=["xlsx"])
modo = st.radio("¿Qué tipo de archivo estás subiendo?", ["Archivo original (complejo)", "Nómina ya procesada"], horizontal=True)

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        resultados = {}

        for sheet in sheet_names:
            df = excel_file.parse(sheet_name=sheet)

            try:
                if modo == "Archivo original (complejo)":
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, skiprows=10)
                    df_nomina = df.iloc[:, [2, 166]]
                    df_nomina.columns = ["Nombre", "Puntaje"]
                    df_nomina["Puntaje"] = pd.to_numeric(df_nomina["Puntaje"], errors="coerce")
                    df_nomina = df_nomina.dropna(subset=["Nombre", "Puntaje"])
                    df_nomina = df_nomina[df_nomina["Puntaje"].between(0, 1000)]
                else:
                    df_nomina = df[["Nombre", "Puntaje SIMCE" if "Puntaje SIMCE" in df.columns else "Puntaje"]]
                    df_nomina.rename(columns={"Puntaje SIMCE": "Puntaje"}, inplace=True)
                    df_nomina = df_nomina.dropna(subset=["Nombre", "Puntaje"])
                    df_nomina["Puntaje"] = pd.to_numeric(df_nomina["Puntaje"], errors="coerce")

                df_nomina["Rendimiento"] = df_nomina["Puntaje"].apply(lambda x: clasificar_rendimiento(x, analisis_tipo))
                resultados[sheet] = df_nomina

            except Exception as e:
                st.warning(f"⚠️ Hoja '{sheet}' no procesada: {e}")

        if resultados:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for sheet, df in resultados.items():
                    df.to_excel(writer, index=False, sheet_name=sheet)

            st.success("✅ Análisis completo. Descarga el archivo con resultados clasificados por rendimiento.")
            st.download_button(
                label="📥 Descargar archivo analizado",
                data=output.getvalue(),
                file_name=f"analisis_{analisis_tipo.lower()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No se pudo procesar ninguna hoja del archivo.")
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")
