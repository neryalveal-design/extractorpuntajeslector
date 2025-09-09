import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
from fpdf import FPDF

st.set_page_config(page_title="Extractor y Analizador SIMCE / PAES", layout="centered")

st.title("📊 Extractor, Analizador y Visualizador de Puntajes SIMCE / PAES")

st.write("""
Esta aplicación te permite:
- Extraer nóminas y puntajes desde archivos Excel complejos.
- Reutilizar nóminas ya procesadas.
- Clasificar automáticamente los rendimientos según criterios **SIMCE** o **PAES**.
- Visualizar gráficos de barras por curso.
- Exportar los gráficos a PDF.
""")

# Selección del tipo de análisis
analisis_tipo = st.selectbox("Selecciona el tipo de análisis:", ["SIMCE", "PAES"])

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

uploaded_file = st.file_uploader("📂 Sube tu archivo Excel (.xlsx)", type=["xlsx"])
modo = st.radio("¿Qué tipo de archivo estás subiendo?", ["Archivo original (complejo)", "Nómina ya procesada"], horizontal=True)

resultados = {}
graficos = {}

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

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

                # Crear gráfico de barras
                conteo = df_nomina["Rendimiento"].value_counts().reindex(["Insuficiente", "Intermedio", "Adecuado"], fill_value=0)
                fig, ax = plt.subplots()
                conteo.plot(kind="bar", ax=ax)
                ax.set_title(f"Rendimiento - {sheet}")
                ax.set_ylabel("Número de estudiantes")
                ax.set_xlabel("Categoría")
                ax.grid(axis='y')
                graficos[sheet] = fig

            except Exception as e:
                st.warning(f"⚠️ Hoja '{sheet}' no procesada: {e}")

        if resultados:
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                for sheet, df in resultados.items():
                    df.to_excel(writer, index=False, sheet_name=sheet)

            st.success("✅ Procesamiento completo.")

            st.download_button(
                label="📥 Descargar Excel con rendimientos",
                data=output_excel.getvalue(),
                file_name=f"analisis_{analisis_tipo.lower()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if st.button("📊 Mostrar gráficos por curso"):
                for sheet, fig in graficos.items():
                    st.pyplot(fig)

            if st.button("🖨️ Exportar gráficos a PDF"):
                pdf_buffer = BytesIO()
                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)

                for sheet, fig in graficos.items():
                    temp_img = BytesIO()
                    fig.savefig(temp_img, format='png', bbox_inches='tight')
                    temp_img.seek(0)

                    pdf.add_page()
                    pdf.set_font("Arial", "B", 16)
                    pdf.cell(0, 10, f"Rendimiento - {sheet}", ln=True)
                    pdf.image(temp_img, x=10, y=30, w=180)

                pdf.output(pdf_buffer)
                pdf_buffer.seek(0)

                st.download_button(
                    label="📄 Descargar PDF de gráficos",
                    data=pdf_buffer,
                    file_name="graficos_rendimiento.pdf",
                    mime="application/pdf"
                )

        else:
            st.error("No se pudo procesar ninguna hoja.")

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")