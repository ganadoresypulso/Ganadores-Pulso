import os
import openpyxl
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Ganadores & Pulso - Campeonato de Marcas", layout="wide"
)

# Título y encabezado principal
st.title("Ganadores & Pulso - Campeonato de Marcas")
st.subheader("Director: Nelson Osorio")

# Archivo de control de Excel
archivo_excel = "CONTROL CAMPEONATO DE MARCAS.xlsx"

# Verificar si el archivo existe
if os.path.exists(archivo_excel):
  try:
    wb = openpyxl.load_workbook(archivo_excel)

    st.sidebar.header("Menú de Navegación")
    opcion = st.sidebar.selectbox(
        "Seleccione una opción:", ["Registro de Marcas", "Ver Mis Marcas"]
    )

    if opcion == "Registro de Marcas":
      st.markdown("### Listado Oficial de Participantes y Carga de Marcas")

      # Formulario de registro de marcas para el público
      with st.form("form_marcas"):
        participante = st.text_input("Nombre del Participante / Marca:")
        clave = st.text_input("Clave de seguridad:", type="password")
        ejemplar = st.text_input("Ejemplar o Marca a Registrar:")

        submitted = st.form_submit_button("Enviar Marca")

        if submitted:
          if clave:
            sheet = wb.active
            sheet.append([participante, ejemplar])
            wb.save(archivo_excel)
            st.success(
                f"¡Marca de **{participante}** guardada con éxito en el sistema!"
            )
          else:
            st.warning("⚠️ Por favor ingresa tu clave de seguridad.")

    elif opcion == "Ver Mis Marcas":
      st.markdown("### Resumen de Marcas Registradas")
      df = pd.read_excel(archivo_excel)
      st.dataframe(df)

  except Exception as mi:
    st.error(f"Error al procesar el archivo: {mi}")
else:
  st.error("No se encuentra el archivo de Excel en la carpeta.")
