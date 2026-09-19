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

# Verificar si el archivo existe en el repositorio
if os.path.exists(archivo_excel):
  try:
    wb = openpyxl.load_workbook(archivo_excel)

    st.sidebar.header("Menú de Navegación")
    opcion = st.sidebar.selectbox(
        "Seleccione una opción:",
        ["Registro de Marcas del Día", "Listado de Participantes"],
    )

    if opcion == "Registro de Marcas del Día":
      st.markdown(
          "### 🏇 Registro Oficial de Marcas - Hipódromo La Rinconada"
      )

      with st.form("form_campeonato"):
        # Datos generales del participante
        st.markdown("#### Datos del Participante")
        col1, col2 = st.columns(2)
        with col1:
          participante = st.text_input("Nombre del Participante / Marca:")
        with col2:
          clave_participante = st.text_input(
              "Clave de seguridad personal:", type="password"
          )

        st.markdown("---")
        st.markdown("#### Selección de Marcas por Carrera (3 por carrera)")

        # Selector de cantidad de carreras del meeting
        num_carreras = st.selectbox(
            "Seleccione la cantidad de carreras de la reunión:",
            options=list(range(1, 11)),
            index=5,
        )  # Por defecto 6 carreras

        # Diccionario para almacenar las marcas ingresadas por carrera
        marcas_carreras = {}

        # Generar dinámicamente las cajitas para cada carrera (3 marcas + indicador de fija/superfija)
        for i in range(1, num_carreras + 1):
          st.markdown(f"**Carrera #{i}**")
          c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
          with c1:
            marca_1 = st.text_input(f"1era Marca (C-{i})", key=f"m1_{i}")
          with c2:
            marca_2 = st.text_input(f"2da Marca (C-{i})", key=f"m2_{i}")
          with c3:
            marca_3 = st.text_input(f"3era Marca (C-{i})", key=f"m3_{i}")
          with c4:
            tipo_fija = st.selectbox(
                f"Selección (C-{i})",
                ["Normal", "Fija", "Superfija"],
                key=f"fija_{i}",
            )

          marcas_carreras[f"Carrera_{i}"] = {
              "m1": marca_1,
              "m2": marca_2,
              "m3": marca_3,
              "tipo": tipo_fija,
          }

        st.markdown("---")
        submitted = st.form_submit_button(
            "🚀 Enviar Todas las Marcas al Sistema"
        )

        if submitted:
          if participante and clave_participante:
            sheet = wb.active
            # Guardamos la información detallada por cada carrera registrada
            for carrera, datos in marcas_carreras.items():
              if (
                  datos["m1"] or datos["m2"] or datos["m3"]
              ):  # Si al menos cargó una marca
                sheet.append(
                    [
                        participante,
                        carrera,
                        datos["m1"],
                        datos["m2"],
                        datos["m3"],
                        datos["tipo"],
                    ]
                )
            wb.save(archivo_excel)
            st.success(
                f"¡Excelente, {participante}! Tus marcas han sido registradas"
                " con éxito en el sistema."
            )
          else:
            st.warning(
                "⚠️ Por favor ingresa tu nombre y tu clave de seguridad antes"
                " de enviar."
            )

    elif opcion == "Listado de Participantes":
      st.markdown("### 📋 Listado de Participantes")
      df = pd.read_excel(archivo_excel)
      st.dataframe(df, use_container_width=True)

  except Exception as mi:
    st.error(f"Error al procesar el archivo del sistema: {mi}")
else:
  st.error(
      "No se encuentra el archivo de control de Excel en la carpeta del"
      " sistema."
  )
