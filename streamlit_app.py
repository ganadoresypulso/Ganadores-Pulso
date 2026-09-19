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
    # Cargar el libro de Excel para trabajar con él
    wb = openpyxl.load_workbook(archivo_excel)
    hojas = wb.sheetnames

    st.sidebar.header("Menú de Navegación")
    opcion = st.sidebar.selectbox(
        "Seleccione una opción:", ["Registro de Marcas", "Ver Mis Marcas"]
    )

    if opcion == "Registro de Marcas":
      st.markdown("### Listado Oficial de Participantes y Carga de Marcas")

      # Formulario de registro de marcas
      with st.form("form_marcas"):
        participante = st.text_input("Nombre del Participante / Marca:")
        clave = st.text_input("Clave de seguridad:", type="password")

        # Supongamos que pedimos una selección de ejemplares o carreras
        ejemplar = st.text_input("Ejemplar o Marca a Registrar:")

        submitted = st.form_submit_button("Enviar Marca")

        if submitted:
          if clave:  # Aquí puedes validar tu clave si lo deseas
            # Lógica para guardar en el Excel
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
    st.error(f"Error al procesar el archivo de Excel: {mi}")
else:
  st.error("No se encuentra el archivo de Excel en la carpeta.")

# ---------------------------------------------------------
# BOTÓN DE DESCARGA DIRECTA DEL EXCEL ACTUALIZADO
# ---------------------------------------------------------
st.markdown("---")
st.markdown("### 📥 Panel de Control - Descarga de Datos")

if os.path.exists(archivo_excel):
  with open(archivo_excel, "rb") as f:
    st.download_button(
        label="📥 Descargar Excel con las Marcas Actualizadas",
        data=f,
        file_name="CONTROL_CAMPEONATO_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
  st.warning(
      "El archivo de Excel aún no está disponible para descarga en este"
      " momento."
  )
# ---------------------------------------------------------
# PANEL DE CONTROL SEGURO (SOLO PARA EL DIRECTOR)
# ---------------------------------------------------------
st.markdown("---")
with st.expander("🔒 Acceso Administrativo (Solo Director)"):
  admin_key = st.text_input(
      "Ingrese clave de administrador:", type="password"
  )
  # Cambia 'tu_clave_secreta' por la contraseña que tú quieras usar
  if admin_key == "tu_clave_secreta":
    st.success("¡Acceso concedido!")
    if os.path.exists(archivo_excel):
      with open(archivo_excel, "rb") as f:
        st.download_button(
            label="📥 Descargar Excel con las Marcas Actualizadas",
            data=f,
            file_name="CONTROL_CAMPEONATO_ACTUALIZADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
      st.warning("El archivo de Excel aún no está disponible.")
  elif admin_key:
    st.error("Clave incorrecta.")
