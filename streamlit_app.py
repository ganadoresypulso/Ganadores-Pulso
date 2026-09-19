import streamlit as st
import pandas as pd
import openpyxl
import os
from datetime import datetime

st.set_page_config(page_title="Campeonato de Marcas - Ganadores & Pulso", layout="wide")

st.title("🏇 Campeonato de Marcas - Ganadores & Pulso")
st.markdown("### Portal de Participantes y Carga de Marcas")

excel_file = "CONTROL CAMPEONATO DE MARCAS.xlsx"

if os.path.exists(excel_file):
    try:
        # Cargamos el libro con openpyxl para leer coordenadas exactas
        wb_lee = openpyxl.load_workbook(excel_file, data_only=True)
        if "CONTROL" not in wb_lee.sheetnames:
            st.error("No se encontró la hoja 'CONTROL' en el archivo de Excel.")
            st.stop()
            
        ws_lee = wb_lee["CONTROL"]
        
        # Leemos los encabezados de la Fila 3, pero limitamos estrictamente hasta la columna R (Columna 18)
        max_col_permitida = 18 # Columna R
        headers = [ws_lee.cell(row=3, column=col).value for col in range(1, max_col_permitida + 1)]
        headers_limpios = [str(h).strip() if h is not None else f"Col_{i+1}" for i, h in enumerate(headers)]
        
        # Leemos los datos desde la Fila 4 hasta la Fila 53 (o el máximo real si fuera menor)
        max_fila_permitida = min(53, ws_lee.max_row)
        datos_filas = []
        participantes = []
        
        for r in range(4, max_fila_permitida + 1):
            val_item = ws_lee.cell(row=r, column=1).value   # Columna A (ITEM)
            val_codigo = ws_lee.cell(row=r, column=2).value # Columna B (Código)
            val_nombre = ws_lee.cell(row=r, column=3).value # Columna C (Nombre)
            
            # Si el código o el nombre están vacíos, saltamos la fila para evitar filas en blanco
            if val_codigo is None or str(val_codigo).strip() == "" or str(val_codigo).lower() in ['none', 'nan']:
                continue
            if val_nombre is None or str(val_nombre).strip() == "" or str(val_nombre).lower() in ['none', 'nan']:
                continue
                
            # Construimos la fila acotada exactamente hasta la columna R (18)
            fila_valores = [ws_lee.cell(row=r, column=c).value for c in range(1, max_col_permitida + 1)]
            
            # Limpiamos valores None para que se vean ordenados en la tabla
            fila_valores_limpios = ["" if v is None else v for v in fila_valores]
            datos_filas.append(fila_valores_limpios)
            
            # Preparamos el selector para el menú
            item_str = str(val_item).strip() if val_item is not None else "S/I"
            codigo_str = str(val_codigo).strip()
            nombre_str = str(val_nombre).strip()
            participantes.append(f"{item_str} - {codigo_str} - {nombre_str}")
            
        # Creamos el DataFrame final limpio hasta la columna R y fila 53
        df_final = pd.DataFrame(datos_filas, columns=headers_limpios)

        st.sidebar.header("🔐 Menú de Navegación")
        modo = st.sidebar.radio("Seleccione una opción:", ["Ver Tabla General", "Cargar Mis Marcas"])
        
        if modo == "Ver Tabla General":
            st.subheader("📋 Listado Oficial de Participantes")
            st.dataframe(df_final, use_container_width=True, hide_index=True)
            
        elif modo == "Cargar Mis Marcas":
            st.subheader("✍️ Módulo de Ingreso de Marcas por Participante")
            
            if participantes:
                participante_seleccionado = st.selectbox("Seleccione su Participante:", participantes)
                # Extraemos el nombre limpio para buscarlo en las hojas de jornadas
                partes = participante_seleccionado.split(" - ")
                nombre_limpio = partes[-1] if len(partes) >= 3 else participante_seleccionado
                
                st.markdown("---")
                st.markdown("#### ⚙️ Configuración de la Jornada")
                
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    fecha_reunion = st.date_input("Fecha de la Reunión (Domingo):", value=datetime.today())
                with col_f2:
                    num_carreras = st.slider("Cantidad de carreras programadas:", min_value=6, max_value=16, value=14)
                
                fecha_str = fecha_reunion.strftime("%d-%m-%Y")
                st.info(f"Participante activo: **{participante_seleccionado}** | Reunión: **{fecha_str}** | Carreras: **{num_carreras}**")
                
                st.markdown("---")
                st.markdown("### 📌 Reglamento de Marcas:")
                st.markdown("- **Carreras normales:** 3 marcas (1ra, 2da y 3ra).")
                st.markdown("- **Súper Fijo (SF):** Solo se marca **1 ejemplar** en la carrera seleccionada.")
                st.markdown("- **Fijas (F):** Una en las carreras no válidas y otra en las válidas del 5y6.")
                
                with st.form("form_marcas"):
                    marcas_por_carrera = {}
                    
                    for i in range(1, num_carreras + 1):
                        st.markdown(f"#### 🏁 Carrera {i}")
                        
                        col_m1, col_m2, col_m3, col_m4 = st.columns([2, 2, 2, 2])
                        
                        with col_m1:
                            m1 = st.text_input(f"1ra Marca C-{i}", key=f"c{i}_m1")
                        with col_m2:
                            m2 = st.text_input(f"2da Marca C-{i}", key=f"c{i}_m2")
                        with col_m3:
                            m3 = st.text_input(f"3ra Marca C-{i}", key=f"c{i}_m3")
                        with col_m4:
                            tipo_jugada = st.selectbox(f"Tipo C-{i}", ["Normal", "Fijo (F)", "Súper Fijo (SF)"], key=f"c{i}_tipo")
                        
                        marcas_por_carrera[i] = {
                            "m1": m1, "m2": m2, "m3": m3, "tipo": tipo_jugada
                        }
                        st.markdown("---")
                    
                    clave_seguridad = st.text_input("Ingrese su Clave de Seguridad:", type="password")
                    boton_guardar = st.form_submit_button("Guardar y Enviar Marcas")
                    
                    if boton_guardar:
                        if clave_seguridad:
                            sf_count = sum(1 for c in marcas_por_carrera.values() if c["tipo"] == "Súper Fijo (SF)")
                            
                            if sf_count > 1:
                                st.error("⚠️ Solo puedes seleccionar un (1) Súper Fijo (SF) por reunión.")
                            else:
                                error_sf = False
                                for num, datos in marcas_por_carrera.items():
                                    if datos["tipo"] == "Súper Fijo (SF)" and (datos["m2"].strip() != "" or datos["m3"].strip() != ""):
                                        st.error(f"⚠️ En la Carrera {num} seleccionaste Súper Fijo (SF), la 2da y 3ra marca deben estar vacías.")
                                        error_sf = True
                                        break
                                        
                                if not error_sf:
                                    try:
                                        wb = openpyxl.load_workbook(excel_file)
                                        
                                        if fecha_str in wb.sheetnames:
                                            ws = wb[fecha_str]
                                        else:
                                            ws = wb.create_sheet(title=fecha_str)
                                            cabecera = ["NOMBRE PARTICIPANTE"]
                                            for c in range(1, 17):
                                                cabecera.extend([f"C{c}_1ra", f"C{c}_2da", f"C{c}_3ra", f"C{c}_Tipo"])
                                            ws.append(cabecera)
                                        
                                        fila_encontrada = None
                                        for row_idx in range(4, ws.max_row + 1):
                                            val_celda = ws.cell(row=row_idx, column=3).value
                                            if val_celda and str(val_celda).strip().upper() == nombre_limpio.strip().upper():
                                                fila_encontrada = row_idx
                                                break
                                        
                                        datos_fila = [nombre_limpio]
                                        for c in range(1, 17):
                                            if c <= num_carreras:
                                                m_data = marcas_por_carrera[c]
                                                tipo_corto = "N"
                                                if m_data["tipo"] == "Fijo (F)":
                                                    tipo_corto = "F"
                                                elif m_data["tipo"] == "Súper Fijo (SF)":
                                                    tipo_corto = "SF"
                                                    
                                                datos_fila.extend([m_data["m1"], m_data["m2"], m_data["m3"], tipo_corto])
                                            else:
                                                datos_fila.extend(["", "", "", ""])
                                        
                                        if fila_encontrada:
                                            for col_idx, val in enumerate(datos_fila, start=1):
                                                ws.cell(row=fila_encontrada, column=col_idx, value=val)
                                        else:
                                            ws.append(datos_fila)
                                        
                                        wb.save(excel_file)
                                        st.success(f"¡Marcas de **{nombre_limpio}** guardadas con éxito en la hoja '{fecha_str}'!")
                                    except Exception as err:
                                        st.error(f"Error al escribir en el Excel: {err}")
                        else:
                            st.error("Por favor ingrese su clave de seguridad.")
            else:
                st.warning("No se encontraron participantes válidos en las filas a partir de la fila 4.")
        
    except Exception as e:
        st.error(f"Error al procesar el archivo de Excel: {e}")
else:
    st.error("No se encuentra el archivo de Excel en la carpeta.")
