import streamlit as st
import pandas as pd
import re
from collections import Counter
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Filtro de Pallets", page_icon="📊", layout="wide")

# --- LISTA DE CAMPOS ---
campos = [
    "Contenedor - Folio", "Folio", "N° Semana", "Fecha Análisis", "Fecha Etiqueta", "Analista", 
    "Turno", "Lote", "Cliente", "Tipo de producto", "Condición GF/convencional", 
    "Espesor inferior", "Espesor superrior", "% Humedad inferior FT", "% Humedad superior FT", 
    "Hora", "Cantidad sacos/maxisaco", "Peso saco/maxisaco", "Kilos producidos", "Humedad", 
    "Temperatura producto", "Enzimática", "Peso hectolitro", "Filamentos", "Cáscaras", 
    "Semillas Extrañas", "Gelatinas", "Quemadas", "Granos sin aplastar", 
    "Granos Parcialmente Aplastados", "Trigos", "Cebada", "Centeno", "Materiales extraños", 
    "Retención malla 7", "Bajo malla 25", "Espesor 1", "Espesor 2", "Espesor 3", 
    "Espesor 4", "Espesor 5", "Espesor 6", "Espesor 7", "Espesor 8", "Espesor 9", 
    "Espesor 10", "Promedio espesor", "Sacos detector de metales", 
    "Verificación de patrones PCC", "ESTADO", "Motivo Retención"
]

# --- FUNCIÓN DE DETECCIÓN (Misma lógica) ---
def detectar_patron_inteligente(texto_sucio):
    texto_sin_fechas = re.sub(r'\d{1,2}/\d{1,2}/\d{2,4}', '', texto_sucio)
    candidatos = re.findall(r'\b\d{10,14}\b', texto_sin_fechas)
    
    if not candidatos:
        return None, None
    
    prefijos = [c[:4] for c in candidatos]
    sufijos = [c[-2:] for c in candidatos]

    comun_prefix = Counter(prefijos).most_common(1)[0][0]
    comun_suffix = Counter(sufijos).most_common(1)[0][0]

    # Usamos rf string
    patron_generado = rf"{comun_prefix}(\d+?){comun_suffix}"
    
    return patron_generado, len(candidatos)

# --- INTERFAZ DE USUARIO ---
st.title("📊 Generador de Reportes de Hojuela")
st.markdown("Sube el archivo maestro, ingresa el contenedor y pega los pallets desordenados.")

# 1. Subir Archivo Maestro
archivo_maestro = st.file_uploader("📂 Cargar 'Control de producto planta Coihue 2026.xlsx'", type=["xlsx"])

col1, col2 = st.columns(2)

with col1:
    contenedor = st.text_input("📦 Identificador Contenedor", placeholder="Ej: MNBU123456")

with col2:
    pallets = st.text_area("📝 Lista de Pallets (Pegar texto sucio)", height=150, placeholder="Pega aquí el texto copiado del correo o sistema...")

# --- BOTÓN DE PROCESAR ---
if st.button("🚀 Procesar y Generar Excel"):
    if not archivo_maestro:
        st.error("⚠️ Por favor, sube primero el archivo Excel maestro.")
    elif not contenedor:
        st.error("⚠️ Falta el número de contenedor.")
    elif not pallets:
        st.error("⚠️ No has ingresado la lista de pallets.")
    else:
        try:
            with st.spinner('Leyendo base de datos...'):
                # Cargar Excel
                df_hojuelaavena = pd.read_excel(archivo_maestro, sheet_name="HOJUELA", header=1)
            
            st.success("Base de datos cargada. Analizando patrones...")
            
            # Detectar Patrón
            patron, num_candidatos = detectar_patron_inteligente(pallets)
            
            if patron:
                st.info(f"Patrón detectado en {num_candidatos} códigos. Regex: `{patron}`")
                
                lista_limpia = re.findall(patron, pallets)
                lista_int = [int(x) for x in lista_limpia]
                lista_int.sort()
                
                filas_encontradas = []
                coincidencias = 0
                
                # Barra de progreso
                barra = st.progress(0)
                total_items = len(lista_int)

                for idx, folio_buscado in enumerate(lista_int):
                    fila_match = df_hojuelaavena[df_hojuelaavena["Folio"] == folio_buscado]
                    
                    if not fila_match.empty:
                        coincidencias += 1
                        datos_fila = fila_match.iloc[0].to_dict()
                        datos_fila["Contenedor - Folio"] = f"{contenedor} - {folio_buscado}"
                        filas_encontradas.append(datos_fila)
                    
                    # Actualizar barra
                    barra.progress((idx + 1) / total_items)
                
                st.write(f"**Resultados:** {coincidencias} coincidencias de {total_items} códigos buscados.")

                # Generar Excel en memoria
                if filas_encontradas:
                    df_exportar = pd.DataFrame(filas_encontradas)
                    df_final = df_exportar.reindex(columns=campos)
                    
                    # Mostrar vista previa
                    st.dataframe(df_final.head())
                    
                    # Crear buffer en memoria (BytesIO)
                    output = io.BytesIO()
                    
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Reporte')
                        worksheet = writer.sheets['Reporte']
                        
                        # Formatos avanzados (Filtros, Paneles, Ancho)
                        worksheet.auto_filter.ref = worksheet.dimensions
                        worksheet.freeze_panes = 'B2'
                        
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = (max_length + 2)
                            worksheet.column_dimensions[column_letter].width = adjusted_width
                    
                    # Obtener valor del buffer para descargar
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 Descargar Reporte Excel",
                        data=excel_data,
                        file_name=f"Reporte_Contenedor_{contenedor}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                else:
                    st.warning("No se encontraron coincidencias en la base de datos con los folios extraídos.")
            else:
                st.error("No se pudo detectar un patrón de folios válido en el texto ingresado.")
                
        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
