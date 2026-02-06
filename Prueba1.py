import numpy as np
import pandas as pd
import re
from collections import Counter

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
control_producto = "Control de producto/Control de producto planta Coihue 2026.xlsx"
print("ingresar identificador contenedor")
contenedor = str(input())
print("ingresar lista de pallets")
pallets = input()
coincidencias = 0

df_hojuelaavena = pd.read_excel(io = control_producto, sheet_name = "HOJUELA", header = 1)

def detectar_patron_inteligente(texto_sucio):
    texto_sin_fechas = re.sub(r'\d{1,2}/\d{1,2}/\d{2,4}', '', texto_sucio)
    candidatos = re.findall(r'\b\d{10,14}\b', texto_sin_fechas)
    if not candidatos:
        return None, None
    prefijos = [c[:4] for c in candidatos]
    sufijos = [c[-2:] for c in candidatos]

    comun_prefix = Counter(prefijos).most_common(1)[0][0]
    comun_suffix = Counter(sufijos).most_common(1)[0][0]

    patron_generado = rf"{comun_prefix}(\d+?){comun_suffix}"
    
    print(f"Detectados {len(candidatos)} números candidatos.")
    print(f"Patrón dominante identificado: Empieza con '{comun_prefix}' y termina con '{comun_suffix}'")
    print(f"Regex generado: {patron_generado}\n")
    
    return patron_generado

print(df_hojuelaavena.head())
patron = detectar_patron_inteligente(pallets)
lista_limpia = re.findall(patron, pallets)
lista_int = [int(x) for x in lista_limpia]
lista_int.sort()

filas_encontradas = []
coincidencias = 0

for folio_buscado in lista_int:
    fila_match = df_hojuelaavena[df_hojuelaavena["Folio"] == folio_buscado]
    if not fila_match.empty:
        coincidencias += 1
        datos_fila = fila_match.iloc[0].to_dict()
        datos_fila["Contenedor - Folio"] = f"{contenedor} - {folio_buscado}"
        filas_encontradas.append(datos_fila)
        print(f"Encontrado: Folio {folio_buscado}")
    else:
        print(f"NO Encontrado: Folio {folio_buscado}")

print(f"Total procesados: {len(lista_int)}")
print(f"Total coincidencias encontradas: {coincidencias} / {len(lista_int)}")

if filas_encontradas:
    df_exportar = pd.DataFrame(filas_encontradas)
    df_final = df_exportar.reindex(columns=campos)
    nombre_archivo_salida = f"Reporte_Contenedor_{contenedor}.xlsx"
    try:
        with pd.ExcelWriter(nombre_archivo_salida, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Reporte')
                worksheet = writer.sheets['Reporte']
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
        print(f"¡Éxito! Archivo guardado con filtros y formato: {nombre_archivo_salida}")
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")
else:

    print("No se encontraron coincidencias para generar el archivo.")
