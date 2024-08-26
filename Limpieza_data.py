import pandas as pd
import datetime
import re
import tkinter as tk
from tkinter import filedialog
#____________________________________________________________________________________________#
# Cargar los tres archivos Excel NOKIA
file_path_onair = r'C:\Users\ivanf\Downloads\sites_list_onair.xlsx'
df_onair = pd.read_excel(file_path_onair, sheet_name='Export')

file_path_seguimiento_5G = r'C:\Users\ivanf\Downloads\onair_seguimiento_5g.xlsx'
df_seguimiento_5G = pd.read_excel(file_path_seguimiento_5G, sheet_name='Export')

# Cargar archivos Excel UMBRELLA
def select_file():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    file_path_UMB = filedialog.askopenfilename(title="Selecciona un archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path_UMB

file_path_UMB = select_file()
if file_path_UMB:
    df_umb = pd.read_excel(file_path_UMB)
else:
    print("No se ha seleccionado ningún archivo")
#____________________________________________________________________________________________#
# Definir las posiciones de las columnas a eliminar
columnas_a_eliminar_onair = [0, 1, 7, 8, 9, 12, 17, 20] + list(range(35, 51)) 
columnas_a_eliminar_5G = [0, 4, 12, 13, 14, 24, 25, 26] + list(range(28, 41))    
columnas_a_eliminar_UMB = [0, 1, 4, 6, 7, 8, 10, 11, 12, 13, 14, 15, 17, 18, 21, 22, 23, 25, 26, 27, 28, 29, 30, 34,]        
# Eliminar columnas de los tres dataframes
df_onair.drop(df_onair.columns[columnas_a_eliminar_onair], axis=1, inplace=True)
df_seguimiento_5G.drop(df_seguimiento_5G.columns[columnas_a_eliminar_5G], axis=1, inplace=True)
df_umb.drop(df_umb.columns[columnas_a_eliminar_UMB], axis=1, inplace=True)
#____________________________________________________________________________________________#
#Organizar DATA sites OnAir 

# Nombre específico en la columna "Site Name" que deseas eliminar
site_name_to_remove = 'BCA.Terminal'
df = df_onair[df_onair['Site Name'] != site_name_to_remove]

# Extraer el año y el número de la semana de la cadena '2024_Wxx'

df['Year'] = df['W OnAir'].str.extract(r'(\d{4})', expand=False).astype(float)
df['Week'] = df['W OnAir'].str.extract(r'W(\d+)', expand=False).astype(float)

# Obtener la semana actual y la semana anterior del año actual
current_year, actual_week = datetime.datetime.now().isocalendar()[:2]
current_week = actual_week + 1
previous_week = current_week - 1

# Filtrar el DataFrame para dejar solo las filas sin contenido en 'W OnAir' y las filas de la semana actual o anterior
df_filtered = df[(df['W OnAir'].isna()) | 
                ((df['Year'] == current_year) & (df['Week'].isin([current_week, previous_week])))].copy()

# Eliminar las columnas temporales 'Year' y 'Week'
df_filtered.drop(columns=['Year', 'Week'], inplace=True)

# Crear la columna 'Condicion ODH'
def determine_condicion(row):
    if row['Proyecto'] == 'ODH_Nuevos':
        comentario = str(row['Comentario']).lower()
        if 'provisional' in comentario or 'temporal' in comentario:
            return 'Provisional'
        else:
            return 'Definitivo'
    return ''

df_filtered['Condicion ODH'] = df_filtered.apply(determine_condicion, axis=1)

# Mover el código inicial a 'OT OnAir' si el comentario no empieza con una fecha y luego borrar 'Comentario'
def move_code_to_ot_onair(row):
    comentario = str(row['Comentario'])
    if re.match(r'^\d{2}/\d{2}/\d{4}', comentario):
        return row['OT OnAir'], comentario
    else:
        codigo = comentario.split(' ')[0]
        return codigo, ''

df_filtered['OT OnAir'], df_filtered['Comentario'] = zip(*df_filtered.apply(move_code_to_ot_onair, axis=1))

# Eliminar la columna 'Comentario'
df_filtered.drop(columns=['Comentario'], inplace=True)

# Borrar los valores en 'OT OnAir' que no coincidan con el patrón similar a '1F154A-D549D7'
pattern = re.compile(r'^[A-Za-z0-9]+-[A-Za-z0-9]+$')
df_filtered['OT OnAir'] = df_filtered['OT OnAir'].apply(lambda x: x if pattern.match(str(x)) else '')

# Reemplazar la columna 'Integracion' con los valores de 'Integracion ACK'
df_filtered['Integracion'] = df_filtered['Integracion ACK']

# Convertir las columnas de fecha a datetime y eliminar la hora
df_filtered['Fecha Ult Cambio Est'] = pd.to_datetime(df_filtered['Fecha Ult Cambio Est'], errors='coerce', dayfirst=True).dt.normalize()
df_filtered['Integracion'] = pd.to_datetime(df_filtered['Integracion'], errors='coerce', dayfirst=True).dt.normalize()

# Verificar las conversiones de fecha
print("Tipos de datos después de convertir fechas:")
print(df_filtered.dtypes)

# Crear un diccionario para el mapeo de estado a owner
estado_owner_mapping = {
    "10. En Revisión Calidad (NI)": "NI",
    "11. Pend HW (NI)": "Claro",
    "12. Activación OnGoing (NI)": "NI",
    "20. Primera Revisión NPO": "Calidad",
    "21. En Revisión Optimización": "NPO",
    "22. Segunda Revisión NPO": "Calidad",
    "23. Revisión Caso Especial NPO": "NPO",
    "24. 5G Esperando Concepto Actividad Sinergia": "NPO",
    "31. Falla HW": "Aliado",
    "32. Alarmas": "Aliado",
    "33. Diferencia RTWP Entre Puertos": "Aliado",
    "34. Alto RTWP": "Aliado",
    "35. Otros KPIs": "Aliado",
    "36. Instalación/Integración": "Aliado",
    "37. Pend Reporte Radiante Aprobado": "Aliado",
    "40. Pend OT INT UMB": "DEC",
    "41. Pend OT Acceso UMB": "DEC",
    "42. Cargando Evidencias": "NI",
    "43. Cargando Evidencias": "NI",  # SSV NPO para 5G
    "45. Segunda Revisión SSV - Pend Espectro": "NPO",
    "46. OK SSV Pend Espectro": "NPO",
    "51. Falla Tx": "Claro",
    "52. Falla Energia": "Claro",
    "53. Falla HW Existente": "Claro",
    "54. Problema HSEQ": "Claro",
    "55. Riesgo Biológico": "Claro",
    "56. Problema Acceso": "Claro",
    "57. Problema Orden Público": "Claro",
    "58. Problema de RF Claro": "RF Claro",
    "60. Pend Revisión RF-NOC": "RF Claro",
    "61. Rechazado RF. Revisita": "Aliado",
    "62. Rechazado RF. Falla HW": "Aliado",
    "63. Rechazado RF. Optimización": "NPO",
    "64. Rechazado RF. Reiniciado": "RF Claro",
    "65. Pend Revisión NOC": "NOC",
    "66. Rechazado NOC. Falla HW": "Aliado",
    "67. Rechazado NOC. Revisita": "Aliado",
    "68. Rechazado NOC. Reiniciado": "NOC",
    "69. Pend Marcar RFTool": "NOC",
}

def determine_owner(row):
    sub_estado = str(row['Sub Estado Insrv']).strip()
    proyecto = str(row['Proyecto']).strip()
    condicion_od = row['Condicion ODH']
    
    # Caso especial para 43. Cargando Evidencias con 5G
    if sub_estado == "43. Cargando Evidencias" and proyecto == "5G":
        return "SSV NPO"
    
    # Caso especial para Condicion ODH Provisional
    if condicion_od == 'Provisional':
        return 'Claro'
    
    # Mapeo general de estado a owner
    if sub_estado in estado_owner_mapping:
        return estado_owner_mapping[sub_estado]
    
    return ''

# Aplicar la función y depurar
df_filtered['Owner'] = df_filtered.apply(determine_owner, axis=1)

# Crear la columna 'Aging NPO'
def calculate_aging_npo(row):
    if row['Sub Estado Insrv'] in ["21. En Revisión Optimización", "23. Revisión Caso Especial NPO", "24. 5G Esperando Concepto Actividad Sinergia", "63. Rechazado RF. Optimización"]:
        fecha_ult_cambio = row['Fecha Ult Cambio Est']
        if pd.notna(fecha_ult_cambio):
            return (pd.Timestamp.today().normalize() - fecha_ult_cambio).days
    return ''

df_filtered['Aging NPO'] = df_filtered.apply(calculate_aging_npo, axis=1)

# Crear la columna 'Aging Produccion' utilizando 'Integracion ACK'
def calculate_aging_produccion(row):
    if row['Sub Estado Insrv'] != "7. Producción":
        integracion = row['Integracion']
        if pd.notna(integracion):
            days_diff = (pd.Timestamp.today().normalize() - integracion).days
            return max(0, days_diff)  # Asegurarse de que no sea negativo
    return ''

df_filtered['Aging Produccion'] = df_filtered.apply(calculate_aging_produccion, axis=1)
df_filtered.loc[df_filtered['Sub Estado Insrv'] == "70. Producción", 'Aging Produccion'] = ''

# Verificar los resultados del cálculo
print("Resultados de 'Aging Produccion':")
print(df_filtered[['Sub Estado Insrv', 'Integracion', 'Aging Produccion']].head())

#Crear columna 'Aging Claro'
def calculate_aging_Claro(row):
    if row ['Sub Estado Insrv'] in ["51. Falla Tx", "52. Falla Energia", "53. Falla HW Existente", "54. Problema HSEQ", "55. Riesgo Biológico", "56. Problema Acceso", "57. Problema Orden Público", "58. Problema de RF Claro"]:
        fecha_ult_cambio_claro = row ['Fecha Ult Cambio Est']
        if pd.notna(fecha_ult_cambio_claro):
            return(pd.Timestamp.today().normalize() - fecha_ult_cambio_claro).days
        return ''
    
df_filtered['Aging Claro'] = df_filtered.apply(calculate_aging_Claro, axis=1)

#Crear columna 'Aging DEC'
def calculate_aging_Dec(row):
    if row ['Sub Estado Insrv'] in ["40. Pend OT INT UMB", "41. Pend OT Acceso UMB"]:
        fecha_ult_cambio_dec = row ['Fecha Ult Cambio Est']
        if pd.notna(fecha_ult_cambio_dec):
            return(pd.Timestamp.today().normalize() - fecha_ult_cambio_dec).days
        return ''
    
df_filtered['Aging DEC'] = df_filtered.apply(calculate_aging_Dec, axis=1)

# Crear la columna 'W Meta' con las probabilidades de paso a producción
def calculate_prob_insr(row):
    estado_insrv = str(row['Estado Insrv']).strip()
    sub_estado_insrv = str(row['Sub Estado Insrv']).strip()
    owner = row['Owner']
    probabilidad = 0

    if estado_insrv:
        if owner == "RF Claro":
            probabilidad = 80
        elif owner == "NOC":
            probabilidad = 90
        elif owner == "NPO":
            probabilidad = 60
        elif owner == "Aliado":
            probabilidad = 40
    if estado_insrv:
        if owner == "RF Claro":
            probabilidad = 80
        elif owner == "NOC":
            probabilidad = 90
        elif owner == "NPO":
            probabilidad = 60
        elif owner == "Aliado":
            probabilidad = 50

    if sub_estado_insrv == "41. Pend OT Acceso UMB" and owner == "DEC":
        probabilidad = max(probabilidad, 65)

    if sub_estado_insrv and owner == "RF Claro":
        probabilidad = max(probabilidad, 70)

    return probabilidad

df_filtered['Probabilidad InSrv en la semana'] = df_filtered.apply(calculate_prob_insr, axis=1)
df_filtered.loc[df_filtered['Sub Estado Insrv'] == "70. Producción", 'Probabilidad InSrv en la semana'] = ''
# Convertir la columna 'Probabilidad InSrv en la semana' a números, reemplazando cualquier valor no numérico con NaN
df_filtered['Probabilidad InSrv en la semana'] = pd.to_numeric(df_filtered['Probabilidad InSrv en la semana'], errors='coerce')

# Crear la columna 'META W' basada en la condición de la probabilidad
df_filtered['META W'] = df_filtered['Probabilidad InSrv en la semana'].apply(lambda x: 1 if pd.notna(x) and x > 60 else 0)
df_filtered.loc[df_filtered['Sub Estado Insrv'] == "70. Producción", 'META W'] = ''
#________________________________________________________________________________________________________________________________#
#________________________________________________________________________________________________________________________________#
#________________________________________________________________________________________________________________________________#

#Reorganización DATA de archivo UMB
df_umb_new = df_umb[
    df_umb['Plantilla'].str.startswith('Control') |
    df_umb['Plantilla'].isin(['OT_Acceso', 'OT_Integración Infraestructura', 'OT_Recepción Infraestructura Acceso'])
]
df_umb_new = df_umb_new.copy()
# Diccionario de mapeo UMB
mapping = {
    'OT_Acceso': 0,
    'OT_Integración Infraestructura': 0,
    'OT_Recepción Infraestructura Acceso': 0,
    'Control 1': 1,
    'Control 1.1': 1.1,
    'Control 1.2': 1.2,
    'Control RF_Nuevo RI': 2,
    'Control RF 1_Nuevo RI': 2.1,
    'Control RF 2_Nuevo RI': 2.2,
    'Control 2': 2,
    'Control 2.1': 2.1,
    'Control 2.2': 2.2,
    'Control NOC_Nuevo RI': 3,
    'Control NOC 1_Nuevo RI': 3.1,
    'Control NOC 2_Nuevo RI': 3.2,
    'Control 3': 3,
    'Control 3.1': 3.1,
    'Control 3.2': 3.2,
    'Control 4': 4,
    'Control 4.1': 4.1,
    'Control 4.2': 4.2
}
# Agregar una nueva columna con los valores mapeados UMB
df_umb_new.loc[:, 'Estatus numerico'] = df_umb_new['Plantilla'].map(mapping)
# Encontrar el valor máximo de "Estatus numerico" para cada combinación de "Nombre_Sitio" y "Proyecto" UMB
max_status_per_site_project = df_umb_new.groupby(['Nombre_Sitio', 'Proyecto'])['Estatus numerico'].transform(max)
# Filtrar filas que tienen el estatus numérico máximo para cada combinación de "Nombre_Sitio" y "Proyecto"
df_umb_max_new= df_umb_new[df_umb_new['Estatus numerico'] == max_status_per_site_project]
# Crear una nueva columna combinando "Plantilla" y "Estado" con una barra inclinada (/) UMB
df_umb_max_new['Plantilla_Estado'] = df_umb_max_new['Plantilla'] + '/' + df_umb_max_new['Estado']
# Reorganizar las columnas para colocar "Plantilla_Estado" al lado de "Plantilla"
cols = list(df_umb_max_new.columns)
plantilla_index = cols.index('Plantilla')
cols.insert(plantilla_index + 1, cols.pop(cols.index('Plantilla_Estado')))
df_umb_max_new = df_umb_max_new[cols]
df_umb_max_new.drop(columns=['Plantilla', 'Estado', 'Estatus numerico'], inplace=True)
#Cambiar nombre de columna 
df_umb_max_new.rename(columns={'Flujo_UUID': 'OT OnAir'}, inplace=True)
#____________________________________________________________________________________________#

# Unir los DataFrames Sities OnAir y UMS
merged_df = df_filtered.merge(df_umb_max_new, on='OT OnAir', how='outer')
merged_df.rename(columns={'SMP_x': 'SMP'}, inplace=True)
merged_df = merged_df.merge(df_seguimiento_5G, on = 'SMP', how= 'outer')
output_data_final= r'C:\Users\ivanf\Downloads\data_final.xlsx'
merged_df.to_excel(output_data_final, index=False)
#____________________________________________________________________________________________#

#____________________________________________________________________________________________#

#Validación individual de cada dataset

#output_onair = r'C:\Users\ivanf\Downloads\sites_list_onair_NEW.xlsx'
#output_seguimiento_5G = r'C:\Users\ivanf\Downloads\onair_seguimiento_5g_NEW.xlsx'
#output_UMB = r'C:\Users\ivanf\Downloads\UMBRELLA_NEW.xlsx'
#df_onair.to_excel(output_onair, index=False)
#df_seguimiento_5G.to_excel(output_seguimiento_5G, index=False)
#df_umb_max_new.to_excel(output_UMB, index=False)
#____________________________________________________________________________________________#