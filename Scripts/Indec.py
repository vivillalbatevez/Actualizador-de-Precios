#!/usr/bin/env python
# coding: utf-8

# In[32]:


import os
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
import pandas as pd
import time
import zipfile
import urllib.request
import re
import shutil


# # Rutas

# In[33]:


# Rutas relativas para producto final (ajusta si es necesario en el futuro)
#Ruta del driver
driver_path = os.path.join(os.path.dirname(__file__), "../WebDriver/msedgedriver.exe")

#Ruta del archivo de inputs
input_user_file = os.path.join(os.path.dirname(__file__), "../Input/Input.xlsx")

#Ruta del archivo de outputs
IPC_output_file = os.path.join(os.path.dirname(__file__), "../Intermedio/IPC_INDEC.xlsx")

IPIM_output_file = os.path.join(os.path.dirname(__file__), "../Intermedio/IPIM_INDEC.xlsx")


# #Rutas absolutas para produccion
# #Ruta del driver
# driver_path = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\WebDriver\msedgedriver.exe"
# 
# #Ruta del archivo de inputs
# input_user_file = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Input\Input.xlsx"  # Archivo proporcionado por el usuario
# 
# #Ruta del archivo de outputs
# IPC_output_file = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Intermedio\IPC_INDEC.xlsx"  # Archivo final de salida
# IPIM_output_file = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Intermedio\IPIM_INDEC.xlsx"  # Archivo final de salida

# In[35]:


#Ruta de la carpeta basanodse en driver path
Driverfolder_path = os.path.dirname(driver_path)


# # Funciones

# In[36]:


# Función para obtener la versión de Microsoft Edge en Windows
def get_edge_version():
    try:
        # Ejecuta un comando para obtener la versión de Edge
        command = r'reg query "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon" /v version'
        stream = os.popen(command)
        output = stream.read()

        # Filtra la versión desde la salida del comando
        for line in output.splitlines():
            if "version" in line:
                version = line.split()[-1]
                return version
    except Exception as e:
        print(f"Error al obtener la versión de Edge: {e}")
        return None


# In[37]:


def procesar_y_guardar_archivo(df, input_user_file, output_file):
    # 1. Extraer las fechas correctas desde la primera fila
    fechas = df.iloc[0, 1:].values  # Extraer las fechas desde la primera fila, ignorando la primera columna
    fechas = pd.to_datetime(fechas).to_period('M')  # Convertir las fechas al formato YYYY-MM

    # Verificar que las fechas sean correctas
    print("\nFechas extraídas del DataFrame original:")
    print(fechas)

    # 2. Extraer las filas correspondientes a "Nivel general" y "Alimentos y bebidas no alcohólicas"
    nivel_general = df.iloc[2, 1:].values  # Fila correspondiente a "Nivel general" (fila 9)
    alimentos_bebidas = df.iloc[3, 1:].values  # Fila correspondiente a "Alimentos y bebidas no alcohólicas" (fila 10)

    print("\nNivel general:")
    print(nivel_general)
    print("\nAlimentos y bebidas no alcohólicas:")
    print(alimentos_bebidas)

    # 3. Leer los meses proporcionados por el usuario
    user_df = pd.read_excel(input_user_file, sheet_name = "Fechas a act")
    print("\nDatos proporcionados por el usuario:")
    print(user_df.head())
    user_df['Fechas'] = user_df['Fechas'] - pd.DateOffset(months=1)
    

    # Convertir la columna 'fecha' a formato de fecha YYYY-MM
    user_df['Fechas'] = pd.to_datetime(user_df['Fechas']).dt.to_period('M')
    print("\nMeses proporcionados por el usuario después de la conversión:")
    print(user_df['Fechas'])

    # 4. Filtrar los datos que coinciden con los meses proporcionados por el usuario
    # Creamos un DataFrame temporal con las fechas, "Nivel general" y "Alimentos y bebidas"
    datos = pd.DataFrame({
        'Fechas': fechas,
        'Nivel general': nivel_general,
        'Alimentos y bebidas no alcohólicas': alimentos_bebidas
    })

    # Hacemos un merge entre los meses del usuario y los datos filtrados
    resultado = pd.merge(user_df, datos, on='Fechas', how='inner')

    # Verificar si hay coincidencias en el `merge`
    if resultado.empty:
        print("\nNo se encontraron coincidencias de meses entre los datos.")
    else:
        print("\nDatos tras el merge:")
        print(resultado)
        
    #Ajustar formato fecha al deseado
    resultado['Fechas'] = resultado['Fechas'].dt.to_timestamp().dt.strftime('%d/%m/%Y')
    
    # Convertir la columna de porcentaje a decimal
    resultado['Nivel general'] = resultado['Nivel general'] / 100
    resultado['Alimentos y bebidas no alcohólicas'] = resultado['Alimentos y bebidas no alcohólicas'] / 100

    # 5. Guardar el archivo de salida con los meses filtrados en el directorio deseado
    resultado.to_excel(output_file, index=False)
    print(f"\nArchivo final procesado y guardado en: {output_file}")


# In[38]:


def calcular_variacion_mensual(IPIM_df, input_user_file, output_file):
    # Leer las fechas desde el archivo Excel
    user_df = pd.read_excel(input_user_file, sheet_name="Fechas a act")
    user_df['Fechas'] = user_df['Fechas'] - pd.DateOffset(months=1)
    fechas = user_df.iloc[:, 0].tolist() 
    
    # Filtrar el DataFrame para las descripciones específicas
    IPIM_df_filtrado = IPIM_df[IPIM_df[('Descripción-Unnamed: 1_level_1')].isin(['Nivel general', ' Alimentos y bebidas'])]
    
    # Crear un nuevo DataFrame para almacenar los resultados
    resultados = []

    # Mapeo de los meses en español
    meses = {
        'Jan': 'Ene', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Abr',
        'May': 'May', 'Jun': 'Jun', 'Jul': 'Jul', 'Aug': 'Ago',
        'Sep': 'Sep', 'Oct': 'Oct', 'Nov': 'Nov', 'Dec': 'Dic'
    }

    # Calcular la variación mensual para cada concepto
    for index, row in IPIM_df_filtrado.iterrows():
        for fecha in fechas:
            fecha_str = str(fecha)  # Convertir a string
            anio, mes, _ = fecha_str.split('-')  # Dividir la fecha en anio, mes y dia
            mes_nombre = meses[pd.to_datetime(fecha_str).strftime('%b')]  # Obtener el nombre del mes en español
            mes_columna = f"{anio}-{mes_nombre}"

            if mes_columna in IPIM_df.columns:
                # Calcular el mes anterior
                mes_anterior_date = pd.to_datetime(fecha_str) - pd.DateOffset(months=1)
                mes_anterior_nombre = meses[mes_anterior_date.strftime('%b')]  # Mes anterior en español
                mes_anterior = f"{mes_anterior_date.year}-{mes_anterior_nombre}"

                if mes_anterior in IPIM_df.columns:
                    variacion = (row[mes_columna] - row[mes_anterior]) / row[mes_anterior] if row[mes_anterior] != 0 else None #variacion en decimal
                    resultados.append({
                        'fecha': mes_columna,
                        'concepto': row[('Descripción-Unnamed: 1_level_1')],
                        'variación': variacion
                    })
                else:
                    print(f"Columna no encontrada: {mes_anterior}")
            else:
                print(f"Columna no encontrada: {mes_columna}")

    # Convertir los resultados en un DataFrame
    resultados_df = pd.DataFrame(resultados)

    # Pivotar el DataFrame para obtener el formato deseado
    pivot_df = resultados_df.pivot_table(index='fecha', columns='concepto', values='variación')

    # Resetear el índice para que 'fecha' sea una columna normal
    pivot_df.reset_index(inplace=True)
    
    #Ajustar formato fecha al deseado
    pivot_df['fecha'] = pivot_df['fecha'].str.strip()
    # Diccionario para reemplazar los nombres de los meses en español por sus equivalentes en inglés
    meses = {
    'Ene': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Abr': 'Apr',
    'May': 'May', 'Jun': 'Jun', 'Jul': 'Jul', 'Ago': 'Aug',
    'Sep': 'Sep', 'Oct': 'Oct', 'Nov': 'Nov', 'Dic': 'Dec'
    }

    # Reemplazar los nombres de los meses en la columna 'fecha'
    pivot_df['fecha'] = pivot_df['fecha'].replace(meses, regex=True)
    
    pivot_df['fecha'] = pd.to_datetime(pivot_df['fecha'], format='%Y-%b', errors='coerce')
    # Convertir las fechas al formato 'día/mes/año'
    pivot_df['fecha'] = pivot_df['fecha'].dt.strftime('%d/%m/%Y')

    # Guardar los resultados en un archivo Excel
    pivot_df.to_excel(output_file, index=False)
    print(f"Resultados guardados en '{output_file}'.")


# In[39]:


def limpiar_carpeta(ruta_carpeta):
    try:
        # Verificar si la carpeta existe
        if os.path.exists(ruta_carpeta):
            # Eliminar todos los archivos y subcarpetas
            for archivo in os.listdir(ruta_carpeta):
                ruta_archivo = os.path.join(ruta_carpeta, archivo)
                # Si es un archivo, lo elimina
                if os.path.isfile(ruta_archivo) or os.path.islink(ruta_archivo):
                    os.unlink(ruta_archivo)
                # Si es una carpeta, la elimina junto con su contenido
                elif os.path.isdir(ruta_archivo):
                    shutil.rmtree(ruta_archivo)
            print(f"Se ha limpiado la carpeta: {ruta_carpeta}")
        else:
            print(f"La carpeta {ruta_carpeta} no existe.")
    except Exception as e:
        print(f"Ocurrió un error al limpiar la carpeta: {e}")


# # Web scraping con selenium

# In[40]:


limpiar_carpeta(Driverfolder_path)


# #### Obtener la version de edge y decargar los drivers

# In[41]:


# Obtener la versión de Edge instalada en el sistema
edge_version = get_edge_version()
print(f"Versión detectada de Microsoft Edge: {edge_version}")


# In[ ]:





# In[42]:


# Si el controlador no existe, descargarlo
if not os.path.exists(driver_path):
    # URL directa de descarga para el EdgeDriver (versión específica o adaptativa según tu versión de Edge)
    edge_driver_url = f"https://msedgedriver.azureedge.net/{edge_version}/edgedriver_win64.zip"
    zip_path = os.path.join(Driverfolder_path, "edgedriver.zip")

    # Descargar el archivo zip
    urllib.request.urlretrieve(edge_driver_url, zip_path)

    # Extraer el contenido del archivo zip
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(Driverfolder_path)

    # Elimina el archivo zip después de extraerlo
    os.remove(zip_path)


# #### Configurar las opciones de edge y activar el driver

# In[43]:


# Configuración de las opciones para Edge
edge_options = Options()
edge_options.add_argument("--headless")  # Opcional si deseas que se ejecute sin interfaz gráfica
edge_options.add_argument("--disable-gpu")  # Ayuda en entornos sin GPU
edge_options.add_argument("--no-sandbox")
edge_options.add_argument("--disable-dev-shm-usage")


# In[44]:


# Usa el EdgeDriver descargado manualmente
service = Service(driver_path)
driver = webdriver.Edge(service=service, options=edge_options)


# #### Navegamos a la pagina del indec que queremos, obtenemos com xpath donde esta el excel que buscamos y con el link hacemos un df en pandas

# In[45]:


# Navega a la página del INDEC
url = "https://www.indec.gob.ar/indec/web/Nivel4-Tema-3-5-31"
driver.get(url)


# In[46]:


# Espera que la página cargue completamente
time.sleep(5)


# In[47]:


# Encuentra el primer enlace al archivo .xls utilizando XPath
excel_link_element = driver.find_element(By.XPATH, "//a[contains(@href, '.xls')]")
excel_link = excel_link_element.get_attribute("href")


# In[48]:


#Ahora buscamos IPIM
url = "https://www.indec.gob.ar/indec/web/Nivel4-Tema-3-5-32"
driver.get(url)


# In[49]:


# Espera que la página cargue completamente
time.sleep(5)


# In[50]:


# Encuentra el primer enlace al archivo .xls utilizando XPath
IPIM_excel_link_element = driver.find_element(By.XPATH, "//a[contains(@href, '.xls')]")
IPIM_excel_link = IPIM_excel_link_element.get_attribute("href")


# In[51]:


# Cierra el navegador
driver.quit()


# In[52]:


# Mostrar el enlace del archivo Excel descargado
print(f"Enlace del archivo Excel IPC: {excel_link}")
print(f"Enlace del archivo Excel IPIM: {IPIM_excel_link}")


# In[53]:


# Leer el archivo Excel directamente desde la URL extraída
IPC_df = pd.read_excel(excel_link, sheet_name = "Variación mensual IPC Nacional", header = None)


# In[54]:


IPIM_df = pd.read_excel(IPIM_excel_link, sheet_name = "IPIM", header =[3, 4])


# In[55]:


# Mostrar las primeras filas del DataFrame para verificar la carga
IPC_df.head(6)


# In[56]:


IPIM_df.head(6)


# # Procesamiento de IPC DF para obtener indices que buscamos

# In[57]:


#creamos indices para identificar los encabezados de las tablas
region_keywords = ['Total nacional', 'Región GBA','Región Pampeana','Región Noroeste','Región Noreste','Región Cuyo','Región Patagonia']

#obtenemos el indice de las filas correspondientes a los encabezados
region_rows = IPC_df[IPC_df.apply(lambda row: any(keyword in str(cell) for cell in row for keyword in region_keywords), axis=1)]

region_rows


# In[58]:


#convertimos el resultado en una lista
region_indices = region_rows.index.tolist()

tables = []
#recorremos la lista para extraer las tablas
for i in range(len(region_indices)):
    start_idx = region_indices[i]
    end_idx = region_indices[i+1] if i+1 < len(region_indices) else len(IPC_df)

    #estraemos la tabla entre los indices y eliminamos las filas vacias
    table = IPC_df.iloc[start_idx:end_idx].dropna(how='all')
    tables.append(table)

IPC_df_nacional = tables[0]


# # Procesamiento IPIM

# In[59]:


# Crear un nuevo índice de fechas
new_columns = []
for year, month in IPIM_df.columns:
    if isinstance(year, int) and isinstance(month, str):
        month = re.sub(r'[^a-zA-Z]', '', month)  # Solo mantendremos letras
        new_columns.append(f"{year}-{month}")
    else:
        new_columns.append(f"{year}-{month}")

# Asignar el nuevo índice de columnas
IPIM_df.columns = new_columns

# Verificar las nuevas columnas
print(IPIM_df.columns)


# # ejecutar funciones

# In[60]:


calcular_variacion_mensual(IPIM_df, input_user_file, IPIM_output_file)


# In[61]:


# Ejecutar la función para procesar y guardar el archivo
procesar_y_guardar_archivo(IPC_df_nacional, input_user_file, IPC_output_file)

