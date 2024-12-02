#!/usr/bin/env python
# coding: utf-8

# In[69]:


import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta
from datetime import datetime
import xlsxwriter
import os


# ## Rutas

# In[ ]:


# Rutas relativas para producto final (ajusta si es necesario en el futuro)
#Inputs
input_user_file = os.path.join(os.path.dirname(__file__), "../Input/Input.xlsx")
tabla_base_path = os.path.join(os.path.dirname(__file__), "../Config/Tabla_Base.xlsx")
IPC_output_file = os.path.join(os.path.dirname(__file__), "../Intermedio/IPC_INDEC.xlsx")
IPIM_output_file = os.path.join(os.path.dirname(__file__), "../Intermedio/IPIM_INDEC.xlsx")
porcentaje_act_file = os.path.join(os.path.dirname(__file__), "../Config/Porcentaje_ACT.xlsx")
tabla_gatillo_file_path = os.path.join(os.path.dirname(__file__), "../Config/Gatillo.xlsx")
tabla_diferentes_file_path = os.path.join(os.path.dirname(__file__), "../Config/Diferentes.xlsx")

# Rutas de archivos intermedios
intermedio_file_path = os.path.join(os.path.dirname(__file__), "../Intermedio/tabla_coeficiente.xlsx")
tabla_act_file_path = os.path.join(os.path.dirname(__file__), "../Intermedio/tabla_actualizacion.xlsx")

# Ruta del archivo de outputs
output_file_path = os.path.join(os.path.dirname(__file__), "../Output/tabla_Final.xlsx")


# # Inputs
# input_user_file = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Input\Input.xlsx"
# tabla_base_path = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Config\Tabla_Base.xlsx"
# IPC_output_file = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Intermedio\IPC_INDEC.xlsx"
# IPIM_output_file = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Intermedio\IPIM_INDEC.xlsx"
# porcentaje_act_file = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Config\Porcentaje_ACT.xlsx"
# tabla_gatillo_file_path = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Config\Gatillo.xlsx"
# tabla_diferentes_file_path = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Config\Diferentes.xlsx"
# 
# # Rutas de archivos intermedios
# intermedio_file_path = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Intermedio\tabla_coeficiente.xlsx"
# tabla_act_file_path = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Intermedio\tabla_actualizacion.xlsx"
# 
# # Ruta del archivo de outputs
# output_file_path = r"D:\MAXIMIA\PROYECTO ESTIMACIONES\Actualizador de precios\Output\tabla_Final.xlsx"

# In[71]:


# Leer los archivos de Excel
fecha_act = pd.read_excel(input_user_file, sheet_name='Fechas a act')
mano_obra = pd.read_excel(input_user_file, sheet_name='MO')
redeterminaciones = pd.read_excel(input_user_file, sheet_name='Redeterminaciones')
IPC = pd.read_excel(IPC_output_file)
IPIM = pd.read_excel(IPIM_output_file)
porcentaje_act = pd.read_excel(porcentaje_act_file)
tabla_base = pd.read_excel(tabla_base_path,sheet_name='Tabla_Base')
tabla_gatillo = pd.read_excel(tabla_gatillo_file_path)
tabla_diferentes = pd.read_excel(tabla_diferentes_file_path)


# In[72]:


# Convertir la columna de fechas a datetime
IPC['Fechas'] = pd.to_datetime(IPC['Fechas'], dayfirst=True)
IPIM['fecha'] = pd.to_datetime(IPIM['fecha'], dayfirst=True)
fecha_act['Fechas'] = pd.to_datetime(fecha_act['Fechas'], dayfirst=True)


# In[73]:


def obtener_valor_mano_obra(concepto):
    return mano_obra.loc[mano_obra['Conceptos'] == concepto, 'Act'].values[0] if concepto in mano_obra['Conceptos'].values else 0


# In[74]:


def obtener_valor_ipc(fecha, concepto):
    if concepto == 'IPC-GENERAL':
        valor = IPC.loc[IPC['Fechas'] == fecha, 'Nivel general']
    elif concepto == 'IPC-ALIM_BEB':
        valor = IPC.loc[IPC['Fechas'] == fecha, 'Alimentos y bebidas no alcohólicas']
    else:
        valor = pd.Series([0])
    return valor.values[0] if not valor.empty else 0


# In[75]:


def obtener_valor_ipim(fecha, concepto):
    if concepto == 'IPIM-GENERAL':
        valor = IPIM.loc[IPIM['fecha'] == fecha, 'Nivel general']
    elif concepto == 'IPIM-ALIM_BEB':
        valor = IPIM.loc[IPIM['fecha'] == fecha, 'Alimentos y bebidas no alcohólicas']
    else:
        valor = pd.Series([0])
    return valor.values[0] if not valor.empty else 0


# In[76]:


def calcular_coeficiente_actualizacion(cliente, ccosto):
    # Filtrar los registros correspondientes en porcentaje_act
    registros_cliente = porcentaje_act[(porcentaje_act['Cod cliente'] == cliente) & 
                                       (porcentaje_act['Codigo CC'] == ccosto)]
    
    resultados = []
    # Iterar sobre todas las fechas de IPC e IPIM
    fechas_ipc = IPC['Fechas'].unique()
    fechas_ipim = IPIM['fecha'].unique()
    todas_fechas = sorted(set(fechas_ipc).union(fechas_ipim))
    
    # Calcular el coeficiente de actualización para cada fecha
    for fecha in todas_fechas:
        porcentaje_actualizacion = 0
        
        for _, registro in registros_cliente.iterrows():
            concepto = registro['Concepto']
            peso = registro['Porcentaje']
            
            if 'MO' in concepto:
                valor = obtener_valor_mano_obra(concepto) #, fecha)
            elif 'IPC' in concepto:
                valor = obtener_valor_ipc(fecha, concepto)
            elif 'IPIM' in concepto:
                valor = obtener_valor_ipim(fecha, concepto)
            else:
                valor = 0
            
            porcentaje_actualizacion += peso * valor

            nombre_ccosto = registro['Ccosto']
        fecha_ajustada = (fecha + relativedelta(months=1)).strftime('%d/%m/%Y')
        
        # Guardar el coeficiente calculado junto con la fecha
        resultados.append({
            'Fecha': fecha_ajustada,
            'Código Cliente': cliente,
            'Código CC': ccosto,
            'Ccosto': nombre_ccosto,
            'Porcentaje de Actualización': porcentaje_actualizacion
        })
    
    # Convertir los resultados a DataFrame
    resultados_df = pd.DataFrame(resultados)
    
    return resultados_df


# In[77]:


nueva_tabla = []

# Calcular valores para cada combinación única de cliente, ccosto y col_apoyo
for (cliente, ccosto), group in porcentaje_act.groupby(['Cod cliente', 'Codigo CC']):
    servicios = group['Col apoyo'].unique()
    if len(servicios) == 1 and pd.isna(servicios[0]):
        # Caso cuando no hay 'Col apoyo' (solo un servicio por ccosto)
        porcentajes_df = calcular_coeficiente_actualizacion(cliente, ccosto)
        
        # Agregar cada registro de porcentaje y fecha en nueva_tabla
        for _, row in porcentajes_df.iterrows():
            nueva_tabla.append({
                'Código Cliente': cliente,
                'Nombre Cliente': group['Cliente'].iloc[0],
                'Código CC': ccosto,
                'Ccosto': row['Ccosto'],
                'Servicio': ' ',
                'Fecha': row['Fecha'],
                'Porcentaje de Actualización': row['Porcentaje de Actualización'],
                'Coeficiente de Actualización': row['Porcentaje de Actualización'] + 1
            })
    else:
        for col_apoyo in servicios:
            porcentajes_df = calcular_coeficiente_actualizacion(cliente, ccosto)
            
            for _, row in porcentajes_df.iterrows():
                nueva_tabla.append({
                    'Código Cliente': cliente,
                    'Nombre Cliente': group['Cliente'].iloc[0],
                    'Código CC': ccosto,
                    'Ccosto': row['Ccosto'],
                    'Servicio': col_apoyo,
                    'Fecha': row['Fecha'],
                    'Porcentaje de Actualización': row['Porcentaje de Actualización'],
                    'Coeficiente de Actualización': row['Porcentaje de Actualización'] + 1
                })

# Convertir a DataFrame
tabla_coeficiente = pd.DataFrame(nueva_tabla)

# Guardar la nueva tabla en un archivo Excel
tabla_coeficiente.to_excel(intermedio_file_path, index=False)

print("Tabla de actualización generada y guardada correctamente.")


# In[78]:


def calcular_precio_actualizado(tabla_coeficiente, tabla_base):
    # Convertir las fechas de la tabla base a datetime si no lo están
    tabla_base['Fecha'] = pd.to_datetime(tabla_base['Fecha'], dayfirst=True)
    tabla_coeficiente['Fecha'] = pd.to_datetime(tabla_coeficiente['Fecha'], dayfirst=True)

    resultados = []

    # Iterar sobre cada fila en la tabla_coeficiente
    for _, row in tabla_coeficiente.iterrows():
        cliente = row['Código Cliente']
        nombre_cliente = row['Nombre Cliente']
        ccosto = row['Código CC']
        nombre_ccosto = row['Ccosto']
        servicio = row['Servicio'] if row['Servicio'] not in [None, ''] else 'Sin servicio'
        fecha = row['Fecha']
        coeficiente = row['Coeficiente de Actualización']

        # Obtener el precio del mes anterior de la tabla base
        fecha_mes_anterior = fecha - relativedelta(months=1)

        # Asegúrate de que la columna de servicio se llame correctamente
        precio_anterior = tabla_base[
            (tabla_base['Codigo Cliente'] == cliente) & 
            (tabla_base['Cliente'] == nombre_cliente) &
            (tabla_base['Codigo Ccosto'] == ccosto) & 
            ((tabla_base['Col  apoyo'] == servicio) | (tabla_base['Col  apoyo'].isna()))  & 
            (tabla_base['Fecha'] == fecha_mes_anterior) &
            (tabla_base['Ccosto'] == nombre_ccosto)
        ]['Precio'].values
        
        if len(precio_anterior) > 0:
            precio_anterior = precio_anterior[0]
            precio_actualizado = precio_anterior * coeficiente
            
            # Agregar el resultado a la lista
            resultados.append({
                'Fecha': fecha.strftime('%d/%m/%Y'),
                'Código Cliente': cliente,
                'Cliente' : nombre_cliente,
                'Código CC': ccosto,
                'Ccosto': nombre_ccosto,
                'Servicio': servicio,
                'Precio Anterior': precio_anterior,
                'Coeficiente': coeficiente,
                'Precio Actualizado': precio_actualizado
            })
        else:
            # Si no se encuentra el precio anterior, se puede agregar un registro con el precio como None
            resultados.append({
                'Fecha': fecha.strftime('%d/%m/%Y'),
                'Código Cliente': cliente,
                'Cliente' : nombre_cliente,
                'Código CC': ccosto,
                'Ccosto': nombre_ccosto,
                'Servicio': servicio,
                'Precio Anterior': None,
                'Coeficiente': coeficiente,
                'Precio Actualizado': None
            })

    # Convertir los resultados a DataFrame
    resultados_df = pd.DataFrame(resultados)
    
    return resultados_df


# In[79]:


tabla_actualizacion = calcular_precio_actualizado(tabla_coeficiente, tabla_base)
# Ordenar la tabla de actualización por mes (fecha)
tabla_actualizacion = tabla_actualizacion.sort_values(by='Fecha').reset_index(drop=True)

tabla_actualizacion.to_excel(tabla_act_file_path, index=False)

if tabla_actualizacion.empty:
    print("No se encontraron precios actualizados.")
else:
    print("Precios actualizados generados y guardados correctamente.")


# # trabajando en la tabla final

# In[80]:


# Crear una copia inicial de la tabla base como punto de partida para la tabla final
tabla_final = tabla_base.copy()

# Formatear la columna 'Fecha' al formato '%d/%m/%Y'
tabla_final['Fecha'] = tabla_final['Fecha'].dt.strftime('%d/%m/%Y')


# In[81]:


def aplicar_logica_gatillo(precio_anterior, coeficiente_anterior, coeficiente_actual, sin_act_por_gatillo, cliente, tabla_gatillo):
    
    hubo_gatillo = False
    precio_actualizado = precio_anterior  

    if sin_act_por_gatillo:
        # Si no hay valores NaN, aplica actualización
        if not any(np.isnan([precio_anterior, coeficiente_anterior, coeficiente_actual])):
            precio_actualizado = (precio_anterior * coeficiente_anterior) * coeficiente_actual

    elif cliente in tabla_gatillo['Codigo cliente'].values:
        # Cliente en tabla_gatillo, aplica lógica de gatillo
        porcentaje_gatillo = float(tabla_gatillo.loc[tabla_gatillo['Codigo cliente'] == cliente, 'Gatillo'].values[0])
        porcentaje_actualizacion = coeficiente_actual - 1

        if porcentaje_actualizacion >= porcentaje_gatillo:
            precio_actualizado = precio_anterior * coeficiente_actual
        else:
            hubo_gatillo = True

    else:
        # Cliente no está en tabla_gatillo, aplica actualización por defecto
        precio_actualizado = precio_anterior * coeficiente_actual

    return precio_actualizado, hubo_gatillo


# In[82]:


def actualizar_precio_si_corresponde(codigo_cliente, meses_sin_act_por_contrato, tabla_diferentes):
    
    # Filtrar la tabla para obtener la demora para el cliente específico
    fila_cliente = tabla_diferentes[tabla_diferentes['Codigo cliente'] == codigo_cliente]
    
    # Verificar si se encontró el cliente en la tabla 'diferentes'
    if not fila_cliente.empty:
        demora = int(fila_cliente['Demora-ACT'].values[0])  # Obtener la demora especificada para este cliente
        umbral_actualizacion = demora - 1
        
        # Comprobar si es el momento de actualizar
        if meses_sin_act_por_contrato < umbral_actualizacion:
            # Incrementar el contador y no actualizar
            corresponde_act = False
            meses_sin_act_por_contrato += 1  
            
            return corresponde_act, meses_sin_act_por_contrato
        else:
            corresponde_act = True
            meses_sin_act_por_contrato = 0  
            
            return corresponde_act, meses_sin_act_por_contrato


# In[83]:


def actualizar_tabla_final(tabla_actualizacion, tabla_final, tabla_gatillo, tabla_diferentes,redeterminaciones):    
    nuevas_filas = []

    for index, row in tabla_actualizacion.iterrows():
        cliente = row['Código Cliente']
        ccosto = row['Código CC']
        nombre_ccosto = row['Ccosto']
        servicio = row['Servicio']
        fecha_actual = pd.to_datetime(row['Fecha'], format='%d/%m/%Y')
        coeficiente_actual = row['Coeficiente']
        

        # Calcular el mes anterior
        fecha_mes_anterior = (fecha_actual - relativedelta(months=1)).strftime('%d/%m/%Y')

        # Buscar el registro del mes anterior en la tabla final
        precio_anterior_row = tabla_final[
            (tabla_final['Codigo Cliente'] == cliente) &
            (tabla_final['Codigo Ccosto'] == ccosto) &
            (tabla_final['Col  apoyo'] == servicio) &
            (tabla_final['Fecha'] == fecha_mes_anterior) &
            (tabla_final['Ccosto'] == nombre_ccosto)
        ]

        if not precio_anterior_row.empty:
            
            precio_anterior = precio_anterior_row['Precio'].values[0]
            meses_sin_act = precio_anterior_row['MesesSinActPorContrato'].values[0]
            
            # Convertir sin_act_por_gatillo a booleano explícitamente
            sin_act_por_gatillo = bool(precio_anterior_row['SinActPorGatillo'].values[0])
            coeficiente_anterior = precio_anterior_row['Coeficiente'].values[0]
            cod_articulo = precio_anterior_row['Cod Articulo'].values[0]
            articulo = precio_anterior_row['Articulo'].values[0]

            concatenado = f"{ccosto}-{cliente}-{servicio}"

            if concatenado in redeterminaciones['Concatenacion'].values:
                precio_actualizado = precio_anterior
                hubo_gatillo = False

            else:

                if cliente in tabla_diferentes['Codigo cliente'].values:

                    corresponde_act,meses_sin_act =  actualizar_precio_si_corresponde(cliente, meses_sin_act, tabla_diferentes)

                    if corresponde_act:
                        # Llamar a la lógica de gatillo para calcular el precio actualizado
                        precio_actualizado, hubo_gatillo = aplicar_logica_gatillo(precio_anterior, coeficiente_anterior,
                                                                      coeficiente_actual, sin_act_por_gatillo, cliente, tabla_gatillo)
                    else:
                        precio_actualizado = precio_anterior 
                        hubo_gatillo = False
                else:
                    # Llamar a la lógica de gatillo para calcular el precio actualizado
                    precio_actualizado, hubo_gatillo = aplicar_logica_gatillo(precio_anterior, coeficiente_anterior,
                                                                      coeficiente_actual, sin_act_por_gatillo, cliente, tabla_gatillo)
                
            nueva_fila = {
                'Codigo Cliente': cliente,
                'Cliente': row.get('Cliente', ''),
                'Codigo Ccosto': ccosto,
                'Ccosto': nombre_ccosto,
                'Cod Articulo' : cod_articulo, 
                'Articulo' : articulo,
                'Fecha': fecha_actual.strftime('%d/%m/%Y'),
                'Precio': precio_actualizado,
                'Coeficiente': coeficiente_actual,
                'SinActPorGatillo': hubo_gatillo,
                'MesesSinActPorContrato': meses_sin_act,
                'Col  apoyo': servicio
            }
            nuevas_filas.append(nueva_fila)
            
            
    # Convertir la lista de nuevas filas en un DataFrame y añadirlas a la tabla final
    df_nuevas_filas = pd.DataFrame(nuevas_filas)
    tabla_final = pd.concat([tabla_final, df_nuevas_filas], ignore_index=True)

    return tabla_final


# In[84]:


# Actualizar la tabla final usando la tabla de coeficiente y tabla de gatillo
tabla_final = actualizar_tabla_final(tabla_actualizacion, tabla_final, tabla_gatillo, tabla_diferentes, redeterminaciones)
tabla_final['SinActPorGatillo'] = tabla_final['SinActPorGatillo'].astype(str)
# Guardar la tabla final actualizada en la primera hoja, con el nombre "tabla final"
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    tabla_final.to_excel(writer, sheet_name='tabla final', index=False)

print("Tabla final actualizada y guardada correctamente.")

