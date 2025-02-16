### Importar librerias
import pandas as pd
import numpy as np
import os
import sys
import time
import powerfactory as pf
from datetime import datetime

def createFolder(dirName):
    try:
        # Create directory only if it doesn't exist already
        if not os.path.exists(dirName):
            os.makedirs(dirName)
    except OSError:
        print('Error creating directory: ', dirName)


### Conexión Power Factory
app = pf.GetApplication()
if not app:
    raise Exception("No se pudo conectar con PowerFactory")

dirRoot = os.path.dirname(os.path.abspath(__file__))
Folder_resultados = os.path.join(dirRoot, 'Resultados_Variaciones')
createFolder(Folder_resultados)

# Información basica
app.PrintPlain('Directorio: %s' %(dirRoot))

user=app.GetCurrentUser()
app.PrintPlain('Usuario: %s' %(user))

oFolders= user.GetContents('*.IntFolder',0)
app.PrintPlain('Carpetas: %s' %(oFolders))

oFoldersStudy = app.GetProjectFolder('study')
app.PrintPlain('Carpeta Study: %s' %(oFoldersStudy))
Contents=oFoldersStudy.GetContents()

oFoldersScheme = app.GetProjectFolder('scheme')

# variaciones activas
variaciones_activas = app.GetActiveNetworkVariations()
sstage_activos =app.GetActiveStages()
def explorar_objetos(folder, ruta_actual=''):
    resultados = []
    
    # Obtener todos los objetos dentro de la carpeta actual
    contenidos = folder.GetContents()
    
    # Si no hay objetos, salir de la función retornando lista vacía
    if not contenidos:
        return []

    for item in contenidos:
        # Guardar información de cada objeto encontrado
        if item.GetClassName() == 'IntFolder':
            
            resultados.append({
                'Objeto': item,
                'Nombre': item.loc_name,
                'Tipo': item.GetClassName(),
                'Ruta Jerárquica': ruta_actual,
                'Ruta': item.GetFullName(),
                
            })
            nueva_ruta = os.path.join(ruta_actual, item.loc_name)
            resultados += explorar_objetos(item, nueva_ruta)  # Concatenar resultados de la subcarpeta
        
        elif item.GetClassName() == 'IntScheme':
            estado_variacion = 'Activo' if item in variaciones_activas else 'Inactivo'
            resultados.append({
                'Objeto': item,
                'Nombre': item.loc_name,
                'Tipo': item.GetClassName(),
                'Ruta Jerárquica': ruta_actual,
                'Ruta': item.GetFullName(),
                'Estado': estado_variacion,
                'Activation Time Starting': datetime.utcfromtimestamp(item.tFromAc).strftime('%Y-%m-%d %H:%M:%S'),
                'Activation Time Completed': datetime.utcfromtimestamp(item.tToAc).strftime('%Y-%m-%d %H:%M:%S') ,
            })
            nueva_ruta = os.path.join(ruta_actual, item.loc_name)
            resultados += explorar_objetos(item, nueva_ruta)  # Concatenar resultados de la subcarpeta
        elif item.GetClassName() == 'IntSstage':
            estado_sstage = 'Activo' if item in sstage_activos else 'Inactivo'
            resultados.append({
                'Objeto': item,
                'Nombre': item.loc_name,
                'Tipo': item.GetClassName(),
                'Ruta Jerárquica': ruta_actual,
                'Ruta': item.GetFullName(),
                'Estado': estado_sstage,
                'Variación': item.GetVariation(),
                'Activation Time': datetime.utcfromtimestamp(item.tAcTime).strftime('%Y-%m-%d %H:%M:%S')
            })       

    return resultados  # Retornar la lista completa de objetos encontrados

# Ejecutar la función recursiva
resultados = explorar_objetos(oFoldersScheme)
for i in resultados:
    app.PrintPlain(i)
# Convertir resultados a DataFrame y guardar en Excel
df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel(os.path.join(Folder_resultados,'Listado_Variaciones_2025_v1.xlsx'), index=False)

app.PrintPlain('Exploración finalizada. Resultados guardados en Objetos_PowerFactory.xlsx')
