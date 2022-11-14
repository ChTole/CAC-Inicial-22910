"""
El sitio https://jsonplaceholder.typicode.com/users tiene alojados datos de usuarios (ficticios) en formato JSON. 
Utilizando la librería requests, capturar los datos y generar un libro de MS Excel con la siguiente información:
- id / name / email / phone / website
Tener en cuenta los tipos de colecciones a recorrer (y cómo acceder a los datos).
"""

import requests
from openpyxl import Workbook

solicitud = requests.get('https://jsonplaceholder.typicode.com/users')


if solicitud.status_code == 200:
    datos = solicitud.json()
    print('La conexión fue exitosa!')
else:
    print('No respondió el servidor!')
    
##########################################

libro = Workbook()
hoja = libro.active
encabezado = ['id', 'name', 'email',  'phone', 'website']
hoja.append(encabezado)

for dato in datos:
   hoja.append([dato['id'],dato['name'],dato['email'],dato['phone'],dato['website']]) 

libro.save(filename="D:\CodoACodo\Comision 22910\Clase29\ejercicios\Datos.xlsx")
libro.close()