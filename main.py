#Importar matplotlib.pyplot, openpyxl y numpy

import numpy
import openpyxl
import matplotlib.pyplot

#Abrir archivo1

m = openpyxl.load_workbook('Archivo1.xlsx',data_only = True)

#Activar la hoja del archivo1 en la que se encuentran los datos

hoja1 = m.active

#Acceder a las celdas que continen la información a utilizar

celdas = hoja1['A1':'B4']

#Recorrer y obtener el contenido de las celdas para añadirlo a una lista

listam = []

for fila in celdas:
  a = [celda.value for celda in fila] 
  listam.append(a)

#print(listam)

#Cerrar archivo 1

m.close()

#Abrir archivo2

n = openpyxl.load_workbook('Archivo2.xlsx',data_only = True)

#Activar la hoja del archivo1 en la que se encuentran los datos

hoja2 = n.active

#Acceder a las celdas que continen la información a utilizar

celdas2 = hoja2['A1':'B3']

#Recorrer y obtener el contenido de las celdas para añadirlo a una lista

listan = []

for fila in celdas2:
  b = [celda.value for celda in fila] 
  listan.append(b)

#print(listan)

#Cerrar archivo 2

n.close()

#Crear el archivo 3 en donde se van a guardar las respuestas del archivo 1

e = open('Archivo3.xlsx','w+')

##Trabajo en el archivo 3

#Escribir el contenido del archivo 1 en el archivo 3

e.write("El contenido del archivo 1 es:\n")
e.write(str(listam))

#Convertir la lista de datos del archivo 1 a un diccionario

d1 = dict(listam)

#Escribir el diccionario 1 en el archivo 3

e.write("\n\nEl diccionario con la lista anterior es :\n")
e.write(str(d1))

#Realizar gráfico del diccionario 1 de color verde y con tringulos

a = matplotlib.pyplot.plot(d1,'g^')

#Nombrar los ejes del grafico

b = matplotlib.pyplot.xlabel ('Tiempo')
c =matplotlib.pyplot.ylabel('Valor')

#Colocarle nombre al grafico

d = matplotlib.pyplot.title('Grafico del diccionario 1')

# mostrar gráfico
g = matplotlib.pyplot.show()

#Escribir la grafica obtenida en el archivo 3

e.write("\n\nEl gráfico del diccionario 1 es :\n")
e.write(g)

#Cerrar y guardar la infromación del archivo 3

e.close()

#Crear el archivo 4 en donde se van a guardar las respuestas del archivo 2

f = open('Archivo4.xlsx','w+')

##Trabajo en el archivo 4

#Escribir el contenido del archivo 2 en el archivo 4

f.write("El contenido del archivo 2 es:\n")
f.write(str(listan))

#Convertir la lista de datos del archivo 1 a un diccionario

d2 = dict(listan)

#Escribir el diccionario 2 en el archivo 4

f.write("\n\nEl diccionario con la lista anterior es :\n")
f.write(str(d2))

#Realizar gráfico ddel diccionario 2 realizado con una linea punteada y de color azul 

h = matplotlib.pyplot.plot(d2,'b--')

#Nombrar los ejes del grafico

g = matplotlib.pyplot.xlabel ('Categoría')
l =matplotlib.pyplot.ylabel('Alimento')

#Colocarle nombre al grafico

s = matplotlib.pyplot.title('Grafico del diccionario 2')

# mostrar gráfico
z = matplotlib.pyplot.show()

#Escribir la grafica obtenida en el archivo 3

f.write("\n\nEl gráfico del diccionario 1 es :\n")
f.write(z)

#Cerrar y guardar la infromación del archivo 4

f.close()
