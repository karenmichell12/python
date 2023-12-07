from os import listdir, makedirs
from os.path import isfile, join, exists
from shutil import move
import pandas as pd
import os

Excel = r'C:\Users\AUXILIAR ING\Documents\karen\Libro11.xlsx'
Ruta_creacion = input("Ingrese donde quiere crear los archivos: ")

df = pd.read_excel(Excel)
Lista = pd.Series.tolist(df["ACTIVIDADES"])

try:
    for i in range(0,len(Lista)-1):
        makedirs(Ruta_creacion+'\\'+Lista[i].replace('-','\\'))
    
except OSError:
    if not os.path.isdir(Ruta_creacion+'\\'+Lista[i].replace('-','\\')):
        raise


def archivos( nombre ):
  if ( len(nombre)<3 ):
    return False
  archiv = nombre[0:3]
  return archiv

def moverFichero( nombre, dirBase ):
  directorio = nombre[0:-7]
  destino = join( Ruta_creacion, directorio.replace('-','\\'))
  if ( not exists(destino) ):
    makedirs(destino)
  origen = join( dirBase, nombre )
  move ( origen, destino )

dirBase= input("Ingrese el directorio donde se encuentran los archivos: ")
ficheros = [ f for f in listdir(dirBase) if isfile(join(dirBase,f)) ]
for fich in ficheros:
  if ( archivos( fich ) ):
    moverFichero( fich, dirBase )