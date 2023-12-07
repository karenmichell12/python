import os
from colorama import init, Fore, Back
import pandas as pd


 
def check_name_ex(ex):
    if not "." in ex:
        ex = "."+ex
    return ex
 
def BMP(s):
    return "".join((i if ord(i) < 10000 else '\ufffd' for i in s))
 
def change_dir():
    while True:
        dire = input("Introduzca directorio base: ").strip()
        if os.path.isdir(dire):
            os.chdir(dire)
            break
        else:
            print(Fore.RED+"ERROR, DIRECTORIO NO VÁLIDO"+Fore.RESET)
 
def show_dir(direc):
    global showed_dir
    if showed_dir == False:
        print(Fore.BLUE+Back.WHITE+direc+Fore.RESET+Back.RESET)
        showed_dir = True
 
 
def ns(c):
    while c.lower():
        print(chr(7))
    return(c.lower())



A = []
R = []
conti = "s"
while(conti == "s"):
    init()
    print("Directorio actual: {} ".format(os.getcwd()))
    count = 0
    showed_dir = False

    change_dir()
    sep = ns
    print("BUSCANDO...\n")
    for dirpath, dirnames, filenames in os.walk(os.getcwd()):
            for file in filenames:
                name,ex = os.path.splitext(file)
                if sep:
                        show_dir(dirpath)
                count+=1
                conti = (Fore.GREEN+'{}-'.format(count)+os.path.join(dirpath,BMP(file)))
                conti2 = (os.path.join(dirpath,BMP(file)))
                cont = dirpath
                A.append(conti2)
                R.append(conti)
                print(conti)
                
            showed_dir = False
    else:
            print(Fore.BLACK+Back.GREEN+"\n{} ARCHIVOS ENCONTRADOS.".format(count))
    print(Fore.RESET+Back.RESET+"")

folder = os.getcwd()


ruta = input("Ingrese la ruta donde quiere crear el excel:")
df = pd.DataFrame(A,columns=['Rutas'])
ruta1 = ruta + '\\' + 'rutas1.xlsx'
df.to_excel( ruta1 ,sheet_name='hoja1')


input("Presione una tecla para continuar...")
df = pd.read_excel(ruta1 , 
                   sheet_name='hoja1',
                   index_col = 2) 

File = pd.ExcelFile(ruta1)
df = File.parse('hoja1')



Lista = pd.Series.tolist(df["Rutas"])
Lista1= pd.Series.tolist(df["Unnamed: 2"])


f = []

for i in range(0,len(Lista)):
    R = (''.join(Lista[i]).split('\\')[-1])
    archivo, ext = os.path.splitext(R)
    f.append(archivo + ext)

for i in range(0,len(f)):
    os.rename(f[i],Lista1[i])
    

print(f)


