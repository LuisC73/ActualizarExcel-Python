from openpyxl import Workbook
import win32com.client
import time
import glob

def main():
    actualizar_archivo()

def actualizar_archivo():   
    # Hacemos uso del glob que nos permite buscar una lista de archivos en el sistema de archivos con nombres que coinciden con un patrón
    archivos = glob.glob(r"D:/Users/practicante.geserv1/OneDrive - Centro de Servicios Mundial SAS/Imágenes/lm/Python/actualizarExcel-Python/input/"+"*.xlsm")
    File = win32com.client.Dispatch("Excel.Application")
    desicion = input("Quieres Actualizar todos los archivos: \n si \n no \n -")
    
    if desicion == "si":
        for f in archivos:
            File.visible = 1
            print("Abriendo Archivo.....")
            Workbook = File.Workbooks.open(f)
            print("Actualizando Archivo.....")
            Workbook.RefreshAll()
            time.sleep(8)
            Workbook.Save()
            File.Quit()
    elif desicion == " ":
        print("Porfavor ingrese un valor valido")   
    else:
        quit()         

if __name__ == "__main__":
    main()
    input("\tPROCESO TERMINADO, presione enter para salir...")    
