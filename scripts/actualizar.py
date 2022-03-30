from openpyxl import Workbook
import win32com.client

def main():
    actualizar_archivo()


def actualizar_archivo():    

    File = win32com.client.Dispatch("Excel.Application")

    File.visible = 1
    print("Abriendo Archivo.....")
    Workbook = File.Workbooks.open(r"C:\Users\Alejo\Documents\Luis\solo_python\actualizar\output\excel.xlsx")

    Workbook.RefreshAll()

    Workbook.Save()

    File.Quit()


if __name__ == "__main__":
    main()
    input("\tPROCESO TERMINADO, presione enter para salir...")    

