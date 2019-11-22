import xlrd
import sys
import os
import getopt
from pathlib import Path

def encontrarCelda(sh, searchedValue):
    ret = ""
    for row in range(sh.nrows):
        for col in range(sh.ncols):
            celda = sh.cell(row, col)
            if sh.cell_type(row, col) == 1 and searchedValue in celda.value:
                ret += xlrd.formula.cellname(row, col) + ","
    return ret

def procesaArchivos(dir, buscar):
    numBuscados = 0
    archBuscar = sorted(os.listdir( dir ))
    Archivos = filter(isExcel,archBuscar)
    for Archivo in Archivos:
        if "~lock" in Archivo:
            continue
        numBuscados += 1
        try:
            for sh in xlrd.open_workbook(os.path.join(dir,Archivo)).sheets():
                ret = encontrarCelda(sh, buscar)
                if len(ret) > 0:
                    print("\tarchivo:  '" + Archivo + "'->" + sh.name + "\n\t\t" + splitRet(ret))
        except:
            print("\tNo se pudo abrir el archivo:  '" + Archivo + "'")

    return numBuscados


def splitRet(ret):
    linea = ""
    retorno = ""
    for elemen in ret.split(","):
        if len(linea) + len(elemen) + 17 < columnas:
            if len(linea) > 0:
                linea += ","
            linea += elemen
        else:
            retorno += linea + "\n\t\t"
            linea = ""

    retorno += linea
    return retorno

def get_windows_terminal():
    from ctypes import windll, create_string_buffer
    h = windll.kernel32.GetStdHandle(-12)
    csbi = create_string_buffer(22)
    res = windll.kernel32.GetConsoleScreenBufferInfo(h, csbi)
 
    if not res: return 80, 25 
 
    import struct
    (bufx, bufy, curx, cury, wattr, left, top, right, bottom, maxx, maxy)\
    = struct.unpack("hhhhHhhhhhh", csbi.raw)
    width = right - left + 1
    height = bottom - top + 1
 
    return width, height


def isExcel(variable): 
    if ('.xls' in variable): 
        return True
    else: 
        return False

def printInstrucciones():
     print ('\tusar: BuscarContenidoExcel.pyz -o <buscar en> -b <texto a buscar>')
     print ('\t\t-h\tAyuda')
     print ('\t\t-o\tDirectorio donde se inicia la busqueda')
     print ('\t\t-b\Text a buscar')
     print ('  ej:   BuscarContenidoExcel.pyz -o "C:/Datos" -b "planilla"')

def main(argv):
    ''' Crear el archivo '.pyz' con python -m zipapp BuscarContenidoExcel '''
    global columnas
    columnas, filas = get_windows_terminal()
    origen = ''
    buscar = ''
    try:
        opts, args = getopt.getopt(argv,"ho:b:")
    except getopt.GetoptError:
        printInstrucciones()
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            printInstrucciones()
            sys.exit()
        elif opt in ("-o"):
            origen = arg
        elif opt in ("-b"):
            buscar = arg

    print("")
    print(' '*5 +"origen : " + origen)
    print(' '*5 +"buscar : " + buscar)

    if len(origen) == 0 or len(buscar) == 0:
        printInstrucciones()
        sys.exit(2)

    if not os.path.exists(os.path.join(origen)):
        print ('-o: error en origen: ' + origen)
        sys.exit(2)

    numBuscados = 0
    print("\n")
    for (root,dirs,files) in os.walk(os.path.join(origen)):
        numBuscados += procesaArchivos(os.path.join(root), buscar)
        for dir in dirs:
            numBuscados += procesaArchivos(os.path.join(root,dir), buscar)
             
    print("\nArchivos buscados: " + str(numBuscados))

if __name__ == "__main__":
   main(sys.argv[1:])