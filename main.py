# -*- coding: utf-8 -*-
# Source: https://www.datacamp.com/community/tutorials/python-excel-tutorial (no vale para excel viejos)
# https://stackoverflow.com/questions/13437727/python-write-to-excel-spreadsheet (excel viejos)
# https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/
import os

from xlrd import open_workbook
from xlutils.copy import copy
import shutil

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()
cwd
version = "0.3 alpha"
# Change directory
# os.chdir("/path/to/your/folder")

# List all files and directories in current directory
files = os.listdir('.')

c = 0
d = 0
archivos = []


def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell


def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
            # END HACK


def list(array):
    c = 0
    for j in array:
        print "[" + str(c) + "] " + j
        c = c + 1
    return 0;


def seleccion(archivo):
    print archivo
    abierto = open_workbook("./" + archivo, on_demand=True, formatting_info=True)
    wb = copy(abierto)
    xl_sheet = abierto.sheet_by_index(0)
    hojan = wb.get_sheet(0)
    hojan.name = u'HojasUsua'
    #print ('Sheet name: %s' % xl_sheet.name)
    zz = raw_input("La clave en SGIPE es zz? s/n\n")
    vlc = raw_input("El diseño es de Valencia? s/n\n")
    inte = int(raw_input("No. de interiores\n"))
    neq = int(raw_input("No. de equipos\n"))
    ncto = int(raw_input("No de CTO\n"))

    for j in range(0, xl_sheet.nrows):
        #tarea = xl_sheet.cell(j, 1)
        #valor = xl_sheet.cell(j, 2)
        #print ('Tarea: [%s] Valor: %s' % (tarea, valor))
        ron = xl_sheet.row(j)
        if zz == "s" or zz == "S":
            if ron[1].value == '160253':
                setOutCell(hojan, 2, j, 1)
            if ron[1].value == "160245":
                setOutCell(hojan, 2, j, inte)
        if ron[1].value == "166006":
            setOutCell(hojan, 2, j, 2)
        if ron[1].value == "165069":
            setOutCell(hojan, 2, j, neq)
        if ron[1].value == "165051":
            setOutCell(hojan, 2, j, neq*2)
        if ron[1].value == "165077":
            setOutCell(hojan, 2, j, neq)
        if ron[1].value == "160024":
            setOutCell(hojan, 2, j, 1)
        if vlc == "s" or vlc == "S":
            if ron[1].value == "740314":
                setOutCell(hojan, 2, j, 1)
            if ron[1].value == "720488":
                setOutCell(hojan, 2, j, 1)
        if ron[1].value == "160270":
            setOutCell(hojan, 2, j, ncto)
    wb.save('Alta_masiva_hojas_usua.xls')
    abierto.release_resources()
    os.remove(archivo)
    shutil.copy2('Alta_masiva_hojas_usua.xls', archivo)
    return 0;

print ("SGIPE converter version " + version + " by borekon")
for i in files:
    a = i.find('xls')
    if a > 0 and i.__len__() == 12:
        archivos.append(i)
    elif a > 0:
        archivos.append(i)
if archivos.__len__() == 1:
    respuesta = str(raw_input("Es %s el archivo a editar? (s/n)\n" % archivos[0]))
    if respuesta == "n" or respuesta == "N":
        list(files)
        respuesta = int(raw_input("Elige el archivo\n"))
        seleccion(files[respuesta])
    elif respuesta == "s" or respuesta == "S":
        seleccion(archivos[0])
    else:
        print "Respuesta inválida"
elif archivos.__len__() > 1:
    list(archivos)
    print "[" + str(archivos.__len__()) + "] Ver todos"
    respuesta = int(raw_input("Elige el archivo a editar\n"))
    if respuesta == archivos.__len__():
        list(files)
        respuesta = int(raw_input("Elige el archivo\n"))
        seleccion(files[respuesta])
    else:
        seleccion(archivos[respuesta])
else:
    print("No hay archivos excel en el directorio.")



