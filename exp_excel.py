#!/usr/bin/python
#coding: utf8 

import xlrd
import sys
import glob
import string
import mysql.connector
import unicodedata
from mysql.connector import errorcode




def elimina_tildes(s): 
  return ''.join((c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')) 


#telefono1,nss,nombre,fecha_Salida,afore,monto,telefono,razonsocial,domicilio,ubicacion
def title_execl(filename):
  myexcel = xlrd.open_workbook(filename)

  list_tables = myexcel.sheet_names()

  list_tables = [lista.replace(' ','_') for lista in list_tables]
  list_tables = [lista.replace(',','_') for lista in list_tables]
  list_tables = [lista.replace('.','_') for lista in list_tables]

#  print sheet.nrows
  i = 0
  ddl_tabla = ""
  list_col = []

  for tabla in list_tables:
    sheet = myexcel.sheet_by_index(i)
    i = i + 1
    ddl_tabla = "create table t_" + tabla + "("
    list_col = []
    for columna in range(0,sheet.ncols-1):
      colname = sheet.cell(0,columna).value
      if colname:
        if columna<sheet.ncols-1:
          colname = elimina_tildes(unicode(colname))
          colname = colname.replace('.','')
          colname = colname.replace(' ','_') 
          colname = colname + str(columna)
          ddl_tabla = ddl_tabla + colname.lower() +" varchar(255)," 
          list_col.append(colname.lower)
    colname = sheet.cell(0,sheet.ncols-1).value
    if colname:
      colname = elimina_tildes(unicode(colname))
      colname = colname.replace('.','')
      colname = colname.replace(' ','_') 
      colname = colname + str(columna)
      ddl_tabla = ddl_tabla + colname.lower() +" varchar(255)"
      list_col.append(colname.lower)
    ddl_tabla = ddl_tabla + ");"
    ddl_tabla = ddl_tabla.replace(',);',');')
    print ""
    print "Creando Tabla %s " % tabla
    print ddl_tabla 
    print ""
    try:
      cnx = mysql.connector.connect(user='dbadmin', password='hola123',host='127.0.0.1',database='mails')
      cursor = cnx.cursor()
    except mysql.connector.Error as err:
      if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
        print("Something is wrong with your user name or password")
      elif err.errno == errorcode.ER_BAD_DB_ERROR:
        print("Database does not exists")
      else:
        print(err)
    else:
      print "Conectado a la base de datos"
    cursor.execute(ddl_tabla)
    print "insertando a la tabla "
    add_dml = "insert into " + tabla 
    campos_insert = "("
    for x in range(0,len(list_col)-1):
      campos_insert = campos_insert + str(list_col[x]) + ","
    campos_insert = campos_insert + str(list_col[-1]) + ") "

    for row_index in range(1,sheet.nrows-1):
      values_dml = "values ("
      for col_num in range(0,sheet.ncols-1):
        print col_num
        print row_index
        print values_dml
        values_dml = values_dml + "'" + sheet.cell(row_index,col_num).value + "',"
      values_dml = values_dml + "'" + sheet.cell(row_index,sheet.ncols).value + "')"
      add_dml = add_dml + values_dml 
      cursor.execute(add_dml)
      cnx.commit()
    cursor.close()
    cnx.close()   

title_execl('BD_Afore.xls')