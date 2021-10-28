import openpyxl
import sqlite3
import math
import sys
import time

class ManagerSqlite(object):

    def __init__(self, formato, estado, municipio="", circunscripcion=""):
        print("Inicializar la instancia de la base de datos")
        super(ManagerSqlite, self).__init__()
        self.formato = formato
        self.estado = estado
        self.municipio = municipio
        self.circunscripcion = circunscripcion
        self.fileNames = []
        dbName = self.GetdbName(self.formato)
        print("Base de datos seleccionada: " + dbName)
        self.conn = sqlite3.connect(dbName)
        self.cursor = self.conn.cursor()

    def __del__(self):
        print("Liberar la instancia de la base de datos")
        self.cursor.close()
        self.conn.close()

    def executeSqlCommand(self, sqlCommand):
        """
                 Ejecute el comando sql ingresado
                 : param sqlCommand: comando sql
        """
        print("Ejecutar sql personalizado")
        print(sqlCommand)
        self.cursor.execute(sqlCommand)
        results = self.cursor.fetchall()
        self.conn.commit()
        return results

    def GetdbName(self, formato):
        if formato == "CVC Gobernador" or formato == "CVC Alcalde":
           dbName = "ver_ciu.db"
        else:
           dbName = "act_cont.db"
        return dbName

    def GettableName(self, formato):
        if formato == "CVC Gobernador" or formato == "AEC Gobernador" or formato == "CREC Gobernador":
           tableName = "Gobernador"
        elif formato == "CVC Alcalde" or formato == "AEC Alcalde" or formato == "CREC Alcalde":
           tableName = "Alcalde"
        elif formato == "AEC Legislador Lista" or formato == "CREC Legislador Lista":
           tableName = "Legislador_lista"
        elif formato == "AEC Legislador Nominal" or formato == "CREC Legislador Nominal":
           tableName = "Legislador_Nominal"
        elif formato == "AEC Concejal Lista" or formato == "CREC Concejal Lista":
           tableName = "Concejal_lista"
        elif formato == "AEC Concejal Nominal" or formato == "CREC Concejal Nominal":
           tableName = "Concejal_Nominal"
        else:
           tableName = "NA"
           print("Formato no encontrado")

        return tableName

    def GettableColumn(self, formato):
        if formato == "AEC Legislador Lista" or formato == "CREC Legislador Lista" \
           or formato == "AEC Concejal Lista" or formato == "CREC Concejal Lista":
           columnNames = "siglas, cod_par"
           formatoList = True
        else:
           columnNames = "nombre_boleta, siglas, cod_par"
           formatoList = False

        return columnNames, formatoList

    def GetNameFormato(self, formato):
        if formato == "CREC Concejal Nominal":
           nameFormato = "CREC_Concejal_Nominal"
           numeroLine = 73
           numeroLineaData = 25
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H12"
           PosPagina = "J134"

        elif formato == "AEC Concejal Nominal":
           nameFormato = "AEC_Concejal_Nominal"
           numeroLine = 73
           numeroLineaData = 23
           PosEstado = "H4"
           PosMunicipio = "A5"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "CREC Concejal Lista":
           nameFormato = "CREC_Concejal_Lista"
           numeroLine = 74
           numeroLineaData = 26
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "AEC Concejal Lista":
           nameFormato = "AEC_Concejal_Lista"
           numeroLine = 73
           numeroLineaData = 26
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "CVC Gobernador":
           nameFormato = "CVC_Gobernador"
           numeroLine = 61
           numeroLineaData = 24
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "AEC Gobernador":
           nameFormato = "AEC_Gobernador"
           numeroLine = 72
           numeroLineaData = 24
           PosEstado = "H4"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "CREC Gobernador":
           nameFormato = "CREC_Gobernador"
           numeroLine = 74
           numeroLineaData = 25
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "CVC Alcalde":
           nameFormato = "CVC_Alcalde"
           numeroLine = 62
           numeroLineaData = 23
           PosEstado = "H4"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "AEC Alcalde":
           nameFormato = "AEC_Alcalde"
           numeroLine = 75
           numeroLineaData = 23
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "CREC Alcalde":
           nameFormato = "CREC_Alcalde"
           numeroLine = 73
           numeroLineaData = 25
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H12"
           PosPagina = "J134"

        elif formato == "CREC Legislador Nominal":
           nameFormato = "CREC_Legislador_Nominal"
           numeroLine = 72
           numeroLineaData = 25
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H12"
           PosPagina = "J134"

        elif formato == "AEC Legislador Nominal":
           nameFormato = "AEC_Legislador_Nominal"
           numeroLine = 72
           numeroLineaData = 24
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "CREC Legislador Lista":
           nameFormato = "CREC_Legislador_Lista"
           numeroLine = 75
           numeroLineaData = 25
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        elif formato == "AEC Legislador Lista": 
           nameFormato = "AEC_Legislador_Lista"         
           numeroLine = 74
           numeroLineaData = 25
           PosEstado = "H5"
           PosMunicipio = "A6"
           PosCirscuncripcion = "H11"
           PosPagina = "J134"

        return nameFormato, numeroLine, numeroLineaData, PosEstado, PosMunicipio, PosCirscuncripcion, PosPagina

    def ExecuteOutput(self):
         tableName = self.GettableName(self.formato)
         columnNames, formatoList = self.GettableColumn(self.formato)

         if len(self.estado) == 0 and len(self.municipio) == 0 and len(self.circunscripcion) ==0:
            sqlQuery = "select " + columnNames + " from " + tableName + " where des_edo is not null"

         if len(self.estado) > 0 and len(self.municipio) == 0 and len(self.circunscripcion) == 0:
            sqlQuery = "select " + columnNames + " from " + tableName  + " where des_edo = '" + self.estado + "'" + \
                      " and des_edo is not null"

         if len(self.estado) > 0 and len(self.municipio) == 0 and len(self.circunscripcion) > 0:
            sqlQuery = "select " + columnNames + " from " + tableName  + " where des_edo = '" + self.estado + "'" + \
                      " and cod_circunscripcion = " + self.circunscripcion + " and des_edo is not null"

         if len(self.estado) > 0 and len(self.municipio) > 0 and len(self.circunscripcion) == 0:
            sqlQuery = "select " + columnNames + " from " + tableName  + " where des_edo = '" + self.estado + "' and des_mun = '" + \
            self.municipio  + "' and des_edo is not null"

         if len(self.estado) > 0 and len(self.municipio) > 0 and len(self.circunscripcion) > 0:
             sqlQuery = "select " + columnNames + " from " + tableName  + " where des_edo = '" + self.estado + "' and des_mun = '" + \
                        self.municipio  + "' and cod_circunscripcion = " + self.circunscripcion  + "  and des_edo is not null"

         if self.formato == "AEC Legislador Nominal" or self.formato == "CREC Legislador Nominal":
            sqlQuerypn = "select distinct orden_nominal from " + tableName  + " where des_edo = '" + self.estado + \
                        "' and cod_circunscripcion = " + self.circunscripcion  + "  and des_edo is not null"
            outconpn = self.executeSqlCommand(sqlQuerypn)
            range_max=outconpn.__len__()
         elif self.formato == "AEC Concejal Nominal" or self.formato == "CREC Concejal Nominal":
            sqlQuerypn = "select distinct orden_nominal from " + tableName  + " where des_edo = '" + self.estado + \
                        "' and des_mun = '" + self.municipio  + "' and cod_circunscripcion = " + \
                        self.circunscripcion  + "  and des_edo is not null"
            outconpn = self.executeSqlCommand(sqlQuerypn)
            range_max=outconpn.__len__()
         else:
            range_max = 1

         self.fileNames = []

         print('range_max', range_max)
         for indexmax in range(0,range_max):

            if self.formato == "AEC Concejal Nominal" or self.formato == "CREC Concejal Nominal" \
               or self.formato == "AEC Legislador Nominal" or self.formato == "CREC Legislador Nominal":
                sqlQuerySearch = sqlQuery + " and orden_nominal = " + str(indexmax + 1)
            else:
                sqlQuerySearch = sqlQuery

            output = self.executeSqlCommand(sqlQuerySearch)
            nameFormato, numeroLine, numeroLineaData, PosEstado, PosMunicipio, PosCirscuncripcion, PosPagina = self.GetNameFormato(self.formato)
            wb = openpyxl.load_workbook('Formatos/' + nameFormato + '.xlsx')
            sheet = wb['Hoja1']

            print("Numero total de lineas: " + str(output.__len__()))
            numeroPag = math.ceil(output.__len__()/numeroLine)
            print("Numero total de paginas: " + str(numeroPag))

            timestr = time.strftime("%Y%m%d-%H%M%S")

            if numeroPag == 1:
                for index in range(0, output.__len__()):
                    print(str(output[index]))
                    if formatoList:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        sheet['B' + str(index + numeroLineaData)] = output[index][1]
                    else:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        if str(output[index][1]) == "None":
                           siglas = ""
                        else:
                           siglas = str(output[index][1])
                        sheet['B' + str(index + numeroLineaData)] = siglas
                        sheet['C' + str(index + numeroLineaData)] = output[index][2]
                sheet[PosEstado] = self.estado
                sheet[PosMunicipio] = '                    ' + self.municipio
                if len(self.circunscripcion) > 0:
                   sheet[PosCirscuncripcion] = 'CIRCUNSCRIPCION: ' + self.circunscripcion
                else:
                   sheet[PosCirscuncripcion] = ''

                sheet[PosPagina] = ""
                if range_max > 1:
                   if self.formato == "AEC Concejal Nominal" or self.formato == "CREC Concejal Nominal" \
                      or self.formato == "AEC Legislador Nominal" or self.formato == "CREC Legislador Nominal":
                      sheet[PosPagina] = str(indexmax+1) + '/' + str(range_max)

                wb.save('Salidas/' + nameFormato + '-' + timestr + '-' + str(indexmax) + '.xlsx')
                self.fileNames.append(nameFormato + '-' + timestr + '-' + str(indexmax))

            elif numeroPag == 2:
                for index in range(0, numeroLine):
                    print(str(index), str(output[index]))
                    if formatoList:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        sheet['B' + str(index + numeroLineaData)] = output[index][1]
                    else:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        if str(output[index][1]) == "None":
                           siglas = ""
                        else:   
                           siglas = str(output[index][1]) 
                        sheet['B' + str(index + numeroLineaData)] = siglas
                        sheet['C' + str(index + numeroLineaData)] = output[index][2]
                sheet[PosEstado] = self.estado
                sheet[PosMunicipio] = '                    ' +  self.municipio
                if len(self.circunscripcion) > 0:
                   sheet[PosCirscuncripcion] = 'CIRCUNSCRIPCION: ' + self.circunscripcion
                else:
                   sheet[PosCirscuncripcion] = ''
                sheet[PosPagina] = "1/2"
                wb.save('Salidas/' + nameFormato + '-' + timestr + '-' + str(indexmax) + 'p1.xlsx')
                self.fileNames.append(nameFormato + '-' + timestr + '-' + str(indexmax) + 'p1')

                for index in range(numeroLine, output.__len__()):
                    print(str(index), str(output[index]))
                    if formatoList:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = str(output[index][0])
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = output[index][1]
                    else:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = str(output[index][0])
                        if str(output[index][1]) == "None":
                           siglas = ""
                        else:
                           siglas = str(output[index][1])
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = siglas
                        sheet['C' + str(index - numeroLine + numeroLineaData)] = output[index][2]
                for index in range(output.__len__(), numeroLine*2):
                    print(str(index))
                    if formatoList:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = ""
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = ""
                    else:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = ""
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = ""
                        sheet['C' + str(index - numeroLine + numeroLineaData)] = ""
                sheet[PosEstado] = self.estado
                sheet[PosMunicipio] = '                    ' +  self.municipio
                if len(self.circunscripcion) > 0:
                   sheet[PosCirscuncripcion] = 'CIRCUNSCRIPCION: ' + self.circunscripcion
                else:
                   sheet[PosCirscuncripcion] = ''
                sheet[PosPagina] = "2/2"
                wb.save('Salidas/' + nameFormato + '-' + timestr + '-' + str(indexmax) + 'p2.xlsx')
                self.fileNames.append(nameFormato + '-' + timestr + '-' + str(indexmax) + 'p2')

            elif numeroPag == 3:
                for index in range(0, numeroLine):
                    print(str(index), str(output[index]))
                    if formatoList:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        sheet['B' + str(index + numeroLineaData)] = output[index][1]
                    else:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        if str(output[index][1]) == "None":
                           siglas = ""
                        else:
                           siglas = str(output[index][1])
                        sheet['B' + str(index + numeroLineaData)] = siglas
                        sheet['C' + str(index + numeroLineaData)] = output[index][2]
                sheet[PosEstado] = self.estado
                sheet[PosMunicipio] = '                    ' +  self.municipio
                if len(self.circunscripcion) > 0:
                   sheet[PosCirscuncripcion] = 'CIRCUNSCRIPCION: ' + self.circunscripcion
                else:
                   sheet[PosCirscuncripcion] = ''
                sheet[PosPagina] = "1/3"
                wb.save('Salidas/' + nameFormato + '-' + timestr + '-' + str(indexmax) + 'p1.xlsx')
                self.fileNames.append(nameFormato + '-' + timestr + '-' + str(indexmax) + 'p1')

                for index in range(numeroLine, numeroLine*2):
                    print(str(index), str(output[index]))
                    if formatoList:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        sheet['B' + str(index + numeroLineaData)] = output[index][1]
                    else:
                        sheet['A' + str(index + numeroLineaData)] = str(output[index][0])
                        if str(output[index][1]) == "None":
                           siglas = ""
                        else:
                           siglas = str(output[index][1])
                        sheet['B' + str(index + numeroLineaData)] = siglas
                        sheet['C' + str(index + numeroLineaData)] = output[index][2]
                sheet[PosEstado] = self.estado
                sheet[PosMunicipio] =  '                    ' + self.municipio
                if len(self.circunscripcion) > 0:
                   sheet[PosCirscuncripcion] = 'CIRCUNSCRIPCION: ' + self.circunscripcion
                else:
                   sheet[PosCirscuncripcion] = ''
                sheet[PosPagina] = "2/3"
                wb.save('Salidas/' + nameFormato + '-' + timestr + '-' + str(indexmax) + 'p2.xlsx')
                self.fileNames.append(nameFormato + '-' + timestr + '-' + str(indexmax) + 'p2')

                for index in range(numeroLine*2, output.__len__()):
                    print(str(index), str(output[index]))
                    if formatoList:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = str(output[index][0])
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = output[index][1]
                    else:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = str(output[index][0])
                        if str(output[index][1]) == "None":
                           siglas = ""
                        else:
                           siglas = str(output[index][1])
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = siglas
                        sheet['C' + str(index - numeroLine + numeroLineaData)] = output[index][2]

                for index in range(output.__len__(), numeroLine*3):
                    print(str(index))
                    if formatoList:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = ""
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = ""
                    else:
                        sheet['A' + str(index - numeroLine + numeroLineaData)] = ""
                        sheet['B' + str(index - numeroLine + numeroLineaData)] = ""
                        sheet['C' + str(index - numeroLine + numeroLineaData)] = ""
                sheet[PosEstado] = self.estado
                sheet[PosMunicipio] =  '                    ' + self.municipio
                if len(self.circunscripcion) > 0:
                   sheet[PosCirscuncripcion] = 'CIRCUNSCRIPCION: ' + self.circunscripcion
                else:
                   sheet[PosCirscuncripcion] = ''
                sheet[PosPagina] = "3/3"
                wb.save('Salidas/' + nameFormato + '-' + timestr + '-' + str(indexmax) + 'p3.xlsx')
                self.fileNames.append(nameFormato + '-' + timestr + '-' + str(indexmax) + 'p3')
