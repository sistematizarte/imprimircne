from PyQt5.QtWidgets import QApplication,QLineEdit,QWidget,QFormLayout,QComboBox,QLabel,QPushButton
from PyQt5.QtGui import QIntValidator,QDoubleValidator,QFont
from PyQt5.QtCore import Qt
from FilterFormatData import ManagerSqlite
from ExportData import exportXLSdb
from printerExcel import PrinterXLS
import os
import sys
import sqlite3
import openpyxl
import time

class ManagerFilter():

    def __init__(self, dbName):
        print("Inicializar la instancia de la base de datos")
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

class formularioMain(QWidget):
        def __init__(self,parent=None):
                super().__init__(parent)
                self.desde = QLineEdit()
                self.desde.setValidator(QIntValidator())
                self.desde.setMaxLength(4)
                self.desde.setAlignment(Qt.AlignRight)
                self.desde.setFont(QFont("Arial",20))

                self.desde.textChanged.connect(self.textchangedDesde)

                self.hasta = QLineEdit()
                self.hasta.setValidator(QIntValidator())
                self.hasta.setMaxLength(4)
                self.hasta.setAlignment(Qt.AlignRight)
                self.hasta.setFont(QFont("Arial",20))

                self.hasta.textChanged.connect(self.textchangedHasta)

                self.formato = QComboBox(self)
                self.formato.addItem("CVC Gobernador")
                self.formato.addItem("CVC Alcalde")
                self.formato.addItem("AEC Gobernador")
                self.formato.addItem("AEC Alcalde")
                self.formato.addItem("AEC Legislador Lista")
                self.formato.addItem("AEC Legislador Nominal")
                self.formato.addItem("AEC Concejal Lista")
                self.formato.addItem("AEC Concejal Nominal")
                self.formato.addItem("CREC Gobernador")
                self.formato.addItem("CREC Alcalde")
                self.formato.addItem("CREC Legislador Lista")
                self.formato.addItem("CREC Legislador Nominal")
                self.formato.addItem("CREC Concejal Lista")
                self.formato.addItem("CREC Concejal Nominal")

                self.formato.activated[str].connect(self.onChangedFor)

                self.linea = QComboBox(self)
                self.linea.addItem("LINEA1")
                self.linea.addItem("LINEA2")
                self.linea.addItem("LINEA3")
                self.linea.addItem("LINEA4")
                self.linea.addItem("LINEA5")
                self.linea.addItem("LINEA6")
                self.linea.addItem("LINEA7")
                self.linea.addItem("LINEA8")
                self.linea.addItem("LINEA9")
                self.linea.addItem("LINEA10")

                self.linea.activated[str].connect(self.onChangedLin)

                self.estado = QComboBox(self)
                dbFilter = ManagerFilter("edo_mun.db")
                sqlQuery = "select distinct estado from estados"
                output = dbFilter.executeSqlCommand(sqlQuery)
                self.estado.addItem("Seleccione estado")
                for index in range(0, output.__len__()):
                     print(str(output[index][0]))
                     self.estado.addItem(str(output[index][0]))

                self.estado.activated[str].connect(self.onChangedEdo)

                self.municipio = QComboBox(self)
                self.municipio.setFixedSize(210,25)
                self.municipio.activated[str].connect(self.onChangedMun)

                self.circunscripcion = QComboBox(self)
                self.circunscripcion.setFixedSize(210,25)
                self.circunscripcion.activated[str].connect(self.onChangedCir)

                self.totalNum = QLabel(self)
                self.totalNum.setFont(QFont("Arial",20))
                self.desde.textChanged.connect(self.textchangedTotal)

                self.btnPrint = QPushButton('IMPRIMIR', self)
                self.btnPrint.setGeometry(20, 300, 150, 35)
                self.btnPrint.clicked.connect(self.onClickPrint)
                self.btnPrint.setEnabled(False)

                self.btnData = QPushButton('RECARGAR', self)
                self.btnData.setGeometry(20, 300, 150, 35)
                self.btnData.clicked.connect(self.onClickData)
                
                self.btnStop = QPushButton('DETENER', self)
                self.btnStop.setGeometry(20, 300, 150, 35)
                self.btnStop.clicked.connect(self.onClickStop)                

                flo = QFormLayout()
                flo.addRow("IMPRESORA: ",self.linea)
                flo.addRow("FORMATO: ",self.formato)
                flo.addRow("ESTADO: ",self.estado)
                flo.addRow("MUNICIPIO: ",self.municipio)
                flo.addRow("CIRCUNSCRIPCION: ",self.circunscripcion)
                flo.addRow("Desde",self.desde)
                flo.addRow("Hasta",self.hasta)
                flo.addRow("Total",self.totalNum)
                flo.addRow("",self.btnPrint)
                flo.addRow("",self.btnData)
                flo.addRow("",self.btnStop)

                self.linea1 = True                  
                self.linea2 = True 
                self.linea3 = True 
                self.linea4 = True                  
                self.linea5 = True 
                self.linea6 = True 
                self.linea7 = True                  
                self.linea8 = True 
                self.linea9 = True 
                self.linea10 = True 
                
                self.setLayout(flo)
                self.setWindowTitle("Impresion de planilla")

        def onClickData(self):
               print("Recargar la data")
               exportxls = exportXLSdb()
               exportxls.exportDB("cont.xlsx","act_cont.db")
               exportxls.exportDB("cvc.xlsx","ver_ciu.db")

        def onClickPrint(self):
                print("Ha hecho click en Imprimir")
                if self.formato.currentText() == "CVC Gobernador" or self.formato.currentText() == "CREC Gobernador" \
                   or self.formato.currentText() == "CREC Legislador Lista" or self.formato.currentText() == "AEC Gobernador" \
                   or self.formato.currentText() == "AEC Legislador Lista":
                      print("Generando archivo de salida")
                      manSql = ManagerSqlite(self.formato.currentText(),self.estado.currentText())
                      self.printInstancia(manSql)

                elif self.formato.currentText() == "CVC Alcalde" or self.formato.currentText() == "CREC Alcalde" \
                   or self.formato.currentText() == "CREC Concejal Lista" or self.formato.currentText() == "AEC Alcalde" \
                   or self.formato.currentText() == "AEC Concejal Lista":
                      print("Generando archivo de salida")
                      manSql = ManagerSqlite(self.formato.currentText(),self.estado.currentText(),self.municipio.currentText())
                      self.printInstancia(manSql) 

                elif self.formato.currentText() == "CREC Legislador Nominal" or self.formato.currentText() == "AEC Legislador Nominal":
                      print("Generando archivo de salida")
                      manSql = ManagerSqlite(self.formato.currentText(),self.estado.currentText(),"",self.circunscripcion.currentText())
                      self.printInstancia(manSql) 

                elif self.formato.currentText() == "CREC Concejal Nominal" or self.formato.currentText() == "AEC Concejal Nominal":
                      print("Generando archivo de salida")
                      manSql = ManagerSqlite(self.formato.currentText(),self.estado.currentText(),self.municipio.currentText(),self.circunscripcion.currentText())
                      self.printInstancia(manSql)

        def printInstancia(self, instanciaSQL):
            instanciaSQL.ExecuteOutput()
            printFile = PrinterXLS()
            printerName = printFile.selectPrinter(self.linea.currentText())
            timestr = time.strftime("%Y%m%d-%H%M%S")
            self.activeLinea(True)
            for index in range(int(self.desde.text()), int(self.hasta.text())+1):
                print('instanciaSQL.fileNames', instanciaSQL.fileNames)
                for fileName in instanciaSQL.fileNames:
                    currentFiles = []
                    wb = openpyxl.load_workbook('Salidas/' + fileName + ".xlsx")
                    sheet = wb['Hoja1']
                    sheet['L134'] = index
                    fileNamestr = fileName + "-" + str(index) + ".xlsx" 
                    wb.save('Salidas/' + fileNamestr)
                    currentFiles.append(fileNamestr)
                    print('currentFiles', currentFiles)
                    printFile.PrintFile(currentFiles, printerName)
                    if self.linea.currentText() == "LINEA1":                    
                       if self.linea1 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA2":                    
                       if self.linea2 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA3":                    
                       if self.linea3 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA4":                    
                       if self.linea4 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA5":                    
                       if self.linea5 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA6":                    
                       if self.linea6 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA7":                    
                       if self.linea7 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA8":                    
                       if self.linea8 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA9":                    
                       if self.linea9 == False:
                          print("Abortando impresion ......!!!")  
                          break
                    if self.linea.currentText() == "LINEA10":                    
                       if self.linea10 == False:
                          print("Abortando impresion ......!!!")  
                          break                      
                      
            #for fileName in manSql.fileNames:
            #    os.system("del " + fileName)
            #for fileName in currentFiles:
            #    os.system("del " + fileName)

        def activeLinea(self,bandera):
            if self.linea.currentText() == "LINEA1":
               self.linea1 = bandera  
            if self.linea.currentText() == "LINEA2":
               self.linea2 = bandera
            if self.linea.currentText() == "LINEA3":
               self.linea3 = bandera  
            if self.linea.currentText() == "LINEA4":
               self.linea4 = bandera  
            if self.linea.currentText() == "LINEA5":
               self.linea5 = bandera  
            if self.linea.currentText() == "LINEA6":
               self.linea6 = bandera  
            if self.linea.currentText() == "LINEA7":
               self.linea7 = bandera  
            if self.linea.currentText() == "LINEA8":
               self.linea8 = bandera  
            if self.linea.currentText() == "LINEA9":
               self.linea9 = bandera  
            if self.linea.currentText() == "LINEA10":
               self.linea10 = bandera                 

        def onClickStop(self):
            print("Ha realizado un click en detener")                
            self.activeLinea(False)

        def validatePrint(self):
                self.btnPrint.setEnabled(False)
                print("Validacion de impresion")
                if self.formato.currentText() == "CVC Gobernador" or self.formato.currentText() == "CREC Gobernador" \
                   or self.formato.currentText() == "CREC Legislador Lista" or self.formato.currentText() == "AEC Gobernador" \
                   or self.formato.currentText() == "AEC Legislador Lista":
                      if self.estado.currentIndex() > 0 and len(self.desde.text()) and len(self.hasta.text()) > 0 \
                         and self.totalNum.text() != "Valor invalido":
                            print("Habilitar boton para imprimir")
                            self.btnPrint.setEnabled(True)
                elif self.formato.currentText() == "CVC Alcalde" or self.formato.currentText() == "CREC Alcalde" \
                   or self.formato.currentText() == "CREC Concejal Lista" or self.formato.currentText() == "AEC Alcalde" \
                   or self.formato.currentText() == "AEC Concejal Lista":
                      if self.estado.currentIndex() > 0 and len(self.desde.text()) and len(self.hasta.text()) > 0 \
                         and self.totalNum.text() != "Valor invalido" and self.municipio.currentIndex() > 0:
                            print("Habilitar boton para imprimir")
                            self.btnPrint.setEnabled(True)
                elif self.formato.currentText() == "CREC Legislador Nominal" or self.formato.currentText() == "AEC Legislador Nominal":
                      if self.estado.currentIndex() > 0 and len(self.desde.text()) and len(self.hasta.text()) > 0 \
                         and self.totalNum.text() != "Valor invalido" and self.circunscripcion.currentIndex() > 0:
                            print("Habilitar boton para imprimir")
                            self.btnPrint.setEnabled(True)
                elif self.formato.currentText() == "CREC Concejal Nominal" or self.formato.currentText() == "AEC Concejal Nominal":
                      if self.estado.currentIndex() > 0 and len(self.desde.text()) and len(self.hasta.text()) > 0 \
                         and self.totalNum.text() != "Valor invalido" and self.circunscripcion.currentIndex() > 0 \
                         and self.municipio.currentIndex() > 0:
                            print("Habilitar boton para imprimir")
                            self.btnPrint.setEnabled(True)

        def onChangedCir(self, text):
                self.validatePrint()
                print("Circunscripcion seleccionada: " + text)

        def onChangedLin(self, text):
                self.validatePrint()
                print("Impresora seleccionada: " + text)

        def onChangedFor(self, text):
                self.validatePrint()
                self.estado.setCurrentIndex(0)
                self.municipio.clear()
                self.circunscripcion.clear()

        def onChangedEdo(self, text):
                self.validatePrint()
                if text != "Seleccione estado":
                   print("Estado seleccionado: " + text)
                   if self.formato.currentText() == "CVC Alcalde" or self.formato.currentText() == "CREC Alcalde" \
                      or self.formato.currentText() == "CREC Concejal Nominal" or self.formato.currentText() == "CREC Concejal Lista" \
                      or self.formato.currentText() == "AEC Alcalde" or self.formato.currentText() == "AEC Concejal Nominal" \
                      or self.formato.currentText() == "AEC Concejal Lista":
                      dbFilter = ManagerFilter("edo_mun.db")
                      sqlQuery = "select distinct municipio from municipios where cod_estado = " + \
                                 "(select cod_estado from estados where estado = '" + text + "')"
                      output = dbFilter.executeSqlCommand(sqlQuery)
                      self.municipio.clear()
                      self.municipio.addItem("Seleccione municipio")
                      for index in range(0, output.__len__()):
                         print(str(output[index][0]))
                         self.municipio.addItem(str(output[index][0]))
                   elif  self.formato.currentText() == "CREC Legislador Nominal" \
                         or self.formato.currentText() == "AEC Legislador Nominal":
                      dbFilter = ManagerFilter("edo_mun.db")
                      sqlQuery = "select cod_estado from estados where estado = '" + text  + "'"
                      output = dbFilter.executeSqlCommand(sqlQuery)
                      cod_estado = str(output[0][0])

                      dbFilter = ManagerFilter("act_cont.db")
                      sqlQuery = "select distinct cod_circunscripcion from Legislador_Nominal " + \
                      " where cod_estado = " + cod_estado
                      output = dbFilter.executeSqlCommand(sqlQuery)

                      self.circunscripcion.clear()
                      self.circunscripcion.addItem("Seleccione circunscripcion")
                      for index in range(0, output.__len__()):
                         print(str(output[index][0]))
                         self.circunscripcion.addItem(str(int(output[index][0])))
                else:
                   self.municipio.clear()
                   self.circunscripcion.clear()

        def onChangedMun(self, text):
              print("Municio seleccionado: " + text)
              self.validatePrint()
              if text != "Seleccione municipio":
                 if self.formato.currentText() == "CREC Concejal Nominal" or  self.formato.currentText() == "AEC Concejal Nominal":
                   dbFilter = ManagerFilter("edo_mun.db")
                   sqlQuery = "select cod_estado from estados where estado = '" + self.estado.currentText()  + "'"
                   output = dbFilter.executeSqlCommand(sqlQuery)
                   cod_estado = str(output[0][0])

                   sqlQuery = "select cod_municipio from municipios where municipio = '" + text  + "'"
                   output = dbFilter.executeSqlCommand(sqlQuery)
                   cod_municipio = str(output[0][0])

                   dbFilter = ManagerFilter("act_cont.db")
                   sqlQuery = "select distinct cod_circunscripcion from Concejal_Nominal " + \
                              " where cod_estado = " + cod_estado + \
                              " and cod_municipio = " + cod_municipio

                   print(sqlQuery)
                   output = dbFilter.executeSqlCommand(sqlQuery)
                   self.circunscripcion.clear()
                   self.circunscripcion.addItem("Seleccione circunscripcion")
                   for index in range(0, output.__len__()):
                       print(str(output[index][0]))
                       self.circunscripcion.addItem(str(int(output[index][0])))

        def textchangedTotal(self,text):
                print("Etiqueta de total cambiada:" + text)
                self.validatePrint()

        def textchangedDesde(self,text):
                print("Changed: " + text)
                self.validatePrint()
                if len(self.desde.text()) > 0 and len(self.hasta.text()) > 0:
                   print("Validado etiquetas con informacion")
                   if int(self.hasta.text()) >= int(self.desde.text()):
                      print("Asignando valor a etiqueta")
                      self.totalNum.setText(str(int(self.hasta.text()) - int(self.desde.text()) + 1))
                   else:
                     self.totalNum.setText("Valor invalido")

        def textchangedHasta(self,text):
                print("Changed: " + text)
                self.validatePrint()
                if len(self.desde.text()) > 0 and len(self.hasta.text()) > 0:
                   print("Validado etiquetas con informacion")
                   if int(self.hasta.text()) >= int(self.desde.text()):
                      print("Asignando valor a etiqueta")
                      self.totalNum.setText(str(int(self.hasta.text()) - int(self.desde.text()) + 1))
                   else:
                     self.totalNum.setText("Valor invalido")


if __name__ == "__main__":
        app = QApplication(sys.argv)
        win = formularioMain()
        win.show()
        sys.exit(app.exec_())
