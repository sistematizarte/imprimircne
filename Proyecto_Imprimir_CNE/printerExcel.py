import os
import win32api
import win32print
import time

class PrinterXLS():

  def selectPrinter(self, impresora):
      index = 0
      Printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 2)
      print("[Print List]")

      for Printer in Printers:
         index += 1
         print(str(index) + ". " + Printer['pPrinterName'])

         if Printer['pPrinterName'] == impresora:
            win32print.SetDefaultPrinter(Printers[index-1]['pPrinterName'])
            print("Setting Printer: " + win32print.GetDefaultPrinter())

      return  win32print.GetDefaultPrinter()

  def PrintFile(self,fileNames, printerName):
      for fileName in fileNames:
          full_filename = os.path.join(os.getcwd() + '/Salidas/', fileName)
          print("Imprimiendo: " + full_filename)
          win32api.ShellExecute(0, 'printto', full_filename, '"' + printerName + '"', None,  0)
          time.sleep(4)
