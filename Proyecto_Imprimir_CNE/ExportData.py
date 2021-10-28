import sqlite3
import pandas as pd
import sys


class exportXLSdb():
   def exportDB(self, ruta_xls, ruta_db):
      con=sqlite3.connect(ruta_db)
      xls = pd.ExcelFile(ruta_xls)

      for sheet in xls.sheet_names:
          df=xls.parse(sheet, parse_dates=True)
          df.to_sql(sheet, con, index=False,if_exists="replace")

      con.commit()
      con.close()
