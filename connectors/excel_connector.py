import openpyxl as px
from openpyxl import load_workbook
import pandas

class ExcelConnector:
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = self._load_file()
        self.ws = self.wb.active
        self.df = self._load_dataframe()
  

    def _load_file(self):
     
     try:
            wb = load_workbook(self.file_path)
            return wb
     except:
         wb = px.Workbook()
         wb.save(self.file_path)
         return wb
    def _load_dataframe(self):
        try:
            df = pandas.read_excel(self.file_path)
            return df
        except:
            df = pandas.DataFrame()
            return df
    def save_dataframe(self, df):
        df.to_excel(self.file_path, index=False)
        return True
    


