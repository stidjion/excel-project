import openpyxl as px
from openpyxl import load_workbook
import pandas as pd

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
            df = pd.read_excel(self.file_path, sheet_name=self.ws.title)
            return df
        except:
            df = pd.DataFrame()
            return df
    def save_dataframe(self):
        with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
          self.df.to_excel(writer, sheet_name=self.ws.title, index=False)

        return True
    def set_active_sheet(self, sheet_name):
        if sheet_name in self.wb.sheetnames:
            self.ws = self.wb[sheet_name]
            self.df = self._load_dataframe()
            return True
        else:
                print(f"Sheet '{sheet_name}' does not exist.")
                return False
    def create_sheet(self, sheet_name):
        try:
            self.wb.create_sheet(title=sheet_name)
            self.wb.save(self.file_path)
            self.ws = self.wb[sheet_name]
            self.df = pd.DataFrame()
            return True
        except:
            print(f"Failed to create sheet '{sheet_name}'.")
            return False

    


