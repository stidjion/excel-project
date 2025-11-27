import openpyxl as px
from openpyxl import load_workbook
import pandas
class ExcelConnector:
    def __init__(self, file_path):
      self.file_path = file_path
      
    def load_workbook(self):
        try:
            workbook = load_workbook(filename=self.file_path)
            return workbook
        except Exception as e:
            print(f"Error loading workbook: {e}")
            return None
    def active_sheet(self, workbook, sheet_name):
        try:
            wb =  workbook()
            ws = wb.active if sheet_name is None else wb[sheet_name]
            ws.title = sheet_name
            return ws
        except Exception as e:
            print(f"Error accessing sheet: {e}")
            return None
    