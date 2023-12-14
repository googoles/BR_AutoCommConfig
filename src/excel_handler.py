import pandas as pd

class ExcelHandler:
    @staticmethod
    def read_excel(file_path):
        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            raise ValueError(f"Error reading Excel file: {str(e)}")
