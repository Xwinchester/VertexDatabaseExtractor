import os
import pandas as pd
from openpyxl import load_workbook


def print_header():
    chars = 50
    print("-" * chars)
    print(f"{'Running Vertex Database Extractor':^{chars}}")
    print("-" * chars)
    print(f"{'Programmed by Drew Winchester':^{chars}}")
    print("-" * chars)

class VertextDatabaseExtractor:
    """
    This class represents a Vertex Database Extractor, designed to crawl a specified folder,
    extract data from Excel files with a '.xlsx' extension, and output the aggregated data
    into a new Excel file named 'vertex_database.xlsx'. 
    """
        
    FOLDER_PATH = r"C:\Users\dwinc\Desktop\parts"# Use the 'r' (to use the raw sting) before the folder path to not have to double the '\'
    OUTPUT_FILE_NAME = 'vertex_database.xlsx'
    
    def __init__(self):
        self.extensions = [".xlsx"] 
        self.data = []
        
        # just shows the user info on total and current file
        self.FILE_COUNTER = {'total':0, 'current':0}
        
        self.get_file_count(self.FOLDER_PATH)
        
        # Crawl the folder
        self.crawl_folder(self.FOLDER_PATH)
        print("")
        
        # Output the data into the database
        self.output()
        
         # Add autofilters to headers using openpyxl
        self.add_autofilters()
        
    def get_file_count(self, folder_path):
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isdir(file_path):
                self.get_file_count(file_path)
            elif any(file_path.endswith(ext) for ext in self.extensions):
                self.FILE_COUNTER['total'] += 1      

    def crawl_folder(self, folder_path):
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isdir(file_path):
                self.crawl_folder(file_path)
            elif any(file_path.endswith(ext) for ext in self.extensions):
                self.FILE_COUNTER['current'] += 1
                print(f"File {self.FILE_COUNTER['current']} of {self.FILE_COUNTER['total']}.", end='\r')
                self.extract_data(file_path)
                
    def extract_data(self, file_path):
        # Skip the first row (header) when skiprows = 1
        df = pd.read_excel(file_path, header=None, skiprows=1) 
        self.data.extend([{'Serial Number':row[0], 'Part Number':row[1], 'Operator':row[2], 'Vertex Height':row[3]} for _, row in df.iterrows()])         

    def output(self):
        result_df = pd.DataFrame(self.data)
        result_df.to_excel(self.OUTPUT_FILE_NAME, engine='xlsxwriter', index=False)                

    def add_autofilters(self):
        workbook = load_workbook(self.OUTPUT_FILE_NAME)
        sheet = workbook.active
        sheet.auto_filter.ref = sheet.dimensions
        workbook.save(self.OUTPUT_FILE_NAME)

if __name__ == "__main__":
    print_header()
    VertextDatabaseExtractor()
