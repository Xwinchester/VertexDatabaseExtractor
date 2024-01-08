import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from config import MAKE_HTML, MAKE_XLSX, OUTPUT_FILE_NAME, SEARCHING_FOLDER_PATH, COMPANY_NAME, DEPARTMENT_NAME, DEBUG_FOLDER_PATH
from config import DEBUG as CONFIG_DEBUG


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
        
    #FOLDER_PATH = r"A:\Servomex"# Use the 'r' (to use the raw sting) before the folder path to not have to double the '\'
    DATA_SHEET_FOLDER_NAME = "Data Sheet"
    SEARCHED_FILE_NAME = "PART_NUMBER.XLS"
    
    def __init__(self):
        self.data = []
                
        # just shows the user info on total and current file
        self.FILE_COUNTER = {'total':0, 'current':0}
                
        # Crawl the folder
        if CONFIG_DEBUG:
            self.crawl_folder_debugging(DEBUG_FOLDER_PATH)
        else:
            self.crawl_folder(SEARCHING_FOLDER_PATH)
        print("")
        
        # Output the data into the database
        if self.FILE_COUNTER['current'] == 0:
                print("No data found, no output.")
        else:
            if MAKE_HTML:
                print("Outputting to HTML file.")
                HTMLOutput(pd.DataFrame(self.data))
            if MAKE_XLSX:
                print("Outputting to XLSX file.")
                ExcelOutput(pd.DataFrame(self.data))

    def crawl_folder_debugging(self, folder_path):
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isdir(file_path):
                self.crawl_folder_debugging(file_path)
            elif any(file_path.endswith(ext) for ext in ['xlsx']):
                self.FILE_COUNTER['current'] += 1
                print(f"File {self.FILE_COUNTER['current']}.", end='\r')
                self.extract_data_debugging(file_path)
            

    def crawl_folder(self, folder_path):
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isdir(file_path):
                self.crawl_folder(file_path)
            elif self.DATA_SHEET_FOLDER_NAME.lower().replace(" ", "") in file_path.lower().replace(" ", ""):
                if file_name.upper() in self.SEARCHED_FILE_NAME:
                    self.FILE_COUNTER['current'] += 1
                    print(f"File {self.FILE_COUNTER['current']}.", end='\r')
                    self.extract_data(file_path)

    def extract_data_debugging(self, file_path):
        # Skip the first row (header) when skiprows = 1
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, header=None, skiprows=0, engine='xlrd')
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, header=None, skiprows=0, engine='openpyxl')
        self.data.extend([{ 'Part Number':row[1], 'Serial Number':row[0], 'Vertex Height':row[3]} for _, row in df.iterrows() if pd.notna(row[0]) and pd.notna(row[2]) ])                        
                
    def extract_data(self, file_path):
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, header=None, skiprows=0, engine='xlrd')
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, header=None, skiprows=0, engine='openpyxl')
            else:
                print(f"Unsupported file format: {file_path}")
                return
            # Extracting Part Number from cell E3
            part_number = df.iloc[2, 5]
            # Skip the top 11 rows after extracting the part number
            df = df[11:]
            # Continue with data extraction using df
            d = [{'Part Number': part_number, 'Serial Number': row[0], 'Vertex Height': row[2]} for _, row in df.iterrows() if pd.notna(row[0]) and pd.notna(row[2])]
            self.data.extend(d)
        except Exception as e:
            print(e)

    
class HTMLOutput:

    BASE_HTML_FILE = "base.html"
    EXTENSION = "html"
    
    def __init__(self, df):  
        self.df = df
        # Read the base HTML template
        with open(self.BASE_HTML_FILE, 'r') as file:
            self.html_data = file.read()

        # Update web-page title
        self.update_webpage_title()
        
        # Update page header
        self.update_header()
        
        # Inject the table data into the html file
        self.inject_table_data()
        
        # Update icon for the page
        self.update_icon()
                          
        # Save the modified HTML to a file
        with open(self.get_file_path(), 'w') as html_file:
            html_file.write(self.html_data)
    
    def inject_table_data(self):
        # Update Headers
        header_columns = ''.join(f'<th>{col}</th>' for col in self.df.columns)
        self.update_html(before='<!-- Dropdowns will be dynamically added here -->', after=header_columns)
        
        # Update data inside of the tables
        data_rows = ''.join('<tr>' + ''.join(f'<td>{row[col]}</td>' for col in self.df.columns) + '</tr>' for _, row in self.df.iterrows())                      
        self.update_html(before='<!-- Data rows will be dynamically added here -->', after=data_rows)

    def update_webpage_title(self):
        self.update_html(before='<!-- CUSTOME TITLE -->', after=f"{COMPANY_NAME} - {DEPARTMENT_NAME}")

    def update_header(self):
        # format the current time to let the operators know when the macro was run last
        formatted_datetime = datetime.now().strftime("%A, %d %B %Y %H:%M")
        # Message in the top of the html file showing the operators
        header_string = f"{COMPANY_NAME} {DEPARTMENT_NAME}<br>Run on {formatted_datetime}"       
        self.update_html(before='<!-- HEADER will be dynamically added here -->', after=header_string)
            
    def update_icon(self):
        # Udpate icon
        icon_path = os.path.join(os.getcwd(), "icon.ico")
        self.update_html(before='<ICON_PATH>', after=icon_path) 
            
    def update_html(self, before, after):
        self.html_data = self.html_data.replace(before, after)
            
    def get_file_path(self):
        # Get the current working directory
        current_directory = os.getcwd()
        # Move one level up to the parent directory
        parent_directory = os.path.dirname(current_directory)
        return os.path.join(parent_directory, f"{OUTPUT_FILE_NAME}.{self.EXTENSION}")

class ExcelOutput:
    
    EXTENSION = "xlsx"
        
    def __init__(self, df):  
         self.output()
         self.add_autofilters()
         
    def get_file_path(self):
         # Get the current working directory
         current_directory = os.getcwd()
         # Move one level up to the parent directory
         parent_directory = os.path.dirname(current_directory)       
         return os.path.join(parent_directory, f"{OUTPUT_FILE_NAME}.{self.EXTENSION}")
        
    def output(self):
        self.df.to_excel(self.get_file_path(), engine='xlsxwriter', index=False)
        
    def add_autofilters(self):
        workbook = load_workbook(self.get_file_path())
        sheet = workbook.active
        sheet.auto_filter.ref = sheet.dimensions
        workbook.save(self.get_file_path())  


if __name__ == "__main__":
    print_header()
    VertextDatabaseExtractor()
