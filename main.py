import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from config import CrawlerConfig
import logging



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
    
    def __init__(self):
        # Configure logging
        crawlerConfig = CrawlerConfig()
        
        # Log a message indicating the start of the script
        logging.info("Starting Vertex Database Extractor")
    
        self.data = []
                
        # just shows the user info on total and current file
        self.FILE_COUNTER = 0
                
        # Crawl the folder
        if crawlerConfig.DEBUG:
            self.crawl_folder_debugging(crawlerConfig.DEBUG_FOLDER_PATH)
            logging.warning("Running in debug mode.")
        else:
            self.crawl_folder(crawlerConfig.SEARCHING_FOLDER_PATH)
        print("")
        
        # Output the data into the database
        if self.FILE_COUNTER == 0:
                logging.warning("No data found, no output.")
        else:
            logging.info(f"Extracted data from {self.FILE_COUNTER} files.")
            if crawlerConfig.MAKE_HTML:
                logging.info("Outputting to HTML file.")
                HTMLOutput(pd.DataFrame(self.data), crawlerConfig)
            if crawlerConfig.MAKE_XLSX:
                logging.info("Outputting to XLSX file.")
                ExcelOutput(pd.DataFrame(self.data), crawlerConfig)

        # Log a message indicating the end of the script
        logging.info("Vertex Database Extractor completed")                

    def crawl_folder_debugging(self, folder_path):
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isdir(file_path):
                self.crawl_folder_debugging(file_path)
            elif any(file_path.endswith(ext) for ext in ['xlsx']):
                self.FILE_COUNTER += 1
                print(f"File {self.FILE_COUNTER}.", end='\r')
                self.extract_data_debugging(file_path)
            
    def crawl_folder(self, folder_path):
        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isdir(file_path):
                self.crawl_folder(file_path)
            elif crawlerConfig.DATA_SHEET_FOLDER_NAME.lower().replace(" ", "") in file_path.lower().replace(" ", ""):
                if file_name.upper() in crawlerConfig.SEARCHED_FILE_NAME:
                    self.FILE_COUNTER += 1
                    print(f"File {self.FILE_COUNTER}.", end='\r')
                    self.extract_data(file_path)

    def extract_data_debugging(self, file_path):
        # Skip the first row (header) when skiprows = 1
        if file_path.endswith('.xls'):
            df = pd.read_excel(file_path, header=None, skiprows=0, engine='xlrd')
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, header=None, skiprows=0, engine='openpyxl')
        self.data.extend([{'Part Number': row[1], 'Serial Number': row[0], 'Vertex Height': row[3]} for _, row in df.iterrows() if pd.notna(row[0]) and pd.notna(row[2])])                        
                
    def extract_data(self, file_path):
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, header=None, skiprows=0, engine='xlrd')
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, header=None, skiprows=0, engine='openpyxl')
            else:
                logging.CRITICAL(f"Unsupported file format: {file_path}")
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
    
    def __init__(self, df, crawlerConfig):  
        self.df = df
        self.crawlerConfig = crawlerConfig
        # Read the base HTML template
        with open(crawlerConfig.BASE_HTML_FILE, 'r') as file:
            self.html_data = file.read()

        # Update web-page title
        self.update_webpage_title()
        
        # Update the js and css paths
        self.update_js_css_paths()
        
        # Update page header
        self.update_header()
        
        # Inject the table data into the html file
        self.inject_table_data()
        
        # Update icon for the page
        self.update_icon()
                          
        # Save the modified HTML to a file
        logging.info(f"Outputting html to file: {self.get_file_path()}")
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
        self.update_html(before='<!-- CUSTOME TITLE -->', after=f"{self.crawlerConfig.HTML_HEADER_NAME}")

    def update_header(self):
        # Message in the top of the html file showing the operators
        header_string = self.crawlerConfig.update_header_string()      
        self.update_html(before='<!-- HEADER will be dynamically added here -->', after=header_string)
            
    def update_icon(self):
        # Update icon
        self.update_html(before='<ICON_PATH>', after=self.crawlerConfig.ICON_PATH) 

    def update_js_css_paths(self):
        # Update icon
        self.update_html(before='<DATATABLE_SELECT_PATH>', after=self.crawlerConfig.DATATABLE_SELECT_PATH)         
        self.update_html(before='<JQUERY_DATATABLE_CSS_PATH>', after=self.crawlerConfig.JQUERY_DATATABLE_CSS_PATH)        
        self.update_html(before='<JQUERY_DATATABLE_JS_PATH>', after=self.crawlerConfig.JQUERY_DATATABLE_JS_PATH)        
        self.update_html(before='<JQUERY_THREE_SIX_ZERO_JS>', after=self.crawlerConfig.JQUERY_THREE_SIX_ZERO_JS)        
        self.update_html(before='<SELECT_DATATABLES_CSS>', after=self.crawlerConfig.SELECT_DATATABLES_CSS)        
            
    def update_html(self, before, after):
        self.html_data = self.html_data.replace(before, after)
            
    def get_file_path(self):
           return self.crawlerConfig.get_html_output_file_path()

class ExcelOutput:
            
    def __init__(self, df, crawlerConfig):  
         self.crawlerConfig = self.crawlerConfig
         self.output()
         self.add_autofilters()
         
    def get_file_path(self):
        return self.crawlerConfig.get_xlsx_output_file_path()
        
    def output(self):
        logging.info(f"Outputting Excel to file: {self.get_file_path()}")
        self.df.to_excel(self.get_file_path(), engine='xlsxwriter', index=False)
        
    def add_autofilters(self):
        workbook = load_workbook(self.get_file_path())
        sheet = workbook.active
        sheet.auto_filter.ref = sheet.dimensions
        workbook.save(self.get_file_path())  

if __name__ == "__main__":
    print_header()
    VertextDatabaseExtractor()
