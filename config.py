import os
from datetime import datetime
import logging
import json

class CrawlerConfig:
    """
    Configuration class for the Vertex Database Extractor script.
    """
    # Default values if the configuration file is not present
    SEARCHING_FOLDER_PATH = "A:\\Servomex"
    MAKE_HTML = True
    MAKE_XLSX = False
    OUTPUT_FILE_NAME = "VertexDatabase"
    DEBUG_FOLDER_PATH = "C:\\Users\\dwinc\\Desktop\\parts"
    DEBUG = True
    LOGGING_LEVEL = "DEBUG"
    LOG_FILE_PATH = "crawler_2022_01_01_12_30_45.log"
    OUTPUT_LOG_FOLDER_PATH = "C:\\path\\to\\logs"
    OUTPUT_FOLDER_PATH = "C:\\path\\to\\output"
    HTML_EXTENSION = "html"
    XLSX_EXTENSION = "xlsx"
    BASE_HTML_FILE = "base.html"
    HTML_HEADER_NAME = "New And Improved Department!!"
    HTML_FORMATTED_TIME = "Sunday, 01 January 2022 12:30"
    ICON_PATH = "C:\\path\\to\\icon.ico"
    DATA_SHEET_FOLDER_NAME = "Data Sheet"
    SEARCHED_FILE_NAME = "PART_NUMBER.XLS"
    DATATABLE_SELECT_PATH = "dataTables.select.js"
    JQUERY_DATATABLE_CSS_PATH = "jquery_datatable.css"
    JQUERY_DATATABLE_JS_PATH = "jquery_dataTables.js"
    JQUERY_THREE_SIX_ZERO_JS = "jquery-3.6.0.js"
    SELECT_DATATABLES_CSS = "select_dataTables.css"
    
    
    
    def __init__(self):
        # Load configuration from file if present
        config_file_path = 'config.json'
        if os.path.exists(config_file_path):        
            worked = True
            with open(config_file_path, 'r') as file:
                config_data = json.load(file)
            self.__dict__.update(config_data)
        else:            
            worked = False
        self.configure_logging()
        if worked:
            logging.info(f"Loaded config: {config_file_path}")
        else:
            logging.critical(f"[ERROR]Loaded config: {config_file_path}")
            
    def configure_logging(self):
        if not os.path.exists(self.OUTPUT_LOG_FOLDER_PATH):
            os.makedirs(self.OUTPUT_LOG_FOLDER_PATH)
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")            
        file_path = os.path.join(self.OUTPUT_LOG_FOLDER_PATH, self.LOG_FILE_PATH).replace("[timestamp]", timestamp) 
        logging.basicConfig(level=self.LOGGING_LEVEL, filename=file_path, filemode='w', format='%(asctime)s - %(levelname)s - %(message)s')

    def get_output_file_path(self, extension):
        return os.path.join(self.OUTPUT_FOLDER_PATH, f"{self.OUTPUT_FILE_NAME}.{extension}")

    def get_html_output_file_path(self):
        return self.get_output_file_path(self.HTML_EXTENSION)

    def get_xlsx_output_file_path(self):
        return self.get_output_file_path(self.XLSX_EXTENSION)
        
    def create_output_folder(self):
        if not os.path.exists(self.OUTPUT_FOLDER_PATH):
            os.makedirs(self.OUTPUT_FOLDER_PATH)

    def get_formatted_datetime(self):
        return datetime.now().strftime("%A, %d %B %Y %H:%M")

    def update_header_string(self):
        formatted_datetime = datetime.now().strftime(self.HTML_FORMATTED_TIME)
        return f"{self.HTML_HEADER_NAME}<br>Run on {formatted_datetime}"

