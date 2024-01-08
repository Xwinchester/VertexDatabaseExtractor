# Folder path we are searching in
# Use the 'r' (to use the raw sting) before the folder path to not have to double the '\'
# If Debug is True, grabs data from DEBUG_FOLDER_PATH
SEARCHING_FOLDER_PATH = r"A:\Servomex"

# True = makes html database
# False = Will not make the html file
MAKE_HTML = True

# True = makes xlsx database
# False = Will not make the xlsx file
MAKE_XLSX = False

# Extension will come from the output class
OUTPUT_FILE_NAME = "output_file_name"


# Company Name in header of HTML File
COMPANY_NAME = "[company_name]"

# Department Name in HTML File
DEPARTMENT_NAME = "[department_name]" 

# Debugging
# This will run from the parts folder located in the same directory
# And use the test data I created instead
DEBUG = True
DEBUG_FOLDER_PATH = r"DEBUGGING\PATH"
