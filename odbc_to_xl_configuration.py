import json

#   a JSON file can be added to a gitIgnore
#   and modified as desired without altering the PY files
with open('odbc_to_xl.json', 'r') as file:
    configuration = json.load(file)

# set variables
connection_string = configuration['connection_string']
""" see https://code.google.com/archive/p/pyodbc/wikis/ConnectionStrings.wiki """
query = configuration['query']
""" query is the query to execute against database_name on server_name """
results_layout = configuration['layout']
""" this is a layout of dictionaries """
""" each dictionary must have a numeric column, values between 1 and 16384 """
""" each dictionary optionally can have a header, that will appear in Row 1  """
""" each dictionary must have a field and that field must be returned by the database query  """

file_path = configuration['file_path']
file_name = configuration['file_name']
""" file_path and file_name will be the XLSX output file """
