#!/usr/bin/python

"""
This code will query a SQL Server database, and export results to an XLSX file
The query results will be on the first Sheet
The text of the query will be on the second Sheet

It was created while learning from Al Sweigart's Automate the Boring Stuff with Python
and this web site: https://datatofish.com/how-to-connect-python-to-sql-server-using-pyodbc/
"""


import access_odbc
import query_results_to_xl
import odbc_to_xl_configuration

# process of querying database and outputting to a worksheet
workbook = query_results_to_xl.get_workbook()
conn = access_odbc.get_database_connection(odbc_to_xl_configuration.connection_string)
cursor = conn.cursor()
query_results_to_xl.create_query_results_sheets_in_workbook(workbook, odbc_to_xl_configuration.query)
""" the above creates the XLSX in memory, with the two Sheets """
result_set = access_odbc.get_database_result_set(cursor, odbc_to_xl_configuration.query)
""" result_set is the data returned by the query """
query_results_to_xl.write_results_to_workbook(workbook, result_set, odbc_to_xl_configuration.results_layout)
""" the above writes the query results to the XLSX that is in memory """
query_results_to_xl.save_workbook_to_file(workbook, odbc_to_xl_configuration.file_path, odbc_to_xl_configuration.file_name)
""" the above saves the XLSX file """
access_odbc.close_cursor_and_connection(cursor, conn)


__author__ = 'Joseph Mate'
__copyright__ = '2019'
__credits__ = ['Joseph Mate']
__license__ = 'GPL'
__version__ = '1.0.1'
__maintainer__ = 'Joseph Mate'
__email__ = 'N/A'
__status__ = 'Development'
