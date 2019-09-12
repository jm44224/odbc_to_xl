import pyodbc


def get_database_connection(connection_string):
    """ creates a connection for use in pyodbc calls

    Args:

        connection_string (str): full strinng, so the user can make it trusted or not

    Returns:

        a pyodbc connect object
    """
    conn = pyodbc.connect(connection_string)

    try:
        conn = pyodbc.connect(connection_string)
    except pyodbc.Error:
        raise Exception(
            f"A connection was not successful to the database.")
    return conn


def get_database_result_set(qry_cursor, qry_text):
    """ queries the database
        adds dictionary to result

    Args:
        qry_cursor (openpyxl connection cursor)
        qry_text (str): The query that will be executed
    Returns:
        results from query (dictionary)
    """
    # query database
    qry_cursor.execute(qry_text)
    # reformat results into a dictionary
    results = []
    # https://stackoverflow.com/questions/16519385/output-pyodbc-cursor-results-as-python-dictionary
    columns = [column[0] for column in qry_cursor.description]
    for row in qry_cursor.fetchall():
        results.append(dict(zip(columns, row)))
    return results


def close_cursor_and_connection(csr, connection):
    """ close pyodbc cursor and connection

    Args:
        csr (pyodbc cursor object)
        connection (pyodbc connection object)

    Returns:
        None
    """
    csr.close()
    connection.close()
