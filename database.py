import mysql.connector
from mysql.connector import MySQLConnection, Error
from python_mysql_dbconfig import read_db_config

import openpyxl


db_config = read_db_config()
conn = None

def connect():
    """ Connect to MySql database"""

    try:
        print("Connecting to MySQL database...")
        conn = MySQLConnection(**db_config)

        if conn.is_connected():
            print("Connected established.")
        else:
            print("Connected fail.")

    except Error as error:
        print(error)

    finally:
        if conn is not None and conn.is_connected():
            conn.close()
            print("Connection closed.")

def create_table():
    """ Create table """

    try:
        conn = MySQLConnection(**db_config)
        cursor = conn.cursor()

        query = f"CREATE TABLE Users (id int PRIMARY KEY AUTO_INCREMENT, name VARCHAR(50), passwd VARCHAR(50))"

        cursor.execute(query)

        for x in cursor:
            print(x)
    except Error as error:
        print(error)
    
    finally:
        cursor.close()
        conn.close()

def insert_into_table(name: str, password: str):
    """ Insert into table """

    try:
        conn = MySQLConnection(**db_config)
        cursor = conn.cursor()

        query = "INSERT INTO Users (name, passwd) VALUES(%s, %s)"
        args = (name, password)
        cursor.execute(query, args)

        if cursor.lastrowid:
            print("Last insert id: ", cursor.lastrowid)
        else:
            print("Last insert id not found")

        conn.commit()

    except Error as error:
        print(error)
    
    finally:
        cursor.close()
        conn.close()

def update_password(password, id_):
    """ Update password from table using id """
    
    query = """UPDATE Users 
                SET passwd = %s
                WHERE id = %s """

    args = (password, id_)

    try:
        conn = MySQLConnection(**db_config)
        cursor = conn.cursor()
        cursor.execute(query, args)

        conn.commit()

    except Error as e:
        print(e)

    finally:
        cursor.close()
        conn.close()

def update_name(name, id_):
    """ Update name from table using id """
    
    query = """UPDATE Users 
                SET name = %s
                WHERE id = %s """

    args = (name, id_)

    try:
        conn = MySQLConnection(**db_config)
        cursor = conn.cursor()
        cursor.execute(query, args)

        conn.commit()

    except Error as e:
        print(e)

    finally:
        cursor.close()
        conn.close()

def delete_from_table(id_):
    """ Delete from table using the id"""

    query = "DELETE FROM Users WHERE id = %s"
    args = (id_,)

    try:
        conn = MySQLConnection(**db_config)

        cursor = conn.cursor()
        cursor.execute(query, args)

        conn.commit()

    except Error as error:
        print(error)

    finally:
        cursor.close()
        conn.close()

def fetchall():
    """ View table """

    try:
        conn = MySQLConnection(**db_config)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Users")
        rows = cursor.fetchall()

        print("Total Row(s): ", cursor.rowcount)
        for row in rows:
            print(row)

    except Error as error:
        print(error)

    finally:
        cursor.close()
        conn.close()

def export_to_excel():
    """ Export to excel """

    conn = MySQLConnection(**db_config)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Users")
    rows = cursor.fetchall()
    records = rows
    workbok = openpyxl.Workbook()
    sheet = workbok.active

    for row_index, row in enumerate(records):
        sheet.insert_rows(idx=row_index + 1)
        for column_index, cell_value in enumerate(row):
            sheet.cell(row=row_index + 1, column=column_index + 1, value=cell_value)
    
    workbok.save("exports/users.xls")
    return "Exported to file users.xls"