import pandas as pd
from mysql.connector import MySQLConnection, Error

insertSql = "insert into students(id, ma_sv, first_name, last_name, birthday, toan, ly, hoa) values (%s,%s,%s,%s,%s,%s,%s,%s)"
readSql = "select * from students"
deleteSql = "delete from students where last_name = %s"
updateSql = "update students set first_name = %s, last_name = %s where id = %s"

data = []

def connect():
    db_config = {
        'host': 'localhost',
        'database': 'py',
        'user': 'root',
        'password': ''
    }
    conn = None
    try:
        conn = MySQLConnection(**db_config)
        if conn.is_connected():
            return conn
    except Error as error:
        print(error)
    return conn

def readExcel():
    try:
        conn = connect()
        cursor = conn.cursor()
        df = pd.read_excel('input.xlsx', sheet_name='MAU', usecols='A:H', skiprows=11, nrows=52)
        for row in df.iterrows():
            row_data = []
            for value in row[1]:
                row_data.append(value)
            val = (row_data[0],row_data[1],row_data[2],row_data[3],row_data[4],row_data[5],row_data[6],row_data[7])
            cursor.execute(insertSql, val)
            conn.commit()
            data.append(row_data)
    except:
        conn.rollback()
        conn.close()

def insert(val):
    try:
        conn = connect()
        cursor = conn.cursor()
        cursor.execute(insertSql, val)
        conn.commit()
    except:
        conn.rollback()
        conn.close()

def getAll():
    try:
        conn = connect()
        cursor = conn.cursor()
        cursor.execute(readSql)
        
        rs = cursor.fetchall()
        
        for x in rs:
            print(x)
    except:
        conn.rollback()
        conn.close()

def update(val):
    try:
        conn = connect()
        cursor = conn.cursor()
        cursor.execute(updateSql, val)
        conn.commit()
    except:
        conn.rollback()
        conn.close()

def delete(val):
    try:
        conn = connect()
        cursor = conn.cursor()
        cursor.execute(deleteSql, val)
        conn.commit()    
    except:
        conn.rollback()
        conn.close()
readExcel()