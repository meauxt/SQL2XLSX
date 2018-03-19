
#########################################################################################
##                                                                                     ##
##                                  Script Parameters                                  ##
##                                                                                     ##
#########################################################################################

databaseName = "databasename"
username = "username"
password = "password"
host= ""
port= 0
query=""""SELECT * FROM TABLE"""
fileName=""

#########################################################################################
##                                                                                     ##
##                                  Script Libraries                                   ##
##                                                                                     ##
#########################################################################################

import psycopg2
import xlsxwriter
import numbers
import datetime

#########################################################################################
##                                                                                     ##
##                   Connecting to DB and generating xlsx files                        ##
##                                                                                     ##
#########################################################################################
try:

    if databaseName is ''  or username is '' or host is '' or password is '':
        raise Exception("Script Parameters are missin...please edit SQL2XLSX file and add the Database information")
    conn = psycopg2.connect("dbname='"+databaseName+"' user='"+username+"' host='"+host+"' port='"+str(port)+"' password='"+password+"'")
    cur = conn.cursor()
    workbook   = xlsxwriter.Workbook(fileName+".xlsx")
    worksheet = workbook.add_worksheet()
    cur.execute(query)
    elements = cur.fetchall()
    # Adding the the column name from the query as first column 
    colnames = [desc[0] for desc in cur.description]
    cell_format = workbook.add_format({'bold': True, 'locked':True})
    for idx, colName in enumerate(colnames): 
        worksheet.write(0, idx, str(colName).decode('utf-8').strip(),cell_format)  
    # Looping in result set and writing the rows to the excel sheet    
    row = 1
    for element in elements:
        column = 0
        for cell in element:
            # check if it's number then write the cell as is
            if isinstance(cell, numbers.Number):
                worksheet.write(row, column, cell) 
            else:
                worksheet.write(row, column, str(cell).decode('utf-8').strip())  
    workbook.close()
    conn.close()
except Exception as e:
    print e