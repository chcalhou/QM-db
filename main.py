#!/usr/bin/python
# -*- coding: utf-8 -*-

import psycopg2
import sys
from xlrd import open_workbook
con = None
program_list =[]
measure_list = []

##This is a bit of code that will parse an excel file and read records to a postgreSQL db.  For this example the data consists
##of clinical quality measures and the government programs asscociated with them.


##Here the function opens the excel workbook .xlsx and the specific spreadsheet then reads the list of program names at the
##top of the sheet.  It iterates over the correct cells using a for loop over a range of the correct cells and returns a list
##of the desired cell values.

def readProgs():
  book = open_workbook('aug_3_measures.xlsx','r')
  sheet = book.sheet_by_index(1)

  for col_index in range(5,sheet.ncols):
    pName = sheet.cell(0,col_index).value
    pDes = sheet.cell(1,col_index).value
    pLink = sheet.cell(2,col_index).value
    program = (pName,pDes,pLink)
    program_list.append(program)
  return program_list


##Here a connection with a db is opened using the psycopg2 module.  Then a create query for a table called programs is written here.
##The readProgs() function is then called to return the correct cell values.  Then the function disects the the list and puts
##the values into an insert statement using a for loop.  Lastly the .commit() writes the statement.

def writeProgs():
    con = psycopg2.connect(database='chriscalhoun', user='chriscalhoun') 
    cur = con.cursor()
    cur.execute("CREATE TABLE programs(\
                id serial PRIMARY KEY,\
                program_name VARCHAR(200), \
                program_description TEXT, \
                program_link VARCHAR(200))")
    program_list=readProgs()

    for i in program_list:
      program_name = i[0]
      program_description = i[1]
      program_link = i[2]
      cur.execute("INSERT INTO programs(program_name, program_description, program_link)\
                  VALUES\
                  ('" + program_name + "', '" + program_description + "', '" + program_link + "')")
      con.commit()
    
##Here we are grabbing a different section of data from the excel spreadsheet using the same method as the readProgs() and returning
##a list with the desired values.

def readMeasures():
  book = open_workbook('aug_3_measures.xlsx','r')
  sheet = book.sheet_by_index(1)

  for row_index in range(3,sheet.nrows):
    u'\xae'.encode('utf-8')

    measure_description = sheet.cell(row_index,3).value
    care_setting = sheet.cell(row_index,4).value
    cms_id = sheet.cell(row_index,0).value
    nqf_id = sheet.cell(row_index,1).value
    pqrs_id = sheet.cell(row_index,2).value
    measure = (measure_description.encode('ascii', 'ignore'),
               care_setting.encode('ascii', 'ignore'),
               str(nqf_id),
               str(pqrs_id),
               str(cms_id))

    measure_list.append(measure)
  return measure_list

##Here we are opening a db and creating a table again called measures this time. Then calling the readMeasures() function
##and using a for loop to itereate over the returned list and writing multiple insert statements before commiting the records.

def writeMeasures():
    con = psycopg2.connect(database='chriscalhoun', user='chriscalhoun') 
    cur = con.cursor()
    cur.execute("CREATE TABLE measures(\
                measure_id serial PRIMARY KEY NOT NULL,\
                measure_description TEXT NOT NULL,\
                care_setting TEXT,\
                nqf_id varchar(20), \
                pqrs_id varchar(30),\
                cms_id varchar(60))")
    
    measure_list = readMeasures()
    for i in measure_list:
      measure_description = i[0]
      care_setting = i[1]
      nqf_id = i[2]
      pqrs_id = i[3]
      cms_id = i[4]

      cur.execute("INSERT INTO measures(measure_description,\
                  care_setting, nqf_id, pqrs_id, cms_id) \
                  VALUES\
                  ('"+measure_description + "', '" + care_setting + "', '" + nqf_id + "', '" + pqrs_id + "', '" + cms_id + "')")
      con.commit()

## Here we are grabbing the data from the excel sheet and checking the cell for certain values. In this example either a cell with 
##data or a blank cell.  Be aware that excel and python data types can vary so it is usually the case that data in your excel file
##will need to be converted to text format.

##The function goes through the correct cells in the specified range of rows and cols.  Then checks if the cell has anything in it.
##If the cell has data in it the data is appended to a list along with the measure name and the program for that measure.
##Then the list is returned.

def measure_program_check():
    checkProgram_list =[]
    book = open_workbook('aug_3_measures.xlsx','r')
    sheet = book.sheet_by_index(0)
    
    
    for row_index in range(3,sheet.nrows):
        for col_index in range(5, 26):
            cell = sheet.cell(row_index,col_index)
            check = cell.value
            if check != '':
                pName = sheet.cell(0,col_index).value
                mDes = sheet.cell(row_index, 3).value
                checkProgram = (True, pName.encode('ascii', 'ignore'), mDes.encode('ascii','ignore'))
                checkProgram_list.append(checkProgram)
##            elif check == '':
##                pName = sheet.cell(0,col_index).value
##                mDes = sheet.cell(row_index, 3).value
##                checkProgram = (False, pName.encode('ascii', 'ignore'), mDes.encode('ascii','ignore'))
##                checkProgram_list.append(checkProgram)
                
                
    return checkProgram_list

##Here we create the table for measure_programs and also create the relationships between the measures and programs tables.
##This table uses to foreign keys as the primary key. 



def measure_program_CreateInsert():
    measure_program_list = []
    con = psycopg2.connect(database='chriscalhoun', user='chriscalhoun') 
    cur = con.cursor()
    cur.execute("CREATE TABLE measure_program(measure_id integer REFERENCES measures (measure_id) ON UPDATE RESTRICT NOT NULL,\
                program_id integer REFERENCES programs (id) ON UPDATE RESTRICT,\
                value BOOLEAN,\
                PRIMARY KEY (measure_id, program_id))")
                
    checkProgram_list = measure_program_check()
    
    ##A query is ran to retrieve the program names and the primary keys associated with them from the programs table.  
    ##They are both paired into a tuple and placed into a list.  The same thing is then repeated with the measure names and their primary keys
    cur.execute("SELECT id, program_name FROM programs")
    prog_idNameList = cur.fetchall()
    
    cur.execute("SELECT measure_id, measure_description FROM measures")
    measure_idNameList = cur.fetchall()
    
    ##Here the relationships for the measure_program table are created.  The list created by calling measure_program_check()
    ##is iterated over using a for loop.  There are two other for loops that compare the values in the lists programs and ids, and the
    ##measures and their ids.  This is to see if the names are the same and when they are assign the appropriate key value in the new table.
    ## A list of tuples with the measure ids and program ids for the correct measures are then returned.
    
    measure_program = []
    for i in range(len(checkProgram_list)):
      for m in measure_idNameList:
        if checkProgram_list[i][2]==m[1]:
          m_id = m[0]
      for p in prog_idNameList:
        if checkProgram_list[i][1] == p[1]:
          p_id = p[0]
          measure_info = (m_id, p_id)
          measure_program.append(measure_info)
    ##Using the execute many function we can simply use string formatting inside the insert statement to populate the query with the 
    ##correct values in each tuple in the list.  
    query = "INSERT INTO measure_program (measure_id, program_id) VALUES(%s, %s)"
   
    cur.executemany(query, measure_program)
    con.commit()
                
def main():
    writeProgs()
    writeMeasures()
    measure_program_check()  
    measure_program_CreateInsert()
    

try:
    main()
    
except psycopg2.DatabaseError, e:
    print('Error %s' % e)    
    sys.exit(1)
    
    
finally:
    
    if con:
        con.close()
