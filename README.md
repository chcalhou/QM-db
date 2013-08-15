QM-db
=====

This is a way of using python with xlrd to read an excel spreadsheet and commit records to a postgreSQL database.  


What you'll need
================
 -python2.7
 -psycopg2
 -xlrd
 -data in an excel file
  ==> link to my example data (https://iu.box.com/s/etfr6kp0m0u1eg8tme13)

Example
=======

The data used in my example is of clinical quality measure data.  The spreadsheet has a list of government clinical quality 
programs in the top row.  Then going down the rows there is a quality measure on each row with associated id numbers and info.
Under each program if there is a "1" in a cell in the same column indicates that measure is in that program.  
