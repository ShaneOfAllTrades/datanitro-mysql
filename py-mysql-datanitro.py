#!/usr/bin/python
import MySQLdb
import sys
import csv

active_sheet("GUI")
db = MySQLdb.connect(host="localhost", user="mysql_user",
                      passwd="noobnoob",db="mysql_db")
# the if elif can be a switch in Excel
if Cell("G1").value == "insert/update":
        active_sheet("active_1")
        row_up = 10
	c_up = db.cursor()
	
	sqy = "INSERT INTO table_name (col_1, col_2) \
	VALUES (%s,) \
	ON DUPLICATE KEY UPDATE \
	structure_for=VALUES(col_1, col_2);"
	while Cell(row_up,1).value is not None:
                c_up.execute(sqy,
                (CellRange((row_up,1),(row_up,2)).value)) #this range will extend to how many columns
                row_up = row_up + 1
        active_sheet("GUI")
# this one will import from database to Excel spreadsheet
elif Cell("G1").value == "select":
        active_sheet("active_1")
        CellRange("A10:W1090").value = None # cleanup existing data        
	c = db.cursor()
	exstring = "select * from table_name where some_id = 1"
	# whatever SELECT you want to use
	c.execute(exstring)
	sh = c.fetchall()
	for i, pos in enumerate(sh):
                Cell(10+i, 1).horizontal = pos
                if Cell(10+i, 12).value is None:
                        Cell(10+i, 12).value = "01/01/2000 12:00"

        active_sheet("GUI")

if Cell("G1").value == "select":
        db.commit()
db.close()