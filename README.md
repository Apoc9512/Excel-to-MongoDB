_____________________________________________________
Features of the script:
–	Imports data from an Excel file into a MongoDB collection.
–	Updates existing records in the collection based on matching names.
–	Adds new records to the collection.
–	Removes records from the collection if the name is not present in the new data.
–	Generates a report showing the total number of records, updated records, unchanged records, added names, and removed names.
–	SOR to add in the record along with Date in commands.
_____________________________________________________
Commands/ how to use:
Ensure you have Python installed on your system.
Install the required dependencies by running “pip install pandas pymongo”
_____
Commands:

import - Insert new records from Excel into MongoDB
update - Sync MongoDB records by comparing to Excel
Arguments:

-f FILE - Path to Excel file
-d DB_NAME - MongoDB database name
-c COLL_NAME - MongoDB collection name
-s SOR - Value for SOR (Source of Records) field
-dt DATE - Value for DATE field
-sn SHEET_NAME - Name or index of sheet in Excel file
-o FILE - Path to write update report (update only)
import usage:

Inserts all records from the Excel sheet into the specified collection
Adds SOR and DATE fields to each inserted document

EX:
python script.py update -f newdata.xlsx -d mydb -c records -s ERP -dt 2022-02-01 -sn 'Data' -o report.txt

This compares the Excel sheet 'Data' to the mydb.records collection and updates accordingly. The report is written to report.txt.
