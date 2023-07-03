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
Install the required dependencies by running “pip install pandas pymongo openpyxl”
Save the script to a file named excel_to_mongodb.py.
Open a command prompt or terminal and navigate to the directory where the script is saved.
Execute the script using the following command format:
python excel_to_mongodb.py <command> -f <filename> -d <database> -c <collection> -s <source> -dt <date> -sn <sheet_name>
_____
-f or --file: Specifies the path to the Excel file that you want to import or update.
-d or --database: Specifies the name of the MongoDB database where the data will be imported or updated.
-c or --collection: Specifies the name of the MongoDB collection where the data will be imported or updated.
-s or --sor: Specifies the source of the data (e.g., ADP, ADPR, HR, etc.).
-dt or --date: Specifies the date associated with the data in the format 'YYYY-MM-DD'.
-sn or --sheet_name: Specifies the name of the Excel sheet within the file that contains the data.
_______________________________________________________
