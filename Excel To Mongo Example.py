import pandas as pd
import json
from pymongo import MongoClient

# Set up MongoDB connection
client = MongoClient()
db = client['IdentitySystems']
collection = db['data']

# Read the Excel files
#The PID_1589506.xlsx is the file that contains the data, you can replace the name with your own file, and replace
#the sheet names that are needed.
df1 = pd.read_excel('C:\\Internship\\PID_1589506.xlsx', sheet_name='Training_APU_ID_27865')
df2 = pd.read_excel('C:\\Internship\\PID_1589506.xlsx', sheet_name='Production_APU_ID_28686')
df3 = pd.read_excel('C:\\Internship\\PID_1589506.xlsx', sheet_name='From OneLogin 2023 03 28')

# Convert the data to JSON format and add tags
json_data1 = json.loads(df1.to_json(orient='records'))
for record in json_data1:
    record['Application'] = 'eClinicalWorks'
    record['Environment'] = 'Training'

json_data2 = json.loads(df2.to_json(orient='records'))
for record in json_data2:
    record['Application'] = 'eClinicalWorks'
    record['Environment'] = 'Production'

json_data3 = json.loads(df3.to_json(orient='records'))
for record in json_data3:
    record['Application'] = 'OneLogin'
    record['Environment'] = 'Production'

# Insert the JSON data into MongoDB
collection.insert_many(json_data1)
collection.insert_many(json_data2)
collection.insert_many(json_data3)

#The data will import with the collumns being 'tags'