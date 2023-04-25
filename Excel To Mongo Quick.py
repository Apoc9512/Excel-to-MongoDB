import pandas as pd
import json
from pymongo import MongoClient

# Load Excel sheets
training_data = pd.read_excel('training_data.xlsx', sheet_name='Training_APU_ID_27865')
production_data = pd.read_excel('production_data.xlsx', sheet_name='Production_APU_ID_28686')

# Convert to JSON format
training_data_json = json.loads(training_data.to_json(orient='records'))
production_data_json = json.loads(production_data.to_json(orient='records'))

# Connect to MongoDB
client = MongoClient('localhost', 27017)
db = client['mydatabase']
collection = db['mycollection']

# Insert data into MongoDB
collection.insert_many(training_data_json)
collection.insert_many(production_data_json)