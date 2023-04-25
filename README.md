# Excel-to-MongoDB

The code inputed gives an example of how to transfer the document into mongodb.

After that you can clone the databse and use the pipelines to compare databases or use python scripts which I can provide depending on what we want to accomplish and compare.

client = MongoClient()
db = client['IdentitySystems']
collection = db['data']

This is settuping up to a already known client and database called IdentitySystems and data, you can change the name to what you create.

You'll have to install MongoDB compass locally to do this locally. if it's online you'll have to change theMongoClient() accordingly.

df1 = pd.read_excel('C:\\Internship\\PID_1589506.xlsx', sheet_name='Training_APU_ID_27865')

This just goes to the excel file and the name of the sheet you want to transfer over in the DF (Datafile 1)
It's important that it's a variable since we want to sort by adding records so we know exactly where it's from:
json_data1 = json.loads(df1.to_json(orient='records'))
for record in json_data1:
    record['Application'] = 'eClinicalWorks'
    record['Environment'] = 'Training'

The code above shows how I sorted it.
