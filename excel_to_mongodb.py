import argparse
import pandas as pd
from pymongo import MongoClient

# MongoDB configuration
MONGO_HOST = 'localhost'
MONGO_PORT = 27017

def import_records(file, database, collection, sor, date, sheet_name):

    df = pd.read_excel(file, sheet_name=sheet_name)

    records = df.to_dict(orient='records')

    # Connect to MongoDB
    client = MongoClient(MONGO_HOST, MONGO_PORT)
    db = client[database]
    coll = db[collection]
    coll.insert_many(records)

    # Update SOR and DATE fields in each record
    coll.update_many({}, {"$set": {"SOR": sor, "DATE": date}})

    print(f"Successfully imported {len(records)} records into MongoDB collection '{collection}'")

def compare_and_update_records(file, database, collection, sor, date, sheet_name):
    # Load excel/read
    df = pd.read_excel(file, sheet_name=sheet_name)

    # Convert
    new_records = df.to_dict(orient='records')

    # Connect to MongoDB
    client = MongoClient(MONGO_HOST, MONGO_PORT)
    db = client[database]
    coll = db[collection]

    existing_records = coll.find()

    existing_names = set(record['NAME'] for record in existing_records)

    # Compare and update records
    updated_count = 0
    added_names = []
    removed_names = []
    for record in new_records:
        name = record['NAME']
        if name in existing_names:
            coll.update_one({"NAME": name}, {"$set": {"DATE": date, **record}})
            updated_count += 1
        else:
            coll.insert_one({**record, "SOR": sor, "DATE": date})
            added_names.append(name)

    coll.delete_many({"NAME": {"$nin": [record['NAME'] for record in new_records]}})
    removed_names = existing_names - set(record['NAME'] for record in new_records)

    updated_names = set(record['NAME'] for record in new_records) & existing_names

    print(f"Successfully updated {updated_count} records in MongoDB collection '{collection}'")

    # Convert added_names and removed_names to strings
    added_names = [str(name) for name in added_names]
    removed_names = [str(name) for name in removed_names]

    # Generate report
    report = {
        'Total Records': len(new_records),
        'Updated Records': updated_count,
        'Unchanged Records': len(new_records) - updated_count,
        'Added Names': added_names,
        'Removed Names': removed_names
    }

    print("\nReport:")
    for key, value in report.items():
        if isinstance(value, list):
            value = ', '.join(value)
        print(f"{key}: {value}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Import or update records in MongoDB from an Excel file')
    subparsers = parser.add_subparsers(dest='command')

    # Import command
    import_parser = subparsers.add_parser('import', help='Import records into MongoDB from an Excel file')
    import_parser.add_argument('-f', '--file', required=True, help='Path to the Excel file')
    import_parser.add_argument('-d', '--database', required=True, help='Name of the MongoDB database')
    import_parser.add_argument('-c', '--collection', required=True, help='Name of the MongoDB collection')
    import_parser.add_argument('-s', '--sor', required=True, help='Value for the SOR field')
    import_parser.add_argument('-dt', '--date', required=True, help='Value for the DATE field')
    import_parser.add_argument('-sn', '--sheet_name', default=0, help='Name or index of the sheet in the Excel file')

    # Update command
    update_parser = subparsers.add_parser('update', help='Update records in MongoDB from an Excel file')
    update_parser.add_argument('-f', '--file', required=True, help='Path to the Excel file')
    update_parser.add_argument('-d', '--database', required=True, help='Name of the MongoDB database')
    update_parser.add_argument('-c', '--collection', required=True, help='Name of the MongoDB collection')
    update_parser.add_argument('-s', '--sor', required=True, help='Value for the SOR field')
    update_parser.add_argument('-dt', '--date', required=True, help='Value for the DATE field')
    update_parser.add_argument('-sn', '--sheet_name', default=0, help='Name or index of the sheet in the Excel file')

    args = parser.parse_args()

    if args.command == 'import':
        import_records(args.file, args.database, args.collection, args.sor, args.date, args.sheet_name)
    elif args.command == 'update':
        compare_and_update_records(args.file, args.database, args.collection, args.sor, args.date, args.sheet_name)