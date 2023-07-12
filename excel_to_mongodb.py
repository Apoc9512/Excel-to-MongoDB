import pandas as pd
import argparse
from pymongo import MongoClient

# Config
HOST = 'localhost' 
PORT = 27017

def import_records(file, db_name, coll_name, sor, date, sheet_name):

  df = pd.read_excel(file, sheet_name=sheet_name)
  records = df.to_dict(orient='records')

  client = MongoClient(HOST, PORT)
  db = client[db_name]
  coll = db[coll_name]

  coll.insert_many(records)
  coll.update_many({}, {'$set': {'SOR': sor, 'DATE': date}})

  print(f'Imported {len(records)} records')

def update_records(file, db_name, coll_name, sor, date, sheet_name, output_file):
  
  df = pd.read_excel(file, sheet_name=sheet_name)
  df['NAME'] = df['NAME'].astype(str)

  docs = df.to_dict(orient='records')

  client = MongoClient(HOST, PORT)
  db = client[db_name]
  coll = db[coll_name]

  added = []
  removed = []

  for doc in docs:
    name = doc['NAME']
    if not isinstance(name, str):
      print(f'Invalid name: {name}')
      continue

    if coll.count_documents({'NAME': name}) > 0:
      coll.update_one({'NAME': name}, {'$set': doc})
    else:
      coll.insert_one({**doc, 'SOR': sor, 'DATE': date})
      added.append(name)

  removed = [n for n in coll.distinct('NAME') if n not in [d['NAME'] for d in docs]]

  report = {
    'Total': len(docs),
    'Updated': len(docs) - len(added),
    'Added': added, 
    'Removed': removed
  }

  if output_file:
    with open(output_file, 'w') as f:
      for k, v in report.items():
        try:
          v = ', '.join(v) 
        except TypeError:
          v = str(v)
        f.write(f'{k}: {v}\n')

  else:  
    print(report)


if __name__ == '__main__':
  
  parser = argparse.ArgumentParser()
  subparsers = parser.add_subparsers(dest='command')

  import_parser = subparsers.add_parser('import')
  import_parser.add_argument('-f', required=True)
  import_parser.add_argument('-d', required=True)
  import_parser.add_argument('-c', required=True)
  import_parser.add_argument('-s', required=True)
  import_parser.add_argument('-dt', required=True)
  import_parser.add_argument('-sn', default=0)

  update_parser = subparsers.add_parser('update')
  update_parser.add_argument('-f', required=True)
  update_parser.add_argument('-d', required=True)
  update_parser.add_argument('-c', required=True)
  update_parser.add_argument('-s', required=True)
  update_parser.add_argument('-dt', required=True)
  update_parser.add_argument('-sn', default=0)
  update_parser.add_argument('-o', required=False)

  args = parser.parse_args()

  if args.command == 'import':
    import_records(args.f, args.d, args.c, args.s, args.dt, args.sn)

  elif args.command == 'update':
    update_records(args.f, args.d, args.c, args.s, args.dt, args.sn, args.o)