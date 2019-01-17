import sys
import os
import pandas as pd
from pymongo import MongoClient
import json

client = MongoClient()
col = client['test']['test']

df = pd.read_csv('./data.csv')
data = df.to_dict(orient='records')
col.insert_many(data)
