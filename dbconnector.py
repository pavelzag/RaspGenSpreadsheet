import collections
from logger import logging_handler
from os import uname
from pymongo import MongoClient, errors
from configuration import get_db_creds

env = get_db_creds('env')
test_uri = get_db_creds('test_uri')
prod_uri = get_db_creds('prod_uri')
client = MongoClient(prod_uri)
db = client.raspgen


def get_time_spent(single_date):
    """"Returns a time spent per single date"""
    time_sum_seconds =[]
    cursor = db.time_spent.find({})
    for document in cursor:
        if single_date == document['time_stamp'].date():
            time_sum_seconds.append(document['time_span'])
    return sum(time_sum_seconds)

