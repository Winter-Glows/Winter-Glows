import json
from pprint import pprint


def jsontodict(file: str) -> dict:
    if file.endswith('.json'):
        with open(file, encoding='UTF-8') as f:
            record = json.load(f)
            # print(record)
    else:
        raise TypeError("Only support .json !")
    return record

if __name__ == '__main__':
    file = 'json/test_immunosuppression.json'
    pprint(jsontodict(file))

# import os
# print(os.getcwd())