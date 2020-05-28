import json
import os

file = os.path.abspath('../res/report_setup.json')
with open(file, 'r', encoding='utf-8') as task:
    setup_dict = json.load(task)
print (setup_dict)
