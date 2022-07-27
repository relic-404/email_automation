import requests
import random
import time
import datetime
import os
from datetime import timedelta
from pathlib import Path
from dotenv import load_dotenv

current_epoch_time = int(time.time())
current_epoch_time = str(current_epoch_time)
og_epoch_time='86406'
num_days = 3
view_data = []
counter = 0


env_path = Path('.') / '.env' #('.') represents current directory
load_dotenv(dotenv_path=env_path)
token = os.environ['ACCESS_TOKEN']

headers = {
    'Authorization': f'Bearer {token}'
}

while True:
    counter+=1
    if counter < num_days:
        lower_timestamp = int(current_epoch_time) - int(og_epoch_time)
        data_field = f"?created[lt]={current_epoch_time}&created[gt]={lower_timestamp}"
        api_base_url = "https://api.wootric.com/v1"
        endpoint_path = f"/responses/{data_field}"
        endpoint = f"{api_base_url}{endpoint_path}"
        r = requests.get(endpoint, headers=headers)
        data = r.json()
        view_data.extend(data)
        
        print(f"Day: {counter}")
        current_epoch_time = lower_timestamp

    if counter > num_days:
        break

# Initialize rows with an initial header row
rows = [['ID', 'Score', 'Text', 'Created At', 'Email', 'CSM', 'Firm ID', 'Firm Name']]

# create keys for each value from the surveys

for survey in view_data:
    row = [
        survey['id'],
        survey['score'],
        survey['text'],
        survey['created_at'],
        survey['end_user']['email'],
    ]
    rows.append(row)


print(view_data)






"""
//r = requests.post(endpoint, headers=headers, json=json_payload)
print(r)
print(rows[1][4])
"""
