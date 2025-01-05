import requests
from datetime import datetime, timedelta

# Go to https://developer.microsoft.com/en-us/graph/graph-explorer
# Constants

access_token=\
'token-here'

# Define the dates
source_date_str = '2024-12-15'
target_date_str = '2025-01-05'

# Calculate the date difference
source_date = datetime.strptime(source_date_str, '%Y-%m-%d')
target_date = datetime.strptime(target_date_str, '%Y-%m-%d')
date_diff = target_date - source_date

# Microsoft Graph API URL
base_url = 'https://graph.microsoft.com/v1.0/me/calendar/events'

# Headers for the request
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# Get the events on the source date
params = {
    '$select': f"subject,categories,start,end",
    '$filter': f"start/dateTime ge '{source_date.isoformat()}' and start/dateTime lt '{(source_date + timedelta(days=1)).isoformat()}'"
}

url = base_url

while url is not None:

    response = requests.get(url, headers=headers, params=params).json() if url is base_url else requests.get(url, headers=headers).json()
    first_time = False

    url = response.get('@odata.nextLink', None)
    # url = None

    events = response.get('value', [])

    print("------")
    

    # Move only the events that are categorized as "Orange"

    for event in events:
        
        if 'categories' in event and 'Orange category' in event['categories']:
            
            start_time = datetime.fromisoformat(event['start']['dateTime'].rstrip('0').rstrip('.'))
            end_time = datetime.fromisoformat(event['end']['dateTime'].rstrip('0').rstrip('.'))
            
            new_start_time = start_time + date_diff
            new_end_time = new_start_time + (end_time - start_time)

            print(event["subject"], start_time, end_time, new_start_time, new_end_time)

            # Update event times

            event_id = event['id']
            update_data = {
                'start': {'dateTime': new_start_time.isoformat(), 'timeZone': event['start']['timeZone']},
                'end': {'dateTime': new_end_time.isoformat(), 'timeZone': event['end']['timeZone']}
            }
            update_url = f'{base_url}/{event_id}'
            update_response = requests.patch(update_url, headers=headers, json=update_data)
            print(f'Event "{event["subject"]}" marked with "Orange" category moved to {target_date_str}.')
    
    print("-----")

print(f'All "Orange" category events on {source_date_str} have been moved to {target_date_str}.')
