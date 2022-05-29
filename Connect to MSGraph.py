import requests
import json

TenantID = ""
ClientID = ""
ClientSecret = ""

tokenbody = {
    'Grant_Type': 'client_credentials', 
    'Scope': 'https://graph.microsoft.com/.default', 
    'Client_Id': ClientID,
    'Client_Secret': ClientSecret
}
tokenResponse = requests.post(f"https://login.microsoftonline.com/{TenantID}/oauth2/v2.0/token", data=tokenbody)
tokenResponse = tokenResponse.json()
token = tokenResponse['access_token']

headers = {
     'Authorization': f'Bearer {token}',
     'Content-type': 'application/json'
}
url_get = "https://graph.microsoft.com/v1.0/users/"
Users = requests.get(url_get, headers=headers)
