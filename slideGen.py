from __future__ import print_function

from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools
import json
import requests

SCOPES = (
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
)

store = file.Storage('storage.json')
creds = store.get()
if not creds or creds.invalid:
    print('invalid')
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)
HTTP = creds.authorize(Http())
print("authorized")
SHEETS = discovery.build('sheets', 'v4', http=HTTP)
SLIDES = discovery.build('slides', 'v1', http=HTTP)
service = discovery.build('drive', 'v3', http=HTTP)

print('** Fetch Sheets data')
sheetID = '1ze6RSxdDzTyK3Zo7wN9NPm6rDCnjft76d3SzUaJq3vw'   # UPDATE WITH NEW SPREADSHEET
orders = SHEETS.spreadsheets().values().get(range='Open House',
        spreadsheetId=sheetID).execute().get('values')
print(orders)

print('** Fetch chart info from Sheets')
sheet = SHEETS.spreadsheets().get(spreadsheetId=sheetID,
        ranges=['Open House']).execute().get('sheets')[0]

print('** Create new slide deck')

slideshowid = "1QOr6e6Ple7z5ND9eGFy1_OxLcKWaImpgYRvQ-h1PDDs"
rsp = SLIDES.presentations().get(presentationId=slideshowid).execute()

deckID = rsp['presentationId']
titleSlide = rsp['slides'][0]
titleID = titleSlide['pageElements'][0]['objectId']
subtitleID = titleSlide['pageElements'][1]['objectId']

print('** Create slides & insert slide deck title+subtitle')



reqs = [
  {'insertText': {'objectId': titleID,    'text': 'Rush 2024'}},
 #{'insertText': {'objectId': subtitleID, 'text': ''}},
]
rsp = SLIDES.presentations().batchUpdate(body={'requests': reqs},
        presentationId=deckID).execute().get('replies')

i = 1


while i < len(orders):
    rusheeInfo = orders[i]
    rusheeName = rusheeInfo[1] + " " + rusheeInfo[2] + " #" + str(i)
    print(i)
    print(rusheeName)
    bio = "Year: " + rusheeInfo[7] + "\nHometown: " +  rusheeInfo[8] + "\nHigh School: " + rusheeInfo[9] + "\nHidden Talent: " + rusheeInfo[11]
    image = rusheeInfo[3]
    image_id = image.index("id=") + 3
    image_url = 'https://drive.google.com/uc?export=view&id=' + image[image_id:]  #https://drive.google.com/uc?export=view&id=' + image[image_id:]

                 
    reqs = [
            {'createSlide': {'slideLayoutReference' : {'predefinedLayout': 'TITLE_AND_TWO_COLUMNS'}}}
        ]
    rsp = SLIDES.presentations().batchUpdate(body={'requests': reqs},
        presentationId=deckID).execute().get('replies')

    nextSlideId = rsp[0]['createSlide']['objectId']

    rsp = SLIDES.presentations().pages().get(presentationId=deckID,
        pageObjectId=nextSlideId).execute().get('pageElements')

    titleBoxId = rsp[0]['objectId']
    textBoxIDLeft = rsp[1]['objectId']
    textBoxIDRight = rsp[2]['objectId']

    reqs = [
        {'insertText': {'objectId' : titleBoxId, 'text' : rusheeName}},
        {'insertText': {'objectId' : textBoxIDRight, 'text' : bio}},
        {'createImage' : { 'url': image_url, 'elementProperties': {'pageObjectId': nextSlideId, "size": {
                "height": {
                    "magnitude": 4000000,
                    "unit": "EMU"
                },
                "width": {
                    "magnitude": 3000000,
                    "unit": "EMU"
                }
            },
            "transform": {
                "unit": "EMU",
                "scaleX": 1.3449,
                "scaleY": 1.3031,
                "translateX": 4671925/16,
                "translateY": 490150
            }
        }
    }}
    ]
    rsp = SLIDES.presentations().batchUpdate(body = {'requests' : reqs},
            presentationId=deckID).execute().get('replies')
    i += 1

