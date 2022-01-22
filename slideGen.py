from __future__ import print_function

from apiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools

SCOPES = (
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
)
store = file.Storage('storage.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)
HTTP = creds.authorize(Http())
SHEETS = discovery.build('sheets', 'v4', http=HTTP)
SLIDES = discovery.build('slides', 'v1', http=HTTP)

print('** Fetch Sheets data')
sheetID = '1zrh5YL5zmUa5QSVZkah24ShrQoy4e0pu1KXCc-Z5ARM'   # use your own!
orders = SHEETS.spreadsheets().values().get(range='Form Responses 1',
        spreadsheetId=sheetID).execute().get('values')

print('** Fetch chart info from Sheets')
sheet = SHEETS.spreadsheets().get(spreadsheetId=sheetID,
        ranges=['Form Responses 1']).execute().get('sheets')[0]

print('** Create new slide deck')
DATA = {'title': 'Rush Powerpoint'}
rsp = SLIDES.presentations().create(body=DATA).execute()
deckID = rsp['presentationId']
titleSlide = rsp['slides'][0]
titleID = titleSlide['pageElements'][0]['objectId']
subtitleID = titleSlide['pageElements'][1]['objectId']

print('** Create slides & insert slide deck title+subtitle')



reqs = [
  {'insertText': {'objectId': titleID,    'text': 'Rush 2022'}},
 #{'insertText': {'objectId': subtitleID, 'text': ''}},
]
rsp = SLIDES.presentations().batchUpdate(body={'requests': reqs},
        presentationId=deckID).execute().get('replies')

i = 1



while i < len(orders):
    rusheeInfo = orders[i]
    rusheeName = rusheeInfo[1] + " " + rusheeInfo[2] + " #" + str(i)
    bio = "Year: " + rusheeInfo[6] + "\nHometown: " + rusheeInfo[7] + "\nHigh School: " + rusheeInfo[8] + "\nUgliest Brother: " + rusheeInfo[10] + "\nHeads or Tails: " + rusheeInfo[11]
    image = rusheeInfo[4]
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
#tableSlideID = rsp[0]['createSlide']['objectId']
#chartSlideID = rsp[1]['createSlide']['objectId']
