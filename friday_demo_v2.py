#!/usr/bin/env python

# title: friday_demo.py
# description: Downloads info from supplied URL from product pages and inserts it into copied google slides template
# author: mvrchota
# usage: python friday_demo.py

# importing google slides API part
from __future__ import print_function

from apiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools

# required imports
import datetime
import os
import requests
import re
import smartsheet
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.MIMEBase import MIMEBase

# initialize Smartsheet Auth
token = '4s77fa8n23jfjvlvftse1ioxwi'
ss = smartsheet.Smartsheet(token)

# PP part of the code
PP_API_BASE_URL = 'https://pp.engineering.redhat.com/pp/api/v3'

# Input PP URL from within release info. URl needs to contain .../release/***/.
urlInput = raw_input('Please input release URL.\n')
try:
    buShortname = re.search('/release/(.+?)/', urlInput).group(1)
except AttributeError:
    buShortname = ''
    print('Sorry, you need to input URL link to release')
    quit()
# input date in format YYYY-MM-DD
start_date = raw_input('Please input start date of release (YYYY-MM-DD).\n')
duration = raw_input('How long will it run (days)?\n')
# select which SS template to work with
print("Smartsheet template choices: ")
print("1. Sprint based Schedule")
print("2. Waterfall based Schedule")
template_selection = int(raw_input("Please enter the number corresponding to the workspace in which the sheet is stored: "))

# Setting template sid based on user selection
if template_selection == 1:
    template_id = 2149846049154948 # Sprint based template's sid in WCR_Development workspace
elif template_selection == 2:
    template_id = 5727759965153156 # Waterfall based template's sid in WCR_Development workspace
else:
    print('Wrong selection, template does not exist. Good bye')
    quit()

print("Thank you, we will now proceed with your request.")

def _get_json(url):
    response = requests.get(url,
                            headers=dict(Accept='application/json'),
                            verify=False)
    return response.json()

class Release(object):
    """what info are we getting from Releases API"""
    def __init__(self, rel):
        self.shortname = rel['shortname']
        self.name = rel['name']
        self.rel_id = rel['id']
        self.date = rel['ga_date']
        self._ppl_pm = None
        self._ppl_pmm = None
        self._ppl_pgm = None

        if self.date is None:
            self.date = 'Not specified'

    def __str__(self):
        return "%s (%s (%s))" % (self.name, self.date, self.rel_type)

    @property
    def ppl_pm(self):
        if self._ppl_pm is None:
            self._get_ppl_pm()
        return self._ppl_pm

    def _get_ppl_pm(self):
        prodmans = _get_json(
                '%s/releases/%s/people/' %(PP_API_BASE_URL, self.rel_id))
        rcmlist = set()

        for prodman in prodmans:
            if 'Product Management' in prodman['function__name']:
                rcmlist.add(prodman['user_full_name'])

        self._ppl_pm = ", ".join(rcmlist)
        if not rcmlist:
            print('- - - - - No PM found')
            rcmlist.add('No manager found')
            self._ppl_pm = ", ".join(rcmlist)

    @property
    def ppl_pmm(self):
        if self._ppl_pmm is None:
            self._get_ppl_pmm()
        return self._ppl_pmm

    def _get_ppl_pmm(self):
        markmans = _get_json(
                '%s/releases/%s/people/' %(PP_API_BASE_URL, self.rel_id))
        rcmlist = set()

        for markman in markmans:
            if 'Product Marketing' in markman['function__name']:
                rcmlist.add(markman['user_full_name'])

        self._ppl_pmm = ", ".join(rcmlist)
        if not rcmlist:
            print('- - - - - No PMM found')
            rcmlist.add('No manager found')
            self._ppl_pmm = ", ".join(rcmlist)

    @property
    def ppl_pgm(self):
        if self._ppl_pgm is None:
            self._get_ppl_pgm()
        return self._ppl_pgm

    def _get_ppl_pgm(self):
        progmans = _get_json(
                '%s/releases/%s/people/' %(PP_API_BASE_URL, self.rel_id))
        rcmlist = set()

        for progman in progmans:
            if 'Program Management' in progman['function__name']:
                rcmlist.add(progman['user_full_name'])

        self._ppl_pgm = ", ".join(rcmlist)
        if not rcmlist:
            print('- - - - - No PgM found')
            rcmlist.add('No manager found')
            self._ppl_pgm = ", ".join(rcmlist)

def get_releases():
    """getting releases from PP API in to array"""
    all_rels = _get_json(
        '%s/releases/?shortname=%s' % (PP_API_BASE_URL, buShortname))

    releases = []
    for rel in all_rels:
        releases.append(Release(rel))

    return sorted(releases, key=lambda x: x.date)

# Google Drive Slide template name (needs to be in root)
TMPLFILE = 'friday demo template'
SCOPES = (
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/presentations',
)
store = file.Storage('storage.json')
creds = store.get()
if not creds or creds.invalid:
  flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
  creds = tools.run_flow(flow, store)
HTTP = creds.authorize(Http())
DRIVE = discovery.build('drive', 'v3', http=HTTP)
SLIDES = discovery.build('slides', 'v1', http=HTTP)

nowDate = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
rsp = DRIVE.files().list(q="name='%s'" % TMPLFILE).execute()['files'][0]
for rel in get_releases():
    DATA = {'name': '%s - %s' %(rel.name, nowDate)}
    releaseId = {'id': '%s' %rel.rel_id}
print('- - - - - Copying template %r as %r (%s)' % (rsp['name'], DATA ['name'], releaseId['id']))
DECK_ID = DRIVE.files().copy(body=DATA, fileId=rsp['id']).execute()['id']

print('- - - - - Grabing title slide objects')
slide = SLIDES.presentations().get(presentationId=DECK_ID,
        fields='slides').execute().get('slides', [])[0]
obj = None
for obj in slide['pageElements']:
    if obj['shape']['shapeType'] == 'RECTANGLE':
        break

print('- - - - - Inserting data from release')
for rel in get_releases():
    reqs = [
        {'replaceAllText': {
            'containsText': {'text': '{{product}}'},
            'replaceText': '%s' %rel.name
            }},
        {'replaceAllText': {
            'containsText': {'text': '{{PM}}'},
            'replaceText': '%s' %rel.ppl_pm
            }},
        {'replaceAllText': {
            'containsText': {'text': '{{PMM}}'},
            'replaceText': '%s' %rel.ppl_pmm
            }},
        {'replaceAllText': {
            'containsText': {'text': '{{PgM}}'},
            'replaceText': '%s' %rel.ppl_pgm
            }},
        {'replaceAllText': {
            'containsText': {'text': '{{ga}}'},
            'replaceText': '%s' %(str(rel.date))
            }},
        {'replaceAllText': {
            'containsText': {'text': '{{url}}'},
            'replaceText': '%s' %urlInput
        }},
    ]
    SLIDES.presentations().batchUpdate(body={'requests': reqs},
            presentationId=DECK_ID, fields='').execute()
    print('- - - - - Google Slide Presentation has been created')

# Smartsheet part of script
schedule_name = "schedule_name_" + "%s" % rel.name
# shortlink = buShortname

response = ss.Sheets.copy_sheet(
    template_id,                                    # template sheet id
    ss.models.ContainerDestination({
        'destination_type': 'workspace',
        'destination_id': 8330903337363332,         # wid = WCR_Development
        'new_name': schedule_name,
    }),include = 'all'
)

sid = response.result.id             # response.result.id  # new sheet id
sheet = ss.Sheets.get_sheet(sid)


# Get SS row ids
rows_dict={}
for row in sheet.rows:
    rows_dict[row.row_number]=row.id

# Get SS columns
action1 = ss.Sheets.get_columns(sid)
columns = action1.data
col_dict = {}
for acol in columns:
    col_dict[acol.index]=acol.id

# Update Product name release in template
cell_a = ss.models.Cell()
cell_a.column_id = col_dict[0]
cell_a.value = rel.name
cell_a.strict = False

# Update Finish Date in template
cell_b = ss.models.Cell()
cell_b.column_id = col_dict[2]
cell_b.value = start_date
cell_b.strict = False

# Update the Duration in template
cell_c = ss.models.Cell()
cell_c.column_id = col_dict[1]
cell_c.value = duration
cell_c.strict = False

row_a = ss.models.Row()
row_a.id = rows_dict[1]
row_a.cells.append(cell_a)
row_a.cells.append(cell_b)
row_a.cells.append(cell_c)

row_b = ss.models.Row()
row_b.id = rows_dict[2]
row_b.cells.append(cell_b)
action2 = ss.Sheets.update_rows(sid,[row_a])

print("- - - - - SmartSheet schedule created and filled with data")

# mailing part
# mail properties, sender, receiver, body
print('- - - - - Trying to send an email')
datestamp = str(datetime.datetime.today().strftime('%Y-%m-%d %H:%M:%S'))
mailSender = 'mvrchota.mailing@gmail.com'
mailReciever = 'mvrchota@redhat.com' #,snanda@redhat.com,myarboro@redhat.com'
msg = MIMEMultipart()
msg['Subject'] = 'Demo "Release into Slides and SS" has finished'
msg['From'] = mailSender
msg['To'] = mailReciever
body = 'Script has successfully finished.\n\nYou have requested action for %s release.\n\nPresentation is to be found in Google Drive\nSchedule is located in WCR Development workspace in Smartsheet web UI\n\n%s' % (rel.name, datestamp)
content = MIMEText(body)
msg.attach(content)
# connecting to the server and sending the mail
server = smtplib.SMTP('smtp.gmail.com:587')
server.starttls()
server.login('mvrchota.mailing@gmail.com','mattytestingmail')
server.sendmail(mailSender, mailReciever.split(','), msg.as_string())
print('- - - - - Mail sent to %s at %s' %(mailReciever, datestamp))
server.quit()
