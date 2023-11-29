import base64
import os.path
import json
from unidecode import unidecode

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
from email.message import EmailMessage
import mimetypes

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly",
"https://www.googleapis.com/auth/gmail.send"]

sheetsToken = "token.json"

sheetService = None
mailService = None

def init():
	global sheetService
	global mailService
	if sheetService : return
	if os.path.exists(sheetsToken):
		creds = Credentials.from_authorized_user_file(sheetsToken, SCOPES)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				"credentials.json", SCOPES
			)
			creds = flow.run_local_server(port=0)

	# Save the credentials for the next run
	with open(sheetsToken, "w") as token:
		token.write(creds.to_json())

	sheetService = build("sheets", "v4", credentials=creds)
	mailService = build('gmail', 'v1', credentials=creds)


def getSheet( sheetID, sheetRange ):
	# Call the Sheets API
	sheet = sheetService.spreadsheets()
	result = (
		sheet.values()
		.get(spreadsheetId=sheetID, range=sheetRange)
		.execute()
	)
	return result.get("values", [])


def sendMail( msg ):
	init()
	message = EmailMessage()
	for field in [ "To", "Subject", "Cc" ]: message[ field ] = msg[ field ]
	message.set_content( msg["Body"] )

	if "Attachments" in msg:
		files = msg[ "Attachments" ]
		for file in files:
			# guessing the MIME type
			type_subtype, _ = mimetypes.guess_type(file)
			maintype, subtype = type_subtype.split("/")
			message.add_attachment(files[ file ], maintype, subtype,  filename=file)

	create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}
	message = (mailService.users().messages().send(userId="me", body=create_message).execute())
	print(F'sent message to {message} Message Id: {message["id"]}')


def getEmails():
	init()
	SAMPLE_SPREADSHEET_ID = "1GgqSww81cxa0vbT3VTPy3Aq-9gR4wAwxqA48A5gHnoU"
	SAMPLE_RANGE_NAME = "CC_23-24!A:Z"
	SAMPLE_RANGE_NAME2 = "CC_T1!A:Z"
	s = getSheet( SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME )
#	print( s)
	s2 = getSheet( SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME2 )
#	print( s2)

	emails = {}
	for n, line in enumerate( s2[ 1:] ):
		if len( line ) < 8: continue
#		print( "classe :", line[ 0 ] )
		people = []
		for name in [ line[6], line[ 7 ] ]:
			if len( name ) < 4 : continue
			cleanName = unidecode( name ).strip()
			found = False
			email = None
			for l in s:
				if len( l ) < 4 : continue
				if unidecode( l[ 3 ] ).strip() == cleanName:
					found = True
					email = l[ 4 ]

			if not found :
				print( "Error : name not found : " + name )
				exit( 1 )

#			print( name, ":" ,email )
			people.append( { "name" : cleanName, "email" : email } )
		emails[ line[ 0 ] ] = people
	return emails

if __name__ == "__main__":
	msg = {
		"To" : "sebastien.valette.perso@gmail.com",
		"Subject" : "Hello",
		"Body" : "Test de message \n Sebastien"
	}
#	sendMail( msg )
	e =  getEmails()
	print( json.dumps( e, indent = 4 ) )
	
