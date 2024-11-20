import base64
import os.path
import json
from unidecode import unidecode
import yaml

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
	for field in [ "To", "Subject", "Cc" ]:
		if field in msg : message[ field ] = msg[ field ]
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


def getEmails( config, ignore = [], verbose=False ):
	if verbose : print( "Récuperation des emails..." )
	init()
	SAMPLE_SPREADSHEET_ID = config["spreadsheet"]["id"]
	SAMPLE_RANGE_NAME = config["spreadsheet"]["names_range"]
	SAMPLE_RANGE_NAME2 = config["spreadsheet"]["current_range"]
	members = getSheet( SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME )
	if verbose : print( "Membres : ", members)
	classes = getSheet( SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME2 )
	if verbose : print( "Classes : ", classes)

	emails = {}
	for n, line in enumerate( classes[ 1:] ):
		if len( line ) < 6: continue
		group = line[ 0 ].split( " " )[ 0 ]
		if group in ignore : continue
		if verbose : print( "classe :", group )
		people = []

		for name in line[ 6: ]:
			if len( name ) < 4 : continue
			cleanName = unidecode( name.replace( "?", "" ) ).strip()
			cleanName = " ".join( cleanName.split( "-" ) )
			found = False
			email = None
			for l in members:
				if len( l ) < 4 : continue
				tableName = " ".join( unidecode( l[ 3 ] ).strip().split( "-" ) )
				tableName = unidecode( l[ 2 ] ) + " " + unidecode( l[ 1 ] )
				tableName = " ".join( tableName.split( "-" ) )

				if tableName == cleanName:
					found = True
					email = l[ 4 ]

				tableName = " ".join( unidecode( l[ 3 ] ).strip().split( "-" ) )
				if tableName == cleanName:
					found = True
					email = l[ 4 ]

				tableName = unidecode( l[ 1 ] ) + " " + unidecode( l[ 2 ] )
				tableName = " ".join( tableName.split( "-" ) )
				if tableName == cleanName:
					found = True
					email = l[ 4 ]

			if not found :
				print( "Error : name not found : " + name )
				exit( 1 )

#			print( name, ":" ,email )
			people.append( { "name" : cleanName, "email" : email } )
		if verbose : print( "emails :", people )
		emails[ group ] = people
	if verbose : print( "Emails trouvés" )
	return emails

if __name__ == "__main__":
	msg = {
		"To" : "sebastien.valette.perso@gmail.com",
		"Subject" : "Hello",
		"Body" : "Test de message \n Sebastien"
	}
#	sendMail( msg )
	with open( 'config.yml', 'r' ) as file:
		config = yaml.safe_load(file)
	e =  getEmails( config )
	print( json.dumps( e, indent = 4 ) )
	
