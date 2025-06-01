import argparse
import csv
import emoji
import json
from fpdf import FPDF
from gapi import sendMail, getEmails
from datetime import datetime
import yaml

with open( 'config.yml', 'r' ) as file:
    config = yaml.safe_load(file)

parser = argparse.ArgumentParser( description = 'Distances computation', formatter_class=argparse.ArgumentDefaultsHelpFormatter )
parser.add_argument( "-to", dest= "sendTo", help="Send to this adress" )
parser.add_argument( "-cc", dest= "copyTo", help="Send copies to this adress", default = [ config[ "APERD_Email" ] ], action = "append" )
parser.add_argument( "-g", "--go", dest= "go", help="send the mails", action = "store_true" )
parser.add_argument( "--pdf", help="only generate pdfs", action = "store_true" )
parser.add_argument( "-v", "--verbose", dest= "verbose", help="verbose output", action = "store_true" )
parser.add_argument('-i','--ignore', nargs='+', help='Ignored groups', default = [])
parser.add_argument('-o','--only', nargs='+', help='Only groups, \"all\" for global digest', default = [])
args = parser.parse_args()

allGroups = "all"

parent_row = config[ "poll"]["parent_row"]
font = config[ "text" ][ "font"]
height = config[ "text" ][ "height" ]
titles = config[ "question_titles" ]

def clean( txt ):
	txt = txt.translate({ord(i): None for i in '\"\u200b'})
	txt = txt.replace( "\u2019", "'" )
	txt = txt.replace( "\u0153", "oe" )
	txt = txt.replace( "\u2026", "... ")
	txt = txt.replace( "\u2014", "-")
	return txt

def getGroup( lines, gr ) :
	group = []
	found = False
	for line in lines[ 1: ]:
		if gr == allGroups or line[ config[ "poll" ][ "class_row" ] ] == gr :
			group.append( line )
			found = True

	if gr == allGroups : group.sort( key= lambda l : l[ config[ "poll" ][ "class_row" ] ] )
	group.insert( 0, lines[ 0 ] )

	if not found :
		print( "Pas de retour pour cette classe" )
		return None
	return group

def printGroup( lines, group, returnPDF = False ) :
	lines = getGroup( lines, group )
	if not lines : return None
	pdf = FPDF()
	pdf.add_page()
	pdf.image( "logo.png", x = 160, w = 35)
	pdf.set_font( font, "BU", size=24)
	if group == allGroups:
		pdf.text( x=10, y=25, txt= "Retour des sondages")
	else:
		pdf.text( txt= "Retour des sondages pour la classe " + group, x=10, y=25)

	pdf.set_font( font, "B", size=18)
	if group == allGroups:
		pdf.write( height, txt= "\n{} réponses: \n\n".format( len(lines) - 1 ))
	else:
		pdf.write( height, txt= "\n{} réponse(s) pour cette classe: \n".format( len(lines) - 1 ))
		pdf.set_font( font, "B", size=10)
		l = lines[ 0 ][ parent_row : parent_row + 4 ]
		pdf.write( height, txt= "\n" + ", ".join( l ) + "\n\n" )

	lastGroup = ""

	# list pupils and parents names
	for i, line in enumerate( lines ):
		if i == 0 : continue
		gr = line[ config[ "poll" ][ "class_row" ] ]
		if gr != lastGroup :
			pdf.set_text_color(0, 0, 0)
			pdf.set_font( font, "BU", size=14)
			if group == allGroups : pdf.write( height, txt= "Classe " + str( gr ) + "\n\n" )
			lastGroup = gr

		pdf.set_text_color(0, 0, 255)
		pdf.set_font( font, "B", size=12)
		pdf.write( height, txt= "(" + str( i ) + ")   " )
		pdf.set_font( font, size=12)
		pdf.write( height, txt= line[ 0 ]  + ", ")
		pdf.write( height, txt= line[ 2 ] + ", ")
		if group == allGroups :
			pdf.write( height, txt= line[ config[ "poll" ][ "class_row" ] ] )
		else:
			l = line[ parent_row : parent_row + 4 ]
			pdf.write( height, txt= ", ".join( l ) )
		pdf.write( height, txt= "\n\n")
		pdf.set_text_color(0, 0, 0)


	# list answers
	categories = lines[ 0 ]
	for col, category in enumerate( categories ):
		if col > 19 and col < 27 : continue

		lastGroup = ""

		for line in range( len( lines ) ):
			if col < config[ "poll" ][ "answers_start_row" ] : continue
			if line == 0 :
				pdf.set_font( font, "B", size=16)
				pdf.write( height, txt= "\n")
				question = str(col)
				if args.verbose : pdf.write( height, txt= question + "\n")
				if col in titles:
					title = titles[ col ]
					pdf.set_font( font, "BU", size=16)
					pdf.write( height, txt= clean(title[0]) + "\n")
					pdf.set_font( font, "I", size=12)
					pdf.write( height, txt= clean(title[1]) + "\n\n")
					pdf.set_font( font, "B", size=16)
					if len( title ) > 3:
						pdf.set_font( font, "BU", size=16)
						pdf.write( height, txt= clean(title[2]) + "\n")
						pdf.set_font( font, "I", size=12)
						pdf.write( height, txt= clean(title[3]) + "\n\n")
						pdf.set_font( font, "B", size=16)

			if line > 0:
				gr = lines[ line ][ config[ "poll" ][ "class_row" ] ]
				if gr != lastGroup :
					pdf.set_text_color(0, 0, 0)
					pdf.set_font( font, "BU", size=14)
					if group == allGroups : pdf.write( height, txt= "Classe " + str( gr ) + "\n\n" )
					lastGroup = gr

				pdf.set_text_color(0, 0, 255)
				pdf.set_font( font, "B", size=12)
				pdf.write( height, txt= "(" + str( line ) + ")   " )
				pdf.set_font( font, size=12)
				pdf.set_text_color(0, 0, 0)

			value = lines[ line ][ col ]
			value = emoji.demojize( value)
			pdf.write( height, txt= value )
			pdf.write( height, txt= "\n")
			if line > 0 and col == 19 and len(value ) > 0 :
				pdf.set_font( font, "B", size=12)
				pdf.write( height, txt= "Raisons évoquées:\n" )
				found = False
				pdf.set_font( font, size=12)
				for reason in range( 20, 27 ):
					if len( lines[ line ][ reason ] ) > 0:
						found = True
						pdf.write( height, txt= categories[ reason ] + "\n" )

				if not found :
					pdf.write( height, txt= "aucune\n" )

			pdf.write( height, txt= "\n")

	if not returnPDF : pdf.output( group + ".pdf")
	else : return pdf.output( dest = "S" )

def getAllGroups( file ):
	data_line = config[ "poll" ][ "data_line" ]
	lines = []
	with open( file ) as f:
		l = csv.reader( f, delimiter="\t" )
		for n, line in enumerate( l ):
			if n < data_line -1 : continue
			if n >= data_line :
				if int( line[ 0 ] ) < config[ "poll" ][ "start" ] : continue
			lines.append( list( map( clean, line ) ) )
		return lines

lines = getAllGroups( config[ "poll" ][ "file" ] )


bodyLines = []
bodyLines.extend( config["email"]["begin"] )
bodyLines.extend( config["email"]["has_polls"] )
bodyLines.extend( config["email"]["end"] )

bodyNoPollsLines = []
bodyNoPollsLines.extend( config["email"]["begin"] )
bodyNoPollsLines.extend( config["email"]["has_no_polls"] )
bodyNoPollsLines.extend( config["email"]["end"] )

body = "\n".join( bodyLines )
bodyNoPolls = "\n".join( bodyNoPollsLines )


groups = args.only
if len( groups ) == 0:
	for n in range( 3, 7 ) :
		for i in range( 1, 5 ) : groups.append( str( n ) + "0" + str( i ) )

groups = list( filter( lambda g : not g in args.ignore, groups ) )

emails = None
if not args.sendTo :
	emails = getEmails( config, groups, args.verbose )
	if args.verbose : print( "Emails : ", emails )

for c in groups:
	print( "Classe : " + c )
	content = printGroup( lines, c, not args.pdf )
#	if c == allGroups : exit( 0 )
	fileName = c + ".pdf"
	now = datetime.now()
	dateStr = now.strftime("%d/%m/%Y %H:%M:%S")

	toSend = []
	if args.sendTo:
		toSend.append( args.sendTo )
	else:
		if c in emails :
			for p in emails[ c ]:
				email = p[ "email" ].strip()
				if len ( email ) == 0 :
					print( '******Warning, email for ' + p[ "name" ] + " is empty******" )
				else:
					toSend.append( p[ "email" ] )

	print( toSend )
	if args.pdf : continue
	bodyToSend = body if content else bodyNoPolls

	msg = {
		"Body" : bodyToSend.format( c, dateStr),
		"To" : ",".join( toSend ),
		"Subject" : config[ "email"][ "subject" ].format( c ),
	}

	if content:	msg[ "Attachments" ] = { fileName : content }

	if not args.sendTo and len( args.copyTo ) : msg[ "Cc" ] = ",".join( args.copyTo )

	if not args.go :
		if content: msg[ "Attachments" ] = fileName
		print( json.dumps( msg, indent = 4 ) )
		print( "Dry mode : email not sent, use -g option to really send mails" )
		continue

	sendMail( msg )

if args.verbose :
	print( "Config : ")
	print( config )
