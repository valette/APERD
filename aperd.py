import argparse
import csv
import json
from fpdf import FPDF
from gapi import sendMail, getEmails
from datetime import datetime

parser = argparse.ArgumentParser( description = 'Distances computation', formatter_class=argparse.ArgumentDefaultsHelpFormatter )
parser.add_argument( "-to", dest= "sendTo", help="Send to this adress" )
parser.add_argument( "-cc", dest= "copyTo", help="Send copies to this adress", default = [], action = "append" )
parser.add_argument( "-cr", dest= "classRow", help="class row in file", default = 9, type = int )
parser.add_argument( "-sr", dest= "startRow", help="start row in file", default = 14, type = int )
parser.add_argument( "-pr", dest= "parentRow", help="start row in file", default = 10, type = int )
parser.add_argument( "-l", dest= "linesToDelete", help="start line in file", default = 2, type = int )
parser.add_argument( "-s", dest= "start", help="start line in file", default = 0, type = int )
parser.add_argument( "-g", "--go", dest= "go", help="send the mails", action = "store_true" )
parser.add_argument( "-pdf", dest= "pdf", help="only generate pdfs", action = "store_true" )
parser.add_argument( "-v", "--verbose", dest= "verbose", help="verbose output", action = "store_true" )
parser.add_argument('-i','--ignore', nargs='+', help='Ignored groups', default = [])
args = parser.parse_args()

font = "times"
classRow = args.classRow

titles = {
"14" : [
"La Classe en général", "Avez-vous des remarques à faire relatives aux horaires, à l’ambiance de la classe, la discipline, les contrôles, l’emploi du temps, le déroulement des cours, les absences des professeurs…."
],
"15":[
"La scolarité de votre enfant",
"Ces questions visent à recueillir votre avis sur l'épanouissement de votre enfant dans le collège.",
"D'un point de vue \"scolaire\"",
"Au sujet du programme, du rythme, des horaires, du contrôle des connaissances…",
],
"17":[
"D’un point de vue \"relationnel\"",
"Au sujet de l’ambiance générale, de la discipline, des relations avec les autres élèves ou les professeurs."
],
"19":[
"Dans certaines matières ?",
"Est-ce que dans certaines matières, votre en​fa​nt rencontr​e​ de​s ​difficultés, ou bien vous avez des questions concernant à leur sujet ?​"],
"28":[
"La vie au collège & expression libre",
"​A​v​ez-vo​us d​es ​r​e​m​a​rqu​es à ​f​a​ire c​o​n​cer​n​a​nt l​a ​cantine​, ​l​es re​p​a​s​, ​l​'​a​s​soc​i​a​t​io​n sporti​v​e​, ​le​s ​r​é​création​s​, ​la sécu​r​it​é, ​le s​e​r​v​ice médic​a​l​, ​l​'​accue​i​l​, ​le ​C​DI​, ​les a​c​ti​v​ités ​s​p​or​ti​ves e​t ​c​ulturell​es ​du midi​, ​l​es ​sor​t​ie​s ​p​é​da​g​o​g​iqu​e​s​, ​l​e​s ​é​tudes, l’association des parents… ?"]
}


def clean( txt ):
	txt = txt.translate({ord(i): None for i in '\"\u200b'})
	txt = txt.replace( "\u2019", "'" )
	txt = txt.replace( "\u0153", "oe" )
	txt = txt.replace( "\u2026", "... ")
	txt = txt.replace( "\u2014", "-")
	return txt

def getGroup( lines, gr ) :
	group = []
	group.append( lines[ 0 ] )
	for line in lines[ 1: ]:
		if line[ args.classRow ] == gr : group.append( line )
	return group

def printGroup( lines, group, returnPDF = False ) :
	lines = getGroup( lines, group )
	pdf = FPDF()
	pdf.add_page()
	pdf.image( "logo.png", x = 160, w = 35)
	pdf.set_font( font, "BU", size=24)
	pdf.text( x=10, y=25, text= "Retour des sondages pour la classe " + group)

	pdf.set_font( font, "B", size=18)
	pdf.write( text= "\n{} réponse(s) pour cette classe: \n".format( len(lines) - 1 ))
	pdf.set_font( font, "B", size=10)
	l = lines[ 0 ][ args.parentRow : args.parentRow + 4 ]
	pdf.write( text= "\n" + ", ".join( l ) + "\n\n" )

	# list pupils and parents names
	for i, line in enumerate( lines ):
		if i == 0 : continue
		pdf.set_text_color(0, 0, 255)
		pdf.set_font( font, "B", size=12)
		pdf.write( text= "(" + str( i ) + ")   " )
		pdf.set_font( font, size=12)
		pdf.write( text= line[ 0 ]  + ", ")
		pdf.write( text= line[ 2 ] + ", ")
		l = line[ args.parentRow : args.parentRow + 4 ]
		pdf.write( text= ", ".join( l ) )
		pdf.write( text= "\n\n")
		pdf.set_text_color(0, 0, 0)


	# list answers
	categories = lines[ 0 ]
	for col, category in enumerate( categories ):
		if col > 19 and col < 27 : continue

		for line in range( len( lines ) ): 
			if col < args.startRow : continue
			if line == 0 :
				pdf.set_font( font, "B", size=16)
				pdf.write( text= "\n")
				question = str(col)
				if args.verbose : pdf.write( text= question + "\n")
				if question in titles:
					title = titles[ question ]
					pdf.set_font( font, "BU", size=16)
					pdf.write( text= clean(title[0]) + "\n")
					pdf.set_font( font, "I", size=12)
					pdf.write( text= clean(title[1]) + "\n\n")
					pdf.set_font( font, "B", size=16)
					if len( title ) > 3:
						pdf.set_font( font, "BU", size=16)
						pdf.write( text= clean(title[2]) + "\n")
						pdf.set_font( font, "I", size=12)
						pdf.write( text= clean(title[3]) + "\n\n")
						pdf.set_font( font, "B", size=16)

			if line > 0:
				pdf.set_text_color(0, 0, 255)
				pdf.set_font( font, "B", size=12)
				pdf.write( text= "(" + str( line ) + ")   " )
				pdf.set_font( font, size=12)
				pdf.set_text_color(0, 0, 0)

			value = lines[ line ][ col ]
			pdf.write( text= value )
			pdf.write( text= "\n")
			if line > 0 and col == 19 and len(value ) > 0 :
				pdf.set_font( font, "B", size=12)
				pdf.write( text= "Raisons évoquées:\n" )
				found = False
				pdf.set_font( font, size=12)
				for reason in range( 20, 27 ):
					if len( lines[ line ][ reason ] ) > 0:
						found = True
						pdf.write( text= categories[ reason ] + "\n" )

				if not found :
					pdf.write( text= "aucune\n" )

			pdf.write( text= "\n")

	if not returnPDF : pdf.output( group + ".pdf")
	else : return pdf.output( dest = "S" )

def getAllGroups( file ):
	lines = []
	classes = {}
	with open( file ) as f:
		l = csv.reader( f, delimiter="\t" )
		for n, line in enumerate( l ):
			if n < args.linesToDelete: continue
			if n > args.linesToDelete :
				if int( line[ 0 ] ) < args.start : continue
			lines.append( list( map( clean, line ) ) )
			if n == args.linesToDelete: continue
			classes[ line[ args.classRow ] ] = True
		return lines, classes

lines, classes = getAllGroups( 'aperd.tsv' )

subject = "APERD Dufy {} : Retours Sondage Conseils de Classe"
body = "\n".join( ["Bonjour",
"Ceci est un mail automatique pour vous faire part des remontées du sondage des parents de la classe {} ({}), en vue de préparer le conseil de classe.",
"Pour toute question utilisez de préférence le canal Whatsapp \"APERD Conseils de Classe\"",
"Bonne journée",
"Pour l'APERD",
"Sébastien VALETTE"] )

emails = None
if not args.sendTo :
	emails = getEmails( args.verbose )
	if args.verbose : print( "Emails : ", emails )

for c in sorted( classes ):
	if c in args.ignore : continue
	print( "Classe : " + c )
	content = printGroup( lines, c, not args.pdf )
	fileName = c + ".pdf"
	now = datetime.now()
	dateStr = now.strftime("%d/%m/%Y %H:%M:%S")
	
	toSend = []
	if args.sendTo:
		toSend.append( args.sendTo )
	else:
		for p in emails[ c ]:
			toSend.append( p[ "email" ] )

	print( toSend )
	if args.pdf : continue

	msg = {
		"Body" : body.format( c, dateStr),
		"To" : ",".join( toSend ),
		"Subject" : subject.format( c ),
		"Attachments" : { fileName : content }
	}

	if len( args.copyTo ) : msg[ "Cc" ] = ",".join( args.copyTo )

	if not args.go :
		msg[ 'Attachments' ] = fileName
		print( json.dumps( msg, indent = 4 ) )
		print( "Dry mode : email not sent, use -g option to really send mails" )
		continue

	sendMail( msg )
