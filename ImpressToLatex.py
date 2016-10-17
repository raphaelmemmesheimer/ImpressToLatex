#!/usr/bin/env python
# coding: utf-8

#This program is free software; you can redistribute it and/or modify
#it under the terms of the GNU General Public License as published by
#the Free Software Foundation; either version 2 of the License, or
#(at your option) any later version.

#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.

#You should have received a copy of the GNU General Public License
#along with this program; if not, write to the Free Software
#Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.

#Copyright (C) Raphael Memmesheimer , 2012

	# start libreoffice like : libreoffice --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"

import uno
import unohelper
import os
import sys
import subprocess
import getopt
import argparse
import ast
from com.sun.star.beans import PropertyValue
from com.sun.star.beans.PropertyState import DIRECT_VALUE

impressToLatexVersion = "0.0"


#found @ http://user.services.openoffice.org/en/forum/viewtopic.php?f=45&t=44220
def writePNG(desktop, doc, url, ctx, element):
	smgr = ctx.getServiceManager()
	# choose drawings which you want to export
	collection = smgr.createInstanceWithContext("com.sun.star.drawing.ShapeCollection", ctx)
	collection.add(element)

	args = []
	args.append(PropertyValue("MediaType", 0, "image/png", DIRECT_VALUE))
	args.append(PropertyValue("URL", 0, url, DIRECT_VALUE))

	graphic_filter = smgr.createInstanceWithContext(
			"com.sun.star.drawing.GraphicExportFilter", ctx)
	graphic_filter.setSourceDocument(collection)
	#print( "[GRAPHIC SUPPORTED EXPORT FILES]"+str(graphic_filter.getSupportedMimeTypeNames()))
	graphic_filter.filter(tuple(args))


def writeEPS(desktop, doc, url, ctx, element):
	smgr = ctx.getServiceManager()
	collection = smgr.createInstanceWithContext("com.sun.star.drawing.ShapeCollection", ctx)
	collection.add(element)

	filterData = []
	filterData.append(PropertyValue("Level", 0, 2, DIRECT_VALUE))
	filterData.append(PropertyValue("ColorFormat", 0, 1, DIRECT_VALUE))
	filterData.append(PropertyValue("TextMode", 0, 2, DIRECT_VALUE))
	filterData.append(PropertyValue("Preview", 0, 1, DIRECT_VALUE))
	filterData.append(PropertyValue("CompressionMode", 0, 1, DIRECT_VALUE))
	filterData.append(PropertyValue("FileFormatVersion", 0, 2, DIRECT_VALUE))

	args = []
	args.append(PropertyValue("MediaType", 0, "image/x-eps", DIRECT_VALUE))
	args.append(PropertyValue("URL", 0, url, DIRECT_VALUE))

	graphic_filter = smgr.createInstanceWithContext(
			"com.sun.star.drawing.GraphicExportFilter", ctx)
	graphic_filter.setSourceDocument(collection)
	graphic_filter.filter(tuple(args))


def processText(s): # replace special characters
	for r in replaceCharTable:
		s = s.replace(r[0], r[1])
	return str(s)

def isFloat(s):
    return True
	# try: return (float(s) == float(s))
	# except (ValueError, TypeError), e: return False

#info
#if len(sys.argv) <= 1:
	#usage()
	#sys.exit(0)

#os.system('/opt/openoffice.org3/program/soffice -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" &')

inputFilename = ""
outputFilename = ""
parse_section = False
verbose = False
debug = False
if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='Converting an (Libre|Open) Office Impress readable presentation to a Latex file',epilog='https://github.com/airglow/ImpressToLatex')
	parser.add_argument('input', metavar='InputFilename', type=str,  help='the file to convert, preferably .odp')
	parser.add_argument('output', metavar='OutputFilename', type=str,  help='the resulting .tex file')
	parser.add_argument('-start', metavar='start', type=int,  help='Start Page', default = 0)
	parser.add_argument('-end', metavar='end', type=int,  help='End Page', default = -1)
	parser.add_argument('--version',dest="version", action="store_true", help="print( version")
	parser.add_argument('-v', '--verbose', dest='verbose', action='store_true', help="show some processing info")
	parser.add_argument('-d', '--debug', dest='debug', action='store_true', help="show some debugging info")
	parser.add_argument('-ps', '--parse_section', dest='parse_section', action='store_true', help="trying to parse section, subsection, subsubsection from slide title")
	parser.add_argument('-pi', '--parse_items', dest='parse_item', action='store_true', help="trying to parse nested items")
	parser.add_argument('-ur', '--use_rput', dest='use_rput', action='store_true', help="image positioning will be done using rput")
	parser.add_argument('-t', metavar='templateFileName', dest="templateFileName", type=str,  help='the generated latex code will be placed at the bottom of the file')

	#TODO parameter ideas
	#DONE -tex template e.g with placeholder, where the content should be placed
	#-user defined image path
	#-user defined image export format
	# ...
	args = parser.parse_args()
	inputFilename = args.input
	outputFilename = args.output
	startPageNo = args.start
	endPageNo = args.end
	debug = args.debug
	verbose = args.verbose
	use_rput = args.use_rput
	parse_section = args.parse_section
	parse_item = args.parse_item
	templateFileName = args.templateFileName
	print((templateFileName))
	if args.version:
		print( "ImpressToLatex Version: "+impressToLatexVersion)
		sys.exit(3)

inputFileUrl = unohelper.systemPathToFileUrl(inputFilename)

#establish conection
local = uno.getComponentContext();
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
#document = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ()) # new document
document = desktop.loadComponentFromURL(inputFileUrl ,"_blank", 0, ()) # existing document

pageCnt = document.DrawPages.Count
print( pageCnt)

texHead = '''\documentclass[10pt,a4paper]{beamer}
TODO: fill latex header code
\\begin{document}'''
# \maketitle\n'''

#texFile = open(os.path.splitext(inputFilename)[0]+"tex", "w")
#if len(sys.argv) >= 2:
	#outputFilename = sys.argv[2]
#else:
	#outputFilename = os.path.basename(os.path.splitext(inputFilename)[0] + '.tex')
#outputFilename = "output.tex"
#os.path.isfile(fname)
#val = open("char_table.txt","w")
#val.write("[")
#val.write(",\n".join(str(elem) for elem in replaceCharTable))
#val.write("]")
charTable = open("char_table.txt","r")
replaceCharTable = eval(charTable.read())

templateFile = open(templateFileName,"r")
texHead = templateFile.read()

print( texHead)

if os.path.exists(outputFilename):
	os.remove(outputFilename)

texFile = open(outputFilename, "w")
texFile.write(texHead)
texFile.write("");

pageBody = ""


firstTitle = True
frameOpened = False
section = ""
subSection = ""
subSubSection = ""
#if len(sys.argv) >= 3:
	#pageCnt = int(sys.argv[3])
	#print( pageCnt)

def closeFrame(force = False):
	global frameOpened
	global pageBody
	print( "!!!!!!!!!!!!!!!!!!!!!!!!!!!!NEWPAGE")
	print( firstTitle, frameOpened, force)
	if (not firstTitle and frameOpened) or force:
		texFile.write(pageBody)
		texFile.write("\\end{frame} %close frame \n\n\n") 
		pageBody = ""
		frameOpened = False

def writeLinesAsItemize(lines, prefix = ""):
	global pageBody
	#old implementation (each line results in an itemize entry)				
	pageBody += prefix+"\t\\begin{itemize} \n"
	for line in lines:
		if line.strip() == "": # empty items
			continue
		line = processText(unicode(line).encode("utf-8")) #line.replace("\\", "\\textbackslash ") # replace backslash
		#print( line)
		pageBody += prefix+"\t\t\item "+str(line)+"\n"
	pageBody += prefix+"\t\\end{itemize} \n"




def addImageToPageBody(imgFileName, posx = 0, posy = 0, comment = ""):
	global pageBody
	global use_rput
	line = "\t"
	if use_rput:
		line += "\\rput("+str(posx)+", "+str(posy)+"){"
	line += "\includegraphics[width=.4\linewidth]{"+imgFileName+"}"
	if use_rput:
		line+="}"
	if comment:
		line += "%"+comment
	pageBody += line+"\n"


#sys.exit(0)
#try:
if endPageNo == -1:
	endPageNo = pageCnt

for i in range(startPageNo, endPageNo): # iterate over pages range(pageCnt)
	closeFrame()
	page = document.DrawPages.getByIndex(i)
	if verbose:
		print( "\n[INFO] Processing Page %(page)04d/%(pages)04d\n" %{"page": i, "pages": pageCnt - 1})
	for j in range(page.Count): #iterate over page elements
		#print( page.Count, j)
		element = page.getByIndex(j)
		#print( type(element))
		if debug:
			print( element.Name)
			print( element)

		#print( element.Text.getString())
		#print( "\t"+str(element))

		#TODO: Calc a better, more correct x/y position for images
		x = element.getPosition().X/10000.0
		y = element.getPosition().Y/10000.0

		#
		if  'com.sun.star.presentation.TitleTextShape'  in element.SupportedServiceNames: #title
			#print( "[Title]")
			#closeFrame();
			s = element.Text.getString()
			print( s)
			s = processText(unicode(s).encode("utf-8"))

			if parse_section: #TODO: ugly
				lines = s.split("\n")
				if len(lines) > 0 and len(lines[0].split()) > 1 and isFloat(lines[0].split()[0]):
					newSection = lines[0].split()[1].strip()
					#print( "xxx"+newSection)
					if newSection != section:
						texFile.write("\\section{"+newSection+"}\n")
						section = newSection
						#print( newSection+"\n\n")
						#s = lines[1:]
					#section = lines[0];
					if len(lines) > 1:
						newSubSection = lines[1].strip()
						#print( lines[1])
						if newSubSection != subSection:
							texFile.write("\\subsection{"+newSubSection+"}\n")
							subSection = newSubSection
							#print( newSubSection;)
						if len(lines) > 2:
							newSubSubSection = lines[2].strip()
							if newSubSubSection != subSubSection:
								texFile.write("\\subsubsection{"+newSubSubSection+"}\n")
								subSubSection = newSubSubSection
								#print( newSubSubSection;)

			texFile.write("\\begin{frame}\n")
			if parse_section:
				title = subSubSection
				if not title:
					title = subSection
				if not title:
					title = section
				if title:
					print( "Title", title)
					texFile.write("\t\\frametitle{"+subSubSection+"} \n")
			else:
				texFile.write("\t\\frametitle{"+s+"} \n")
			firstTitle = False
			frameOpened = True
			#texFile.write(pageBody)
			#print( pageBody)
			#pageBody = ""

		elif 'com.sun.star.drawing.TextShape' in element.SupportedServiceNames:#text as itemize
			if verbose:
				print( "[TEXTSHAPE]")
			lines = element.Text.getString().split("\n")
			#print( lines)
			#texFile.write(element.Text.getString()+"\n")
			#print( lines )
			if debug:
				print( element)

			if lines[0] == "<number>": # page numbers?
				continue

			if parse_item:
				a = element.createEnumeration();
				lastNumberingLevel = -1
				#el = a.nextElement()
				while a.hasMoreElements():
					#print( el)
					el = a.nextElement()
					text = processText(unicode(el.getString()).encode("utf-8")).strip()
					if text :
						print( text)
					#if el.ImplementationName == "SvxUnoTextContent":
						#texFile.write(text)
						#if debug:
							#print( text)
						#continue
					tabIndent = "\t"*el.NumberingLevel
					print( "numberinglevel", el.NumberingLevel, lastNumberingLevel)
					if el.NumberingLevel > lastNumberingLevel and el.NumberingLevel >= 1:
						pageBody += tabIndent+"\\begin{itemize} \n"
					if el.NumberingLevel < lastNumberingLevel:
						tabIndent+"\t\\end{itemize} \n" #close all opened begin blocks
					if text:
						if el.NumberingLevel >= 1: #enumerations
							pageBody += tabIndent+"\t\item "+text+"\n"
						else:
							pageBody += text+"\n" #standard text
					print( el.NumberingLevel, text)
					#el = a.nextElement()
					print( "aaaaaaaaaaaa", (not a.hasMoreElements()) or el.NumberingLevel < lastNumberingLevel, el.NumberingLevel < lastNumberingLevel, not a.hasMoreElements(), el.NumberingLevel, lastNumberingLevel)
					#print( a)
					lastNumberingLevel = el.NumberingLevel

					if (not a.hasMoreElements()) or el.NumberingLevel < lastNumberingLevel:
						if not a.hasMoreElements():
							levelDifference = lastNumberingLevel
						else:
							levelDifference = lastNumberingLevel - el.NumberingLevel
						print( el.NumberingLevel, lastNumberingLevel, levelDifference)
						for cnt in range(levelDifference):
							if not a.hasMoreElements():
								tabCnt = levelDifference - cnt 
							else:
								tabCnt = lastNumberingLevel
							pageBody += "\t"*(tabCnt)+"\\end{itemize} \n" #close all opened begin blocks
			else:
			#old implementation (each line results in an itemize entry)				
				writeLinesAsItemize(lines)


		elif 'com.sun.star.drawing.GraphicObjectShape' in element.SupportedServiceNames: #graphics #TODO: may cut
			graphicsFileName = element.GraphicStreamURL
			#graphicsFileName = "images/%(page)04d_%(element)03d_image.png" %{"page": i, "element": j}
			if graphicsFileName:
				releps_filename = "images/"+os.path.basename(os.path.splitext(graphicsFileName)[0] + '.eps')
				if verbose:
					print( "[GRAPHIC] "+ releps_filename)

				eps_filename = os.path.expanduser(releps_filename)
				eps_filename = os.path.abspath(eps_filename)
				eps_url = unohelper.systemPathToFileUrl(eps_filename)
				writeEPS(desktop, document, eps_url, context, element)
				addImageToPageBody(os.path.splitext(os.path.basename(releps_filename))[0], x, y, "IMAGE")
				#pageBody += "\t\\rput("+str(x)+", "+str(y)+"){\includegraphics[width=.4\linewidth]{"+os.path.splitext(os.path.basename(releps_filename))[0]+"}} %IMAGE \n"

		elif 'com.sun.star.drawing.OLE2Shape' in element.SupportedServiceNames: #hopefully stuff like visio drawings, there is an error in exporting these kind of data with libreoffice (blank file)
			relemf_filename = "images/%(page)04d_%(element)03d_diagram.eps" %{"page": i, "element": j}
			emfFilename = os.path.expanduser(relemf_filename)
			emfFilename = os.path.abspath(emfFilename)
			emfUrl = unohelper.systemPathToFileUrl(emfFilename)
			if verbose:
				print( "[EPS] "+relemf_filename)
			if debug:
				width =  int(element.getSize().Width)
				height = int(element.getSize().Height)
				print( width, height)

			if relemf_filename not in ["images/0048_004_diagram.eps", "images/0058_003_diagram.eps"] :
				writeEPS(desktop, document, emfUrl, context, element)
			#print( textLine)
			addImageToPageBody(os.path.splitext(os.path.basename(relemf_filename))[0], x, y, "VECTOR GRAPHIC")	
			#pageBody += "\t\\rput("+str(x)+", "+str(y)+"){\includegraphics[width=.4\linewidth]{"+os.path.splitext(os.path.basename(relemf_filename))[0]+"}} %VECTOR GRAPHIC \n"


		elif 'com.sun.star.drawing.Shape' in element.SupportedServiceNames: #485 479
			# filter 
			if element.getSize().Width == 3197 and element.getSize().Height == 3201:
				if verbose:
					print( "[FILTER CIRCLE]")
				continue # filter ugly circles
			relpng_filename = "images/%(page)04d_%(element)03d_shape.eps" %{"page": i, "element": j}
			png_filename = os.path.expanduser(relpng_filename)
			png_filename = os.path.abspath(png_filename)
			png_url = unohelper.systemPathToFileUrl(png_filename)
			if verbose:
				print( "[SHAPE] "+relpng_filename, png_filename, element.getSize().Width, element.getSize().Height)
			comment = ""
			if "com.sun.star.drawing.Text" in element.SupportedServiceNames:
				text = element.Text.getString()
				if debug:
					print( text)
				if text:
					lines = text.split("\n")
					writeLinesAsItemize(lines, "%")
					#comment = "%"
			writeEPS(desktop, document, png_url, context, element);
			addImageToPageBody(os.path.splitext(os.path.basename(relpng_filename))[0], x, y, "SHAPE")
			#pageBody += comment+"\t\\rput("+str(x)+", "+str(y)+"){\includegraphics[width=.4\linewidth]{"+os.path.splitext(os.path.basename(relpng_filename))[0]+"}} %SHAPE \n"
		#print( pageBody)


closeFrame(True)
#except Exception, e:
	#raise e
#finally:
document.dispose()
texFile.write("\end{document}")
texFile.close()

if verbose:
	print( "\n[INFO] Document converted")

