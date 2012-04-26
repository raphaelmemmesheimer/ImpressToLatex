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
from com.sun.star.beans import PropertyValue
from com.sun.star.beans.PropertyState import DIRECT_VALUE


#found @ http://user.services.openoffice.org/en/forum/viewtopic.php?f=45&t=44220
def write_png(desktop, doc, url, ctx, element):
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
	#args.append(PropertyValue("FilterData", 0 ,tuple(filterData), DIRECT_VALUE))
	
	graphic_filter = smgr.createInstanceWithContext(
	    "com.sun.star.drawing.GraphicExportFilter", ctx)
	graphic_filter.setSourceDocument(collection)
	graphic_filter.filter(tuple(args))


#character replacement table, order dependend
replaceCharTable = [["\\","\\textbackslash "],
					["%","\\%"],
					["$", " !SPECIALCHAR! "],
					["≠"," $ \\neq $ "],
					["„", "\\quotedblbase "],
					["“","\\textquotedbl "],
					["Ü",'\\"U'],
					["ü",'\\"u'],
					["Ö",'\\"O'],
					["ö",'\\"o'],
					["Ä",'\\"A'],
					["ä",'\\"a'],
					["€","EUR"],
					["ß","\\ss "],
					[" ", " $ \cong $ "],
					["", " $ \lambda $ "],
					["µ"," $\mu $ "],
					["μ"," $\mu$ "],
					["&", "\&"],
					[" "," "], #different spaces
					#[" "," !SPECIAL CHAR! "],
					["m","a"], #todo
					#[")","!SPECIAL CHAR! "],
					[""," $\\rightarrow$ "],
					["°", "\degree"], 
					["↑"," $\\uparrow$ "], 
					["↓","$ \\downarrow$ "],
					["↔"," $\leftrightarrow$ "],
					[""," $\Omega$ "],
					[""," $A$ "],
					["<<"," $<<$ "],
					[""," "], #yet another space
					["^2", " $^2$ "],
					["²", " $^2$ "],
					[""," $\\alpha$ "],
					["", " $\\theta$ "],
					["", " !!!SPECIALCHAR!!! !(underscore 2)! "],
					["", " $\\rho$ "],
					[""," $\\tau$ "],
					["", " $\\beta$ "],
					["", " !!!SPECIALCHAR!!! !(underscore c)! "],
					#["", " !SPECIALCHAR! "],
					["", " !SEPCIALCHAR! "],
					["≈", "!SPECIALCHAR! "],
					[""," !SPECIALCHAR! "],
					[""," !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["θ", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["≥", " !SPECIALCHAR! "],
					["_", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["½", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " !SPECIALCHAR! "],
					["", " "],
					["", " "],
					[""," "],
					["", " "]
					] 

def processText(s): # replace special characters
	for r in replaceCharTable:
		s = s.replace(r[0], r[1])
	return str(s) 


#info
if len(sys.argv) <= 1:
	print "Usage:"
	print "\tpython toLatex.py <ImpressFile> <OutputFile>"
	sys.exit(0)	
else:
	inputFilename = sys.argv[1]
	inputFileUrl = unohelper.systemPathToFileUrl(inputFilename)


#establish conection
local = uno.getComponentContext();
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

#document = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ()) # new document
document = desktop.loadComponentFromURL(inputFileUrl ,"_blank", 0, ()) # existing document

pageCnt = document.DrawPages.Count;

texHead = '''\documentclass[10pt,a4paper]{beamer}
\usepackage[utf8]{inputenc}
%\usepackage{eurosym}
\usepackage{amsmath}
\usepackage{amsfonts}
\usepackage{amssymb}
\usetheme{Berlin}
\usepackage[T1]{fontenc}
\usepackage{epstopdf}
\usepackage{pstricks}
\usepackage{gensymb}

\graphicspath{{./images/}}

\\author{Name}
\\begin{document}
\maketitle\n'''

if len(sys.argv) >= 2:
	outputFilename = sys.argv[2]
else:
	outputFilename = os.path.basename(os.path.splitext(inputFilename)[0] + '.tex')
if os.path.exists(outputFilename):
	os.remove(outputFilename)

texFile = open(outputFilename, "w")
texFile.write(texHead)
texFile.write("");



firstTitle = True
frameOpened = False
try:
	for i in range(1, pageCnt): # iterate over pages range(pageCnt)
		page = document.DrawPages.getByIndex(i)
		for j in range(page.Count): #iterate over page elements
			element = page.getByIndex(j)
			if 'com.sun.star.presentation.TitleTextShape' in element.SupportedServiceNames: #title
				s = element.Text.getString()
				s = processText(unicode(s).encode("utf-8"))
				print "\n"+ s +"\n"
				texFile.write("\\section{"+s+"}\n")
				texFile.write("\\frame{\\frametitle{"+s+"} \n")
				firstTitle = False
				frameOpened = True
	
			elif 'com.sun.star.drawing.TextShape' in element.SupportedServiceNames:#text as itemize
				lines = element.Text.getString().split("\n")
				texFile.write("\\begin{itemize} \n")
				for line in lines:
					if line.strip() == "": # empty items
						continue
					line = processText(unicode(line).encode("utf-8"))
					texFile.write("\t \item "+str(line)+"\n")
				texFile.write("\\end{itemize} \n")
	
			elif 'com.sun.star.drawing.GraphicObjectShape' in element.SupportedServiceNames: #graphics
				graphicsFileName = element.GraphicStreamURL
				if graphicsFileName:
					relpng_filename = "images/"+os.path.basename(os.path.splitext(graphicsFileName)[0] + '.png')
					
					print "[GRAPHIC] "+ relpng_filename
					png_filename = os.path.expanduser(relpng_filename)
					png_filename = os.path.abspath(png_filename)
					png_url = unohelper.systemPathToFileUrl(png_filename)
					write_png(desktop, document, png_url, context, element)
					x = element.getPosition().X/10000.0
					y = element.getPosition().Y/10000.0
					texFile.write("\\rput("+str(x)+", "+str(y)+"){\includegraphics[scale=0.2]{"+os.path.basename(relpng_filename)+"}} %IMAGE \n")


			elif 'com.sun.star.drawing.OLE2Shape' in element.SupportedServiceNames: #hopefully stuff like visio drawings, there is an error in exporting these kind of data with libreoffice (blank file)
				relemf_filename = "images/%(page)04d_%(element)03d_diagram.eps" %{"page": i, "element": j}
				emfFilename = os.path.expanduser(relemf_filename)
				emfFilename = os.path.abspath(emfFilename)
				emfUrl = unohelper.systemPathToFileUrl(emfFilename)
				print "[EPS FILE] "+relemf_filename
				width =  int(element.getSize().Width)
				height = int(element.getSize().Height)
				#print width, height
				writeEPS(desktop, document, emfUrl, context, element)
				texFile.write("\includegraphics[width=10mm]{"+os.path.basename(relemf_filename)+"} % \n")
	
	
			elif 'com.sun.star.drawing.Shape' in element.SupportedServiceNames: #
				relpng_filename = "images/%(page)04d_%(element)03d_shape.png" %{"page": i, "element": j}
				png_filename = os.path.expanduser(relpng_filename)
				png_filename = os.path.abspath(png_filename)
				png_url = unohelper.systemPathToFileUrl(png_filename)
				print "[SHAPE] "+relpng_filename, png_filename, element.getSize().Width, element.getSize().Height
				write_png(desktop, document, png_url, context, element);
				texFile.write("\includegraphics[width=10mm]{"+os.path.basename(relpng_filename)+"} %SHAPE \n")

	
		if not firstTitle and frameOpened:
			texFile.write("} %close frame \n\n\n") 
			frameOpened = False

except Exception, e:
	raise e
finally:
	document.dispose()
	texFile.write("\end{document}")
	texFile.close()

