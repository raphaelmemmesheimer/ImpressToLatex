# coding: utf-8

# start libreoffice like : libreoffice -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"

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
	#print "[GRAPHIC SUPPORTED EXPORT FILES]"+str(graphic_filter.getSupportedMimeTypeNames())
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
	#print "[asdasdasd]"+str(graphic_filter.getSupportedMimeTypeNames())
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
					["°", "\degree"], #todo
					["↑"," $\\uparrow$ "], #todo
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
					] #todo       ½      

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

\def\\twolang#1#2{#1}
\let\\2=\\twolang


\graphicspath{{./images/}}

\\author{Dr.-Ing. Yannick Caulier }
\\begin{document}
\maketitle\n'''

#texFile = open(os.path.splitext(inputFilename)[0]+"tex", "w")
if len(sys.argv) >= 2:
	outputFilename = sys.argv[2]
else:
	outputFilename = os.path.basename(os.path.splitext(inputFilename)[0] + '.tex')
#outputFilename = "output.tex"
#os.path.isfile(fname)
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
			#print type(element)
			#print element.Name
			#print element.Text.getString()
			#print "\t"+str(element)
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
				print lines
				#texFile.write(element.Text.getString()+"\n")
				#print lines 
				texFile.write("\\begin{itemize} \n")
				for line in lines:
					if line.strip() == "": # empty items
						continue
					line = processText(unicode(line).encode("utf-8")) #line.replace("\\", "\\textbackslash ") # replace backslash
					texFile.write("\t \item "+str(line)+"\n")
				texFile.write("\\end{itemize} \n")
	
			elif 'com.sun.star.drawing.GraphicObjectShape' in element.SupportedServiceNames: #graphics
				graphicsFileName = element.GraphicStreamURL
				#graphicsFileName = "images/%(page)04d_%(element)03d_image.png" %{"page": i, "element": j}
				if graphicsFileName:
					relpng_filename = "images/"+os.path.basename(os.path.splitext(graphicsFileName)[0] + '.png')
					
					print "[GRAPHIC] "+ relpng_filename
					
					#print png_filename
					png_filename = os.path.expanduser(relpng_filename)
					png_filename = os.path.abspath(png_filename)
					png_url = unohelper.systemPathToFileUrl(png_filename)
					write_png(desktop, document, png_url, context, element)
					x = element.getPosition().X/10000.0
					y = element.getPosition().Y/10000.0
					viewport = "viewport= %(x)d %(y)d %(width)d %(height)d" %{"x": x, "y": y,"width": x+(element.getSize().Width/1000),"height": y+(element.getSize().Height/1000)} 
					print "[Viewport] "+ viewport+"\t" +str(element.getSize().Width)+"\t"+ str(element.getSize().Height)
					texFile.write("\\rput("+str(x)+", "+str(y)+"){\includegraphics[scale=0.2]{"+os.path.basename(relpng_filename)+"}} %IMAGE \n")
					#texFile.write("\includegraphics[width=10mm]{"+relpng_filename+"} %IMAGE\n")
					#texFile.write("\includegraphics["+viewport+"]{"+relpng_filename+"} %IMAGE\n")
	
			elif 'com.sun.star.drawing.OLE2Shape' in element.SupportedServiceNames: #hopefully stuff like visio drawings, there is an error in exporting these kind of data with libreoffice (blank file)
				print ""
				relemf_filename = "images/%(page)04d_%(element)03d_diagram.eps" %{"page": i, "element": j}
				emfFilename = os.path.expanduser(relemf_filename)
				emfFilename = os.path.abspath(emfFilename)
				emfUrl = unohelper.systemPathToFileUrl(emfFilename)
				print "[EPS FILE] "+relemf_filename
				width =  int(element.getSize().Width)
				height = int(element.getSize().Height)
				print width, height
				if relemf_filename not in ["images/0048_004_diagram.eps", "images/0058_003_diagram.eps"] :
					writeEPS(desktop, document, emfUrl, context, element)
				#textLine = "\includegraphics[width="+str(width)+"mm, height="+str(height)+"mm]{"+relemf_filename+"} %DIAGRAM \n" # %{"width": width,"height": height}
				#print textLine
				#texFile.write(textLine)includegraphics
				texFile.write("\includegraphics[width=10mm]{"+os.path.basename(relemf_filename)+"} % \n")
	
	
			elif 'com.sun.star.drawing.Shape' in element.SupportedServiceNames: #
				if element.getSize().Width == 1998 and element.getSize().Height == 2002:
					print "[FILTER CIRCLE]"
					continue # filter ugly circles
				relpng_filename = "images/%(page)04d_%(element)03d_shape.png" %{"page": i, "element": j}
				png_filename = os.path.expanduser(relpng_filename)
				png_filename = os.path.abspath(png_filename)
				png_url = unohelper.systemPathToFileUrl(png_filename)
				print "[SHAPE] "+relpng_filename, png_filename, element.getSize().Width, element.getSize().Height
				write_png(desktop, document, png_url, context, element);

				#viewport = "viewport %(x) %(y) %(width) %(height)" %{"x": element.getPosition().X, "y": element.getPosition().Y,"width": element.getSize().Width,"height": element.getSize().Height} 
				#print "[Viewport] "+ viewport+"\t" +str(element.getSize().Width)+"\t"+ str(element.getSize().Height)
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

