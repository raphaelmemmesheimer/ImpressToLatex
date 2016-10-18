ImpressToLatex
==============

A small script for exporting Latex from OO/Libre- Impress presentations.
(This can also be used to convert PPT to LaTeX, when you first store your PPT in ODP format.)

Its intention is not to create a 1:1 representation of the input file, just converting should be made a bit more comfortable. Actually it's the result of a quick hack onto a griven problem thus the current state is not meant as an allround solution  

Introductions
=============

Dependecies: pyUNO 

on debian like machines, the next line should do the job:

`$ apt-get install python-uno`


## 1) Start Open/Libre Office like:

`$ <executable> --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"`

for instance `<executable> = libreoffice`

## 2) Run the Script

`$ ImpressToLatex.py <InputFileName> <OutPutFileName>`

Now the script iterates over all elements on all pages and handels them speratly.

* Title elements will be exported as beamer frames
* Text elements will be exported as an itemize block
* Shapes and Images will be exported as *.png image files ans used as an includegraphics tag 
* OLE2 Shapes such as Visio Drawings will be exported as *.eps files

All images will be stored in an "images" subfolder

There is also a processText function, which handles a bunch of unicode or reserved latex charachters (these are at least the chars I needed that far). Just extend the replaceCharTable array at the top of the script in case you need some more characters. 

## 3) Create PDF

I used the following command:

`$ pdflatex -shell-escape <texfile>`

the "-shell-escape" parameter is needed in case the script includes some *.eps files

Alternatively use latex and dvips.

If you start the program using the -p option,
the program will create PNG files for each graphics object which can be
included directly into the PDF file when you run pdflatex

$ ImpressToLatex.py -p <InputFileName> <OutPutFileName>`

## 4) Tested 

Tested with Libreoffice
