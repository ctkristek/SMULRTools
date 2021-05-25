from pprint import pprint as pp
from flask import Flask, flash, redirect, render_template, request, url_for, Response, make_response, send_from_directory, send_file
from werkzeug import secure_filename
import sys
from zipfile import ZipFile
import re
import xlwt
import xlsxwriter
from bs4 import BeautifulSoup
from urlextract import URLExtract
from pypermacc import Permacc
from io import BytesIO
import os
#import ocrmypdf
#from fpdf import FPDF

import requests
import json

UPLOAD_FOLDER = '/tmp'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def listToString(s):
    str1 = ""
    for ele in s:
        str1 += ele
    return str1

def docxToBTLPerma(f, p):
    #Read the uploaded .docx file
    with ZipFile(BytesIO(f.read())) as my_zip_file:
        for contained_file in my_zip_file.namelist():
            if contained_file == "word/footnotes.xml":
                for line in my_zip_file.open(contained_file).readlines():
                    soup = BeautifulSoup(line, 'xml')
            else:
                print('Please Upload a .docx file')

    extractor = URLExtract()
    Permafolder = p


    #Create in-memory buffer to write the excel file to
    output=BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 50)

    # Set up some formats to use.
    bold = workbook.add_format({'bold': True})
    italic = workbook.add_format({'italic': True})
    text_wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})


    footnotes = soup.findAll('w:footnote')

    #Go through each footnote
    for ft in footnotes:
        #Empty variabel to construct footnote parts
        #footnoteObj = []

        #Get the footnote number. Have to subtract 1 for some unkown reason
        ftId = ft.get('w:id')
        cellId = 'B' + ftId
        urlID = 'C' + ftId
        ftId = int(ftId)-1


        if int(ftId) < 0:
            continue

        #Gets to the actual text of the footnote in the XML file
        ftTags = ft.findAll('w:r')
        ftStyleText = []
        permaLinksXLSX = []
        extractor = URLExtract()
        styleCount = 0

        for ftag in ftTags:

            ftTagText = ftag.findAll('w:t')

            footnoteString = ""
            for f in ftTagText:
                footnoteString+=f.text

            ftURLs = extractor.find_urls(footnoteString)

            if footnoteString == "":
                continue

            for url in ftURLs:
                if url in footnoteString:
                    u = url.rstrip('.')
                    headers = {
                        'Content-Type': 'application/json; charset=utf-8'
                    }

                    params = (
                        ('api_key', '9b721250766b6f80ce7df45c8fa0a3c63c231079'),
                    )

                    data = '{"url":"' +str(u)+ '", "folder":' +str(Permafolder)+ '}'

                    response = requests.post('https://api.perma.cc/v1/archives/', headers=headers, params=params, data=data.encode('utf-8'))
                    if response.status_code == 201:
                        permaResponse = response.json()
                        Permaguid = permaResponse.get("guid")
                        permaLink = "perma.cc/" + str(Permaguid)
                        permaLinksXLSX.append(permaLink)
                    else:
                        permaLinksXLSX.append(str(response.content))

            if ftag.find ('w:i'):
                ftStyleText.append(italic)
                ftStyleText.append(footnoteString)
                styleCount = 1
            elif ftag.find ('w:smallCaps'):
                ftStyleText.append(bold)
                ftStyleText.append(footnoteString)
                styleCount = 1
            else:
                ftStyleText.append(footnoteString)

        #print(ftStyleText)




        #writePermaText = listToString(ftPermaLinks)
        linkText=listToString(permaLinksXLSX)

        if styleCount == 0:
            writeText=listToString(ftStyleText)
            worksheet.write(cellId, writeText, text_wrap)
            worksheet.write(urlID, linkText)
            #worksheet.write(pcellId, writePermaText, text_wrap)
        else:
            ftStyleText.append(text_wrap)
            worksheet.write_rich_string(cellId, *ftStyleText)
            worksheet.write(urlID, linkText)
            #worksheet.write(pcellId, writePermaText, text_wrap)



    workbook.close()
    output.seek(0)
    #print(output)
    response = make_response(output.read())
    response.headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.headers.set('Content-Disposition', 'attachment', filename='BTLedits.xlsx')
    return response


def docxToBTL(f):
    #Read the uploaded .docx file
    with ZipFile(BytesIO(f.read())) as my_zip_file:
        for contained_file in my_zip_file.namelist():
            if contained_file == "word/footnotes.xml":
                for line in my_zip_file.open(contained_file).readlines():
                    soup = BeautifulSoup(line, 'xml')
            else:
                print('Please Upload a .docx file')

    extractor = URLExtract()

    #Create in-memory buffer to write the excel file to
    output=BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 50)

    # Set up some formats to use.
    bold = workbook.add_format({'bold': True})
    italic = workbook.add_format({'italic': True})
    text_wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})


    footnotes = soup.findAll('w:footnote')

    #Go through each footnote
    for ft in footnotes:
        #Empty variabel to construct footnote parts
        #footnoteObj = []

        #Get the footnote number. Have to subtract 1 for some unkown reason
        ftId = ft.get('w:id')
        cellId = 'B' + ftId
        urlID = 'C' + ftId
        ftId = int(ftId)-1


        if int(ftId) < 0:
            continue

        #Gets to the actual text of the footnote in the XML file
        ftTags = ft.findAll('w:r')
        ftStyleText = []
        extractor = URLExtract()
        styleCount = 0

        for ftag in ftTags:

            ftTagText = ftag.findAll('w:t')

            footnoteString = ""
            for f in ftTagText:
                footnoteString+=f.text

            ftURLs = extractor.find_urls(footnoteString)

            if footnoteString == "":
                continue

            if ftag.find ('w:i'):
                ftStyleText.append(italic)
                ftStyleText.append(footnoteString)
                styleCount = 1
            elif ftag.find ('w:smallCaps'):
                ftStyleText.append(bold)
                ftStyleText.append(footnoteString)
                styleCount = 1
            else:
                ftStyleText.append(footnoteString)


        if styleCount == 0:
            writeText=listToString(ftStyleText)
            worksheet.write(cellId, writeText, text_wrap)
        else:
            ftStyleText.append(text_wrap)
            worksheet.write_rich_string(cellId, *ftStyleText)

    workbook.close()
    output.seek(0)
    #print(output)
    response = make_response(output.read())
    response.headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.headers.set('Content-Disposition', 'attachment', filename='BTLedits.xlsx')
    return response

def OCRaPDF(pdfFile):

    OCRoutput=FPDF()

    ocrmypdf.ocr(input_file=pdfFile, output_file=OCRoutput, deskew=True, force_ocr=True)
    response = ocrmypdf.output(dest='S').encode('latin-1')
    response.headers.set('Content-Disposition', 'attachment', filename=name + '.pdf')
    response.headers.set('Content-Type', 'application/pdf')
    return response




@app.route('/')
def index():
    return render_template('index.html')

@app.route('/uploaddocx', methods=['POST','GET'])
def upload_pdf():
    if request.method == 'POST':
        f = request.files.get('file')
        p = request.form['perma']
        if p == "":
            response = docxToBTL(f)
        else:
            response = docxToBTLPerma(f, p)

        f.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename)))
        return response

@app.route('/OCRaPDF', methods=['POST','GET'])
def upload_file():
    if request.method == 'POST':
        f = request.files.get('ocr_file')
        OCRresponse = OCRaPDF(f)
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename)))
        return OCRresponse
