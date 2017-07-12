# -*- coding: utf-8 -*-
"""
Created on Mon Jul  3 12:04:33 2017

@author: cartemic
"""

import requests
from lxml import html
import xlsxwriter

url_list = ['http://shepherd.caltech.edu/detn_db/html/db_121.html',
            'http://shepherd.caltech.edu/detn_db/html/db_122.html',
            'http://shepherd.caltech.edu/detn_db/html/db_123.html',
            'http://shepherd.caltech.edu/detn_db/html/db_124.html',
            'http://shepherd.caltech.edu/detn_db/html/db_125.html',
            'http://shepherd.caltech.edu/detn_db/html/db_126.html',
            'http://shepherd.caltech.edu/detn_db/html/db_127.html',
            'http://shepherd.caltech.edu/detn_db/html/db_128.html']

baseUrl = 'http://shepherd.caltech.edu/detn_db/'
refUrl = 'http://shepherd.caltech.edu/detn_db/html/references.html'


class stringFind():
    def __new__(self, theString, startKey, stopKey, startMod=0, stopMod=0):
        dataLoc = [0, 0]
        dataLoc[0] = theString.find(startKey)+startMod
        dataLoc[1] = theString[dataLoc[0]+1:].find(stopKey)+dataLoc[0]+stopMod
        return(theString[dataLoc[0]:dataLoc[1]])

for url in url_list:
    # pull page HTML
    page = requests.get(url)
    refpage = requests.get(refUrl)

    pagetext = page.text
    reftext = refpage.text
    key = ['<BLOCKQUOTE>', '</BLOCKQUOTE>']
    offset = len(key[0])
    num_elements = pagetext.count(key[0])

    theTitle = pagetext[pagetext.find('<TITLE>')+7:pagetext.find('</TITLE>')]
    print(theTitle)
    fuelType = stringFind(pagetext[pagetext.find('H3'):pagetext.find('/H3')],
                          '-', 'Fuel', 2)
    workbook = xlsxwriter.Workbook(theTitle+'.xlsx')

    loc = [0, 0]
    for i in range(num_elements):
        loc[0] = loc[1]+1
        loc[0] = pagetext[loc[0]:].find(key[0]) + loc[0]
        loc[1] = pagetext[loc[0]:].find(key[1]) + loc[0]
        inspectStr = pagetext[loc[0]+offset:loc[1]]

        if i % 2 == 0:
            # find URL of text file containing data points
            dataUrl = stringFind(inspectStr, '"', '"', 4, 1)

            # find the name of the dataset
            dataName = stringFind(inspectStr, '.txt', '</A>', 6, -3)

            # find the autor
            dataAuthor = stringFind(inspectStr, '&#160;', '&#160;', 6, 1)

            # find the reference number and corresponding citation
            refNumber = stringFind(inspectStr, '[', ']', 1, 1)
            if refNumber != '130':
                dataRef = stringFind(reftext, '['+refNumber+']', '<P>',
                                     12+len(refNumber))
            else:
                dataRef = stringFind(reftext, '['+refNumber+']',
                                     '</DL>', 12+len(refNumber))
            dataRef = html.fromstring(dataRef).text_content()
            dataRef = dataRef.replace('\n', ' ').replace('   ', ' ')

        else:
            # find category
            dataCategory = stringFind(inspectStr, 'LEFT', '</TD>', 7, 1)

            # find fuel
            dataFuel = stringFind(inspectStr, 'Fuel', '</TD>', 35+6, 1)

            # find sub-category
            dataSubCategory = stringFind(inspectStr, 'Sub-', '</TD>', 35+14, 1)

            # find oxidizer
            dataOxidizer = stringFind(inspectStr, 'Oxidizer:', '</TD>', 35+10)

            # find initial pressure
            dataPressure = stringFind(inspectStr, 'Pressure', '</TD>', 35+10,
                                      1)

            # find diluent
            dataDiluent = stringFind(inspectStr, 'Diluent', '</TD>', 35+9, 1)

            # find initial temperature
            dataTemperature = stringFind(inspectStr, 'Temp', '</TD>', 35+13, 1)

            # find equivalence ratio
            dataEquivalence = stringFind(inspectStr, 'Equiv', '</TD', 35+19, 1)

            if dataOxidizer == 'Air':
                # collect the actual data
                theData = requests.get(baseUrl+dataUrl)
                dataOut = theData.text[1:].split('\n')
                dataOut = list(filter(None, dataOut))
                worksheet = workbook.add_worksheet(dataName)
                worksheet.write(0, 0, 'Dataset '+dataName)
                worksheet.write(1, 0, 'Author '+dataAuthor)
                worksheet.write(2, 0, 'Source '+dataRef)
                worksheet.write(4, 0, 'Category '+dataCategory)
                worksheet.write(5, 0, 'Sub-Category '+dataSubCategory)
                worksheet.write(7, 0, 'Oxidizer '+dataOxidizer)
                worksheet.write(8, 0, 'Diluent '+dataDiluent)
                worksheet.write(10, 0, 'Temperature '+dataTemperature)
                worksheet.write(11, 0, 'Pressure '+dataPressure)
                worksheet.write(12, 0, 'Equivalence Ratio '+dataEquivalence)
                headerRows = 14
                for j in range(len(dataOut)):
                    if j != 0:
                        dataOut[j] = dataOut[j].replace(' ', '')
                    else:
                        dataOut[j] = dataOut[j].replace(', ', ',')
                    dataOut[j] = dataOut[j].split(',')
                    for k in range(len(dataOut[0])):
                        worksheet.write(j+headerRows, k, dataOut[j][k])
    workbook.close()
