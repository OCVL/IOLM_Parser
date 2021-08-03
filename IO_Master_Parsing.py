# Libraries needed
import os
from os import path
import PyPDF2
import re
import csv
from tkinter import filedialog as fd



class CSV_Class:
    def __init__(self, fPath, fDes):
        self.filepath = fPath
        self.fDest = fDes

    def create(self):
        pdfObj = open(self.filepath, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfObj)
        numPages = pdfReader.numPages
        pageObj = pdfReader.getPage(numPages - 1)
        pageContents = pageObj.extractText()

        # Use the page contents of the PDF and split into individual strings to be checked for
        stringList = pageContents.split('\n')
        s_compAL = "Comp. AL:"
        s_avg = "Avg:"
        s_acd = "ACD:"
        r = "OD (right)"
        l = "OS (left)"
        pID = '[0-9]-[0-9]'
        s_date = '[0-1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]'

        # Headers to be used in the CSV document
        header1 = "Axial Length (mm)"
        header2 = "Corneal Curvature K1 (D)"
        header3 = "Corneal Curvature K2 (D)"
        header4 = "Anterior Chamber Depth (mm)"
        abstractedContent = []
        d = []

        # Abstract the data needed from the pdf
        for x in stringList:
            a = re.search(s_compAL, x)
            b = re.search(s_avg, x)
            c = re.search(s_acd, x)
            i = re.search(pID, x)
            dates = re.search(s_date, x)

            if a:
                dString = a.string.split(' ')
                abstractedContent.append(dString[2])
            if b:
                dString = b.string.split(' ')
                e = dString[1].split('/')
                abstractedContent.append(e[0])
                abstractedContent.append(e[1])
            if c:
                dString = c.string.split(' ')
                abstractedContent.append(dString[2])

            if dates:
                d.append(dates.string)
                # print(dates.string)
            if i:
                id = i.string
                # print(id)

        # The exam date is always going to be the second date
        eDate = d[1]

        # Check to make sure that there is an ID otherwise leave it blank
        if "id" in locals():
            id = id
        else:
            id = ""

        # Make the CSV File
        csv_name = id + '_IOMasterInfo.csv'

        # Change to the directory where the csv should be saved
        os.chdir(self.fDest)

        # Check to see if the file already exists with the name generated above if so add a leading number to the
        # file name
        if path.exists(csv_name):
            i = 0
            csv_name2 = csv_name
            while path.exists(csv_name2):
                i = i + 1
                csv_name2 = str(i) + '_' + csv_name
        else:
            csv_name2 = csv_name

        # Layout the information into the way it will be written to the CSV
        header = [stringList[3], "Eye", header1, header2, header3, header4]
        row1 = [eDate, r, abstractedContent[0], abstractedContent[2], abstractedContent[3], abstractedContent[6]]
        row2 = [eDate, l, abstractedContent[1], abstractedContent[4], abstractedContent[5], abstractedContent[7]]

        # Create and populate the CSV with the name from above
        with open(csv_name2, 'w', encoding='UTF8', newline='') as f:
            writer = csv.writer(f)
            # Write the measured values
            writer.writerow(header)
            writer.writerow(row1)
            writer.writerow(row2)


if __name__ == "__main__":
    file = fd.askopenfilename()
    fileDestination = fd.askdirectory()
    # print(fileDestination)
    c = CSV_Class(file, fileDestination)
    c.create()
