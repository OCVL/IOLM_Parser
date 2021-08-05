# Libraries needed
import os
from os import path
import PyPDF2
import re
import csv
from tkinter import filedialog as fd


class CSV_Class:
    # Constructor that check to make sure the PDF exists
    def __init__(self, fPath, fDes, fName):
        self.filepath = fPath
        self.fDest = fDes
        self.name = fName

    # This is the function that opens the PDF and checks to make sure it exists
    def openFile(self):
        ex = path.exists(self.filepath)
        if ex:
            pdfObj = open(self.filepath, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfObj)
            numPages = pdfReader.numPages
            pageObj = pdfReader.getPage(numPages - 1)
            pageContents = pageObj.extractText()
            return pageContents
        else:
            return -1

    def create(self, pCont):
        # Use the page contents of the PDF and split into individual strings to be checked for
        stringList = pCont.split('\n')
        # print(stringList)
        s_compAL = "Comp. AL:"
        s_avg = "Avg:"
        s_acd = "ACD:"
        r = "OD (right)"
        l = "OS (left)"
        pID1 = '[0-9]-[0-9]'
        pID2 = '\S-[0-9]+'
        pID3 = '_[0-9]'
        pID = pID1 + '|' + pID2 + '|' + pID3  # Covers both of the IDs seen so far
        s_date = '[0-1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]'

        # Headers to be used in the CSV document
        header1 = "Axial Length (mm)"
        header2 = "Corneal Curvature K1 (D)"
        header3 = "Corneal Curvature K2 (D)"
        header4 = "Anterior Chamber Depth (mm)"
        abstractedContent = []
        d = []
        v1 = []
        v2 = []
        v3 = []

        # Abstract the data needed from the pdf
        for x in stringList:
            a = re.search(s_compAL, x)
            b = re.search(s_avg, x)
            c = re.search(s_acd, x)
            i = re.search(pID, x)
            dates = re.search(s_date, x)

            if a:  # Finds the axial length values
                v1.append(a.string)
            if b:  # Finds the corneal curvature (k1 and k2) values
                v2.append(b.string)
            if c:  # Finds the anterior chamber depth values
                v3.append(c.string)
            if dates:  # Finds the dates
                d.append(dates.string)
            if i:  # Finds the ID
                id = i.string
                # print(id)

        # Make sure that the correct values are taken from the values of interest
        if len(v1) == 2:  # axial values
            temp = v1[0].split(' ')
            # print(temp)
            abstractedContent.append(temp[2])
            temp = v1[1].split(' ')
            abstractedContent.append(temp[2])
        elif len(v1) == 1:
            temp = v1[0].split(' ')
            abstractedContent.append(temp[2])
            abstractedContent.append("")

        if len(v2) == 2:  # k1 and k2 values
            temp = v2[0].split(' ')
            t = temp[1].split('/')
            abstractedContent.append(t[0])
            abstractedContent.append(t[1])

            temp = v2[1].split(' ')
            t = temp[1].split('/')
            abstractedContent.append(t[0])
            abstractedContent.append(t[1])
        elif len(v2) == 1:
            temp = v2[0].split(' ')
            t = temp[1].split('/')
            abstractedContent.append(t[0])
            abstractedContent.append(t[1])
            abstractedContent.append("")
            abstractedContent.append("")

        if len(v3) == 2:  # chamber depth
            temp = v3[0].split(' ')
            abstractedContent.append(temp[2])
            temp = v3[1].split(' ')
            abstractedContent.append(temp[2])
        elif len(v3) == 1:
            temp = v3[0].split(' ')
            abstractedContent.append(temp[2])
            abstractedContent.append("")

        # The exam date is always going to be the second date
        if len(d) == 3:
            eDate = d[1]
        else:
            eDate = d[0]

        # Check to make sure that there is an ID otherwise leave it blank
        if "id" not in locals():
            v = input('No ID was detected. Do you want to enter one? Y - N\n')
            if v == 'Y':
                q = input('Enter the id:\n')
                id = q
            else:
                id = ""

        # Did the user give a name for the CSV?
        if self.name == '':
            # Make the CSV File if needed
            csv_name = id + '_IOMasterInfo.csv'
        else:
            csv_name = self.name + '_IOMasterInfo.csv'

        # Change to the directory where the csv should be saved
        os.chdir(self.fDest)

        # Check to see if the file already exists with the name generated above if so add a leading number to the
        # file name
        if path.exists(csv_name):
            i = 0
            csv_name2 = csv_name
            while path.exists(csv_name2):
                csv_name2 = str(i) + '_' + csv_name
                i = i + 1
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

        # Set up the directory to be returned on top of the creation of the CSV
        directory = dict([
            ('ID', id),  # ID number if applicable
            ('R Eye', r),  # Right eye label
            ('AL R', abstractedContent[0]),  # Axial Length of the right eye
            ('K1 R', abstractedContent[2]),  # Corneal Curvature Values k1 for the right eye
            ('K2 R', abstractedContent[3]),  # Corneal Curvature Values k2 for the right eye
            ('ACD R', abstractedContent[6]),  # Anterior Chamber Depth Values for the right eye
            ('L Eye', l),  # Left eye label
            ('Al L', abstractedContent[1]),  # Axial Length of the left eye
            ('K1 L', abstractedContent[4]),  # Corneal Curvature Values K1 for the left eye
            ('K2 L', abstractedContent[5]),  # Corneal Curvature Values K2 for the left eye
            ('ACD L', abstractedContent[7])  # Anterior Chamber Depth Values for the left eye
        ])
        return directory


if __name__ == "__main__":
    file = fd.askopenfilename()
    fileDestination = fd.askdirectory()

    option = input("Create your own name for the CSV? Y or N\n")

    if option == 'Y':
        name = input('Enter name of the CSV\n')
    else:
        name = ''

    print(name)

    c = CSV_Class(file, fileDestination, name)

    if c == -1:
        print("PDF selected does not exist!\n")
    else:
        content = c.openFile()
        d = c.create(content)

    print(d)
