# Libraries needed
import os
from os import path
from pdfminer.high_level import extract_text
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
            text = extract_text(self.filepath)
            return text
        else:
            return -1

    def create_v5(self, pCont):
        # Use the page contents of the PDF and split into individual strings to be checked for
        stringList = pCont.split('\n')

        # Comparisons to use for the version 5 documents
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

        # Arrays used to store information
        extractedContent = []
        d = []
        v1 = []
        v2 = []
        v3 = []

        # Extract the data needed from the pdf content
        for x in stringList:
            # Regular expressions found for the current line
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

        # Check to see if any of the v arrays are empty or missing an eye's measurement - if so add N/As to arrays
        if not v1:
            v1.append("N/A")
            v1.append("N/A")
        elif len(v1) == 1:
            v1.append("N/A")

        if not v2:
            v2.append("N/A/N/A")
            v2.append("N/A/N/A")
        elif len(v2) == 1:
            v2.append("N/A/N/A")

        if not v3:
            v3.append("N/A")
            v3.append("N/A")
        elif len(v3) == 1:
            v3.append("N/A")

        # Make sure that the correct values are taken from the values of interest
        # Axial Length Information
        if len(v1) == 2:  # axial values
            if v1[0] == "N/A":
                extractedContent.append(v1[0])
            else:
                temp = v1[0].split(' ')
                extractedContent.append(temp[2])

            if v1[1] == "N/A":
                extractedContent.append(v1[1])
            else:
                temp = v1[1].split(' ')
                extractedContent.append(temp[2])

        # Corneal Curvature Values
        if len(v2) == 2:  # k1 and k2 values
            if v2[0] == "N/A/N/A":
                extractedContent.append("N/A")
                extractedContent.append("N/A")
            else:
                temp = v2[0].split(' ')
                t = temp[1].split('/')
                extractedContent.append(t[0])
                extractedContent.append(t[1])
            if v2[1] == "N/A/N/A":
                extractedContent.append("N/A")
                extractedContent.append("N/A")
            else:
                temp = v2[1].split(' ')
                t = temp[1].split('/')
                extractedContent.append(t[0])
                extractedContent.append(t[1])

        # Anterior Chamber Depth Values
        if len(v3) == 2:  # chamber depth
            if v3[0] == "N/A":
                extractedContent.append(v3[0])
            else:
                temp = v3[0].split(' ')
                extractedContent.append(temp[2])
            if v3[1] == "N/A":
                extractedContent.append(v3[1])
            else:
                temp = v3[1].split(' ')
                extractedContent.append(temp[2])

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
        row1 = [eDate, r, extractedContent[0], extractedContent[2], extractedContent[3], extractedContent[6]]
        row2 = [eDate, l, extractedContent[1], extractedContent[4], extractedContent[5], extractedContent[7]]

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
            ('Exam Date', eDate),  # Exam Date
            ('R Eye', r),  # Right eye label
            ('AL R', extractedContent[0]),  # Axial Length of the right eye
            ('K1 R', extractedContent[2]),  # Corneal Curvature Values k1 for the right eye
            ('K2 R', extractedContent[3]),  # Corneal Curvature Values k2 for the right eye
            ('ACD R', extractedContent[6]),  # Anterior Chamber Depth Values for the right eye
            ('L Eye', l),  # Left eye label
            ('Al L', extractedContent[1]),  # Axial Length of the left eye
            ('K1 L', extractedContent[4]),  # Corneal Curvature Values K1 for the left eye
            ('K2 L', extractedContent[5]),  # Corneal Curvature Values K2 for the left eye
            ('ACD L', extractedContent[7])  # Anterior Chamber Depth Values for the left eye
        ])

        return directory

    def create_v7(self, pCont):
        stringList = pCont.split('\n')

        # Remove blanks from list
        while '' in stringList:
            stringList.remove('')

        teeth = len(stringList)

        # Regular Expressions to find certain strings
        al_val = '[0-9]'
        s_date = '[0-3][0-9]/[0-1][0-9]/[0-2][0-9][0-9][0-9]'
        OS = 'left'
        OD = 'right'

        # Headers to be used in the CSV document - might not use
        header1 = "Axial Length (mm)"
        header2 = "CCT (um)"
        header3 = "Anterior Chamber Depth (mm)"
        header4 = "LT (mm)"
        header5 = "SE (D)"
        header6 = "Corneal Curvature K1 (D)"
        header7 = "Corneal Curvature K2 (D)"
        header8 = "Delta K (D)"

        # Arrays to store values
        i = 0  # index into the stringList
        extractedContent = []  # For the CSV down below
        bioVals = []  # Want the first 8 entries that correspond to the analyze page of the PDF
        dates = []  # Date of exam seems to be the third date found
        eye = []  # Eye from the analyze page is the first one found
        t = 0

        for x in stringList:
            # Find the Patient ID
            if x == 'Physician':
                id = stringList[i + 2]

            # Find all the dates in the strings
            d = re.search(s_date, x)
            if d:
                dates.append(d.string)

            # Find the eye for the csv
            r = re.search(OD, x)
            l = re.search(OS, x)
            if r:
                eye.append(r.string)
            elif l:
                eye.append(l.string)

            # Find the four following Biometric values AL, CCT, ACD and LT
            if x == 'AL:':
                bioVals.append(x)
                # First number following tag is the value of it
                p = i + 1
                while t == 0:
                    temp = stringList[p]  # get the next string in the list
                    a = re.match(al_val, temp)  # match it to be a number
                    if a:
                        al = a.string
                        bioVals.append(al)  # Add it to the bioVal statistics array
                        t = 1  # Break the while loop
                    p += 1
                t = 0
            elif x == "CCT:":
                bioVals.append(x)
                # The number after the AL value number is the one we want
                cct = findValue(i, stringList, al, teeth)
                bioVals.append(cct)
            elif x == "ACD:":
                bioVals.append(x)
                acd = findValue(i, stringList, cct, teeth)
                bioVals.append(acd)
            elif x == "LT:":
                bioVals.append(x)
                lt = findValue(i, stringList, acd, teeth)
                bioVals.append(lt)
            elif x == "SE:":
                bioVals.append(x)
                # First number following tag is the value of it
                p = i + 1
                while t == 0:
                    temp = stringList[p]  # get the next string in the list
                    b = re.match(al_val, temp)  # match it to be a number
                    if b:
                        se = b.string
                        bioVals.append(se)  # Add it to the bioVal statistics array
                        t = 1  # Break the while loop
                    p += 1
                t = 0
            elif x == "K1:":
                bioVals.append(x)
                k1 = findValue(i, stringList, se, teeth)
                bioVals.append(k1)
            elif x == "K2:":
                bioVals.append(x)
                k2 = findValue(i, stringList, k1, teeth)
                bioVals.append(k2)
            elif x == "ΔK:":
                bioVals.append(x)
                dk = findValue(i, stringList, k2, teeth)
                bioVals.append(dk)

            i += 1  # Increment the index as the for loop goes through

        # Get the data from the arrays
        eDate = dates[2]

        if eye[0] == 'right':
            eyeFound = "OD (right)"
        elif eye[0] == 'left':
            eyeFound = "OS (left)"

        for i in range(0, 16):
            extractedContent.append(bioVals[i])

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
        header = ["Exam Date", "Eye", header1, header2, header3, header4, header5, header6, header7, header8]
        row1 = [eDate, eyeFound, extractedContent[1], extractedContent[3], extractedContent[5], extractedContent[7],
                extractedContent[9], extractedContent[11], extractedContent[13], extractedContent[15]]

        # Create and populate the CSV with the name from above
        with open(csv_name2, 'w', encoding='UTF8', newline='') as f:
            writer = csv.writer(f)

            # Write the measured values
            writer.writerow(header)
            writer.writerow(row1)

        # Set up the directory to be returned on top of the creation of the CSV
        directory = dict([
            ('ID', id),  # ID number if applicable
            ('Exam Date', eDate),  # Exam Date
            ('Eye', eyeFound),  # Right eye label
            ('AL', extractedContent[1]),  # Axial Length of the right eye
            ('CCT', extractedContent[3]),
            ('ACD', extractedContent[5]),  # Anterior Chamber Depth Values for the right eye
            ('LT', extractedContent[7]),
            ('SE', extractedContent[9]),
            ('K1', extractedContent[11]),  # Corneal Curvature Values k1 for the right eye
            ('K2', extractedContent[13]),  # Corneal Curvature Values k2 for the right eye
            ('ΔK', extractedContent[15])  # Anterior Chamber Depth Values for the left eye
        ])

        return directory


def findValue(i, arr, val, lenlen):
    p = i + 1
    t = 0
    while t == 0:
        temp = arr[p]
        if temp == val:
            d = arr[p + 1]
            found = d
            t = 1
        p += 1
        if p >= lenlen:
            found = 0
            return found

    return found


if __name__ == "__main__":
    file = fd.askopenfilename()
    fileDestination = fd.askdirectory()

    option = input("Create your own name for the CSV? Y or N\n")

    if option == 'Y':
        name = input('Enter name of the CSV\n')
    else:
        name = ''

    c2 = CSV_Class(file, fileDestination, name)

    if c2 == -1:
        print("PDF selected does not exist!\n\n")
    else:
        content = c2.openFile()
        version = input("What version are you using? 5 or 7?\n")
        if version == "5":
            doc = c2.create_v5(content)
            print(doc)
        elif version == "7":
            doc2 = c2.create_v7(content)
            print(doc2)
