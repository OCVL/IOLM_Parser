# IOLM_Parser - A Python script for Extracting data from IOLMaster PDF documents
A repository that facilitates the parsing and extraction of information from IOLMaster PDFs.

### It performs 3 main functions:
* Allows for the PDF to be selected from the window explorer
* Allows for the CSV to be saved to a specific folder
* Abstracts the following information into a CSV and directory
     * Patient ID
     * Exam Date
     * Axial Length values
     * Corneal Curvature values
     * Anterior Chamber Depth Values

### Python 3.9 with the Following Libraries Needed:
* OS
* PyPDF2
* re
* csv
* tkinter

## Information on how to operate the script:
### Run IO_Master_Parsing.py:
When ran, the script prompts the user to select the PDF to be used in the first pop-up file explorer.
**Presently, there is no constraints to how the PDF file to be used needs to be named**

The next file explorer pop-up will prompt the user to navigate to where the CSV will be saved.

Once the destination of the CSV is selected, the script will ask the user via the terminal if they would like to enter their own name for the generated CSV. If this is desired **_'Y'_** should be entered. If **_'Y'_** was entered, the user is then asked to enter the name of the CSV file via the terminal. **_Note:_ This name should not have the extension of csv added, it is done inside the script** 

**If the user does not wish to name the CSV themselves, the ID from the PDF is used along with the string __IOMasterInfo.csv_. If an ID is not found during the PDF abstraction, the user will be prompted to enter one if it is known**

The values from the PDF will then be added to the CSV and saved to the selected destination. A directory is also displayed to the terminal that contains all the values that were entered into the CSV, along with the Patient's ID.


