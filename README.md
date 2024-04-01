**Title: Excel VBA Script for Daily Automatic Copying of CRMS Data to Analysis Workbook**

**Description:** 
The `VbaAutoCopy.txt` file provides a detailed explanation of the VBA code designed for Excel. This code automates the process of copying data from downloaded files (originating from Suzlon's CRMS) into an analysis workbook. The primary purpose is to analyze data related to clients' Wind Turbine Generators (WTGs).

**What to do before running the code:**

_File Naming:_
   When saving the downloaded files into a pre-decided location, the the user needs to change the file name to a specific format. For example, for a client named "gc," the file is saved with a name like "gc13-01-2022" (representing today's date). after that further steps are executed VBA.

**Features of this code:**

1. **Verification of Existing Data:**
   The code first checks for already present data in the analysis workbook before proceeding with the copying process.

3. **File Identification and Opening:**
   The code then identifies the file by its name in the system's location and opens it for further processing.

4. **Information Copying:**
   Relevant information from the opened file is copied into the analysis workbook.

5. **User Interaction via Message Boxes:**
   The code interacts with the user through strategically placed message boxes.
   

**UPDATES:**
   
New code file named ______ now automates data extraction from the CRMS website using Selenium, prompting the user for data retrieval confirmation. It compares data integrity before proceeding, then downloads and copies data from specified sources into a master data sheet. The testselenium subroutine interacts with the CRMS site using Selenium, while RenameTodayDownloadedFile renames the latest downloaded file. The call_fns subroutine orchestrates the sequence of extraction and renaming functions. Finally, MyPublicSub makes the functionality accessible via a button in Excel.







Note: 

- Refer to the `VbaAutoCopy.txt` file for the complete VBA code and a detailed explanation of each code line.

- This code will work as per the destination of folder where the files are downloaded, please edit the path and range as per personal needs.

- Adjust the file naming conventions and any client-specific details as needed in the VBA code.
