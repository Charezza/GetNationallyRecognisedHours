import os
import requests
from bs4 import BeautifulSoup
import re
import pandas as pd
# import tkinter as tk
# from tkinter import filedialog
import csv
import numpy as np
import glob
import csv
from xlsxwriter.workbook import Workbook
from xlsxwriter.utility import xl_rowcol_to_cell

### GET THE FILE ###
NCVERlink = "https://www.ncver.edu.au/rto-hub/statistical-standard-software/nationally-agreed-nominal-hours"

def getSaveFolder():
	# root = tk.Tk()
	# root.withdraw()
	WorkingDir = os.getcwd()
	return WorkingDir

def getHTMLdocument(url):
    # request for HTML document of given url
    response = requests.get(url)
      
    # response will be provided in JSON format
    return response.text

html_document = getHTMLdocument(NCVERlink)
soup = BeautifulSoup(html_document, 'html.parser')

for link in soup.find_all('a', attrs={'href': re.compile(".txt")}):
    # display the actual urls
    link = (link.get('href'))  

# Download the file from the NVCER site
def download(url: str, dest_folder: str):
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)  # create folder if it does not exist

    filename = url.split('/')[-1].replace(" ", "_")  # be careful with file names
    file_path = os.path.join(dest_folder, filename)
    absPath = os.path.abspath(file_path)

    r = requests.get(url, stream=True)
    if r.ok:
        print("saving to", absPath)
        with open(file_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=1024 * 8):
                if chunk:
                    f.write(chunk)
                    f.flush()
                    os.fsync(f.fileno())
    else:  # HTTP status code 4XX/5XX
        print("Download failed: status code {}\n{}".format(r.status_code, r.text))
    base = os.path.splitext(os.path.basename(absPath))[0]
    csvFile = dest_folder+'/'+base+'.csv'
    xlsxFile = dest_folder+'/'+base+'.xlsx'
    read_file = pd.read_csv (absPath, sep = '\t')
    read_file.to_csv (csvFile, index=None, encoding='utf-8')
    os.remove (absPath)
# Convert to .xlsx
    for csvfile in glob.glob(os.path.join('.', '*.csv')):
        workbook = Workbook(csvfile[:-4] + '.xlsx')

#Add search formula to new sheet
        formulaSheet = workbook.add_worksheet("Search")
        formulaSheet.write(0,0,'Paste Units to Search here')
        formulaSheet.write(0,1,'Unit Name')
        formulaSheet.write(0,2,'Hours')
        formulaSheet.set_column(0,2,25)
        for row_num in range(1, 100):
            cell = xl_rowcol_to_cell(row_num, 0)
            formula = '=IFERROR(VLOOKUP(%s,Hours!A1:C70000,3,0),"")' % cell
            formulaSheet.write(row_num, 2, formula)
#Add the hours to a new sheet
        worksheet = workbook.add_worksheet("Hours")
        with open(csvfile, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)


        workbook.close()
    os.remove (csvFile)
    os.startfile(xlsxFile)

SaveFolder = getSaveFolder()
download(link, dest_folder=SaveFolder)