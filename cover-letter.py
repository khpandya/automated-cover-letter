import docx
import shutil
import os
from pathlib import Path
from docx.shared import Pt
from datetime import date
from docx2pdf import convert

def manageDocs(companyName,jobTitle):
    """
    makes a folder named the company name if it doesn't exist and copies the cover letter template there and renames it to the job title
    """
    pathCompany=companyName.replace(" ","")
    pathJobTitle=jobTitle.replace(" ","")
    newPath=os.path.join(os.getcwd(),pathCompany)
    Path(newPath).mkdir(parents=True,exist_ok=True)
    destination=newPath+'\\'+pathJobTitle+'.docx'
    shutil.copy('.\\cover-letter.docx',destination)
    return destination

def replace_string(filename,find,replace):
    """
    for a specified path to a file, finds all instances of a string and replaces it with desired string
    """
    doc = docx.Document(filename)
    style=doc.styles['Normal']
    font=style.font 
    font.name='Garamond'
    font.size=Pt(12)
    for p in doc.paragraphs:
        if find in p.text:
            print('SEARCH FOUND!!')
            text = p.text.replace(find, replace)
            p.text = text
            p.style = doc.styles['Normal']
    doc.save(filename)

def convertToPDF(docxpath):
    convert(docxpath)

companyName=input("company name: ")
jobTitle=input("job title: ")
jobId=input("job ID: ")
if jobId=="d":
    jobId=""
else:
    jobId=" with job id "+jobId


replaceDict={'#companyName#':companyName,
            '#date#':date.today().strftime("%B %d, %Y"),
            '#jobTitle#':jobTitle,
            '#jobId#':jobId
            }

destination=manageDocs(companyName,jobTitle)

for find in replaceDict:
    replace_string(destination,find,replaceDict[find])

convertToPDF(destination)

