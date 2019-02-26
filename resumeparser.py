import csv
import re
import spacy
import sys
reload(sys)
import pandas as pd
sys.setdefaultencoding('utf8')
from StringIO import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import os
import sys, getopt
import numpy as np
from bs4 import BeautifulSoup
import urllib2
from urllib2 import urlopen
from xlsxwriter import Workbook
import glob

workbook = Workbook('first_file.xlsx')
worksheet = workbook.add_worksheet()
files_list = glob.glob('./Files/*.pdf')
print(files_list)

#Function converting pdf to string
def convert(fname, pages=None):
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)
    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)
    infile = file(fname, 'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)
    infile.close()
    converter.close()
    text = output.getvalue()
    output.close
    return text
    
#Function to extract names from the string using spacy
def extract_name(string):
    r1 = unicode(string)
    nlp = spacy.load('xx')
    doc = nlp(r1)
    for ent in doc.ents:
        if(ent.label_ == 'PER'):
            first = str(ent.text)
            print(first)
            break
    return first

#Function to extract Phone Numbers from string using regular expressions
def extract_phone_numbers(string):
    r = re.compile(r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{6}|\d{3}[-\.\s]??\d{3}[-\.\s]??\d{5}|\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4})')
    phone_numbers = r.findall(string)
    return [re.sub(r'\D', '', number) for number in phone_numbers]
#Function to extract Email address from a string using regular expressions
def extract_email_addresses(string):
    r = re.compile(r'[\w\.-]+@[\w\.-]+')
    return r.findall(string)
#Converting pdf to string
worksheet.write(0, 0, "Name")
worksheet.write(0, 1, "Phone Number")
worksheet.write(0, 2, "Email")

try:
    for var, i in enumerate(files_list):
        resume_string = convert(i)
        resume_string1 = resume_string
        #Removing commas in the resume for an effecient check
        resume_string = resume_string.replace(',',' ')
        #Converting all the charachters in lower case
        resume_string = resume_string.lower()
        first = extract_name(resume_string1)
        print('\n')
        print('Phone Number is')
        y = extract_phone_numbers(resume_string)
        y1 = []
        for i in range(len(y)):
            if(len(y[i])>9):
                y1.append(y[i])
        print(y1)
        print('\n')
        print('Email id is')
        print(extract_email_addresses(resume_string))
        third = extract_email_addresses(resume_string)
        l1 = var+1
        if first[0:25]:
            worksheet.write(l1, 0, first[0:25])
        if y1:
            worksheet.write(l1, 1, y1[0])
        if third:
            worksheet.write(l1, 2, third[0])
    workbook.close()

except Exception as ex:
        print (str(ex))
        workbook.close()


