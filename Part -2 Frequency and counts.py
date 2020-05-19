import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import numpy as np
import re
import time
import csv
import docx
from docx import Document
import os
import subprocess
import collections
import glob
import sys
import string
import enchant
from pprint import pprint
import pandas as pd
from pandas import DataFrame

##Read the idioms dictionary that has been created by the python script of part 1.
d = pd.read_csv("Idioms.csv")
bolds = d["Idioms"].tolist()
            
file_names = []
number_of_idioms = []
directory = []
list_of_files = []
################Creating files containing frequencies of Idioms#######

##Reading books contained in the Docx files folder.
for d in os.listdir("Docx Files/"):
    file_names = []
    ##Skip the .DS_Store file
    if d!=".DS_Store":
        print(d)
        ##Pick books from all the categories.
        for f in glob.glob("Docx Files/"+d+"/*.docx"):
            ##Creating a list of files contained in each folder
            file_names.append(os.path.basename(f))
            list_of_files.append(os.path.basename(f))
        for f_name in file_names:
            ##Evaluating the text in contained in the books using Docx package
            file = docx.Document("Docx Files/"+d+"/"+f_name)
            l = []
            ##Picking up paragraphs in the books.
            for para1 in file.paragraphs:
                for run in para1.runs:
                    ##Iterating the list of idioms that are contained in Idioms.csv file
                    for idiom in bolds:
                        ##Checking if any of the idiom in that list is conatined in text of all the paragraphs.
                        if idiom in run.text:
                            ##appending the idioms found to a list. 
                            l.append(idiom)
            ##Calculating the number of idioms present in a book
            number_of_idioms.append(len(l))
            ##List to capture the containing folder of each book. 
            directory.append(d)
            ##Frequency count of each idiom in the book
            counter = collections.Counter(l)
            freq_dict = dict(counter)
            filename = f_name[:-5]+'.csv'
            ##Creating a csv file with the same name as that of the back, which will contain all the results obtained until now.
            with open("Count files/"+filename, 'w', newline='') as f:
                print(filename)
                writer = csv.writer(f, delimiter=',')
                ##header for the csv file.
                header = ['Book Name', 'Idiom', 'Frequency']
                writer.writerow(header)
                ##transferring the results in the form of rows to the csv file.
                for key in freq_dict.keys():
                    l = [f_name, key, freq_dict[key]]
                    writer.writerow(l)
        
################Creating Files with word and Idiom count#######################      
wordcountdict = {}
number_of_words = []

##For comparing the words with the words captured in the US dictionary.
d = enchant.Dict("en_US")

##Acessing txt format of the same books.
for file in os.listdir("TXT Files/"):
    if file.endswith(".txt"):
        words = []
        punct = ['\n','.',',','...']
        ##Extracting all the words contained in the books
        with open("TXT Files/"+file,'r', encoding='utf-8') as fi:
            for line in fi:
                for word in line.split(" "):
                    if(word and d.check(word)):
                       words.append(word)
        ##Adding the number of words in each book to a list.
        number_of_words.append(len(words))
        ##Creating a dictionary containing name of the book and the number of words in it.
        wordcountdict.update({file : len(words)})


######################################Printing the results to a CSV file.##################################
header = ['Category', 'Book name', 'Number of Idioms', 'Number of words']

##File containing the final results.
with open("Idiom_count.csv", "w", newline='') as f:
    writer = csv.writer(f, delimiter=',')
    writer.writerow(header) # write the header
    for i, file in enumerate(list_of_files):
        l = [directory[i], file, number_of_idioms[i], number_of_words[i]]
        writer.writerow(l)

######################################Average Calculation###################################################
noOfBooks = 4
idiomcount = pd.read_csv("Idiom_count.csv")
cat = idiomcount["Category"].tolist()
icount = idiomcount["Number of Idioms"]
books_list = idiomcount["Book name"]

iartstotal = 0
isciencetotal = 0
iedutotal = 0
ihisttotal = 0
ificttotal = 0
for i in range(len(cat)):
	if "Arts" == cat[i]:
		iartstotal = iartstotal + icount[i]
	if "Science" == cat[i]:
		isciencetotal = isciencetotal + icount[i]
	if "Education" == cat[i]:
		iedutotal = iedutotal + icount[i]
	if "History" == cat[i]:
		ihisttotal = ihisttotal + icount[i]	
	if "Fiction" == cat[i]:
		ificttotal = ificttotal + icount[i]
fictionavg = ificttotal / noOfBooks
artsavg = iartstotal / noOfBooks
scienceavg = isciencetotal / noOfBooks
edusavg = iedutotal / noOfBooks
histsavg = ihisttotal / noOfBooks
categories = list(set(cat))
average = [artsavg, edusavg, scienceavg, histsavg, fictionavg]
with open("Average Number of Idioms in each Category.csv", 'w', newline='') as f:
    writer = csv.writer(f, delimiter=',')
    header = ['Book Category', 'Average Number of Idioms']
    writer.writerow(header)
    for i in range(len(categories)):
        l = [categories[i], average[i]]
        writer.writerow(l)

print("Resultant files have been created.")
