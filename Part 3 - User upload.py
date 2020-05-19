import csv
import docx
from docx import Document
import pandas as pd
import collections
import enchant
import docx2txt
import string
import os

##Read the idioms dictionary that has been created by the python script of part 1.
d = pd.read_csv("Idioms.csv")
bolds = d["Idioms"].tolist()

file_names = []
words = []
l = []
book_exist_indicator = 0
##Pick the book/article that has been uploaded by the user. 
uploaded_file = os.listdir("Upload Directory/")
try:
    fname = uploaded_file[0]
except IndexError:
    print("File has not been uploaded")
    exit(0)
user_file = "Upload Directory/" + fname
if ".docx" not in fname:
    print("Invalid file")
    exit(0)
print(fname)
file = docx.Document(user_file)

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
number_of_idioms = len(l)

##Frequency count of each idiom in the book
counter = collections.Counter(l)
freq_dict = dict(counter)
filename = fname[:-5] +'.csv'
print("File containing the results :- Output.csv")
        
##########################Creating Files with word and Idiom count#############################       

##For comparing the words with the words captured in the US dictionary.
d = enchant.Dict("en_US")

##Convert format of the book that has to be analyzed from Docx to TXT.
MY_TEXT = docx2txt.process(user_file)

##Creating txt format of the same book.
with open(fname[:-5] +'.txt', "w") as textfile:
    print(MY_TEXT, file=textfile)

##Acessing the TXT format to count the number of words.
with open(fname[:-5] +'.txt' , 'r', encoding='utf-8') as fi:
    for line in fi:
        for word in line.split(" "):
            if(word and d.check(word)):
               words.append(word)

##Calculating the number of words in the book.
number_of_words = len(words)

######################################Printing the results to a CSV file.##################################
indicator = 0
if os.path.isfile("Output.csv"):
    d = pd.read_csv("Output.csv")
    books = set(d["Book name"])
    if fname in books:
        book_exist_indicator = 1
    id = len(books) + 1
else:
    indicator = 1

######Checking if the stats of an uploaded book already exist in the CSV file####
if book_exist_indicator != 1:
    with open("Output.csv", 'a+', newline='') as f:
        writer = csv.writer(f, delimiter=',')
        ##header for the csv file.
        header = ['Book Id', 'Book name', 'Number of Idioms', 'Number of words', 'Idiom', 'Frequency']
        if indicator == 1:
            writer.writerow(header)
            id = 1
        
        ##transferring the results in the form of rows to the csv file.
        for key in freq_dict.keys():
            l = [id, fname, number_of_idioms, number_of_words, key, freq_dict[key]]
            writer.writerow(l)

################Avoid Repetition###################################################        

if book_exist_indicator == 1:
    print("Book exists")
    lines = []
    books_list = []
    ####Rewrite the output file with just one occurence of the repeated book####
    with open("Output.csv", 'r') as readFile:
        reader = csv.reader(readFile)
        for row in reader:
            ####Skip the previous stats of the repeated book while copying the previous output file to the new one.##### 
            if row[1] != fname:
                lines.append(row)
                if row[0] != 'Book Id' and row[1] not in books_list:
                    books_list.append(row[1])
    new_id = len(books_list) +1
    ####Re-ordering the book ids of the books contained with the newest entry of repeated book######
    for l in lines:
        if l[0] != 'Book Id':
            l[0] = books_list.index(l[1]) + 1
    ###Writing the output file.######
    with open("Output.csv", 'w') as writeFile:
        writer = csv.writer(writeFile)
        writer.writerows(lines)
        for key in freq_dict.keys():
            l = [new_id, fname, number_of_idioms, number_of_words, key, freq_dict[key]]
            writer.writerow(l)

####Delete the book after the stats have been obtained####
os.remove("Upload Directory/"+ uploaded_file[0])
os.remove(fname[:-5] +'.txt')
