import docx
from docx import Document
import os
import subprocess
import collections
import glob
import sys
import string
import enchant
import csv
import pandas as pd
from pandas import DataFrame

############################Read the Oxford Idioms dictionary using the docx package.################################
document = docx.Document('Idioms_dictionary.docx')

bolds=[]
meaning = []
meaning2 = []
meaning3 = []
meaning4 = []
meaning5 = []
italics=[]
txtfiles = []
i=0
indicator = 0
index = 0

sentence_list1 = []
sentence_list2 = []
sentence_list3 = []
sentence_list4 = []

break_indicator = 0
idiom_words = []

##################################Breaking up the dictionary into paragraphs.##################################
for para in document.paragraphs:
    a = 0
    m2 = ""
    m3 = ""
    m4 = ""
    m5 = ""
    break_indicator_m3 = 0
    break_indicator_m4 = 0
    break_indicator_m5 = 0
    idiom_words = []
    break_indicator_m2 = 0
    b_i = 0
    i=i+1
    ##Started from the 40th paragraph in order to skip the Introductory part of the dictionary.
    if i>=40:
        ##Capture complete text stored inside each paragraph.
        for run in para.runs:
            ##Capture the bold part of the text contained in a paragraph, i.e. the Idiom.
            if run.bold:
                ##Ignore the page numbers and titles 
                run.text = ''.join([i for i in run.text if not i.isdigit()])
                run.text = ''.join([i for i in run.text if not i == '.'])
                if len(run.text) > 2:
                    ##Strip the text contained into parts of words having the same font style 
                    run.text = run.text.rstrip()
                    if run.text not in bolds:
                        ##Add the words in bol(idiom) to the list bold if they haven't already been added.
                        bolds.append(run.text)
                        index = index + 1
                else:
                    ##If length is lesser than 2 then it indicates the captured text is a heading in the dictionary which is meant to be skipped.
                    b_i = 1
                    break
            ##If the captured text is neither bold nor italics, this indicates that it is the meaning of the idiom.
            elif not run.bold and not run.italic:
                run.text = ''.join([i for i in run.text if not i.isdigit()])
                if len(run.text) > 3:
                    run.text = run.text.rstrip()
                    word_list = run.text.split()
                    last = len(word_list) - 1
                    ##Skipping unnecassary parts of the dictionary
                    if "See" == word_list[0] or "Copyright" == word_list[0]:
                        continue
                    if "AnD" == word_list[0]:
                        indicator = 1
                        continue
                    ##Splitting different meanings on the basis of a semi-colon
                    if ";" in run.text:
                        run.text,m2 = run.text.split(';',1)
                        if ";" in m2:
                            m2,m3 = m2.split(';',1)
                            if ";" in m3:
                                m3,m4 = m3.split(';',1)
                                if ";" in m4:
                                    m4,m5 = m4.split(';',1)
                        ##Adding different meanings to different lists
                        if m2 != "":
                            break_indicator_m2 = 1
                            meaning2.append(m2)
                        if m3 != "":
                            break_indicator_m3 = 1
                            meaning3.append(m3)
                        if m4 != "":
                            break_indicator_m4 = 1
                            meaning4.append(m4)
                        if m5 != "":
                            break_indicator_m5 = 1
                            meaning5.append(m5)
                    meaning.append(run.text)
                    ##adding spaces if there is just 1 meaning to an idiom
                    if break_indicator_m2 == 0:
                        meaning2.append("")
                    if break_indicator_m3 == 0:
                        meaning3.append("")
                    if break_indicator_m4 == 0:
                        meaning4.append("")
                    if break_indicator_m5 == 0:
                        meaning5.append("")
                    if indicator == 1:
                        meaning.append(run.text)
                        indicator = 0

        sl1=0
        sl2=0
        sl3=0
        sl4=0
        inc1 = 0
        inc2 = 0
        if b_i == 1:
            continue
        for r in para.runs:
            if r.text.isspace():
               continue
            ##Capturing given sentences that have been framed using the idioms. They are written in Italics format.
            if r.italic and sl1 == 0 and sl2 == 0 and sl3 == 0 and sl4 == 0:
                if r.text[len(r.text)-1] != "." and r.text[len(r.text)-2] != "." and r.text[len(r.text)-1] != "?" and r.text[len(r.text)-2] != "?" and r.text[len(r.text)-1] != "!" and r.text[len(r.text)-2] != "!":
                    incomp = "".join([incomp,r.text])
                    inc1 = 1
                    continue
                if inc1 == 1:
                    r.text = "".join([incomp,r.text])
                ##Adding sentences to their respective sentence list.
                sentence_list1.append(r.text)
                sl1 = sl1+1
                inc1 = 0
                incomp = ""
                continue
            ##Checking if there are more than one example sentences for a particular idiom
            if r.italic and sl1 == 1 and sl2 == 0 and sl3 == 0 and sl4 == 0:
                if r.text[len(r.text)-1] != "." and r.text[len(r.text)-2] != "." and r.text[len(r.text)-1] != "?" and r.text[len(r.text)-2] != "?" and r.text[len(r.text)-1] != "!" and r.text[len(r.text)-2] != "!":
                    incomp = "".join([incomp,r.text])
                    inc2 = 1
                    continue
                if inc2 == 1:
                    r.text = "".join([incomp,r.text])
                sentence_list2.append(r.text)
                sl2 = sl2+1
                inc2 = 0
                incomp = ""
                continue
            if r.italic and sl1 == 1 and sl2 == 1 and sl3 == 0 and sl4 == 0:
                sentence_list3.append(r.text)
                sl3 = sl3+1
                continue
            if r.italic and sl1 == 1 and sl2 == 1 and sl3 == 1 and sl4 == 0:
                sentence_list4.append(r.text)
                sl4 = sl4+1
        ##Append blank spaces in case sentences are not avaialbale
        if sl1 == 0:
           sentence_list1.append("")
        if sl2 == 0:
           sentence_list2.append("")   
        if sl3 == 0:
           sentence_list3.append("")
        if sl4 == 0:
           sentence_list4.append("")



##Creating a data frame to contain the list of idioms, their meanings, and there usage in different statements.
df1 = DataFrame(meaning)
df = DataFrame(bolds)
df2 = DataFrame(meaning2)
df3 = DataFrame(meaning3)
df4 = DataFrame(meaning4)
df5 = DataFrame(meaning5)
df6 = DataFrame(sentence_list1)
df7 = DataFrame(sentence_list2)
df8 = DataFrame(sentence_list3)
df9 = DataFrame(sentence_list4)
ct = pd.concat([df,df1,df2,df3,df4,df5,df6,df7,df8,df9], axis = 1)

##Creating a csv file to contain the lists that have been created.
h = ["Idioms", "Meaning 1", "Meaning 2", "Meaning 3", "Meaning 4", "Meaning 5", "Sentence 1", "Sentence 2", "Sentence 3", "Sentence 4"]
ct.to_csv('Idioms.csv', header = h, index = None, encoding='utf-8-sig')
###############################End######################################
