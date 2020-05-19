import pandas as pd
import csv

######Average Calculation###############
noOfBooks = 20
idiomcount = pd.read_csv("Idiom_count.csv")
cat = idiomcount["Category"].tolist()
icount = idiomcount["Number of Idioms"]
books_list = idiomcount["Book name"]

iartstotal = 0
isciencetotal = 0
iedutotal = 0
ihisttotal = 0
ificttotal = 0

########Calculating the total number of idioms in the Collection of books#######
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
dict = {"Arts" : artsavg, "Science" : scienceavg, "Education" : edusavg, "History" :histsavg, "Fiction" : fictionavg}

########Printing the average number of Idioms to a CSV file######

with open("Average Number of Idioms.csv", 'w', newline='') as f:
    writer = csv.writer(f, delimiter=',')
    header = ['Book Category', 'Average']
    writer.writerow(header)
    for key, value in dict.items():
        writer.writerow([key, round(value)])

