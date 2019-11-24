import os
import csv

election_data_csv = os.path.join("..","..","Instructions","PyPoll","Resources","election_data.csv")

#The total number of votes cast
with open(election_data_csv, newline='') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',')
        csv_header = next(csvreader)
        #print(csv_header)

        count_of_rows = int(sum(1 for line in csvfile))

#list of candinates who recieved votes
candidates = []
unique_candidates = []
with open(election_data_csv, newline='') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')
    next(csvreader)
    for i in csvreader:
        candidates.append(i[2])
for row in candidates:
    if row not in unique_candidates:
        unique_candidates.append(row)
        

count_for_name = []
#The number of votes each candidate won
for index in range(0,len(unique_candidates)):
        name = str(unique_candidates[index])
        index_C = [candidates for candidates in candidates if  candidates == name]
        to_add = [(len(index_C)),name]
        count_for_name.append(to_add)
#The precentage of votes each candidate won
precentages = []
for c in count_for_name:
        p = (int(c[0])/count_of_rows)*100
        adding_this = str(round(p,2))+"%"
        precentages.append(adding_this)

#The winner of the election based on popular vote.
        winning_number_of_votes = precentages.index(max(precentages))
        winner = str(unique_candidates[winning_number_of_votes])


#output

def elections_results():
        print("Election Results")
        print("---------------------------")
        print("Total Votes: "+str(count_of_rows))
        print("---------------------------")
        index_to_use = 0
        for row in count_for_name: 
                print(str(row[1])+" with "+str(row[0])+" votes, which is "+str(precentages[index_to_use]))
                index_to_use = index_to_use +1
        print("---------------------------")
        print("The winner is: " + winner)

elections_results()

file = open("elections.txt","w")
file.write(f"{elections_results()}")
file.close()


