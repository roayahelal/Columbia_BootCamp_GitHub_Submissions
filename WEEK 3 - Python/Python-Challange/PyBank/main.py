import os
import csv

budget_data_csv = os.path.join("..","..","Instructions","PyBank","Resources","budget_data.csv")

#Total number of months
with open(budget_data_csv, newline='') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',')
        csv_header = next(csvreader)
        print(f"CSV Header: {csv_header}")
        #Total Months below
        count_of_rows = int(sum(1 for line in csvfile))

#Total amount
total_amount = 0
with open(budget_data_csv, newline='') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')
    next(csvreader)
    for line in csvreader:
        total_amount = int(line[1]) + total_amount
#Average change
list_of_changes = []
with open(budget_data_csv, newline='') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')
    next(csvreader)
    list_of_amounts = [ line[1] for line in csvreader]
previous_amount =[0]
previous_amount.extend(list_of_amounts)
for i in range(0,int(count_of_rows)): 
    list_of_amounts[i] = int(list_of_amounts[i]) 
    previous_amount[i] = int(previous_amount[i])
for i in range(0,int(count_of_rows)):
    list_of_changes.append(list_of_amounts[i]- previous_amount[i])
list_of_changes.remove(list_of_changes[0])
sum_of_monthly_changes = sum(list_of_changes)
Average_monthly_changes = round((sum_of_monthly_changes/len(list_of_changes)),2)


# greatest increase in profit 
increase_profits = ((max(list_of_changes)))
index_needed = list_of_changes.index(max(list_of_changes))
Date = []
ProfitLoss = []

with open(budget_data_csv, newline='') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',')
        for row in csvreader:
            Date.append(row[0])
            ProfitLoss.append(row[1])

increase_months = (Date[index_needed+2])

        
#greatest decrease in profit
decrease_profits = (min(list_of_changes))
min_index = list_of_changes.index(min(list_of_changes))
decrease_months = (Date[min_index+2])

#output
def financial_analysis():
    print("Financial Analysis")
    print("------------------------------")
    print("Total Months: " + str(count_of_rows))
    print("Total: $" + str(total_amount))
    print("Average Change: $" + str (Average_monthly_changes))
    print("Greatest Increase in Profits: " + str(increase_months) +  " $(" +  str(increase_profits)+")")
    print("Greatest Decrease in Profits: " + str(decrease_months) + " $(" + str(decrease_profits)+")")

financial_analysis()

file = open("financial.txt","w")
file.write(f"Financial Analysis\n")
file.write(f"------------------------------\n")
file.write(f"Total Months: {count_of_rows}\n")
file.write(f"Total: ${total_amount}\n")
file.write(f"Average Change: ${Average_monthly_changes}\n")
file.write(f"Greatest Increase in Profits: {increase_months} ${increase_profits}\n")
file.write(f"Greatest Decrease in Profits: {decrease_months} ${decrease_profits}\n")
file.close()


    





    

    
