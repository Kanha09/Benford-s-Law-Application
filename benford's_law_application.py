import csv
from openpyxl import load_workbook

my_vals= []
my_strs= []
distribution= {}
my_percentages= []

wb= load_workbook(filename='transaction1.xlsx')
sheetranges= wb['Sheet1']
print(type(sheetranges['A1']))

for cell in sheetranges['B'][1:]:
    my_vals.append(cell.value)

def distribution_check():
    str1,str2,str3,str4,str5,str6,str7,str8,str9= 0,0,0,0,0,0,0,0,0

    for change in my_vals:
        change= str(change)
        my_strs.append(change)

    for check in my_strs:
        firstNum= check[0]
        firstNum= int(firstNum)

        if firstNum==1:
            str1+=1
        elif firstNum==2:
            str2+=1

        elif firstNum==3:
            str3+=1

        elif firstNum==4:
            str4+=1

        elif firstNum==5:
            str5+=1
        elif firstNum==6:
            str6+=1

        elif firstNum==7:
            str7+=1

        elif firstNum==8:
            str8+=1

        elif firstNum==9:
            str9+=1
        all_res= [str1,str2,str3,str4,str5,str6,str7,str8,str9]

    def calculate_percentage():
        for fin_res in all_res:
            my_percen= fin_res/len(my_strs)*100
            my_percentages.append(my_percen)

        for display in my_percentages:
            display = str(round(display, 2))
            print(display,'%')

        distribution[1]= round(my_percentages[0],2)
        distribution[2]= round(my_percentages[1],2)
        distribution[3]= round(my_percentages[2],2)
        distribution[4]= round(my_percentages[3],2)
        distribution[5]= round(my_percentages[4],2)
        distribution[6]= round(my_percentages[5],2)
        distribution[7]= round(my_percentages[6],2)
        distribution[8]= round(my_percentages[7],2)
        distribution[9]= round(my_percentages[8],2)
    calculate_percentage()

distribution_check()

print(distribution)

with open('percens.csv', 'w', newline="") as csv_file:
    writer = csv.writer(csv_file)
    for key, value in distribution.items():
        writer.writerow([key, value])
