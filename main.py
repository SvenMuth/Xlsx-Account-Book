import xlsxwriter
import os
import numpy as np
import subprocess as sp
from datetime import datetime

from functions import *

# Function to randomly write a salary per months and write entries in different cost types with different values
# For testing remove hashtag, also remove the marked hashtags below
# from testing import test
# repeat_test = 100

# Create list for different months and position in Xlsx-Document [col, row, month, index]
months = [[0, 2, "January", 0], [0, 3, "February", 1], [0, 4, "March", 2], [0, 5, "April", 3],
          [0, 6, "May", 4], [0, 7, "June", 5], [0, 8, "July", 6], [0, 9, "August", 7],
          [0, 10, "September", 8], [0, 11, "October", 9], [0, 12, "November", 10], [0, 13, "December", 11]]

# Create list for costs and index
costs_list = [['Rent', '6'], ['Credit', '7'], ['Car', '8'], ['Foods', '9'],
              ['Amazon', '10'], ['Sport', '11'], ['Other', '12']]

# Create empty lists which are later needed
# Lists to sort the database in categories
rent, credit, car, foods, amazon, sport, other = [], [], [], [], [], [], []

# List for database types
list_salary, list_expenses, = [], []

# Create variables which are later needed
chosen_month_number, chosen_month_name, replace_entry = "", "", ""

# Declare name of Xlsx-Document and Sheet
workbook = xlsxwriter.Workbook("Account-Book.xlsx")
worksheet = workbook.add_worksheet("2021")

# Create to formats for the workbook which are bold
cell_format11 = workbook.add_format()
cell_format11.set_bold()
cell_format11.set_font_size(11)

cell_format14 = workbook.add_format()
cell_format14.set_bold()
cell_format14.set_font_size(14)

# Create different colors for text in the workbook
f_turquoise = workbook.add_format({'bold': False, 'font_color': '#0B615E'})
f_darkgreen = workbook.add_format({'bold': False, 'font_color': '#088A08'})
f_yellow = workbook.add_format({'bold': False, 'font_color': '#C8CA30'})
f_purple = workbook.add_format({'bold': False, 'font_color': '#8A0886'})
f_grye = workbook.add_format({'bold': False, 'font_color': '#688A08'})
f_darkred = workbook.add_format({'bold': False, 'font_color': '#8A0808'})
f_deeppurple = workbook.add_format({'bold': False, 'font_color': '#3B0B2E'})
f_pinkred = workbook.add_format({'bold': False, 'font_color': '#8A084B'})
f_green = workbook.add_format({'bold': False, 'font_color': '#01DF01'})
f_red = workbook.add_format({'bold': False, 'font_color': '#DF0101'})
f_orange = workbook.add_format({'bold': False, 'font_color': '#F36105'})

# For testing remove hashtags
# test(months, costs_list, repeat_test)
# """

# Chose class
select_class = show_classes(chosen_month_number)

# Loop for main program
while select_class != "0":

    # Create Table if not already exist
    database()

    # Write salary
    if select_class == "1":
        # Collect the month which already have been written in Database
        set_salary = get_salary_month()

        # Check if set with months is empty --> first entry for salary
        if len(set_salary) == 0:
            income_month = int(input("Please insert salary for " + chosen_month_name + ":"))

            # Write information's about your salary into database
            datasql = ("1", chosen_month_name, income_month, 0, "0", "0")
            insert_sql(datasql)

        else:
            # Check if there is already an entry for the chosen month
            if chosen_month_name in set_salary:
                print("")
                replace_entry = str(input("Should the entry for " + chosen_month_name + " be replaced? [y, n]"))
                if replace_entry == "y":

                    # Set a new income for this month
                    income_month = int(input("Please insert new Salary for " + chosen_month_name + ":"))
                    # Update old entry
                    update_salary_entry(income_month, chosen_month_name)
                    print("Salary was updated")
                    print("")

                # Let the old entry and jump in main loop
                elif replace_entry == "n":
                    print("Salary wasn't updated")
                    print("")

                else:
                    print("Input invalid")

            # No entry found so just write a new entry for this month
            else:
                income_month = int(input("Please insert salary for " + chosen_month_name + ":"))

                # Write information's about your salary into database
                datasql = ("1", chosen_month_name, income_month, 0, "0", "0")
                insert_sql(datasql)

        select_class = show_classes(chosen_month_number)

    # Write expenses
    elif select_class == "2":

        # List cost categories
        entry = 1
        for cost, index in costs_list:
            print(str(entry) + ". " + cost)
            entry += 1

        print("You have chosen " + chosen_month_name)
        category_chosen = int(input("Please chose a category: [1 - 7] "))

        if 0 < category_chosen <= 8:

            print("You chose " + costs_list[category_chosen - 1][0])
            expanses_costs = int(input("Please input the costs: "))

            # Get the category
            category_expanses = costs_list[category_chosen - 1][0]

            # Write a commentary to your entry and append the right time of day and date
            commentary = str(input("Please write a commentary to the expense:"))
            commentary += " - " + datetime.now().strftime("%H:%M - %d.%m.%y")

            # Write information's about your costs into database
            datasql = ("2", chosen_month_name, 0, expanses_costs, category_expanses, commentary)
            insert_sql(datasql)

        else:
            print("Input is invalid")

        #  Change between classes
        select_class = show_classes(chosen_month_number)

    # Change month
    elif select_class == "3":
        # Change Month to which the inputs are done
        # Show the different months
        number = 1
        for row, col, month, index in months:
            print(str(number) + ": " + month)
            number += 1

        print("")
        chosen_month_number = int(input("Please select a month: [1-12] "))

        if 0 < chosen_month_number <= 12:
            chosen_month_name = months[chosen_month_number - 1][2]
            print("You have chosen " + chosen_month_name)
            print("")

        else:
            print("Input is invalid")

        #  Change between classes
        select_class = show_classes(chosen_month_number)
    
    # Delete all entries from the chosen month
    elif select_class == "4":
        delete_entry(chosen_month_name)
        print("All entries for " + chosen_month_name + " where deleted")

        select_class = show_classes(chosen_month_number)
        
    # Delete database
    elif select_class == "5":
        reset = str(input("Delete the previous database? [y/n]"))
        if reset == "y":
            # Delete file
            os.remove(r"database\database.db")
            print("Database was removed")
            print("")

        elif reset == "n":
            print("Database wasn't deleted")

        else:
            print("invalid Input")

        select_class = show_classes(chosen_month_number)

    else:
        print("Input is wrong")
        #  Change between classes
        select_class = show_classes(chosen_month_number)

# For testing remove hashtag
# """
# Get all entries from database
cdata = get_data_sql()

# Order them by the type ("1" = salary, "2" = costs)
for element in cdata:
    if element[0] == "1":
        list_salary.append(element)

    elif element[0] == "2":
        list_expenses.append(element)

# Take the costs an append them to the different list which are equal to the category
for expense in list_expenses:
    if expense[4] == costs_list[0][0]:
        rent.append(expense)

    elif expense[4] == costs_list[1][0]:
        credit.append(expense)

    elif expense[4] == costs_list[2][0]:
        car.append(expense)

    elif expense[4] == costs_list[3][0]:
        foods.append(expense)

    elif expense[4] == costs_list[4][0]:
        amazon.append(expense)

    elif expense[4] == costs_list[5][0]:
        sport.append(expense)

    else:
        other.append(expense)


# Collect the sum of the costs per months
costs_per_month = [0] * 12
for element in cdata:
    for col, row, month, index in months:
        if element[1] == month:
            costs_per_month[index] += element[3]

# Write months in worksheet
for col, row, month, index in months:
    worksheet.write(col, row, month, cell_format14)

worksheet.write(1, 0, "Income", cell_format14)
worksheet.write(2, 1, "Salary", cell_format11)
worksheet.write(4, 0, "Costs", cell_format14)

# Write salary per month in Worksheet + sort salary per months in list
expenses_per_month = [0] * 12

for salary_month in list_salary:
    for col, row, month, index in months:
        if salary_month[1] in month:
            expenses_per_month[index] = salary_month[2]
            val = col + 2
            worksheet.write_number(val, row, salary_month[2], f_grye)

# Sort costs by months and write them in the right place in the Xlsx document
worksheet.write(5, 1, costs_list[0][0], cell_format11)
number_a = [5] * 12

for element in rent:
    for col, row, month, index in months:
        val = number_a[index]
        if element[1] in month:
            worksheet.write_number(val, row, element[3], f_turquoise)
            worksheet.write_comment(val, row, element[5])

            # Increase the position +1 at the right position in list
            number_a[index] = val + 1

# Start position for next category --> start at the highest position
t1 = max(number_a)

# Check if an entry was made in the last category, otherwise add the value 1 to the start position
if t1 == 5:
    t1 += 1
    worksheet.write(t1, 1, costs_list[1][0], cell_format11)

else:
    worksheet.write(t1, 1, costs_list[1][0], cell_format11)

number_b = [t1] * 12
for element in credit:
    for col, row, month, index in months:
        val = number_b[index]
        if element[1] in month:
            worksheet.write_number(val, row, element[3], f_darkred)
            worksheet.write_comment(val, row, element[5])
            number_b[index] = val + 1

t2 = max(number_b)

if t1 == t2:
    t2 += 1
    worksheet.write(t2, 1, costs_list[2][0], cell_format11)

else:
    worksheet.write(t2, 1, costs_list[2][0], cell_format11)

number_c = [t2] * 12
for element in car:
    for col, row, month, index in months:
        val = number_c[index]
        if element[1] in month:
            worksheet.write_number(val, row, element[3], f_deeppurple)
            worksheet.write_comment(val, row, element[5])
            number_c[index] = val + 1

t3 = max(number_c)

if t2 == t3:
    t3 += 1
    worksheet.write(t3, 1, costs_list[3][0], cell_format11)

else:
    worksheet.write(t3, 1, costs_list[3][0], cell_format11)

number_d = [t3] * 12
for element in foods:
    for col, row, month, index in months:
        val = number_d[index]
        if element[1] in month:
            worksheet.write_number(val, row, element[3], f_darkgreen)
            worksheet.write_comment(val, row, element[5])
            number_d[index] = val + 1

t4 = max(number_d)

if t3 == t4:
    t4 += 1
    worksheet.write(t4, 1, costs_list[4][0], cell_format11)

else:
    worksheet.write(t4, 1, costs_list[4][0], cell_format11)

number_e = [t4] * 12
for element in amazon:
    for col, row, month, index in months:
        val = number_e[index]
        if element[1] in month:
            worksheet.write_number(val, row, element[3], f_purple)
            worksheet.write_comment(val, row, element[5])
            number_e[index] = val + 1

t5 = max(number_e)

if t4 == t5:
    t5 += 1
    worksheet.write(t5, 1, costs_list[5][0], cell_format11)

else:
    worksheet.write(t5, 1, costs_list[5][0], cell_format11)

number_f = [t5] * 12
for element in sport:
    for col, row, month, index in months:
        val = number_f[index]
        if element[1] in month:
            worksheet.write_number(val, row, element[3], f_yellow)
            worksheet.write_comment(val, row, element[5])
            number_f[index] = val + 1

t6 = max(number_f)

if t5 == t6:
    t6 += 1
    worksheet.write(t6, 1, costs_list[6][0], cell_format11)

else:
    worksheet.write(t6, 1, costs_list[6][0], cell_format11)

number_g = [t6] * 12
for element in other:
    for col, row, month, index in months:
        val = number_g[index]
        if element[1] in month:
            worksheet.write_number(val, row, element[3], f_pinkred)
            worksheet.write_comment(val, row, element[5])
            number_g[index] = val + 1

t7 = max(number_g)

# Calculate the difference between income and costs using numpy
difference_per_month = np.array(expenses_per_month) - np.array(costs_per_month)

# Write sum up and difference per month in the worksheet
if t6 == t7:
    t7 += 1
    worksheet.write(t7 + 1, 0, "Calculation", cell_format14)
    worksheet.write(t7 + 2, 1, "Sum Costs", cell_format11)
    worksheet.write(t7 + 3, 1, "Difference", cell_format11)

    for col, row, month, index in months:
        worksheet.write_number(t7 + 2, row, costs_per_month[index], f_orange)

    for col, row, month, index in months:
        if difference_per_month[index] >= 0:
            worksheet.write_number(t7 + 3, row, difference_per_month[index], f_green)

        else:
            worksheet.write_number(t7 + 3, row, difference_per_month[index], f_red)

else:
    worksheet.write(t7 + 1, 0, "Calculation", cell_format14)
    worksheet.write(t7 + 2, 1, "Sum Costs", cell_format11)
    worksheet.write(t7 + 3, 1, "Difference", cell_format11)

    for i in expenses_per_month:
        for col, row, month, index in months:
            worksheet.write_number(t7 + 2, row, costs_per_month[index], f_orange)

    for i in difference_per_month:
        for col, row, month, index in months:
            if difference_per_month[index] >= 0:
                worksheet.write_number(t7 + 3, row, difference_per_month[index], f_green)

            else:
                worksheet.write_number(t7 + 3, row, difference_per_month[index], f_red)

# Print entries from database legible in terminal
for salary_month in list_salary:
    print("Salary in " + salary_month[1] + " was: " + str(salary_month[2]) + "€")

for cost_month in list_expenses:
    print("Costs for " + cost_month[4] + " in " + cost_month[1] + " was: " + str(cost_month[3]) + "€")

# Close workbook and write changes
workbook.close()

# Check if path to Excel is correct
if os.path.exists(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"):

    # Open the generated file with excel, python program will be ended, when the excel file is closed
    sp.call([r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", "Account-Book.xlsx"])

else:
    print("Please correct the path to your excel application, so that the file will be open automatically.")

print("Program is ended")
