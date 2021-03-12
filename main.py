import xlsxwriter
import os
import numpy as np
import subprocess as sp

from terminaltables import AsciiTable
from datetime import datetime
from pyfiglet import Figlet
from decimal import Decimal
from functions import *


# Function to randomly write a salary per months and write entries in different cost types with different values
# For testing remove hashtag, also remove the marked hashtags below
# from testing import test
# repeat_test = 0

# Initial new database
database()

# Create list for different months and position in Xlsx-Document [col, row, month, index]
months = [[0, 2, "January", 0], [0, 3, "February", 1], [0, 4, "March", 2], [0, 5, "April", 3],
          [0, 6, "May", 4], [0, 7, "June", 5], [0, 8, "July", 6], [0, 9, "August", 7],
          [0, 10, "September", 8], [0, 11, "October", 9], [0, 12, "November", 10], [0, 13, "December", 11]]

# Create list for categories
costs_list = ["Rent", "Credit", "Car", "Foods", "Amazon", "Sport", "Other"]

# Get category entries from database
category_sql = get_category_sql()

# Check if there any entries, otherwise write the standard to database
if not category_sql:
    for cost in costs_list:
        datasql = (0, "0", "0", "0", cost, "0")
        insert_sql(datasql)
else:
    costs_list.clear()
    for category in category_sql:
        costs_list.append(category[0])

# Create variables which are later needed
chosen_month_name, replace_entry = "", ""
chosen_month_number = 0

# Get the ID's of the different entries
id_data = get_id_sql()
id_item = 0

# Set id_item to the highest ID value in database
for value in id_data:
    if value[0] > id_item:
        id_item = value[0]

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

# Font for the comment box
c_white = {'color': '#FFFFFF'}

# For testing remove hashtags
# test(months, costs_list, repeat_test, id_item)
# """

# Print account book via figlet plugin
format_figlet = Figlet(font="slant")
print(format_figlet.renderText('Account-Book'))

# Chose class
select_class = show_classes(chosen_month_number)

# Loop for main program
while True:
    # Create Table if not already exist
    database()

    # Write salary
    if select_class == "1":
        # Collect the month which already have been written in Database
        set_salary = get_salary_month_sql()

        # Check if set with months is empty --> first entry for salary
        if len(set_salary) == 0:
            income_month = str(input("Please insert salary for " + chosen_month_name + ":"))
            income_month = check_number_input(income_month)

            if income_month != "0":
                # Write information's about your salary into database
                id_item += 1
                datasql = (id_item, chosen_month_name, income_month, "0", "0", "0")
                insert_sql(datasql)
                print("Salary was updated\n")

        else:
            # Check if there is already an entry for the chosen month
            if chosen_month_name in set_salary:

                replace_entry = str(input("\nShould the entry for " + chosen_month_name + " be replaced? [y, n]"))
                if replace_entry == "y":

                    # Set a new income for this month
                    income_month = str(input("Please insert new Salary for " + chosen_month_name + ":"))
                    income_month = check_number_input(income_month)

                    if income_month != "0":
                        # Update old entry
                        update_salary_sql(income_month, chosen_month_name)
                        print("Salary was updated\n")

                # Let the old entry and jump in main loop
                elif replace_entry == "n":
                    print("Salary wasn't updated\n")

                else:
                    print("Input invalid\n")

            # No entry found so just write a new entry for this month
            else:
                income_month = str(input("Please insert salary for " + chosen_month_name + ": "))
                income_month = check_number_input(income_month)

                if income_month != "0":
                    # Write information's about your salary into database
                    id_item += 1
                    datasql = (id_item, chosen_month_name, income_month, "0", "0", "0")
                    insert_sql(datasql)

        select_class = show_classes(chosen_month_number)

    # Write expenses
    elif select_class == "2":

        # List cost categories
        entry = 1
        for cost in costs_list:
            print(str(entry) + ". " + cost)
            entry += 1

        print("\nYou have chosen " + chosen_month_name)

        try:
            category_chosen = int(input("Please chose a category: [1 - " + str(len(costs_list)) + "] "))
            if category_chosen == 0:
                print("Invalid input\n")

            else:
                # Get the category
                category_expanses = costs_list[category_chosen - 1]
                expanses_costs = str(input("Please input the costs for " + category_expanses + ": "))
                expanses_costs = check_number_input(expanses_costs)

                if expanses_costs != "0":
                    id_item += 1

                    # Write a commentary to your entry and append the right time of day and date
                    commentary = str(input("Please write a commentary to the expense: "))
                    print()
                    commentary += "\nCreated on:\n" + datetime.now().strftime("%d.%m.%y - %H:%M")

                    # Write information's about your costs into database
                    datasql = (id_item, chosen_month_name, "0", expanses_costs, category_expanses, commentary)
                    insert_sql(datasql)

        except IndexError:
            print("Input is out of range\n")

        except ValueError:
            print("Invalid input\n")

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

        try:
            chosen_month_number = int(input("\nPlease select a month: [1-12] "))
            if chosen_month_number == 0:
                print("Invalid Input\n")

            else:
                chosen_month_name = months[chosen_month_number - 1][2]
                print("You have chosen " + chosen_month_name + "\n")

        except ValueError:
            chosen_month_number = 0
            print("Invalid Input\n")

        except IndexError:
            chosen_month_number = 0
            print("Input is out of range\n")

        select_class = show_classes(chosen_month_number)

    # Option to add or delete a cost category
    elif select_class == "4":
        print("\nAdd a new category for expenses [1]")
        print("Delete a category (All entries for this category will be deleted) [2]")

        chosen_category = str(input("Please chose: [1, 2]"))
        if chosen_category == "1":
            new_category = str(input("\nPlease write a new category you want to add: "))
            new_category = new_category.strip()

            if new_category != "":
                # Check if category already exist and add to database
                if new_category not in costs_list:
                    datasql = (0, "0", "0", "0", new_category, "0")
                    insert_sql(datasql)
                    costs_list.append(new_category)

                    print("New category was saved.")

                else:
                    print("\nCategory already exists.")

            else:
                print("Invalid input\n")

        elif chosen_category == "2":

            i = 1
            for cost in costs_list:
                print(str(i) + ". " + cost)
                i += 1

            try:
                delete_category = int(input("\nPlease chose a category to delete: [1-" + str(len(costs_list)) + "]"))
                # Delete a category and all entries which it belongs to
                delete_category_sql(costs_list, delete_category)
                costs_list.remove(costs_list[delete_category - 1])
                print("Category [" + costs_list[delete_category - 1] + "] was successfully deleted.")

            except ValueError:
                print("Invalid input\n")

            except IndexError:
                print("Input is out of range\n")

        else:
            print("Invalid Input\n")

        select_class = show_classes(chosen_month_number)

    # Delete a specific entry in the database
    elif select_class == "5":
        print("Delete an entry from salary [1]")
        print("Delete an entry from costs [2]")

        # Displays the salary entries with an id
        choice = str(input("Please chose: [1, 2] "))
        if choice == "1":
            salary_data = get_salary_sql()

            # Check if there are any entries
            if len(salary_data) == 0:
                print("No entries for salary yet.\n")

            else:
                table_data = [['ID', 'Month', 'Salary']]

                for salary in salary_data:
                    var_digit = digit_correct(salary[2])
                    table_data.append([salary[0], salary[1], var_digit])

                table = AsciiTable(table_data)
                print(table.table)

                try:
                    select = int(input("Please enter the ID, of the entry you want to delete: "))

                    # Check if an entry with this ID exist
                    id_item = check_id_sql(select)

                    if id_item is None:
                        print("Entry does not exist.")

                    else:
                        # Delete the entry
                        delete_id_sql(select)
                        print("Entry was successfully deleted\n")

                except ValueError:
                    print("Invalid input\n")

        # Displays the cost entries with an id
        elif choice == "2":
            costs_data = get_costs_sql()

            # Check if there are any entries
            if len(costs_data) == 0:
                print("No entries for costs yet.\n")

            else:
                table_data = [['ID', 'Month', 'Costs', 'Category', 'Comment']]

                for cost in costs_list:
                    category_data = get_expense_by_category(cost)
                    for cost_sql in category_data:
                        var_digit = digit_correct(cost_sql[3])
                        table_data.append([cost_sql[0], cost_sql[1], var_digit, cost_sql[4], cost_sql[5]])

                table = AsciiTable(table_data)
                print(table.table)

                try:
                    select = int(input("Please enter the ID, of the entry you want to delete: "))
                    # Check if an entry with this ID exist
                    id_item = check_id_sql(select)
                    if id_item is None:
                        print("Entry does not exist.\n")

                    else:
                        # Delete the entry
                        delete_id_sql(select)
                        print("Entry was successfully deleted\n")

                except ValueError:
                    print("Invalid input\n")

        else:
            print("Invalid Input\n")

        select_class = show_classes(chosen_month_number)

    # Delete all entries from the chosen month
    elif select_class == "6":
        input_delete = str(input("You are sure to delete all entries for " + chosen_month_name + " [y, n]"))

        if input_delete == "y":
            delete_entry_sql(chosen_month_name)
            print("All entries for " + chosen_month_name + " were deleted\n")

        elif input_delete == "n":
            print("Operation was successfully aborted\n")

        else:
            print("Invalid Input\n")

        select_class = show_classes(chosen_month_number)
        
    # Delete database
    elif select_class == "7":
        reset = str(input("Delete the previous database? [y/n]"))
        if reset == "y":
            # Delete file
            os.remove(r"database\database.db")
            print("Database was removed\n")

            # Create a new table
            database()

            reset_categories = str(input("Do you want to reset the categories too? [y, n] "))
            while True:

                if reset_categories == "y":

                    # Write standard categories to database
                    costs_list = ["Rent", "Credit", "Car", "Foods", "Amazon", "Sport", "Other"]
                    for cost in costs_list:
                        datasql = (0, "0", "0", "0", cost, "0")
                        insert_sql(datasql)

                    print("Operation successfully\n")
                    reset_categories = "0"

                elif reset_categories == "n":
                    for cost in costs_list:
                        datasql = (0, "0", "0", "0", cost, "0")
                        insert_sql(datasql)

                    print("Operation successfully\n")
                    reset_categories = "0"

                elif reset_categories == "0":
                    break

                else:
                    print("Invalid Input")

        elif reset == "n":
            print("Database wasn't deleted")

        else:
            print("invalid Input")

        select_class = show_classes(chosen_month_number)

    elif select_class == "0":
        break

    else:
        print("Invalid Input")
        #  Change between classes
        select_class = show_classes(chosen_month_number)

# For testing remove hashtag
# """

# Write months in worksheet
for col, row, month, index in months:
    worksheet.write(col, row, month, cell_format14)

worksheet.write(1, 0, "Income", cell_format14)
worksheet.write(2, 1, "Salary", cell_format11)
worksheet.write(4, 0, "Costs", cell_format14)

# Collect the sum of the costs per months
costs_per_month = [0] * 12
cost_data = get_costs_sql()

for cost in cost_data:
    for col, row, month, index in months:
        if cost[1] == month:
            costs_per_month[index] += Decimal(cost[3])
            

# Write salary per month in Worksheet + sort salary per months in list
expenses_per_month = [0] * 12
salary_data = get_salary_sql()

for salary_month in salary_data:
    for col, row, month, index in months:
        if salary_month[1] in month:
            expenses_per_month[index] = Decimal(salary_month[2])
            val = col + 2

            var_digit = digit_correct(salary_month[2])
            worksheet.write(val, row, var_digit, f_grye)


# Sort costs by months and write them in the right place in the Xlsx document
# List for the different colors, which where declared in the beginning
colorlist = [f_darkgreen, f_yellow, f_purple, f_darkred, f_deeppurple, f_pinkred, f_turquoise]

# Variables to check if the position has changed and when the first entry will be made
startposold = 5
startpos = 5
count = 0
# Get length of color list
length = len(colorlist) - 1

for cost in costs_list:
    # Change colors automatically and reset counter to 1, so that the colors begin from 1
    if count == 0:
        color = colorlist[length]
    elif count == length:
        color = colorlist[length]
        count = 1
    else:
        color = colorlist[count - 1]

    # Get the right entries to the cost category
    category_data = get_expense_by_category(cost)

    # Check if first entry
    if count == 0:

        worksheet.write(startpos, 1, cost, cell_format11)

        # List for position for the different months
        startposlist = [startpos] * 12
        for cost_sql in category_data:

            # Sort costs by months and write them in the right place in the Xlsx document
            for col, row, month, index in months:
                val = startposlist[index]
                if cost_sql[1] in month:

                    var_digit = digit_correct(cost_sql[3])
                    worksheet.write(val, row, var_digit, color)
                    worksheet.write_comment(val, row, cost_sql[5], c_white)

                    # Increase the position +1 at the right position in list
                    startposlist[index] = val + 1

                # Start position for next category --> start at the highest position
                startpos = max(startposlist)

    else:
        if startpos == startposold:
            startpos += 1
            worksheet.write(startpos, 1, cost, cell_format11)

        else:
            worksheet.write(startpos, 1, cost, cell_format11)

        # Check if there are any entries for the first category
        if startpos != 5:
            # Otherwise use the variable to check if there any entries
            startposold = startpos

        startposlist = [startpos] * 12

        for cost_sql in category_data:

            for col, row, month, index in months:
                val = startposlist[index]
                if cost_sql[1] in month:
                    var_digit = digit_correct(cost_sql[3])
                    worksheet.write(val, row, var_digit, color)
                    worksheet.write_comment(val, row, cost_sql[5], c_white)

                    startposlist[index] = val + 1

            startpos = max(startposlist)

    # Clear the position list and the list for the costs per category
    startposlist.clear()
    category_data.clear()

    count += 1

# Calculate the difference between income and costs using numpy
difference_per_month = np.subtract(np.array(expenses_per_month), np.array(costs_per_month))

# Write sum up and difference per month in the worksheet
if startpos == startposold:
    startpos += 1
    worksheet.write(startpos + 1, 0, "Calculation", cell_format14)
    worksheet.write(startpos + 2, 1, "Sum Costs", cell_format11)
    worksheet.write(startpos + 3, 1, "Difference", cell_format11)

    for col, row, month, index in months:
        var_digit = digit_correct(costs_per_month[index])
        worksheet.write(startpos + 2, row, var_digit, f_orange)

    # Check if difference between income and cost are positive or negative
    for col, row, month, index in months:
        if difference_per_month[index] >= 0:
            var_digit = digit_correct(difference_per_month[index])
            worksheet.write(startpos + 3, row, var_digit, f_green)

        else:
            var_digit = digit_correct(difference_per_month[index])
            worksheet.write(startpos + 3, row, var_digit, f_red)

else:
    worksheet.write(startpos + 1, 0, "Calculation", cell_format14)
    worksheet.write(startpos + 2, 1, "Sum Costs", cell_format11)
    worksheet.write(startpos + 3, 1, "Difference", cell_format11)

    for i in expenses_per_month:
        for col, row, month, index in months:
            var_digit = digit_correct(costs_per_month[index])
            worksheet.write(startpos + 2, row, var_digit, f_orange)

    for i in difference_per_month:
        for col, row, month, index in months:
            if difference_per_month[index] >= 0:
                var_digit = digit_correct(difference_per_month[index])
                worksheet.write(startpos + 3, row, var_digit, f_green)

            else:
                var_digit = digit_correct(difference_per_month[index])
                worksheet.write(startpos + 3, row, var_digit, f_red)

# Close workbook and write changes
workbook.close()

try:
    # Open the generated file with excel, python program will be ended, when the excel file is closed
    sp.call([r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", "Account-Book.xlsx"])

except FileNotFoundError:
    print("Path to your Excel application is wrong.")

print("Program is ended")