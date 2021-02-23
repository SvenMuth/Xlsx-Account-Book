import random
import os
from functions import *


# Function to test the program
def test(months, costslist, repeat):
    list_of_months = []

    # Delete Database
    if os.path.exists(r"database\database.db"):
        # Delete file
        os.remove(r"database\database.db")

    # Initial new database
    database()

    # Create for every single month exactly one entry for a salary and write them to database
    for col, row, month, index in months:
        list_of_months.append(month)
        income = random.randint(2500, 3000)
        datasql = ("1", month, income, 0, "0", "0")
        insert(datasql)

    # Create a numerous variety of expanses and write them to database
    i = 0
    while i < repeat:
        chosen_month_name = list_of_months[random.randint(0, 11)]
        category_expanses = costslist[random.randint(0, 6)][0]
        expanses_costs = random.randint(0, 500)

        datasql = ("2", chosen_month_name, 0, expanses_costs, category_expanses, "X")
        insert(datasql)
        i += 1






