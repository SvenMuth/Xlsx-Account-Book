import random
from functions import *


# Function to test the program
def test(months, costs_list, repeat, id_item):
    list_of_months = []

    salary_data = get_salary_sql()

    for col, row, month, index in months:
        list_of_months.append(month)
        # Check if there are any entries for salary per month
        if len(salary_data) == 0:
            income = random.randint(2500, 3000)
            id_item += 1
            datasql = (id_item, month, income, 0, "0", "0")
            insert_sql(datasql)

# Create a numerous variety of expanses and write them to database
    i = 0
    length = len(costs_list) - 1
    while i < repeat:
        chosen_month_name = list_of_months[random.randint(0, 11)]
        category_expanses = costs_list[random.randint(0, length)]
        expanses_costs = random.randint(0, 500)
        id_item += 1
        datasql = (id_item, chosen_month_name, 0, expanses_costs, category_expanses, "test")
        insert_sql(datasql)
        i += 1