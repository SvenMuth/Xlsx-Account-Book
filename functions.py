import sqlite3


def show_classes(chosen_month_number):

    # At first let the user chose a month to make entries
    if chosen_month_number == 0:
        print("At first a month has to be chosen:")
        print()
        select_class = "3"
        return select_class

    # Shows the menu and return the select class
    else:
        # Print the main menu
        print("-----------------Menu-----------------")
        print("Write salary [1]")
        print("Write costs [2]")
        print("Change month for entries [3]")
        print("Edit expense categories [4]")
        print("Delete specific entry [5]")
        print("Delete all entries for this month [6]")
        print("Delete all entries in Database [7]")
        print("Exit and write to Worksheet [0]\n")

        select_class = str(input("Please chose a class: [1, 2, 3, 4, 5, 6, 7, 0] "))
        print()

        return select_class


# Database to save all changes
def database():
    # Connect to database and create cursor
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Create a table if it not already exists
    c.execute("""CREATE TABLE IF NOT EXISTS data (
            id_item INTEGER,
            month TEXT,
            salary INTEGER,
            expenses INTEGER,
            class_expenses TEXT,
            commentary TEXT
        )""")

    # Push in database an close connection
    conn.commit()
    conn.close()


# Insert data to database
def insert_sql(datasql):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Insert variables from main.py when function is called
    c.execute("""INSERT INTO data 
                    VALUES (?,?,?,?,?,?)
                    """,
              datasql)

    conn.commit()
    conn.close()


# Only get the ID of the different entries
def get_id_sql():
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Select all data from database
    c.execute("SELECT id_item FROM data")
    id_data = c.fetchall()

    conn.close()

    return id_data


# Check if an ID exists
def check_id_sql(select):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Update the old salary entry
    sql_query = "SELECT * FROM data WHERE id_item=?"
    c.execute(sql_query, (select, ))
    id_item = c.fetchone()

    conn.close()
    return id_item


# Get the different categories, which stored
def get_category_sql():
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Select all data from database
    c.execute("SELECT class_expenses FROM data WHERE expenses=0 AND salary=0")
    category_sql = c.fetchall()

    conn.close()

    category_sql = set(category_sql)
    return category_sql


# Get all cost entries
def get_costs_sql():
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Select all data from database
    c.execute("SELECT * FROM data WHERE expenses > 0")
    costs_data = c.fetchall()

    conn.close()

    return costs_data


# Get all salary entries
def get_salary_sql():

    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Select all entries which belong to salary
    c.execute("SELECT * FROM data WHERE salary > 0")
    salary_data = c.fetchall()

    conn.close()

    return salary_data


# Get the different month which got an salary entry
def get_salary_month_sql():

    # Create set to store month name, when a salary was written to
    set_salary = set()

    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Select all entries which belong to salary
    c.execute("SELECT month FROM data WHERE salary > 0")
    months_used = c.fetchall()

    # Add the entries to the set
    for month in months_used:
        set_salary.add(month[0])

    conn.close()

    return set_salary


# Update a salary for a month
def update_salary_sql(income_month, chosen_month_name):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Update the old salary entry
    sql_update_query = "UPDATE data SET salary=? WHERE month=?"
    c.execute(sql_update_query, (income_month, chosen_month_name))

    conn.commit()
    conn.close()


# Get expenses by category to the element in cost list
def get_expense_by_category(cost):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()
    sql_query = "SELECT * FROM data WHERE class_expenses=? AND expenses > 0"
    c.execute(sql_query, (cost,))
    category_data = c.fetchall()
    conn.commit()
    conn.close()

    return category_data


# Delete entries for a specific month
def delete_entry_sql(chosen_month_name):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Update the old salary entry
    sql_delete_query = "DELETE FROM data WHERE month=?"
    c.execute(sql_delete_query, (chosen_month_name, ))

    conn.commit()
    conn.close()


# Delete a entry for a specific ID
def delete_id_sql(select):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    sql_delete_query = "DELETE FROM data WHERE id_item=?"
    c.execute(sql_delete_query, (select, ))

    conn.commit()
    conn.close()


# Remove a category
def delete_category_sql(costs_list, delete_category):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Update the old salary entry
    sql_delete_query = "DELETE FROM data WHERE class_expenses=?"
    category = costs_list[delete_category - 1]
    c.execute(sql_delete_query, (category, ))

    conn.commit()
    conn.close()