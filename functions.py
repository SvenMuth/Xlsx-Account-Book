import sqlite3


def show_classes(chosen_month_number):

    # At first let the user chose a month to make entries
    if chosen_month_number == "":
        print("At first a month has to be chosen:")
        print()
        select_class = "3"
        return select_class

    # Shows the menu and return the select class
    else:
        # Print the main menu
        print("-----------Menu-----------")
        print("Salary [1]")
        print("Costs [2]")
        print("Change month [3]")
        print("Delete all entries for the chosen month [4]")
        print("Delete all entries in Database [5]")
        print("Exit and write to Worksheet [0]")
        print()
        select_class = str(input("Please chose a class: [1, 2, 3, 4, 5, 0] "))
        return select_class


# Database to save all changes
def database():
    # Connect to database and create cursor
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Create a table if it not already exists
    c.execute("""CREATE TABLE IF NOT EXISTS data (
            type TEXT,
            month TEXT,
            salary INTEGER,
            expenses INTEGER,
            class_expenses TEXT,
            commentary TEXT
        )""")

    # Push in database an close connection
    conn.commit()
    conn.close()


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


def get_data_sql():
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Select all data from database
    c.execute("SELECT * FROM data")
    cdata = c.fetchall()

    conn.close()

    return cdata


def get_salary_month():

    # Create set to store month name, when a salary was written to
    set_salary = set()

    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Select all entries which belong to salary
    c.execute("SELECT month FROM data WHERE type='1'")
    ddata = c.fetchall()

    # Add the entries to the set
    for element in ddata:
        set_salary.add(element[0])

    conn.close()

    return set_salary


def update_salary_entry(income_month, chosen_month_name):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Update the old salary entry
    sql_update_query = "UPDATE data SET salary=? WHERE month=?"
    c.execute(sql_update_query, (income_month, chosen_month_name))

    conn.commit()
    conn.close()


def delete_entry(chosen_month_name):
    conn = sqlite3.connect(r"database\database.db")
    c = conn.cursor()

    # Update the old salary entry
    sql_delete_query = "DELETE from data WHERE month=?"
    c.execute(sql_delete_query, (chosen_month_name, ))

    conn.commit()
    conn.close()
