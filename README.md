# Xlsx-Account-Book
Just a fun project to sort expenses per month in an excel sheet via python.
In order to learn the fundamentals of python I started writing this program. It could be seen as my first proper project. 

The account book will write the entries from the user to a xlsx-sheet. 
All entries will be saved in a SQLite database. 
In the terminal new entries can be created, the month can be changed during the process. 
Further option is to delete specific entries, all entries for a month or the whole database.
The entries from the user will be sorted to the right month. Also, it will be checked, that only one entry for salary per month exist.
For the expenses a commentary and the time + date will be added.
For the input of the money for salary and costs are many formats accepted.
For example: 23, 23., 23.0 or 23.00 are valid formats. The input will be formatted automatically and an euro symbol will be added.

![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/excel.PNG?raw=true)

At first step a month has to been chosen. After this it is possible to change the month.

![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/months.PNG?raw=true)

In the menu you can select between different operations.

![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/menu.PNG?raw=true)

Also an option was added to delete a specific entry.

![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/delete_entry.PNG?raw=true)

After leaving the menu, the workbook will be created. 

I will work on this project to include more features and remove bugs if the occur.

