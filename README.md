# Xlsx-Account-Book
Just a fun project to sort expenses per month in a excel sheet via python (It could contain bugs which i haven't found yet)

In order to learn the fundaments of python i started writing this program. It could be seen as my first "real" program. 
Maybe the job could be done easier with just Excels own tools. But i thinks its also a cool way to keep track on expenses and income per month.


The entries will be sorted to the right month. Also it will be checked, that only one entry for salary per month exist.
For the expenses a commentar and the time + date will be added.
![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/excel.PNG?raw=true)

At first step a month has to been choosen. After this it is possible to change the month.
All entries will be saved in a sqlite database.
![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/months.PNG?raw=true)

In the menue you can select between: Salary, Expenses, Change month and delete the database. 
![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/menue.PNG?raw=true)

At the end of the program, all entries will be shown. Also the workbook will be created. 
![alt text](https://github.com/SvenMuth/Xlsx-Account-Book/blob/main/pictures/changes.PNG?raw=true)

I will work on this project to include more features and remove bugs if the occur.
I also wrote a test program, to write entries to the database. 
The menue will be skipped, when the test program is active.
