My Shopping List application repository.

This is my GitHub repository for a shopping list app written in Python.

My first version "Shopping-List-python.py" is exclusively written in python and uses openpyxl to store recipe information in an excel file.
"ACCESS db Shopping-List.py" uses pyodbc to connect with Access and uses SQL quieries to interact with the database.
The app allows selection of multiple recipes and generates a shopping list of the required ingredients and quantities.
New recipes can be added and deleted from Excel/Access.

To use the excel version, update the global file-path at the beginning of the code and save an excel file as "Recipes.xlsx" in the same file path.

To use the Access version, update the global file-path and conn_str at the beginning of the code and copy the Access database into the same file path.

Features:
-Selection of recipes and creation of a shopping list .txt file.
-Addition and deletion of recipes.


This is an ongoing project with plenty of scope for additions and improvements.
Future Features:
-unit testing
-Editing of stored recipes
-Randomly generated recipe selections
-Connection to online recipe databases
-Importing of new recipes from URLs 
