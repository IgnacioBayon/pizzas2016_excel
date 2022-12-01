# PIZZAS 2016

### INTRODUCTION

This code is an ETL which predicts the ingredients that a certain pizzeria
has to buy in a specific week based on various DataFrames from the orders
made in the year 2016. We have to manually clean the datasets using re and pd

Disclaimer: some variables and comments are in spanish as it is my native tongue

### 1. Libraries used
- 'pandas': used for managing the DataFrame
- 're': used for splitting strings and detecting re patterns
- 'matplotlib.pyplot': drawing plots
- 'seaborn': drawing plots
- 'datetime': used for formatting times
- 'warnings': to ignore certain warnings concerning data types
- 'openpyxl': for dealing with xlsx files

### 2. ETL
#### 1) Extract
We simply read the csvs with the names on the list 'file_names' with the 'read_csv' function from pandas, controlling whether the separation is ',' or ';'

#### 2) Transform
We have created different functions to make the Transform function more understandable.
- a) Clean Data
We first clean the data, as the dataframe is not entirely usable. We drop some rows and columns which are either not needed or do not provide any information. We then proceed to clean the dataframe with regex

- b) Get Pizzas Year
We get the pizzas ordered throughout the whole year

- c) Get Pizzas Weeks
We get the pizzas ordered by weeks and weekdays

We return all the dataframes

#### 3) Load
We have also separated this function into different smaller functions:

- a) Load Data Excel
This Function uses the pandas function 'ExcelWriter', aswell as the 'to_excel', to write the DataFrames on the different sheets from the xlsx file. 

- b) BarChart Excel
It plots a bar chart using 'openpyxl'. it receives as arguments the sheet it has to plot it on (as well as where the data is), a dictionary called 'chart_data', which contains the possible descriptive data necessary (Title, Axis' Titles, Height, Width and Cell) and a boolean which determines the orientation of the bar chart

- c) PieChart Excel
In the same way as the BarChart Function, it generates a pie chart

- d) LineChart Excel
In the same way as the BarChart Function, it generates a line chart

- e) Load
We finally call all the different functions to create the final Excel
