# Excel VBA Projects

## Ascending-Descending Order

1. Take integer numbers displayed in sheet1 and put into 1-dimensional arrays
2. Use nested for loops to arrange the integer numbers into ascending/descending orders
3. Put results back to sheet1.

## First-Last Name

1. Take 10 1-dimensional arrays of people names, each one is first name and last name seperated by space/no space
2. Use nested for loops to rearrange the names by length, remove the names with no spaces and call a sub procedure that makes the last name first, space, then first name

## Import Data From SQL

1. Connect to SQL Server using ADODB connection
2. Store result set from SQL statement into ADODB recordset
3. Access each record in recordset and show results on a new worksheet

## Incremental Numbers

Generate incremental numbers by 1000 from 1000 to 5000 in sheet1 using for loops and 1-dimensional arrays

## Matching Letters

Take 10 created arrays of names and list all common letters between the names using nested for loops and 1-dimensional arrays

For example, Helen and Sara have nothing common but Alex and Sara share "a"
 
## Name VLookup

Use the ID in sheet1 to lookup for the corresponding name in sheet2 according to the asterisk position. 

For example, if the asterisk is located in the cell I30, the macro will only do a lookup against the data within ranges "A1" to "I30" in sheet2

## Pivot Table

Create a pivot table that has countries as columns, customer names as rows and total sales amount as values from sales data in "Data" sheet using nested for loops

## Postal Code

Generate distinct postal code in sheet1 using nested for loops

## Random Numbers

Generate 6 random numbers on Sheet1 using for loops and 2-dimensional arrays

## Read-Write Text File

1. Read each line of people name (first and last name seperated by space) from a text file into a 1-dimensional array
2. Rearrange the people names by length using for loops
3. Write the results back to another text file

## Remove Non Alpha Numeric

Function that removes all non alpha-numerical characters and digits if they are not at the start of the string
For example, "12 3nje dfBS<f    678jk;ji#*AAfnj" should return "123njedfBSfjkjiAAfnj"
