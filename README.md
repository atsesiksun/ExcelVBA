# Excel VBA Projects

## Ascending-Descending Order

1. Take integer numbers displayed in sheet1 and put into 1-dimensional array
2. Use nested (FOR) loops to arrange the integer numbers into ascending/descending orders
3. Put results back to sheet1

## First-Last Name

1. Take 1-dimensional array of 10 people names, each one is first name and last name either seperated by a space or not seperated
2. Use nested (FOR) loops to rearrange the names by length, remove the names with no spaces and call a sub procedure that shows last name first, then space, then first name

## Import Data From SQL

1. Connect to SQL Server using ADODB connection
2. Store result set from SQL statement in ADODB recordset
3. Access each record in recordset and place results on a new worksheet

## Incremental Numbers

Generate incremental numbers by 1000 from 1000 to 5000 in sheet1 using for loops and 1-dimensional array

## Matching Letters

Take 1-dimentional array of 10 people names and list all common letters between the names using nested (FOR) loops

For example, names Helen and Sara have no common letters but Alex and Sara share "a"
 
## Name VLookup

Use the ID in sheet1 to lookup for corresponding name in sheet2 according to the asterisk position. 

For example, if asterisk is located in cell I30, the macro will only do a lookup against the data within ranges "A1" to "I30" in sheet2

## Pivot Table

Create a pivot table that has countries as columns, customer names as rows and total sales amount as values from sales data in "Data" sheet using nested (FOR) loops

## Postal Code

Generate distinct postal code in sheet1 using randomized numbers and nested (FOR) loops

## Random Numbers

Generate 6 random numbers on Sheet1 using (FOR) loops and 2-dimensional array

## Read-Write Text File

1. Read each line of a text file that contains list of first and last name seperated by space into a 1-dimensional array
2. Rearrange people names by length using (FOR) loops
3. Write results to another text file

## Remove Non Alphanumeric Characters

Function that removes all non alphanumeric characters and digits if they are not at the start of the string

For example, "12 3nje dfBS<f    678jk;ji#*AAfnj" should return "123njedfBSfjkjiAAfnj"
