ExceleratedSQL
==============

This is a program that will move Excel data into a SQL Database. 

##Problem Statement:
"Data needs to be moved from a spreadsheet to a database automatically. Write a program that allows the user to select which .xls or .xlsx file to import, grab the data from the spreadsheet, and send the data to a stored procedure on a SQL Server. You are allowed to use whatever tools/methods you can find on the internet to assist in the project. Use whatever language/tool set you are comfortable with but you must tell me why a tool was used."
    
##Features:
- No SQL Injection! Sanitized inputs! (I think...)
- Easy to run, flawless execution.
- Lots of error checking. Some of the errors are kind of vague though. 
- Can account for blank rows or not enough/too many columns!

##In Devlopment:
- Better error reporting
- More useful comments!
- A GUI?!

##Bugs
- If you have a formatting error (garbage data or blank cell) the entire program will stop. It currently cannot move on and attempt more inserts after an error.

- Sometimes extra Excel processes are left running after the programs completion. This is likely caused by errors that exit the program without properly closing connections and quitting Excel. Needs investigation. 
