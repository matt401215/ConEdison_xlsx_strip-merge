## ConEdison_xlsx_strip-merge

The purpose of this code is to format a large number of Con Edison utility Summary Statements, and combine them into a single xlsx file for processing. This was accomplished using two codes, xlsx_Read and MergeExcel.

![Sample Summary Statement](ConEdison_xlsx_strip-merge/images/Screen Shot 2020-12-21 at 3.31.34 PM.png)

# xlsx_Read_v3.py

For processing utility Summary Statements as formatted by Con Edison Energy Company. Returns the information in the main table, as well as the billing address and the account number from the heading. The purpose of this code is to get the data in an easy to use format for future processing.


# MergeExcel.py

For combining the formated data from returned from xlsx_Read_v3.py into a single xlsx file.
