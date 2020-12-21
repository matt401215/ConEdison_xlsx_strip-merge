## ConEdison_xlsx_strip-merge

The purpose of this code is to format a large number of Con Edison utility Summary Statements (shown below), and combine them into a single xlsx file for processing. This was accomplished using two codes, xlsx_Read and MergeExcel.


![Example Summary Statement](https://github.com/matt401215/ConEdison_xlsx_strip-merge/blob/main/images/sumStatementSample.png)

# xlsx_Read_v3.py

For processing utility Summary Statements as formatted by Con Edison Energy Company. Returns the information in the main table, as well as the billing address and the account number from the heading. The purpose of this code is to get the data in an easy to use format for future processing.
The Summary Statements come in 4 types depending on what columns are included. 

  - Type 1: 11 columns, no ESCO
  - Type 2: 12 columns, electric ESCO
  - Type 3: 12 columns, gas ESCO
  - Type 4: 13 columns, electric and gas ESCO

# MergeExcel.py

For combining the formated data from returned from xlsx_Read_v3.py into a single xlsx file.
