# Combine2CodeGeneric
This is a simple process to combine 2 csv files and export them to Excel. 
App.config contains the folder location. 
Asumptions: 
	-The files contain the words 'sub_csv' or 'master_csv' in ReadFilesToList() class.
	-No validation for multiple files 
	-Add fields to the Record class and add the binding in the classes below
	-Add Column headings for the data table in RecordListToDataTable() and StringListToDataTable()
Additions in the future:
	-SQL 2 SQL tables
	-Csv 2 SQL tables
	-SQL to Blockchain :)
Side notes:
By no means is this perfect and these functions have been written a millions times.
However it's a good 'starter for 10' for me and anyone else who combines lists
	-I know for a fact theres a generic way of adding tables to excel, but this explains the process well for learning.
