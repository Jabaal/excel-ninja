# Description
Excel VBA
Choose a table, filter and display it and save the filtered data to another workbook.

More or less my first project, please be patient.

# current state
- reads sheetnames(tables) via ADO SQL into a combobox from the active workbook
- choose a predefined Filter in a second combobox
- displays filtered data with additional
- copy filtered data to a temporary sheet
- moves this "shadow" sheet to a predefined workbook and names this sheet like the original sheet

# To do/wishlist
- read sheetnames from a workbook you can choose yourself
- make your own Filter and save them for later use
- chose your own save destination
- query a different workbook via SQL
- speed the process up 

# Known issues
- cannot save after using update dropdowns
- cannot rename the sheet, if the name already exists
- slow
