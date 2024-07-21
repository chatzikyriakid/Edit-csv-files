# Edit-csv-files
## Overview
I created this script in order to format and edit CSV files. The output file is of .xlsx format. Some parts of the program are hard-coded because i use the script for editing files related to my job.

## Code explanation
The program does the following:

- Removes unnecessary columns
- Makes the column headers bold
- Adds borders with thin lines to the whole sheet
- Filters the file based on column **Record Date**. It uses the datetime.now() function to obtain the current system date and filter all columns based on current date

In future version i plan to add different background colors to different cells

