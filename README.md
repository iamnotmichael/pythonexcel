# pythonexcel
Python Code to Read and Write Excel

This program is used to copy specific categories of information from an Excel Phone Report emailed daily and write them to seperate reports. The form has many stanard fields but the row positions change every day.

Overall, the script cylces through all files in a specified folder, finds the date the form was created for, finds the row positions for the caterogies needed, copies and writes the information to the appropriate sheets, and tracks the dates copied so far.

Upcoming Changes:

+ Error checking for common issues that might accure.

+ Splitting up exceptions and functions to seperate py files for clearer code

+ Automating the VBA portion of the reports (in addition to Python code, there is a VBA script in Outlook that downloads the Excel attachments and another in Excel that converts the original binary type files to xlsx file time.
