# read_excel
Windows app to find time &amp; expense info for employees in a given invoice

### Company owner asked that this application:
* Adds information to a tab in the Excel workbook based invoice explaining where each item was found/wasn't found.
* Copy pertinent pages from backup pdfs and include them in ts/exp directories with the invoice
* Provide a loader that allows users to see how far along their process is

_it does all 3._

## How it works:

![A picture of the loader](http://res.cloudinary.com/ralst0n/image/upload/v1510071837/loading_bar_hrnodj.jpg)

A vba script calls the exe passing it the necessary arguments to find out which information it should seek.
the program navigates to the timesheet or expense sheet backup folders then using the "period ending dates" on the excel sheet, determines which files it needs to search for a given person within.
It uses iTextSharp to match the information it knows to a pdf page then prints out the matching information to the excel sheet. or leaving a message that it could not be found.

_note: double clicking the Prudent Engineering Logo brings up the about information for the program_
![a picture of fully loading message saying it was reopening the excel workbook](http://res.cloudinary.com/ralst0n/image/upload/v1510071836/fully_loaded_iuw1j6.jpg)

upon completion the Excel workbook is reopened with each employees backup data printed in excel and also saved as individual pdfs named for the employee and pay period

![completed search of timesheets](http://res.cloudinary.com/ralst0n/image/upload/v1510071837/results_ldvsia.jpg)
