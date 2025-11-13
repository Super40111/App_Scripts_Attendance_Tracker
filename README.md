# App Scripts Attendance Tracker
This is a program that utilizes Google App Scripts, Google Sheets, and barcodes in order to create an attendance tracker.

# Part 1: The Project

This project was made to help a Kumon Center manager track which students were scheduled to attend each day of the week, whenever they are in class on not, and how long they have spent in class. 

# Part 2: How to Setup

Due to how Google App Scripts integration works, I cannot provide a xlsx file that already has the integration built in. Due to this, there are a few steps that need to be taken to replicate the functionality. Essentially, what needs to be done is intergate the script and setup the buttons.

The script needs to be copied into Google App Scripts by going into Extensions -> App Scripts. Afterwards, within the Google App Scripts tab, you need to make a New Deployment under a web app. If it says that the script needs Authorization, you need to procede with the underlined Advance option as progressing, otherwise Google App Scripts throws an error due to the script being Unauthorized. In addition, each button needs a specific script added to it in order for the button to work. For each button, click the three dots connected to it and choose the 'Assign Script' option. Then copy and paste in the following text based on the position of the button:

L29: GETSMALLESTUNUSEDID <br/>
L32: CLEANUPINPUTS
L42: FIXGAPS
L47: ROWCOUNT
L50: BINDERCOUNT
M5: ADDBARCODE
M9: REMOVEBARCODE
M13: ENDDAY
M25: ADDSTUDENT
M29: GETSMALLESTUNUSEDBARCODE
M32: BACKUPDATA
M38: SEARCHSTUDENT
O5: UPDATENAME
O11: UPDATEID
O17: UPDATESUBJECTCOUNT
O27: ADJUSTSTUDENTSCHEDULE
O32: REMOVESTUDENT
O38: SORTSTUDENTS
Q5: IMPORTBACKUPDATA
Q8: FIXSTARTUP
Q30: UPDATEINFOGRAPHIC
Q38: UPDATEATTENDANCETIME

# Part 3: Using the Spreadsheet

The spreadsheet was made with the primary way of checking students in and out using a barcode scanner. How barcode scanners work with Google Sheets is that when a connected barcode scanner scans a barcode, it copies the barcode data into the selected cell, the shifts the cell down one cell as if the 'Enter' Key is pressed. When scanning barcodes, the 'C1' cell should be highlighted. When data is entered and the cell shifts to 'C2', the script will reset the cell to 'C1'. There is a small delay to this however, so it is recommended to wait a few seconds before entering data again.


