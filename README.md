# App Scripts Attendance Tracker
This is a program that utilizes Google App Scripts, Google Sheets, and barcodes in order to create an attendance tracker.

# Part 1: The Project

This project was made to help a Kumon Center manager track which students were scheduled to attend each day of the week, whenever they are in class on not, and how long they have spent in class. The script operates by using a barcode scanner attached to a device with the spreadsheet open to scan barcodes attached to the folders of students whenever they check in and check out. While checked in, the background color of the data will change based on how long the student The script was designed based onthe schedule of the Kumon center in mind, which lead to certain design choices in the script. For example, the script is designed for students to attend on only Mondays, Tuesdays, Thursdays, and Saturdays.

# Part 2: How to Setup

Due to how Google App Scripts integration works, I cannot provide a xlsx file that already has the integration built in. Due to this, there are a few steps that need to be taken to replicate the functionality. Essentially, what needs to be done is intergate the script and setup the buttons.

The script needs to be copied into Google App Scripts by going into Extensions -> App Scripts. Afterwards, within the Google App Scripts tab, you need to make a New Deployment under a web app. If it says that the script needs Authorization, you need to procede with the underlined Advance option as progressing, otherwise Google App Scripts throws an error due to the script being Unauthorized. In addition, each button needs a specific script added to it in order for the button to work. For each button, click the three dots connected to it and choose the 'Assign Script' option. Then copy and paste in the following text based on the position of the button:

L29: GETSMALLESTUNUSEDID <br/>
L32: CLEANUPINPUTS <br/>
L42: FIXGAPS <br/>
L47: ROWCOUNT <br/>
L50: BINDERCOUNT <br/>
M5: ADDBARCODE <br/>
M9: REMOVEBARCODE <br/>
M13: ENDDAY <br/>
M25: ADDSTUDENT <br/>
M29: GETSMALLESTUNUSEDBARCODE <br/>
M32: BACKUPDATA <br/>
M38: SEARCHSTUDENT <br/>
O5: UPDATENAME <br/>
O11: UPDATEID <br/>
O17: UPDATESUBJECTCOUNT <br/>
O27: ADJUSTSTUDENTSCHEDULE <br/>
O32: REMOVESTUDENT <br/>
O38: SORTSTUDENTS <br/>
Q5: IMPORTBACKUPDATA <br/>
Q8: FIXSTARTUP <br/>
Q30: UPDATEINFOGRAPHIC <br/>
Q38: UPDATEATTENDANCETIME

# Part 3: Using the Spreadsheet

The spreadsheet was made with the primary way of checking students in and out using a barcode scanner. How barcode scanners work with Google Sheets is that when a connected barcode scanner scans a barcode, it copies the barcode data into the selected cell, the shifts the cell down one cell as if the 'Enter' Key is pressed. When scanning barcodes, the 'C1' cell should be highlighted. When data is entered and the cell shifts to 'C2', the script will reset the cell to 'C1'. There is a small delay to this however, so it is recommended to wait a few seconds before entering data again. The right side of the form is for editing the data via scripts instead of editing it directly. For each command, you enter the required data for the command before hitting the button below to run the associated function. When editing the forms on the side for editing data, any cell with a red background only cares if there is any text present and what exactly is inserted in does not matter. Any cells marked in orange text are not meant to be edited, and will automatically be filled out by running certain commands. These commands are usually right below these orange cells. If you need a site to get barcodes from, I used the site link in the bottom right corner to get mine.


