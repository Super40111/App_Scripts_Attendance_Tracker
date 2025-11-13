# App Scripts Attendance Tracker
This is a program that utilizes Google App Scripts, Google Sheets, and barcodes in order to create an attendance tracker.

# Part 1: The Project

This project was made to help a Kumon Center manager track which students were scheduled to attend each day of the week, whenever they are in class on not, and how long they have spent in class. 

# Part 2: How to Setup

Due to how Google App Scripts integration works, I cannot provide a xlsx file that already has the integration built in. Due to this, there are a few steps that need to be taken to replicate the functionality. Essentially, what needs to be done is intergate the script and setup the buttons.

# Part 3: Using the Spreadsheet

The spreadsheet was made with the primary way of checking students in and out using a barcode scanner. How barcode scanners work with Google Sheets is that when a connected barcode scanner scans a barcode, it copies the barcode data into the selected cell, the shifts the cell down one cell as if the 'Enter' Key is pressed. When scanning barcodes, the 'C1' cell should be highlighted. When data is entered and the cell shifts to 'C2', the script will reset the cell to 'C1'. There is a small delay to this however, so it is recommended to wait a few seconds before entering data again.


