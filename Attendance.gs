var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance_Sheet"); //Gets the sheet used to log attendance.
var maximumRows = parseInt(sheet.getRange('L46').getDisplayValue()) + 10; //The number of data entries present plus a few extra rows.
var maximumDataEntries = 597; //Maximum possible rows throughout the data set's use (Should be 2 * (Maximum number of possible students) - 3).
var minimumID = 1; //Minimum possible ID value.
var minimumBarcode = 1; //Minimum possible barcode value.
var clearRow = [["","","","","","","","",""]]; //Variable used to clear a row of data.
var jFunction = ["=IF(ISBLANK(A","),\"\",IF(I","=\"IN_CLASS\",$K$1-H",",I","-H","))"]; //The formula for the J Column Function (Without the row number).
var kFunction = ["=IF(ISBLANK(A","),\"\",IF(F","=2,H","+TIME(1,0,0),H","+TIME(0,30,0)))"]; //The formula for the K Column Function (Without the row number).


function onEdit(e) { //Checks if a barcode has been scanned to C1.
  if (sheet.getRange('C3').getDisplayValue() !== "Barcode ID") { //Prevents the 'C3' Cells from being accidentally updated while logging in/out students.
    sheet.getRange('C1').activate();
    sheet.getRange('C3').setValue("Barcode ID");
    sheet.getRange('C3').setUnderline(true);
  }
  if (e.range.getA1Notation() === 'C1' && e.range.getDisplayValue() !== "") { // Updates Attendance when C1 is edited.
    UPDATEATTENDANCE();
    e.range.activate(); //Sets the highlighted cell back to C1 (Since ENTER sets the highlighted cell to C2).
    e.range.setValue(""); //Wipes C1 after running.
    let displacedCell = e.range.offset(1,0);
    if (displacedCell.getDisplayValue() !== "") { //Runs in case two students signed in too quickly between each other.
      let displacedValue = displacedCell.getDisplayValue();
      displacedCell.setValue("");
      e.range.setValue(displacedValue);
      UPDATEATTENDANCE();
      e.range.setValue("");
    }
  }
  else if (e.range.getA1Notation() === 'A1' || e.range.getA1Notation() === 'A2' || e.range.getA1Notation() === 'D1' || e.range.getA1Notation() === 'D2') { //Prevents the A1, A2, D1, and D2 cells from being edited.
    e.range.setValue("HIGHLIGHT C1 BEFORE SCANNING!!!");
    sheet.getRange('C1').activate();
  }
  else if (e.range.getA1Notation() === 'B1' || e.range.getA1Notation() === 'B2' || e.range.getA1Notation() === 'E1' || e.range.getA1Notation() === 'E2') { //Prevents the B1, B2, E1, and E2 cells from being edited.
    e.range.setValue("");
    sheet.getRange('C1').activate();
  }
  if (!(e.range.getA1Notation() === 'C1' || e.range.getA1Notation() === 'C2' || e.range.getA1Notation() === 'Q29')) { //Updates the infographic.
    sheet.getRange('Q29').setValue(e.range.getA1Notation());
  }
}


function UPDATEATTENDANCE() { //Updates Attendance based on the scanned barcode.
  if (STARTUP()) {
    let studentIndex = sheet.getRange('A4'); //Gets the current student.
    let barcodeIndex = sheet.getRange('C4'); //Gets the first barcode cell.
    let barcodeValue = sheet.getRange('C1').getDisplayValue(); //Gets the entered barcode value.
    let getDay = sheet.getRange('I1').getDisplayValue(); //Gets the current day of the week.
    let alternateName = "";
    let alternateDay = ""; //Logs the correct day of the week if the student came in on an off-day.
    let alternateCheck = false; //True if the associated student came on an off-day.
    if (barcodeValue === "NOT_ASSIGNED") { //Checks that the input is an actual barcode.
      FUNCTIONSTATUS(false, false, "Error: Cannot Check Unassigned Barcodes");
      return;
    }
    for (let checkLimit = 0; checkLimit < maximumRows && studentIndex.getDisplayValue() !== ""; checkLimit++) {
      if ((studentIndex.getDisplayValue() === barcodeValue || barcodeIndex.getDisplayValue() === barcodeValue) && barcodeIndex.offset(0,2).getDisplayValue() === getDay){ //Checks for matching barcode & day of the week.
        SETTIME(studentIndex.getDisplayValue(), barcodeIndex, false); 
        return;
      } 
      else if ((studentIndex.getDisplayValue() === barcodeValue || barcodeIndex.getDisplayValue() === barcodeValue) &&  !alternateCheck) { //Logs a potential offday if there is a matching barcode, but the current day is an off-day.
        alternateCheck = true;
        alternateName = studentIndex.getDisplayValue();
        alternateDay = barcodeIndex;
      }
      studentIndex = studentIndex.offset(1,0);
      barcodeIndex = barcodeIndex.offset(1,0); //Checks the next data entry.
    }
    if (alternateCheck) { //Logs the student ONLY IF it is an off-day.
      SETTIME(alternateName, alternateDay, true); 
    }
    else { //Runs if no student is found.
      FUNCTIONSTATUS(false, false, "Error: Student Not Found");
    }
  }
}


function SETTIME(name, barcodeIndex, alternateDay) { //Updates the sign-in or sign-out times of a student.
  let getTime = sheet.getRange('K1').getDisplayValue(); //Gets the current time.
  let startTime = barcodeIndex.offset(0,5);
  let endTime = barcodeIndex.offset(0,6); //Gets the columms that need adjusting.
  let getDay = sheet.getRange('G1').getDisplayValue();
  let todayAttendance = sheet.getRange('Q11');
  let remainingStudents = sheet.getRange('Q12');
  let lastDayCheckin = sheet.getRange('Q15').getDisplayValue();
  let lastDayCheckout = sheet.getRange('Q16').getDisplayValue();
  if (getDay === lastDayCheckin || getDay === lastDayCheckout) { //Updates the Infographic based on the given information.
    if (endTime.getDisplayValue() !== "IN_CLASS") {
      todayAttendance.setValue(parseInt(todayAttendance.getDisplayValue()) + 1);
    }
    else if (!alternateDay && remainingStudents.getDisplayValue() > 0) {
      remainingStudents.setValue(parseInt(remainingStudents.getDisplayValue()) - 1);
    }
  }
  else {
    todayAttendance.setValue("1");
    remainingStudents.setValue(parseInt(GETATTENDANCEBYDAY(sheet.getRange('I1').getDisplayValue())) - 1);
  }
  if (endTime.getDisplayValue() === "IN_CLASS") { //Signs the student out.
    endTime.setValue(getTime); 
    if (alternateDay) { //Checks if the student signed in on a day they don't attend.
      FUNCTIONSTATUS(false, true, "Sign-Out Successful. NOTE: Incorrect Day");
    }
    else {
      FUNCTIONSTATUS(false, true, "Sign-Out Successful");
    }
    sheet.getRange('Q14').setValue(name);
    sheet.getRange('Q16').setValue(sheet.getRange('G1').getDisplayValue());
    sheet.getRange('Q18').setValue(getTime);
  }
  else { //Signs the student in.
    startTime.setValue(getTime); 
    endTime.setValue("IN_CLASS");
    if (alternateDay) { //Checks if the student signed in on a day they don't attend.
      FUNCTIONSTATUS(false, true, "Sign-In Successful. NOTE: Incorrect Day");
    }
    else {
      FUNCTIONSTATUS(false, true, "Sign-In Successful");
    }
    sheet.getRange('Q13').setValue(name);
    sheet.getRange('Q15').setValue(sheet.getRange('G1').getDisplayValue());
    sheet.getRange('Q17').setValue(getTime);
  }
}


function GETATTENDANCEBYDAY(day) { //Returns the current Infographic value of a specific day.
  if (day === "Monday") { //Returns the Monday Value.
    return sheet.getRange('Q20').getDisplayValue();
  }
  if (day === "Tuesday") { //Returns the Tuesday Value.
    return sheet.getRange('Q21').getDisplayValue();
  }
  if (day === "Thursday") { //Returns the Thursday Value.
    return sheet.getRange('Q22').getDisplayValue();
  }
  if (day === "Saturday") { //Returns the Saturday Value.
    return sheet.getRange('Q23').getDisplayValue();
  }
  return sheet.getRange('L44'.getDisplayValue());
}


function UPDATEATTENDANCETIME() { //Updates the current attendance time of a student in the data.
  if (STARTUP()) {
    let name = sheet.getRange('Q33').getDisplayValue();
    let id = sheet.getRange('Q34').getDisplayValue();
    let updateCheckIn = (sheet.getRange('Q35').getDisplayValue() !== "");
    let updateDay = sheet.getRange('Q36').getDisplayValue();
    let newTime = sheet.getRange('Q37').getDisplayValue();
    let studentIndex = sheet.getRange('A4');
    if (name === "" || id === "" || updateDay === "" || newTime === "") { //Checks for all of the mandatory inputs.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    if (!(updateDay === "Monday" || updateDay === "Tuesday" || updateDay === "Thursday" || updateDay === "Saturday")) { //Checks for valid days.
      FUNCTIONSTATUS(false, false, "Error: Invalid Day");
    }
    if (!CHECKVALIDTIME(newTime, updateDay)) { //Checks for a valid time.
      FUNCTIONSTATUS(false, false, "Error: Invalid Time");
    }
    sheet.getRange('Q35').setValue("");
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      if (studentIndex.getDisplayValue() === name && studentIndex.offset(0,1).getDisplayValue() === id && studentIndex.offset(0,4).getDisplayValue() === updateDay) { //Checks for the correct student and day
        if (studentIndex.offset(0,6).getDisplayValue() === "IN_CLASS") { //Checks if the student is still in class.
          if (updateCheckIn) { //Checks to make sure we are not updating the check out time of a student who is still in class.
            studentIndex.offset(0,7).setValue(newTime);
          }
          else {
            FUNCTIONSTATUS(false, false, "Error: Student Still In Class");
            return;
          }
        }
        else if (updateCheckIn) {
          studentIndex.offset(0,7).setValue(newTime);
        }
        else {
          studentIndex.offset(0,8).setValue(newTime);
        }
        FUNCTIONSTATUS(false, true, "Attendance Time Updated Successfully")
        return;
      }
      studentIndex = studentIndex.offset(1,0);
    }
    FUNCTIONSTATUS(false, false, "Error: Student Not Found");
  }
}


function UPDATEINFOGRAPHIC() { //Updates the Infographic.
  if (STARTUP()) {
    let studentIndex = sheet.getRange('A4');
    let mondayCount = 0;
    let tuesdayCount = 0;
    let thursdayCount = 0;
    let saturdayCount = 0;
    let singleDayCount = 0;
    let multiDayCount = 0;
    let singleSubjectCount = 0;
    let multiSubjectCount = 0;
    let currentDayCombo = "";
    let currentDay = "";
    let subjectCount = "";
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      currentDayCombo = studentIndex.offset(0,3).getDisplayValue();
      currentDay = studentIndex.offset(0,4).getDisplayValue();
      subjectCount = studentIndex.offset(0,5).getDisplayValue();
      if (currentDay === "Monday") { //Adds to the Monday Count.
        mondayCount++;
      }
      if (currentDay === "Tuesday") { //Adds to the Tuesday Count.
        tuesdayCount++;
      }
      if (currentDay === "Thursday") { //Adds to the Thursday Count.
        thursdayCount++;
      }
      if (currentDay === "Saturday") { //Adds to the Saturday Count.
        saturdayCount++;
      }
      if (UNIQUESTUDENT(currentDayCombo, currentDay)) { //Only adds to the upcoming variables if the current student has not shown up yet.
        if (currentDayCombo === "Mondays" || currentDayCombo === "Tuesdays" || currentDayCombo === "Thursdays" || currentDayCombo === "Saturdays") { //Adds to the single-day count
          singleDayCount++;
        }
        else { //Adds to the multi-day count.
          multiDayCount++;
        }
        if (parseInt(subjectCount) === 1) { //Adds to the single-subject count.
          singleSubjectCount++;
        }
        else { //Adds to the multi-subject count.
          multiSubjectCount++;
        }
      }
      studentIndex = studentIndex.offset(1,0);
    }
    sheet.getRange('Q20').setValue(mondayCount);
    sheet.getRange('Q21').setValue(tuesdayCount);
    sheet.getRange('Q22').setValue(thursdayCount);
    sheet.getRange('Q23').setValue(saturdayCount);
    sheet.getRange('Q24').setValue(singleDayCount);
    sheet.getRange('Q25').setValue(multiDayCount);
    sheet.getRange('Q26').setValue(singleSubjectCount);
    sheet.getRange('Q27').setValue(multiSubjectCount); //Updates all of the infographic sections that can be updated.
    FUNCTIONSTATUS(false, true, "Infographic Updated Sucessfully");
  }
}


function SEARCHSTUDENT() { //Searches the data for students that match a name and/or id query.
  if (STARTUP()) {
    let studentIndex = sheet.getRange('A4');
    let name = sheet.getRange('M36').getDisplayValue();
    let id = sheet.getRange('M37').getDisplayValue();
    if (name === "" && id === "") { //Name or ID is required.
      sheet.getRange('L40').setValue("Error: Missing Input");
      FUNCTIONSTATUS(false, false, "Error: Missing Input");
      return;
    }
    if (name.length < 3 && (name !== "" || id === "")) { //Name Query must be at least 3 characters long.
      sheet.getRange('L40').setValue("Error: Search Query Too Short");
      FUNCTIONSTATUS(false, false, "Error: Search Query Too Short");
      return;
    }
    let foundStudent = false;
    let multipleIDs = false;
    let idSearch = false;
    let foundID = "";
    let outputString = "Query \"".concat(name); //The base of the output.
    if (id !== "") { //Adds additional info if an id query is added on.
      if (name !== "") { //Checks if it is an id-only query
        foundID = id;
        outputString = outputString.concat("\" with ID \"",id);
      }
      else {
        outputString = "Query with ID \"".concat(id);
        idSearch = true;
      }
    }
    outputString = outputString.concat("\" can be found on row(s): ");
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      if (!idSearch && studentIndex.getDisplayValue().includes(name) && (id === "" || studentIndex.offset(0,1).getDisplayValue  () == id)) { //Checks for matching students
        if (!(foundID === "" || studentIndex.offset(0,1).getDisplayValue() === foundID)) { //Checks if multiple different students have been found.
          multipleIDs = true;
        }
        if (!foundStudent) { //Checks if the first student is found.
          if (foundID === "") { //Checks if the ID input has not been filled (To check for multiple different students).
           foundID = studentIndex.offset(0,1).getDisplayValue();
          }
          foundStudent = true;
        }
        else { //Modifies the addition to the output string if not the first entry found.
          outputString = outputString.concat(", ");
        }
        outputString = outputString.concat(studentIndex.getRow());
      }
      else if (idSearch && studentIndex.offset(0,1).getDisplayValue() === id) { //Checks for matching students based on an id-only search.
        if (!foundStudent) { //Checks if it is the first student found.
          foundStudent = true;
        }
        else {
          outputString = outputString.concat(", ");
        }
        outputString = outputString.concat(studentIndex.getRow());
      }
      studentIndex = studentIndex.offset(1,0);
    }
    if (foundStudent) { //Checks if any valid entries were found.
      if (!idSearch) { //Checks if it is an id-only search.
        if (multipleIDs) { //Checks if multiple different students were found to make a note of it.
          outputString = outputString.concat(". NOTE: Multiple Different Students Found.");
        }
        else if (id === "") { //Checks if it is a name-only search.
          outputString = outputString.concat(" (ID: ", foundID, ")")
        }
      }
      sheet.getRange('L40').setValue(outputString);
      FUNCTIONSTATUS(false, true, "Query Completed Sucessfully");
      return;
    }
    let noResultsMessage = "No results for \"".concat(name); //Outputs a different message if no entries were found.
    if (idSearch) { //Checks if it is an id-only search.
      noResultsMessage = "No results for ID \"" .concat(id, "\" found");
    }
    else {
      if (id !== "") { //Checks if it is a name-only search.
        noResultsMessage = noResultsMessage.concat("\" with ID \"", id);
      }
      noResultsMessage =  noResultsMessage.concat("\" found.");
    }
    sheet.getRange('L40').setValue(noResultsMessage);
    FUNCTIONSTATUS(false, true, "No Students Found");
  }
}


function SORTSTUDENTS(manual = true) { //Sorts all students based on day of the week of their attendance, then by either their name, their student ID, or their expected time.
  if (STARTUP(manual)) {
    let mondayStudents = [];
    let tuesdayStudents = [];
    let thursdayStudents = [];
    let saturdayStudents = []; //Creates an array for each possible day of the week.
    let lateSaturdayStudents = []; //Logs students that come in at 1:00pm or later on Saturdays for sorting reasons.
    let addIndex = sheet.getRange('A4:I4'); //Gets the range of data that needs to be sorted
    let sortByName = (sheet.getRange('O35').getDisplayValue() !== "");
    let sortByID = (sheet.getRange('O36').getDisplayValue() !== "");
    let sortByTime = (sheet.getRange('O37').getDisplayValue() !== "");
    if (!manual) { //Checks if the function is being run by another function (Sorts by time if so).
      sortByName = false;
      sortByID = false;
      sortByTime = true;
    }
    if (!(sortByName || sortByID || sortByTime)) { //Checks that at least one sort option is true.
      FUNCTIONSTATUS(false, false, "Error: Missing Input");
      return;
    }
    if ((sortByName && sortByID) || (sortByName && sortByTime) || (sortByID && sortByTime)) { //Checks if ONLY 1 sort option is true.
      FUNCTIONSTATUS(false, false, "Error: Multiple Sort Queries Selected");
      return;
    }
    for (let checkLimit = 0; checkLimit < maximumRows && addIndex.getDisplayValues()[0][0] !== ""; checkLimit++) {
      let dayIndex = addIndex.getDisplayValues()[0][4];
      if (dayIndex === "Monday") { //Adds Monday students to Monday array.
        mondayStudents.push(addIndex.getDisplayValues());
      } 
      else if (dayIndex === "Tuesday") { //Adds Tuesday students to Tuesday array.
        tuesdayStudents.push(addIndex.getDisplayValues());
      } 
      else if (dayIndex === "Thursday") { //Adds Thursday students to Thursday array.
        thursdayStudents.push(addIndex.getDisplayValues());
      } 
      else if (dayIndex === "Saturday") { //Adds Saturday students to Saturday array.
        let checkTime = addIndex.getDisplayValues()[0][6];
        if (sortByTime && checkTime.substring(0,2) === "1:") { //Checks if the Saturday student comes in 1:00PM or later due to sorting issues.
          lateSaturdayStudents.push(addIndex.getDisplayValues());
        }
        else {
          saturdayStudents.push(addIndex.getDisplayValues());
        }
      } 
      else { //Recheck column E for mispellings if you are getting this error.
        if (manual) {
          FUNCTIONSTATUS(false, false, "Error: Unknown Day Found in Row ".concat(addIndex.getRow().toString())); 
        }
        return;
      }
      addIndex = addIndex.offset(1,0); //Moves indexes to the next data point.
    }
    if (sortByName) { //Sorts each array based on student name.
      mondayStudents.sort(function(a,b) {return a[0][0].localeCompare(b[0][0])});
      tuesdayStudents.sort(function(a,b) {return a[0][0].localeCompare(b[0][0])});
      thursdayStudents.sort(function(a,b) {return a[0][0].localeCompare(b[0][0])});
      saturdayStudents.sort(function(a,b) {return a[0][0].localeCompare(b[0][0])});
    }
    else if (sortByID) { //Sorts each array based on student ID.
      mondayStudents.sort(function(a,b) {return parseInt(a[0][1]) - parseInt(b[0][1])});
      tuesdayStudents.sort(function(a,b) {return parseInt(a[0][1]) - parseInt(b[0][1])});
      thursdayStudents.sort(function(a,b) {return parseInt(a[0][1]) - parseInt(b[0][1])});
      saturdayStudents.sort(function(a,b) {return parseInt(a[0][1]) - parseInt(b[0][1])}); 
    }
    else { //Sorts each array based on attendance time.
      mondayStudents.sort(function(a,b) {return a[0][6].localeCompare(b[0][6])});
      tuesdayStudents.sort(function(a,b) {return a[0][6].localeCompare(b[0][6])});
      thursdayStudents.sort(function(a,b) {return a[0][6].localeCompare(b[0][6])});
      saturdayStudents.sort(function(a,b) {return a[0][6].localeCompare(b[0][6])});
      lateSaturdayStudents.sort(function(a,b) {return a[0][6].localeCompare(b[0][6])}); 
      saturdayStudents = saturdayStudents.concat(lateSaturdayStudents); //Combines both sorted Saturday arrays.
    }
    addIndex = sheet.getRange('A4:I4');
    let studentArray = [mondayStudents, tuesdayStudents, thursdayStudents, saturdayStudents]; //Puts the day arrays into a single 4d array. (Note: for studentArray[a][b][c][d], c is always 0 due to the way Ranges stores values).
    let dayCount = 0;
    let replaceIndex = 0;
    for (let checkIndex = 0; checkIndex < maximumRows && dayCount < 4; checkIndex++) {
      addIndex.setValues(studentArray[dayCount][replaceIndex]) //Adds the values back based on the order of the day of the week (Monday, Tuesday, Thursday, Saturday).
      replaceIndex++;
      if (replaceIndex >= studentArray[dayCount].length) { //When each array is fully added, it moves on to the next array.
        dayCount++;
        replaceIndex = 0;
      }
      addIndex = addIndex.offset(1,0); //Moves indexes to the next data point.
    }
    sheet.getRange('O35:O37').setValues([[""],[""],[""]]);
    if (manual) {
      FUNCTIONSTATUS(false, true, "Students Sorted Successfully");
    }
  }
}


function UPDATENAME() { //Updates a student's Name.
  if (STARTUP()) {
    let oldName = sheet.getRange('O2').getDisplayValue();
    let id = sheet.getRange('O3').getDisplayValue();
    let newName = sheet.getRange('O4').getDisplayValue();
    let studentIndex = sheet.getRange('A4');
    let studentFound = false;
    if (oldName === "" || id === "" || newName === "") { //Checks that all necessary fields were marked
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    if (oldName === newName) { //Checks that the new name is different.
      FUNCTIONSTATUS(false, false, "Error: New Name Matches Old Name");
      return;
    }
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      if (studentIndex.getDisplayValue() === oldName && studentIndex.offset(0,1).getDisplayValue() === id) { //Checks for matching student name and ID.
        studentIndex.setValue(newName);
        if (studentFound) { //Ends the loop early once the maximum number of entries have been found.
          break;
        }
        studentFound = true;
      }
      studentIndex = studentIndex.offset(1,0);
    }
    if (studentFound) { //Checks if any valid entries were found.
      sheet.getRange('Q28').setValue(newName);
      FUNCTIONSTATUS(false, true, "Name Updated Successfully");
      return;
    }
    FUNCTIONSTATUS(false, false, "Error: Student Not Found");
  }
}


function UPDATEID() { //Updates a student's ID.
  if (STARTUP()) {
    let name = sheet.getRange('O8').getDisplayValue();
    let oldID = sheet.getRange('O9').getDisplayValue();
    let newID = parseInt(sheet.getRange('O10').getDisplayValue());
    let studentIndex = sheet.getRange('A4');
    let validArray = []; //Checks for all valid entries to update.
    let studentFound = false;
    if (name === "" || oldID === "" || newID === "") { //Checks that all necessary fields were marked.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    if (oldID === newID) { //Checks that the new ID is different.
      FUNCTIONSTATUS(false, false, "Error: New ID Matches Old ID");
      return;
    }
    if (!Number.isInteger(newID) || newID < minimumID) { //Checks if the newID is valid.
      FUNCTIONSTATUS(false, false, "Error: Invalid ID");
      return;
    }
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      if (studentIndex.getDisplayValue() === name && studentIndex.offset(0,1).getDisplayValue() === oldID) { //Checks for matching student name and ID.
        validArray.push(studentIndex.offset(0,1));
        if (studentFound) { //Ends the loop early once the maximum number of entries have been found.
          break;
        }
        studentFound = true;
      }
      else if (studentIndex.offset(0,1).getDisplayValue() === newID) { //Checks if the new ID is already used.
        FUNCTIONSTATUS(false, false, "Error: New ID Already in Use");
        return;
      }
      studentIndex = studentIndex.offset(1,0);
    }
    if (studentFound) { //Checks if any valid entries were found.
      for (let arrayIndex = 0; arrayIndex < validArray.length; arrayIndex++) { //Updates the respective entries.
        validArray[arrayIndex].setValue(newID);
      }
      if (parseInt(sheet.getRange('L28').getDisplayValue()) > parseInt(oldID)) { //Updates the lowest displayed ID if needed.
        GETSMALLESTUNUSEDID(false);
      }
      sheet.getRange('Q28').setValue(name);
      FUNCTIONSTATUS(false, true, "ID Updated Successfully");
      return;
    }
    FUNCTIONSTATUS(false, false, "Error: Student Not Found");
  }
}


function UPDATESUBJECTCOUNT() { //Updates a student's subject count.
  if (STARTUP()) {
    let name = sheet.getRange('O14').getDisplayValue();
    let id = sheet.getRange('O15').getDisplayValue();
    let newCount = sheet.getRange('O16').getDisplayValue();
    let studentIndex = sheet.getRange('A4');
    let studentFound = false;
    if (name === "" || id === "" || newCount === '') { //Checks for any missing inputs.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    if (!(Number(newCount) === 1 || Number(newCount) === 2)) { //Checks for valid subject counts.
      FUNCTIONSTATUS(false, false, "Error: Invalid Subject Count");
      return;
    }
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      if (studentIndex.getDisplayValue() === name && studentIndex.offset(0,1).getDisplayValue() === id) { //Checks for the matching student.
        if (studentIndex.offset(0,5).getDisplayValue() === newCount) { //Checks if the current subject count is the same as the new subject count.
          FUNCTIONSTATUS(false, false, "Error: New Subject Count Equals Old Subject Count");
          return;
        }
        studentIndex.offset(0,5).setValue(newCount);
        if (studentFound) { //Ends the loop early once the maximum number of entries have been found.
          break;
        }
        studentFound = true;
      }
      studentIndex = studentIndex.offset(1,0);
    }
    if (studentFound) {
      BINDERCOUNT(false);
      sheet.getRange('Q36').setValue(name);
      if (newCount === "2") {
        UPDATECOUNTINFO(true, "S->M");
      }
      else {
        UPDATECOUNTINFO(true, "M->S");
      }
      FUNCTIONSTATUS(false, true, "Subject Count Updated Sucessfully");
      return;
    }
    FUNCTIONSTATUS(false, false, "Error: Student Not Found");
  }
}


function ADDBARCODE() { //Attaches a new barcode to a specific student.
  if (STARTUP()) {
    let nameInput = sheet.getRange('M2').getDisplayValue(); //Gets the name input.
    let idInput = sheet.getRange('M3').getDisplayValue(); //Gets the student ID input.
    let barcodeInput = Number(sheet.getRange('M4').getDisplayValue()); //Gets the barcode input.
    let nameIndex = sheet.getRange('A4'); //Gets the name of each data point.
    let barcodeIndex = sheet.getRange('C4'); //Gets the barcode of each data point.
    let validReplacements = [];
    let addedBarcode = false;
    let oldBarcode = "";
    if (nameInput === "" || idInput === "" || barcodeInput === "") { //Checks for necessary inputs.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    if (!Number.isInteger(barcodeInput) || barcodeInput < minimumBarcode) { //Checks for a valid barcode.
      FUNCTIONSTATUS(false, false, "Error: Invalid Barcode");
      return;
    }
    for (let checkLimit = 0; checkLimit < maximumRows && nameIndex.getDisplayValue() !== ""; checkLimit++) {
      if (barcodeIndex.getDisplayValue() === barcodeInput) { //Checks for duplicate barcodes.
        if (sheet.getRange('M28').getDisplayValue() >= barcodeInput) {
          GETSMALLESTUNUSEDBARCODE(false);
        }
        FUNCTIONSTATUS(false, false, "Error: Barcode Already Used");
        return;
      } 
      if (nameIndex.getDisplayValue() === nameInput && nameIndex.offset(0,1).getDisplayValue() === idInput) { //Checks for the  data entries to add the barcode to.
        validReplacements.push(barcodeIndex);
        oldBarcode = nameIndex.offset(0,2).getDisplayValue();
        if (addedBarcode) { //Ends the loop early once the maximum number of entries have been found.
          break;
        }
        addedBarcode = true;
      } 
      nameIndex = nameIndex.offset(1,0);
      barcodeIndex = barcodeIndex.offset(1,0); //Moves indexes to the next data point.
    }
    if (addedBarcode) {
      for (let insertIndex = 0; insertIndex < validReplacements.length; insertIndex++) { //Inserts the barcode replacements.
        validReplacements[insertIndex].setValue(barcodeInput);
      }
      let oldSmallestUnusedBarcode = Number(sheet.getRange('M28').getDisplayValue());
      if (Number(oldBarcode) < oldSmallestUnusedBarcode || Number(barcodeInput) === oldSmallestUnusedBarcode) { //Updates the smallest unused barcode if needed.
        GETSMALLESTUNUSEDBARCODE(false);
      }
      sheet.getRange('Q28').setValue(nameInput);
      FUNCTIONSTATUS(false, true, "Updated Barcode(s) Successfully");
      return;
    }
    FUNCTIONSTATUS(false, false, "Error: Unknown Student");
  }
}


function REMOVEBARCODE() { //Removes all instance of a specific barcode.
  if (STARTUP()) {
    let barcodeCheck = sheet.getRange('C4'); //Gets the barcode for each data point
    let removalInput = sheet.getRange('M8'); 
    let barcodeRemoval = removalInput.getDisplayValue(); //Gets the barcode to be removed.
    let removed = false;
    if (barcodeRemoval === "") { //Checks if there is a valid input.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    removalInput.setValue("");
    for (let checkLimit = 0; checkLimit < maximumRows && barcodeCheck.getDisplayValue() !== ""; checkLimit++) {
      if (barcodeCheck.getDisplayValue() === barcodeRemoval) { //Removes the barcode if it is the targeted barcode and sets it to a default value.
        barcodeCheck.setValue("NOT_ASSIGNED");
        if (removed) { //Ends the loop early once the maximum number of entries have been found.
          break;
        }
        removed = true;
      }
      barcodeCheck = barcodeCheck.offset(1,0); //Moves barcode index to the next data point.
    }
    if (removed) { //Checks if a barcode was removed or not.
      if (Number(sheet.getRange('M28').getDisplayValue()) > Number(barcodeRemoval)) { //Updates the smallest unused barcode if needed.
        GETSMALLESTUNUSEDBARCODE(false);
      }
      sheet.getRange('Q28').setValue(name);
      FUNCTIONSTATUS(false, true, "Barcode(s) Removed Successfully");
      return;
    }
    FUNCTIONSTATUS(false, false, "Error: Barcode Not Found");
  }
}


function ENDDAY() { //Signs-out any students that are still signed-in.
  if (STARTUP()) {
    let logoutIndex = sheet.getRange('I4'); //Gets the value that determines if a student is in class or not.
    let logoutTime = sheet.getRange('K1').getDisplayValue(); //Gets the current time.
    let logoutConfirmation = sheet.getRange('M12'); //Checks confirmation field.
    let day = sheet.getRange('I1').getDisplayValue();
    if (logoutConfirmation.getDisplayValue === "") { //Checks for confirmation.
      FUNCTIONSTATUS(false, false, "Error: Confirmation Not Given");
      return;
    }
    logoutConfirmation.setValue("");
    for (let checkLimit = 0; checkLimit < maximumRows && logoutIndex.getDisplayValue() !== ""; checkLimit++) {
      if (logoutIndex.getDisplayValue() === "IN_CLASS") { //Signs-out students if they are in class.
        logoutIndex.setValue(logoutTime);
      }
      logoutIndex = logoutIndex.offset(1,0) //Moves index to the next data point.
    }
    sheet.getRange('Q11').setValue("0");
    FUNCTIONSTATUS(false, true, "Student(s) Signed-Out Successfully");
  }
}


function BACKUPDATA() { //Backs up the current data.
  if (STARTUP()) {
    SORTSTUDENTS(false);
    let lastRow = parseInt(sheet.getRange('L46').getDisplayValue());
    let currentDataRange = "A4:I".concat((lastRow + 3).toString()); //Gets the range of all the current data.
    let backupData = sheet.getRange('R2:Z'.concat((lastRow + 1).toString())); //Gets the data range to backup the current data to.
    backupData.setValues(sheet.getRange(currentDataRange).getDisplayValues()); //Copies the current data to the backup data.
    let cleanExcessData = sheet.getRange('R'.concat((lastRow + 2).toString(),':Z',(lastRow + 2).toString()));
    for (let excessCount = 0; excessCount < maximumDataEntries - (lastRow + 2); excessCount++) { //Clears any potential remaining old backup data.
      cleanExcessData.setValues(clearRow);
      cleanExcessData = cleanExcessData.offset(1,0);
    }
    sheet.getRange('Q19').setValue(sheet.getRange('G1').getDisplayValue());
    FUNCTIONSTATUS(false, true, "Data Backed Up Sucessfully");
  }
}


function IMPORTBACKUPDATA() { //Replaces the current data back with the backup data.
  if (STARTUP()) {
    let importConfirmation = sheet.getRange('Q4').getDisplayValue();
    let backupStudentIndex = sheet.getRange('R2');
    if (importConfirmation === "") {
      FUNCTIONSTATUS(false, false, "Error: Confirmation Not Given");
      return;
    }
    for (let studentCount = 0; studentCount < maximumDataEntries && backupStudentIndex.getDisplayValue() !== ""; studentCount++) {
      backupStudentIndex = backupStudentIndex.offset(1,0);
    }
    backupStudentIndex = backupStudentIndex.offset(-1,0);
    let backupRowCount = backupStudentIndex.getRow();
    if (backupRowCount <= 1) {
      FUNCTIONSTATUS(false, false, "Error: Not Backup Data Found");
      return;
    }
    let backupDataRange = sheet.getRange('R2:Z'.concat(backupRowCount.toString()));
    let insertBackupRange = sheet.getRange('A4:I'.concat((backupRowCount + 2).toString()));
    insertBackupRange.setValues(backupDataRange.getDisplayValues());
    sheet.getRange('Q4').setValue("");
    BINDERCOUNT(false);
    GETLOWESTUNUSEDBARCODE(false);
    GETLOWESTUNUSEDID(false);
    ROWCOUNT(false);
    SORTSTUDENTS(false);
    FUNCTIONSTATUS(false, true, "Imported Backup Data Successfully");
  }
}


function ADDSTUDENT() { //Adds a student to the database.
  if (STARTUP()) {
    let name = sheet.getRange('M16').getDisplayValue(); //Gets new student's name.
    let id = sheet.getRange('M17').getDisplayValue(); //Gets new student's id.
    let barcode = sheet.getRange('M18').getDisplayValue(); //Gets new student's barcode (If available).
    let dayCount = parseInt(sheet.getRange('M19').getDisplayValue()); //Gets the number of days in the week that the student will attend.
    let dayOne = sheet.getRange('M20').getDisplayValue(); //Gets the first day of the week.
    let dayOneTime = sheet.getRange('M21').getDisplayValue(); //Gets the time for the first day of the week.
    let dayTwo = sheet.getRange('M22').getDisplayValue(); //Gets the second day of the week (If applicable).
    let dayTwoTime = sheet.getRange('M23').getDisplayValue(); //Gets the time for the second day of the week (If applicable).
    let subjectCount = parseInt(sheet.getRange('M24').getDisplayValue()); //Gets the number of subjects the student is learning.
    if (id === "") { //Sets the id to the lowest unused ID if left blank.
      GETSMALLESTUNUSEDID(false);
      id = sheet.getRange('L28').getDisplayValue();
    }
    if (name === "" || dayCount === "" || dayOne === "" || dayOneTime === "" || subjectCount === "" || (dayCount === 2 && (dayTwo === "" || dayTwoTime === ""))) { //Checks for missing inputs.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    if (name.length < 3) { //Checks if the name meets the minimum length.
      FUNCTIONSTATUS(false, false, "Error: Name Too Short");
      return;
    }
    let testMinimum = (id < minimumID);
    let testVar = (!Number.isInteger(parseInt(id)));
    if (!Number.isInteger(parseInt(id)) || id < minimumID) { //Checks if the newID is valid.
      FUNCTIONSTATUS(false, false, "Error: Invalid ID");
      return;
    }
    if (!((Number.isInteger(parseInt(barcode)) && barcode >= minimumBarcode) || barcode === "")) { //Checks for a valid barcode.
      FUNCTIONSTATUS(false, false, "Error: Invalid Barcode");
      return;
    }
    if (!(Number(dayCount) === 1 || Number(dayCount) === 2)) { //Checks for correct day count.
      FUNCTIONSTATUS(false, false, "Error: Invalid Day Count");
      return;
    }
    if ((Number(dayCount) === 1 && dayTwo !== "") || (Number(dayCount) === 2 && dayTwo === "")) {
      FUNCTIONSTATUS(false, false, "Error: Invalid Number Of Day(s) Inserted");
    }
    if (!(dayOne === "Monday" || dayOne === "Tuesday" || dayOne === "Thursday" || dayOne === "Saturday") || (dayCount === 2 && !(dayTwo === "Monday" || dayTwo === "Tuesday" || dayTwo === "Thursday" || dayTwo === "Saturday"))) { //Checks for correct days of the week.
      FUNCTIONSTATUS(false, false, "Error: Invalid Day(s)");
      return;
    }
    if (!CHECKVALIDTIME(dayOneTime, dayOne) || (dayTwoTime !== "" && !CHECKVALIDTIME(dayTwoTime, dayTwo))) { //Checks if the given times are valid.
      FUNCTIONSTATUS(false, false, "Error: Invalid Time(s)");
      return;
    }
    if (Number(dayCount) === 2 && dayOne === dayTwo) { //Checks for duplicate days of the week.
      FUNCTIONSTATUS(false, false, "Error: Matching Days");
      return;
    }
    if (!(Number(subjectCount) === 1 || Number(subjectCount) === 2)) { // Checks for correct subject count.
      FUNCTIONSTATUS(false, false, "Error: Invalid Subject Count");
      return;
    }
    if (barcode === "") { //Sets the barcode to a default value if left blank.
      barcode = "NOT_ASSIGNED";
    }
    let findNewInput = sheet.getRange('A4'); //Gets the index in order to find the next available row for data entry.
    for (let newIndex = 0; newIndex < maximumRows && findNewInput.getDisplayValue() !== ""; newIndex++) {
      if (findNewInput.offset(0,1).getDisplayValue() === id) { //Checks for duplicate ID.
        if (sheet.getRange('L28').getDisplayValue() >= id) {
          GETSMALLESTUNUSEDID(false);
        }
        FUNCTIONSTATUS(false, false, "Error: Student ID Already Used");
        return;
      }
      else if (findNewInput.offset(0,2).getDisplayValue() === barcode) { //Checks for duplicate barcodes.
        if (sheet.getRange('M28').getDisplayValue() === barcode) {
          GETSMALLESTUNUSEDBARCODE(false);
        }
        FUNCTIONSTATUS(false, false, "Error: Barcode Already Used");
        return;
      }
      findNewInput = findNewInput.offset(1,0); //Moves index to the next data point.
    }
    let insertDay = dayOne;
    let insertTime = dayOneTime;
    let secondDayInput = findNewInput.offset(1,0);
    let currRow = findNewInput.getRow().toString();
    for (let dayIndex = 0; dayIndex < dayCount; dayIndex++) { //Adds a new data entry for each day of the week the student attends.
      findNewInput.setValue(name); findNewInput = findNewInput.offset(0,1);
      findNewInput.setValue(id); findNewInput = findNewInput.offset(0,1);
      findNewInput.setValue(barcode); findNewInput = findNewInput.offset(0,1);
      if (dayCount === 1) { //Adds input based on the number of days of the week the student attends.
        findNewInput.setValue(insertDay.concat("s"));
      }
      else {
        findNewInput.setValue(GETDAYCOMBO(dayOne,dayTwo));
      } findNewInput = findNewInput.offset(0,1);
      findNewInput.setValue(insertDay); findNewInput = findNewInput.offset(0,1);
      findNewInput.setValue(subjectCount.toString()); findNewInput = findNewInput.offset(0,1);
      for (let insertTimeCount = 0; insertTimeCount < 3; insertTimeCount++) { //Adds the check in time for three rows in a row.
        findNewInput.setValue(insertTime); findNewInput = findNewInput.offset(0,1);
      }
      findNewInput.setValue(jFunction[0].concat(currRow,jFunction[1],currRow,jFunction[2],currRow,jFunction[3],currRow,jFunction[4],currRow,jFunction[5])); findNewInput = findNewInput.offset(0,1);
      findNewInput.setValue(kFunction[0].concat(currRow,kFunction[1],currRow,kFunction[2],currRow,kFunction[3],currRow,kFunction[4])); //Adds the two Sheet functions into their respective columns.
      if (dayCount === 2 && dayIndex === 0) { //Gets the data ready for the second data entry (If applicable).
        insertDay = dayTwo;
        insertTime = dayTwoTime;
        findNewInput = secondDayInput;
        currRow = findNewInput.getRow().toString();
      }
    }
    BINDERCOUNT(false);
    if (sheet.getRange('M28').getDisplayValue() === barcode) { //Updates the lowest unused barcode if needed.
      GETSMALLESTUNUSEDBARCODE(false);
    }
    if (sheet.getRange('L28').getDisplayValue() === id) { //Updates the lowest unused ID if needed.
      GETSMALLESTUNUSEDID(false);
    }
    ROWCOUNT(false);
    SORTSTUDENTS(false);
    if (dayTwo === "") {
      UPDATEDAYINFO(dayOne, "+");
      UPDATECOUNTINFO(false, "S+");
    }
    else {
      UPDATEDAYINFO(dayOne, "+", dayTwo, "+");
      UPDATECOUNTINFO(false, "M+");
    }
    if (subjectCount === 1) {
      UPDATECOUNTINFO(true, "S+");
    }
    else {
      UPDATECOUNTINFO(true, "M+");
    }
    sheet.getRange('Q28').setValue(name);
    FUNCTIONSTATUS(false, true, "Student Added Successfully");
  }
}


function REMOVESTUDENT() { //Removes a student from the database.
  if (STARTUP()) {
    let removeName = sheet.getRange('O30').getDisplayValue(); //Gets the name of the student to be removed.
    let removeID = sheet.getRange('O31').getDisplayValue(); //Gets the id of the student to be removed.
    if (removeName === "" || removeID === "") { //Checks if the name of the student is present.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs");
      return;
    }
    let removedRows = []; //Gets the rows that will have the data removed (There is a hole if there are any rows below it with data).
    let removedDays = [];
    let removed = false;
    let removeCheck = sheet.getRange('A4');
    let removeBarcode = "";
    let removeDayCombo = "";
    let removeSubjectCount = "";
    for (let removeIndex = 0; removeIndex < maximumRows && removeCheck.getDisplayValue() !== ""; removeIndex++) {
      if (removeCheck.getDisplayValue() === removeName && removeCheck.offset(0,1).getDisplayValue() === removeID) { //Checks if the current student matches the one to be removed.
        let rowRemoval = removeCheck.getRow();
        let deleteStudent = sheet.getRange("A".concat(rowRemoval.toString(),":I",rowRemoval.toString()));
        removeBarcode = removeCheck.offset(0,2).getDisplayValue();
        removeDayCombo = removeCheck.offset(0,3).getDisplayValue();
        removeDay = removeCheck.offset(0,4).getDisplayValue();
        removeSubjectCount = parseInt(removeCheck.offset(0,5).getDisplayValue());
        removedRows.push(rowRemoval);
        removedDays.push(removeDay);
        deleteStudent.setValues(clearRow);
        removed = true;
      }
      removeCheck = removeCheck.offset(1,0); //Moves index to the next data point.
    }
    if (removeCheck.getDisplayValue() === "") { //Pushes the index up a row to try to get it on a data point if it is not on one already.
      removeCheck = removeCheck.offset(-1,0);
    }
    while (removeCheck.getDisplayValue() === "" && removedRows.length > 0) { //Removes any rows from the array that were removed from the very bottom row(s) (These rows are not holes in the data).
      removedRows.pop();
      removeCheck.offset(-1,0);
    }
    if (removed) { //Checks if any rows are left in the array.
      for (let arrayIndex = 0; arrayIndex < removedRows.length; arrayIndex++) { //Fills any holes in the data with data from the end of the database.
        let emptyRow = removeCheck.getRow();
        let emptyData = sheet.getRange("A".concat(emptyRow.toString(),":I",emptyRow.toString())); //Gets the row to be moved
        let fillRow = removedRows[arrayIndex];
        let fillData = sheet.getRange("A".concat(fillRow.toString(),":I",fillRow.toString())); //Gets the earliest row that needs to be filled
        fillData.setValues(emptyData.getDisplayValues());
        emptyData.setValues(clearRow);
        removeCheck = removeCheck.offset(-1,0);
      }
      BINDERCOUNT(false);
      if (Number(sheet.getRange('M28').getDisplayValue()) > Number(removeBarcode)) { //Updates the lowest unused barcode if needed.
        GETSMALLESTUNUSEDBARCODE(false);
      }
      if (Number(sheet.getRange('L28').getDisplayValue()) > Number(removeID)) { //Updates the lowest unused ID if needed.
        GETSMALLESTUNUSEDID(false);
      }
      ROWCOUNT(false);
      SORTSTUDENTS(false);
      for (let removalIndex = 0; removalIndex < removedRows.length; removalIndex++) {
        UPDATEDAYINFO(removedDays[removalIndex],"-")
      }
      if (removeDayCombo === "Mondays" || removeDayCombo === "Mondays" || removeDayCombo === "Mondays" || removeDayCombo === "Mondays") {
        UPDATECOUNTINFO(false, "S-");
      }
      else {
        UPDATECOUNTINFO(false, "M-");
      }
      if (removeSubjectCount === 1) {
        UPDATECOUNTINFO(true, "S-");
      }
      else {
        UPDATECOUNTINFO(true, "M-");
      }
      sheet.getRange('Q28').setValue(removeName);
      FUNCTIONSTATUS(false, true, "Student Removed Successfully");
      return;
    }
    FUNCTIONSTATUS(false, false, "Error: Student Not Found");
  }
}


function ADJUSTSTUDENTSCHEDULE() { //Adjusts the schedule of a student.
  if (STARTUP()) {
    let name = sheet.getRange('O20').getDisplayValue();
    let id = sheet.getRange('O21').getDisplayValue();
    let add = (sheet.getRange('O22').getDisplayValue() !== "");
    let removal = (sheet.getRange('O23').getDisplayValue() !== "");
    let oldDay = sheet.getRange('O24').getDisplayValue();
    let newDay = sheet.getRange('O25').getDisplayValue();
    let newTime = sheet.getRange('O26').getDisplayValue();
    let studentIndex = sheet.getRange('A4');
    let studentFound = false;
    if (name === "" || id === "" || (add && (newDay === "" || newTime === "")) || (removal && oldDay === "") || (!add && !removal && (oldDay === "" || (newDay === "" && newTime === "")))) { //Checks for valid inputs.
      FUNCTIONSTATUS(false, false, "Error: Missing Inputs"); 
      return;
    }
    if (add && removal) { //Checks if both add and removal are not both active.
      FUNCTIONSTATUS(false, false, "Error: Add and Removal Both Selected");
      return;
    }
    if ((oldDay !== "" && !(oldDay === "Monday" || oldDay === "Tuesday" || oldDay === "Thursday" || oldDay === "Saturday")) || (newDay !== "" && !(newDay === "Monday" || newDay === "Tuesday" || newDay === "Thursday" || newDay === "Saturday"))) { //Checks for the correct days.
      FUNCTIONSTATUS(false, false, "Error: Invaid Day(s)");
      return;
    }
    if (newTime !== "" && ((newDay === "" && !CHECKVALIDTIME(newTime, oldDay)) || (newDay !== "" && !CHECKVALIDTIME(newTime, newDay)))) { //Checks if the new time is valid if needed.
      FUNCTIONSTATUS(false, false, "Error: Invaid Time");
      return;
    }
    if (add) { //Runs if a day is being added.
      let barcode = "";
      let subjectCount = "";
      let dayPair = "";
      let getFirstDayPair = "";
      for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
        if (studentIndex.getDisplayValue() === name && studentIndex.offset(0,1).getDisplayValue() === id) { //Checks if it is the correct student to gather the right information for the second entry.
          if (!studentFound) { //Checks if the student is going from one day to two days.
            if (studentIndex.offset(0,4) === newDay) { //Checks if the new day matches the old day
              FUNCTIONSTATUS(false, false, "Error: Student Already Attends Given Day");
              return;
            }
            barcode = studentIndex.offset(0,2).getDisplayValue();
            dayPair = GETDAYCOMBO(studentIndex.offset(0,4).getDisplayValue(), newDay);
            subjectCount = studentIndex.offset(0,5).getDisplayValue();
            getFirstDayPair = studentIndex.offset(0,3);
            studentFound = true;
          }
          else {
            FUNCTIONSTATUS(false,false, "Error:  Student at Maximum Attendance");
            return;
          }
        }
        studentIndex = studentIndex.offset(1,0);
      }
      if (!studentFound) { //Checks to see if a student has been found.
        FUNCTIONSTATUS(false, false, "Error: Student not found");
        return;
      }
      getFirstDayPair.setValue(dayPair); //Inputs all the correct information for the first and second entry.
      let currRow = studentIndex.getRow().toString();
      studentIndex.setValue(name); studentIndex = studentIndex.offset(0,1);
      studentIndex.setValue(id); studentIndex = studentIndex.offset(0,1);
      studentIndex.setValue(barcode); studentIndex = studentIndex.offset(0,1);
      studentIndex.setValue(dayPair); studentIndex = studentIndex.offset(0,1);
      studentIndex.setValue(newDay); studentIndex = studentIndex.offset(0,1);
      studentIndex.setValue(subjectCount); studentIndex = studentIndex.offset(0,1);
      for (let inputLoop = 0; inputLoop < 3; inputLoop++) {
        studentIndex.setValue(newTime); studentIndex = studentIndex.offset(0,1);
      }
      studentIndex.setValue(jFunction[0].concat(currRow,jFunction[1],currRow,jFunction[2],currRow,jFunction[3],currRow,jFunction[4],currRow,jFunction[5])); studentIndex = studentIndex.offset(0,1);
      studentIndex.setValue(kFunction[0].concat(currRow,kFunction[1],currRow,kFunction[2],currRow,kFunction[3],currRow,kFunction[4])); //Adds the two Sheet functions into their respective columns.
      sheet.getRange('O22').setValue("");
      ROWCOUNT(false);
      SORTSTUDENTS(false);
      UPDATEDAYINFO(newDay, "+");
      UPDATECOUNTINFO(false, "S->M");
      sheet.getRange('Q28').setValue(name);
      FUNCTIONSTATUS(false, true, "Day Added Successfully");
      return;
    }
    let getChangeRow = "";
    let newSingleDay = "";
    let adjustDayCombo = "";
    let dayComboCells = [];
    let changedDayCombo = false;
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      if (studentIndex.getDisplayValue() === name && studentIndex.offset(0,1).getDisplayValue() === id) { //Checks for the correct students.
        let checkDayCombo = studentIndex.offset(0,3).getDisplayValue();
        dayComboCells.push(studentIndex);
        if (studentIndex.offset(0,4).getDisplayValue() === oldDay) { //Checks for the matching old day.
          if (removal && (checkDayCombo === "Mondays" || checkDayCombo === "Tuesdays" || checkDayCombo === "Thursdays" || checkDayCombo === "Saturdays")) { //Checks that the student will still have at least 1 day in attendance.
            FUNCTIONSTATUS(false, false, "Error: Student at Minimum Attendance");
            return;
          }
          getChangeRow = studentIndex.offset(0,0);
          studentFound = true;
        }
        else if (!removal && oldDay !== newDay && !(checkDayCombo === "Monday" || checkDayCombo === "Tuesday" || checkDayCombo === "Thursday" || checkDayCombo === "Saturday")) { //Checks if a day combo needs to be updated.
          adjustDayCombo = studentIndex.offset(0,4).getDisplayValue();
          changedDayCombo = true;
        }
        else if (removal) { //Gets the cell for the day combo for the student's day that is not being removed to be updated.
          newSingleDay = sheet.getRange("D".concat(studentIndex.getRow().toString()));
        }
      }
      studentIndex = studentIndex.offset(1,0);
    }
    if (!studentFound) { //Checks to see if the student has been found.
      FUNCTIONSTATUS(false, false, "Error: Student not found");
      return;
    }
    if (removal) { //Checks if a day is being removed.
      let currRow = getChangeRow.getRow();
      let lastRow = studentIndex.offset(-1,0).getRow().toString();
      let getNewDay = GETSINGLEDAY(oldDay,newSingleDay.getDisplayValue()); //Converts a day pair into a single day based on   what is being removed.
      newSingleDay.setValue(getNewDay);
      let shiftTo = sheet.getRange("A".concat(currRow.toString(),":I",currRow.toString()));
      let shiftFrom = sheet.getRange("A".concat(lastRow,":I",lastRow));
      shiftTo.setValues(shiftFrom.getDisplayValues());
      shiftFrom.setValues(clearRow);
      sheet.getRange('O23').setValue("");
      ROWCOUNT(false);
      SORTSTUDENTS(false);
      sheet.getRange('Q36').setValue(name);
      UPDATEDAYINFO(oldDay, "-");
      UPDATECOUNTINFO(false, "M->S");
      FUNCTIONSTATUS(false, true, "Day Removed Successfully");
      return;
    }
    if (changedDayCombo) { //Checks if a day pair needs to be changed.
      if (adjustDayCombo === newDay) { //Checks if the old day is being assigned to a day the student is already attending
        FUNCTIONSTATUS(false, false, "Error: Student Already Attending on that Day");
        return;
      }
      let newDayCombo = GETDAYCOMBO(adjustDayCombo, newDay);
      for (let arrayIndex = 0; arrayIndex < dayComboCells.length; arrayIndex++) {
        dayComboCells[arrayIndex].offset(0,3).setValue(newDayCombo);
      }
    }
    else if (!(newDay === "" || newDay === oldDay)) { //Updates the old day to the new day if needed.
      getChangeRow.offset(0,3).setValue(newDay.concat("s"));
      getChangeRow.offset(0,4).setValue(newDay);
      UPDATEDAYINFO(oldDay, "-", newDay, "+");
    }
    if (newTime !== "") { //Runs if a new time needs to be set.
      getChangeRow.offset(0,6).setValue(newTime);
    }
    SORTSTUDENTS(false);
    sheet.getRange('Q28').setValue(name);
    FUNCTIONSTATUS(false, true, "Attendance Adjusted Successfully");
  }
}


function GETSINGLEDAY(removedDay, value) { //Returns a single day given a day pair and a day to remove.
  let days = [];
  if (value.includes("Monday")) { //Check if the pair includes Monday.
    days.push("Mondays");
  }
  if (value.includes("Tuesday")) { //Check if the pair includes Tuesday.
    days.push("Tuesdays");
  }
  if (value.includes("Thursday")) { //Check if the pair includes Thursday.
    days.push("Thursdays");
  }
  if (value.includes("Saturday")) { //Check if the pair includes Saturday.
    days.push("Saturdays");
  }
  if (days[0].includes(removedDay)) { //returns the value that is not the designated day to remove.
    return days[1];
  }
  return days[0];
}


function UPDATEDAYINFO(firstDay, firstAdjust, secondDay = "", secondAdjust = "+") { //Updates the enrollment per day portion of the infographic.
  let adjustDayArray = [firstDay];
  let adjusterArray = [firstAdjust];
  if (secondDay !== "") {
    adjustDayArray.push(secondDay);
    adjusterArray.push(secondAdjust);
  }
  for (let arrayIndex = 0; arrayIndex < adjustDayArray.length; arrayIndex++) {
    let adjustDay = adjustDayArray[arrayIndex];
    let adjuster = adjusterArray[arrayIndex];
    if (adjustDay === "Monday") {
      if (adjuster === "+") {
        sheet.getRange('Q20').setValue(parseInt(sheet.getRange('Q20').getDisplayValue()) + 1);
      }
      else {
        sheet.getRange('Q20').setValue(parseInt(sheet.getRange('Q20').getDisplayValue()) - 1);
      }
    }
    else if (adjustDay === "Tuesday") {
      if (adjuster === "+") {
        sheet.getRange('Q21').setValue(parseInt(sheet.getRange('Q21').getDisplayValue()) + 1);
      }
      else {
        sheet.getRange('Q21').setValue(parseInt(sheet.getRange('Q21').getDisplayValue()) - 1);
      }
    }
    else if (adjustDay === "Thursday") {
      if (adjuster === "+") {
        sheet.getRange('Q22').setValue(parseInt(sheet.getRange('Q22').getDisplayValue()) + 1);
      }
      else {
        sheet.getRange('Q22').setValue(parseInt(sheet.getRange('Q22').getDisplayValue()) - 1);
      }
    }
    else {
      if (adjuster === "+") {
        sheet.getRange('Q23').setValue(parseInt(sheet.getRange('Q23').getDisplayValue()) + 1);
      }
      else {
        sheet.getRange('Q23').setValue(parseInt(sheet.getRange('Q23').getDisplayValue()) - 1);
      }
    }
  }
}


function UPDATECOUNTINFO(subject, modifier) { //Updates the day count or the subject count portion of the infographic.
  let singleRange = sheet.getRange('Q24');
  let multiRange = sheet.getRange('Q25')
  if (subject) {
    singleRange = sheet.getRange('Q26');
    multiRange = sheet.getRange('Q27');
  }
  let singleValue = parseInt(singleRange.getDisplayValue());
  let multiValue = parseInt(multiRange.getDisplayValue());
  if (modifier === "S+") {
    singleRange.setValue(singleValue + 1);
  }
  else if (modifier === "S-") {
    singleRange.setValue(singleValue - 1);
  }
  else if (modifier === "M+") {
    multiRange.setValue(multiValue + 1);
  }
  else if (modifier === "M-") {
    multiRange.setValue(multiValue - 1);
  }
  else if (modifier === "S->M") {
    singleRange.setValue(singleValue - 1);
    multiRange.setValue(multiValue + 1);
  }
  else if (modifier === "M->S") {
    singleRange.setValue(singleValue + 1);
    multiRange.setValue(multiValue - 1);
  }
}


function GETDAYCOMBO(dayOne,dayTwo) { //Gets a string combination of two different days of the week.
  if ((dayOne === "Tuesday" && dayTwo === "Monday") || (dayOne === "Thursday" && dayTwo === "Monday") || (dayOne === "Saturday" && dayTwo === "Monday") || (dayOne === "Thursday" && dayTwo === "Tuesday") || (dayOne === "Saturday" && dayTwo === "Tuesday") || (dayOne === "Saturday" && dayTwo === "Thursday")) { //Checks if the two days of the week are out of order.
  [dayOne, dayTwo] = [dayTwo, dayOne];
  }
  return dayOne.concat("s/",dayTwo,"s");
}


function CHECKVALIDTIME(value, day) { //Check if a time is within Kumon hours given a time and day.
  let dividerIndex = 1; 
  if (value.charAt(2) === ":") { //Gets the string length of the hours hand.
    dividerIndex = 2; 
  } 
  if (!(value.length === 7 || (dividerIndex === 2 && value.length === 8))) { //Checks for the correct string length.
    return false;
  }
  let firstNumber = parseInt(value.substring(0, dividerIndex));
  let firstColon = value.substring(dividerIndex, dividerIndex + 1); dividerIndex++;
  let secondNumber = parseInt(value.substring(dividerIndex, dividerIndex + 1)); dividerIndex++;
  let thirdNumber = parseInt(value.substring(dividerIndex, dividerIndex + 1)); dividerIndex = dividerIndex + 2;
  let meridiem = value.substring(dividerIndex, value.length);
  if (firstColon !== ":") { //Checks for a colon.
    return false;
  }
  if (!(Number.isInteger(firstNumber) || Number.isInteger(secondNumber) || Number.isInteger(thirdNumber))) { //Checks for all numbers in the time are actually numbers.
    return false;
  }
  if (secondNumber < 0 || secondNumber > 5) { //Checks for a valid tens-minute.
    return false;
  }
  if (!(meridiem === "PM" || meridiem === "AM")) { //Checks if AM or PM is present.
    return false;
  }
  if (meridiem === "AM" && !((firstNumber === 10 || firstNumber === 11) && day === "Saturday")) { //Checks for a valid day & time comination for AM.
    return false
  }
  return (!(meridiem === "PM" && !(((firstNumber >= 2 && firstNumber < 7) && (day === "Monday" || day === "Tuesday" || day === "Thursday")) || ((firstNumber === 12 || firstNumber === 1) && day === "Saturday")))); //Checks for a valid day & time comination for PM; The last condition for a valid time.
}


function GETSMALLESTUNUSEDID(manual = true) { //Finds the lowest unused student ID.
  if (STARTUP(manual)) {
    let idIndex = sheet.getRange('B4');
    let idArray = [];
    let unusedMinimumID = minimumID;
    for (let studentCount = 0; studentCount < maximumRows && idIndex.offset(0,-1).getDisplayValue() !== ""; studentCount++) { //Gets every used ID and puts them into an array.
      idArray.push(parseInt(idIndex.getDisplayValue())); 
      idIndex = idIndex.offset(1,0);
    }
    idArray.sort(function (a,b) { return a - b}) //Sorts the array in ascending order.
    for (let arrayIndex = 0; arrayIndex < idArray.length; arrayIndex++) {
      if (idArray[arrayIndex] === unusedMinimumID || (unusedMinimumID > 0 && idArray[arrayIndex] === unusedMinimumID - 1)) { //Checks if the current ID skips over an integer (The integer skipped over is the lowest unused ID).
        unusedMinimumID = idArray[arrayIndex] + 1;
      }
      else {
        break;
      }
    }
    sheet.getRange('L28').setValue(unusedMinimumID);
    if (manual) {
      FUNCTIONSTATUS(false, true, "Smallest Unused ID Found: ".concat(unusedMinimumID.toString()));
    }
  }
}


function GETSMALLESTUNUSEDBARCODE(manual = true) { //Finds the lowest unused student ID.
  if (STARTUP(manual)) {
    if (manual) {
      sheet.getRange('Q32').setValue("Smallest Unused Barcode");
    }
    let barcodeIndex = sheet.getRange('C4');
    let barcodeArray = [];
    let unusedMinimumBarcode = minimumBarcode;
    for (let studentCount = 0; studentCount < maximumRows && barcodeIndex.offset(0,-2).getDisplayValue() !== ""; studentCount++) { //Gets every used barcode and puts them into an array.
      if (barcodeIndex.getDisplayValue() !== "NOT_ASSIGNED") {
        barcodeArray.push(parseInt(barcodeIndex.getDisplayValue())); 
      }
      barcodeIndex = barcodeIndex.offset(1,0);
    }
    barcodeArray.sort(function (a,b) { return a - b }) //Sorts the array in ascending order.
    for (let arrayIndex = 0; arrayIndex < barcodeArray.length; arrayIndex++) {
      if (barcodeArray[arrayIndex] === unusedMinimumBarcode || (unusedMinimumBarcode > 0 && barcodeArray[arrayIndex] === unusedMinimumBarcode - 1)) { //Checks if the current ID skips over an integer (The integer skipped over is the lowest unused ID).
        unusedMinimumBarcode = barcodeArray[arrayIndex] + 1;
      }
      else {
        break;
      }
    }
    sheet.getRange('M28').setValue(unusedMinimumBarcode);
    if (manual) {
      FUNCTIONSTATUS(false, true, "Smallest Unused Barcode Found: ".concat(unusedMinimumBarcode.toString()));
    }
  }
}

function FIXSTARTUP() { //Fixes any startup issues.
  sheet.getRange('F2:K2').setBackground("Red");
  FIXGAPS(false);
}


function FIXGAPS(manual = true) { //Removes any gaps in the data automatically. 
  if (!CHECKINPROGRESS()) {
    if (manual) {
      FUNCTIONSTATUS(true);
    }
    let emptyRows = [];
    let replaceIndex = sheet.getRange('A4:I4');
    let checkExistence = sheet.getRange('A4');
    let lastRow = 4;
    for (let dataCount = 0; dataCount < maximumDataEntries; dataCount++) { //Moves all non-empty data into an array.
      if (checkExistence.getDisplayValue() === "") { //Checksto see if the current row is empty.
        emptyRows.push(checkExistence.getRow());
      }
      else {
        lastRow = checkExistence.getRow(); //Gets the last row with data on it.
      }
      checkExistence = checkExistence.offset(1,0);
    }
    let lastDataIndex = sheet.getRange("A".concat(lastRow.toString(),":I",lastRow.toString()));
    for (let endCount = 0; endCount < maximumDataEntries && lastDataIndex.getDisplayValue() === "" && lastDataIndex.getRow() >= 4; endCount++) { //Moves the data index up to a point where there is a data entry
      lastDataIndex = lastDataIndex.offset(-1,0);
    }
    for (let insertCount = 0; insertCount < emptyRows.length; insertCount++) { //Puts the data in the array back onto the sheet.
      replaceIndex = sheet.getRange("A".concat(emptyRows[insertCount].toString(),":I",emptyRows[insertCount].toString()));
      if (replaceIndex.getRow() < lastDataIndex.getRow()) { //Checks if the two indexes pass each other.
        replaceIndex.setValues(lastDataIndex.getDisplayValues());
        lastDataIndex.setValues(clearRow);
        for (let endCount = 0; endCount < maximumDataEntries && lastDataIndex.getDisplayValue() === "" && lastDataIndex.getRow() >= 4; endCount++) {
          lastDataIndex = lastDataIndex.offset(-1,0);
        }
      }
      else {
        break;
      }
    }
    ROWCOUNT(false);
    if (manual) {
      FUNCTIONSTATUS(false, true, "Fixed Gaps in Data");
    }
  }
}


function ROWCOUNT(manual = true) { //Counts the number of rows and the number of unique students in the data.
  if (STARTUP(manual)) {
    let totalStudents = 0;
    let totalRows = 0;
    let studentCheck = sheet.getRange('A4');
    for (let studentIndex = 0; studentIndex < maximumDataEntries && studentCheck.getDisplayValue() !== ""; studentIndex++) {
      if (UNIQUESTUDENT(studentCheck.offset(0,3).getDisplayValue(),studentCheck.offset(0,4).getDisplayValue())) { //Checks if a student appeared a second time.
        totalStudents++;
      }
      totalRows++;
      studentCheck = studentCheck.offset(1,0);
    }
    sheet.getRange('L44').setValue(totalStudents);
    sheet.getRange('L46').setValue(totalRows);
    if (manual) {
      FUNCTIONSTATUS(false, true, "Row Count Updated Sucessfully");
    }
  }
}


function BINDERCOUNT(manual = true) { //Counts the number of subject binders being used.
  if (STARTUP(manual)) {
    let studentIndex = sheet.getRange('A4');
    let binderIndex = sheet.getRange('F4');
    let binderCount = 0;
    for (let studentCount = 0; studentCount < maximumRows && studentIndex.getDisplayValue() !== ""; studentCount++) {
      if (UNIQUESTUDENT(studentIndex.offset(0,3).getDisplayValue(), studentIndex.offset(0,4).getDisplayValue())) {
        binderCount += parseInt(binderIndex.getDisplayValue());
      }
      studentIndex = studentIndex.offset(1,0);
      binderIndex = binderIndex.offset(1,0);
    }
    sheet.getRange('L49').setValue(binderCount);
    if (manual) {
      FUNCTIONSTATUS(false, true, "Subject Binders Counted Sucessfully")
    }
  }
}


function UNIQUESTUDENT(checkDays,currentDay) { //Checks if a specific student has already shown up already in a row.
  if (!(checkDays === "Mondays" || checkDays === "Tuesdays" || checkDays === "Thursdays" || checkDays === "Saturdays")) { //Checks if a student attends two days (returns true if not).
    if (checkDays.includes("Monday")) { //Checks if the first day matches Monday.
      return (currentDay === "Monday"); 
    }
    else if (checkDays.includes("Tuesday")) { //Checks if the first day matches Tuesday.
      return (currentDay === "Tuesday"); 
    }
    return (currentDay === "Thursday"); //Checks if the first day matches Thursday. (At this point the only remaining true combination is "Thursday/Saturday" and "Thursday").
  }
  return true;
}


function CLEANUPINPUTS() { //Clears any active input fields.
  let fixInProgress = CHECKINPROGRESS();
  if (!fixInProgress) {
    FUNCTIONSTATUS(true);
  }
  else {
    sheet.getRange('F2:K2').setBackground("Red");
  }
  sheet.getRange('L29:M29').setValue(["",""]);
  sheet.getRange('L32:M32').setValue(["",""]);
  sheet.getRange('L40').setValue("Waiting for Query...");
  sheet.getRange('L42').setValue("");
  sheet.getRange('L47').setValue("");
  sheet.getRange('L50').setValue("");
  sheet.getRange('M2:M5').setValues([[""],[""],[""],[""]]);
  sheet.getRange('M8:M9').setValues([[""],[""]]);
  sheet.getRange('M12:M13').setValues([[""],[""]]);
  sheet.getRange('M16:M25').setValues([[""],[""],[""],[""],[""],[""],[""],[""],[""],[""]]);
  sheet.getRange('M36:M38').setValues([[""],[""],[""]]);
  sheet.getRange('O2:O5').setValues([[""],[""],[""],[""]]);
  sheet.getRange('O8:O11').setValues([[""],[""],[""],[""]]);
  sheet.getRange('O14:O17').setValues([[""],[""],[""],[""]]);
  sheet.getRange('O20:O27').setValues([[""],[""],[""],[""],[""],[""],[""],[""]]);
  sheet.getRange('O30:O32').setValues([[""],[""],[""]]);
  sheet.getRange('O35:O38').setValues([[""],[""],[""],[""]]);
  sheet.getRange('Q4:Q5').setValues([[""],[""]]);
  sheet.getRange('Q30').setValue("");
  sheet.getRange('Q33:Q38').setValues([[""],[""],[""],[""],[""],[""]]);
  if (!fixInProgress) {
    FUNCTIONSTATUS(false, true, "Inputs Cleared Successfully");
  }
}


function STARTUP(manual = true) {
  //Checks if the two main criteria for most functions to run are met.
  if (manual) {
    if (CHECKINPROGRESS()) { 
      return false; 
    }
    FUNCTIONSTATUS(true);
    if (GAPCHECK()) {
      FUNCTIONSTATUS(false, false, "Error: Gaps in Data");
      return false;
    }
  }
  return true;
}


function CHECKINPROGRESS() { //Checks if any script is currently running.
  return (sheet.getRange('G2').getBackground() === "#ffa500");
}


function GAPCHECK() { //Checks for any gaps in the data.
  let gapPotential = false;
  let gapIndex = sheet.getRange('A4');
  for (let gapCount = 0; gapCount < maximumRows; gapCount++) {
    if (gapIndex.getDisplayValue() === "" && gapPotential === false) { //Checks for when the data is blank.
      gapPotential = true;
    }
    else if (gapIndex.getDisplayValue() !== "" && gapPotential === true) { //Checks when blank data stops and a data entry emerges (There is a gap in the data).
      return true;
    }
    gapIndex = gapIndex.offset(1,0);
  }
  return false;
}


function FUNCTIONSTATUS(active = false, successCheck = false, outputMessage = "Error: FUNCTIONSTATUS Incorrectly Called") { //Updates the error logs in the spreadsheet.
  let statusCell = sheet.getRange('G2');
  let successCell = sheet.getRange('I2');
  let logCell = sheet.getRange('K2');
  let logColors = sheet.getRange('F2:K2');
  if (active) { //Runs if any script is running (sucessCheck and outputMessage are not important if so).
    statusCell.setValue("YES");
    successCell.setValue("Loading...");
    logCell.setValue("Loading...");
    logColors.setBackground("Orange");
  }
  else { //Runs once a script completes.
    if (successCheck) { //Runs if the script completed without any errors.
      successCell.setValue("Success");
      logColors.setBackground("Green");
    }
    else { //Runs if there is an error.
      successCell.setValue("Failure");
      logColors.setBackground("Red");
    }
    logCell.setValue(outputMessage);
    statusCell.setValue("NO");
  }
}
