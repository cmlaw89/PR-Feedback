function onOpen() {
  //Adds the feedback menu to the Translation Proofreading Sheet (Schedule)
  
  SpreadsheetApp.getUi()
      .createMenu('Wallace')
      .addItem('Submit Feedback', 'openDatePicker')
      .addItem('View Feedback', 'viewFeedback')
      .addItem("New Month", 'makeMonth')
      .addToUi();
}

function openDatePicker() {
  var users = getUsers();
  var user =  users[Session.getActiveUser().getEmail().toLowerCase()];
  var html = HtmlService.createTemplateFromFile('Date_Picker')
  html.user = user;
  html = html.evaluate().setHeight(300).setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Please select the feedback date.");
}

function viewFeedback() {
  var html = HtmlService.createTemplateFromFile("View_Feedback_Index");
  var database_editors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PR Feedback").getRange("M:M").getValues();
  var editors = []
  for (var i = 0; i < database_editors.length; i++) {
    if (editors.indexOf(database_editors[i][0]) == -1) {
      editors.push(database_editors[i][0]);
    }
  }
  html.editors = editors;
  html = html.evaluate()
  .setTitle("View Feedback")
  .setHeight(450)
  .setWidth(750)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  
  SpreadsheetApp.getUi().showModalDialog(html, "Submitted Feedback")
}

function prFeedback(date, user) {
  //Opens the feedback sidebar
  
  date = new Date(date)
  var month = date.getMonth()
  
  var months = {0: "January", 
                1: "February",
                2: "March",
                3: "April",
                4: "May",
                5: "June",
                6: "July",
                7: "August",
                8: "September",
                9: "October",
                10: "November",
                11: "December"}
  

  var extracted_data = getCases(user, date);
  if (extracted_data) {
    var cases = extracted_data[0];
    var word_count = extracted_data[1];
    var year = SpreadsheetApp.getActiveSpreadsheet().getName().split(" ")[2].slice(2, 4);
    
    var html = HtmlService.createTemplateFromFile('Index');
    html.cases = cases;
    html.user_word_count = word_count
    html.user = user;
    html.month_year = "-" + pad(month+1) + "-" + year;
    html.month = months[month]
    html.date = date
    html = html.evaluate().setTitle("Proofreading Feedback");
    SpreadsheetApp.getUi().showSidebar(html);
  }
  else {
    SpreadsheetApp.getUi().alert("You have no cases assigned for this day.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getCases(user, date) {
  //Searches the case ID columns of the month sheet for the chosen date.
  //Returns an array of the case IDs and editors for the specified proofreader (user)
   
  var months = {0: "January", 
                1: "Februay",
                2: "March",
                3: "April",
                4: "May",
                5: "June",
                6: "July",
                7: "August",
                8: "September",
                9: "October",
                10: "November",
                11: "December"};
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month_sheet = ss.getSheetByName(months[date.getMonth()]);
  
  var cases = [];
  var word_counts = [];
  
  //Find the start row for each week in the schedule sheet
  var week_indexes = []
  var column_A_values = month_sheet.getRange("A:A").getValues();
  var sheet_height = column_A_values.length;
  for (var i = 0; i < sheet_height; i++) {
    if (column_A_values[i][0] == "Date") {
      week_indexes.push(i);
    }
  }
  week_indexes.push(sheet_height)
  
  //Loop over days and week indexes
  //Use quotient (Math.floor) and remainder (%) on division by 7 to index rows and columns
  //Add cases and word counts for the selected date
  var found = false;
  var i = 0;
  while ((found == false)&&(i < 41)) {
    var week = Math.floor(i / 7);
    var column = (7 * (i % 7)) + 5;
    if (month_sheet.getRange(week_indexes[week] + 1, column, 1, 1).getValues()[0][0] == date.getDate()) {
      found = true
      cases = cases.concat(month_sheet.getRange(week_indexes[week] + 1, column, week_indexes[week+1] - week_indexes[week], 4).getValues());
      word_counts = word_counts.concat(month_sheet.getRange(week_indexes[week] + 1, 2, week_indexes[week + 1] - week_indexes[week], 2).getValues());
    }
    i += 1;
  }
  
  //Unwrap word counts extract word count for user
  word_counts = [].concat.apply([], word_counts);
  var word_count = word_counts[word_counts.indexOf(user) + 1];

  //Unwrap cases and map proofreader names to title case
  var cases = [].concat.apply([], cases.map(
                        function (entry) {
                          var proofreader_letters = entry[2].toString().trim().split("");
                          if (proofreader_letters.length > 0) {
                            var proofreader = proofreader_letters[0].toUpperCase() + proofreader_letters.slice(1, proofreader_letters.length).join("").toLowerCase();
                            return [entry[0], entry[1], proofreader, entry[3]];
                          }
                        }));
  
  //Add assigned cases for user (use regex to check case ID)
  var assigned_cases = [];
  var indexes = getAllIndexes(cases, user);
  var regex = new RegExp('^O[0-9]{6}$')
  for (var i = 0; i < indexes.length; i++) {
    if (regex.test(cases[indexes[i] - 2].toString().trim())) {
      assigned_cases.push([cases[indexes[i] - 2].toString().trim(), cases[indexes[i] + 1]]);
    }
  }
  
  if (assigned_cases.length > 0) {
    return [assigned_cases, word_count];
  }
  else {
    return false;
  }
}

function getFeedback() {
  
  var users = getUsers();
  var user =  users[Session.getActiveUser().getEmail().toLowerCase()];
  //Extracts the submitted feedback for the users in the list
  
  var PR_feedback_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PR Feedback");
  var all_feedback = PR_feedback_sheet.getRange(2, 2, PR_feedback_sheet.getLastRow(), 12).getValues();
  all_feedback = [].concat.apply([], all_feedback);
  var indexes = getAllIndexes(all_feedback, user);
  var feedback = []
  for (var i = 0; i < indexes.length; i++) {
    if (indexes[i] % 12 == 0) {
      var entry = all_feedback.slice(indexes[i] + 1, indexes[i] + 2)
                  .concat(all_feedback.slice(indexes[i] + 11, indexes[i] + 12))
                  .concat(all_feedback.slice(indexes[i] + 2, indexes[i] + 3))
                  .concat(all_feedback.slice(indexes[i] + 4, indexes[i] + 11));
      entry = entry.map( function (x) {return x.toString()} );
      feedback.push(entry);
    }
  }
  return feedback.reverse();
}

function getFeedbackCase(caseId) {
  //Returns the feedback that was submitted for the given case ID
  
  var users = getUsers();
  var user =  users[Session.getActiveUser().getEmail().toLowerCase()];
  
  var cases = getFeedback();
  cases = [].concat.apply([], cases);
  var index = cases.indexOf(caseId);
  if (index != -1) {
    return cases.slice(index, index + 10)
  }
}

function submitFeedback(values) {
  
  //Delete existing feedback with the same case ID
  var PR_feedback_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PR Feedback");
  var case_ID = values[2];
  var case_ID_column = PR_feedback_sheet.getRange(2, 3, PR_feedback_sheet.getLastRow() - 1, 1).getValues();
  var row_index = -1;
  var i = 0;
  while (row_index == -1 && i < case_ID_column.length) {
    if (case_ID_column[i][0] == case_ID) {
      row_index = i;
    }
    i += 1;
  }
  if (row_index != -1) {
    PR_feedback_sheet.deleteRow(row_index + 2);
  }
  
  //Add the new feedback
  PR_feedback_sheet.getRange(PR_feedback_sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
}

function getUsers() {
  // Creates a dictionary of users' names and email addresses (e.g., users = {"adamhuang@wallace.tw": "Adam", ...})
  
  var list_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proofreaders");
  var emails = list_sheet.getRange(3, 1, list_sheet.getLastRow()-2, 2).getValues();
  var users = {};
  for (var i = 0; i < emails.length; i++) {
    users[emails[i][1].toLowerCase()] = emails[i][0];
  }
  
  return users
}



//Add New Month
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////

function makeMonth() {
  var ui = SpreadsheetApp.getUi();
  
  var prompt = ui.prompt("Please enter the month name (e.g., August):", ui.ButtonSet.OK_CANCEL)
  
  if (prompt.getSelectedButton() == ui.Button.OK) {
    var months = {
        "January" :0,
        "February" :1,
        "March" :2,
        "April" :3,
        "May" :4,
        "June" :5,
        "July" :6,
        "August" :7,
        "September" :8,
        "October" :9,
        "November" :10,
        "December" :11
      };
    var month_name = prompt.getResponseText();
    if (Object.keys(months).indexOf(month_name) != -1) {
      var month = months[month_name];
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheets = ss.getSheets();
      var sheet_names = [];
      for (var i = 0; i < sheets.length; i++) {
        sheet_names.push(sheets[i].getName());
      }
      if (sheet_names.indexOf(month_name) == -1) {
        var template = ss.getSheetByName("Template");
        var new_month = template.copyTo(ss)
        new_month.setName(month_name)
        setDates(month, new_month);
        
        //Set permissions
        var all_editors = ss.getSheetByName("Proofreaders").getRange(2, 2, ss.getSheetByName("Proofreaders").getMaxRows(), 4).getValues();
        var protection = new_month.protect()
        Logger.log(all_editors.length)
        for (var i = 0; i < all_editors.length; i++) {
          if (all_editors[i][3] == "Y") {
            protection.removeEditor(all_editors[i][0]);
          }
        }
        protection.addEditor("felix@wallace.tw").setUnprotectedRanges([new_month.getRange("D:D"),
                                                                       new_month.getRange("K:K"), 
                                                                       new_month.getRange("R:R"), 
                                                                       new_month.getRange("Y:Y"), 
                                                                       new_month.getRange("AF:AF"), 
                                                                       new_month.getRange("AM:AM"), 
                                                                       new_month.getRange("AT:AT")]);
        
      }
      else {
        ui.alert("The sheet for this month has already been created.");
      }
    }
    else {
      ui.alert("The entered month was not correct. Please try again.");
    }
  }
}

function setDates(month, sheet) {
  //Find the start row for each week in the schedule sheet
  var week_indexes = []
  var column_A_values = sheet.getRange("A:A").getValues();
  var sheet_height = column_A_values.length;
  for (var i = 0; i < sheet_height; i++) {
    if (column_A_values[i][0] == "Date") {
      week_indexes.push(i);
    }
  }
  week_indexes.push(sheet_height)

  //Add dates to the template 
  var date = new Date();
  date.setMonth(month);
  date.setDate(1);
  var bodge = [1, 2, 3, 4, 5, 6, 0];
  var start_day = bodge.indexOf(date.getDay())
  for (var i = 0; i < week_indexes.length - 1; i++) {
    for (var j = 0; j < 7; j++) {
      if (j + 7*i < start_day) {
        sheet.getRange(week_indexes[i] + 2, 5 + 7*j, week_indexes[i + 1] - week_indexes[i] - 3, 6).setBackground("#d9d9d9");
      }
      else if ((j + 7*i >= start_day)&&(j + 7*i < start_day + daysInMonth(month + 1, date.getFullYear()))) {
        sheet.getRange(week_indexes[i] + 1, 5 + 7*j, 1, 1).setValue((j + 7*i - start_day + 1).toString())
        var end_row = i
      }
      else if (j + 7*i >= start_day + daysInMonth(month + 1, date.getFullYear())) {
        sheet.getRange(week_indexes[i] + 2, 5 + 7*j, week_indexes[i + 1] - week_indexes[i] - 3, 6).setBackground("#d9d9d9");
      }
    }
  }
}



//Auxiliary functions
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////

function include(filename) {
  //Adds stylesheet and javascript to Index.html
  
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function pad(n) {
  //Adds a leading zero to single digit numbers
  return (n < 10) ? ("0" + n) : n;
}

function getAllIndexes(arr, val) {
  //Returns all indices of a value (val) in a 1D array (arr)
  var indexes = [], i = -1;
  while ((i = arr.indexOf(val, i+1)) != -1){
    indexes.push(i);
  }
  return indexes;
}

function getWeekNo(date) {
  //Returns the week number in the month for the specified date
  //Weeks are defined as Mon -- Sun
  var day = date.getDate()
  day += (date.getDay() == 0 ? 0 : 7 - date.getDay());
  
  return Math.ceil(parseFloat(day) / 7)
}

function daysInMonth (month, year) {
  //Returns the number of days in the month for the specified year
  return new Date(year, month, 0).getDate();
}
