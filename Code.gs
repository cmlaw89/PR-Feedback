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
  var html = HtmlService.createTemplateFromFile('Date_Picker').evaluate().setHeight(300).setWidth(300);
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


function prFeedback(date) {
  //Opens the feedback sidebar
  
  date = new Date(date)
  var day = date.getDate()
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
  
  var users = getUsers();
  var user =  users[Session.getActiveUser().getEmail()];

  var extracted_data = getCases(user, month);
  var all_cases = extracted_data[0]
  var week_indexes = extracted_data[1]
  if (all_cases == false) {
    SpreadsheetApp.getUi().alert("You have no cases assigned for this month.", SpreadsheetApp.getUi().ButtonSet.OK);
  }  
  else {
    var cases = all_cases[day]
    if (cases == undefined) {
      SpreadsheetApp.getUi().alert("You have no cases assigned for this day.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
    else {
      var year = SpreadsheetApp.getActiveSpreadsheet().getName().split(" ")[2].slice(2, 4);
      
      //Get 8H word count for the user
      var week_no = getWeekNo(date)
      var word_counts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(months[month]).getRange(week_indexes[week_no - 1] + 2, 2, week_indexes[week_no] - week_indexes[week_no - 1], 2).getValues();
      word_counts = [].concat.apply([], word_counts);
      var user_word_count = word_counts[word_counts.indexOf(user) + 1]
      
      var html = HtmlService.createTemplateFromFile('Index');
      html.cases = cases;
      html.user_word_count = user_word_count
      html.all_cases = all_cases;
      html.user = user;
      html.month_year = "-" + pad(month+1) + "-" + year;
      html.month = months[month]
      html.date = date
      //html.editors = editors
      html = html.evaluate().setTitle("Proofreading Feedback");
      SpreadsheetApp.getUi().showSidebar(html);
    }
  }
}


function getCases(user, month) {
  //Searches the case ID columns of the active month sheet.
  //Returns an array of the case IDs for the specified proofreader (user)

  var today = new Date();
  var today_date = today.getDate();
  var today_month = today.getMonth();
   
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
                11: "December"}  
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month_sheet = ss.getSheetByName(months[month]);
  
  var cases = [];
  var num_rows = month_sheet.getMaxRows();
  
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
  
  
  for (var i = 0; i < week_indexes.length - 1; i++) {
    for (var j = 0; j < 7; j++) {
      cases = cases.concat(month_sheet.getRange(week_indexes[i] + 1, (7*j+5), week_indexes[i+1] - week_indexes[i], 4).getValues());
    }
  }
  
  //Find the index of tomorrows date if the current month is selected
  var date_index = 0;
  if (today_month == month) {
    var found = false;
    while (!found && date_index < cases.length) {
      if (cases[date_index][0] == today_date+1) {
        found = true;
      }
      else {
        date_index += 1;
      }
    }
  }
  else {
    date_index = cases.length;
  }
  
  var assigned_cases = {};
  
  var day_indexes = [];
  for (i = 0; i < week_indexes.length - 1; i++) {
    for (j = 0; j < 7; j++) {
      day_indexes.push(j*(week_indexes[i + 1]-week_indexes[i]) + 7*(week_indexes[i] - week_indexes[0]));
    }
  }
  
  var day = "";
  
  //Use regular expression to match case numbers (e.g., O001881 or O001881-1)
  var regex = new RegExp('^O[0-9]{5}.*[0-9]$')
  for (var i = 0; i < date_index; i++) {
    if (day_indexes.indexOf(i) != -1) {
      day = cases[i][0];
    }
    Logger.log(cases[i][0])
    if (regex.test(cases[i][0].toString().trim())&&(typeof cases[i][2] == "string")) {
      
      //Convert name to title case
      var proofreader_letters = cases[i][2].trim().split("");
      var proofreader = proofreader_letters[0].toUpperCase() + proofreader_letters.slice(1, proofreader_letters.length).join("").toLowerCase();
      
      //Create case dictionary
      if (Object.keys(assigned_cases).indexOf(proofreader) == -1) {
        assigned_cases[proofreader] = {};
        assigned_cases[proofreader][day] = [[cases[i][0].trim(), cases[i][3]]]
      }
      else {
        if (Object.keys(assigned_cases[proofreader]).indexOf(day.toString()) == -1) {
          assigned_cases[proofreader][day] = [[cases[i][0].trim(), cases[i][3]]]
        }
        else {
          assigned_cases[proofreader][day].push([cases[i][0].trim(), cases[i][3]]);
        }
      }
    }
  }
  
  if (Object.keys(assigned_cases).indexOf(user) != -1) {
    return [assigned_cases[user], week_indexes];
  }
  else {
    return [false, week_indexes];
  }
}


function getOutstanding(cases) {
  //Get list of cases with incomplete feedback for the selected month
  
  var TPR_Feedback_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PR Feedback");
  var complete = TPR_Feedback_sheet.getRange(2, 3, TPR_Feedback_sheet.getLastRow() - 1, 1).getValues();
  var completed_cases = [];
  for (i = 0; i < complete.length; i++) {
    completed_cases.push(complete[i][0].slice(0, 7));
  }
  
  var all_cases = [];
  for (var i = 0; i < Object.keys(cases).length; i++) {
    var day = Object.keys(cases)[i];
    all_cases = all_cases.concat(cases[day]);
  }
  
  var incomplete_cases = []
  for (i = 0; i < all_cases.length; i++) {
    if (completed_cases.indexOf(all_cases[i][0]) == -1) {
      incomplete_cases.push(all_cases[i]);
    }
  }
  
  return incomplete_cases
}

function getFeedback(proofreaders, editors) {
  if (typeof editors === "undefined") {
    editors = [];
  }
  
  var users = getUsers();
  var user =  users[Session.getActiveUser().getEmail()];
  proofreaders = [user]
  //Extracts the submitted feedback for the users in the list
  
  var PR_feedback_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PR Feedback");
  var all_feedback = PR_feedback_sheet.getRange(2, 2, PR_feedback_sheet.getLastRow(), 12).getValues();
  all_feedback = [].concat.apply([], all_feedback);
  var indexes = []
  for (var i = 0; i < proofreaders.length; i++) {
    indexes = indexes.concat(getAllIndexes(all_feedback, proofreaders[i]))
  }
  indexes = indexes.sort(function(a, b) {return a-b});
  var feedback = []
  for (var i = 0; i < indexes.length; i++) {
    var entry = all_feedback.slice(indexes[i] + 1, indexes[i] + 2)
                .concat(all_feedback.slice(indexes[i] + 11, indexes[i] + 12))
                .concat(all_feedback.slice(indexes[i] + 2, indexes[i] + 3))
                .concat(all_feedback.slice(indexes[i] + 4, indexes[i] + 11));
    entry = entry.map( function (x) {return x.toString()} );
    if (editors.length > 0) {
      if (editors.indexOf(entry[1]) != -1) {
        feedback.push(entry);
      }
    }
    else {
      feedback.push(entry);
    }
  }
  Logger.log(feedback.reverse())
  return feedback;
}

function getFeedbackCase(caseId) {
  //Returns the feedback that was submitted for the given case ID
  
  var users = getUsers();
  var user =  users[Session.getActiveUser().getEmail()];
  
  var cases = getFeedback([user]);
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
    users[emails[i][1]] = emails[i][0];
  }
  
  return users
}






function include(filename) {
  //Adds stylesheet and javascript to Index.html
  
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function pad(n) {
    return (n < 10) ? ("0" + n) : n;
}

function getAllIndexes(arr, val) {
    var indexes = [], i = -1;
    while ((i = arr.indexOf(val, i+1)) != -1){
        indexes.push(i);
    }
    return indexes;
}


function getWeekNo(date) {
  
  var day = date.getDate()
  
  //get weekend date
  day += (date.getDay() == 0 ? 0 : 7 - date.getDay());
  
  return Math.ceil(parseFloat(day) / 7)
}


//Add New Month
///////////////


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

function daysInMonth (month, year) {
    return new Date(year, month, 0).getDate();
}
