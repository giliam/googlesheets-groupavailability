var CAL_DAYS_LABEL = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

var CAL_MONTHS_LABEL = ['January', 'February', 'March', 'April',
                     'May', 'June', 'July', 'August', 'September',
                     'October', 'November', 'December'];
var CAL_DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

var YES = "yes".toLowerCase();
var NO = "no".toLowerCase();

var ROW_SUPPORT_MEMBERS = 4; 
var ROW_SUPPORT_DEFAULT = 6; 
var ROW_EMAILS = 8; 


/** 
 * onOpen
 * Called on the opening of the file.
 */
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('SupportMenu')
      .addItem('Add a month', 'chooseMonthDialog')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
  showSidebar();
}


/**
 * onEditTrigger
 * Adds an event to the edit cell event.
 */
function onEditTrigger(e){
  var range = e.range;
  var day = parseInt(range.getRow())-1;
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var current_sheet = active.getActiveSheet();
  if(current_sheet.getName() != "Utils"){
    var extract = extractYearMonth(current_sheet.getName());
    checkDay(extract[1],extract[0],day);
    Logger.log("Tested");
    active.setActiveRange(range);
  }
}


/** 
 * getEmails
 * Returns the emails specified in the Utils sheet.
 */
function getEmails(){
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var utils = active.getSheetByName("Utils");
  var emails = [];
  var i = 0;
  while(i<100){
    var r = utils.getRange(ROW_EMAILS, i+1).getValue();
    if(r=="") break;
    emails[i++] = r;
  }
  return emails;
}


/** 
 * getSupportMembers
 * Returns the list of support members specified in the Utils sheet.
 */
function getSupportMembers(){
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var utils = active.getSheetByName("Utils");
  var n_people = utils.getRange(2, 1, 1, 1).getValue();
  var support_members = [];
  for( var i=0;i<n_people;i++){
    support_members[i] = utils.getRange(ROW_SUPPORT_MEMBERS, i+1).getValue();
  }
  return support_members;
}

/** 
 * getMonths
 * Returns the list of months.
 */
function getMonths(){
  return CAL_MONTHS_LABEL;
}


/** 
 * getCurrentMonth
 * Returns the number of the current month (January is 0).
 */
function getCurrentMonth(){
  var today = new Date();
  return today.getMonth();
}


/** 
 * getYears
 * Returns the current year and the following one.
 */
function getYears(){
  var today = new Date();
  var year = today.getFullYear();
  return [year, year+1];
}


/**
 * extractYear
 * Returns the 
 */
function extractYearMonth(sheetname)
{
  var current_month = sheetname.substring(0, sheetname.length - 4);
  var year = sheetname.substring(sheetname.length-4);
  var month_id = -1;
  for(var key in CAL_MONTHS_LABEL) {
    if(CAL_MONTHS_LABEL[key] == current_month) {
      month_id = key;
    }
  }
  return [month_id,year];
}

/** 
 * available
 * @boolean available
 * @string who: name of the support member. Has to figure on the current sheet
 * Modifies the current sheet to set the availability of the member referenced by "who" to "available".
 */
function available(available,who,tomorrow)
{
  var today = new Date();
  if(tomorrow)
    today.setDate(today.getDate() + 1); 
  var day = today.getDate();
  var month = today.getMonth();
  var year = today.getFullYear();
  try{
    var active = SpreadsheetApp.getActiveSpreadsheet();
    var utils = active.getSheetByName("Utils");
    var n_people = utils.getRange(2, 1, 1, 1).getValue();
    var active_s = active.setActiveSheet(active.getSheetByName(CAL_MONTHS_LABEL[month] + "" + year));
    var n = n_people+2;
    for(var i=1;i<=n_people;i++){
      if(active_s.getRange(1, i+1).getValue()==who){
        n = i+1;
      }
    }
    if((n-n_people)==2){
      Logger.log("Support member not found on this month sheet!");
    }else{
      if(available){
        active_s.getRange(day+1, n,1,1).setValue(YES);
      }
      else{
        active_s.getRange(day+1, n,1,1).setValue(NO);
      }
      Logger.log("Status of " + who + " has been set to " + available.toString());
    }
  }
  catch(e){
    Logger.log(e);
    Logger.log("No sheet found for today's support availabilities (" + CAL_MONTHS_LABEL[month] + " " + year + ") !");
  }
}


/** 
 * showSidebar
 * Shows the sidebar base on the template file Month.html
 */
function showSidebar() {
  var template = HtmlService.createTemplateFromFile("Month.html");
  var html = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Support Sidebar')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}


/** 
 * chooseMonthDialog
 * Displays a dialog to choose year and month and then triggers the creation of a new sheet on this month.
 */
function chooseMonthDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Choose the month', 'Month:', ui.ButtonSet.YES_NO);

  if (response.getSelectedButton() == ui.Button.YES) {
    var month_name = response.getResponseText();
    var month_id = -1;
    for(var key in CAL_MONTHS_LABEL) {
      if(CAL_MONTHS_LABEL[key] == month_name) {
        month_id = key;
      }
    }
    if(month_id == -1){
      SpreadsheetApp.getUi().alert('Month not found (' + month_name + ').');
      Logger.log('Month not found (' + month_name + ').');
      chooseMonthDialog();
    }else{
      var year = 2008;//parseInt(chooseYearDialog());
      createNewMonth(month_id,year);
    }
  } else if (response.getSelectedButton() == ui.Button.NO) {
    Logger.log('The user didn\'t want to provide a name.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}



/** 
 * chooseYearDialog
 * Prompts the user for the year.
 */
function chooseYearDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Choose the year', 'Year:', ui.ButtonSet.YES_NO);

  if (response.getSelectedButton() == ui.Button.YES) {
    return response.getResponseText();
  } else if (response.getSelectedButton() == ui.Button.NO) {
  } else {
  }
}



/** 
 * createNewMonth
 * @integer month
 * @integer year
 * Creates a new sheet for the month and year sent by parameters.
 */
function createNewMonth(month,year) {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var utils = active.getSheetByName("Utils");
  try{
    active.setActiveSheet(active.getSheetByName(CAL_MONTHS_LABEL[month] + "" + year));
  }catch(e){
    var n_people = utils.getRange(2, 1, 1, 1).getValue();
    var _range = utils.getRange(ROW_SUPPORT_MEMBERS, 1, 1, n_people);
    var range_defaults_values = utils.getRange(ROW_SUPPORT_DEFAULT, 1, 1, n_people);
    var new_active = active.insertSheet();
    new_active.setName(CAL_MONTHS_LABEL[month] + "" + year);
    _range.copyTo(new_active.getRange(1, 2, 1, n_people));
    var n_days = CAL_DAYS_IN_MONTH[month];
    if( month ==1 && isBissextile(year)) { // if the year is bissextile
      n_days++;
    }
    for(var i=1;i<=n_days;i++)
    {
      var n = i+1;
      var cell = new_active.getRange("A" + n.toString());
      cell.setValue(i);
      range_defaults_values.copyTo(new_active.getRange("B" + n.toString()));
    }
  }
}



/** 
 * isBissextile
 * Determines if the year is bissextile (https://en.wikipedia.org/wiki/Leap_year#Algorithm)
 */
function isBissextile(year)
{
  return year%400==0 || ( year%100 != 0 && year%4 == 0 );
}


/** 
 * checkMonth
 * Checks if the month of the current sheet is scheduled by the support and if there is at least one person each evening.
 */
function checkMonth()
{
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var current_sheet = active.getActiveSheet();
  if(current_sheet.getName() == "Utils"){
    Logger.log("You are on the Utils spreadsheet.");
  }else{
    var extract = extractYearMonth(current_sheet.getName());
    var year = extract[1];
    var month_id = extract[0];
    if(month_id == -1){
      Logger.log("Month not found! ("+CAL_MONTHS_LABEL[month_id]+")");
    }else{
      for(var i=1;i<=CAL_DAYS_IN_MONTH[month_id];i++)
      {
        Logger.log("Checking availabilities for " + i + "/" + CAL_MONTHS_LABEL[month_id] + "/" + year + ".");
        checkDay(year,month_id,i);
      }
    }
  }
}


/** 
 * checkToday
 * Checks today's availabilities.
 */
function checkToday()
{
  Logger.log("Checking today availability.");
  var today = new Date();
  var day = today.getDate();
  var month = today.getMonth();
  var year = today.getFullYear();
  checkDay(year,month,day);
}


/** 
 * checkDay
 * @string year
 * @string month
 * @string day
 * Checks if at least one person is available today. Sends a mail otherwise.
 */
function checkDay(year,month,day)
{
  var active = SpreadsheetApp.getActiveSpreadsheet();
  try{
    if(active.getActiveSheet().getName() != CAL_MONTHS_LABEL[month] + "" + year)
      active.setActiveSheet(active.getSheetByName(CAL_MONTHS_LABEL[month] + "" + year));
    var utils = active.getSheetByName("Utils");
    var n_people = utils.getRange(2, 1, 1, 1).getValue();
    var _range = active.getActiveSheet().getRange(day+1, 1, 1, n_people);
    var full = true;
    var all_no = false;
    for(var i=1;i<=n_people;i++)
    {
      var cell = active.getActiveSheet().getRange(day+1,i+1,1,1);
      if(cell.getValue().toLowerCase() == YES) { // if someone stays
        all_no = true;
      }
      if(cell.getValue().toLowerCase() != YES && cell.getValue().toLowerCase() != NO) { //if one cell is not completed
        full = false;
      }
    }
    if(!all_no && full) {
       var emails = getEmails();
       for(var i=0;i<emails.length;i++){
         MailApp.sendEmail(emails[i],
                   "No support for "+day+" "+CAL_MONTHS_LABEL[month]+" "+year+"!",
                   "Where are the supports on the "+day+" "+CAL_MONTHS_LABEL[month]+" "+year+"?");
         Logger.log("Mail sent for "+day+"/"+CAL_MONTHS_LABEL[month]+"/"+year+" to " + emails[i]);
       }
    }
    Logger.log("Found ("+day+"/"+CAL_MONTHS_LABEL[month]+"/"+year+"): " + all_no.toString());
  }catch(e){
    Logger.log(e);
    Logger.log("No sheet found for " + day + "/" + CAL_MONTHS_LABEL[month] + "/" + year + "'s support availabilities (" + CAL_MONTHS_LABEL[month] + " " + year + ") !");
  }
}