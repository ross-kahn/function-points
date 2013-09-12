function Properties(){
  this.alph = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
  
  this.fp_col = 3;
  
  this.fp_col_letter = this.alph[3];
  this.timestamp_col = this.alph[0];
  
  // The user has to specify which row number the current term starts in
  this.term_start_row_str = "Current Term Starts In Row #:";
  this.term_start_row_col = this.alph[this.fp_col + 8];
  
  this.latest_entry_str = "Latest Entry";
  this.latest_entry_col = this.alph[this.fp_col + 5];
 
  this.last_modified_str = "Last Modified";
  this.last_modified_col = this.alph[this.fp_col + 6];
  
  // Histories are used mainly to graph progress
  this.st_average_history_str = "Weekly Averages History";
  this.st_average_history_col = this.alph[this.fp_col + 9];  // A list of the term's weekly averages
  
  // The starting average for the current term. Must be manually entered
  this.starting_average_str = "Starting Term Average";
  this.starting_average_col = this.alph[this.fp_col + 10];
  
  // NOTE: THIS IS EQUIVALENT TO TAKING THE AVERAGE OF ALL RAW FP SCORES
  // An average of the long-term(lt) average history
  // Shows the average FP score based on the entire history of the position
  this.lt_average_str = "Total Cumulative Average";
  this.lt_average_col = this.alph[this.fp_col + 7];          
  
  // The current term average  
  // Shows the average FP score for the current term
  this.st_average_str = "Current Term Average";
  this.st_average_col = this.alph[this.fp_col + 1];          
  
  // Percent increase of st_average from starting term average
  // Shows on-average progress since the beginning of the term
  this.st_average_perc_str = "Term Average Increase";
  this.st_average_perc_col = this.alph[this.fp_col + 3];     
  
  // Percent increase of st_average from the lt_average
  // Shows progress in the context of long-term history
  this.lt_average_perc_str = "Total Average Increase";
  this.lt_average_perc_col = this.alph[this.fp_col + 4];
  
  // Percent increase of curFP from the st_average
  // Shows whether the current week was above or below average
  this.weekly_perc_str = "Current Week Increase";
  this.weekly_perc_col = this.alph[this.fp_col + 2];         
  
  var c = this.fp_col_letter;
  this.dataRange = c + '2:' + c;
  
  c = this.timestamp_col;
  this.timestampRange = c + ':' + c;
  
  c = this.term_start_row_col;
  this.termStartRange = c+'2';
  
  c = this.starting_average_col;
  this.startingAverageRange = c+'2';
}

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 *
 * NOTE: This is a debug function. Not used in the main functioning of the program
 */
function readRows() {
  var fPointIndex = 3;  
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  var functionPoints = new Array();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    //Logger.log(row);
  }
};

function cleanSheet(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(1, 5, 1000, 50).clearContent();
  makeHeaders(new Properties(), sheet);
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 *
 * NOTE: This is the main method of the program
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var properties = new Properties();
  
  // Adds a menu option on the spreadsheet
  var entries = [{
    name : "Calculate Data",
    functionName : "onOpen"
  },{
    name : "Clean Sheet",
    functionName: "cleanSheet"}];
  sheet.addMenu("Script Center Menu", entries);
  
  // Retrieves the raw weekly function point data
  var fpData = cleanArray(sheet.getRange(properties.dataRange).getValues());

  // Find the termStart, which must be set by the user. If it is not, default to termStart = 1
  var termStart = sheet.getRange(properties.termStartRange).getValue();
  if (termStart == "" || termStart < 2) {
    setTermStart(2, properties);
  }
  
  
  var startingAvg = sheet.getRange(properties.startingAverageRange).getValue();
  if (startingAvg == "" || startingAvg < 0) {
    // Set error to "N/A, Starting average not set";
    startingAvg = 0;
  }
  
  var totalAverage = longtermAverage(fpData);
  
  // Calculate the current term's average. If there's a problem with the
  // calculation (such as, termStart was not set by the user),
  var termAverage = shorttermAverage(fpData, termStart);
  if(termAverage == -1){
    termAverage = totalAverage;
  }
  
  var wAverages = weeklyAverages(fpData);
  var newestEntry = getNewestEntry(fpData);
  var lastModified = getLastModified(sheet, properties);
  
  var stAveragePercent = percentIncrease(startingAvg, termAverage)
  var ltAveragePercent = percentIncrease(totalAverage, termAverage)
  var weeklyPercent = percentIncrease(termAverage, newestEntry)
  
  
  //var tDate = new Date(getLastModified(sheet));
  //Logger.log(tDate.getMonth());

  sheet.getRange(properties.lt_average_col + "2").setValue(totalAverage);
  sheet.getRange(properties.weekly_perc_col + "2").setValue(weeklyPercent);
  sheet.getRange(properties.latest_entry_col + "2").setValue(newestEntry);
  sheet.getRange(properties.last_modified_col + "2").setValue(lastModified);
  sheet.getRange(properties.lt_average_perc_col + "2").setValue(ltAveragePercent);
  sheet.getRange(properties.st_average_perc_col + "2").setValue(stAveragePercent);
  sheet.getRange(properties.st_average_col + "2").setValue(termAverage);

  for(var i=2; i<wAverages.length+2; i++){
    sheet.getRange(properties.st_average_history_col+i).setValue(wAverages[i-2]);
  }

  makeHeaders(properties, sheet);
  
  //makeChart(sheet);  
};

function makeHeaders(properties, sheet){
  sheet.getRange(properties.weekly_perc_col + "1").setValue(properties.weekly_perc_str);
  sheet.getRange(properties.latest_entry_col + "1").setValue(properties.latest_entry_str);
  sheet.getRange(properties.last_modified_col + "1").setValue(properties.last_modified_str);
  sheet.getRange(properties.st_average_history_col + "1").setValue(properties.st_average_history_str);
  sheet.getRange(properties.lt_average_perc_col + "1").setValue(properties.lt_average_perc_str);
  sheet.getRange(properties.st_average_perc_col + "1").setValue(properties.st_average_perc_str);
  sheet.getRange(properties.st_average_col + "1").setValue(properties.st_average_str);
  sheet.getRange(properties.lt_average_col + "1").setValue(properties.lt_average_str);
  sheet.getRange(properties.term_start_row_col + "1").setValue(properties.term_start_row_str);
  sheet.getRange(properties.starting_average_col + "1").setValue(properties.starting_average_str);
}

// Takes an average of every function point in the document's history
function longtermAverage(data){
  //Logger.log(data);
  
  if(data.length < 1){
    return 0;
  }
  
  var sum = 0;
  for(var i=0; i < data.length; i++){
    sum += data[i];
  }
  return sum / data.length;
}

// Sets the term start property to a defined one
function setTermStart(termStart, properties){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(properties.termStartRange).setValue(termStart);
}

// WARNING: ASSUMES ALL DATA IS IN ORDER WITH NO BREAKS
function shorttermAverage(data, termStart){
  
  // Take the average of everything AFTER AND INCLUDING the termStart
  var sum = 0;
  var count = 0;
  for (var i=termStart-2; i<data.length; i++) {
    sum += data[i];
    count += 1;
  }
  
  if (count < 1) {
    return -1;
  }else{
    return sum / count;
  }
}

/**
 * A = latest weekly average
 * B = newest function point entry
 * return (B-A)/A
 * NOTE: check for A == 0
*/
function percentIncrease(startAvg, newest) {
  if( startAvg == 0 ){
    return newest * 100;
  }
  
  return ((newest - startAvg) / startAvg) * 100;
}


function weeklyAverages(fpData){
  var averages = new Array();
  var curSum = 0;
  var weekNum = 1;
  
  for(var i=0; i<fpData.length; i++){
    curSum += fpData[i];
    var wkAvg = curSum / weekNum;
    averages.push(wkAvg);
    weekNum++;
  }
  
  return averages;
}



function cleanArray(actual){
  var newArray = new Array();
  var data = "";
  
  for(var i = 0; i<actual.length; i++){
      data = actual[i][0];
      strTest = new String(data);
      if (strTest != ""){
        newArray.push(data);
        Logger.log(data);
    }
  }
  return newArray;
}

function getLastModified(sheet, properties){
  var c = properties.timestampRange;
  var data = cleanArray(sheet.getRange(c).getValues());
  
  if( data.length < 1 ){
   return "-No Value-"; 
  }
  
  return data[ data.length - 1];
}

function getNewestEntry(fpData){
  if( fpData.length < 1 ){
    return "-No Value-";
  }
  
  return fpData[ fpData.length - 1];
}

function makeChart(sheet) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var chart = null;
  
  if( sheet.getCharts().length > 0){
    return;
  }
  
  for(var i=0; i<sheet.getCharts().length; i++){
    if( sheet.getCharts()[i].getOptions().get("name") == "auto" ){
      chart = sheet.getCharts()[i];
    }
  }
  
  if( null == chart) {
    var chartBuilder = sheet.newChart().asLineChart();
    chartBuilder.setOption("name", "auto");
    chartBuilder.addRange(sheet.getRange("D:D"));
    chartBuilder.addRange(sheet.getRange("I:I"));
    chartBuilder.setPosition(3,1,1,1);
    chartBuilder.setTitle("Productivity over Time");
    chartBuilder.setXAxisTitle("Time");
    chart = chartBuilder.build();
    sheet.insertChart(chart);
  }
   
}