// add project trigger for this script:
// time-based, minutes, every minute

function DeleteExpiredEntries() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Broadcasts");
  var datarange = sheet.getDataRange();
  var lastrow = datarange.getLastRow();
  var nowTimeStamp = new Date().getTime()/1000;
  
  for (i=lastrow;i>=2;i--) {
    var curRowTimeStamp = sheet.getRange(i, 1).getValue();
    var diff = nowTimeStamp - curRowTimeStamp
    var dur = 3600*sheet.getRange(i, 5).getValue(); 
    if(diff > dur) {
      sheet.deleteRow(i);
    }  
  }
}