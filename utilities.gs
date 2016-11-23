function formMule_checkForSourceChanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName();
  var activeRange = activeSheet.getActiveCell();
  var activeRow = activeRange.getRow();
  var emailConditionString = ScriptProperties.getProperty('emailConditions');
  if (emailConditionString) {
    var emailConditionObject = Utilities.jsonParse(emailConditionString);
    if (activeRow == 1) {
      var numSelected = ScriptProperties.getProperty("numSelected");
      for (var i = 0; i<numSelected; i++) {
        var sheetName = emailConditionObject['sht-'+i].trim();
        var sheet = ss.getSheetByName(sheetName);
        formMule_setAvailableTags(sheet);
      }
    }
  }
}


// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org

function formMule_institutionalTrackingUi() {
  NVSL.openInstitutionalTrackingUi();
}


function formMule_logCalEvent()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Auto-Created%20Calendar%20Event", scriptName, scriptTrackingId, systemName)
}

function formMule_logCalUpdate()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Auto-Updated%20Calendar%20Event", scriptName, scriptTrackingId, systemName)
}


function formMule_logSMS()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("SMS%20Message%20Sent", scriptName, scriptTrackingId, systemName)
}


function formMule_logVoiceCall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Voice%20Call%20Made", scriptName, scriptTrackingId, systemName)
}


function formMule_logEmail()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Mailed%20Templated%20Email", scriptName, scriptTrackingId, systemName)
}


function logRepeatInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function logFirstInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
}


function setSid() { 
  var scriptNameLower = scriptName.toLowerCase();
  var sid = ScriptProperties.getProperty(scriptNameLower + "_sid");
  if (sid == null || sid == "")
  {
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    ScriptProperties.setProperty(scriptNameLower + "_sid", ms_str);
    var uid = UserProperties.getProperty(scriptNameLower + "_uid");
    if (uid) {
      logRepeatInstall();
    } else {
      logFirstInstall();
      UserProperties.setProperty(scriptNameLower + "_uid", ms_str);
    }      
  }
}



function formMule_clearAllFlags() {
  formMule_clearMergeFlags();
  formMule_clearEventUpdateFlags();
  formMule_clearEventCreateFlags()
}


function formMule_clearSpecificTemplateMergeFlags() {
  var templateName = "First Reminder";  //Hand code the name of the template whose status you want deleted
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var emailConditionString = ScriptProperties.getProperty('emailConditions');  
  var emailConditionObject = Utilities.jsonParse(emailConditionString);    
  var numSelected = emailConditionObject['max'];
  var copyDownOn = ScriptProperties.getProperty('copyDownOption');
  var k = 2;
  if (copyDownOn=="true") {
    k=3;
  }
  var col = headers.indexOf(templateName + " Status") + 1;
  var lastRow = sheet.getLastRow();
  if ((col!=0)&&(lastRow>k-1)) {
    var range = sheet.getRange(k, col, lastRow-(k-1), 1).clear();
  }
}



function formMule_clearMergeFlags() {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var emailConditionString = ScriptProperties.getProperty('emailConditions');  
  if (emailConditionString) {
    var emailConditionObject = Utilities.jsonParse(emailConditionString);  
    var numSelected = emailConditionObject['max'];
  }
  var copyDownOn = ScriptProperties.getProperty('copyDownOption');
  var k = 2;
  if (copyDownOn=="true") {
    k=3;
  }
  for (var i=0; i<numSelected; i++) {
    var col = headers.indexOf(emailConditionObject['sht-'+i] + " Status") + 1;
    var lastRow = sheet.getLastRow();
    if ((col!=0)&&(lastRow>k-1)) {
      var range = sheet.getRange(k, col, lastRow-(k-1), 1).clear();
    }
  }
}



function formMule_clearEventUpdateFlags() {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  if ((sheetName)&&(sheetName!='')) {
    var sheet = ss.getSheetByName(sheetName);
    var headers = formMule_fetchHeaders(sheet);
    var eventUpdateIndex = headers.indexOf("Event Update Status");
    var lastRow = sheet.getLastRow();
    var copyDownOn = ScriptProperties.getProperty('copyDownOption');
    var k = 2;
    if (copyDownOn=="true") {
      k=3;
    }
    if ((eventUpdateIndex!=-1)&&(eventUpdateIndex>k-1)) {
      if (lastRow-(k-1)>0) {
        var range = sheet.getRange(k, eventUpdateIndex+1, lastRow-(k-1), 1).clear();
      }
    }
  }
}


function formMule_clearEventCreateFlags() {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  if ((sheetName)&&(sheetName!='')) {
    var sheet = ss.getSheetByName(sheetName);
    var headers = formMule_fetchHeaders(sheet);
    var eventUpdateIndex = headers.indexOf("Event Creation Status");
    var lastRow = sheet.getLastRow();
    var copyDownOn = ScriptProperties.getProperty('copyDownOption');
    var k = 2;
    if (copyDownOn=="true") {
      k=3;
    }
    if ((eventUpdateIndex!=-1)&&(eventUpdateIndex>k-1)) {
      if (lastRow-(k-1)>0) {
        var range = sheet.getRange(k, eventUpdateIndex+1, lastRow-(k-1), 1).clear();
      }
    }
  }
}

function rangetotable(input, headers) {
  var output = '<table cellpadding="8" cellspacing="0" style="border: 1px solid grey;">';
  var rows = input.length;
  var cols = input[0].length;
  if (headers) {
    if (headers[0].length != cols) {
      return "input and header arrays must be the same width";
    }
    output += "<tr>";
    for (var i=0; i<headers[0].length; i++) {
      output += "<th style=\"background-color:#DDDDDD; border: 1px solid grey;\">"+headers[0][i]+"</th>";
    }
    output += "</tr>";
  }
  for (var i=0; i<input.length; i++) {
    output += "<tr>";
    for (var j=0; j<input[0].length; j++) {
        output += "<td style=\"border: 1px solid grey;\">"+input[i][j]+"</td>";
      }
    output += "</tr>";
  }
  output += "</table>";
  return output;
}


function rangetoverticaltable(input, headers) {
  var output = '<table cellpadding="8" cellspacing="0" style="border: 1px solid grey;">';
  var rows = input[0].length;
  if (headers) {
    if (headers[0].length != rows) {
      return "input and header arrays must be the same width";
    }
  }
  for (var i=0; i<input[0].length; i++) {
    output += "<tr>";
     if (headers) {
      output += "<th style=\"background-color:#DDDDDD; border: 1px solid grey;\">"+headers[0][i]+"</th>";
    }
    output += "<td style=\"border: 1px solid grey;\">"+input[0][i]+"</td>";
    output += "</tr>";
  }
  output += "</table>";
  return output;
}