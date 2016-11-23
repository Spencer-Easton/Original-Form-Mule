function formMule_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var propertyString = '';
  var excludeProperties = ['formmule_sid', 'caseNo', 'ssId', 'SSId', 'calendarToken', 'webAppUrl','twilioNumber','accountSID','lastPhone','authToken','preconfigStatus', 'ssId', 'ssKey', 'formulaTriggerSet'];
  for (var key in properties) {
    if (excludeProperties.indexOf(key)==-1) {
      var keyProperty = properties[key]; //.replace(/[/\\*]/g, "\\\\");                                     
      propertyString += "   ScriptProperties.setProperty('" + key + "','" + keyProperty + "');\n";
    }
  }
  var app = UiApp.createApplication().setHeight(500).setTitle("Export preconfig() settings");
  var panel = app.createVerticalPanel().setWidth("100%").setHeight("100%");
  var labelText = "Copying a Google Spreadsheet copies scripts along with it, but without any of the script settings saved.  This normally makes it hard to share full, script-enabled Spreadsheet systems. ";
  labelText += " You can solve this problem by pasting the code below into a script file called \"paste preconfig here\" (go to Script Editor and look in left sidebar) prior to publishing your Spreadsheet for others to copy. \n";
  labelText += " After a user copies your spreadsheet, they will select \"Run initial installation.\"  This will preconfigure all needed script settings.  If you got this workflow from someone as a copy of a spreadsheet, this has probably already been done for you.";
  var label = app.createLabel(labelText);
  var window = app.createTextArea();
  var codeString = "//This section sets all script properties associated with this formMule profile \n";
  codeString += "var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');\n";
  codeString += "if (preconfigStatus!='true') {\n";
  codeString += propertyString; 
  codeString += "};\n";
  codeString += "ScriptProperties.setProperty('preconfigStatus','true');\n";
  codeString += "var ss = SpreadsheetApp.getActiveSpreadsheet();\n";
  if (properties.formulaTriggerSet == "true") {
    codeString += "setCopyDownTrigger(); \n";
  }
  
 //generate msgbox warning code if automated email or calendar is enabled in template 
  if ((properties.calendarStatus == 'true')||(properties.emailStatus == 'true')) {
    codeString += "\n \n";
    codeString += "//Custom popup and function calls to prompt user for additional settings \n";
    if (ScriptProperties.getProperty('emailStatus')=="true") {
       codeString += "Browser.msgBox(\"You will want to check the email template sheets to ensure the correct sender and recipients are listed before using.\");\n";
    }
    if (ScriptProperties.getProperty('calendarStatus')=="true") {
    codeString += "Browser.msgBox(\"You need to set a new calendarID for this script before it will work.\");\n";
    codeString += "formMule_setCalendarSettings();";
    }
  }
  codeString += "ss.toast('Custom formMule preconfiguration ran successfully. Please check formMule menu options to confirm system settings.');\n";
  window.setText(codeString).setWidth("100%").setHeight("400px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}
