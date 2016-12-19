var scriptTitle = "formMule Script V6.5.3 (1/30/14)";
var scriptName = 'formMule'
var scriptTrackingId = 'UA-30976195-1'
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Support and contact at http://www.youpd.org/formmule
// 

var ss = SpreadsheetApp.getActiveSpreadsheet();
var MULEICONURL = 'https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/formMule6_icon.gif?attachauth=ANoY7cpzkYpzCxNW5Xfg-odHtcod8xUm5hP7l1Q7JgEFVbem83zgoz0mAq40-26SBZM7LlpODEluBzq6-0j9I51YFHcV1ex-DYen2UyYg7mXfuSJkcqJFE5T9yD9La27lk1Wh3oBwjeozxJLUQIuFWPd2dSTSs_eFF8v-t4EYrEJ1bJjifRHLalWrZQUilFs9HjkNWh_1x7IhGgqNWhZdDN_PAUZn1Dd0niCuUtX4gTcFl6obSZ-dFBuCHB0IEg3TjEVtKtZbKGj&attredirects=0';

function onInstall () {
  Browser.msgBox('To complete initialization, please select \"Run initial installation\" from the formMule script menu above');
  var menuEntries = [];
      menuEntries.push({name: "What is formMule?", functionName: "formMule_whatIs"});
      menuEntries.push({name: "Run initial installation", functionName: "formMule_completeInstall"});
  onOpen();
}

function onOpen() {
  var menuEntries = [];
  var installed = ScriptProperties.getProperty('installedFlag');
  var sheetName = ScriptProperties.getProperty('sheetName');
  var webAppUrl = ScriptProperties.getProperty('webAppUrl');
  var twilioNumber = ScriptProperties.getProperty('twilioNumber');
  if (!(installed)) {
      menuEntries.push({name: "Run initial installation", functionName: "formMule_completeInstall"});
  } else {
      menuEntries.push({name: "What is formMule?", functionName: "formMule_whatIs"});
      menuEntries.push({name: "Step 1: Define merge source settings", functionName: "formMule_defineSettings"});
    if ((sheetName) && (sheetName!='')) {
      menuEntries.push({name: "Step 2a: Set up email merge", functionName: "formMule_emailSettings"});
      menuEntries.push({name: "Step 2b: Set up calendar merge", functionName: "formMule_setCalendarSettings"});
      if (!webAppUrl) {
        menuEntries.push({name: "Step 2c: Set up SMS and Voice Message merge", functionName: "formMule_howToPublishAsWebApp"});
      }
      if ((webAppUrl)&&(!twilioNumber)) {
        menuEntries.push({name: "Step 2c: Set up SMS and Voice Message merge", functionName: "formMule_howToSetUpTwilio"});
      }
      if ((webAppUrl)&&(twilioNumber)) {
        menuEntries.push({name: "Step 2c: Set up SMS and Voice Message merge", functionName: "formMule_smsAndVoiceSettings"});
      }     
      menuEntries.push({name: "Preview and perform manual merge", functionName: "formMule_loadingPreview"});
      menuEntries.push({name: "Advanced options", functionName: "formMule_advanced"});
    }
  }
  this.ss.addMenu("formMule", menuEntries);
}

function formMule_advanced() {
  var app = UiApp.createApplication().setTitle("Advanced options").setHeight(130).setWidth(290);
  var quitHandler = app.createServerHandler('formMule_quitUi');
  var handler1 = app.createServerHandler('formMule_detectFormSheet');
  var button1 = app.createButton('Copy down formulas on form submit').addClickHandler(quitHandler).addClickHandler(handler1);
  var handler2 = app.createServerHandler('formMule_extractorWindow');
  var button2 = app.createButton('Package this workflow for others to copy').addClickHandler(quitHandler).addClickHandler(handler2);
  var handler3 = app.createServerHandler('formMule_institutionalTrackingUi');
  var button3 = app.createButton('Manage your usage tracker settings').addClickHandler(quitHandler).addClickHandler(handler3);
  var panel = app.createVerticalPanel();
  panel.add(button1);
  panel.add(button2);
  panel.add(button3);
  app.add(panel);
  ss.show(app);
  return app;
}

function formMule_quitUi(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}


function formMule_completeInstall() {
  formMule_preconfig();
  ScriptProperties.setProperty('installedFlag', 'true');
  var triggers = ScriptApp.getScriptTriggers();
  var formTriggerSetFlag = false;
  var editTriggerSetFlag = false;
  for (var i = 0; i<triggers.length; i++) {
    var eventType = triggers[i].getEventType();
    var triggerSource = triggers[i].getTriggerSource();
    var handlerFunction = triggers[i].getHandlerFunction();
    if ((handlerFunction=='formMule_onFormSubmit')&&(eventType=="ON_FORM_SUBMIT")&&(triggerSource=="SPREADSHEETS")) {
      formTriggerSetFlag = true;
    }
    if ((handlerFunction=='formMule_checkForSourceChanges')&&(eventType=="ON_EDIT")&&(triggerSource=="SPREADSHEETS")) {
      editTriggerSetFlag = true;
    }
  }
  if (formTriggerSetFlag==false) {
    formMule_setFormTrigger();
  }
  if (editTriggerSetFlag==false) {
    formMule_setEditTrigger();
  }
//ensure console and readme sheets exist

 onOpen();
}

function formMule_manualSend () {
  var lock = LockService.getPublicLock();
  lock.waitLock(120000);
  var manual = true;
  formMule_sendEmailsAndSetAppointments(manual);
  lock.releaseLock();
}


function formMule_setFormTrigger() {
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptApp.newTrigger('formMule_onFormSubmit').forSpreadsheet(ssKey).onFormSubmit().create();
}

function formMule_setEditTrigger() {
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptApp.newTrigger('formMule_checkForSourceChanges').forSpreadsheet(ssKey).onEdit().create();
}

function formMule_lockFormulaRow() {
// setFrozenRows function was once broken in Apps Script...seeing if it's fixed.
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetName = ScriptProperties.getProperty('sheetName');
var sheet = ss.getSheetByName(sheetName);
sheet.setFrozenRows(2);
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetName = ScriptProperties.getProperty('sheetName');
var sheet = ss.getSheetByName(sheetName)
ss.setActiveSheet(sheet);
  var frozenRows = sheet.getFrozenRows();
  if (frozenRows!=2) {
    Browser.msgBox("To avoid issues, it is highly recommended you freeze the first two rows of this sheet.");
  }
}


function formMule_setCalendarSettings () {
  var calendarStatus = ScriptProperties.getProperty('calendarStatus');
  var calendarUpdateStatus = ScriptProperties.getProperty('calendarUpdateStatus');
  var calendarToken  = ScriptProperties.getProperty('calendarToken');
  var eventTitleToken = ScriptProperties.getProperty('eventTitleToken');
  var locationToken = ScriptProperties.getProperty('locationToken');
  var guests = ScriptProperties.getProperty('guests');
  var emailInvites = ScriptProperties.getProperty('emailInvites');
  var allDay = ScriptProperties.getProperty('allDay');
  var startTimeToken = ScriptProperties.getProperty('startTimeToken');
  var endTimeToken = ScriptProperties.getProperty('endTimeToken');
  var descToken = ScriptProperties.getProperty('descToken');
  var reminderType = ScriptProperties.getProperty('reminderType');
  var minBefore = ScriptProperties.getProperty('minBefore');
 
  var app = UiApp.createApplication().setTitle('Step 2b: Set up calendar merge').setHeight(450).setWidth(800);
  var panel = app.createHorizontalPanel().setId('calendarPanel').setSpacing(20).setStyleAttribute('borderColor', 'grey');
 // var helpLabel = app.createLabel().setText('Indicate whether a calendar event should be generated on form submission, and (optionally) set an additional condition below. Fields for calendar id, and date fields can be populated with static values or dynamically using the variables to the right.');
 // var popupPanel = app.createPopupPanel();
 // popupPanel.add(helpLabel);
  var scrollPanel = app.createScrollPanel().setHeight("230px");
  var verticalPanel = app.createVerticalPanel();
  var topSettingsGrid = app.createGrid(4, 2).setId('topSettingsGrid');
  
  // Create calendar event conditions grid
  var conditionsGrid = app.createGrid(2,3).setId('conditionsGrid').setCellPadding(0);
  var conditionLabel = app.createLabel('Create Event Condition');
  var dropdown = app.createListBox().setId('col-0').setName('col-0').setWidth("150px").setEnabled(false);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = ScriptProperties.getProperty('sheetName');
  if ((sourceSheetName)&&(sourceSheetName!='')) {
     var sourceSheet = ss.getSheetByName(sourceSheetName);
  } else {
    Browser.msgBox('You must select a source sheet before this menu item can be selected.');
    formMule_defineSettings();
    app().close();
    return app;
  }
  var lastCol = sourceSheet.getLastColumn();
  if (lastCol > 0) {
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  } else {
    Browser.msgBox('You must have headers in your source data sheet before this menu item can be selected.');
    formMule_defineSettings();
    app().close();
    return app;
  }
  for (var i=0; i<headers.length; i++) {
    dropdown.addItem(headers[i]);
  }
  var equalsLabel = app.createLabel('=');
  var textbox = app.createTextBox().setId('val-0').setName('val-0').setEnabled(false);
  var conditionHelp = app.createLabel('Leave blank to always create an event. Use NULL for empty.  NOT NULL for not empty.').setStyleAttribute('fontSize', '10px');
  conditionsGrid.setWidget(0, 0, conditionLabel);
  conditionsGrid.setWidget(1, 0, dropdown);
  conditionsGrid.setWidget(1, 1, equalsLabel);
  conditionsGrid.setWidget(1, 2, textbox);
  conditionsGrid.setWidget(0, 2, conditionHelp);
  
  var extraHelp1 = app.createLabel('Note: Events are created only for rows where the condition is met AND \"Event Creation Status\" is blank').setStyleAttribute('fontSize', '10px').setStyleAttribute('color', '#0099FF').setStyleAttribute('fontWeight', 'bold').setId('extraHelp1').setVisible(false);
  var extraHelp2 = app.createLabel('Note: Events are updated or deleted only for rows where the condition is met AND \"Event Update Status\" is blank').setStyleAttribute('fontSize', '10px').setStyleAttribute('color', '#66C285').setStyleAttribute('fontWeight', 'bold').setId('extraHelp2').setVisible(false);
 
  var createPanel = app.createVerticalPanel();
  createPanel.add(conditionsGrid);
  createPanel.add(extraHelp1);
  
  var calendarConditionString = ScriptProperties.getProperty('calendarConditions');
  if ((calendarConditionString)&&(calendarConditionString!='')) {
    var calendarConditionsObject = Utilities.jsonParse(calendarConditionString);
    var selectedHeader = calendarConditionsObject['col-0'];
    var selectedIndex = 0;
    for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedHeader) {
        selectedIndex = i;
        break;
      }
    }
    dropdown.setSelectedIndex(selectedIndex);
    var selectedValue = calendarConditionsObject['val-0'];
    textbox.setValue(selectedValue);
  }
  
  //Update calendar conditions grid
  var updateConditionsGrid = app.createGrid(4,3).setId('updateConditionsGrid');
  var updateConditionLabel = app.createLabel('Update Event Condition');
  var deleteConditionLabel = app.createLabel('Delete Event Condition');
  var updateDropdown = app.createListBox().setId('update-col-0').setName('update-col-0').setWidth("150px").setEnabled(false);
  for (var i=0; i<headers.length; i++) {   
    updateDropdown.addItem(headers[i]);
  }
  
  var deleteDropdown = app.createListBox().setId('delete-col-0').setName('delete-col-0').setWidth("150px").setEnabled(false);
  for (var i=0; i<headers.length; i++) {
    deleteDropdown.addItem(headers[i]);
  }
  var equalsLabel = app.createLabel('if');
  var equalsLabel2 = app.createLabel('=');
  var updateTextbox = app.createTextBox().setId('update-val-0').setName('update-val-0').setEnabled(false);
  var deleteTextbox = app.createTextBox().setId('delete-val-0').setName('delete-val-0').setEnabled(false);
  var conditionHelp = app.createLabel('Leave blank to update no events. Use \"NULL\" for empty.  \"NOT NULL\" for not empty.').setStyleAttribute('fontSize', '10px');
  var deleteConditionHelp = app.createLabel('Leave blank to delete no events. Use \"NULL\" for empty.  \"NOT NULL\" for not empty.').setStyleAttribute('fontSize', '10px');
  updateConditionsGrid.setWidget(0, 0, updateConditionLabel);
  updateConditionsGrid.setWidget(1, 0, updateDropdown);
  updateConditionsGrid.setWidget(1, 1, equalsLabel);
  updateConditionsGrid.setWidget(1, 2, updateTextbox);
  updateConditionsGrid.setWidget(2, 0, deleteConditionLabel);
  updateConditionsGrid.setWidget(3, 0, deleteDropdown);
  updateConditionsGrid.setWidget(3, 1, equalsLabel2);
  updateConditionsGrid.setWidget(3, 2, deleteTextbox);
  updateConditionsGrid.setWidget(0, 2, conditionHelp);
  updateConditionsGrid.setWidget(2, 2, deleteConditionHelp);
  
  var calendarUpdateConditionString = ScriptProperties.getProperty('calendarUpdateConditions');
  if ((calendarUpdateConditionString)&&(calendarUpdateConditionString!='')) {
    var calendarUpdateConditionsObject = Utilities.jsonParse(calendarUpdateConditionString);
    var selectedUpdateHeader = calendarUpdateConditionsObject['col-0'];
    var selectedIndex = 0;
    for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedUpdateHeader) {
        selectedIndex = i;
        break;
      }
    }
    updateDropdown.setSelectedIndex(selectedIndex);
    var selectedValue = calendarUpdateConditionsObject['val-0'];
    updateTextbox.setValue(selectedValue);
  }
  
  var calendarDeleteConditionString = ScriptProperties.getProperty('calendarDeleteConditions');
  if ((calendarDeleteConditionString)&&(calendarDeleteConditionString!='')) {
    var calendarDeleteConditionsObject = Utilities.jsonParse(calendarDeleteConditionString);
    var selectedDeleteHeader = calendarDeleteConditionsObject['col-0'];
    var selectedIndex = 0;
    for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedDeleteHeader) {
        selectedIndex = i;
        break;
      }
    }
    deleteDropdown.setSelectedIndex(selectedIndex);
    var selectedValue = calendarDeleteConditionsObject['val-0'];
    deleteTextbox.setValue(selectedValue);
  }
  

  var eventIdPanel = app.createVerticalPanel();
  var eventIdLabel = app.createLabel('Column containing Event Id to be used for update');
  var eventIdCol = app.createListBox().setId('eventIdCol').setName('eventIdCol').setWidth("150px").setEnabled(false);
  for (var i=0; i<headers.length; i++) {
    eventIdCol.addItem(headers[i]);
  } 
  var selectedEventIdHeader = ScriptProperties.getProperty('eventIdCol');
  var eventIdExists = headers.indexOf("Event Id");
  if (eventIdExists==-1) {
    eventIdCol.addItem('Event Id');
  } 
  if ((!(selectedEventIdHeader))||(selectedEventIdHeader=='')) {
    selectedEventIdHeader = "Event Id";
  }
  
  var selectedEventIdUpdateIndex = 0;
  for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedEventIdHeader) {
        selectedEventIdUpdateIndex = i;
        break;
      }
    }
  eventIdCol.setSelectedIndex(selectedEventIdUpdateIndex);
  
  
  eventIdPanel.add(eventIdLabel);
  eventIdPanel.add(eventIdCol);
  
  var settingsGrid = app.createGrid(14, 2).setId('settingsGrid').setWidth(520);
  var calendarLabel = app.createLabel().setText('Calendar Id (xyz@sample.org)');
  var calendarTextBox = app.createTextBox().setName('calendarToken').setWidth("100%");
  if (calendarToken) { calendarTextBox.setValue(calendarToken); }
  var eventTitleLabel = app.createLabel().setText('Event title');
  var eventTitleTextBox = app.createTextBox().setName('eventTitleToken').setWidth("100%");
  if (eventTitleToken) { eventTitleTextBox.setValue(eventTitleToken); }
  var eventLocationLabel = app.createLabel().setText('Location');
  var eventLocationTextBox = app.createTextBox().setName('locationToken').setWidth("100%");
  if (locationToken) { eventLocationTextBox.setValue(locationToken); }
  var guestsLabel = app.createLabel().setText('Guests (comma separated email addresses)');
  var guestsTextBox = app.createTextBox().setName('guests').setWidth("100%");
  if (guests) { guestsTextBox.setValue(guests);  }
  var emailInvitesCheckBox = app.createCheckBox().setText('Email invitations').setName('emailInvites');
  if (emailInvites=="true") { emailInvitesCheckBox.setValue(true); }
  var startTimeLabel = app.createLabel().setText('Start time (must use a spreadsheet formatted datetime value)');
  var startTimeTextBox = app.createTextBox().setName('startTimeToken').setWidth("100%");
  if (startTimeToken) { startTimeTextBox.setValue(startTimeToken); }
  var endTimeLabel = app.createLabel().setText('End time (must use a spreadsheet formatted datetime value)');
  var endTimeTextBox = app.createTextBox().setName('endTimeToken').setWidth("100%");
  if (endTimeToken) { endTimeTextBox.setValue(endTimeToken); }
  var allDayCheckBox = app.createCheckBox().setText('All day event').setId('allDayTrue').setName('allDay').setVisible(false);
  var allDayCheckBoxFalse = app.createCheckBox().setText('All day event').setId('allDayFalse').setVisible(true);
  if (allDay=="true") { 
    allDayCheckBox.setValue(true).setVisible(true);
    allDayCheckBoxFalse.setVisible(false);
    endTimeLabel.setVisible(false);
    endTimeTextBox.setVisible(false);
   }
  var uncheckClientHandler = app.createClientHandler().forTargets(endTimeLabel, endTimeTextBox).setVisible(true);
                                                                                                           
  var uncheckServerHandler = app.createServerHandler('formMule_uncheck').addCallbackElement(allDayCheckBox);
  
  var checkClientHandler = app.createClientHandler().forTargets(endTimeLabel, endTimeTextBox).setVisible(false);                                                   
 
  var checkServerHandler = app.createServerHandler('formMule_check').addCallbackElement(allDayCheckBox);
  
  allDayCheckBox.addClickHandler(uncheckClientHandler).addClickHandler(uncheckServerHandler);
  allDayCheckBoxFalse.addClickHandler(checkClientHandler).addClickHandler(checkServerHandler);

  var descLabel = app.createLabel().setText('Event description  (HTML accepted, however tags will unfortunately get stripped upon next edit of these settings.)');
  var descTextArea = app.createTextArea().setName('descToken').setWidth("100%").setHeight(75);
  if (descToken) { descTextArea.setValue(descToken); }
  var reminderLabel = app.createLabel().setText('Set reminder type');
  var reminderListBox = app.createListBox().setName('reminderType');
  reminderListBox.addItem('None');
  reminderListBox.addItem('Email reminder');
  reminderListBox.addItem('Popup reminder');
  reminderListBox.addItem('SMS reminder');
   if (reminderType) { 
    switch(reminderType) {
    case "Email reminder": 
       reminderListBox.setSelectedIndex(1);
    break;
    case "Popup reminder":
       reminderListBox.setSelectedIndex(2);
    case "SMS reminder":
       reminderListBox.setSelectedIndex(3);
    break;
    default:
        reminderListBox.setSelectedIndex(0);
    }
  }
  
  var reminderMinLabel = app.createLabel().setText('Minutes before');
  var reminderMinTextBox = app.createTextBox().setName('minBefore');
  if (minBefore) { 
    reminderMinTextBox.setValue(minBefore);
  }
  var repeatLabel = app.createLabel('Repeat for how many weeks (leave blank for single event)');
  var repeatTextBox = app.createTextBox().setName('calendarWeeklyRepeats');
  var calendarWeeklyRepeats = ScriptProperties.getProperty('calendarWeeklyRepeats');
  if (calendarWeeklyRepeats) {
    repeatTextBox.setValue(calendarWeeklyRepeats)
  }
  var daysLabel = app.createLabel('Comma separated days of the week to repeat the event on (e.g. Monday,Wednesday)');
  var daysBox = app.createTextBox().setName('calendarWeekdays');
  var calendarWeekdays = ScriptProperties.getProperty('calendarWeekdays');
  if (calendarWeekdays) {
    daysBox.setValue(calendarWeekdays);
  }
  if (minBefore) { reminderMinTextBox.setValue(minBefore); }
  settingsGrid.setWidget(0, 0, calendarLabel).setWidget(0, 1, calendarTextBox)
              .setWidget(1, 0, eventTitleLabel).setWidget(1, 1, eventTitleTextBox)
              .setWidget(2, 0, eventLocationLabel).setWidget(2, 1, eventLocationTextBox) 
              .setWidget(3, 0, guestsLabel).setWidget(3, 1, guestsTextBox) 
              .setWidget(4, 0, emailInvitesCheckBox)
              .setWidget(5, 0, allDayCheckBox)
              .setWidget(6, 0, allDayCheckBoxFalse)
              .setWidget(7, 0, startTimeLabel).setWidget(7, 1, startTimeTextBox)
              .setWidget(8, 0, endTimeLabel).setWidget(8, 1, endTimeTextBox)
              .setWidget(9, 0, descLabel).setWidget(9, 1, descTextArea)
              .setWidget(10, 0, reminderLabel).setWidget(10, 1, reminderListBox)
              .setWidget(11, 0, reminderMinLabel).setWidget(11, 1, reminderMinTextBox)
              .setWidget(12, 0, repeatLabel)
              .setWidget(12, 1, repeatTextBox)
              .setWidget(13, 0, daysLabel)
              .setWidget(13, 1, daysBox);
  
  settingsGrid.setStyleAttribute("backgroundColor", "whiteSmoke").setStyleAttribute('textAlign', 'right').setStyleAttribute('padding', '8px');
  settingsGrid.setStyleAttribute(0, 1, 'width', '200px');
  
  
  
  var checkBox = app.createCheckBox().setText("Turn on calendar-event merge feature.").setId('calendarStatus').setName('calendarStatus').setStyleAttribute('color', '#0099FF').setStyleAttribute('fontWeight', 'bold');
 
  var clickHandler1 = app.createServerHandler('toggle1').addCallbackElement(topSettingsGrid);
  checkBox.addValueChangeHandler(clickHandler1);
      
  var calTrigger = ScriptProperties.getProperty('calTrigger');
  var calTriggerCheckBox = app.createCheckBox().setText("Trigger event creation on form submit.").setId('calTriggerCheckBox').setName('calTrigger').setStyleAttribute('color', '#0099FF').setStyleAttribute('fontWeight', 'bold').setEnabled(false);
  if (calTrigger=="true") {
    calTriggerCheckBox.setValue(true);
  }
  if (calendarStatus=="true") {
    checkBox.setValue(true);
    calTriggerCheckBox.setEnabled(true);
    dropdown.setEnabled(true);
    textbox.setEnabled(true);
    eventIdCol.setEnabled(true);
    extraHelp1.setVisible(true);
  }
 
       
  var updateCheckBox = app.createCheckBox().setText("Turn on calendar-event update feature.").setId('calendarUpdateStatus').setName('calendarUpdateStatus').setStyleAttribute('color', '#66C285').setStyleAttribute('fontWeight', 'bold');
 
  
  var clickHandler2 = app.createServerHandler('toggle2').addCallbackElement(topSettingsGrid);
  updateCheckBox.addValueChangeHandler(clickHandler2);
  
  var calUpdateTrigger = ScriptProperties.getProperty('calUpdateTrigger');
  var calUpdateTriggerCheckBox = app.createCheckBox().setText("Trigger event update on form submit.").setId('calUpdateTriggerCheckBox').setName('calUpdateTrigger').setStyleAttribute('color', '#66C285').setStyleAttribute('fontWeight', 'bold').setEnabled(false);
  if (calUpdateTrigger=="true") {
    calUpdateTriggerCheckBox.setValue(true);
  }
  
  if (calendarUpdateStatus=="true") {
    updateCheckBox.setValue(true);
    eventIdCol.setEnabled(true);
    calUpdateTriggerCheckBox.setEnabled(true);
    updateDropdown.setEnabled(true);
    updateTextbox.setEnabled(true);
    deleteDropdown.setEnabled(true);
    deleteTextbox.setEnabled(true);
    extraHelp2.setVisible(true);
  }
  
  
  //app.add(popupPanel);
  topSettingsGrid.setWidget(0, 0, checkBox);
  topSettingsGrid.setWidget(1, 0, calTriggerCheckBox);
  topSettingsGrid.setWidget(2, 0, createPanel);
  topSettingsGrid.setWidget(3, 0, extraHelp2);
  topSettingsGrid.setWidget(0, 1, updateCheckBox);
  topSettingsGrid.setWidget(1, 1, calUpdateTriggerCheckBox);
  topSettingsGrid.setWidget(2, 1, updateConditionsGrid);
  topSettingsGrid.setWidget(3, 1, eventIdPanel); 
  var verticalPanel = app.createVerticalPanel();
  var variablesPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px');
  var variablesScrollPanel = app.createScrollPanel().setHeight(240);
  var variablesLabel = app.createLabel().setText("Choose from the following variables: ").setStyleAttribute('fontWeight', 'bold');
  var tags = formMule_getAvailableTags();
  var flexTable = app.createFlexTable()
  for (var i = 0; i<tags.length; i++) {
    var tag = app.createLabel().setText(tags[i]);
    flexTable.setWidget(i, 0, tag);
    }
  variablesPanel.add(variablesLabel);
  variablesScrollPanel.add(flexTable);
  variablesPanel.add(variablesScrollPanel);
  panel.add(settingsGrid);
  panel.add(variablesPanel);
  verticalPanel.add(panel);
  var mainPanel = app.createVerticalPanel();
  mainPanel.add(topSettingsGrid);
  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "220px");
  spinner.setId("dialogspinner");
  var submitHandler = app.createServerHandler('formMule_saveEmailSettings').addCallbackElement(topSettingsGrid).addCallbackElement(panel);
  var backgroundHandler = app.createClientHandler().forTargets(mainPanel, settingsGrid).setStyleAttribute('opacity', '0.5');
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true);
  var buttonHandler = app.createServerHandler('formMule_saveCalendarSettings').addCallbackElement(scrollPanel).addCallbackElement(topSettingsGrid);
  var button = app.createButton().setText('Save Calendar Settings').addClickHandler(buttonHandler);
  button.addMouseDownHandler(backgroundHandler).addMouseDownHandler(spinnerHandler);
  verticalPanel.add(button);
  scrollPanel.add(verticalPanel);
  mainPanel.add(scrollPanel);
  app.add(mainPanel);
  app.add(spinner);
  ss.show(app);
  return app;
}

function toggle1(e) {
  var app = UiApp.getActiveApplication();
  var calTriggerCheckBox = app.getElementById('calTriggerCheckBox');
  var textBox = app.getElementById('val-0');
  var dropDown = app.getElementById('col-0');
  var extraHelp1 = app.getElementById('extraHelp1');
  var calendarStatus = e.parameter.calendarStatus;
  if (calendarStatus == "true") {
    calTriggerCheckBox.setEnabled(true);
    textBox.setEnabled(true);
    dropDown.setEnabled(true);
    extraHelp1.setVisible(true);
  } else {
    calTriggerCheckBox.setValue(false);
    calTriggerCheckBox.setEnabled(false);
    textBox.setEnabled(false);
    dropDown.setEnabled(false);
    extraHelp1.setVisible(false);
  }
  return app;
}


function toggle2(e) {
  var app = UiApp.getActiveApplication();
  var calTriggerCheckBox = app.getElementById('calUpdateTriggerCheckBox');
  var deleteTextBox = app.getElementById('delete-val-0');
  var deleteDropDown = app.getElementById('delete-col-0');
  var textBox = app.getElementById('update-val-0');
  var dropDown = app.getElementById('update-col-0');
  var extraHelp2 = app.getElementById('extraHelp2');
  var eventIdCol = app.getElementById('eventIdCol');
  var calendarUpdateStatus = e.parameter.calendarUpdateStatus;
  if (calendarUpdateStatus == "true") {
    calTriggerCheckBox.setEnabled(true);
    textBox.setEnabled(true);
    dropDown.setEnabled(true);
    deleteTextBox.setEnabled(true);
    deleteDropDown.setEnabled(true);
    eventIdCol.setEnabled(true);
    extraHelp2.setVisible(true);
  } else {
    calTriggerCheckBox.setValue(false);
    calTriggerCheckBox.setEnabled(false);
    textBox.setEnabled(false);
    dropDown.setEnabled(false);
    deleteTextBox.setEnabled(false);
    deleteDropDown.setEnabled(false);
    eventIdCol.setEnabled(false);
    extraHelp2.setVisible(false);
  }
  return app;
}





function formMule_getAvailableTags() {
  var sheetName = ScriptProperties.getProperty("sheetName");
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Browser.msgBox('You must select a source data sheet before this menu item can be selected.');
    formMule_defineSettings();
  }
  var lastCol = sheet.getLastColumn();
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getDisplayValues();
  var availableTags = [];
  
  for (var i=0; i<headers[0].length; i++) {
    availableTags[i] = "${\""+headers[0][i]+"\"}";
  } 
  var k = availableTags.length;
  availableTags[k] = "$currDay  (current day)";
  availableTags[k+1] = "$currMonth  (current month)";
  availableTags[k+2] = "$currYear  (current year)";
  availableTags[k+3] = "$eventId  (available in description only)";
  availableTags[k+4] = "$formUrl  (link to form)"
  return availableTags;
}


function formMule_check() {
//  Browser.msgBox("hellow");
  var app = UiApp.getActiveApplication();
  var allDayCheckBox = app.getElementById('allDayTrue');
  allDayCheckBox.setValue(true).setVisible(true);
  var allDayCheckBoxFalse = app.getElementById('allDayFalse');
  allDayCheckBoxFalse.setValue(false).setVisible(false);
  return app;
}

function formMule_uncheck() {
//  Browser.msgBox("hello");
  var app = UiApp.getActiveApplication();
  var allDayCheckBox = app.getElementById('allDayTrue');
  allDayCheckBox.setValue(false).setVisible(false);
  var allDayCheckBoxFalse = app.getElementById('allDayFalse');
  allDayCheckBoxFalse.setValue(false).setVisible(true);
  return app;
}

function formMule_fetchHeaders(sheet) {
  if (sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([['Dummy Header 1','Dummy Header 2','Dummy Header 3']]);
    SpreadsheetApp.flush();
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return headers;  
}


function formMule_saveCalendarSettings (e) {
  var app = UiApp.getActiveApplication();
  var calendarStatus = e.parameter.calendarStatus;
  ScriptProperties.setProperty('calendarStatus', calendarStatus);
  var calendarUpdateStatus = e.parameter.calendarUpdateStatus;
  ScriptProperties.setProperty('calendarUpdateStatus', calendarUpdateStatus);
  var calTrigger = e.parameter.calTrigger;
  ScriptProperties.setProperty('calTrigger', calTrigger);
  var calUpdateTrigger = e.parameter.calUpdateTrigger;
  ScriptProperties.setProperty('calUpdateTrigger', calUpdateTrigger);
  var conditionObject = new Object();
  conditionObject['col-0'] = e.parameter['col-0'];
  conditionObject['val-0'] = e.parameter['val-0'].trim();
  conditionObject['sht-0'] = "Event Creation";
  var conditionString = Utilities.jsonStringify(conditionObject);
  ScriptProperties.setProperty('calendarConditions', conditionString);
  var updateConditionObject = new Object();
  updateConditionObject['col-0'] = e.parameter['update-col-0'];
  updateConditionObject['val-0'] = e.parameter['update-val-0'].trim();
  updateConditionObject['sht-0'] = "Event Update";
  var updateConditionString = Utilities.jsonStringify(updateConditionObject);
  ScriptProperties.setProperty('calendarUpdateConditions', updateConditionString);
  var deleteConditionObject = new Object();
  deleteConditionObject['col-0'] = e.parameter['delete-col-0'];
  deleteConditionObject['val-0'] = e.parameter['delete-val-0'].trim();
  deleteConditionObject['sht-0'] = "Event Update";
  var deleteConditionString = Utilities.jsonStringify(deleteConditionObject);
  ScriptProperties.setProperty('calendarDeleteConditions', deleteConditionString);
  
  //check for illegal choices
  if ((e.parameter.calendarStatus=="true")&&(e.parameter.calendarUpdateStatus=="true")&&(e.parameter['col-0']==e.parameter['update-col-0'])&&(e.parameter['val-0']==e.parameter['update-val-0'])&&(e.parameter['update-val-0']!='')) {
    Browser.msgBox("You cannot set identical conditions for calendar event creation and calendar event update, as this would create and update an event at the same time.");
    app.close();
    formMule_setCalendarSettings();
    return app;
  }
  if ((e.parameter.calendarStatus=="true")&&(e.parameter.calendarUpdateStatus=="true")&&(e.parameter['col-0']==e.parameter['delete-col-0'])&&(e.parameter['val-0']==e.parameter['delete-val-0'])&&(e.parameter['delete-val-0']!='')) {
    Browser.msgBox("You cannot set identical conditions for calendar event creation and calendar event deletion, as this would create and delete an event at the same time.");   
    app.close();
    formMule_setCalendarSettings();
    return app;
  }
  
  if ((e.parameter.calendarUpdateStatus=="true")&&(e.parameter['update-col-0']==e.parameter['delete-col-0'])&&(e.parameter['update-val-0']==e.parameter['delete-val-0'])&&(e.parameter['delete-val-0']!='')) {
    Browser.msgBox("You cannot set identical conditions for calendar event update and calendar event deletion, as this would update and delete an event at the same time.");   
    app.close();
    formMule_setCalendarSettings();
    return app; 
  }  
  
  var eventIdCol = e.parameter.eventIdCol;
  ScriptProperties.setProperty('eventIdCol', eventIdCol);
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  if (calendarStatus=="true") {
    headers = formMule_fetchHeaders(sheet);
    var eventIdIndex = headers.indexOf("Event Id");
    var statusIndex = headers.indexOf("Event Creation Status");
    if((statusIndex==-1)&&(eventIdIndex==-1)) {
      var lastCol = sheet.getLastColumn();
      sheet.insertColumnAfter(lastCol);
      sheet.getRange(1, lastCol+1).setValue('Event Id').setBackground("#0099FF").setFontColor("black").setComment("Don't change the name of this column.");
      var copyDownOption = ScriptProperties.getProperty('copyDownOption');
      if (copyDownOption=="true") {
        sheet.getRange(2, lastCol+1).setValue('N/A: This is the formula row.').setBackground("#0099FF").setFontColor("black");
      }
      sheet.insertColumnAfter(lastCol+1);
      sheet.getRange(1, lastCol+2).setValue('Event Creation Status').setBackground("#0099FF").setFontColor("black").setComment("Don't change the name of this column.");
      var copyDownOption = ScriptProperties.getProperty('copyDownOption');
      if (copyDownOption=="true") {
        sheet.getRange(2, lastCol+2).setValue('N/A: This is the formula row.').setBackground("#0099FF").setFontColor("black");
      }
    }
  }
    if (calendarUpdateStatus=="true") {
    headers = formMule_fetchHeaders(sheet);
    var eventUpdateIndex = headers.indexOf("Event Update Status");
    var eventIdIndex = headers.indexOf("Event Id");
    if((eventIdIndex!=-1)&&(eventUpdateIndex==-1)) {
      var lastCol = sheet.getLastColumn();
      sheet.insertColumnAfter(lastCol);
      sheet.getRange(1, lastCol+1).setValue('Event Update Status').setBackground("#66C285").setFontColor("black").setComment("Don't change the name of this column.");
      var copyDownOption = ScriptProperties.getProperty('copyDownOption');
      if (copyDownOption=="true") {
        sheet.getRange(2, lastCol+1).setValue('N/A: This is the formula row.').setBackground("#66C285").setFontColor("black");
      }
    }
  }
    var calendarToken = e.parameter.calendarToken;
    var eventTitleToken = e.parameter.eventTitleToken;
    var locationToken = e.parameter.locationToken;
    var allDay = e.parameter.allDay;
    var startTimeToken = e.parameter.startTimeToken;
    var endTimeToken = e.parameter.endTimeToken;
    var descToken = e.parameter.descToken;
    var guests = e.parameter.guests;
    var emailInvites = e.parameter.emailInvites;
    var reminderType = e.parameter.reminderType;
    var sendInvites = e.parameter.sendInvites;
    var minBefore = e.parameter.minBefore;
    var calendarWeeklyRepeats = e.parameter.calendarWeeklyRepeats;
    var calendarWeekdays = e.parameter.calendarWeekdays;
  
    ScriptProperties.setProperty('calendarToken', calendarToken);
    ScriptProperties.setProperty('eventTitleToken', eventTitleToken);
    ScriptProperties.setProperty('locationToken', locationToken);
    ScriptProperties.setProperty('guests', guests);
    ScriptProperties.setProperty('emailInvites', emailInvites);
    ScriptProperties.setProperty('allDay', allDay);
    ScriptProperties.setProperty('startTimeToken', startTimeToken);
    ScriptProperties.setProperty('endTimeToken', endTimeToken);
    ScriptProperties.setProperty('descToken', descToken);
    ScriptProperties.setProperty('reminderType', reminderType);
    ScriptProperties.setProperty('minBefore', minBefore);
    ScriptProperties.setProperty('calendarWeeklyRepeats', calendarWeeklyRepeats);
    ScriptProperties.setProperty('calendarWeekdays', calendarWeekdays);
  
    var errMsg = '';
    if (calendarToken=='') { errMsg += "You forgot to enter a Calendar Id, "; }
    if (eventTitleToken=='') { errMsg += "You forgot to enter an event title, "; }
    if ((reminderType!="None")&&(minBefore=='')) { errMsg += "You forgot to specify the numer of minutes before the event for reminders, "; }
    if ((allDay!='true')&&(startTimeToken=='')) { errMsg += "You forgot to enter a start time, "; }
    if ((allDay!='true')&&(endTimeToken=='')) { errMsg += "You forgot to enter an end time, "; }
    if (errMsg !='') {
      Browser.msgBox(errMsg);
      formMule_setCalendarSettings ();
      app.close();
      return app;
    }
   app.close();
   return app;
}
  

function formMule_onFormSubmit () {
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
     ss = SpreadsheetApp.openById(properties.ssId);
  }
  var submissionRow = ss.getActiveRange().getRow();
  var lock = LockService.getPublicLock();
  lock.waitLock(120000);
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  var caseNoSetting = ScriptProperties.getProperty('caseNoSetting');
  copyDownFormulas(submissionRow, properties);
  if (caseNoSetting == "true") {
    var headers = formMule_fetchHeaders(sheet);
    var caseNoIndex = headers.indexOf("Case No");
    var cellRange = sheet.getRange(submissionRow, caseNoIndex+1);
    var cellValue = cellRange.getValue();
    if (cellValue=="") {
      cellRange.setValue(formMule_assignCaseNo());
    }
  }
  formMule_sendEmailsAndSetAppointments();
  lock.releaseLock();
}


function urlencode(inNum) {
  // Function to convert non URL compatible characters to URL-encoded characters
  var outNum = 0;     // this will hold the answer
  outNum = escape(inNum); //this will URL Encode the value of inNum replacing whitespaces with %20, etc.
  return outNum;  // return the answer to the cell which has the formula
}




function formMule_assignCaseNo(resetToValue) {
  if (resetToValue) {
  var caseNo = resetToValue;
  } else { 
  var caseNo = ScriptProperties.getProperty('caseNo')
  };
  if(caseNo==null) {
      ScriptProperties.setProperty('caseNo','0');
      caseNo = 0;
   } else {
   caseNo = parseInt(caseNo) + 1;
   ScriptProperties.setProperty('caseNo', caseNo);
  }
return caseNo; 
}


function formMule_emailSettings () {
  var app = UiApp.createApplication().setTitle("Step 2a: Set up email merge").setHeight(450).setWidth(700);
  var panel = app.createVerticalPanel().setId("emailPanel").setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '5px').setStyleAttribute('marginTop', '5px');
  var scrollPanel = app.createScrollPanel().setHeight("150px");
  var helpLabel = app.createLabel().setText('Indicate below whether you want templated emails to be generated upon form submit, how many different emails you want to create, and (optionally) any additional conditions that must be met before sending');
  var helpPopup = app.createPopupPanel();
  helpPopup.add(helpLabel);
  app.add(helpPopup);
  
 
  var emailStatus = ScriptProperties.getProperty('emailStatus');
  var emailStatusCheckBox = app.createCheckBox().setText('Turn-on email merge feature').setName('emailStatus').setStyleAttribute('color', '#FF9999').setStyleAttribute('fontWeight','bold');
  if (emailStatus=="true") {
    emailStatusCheckBox.setValue(true);
  }
  var emailTrigger = ScriptProperties.getProperty('emailTrigger');
  var emailTriggerCheckBox = app.createCheckBox().setText('Trigger this feature on form submit').setName('emailTrigger').setStyleAttribute('color', '#FF9999').setStyleAttribute('fontWeight','bold');
  if (emailTrigger=="true") {
    emailTriggerCheckBox.setValue(true);
  }
  var numSelectLabel = app.createLabel().setText("How many different unique, templated emails do you want to send out to different recipients when the form is submitted?");
  var grid = app.createGrid().setId('emailConditionGrid');
  formMule_setEmailConditionGrid(app);
  var numSelectChangeHandler = app.createServerHandler('formMule_refreshEmailConditions').addCallbackElement(panel);
  var numSelect = app.createListBox().setId("numSelect").setName("numSelectValue").addChangeHandler(numSelectChangeHandler);
  numSelect.addItem('1');
  numSelect.addItem('2');
  numSelect.addItem('3');
  numSelect.addItem('4');
  numSelect.addItem('5');
  numSelect.addItem('6');
  var preSelectedNum = ScriptProperties.getProperty('numSelected');
  switch (preSelectedNum) {
    case "1":
     numSelect.setSelectedIndex(0);
    break;
    case "2":
     numSelect.setSelectedIndex(1);
    break;
    case "3":
     numSelect.setSelectedIndex(2);
    break;
     case "4":
     numSelect.setSelectedIndex(3);
    break;
     case "5":
     numSelect.setSelectedIndex(4);
    break;
    case "6":
     numSelect.setSelectedIndex(5);
    break;
    default: 
     numSelect.setSelectedIndex(0);
  }
  
  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "220px");
  spinner.setId("dialogspinner");
  
  
  var submitHandler = app.createServerHandler('formMule_saveEmailSettings').addCallbackElement(panel);
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  var button = app.createButton().setText("Submit settings").addClickHandler(submitHandler).addMouseDownHandler(spinnerHandler);
  var extraHelp1 = app.createLabel('Note: Emails are generated only for rows where the condition is met AND the \"Send Status\" for that template is blank').setStyleAttribute('fontSize', '10px').setStyleAttribute('color', '#FF9999').setStyleAttribute('fontWeight', 'bold').setId('extraHelp1');
  var numSelectNote = app.createLabel().setStyleAttribute('marginTop', '20px').setText("The the first time you submit this form, it will auto-generate a new sheet with a blank template for each email. You will need to go to these sheets and complete the the \"To:\", \"CC:\", \"Subject:\", and \"Body:\" sections of the template. If you return later and increase the number here, the script will generate additional sheets. Decreasing this value will delete the corresponding sheets.  Do not change the names of the template sheets.");
  panel.add(emailStatusCheckBox);
  panel.add(emailTriggerCheckBox);
  panel.add(numSelectLabel);
  panel.add(numSelect);
  scrollPanel.add(grid);
  panel.add(scrollPanel);
  panel.add(extraHelp1);
  panel.add(numSelectNote); 
  panel.add(button);
  app.add(panel);
  app.add(spinner);
  this.ss.show(app);
}

function formMule_setEmailConditionGrid(app) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = ScriptProperties.getProperty('sheetName');
  if ((!sourceSheetName)||(sourceSheetName=='')) {
    Browser.msgBox('You must select a source data sheet before this menu item can be selected.');
    formMule_defineSettings();
    app().close();
    return app;
  } else {
    var sourceSheet = ss.getSheetByName(sourceSheetName);
  }
  var lastCol = sourceSheet.getLastColumn();
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var numSelected = ScriptProperties.getProperty('numSelected');
  if ((!numSelected)||(numSelected=='')) {
    numSelected = 1;
  }
  numSelected = parseInt(numSelected);
  var grid = app.getElementById('emailConditionGrid');
  grid.resize(numSelected+1, 5);
  grid.setWidget(0,0,app.createLabel('Send Email Template'));
  grid.setWidget(0,2,app.createLabel('Column'));
  grid.setWidget(0,4,app.createLabel('Leave blank to send for all rows. Use NULL for empty.  NOT NULL for not empty.').setStyleAttribute('fontSize', '10px'));
  var dropdown = [];
  var textbox = [];
  var namebox = [];
  for (var i=0; i<numSelected; i++) {
    var label = 'Condition for Email ' + (i+1);
    grid.setWidget(i+1, 0, app.createLabel(label)).setStyleAttribute(i+1, 0, 'width','100px');
    dropdown[i] = app.createListBox().setId('conditionCol-'+i).setName('conditionCol-'+i).setWidth("150px");
    for (var j=0; j<headers.length; j++) {
        dropdown[i].addItem(headers[j]);
    }
    namebox[i] = app.createTextBox().setId('templateName-'+i).setName('templateName-'+i);
    textbox[i] = app.createTextBox().setId('value-'+i).setName('value-'+i);
    grid.setWidget(i+1, 0, namebox[i]);
    grid.setWidget(i+1, 1, app.createLabel('if'));
    grid.setWidget(i+1, 2, dropdown[i]);
    grid.setWidget(i+1, 3, app.createLabel('='));
    grid.setWidget(i+1, 4, textbox[i]);
  }
  var emailConditions = ScriptProperties.getProperty('emailConditions');
  if ((emailConditions)&&(emailConditions!='')) {
    emailConditions = Utilities.jsonParse(emailConditions);
    numSelected = parseInt(numSelected);
    for (var i=0; i<numSelected; i++) {
      var preset = 0;
      for (var j=0; j<headers.length; j++) {
        if (emailConditions["col-"+i]==headers[j]) {
          preset = j;
          break;
        }
      }
      dropdown[i].setSelectedIndex(preset);
      textbox[i].setValue(emailConditions["val-"+i]);
      if ((emailConditions["sht-"+i])&&(emailConditions["sht-"+i]!='')) {
        namebox[i].setValue(emailConditions["sht-"+i]);
      } else {
        namebox[i].setValue("Email"+(i+1)+" Template");
      }
    }
  }
  
  if (!(emailConditions)) {
    namebox[0].setValue("Email1 Template");
  }
  return app; 
}


function formMule_refreshEmailConditions(e) {
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('emailConditionGrid');
  var numSelected = e.parameter.numSelectValue;
  var oldNumSelected = ScriptProperties.getProperty('numSelected');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = ScriptProperties.getProperty('sheetName');
  if ((!sourceSheetName)||(sourceSheetName=='')) {
    var sourceSheet = ss.getSheets()[0];
  } else {
    var sourceSheet = ss.getSheetByName(sourceSheetName);
  }
  var lastCol = sourceSheet.getLastColumn();
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  if ((!numSelected)||(numSelected=='')) {
    numSelected = 1;
  }
  numSelected = parseInt(numSelected);
  grid.resize(numSelected+1, 5);
  grid.setWidget(0,0,app.createLabel('Send Email Template'));
  grid.setWidget(0,2,app.createLabel('Column'));
  grid.setWidget(0,4,app.createLabel('Leave blank to ignore. Use NULL for empty.  NOT NULL for not empty.').setStyleAttribute('fontSize', '10px'));
  var dropdown = [];
  var namebox = [];
  var textbox = [];
  for (var i=0; i<numSelected; i++) {
    var label = 'Condition for Email ' + (i+1);
    grid.setWidget(i+1, 0, app.createLabel(label)).setStyleAttribute(i+1, 0, 'width','100px');
    dropdown[i] = app.createListBox().setId('conditionCol-'+i).setName('conditionCol-'+i).setWidth("150px");
    for (var j=0; j<lastCol; j++) {
        dropdown[i].addItem(headers[j]);
    }
    namebox[i] = app.createTextBox().setId('templateName-'+i).setName('templateName-'+i);
    textbox[i] = app.createTextBox().setId('value-'+i).setName('value-'+i);
    grid.setWidget(i+1, 0, namebox[i]);
    grid.setWidget(i+1, 1, app.createLabel('if'));
    grid.setWidget(i+1, 2, dropdown[i]);
    grid.setWidget(i+1, 3, app.createLabel('='));
    grid.setWidget(i+1, 4, textbox[i]);
  }
  var emailConditions = new Object();
  for (var i=0; i<numSelected; i++) {
    var condCol = e.parameter["conditionCol-"+i];
    var condVal = e.parameter["value-"+i];
    var shtName = e.parameter["templateName-"+i];
    emailConditions["col-"+i] = condCol;
    emailConditions["val-"+i] = condVal;
    emailConditions["sht-"+i] = shtName;
  }
  emailConditions = Utilities.jsonStringify(emailConditions);
  if ((emailConditions)&&(emailConditions!='')) {
    emailConditions = Utilities.jsonParse(emailConditions);
    if (numSelected>=oldNumSelected) {
    for (var i=0; i<oldNumSelected; i++) {    
      var preset = 0;
      for (var j=0; j<headers.length; j++) {
        if (emailConditions["col-"+i]==headers[j]) {
          preset = j;
          break;
        }
      }
      dropdown[i].setSelectedIndex(preset);
      textbox[i].setValue(emailConditions["val-"+i]);
     }
   }
   if (numSelected>=oldNumSelected) {
     for (var i=0; i<numSelected; i++) {    
       if ((emailConditions["sht-"+i]&&emailConditions["sht-"+i]!='')) {
        namebox[i].setValue(emailConditions["sht-"+i]);
      } else {
        namebox[i].setValue("Email"+(i+1)+" Template");
      }
     }
    }
    if (numSelected<oldNumSelected) {
    for (var i=0; i<numSelected; i++) {    
      var preset = 0;
      for (var j=0; j<headers.length; j++) {
        if (emailConditions["col-"+i]==headers[j]) {
          preset = j;
          break;
        }
      }
      dropdown[i].setSelectedIndex(preset);
      textbox[i].setValue(emailConditions["val-"+i]);
      if ((emailConditions["sht-"+i]&&emailConditions["sht-"+i]!='')) {
        namebox[i].setValue(emailConditions["sht-"+i]);
      } else {
        namebox[i].setValue("Email"+(i+1)+" Template");
      }
     }
   }
  }
  return app; 
}


//returns true if testval meets the condition 
function formMule_evaluateSMSConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["smsCol-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["smsVal-"+i];
  if (condObject["smsName-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["smsName-"+i]+" SMS Status");
    if (rowData[statusCol]!='') {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = true;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}


function formMule_evaluateVMConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["vmCol-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["vmVal-"+i];
  if (condObject["vmName-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["vmName-"+i]+" VM Status");
    if (rowData[statusCol]!='') {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = true;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}







//returns true if testval meets the condition 
function formMule_evaluateConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["col-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["val-"+i];
  if (condObject["sht-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["sht-"+i]+" Status");
    if (rowData[statusCol]!='') {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = true;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}


//returns true if testval meets the condition 
function formMule_evaluateUpdateConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["col-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["val-"+i];
  if (condObject["sht-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["sht-"+i]+" Status");
    if ((rowData[statusCol]!=null)&&(rowData[statusCol]!='')) {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = false;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}


function formMule_saveEmailSettings(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.getActiveApplication();
  var emailStatus = e.parameter.emailStatus;
  var emailTrigger = e.parameter.emailTrigger;
  ScriptProperties.setProperty('emailStatus', emailStatus);
  ScriptProperties.setProperty('emailTrigger', emailTrigger);
  var numSelected = e.parameter.numSelectValue;
  var emailConditions = new Object();
  emailConditions["max"]=numSelected;
  for (var i=0; i<numSelected; i++) {
    var condCol = e.parameter["conditionCol-"+i];
    var condVal = e.parameter["value-"+i].trim();
    var templateName = e.parameter["templateName-"+i].trim();
    Logger.log(templateName);
    emailConditions["col-"+i] = condCol;
    emailConditions["val-"+i] = condVal;
    emailConditions["sht-"+i] = templateName;
  }
  emailConditions = Utilities.jsonStringify(emailConditions);
  ScriptProperties.setProperty('emailConditions', emailConditions);
  emailConditions = Utilities.jsonParse(emailConditions);
  ScriptProperties.setProperty('numSelected', numSelected);
  var num = parseInt(numSelected);
  var sheets = ss.getSheets();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var copyDownOption = ScriptProperties.getProperty('copyDownOption');
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  for (var i=0; i<num; i++) {
    var alreadyExists = '';
    for (var j=0; j<sheets.length; j++) {
      if (sheets[j].getName()==emailConditions["sht-"+i]) {
        alreadyExists = true;
        break;
      }
    }
    if (!(alreadyExists==true)) {
    var newSheet = ss.insertSheet().setName(emailConditions["sht-"+i]);
    var newSheetName = newSheet.getName();
    newSheet.getRange(1, 1).setValue("Reply to:").setBackground("yellow");
    newSheet.getRange(2, 1).setValue("To:").setBackground("yellow");
    newSheet.getRange(3, 1).setValue("CC:").setBackground("yellow");
    newSheet.getRange(4, 1).setValue("Subject:").setBackground("yellow"); 
    newSheet.getRange(5, 1).setValue("Body: \n Ctrl+Return gives newline. HTML-friendly!").setVerticalAlignment("top").setBackground("yellow");
    newSheet.getRange(6, 1).setValue("Translate code:").setBackground("yellow").setVerticalAlignment("top");
    newSheet.getRange(5, 2).setVerticalAlignment("top");
    newSheet.setRowHeight(5, 200); 
    newSheet.getRange(1, 3).setValue("<--Optional: Use a single valid email or email token from the list below. Even when this is set, sender address will always appear as the installer of this script.");
    newSheet.getRange(2, 3).setValue("<--Required: Use a single email address or comma separated email addresses, or use email tokens from the list below.");
    newSheet.getRange(3, 3).setValue("<--Optional: Complete me with valid, comma separated email addresses or email tokens from the list below.");
    newSheet.getRange(4, 3).setValue("<--Use tokens below for dynamic values.").setVerticalAlignment("middle");
    newSheet.getRange(5, 3).setValue("<--Use tokens below for dynamic values.").setVerticalAlignment("middle");
    newSheet.getRange(6, 3).setValue("<--Optional: E.g. 'es' for Spanish. Use token for dynamic value. Value must be available in Google Translate: https://developers.google.com/translate/v2/using_rest#language-params");
    newSheet.setColumnWidth(1, 75);
    newSheet.setColumnWidth(2, 350);
    newSheet.setColumnWidth(3, 340);
    formMule_setAvailableTags(newSheet);
  } 
  var headerIndex = headers.indexOf(emailConditions["sht-"+i] + " Status");
  if (headerIndex==-1) {
    var sheet = ss.getSheetByName(sheetName);
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1,1,1).setValue(emailConditions["sht-"+i] + " Status").setFontColor('black').setBackground('#FF9999').setComment("Don't change the name of this column. It is used to log the send status of an email template, whose name must be a match.");
    if (copyDownOption == "true") {
     sheet.getRange(2, sheet.getLastColumn(),1,1).setValue("N/A This is the formula row").setFontColor('black').setBackground('#FF9999'); 
    }
  }    
  }
  app.close();
  return app;
}


function formMule_setAvailableTags(templateSheet) {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var lastCol = sheet.getLastColumn();
  if(lastCol==0){
    lastCol = 1;
    var noTags = "You have no headers in your data sheet";
  }
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getDisplayValues();
  var availableTags = [];
  var j=0;
  for (var i=0; i<headers[0].length; i++) {
    if (headers[0][i]=='') {
      Browser.msgBox('Column ' + parseInt(i+1) + ' has a blank header. You cannot have blank headers in your source data. Please fix this before proceeding.');
      break;
    }
    if (headers[0][i]!='Event Update Status') {
      availableTags[j] = [];
      availableTags[j][0] = "${\""+headers[0][i]+"\"}";
      j++; 
    }
  }
  
  var k = availableTags.length;
  availableTags[k] = [];
  availableTags[k][0] = "$currMonth";
  availableTags[k+1] = [];
  availableTags[k+1][0] = "$currDay";
  availableTags[k+2] = [];
  availableTags[k+2][0] = "$currYear";
  availableTags[k+3] = [];
  availableTags[k+3][0] = "$formUrl";
    
  if (noTags) {
    availableTags[0][0] = noTags;
  }
 templateSheet.getRange(7, 1).setValue("Available merge variables").setFontWeight("bold").setBackground("yellow");;
 templateSheet.getRange(7, 2).setValue("Use in any field").setBackground("yellow");
 templateSheet.getRange(8, 1, availableTags.length,1).setBackground("yellow");
 templateSheet.getRange(8, 2, availableTags.length,1).setValues(availableTags).setBackground("yellow");
 templateSheet.getRange(9+availableTags.length, 1).setValue("Handy HTML Tags").setBackground("pink").setFontWeight("bold");
 templateSheet.getRange(9+availableTags.length, 2).setValue("Use in email body only").setBackground("pink");
 var htmlTags = [["Hyperlink","<a href= \"URL\">Link Text</a>"],["Header Level 3","<h3>My header</h3>"],['=hyperlink("http://www.w3schools.com/html/html_basic.asp","Learn more!")','']];
 templateSheet.getRange(10+availableTags.length, 1, htmlTags.length, 2).setValues(htmlTags).setBackground("pink");
 
}



function formMule_sendEmailsAndSetAppointments(manual) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var properties = ScriptProperties.getProperties();
  var calendarStatus = properties.calendarStatus;
  var calendarUpdateStatus = properties.calendarUpdateStatus;
  var calTrigger = properties.calTrigger;
  var calUpdateTrigger = properties.calUpdateTrigger;
  var copyDownOption = properties.copyDownOption;
  var sheetName = properties.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  ss.setActiveSheet(sheet);
  var headers = formMule_fetchHeaders(sheet);
  if (headers.indexOf("Formula Copy Down Status") != -1) {
    var copyDownFormulasCol = true;
  }
  var days = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
  if (((calendarStatus == "true")&&(manual==true))||((calendarStatus == "true")&&(calTrigger=="true"))||((calendarUpdateStatus == "true")&&(manual==true))||((calendarUpdateStatus == "true")&&(calUpdateTrigger=="true"))) {
    CacheService.getPrivateCache().remove('lastCalendarId');
    var eventIdCol = properties.eventIdCol;
    var calendarToken  = properties.calendarToken;
    var eventTitleToken = properties.eventTitleToken;
    var locationToken = properties.locationToken;
    var guestsToken = properties.guests;
    var emailInvites = properties.emailInvites;
    if (emailInvites == "true") { emailInvites = true; } else { emailInvites = false;}
    var allDay = properties.allDay;
    var startTimeToken = properties.startTimeToken;
    var endTimeToken = properties.endTimeToken;
    var descToken = properties.descToken;
    var reminderType = properties.reminderType;
    var minBefore = properties.minBefore;
    var calConditionsString = properties.calendarConditions;
    var calCondObject = Utilities.jsonParse(calConditionsString);
    var calUpdateConditionsString = properties.calendarUpdateConditions;
    var calUpdateCondObject = Utilities.jsonParse(calUpdateConditionsString);
    var calDeleteConditionsString = properties.calendarDeleteConditions;
    var calWeeklyTimes = properties.calendarWeeklyRepeats;
    var calRepeats = properties.calendarWeeklyRepeats;
    var calWeekdays = properties.calendarWeekdays;
    var calDeleteCondObject = Utilities.jsonParse(calDeleteConditionsString);
    var eventUpdateStatusIndex = headers.indexOf("Event Update Status");
    var eventCreateStatusIndex = headers.indexOf("Event Creation Status");
    var eventIdIndex = headers.indexOf("Event Id");
    var updateEventIdIndex = headers.indexOf(eventIdCol);
    eventIdCol = formMule_normalizeHeader(eventIdCol);
  }
  var emailStatus = properties.emailStatus;
  var emailTrigger = properties.emailTrigger;
  var emailCondString = properties.emailConditions;
  var smsStatus = properties.smsEnabled;
  var smsTrigger = properties.smsTrigger;
  var vmStatus = properties.vmEnabled;
  var vmTrigger = properties.vmTrigger;
  var accountSID = properties.accountSID;
  var authToken = properties.authToken;
  var maxTexts = properties.smsMaxLength;
      
  if (((emailStatus=="true")&&(manual==true))||((emailStatus=="true")&&(emailTrigger=="true"))) {
    var emailCondString = properties.emailConditions;
    var emailCondObject = Utilities.jsonParse(emailCondString);
    var numSelected = properties.numSelected;
    var templateSheetNames = [];
    var normalizedSendColNames = [];
    var sendersTemplates = [];
    var recipientsTemplates = [];
    var ccRecipientsTemplates = [];
    var subjectTemplates = [];
    var bodyTemplates = [];  
    var langTemplates = [];
    var sendColNames = [];
    for (var i=0; i<numSelected; i++) {
       var templateSheet = ss.getSheetByName(emailCondObject['sht-'+i]);
       sendColNames[i] = emailCondObject['sht-'+i] + " Status";
       normalizedSendColNames[i] = formMule_normalizeHeader(emailCondObject['sht-'+i] + " Status");
       templateSheetNames[i] = templateSheet.getName();
       sendersTemplates[i] = templateSheet.getRange("B1").getValue();
       recipientsTemplates[i] = templateSheet.getRange("B2").getValue();
       ccRecipientsTemplates[i] = templateSheet.getRange("B3").getValue();
       subjectTemplates[i] = templateSheet.getRange("B4").getValue();
       bodyTemplates[i] = templateSheet.getRange("B5").getValue();
       langTemplates[i] = templateSheet.getRange("B6").getValue();
    }
  }
  var lastHeader = sheet.getRange(1, sheet.getLastColumn());
  var lastHeaderValue = formMule_normalizeHeader(lastHeader.getValue());  
  var k=2;
  dataRange = sheet.getRange(k, 1, sheet.getLastRow()-(k-1), sheet.getLastColumn());
  // Create one JavaScript object per row of data.
  var objects = formMule_getRowsData(sheet, dataRange, 1);
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var j = 0; j < objects.length; ++j) {
    if ((properties.copyDownFormulas == "true")&&(!copyDownFormulasCol)) { //self-repair if spreadsheet header for copydown of formulas has been deleted
      returnCopydownStatusColIndex();
      break;
    }
    if (((properties.copyDownFormulas == "true")||(copyDownFormulasCol))&&(objects[j].formulaCopyDownStatus == "")) {
      copyDownFormulas(j+k, properties);
      var thisRowRange = sheet.getRange(j+k, 1, 1, sheet.getLastColumn());
      objects[j] = formMule_getRowsData(sheet, thisRowRange, 1);
    }
    var error = false;
    // Get a row object
    var rowData = objects[j];        
    var confirmation = '';
    var calUpdateConfirmation = '';
    var found = '';
    var calConditionTest = false;
    var calUpdateConditionTest = false;
    var calDeleteConditionTest = false;
    if (calendarStatus=="true") {    
      //test calendar event creation conditions
      if ((calConditionsString)&&(calConditionsString!='')) {
        var calConditionTest = formMule_evaluateConditions(calCondObject, 0, rowData);
      }
    }
    if (calendarUpdateStatus=="true") {
      var calUpdateStatus = rowData.eventUpdateStatus.replace(/ /g,'');
      if (calUpdateStatus=='') {
      //test up calendar update conditions
      if ((calUpdateConditionsString)&&(calUpdateConditionsString!='')) {
        var calUpdateConditionTest = formMule_evaluateUpdateConditions(calUpdateCondObject, 0, rowData);    
      } 
        //test calendar event deletion condition
        if ((calDeleteConditionsString)&&(calDeleteConditionsString!='')) {
          var calDeleteConditionTest = formMule_evaluateUpdateConditions(calDeleteCondObject, 0, rowData);    
        } 
      }
    }
    if ((calendarToken)&&((calUpdateConditionTest)||(calConditionTest)||(calDeleteConditionTest))) {
      var lastCalendarId = CacheService.getPrivateCache().get('lastCalendarId');
      var calendarId  = formMule_fillInTemplateFromObject(calendarToken, rowData);
      if (lastCalendarId!=calendarId) {
        var calendar = CalendarApp.getCalendarById(calendarId);
      }
    }
     
    //Either the event is being created or it's being updated
    if (((calConditionTest == true)||(calUpdateConditionTest == true))&&(calendar)) {
        var eventTitle = formMule_fillInTemplateFromObject(eventTitleToken, rowData);
        var location = formMule_fillInTemplateFromObject(locationToken, rowData);
        var desc = formMule_fillInTemplateFromObject(descToken, rowData);
        var guests = formMule_fillInTemplateFromObject(guestsToken, rowData);
        var repeats = formMule_fillInTemplateFromObject(calRepeats, rowData);
        var weekdays = formMule_fillInTemplateFromObject(calWeekdays, rowData);
        var startTimeStamp;
        var endTimeStamp;
        if (!(allDay=="true")) {
          var startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData) + " EST");
          var endTimeStamp = new Date(formMule_fillInTemplateFromObject(endTimeToken, rowData)  + " EST");
        } else {
          startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData) + " EST");
        }
      
      
       Logger.log(formMule_fillInTemplateFromObject(startTimeToken, rowData) + " EST");
       Logger.log(formMule_fillInTemplateFromObject(endTimeToken, rowData) + " EST");
       
        
        if (guests!='') {
          var options = {guests:guests, location:location, description:desc, sendInvites:emailInvites}; 
        } else {
          var options = {location:location, description:desc,};
        }
        //calendar event creation condition
        if ((calConditionTest == true)&&(calUpdateConditionTest == false)) {
          var event = '';
          var eventId = ''; 
          var calConfirmation = '';     
        try {
          if ((allDay=="true")&&(startTimeStamp != "Invalid Date")) {
            if ((!repeats)||(repeats=='')||(repeats==1)) {
              eventId = calendar.createAllDayEvent(eventTitle, startTimeStamp, options).getId();
              calConfirmation += "All day event added to calendar: "+ calendar.getName() + " for " + startTimeStamp;
            } else {
              if (repeats) {
                repeats = parseInt(repeats); 
                if (weekdays) {
                  weekdays = weekdays.split(",");
                  var weekdayArray = []
                  for (var h=0; h<weekdays.length; h++) {
                    var weekday = weekdays[h].trim();
                    weekday = weekday.toUpperCase();
                    if (days.indexOf(weekday)!=-1) {
                      weekdayArray.push(CalendarApp.Weekday[weekday]);
                    }
                  }
                  if (weekdayArray.length>0) { 
                    repeats = weekdayArray.length * repeats;
                    eventId = calendar.createAllDayEventSeries(eventTitle, startTimeStamp, CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), options).getId();
                  }
                } else {
                 eventId = calendar.createAllDayEventSeries(eventTitle, startTimeStamp, CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), options).getId();
                }
              }
              calConfirmation += "Recurring all day event added to calendar: "+ calendar.getName() + " starting " + startTimeStamp;
            }
            try {
              formMule_logCalEvent();
            } catch(err) {
            }
            sheet.getRange(j+k, eventIdIndex+1).setValue(eventId);
          } else if (startTimeStamp != "Invalid Date") {
             if ((!repeats)||(repeats=='')||(repeats==1)) {
               eventId = calendar.createEvent(eventTitle, startTimeStamp, endTimeStamp, options).getId();
               calConfirmation += "Event added to calendar: "+ calendar.getName() + " for " + startTimeStamp;
             } else {
                if (repeats) {
                  repeats = parseInt(repeats); 
                    if (weekdays) {
                      weekdays = weekdays.split(",");
                      var weekdayArray = []
                      for (var h=0; h<weekdays.length; h++) {
                        var weekday = weekdays[h].trim();
                        weekday = weekday.toUpperCase();
                        if (days.indexOf(weekday)!=-1) {
                          weekdayArray.push(CalendarApp.Weekday[weekday]);
                        }
                      }
                      if (weekdayArray.length>0) { 
                        repeats = weekdayArray.length * repeats;
                        eventId = calendar.createEventSeries(eventTitle, startTimeStamp, endTimeStamp, CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), options).getId();
                      }
                    } else {
                     eventId = calendar.createEventSeries(eventTitle, startTimeStamp, endTimeStamp, CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), options).getId();
                    }
                  }
               calConfirmation += "Recurring event added to calendar: "+ calendar.getName() + " starting " + startTimeStamp;
            }
            try {
              formMule_logCalEvent();
            } catch(err) {
            }
            sheet.getRange(j+k, eventIdIndex+1).setValue(eventId);
          } else {
            calConfirmation = "No event created. Invalid date/time parameters.";
            error = true;
          }
          } catch(err) {
            calConfirmation = "Create calendar event failed: " + err;
            error = true;
          }
        if (eventId) {
          event = calendar.getEventSeriesById(eventId);
          if (reminderType) { 
            switch(reminderType) {
            case "Email reminder": 
              event = calendar.getEventSeriesById(eventId).addEmailReminder(minBefore);
            break;
            case "Popup reminder":
              event = calendar.getEventSeriesById(eventId).addPopupReminder(minBefore);
            break;
            case "SMS reminder":
              event = calendar.getEventSeriesById(eventId).addSmsReminder(minBefore);
            break;
            default:
              event = calendar.getEventSeriesById(eventId).removeAllReminders();
            }
         }
         rowData.eventId = eventId;
         desc = formMule_fillInTemplateFromObject(descToken, rowData);
         event = calendar.getEventSeriesById(eventId).setDescription(desc); 
       }
      sheet.getRange(j+k, eventCreateStatusIndex+1).setValue(calConfirmation).setFontColor("black");
      if (error==true) {
        sheet.getRange(j+k, eventCreateStatusIndex+1).setFontColor("red");
      }
      error = false;
     } 
    }
    if ((calUpdateConditionTest == true)||(calDeleteConditionTest == true)) {
      var eventId = rowData[eventIdCol];
      found = false;
      try {
        var future = new Date();
        future.setDate(future.getDate()+360);
        var past = new Date();
        past.setDate(past.getDate()-360);
        var events = calendar.getEvents(past, future);
        for (var g = 0; g<events.length; g++) {
          if (events[g].getId()==eventId) {
            var event = events[g];
            found = true;
          }
        }
      } catch(err) {
        found = false;
      }
    }
    if ((calUpdateConditionTest == true)&&(found == true)) {
      try {
        if (allDay!="true") {
          if ((!repeats)||(repeats=='')||(repeats==1)) {
            var event = event.setTime(startTimeStamp, endTimeStamp);
          } else {
              if (repeats) {
                  repeats = parseInt(repeats); 
                    if (weekdays) {
                      weekdays = weekdays.split(",");
                      var weekdayArray = []
                      for (var h=0; h<weekdays.length; h++) {
                        var weekday = weekdays[h].trim();
                        weekday = weekday.toUpperCase();
                        if (days.indexOf(weekday)!=-1) {
                          weekdayArray.push(CalendarApp.Weekday[weekday]);
                        }
                      }
                      if (weekdayArray.length>0) { 
                        repeats = weekdayArray.length * repeats;
                        event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), startTimeStamp, endTimeStamp);
                      }
                    } else {
                      event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), startTimeStamp, endTimeStamp);
                    }
                  }
          }
        } else {
          if ((!repeats)||(repeats=='')||(repeats==1)) {
            var event = event.setAllDayDate(startTimeStamp);
          } else {
              if (repeats) {
                  repeats = parseInt(repeats); 
                    if (weekdays) {
                      weekdays = weekdays.split(",");
                      var weekdayArray = []
                      for (var h=0; h<weekdays.length; h++) {
                        var weekday = weekdays[h].trim();
                        weekday = weekday.toUpperCase();
                        if (days.indexOf(weekday)!=-1) {
                          weekdayArray.push(CalendarApp.Weekday[weekday]);
                        }
                      }
                      if (weekdayArray.length>0) { 
                        repeats = weekdayArray.length * repeats;
                        event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), startTimeStamp);
                      }
                    } else {
                      event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), startTimeStamp);
                    }
                  }
          }
        }
        event = calendar.getEventSeriesById(eventId).setDescription(desc)
        event = calendar.getEventSeriesById(eventId).setLocation(location);
        event = calendar.getEventSeriesById(eventId).setTitle(eventTitle);
        event = calendar.getEventSeriesById(eventId).removeAllReminders();
        calUpdateConfirmation += eventTitle + ' successfully moved/updated to ' + startTimeStamp;
        try {
          formMule_logCalUpdate();
        } catch(err) {
        }
      } catch(err) {
        calUpdateConfirmation += eventId + ' not found on calendar: ' + calendar.getName();
      }
      sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation).setFontColor("black");
      if (found==false) {
        sheet.getRange(j+k, eventUpdateStatusIndex+1).setFontColor("red");
      }
      if ((eventId)&&(eventId!="")&&(found==true)) {
        event = calendar.getEventSeriesById(eventId);    
        if (reminderType) { 
        switch(reminderType) {
        case "Email reminder": 
          event = calendar.getEventSeriesById(eventId).addEmailReminder(minBefore);
        break;
        case "Popup reminder":
          event = calendar.getEventSeriesById(eventId).addPopupReminder(minBefore);
        break;
        case "SMS reminder":
          event = calendar.getEventSeriesById(eventId).addSmsReminder(minBefore);
        break;
        default:
          event = calendar.getEventSeriesById(eventId).removeAllReminders();
        }
       }
       rowData.eventId = eventId;
       desc = formMule_fillInTemplateFromObject(descToken, rowData);
       event = calendar.getEventSeriesById(eventId).setDescription(desc); 
     }
     if (guests!='') {
       guests = guests.split(",");
       for (var g=0; g<guests.length; g++) {
         event.addGuest(guests[g]);
       }
      }
    }
  if ((found==true)&&(calDeleteConditionTest == true)&&(calUpdateConditionTest == false)) {
    var eventTitle = event.getTitle();
    event = calendar.getEventSeriesById(eventId).deleteEventSeries();
    calUpdateConfirmation = eventTitle + ' deleted from calendar: ' + calendar.getName();
    sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation).setFontColor("black");
  }
  if ((found==false)&&(calendar)&&((calUpdateConditionTest == true)||(calDeleteConditionTest == true))) {
     calUpdateConfirmation = eventId + ' not found on calendar: ' + calendar.getName();
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation);
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setFontColor("red");
  }
  if ((!(calendar))&&(calUpdateConditionTest == true)) {
     calUpdateConfirmation = "Calendar not found";
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation);
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setFontColor("red");
  }



  if (((emailStatus=="true")&&(manual==true))||((emailStatus=="true")&&(emailTrigger=="true"))) {
  for (var i=0; i<numSelected; i++) {
    var confirmation = '';
    if ((emailCondString)&&(emailCondString!='')) {
      var emailConditionTest = formMule_evaluateConditions(emailCondObject, i, rowData);
    }
    if ((emailConditionTest == true)||(!emailCondString)) {
      var sendersTemplate = sendersTemplates[i];
      var recipientsTemplate = recipientsTemplates[i];
      var ccRecipientsTemplate = ccRecipientsTemplates[i];
      var subjectTemplate = subjectTemplates[i];
      var bodyTemplate = bodyTemplates[i];
      var langTemplate = langTemplates[i];

    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var from = formMule_fillInTemplateFromObject(sendersTemplate, rowData);
    var to = formMule_fillInTemplateFromObject(recipientsTemplate, rowData);
    try {
      if (to!='') {
        var cc = formMule_fillInTemplateFromObject(ccRecipientsTemplate, rowData);
        var subject = formMule_fillInTemplateFromObject(subjectTemplate, rowData);
        var body = formMule_fillInTemplateFromObject(bodyTemplate, rowData); 
        var lang = formMule_fillInTemplateFromObject(langTemplate, rowData);
        body = body.replace("\n", "<br />","g");
        if ((lang)&&(lang!='')) {
          var translation = LanguageApp.translate("Automated Translation", '', lang);
          var divider = '<h3>####### ' + translation + ' #######</h3>';
          body += divider + LanguageApp.translate(body, '', lang)
        }
      if (from=='') {
        MailApp.sendEmail(to, subject, '', {htmlBody: body, cc: cc});
      } else {
        MailApp.sendEmail(to, subject, '', {htmlBody: body, cc: cc, replyTo: from});
      } 
      var now = new Date();
      now = Utilities.formatDate(now, timeZone, "MM/dd/yy' at 'h:mm:ss a")
      if (cc!='') { var ccMsg = ", and cc'd to "+cc; } else {ccMsg = ''}
      if ((i>0)&&(i<numSelected-1)) { var addSemiColon = "; "} else { var addSemiColon = ""; }
        confirmation += templateSheetNames[i]+" sent to "+to+ccMsg+' on '+ now + addSemiColon;
        try {
          formMule_logEmail();
        } catch(err) {
        }
      } else { 
        confirmation += templateSheetNames[i] + " error: Template missing \"To\" address";
        var error = true;
      }
    } catch(err) {
      confirmation += templateSheetNames[i] +" error: " + err;
      var error = true;
    }
    var statusIndex = headers.indexOf(sendColNames[i]);
    var statusCell = sheet.getRange(j+k, statusIndex+1);
    if (confirmation!='') {
    var statusCell = sheet.getRange(j+k, statusIndex+1);
      statusCell.setValue(confirmation).setFontColor("black");
    }
    if (error==true) {
      statusCell.setFontColor("red");
    }
    error == false
   } // end per email conditional check
  } // end i loop through email templates
  } // end conditional test for confirmation email
    //begin SMS section
    if (((smsStatus=="true")&&(manual==true))||((smsStatus=="true")&&(smsTrigger=="true"))) { 
      var twilioNumber = properties.twilioNumber;
      var smsPropertyString = properties.smsPropertyString;
      for (var i=0; i<properties.smsNumSelected; i++) {
        if ((smsPropertyString)&&(emailCondString!='')) {
          var smsPropertyObject = Utilities.jsonParse(smsPropertyString);
          var smsConditionTest = formMule_evaluateSMSConditions(smsPropertyObject, i, rowData);  
          if ((smsConditionTest == true)||(!smsPropertyString)) {
            var phoneNumber = formMule_fillInTemplateFromObject(smsPropertyObject['smsPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(smsPropertyObject['smsLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(smsPropertyObject['smsBody-'+i], rowData, true);
            var flagObject = new Object();
            flagObject.sheetName = sheetName;
            flagObject.header = urlencode(smsPropertyObject['smsName-'+i] + " SMS Status");
            flagObject.row = j+k;
            var args = new Object();
            formMule_sendAText(phoneNumber, body, accountSID, authToken, flagObject, args, lang, maxTexts);
            try {
              formMule_logSMS();
            } catch(err) {
            }
          }
        }
      }
    }
    //end SMS Section
    //begin VM section
    if (((vmStatus=="true")&&(manual==true))||((vmStatus=="true")&&(vmTrigger=="true"))) {
      var vmPropertyString = properties.vmPropertyString;
      for (var i=0; i<properties.vmNumSelected; i++) {
        if ((vmPropertyString)&&(emailCondString!='')) {
          var vmPropertyObject = Utilities.jsonParse(vmPropertyString);
          var vmConditionTest = formMule_evaluateVMConditions(vmPropertyObject, i, rowData);  
          if ((vmConditionTest == true)||(!vmPropertyString)) {
            var phoneNumber = formMule_fillInTemplateFromObject(vmPropertyObject['vmPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(vmPropertyObject['vmLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(vmPropertyObject['vmBody-'+i], rowData, true);
            var flagObject = new Object();
            flagObject.sheetName = sheetName;
            flagObject.header = urlencode(vmPropertyObject['vmName-'+i] + " VM Status");
            flagObject.row = j+k;
            var args = new Object();
            if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
              args.RequestResponse = "TRUE";
            }
            if ((vmPropertyObject['vmSoundFile-'+i])&&(vmPropertyObject['vmSoundFile-'+i])!='') {  
              args.PlayFile = vmPropertyObject['vmSoundFile-'+i];
            }
            formMule_makeRoboCall(phoneNumber, body, accountSID, authToken, flagObject, args, lang);
            try {
              formMule_logVoiceCall();
            } catch(err) {
            }
          }
        }
      }
    } 
    //send VM section
  } //end j loop through spreadsheet rows
} // end function



// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function formMule_fillInTemplateFromObject(template, data, newline) {
  var newString = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);
  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  if (templateVars) {
  for (var i = 0; i < templateVars.length; ++i) {
    // formMule_normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[formMule_normalizeHeader(templateVars[i])];
    newString = newString.replace(templateVars[i], variableData || "");
  }
  }
  var currentTime = new Date();
  var month = (currentTime.getMonth() + 1).toString();
  var day = currentTime.getDate().toString();
  var year = currentTime.getFullYear().toString();
  
  if (!newline) {  newString = newString.replace("\n", "<br />","g"); }
  if (newString.indexOf("$currMonth")!=-1) {
    newString = newString.replace("$currMonth", month);
  }
  if (newString.indexOf("$currDay")!=-1) {
    newString = newString.replace("$currDay", day);
  }
  if (newString.indexOf("$currYear")!=-1) {
    newString = newString.replace("$currYear", year);
  }
  if (newString.indexOf("$formUrl")!=-1) {
    var formUrl = CacheService.getPrivateCache().get('formUrl');
      if (!formUrl) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      formUrl = ss.getFormUrl();
      CacheService.getPrivateCache().put('formUrl', formUrl);
    }
    newString = newString.replace("$formUrl", formUrl);
  }
  if (data.eventId) {
    newString = newString.replace("$eventId", data.eventId);
  }
  return newString;
}





//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// formMule_getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function formMule_getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getDisplayValues()[0];
  return formMule_getObjects(range.getDisplayValues(), formMule_normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function formMule_getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
    //  if (formMule_isCellEmpty(cellData)) {
    //    continue;
    //  }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function formMule_normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = formMule_normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function formMule_normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!formMule_isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && formMule_isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function formMule_isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function formMule_isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    formMule_isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function formMule_isDigit(char) {
  return char >= '0' && char <= '9';
}



function formMule_defineSettings() {
  setSid();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var app = UiApp.createApplication().setTitle("Step 1: Define merge source settings").setHeight(300);
  var panel = app.createVerticalPanel().setId("settingsPanel");  
  var sheetLabel = app.createLabel().setText("Choose the sheet that you want the script to use as a data source for merged emails and/or calendar events.");
  sheetLabel.setStyleAttribute("background", "#E5E5E5").setStyleAttribute("marginTop", "20px").setStyleAttribute("padding", "5px");
  var sheetChooser = app.createListBox().setId("sheetChooser").setName("sheet");
    for (var i=0; i<sheets.length; i++) {
      if ((!sheets[i].getName().match("Template"))&&(!sheets[i].getName().match("formMule Read Me"))&&(!sheets[i].getName().match("Forms in same folder"))) {
      sheetChooser.addItem(sheets[i].getSheetName());   
      }
    }
   if (ScriptProperties.getProperty('sheetName')) {
   var sheetName = ScriptProperties.getProperty('sheetName');
   }
  if (sheetName) {
    var sheetIndex = formMule_getSheetIndex(sheetName);
    sheetChooser.setSelectedIndex(sheetIndex);
  }
  
  var optionsLabel = app.createLabel("Optional").setStyleAttribute("background", "#E5E5E5").setStyleAttribute("marginTop", "20px").setStyleAttribute("padding", "5px");
 
  var caseNoSetting = ScriptProperties.getProperty('caseNoSetting'); 
  var caseNoCheckBox = app.createCheckBox().setId("caseNoCheckBox").setName("caseNoSetting").setText("Auto-create a unique case number for each form submission")
  if (caseNoSetting=="true") {
   caseNoCheckBox.setValue(true);
  }
  var helpPopup = app.createPopupPanel();
  var fieldsLabel = app.createLabel().setText("The formMule script can be used to merge emails and calendar events from any data sheet, however it gains additional power when coupled with a Google Form, VLOOKUP, and other formulas.  See the \"Read Me\" tab in this spreadsheet or visit http://www.youpd.org/formmule for more information about how to use it.");                                         

                                          
  helpPopup.add(fieldsLabel);
  panel.add(helpPopup);
  panel.add(sheetLabel);
  panel.add(sheetChooser);
  panel.add(optionsLabel);
  panel.add(caseNoCheckBox);


  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "190px");
  spinner.setId("dialogspinner");
  
  var copyDownLabel = app.createLabel("Looking to copy formulas down on form submit? This feature is now located in the \"Advanced options\" menu option.").setStyleAttribute('marginTop', '10px');
  panel.add(copyDownLabel);
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  app.add(panel);
  
  var buttonHandler = app.createServerClickHandler('formMule_saveSettings').addCallbackElement(panel);
  var button = app.createButton("Save settings", buttonHandler).setId('saveButton');
  button.addMouseDownHandler(spinnerHandler); 
  panel.add(button);
  app.add(spinner);
  this.ss.show(app);
}


function formMule_saveSettings(e) {
  var app = UiApp.getActiveApplication();
  var oldSheetName = ScriptProperties.getProperty('sheetName');
  var sheetName = e.parameter.sheet;
  ScriptProperties.setProperty('sheetName', sheetName);
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  var lastCol = sheet.getLastColumn();
  
  if (lastCol==0) { 
    Browser.msgBox("You have no headers in your data sheet. Please fix this and come back."); 
    formMule_defineSettings();
    app.close();
    return app; 
  }
    
  var caseNoSetting = e.parameter.caseNoSetting;
  ScriptProperties.setProperty('caseNoSetting', caseNoSetting);
  
  var caseNo = ScriptProperties.getProperty('caseNo');
  var caseNoIndex = headers.indexOf("Case No");
  
  if (caseNoSetting=="true") {
  if (caseNoIndex == -1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1,sheet.getLastColumn()+1).setBackground("orange").setFontColor("black").setValue("Case No").setComment("Don't change the name of this column. It is used to log a unique case number for each form submission.");
   }
  }
  if (caseNoSetting=="false") {
    if ((caseNoIndex != -1)) {
    sheet.deleteColumn(caseNoIndex+1);
    }
  }
  
  var lastCol = sheet.getLastColumn();
  
  var sheetName = ScriptProperties.getProperty('sheetName');
  if ((sheetName)&&(!(oldSheetName))) {
    onOpen();
    Browser.msgBox('Auto-email and auto-calendar-event options are now available in the formMule menu');
  }
  onOpen();
  app.close()
  return app;
}



function formMule_getSheetIndex(sheetName) {
  var app = UiApp.getActiveApplication();
  var sheets = ss.getSheets();
  var bump = 0;
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getName()==sheetName) {
     var index = i;
     break;
    }
     if ((sheets[i].getName()== "formMule Read Me")||(sheets[i].getName()=="Email1 Template")||(sheets[i].getName()== "Email2 Template")||(sheets[i].getName()=="Email3 Template")||(sheets[i].getName()== "Forms in same folder"))  {
        bump = bump-1;
     }
    index = 0;
  }
index = index+bump;
return index;
}

function formMule_refreshGrid (e) {
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById('settingsPanel');
  var gridPanel = app.getElementById('gridPanel');
  gridPanel.setStyleAttribute('opacity', '1');
  var grid = app.getElementById("checkBoxGrid");
  var spinner = app.getElementById("dialogspinner");
  spinner.setVisible(false);
  panel.setStyleAttribute('opacity', '1');
  var sheetName = e.parameter.sheet;
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  grid.resize(headers.length, 1);
  var copyDownOption = e.parameter.copyDownOption;
  if (copyDownOption=="true") {
    gridPanel.setVisible(true);
    gridPanel.setEnabled(true);
  } else {
    gridPanel.setVisible(false);
    gridPanel.setEnabled(false);
  }
  var count = 0;
  for (var i=0; i<headers.length; i++) {
    var normalizedHeader = formMule_normalizeHeader(headers[i]);
    var checkBox = app.createCheckBox().setText(headers[i]).setName(normalizedHeader).setId("checkBox-"+i);
    var cell = sheet.getRange(1, i+1);
    if ((cell.getBackground()=="#DDDDDD")||(cell.getBackground()=="black")||(cell.getBackground()=="orange")||(cell.getBackground()=="#0099FF")||(cell.getBackground()=="#FF9999")||(cell.getBackground()=="#66C285")) { 
      checkBox.setVisible(false); 
    } else {
      count++;
    }
    grid.setWidget(i,0, checkBox); 
   }
  if (count==0) {
    var noneLabel = app.createLabel("You have no columns to the right of your form data.");
    grid.setWidget(0,0, noneLabel);
  }
  return app;
}