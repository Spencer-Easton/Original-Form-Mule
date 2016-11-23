function formMule_loadingPreview() {
  var app = UiApp.createApplication().setWidth(140).setHeight(140);
  var label = app.createLabel("Loading preview...").setStyleAttribute('textAlign', 'center');
  var image = app.createImage(MULEICONURL).setHeight("100px").setStyleAttribute('marginLeft', '20px');
  app.add(label);
  app.add(image);
  ss.show(app);
  formMule_previewSend();
  return app;
}


function formMule_previewSend() {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  app.close();
  var manual = true;
  var app = UiApp.createApplication().setTitle("Here's what your merge will look like...").setWidth(590).setHeight(420);
  var panel = app.createVerticalPanel();
  var tabPanel = app.createTabPanel().setHeight("330px");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = properties.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  var copyDownOption = properties.copyDownOption;
  var lastRow = sheet.getLastRow();
  var emailData = new Array();
  var calData = new Array();
  var smsData = new Array();
  var vmData = new Array();
  var calUpData = new Array();
  var oldCalData = new Array();
  var k = 0;
  var m = 0;
  var n = 0;
  var s = 0;
  var v = 0;
  var calendarStatus = properties.calendarStatus;
  var calendarUpdateStatus = properties.calendarUpdateStatus;
  var calTrigger = properties.calTrigger;
  var calUpdateTrigger = properties.calUpdateTrigger;
  var userEmail = Session.getEffectiveUser().getEmail();
  if (((calendarStatus == "true")&&(manual==true))||((calendarStatus == "true")&&(calTrigger=="true"))||((calendarUpdateStatus == "true")&&(manual==true))||((calendarUpdateStatus == "true")&&(calUpdateTrigger=="true"))) {
    var eventIdCol = properties.eventIdCol;
    var calConditionsString = properties.calendarConditions;
    var calCondObject = Utilities.jsonParse(calConditionsString);
    var calUpdateConditionsString = properties.calendarUpdateConditions;
    var calUpdateCondObject = Utilities.jsonParse(calUpdateConditionsString);
    var calDeleteConditionsString = properties.calendarDeleteConditions;
    var calDeleteCondObject = Utilities.jsonParse(calDeleteConditionsString);
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
    var calRepeats = properties.calendarWeeklyRepeats;
    var calWeekdays = properties.calendarWeekdays;
  }
  var emailStatus = properties.emailStatus;
  var emailTrigger = properties.emailTrigger;
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
    for (var i=0; i<numSelected; i++) {
       var templateSheet = ss.getSheetByName(emailCondObject['sht-'+i]);
       if (!templateSheet) {
         Browser.msgBox("An error occurred: It appears one of your email template sheets was deleted or renamed. Go back to step 2a to fix this.");
         app.close();
         return app;
       }
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
  var smsStatus = properties.smsEnabled;
  var smsTrigger = properties.smsTrigger;
  var vmStatus = properties.vmEnabled;
  var vmTrigger = properties.vmTrigger;
  var accountSID = properties.accountSID;
  var authToken = properties.authToken;
  var maxTexts = properties.smsMaxLength;
  var dataSheetName = properties.sheetName;
  var dataSheet = ss.getSheetByName(dataSheetName);
  var headers = formMule_fetchHeaders(dataSheet);
  var eventIdIndex = headers.indexOf("Event Id");
  var updateEventIdIndex = headers.indexOf(eventIdCol);
  if (eventIdCol) {
    eventIdCol = formMule_normalizeHeader(eventIdCol);
  }
  var z=2;
  if (copyDownOption=="true") {
     z=3;
  }
  var dataRange = dataSheet.getRange(z, 1, dataSheet.getLastRow()-(z-1), dataSheet.getLastColumn());
  var lastHeader = dataSheet.getRange(1, dataSheet.getLastColumn());
  var lastHeaderValue = formMule_normalizeHeader(lastHeader.getValue());  
  var test = dataRange.getValues()
  // Create one JavaScript object per row of data.
  var objects = formMule_getRowsData(dataSheet, dataRange, 1);
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var j = 0; j < objects.length; ++j) {
    // Get a row object
    var rowData = objects[j];
    //Get merge and update status for the row
    var confirmation = '';
    var calUpdateConfirmation = '';
    var found = '';
    
    var calConditionTest = false;
    var calUpdateConditionTest = false;
    var calDeleteConditionTest = false;
    if (calendarStatus=="true") {
      //test calendar event create condition
      if ((calConditionsString)&&(calConditionsString!='')) {
        var calConditionTest = formMule_evaluateConditions(calCondObject, 0, rowData);
      } 
    }
    
    if (calendarUpdateStatus=="true") {
      var calUpdateStatus = rowData.eventUpdateStatus;
      if (calUpdateStatus=='') {
        //test calendar event update condition
        if ((calUpdateConditionsString)&&(calUpdateConditionsString!='')) {
          var calUpdateConditionTest = formMule_evaluateUpdateConditions(calUpdateCondObject, 0, rowData);    
        } 
        //test calendar event deletion condition
        if ((calDeleteConditionsString)&&(calDeleteConditionsString!='')) {
          var calDeleteConditionTest = formMule_evaluateUpdateConditions(calDeleteCondObject, 0, rowData);    
        } 
      }
    }
    
    if ((calendarToken)&&(calendarToken!='')&&((calUpdateConditionTest)||(calConditionTest)||(calDeleteConditionTest))&&(m<7)) {
      var calendarId  = formMule_fillInTemplateFromObject(calendarToken, rowData);
      var lastCalendarId = CacheService.getPrivateCache().get('lastCalendarId');
      if (lastCalendarId!=calendarId) {
        var calendar = CalendarApp.getCalendarById(calendarId);
      }
    }
    if ((!calendarToken)&&((calUpdateConditionTest)||(calConditionTest)||(calDeleteConditionTest))) {
      Browser.msgBox("You have not specified a calendar ID for your calendar merge.  Return to step 2b.");
          return;
    }
    // check for calendar status setting
    if (calConditionTest == true) {
      calData[m] = new Object();
      if (calendar) {
        calData[m].calendar = calendar.getName();
        var eventTitle = formMule_fillInTemplateFromObject(eventTitleToken, rowData);
        calData[m].eventTitle = eventTitle;
        var location = formMule_fillInTemplateFromObject(locationToken, rowData);
        calData[m].location = location;
        var desc = formMule_fillInTemplateFromObject(descToken, rowData);
        calData[m].desc = desc;
        var guests = formMule_fillInTemplateFromObject(guestsToken, rowData);
        calData[m].guests = guests;
        var startTimeStamp;
        var endTimeStamp;
        var timeZone = Session.getTimeZone();
      if (!(allDay=="true")) {
        var startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        endTimeStamp = new Date(formMule_fillInTemplateFromObject(endTimeToken, rowData));
        calData[m].allDay = allDay;
        calData[m].startTimeStamp = startTimeStamp;
        calData[m].endTimeStamp = endTimeStamp;
      } else {
        calData[m].allDay = allDay;
        startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        startTimeStamp = Utilities.formatDate(startTimeStamp, timeZone, "MM/dd/yyyy");
        calData[m].startTimeStamp = startTimeStamp;
        calData[m].endTimeStamp = '';
      }
      calData[m].emailInvites = emailInvites;
      calData[m].repeats = formMule_fillInTemplateFromObject(calRepeats, rowData);
      calData[m].weekdays = formMule_fillInTemplateFromObject(calWeekdays, rowData);
      } else {
        calData[m].calendar = "Calendar could not be found. Check your calendar Id";
      }
    m++;
  } //end calendar create condition test
  if ((calUpdateConditionTest == true)||(calDeleteConditionTest == true)) {  
    var eventId = rowData[eventIdCol];
    calUpData[n] = new Object();
    oldCalData[n] = new Object();
    if (n<7) {
    if (calendar) {
    calUpData[n].calendar = calendar.getName();
    var found = false;
    try {
      var future = new Date();
      future.setDate(future.getDate()+360);
      var past = new Date();
      past.setDate(past.getDate()-360);
      var events = calendar.getEvents(past, future);
      for (var g = 0; g<events.length; g++) {
        if (events[g].getId()==eventId) {
          var oldEvent = events[g];
          found = true;
        }
      }
    } catch(err) {
      found = false;
    }
    } else {
      Browser.msgBox("Calendar could not be found");
      return;
    }
    if ((found == true)&&((calUpdateConditionTest==true)||(calDeleteConditionTest==true))) {
      oldCalData[n].eventTitle = oldEvent.getTitle();
      oldCalData[n].location = oldEvent.getLocation(); 
      oldCalData[n].desc = oldEvent.getDescription();
      oldCalData[n].guests = oldEvent.getGuestList().join();
      if ((found==true)&&(calUpdateConditionTest == true)) {
        var eventTitle = formMule_fillInTemplateFromObject(eventTitleToken, rowData);
        calUpData[n].eventTitle = eventTitle;
        var location = formMule_fillInTemplateFromObject(locationToken, rowData);
        calUpData[n].location = location;
        var desc = formMule_fillInTemplateFromObject(descToken, rowData);
        calUpData[n].desc = desc;
        var guests = formMule_fillInTemplateFromObject(guestsToken, rowData);
        calUpData[n].guests = guests;
        var startTimeStamp;
        var endTimeStamp;
        var timeZone = Session.getTimeZone();
      if (!(allDay=="true")) {
        var startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        endTimeStamp = new Date(formMule_fillInTemplateFromObject(endTimeToken, rowData));
        calUpData[n].allDay = allDay;
        calUpData[n].startTimeStamp = startTimeStamp;
        calUpData[n].endTimeStamp = endTimeStamp;
      } else {
        calUpData[n].allDay = allDay;
        startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        startTimeStamp = Utilities.formatDate(startTimeStamp, timeZone, "MM/dd/yyyy");
        calUpData[n].startTimeStamp = startTimeStamp;
        calUpData[n].endTimeStamp = '';
      }
        calUpData[n].emailInvites = emailInvites;
      }      
    // same as above, but for old event settings
    var oldAllDay = oldEvent.isAllDayEvent();
    if (!(oldAllDay==true)) {
      var startTimeStamp = oldEvent.getStartTime();
      endTimeStamp = oldEvent.getEndTime();
      oldCalData[n].allDay = oldAllDay;
      oldCalData[n].startTimeStamp = startTimeStamp;
      oldCalData[n].endTimeStamp = endTimeStamp;
    } else {
      oldCalData[n].allDay = oldAllDay;
      startTimeStamp = oldEvent.getStartTime();
      startTimeStamp = Utilities.formatDate(startTimeStamp, timeZone, "MM/dd/yyyy");
      oldCalData[n].startTimeStamp = startTimeStamp;
      oldCalData[n].endTimeStamp = '';
    }
      oldCalData[n].emailInvites = oldEvent.getEmailReminders(); 
    }
     if ((found==true)&&(calDeleteConditionTest == true)) {
      calUpData[n].eventTitle = "Event will be deleted";
      calUpData[n].location = "Event will be deleted";
      calUpData[n].desc = "Event will be deleted";
      calUpData[n].guests = "Event will be deleted";
      calUpData[n].allDay = "Event will be deleted";
      calUpData[n].startTimeStamp = "Event will be deleted";
      calUpData[n].endTimeStamp = "Event will be deleted";
      calUpData[n].emailInvites = "Event will be deleted";
    }
      
    if (found==false) {
      oldCalData[n].eventTitle = "Event " + eventId + " could not be found. Check your event id settings.";
    }
   }
    n++;
  } //end calendar update condition test
  // check for email status setting 
  if ((emailStatus=="true")&&(manual==true)) {
  for (var i=0; i<numSelected; i++) {
    if ((emailCondString)&&(emailCondString!='')) {
    var emailConditionTest = formMule_evaluateConditions(emailCondObject, i, rowData);
    }
    if ((emailConditionTest == true)||(!emailCondString)) {
    emailData[k] = new Object();
    emailData[k].num = 'Template used: '+templateSheetNames[i];
    if (k<7) {
      var sendersTemplate = sendersTemplates[i];
      var recipientsTemplate = recipientsTemplates[i];
      var ccRecipientsTemplate = ccRecipientsTemplates[i];
      var subjectTemplate = subjectTemplates[i];
      var bodyTemplate = bodyTemplates[i];
      var langTemplate = langTemplates[i];
      emailData[k].from = userEmail;
    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
      var from = formMule_fillInTemplateFromObject(sendersTemplate, rowData);
      emailData[k].replyTo = from;
      var to = formMule_fillInTemplateFromObject(recipientsTemplate, rowData);
      try {
      if (to!='') {
        emailData[k].to = to;
        var cc = formMule_fillInTemplateFromObject(ccRecipientsTemplate, rowData);
        emailData[k].cc = cc;
        var subject = formMule_fillInTemplateFromObject(subjectTemplate, rowData);
        emailData[k].subject = subject;
        var body = formMule_fillInTemplateFromObject(bodyTemplate, rowData);
        var lang = formMule_fillInTemplateFromObject(langTemplate, rowData);
        if ((lang)&&(lang!='')) {
          var translation = LanguageApp.translate("Automated Translation", '', lang);
          var divider = '<h3>####### ' + translation + ' #######</h3>';
          try {
            body += divider + LanguageApp.translate(body, '', lang);
          } catch(err) {
            body = "A translation error occurred.  Check your language code(s)";
          }
        }
        emailData[k].body = body;
        var now = new Date();
        now = Utilities.formatDate(now, timeZone, "MM/dd/yy' at 'h:mm:ss a")
        if (cc!='') { var ccMsg = ", and cc to "+cc; } else {ccMsg = ''}
        if ((i>0)&&(i<numSelected-1)) { var addSemiColon = "; "} else { var addSemiColon = ""; }
        confirmation += " Will attempt to send Email"+(i+1)+" to "+to+ccMsg+addSemiColon;
        } else {
          emailData[k].to = "missing";
          emailData[k].cc = cc;
          emailData[k].subject = subject;
          emailData[k].body = body;
          confirmation += " Email"+(i+1)+" error: Template missing \"To\" address";
          var error = true;
        }
      } catch(err) {
        confirmation += " Email"+(i+1)+" error: " + err;
        var error = true;
      }
    }
      k++;
     } // end per email conditional check
     } // end i loop through email templates
    } // end conditional test for email
      //begin SMS section
    if (((smsStatus=="true")&&(manual==true))||((smsStatus=="true")&&(smsTrigger=="true"))) { 
      var twilioNumber = properties.twilioNumber;
      var smsPropertyString = properties.smsPropertyString;
      for (var i=0; i<properties.smsNumSelected; i++) {
        if ((smsPropertyString)&&(emailCondString!='')) {
          var smsPropertyObject = Utilities.jsonParse(smsPropertyString);
          var smsConditionTest = formMule_evaluateSMSConditions(smsPropertyObject, i, rowData);  
          if ((smsConditionTest == true)||(!smsPropertyString)) {
            smsData[s] = new Object();
            smsData[s].phoneNumber = formMule_fillInTemplateFromObject(smsPropertyObject['smsPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(smsPropertyObject['smsLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(smsPropertyObject['smsBody-'+i], rowData, true);
            try {
              if ((lang) && (lang!='')) {
                body = LanguageApp.translate(body, '', lang);
              }
            } catch(err) {
              body = 'A translation error occurred. Check your language code(s).';
            }
            smsData[s].body = formMule_splitTexts(body, maxTexts).join(" || ");
            smsData[s].num = 'Template used: '+smsPropertyObject['smsName-'+i];
            s++;
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
            vmData[v] = new Object();
            vmData[v].phoneNumber = formMule_fillInTemplateFromObject(vmPropertyObject['vmPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(vmPropertyObject['vmLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(vmPropertyObject['vmBody-'+i], rowData, true);
            try {
              if ((lang) && (lang!='')) {
                body = LanguageApp.translate(body, '', lang);
              }
            } catch(err) {
              body = 'A translation error occurred. Check your language code(s).';
            }
            vmData[v].body = body;
            if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
              vmData[v].RequestResponse = "Yes";
            } else {
              vmData[v].RequestResponse = "No";
            }
            if ((vmPropertyObject['vmSoundFile-'+i])&&(vmPropertyObject['vmSoundFile-'+i])!='') {  
              vmData[v].SoundFile = vmPropertyObject['vmSoundFile-'+i];
            }
            v++;
          }
        }
      }
    }     
  } //end j loop through spreadsheet rows
  
  //begin email status panel
  var emailStatus = ScriptProperties.getProperty("emailStatus");
  var emailPanel = app.createVerticalPanel().setHeight("270px");
  emailPanel.add(app.createLabel("Email merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  if (emailStatus=="true") {
  var emailScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("260px");
  var emailVerticalPanel = app.createVerticalPanel().setWidth("100%");
  if (emailData.length<6) {
    var previewNum = emailData.length;
  } else {
    var previewNum = 6
  }
  if (previewNum>0) {
  for (var i=0; i<previewNum; i++) {
    var emailLabel = app.createLabel("Email #"+(i+1)+" - "+emailData[i].num+" Template").setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid2 = app.createGrid(6, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid2.setWidget(1, 0, app.createLabel('From:')).setStyleAttribute(1, 0, "width", "100px").setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(2, 0, app.createLabel('Reply to:')).setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(0, 0, app.createLabel('To:')).setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(3, 0, app.createLabel('CC:')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(4, 0, app.createLabel('Subject:')).setStyleAttribute(4,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(5, 0, app.createLabel('Body:')).setStyleAttribute(5,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(1, 1, app.createLabel(emailData[i].from)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(2, 1, app.createLabel(emailData[i].replyTo)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(0, 1, app.createLabel(emailData[i].to)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(3, 1, app.createLabel(emailData[i].cc)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(4, 1, app.createLabel(emailData[i].subject)).setStyleAttribute(4,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(5, 1, app.createHTML(emailData[i].body)).setStyleAttribute(5,1,'backgroundColor', 'whiteSmoke');
    emailVerticalPanel.add(emailLabel);
    emailVerticalPanel.add(grid2);
    }
    var label2 = app.createLabel('Check out the first ' + previewNum + ' of the '+ emailData.length +' emails that will be sent.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    emailPanel.add(label2);
   } else {
    var label2 = app.createLabel('No emails will be sent given your current settings.  This could be because of your email conditions (see Step 2a) or because all your data rows already have a status message in the \'Status\' column for your templates.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    emailPanel.add(label2);
   }
    emailScrollPanel.add(emailVerticalPanel);
    emailPanel.add(emailScrollPanel);
 
  
  if (previewNum>0) {
    var label3 = app.createLabel("Note: Sender address is always that of the script installer. Images and links in the email body currently don\'t preview correctly.  We're working on it;)").setStyleAttribute('color', 'blue');
    emailPanel.add(label3);
  }
  } else {
    emailPanel.add(app.createLabel("Email merge is currently disabled. Want merged emails? You can enable this service in Step 2a.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  tabPanel.add(emailPanel, "Emails");
  
  //begin calendar tab
  var calendarPanel = app.createVerticalPanel().setHeight("270px");
  var calendarStatus = ScriptProperties.getProperty("calendarStatus");
    calendarPanel.add(app.createLabel("Calendar Event Merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  if (calendarStatus=="true") {
  var calScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("260px");
  var calVerticalPanel = app.createVerticalPanel().setWidth("560px");
  if (calData.length<6) {
    var previewCalNum = calData.length;
  } else {
    var previewCalNum = 6
  }
  if (previewCalNum>0) {
  var label4 = app.createLabel('Check out the first ' + previewCalNum + ' of the '+ calData.length +' calendar events that will be created').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
  calendarPanel.add(label4);
  calScrollPanel.add(calVerticalPanel);
  calendarPanel.add(calScrollPanel);
  for (var i=0; i<previewCalNum; i++) {
    var calLabel = app.createLabel("Calendar Event #"+(i+1)).setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid3 = app.createGrid(10, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid3.setWidget(0, 0, app.createLabel('Calendar:')).setStyleAttribute(0, 0, "width", "100px").setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(1, 0, app.createLabel('Event title:')).setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(2, 0, app.createLabel('Event type:')).setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(3, 0, app.createLabel('Start time:')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(4, 0, app.createLabel('End Time:')).setStyleAttribute(4,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(5, 0, app.createLabel('Event Description:')).setStyleAttribute(5,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(6, 0, app.createLabel('Guests:')).setStyleAttribute(6,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(7, 0, app.createLabel('Email invites?')).setStyleAttribute(7,0,'backgroundColor', 'whiteSmoke');
    if (calData[i].repeats!='') {
      grid3.setWidget(8, 0, app.createLabel('Number of weeks to repeat')).setStyleAttribute(8,0,'backgroundColor', 'whiteSmoke');
    }
    if (calData[i].weekdays) {
      grid3.setWidget(9, 0, app.createLabel('Weekdays to repeat on')).setStyleAttribute(9,0,'backgroundColor', 'whiteSmoke');
    }
    grid3.setWidget(0, 1, app.createLabel(calData[i].calendar)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(1, 1, app.createLabel(calData[i].eventTitle)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    if(calData[i].allDay=="true"){var text = "All Day";} else {var text = "Time slot"}
    grid3.setWidget(2, 1, app.createLabel(text)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(3, 1, app.createLabel(calData[i].startTimeStamp)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    if (calData[i].allDay=="true"){var text = "n/a";} else {var text = calData[i].endTimeStamp;}
    grid3.setWidget(4, 1, app.createLabel(text)).setStyleAttribute(4,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(5, 1, app.createLabel(calData[i].desc)).setStyleAttribute(5,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(6, 1, app.createHTML(calData[i].guests)).setStyleAttribute(6,1,'backgroundColor', 'whiteSmoke');
    if(calData[i].emailInvites==true){var text = "Yes";} else {var text = "No"};
    grid3.setWidget(7, 1, app.createHTML(text)).setStyleAttribute(7,1,'backgroundColor', 'whiteSmoke');
    if (calData[i].repeats!='') {
      grid3.setWidget(8, 1, app.createHTML(calData[i].repeats)).setStyleAttribute(8,1,'backgroundColor', 'whiteSmoke');
    }
    if (calData[i].weekdays!='') {
      grid3.setWidget(9, 1, app.createHTML(calData[i].weekdays)).setStyleAttribute(9,1,'backgroundColor', 'whiteSmoke');
    }
    calVerticalPanel.add(calLabel);
    calVerticalPanel.add(grid3);
    }
  } else {
   var label5 = app.createLabel('No calendar appointments will be sent given your current settings.  This could be because of your calendar event creation condition setting (see Step 2b) or because all your data rows already have a status message in the \'Event Creation Status\' column.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    calendarPanel.add(label5);
  }
  } else {
  calendarPanel.add(app.createLabel("Calendar merge is currently disabled. Want merged calendar events? You can enable this service in Step 2b.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  tabPanel.add(calendarPanel, "Event Create")
  
  
  // Begin calendar update panel
  var calendarUpdatePanel = app.createVerticalPanel().setHeight("300px");
  var calendarStatus = ScriptProperties.getProperty("calendarUpdateStatus");
  calendarUpdatePanel.add(app.createLabel("Event Update Merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  var calUpdateVerticalPanel = app.createVerticalPanel().setWidth("560px");
  if (calendarUpdateStatus=="true") {
  var calUpdateScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("290px");
  if (calUpData.length<6) {
    var previewUpdateCalNum = calUpData.length;
  } else {
    var previewUpdateCalNum = 6
  }
  if (previewUpdateCalNum>0) {
  var label5 = app.createLabel('Check out the first ' + previewUpdateCalNum + ' of the '+ calUpData.length +' calendar events that will be updated').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
  calendarUpdatePanel.add(label5);
  calUpdateScrollPanel.add(calUpdateVerticalPanel);
  calendarUpdatePanel.add(calUpdateScrollPanel);
  for (var i=0; i<previewUpdateCalNum; i++) {
    var calUpdateLabel = app.createLabel("Calendar Event Update #"+(i+1)).setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid4 = app.createGrid(9, 3).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid4.setWidget(1, 0, app.createLabel('Calendar:')).setStyleAttribute(1, 0, "width", "100px").setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(2, 0, app.createLabel('Event title:')).setStyleAttribute(2, 0, "width", "100px").setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(3, 0, app.createLabel('Event type:')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(4, 0, app.createLabel('Start time:')).setStyleAttribute(4,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(5, 0, app.createLabel('End Time:')).setStyleAttribute(5,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(6, 0, app.createLabel('Event Description:')).setStyleAttribute(6,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(7, 0, app.createLabel('Guests:')).setStyleAttribute(7,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(8, 0, app.createLabel('Email invites?')).setStyleAttribute(8,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(0, 1, app.createLabel('Current Event Info'));
    grid4.setWidget(0, 2, app.createLabel('Updated Event Info'));
    grid4.setWidget(1, 1, app.createLabel(calUpData[i].calendar)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(2, 1, app.createLabel(oldCalData[i].eventTitle)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(2, 2, app.createLabel(calUpData[i].eventTitle)).setStyleAttribute(2,2,'backgroundColor', 'whiteSmoke');
    if(oldCalData[i].allDay=="true"){var oldText = "All Day";} else {var oldText = "Time slot"}
    grid4.setWidget(3, 1, app.createLabel(oldText)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    if(calUpData[i].allDay=="true"){var newText = "All Day";} else {var newText = "Time slot"}
    grid4.setWidget(3, 2, app.createLabel(newText)).setStyleAttribute(3,2,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(4, 1, app.createLabel(oldCalData[i].startTimeStamp)).setStyleAttribute(4,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(4, 2, app.createLabel(calUpData[i].startTimeStamp)).setStyleAttribute(4,2,'backgroundColor', 'whiteSmoke');
    if (oldCalData[i].allDay=="true"){var text = "n/a";} else {var text = oldCalData[i].endTimeStamp;} 
    if (calUpData[i].allDay=="true"){var newText = "n/a";} else {var newText = calUpData[i].endTimeStamp;}
    grid4.setWidget(5, 1, app.createLabel(text)).setStyleAttribute(5,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(5, 2, app.createLabel(newText)).setStyleAttribute(5,2,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(6, 1, app.createLabel(oldCalData[i].desc)).setStyleAttribute(6,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(6, 2, app.createLabel(calUpData[i].desc)).setStyleAttribute(6,2,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(7, 1, app.createHTML(oldCalData[i].guests)).setStyleAttribute(7,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(7, 2, app.createHTML(calUpData[i].guests)).setStyleAttribute(7,2,'backgroundColor', 'whiteSmoke');
    if(oldCalData[i].emailInvites=="true"){var text = "Yes";} else {var text = "No"};
    grid4.setWidget(8, 1, app.createHTML(text)).setStyleAttribute(8,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(8, 2, app.createHTML("Currently not possible")).setStyleAttribute(8,2,'backgroundColor', 'whiteSmoke');
    calUpdateVerticalPanel.add(calUpdateLabel);
    calUpdateVerticalPanel.add(grid4);
    }
  } else {
   var label6 = app.createLabel('No calendar appointments will be updated given your current settings. This could be because of your update condition setting (see Step 2b) or because all your data rows already have a status message in the \'Event Update Status\' column.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    calendarUpdatePanel.add(label6);
  }
  } else {
  calendarUpdatePanel.add(app.createLabel("The calendar update feature is currently disabled. Want to update merged calendar events? You can enable this service in Step 2b.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  tabPanel.add(calendarUpdatePanel, "Event Update");
  //  End calendar update panel
  
  //begin SMS panel
  var smsStatus = properties.smsEnabled;
  var smsPanel = app.createVerticalPanel().setHeight("300px");
  smsPanel.add(app.createLabel("SMS merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  var smsVerticalPanel = app.createVerticalPanel().setWidth("100%").setHeight("280px");
  var smsScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("300px");
  if (smsStatus=="true") {
  if (smsData.length<6) {
    var smsPreviewNum = smsData.length;
  } else {
    var smsPreviewNum = 6
  }
  if (smsPreviewNum>0) {
  for (var i=0; i<smsPreviewNum; i++) {
    var smsLabel = app.createLabel("SMS #"+(i+1)+" - "+smsData[i].num+" Template").setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid5 = app.createGrid(2, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid5.setWidget(0, 0, app.createLabel('Phone Number(s):')).setStyleAttribute(0, 0, "width", "100px").setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid5.setWidget(1, 0, app.createLabel('Body: (|| shows separation into multiple texts)')).setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid5.setWidget(0, 1, app.createLabel(smsData[i].phoneNumber)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid5.setWidget(1, 1, app.createLabel(smsData[i].body)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    smsVerticalPanel.add(smsLabel);
    smsVerticalPanel.add(grid5);
    }
  
    var label5 = app.createLabel('Check out the first ' + smsPreviewNum + ' of the '+ smsData.length +' SMS messages that will be sent.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    smsVerticalPanel.add(label5);
   } else {
    var label5 = app.createLabel('No SMS messages will be sent given your current settings.  This could be because of your SMS merge conditions (see Step 2c) or because all your data rows already have a status message in the \'Status\' column for your SMS templates.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    smsVerticalPanel.add(label5);
   }
   
  
  if (smsPreviewNum>0) {
    var label6 = app.createLabel("Note: Sender phone number is always that of the Twilio Account.").setStyleAttribute('color', 'blue').setStyleAttribute('marginTop', '2px').setStyleAttribute('marginLeft', '2px');
    smsVerticalPanel.add(label6);
  }
  } else {
    smsVerticalPanel.add(app.createLabel("SMS merge is currently disabled. Want merged SMS Messages? You can enable this service in Step 2c.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  smsScrollPanel.add(smsVerticalPanel);
  smsPanel.add(smsScrollPanel);
 
  tabPanel.add(smsPanel, "SMS");
  
  //End SMS panel
  
   //begin VM panel
  var vmStatus = properties.vmEnabled;
  var vmPanel = app.createVerticalPanel().setHeight("300px");
  vmPanel.add(app.createLabel("VM merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  var vmVerticalPanel = app.createVerticalPanel().setWidth("100%").setHeight("300px");
  var vmScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("300px");
  if (vmStatus=="true") {
  if (vmData.length<6) {
    var vmPreviewNum = vmData.length;
  } else {
    var vmPreviewNum = 6
  }
  if (vmPreviewNum>0) {
    var recordingSheet = ss.getSheetByName("MyVoiceRecordings");
    if (recordingSheet.getLastRow()-1>0) {
      var recordingSelectTexts = recordingSheet.getRange(2, 1, recordingSheet.getLastRow()-1, 1).getValues();
      var recordingSelectValues = recordingSheet.getRange(2, 2, recordingSheet.getLastRow()-1, 1).getValues();
    } else {
      var recordingSelectTexts = [['No messages recorded yet']];
      var recordingSelectValues = [['']];
    }
   
  for (var i=0; i<vmPreviewNum; i++) {
    var vmLabel = app.createLabel("VM #"+(i+1)+" - "+vmData[i].num+" Template").setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid6 = app.createGrid(4, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid6.setWidget(0, 0, app.createLabel('Phone Number(s):')).setStyleAttribute(0, 0, "width", "100px").setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(1, 0, app.createLabel('RoboCall Body:')).setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(2, 0, app.createLabel('Play recording:')).setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(3, 0, app.createLabel('Record reply?')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(0, 1, app.createLabel(vmData[i].phoneNumber)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(1, 1, app.createLabel(vmData[i].body)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    var soundFileIndex = -1;
    for (var k=0; k<recordingSelectValues.length; k++) {
      if (recordingSelectValues[k][0]==vmData[i].SoundFile) {
        soundFileIndex = k;
      }
    }
    
    if ((soundFileIndex==-1)&&(!vmData[i].SoundFile||(vmData[i].SoundFile != ''))) {
      var soundFileName = "Not Found";
      grid6.setWidget(2, 1, app.createLabel(soundFileName)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    } 
    else if (soundFileIndex!=-1) {
      var soundFileName = recordingSelectTexts[soundFileIndex];
      grid6.setWidget(2, 1, app.createAnchor(soundFileName, vmData[i].SoundFile)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke'); 
    }
    if ((!vmData[i].SoundFile)||(vmData[i].SoundFile == '')) {
      var soundFileName = "None";
      grid6.setWidget(2, 1, app.createLabel(soundFileName)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    }

    grid6.setWidget(3, 1, app.createLabel(vmData[i].RequestResponse)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    vmVerticalPanel.add(vmLabel);
    vmVerticalPanel.add(grid6);
    }
    
    var label7 = app.createLabel('Check out the first ' + vmPreviewNum + ' of the '+ vmData.length +' Voice messages that will be sent.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    vmVerticalPanel.add(label7);
   } else {
    var label8 = app.createLabel('No Voice messages will be sent given your current settings.  This could be because of your Voice merge conditions (see Step 2c) or because all your data rows already have a status message in the \'Status\' column for your VM templates.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    vmVerticalPanel.add(label8);
   }
    
 
  
  if (vmPreviewNum>0) {
    var label9 = app.createLabel("Note: Caller phone number is always that of the Twilio Account.").setStyleAttribute('color', 'blue');
    vmVerticalPanel.add(label9);
  }
  } else {
    vmVerticalPanel.add(app.createLabel("Voice merge is currently disabled. Want merged Voice Messages? You can enable this service in Step 2c.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  vmScrollPanel.add(vmVerticalPanel);
  vmPanel.add(vmScrollPanel);
  tabPanel.add(vmPanel, "Voice");
  
  //End VM panel

  var runHandler = app.createServerHandler("formMule_manualSend");
  var exitHandler = app.createServerHandler("formMule_quitUi");
  var buttonPanel = app.createHorizontalPanel().setStyleAttribute('marginTop', '10px');
  var button1 = app.createButton("Exit", exitHandler);
  var button2 = app.createButton("Run merge now", runHandler).addClickHandler(exitHandler).setEnabled(false);
  buttonPanel.add(button1).add(button2);
  if ((previewNum>0)||(previewCalNum>0)||(previewUpdateCalNum>0)||(smsPreviewNum>0)||(vmPreviewNum>0)) {
    button2.setEnabled(true);
  }
  tabPanel.selectTab(0);
  panel.add(tabPanel);
  panel.add(buttonPanel);
  app.add(panel);
  ss.show(app);
  return app;
} // end function
