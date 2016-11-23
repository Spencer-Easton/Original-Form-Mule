//define some useful globals
var https = "https://"
var appName = "api.twilio.com/";
var twilioVersion = "2010-04-01/";
var accounts = "Accounts/"
var webName = appName + twilioVersion + accounts;
var googleLangList = ['Afrikaans: af','Albanian: sq','Arabic: ar','Azerbaijani: az','Basque: eu','Bengali: bn','Belarusian: be','Bulgarian: bg','Catalan: ca','Chinese Simplified: zh-CN','Chinese Traditional: zh-TW','Croatian: hr','Czech: cs','Danish: da','Dutch: nl','English: en','Esperanto: eo','Estonian: et','Filipino: tl','Finnish: fi','French: fr','Galician: gl','Georgian: ka','German: de','Greek: el','Gujarati: gu','Haitian Creole: ht','Hebrew: iw','Hindi: hi','Hungarian: hu','Icelandic: is','Indonesian: id','Irish: ga','Italian: it','Japanese: ja','Kannada: kn','Korean: ko','Latin: la','Latvian: lv','Lithuanian: lt','Macedonian: mk','Malay: ms','Maltese: mt','Norwegian: no','Persian: fa','Polish: pl','Portuguese: pt','Romanian: ro','Russian: ru','Serbian: sr','Slovak: sk','Slovenian: sl','Spanish: es','Swahili: sw','Swedish: sv','Tamil: ta','Telugu: te','Thai: th','Turkish: tr','Ukrainian: uk','Urdu: ur','Vietnamese: vi','Welsh: cy','Yiddish: yi'];
var langList = ['en','en-gb','es','fr','de'];



//function used during installation to verify that the user has in fact made their webapp usable by anonymous visitors (i.e. Twilio)
function formMule_verifyPublic() {
  var url = ScriptApp.getService().getUrl();
  var test = "FALSE";
  if (!url) {
    test = "Error: You have not published your script as a web app. Please revisit the SMS and Voice instructions.";
    return test;
  }
  url +=  "?Verify=TRUE"
  var content = UrlFetchApp.fetch(url).getContentText();
  try {
    test = Xml.parse(content);
    test = test.verified.getText();
    return test;
  } catch(err) {
    test = 'Error: You have published your script as a web app, but you have not made it available to "Anyone, even anonymous."  Please revisit the SMS and Voice instructions.';
    return test;
  }
}


// Ui providing step by step instructions for how to publish the script as a webApp
function formMule_howToPublishAsWebApp() {
  var app = UiApp.createApplication().setTitle('Step 2c. Set up SMS and Voice - Publish formMule as a web app').setHeight(540).setWidth(600);
  var thisSs = SpreadsheetApp.getActiveSpreadsheet();
  ScriptProperties.setProperty('ssId', thisSs.getId());
  var public = formMule_verifyPublic();
  // Once it is determined that the webapp is public, proceed to Twilio setup instructions
  if (public=="TRUE") {
    app.close();
    formMule_howToSetUpTwilio();
    return app;
  }
  var panel = app.createVerticalPanel();
  var handler = app.createServerHandler('formMule_howToSetUpTwilio').addCallbackElement(panel);
  var button = app.createButton("Confirm my settings").addClickHandler(handler);
  var scrollpanel = app.createScrollPanel().setHeight("360px");
  var grid = app.createGrid(9, 2).setBorderWidth(0).setCellSpacing(0);
  var html = app.createHTML('<strong>Important to understand:</strong> Using formMule with the Twilio service is OPTIONAL, but requires publishing this script as a web app, which will provide a URL that Twilio will use to communicate with the script. The instructions below explain how to publish your form as a web app.');
  panel.add(html);
  var text1 = app.createLabel("Instructions:").setStyleAttribute("width", "100%").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setStyleAttribute("padding", "5px 5px 5px 5px");
  panel.add(text1);
  grid.setWidget(0, 0, app.createHTML('1. Go to \'Tools->Script editor\' from the Spreadsheet that contains your form.</li>'));
  grid.setWidget(0, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/tools%20menu.png?attachauth=ANoY7crn7wMOEG9WpO7uZVHw6b6t5H2WtOs3S_lZ7qV8i2uQsizc1DtFE0HUGG3zXDbqa65wbz3YVJ_9m9zGr-wujYf0jRgz-U45s9Sids1HmI02fPwWNCjkp1GDmzViBwcuL-MwlPNu_WOz4J282Sv1EdVOFQ9l2dKrhPwwOPko1QcWqJ8-v2hS1jOxJaM6VC7EbybBeeLJ1Fy7Z9o56lg1fGcqc28WRklYD3pb_zSh_q_WLfwnl-l0lVQ8HjiWzGNyE5kn4J0Y&attredirects=0').setWidth("430px"));
  grid.setStyleAttribute(1, 0, "backgroundColor", "grey").setStyleAttribute(1, 1, "backgroundColor", "grey");
  grid.setWidget(2, 0, app.createHTML('2. Under the \'File\' menu in the Script Editor, select \'Manage versions\' and save a new version of the script. Because it\'s optional, you can leave the \'Describe what changed\' field blank.</li>'));
  grid.setWidget(2, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/file%20menu.png?attachauth=ANoY7cqYPycZqgCH7TseuZRL9wq2lv7xI2bhXhDuQ1BVMp5VkB83wpbQRz3X2q36ouPGZyds-CPzEdX5_o-RfDJfVruf7TUxZddxm9sQQwlY6sUab9nVn8KbDgiqiED0itUzfWKvMfP86ya0twRDnONOtGQmq1XWw26vFLQ82DeaKLASttZ6Mo_Qg-FBloN0iVsMlmLLd5o40C4RbaK7qkXUp5eoEE3eYPf_dpx7u3gBb6gyjIZ8yCSKbajC-HfG32MeS0UtrxTW&attredirects=0').setWidth("430px"));
  grid.setWidget(3, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/manage%20versions.png?attachauth=ANoY7cpVfU-WEwiPuOMKMIAzbK6EdA8xkmv_M2R8GKdlcGLC7mo00ZJykbBFrtJEZQHDpKVdvizQQnuyfGVc65iigmGuGr_ZwC2Z4rnh1V67_ogOJKXH2TWmDAafxa-q_5fngrasDYYN2w2-hR_eR95GoY6e5Rza-mtWb1iAp97Cm8n9kVHRk67dURdrdD5AIaS8ZOkse1MmfaN-ZJpMv7bLYBKpisq8GldTTjo7W55OUIJhFuDcxLEc__vguXArjfb9Pd_e2bZD&attredirects=0').setWidth("430px"));
  grid.setStyleAttribute(4, 0, "backgroundColor", "grey").setStyleAttribute(4, 1, "backgroundColor", "grey");
  grid.setWidget(5, 0, app.createHTML('3. Under the \'Publish\' menu in the Script Editor, select \'Deploy as web app\'. Choose the version you want to publish (usually #1). Under \'Execute web app as\', choose \'me\'. For twilio to access the script, the URL must be visible to anonymous users, so select \'Anyone, even anonymous\'. In the context of this script, this setting can reveal no data to anonymous users.'));
  grid.setWidget(5, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/publish.png?attachauth=ANoY7crNAwLK0jZOIWN9M0oD2EzLC5pDflZJNnG9pZQcUD4liYn3lBFe_5yr9Hlnc5UgU837A28Mow5hsWQuIJh-ihgu_s_wtDqG51m697X7aFrKPBJoE8WQZRhzgPhLKEh_gYi4Cs9KXJb4Ie-suQb65WbOCvxC4fyStExxZrAcYYcws0U5MJqoStfNqLbvb5iBgys8G6IMSce3tNo-cad8UKVW18aB7J0BLPeTO7spj0u-cJVFXVFFPq22pKwCViQSciRyNYaF&attredirects=0').setWidth("430px"));
  grid.setWidget(6, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/webapp.png?attachauth=ANoY7co9bP7Sb4cK3aN9Vf3l57RFiWpQ8fDG6_XbNkx8hG7offLTHQAPXm6JMgHjgpEsruUr5mVnf6ha-pc6phil_1MFWsPs7nD2eMeKV_hylLf0qwXLKS3psU4Mu36iw1xvTG8lA6yUJLHb8jA-yRW1jol3bhBD9LKxTzjavHqjtMffZXohjXIV_d_diuNzK6DYBySmuDzeXIp4LyEabrwPglXk5rN6LOXcJ9vPTunUe5cSLhKlLLW2Vj17A3qwmWXxf94gx-yN&attredirects=0').setWidth("430px"));
  grid.setWidget(7, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/webapp2.png?attachauth=ANoY7crqHkuxcVazt1Bi1_aJwGchj2iACUhPxWHTCa4c8Qrh_jJSgbQQA1CajpsNvIIllJtUIH3yQxCxH-ZiN0wbk_L03Ow7dJorxbqVU5LUjV1ic7gowIbdKhTqS3b3-Sg81PEppEzUlfA8slzI9goYt1KxAXy86RjNHIKUDV1Zo3nORWV7_18GJ5DMpiNB84JV8yHer041WB18JmXYZzpSSJwN3ej00GWo9hknfayVCNy9EhM-QAH8q7pUYJyzwdIRAcW2cayC&attredirects=0').setWidth("430px"));
  scrollpanel.add(grid);
  panel.add(scrollpanel);
  panel.add(button);
  app.add(panel);
  thisSs.show(app);
  return app; 
}


// Ui showing step by step instructions for how to set up Twilio
function formMule_howToSetUpTwilio() {
  var url = ScriptApp.getService().getUrl();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle("Step 2c. Set up SMS and Voice - Get formMule and Twilio talking").setWidth(600).setHeight(500);
  var public = formMule_verifyPublic();
  if (public!="TRUE") {
    app.close();
    Browser.msgBox(public);
    formMule_howToPublishAsWebApp();
    return app;
  }
  var scrollPanel = app.createScrollPanel().setId("scrollPanel").setWidth("595px").setHeight("410px");
  var panel = app.createVerticalPanel().setId("panel");
  var htmlBody = "Congrats on getting formMule published correctly as a web app.  Now on to configuring Twilio, the 3rd party, paid service that can enable SMS and Voice messages to be sent from this script with a few simple setup steps. Twilio's pricing can be found at http://www.twilio.com/voice/pricing<br>";
  var html = app.createHTML(htmlBody);
  app.add(html);
  var linkLabel = app.createLabel("This script's URL for Twilio SMS and Voice Requests:").setStyleAttribute("margin", "10px 0px 10px 0px").setStyleAttribute("padding", "5px 5px 5px 5px");
  var link = app.createTextBox()
  if (!url) {
    Browser.msgBox("Error: You have not yet published this script as a web app. Please launch Step 2c. again and complete the steps outlined in the instructions.");
    app.close();
    return app;
  } else { 
    ScriptProperties.setProperty('webAppUrl', url);
  }
  link.setText(url).setEnabled(false).setWidth("400px").setStyleAttribute("margin", "10px 10px 10px 10px").setStyleAttribute("padding", "5px 5px 5px 5px");
  var text1 = app.createLabel("Instructions:").setStyleAttribute("width", "100%").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setStyleAttribute("padding", "5px 5px 5px 5px");
  app.add(text1);
  var text2 = app.createLabel("1. Create an account at http://www.twilio.com");
  var image2 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/2.jpg?attachauth=ANoY7comDBRVlYUnW6rFi__cI_prSJtORGUDtTjJcGG9QpEI7kNh2MW-u8sxwuB8Ndvs9Nv2HbKQeN2B83DC3YUMjQW-fVd95aMTI7tChp_lxGwAszP5y4ylN5A4rf2_GQhN5nVxv9NdCVf5SfPN1hQiTIS1lNvjRlzjB5voEhbBXq9JLK_apapm6ZHn6qFWfd1vweN6TCtZPiORGJpz235vvxPDxliNZCOvjrzCjf_i-qATTn0TSq4iYy-hHbuUQO2Wpo8zNgtg&attredirects=0").setWidth("430px");
  var text3 = app.createLabel("2. Verify yourself as a human (Twilio will call you and ask you to verify a code), get assigned a free phone number, and proceed to the \"Numbers\" section of your account.  Click on your Twilio phone number.");
  var image3 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/6.jpg?attachauth=ANoY7cpOcS5hG8YQwcs01G5LZYz417BhmzEa3lmQgkp-sFSS-xEdE5mu0kFkuuiQ0rwsE6H18OhnqfrXCODtG6FD0OHp8wXynSdaC6kgDb8TCH0OZD55QEgr7Uc3ltdWmj7RkP9BFaOsJ6cxtFMDtpzD6gvJ0CJTr6St7HvdAtTAvTV3XHi75HUlYLCl2drEaAKvDZr396m04Hv_Ohb2ueaH2FtNxaFBZI07SqsPu6pgQC63bIAiCTPEOBq88nCLL0xqCab_mLVZ&attredirects=0").setWidth("430px");
  var numberLabel = app.createLabel("Enter your Twilio number here");
  var numberBox = app.createTextBox().setName("twilioNumber");
  var text4 = app.createLabel("3. Paste the URL of your published script (below) into the Voice and SMS Request URL fields.  Change both URL methods to GET.  Save Changes");
  var image4 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/9.jpg?attachauth=ANoY7cp8HufXi1K_3FamZ8SbLFEd3_9DG0ToNsR7vYTvrIzHyKieNlJZn42NJamy9NKxyXdDLgRIl9hXPRx3oWxVRJSVVss3TrFVw2Q2HM28XmylowPDOFyYNLjDW1r-acFaoag74hxzKSLyUAQ6CUB5J_Zheof1Q6XLuE1JZf5dTTXYOtvPANMvrBWrRjOwJ2u9tUnNXAcJfb3BG11MClNWV4d6t1woiexK00zaU6l9bi3QgKDRFZEJXzNhYDnvvlP6qA_SaZQr&attredirects=0").setWidth("430px");
  var fieldPanel = app.createGrid(1,4);
  var accountLabel = app.createLabel("Account SID");
  var accountLabelBox = app.createTextBox().setName('accountSID');
  var accountSID = ScriptProperties.getProperty('accountSID');
  if (accountSID) {
    accountLabelBox.setText(accountSID);
  }
  var authTokenLabel = app.createLabel("Auth Token");
  var authTokenBox = app.createTextBox().setName('authToken');
  var authToken = ScriptProperties.getProperty('authToken');
  if (authToken) {
    authTokenBox.setText(authToken);
  }
  var text5 = app.createLabel("4. Paste and save your Account SID and Auth Token below");
  fieldPanel.setWidget(0, 0, accountLabel).setWidget(0, 1, accountLabelBox).setWidget(0, 2, authTokenLabel).setWidget(0, 3, authTokenBox);
  var image5 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/1.jpg?attachauth=ANoY7cqnvcVDrWwnN9p3T_tZ1fJ2eDb-AS4JszGViXjYExOVyaW8u6m4CdUGGw460rm7eYhgbKZWqwO4DM9ugkI7vTWX23ZCFVPBE7XiPyz8l6Oeie7To7GGOJKXJTo4Lbs6ufiLVST7iQpgW-0Q3uRKvozd3jUFfrRKO4TMozdZdEjQnLAF2XECiBR1gGBNcxfjJJs1sxAFJNeBuNf5u4jvtt9BYXCD7NGRBRAm2Tp2TiTUNctwbQUqHx8i6-AFYsvSmuM-jfiL&attredirects=0").setWidth("430px");
  var saveHandler = app.createServerHandler('formMule_saveTwilioSettings').addCallbackElement(scrollPanel);
  var button = app.createButton("Verify settings").addClickHandler(saveHandler).setId("testButton");
  var grid = app.createGrid(14, 2).setBorderWidth(0).setCellSpacing(0).setId("grid");
  grid.setWidget(0, 0, text2).setWidget(0, 1, image2);
  grid.setStyleAttribute(1, 0, "backgroundColor", "grey").setStyleAttribute(1, 1, "backgroundColor", "grey");
  grid.setWidget(2, 0, text3).setWidget(2, 1, image3);
  grid.setStyleAttribute(3, 0, "backgroundColor", "grey").setStyleAttribute(3, 1, "backgroundColor", "grey");
  grid.setWidget(4, 0, text4).setWidget(4, 1, image4);
  grid.setWidget(5, 0, linkLabel).setWidget(5, 1, link);
  grid.setStyleAttribute(6, 0, "backgroundColor", "grey").setStyleAttribute(6, 1, "backgroundColor", "grey");
  grid.setWidget(7, 0, text5).setWidget(7, 1, image5);
  grid.setWidget(8, 0, button).setWidget(8, 1, fieldPanel);
  scrollPanel.add(grid);
  app.add(scrollPanel);
  ss.show(app);
}


//Saves Twilio settings
function formMule_saveTwilioSettings(e) {
  var app = UiApp.getActiveApplication();
  var testButton = app.getElementById('testButton');
  var accountSID = e.parameter.accountSID;
  var authToken = e.parameter.authToken;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  ScriptProperties.setProperty("ssId", ssId);
  if ((accountSID!='')&&(authToken!='')) {
    ScriptProperties.setProperty("accountSID", accountSID);
    ScriptProperties.setProperty("authToken", authToken);
  }
  var grid = app.getElementById("grid");
  var numberList  = app.createListBox().setName("twilioNumber");
  var phoneNumbers = formMule_getPhoneNumbers(accountSID,authToken);
  var saveHandler = app.createServerHandler('formMule_saveTwilioNumber').addCallbackElement(grid);
  var button = app.createButton("Save settings").setEnabled(false).addClickHandler(saveHandler);
  for (var i = 0; i<phoneNumbers.length; i++) {
    numberList.addItem(phoneNumbers[i]);
  }
  grid.setStyleAttribute(9, 0, "backgroundColor", "grey").setStyleAttribute(9, 1, "backgroundColor", "grey");
  // If no phone numbers are returned by Twilio, warn the user
  if ((!phoneNumbers)||(phoneNumbers.length==0)) {
    button.setEnabled(false);
    testButton.setEnabled(true);   
    grid.setWidget(10, 0, app.createLabel("There appears to be a problem with your settings. Please double check all steps and try again."))
  } else {
    button.setEnabled(true);
    testButton.setEnabled(false);
    grid.setWidget(10, 0, app.createLabel("Twilio phone number(s) successfully detected. If more than one, select which one you want this script to use."));
    var twilioNumber = ScriptProperties.getProperty('twilioNumber');
    if ((twilioNumber)&&(phoneNumbers.indexOf(twilioNumber)!=-1)) {
      numberList.setSelectedIndex(phoneNumbers.indexOf(twilioNumber));
    }
    grid.setWidget(10, 1, numberList);
    grid.setStyleAttribute(11, 0, "backgroundColor", "grey").setStyleAttribute(11, 1, "backgroundColor", "grey");
    var langLabel = app.createLabel("Select default language (from those supported by Twilio) for voice messages");
    var langSelectBox = app.createListBox().setName("defaultLang");
    for (var i=0; i<langList.length; i++) {
      langSelectBox.addItem(langList[i]);
    }
    var langSelected = ScriptProperties.getProperty('defaultVoiceLang');
    if (langSelected) {
      var index = langList.indexOf(langSelected);
      if (index!=-1) {
        langSelectBox.setSelectedIndex(index);
      }
    }
    grid.setWidget(12, 0, langLabel);
    grid.setWidget(12, 1, langSelectBox);
    grid.setWidget(13, 0, button);
  }
  return app;
}

function formMule_saveTwilioNumber(e) {
  var app = UiApp.getActiveApplication();
  var twilioNumber = e.parameter.twilioNumber;
  ScriptProperties.setProperty('twilioNumber', twilioNumber);
  var defaultLang = e.parameter.defaultLang;
  ScriptProperties.setProperty('defaultVoiceLang', defaultLang);
  onOpen();
  Browser.msgBox("Nice work! Now for a bit of magic: Try texting \n \"Woot woot\" to " + twilioNumber + " to see if the script is able to receive and send text messages");
  app.close();
  return app;
}


function formMule_smsAndVoiceSettings(tabIndex) {
  var properties = ScriptProperties.getProperties();
  if (!tabIndex) { 
    tabIndex = 0;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheetName = sheet.getName();
    ScriptProperties.setProperty('sheetName', sheetName);
    Browser.msgBox("No source sheet detected. Source was automatically set to the top sheet: " + sheetName);
  }
  try {
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  } catch(err) {
    sheet.getRange(1,1,1,3).setValues([['Dummy Header 1', 'Dummy Header 2', 'Dummy Header 3']]);
    sheet.setFrozenRows(1);
    Browser.msgBox("Your source sheet must have headers to complete this step");
    formMule_smsAndVoiceSettings();
    return;
  }
  var app = UiApp.createApplication().setTitle("Step 2c: Set up SMS message and voice message merge").setWidth("640").setHeight("450");
  var mainPanel = app.createVerticalPanel().setId("mainPanel").setWidth("640px").setHeight("400px");
  var refreshPanel = app.createVerticalPanel().setId('refreshPanel').setVisible(false);
  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "190px");
  spinner.setId("dialogspinner");
  refreshPanel.add(spinner);
  
  
  var smsGrid = app.createGrid(1, 2);
  var vmGrid = app.createGrid(1, 2);
  var tabPanel = app.createTabPanel();
  
  var smsPanel = app.createVerticalPanel().setWidth("430px").setHeight("360px");
  var smsFeaturePanel = app.createVerticalPanel().setStyleAttribute("backgroundColor", "whiteSmoke").setWidth("100%").setHeight("100px");
  var smsEnabledCheckBox = app.createCheckBox("Turn on SMS merge feature").setName("smsEnabled").setStyleAttribute("color", "purple");
  if (properties.smsEnabled=="true") {
    smsEnabledCheckBox.setValue(true);
  }
  var smsOnTriggerCheckBox = app.createCheckBox("Trigger this feature on form submit").setName("smsTrigger").setStyleAttribute("color", "purple");
  if (properties.smsTrigger=="true") {
    smsOnTriggerCheckBox.setValue(true);
  }
  var smsLengthPanel = app.createHorizontalPanel();
  var smsLengthLabel = app.createLabel("Max SMS length").setStyleAttribute("padding", "5px");
  var smsLengthSelect = app.createListBox().setName('smsMaxLength');
  smsLengthSelect.addItem("160 char (1 msg per send)", '1');
  smsLengthSelect.addItem("320 char (2 msgs per send)", '2');
  smsLengthSelect.addItem("480 char (3 msgs per send)", '3');
  if (properties.smsMaxLength) {
      smsLengthSelect.setSelectedIndex(parseInt(properties.smsMaxLength)-1);
    }
  smsLengthPanel.add(smsLengthLabel);
  smsLengthPanel.add(smsLengthSelect);
  var smsNumSelectPanel = app.createHorizontalPanel();
  var smsNumSelectHandler = app.createServerHandler('refreshSmsTemplateGrid').addCallbackElement(tabPanel);
  var smsNumSelect = app.createListBox().setName('smsNumSelected');
  smsNumSelect.addItem('1')
    .addItem('2')
    .addItem('3');
  var smsNumSelected = 1;
  if(properties.smsNumSelected) {
    smsNumSelected = properties.smsNumSelected;
  }
  smsNumSelect.setSelectedIndex(smsNumSelected-1);
  smsNumSelect.addChangeHandler(smsNumSelectHandler);
  var smsNumLabel = app.createLabel('unique possible SMS message(s) per row').setStyleAttribute("padding", "5px");
  smsNumSelectPanel.add(smsNumSelect).add(smsNumLabel);
  
  smsFeaturePanel.setStyleAttribute("padding", "5px");
  smsFeaturePanel.add(smsEnabledCheckBox);
  smsFeaturePanel.add(smsOnTriggerCheckBox);
  smsFeaturePanel.add(smsLengthPanel);
  smsFeaturePanel.add(smsNumSelectPanel);
  smsPanel.add(smsFeaturePanel)
  
  var smsTemplatePanel = app.createVerticalPanel().setHeight("250px")
  var smsTemplateScrollPanel = app.createScrollPanel().setHeight("250px");
  var smsTemplateGrid = app.createGrid(smsNumSelected*4, 5).setId('smsTemplateGrid').setBorderWidth(0).setCellSpacing(0);
  var smsPropertyString = properties.smsPropertyString;
  var smsPropertyObject = new Object();
  if ((smsPropertyString)&&(smsPropertyString!='')) {
    smsPropertyObject = Utilities.jsonParse(smsPropertyString);
  }
  
  for (var i = 0; i<smsNumSelected; i++) {
    smsTemplateGrid.setWidget(0+(i*4), 0, app.createLabel('Template Name'));
    smsTemplateGrid.setWidget(1+(i*4), 0, app.createLabel('Phone #'));
    smsTemplateGrid.setWidget(2+(i*4), 0, app.createLabel('Body'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 0, 'backgroundColor','grey'); 
    var smsTemplateNameBox = app.createTextBox().setName('smsName-'+i).setWidth("140px");
    if (smsPropertyObject['smsName-'+i]) {
      smsTemplateNameBox.setValue(smsPropertyObject['smsName-'+i]);
    }
    var smsPhoneBox = app.createTextBox().setName('smsPhone-'+i).setWidth("140px");
    if (smsPropertyObject['smsPhone-'+i]) {
      smsPhoneBox.setValue(smsPropertyObject['smsPhone-'+i]);
    }
    var smsBodyArea = app.createTextArea().setName('smsBody-'+i).setWidth("140px").setHeight("110px");
    if (smsPropertyObject['smsBody-'+i]) {
      smsBodyArea.setValue(smsPropertyObject['smsBody-'+i]);
    }
    smsTemplateGrid.setWidget(0+(i*4), 1, smsTemplateNameBox);
    smsTemplateGrid.setWidget(1+(i*4), 1, smsPhoneBox);
    smsTemplateGrid.setWidget(2+(i*4), 1, smsBodyArea);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 1, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(0+(i*4), 2, app.createLabel('send if'));
    var smsCondCol = app.createListBox().setName('smsCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      smsCondCol.addItem(headers[j]);
    }
    if (smsPropertyObject['smsCol-'+i]) {
      var index = headers.indexOf(smsPropertyObject['smsCol-'+i]);
      smsCondCol.setSelectedIndex(index);
    }
    smsTemplateGrid.setWidget(1+(i*4), 2, smsCondCol);
    smsTemplateGrid.setWidget(2+(i*4), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 2, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(1+(i*4), 3, app.createLabel("="));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 3, 'backgroundColor','grey'); 
    var smsCondVal = app.createTextBox().setName('smsVal-'+i).setWidth("100px");
    if (smsPropertyObject['smsVal-'+i]) {
      smsCondVal.setValue(smsPropertyObject['smsVal-'+i]);
    }
    var smsLang = app.createTextBox().setName('smsLang-'+i).setWidth("50px");
    if (smsPropertyObject['smsLang-'+i]) {
      smsLang.setValue(smsPropertyObject['smsLang-'+i]);
    } else {
      smsLang.setValue('en');
    }
    smsTemplateGrid.setWidget(1+(i*4), 4, smsCondVal);
    smsTemplateGrid.setWidget(2+(i*4), 4, smsLang);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 4, 'backgroundColor','grey'); 
  }
  smsTemplateScrollPanel.add(smsTemplateGrid);
  smsTemplatePanel.add(app.createLabel("SMS Template(s)").setStyleAttribute("color", "white").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("width", "100%").setStyleAttribute('padding', '5px'));
  smsTemplatePanel.add(smsTemplateScrollPanel);
  smsPanel.add(smsTemplatePanel);
  smsGrid.setWidget(0, 0, smsPanel);
  
  
  
  var vmPanel = app.createVerticalPanel().setWidth("430px").setHeight("350px");
  var vmFeaturePanel = app.createVerticalPanel().setStyleAttribute("backgroundColor", "whiteSmoke").setWidth("100%").setHeight("80px").setStyleAttribute("padding", "5px");
  var vmEnabledCheckBox = app.createCheckBox("Turn on Voice Message merge feature").setName('vmEnabled').setStyleAttribute("color", "#FF6600");
  if (properties.vmEnabled=="true") {
    vmEnabledCheckBox.setValue(true);
  }
  var vmOnTriggerCheckBox = app.createCheckBox("Trigger this feature on form submit").setName('vmTrigger').setStyleAttribute("color", "#FF6600");
  if (properties.vmTrigger=="true") {
    vmOnTriggerCheckBox.setValue(true);
  }
  var vmNumSelectPanel = app.createHorizontalPanel();
  var vmNumSelectHandler = app.createServerHandler('refreshVmTemplateGrid').addCallbackElement(tabPanel);
  var vmNumSelect = app.createListBox().setName('vmNumSelected');
  vmNumSelect.addItem('1')
    .addItem('2')
    .addItem('3');
  var vmNumSelected = 1;
  if(properties.vmNumSelected) {
    vmNumSelected = properties.vmNumSelected;
  }
  vmNumSelect.setSelectedIndex(vmNumSelected-1);
  vmNumSelect.addChangeHandler(vmNumSelectHandler);
  var vmNumLabel = app.createLabel('unique possible Voice Message(s) per row').setStyleAttribute("padding", "5px");
  vmNumSelectPanel.add(vmNumSelect).add(vmNumLabel);
  
  vmFeaturePanel.add(vmEnabledCheckBox);
  vmFeaturePanel.add(vmOnTriggerCheckBox);
  vmFeaturePanel.add(vmNumSelectPanel);
  vmPanel.add(vmFeaturePanel);

  var vmTemplatePanel = app.createVerticalPanel().setHeight("270px")
  var vmTemplateScrollPanel = app.createScrollPanel().setHeight("270px");
  var vmTemplateGrid = app.createGrid(vmNumSelected*6, 5).setId("vmTemplateGrid").setBorderWidth(0).setCellSpacing(0);
  var recordingSheet = ss.getSheetByName("MyVoiceRecordings");
  if (!recordingSheet) {
    recordingSheet = ss.insertSheet("MyVoiceRecordings");
    var recordingHeaders = [['Recording Name','Recording URL']];
    recordingSheet.getRange(1, 1, 1, 2).setValues(recordingHeaders).setComment("Don't change the name of this column or sheet. formMule needs these to run SMS and Voice merge services.");
    recordingSheet.setColumnWidth(2, 800);
    recordingSheet.setFrozenRows(1);
    Browser.msgBox("formMule just automatically created a sheet to store the URLs of your outbound voice messages");
    formMule_smsAndVoiceSettings();
    app.close();
    return app;
  }
  if (recordingSheet.getLastRow()-1>0) {
    var recordingSelectTexts = recordingSheet.getRange(2, 1, recordingSheet.getLastRow()-1, 1).getValues();
    var recordingSelectValues = recordingSheet.getRange(2, 2, recordingSheet.getLastRow()-1, 1).getValues();
  } else {
     var recordingSelectTexts = [];
    var recordingSelectValues = [];
  }
  
  var vmPropertyString = properties.vmPropertyString;
  var vmPropertyObject = new Object();
  if ((vmPropertyString)&&(vmPropertyString!='')) {
    vmPropertyObject = Utilities.jsonParse(vmPropertyString);
  }
  
  for (var i = 0; i<vmNumSelected; i++) {
    vmTemplateGrid.setWidget(0+(i*6), 0, app.createLabel('Template Name'));
    vmTemplateGrid.setWidget(1+(i*6), 0, app.createLabel('Phone #'));
    vmTemplateGrid.setWidget(2+(i*6), 0, app.createLabel('Read Message (leave blank for none)'));
    vmTemplateGrid.setWidget(3+(i*6), 0, app.createLabel('Play recording'))
    vmTemplateGrid.setWidget(4+(i*6), 0, app.createLabel('Record response?'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 0, 'backgroundColor','grey'); 
    var vmTemplateNameBox = app.createTextBox().setName('vmName-'+i).setWidth("140px");
    if (vmPropertyObject['vmName-'+i]) {
      vmTemplateNameBox.setValue(vmPropertyObject['vmName-'+i]);
    }
    var vmPhoneBox = app.createTextBox().setName('vmPhone-'+i).setWidth("140px");
    if (vmPropertyObject['vmPhone-'+i]) {
      vmPhoneBox.setValue(vmPropertyObject['vmPhone-'+i]);
    }
    var vmBodyArea = app.createTextArea().setName('vmBody-'+i).setWidth("140px").setHeight("90px");
    if (vmPropertyObject['vmBody-'+i]) {
      vmBodyArea.setValue(vmPropertyObject['vmBody-'+i]);
    }
    var vmSoundFileListBox = app.createListBox().setName('vmSoundFile-'+i).setWidth("140px");
    var soundFileArray = [];
    vmSoundFileListBox.addItem("None", "");
    var m = 0;
    for (var k=0; k<recordingSelectTexts.length; k++) {
      vmSoundFileListBox.addItem(recordingSelectTexts[k][0], recordingSelectValues[k][0]);
      soundFileArray.push(recordingSelectValues[k][0]);
    }
    if (vmPropertyObject['vmSoundFile-'+i]) {
      var index = soundFileArray.indexOf(vmPropertyObject['vmSoundFile-'+i]);
      if (index!=-1) {
        vmSoundFileListBox.setSelectedIndex(index+1);
        m = index;
      }
    }
    var vmSoundFileLinkPanel = app.createHorizontalPanel();
    var vmSoundFilePlayLinkRefreshHandler = app.createServerHandler('refreshVmSoundFilePlayLink').addCallbackElement(vmSoundFileListBox).addCallbackElement(vmSoundFileLinkPanel);  
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkRefreshHandler);
    var linkText = "";
    if ((vmPropertyObject['vmSoundFile-'+i]!="")&&(recordingSelectValues.length>0)) {
      linkText = "Play";
      var vmSoundFilePlayLink = app.createAnchor(linkText,recordingSelectValues[m][0]).setId('vmSoundFileLink-'+i);
    } else {
      var vmSoundFilePlayLink = app.createAnchor(linkText,'').setId('vmSoundFileLink-'+i);
    }
    var vmSoundFilePlayLinkClientHandler = app.createClientHandler().forTargets(vmSoundFilePlayLink).setStyleAttribute('color','grey');
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkClientHandler);
    var vmSoundFileIndex = app.createTextBox().setValue(i).setName('vmIndex').setVisible(false);
    vmSoundFileLinkPanel.add(vmSoundFilePlayLink).add(vmSoundFileIndex);
    var vmRecordOption = app.createListBox().setName('vmRecordOption-'+i).addItem("No").addItem("Yes");
    if (vmPropertyObject['vmRecordOption-'+i]) {
      if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
        vmRecordOption.setSelectedIndex(1);
      }
    }
    var vmRecordHandler = app.createServerHandler('recordVm').addCallbackElement(mainPanel);
    var vmRecordButton = app.createButton('Call me to record voice message').addClickHandler(vmRecordHandler);
    vmTemplateGrid.setWidget(0+(i*6), 1, vmTemplateNameBox);
    vmTemplateGrid.setWidget(1+(i*6), 1, vmPhoneBox);
    vmTemplateGrid.setWidget(2+(i*6), 1, vmBodyArea);
    vmTemplateGrid.setWidget(3+(i*6), 1, vmSoundFileListBox);
    vmTemplateGrid.setWidget(3+(i*6), 2, vmSoundFileLinkPanel);
    vmTemplateGrid.setWidget(3+(i*6), 4, vmRecordButton);
    vmTemplateGrid.setWidget(4+(i*6), 1, vmRecordOption);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 1, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(0+(i*6), 2, app.createLabel('call if'));
    var vmCondCol = app.createListBox().setName('vmCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      vmCondCol.addItem(headers[j]);
    }
    if (vmPropertyObject['vmCol-'+i]) {
      var index = headers.indexOf(vmPropertyObject['vmCol-'+i]);
      if (index!=-1) {
        vmCondCol.setSelectedIndex(index);
      }
    }
    vmTemplateGrid.setWidget(1+(i*6), 2, vmCondCol);
    vmTemplateGrid.setWidget(2+(i*6), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 2, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(1+(i*6), 3, app.createLabel("="));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 3, 'backgroundColor','grey'); 
    var vmCondVal = app.createTextBox().setName('vmVal-'+i).setWidth("100px");
    if (vmPropertyObject['vmVal-'+i]) {
      vmCondVal.setValue(vmPropertyObject['vmVal-'+i]);
    }
    var vmLang = app.createTextBox().setName('vmLang-'+i).setWidth("50px");
    if (vmPropertyObject['vmLang-'+i]) {
        vmLang.setValue(vmPropertyObject['vmLang-'+i]);
    } else {
        vmLang.setValue("en");
    }
    vmTemplateGrid.setWidget(1+(i*6), 4, vmCondVal);
    vmTemplateGrid.setWidget(2+(i*6), 4, vmLang);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 4, 'backgroundColor','grey'); 
  }
  vmTemplateScrollPanel.add(vmTemplateGrid);
  vmTemplatePanel.add(app.createLabel("Voice Message(s)").setStyleAttribute("color", "white").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("width", "100%").setStyleAttribute('padding', '5px'));
  vmTemplatePanel.add(vmTemplateScrollPanel);
  vmPanel.add(vmTemplatePanel);
  vmGrid.setWidget(0,0,vmPanel);
 
  
  var inboundPanel = app.createVerticalPanel().setWidth("620px").setHeight("360px");
  var inboundSMSPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setWidth("425px");
  var inboundSMSLabel = app.createLabel("Inbound SMS Handling").setStyleAttribute('width', '425px').setStyleAttribute('backgroundColor', 'grey').setStyleAttribute('color', 'white').setStyleAttribute('margin', '5px').setStyleAttribute('padding', '5px');
  var inboundTextCheckBox = app.createCheckBox("send this text reply to all inbound SMS messages").setName('smsAutoReplyOption').setStyleAttribute('margin', '5px');
  if (properties.smsAutoReplyOption == "true") {
    inboundTextCheckBox.setValue(true);
  }
  var inboundTextBody = app.createTextArea().setName('smsAutoReplyBody').setWidth("425px").setStyleAttribute('margin', '5px');
  if (properties.smsAutoReplyBody) {
    inboundTextBody.setValue(properties.smsAutoReplyBody);
  }
  var inboundTextForwardPanel = app.createHorizontalPanel().setStyleAttribute('margin', '5px')
  var inboundTextForwardCheckBox = app.createCheckBox('forward all inbound SMS messages to this email address').setName('smsAutoForwardOption').setStyleAttribute('padding', '5px');
  if (properties.smsAutoForwardOption=="true") {
    inboundTextForwardCheckBox.setValue(true);
  }
  var inboundTextForwardEmail = app.createTextBox().setName('smsAutoForwardEmail').setStyleAttribute('margin', '5px').setWidth("130px");
  if (properties.smsAutoForwardEmail) {
    inboundTextForwardEmail.setValue(properties.smsAutoForwardEmail);
  }
  inboundTextForwardPanel.add(inboundTextForwardCheckBox).add(inboundTextForwardEmail);
  inboundSMSPanel.add(inboundSMSLabel).add(inboundTextCheckBox).add(inboundTextBody).add(inboundTextForwardPanel);
  var inboundVoicePanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setWidth("425");
  
  var inboundVoiceLabel = app.createLabel("Inbound Call Handling").setStyleAttribute('width', '425px').setStyleAttribute('backgroundColor', 'grey').setStyleAttribute('color', 'white').setStyleAttribute('margin', '5px').setStyleAttribute('marginTop', '8px').setStyleAttribute('padding', '5px');
  var inboundVoiceCheckBox = app.createCheckBox("read this message to all inbound callers").setName('vmAutoReplyReadOption').setStyleAttribute('margin', '5px');
  if (properties.vmAutoReplyReadOption=="true") {
    inboundVoiceCheckBox.setValue(true);
  }
  var inboundVoiceBody = app.createTextArea().setWidth("425px").setName('vmAutoReplyBody').setStyleAttribute('margin', '5px');
  if (properties.vmAutoReplyBody) {
    inboundVoiceBody.setValue(properties.vmAutoReplyBody);
  }
  var inboundVoiceMessagePanel = app.createHorizontalPanel();
  var inboundVoiceMessageCheckBox = app.createCheckBox("play this voice recording").setName('vmAutoReplyPlayOption').setStyleAttribute('margin', '5px');
  if (properties.vmAutoReplyPlayOption=="true") {
    inboundVoiceMessageCheckBox.setValue(true);
  }
  var inboundSoundFileListBox = app.createListBox().setName('vmAutoReplyFile').setWidth("150px");
  var m = 0;
  var inboundSoundFileChangeHandler = app.createServerHandler('refreshInboundPlayLink').addCallbackElement(inboundVoiceMessagePanel);
  inboundSoundFileListBox.addChangeHandler(inboundSoundFileChangeHandler);
  inboundSoundFileListBox.addItem("None", "")
  for (var k=0; k<recordingSelectTexts.length; k++) {
    inboundSoundFileListBox.addItem(recordingSelectTexts[k][0], recordingSelectValues[k][0]);
    if (properties.vmAutoReplyFile) {
      if (properties.vmAutoReplyFile==recordingSelectValues[k][0]) {
        inboundSoundFileListBox.setSelectedIndex(k+1);
        m = k;
      }
    }
  }
  if (tabIndex==2) {
    inboundSoundFileListBox.setSelectedIndex(recordingSelectTexts.length);
  }
  var inboundRecordHandler = app.createServerHandler('inboundRecordVm').addCallbackElement(mainPanel);
  if (recordingSelectValues.length>0) {
    if (properties.vmAutoReplyFile!='') {
      var inboundPlayLink = app.createAnchor("Play", recordingSelectValues[m][0]).setId('inboundPlayLink');
    } else {
      var inboundPlayLink = app.createLabel("").setId('inboundPlayLink');
    }
  } else {
    var inboundPlayLink = app.createLabel("").setId('inboundPlayLink');
  }
  var inboundSoundFileDissapearHandler = app.createClientHandler().forTargets(inboundPlayLink).setStyleAttribute('color','grey');
  inboundSoundFileListBox.addChangeHandler(inboundSoundFileDissapearHandler);
  var inboundRecordButton = app.createButton('Call me to record voice message').addClickHandler(inboundRecordHandler);
  inboundVoiceMessagePanel.add(inboundVoiceMessageCheckBox).add(inboundSoundFileListBox).add(inboundPlayLink).add(inboundRecordButton);
  var inboundVoiceForwardPanel = app.createHorizontalPanel().setStyleAttribute('margin', '5px').setWidth("425px");
  var inboundVoiceForwardCheckBox = app.createCheckBox('forward all inbound calls to this number').setName('vmAutoForwardOption').setStyleAttribute('margin', '5px');
  if (properties.vmAutoForwardOption=="true") {
    inboundVoiceForwardCheckBox.setValue(true);
  }
  var inboundVoiceForwardNumber = app.createTextBox().setName('vmAutoForwardNumber').setStyleAttribute('margin', '5px');
  if (properties.vmAutoForwardNumber) {
    inboundVoiceForwardNumber.setValue(properties.vmAutoForwardNumber);
  }
  inboundVoiceForwardPanel.add(inboundVoiceForwardCheckBox).add(inboundVoiceForwardNumber);
  inboundVoicePanel.add(inboundVoiceLabel).add(inboundVoiceCheckBox).add(inboundVoiceBody).add(inboundVoiceMessagePanel).add(inboundVoiceForwardPanel);
  inboundPanel.add(inboundSMSPanel).add(inboundVoicePanel);
  tabPanel.add(smsGrid, "SMS Merge").add(vmGrid, "Voice Message Merge").add(inboundPanel,"Inbound Traffic Handling").selectTab(tabIndex);
  mainPanel.add(tabPanel);
  var rightSMSPanel = app.createVerticalPanel(); 
  var rightVMPanel = app.createVerticalPanel(); 
  
  var variablesPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("100%");
  var variablesScrollPanel = app.createScrollPanel().setHeight(130);
  var variablesLabel = app.createLabel().setText("Merge tags: ").setStyleAttribute('fontWeight', 'bold');
  var tags = formMule_getAvailableTags();
  var flexTable = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<tags.length; i++) {
    var tag = app.createLabel().setText(tags[i]);
    flexTable.setWidget(i, 0, tag);
  }
  variablesPanel.add(variablesLabel);
  variablesScrollPanel.add(flexTable);
  variablesPanel.add(variablesScrollPanel);
  
  var variablesPanel2 = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("100%");
  var variablesScrollPanel2 = app.createScrollPanel().setHeight(130);
  var variablesLabel2 = app.createLabel().setText("Merge tags: ").setStyleAttribute('fontWeight', 'bold');
  var tags2 = formMule_getAvailableTags();
  var flexTable2 = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<tags2.length; i++) {
    var tag2 = app.createLabel().setText(tags2[i]);
    flexTable2.setWidget(i, 0, tag2);
  }
  variablesPanel2.add(variablesLabel2);
  variablesScrollPanel2.add(flexTable2);
  variablesPanel2.add(variablesScrollPanel2);


  
  var SMSLangPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("184px").setStyleAttribute('marginTop', '2px');
  var SMSLangScrollPanel = app.createScrollPanel().setHeight(130);
  var SMSLangLabel = app.createLabel().setText("SMS Language Codes").setStyleAttribute('fontWeight', 'bold');
  var SMSLangs = googleLangList;
  var SMSLangFlexTable = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<SMSLangs.length; i++) {
    var tag = app.createLabel().setText(SMSLangs[i]);
    SMSLangFlexTable.setWidget(i, 0, tag);
    }
  SMSLangPanel.add(SMSLangLabel);
  SMSLangScrollPanel.add(SMSLangFlexTable);
  SMSLangPanel.add(SMSLangScrollPanel);
  
  var VMLangPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("184px").setStyleAttribute('marginTop', '2px');
  var VMLangScrollPanel = app.createScrollPanel().setHeight(130);
  var VMLangLabel = app.createLabel().setText("Voice Language Codes").setStyleAttribute('fontWeight', 'bold');
  var VMLangs = langList;
  var VMLangFlexTable = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<VMLangs.length; i++) {
    var tag = app.createLabel().setText(VMLangs[i]);
    VMLangFlexTable.setWidget(i, 0, tag);
    }
  VMLangPanel.add(VMLangLabel);
  VMLangScrollPanel.add(VMLangFlexTable);
  VMLangPanel.add(VMLangScrollPanel);
  

  var rightSMSBottomPanel = app.createVerticalPanel();
  var rightVMBottomPanel = app.createVerticalPanel();
  rightVMBottomPanel.add(VMLangPanel);
  rightSMSBottomPanel.add(SMSLangPanel);
  rightSMSPanel.add(variablesPanel2);
  rightSMSPanel.add(rightSMSBottomPanel);
  
  rightVMPanel.add(variablesPanel);
  rightVMPanel.add(rightVMBottomPanel);
  smsGrid.setWidget(0, 1, rightSMSPanel);
  vmGrid.setWidget(0, 1, rightVMPanel);
  var saveHandler = app.createServerHandler('saveSmsAndVoiceSettings').addCallbackElement(mainPanel);
  var saveSpinnerHandler = app.createClientHandler().forTargets(mainPanel).setStyleAttribute('opacity', '0.5').forTargets(refreshPanel).setVisible(true);  
  var button = app.createButton("Save settings").addClickHandler(saveHandler).addClickHandler(saveSpinnerHandler);
  mainPanel.add(button);
  app.add(mainPanel);
  app.add(refreshPanel);
  ss.show(app);
  return app;
}

function refreshInboundPlayLink(e) {
  var app = UiApp.getActiveApplication();
  var link = app.getElementById('inboundPlayLink');
  var newUrl = e.parameter.vmAutoReplyFile;
 if (newUrl != '') {
    link.setText("Play")
    link.setHref(newUrl);
    link.setStyleAttribute('color','blue');
  } else {
    link.setText('');
  }
  return app;
}

function refreshVmSoundFilePlayLink(e) {
  var app = UiApp.getActiveApplication();
  var i = e.parameter.vmIndex;
  var link = app.getElementById('vmSoundFileLink-'+i);
  var newUrl = e.parameter['vmSoundFile-'+i];
  if (newUrl != '') {
    link.setText("Play")
    link.setHref(newUrl);
    link.setStyleAttribute('color','blue');
  } else {
    link.setText('');
  }
  return app;
}


function refreshSmsTemplateGrid(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var properties = ScriptProperties.getProperties();
  var smsNumSelected = e.parameter.smsNumSelected;
  properties.smsNumSelected = smsNumSelected;
  
  var smsTemplateGrid = app.getElementById('smsTemplateGrid');
  smsTemplateGrid.resize(smsNumSelected*4, 5);
  var smsPropertyString = properties.smsPropertyString;
  var smsPropertyObject = new Object();
  for (var i=0; i<properties.smsNumSelected; i++) {
    smsPropertyObject['smsName-'+i] = e.parameter['smsName-'+i];
    smsPropertyObject['smsPhone-'+i] = e.parameter['smsPhone-'+i];
    smsPropertyObject['smsBody-'+i] = e.parameter['smsBody-'+i];
    smsPropertyObject['smsLang-'+i] = e.parameter['smsLang-'+i];
    smsPropertyObject['smsCol-'+i] = e.parameter['smsCol-'+i];
    smsPropertyObject['smsVal-'+i] = e.parameter['smsVal-'+i];
  } 
  for (var i = 0; i<smsNumSelected; i++) {
    smsTemplateGrid.setWidget(0+(i*4), 0, app.createLabel('Template Name'));
    smsTemplateGrid.setWidget(1+(i*4), 0, app.createLabel('Phone #'));
    smsTemplateGrid.setWidget(2+(i*4), 0, app.createLabel('Body'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 0, 'backgroundColor','grey'); 
    var smsTemplateNameBox = app.createTextBox().setName('smsName-'+i).setWidth("140px");
    if (smsPropertyObject['smsName-'+i]) {
      smsTemplateNameBox.setValue(smsPropertyObject['smsName-'+i]);
    }
    var smsPhoneBox = app.createTextBox().setName('smsPhone-'+i).setWidth("140px");
    if (smsPropertyObject['smsPhone-'+i]) {
      smsPhoneBox.setValue(smsPropertyObject['smsPhone-'+i]);
    }
    var smsBodyArea = app.createTextArea().setName('smsBody-'+i).setWidth("140px").setHeight("110px");
    if (smsPropertyObject['smsBody-'+i]) {
      smsBodyArea.setValue(smsPropertyObject['smsBody-'+i]);
    }
    smsTemplateGrid.setWidget(0+(i*4), 1, smsTemplateNameBox);
    smsTemplateGrid.setWidget(1+(i*4), 1, smsPhoneBox);
    smsTemplateGrid.setWidget(2+(i*4), 1, smsBodyArea);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 1, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(0+(i*4), 2, app.createLabel('send if'));
    var smsCondCol = app.createListBox().setName('smsCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      smsCondCol.addItem(headers[j]);
    }
    if (smsPropertyObject['smsCol-'+i]) {
      var index = headers.indexOf(smsPropertyObject['smsCol-'+i]);
      smsCondCol.setSelectedIndex(index);
    }
    smsTemplateGrid.setWidget(1+(i*4), 2, smsCondCol);
    smsTemplateGrid.setWidget(2+(i*4), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 2, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(1+(i*4), 3, app.createLabel("="));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 3, 'backgroundColor','grey'); 
    var smsCondVal = app.createTextBox().setName('smsVal-'+i).setWidth("100px");
    if (smsPropertyObject['smsVal-'+i]) {
      smsCondVal.setValue(smsPropertyObject['smsVal-'+i]);
    }
    var smsLang = app.createTextBox().setName('smsLang-'+i).setWidth("50px");
    if (smsPropertyObject['smsLang-'+i]) {
      smsLang.setValue(smsPropertyObject['smsLang-'+i]);
    } else {
      smsLang.setValue('en');
    }
    smsTemplateGrid.setWidget(1+(i*4), 4, smsCondVal);
    smsTemplateGrid.setWidget(2+(i*4), 4, smsLang);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 4, 'backgroundColor','grey'); 
  }
  return app;
}


function refreshVmTemplateGrid(e) {
  var app = UiApp.getActiveApplication();
  var mainPanel = app.getElementById("mainPanel");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var recordingSheet = ss.getSheetByName("MyVoiceRecordings");
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var recordingSelectTexts = [];
  var recordingSelectValues = [];
  if (recordingSheet.getLastRow()-1>0) {
    recordingSelectTexts = recordingSheet.getRange(2, 1, recordingSheet.getLastRow()-1, 1).getValues();
    recordingSelectValues = recordingSheet.getRange(2, 2, recordingSheet.getLastRow()-1, 1).getValues();
  } 
  var properties = ScriptProperties.getProperties();
  var vmNumSelected = e.parameter.vmNumSelected;
  properties.vmNumSelected = vmNumSelected;
  var vmTemplateGrid = app.getElementById('vmTemplateGrid');
  vmTemplateGrid.resize(vmNumSelected*6, 5);
  var vmPropertyString = properties.vmPropertyString;
  var vmPropertyObject = new Object();
  for (var i=0; i<properties.vmNumSelected; i++) {
    vmPropertyObject['vmName-'+i] = e.parameter['vmName-'+i];
    vmPropertyObject['vmPhone-'+i] = e.parameter['vmPhone-'+i];
    vmPropertyObject['vmBody-'+i] = e.parameter['vmBody-'+i];
    vmPropertyObject['vmSoundFile-'+i] = e.parameter['vmSoundFile-'+i];
    vmPropertyObject['vmRecordOption-'+i] = e.parameter['vmRecordOption-'+i];
    vmPropertyObject['vmLang-'+i] = e.parameter['vmLang-'+i];
    vmPropertyObject['vmCol-'+i] = e.parameter['vmCol-'+i];
    vmPropertyObject['vmVal-'+i] = e.parameter['vmVal-'+i];
  }
  for (var i = 0; i<vmNumSelected; i++) {
    vmTemplateGrid.setWidget(0+(i*6), 0, app.createLabel('Template Name'));
    vmTemplateGrid.setWidget(1+(i*6), 0, app.createLabel('Phone #'));
    vmTemplateGrid.setWidget(2+(i*6), 0, app.createLabel('Read Message'));
    vmTemplateGrid.setWidget(3+(i*6), 0, app.createLabel('Play recording'))
    vmTemplateGrid.setWidget(4+(i*6), 0, app.createLabel('Record response?'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 0, 'backgroundColor','grey'); 
    var vmTemplateNameBox = app.createTextBox().setName('vmName-'+i).setWidth("140px");
    if (vmPropertyObject['vmName-'+i]) {
      vmTemplateNameBox.setValue(vmPropertyObject['vmName-'+i]);
    }
    var vmPhoneBox = app.createTextBox().setName('vmPhone-'+i).setWidth("140px");
    if (vmPropertyObject['vmPhone-'+i]) {
      vmPhoneBox.setValue(vmPropertyObject['vmPhone-'+i]);
    }
    var vmBodyArea = app.createTextArea().setName('vmBody-'+i).setWidth("140px").setHeight("90px");
    if (vmPropertyObject['vmBody-'+i]) {
      vmBodyArea.setValue(vmPropertyObject['vmBody-'+i]);
    }
    var vmSoundFileListBox = app.createListBox().setName('vmSoundFile-'+i).setWidth("140px");
    var soundFileArray = [];
    vmSoundFileListBox.addItem("None", "");
    for (var k=0; k<recordingSelectTexts.length; k++) {
      vmSoundFileListBox.addItem(recordingSelectTexts[k][0], recordingSelectValues[k][0]);
      soundFileArray.push(recordingSelectValues[k][0]);
    }
    var m = 0;
    var n = -1;
    if (vmPropertyObject['vmSoundFile-'+i]) {
      var index = soundFileArray.indexOf(vmPropertyObject['vmSoundFile-'+i]);
      if (index!=-1) {
        vmSoundFileListBox.setSelectedIndex(index + 1);
        m = index;
        n = 0;
      }
    } 
    var vmSoundFileLinkPanel = app.createHorizontalPanel();
    var vmSoundFilePlayLinkRefreshHandler = app.createServerHandler('refreshVmSoundFilePlayLink').addCallbackElement(vmSoundFileListBox).addCallbackElement(vmSoundFileLinkPanel);  
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkRefreshHandler);
    var linkText = "";
    if ((n!=-1)&&(vmPropertyObject['vmSoundFileLink-'+i]!='')&&(recordingSelectValues.length>0)) {
      linkText = "Play";
    }
    var vmSoundFilePlayLink = app.createAnchor(linkText,recordingSelectValues[m][0]).setId('vmSoundFileLink-'+i);
    var vmSoundFilePlayLinkClientHandler = app.createClientHandler().forTargets(vmSoundFilePlayLink).setStyleAttribute('color','grey');
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkClientHandler);
    var vmSoundFileIndex = app.createTextBox().setValue(i).setName('vmIndex').setVisible(false);
    vmSoundFileLinkPanel.add(vmSoundFilePlayLink).add(vmSoundFileIndex);  
    var vmRecordOption = app.createListBox().setName('vmRecordOption-'+i).addItem("No").addItem("Yes");
    if (vmPropertyObject['vmRecordOption-'+i]) {
      if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
        vmRecordOption.setSelectedIndex(1);
      }
    }
    var vmRecordHandler = app.createServerHandler('recordVm').addCallbackElement(mainPanel);
    var vmRecordButton = app.createButton('Call me to record voice message').addClickHandler(vmRecordHandler);
    vmTemplateGrid.setWidget(0+(i*6), 1, vmTemplateNameBox);
    vmTemplateGrid.setWidget(1+(i*6), 1, vmPhoneBox);
    vmTemplateGrid.setWidget(2+(i*6), 1, vmBodyArea);
    vmTemplateGrid.setWidget(3+(i*6), 1, vmSoundFileListBox);
    vmTemplateGrid.setWidget(3+(i*6), 2, vmSoundFileLinkPanel);
    vmTemplateGrid.setWidget(3+(i*6), 4, vmRecordButton);
    vmTemplateGrid.setWidget(4+(i*6), 1, vmRecordOption);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 1, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(0+(i*6), 2, app.createLabel('call if'));
    var vmCondCol = app.createListBox().setName('vmCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      vmCondCol.addItem(headers[j]);
    }
    if (properties['vmCol-'+i]) {
      var index = headers.indexOf(properties['vmCol-'+i]);
      if (index!=-1) {
        vmCondCol.setSelectedIndex(index);
      }
    }
    vmTemplateGrid.setWidget(1+(i*6), 2, vmCondCol);
    vmTemplateGrid.setWidget(2+(i*6), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 2, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(1+(i*6), 3, app.createLabel("="));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 3, 'backgroundColor','grey'); 
    var vmCondVal = app.createTextBox().setName('vmVal-'+i).setWidth("100px");
    if (properties['vmVal-'+i]) {
      vmCondVal.setValue(properties['vmVal-'+i]);
    }
    var vmLang = app.createTextBox().setName('vmLang-'+i).setWidth("50px");
    if (properties['vmLang-'+i]) {
        vmLang.setValue(vmLang);
    } else {
        vmLang.setValue("en");
    }
    vmTemplateGrid.setWidget(1+(i*6), 4, vmCondVal);
    vmTemplateGrid.setWidget(2+(i*6), 4, vmLang);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 4, 'backgroundColor','grey'); 
  }
  return app;
}


function inboundRecordVm(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  e.parameter.num = "2";
  inboundRecordApp(e)
  return app;
}

function recordVm(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  e.parameter.num = "1";
  inboundRecordApp(e)
  return app;
  
}

function inboundRecordApp(e) {
  saveSmsAndVoiceSettings(e);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MyVoiceRecordings');
  ss.setActiveSheet(sheet);
  var app = UiApp.createApplication().setTitle("Record a new message").setHeight(200);
  var panel = app.createVerticalPanel();
  var phoneNumberLabel = app.createLabel("Please enter the phone number you would like Twilio to call to record your voice message");
  var phoneNumberBox = app.createTextBox().setName('phone').setWidth('150px');
  var lastPhone = ScriptProperties.getProperty('lastPhone');
  if (lastPhone) {
    phoneNumberBox.setValue(lastPhone);
  }
  var recordingName = app.createLabel("Please give a unique name to this recording");
  var recordingNameBox = app.createTextBox().setName('name').setWidth('150px');
  var hiddenNumBox = app.createTextBox().setName('num').setVisible(false).setValue(e.parameter.num);
  var panel2 = app.createVerticalPanel().setId('callingPanel').setVisible(false);
  var callingImage = app.createImage('http://status.twilio.com/images/logo.png').setWidth("100px");
  var callingLabel = app.createLabel('Twilio is attempting to call the number you provided.  This can sometimes take 30-45 seconds. Click \'Done\' once you have hung up and can see your recording show up in the \'MyVoiceRecordings\' sheet');
  var doneButtonHandler = app.createServerHandler('doneRecording').addCallbackElement(panel);
  var doneButton = app.createButton('Done').addClickHandler(doneButtonHandler);
  panel2.add(callingImage).add(callingLabel).add(doneButton);
  panel.add(phoneNumberLabel).add(phoneNumberBox).add(recordingName).add(recordingNameBox).add(hiddenNumBox);
  var buttonClientHandler = app.createClientHandler().forTargets(panel2).setVisible(true).forTargets(panel).setVisible(false);
  var buttonHandler = app.createServerHandler('saveRecording').addCallbackElement(panel);
  var button = app.createButton('Call me to record my message').addClickHandler(buttonHandler).addClickHandler(buttonClientHandler);
  panel.add(button);
  app.add(panel);
  app.add(panel2);
  ss.show(app);
  return app;
}


function saveRecording(e) {
  var app = UiApp.getActiveApplication();
  var num = e.parameter.num;
  var phoneNumber = e.parameter.phone;
  ScriptProperties.setProperty('lastPhone', phoneNumber);
  var recordingName = e.parameter.name;
  var accountSID = ScriptProperties.getProperty("accountSID");
  var authToken = ScriptProperties.getProperty("authToken");
  var twilioNumber = ScriptProperties.getProperty("twilioNumber");
  var args = new Object();
  args.RequestResponse = "TRUE";
  args.OutboundMessageRecording = "TRUE";
  args.RecordingName = recordingName;
  var lang = ScriptProperties.getProperty('defaultVoiceLang');
  formMule_makeRoboCall(phoneNumber, "Please record your message after the tone.", accountSID, authToken, '', args, lang);
  return app;
}

function doneRecording(e) {
  var app = UiApp.getActiveApplication();
  var num = parseInt(e.parameter.num);
  app.close();
  formMule_smsAndVoiceSettings(num);
  return app;
}


function saveSmsAndVoiceSettings(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var sheet = ss.getSheetByName(properties.sheetName);
  var headers = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues()[0];
  properties.smsEnabled = e.parameter.smsEnabled;
  properties.smsTrigger = e.parameter.smsTrigger;
  properties.smsMaxLength = e.parameter.smsMaxLength;
  properties.smsNumSelected = e.parameter.smsNumSelected;
  var smsPropertyObject = new Object();
  for (var i=0; i<properties.smsNumSelected; i++) {
    smsPropertyObject['smsName-'+i] = e.parameter['smsName-'+i];
    if ((headers.indexOf(smsPropertyObject['smsName-'+i] + " SMS Status") == -1)&&(properties.smsEnabled=="true")) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(smsPropertyObject['smsName-'+i] + " SMS Status").setBackgroundColor('purple').setFontColor('white').setComment('Don\'t change the name of this column. It is used to log the send status of an SMS Message, whose name must be a match.');
      sheet.setColumnWidth(sheet.getLastColumn(), "400");
    }
    smsPropertyObject['smsPhone-'+i] = e.parameter['smsPhone-'+i];
    smsPropertyObject['smsBody-'+i] = e.parameter['smsBody-'+i];
    smsPropertyObject['smsLang-'+i] = e.parameter['smsLang-'+i];
    smsPropertyObject['smsCol-'+i] = e.parameter['smsCol-'+i];
    smsPropertyObject['smsVal-'+i] = e.parameter['smsVal-'+i];
  }
  var smsPropertyString = Utilities.jsonStringify(smsPropertyObject);
  properties.smsPropertyString = smsPropertyString;
  properties.smsAutoReplyOption = e.parameter.smsAutoReplyOption;
  properties.smsAutoReplyBody = e.parameter.smsAutoReplyBody;
  properties.smsAutoForwardOption = e.parameter.smsAutoForwardOption;
  properties.smsAutoForwardEmail = e.parameter.smsAutoForwardEmail;
  properties.vmAutoReplyReadOption = e.parameter.vmAutoReplyReadOption;
  properties.vmAutoReplyBody = e.parameter.vmAutoReplyBody;
  properties.vmAutoReplyPlayOption = e.parameter.vmAutoReplyPlayOption;
  properties.vmAutoReplyFile = e.parameter.vmAutoReplyFile;
  properties.vmAutoForwardOption = e.parameter.vmAutoForwardOption;
  properties.vmAutoForwardNumber = e.parameter.vmAutoForwardNumber;
  
  properties.vmEnabled = e.parameter.vmEnabled;
  properties.vmTrigger = e.parameter.vmTrigger;
  properties.vmNumSelected = e.parameter.vmNumSelected;
 
  
  var vmPropertyObject = new Object();
  for (var i=0; i<properties.vmNumSelected; i++) {
    vmPropertyObject['vmName-'+i] = e.parameter['vmName-'+i];
    if ((headers.indexOf(vmPropertyObject['vmName-'+i] + " VM Status") == -1)&&(properties.vmEnabled=="true")) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(vmPropertyObject['vmName-'+i] + " VM Status").setBackgroundColor('orange').setComment('Don\'t change the name of this column. It is used to log the status of an outbound Voice Message, whose name must be a match.');
      sheet.setColumnWidth(sheet.getLastColumn(), "400");
    }
    vmPropertyObject['vmPhone-'+i] = e.parameter['vmPhone-'+i];
    vmPropertyObject['vmBody-'+i] = e.parameter['vmBody-'+i];
    vmPropertyObject['vmSoundFile-'+i] = e.parameter['vmSoundFile-'+i];
    vmPropertyObject['vmRecordOption-'+i] = e.parameter['vmRecordOption-'+i];
    vmPropertyObject['vmLang-'+i] = e.parameter['vmLang-'+i];
    vmPropertyObject['vmCol-'+i] = e.parameter['vmCol-'+i];
    vmPropertyObject['vmVal-'+i] = e.parameter['vmVal-'+i];
  }
  var vmPropertyString = Utilities.jsonStringify(vmPropertyObject);
  properties.vmPropertyString = vmPropertyString;
  ScriptProperties.setProperties(properties);
  app.close();
  return app;
}

function generatearray(range) {
  var output = '[';
  for (var i=0; i<range.length; i++) {
    output += "'" + range[i][0] + "',";
  }
  output += ']';
  return output;
}


function doGet(e) {
  var app = UiApp.createApplication();
  var properties = ScriptProperties.getProperties();
  var accountSID = properties.accountSID;
  var authToken = properties.authToken;
  var twilioNumber = properties.twilioNumber;
  var messageId = e.parameter.MessageId;
  var message = CacheService.getPublicCache().get(messageId);
  var language = e.parameter.Language;
  if (!language) {
    language = properties.defaultVoiceLang;
  }
  var ssId = properties.ssId;
  var ss = SpreadsheetApp.openById(ssId);
  if (e.parameter.Verify == "TRUE") {
    var output = ContentService.createTextOutput();
    var content ='<verified>TRUE</verified>';
     output.setContent(content);
     output.setMimeType(ContentService.MimeType.XML);
     return output;
  }
  if (e.parameter.Body == "Woot woot") {
    var args = new Object();
    args.runTest = "TRUE";
    formMule_sendAText(e.parameter.From, "Congratulations! Your script is able to recieve texts via Twilio. Woot woot to you!", accountSID, authToken, '', args, language);
    formMule_makeRoboCall(e.parameter.From, "Congratulations! Your script is able to create voice messages via Twilio. Hoot Hoot to you!", accountSID, authToken, '', args, language);
    onOpen();
    formMule_smsAndVoiceSettings();
    return;
  }
  
  if (e.parameter.Voice == "TRUE") {
    var output = ContentService.createTextOutput();
    var responseArgs = '';
    if ((message)&&(message!='')) {
      responseArgs += '<Say voice="woman" language="'+language+'">'+message+'</Say>';
    }
    if (e.parameter.PlayFile) {
      responseArgs += '<Play>' + e.parameter.PlayFile + '</Play>';
    }
    if (e.parameter.RequestResponse == "TRUE") {
      responseArgs += '<Record />';
    }
    var content ='<Response>'+responseArgs+'</Response>';
    output.setContent(content);
    output.setMimeType(ContentService.MimeType.XML);
    var flagSheetName = e.parameter.flagSheetName;
    var flagHeader = e.parameter.flagHeader;
    var flagRow = e.parameter.flagRow;
    if ((flagSheetName)&&(flagHeader)&&(flagRow)) {
      var sheet = ss.getSheetByName(flagSheetName);
      var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      var statusCol = headers.indexOf(unescape(flagHeader))+1
      var range = sheet.getRange(flagRow, statusCol);
      var sid = e.parameter.CallSid;
      var statusObj = getCallStatus(sid);
      updateCallStatus(range, statusObj);
    }
    return output;
  }
  
  if ((e.parameter.VoiceCallBack == "TRUE")&&(e.parameter.OutboundMessageRecording == "TRUE")) {
    var sheet = ss.getSheetByName("MyVoiceRecordings");
    var lastRow = sheet.getLastRow();
    var callSid = e.parameter.CallSid;
    var recordingUrl = getRecordingUrl(callSid);
    var range = sheet.getRange(lastRow+1, 1, 1, 2).setValues([[e.parameter.RecordingName,recordingUrl]]);
  }
  
   if (e.parameter.VoiceCallBack == "TRUE") {
    var flagSheetName = e.parameter.flagSheetName;
    var flagHeader = e.parameter.flagHeader;
    var flagRow = e.parameter.flagRow;
    if ((flagSheetName)&&(flagHeader)&&(flagRow)) {
      var sheet = ss.getSheetByName(flagSheetName);
      var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      var statusCol = headers.indexOf(unescape(flagHeader))+1
      var range = sheet.getRange(flagRow, statusCol);
      var sid = e.parameter.CallSid;
      var statusObj = getCallStatus(sid);
      var recordingUrl = getRecordingUrl(sid);
      updateCallStatus(range, statusObj, recordingUrl);
    }
     return;
  }
  
  if (e.parameter.CallSid) {
    var output = ContentService.createTextOutput();
    var content = '';
    var callString = '';
    if (properties.vmAutoReplyReadOption=="true") {
      callString += '<Say voice="woman">' + properties.vmAutoReplyBody + '</Say>';
    }
    if (properties.vmAutoReplyPlayOption=="true") {
      callString += '<Play>' + properties.vmAutoReplyFile + '</Play>';
    }
    if (properties.vmAutoForwardOption=="true") {
      callString += '<Dial>' + properties.vmAutoForwardNumber + '</Dial>'
    }
    var content ='<Response>' + callString + '</Response>';
    output.setContent(content);
    output.setMimeType(ContentService.MimeType.XML);
    return output;
  }
  
   if (e.parameter.SmsMessageSid) {
     var output = ContentService.createTextOutput();
     var content = '';
     if (properties.smsAutoReplyOption=="true") {
       content = '<Response><Sms>'+ properties.smsAutoReplyBody + '</Sms></Response>';
     }
     if (properties.smsAutoForwardOption=="true") {
       var email = properties.smsAutoForwardEmail;
       var subj = 'SMS Received from ' + e.parameter.From;
       var body = e.parameter.Body
       MailApp.sendEmail(email, subj, body);
     }
    output.setContent(content);
    output.setMimeType(ContentService.MimeType.XML);
    return output; 
  }
  return app;
}

function updateCallStatus(range, obj, recordingUrl) {
  var app = UiApp.getActiveApplication();
  var ssId = ScriptProperties.getProperty("ssId");
  var ss = SpreadsheetApp.openById(ssId);
  var messages = range.getValue();
  var status = '';
  var sid = obj.Sid.getText();
  status = obj.Status.getText();
  messages = messages.split(";\n");
  var newValues = [];
  var values = [];
  for (var i=0; i<messages.length; i++) {
    values = [];
    values = messages[i].split(", ");
    if (values[1]==("SID="+sid)) {
      values[5] = "Status=" + status;
      if (status=="completed") {
        values[6] = "Duration=" + obj.Duration.getText() + "sec";
        values[7] = "AnsweredBy=" + obj.AnsweredBy.getText();
      }
      if (recordingUrl) {
        values[8] = "RecordingUrl=" + recordingUrl;
      }
      newValues[i] = values.join(", ");
    } else {
      newValues[i] = messages[i];
    }
  }
  newValues = newValues.join(";\n");
  range.setValue(newValues);
  return app;
}


function formMule_splitTexts(longMessage, maxSplits) {
  var charLimit = 160;
  var indicator;
  var breakPoint;
  var messages = [];
  var n;
  var m;
  var numTexts = longMessage.length / charLimit;
  numTexts = Math.ceil(numTexts);
  if (numTexts>maxSplits) {
    numTexts = maxSplits;
  }
  for (n = 0, m = 0; n < numTexts && n < maxSplits; n++) {
    m = n * charLimit;
    // set the indicator so we can now how long it is
    
    if (numTexts>1) {
      indicator = '(' + (n + 1) + '/' + numTexts + ')';
    } else {
      indicator = '';
    }
    // set the breakpoint, taking indicator length into consideration
    breakPoint = m + charLimit - indicator.length;
    // insert the indicator into the correct spot
    longMessage = longMessage.substring(0, breakPoint) + indicator + longMessage.substring(breakPoint);
    // add a message (will be charLimit long and include the indicator)
    messages.push(longMessage.substring(m, m + charLimit));
  }
  return messages;
}


//flagObject contains three parameters: sheetName, header, and row to specify the cell where callback information will be populated
function formMule_sendAText(phoneNumber, message, accountSID, authToken, flagObject, args, lang, maxTexts){
  var ssId = ScriptProperties.getProperty("ssId");
  var ss = SpreadsheetApp.openById(ssId);
  var timeZone = ss.getSpreadsheetTimeZone();
  var triggerFlag = false;
  var flagString = '';
  if (flagObject) {
    var flagSheetName = flagObject.sheetName;
    var flagHeader = flagObject.header;
    var flagRow = flagObject.row;
    flagString = "&flagSheetName="+flagSheetName+"&flagHeader="+flagHeader+"&flagRow="+flagRow;
    var sheet = ss.getSheetByName(flagSheetName);
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf(unescape(flagHeader))+1
    var range = sheet.getRange(flagRow, statusCol);
  } else {
    flagString = '';
  }
  accountSID = accountSID||ScriptProperties.getProperty("accountSID");
  authToken = authToken||ScriptProperties.getProperty("authToken");
  var twilioNumber = ScriptProperties.getProperty("twilioNumber");
  if (accountSID == null || accountSID == undefined){
    ScriptProperties.setProperty("accountSID", Browser.inputBox("Please Enter your Twilio account SID"))
    accountSid = ScriptProperties.getProperty("accountSID");
  }
  if (authToken == null || authToken == undefined){
    ScriptProperties.setProperty("authToken", Browser.inputBox("Please Enter your Twilio authorization token"))
    authToken = ScriptProperties.getProperty("authToken");
  }
  if (lang) {
    message = LanguageApp.translate(message, "", lang);
  }
  if (!maxTexts) {
    maxTexts = 3;
  }
  message = formMule_splitTexts(message, maxTexts);
  var argGetString = '';
  if (!args) {
    var args = new Object();
  }
  var url = ScriptApp.getService().getUrl();
  phoneNumber = phoneNumber.replace(/\s+/g, ' ');
  var phoneNumberArray = phoneNumber.split(",");
  var status = '';
  for (var j=0; j<phoneNumberArray.length; j++) {
    for (var i=0; i<message.length; i++) {
      args.messageNum = encodeURIComponent((i+1) + " of " + (message.length));
      if (args) {
        for (var key in args) {
          argGetString += "&" + key + "=" + args[key];
        }
      }
      if (message[i]!='') {
      var formatPOST = {
       "From" : twilioNumber,
       "To" : phoneNumberArray[j],
       "Body" : message[i],
       };
      var callObj = {
      method: "POST",
      payload: formatPOST,
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)},
      contentType: "application/x-www-form-urlencoded", 
      };
      try {
        var text = UrlFetchApp.fetch(https + webName + accountSID + "/SMS/Messages",callObj);
        var response = text.getContentText();
        response = Xml.parse(response);
        var thisStatus = response.TwilioResponse.SMSMessage.Status.getText();
        if (thisStatus=="queued") {
          triggerFlag = true;
        }
        var created = response.TwilioResponse.SMSMessage.DateCreated.getText();
        created = new Date(created);
        created = Utilities.formatDate(created, timeZone, "MM/dd/yy HH:mm:ss");
        var sid = response.TwilioResponse.SMSMessage.Sid.getText();
        status += "SMSMessage" + (i+1) + "/" + message.length + ", SID=" + sid + ", To=" + phoneNumberArray[j] + ", Created=" + created + ", \"" + message[i] + "\", Status=" + thisStatus;
        if ((i+1)<message.length) {
          status += ";\n";
        }
      }
       catch(e) {
        Logger.log(e);
      }
    }
  }
  var currValue = '';
  if ((j+1)<phoneNumberArray.length) {
    status += "|\n";
  }
}
  if (range) {
    range.setValue(currValue + status);
    if (triggerFlag == true) {
      setUpdateTrigger();
    }
  }
  return;
}


function setUpdateTrigger() {
  var triggers = ScriptApp.getScriptTriggers();
  for (var i=0; i< triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=='triggerUpdateStatus') {
      return;
    }
  }
  var now = new Date();
  var newDateObj = new Date(now.getTime() + (0.5)*60000);
  var oneTimeRerunTrigger = ScriptApp.newTrigger('triggerUpdateStatus').timeBased().at(newDateObj).create();
}

function triggerUpdateStatus() {
  var properties = ScriptProperties.getProperties();
  var sheetName = properties.sheetName;
  var smsPropertyString = properties.smsPropertyString;
  if ((smsPropertyString)&&(smsPropertyString!='')) {
    var smsProperties = Utilities.jsonParse(smsPropertyString)
        for (var i=0; i<properties.smsNumSelected; i++) {
          updateStatus(sheetName, smsProperties['smsName-'+i] + " SMS Status");
        }
  }
  var triggers = ScriptApp.getScriptTriggers();
  for (var i=0; i< triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=='triggerUpdateStatus') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


function updateStatus(sheetName, colName) {
  var triggerFlag = false;
  var ssId = ScriptProperties.getProperty('ssId')
  var ss = SpreadsheetApp.openById(ssId);
  var timeZone = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(colName)+1;
  var range = sheet.getRange(2, col, sheet.getLastRow()-1, 1);
  var values = range.getValues();
  for (var i=0; i<values.length; i++) {
    var value = values[i][0];
    if (value!='') {
      var messages = value.split("|");
      for (var k=0; k<messages.length; k++) {
        var subvalues = messages[k].split(";\n");
        for (var j=0; j<subvalues.length; j++) {
          if (subvalues[j]!='') {
            subvalues[j] = subvalues[j].match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g);
            var subsubvalues = subvalues[j] || [];
            var sid = subsubvalues[1].split("=")[1];
            if (subsubvalues[5]=="Status=queued") {
              var statusObj = getSMSStatus(sid);
              var thisStatus = statusObj.status;
              if (thisStatus=="queued") {
                triggerFlag = true;
              }
              subsubvalues[5] = "Status=" + thisStatus;
              if (thisStatus=="sent") {
                var sentTime = statusObj.sent;
                if (sentTime) {
                  sentTime = new Date(sentTime);
                  sentTime = Utilities.formatDate(sentTime, timeZone, "MM/dd/yy HH:mm:ss");
                }
                subsubvalues[6] = "TimeSent=" + sentTime 
              }
            }
          }
          subvalues[j] = subsubvalues.join(", ");
          }
          messages[k] = subvalues.join(";\n");
        }
       values[i][0] = messages.join("|\n"); 
       range.setValues(values)
      }
    }
    if (triggerFlag==true) {
      setUpdateTrigger();
    }
}

function getCallStatus(sid) {
  var accountSID = ScriptProperties.getProperty("accountSID");
  var authToken = ScriptProperties.getProperty("authToken");
  try {
    var fetchUrl = "https://api.twilio.com/2010-04-01/Accounts/"+accountSID+"/Calls/"+sid+".xml";
    var callObj = {
      method: "GET",
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)}    
    }
        debugger;
    var text = UrlFetchApp.fetch(fetchUrl, callObj);
    var response = text.getContentText();
    response = Xml.parse(response);
    var statusObj = response.TwilioResponse.Call;
    return statusObj;
    }
     catch(e) {
    Logger.log(e);
    }
}

function getRecordingUrl(callSid) {
    var accountSID = ScriptProperties.getProperty("accountSID");
    var authToken = ScriptProperties.getProperty("authToken");
  try {
    var fetchUrl = this.https + this.webName + accountSID + '/Calls/' + callSid + '/Recordings.xml';
    var callObj = {
      method: "GET",
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)}    
    }
    var text = UrlFetchApp.fetch(fetchUrl, callObj);
    var response = text.getContentText();
    response = Xml.parse(response);
    var recordingSid = response.TwilioResponse.Recordings.Recording[0].getElement().getText();
    var recordingUrl = this.https + this.webName + accountSID + '/Recordings/' + recordingSid;
    return recordingUrl;
  } catch(err) {
     Logger.log(err);
  }
}



function getSMSStatus(sid) {
  var accountSID = ScriptProperties.getProperty("accountSID");
  var authToken = ScriptProperties.getProperty("authToken");
  try {
    var fetchUrl = "https://api.twilio.com/2010-04-01/Accounts/"+accountSID+"/SMS/Messages/"+sid+".xml";
    var callObj = {
      method: "GET",
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)}    
    }
    var text = UrlFetchApp.fetch(fetchUrl, callObj);
    var response = text.getContentText();
    response = Xml.parse(response);
    var thisStatus = response.TwilioResponse.SMSMessage.Status.getText();
    var thisSent = response.TwilioResponse.SMSMessage.DateSent.getText();
    var statusObj = new Object();
    statusObj.status = thisStatus;
    statusObj.sent = thisSent;
    return statusObj;
    }
     catch(e) {
    Logger.log(e);
    }
}

//flagObject contains three parameters: sheetName, header, and row to specify the cell where callback information will be populated
function formMule_makeRoboCall(phoneNumber, message, accountSID, authToken, flagObject, args, lang) {
  var ssId = ScriptProperties.getProperty("ssId");
  var ss = SpreadsheetApp.openById(ssId);
  var timeZone = ss.getSpreadsheetTimeZone();
  phoneNumber = phoneNumber.replace(/\s+/g, ' ');
  var phoneNumberArray = phoneNumber.split(",");
  var flagString = '';
  if (flagObject) {
    var flagSheetName = flagObject.sheetName;
    var flagHeader = flagObject.header;
    var flagRow = flagObject.row;
    flagString = "&flagSheetName="+flagSheetName+"&flagHeader="+flagHeader+"&flagRow="+flagRow;
  } else {
    flagString = '';
  }
  var triggerFlag = false;
  if (!(lang)) {
    lang = ScriptProperties.getProperty('defaultVoiceLang');
  }
  message = LanguageApp.translate(message, "", lang);
  var validLangs = ['en', 'en-gb', 'es', 'fr', 'de'];
  if (validLangs.indexOf(lang)==-1) {
    lang = 'en';
  }
  var argGetString = '';
  var record = false;
  if ((args)&&(args!='')) {
    for (var key in args) {
      argGetString += "&" + key + "=" + args[key];
    }
  if (args.RequestResponse == "TRUE") {
    record = true;
  }
  }
  var messageId = new Date();
  messageId = messageId.getTime();
  CacheService.getPublicCache().put(messageId, message, 3600);
  accountSID = accountSID||ScriptProperties.getProperty("accountSID");
  authToken = authToken||ScriptProperties.getProperty("authToken");
  var twilioNumber = ScriptProperties.getProperty("twilioNumber");
  var webAppUrl = ScriptApp.getService().getUrl() + "?Voice=TRUE&MessageId="+messageId+"&Language="+lang + flagString + argGetString;
  var callBackUrl = ScriptApp.getService().getUrl() + "?VoiceCallBack=TRUE" + flagString + argGetString;
  var status = '';
  for (var i=0; i<phoneNumberArray.length; i++) {
    var formatPOST = {
      "From" : twilioNumber,
      "To" : phoneNumberArray[i],
      "Url" : webAppUrl,
      "Method" : "GET",
      "StatusCallback" : callBackUrl,
      "StatusCallbackMethod" : "GET",
      "Record" : record,
      "IfMachine" : "Continue"
     };
    var callObj = {
      //work around to get past the inability to use ":" in the urlfetch app
      method: "POST",
      payload: formatPOST,
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)},    
      contentType: "application/x-www-form-urlencoded",  
      };    
   try{
     var response = UrlFetchApp.fetch(https + webName + accountSID + "/Calls", callObj);
     response = response.getContentText();
     response = Xml.parse(response);
     var thisStatus = response.TwilioResponse.Call.Status.getText();
     if ((thisStatus=="ringing")||(thisStatus=="in-progress")) {
       triggerFlag = true;
     }
     var created = response.TwilioResponse.Call.DateCreated.getText();
     created = new Date(created);
     created = Utilities.formatDate(created, timeZone, "MM/dd/yy HH:mm:ss");
     var sid = response.TwilioResponse.Call.Sid.getText();
     message = message.replace(/,/g, '');
     status += "VoiceMessage, SID=" + sid + ", To=" + phoneNumberArray[i] + ", StartTime=" + created + ", Body=" + message + ", Status=" + thisStatus;
     if ((i+1)<phoneNumberArray.length) {
       status += ";\n";
     }
   }
   catch(err){
     Logger.log(err);
     status += err;
   }
 }
 if (flagSheetName) {
   var sheet = ss.getSheetByName(flagSheetName);
   var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
   var statusCol = headers.indexOf(unescape(flagHeader))+1
   var range = sheet.getRange(flagRow, statusCol);
   range.setValue(status);
  }
 return;
}


//gets the associtated phone numbers with the account 
function formMule_getPhoneNumbers(accountSid,authToken) {
  accountSid = accountSid||ScriptProperties.getProperty("twilioID");
  authToken =  authToken||ScriptProperties.getProperty("securityToken");
  var phoneNumbers = [];
//create checks userproreties using 
  if (accountSid == null || accountSid == undefined){
    Browser.msgBox("You are missing your Account SID");
    formMule_howToSetUpTwilio();
  }
  if (authToken == null || authToken == undefined){
    Browser.msgBox("You are missing your Auth Token");
    formMule_howToSetUpTwilio();
  }
      
  var callObj = {
    //work around to get past the inability to use ":" in the urlfetch app
    headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSid + ":" + authToken)},    
    contentType: "application/xml; charset=utf-8",  
    };

  try{ 
  //change on country will add easy change when fully functional
  //returns an array of JSON objects 1 for each number
    var availableNumbers = UrlFetchApp.fetch(https + webName + accountSid + "/IncomingPhoneNumbers.json", callObj).getContentText();
    availableNumbers = JSON.parse(availableNumbers);
    availableNumbers =  availableNumbers.incoming_phone_numbers;
    if (availableNumbers) {
      for (i=0 ; i < availableNumbers.length ; i++){
        phoneNumbers.push(availableNumbers[i].friendly_name);
      }
    }
    return phoneNumbers;
    }
catch(e){
   Logger.log(e);
   return phoneNumbers;
 }
}

//gets the associtated phone numbers with the account 
function formMule_getOutgoingNumbers(accountSid,authToken) {
  accountSid = accountSid||UserProperties.getProperty("twilioID");
  authToken =  authToken||UserProperties.getProperty("securityToken");
//create checks userproreties using 
  if (accountSid == null || accountSid == undefined){
    UserProperties.setProperty("twilioID", Browser.inputBox("Please Enter your Twilio ID"))
    accountSid = UserProperties.getProperty("twilioID");
    }
  if (authToken == null || authToken == undefined){
    UserProperties.setProperty("securityToken", Browser.inputBox("Please Enter your Twilio autherization key"))
    authToken = UserProperties.getProperty("securityToken");
    }
      
  var callObj = {
    //work around to get past the inability to use ":" in the urlfetch app
    headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSid + ":" + authToken)},    
    contentType: "application/xml; charset=utf-8",  
    };

  try{ 
  //change on country will add easy change when fully functional
  //returns an array of JSON objects 1 for each number
    var availableNumbers = UrlFetchApp.fetch(https + webName + accountSid + "/OutgoingCallerIds.json", callObj).getContentText();
    availableNumbers = JSON.parse(availableNumbers);
    availableNumbers =  availableNumbers.outgoing_caller_ids;
    var phoneNumbers = [];
    for (i=0 ; i < availableNumbers.length ; i++){
      phoneNumbers.push(availableNumbers[i].friendly_name);
  }
    debugger;
   return phoneNumbers;
    }
catch(e){
    Logger.log(e);
  }
  
}
