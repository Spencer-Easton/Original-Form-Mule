function formMule_whatIs() {
  var app = UiApp.createApplication().setHeight(550);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = app.createVerticalPanel();
  var muleGrid = app.createGrid(1, 2);
  var image = app.createImage(this.MULEICONURL);
  image.setHeight("100px");
  var label = app.createLabel("formMule: A flexible email, calendar, SMS, and voice merge utility for use with Google Spreadsheet and/or Form data.");
  label.setStyleAttribute('fontSize', '1.5em').setStyleAttribute('fontWeight', 'bold');
  muleGrid.setWidget(0, 0, image);
  muleGrid.setWidget(0, 1, label);
  var mainGrid = app.createGrid(4, 1);
  var html = "<h3>Features</h3>";
      html += "<ul><li>Set up and generate templated, merged emails from form or spreadsheet data.</li>";
      html += "<li>Optionally set 'send conditions' that trigger up to six different emails based on a column's value in the merged row.  Allows for branching or differentiated outputs based on the value of individual form responses.</li>";
      html += "<li>Can be set to auto-copy-down formula columns (on form submit) that operate to the right of form data.  Great for use with VLOOKUP and IF formulas that reference form data.  For example, look up an email address in another sheet based on a name submitted in the form.  This feature is available under \"Advanced options.\"</li>"; 
      html += "<li>Auto-generate and auto-update calendar events using form or spreadsheet data and conditions.</li>";
      html += "<li>Easily connects to the Twilio SMS and Voice service to enable SMS and Voice merges from Spreadsheet or Form data.</li>";
      html += "<li>Can be triggered on form submit or run manually after merge preview.</li>"; 
      html += "<li>Configuration settings can be exported to allow Spreadsheet-based systems to be packaged for easy replication through File->Copy in Google Drive.</li>";
      html += "<li>Interested in doing Google Document or PDF merges instead? Check out the autoCrat script in the gallery.</li></ul>";
    
  mainGrid.setWidget(0, 0, app.createHTML(html));
  var sponsorLabel = app.createLabel("Brought to you by");
  var sponsorImage = app.createImage("http://www.youpd.org/sites/default/files/acquia_commons_logo36.png");
  var supportLink = app.createAnchor('Get the tutorials!', 'http://cloudlab.newvisions.org/scripts/formmule');
  mainGrid.setWidget(1, 0, sponsorLabel);
  mainGrid.setWidget(2, 0, sponsorImage);
  mainGrid.setWidget(3, 0, supportLink);
  app.add(muleGrid);
  panel.add(mainGrid);
  app.add(panel);
  ss.show(app);
  return app;                                                                    
}