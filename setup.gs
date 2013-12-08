// Magic Ticket 1.0 (12/4/2013)
// By Rachel Cantor (https://plus.google.com/+RachelCantor)
// More info (http://blog.controlgroup.com/2013/12/08/connecting-zendesk-google-spreadsheets-using-google-apps-script)
// File issues, feature requests, and contribute on the GitHub repo (https://github.com/rachelslurs/magic-ticket)

var ss = SpreadsheetApp.getActiveSpreadsheet();

function onInstall() {
  onOpen();
}

function onOpen() {
  var menuEntries = 
    [{
      name: "Settings",
      functionName: "setPreferencesUI"
    },
     {
       name:"Run Magic Ticket",
       functionName: "dateChecker"
     }];
  ss.addMenu("Magic Ticket", menuEntries);
}

function setup() {
  
  // trying to avoid making users open script editor to save a version > deploy as web app
  // not working yet
  
  var url = "";
  try {
    ScriptApp.getService().enable(ScriptApp.getService().Restriction.MYSELF);
    url = ScriptApp.getService().getUrl();
    Browser.msgBox("Your web app is now accessible at the following URL:\n"
                   + url);
  } catch (e) {
    Browser.msgBox("Script authorization expired.\nPlease run it again.");
    ScriptApp.invalidateAuth();
  }  
}

function setPreferencesUI() {
  var app = UiApp.createApplication().setTitle('Magic Ticket Settings').setHeight(430);
  var mainPanel = app.createVerticalPanel().setWidth('100%');
  var settingsPanel = app.createVerticalPanel().setId('settingsPanel');
  var settingsGrid = app.createGrid(7, 2).setCellSpacing(20).setWidth('100%');
  var buttonPanel = app.createHorizontalPanel().setWidth('100%');
  var buttonAlign = app.createHorizontalPanel().setSpacing(10).setStyleAttribute('paddingTop', '20px');
  var close = app.createButton('Close', app.createServerHandler('closeApp_')).setWidth(100);
  var properties = ScriptProperties.getProperty('userSettings');
  // is this still necessary?
  // var user = Session.getEffectiveUser().getEmail();
  
  settingsPanel.add(settingsGrid);
  buttonPanel.add(buttonAlign);
  buttonPanel.setCellHorizontalAlignment(buttonAlign, UiApp.HorizontalAlignment.CENTER);
  settingsPanel.add(buttonPanel);
  mainPanel.add(settingsPanel);
  app.add(mainPanel);
  
  if (properties != null){
    properties = Utilities.jsonParse(properties);
  }
  settingsGrid.setWidget(0, 0, app.createLabel('Select sheet:'));
  settingsGrid.setWidget(1, 0, app.createLabel('Check dates in column(s): (A, B,...)'));
  settingsGrid.setWidget(2, 0, app.createLabel('Open a ticket'));
  settingsGrid.setWidget(3, 0, app.createLabel('Client Name/ID:'));
  settingsGrid.setWidget(4, 0, app.createLabel('Client Secret:'));
  settingsGrid.setWidget(5, 0, app.createLabel('Zendesk URL:'));
  settingsGrid.setWidget(6, 0, app.createLabel('Redirect URL:'));
  
  var sheetList = app.createListBox().setName('sheet');
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    var item = sheetList.addItem(sheetName);
    if (properties != null && sheetName == properties.sheet) {
      sheetList.setSelectedIndex(i);
    }
  }
  var dateColumn = app.createTextBox().setName('dateColumn').setWidth('30');
  var reminderTime = app.createTextBox().setName('reminderTime').setWidth('30');
  var clientName = app.createTextBox().setName('clientName').setWidth('100%');
  var clientSecret = app.createTextBox().setName('clientSecret').setWidth('100%');
  var zdURL = app.createTextBox().setName('zdURL').setWidth('100%');
  var rURL = app.createTextBox().setReadOnly(true).setText(ScriptApp.getService().getUrl()).setWidth('100%');
  if (properties != null) {
    dateColumn.setText(properties.dateColumn);
    reminderTime.setText(properties.reminderTime);
    clientName.setText(properties.clientName);
    clientSecret.setText(properties.clientSecret);
    zdURL.setText(properties.zdURL);
  }
  settingsGrid.setWidget(0, 1, sheetList);
  settingsGrid.setWidget(1, 1, dateColumn);
  settingsGrid.setWidget(2, 1, app.createHorizontalPanel().add(reminderTime).add(app.createLabel('days before date is reached.').setStyleAttribute('marginLeft', '10px')));
  settingsGrid.setWidget(3, 1, clientName);
  settingsGrid.setWidget(4, 1, clientSecret);
  settingsGrid.setWidget(5, 1, zdURL);
  settingsGrid.setWidget(6, 1, rURL);
  var record = app.createButton("Save", app.createServerHandler('saveProperties_').addCallbackElement(settingsGrid)).setEnabled(true);
  buttonAlign.add(record.setWidth(100));
  if (properties != null) {
    var resetReminder = app.createButton("Reset", app.createServerHandler('resetProperties_'));
    buttonAlign.add(resetReminder);
  }
  buttonAlign.add(close);
  
  // need to add validator for Zendesk URL
  
  var handlerOk = app.createClientHandler().validateMatches(dateColumn, "^[a-zA-Z,]{1,}$").validateRange(reminderTime, 1, null);
  handlerOk.forTargets(dateColumn, reminderTime).setStyleAttribute('color', 'black');
  handlerOk.forTargets(record).setEnabled(true);
  var handlerDateColumnOk = app.createClientHandler().validateMatches(dateColumn, "^[a-zA-Z,]{1,}$");
  handlerDateColumnOk.forEventSource().setStyleAttribute('color', 'black');
  var handlerDateColumnNok = app.createClientHandler().validateNotMatches(dateColumn, "^[a-zA-Z,]{1,}$");
  handlerDateColumnNok.forEventSource().setStyleAttribute('color', 'red');
  handlerDateColumnNok.forTargets(record).setEnabled(false);
  var handlerTimeOk = app.createClientHandler().validateRange(reminderTime, 1, null);
  handlerTimeOk.forEventSource().setStyleAttribute('color', 'black');
  var handlerTimeNok = app.createClientHandler().validateNotRange(reminderTime, 1, null);
  handlerTimeNok.forEventSource().setStyleAttribute('color', 'red');
  handlerTimeNok.forTargets(record).setEnabled(false);
  
  dateColumn.addKeyUpHandler(handlerOk).addKeyUpHandler(handlerDateColumnOk).addKeyUpHandler(handlerDateColumnNok);
  reminderTime.addKeyUpHandler(handlerOk).addKeyUpHandler(handlerTimeOk).addKeyUpHandler(handlerTimeNok);
  record.addClickHandler(app.createClientHandler().forEventSource().setEnabled(false).forTargets(close).setEnabled(false));
  
  ss.show(app);
}

function closeApp_() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function resetProperties_(e) {
  var app = UiApp.getActiveApplication();
  ScriptProperties.deleteAllProperties();
  var panel = app.getElementById('settingsPanel').clear();
  var info = app.createLabel('Settings reset.').setStyleAttribute('paddingBottom', '20px');
  panel.add(info).setCellHorizontalAlignment(info, UiApp.HorizontalAlignment.CENTER).setStyleAttribute('paddingTop', '50px').setWidth('100%');
  var currentTriggers = ScriptApp.getScriptTriggers();
  for (i in currentTriggers) {
    ScriptApp.deleteTrigger(currentTriggers[i]);
  }
  var button = app.createButton('Close', app.createServerHandler('closeApp_')).setWidth(100);
  panel.add(button).setCellHorizontalAlignment(button, UiApp.HorizontalAlignment.CENTER);
  return app;
}

function saveProperties_(e) {
  var app = UiApp.getActiveApplication();
  ScriptProperties.setProperty('userSettings', Utilities.jsonStringify(e.parameter));
  var panel = app.getElementById('settingsPanel').clear();
  var info = app.createLabel('Settings saved!').setStyleAttribute('paddingBottom', '20px');
  panel.add(info).setCellHorizontalAlignment(info, UiApp.HorizontalAlignment.CENTER).setStyleAttribute('padding', '50px').setWidth('100%');
  var currentTriggers = ScriptApp.getScriptTriggers();
  for (i in currentTriggers) {
    ScriptApp.deleteTrigger(currentTriggers[i]);
  }
  ScriptApp.newTrigger('dateChecker').timeBased().everyDays(1).atHour(7).create();
  var button = app.createButton('Close', app.createServerHandler('closeApp_')).setWidth(100);
  panel.add(button).setCellHorizontalAlignment(button, UiApp.HorizontalAlignment.CENTER);
  return app;
}
