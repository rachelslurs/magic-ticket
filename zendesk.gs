// Many thanks to Romain Vialard for the guidance his 'Add reminder' Google Apps Script provided in creating this script.
// Many thanks to Arun Nagarajan as well. His Google OAuth2 gists on GitHub were extremely helpful for this first time OAuthor!

// this is the user property where we'll store the token, make sure this is unique across all user properties across all scripts
var tokenPropertyName = 'oAuthToken'; 

function doGet(e) {
  var HTMLToOutput;
  if(e.parameters.code){
    // if we get 'code' as a parameter in, then this is a callback.
    getAndStoreAccessToken(e.parameters.code);
    HTMLToOutput = '<html><h1>Finished with oAuth</h1></html>';
  }
  else if(isTokenValid()){
    // if we already have a valid token, go off and start working with data
    HTMLToOutput = '<html><h1>Already have token</h1></html>';
  }
  else {
    // we are starting from scratch or resetting
    return HtmlService.createHtmlOutput("<html><h1>Lets start with oAuth</h1><a href='"+ getURLForAuthorization() + "'>click here to start</a></html>");
  }
  HTMLToOutput += getData();
  return HtmlService.createHtmlOutput(HTMLToOutput);
}

// check to see if we have a token
function getData(){
  var properties = ScriptProperties.getProperty('userSettings');
  properties = Utilities.jsonParse(properties);
  var zdURL = properties.zdURL;
  var getDataURL = 'https://' + zdURL + '/api/v2/oauth/tokens/current.json';
  var dataResponse = UrlFetchApp.fetch(getDataURL,getUrlFetchOptions()).getContentText();  
  return dataResponse;
}

function getUrlFetchOptions() {
  var token = UserProperties.getProperty(tokenPropertyName);
  return {
    "contentType" : "application/json",
    "headers" : {
      "Authorization" : "Bearer " + token,
      "Accept" : "application/json"
    }
  };
}

function postData(){
  var properties = ScriptProperties.getProperty('userSettings');
  properties = Utilities.jsonParse(properties);
  var zdURL = properties.zdURL;
  var postDataURL = "https://" + zdURL + "/api/v2/tickets.json";
  var dataResponse = UrlFetchApp.fetch(postDataURL,postUrlFetchOptions()).getContentText();  
  return dataResponse;
}

function postUrlFetchOptions(message) {
  var token = UserProperties.getProperty(tokenPropertyName);
  var str = '"';
    for (var i = 0; i < message[0].length; i++) {
      str +=  message[0][i] + ": " + message[1][i] + " ";
    }
  str += '"';
  var options = 
      { "method": "post",
       "contentType" : "application/json",
    "headers": { "Authorization": "Bearer " + token, "Accept" : "application/json" },
    "payload": '{"ticket":{"subject":"Magic Ticket", "comment": { "body": ' + str + ' }, "priority": "urgent", "group_id": "20643765", "type": "task" }}'
      }
  return options;
}

// this is the URL where they'll authorize with zendesk.com
// may need to add a 'scope' param here.
// example scope for google - https://www.googleapis.com/plus/v1/activities

function getURLForAuthorization(){
  var properties = ScriptProperties.getProperty('userSettings');
  properties = Utilities.jsonParse(properties);
  var zdURL = properties.zdURL;
  var clientName = properties.clientName;
  var rURL = ScriptApp.getService().getUrl();
  var authorizeURL = 'https://' + zdURL + '/oauth/authorizations/new'; //step 1. we can actually start directly here if that is necessary
  return authorizeURL + '?response_type=code&client_id='+ clientName + '&redirect_uri='+ rURL +
    '&scope=read%20write';  
}

function getAndStoreAccessToken(code){
  var properties = ScriptProperties.getProperty('userSettings');
  properties = Utilities.jsonParse(properties);
  var zdURL = properties.zdURL;
  var tokenURL = 'https://' + zdURL + '/oauth/tokens'; //step 2. after we get the callback, go get token
  var clientName = properties.clientName;
  var clientSecret= properties.clientSecret;
  var rURL= ScriptApp.getService().getUrl();
  var parameters = {
    method : 'post',
    payload : 'client_id='+ clientName +'&client_secret=' + clientSecret + '&grant_type=authorization_code&redirect_uri=' + rURL + '&code=' + code + '&scope=read'
  };
  var response = UrlFetchApp.fetch(tokenURL,parameters).getContentText();   
  var tokenResponse = JSON.parse(response);
  // store the token for later retrieval
  UserProperties.setProperty(tokenPropertyName, tokenResponse.access_token);
  //reset token option?
}

function isTokenValid() {
  var token = UserProperties.getProperty(tokenPropertyName);
  if(!token){ //if its empty or undefined
    return false;
  }
  return true; //naive check
  
  //if your API has a more fancy token checking mechanism, use it. for now we just check to see if there is a token. 
  /*
  var responseString;
  try{
  responseString = UrlFetchApp.fetch(BASE_URI+'/api/rest/system/session/check',getUrlFetchOptions(token)).getContentText();
  }catch(e){ //presumably an HTTP 401 will go here
  return false;
  }
  if(responseString){
  var responseObject = JSON.parse(responseString);
  return responseObject.authenticated;
  }
  return false;*/
}

function dateChecker(e) {
  var properties = ScriptProperties.getProperty('userSettings');
  properties = Utilities.jsonParse(properties);
  var zdURL = properties.zdURL;
  var postDataURL = "https://" + zdURL + "/api/v2/tickets.json";
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperty('userSettings');
  var logSheet = doc.getSheetByName('Log');
  if (properties != null) {
    properties = Utilities.jsonParse(properties);
    var sheet = doc.getSheetByName(properties.sheet);
    var lastRow = sheet.getLastRow();
    var data = sheet.getDataRange().getValues();
    var dateColumns = properties.dateColumn.replace(/\s/g, '').split(',');
    for (var r = 0; r < dateColumns.length; r++) {
      var dates = sheet.getRange(dateColumns[r] + '1:' + dateColumns[r] + lastRow.toString()).getValues();
      var comments = sheet.getRange(dateColumns[r] + '1:' + dateColumns[r] + lastRow.toString()).getComments();
      for (var i = 1; i < dates.length; i++) {
        var rowResp = [];
        if (comments[i][0] != 'ticket sent') {
          var expiry_date = new Date(dates[i][0]);
          var today = new Date();
          var reminder = properties.reminderTime * 24 * 3600 * 1000;
          if (reminder > 0 && expiry_date.getTime() - today.getTime() < reminder || reminder < 0 && expiry_date.getTime() - today.getTime() < reminder)
          {
            try {
              var dataResponse = UrlFetchApp.fetch(postDataURL,postUrlFetchOptions([data[0], data[i]])); 
              var logResponse=dataResponse.getContentText();
              Logger.log(dataResponse);
              sheet.getRange(dateColumns[r] + (i + 1).toString()).setComment('ticket sent');
              var parseResponse = JSON.parse(logResponse);
              var respCode = dataResponse.getResponseCode();
              if (respCode == 201) // success!
              {
                logSheet.insertRowAfter(1);
                rowResp = [new Date(),parseResponse.ticket["id"], parseResponse.ticket["url"]];
                logSheet.getRange(2, 1, 1, 3).setValues([rowResp]);
              }
            } catch(e) {
              // que problemo?
              {
                var rowResp = [new Date(),"Error",e.message];
              }
            }
          }
          else
          {
            rowResp = [new Date(),"Script finished","No other tickets needed to be created."];
          }
        }
        
      }
      if (doc.getSheetByName('Log')) 
          {
            logSheet.insertRowAfter(1);
            logSheet.getRange(2, 1, 1, 3).setValues([rowResp]);
          }
    }
  }
}
