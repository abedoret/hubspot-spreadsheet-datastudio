/**
 * ###########################################################################
 * # Name: Hubspot Automation                                                #
 * # Description: This script let's you connect to Hubspot CRM and retrieve  #
 * #              its data to populate a Google Spreadsheet.                 #
 * # Date: March 11th, 2018                                                  #
 * # Author: Alexis Bedoret                                                  #
 * # Detail of the turorial: https://goo.gl/64hQZb                           #
 * ###########################################################################
 */

/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # ------------------------------- CONFIG -------------------------------- #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Fill in the following variables
 */
var CLIENT_ID = '';
var CLIENT_SECRET = '';
var SCOPE = 'contacts';
var AUTH_URL = "https://app.hubspot.com/oauth/authorize";
var TOKEN_URL = "https://api.hubapi.com/oauth/v1/token";
var API_URL = "https://api.hubapi.com";

/**
 * Create the following sheets in your spreadsheet
 * "Stages"
 * "Deals"
 */
var sheetNameStages = "Stages";
var sheetNameDeals = "Deals";
var sheetNameLogSources = "Log: Sources";
var sheetNameLogStages = "Log: Stages";



/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # --------------------------- AUTHENTICATION ---------------------------- #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Authorizes and makes a request to get the deals from Hubspot.
 */
function  getOAuth2Access() {
  var service = getService();
  if (service.hasAccess()) {
    // ... do whatever ...
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService().reset();
}

/**
 * Configures the service.
 */
function getService() {
  return OAuth2.createService('hubspot')
      // Set the endpoint URLs.
      .setTokenUrl(TOKEN_URL)
      .setAuthorizationBaseUrl(AUTH_URL)

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope(SCOPE);
}

/**
 * Handles the OAuth2 callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register.
 */
function logRedirectUri() {
  Logger.log(getService().getRedirectUri());
}



/**
 * ###########################################################################
 * # ----------------------------------------------------------------------- #
 * # ------------------------------- GET DATA ------------------------------ #
 * # ----------------------------------------------------------------------- #
 * ###########################################################################
 */

/**
 * Get the different stages in your Hubspot pipeline
 * API & Documentation URL: https://developers.hubspot.com/docs/methods/deal-pipelines/get-deal-pipeline
 */
function getStages() {
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  
  // API request
  var pipeline_id = "default"; // Enter your pipeline id here.
  var url = API_URL + "/crm-pipelines/v1/pipelines/deals";
  var response = UrlFetchApp.fetch(url, headers);
  var result = JSON.parse(response.getContentText());
  var stages = Array();
  
  // Looping through the different pipelines you might have in Hubspot
  result.results.forEach(function(item) {
    if (item.pipelineId == pipeline_id) {
      var result_stages = item.stages;
      // Let's sort the stages by displayOrder
      result_stages.sort(function(a,b) {
        return a.displayOrder-b.displayOrder;
      });
  
      // Let's put all the used stages (id & label) in an array
      result_stages.forEach(function(stage) {
        stages.push([stage.stageId,stage.label]);  
      });
    }
  });
  
  return stages;
}

/**
 * Get the deals from your Hubspot pipeline
 * API & Documentation URL: https://developers.hubspot.com/docs/methods/deals/get-all-deals
 */
function getDeals() {
  // Prepare authentication to Hubspot
  var service = getService();
  var headers = {headers: {'Authorization': 'Bearer ' + service.getAccessToken()}};
  
  // Prepare pagination
  // Hubspot lets you take max 250 deals per request. 
  // We need to make multiple request until we get all the deals.
  var keep_going = true;
  var offset = 0;
  var deals = Array();

  while(keep_going)
  {
    // We'll take three properties from the deals: the source, the stage & the amount of the deal
    var url = API_URL + "/deals/v1/deal/paged?properties=dealstage&properties=source&properties=amount&limit=250&offset="+offset;
    var response = UrlFetchApp.fetch(url, headers);
    var result = JSON.parse(response.getContentText());
    
    // Are there any more results, should we stop the pagination ?
    keep_going = result.hasMore;
    offset = result.offset;
    
    // For each deal, we take the stageId, source & amount
    result.deals.forEach(function(deal) {
      var stageId = (deal.properties.hasOwnProperty("dealstage")) ? deal.properties.dealstage.value : "unknown";
      var source = (deal.properties.hasOwnProperty("source")) ? deal.properties.source.value : "unknown";
      var amount = (deal.properties.hasOwnProperty("amount")) ? deal.properties.amount.value : 0;
      deals.push([stageId,source,amount]);
    });
  }
  
  return deals;
}



/**
* ###########################################################################
* # ----------------------------------------------------------------------- #
* # -------------------------- WRITE TO SPREADSHEET ----------------------- #
* # ----------------------------------------------------------------------- #
* ###########################################################################
*/

/**
 * Print the different stages in your pipeline to the spreadsheet
 */
function writeStages(stages) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameStages);
  
  // Let's put some headers and add the stages to our table
  var matrix = Array(["StageID","StageName"]);
  matrix = matrix.concat(stages);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}

/**
 * Print the different deals that are in your pipeline to the spreadsheet
 */
function writeDeals(deals) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetNameDeals);
  
  // Let's put some headers and add the deals to our table
  var matrix = Array(["StageID","Source", "Amount"]);
  matrix = matrix.concat(deals);
  
  // Writing the table to the spreadsheet
  var range = sheet.getRange(1,1,matrix.length,matrix[0].length);
  range.setValues(matrix);
}



/**
* ###########################################################################
* # ----------------------------------------------------------------------- #
* # -------------------------------- ROUTINE ------------------------------ #
* # ----------------------------------------------------------------------- #
* ###########################################################################
*/

/**
 * This function will update the spreadsheet. This function should be called
 * every hour or so with the Project Triggers.
 */
function refresh() {
  var service = getService();
  
  if (service.hasAccess()) {
    var stages = getStages();
    writeStages(stages);
  
    var deals = getDeals();
    writeDeals(deals);
    
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

/**
 * This function will log the amount of leads per stage over time
 * and print it into the sheet "Log: Stages"
 * It should be called once a day with a Project Trigger
 */
function logStages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Stages: Count");
  var getRange = sheet.getRange("B2:B12");
  var row = getRange.getValues();
  row.unshift(new Date);
  var matrix = [row];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log: Stages");
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
    
    // Writing at the end of the spreadsheet
  var setRange = sheet.getRange(lastRow+1,1,1,row.length);
  setRange.setValues(matrix);
}

/**
 * This function will log the amount of leads per source over time
 * and print it into the sheet "Log: Sources"
 * It should be called once a day with a Project Trigger
 */
function logSources() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sources: Count & Conversion Rates");
  var getRange = sheet.getRange("M3:M13");
  var row = getRange.getValues();
  row.unshift(new Date);
  var matrix = [row];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Log: Sources");
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
    
    // Writing at the end of the spreadsheet
  var setRange = sheet.getRange(lastRow+1,1,1,row.length);
  setRange.setValues(matrix);
}
