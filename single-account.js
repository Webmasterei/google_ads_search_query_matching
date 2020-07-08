var GOOGLE_DOC_URL = "INSERT SPREADSHEET URL HERE"; 
var TIMESPAN = "180"; 
var accountLabel = "Search Query Conversion Report"; 

function main() { 
  var results = runQueryReport(); 
   modifySpreadSheet(results); 
}
// check a query for whether the keyword exists in the account
// returns true or false 

function keywordExists(keyword) {
  var kw = keyword;
  if (kw != null) {
    try{
      kwIter = AdWordsApp.keywords().withCondition("Text = \'"+kw+"\'").withCondition("Status = ENABLED").get();
      var exists = kwIter.totalNumEntities() > 0 ? true : false;
      if(!exists){
        kwIter = AdWordsApp.keywords().withCondition("Text CONTAINS_IGNORE_CASE \'"+kw+"\'").withCondition("Status = ENABLED").get();
        var existsBroad = kwIter.totalNumEntities() > 0 ? true : false;
        if(existsBroad){
        while (kwIter.hasNext()) {
        var keyword = kwIter.next();
          keyword = normalizeKeyword(keyword.getText());
          kw = normalizeKeyword(kw);
          if(keyword == kw){
            exists = true;
            if(exists){
              return exists;
            }
          }
        }
        }
      }
      return exists;
    }
    catch(err){
      Logger.log(err.message)
    }
  }
}

function runQueryReport() {
    var timespan = getTimespan(TIMESPAN);
    var listOfQueries = [];
  
	 var report = AdWordsApp.report(
	     'SELECT Query, CampaignName, AdGroupName, KeywordTextMatchingQuery, QueryMatchTypeWithVariant, Conversions, ConversionValue, Cost, AverageCpc, Clicks, Impressions, Ctr, ConversionRate ' +
	     'FROM SEARCH_QUERY_PERFORMANCE_REPORT ' +
         'WHERE Conversions > 0 ' +
         'DURING ' + timespan["from_date"] +', '+ timespan["to_date"] +' ');    
	       
     var rows = report.rows();

	 while (rows.hasNext()) {
		   var row = rows.next();
              
       var query = row['Query'];
		   var campaign= row['CampaignName'];
       var adgroup = row['AdGroupName'];    
       var keyword = row['KeywordTextMatchingQuery'];
       var matchType = row['QueryMatchTypeWithVariant']
		   var conversions = row['Conversions'];
		   var conversionValue = row['ConversionValue'];
		   var cost = row['Cost'];
       var roas = conversionValue / cost;
		   var averageCpc = row['AverageCpc'];
	     var clicks = row['Clicks'];
       var impressions = row['Impressions'];
		   var ctr = row['Ctr'];
		   var conversionRate = row['ConversionRate'];
       var keyword_exists = keywordExists(query);
       var queryResult = new queryData(query, keyword, matchType, adgroup, campaign, conversions, conversionValue,cost, roas,averageCpc, clicks, impressions, ctr, conversionRate, keyword_exists);
       
           listOfQueries.push(queryResult);
           
	 }  // end of report run
    
	 return listOfQueries;
     
} 

function queryData(query, keyword, matchType, adgroup, campaign, conversions, conversionValue, cost, roas, averageCpc, clicks, impressions, ctr, conversionRate, exists ) {
	this.query = query;
    this.keyword = keyword;
  this.matchType = matchType;
	this.campaign = campaign;
	this.adgroup = adgroup;
	this.conversions = conversions;
	this.conversionValue = conversionValue;
	this.cost = cost;
    this.roas = roas;
	this.averageCpc = averageCpc;
	this.clicks = clicks;
	this.impressions = impressions;
	this.ctr = ctr;
	this.conversionRate = conversionRate;
	this.exists = exists;
} // end of productData
function getOldComments(sheet){
  // Get Old Comments
  var oldComments = {};
  var lastRow = sheet.getLastRow();
  if(!lastRow){return;}
  var lastCol = sheet.getLastColumn();
  if(!lastCol){return;}
  var range = sheet.getRange(2,1, lastRow, lastCol);
  for (var i = 1; i <= lastRow; i++) {
      var oldQuery = range.getCell(i,1).getValue();
      var oldComment = range.getCell(i,3).getValue();
    if(oldComment){
      oldComments[oldQuery] = oldComment;
    }
  }
  return oldComments;
}

function modifySpreadSheet(results) {
  
  var queryResults = results;
  
  var querySS = SpreadsheetApp.openByUrl(GOOGLE_DOC_URL);
  var account = AdWordsApp.currentAccount();
  var accountName = account.getName();
  var sheet = querySS.getSheetByName(accountName);
  if(!sheet){
    querySS.insertSheet(accountName);
    var sheet = querySS.getSheetByName(accountName);
  }
  var oldComments = getOldComments(sheet);
  sheet.clear();
  var columnNames = ["Query", "In Account", "Kommentar", "Keyword", "matchType", "Ad Group", "Campaign", "Conversions", "Conversion Value", "Cost", "ROAS", "Average CPC","Clicks", "Impressions", "Ctr", "ConversionRate"];
  
  var headersRange = sheet.getRange(1, 1, 1, columnNames.length);
  
  headersRange.setFontWeight("bold");
  headersRange.setFontSize(12);
  headersRange.setBorder(false, false, true, false, false, false);
   
   for (i = 0; i < queryResults.length; i++) {
	headersRange.setValues([columnNames]);
     if(queryResults[i].exists == false) {
	   var query = queryResults[i].query;
       var exists = (queryResults[i].exists == true) ? "Added" : "Not Added";
       var keyword = queryResults[i].keyword;
       var matchType = queryResults[i].matchType;
       var campaign = queryResults[i].campaign;
       var adgroup = queryResults[i].adgroup;
       var conversions  = parseFloat(queryResults[i].conversions);
       var conversionValue = parseFloat(queryResults[i].conversionValue);
	   var cost = parseFloat(queryResults[i].cost);
       var roas = (isNaN(queryResults[i].roas)) ? 0.00 : queryResults[i].roas;
	   var averageCpc = parseFloat(queryResults[i].averageCpc);
	   var clicks = queryResults[i].clicks;
	   var impressions = queryResults[i].impressions;
	   var ctr = queryResults[i].ctr;
	   var conversionRate = queryResults[i].conversionRate;
       if(oldComments[query]){
         var comment = oldComments[query];
         } else {
           var comment = "";
         }
      sheet.appendRow([query, exists, comment, "'"+keyword, matchType, adgroup, campaign, conversions, conversionValue, cost, roas, averageCpc, clicks, impressions, ctr, conversionRate]);
     }
    }  
  
  sheet.getRange("A2:N").setFontSize(10);
  
  sheet.getRange("H:H").setNumberFormat("0.00");
  sheet.getRange("I:I").setNumberFormat("0.00");
  sheet.getRange("J:J").setNumberFormat("0.00");
  sheet.getRange("K:K").setNumberFormat("0.00");
  sheet.getRange("L:L").setNumberFormat("0");
  sheet.getRange("M:M").setNumberFormat("0");  
  sheet.getRange("N:N").setNumberFormat("0.00");
  
  sheet.getRange("A2:N").sort([{column: 7, ascending: false}, {column: 13, ascending: false}]);   

}

// Helper functions
function warn(msg) {
  Logger.log('WARNING: '+msg);
}
 
function info(msg) {
  Logger.log(msg);
}

function getTimespan(TIMESPAN){
  var timeZone = AdWordsApp.currentAccount().getTimeZone();
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var now = new Date();
  var to_date = new Date(now.getTime() - MILLIS_PER_DAY);
  to_date = Utilities.formatDate(to_date, timeZone, 'yyyyMMdd')
  var from_date = new Date(now.getTime() - (MILLIS_PER_DAY * TIMESPAN));
  from_date = Utilities.formatDate(from_date, timeZone, 'yyyyMMdd')
  var timespan = {"to_date":to_date, "from_date":from_date}
  return timespan;
}
function toTitleCase(str) {
    return str.replace(/\w\S*/g, function(txt){
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
}

function normalizeKeyword(keyword){
  keyword = keyword.replace(/\[/g, '').replace(/\]/g, '');
  keyword = keyword.toLowerCase();
  keyword = keyword.replace(/\+/g,'');
  return keyword
}
