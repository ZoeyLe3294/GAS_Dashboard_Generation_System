function callsUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet){
  //grab data result after run query
  var queryBody = sourceUrl
  var dataQuery = runQuery(queryBody)
  var extractSheet = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  extractSheet.getRange(extractSheet.getLastRow()+1,1,dataQuery.length,dataQuery[0].length).setValues(dataQuery)
}
function loadQuery(){
 var callSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1vtHfz2H23BB1s3h2F4MP1jGuIU7isG2MRh9e2tYsm1o/edit#gid=981879551").getSheetByName("Tableau")
 var data = runQuery(callSheet.getRange('A1').getNote())
callSheet.getRange(callSheet.getLastRow()+1,1,data.length,data[0].length).setValues(data)
}