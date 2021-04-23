function incidentTaskUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet) {
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var colUsed = colUsed.split(',')
  //check if targetSheet exists in targetUrl
  if(target==null){
  var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  }
  //weekly run
  var dataVal = source.getRange(4,1,source.getLastRow(),source.getLastColumn()).getValues().filter(function(r){return r[0]!=''}).map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1],item[colUsed[2]-1]]})
  dataVal=dataVal.filter(function(r){return timeZone(r[0],'startWeekDate')==timeZone(thisSunday,'startWeekDate')})
  if(dataVal){target.getRange(target.getLastRow()+1,1,dataVal.length,dataVal[0].length).setValues(dataVal)}

  //rerun data if needed
    // target.getRange(2,1,target.getLastRow()-1,3).clear() 
    // var dataVal = source.getRange(4,1,source.getLastRow(),source.getLastColumn()).getValues().filter(function(r){return r[0]!=''}).map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1],item[colUsed[2]-1]]})
    // dataVal=dataVal.filter(function(r){return timeZone(r[0],'startWeekDate')<=timeZone(thisSunday,'startWeekDate')})
    // if(dataVal){target.getRange(2,1,dataVal.length,dataVal[0].length).setValues(dataVal)}
}