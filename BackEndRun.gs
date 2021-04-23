//Backend MAIN FUNCTION
//Use to update all metrics
function runB() {
  //get config + filter config that has output file
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SET UP')
  var configData = configSheet.getRange(startRowB,startColIp,configSheet.getLastRow(),lastColB-startColIp).getValues()
  configData = configData.filter(function(row){return row[targetUrl].indexOf('https://docs.google.com')>-1})
  //loop through config row to run appropriate data process function
  configData.forEach(function(row){
    switch (row[outputName])
    {
        //  case 'OV Forecast': ovForecastUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]); break;
        //  case 'Forecasted Hours': forecastHourUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]); break;
        //  case 'QA Team Level': qaUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
        //  case 'Adherence': adherenceUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
        // case 'Network Logistics Metrics for LC': networkLogisticUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
        // DNU - case 'Calls': callsUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
        // case 'Reimbursement': reimbursementUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break; 
        // case 'Incidents Team Task': incidentTaskUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
        case 'Menu Team Tasks': menuTaskUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break; 
        // case 'Reasons Contacts': runReasonContact(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;   
        //ADD MORE OUTPUT FUNCTION HERE
    }
  })
}
//Use to update specific metric
function updateBackEnd(option) {
  //get config + filter config that has output file
  var configData = configSheet.getRange(startRowB,startColIp,configSheet.getLastRow()-startRowB-1,lastColB-startColIp).getValues()
  configData = configData.filter(function(row){return row[targetUrl].indexOf('https://docs.google.com')>-1 && row[outputName]==option})
  //  Logger.log(configData)
  //loop through config row to run appropriate data process function
  configData.forEach(function(row){
    switch (row[outputName])
    {
         case 'OV Forecast': ovForecastUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]); break;
         case 'Forecasted Hours': forecastHourUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]); break;
         case 'QA Team Level': qaUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
         case 'Adherence': adherenceUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
         case 'Chats Live': chatsLiveUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
        case 'Network Logistics Metrics for LC': networkLogisticUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break; 
        case 'Chats Live': chatsLiveUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;     
        //ADD MORE OUTPUT FUNCTION HERE
    }
  })
}