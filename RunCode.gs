var dailyCOsheetname = 'Daily CO'
var weeklyCOsheetname = 'Weekly CO'
var monthlyCOsheetname = 'Monthly CO'
var dailyCSsheetname = 'Daily CS'
var weeklyCSsheetname = 'Weekly CS'
var monthlyCSsheetname = 'Monthly CS'
var dailyReasonCOsheetname = 'Daily Reason CO'
var weeklyReasonCOsheetname = 'Weekly Reason CO'
var monthlyReasonCOsheetname = 'Monthly Reason CO'
var dailyReasonCSsheetname = 'Daily Reason CS'
var weeklyReasonCSsheetname = 'Weekly Reason CS'
var monthlyReasonCSsheetname = 'Monthly Reason CS'
var dailyCCsheetname = 'Daily CC'
var weeklyCCsheetname = 'Weekly CC'
var monthlyCCsheetname = 'Monthly CC'
var dailyReasonsheetname = 'Daily Reasons Overall'
var weeklyReasonsheetname = 'Weekly Reasons Overall'
var monthlyReasonsheetname = 'Monthly Reasons Overall'
var dailyPropensitysheetname = 'Daily Propensity'
var weeklyPropensitysheetname = 'Weekly Propensity'
var monthlyPropensitysheetname = 'Monthly Propensity'

function runOverallReason(){
  // var setup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Format Master").getRange(1,1).getDataRegion().getValues()
  // var configTitle = setup[0]
  // var runSystem = setup.filter(function(row){return row[configTitle.indexOf("Run?")]})[0]
  // var configSheet = runSystem[configTitle.indexOf("System Name")]
  var configSheet = 'System Setup Overall Reasons'
  var dashboardId = test ? '1-Y7R4ksaexxxBtouvMSUbyKHSikNCX71xnqAcqazb4c' : '1-Y7R4ksaexxxBtouvMSUbyKHSikNCX71xnqAcqazb4c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonsheetname)
  runF(configSheet,dailyDash,weeklyDash,monthlyDash)
  if(dailyDash){runFormat(dailyDash)}
  if(weeklyDash){runFormat(weeklyDash)}
  if(monthlyDash){runFormat(monthlyDash)} 
}
function updateOverallReason(){
  // var setup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Format Master").getRange(1,1).getDataRegion().getValues()
  // var configTitle = setup[0]
  // var runSystem = setup.filter(function(row){return row[configTitle.indexOf("Run?")]})[0]
  // var configSheet = runSystem[configTitle.indexOf("System Name")]
  var configSheet = 'System Setup Overall Reasons'
  var dashboardId = test ? '1-Y7R4ksaexxxBtouvMSUbyKHSikNCX71xnqAcqazb4c' : '1-Y7R4ksaexxxBtouvMSUbyKHSikNCX71xnqAcqazb4c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonsheetname)
  updateWeekF(configSheet,dailyDash,weeklyDash,monthlyDash)
}
function runOverallPropensity(){
  // var setup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Format Master").getRange(1,1).getDataRegion().getValues()
  // var configTitle = setup[0]
  // var runSystem = setup.filter(function(row){return row[configTitle.indexOf("Run?")]})[0]
  // var configSheet = runSystem[configTitle.indexOf("System Name")]
  var configSheet = 'System Setup Overall Propensity'
  var dashboardId = test ? '1kmUpzmsf-_cLj1SpufVA0_k1k23uEeXmY-dz8mjcOJA' : '1kmUpzmsf-_cLj1SpufVA0_k1k23uEeXmY-dz8mjcOJA'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Daily Propensity')
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Weekly Propensity')
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Monthly Propensity')
  runF(configSheet,dailyDash,weeklyDash,monthlyDash)
  if(dailyDash){runFormat(dailyDash)}
  if(weeklyDash){runFormat(weeklyDash)}
  if(monthlyDash){runFormat(monthlyDash)} 
}
function updateOverallPropensity(){
  var configSheet = 'System Setup Overall Propensity'
  var dashboardId = test ? '1kmUpzmsf-_cLj1SpufVA0_k1k23uEeXmY-dz8mjcOJA' : '1kmUpzmsf-_cLj1SpufVA0_k1k23uEeXmY-dz8mjcOJA'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Daily Propensity')
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Weekly Propensity')
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Monthly Propensity')
  updateWeekF(configSheet,dailyDash,weeklyDash,monthlyDash)
}
function runReasonCC(){
  var configSheet = 'System Setup Reason CC'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '10JN1aVPchEu06npWx0LmC7UxXNKffwThAHCsZam4QQY'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Daily Reason CC')
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Weekly Reason CC')
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Monthly Reason CC')
  runF(configSheet,dailyDash,weeklyDash,monthlyDash)
  // if(dailyDash){runFormat(dailyDash)}
  // if(weeklyDash){runFormat(weeklyDash)}
  // if(monthlyDash){runFormat(monthlyDash)} 
  // removeFormat(dailyDash)
  // removeFormat(weeklyDash)
  // removeFormat(monthlyDash)
}
function runReasonLS(){
  var configSheet = 'System Setup Reason LS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1SZq_A1iUAZZUxq486CcjLzpCXrlFeT7IK69nZtv0fus'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Daily Reason LS')
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Weekly Reason LS')
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Monthly Reason LS')
  // updateWeekF(configSheet,dailyDash,weeklyDash,monthlyDash)
  if(dailyDash){runFormat(dailyDash)}
  if(weeklyDash){runFormat(weeklyDash)}
  if(monthlyDash){runFormat(monthlyDash)} 
  // removeFormat(dailyDash)
  // removeFormat(weeklyDash)
  // removeFormat(monthlyDash)
}
function updateReasonCC(){
  var configSheet = 'System Setup Reason CC'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '10JN1aVPchEu06npWx0LmC7UxXNKffwThAHCsZam4QQY'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Daily Reason CC')
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Weekly Reason CC')
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Monthly Reason CC')
  updateWeekF(configSheet,dailyDash,weeklyDash,monthlyDash)
}
function runCOca(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Canada')
}
function runCOuk(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'United Kingdom')
}
function runCOau(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Australia')
}
function runCOie(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Ireland')
}
function runCOnz(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'New Zealand')
}
function runCOit(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Italy')
}
function runCOinit(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
    if(dailyDash){dailyDash.deleteRows(3, dailyDash.getLastRow()-2);dailyDash.clearConditionalFormatRules()}
  if(weeklyDash){weeklyDash.deleteRows(4, weeklyDash.getLastRow()-3);weeklyDash.clearConditionalFormatRules()}
  if(monthlyDash){monthlyDash.deleteRows(4,monthlyDash.getLastRow()-3);monthlyDash.clearConditionalFormatRules()}
}
function runCOfinal(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  // if(dailyDash){runFormat(dailyDash)}
  // if(weeklyDash){runFormat(weeklyDash)}
  if(monthlyDash){runFormat(monthlyDash)}
}
function runCSca(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Canada')
}
function runCSuk(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'United Kingdom')
}
function runCSau(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Australia')
}
function runCSie(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Ireland')
}
function runCSnz(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'New Zealand')
}
function runCSit(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runF(config,dailyDash,weeklyDash,monthlyDash,'Italy')
}
function runCSinit(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
    if(dailyDash){dailyDash.deleteRows(3, dailyDash.getLastRow()-2);dailyDash.clearConditionalFormatRules()}
  if(weeklyDash){weeklyDash.deleteRows(4, weeklyDash.getLastRow()-3);weeklyDash.clearConditionalFormatRules()}
  if(monthlyDash){monthlyDash.deleteRows(4,monthlyDash.getLastRow()-3);monthlyDash.clearConditionalFormatRules()}
}
function runCSfinal(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  // if(dailyDash){runFormat(dailyDash)}
  if(weeklyDash){runFormat(weeklyDash)}
  // if(monthlyDash){runFormat(monthlyDash)}
  // removeFormat(weeklyDash)
}

function updateCOca(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Canada')
}
function updateCOuk(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'United Kingdom')
}
function updateCOau(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Australia')
}
function updateCOie(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Ireland')
}
function updateCOnz(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'New Zealand')
}
function updateCOit(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Italy')
}
function updateCSca(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Canada')
}
function updateCSuk(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'United Kingdom')
}
function updateCSau(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)

  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Australia')
}
function updateCSie(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)

  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Ireland')
}
function updateCSnz(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)

  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'New Zealand')
}
function updateCSit(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)

  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'Italy')
}
function runCC(){
  var config = 'System Setup CC'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCCsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCCsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCCsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash)
}
function updateCC(){
  var config = 'System Setup CC'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCCsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCCsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCCsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash)
}

function runReasonCOinit(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  if(dailyDash){dailyDash.deleteRows(3, dailyDash.getLastRow()-2);dailyDash.clearConditionalFormatRules()}
  if(weeklyDash){weeklyDash.deleteRows(4, weeklyDash.getLastRow()-3);weeklyDash.clearConditionalFormatRules()}
  if(monthlyDash){monthlyDash.deleteRows(4,monthlyDash.getLastRow()-3);monthlyDash.clearConditionalFormatRules()}
}
function runReasonCOfinal(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  var officialDashId = '10JN1aVPchEu06npWx0LmC7UxXNKffwThAHCsZam4QQY'
  var officialDaily = SpreadsheetApp.openById(officialDashId).getSheetByName(dailyReasonCOsheetname)
  var officialWeekly = SpreadsheetApp.openById(officialDashId).getSheetByName(weeklyReasonCOsheetname)
  var officialMonthly = SpreadsheetApp.openById(officialDashId).getSheetByName(monthlyReasonCOsheetname)
  if(officialDaily.getLastRow()!=dailyDash.getLastRow()||officialWeekly.getLastRow()!=weeklyDash.getLastRow()||officialMonthly.getLastRow()!=monthlyDash.getLastRow()){
    SpreadsheetApp.getUi().alert("Unable to update since the number of metrics in the test file is different from the one in the official. Please review and rerun the finalize function.")
  }else{
    if(dailyDash){runFormat(dailyDash)}
    if(weeklyDash){runFormat(weeklyDash)}
    if(monthlyDash){runFormat(monthlyDash)}
    var data = dailyDash.getRange(1,7,dailyDash.getLastRow(),dailyDash.getLastColumn()-6).getValues()
    officialDaily.getRange(1,7,data.length,data[0].length).setValues(data)
    var data = weeklyDash.getRange(1,7,weeklyDash.getLastRow(),weeklyDash.getLastColumn()-6).getValues()
    officialWeekly.getRange(1,7,data.length,data[0].length).setValues(data)
    var data = monthlyDash.getRange(1,7,monthlyDash.getLastRow(),monthlyDash.getLastColumn()-6).getValues()
    officialMonthly.getRange(1,7,data.length,data[0].length).setValues(data)
  }
}
function runReasonCOoverall(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'CO Overall')
}
function runReasonCOca(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'CA')
}
function runReasonCOuk(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'UK')
}
function runReasonCOau(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'AU')
}
function runReasonCOie(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'IE')
}
function runReasonCOnz(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'NZ')
}
function runReasonCOit(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'IT')
}
function updateReasonCOoverall(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'CO Overall')
}
function updateReasonCOca(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'CA')
}
function updateReasonCOuk(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'UK')
}
function updateReasonCOau(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'AU')
}
function updateReasonCOie(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'IE')
}
function updateReasonCOnz(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'NZ')
}
function updateReasonCOit(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'IT')
}

function runReasonCSinit(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  if(dailyDash){dailyDash.deleteRows(3, dailyDash.getLastRow()-2);dailyDash.clearConditionalFormatRules()}
  if(weeklyDash){weeklyDash.deleteRows(4, weeklyDash.getLastRow()-3);weeklyDash.clearConditionalFormatRules()}
  if(monthlyDash){monthlyDash.deleteRows(4,monthlyDash.getLastRow()-3);monthlyDash.clearConditionalFormatRules()}
}
function runReasonCSfinal(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  if(officialDaily.getLastRow()!=dailyDash.getLastRow()||officialWeekly.getLastRow()!=weeklyDash.getLastRow()||officialMonthly.getLastRow()!=monthlyDash.getLastRow()){
    SpreadsheetApp.getUi().alert("Unable to update since the number of metrics in the test file is different from the one in the official. Please review and rerun the finalize function.")
  }else{
    if(dailyDash){runFormat(dailyDash)}
    if(weeklyDash){runFormat(weeklyDash)}
    if(monthlyDash){runFormat(monthlyDash)}
    var data = dailyDash.getRange(1,7,dailyDash.getLastRow(),dailyDash.getLastColumn()-6).getValues()
    officialDaily.getRange(1,7,data.length,data[0].length).setValues(data)
    var data = weeklyDash.getRange(1,7,weeklyDash.getLastRow(),weeklyDash.getLastColumn()-6).getValues()
    officialWeekly.getRange(1,7,data.length,data[0].length).setValues(data)
    var data = monthlyDash.getRange(1,7,monthlyDash.getLastRow(),monthlyDash.getLastColumn()-6).getValues()
    officialMonthly.getRange(1,7,data.length,data[0].length).setValues(data)
  }
}
function runReasonCSoverall(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'CS Overall')
}
function runReasonCSca(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'CA')
}
function runReasonCSuk(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'UK')
}
function runReasonCSau(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'AU')
}
function runReasonCSie(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'IE')
}
function runReasonCSnz(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'NZ')
}
function runReasonCSit(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runF(config,dailyDash,weeklyDash,monthlyDash,'IT')
}
function updateReasonCSoverall(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'CS Overall')
}
function updateReasonCSca(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'CA')
}
function updateReasonCSuk(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'UK')
}
function updateReasonCSau(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'AU')
}
function updateReasonCSie(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'IE')
}
function updateReasonCSnz(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'NZ')
}
function updateReasonCSit(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  updateWeekF(config,dailyDash,weeklyDash,monthlyDash,'IT')
}
function removeFormatReasonCS(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '10JN1aVPchEu06npWx0LmC7UxXNKffwThAHCsZam4QQY'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  removeFormat(dailyDash)
  removeFormat(weeklyDash)
  removeFormat(monthlyDash)
}
function removeFormatReasonCO(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '10JN1aVPchEu06npWx0LmC7UxXNKffwThAHCsZam4QQY'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  removeFormat(dailyDash)
  removeFormat(weeklyDash)
  removeFormat(monthlyDash)
}
function removeFormatCO(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  removeFormat(dailyDash)
  removeFormat(weeklyDash)
  removeFormat(monthlyDash)
}
function removeFormatCS(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  removeFormat(dailyDash)
  removeFormat(weeklyDash)
  removeFormat(monthlyDash)
}
function removeFormatCC(){
  var config = 'System Setup CC'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCCsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCCsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCCsheetname)
  removeFormat(dailyDash)
  removeFormat(weeklyDash)
  removeFormat(monthlyDash)
}
function applyFormatReasonCS(){
  var config = 'System Setup Reason CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '10JN1aVPchEu06npWx0LmC7UxXNKffwThAHCsZam4QQY'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCSsheetname)
  runFormat(dailyDash)
  runFormat(weeklyDash)
  runFormat(monthlyDash)
}
function applyFormatReasonCO(){
  var config = 'System Setup Reason CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '10JN1aVPchEu06npWx0LmC7UxXNKffwThAHCsZam4QQY'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyReasonCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyReasonCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyReasonCOsheetname)
  runFormat(dailyDash)
  runFormat(weeklyDash)
  runFormat(monthlyDash)
}
function applyFormatCO(){
  var config = 'System Setup CO'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'  
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCOsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCOsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCOsheetname)
  runFormat(dailyDash)
  runFormat(weeklyDash)
  runFormat(monthlyDash)
}
function applyFormatCS(){
  var config = 'System Setup CS'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCSsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCSsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCSsheetname)
  // var mapSource = SpreadsheetApp.openById(mapSourceId)
  runFormat(dailyDash)
  runFormat(weeklyDash)
  runFormat(monthlyDash)
}
function applyFormatCC(){
  var config = 'System Setup CC'
  var dashboardId = test ? '1lEvmXQ1w4FeAiTirxmpfSSwILCprT3g8IpFNZJuKV2c' : '1zP6CoGrP1_onnzsRF6ZPmD1HtiLc1km_y5Z0r7-Ff7c'
  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(dailyCCsheetname)
  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(weeklyCCsheetname)
  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName(monthlyCCsheetname)
  runFormat(dailyDash)
  runFormat(weeklyDash)
  runFormat(monthlyDash)
}


