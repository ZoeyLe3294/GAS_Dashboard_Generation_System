function networkLogisticUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet){
  //grab data result after run query
  var queryBody = sourceUrl
  var dataQuery = runQuery(queryBody) //[order_date	tenant_id	delivered_order	deliver_slack_over_20	collect_slack_over_10	ov_rejected]
  var extractSheet = SpreadsheetApp.openByUrl(targetUrl).getSheetByName('Query Result')
  extractSheet.getRange(extractSheet.getLastRow()+1,1,dataQuery.length,dataQuery[0].length).setValues(dataQuery)
  var rawData = extractSheet.getRange(2,1,extractSheet.getLastRow()-1,extractSheet.getLastColumn()).getValues()
  //aggregate data and put into correct target sheet
  var country=[]
  var result={}
  result.daily=rawData.map(function(r){
    if(country.indexOf(r[1])==-1){
      country.push(r[1])
    }
    return [r[0],r[1],r[2],r[3]/r[2],r[4]/r[2],r[5]/r[2],r[2]/r[6],r[7]]
  })

  var targetDaily = SpreadsheetApp.openByUrl(targetUrl).getSheetByName('Daily')
  var targetWeekly = SpreadsheetApp.openByUrl(targetUrl).getSheetByName('Weekly')
  var targetMonthly = SpreadsheetApp.openByUrl(targetUrl).getSheetByName('Monthly')

  targetDaily.getRange(2,1,result.daily.length,result.daily[0].length).setValues(result.daily)
  targetWeekly.getRange(2,1,targetWeekly.getLastRow()-1,targetWeekly.getLastColumn()).clear()
  targetMonthly.getRange(12,1,targetMonthly.getLastRow()-11,targetMonthly.getLastColumn()).clear()
  country.forEach(function(row){
    // Logger.log(rawData)
    var data = rawData.filter(function(r){return r[1]==row})
    result.weekly=[]
    result.monthly=[]
    var week=[]
    var month=[]
    data.forEach(function(entry){
      if(week.indexOf(timeZone(entry[0],'startWeekDate'))==-1){
          week.push(timeZone(entry[0],'startWeekDate'))
          result.weekly.push([timeZone(entry[0],'startWeekDate'),entry[1],entry[2],entry[3],entry[4],entry[5],entry[6],entry[7]])
      }else{
          result.weekly[week.indexOf(timeZone(entry[0],'startWeekDate'))][2]+=entry[2]
          result.weekly[week.indexOf(timeZone(entry[0],'startWeekDate'))][3]+=entry[3]
          result.weekly[week.indexOf(timeZone(entry[0],'startWeekDate'))][4]+=entry[4]
          result.weekly[week.indexOf(timeZone(entry[0],'startWeekDate'))][5]+=entry[5]
          result.weekly[week.indexOf(timeZone(entry[0],'startWeekDate'))][6]+=entry[6]
          result.weekly[week.indexOf(timeZone(entry[0],'startWeekDate'))][7]+=entry[7]
        }
      if(month.indexOf(timeZone(entry[0],'startMonthDate'))==-1){
        month.push(timeZone(entry[0],'startMonthDate'))
        result.monthly.push([timeZone(entry[0],'startMonthDate'),entry[1],entry[2],entry[3],entry[4],entry[5],entry[6],entry[7]])
      }else{
        result.monthly[month.indexOf(timeZone(entry[0],'startMonthDate'))][2]+=entry[2]
        result.monthly[month.indexOf(timeZone(entry[0],'startMonthDate'))][3]+=entry[3]
        result.monthly[month.indexOf(timeZone(entry[0],'startMonthDate'))][4]+=entry[4]
        result.monthly[month.indexOf(timeZone(entry[0],'startMonthDate'))][5]+=entry[5]
        result.monthly[month.indexOf(timeZone(entry[0],'startMonthDate'))][6]+=entry[6]
        result.monthly[month.indexOf(timeZone(entry[0],'startMonthDate'))][7]+=entry[7]
      }
    })
    result.weekly=result.weekly.map(function(r){return [r[0],r[1],r[2],r[3]/r[2],r[4]/r[2],r[5]/r[2],r[2]/r[6],r[7]]})
    result.monthly = result.monthly.map(function(r){return [r[0],r[1],r[2],r[3]/r[2],r[4]/r[2],r[5]/r[2],r[2]/r[6],r[7]]})
  //paste data
  Logger.log(result.weekly)
    targetWeekly.getRange(targetWeekly.getLastRow()+1,1,result.weekly.length,result.weekly[0].length).setValues(result.weekly)
    targetMonthly.getRange(targetMonthly.getLastRow()+1,1,result.monthly.length,result.monthly[0].length).setValues(result.monthly)
  })

}
