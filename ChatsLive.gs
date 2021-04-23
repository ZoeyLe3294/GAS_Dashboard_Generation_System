function chatsLiveUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet) {
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var colUsed = colUsed.split(',')
  //check if targetSheet exists in targetUrl
  if(target==null){
    var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
    
  }
  
  var result=[]
  switch (option)
  {
    case 'ReFormat':
      // get data->filter data last week -> map values needed
      if(sourceSheet.indexOf('weekly')>-1){
        var regExp = new RegExp("([0-9]+\/[0-9]+\/[0-9]*)"); 
        var dataVal = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues()
        .filter(function(r){regExp.exec(r[colUsed[0]-1]) !=null})
        .map(function(item){return [regExp.exec(item[colUsed[0]-1])[0],item[colUsed[1]-1],item[colUsed[2]-1],item[colUsed[3]-1],item[colUsed[4]-1],item[colUsed[5]-1],item[colUsed[6]-1],item[colUsed[7]-1],item[colUsed[7]-1]*item[colUsed[1]-1]]})
        }else {
          var dataVal = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues()
          .map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1],item[colUsed[2]-1],item[colUsed[3]-1],item[colUsed[4]-1],item[colUsed[5]-1],item[colUsed[6]-1],item[colUsed[7]-1],item[colUsed[7]-1]*item[colUsed[1]-1]]})
          .filter(function(r){return typeof(r[1])!='string'&&r[1]!=null})
          }
          Logger.log(dataVal)
        //  date	Total Chats	# Served Chats	# Missed Chats	Chat Satisfaction Rate	Average Chat Wait Time	Average Chat Handle Time	SVL # Chat in SVL
     target.getRange(2,1,dataVal.length,dataVal[0].length).setValues(dataVal)
    break;
    case 'DailyToDaily':
      var dataVal = source.getRange(2,1,source.getLastRow()-1,source.getLastColumn()).getValues()
      // .filter(function(r){return r[1]>=new Date(thisMonday.getTime() - MILLIS_PER_DAY*7)})
      var country=[]
      result=dataVal.map(function(item){
        if(country.indexOf(item[colUsed[0]-1])==-1){
          country.push(item[colUsed[0]-1])
        }
        return [item[colUsed[0]-1],item[colUsed[1]-1],item[colUsed[2]-1],item[colUsed[3]-1],item[colUsed[4]-1],item[colUsed[5]-1],item[colUsed[6]-1]/60,item[colUsed[7]-1]]
      })
      target.getRange(target.getLastRow()+1,1,result.length,result[0].length).setValues(result)
    break;
    case 'DailyToWeekly':
        var country=[]
        var dataVal = source.getRange(2,1,source.getLastRow()-1,source.getLastColumn()).getValues().map(function(item){
          if(country.indexOf(item[colUsed[0]-1])==-1){
            country.push(item[colUsed[0]-1])
          }
          return [item[colUsed[0]-1],item[colUsed[1]-1],item[colUsed[2]-1],item[colUsed[3]-1],item[colUsed[4]-1],item[colUsed[5]-1],item[colUsed[6]-1]/60,item[colUsed[7]-1]]
        })
        // .filter(function(r){return r[1]>=new Date(thisMonday.getTime() - MILLIS_PER_DAY*7)})
        Logger.log(country)
        country.forEach(function(country){
          var data = dataVal.filter(function(r){return r[0]==country})
          var week=[]
          data.forEach(function(entry){
            if(week.indexOf(timeZone(entry[1],'startWeekDate'))==-1){
                week.push(timeZone(entry[1],'startWeekDate'))
                result.push([entry[0],timeZone(entry[1],'startWeekDate'),entry[2],entry[3],entry[4],entry[5],entry[6]])
            }else{
                result[week.indexOf(timeZone(entry[1],'startWeekDate'))][2]+=entry[2]
                result[week.indexOf(timeZone(entry[1],'startWeekDate'))][3]+=entry[3]
                result[week.indexOf(timeZone(entry[1],'startWeekDate'))][4]+=entry[4]
                result[week.indexOf(timeZone(entry[1],'startWeekDate'))][5]+=entry[5]
                result[week.indexOf(timeZone(entry[1],'startWeekDate'))][6]+=entry[6]
              }
          })
        })
        //country date missed total <90 afrt aht svl
          result=result.map(function(item){
            Logger.log(item)
            return [item[0],item[1],item[2],item[3],item[4],item[5]/item[3],item[6]/item[3],item[4]/item[3]]
            })
        target.getRange(target.getLastRow()+1,1,result.length,result[0].length).setValues(result)
      break;
      case 'DailyToMonthly':
        var country=[]
        var dataVal = source.getRange(2,1,source.getLastRow()-1,source.getLastColumn()).getValues().map(function(item){
          if(country.indexOf(item[colUsed[0]-1])==-1){
            country.push(item[colUsed[0]-1])
          }
          return [item[colUsed[0]-1],item[colUsed[1]-1],item[colUsed[2]-1],item[colUsed[3]-1],item[colUsed[4]-1],item[colUsed[5]-1],item[colUsed[6]-1]/60,item[colUsed[7]-1]]
        })
        // .filter(function(r){return r[1]>=new Date(thisMonday.getTime() - MILLIS_PER_DAY*7)})
        country.forEach(function(country){
          var data = dataVal.filter(function(r){return r[0]==country})
          var month=[]
          data.forEach(function(entry){
            if(month.indexOf(timeZone(entry[1],'startMonthDate'))==-1){
              month.push(timeZone(entry[1],'startMonthDate'))
              result.push([timeZone(entry[1],'startMonthDate'),entry[1],entry[2],entry[3],entry[4],entry[5],entry[6]])
            }else{
              result[month.indexOf(timeZone(entry[1],'startMonthDate'))][2]+=entry[2]
              result[month.indexOf(timeZone(entry[1],'startMonthDate'))][3]+=entry[3]
              result[month.indexOf(timeZone(entry[1],'startMonthDate'))][4]+=entry[4]
              result[month.indexOf(timeZone(entry[1],'startMonthDate'))][5]+=entry[5]
              result[month.indexOf(timeZone(entry[1],'startMonthDate'))][6]+=entry[6]
            }
          })
        })
        result=result.map(function(item){
          return [item[0],item[1],item[2],item[3],item[4],item[5]/item[3],item[6]/item[3],item[4]/item[3]]
          })
          target.getRange(target.getLastRow()+1,1,result.length,result[0].length).setValues(result)
      break;  
  }

  
}
