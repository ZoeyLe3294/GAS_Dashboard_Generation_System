
function forecastHourUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet){
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var colUsed = colUsed.split(',')
  var last = new Date(thisSunday.getTime() - MILLIS_PER_DAY*7)
  Logger.log(sourceSheet)
  //check if targetSheet exists in targetUrl
  if(target==null){
  var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)

  }
  
  var data=[]
  switch (option)
  {
    case '1hIntToDaily':
        //INSERT IFELSE FOR startRow and colUsed[0] IN CASE SOURCE SHEET HAS UNIQUE FORMAT
//    var startRow = 2
//    if(sourceSheet=='Custo Forecast'){var colUsed[0] = 13}else{var colUsed[0] = 12}

      var dataVal = source.getRange(2,1,source.getLastRow(),source.getLastColumn()).getValues()
                    .map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1]]})
                    .filter(function(r){return r[0]>last})
                    .filter(function(r){return timeZone(r[0],'startWeekDate')== timeZone(thisSunday,'startWeekDate')})
      var date=[]
      dataVal.forEach(function(row){
        if(row[0]!=[] && typeof row[0]!='string'){
          var day = Utilities.formatDate(row[0], 'America/Winnipeg', 'MM/dd/yyyy')
          var dayIndex = date.indexOf(day)
          if(dayIndex==-1){//if new date
            date.push(day)
            data.push([day,row[1]])
          }else{//if date exist
            data[dayIndex][1]+=row[1]
          }
        }
      })
      break;
    case '30minIntToDaily':
      var dataVal = source.getRange(2,1,source.getLastRow(),source.getLastColumn()).getValues()
                    .map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1]]})
                    .filter(function(r){return r[0]>last})
                    .filter(function(r){return timeZone(r[0],'startWeekDate')== timeZone(thisSunday,'startWeekDate')})
      var date=[]
      dataVal.forEach(function(row){
        if(row[0]!=[] && typeof row[0]!='string'){
          var day = Utilities.formatDate(row[0], 'America/Winnipeg', 'MM/dd/yyyy')
          var dayIndex = date.indexOf(day)
          if(dayIndex==-1){//if new date
            date.push(day)
            data.push([day,row[1]])
          }else{//if date exist
            data[dayIndex][1]+=row[1]
          }
        }
      })
      data=data.map(function(r){return [r[0],r[1]/2]})
      break;
  }
Logger.log(data)
  //Paste to target sheet
  // target.getRange(2,3,data.length,data[0].length).setValues(data)
  // target.getRange(2,4,data.length,1).setNumberFormat('####')
  if(data[0]!=undefined){
    target.getRange(target.getLastRow()+1,3,data.length,data[0].length).setValues(data)
    target.getRange(2,4,data.length,1).setNumberFormat('####')
  }
}
