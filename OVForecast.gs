function ovForecastUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet){
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var colUsed = colUsed.split(',')
  //check if targetSheet exists in targetUrl
  if(target==null){
  var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)

  }

  var lastUpdateDay = target.getRange(target.getLastRow(),3).getValue();
Logger.log(targetSheet)
Logger.log(lastUpdateDay)
  switch (option)
  {
    case 'DailyToDaily':
    //INSERT IFELSE FOR startRow and colUsed[0] IN CASE SOURCE SHEET HAS UNIQUE FORMAT
      var dataVal = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues().map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1]]})
      var dataBg = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getBackgrounds().map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1]]})
      var data=[]
      if(dataVal[0]!=null){
      for (var i=0; i<dataBg.length;i++){
        if(dataBg[i][1]=='#fce5cd'&&dataVal[i][0]>lastUpdateDay&&dataVal[i][0]<thisMonday){
          data.push([Utilities.formatDate(dataVal[i][0], 'America/Winnipeg', 'MM/dd/yyyy'),Math.round(dataVal[i][1])])
        }
      }
        //Paste to target sheet
  target.getRange(target.getLastRow()+1,3,data.length,data[0].length).setValues(data)
  target.getRange(2,4,data.length,1).setNumberFormat('####')
      }

      break;
    case '1hIntToDaily':
      //INSERT IFELSE FOR startRow and colUsed[0] IN CASE SOURCE SHEET HAS UNIQUE FORMAT
      var dataVal = source.getRange(2,1,source.getLastRow(),source.getLastColumn()).getValues().map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1]]})
      dataVal = dataVal.filter(function(r){return r[0]>=lastUpdateDay&&r[0]<thisMonday})
      var data=[]
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
        //Paste to target sheet
  target.getRange(target.getLastRow()+1,3,data.length,data[0].length).setValues(data)
  target.getRange(2,4,data.length,1).setNumberFormat('####')
      break;
  }
  // Logger.log(data)

return data
}