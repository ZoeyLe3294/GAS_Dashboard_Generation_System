function adherenceUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet){
  var teamDrive = DriveApp.getFolderById(sourceUrl);
  var teamName = sourceSheet
  var currentWeek = new Date(today.getTime() - MILLIS_PER_DAY*7)
  var lastWeek = new Date(today.getTime() - MILLIS_PER_DAY*14)
  var currentMonth = new Date(today.getFullYear(), today.getMonth(), 1)
  if(today.getMonth()==1){var lastMonth =new Date(today.getFullYear()-1, 12, 1)}else{ var lastMonth = new Date(today.getFullYear(), today.getMonth()-1, 1)}
  var files = teamDrive.getFiles();
  var result=[]
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  //check if targetSheet exists in targetUrl
  if(target==null){
    var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  }
  while(files.hasNext()){
    var currentFile = files.next()
    var fileID = currentFile.getId();
    Logger.log(currentFile.getName(),currentMonth,lastMonth,Utilities.formatDate(lastMonth, 'America/Winnipeg', 'MMMM'),currentFile.getName().indexOf(Utilities.formatDate(lastMonth, 'America/Winnipeg', 'MMMM')))
    if(currentFile.getName().indexOf('Week '+Math.floor(timeZone(currentWeek,1)))>-1){
      var sheet = SpreadsheetApp.openById(fileID).getSheetByName(teamName);
      if(sheet!=null){
        var data = sheet.getRange(1,5,1,5).getValues()[0]
        data=data.map(function(r){return [r[0],r[2],r[3],0,r[4]]})
        data.unshift(timeZone(currentWeek,'startWeekDate'))
        data.unshift(sourceSheet)
        result.push(data)
      }
    // }else if(currentFile.getName().indexOf('Week '+Math.floor(timeZone(lastWeek,1)))>-1){
    //   var sheet = SpreadsheetApp.openById(fileID).getSheetByName(teamName);
    //   if(sheet!=null){
    //     var data = sheet.getRange(1,5,1,5).getValues()[0]
    //     data.unshift(timeZone(lastWeek,'startWeekDate'))
    //     data.unshift(sourceSheet)
    //     result.push(data)
    //   }
    }else if(currentFile.getName().indexOf(Utilities.formatDate(currentMonth, 'America/Winnipeg', 'MMMM'))>-1){
    var sheet = SpreadsheetApp.openById(fileID).getSheetByName(teamName);
      if(sheet!=null){
        var data = sheet.getRange(3,5,1,5).getValues()[0]
        data.unshift(currentMonth)
        data.unshift(sourceSheet)
        result.push(data)
      }
    }
    // }else if(currentFile.getName().indexOf(Utilities.formatDate(lastMonth, 'America/Winnipeg', 'MMMM'))>-1){
    // var sheet = SpreadsheetApp.openById(fileID).getSheetByName(teamName);
    //   if(sheet!=null){
    //     var data = sheet.getRange(3,5,1,5).getValues()[0]
    //     data.unshift(lastMonth)
    //     data.unshift(sourceSheet)
    //     result.push(data)
    //   }
    // }
  }
  
  Logger.log(sourceSheet,result)
  if(result[0]!=null){
    //    var rowPasteOutput = target.getRange(1,1,target.getLastRow()).getValues().map(function(r){return r[0]}).indexOf(result[0][0])+1
    //    rowPasteOutput>0 ? target.getRange(rowPasteOutput,1,result.length,result[0].length).setValues(result) : target.getRange(2,1,result.length,result[0].length).setValues(result)
    target.getRange(target.getLastRow()+1,1,result.length,result[0].length).setValues(result)
  }
}