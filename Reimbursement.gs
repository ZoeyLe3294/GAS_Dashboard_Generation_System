function reimbursementUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet) {
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var colUsed = colUsed.split(',')
  //check if targetSheet exists in targetUrl
  if(target==null){
  var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  }
  var dataVal = source.getRange(3,1,source.getLastRow(),source.getLastColumn()).getValues().filter(function(r){return r[0]!=''}).map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1]]})
  dataVal=dataVal.filter(function(r){return timeZone(r[0],'startWeekDate')==timeZone(thisSunday,'startWeekDate')})
  if(dataVal){
    var day=[]
    var output =[]
    for(var i=0;i<dataVal.length;i++){
      var index = day.indexOf(timeZone(dataVal[i][0],3))
      if(dataVal[i][1].toString().indexOf("$")>-1||dataVal[i][1].toString().indexOf("€")>-1){dataVal[i][1]=parseFloat(dataVal[i][1].replace("$"||"€",''))}
      
      if(index==-1){
        day.push(timeZone(dataVal[i][0],3))
        typeof dataVal[i][1]=="number" ? output.push([timeZone(dataVal[i][0],3),1,dataVal[i][1]]):output.push([dataVal[i][0],1,0])
      }else{
        output[index][1]++
        typeof dataVal[i][1]=="number" ? output[index][2]+=dataVal[i][1]:output[index][2]+=0
      }
    }
    if(output[0]!=undefined){target.getRange(target.getLastRow()+1,1,output.length,output[0].length).setValues(output)}
    
  }
  // target.getRange(2,1,output.length,output[0].length).setValues(output)
}