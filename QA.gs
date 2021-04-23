function qaUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet) {
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var colUsed = colUsed.split(',')
  //if necessary, filter data for last 4 weeks
  var data=[]
  data = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues().map(function(item){return [item[colUsed[0]-1],item[colUsed[1]-1],item[colUsed[2]-1],item[colUsed[3]-1]]})
  .filter(function(r){return typeof(r[2])!='string'&&r[2]!=null})
  Logger.log([targetSheet,data.length])
  //check if targetSheet exists in targetUrl
  if(target==null){
    var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  }
  var result={}
  result.date=[]
  result.totalAssessment=[]
  result.autofail=[]
  result.score=[]
  result.totalScore=[]
  //functions
  if(data.length>0){
    switch (option)
    {
      case 'AgentToTeamWeekly':
        //[[scoreVal,totalScoreVal,date,autofail]]
       data = data.filter(function(r){return timeZone(r[2],'startWeekDate')==timeZone(new Date(thisMonday.getTime() - MILLIS_PER_DAY*7),'startWeekDate')&&typeof r[0]=='number'&&typeof r[1]=='number'})
        data.forEach(function(r){
          if(typeof(r[2])!='string'&&r[2]!=null){
            var index = result.date.indexOf(timeZone(r[2],'startWeekDate'))
            //          Logger.log(index)
            if(index==-1){//if new week
              result.date.push(timeZone(r[2],'startWeekDate'))
              result.totalAssessment.push(1)
              r[3]==1||r[3]=='Auto-Fail'||r[3]=="Yes" ? result.autofail.push(1) : result.autofail.push(0)
              result.score.push(r[0])
              result.totalScore.push(r[1])
            }else{//if week exists
              result.totalAssessment[index]++
              if(r[3]==1||r[3]=='Auto-Fail'||r[3]=="Yes"){ result.autofail[index]++}
              result.score[index]+=r[0]
              result.totalScore[index]+=r[1]
            }
          }//end if r is date
        })
        
        break;
        
      case 'AgentToTeamMonthly':
        //[[scoreVal,totalScoreVal,date,autofail]]
        // Logger.log(data)
        data=data.filter(function(r){return timeZone(r[2],'startMonthDate')==timeZone(new Date(thisMonday.getTime() - MILLIS_PER_DAY*7),'startMonthDate')&&typeof r[0]=='number'&&typeof r[1]=='number'})
        data.forEach(function(r){
          if(typeof(r[2])!='string'&&r[2]!=null){
            var index = result.date.indexOf(timeZone(r[2],'startMonthDate'))
            //          Logger.log(index)
            if(index==-1){//if new week
              result.date.push(timeZone(r[2],'startMonthDate'))
              result.totalAssessment.push(1)
              r[3]==1||r[3]=='Auto-Fail'||r[3]=="Yes" ? result.autofail.push(1) : result.autofail.push(0)
              result.score.push(r[0])
              result.totalScore.push(r[1])
            }else{//if week exists
              result.totalAssessment[index]++
                if(r[3]==1||r[3]=='Auto-Fail'||r[3]=="Yes"){ result.autofail[index]++}
              result.score[index]+=r[0]
              result.totalScore[index]+=r[1]
            }
          }//end if r is date
        })
        break;
    }//end switch
    // Logger.log([sourceSheet,result])
    //Paste to target sheet //paste at rows match 1st entry
    var output=[]
    for(var i=0; i<result.date.length; i++){
      output.push([result.date[i],result.totalAssessment[i],result.autofail[i],result.score[i]/result.totalScore[i]])
    }
    
    if(output[0]!=null){
      var rowPasteOutput = target.getRange(3,1,target.getLastRow()-2).getDisplayValues().map(function(r){return r[0]}).indexOf(output[0][0],4)
      rowPasteOutput>-1 ? target.getRange(rowPasteOutput+3,1,output.length,output[0].length).setValues(output) : target.getRange(target.getLastRow()+1,1,output.length,output[0].length).setValues(output)
    }
    
    
  }//end if data exists
  return result
}

