////GLOBAL PARAMETERS OF CONFIG PAGE >>> 1 = ColA or Row1
//var config = 'Forecasting Info'
//var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config)
//var dashboardId = '1SZq_A1iUAZZUxq486CcjLzpCXrlFeT7IK69nZtv0fus'
//var MILLIS_PER_DAY = 1000*60*60*24
//var today = new Date()
//var startReportPeriod = new Date(today.getTime() - MILLIS_PER_DAY*30)
//
////Back-end Parameters
//var startColIp = 3
//var startColOp = 9
//var lastColB = 16
//var sourceUrl = 4 -startColIp 
//var sourceSheet = 5 -startColIp
//var colUsed = 6 -startColIp
//var option = 7 -startColIp
//var outputName = 9 -startColIp
//var targetUrl = 10 -startColIp
//var targetSheet = 11 -startColIp
//var converter = 12 -startColIp
//var opLCol = 14 //Column of OUTPUT LIST cell
//var opLRow = 4 //Row of OUTPUT LIST cell
//var startRowB = 5
////Front-end Parameters
//var startColF = 18
//var startRowF = 4
//
////PARAMETERS REF
////Source file,	URL,	Sheet name,	Column # used,	Granularity,	Output file,	URL,	Output sheet name,	Granularity
////ovForecastUpdate(sourceUrl,sourceSheet,dateCol,colUsed,option,targetUrl,targetSheet)
////forecastHourUpdate(sourceUrl,sourceSheet,dateCol,colUsed,option,targetUrl,targetSheet)
//
//
////Frontend MAIN FUNCTION
//function runF(){
//  //Dashboard setup
//  var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Daily Dashboard')
//  var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Weekly Dashboard')
//  var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Monthly Dashboard')
//  //  var dailyCol = dailyDash.getRange('2:2').getValues().map(function(r){return r[0]})
//  //  var weeklyCol = weeklyDash.getRange('2:2').getValues().map(function(r){return r[0]})
//  //  var monthlyCol = monthlyDash.getRange('2:2').getValues().map(function(r){return r[0]})
//  var dailyCol = 7
//  var weeklyCol = 7
//  var monthlyCol = 7
//  dailyDash.getRange(3,1,dailyDash.getLastRow()-2,dailyDash.getLastColumn()).clearContent()
//  weeklyDash.getRange(4,1,weeklyDash.getLastRow()-3,weeklyDash.getLastColumn()).clearContent()
//  monthlyDash.getRange(4,1,monthlyDash.getLastRow()-3,monthlyDash.getLastColumn()).clearContent()
//  
//  //Config info
//  var configData = configSheet.getRange(startRowF,startColF).getDataRegion().getValues() //0=metric, 1=formula, 2=data file, 3=attributes
//  var title = configSheet.getRange(startRowF,startColF,1,configSheet.getLastColumn()+1-startColF).getValues()[0]
//  var dataValidationList = configSheet.getRange(opLRow,opLCol).getDataRegion().getValues()
//  var nameL = dataValidationList.map(function(e){return e[0]})
//  var urlL = dataValidationList.map(function(e){return e[1]})
//  nameL.shift(); urlL.shift()
//  //Result setup  
//  var dailyResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
//  dailyResult.curCountry
//  dailyResult.curCounter = dailyDash.getLastRow()+1
//  dailyResult.days=[] 
//  for(var l=29;l>=0;l--){
//    dailyResult.days.push(timeZone(new Date(today.getTime() - MILLIS_PER_DAY*l)))
//  }
//  var weekResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
//  weekResult.curCountry
//  weekResult.curCounter = weeklyDash.getLastRow()+1
//  weekResult.days=[] 
//  for(var l=3;l>=0;l--){
//    weekResult.days.push(timeZone(new Date(today.getTime() - MILLIS_PER_DAY*7*l),'startWeekDate'))
//  }
//  var monthResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
//  monthResult.curCountry
//  monthResult.curCounter = monthlyDash.getLastRow()+1
//  monthResult.days=[] 
//  if(today.getMonth()==1){ monthResult.days=[timeZone(new Date(today.getFullYear()-1, 12, 1),3),timeZone(new Date(today.getFullYear(), today.getMonth(), 1),3)]
//                         }else{ monthResult.days=[timeZone(new Date(today.getFullYear(), today.getMonth()-1, 1),3),timeZone(new Date(today.getFullYear(), today.getMonth(), 1),3)]}
//  
//  for (var i=0; i<title.length;i++){//loop team
//    if(title[i]=='Applicable'){//if col is Applicable
//      var team = configSheet.getRange(startRowF-2,startColF+i).getValue().split('|')
//      
//      dailyResult.metric=[]
//      dailyResult.value=[]
//      weekResult.metric=[]
//      weekResult.value=[]
//      monthResult.metric=[]
//      monthResult.value=[]
//      
//      if(dailyResult.curCountry!=team[0]){//if country not exists ->print country to sheets
//        dailyResult.curCountry=team[0]
//        weekResult.curCountry=team[0]
//        monthResult.curCountry=team[0]
//        dailyDash.getRange(dailyResult.curCounter,2).setValue(dailyResult.curCountry)
//        weeklyDash.getRange(weekResult.curCounter,2).setValue(weekResult.curCountry)
//        monthlyDash.getRange(monthResult.curCounter,2).setValue(monthResult.curCountry)
//        dailyResult.curCounter++
//          weekResult.curCounter++
//            monthResult.curCounter++
//      }
//      //print team in next row
//      dailyDash.getRange(dailyResult.curCounter,2).setValue(team[1])
//      weeklyDash.getRange(weekResult.curCounter,2).setValue(team[1])
//      monthlyDash.getRange(monthResult.curCounter,2).setValue(team[1])
//      
//      //      Logger.log(team,dailyResult.days,weekResult.days,monthResult.days)
//      
//      var applicableList = configSheet.getRange(startRowF,startColF+i,configData.length,2).getValues() //[dash type, filters(indicator = ',')]
//      
//      //loop metrics
//      for(var j=0;j<applicableList.length;j++){
//        //split filter 
//        var filters = applicableList[j][1].split(',')
//        
//        //if format checked
//        if(configData[j][1]==true){
//          dailyResult.metric.push([configData[j][0]])
//          weekResult.metric.push([configData[j][0]])
//          monthResult.metric.push([configData[j][0]])
//          var res=[]
//          dailyResult.days.forEach(function(r){
//            res.push('')
//          })
//          dailyResult.value.push(res)
//          var res=[]
//          weekResult.days.forEach(function(r){
//            res.push('')
//          })
//          weekResult.value.push(res)
//          var res=[]
//          monthResult.days.forEach(function(r){
//            res.push('')
//          })
//          monthResult.value.push(res)
//          
//        }
//        //DAILY
//        if(applicableList[j][0]=='Daily'){
//          
//          //push metrics || configData[[0=metric, 1=formula, 2=data file, 3=attributes]]
//          dailyResult.metric.push([configData[j][0]])
//          weekResult.metric.push([configData[j][0]])
//          monthResult.metric.push([configData[j][0]])
//          
//          //address file,sheet,col to use in data
//          if(filters.length>=1){
//            var file = SpreadsheetApp.openByUrl(urlL[nameL.indexOf(configData[j][2])]).getSheetByName(filters[0].split(':')[1])
//            var data = file.getRange(1,1).getDataRegion().getValues()
//            if(filters.length>2){//if more filter add in beside sheet name
//              for(var k=1;k<filters.length;k++){
//                data=data.filter(function(r){return filters[k].split(':')[0].indexOf(r[data[0].indexOf(filters[k].split(':')[0])])>-1 || filters[k].split(':')[1].indexOf(r[data[0].indexOf(filters[k].split(':')[0])])>-1})
//              }
//            }
//            //grab data from database
//            var curDat=[]
//            var operator=''
//            var lastIndex=0
//            var attributes = configData[j][3]
//            Logger.log(team)
//            for(var m=0;m<=configData[j][3].length;m++){
//              if(attributes[m]=='/'||attributes[m]=='-'||attributes[m]=='*'||attributes[m]=='+'||m==attributes.length){
//                var dailyVal =[]
//                dailyResult.days.forEach(function(r){
//                  dailyVal.push(0)
//                })
//                var weeklyVal =[]
//                weekResult.days.forEach(function(r){
//                  weeklyVal.push(0)
//                })
//                var monthlyVal =[]
//                monthResult.days.forEach(function(r){
//                  monthlyVal.push(0)
//                })
//                switch(operator){
//                  case '':
//                    curDat=dataGrab(data,attributes.substring(lastIndex,m),dailyResult.days,weekResult.days,monthResult.days,[dailyVal,weeklyVal,monthlyVal],'Daily')
//                    operator=attributes[m]
//                    lastIndex=m+1
//                    break;
//                  case '/':
//                    var nextDat = dataGrab(data,attributes.substring(lastIndex,m),dailyResult.days,weekResult.days,monthResult.days,[dailyVal,weeklyVal,monthlyVal],'Daily')
//                    Logger.log(curDat,nextDat,attributes.substring(lastIndex,m),operator)
//                    curDat=calculateArrays(curDat,'/',nextDat)
//                    Logger.log(curDat)
//                    operator=attributes[m]
//                    lastIndex=m+1
//                    break;
//                  case '+':
//                    curDat=calculateArrays(curDat,'+',dataGrab(data,attributes.substring(lastIndex,m),dailyResult.days,weekResult.days,monthResult.days,[dailyVal,weeklyVal,monthlyVal],'Daily'))
//                    operator=attributes[m]
//                    lastIndex=m+1
//                    break;
//                  case '-':
//                    curDat=calculateArrays(curDat,'-',dataGrab(data,attributes.substring(lastIndex,m),dailyResult.days,weekResult.days,monthResult.days,[dailyVal,weeklyVal,monthlyVal],'Daily'))
//                    operator=attributes[m]
//                    lastIndex=m+1
//                    break;
//                  case '*':
//                    curDat=calculateArrays(curDat,'*',dataGrab(data,attributes.substring(lastIndex,m),dailyResult.days,weekResult.days,monthResult.days,[dailyVal,weeklyVal,monthlyVal],'Daily'))
//                    operator=attributes[m]
//                    lastIndex=m+1
//                    break;
//                }
//              }
//            }
//            //                                Logger.log(curDat)
//            //            curDat=dataGrab(data,configData[j][3].substring(lastIndex,configData[j][3].length),dailyResult.days,weekResult.days,monthResult.days,[dailyVal,weeklyVal,monthlyVal],'Daily')
//            dailyVal=curDat[0]
//            weeklyVal=curDat[1]
//            monthlyVal=curDat[2]
//            
//          }//if filter exists
//          //push val to result
//          dailyResult.value.push(dailyVal)
//          weekResult.value.push(weeklyVal)
//          monthResult.value.push(monthlyVal)
//        }//if daily
//        
//        //WEEKLY+MONTH
//        else {
//          var dashType = applicableList[j][0].split(',')
//          dashType.forEach(function(type){
//            Logger.log(type)
//            switch (type){
//              case 'Weekly':
//                //push metrics || configData[[0=metric, 1=formula, 2=data file, 3=attributes]]
//                weekResult.metric.push([configData[j][0]])
//                if(filters.length>=1){
//                  var file = SpreadsheetApp.openByUrl(urlL[nameL.indexOf(configData[j][2])]).getSheetByName(filters[0].split(':')[1].toString().replace("*", "Weekly"))
//                  var data = file.getRange(1,1).getDataRegion().getValues()
//                  if(filters.length>=2){//if more filter add in beside sheet name
//                    for(var k=1;k<filters.length;k++){
//                      data=data.filter(function(r){return filters[k].split(':')[0].indexOf(r[data[0].indexOf(filters[k].split(':')[0])])>-1 || filters[k].split(':')[1].indexOf(r[data[0].indexOf(filters[k].split(':')[0])])>-1})
//                    }
//                  }
//                  var weeklyVal =[]
//                  weekResult.days.forEach(function(r){
//                    weeklyVal.push(0)
//                  })
//                  var curDat=[]
//                  var operator=''
//                  var lastIndex=0
//                  curDat=dataGrab(data,configData[j][3].substring(lastIndex,configData[j][3].length),dailyResult.days,weekResult.days,monthResult.days,[weeklyVal],'Weekly')
//                  
//                  weeklyVal=curDat[0]
//                  weekResult.value.push(weeklyVal)
//                }
//                break;
//              case 'Monthly':
//                monthResult.metric.push([configData[j][0]])
//                if(filters.length>=1){
//                  var file = SpreadsheetApp.openByUrl(urlL[nameL.indexOf(configData[j][2])]).getSheetByName(filters[0].split(':')[1].toString().replace("*", "Monthly"))
//                  var data = file.getRange(1,1).getDataRegion().getValues()
//                  if(filters.length>2){//if more filter add in beside sheet name
//                    for(var k=1;k<filters.length;k++){
//                      data=data.filter(function(r){return filters[k].split(':')[0].indexOf(r[data[0].indexOf(filters[k].split(':')[0])])>-1 || filters[k].split(':')[1].indexOf(r[data[0].indexOf(filters[k].split(':')[0])])>-1})
//                    }
//                  }
//                  var monthlyVal =[]
//                  monthResult.days.forEach(function(r){
//                    monthlyVal.push(0)
//                  })
//                  var curDat=[]
//                  var operator=''
//                  var lastIndex=0
//                  curDat=dataGrab(data,configData[j][3].substring(lastIndex,configData[j][3].length),dailyResult.days,weekResult.days,monthResult.days,[monthlyVal],'Monthly')
//                  
//                  monthlyVal=curDat[0]
//                  monthResult.value.push(monthlyVal)
//                }
//                break;
//            }
//          })
//          
//        }// if weekly+monthly
//        
//      }//applicable loop
//      Logger.log(team,dailyResult,weekResult,monthResult)
//      //paste result to Daily
//      if(dailyResult.metric[0]!=null){
//        dailyDash.getRange(2,dailyCol,1,dailyResult.days.length).setValues([dailyResult.days]).setNumberFormat('MM/DD/YYYY')
//        dailyDash.getRange(dailyResult.curCounter,3,dailyResult.metric.length,dailyResult.metric[0].length).setValues(dailyResult.metric)
//        dailyDash.getRange(dailyResult.curCounter,dailyCol,dailyResult.value.length,dailyResult.value[0].length).setValues(dailyResult.value)
//        dailyResult.curCounter+=dailyResult.metric.length
//      }else{
//        dailyResult.curCounter++
//      }
//      //paste result to Daily
//      if(weekResult.metric[0]!=null){
//        weeklyDash.getRange(2,weeklyCol,1,weekResult.days.length).setValues([weekResult.days]).setNumberFormat('MM/DD/YYYY')
//        weeklyDash.getRange(weekResult.curCounter,3,weekResult.metric.length,weekResult.metric[0].length).setValues(weekResult.metric)
//        weeklyDash.getRange(weekResult.curCounter,weeklyCol,weekResult.value.length,weekResult.value[0].length).setValues(weekResult.value)
//        weekResult.curCounter+=weekResult.metric.length
//      }else{
//        weekResult.curCounter++
//      }
//      //paste result to Daily
//      if(monthResult.metric[0]!=null){
//        monthlyDash.getRange(2,monthlyCol,1,monthResult.days.length).setValues([monthResult.days]).setNumberFormat('MMMM, YYYY')
//        monthlyDash.getRange(monthResult.curCounter,3,monthResult.metric.length,monthResult.metric[0].length).setValues(monthResult.metric)
//        monthlyDash.getRange(monthResult.curCounter,monthlyCol,monthResult.value.length,monthResult.value[0].length).setValues(monthResult.value)
//        monthResult.curCounter+=monthResult.metric.length
//      }else{
//        monthResult.curCounter++
//      }
//      
//    }//loop team
//    
//  }
//  //  dashboardSheet.getRange(dashboardSheet.getLastRow()+1,4,r.length,r[0].length).setValues(r)
//}
//
////Backend MAIN FUNCTION
//function runB() {
//  //get config + filter config that has output file
//  var configData = configSheet.getRange(startRowB,startColIp,configSheet.getLastRow()-startRowB-1,lastColB-startColIp).getValues()
//  configData = configData.filter(function(row){return row[targetUrl].indexOf('https://docs.google.com')>-1})
//  //  Logger.log(configData)
//  //loop through config row to run appropriate data process function
//  configData.forEach(function(row){
//    switch (row[outputName])
//    {
//        //      case 'OV Forecast': ovForecastUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]); break;
//        //      case 'Forecasted Hours': forecastHourUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]); break;
//        //      case 'QA Team Level': qaUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
//      case 'Adherence': adherenceUpdate(row[sourceUrl],row[sourceSheet],row[colUsed],row[converter],row[targetUrl],row[targetSheet]);break;
//        //ADD MORE OUTPUT FUNCTION HERE
//    }
//  })
//}


//
//
//
//
////grab data
//function dataGrab(data,colName,dayRange,weekRange,monthRange,vals,dashType){
//  switch(dashType)
//  {
//    case 'Daily':
//      var dailyVal = vals[0]
//      var weeklyVal = vals[1]
//      var monthlyVal = vals[2]
//      var dayIndex = data[0].indexOf('Date')>-1 ? data[0].indexOf('Date') : data[0].indexOf('date')
//      data.forEach(function(r){ //values of 30 days for metric j of team i
//        if(r[dayIndex]>=startReportPeriod && r[dayIndex]<=today){
//          //daily data
//          dailyVal[dayRange.indexOf(timeZone(r[dayIndex]))]+=r[data[0].indexOf(colName)]
//          //weekly data
//          if(weekRange.indexOf(timeZone(r[dayIndex],'startWeekDate'))>-1){
//            weeklyVal[weekRange.indexOf(timeZone(r[dayIndex],'startWeekDate'))]+=r[data[0].indexOf(colName)]
//          }            
//          //monthly data
//          if(monthRange.indexOf(timeZone(r[dayIndex],'startMonthDate'))>-1){
//            monthlyVal[monthRange.indexOf(timeZone(r[dayIndex],'startMonthDate',4))]+=r[data[0].indexOf(colName)]
//          }
//        }
//      })
//      
//      return [dailyVal,weeklyVal,monthlyVal]
//      break;
//    case 'Weekly':
//      var weeklyVal=vals[0]
//      var dayIndex = data[0].indexOf('Date')>-1 ? data[0].indexOf('Date') : data[0].indexOf('date')
//      data.forEach(function(r){ //values of 30 days for metric j of team i
//        if(r[dayIndex]>=startReportPeriod && r[dayIndex]<=today){
//          //weekly data
//          if(weekRange.indexOf(timeZone(r[dayIndex],'startWeekDate'))>-1){
//            weeklyVal[weekRange.indexOf(timeZone(r[dayIndex],'startWeekDate'))]=r[data[0].indexOf(colName)]
//          }
//        }
//      })
//      return [weeklyVal]
//      break;
//    case 'Monthly':
//      var monthlyVal=vals[0]
//      var dayIndex = data[0].indexOf('Date')>-1 ? data[0].indexOf('Date') : data[0].indexOf('date')
//      data.forEach(function(r){ //values of 30 days for metric j of team i
//        if(r[dayIndex]>=startReportPeriod && r[dayIndex]<=today){
//          //monthly data
//          if(monthRange.indexOf(timeZone(r[dayIndex],'startMonthDate'))>-1){
//            monthlyVal[monthRange.indexOf(timeZone(r[dayIndex],'startMonthDate',4))]=r[data[0].indexOf(colName)]
//          }  
//        }
//      })
//      return [monthlyVal]
//      break;
//  }
//  
//}























