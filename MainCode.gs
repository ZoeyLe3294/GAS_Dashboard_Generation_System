//Toggle var test to change location
var test = true; //perform testing, location is test file
// var test = false; //not perform testing, location is official file

//GLOBAL PARAMETERS OF CONFIG PAGE >>> 1 = ColA or Row1
var setupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SET UP')
var mapSourceId ='1OgPzxgK6Adrt5xtVu2VmzgzweUCJUzhClje9Ss4VM6k'

var MILLIS_PER_DAY = 1000*60*60*24
var today=new Date()
var thisMonday = new Date(today.getTime() - MILLIS_PER_DAY*(((today.getDay() + 6) % 7) + 1-1))
var thisSunday = new Date(thisMonday.getTime() - MILLIS_PER_DAY)
var dateBeginReport = new Date(2020, 11, 1)
var startReportPeriod = dateBeginReport
var endReportPeriod = thisSunday
//var lastXdays = 7
//var startReportPeriod = new Date(thisSunday.getTime() - MILLIS_PER_DAY*lastXdays)
// var dateBeginReport = new Date(2020, 9, 1)
// var startReportPeriod = dateBeginReport
// var endReportPeriod = new Date(2020, 11, 1)

//Back-end Parameters
var startColIp = 3
var startColOp = 9
var lastColB = 16
var sourceUrl = 4 -startColIp 
var sourceSheet = 5 -startColIp
var colUsed = 6 -startColIp
var option = 7 -startColIp
var outputName = 9 -startColIp
var targetUrl = 10 -startColIp
var targetSheet = 11 -startColIp
var converter = 12 -startColIp
var opLCol = 1 //Column of OUTPUT LIST cell
var opLRow = 4 //Row of OUTPUT LIST cell
var startRowB = 5
//Front-end Parameters
// var startColF = 18
var startColF = 6
var startRowF = 4

//format
var colorDefault = "white"
var jeColor = "#ff8000"
var countryBg = "#0a3847"
var teamBg = "black"

//Frontend MAIN FUNCTION
//Update all --- use when metrics update, formulas update, format update
function runF(config,dailyDash,weeklyDash,monthlyDash,divisionPart){
  //Dashboard setup
  // var configSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1mefNhXK8z0jQI2EuQ7X8T7__V74MEuBiVYv2WsgV3Eo/edit#gid=695943687").getSheetByName(config)
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config)
  //  var dailyCol = dailyDash.getRange('2:2').getValues().map(function(r){return r[0]})
  //  var weeklyCol = weeklyDash.getRange('2:2').getValues().map(function(r){return r[0]})
  //  var monthlyCol = monthlyDash.getRange('2:2').getValues().map(function(r){return r[0]})
  var dailyCol = 7
  var weeklyCol = 7
  var monthlyCol = monthlyDash.getName()=='Monthly CO'||monthlyDash.getName()=='Monthly CS'||monthlyDash.getName()=='Monthly CC'||monthlyDash.getName()=='Monthly Dashboard'||monthlyDash.getName()=='Monthly' ? 9 : 7
  
  //  dailyDash.getRange(3,1,dailyDash.getLastRow()-3,dailyDash.getLastColumn()).clearContent()
  //  weeklyDash.getRange(4,1,weeklyDash.getLastRow()-4,weeklyDash.getLastColumn()).clearContent()
  //  monthlyDash.getRange(4,1,monthlyDash.getLastRow()-4,monthlyDash.getLastColumn()).clearContent()
  if(!divisionPart){
    if(dailyDash){dailyDash.deleteRows(3, dailyDash.getLastRow()-2);dailyDash.clearConditionalFormatRules()}
    if(weeklyDash){weeklyDash.deleteRows(4, weeklyDash.getLastRow()-3);weeklyDash.clearConditionalFormatRules()}
    if(monthlyDash){monthlyDash.deleteRows(4,monthlyDash.getLastRow()-3);monthlyDash.clearConditionalFormatRules()}
  }


  //Config info
  var configData = configSheet.getRange(startRowF,startColF).getDataRegion().getValues() //0=metric, 1=formula, 2=data file, 3=attributes
  var title = configSheet.getRange(startRowF,startColF,1,configSheet.getLastColumn()+1-startColF).getValues()[0]
  var dataValidationList = configSheet.getRange(opLRow,opLCol).getDataRegion().getValues()
  dataValidationList.shift()
  var nameL = dataValidationList.map(function(e){return e[0]})
  var urlL = dataValidationList.map(function(e){return SpreadsheetApp.openByUrl(e[1])})
  // nameL.shift(); urlL.shift()
  //Result setup  
  var dailyResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
  dailyResult.curCountry
  dailyResult.curCounter = dailyDash ? dailyDash.getLastRow()+1 : 4
  dailyResult.days=[] 
  var weekResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
  weekResult.curCountry
  weekResult.curCounter = weeklyDash ? weeklyDash.getLastRow()+1 : 5
  weekResult.days=[] 
  var monthResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
  monthResult.curCountry
  monthResult.curCounter = monthlyDash ? monthlyDash.getLastRow()+1 : 5
  monthResult.days=[] 
  //setup length of dayRange,weekRange,monthRange
  var day=startReportPeriod
  while(day<=endReportPeriod){
    dailyResult.days.push(timeZone(day))
    if(weekResult.days.indexOf(timeZone(day,'startWeekDate'))==-1){ weekResult.days.push(timeZone(day,'startWeekDate'))}
    if(monthResult.days.indexOf(timeZone(day,'startMonthDate'))==-1){ monthResult.days.push(timeZone(day,'startMonthDate'))}
    day = new Date(day.getTime() + MILLIS_PER_DAY)
  }
  weekResult.days.shift()
  for (var i=0; i<title.length;i++){//loop team

    var partition = divisionPart ? title[i]=='Applicable' && configSheet.getRange(startRowF-2,startColF+i).getValue().split('|')[0].indexOf(divisionPart)>-1 : title[i]=='Applicable'
    if(partition){//if col is Applicable
      var team = configSheet.getRange(startRowF-2,startColF+i).getValue().split('|')
      
      dailyResult.metric=[]
      dailyResult.value=[]
      dailyResult.target=[]
      weekResult.metric=[]
      weekResult.value=[]
      weekResult.target=[]
      monthResult.metric=[]
      monthResult.value=[]
      monthResult.target=[]
      
      if(dailyResult.curCountry!=team[0]){//if country not exists ->print country to sheets
        dailyResult.curCountry=team[0]
        weekResult.curCountry=team[0]
        monthResult.curCountry=team[0]
        if(dailyDash){
          dailyDash.getRange(dailyResult.curCounter,2).setValue(dailyResult.curCountry)
          dailyDash.getRange(dailyResult.curCounter+":"+dailyResult.curCounter).setBackground(countryBg).setFontWeight("bold").setFontSize(16).setFontFamily("Poppins").setFontColor(colorDefault)
        }
        if(weeklyDash){
          weeklyDash.getRange(weekResult.curCounter,2).setValue(weekResult.curCountry)
          weeklyDash.getRange(weekResult.curCounter+":"+weekResult.curCounter).setBackground(countryBg).setFontWeight("bold").setFontSize(16).setFontFamily("Poppins").setFontColor(colorDefault)
        }
        if(monthlyDash){
          monthlyDash.getRange(monthResult.curCounter,2).setValue(monthResult.curCountry)
          monthlyDash.getRange(monthResult.curCounter+":"+monthResult.curCounter).setBackground(countryBg).setFontWeight("bold").setFontSize(16).setFontFamily("Poppins").setFontColor(colorDefault)
        }
        
        dailyResult.curCounter++;
        weekResult.curCounter++;
        monthResult.curCounter++;
      }
      //print team in next row
      if(dailyDash){dailyDash.getRange(dailyResult.curCounter,2).setValue(team[1]).setFontSize(12).setFontFamily("Poppins").setFontColor(colorDefault).setBackground(teamBg)}
      if(weeklyDash){weeklyDash.getRange(weekResult.curCounter,2).setValue(team[1]).setFontSize(12).setFontFamily("Poppins").setFontColor(colorDefault).setBackground(teamBg)}
      if(monthlyDash){monthlyDash.getRange(monthResult.curCounter,2).setValue(team[1]).setFontSize(12).setFontFamily("Poppins").setFontColor(colorDefault).setBackground(teamBg)}
      
      if(team[1]!=''){
        dailyResult.curCounter++;
        weekResult.curCounter++;
        monthResult.curCounter++;
      }
      
      var applicableList = configSheet.getRange(startRowF,startColF+i,configData.length,3).getValues() //[dash type, filters(indicator = ',')]
      
      //loop metrics
      for(var j=0;j<applicableList.length;j++){
        
        //split filter 
        var filters = applicableList[j][1].split('&')
        //Filter aggregation 'All'
        if(applicableList[j][0]=='All'){
          // Logger.log([team[0]+" "+team[1],configData[j][0]])
          //if no filters -> include metrics in the list || if filters exist -> push metrics info
          if(filters[0]==''){
            dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            //push target of metrics
            dailyResult.target.push([applicableList[j][2]])
            weekResult.target.push([applicableList[j][2]])
            monthResult.target.push([applicableList[j][2]])
            var res=[]
            dailyResult.days.forEach(function(r){
              res.push('')
            })
            dailyResult.value.push(res)
            var res=[]
            weekResult.days.forEach(function(r){
              res.push('')
            })
            weekResult.value.push(res)
            var res=[]
            monthResult.days.forEach(function(r){
              res.push('')
            })
            monthResult.value.push(res)
            
          }else{
            //push metrics || configData[[0=metric, 1=formula, 2=data file, 3=attributes]]
            dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            //push target of metrics
            dailyResult.target.push([applicableList[j][2]])
            weekResult.target.push([applicableList[j][2]])
            monthResult.target.push([applicableList[j][2]])
            Logger.log([team[0]+" "+team[1],configData[j][0]])
            //grab data from sheets
            var data =[]
            filters.forEach(function(filter){
              var subFilter = filter.split(',')
              var subData = []
              subFilter[0].split(':')[1].split('|').forEach(function(filename){
                var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename)
                subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
              })
              //apply filters to data
              data = data.concat(filterCustom(subData,subFilter))
            })

            //grab data from database
            var curDat=[]
            curDat=attributeOperation(data,configData[j][3],[dailyResult.days,weekResult.days,monthResult.days],applicableList[j][0])

            //push val to result
            dailyResult.value.push(curDat[0])
            weekResult.value.push(curDat[1])
            monthResult.value.push(curDat[2])
          }
          
        }else {//Filter aggregation 'Daily,Weekly,Monthly'

          var dashType = applicableList[j][0].split(',')
          dashType.forEach(function(type){//loop dash type
            switch (type){
              case 'Daily':
              Logger.log([team[0]+" "+team[1],configData[j][0]])
                if(filters[0]==''){
                  dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  //push target of metrics
                  dailyResult.target.push([applicableList[j][2]])
                  
                  var res=[]
                  dailyResult.days.forEach(function(r){
                    res.push('')
                  })
                  dailyResult.value.push(res)
                }else{
                  dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  //push target of metrics
                  dailyResult.target.push([applicableList[j][2]])

                  //grab data from sheets
                  var data =[]
                  filters.forEach(function(filter){
                    var subFilter = filter.split(',')
                    var subData = []
                    subFilter[0].split(':')[1].split('|').forEach(function(filename){
                      var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename.toString().replace('*','Daily'))
                      subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
                    })
                    //apply filters to data
                    data = data.concat(filterCustom(subData,subFilter))
                  })

                  //grab data from database
                  var curDat=[]
                  curDat=attributeOperation(data,configData[j][3],[dailyResult.days],type)

                  //push val to result
                  dailyResult.value.push(curDat[0])                  
                }
                break;
              case 'Weekly':
              Logger.log([team[0]+" "+team[1],configData[j][0]])
                if(filters[0]==''){
                  weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  //push target of metrics
                  weekResult.target.push([applicableList[j][2]])
                  var res=[]
                  weekResult.days.forEach(function(r){
                    res.push('')
                  })  
                  weekResult.value.push(res)
                }else{
                  weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  //push target of metrics
                  weekResult.target.push([applicableList[j][2]])
                  var data =[]
                  filters.forEach(function(filter){
                    var subFilter = filter.split(',')
                    var subData = []
                    subFilter[0].split(':')[1].split('|').forEach(function(filename){
                      var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename.toString().replace('*','Weekly'))
                      subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
                    })
                    //apply filters to data
                    data = data.concat(filterCustom(subData,subFilter))
                  })

                  //grab data from database
                  var curDat=[]
                  curDat=attributeOperation(data,configData[j][3],[weekResult.days],type)

                  //push val to result
                  weekResult.value.push(curDat[0])                  
                }
                break;
              case 'Monthly':
              // Logger.log([team[0]+" "+team[1],configData[j][0]])
                if(filters[0]==''){
                  monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  monthResult.target.push([applicableList[j][2]])
                  var res=[]
                  monthResult.days.forEach(function(r){
                    res.push('')
                  })  
                  monthResult.value.push(res)
                }else{
                  monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  monthResult.target.push([applicableList[j][2]])

                  var data =[]
                  filters.forEach(function(filter){
                    var subFilter = filter.split(',')
                    var subData = []
                    subFilter[0].split(':')[1].split('|').forEach(function(filename){
                      var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename.toString().replace('*','Monthly'))
                      subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
                    })
                    //apply filters to data
                    data = data.concat(filterCustom(subData,subFilter))
                  })

                  //grab data from database
                  var curDat=[]
                  curDat=attributeOperation(data,configData[j][3],[monthResult.days],type)

                  //push val to result
                  monthResult.value.push(curDat[0])                  
                }
                break;
            }
          })
          
        }// if daily,weekly,monthly
      }//end loop metrics

      //PASTE RESULT TO SHEET
      var metricList = configData.map(function(c){return c[0]})
      
      if(dailyResult.metric[0]!=null&&dailyDash){
        //run formula loop
        for(var m=0;m<dailyResult.metric.length;m++){
          var formula = configData[metricList.indexOf(dailyResult.metric[m][1])]
          if(formula[1]!=[]){
            var metricDeArray = dailyResult.metric.map(function(item){return item[1]})
            var formulaResult = formulaLoop(formula[1],metricDeArray,dailyResult.value)
            if(formulaResult!=null){dailyResult.value[m]=formulaResult}
          }else{dailyResult.value[m]=dailyResult.value[m]}
        }
        dailyDash.getRange(2,dailyCol,1,dailyResult.days.length).setValues([dailyResult.days]).setNumberFormat('MM/DD/YYYY').setBackground("white")
        dailyDash.getRange(dailyResult.curCounter,2,dailyResult.metric.length,dailyResult.metric[0].length).setValues(dailyResult.metric)
        dailyDash.getRange(dailyResult.curCounter,5,dailyResult.target.length,dailyResult.target[0].length).setValues(dailyResult.target)
        dailyDash.getRange(dailyResult.curCounter,dailyCol,dailyResult.value.length,dailyResult.value[0].length).setValues(dailyResult.value)
        dailyDash.getRange(dailyResult.curCounter,2,dailyResult.metric.length).setBackground(teamBg).setFontColor("black").setFontSize(10).setFontWeight("normal").setFontFamily("Poppins")
        dailyDash.getRange("C"+dailyResult.curCounter+":"+(dailyResult.curCounter+dailyResult.metric.length)).setBackground("white").setFontColor("black").setFontSize(10).setFontWeight("normal").setFontFamily("Poppins")
        dailyResult.curCounter+=dailyResult.metric.length
      }else{
        dailyResult.curCounter++;
      }
      
      //paste result to Weekly
      if(weekResult.metric[0]!=null&&weeklyDash){
        for(var m=0;m<weekResult.metric.length;m++){
          var formula = configData[metricList.indexOf(weekResult.metric[m][1])]
          if(formula[1]!=[]){
            var metricDeArray = weekResult.metric.map(function(item){return item[1]})
            var formulaResult = formulaLoop(formula[1],metricDeArray,weekResult.value)
            if(formulaResult!=null){weekResult.value[m]=formulaResult}
          }else{weekResult.value[m]=weekResult.value[m]}
        }
        weeklyDash.getRange(2,weeklyCol,1,weekResult.days.length).setValues([weekResult.days]).setBackground("white")
        weeklyDash.getRange(weekResult.curCounter,2,weekResult.metric.length,weekResult.metric[0].length).setValues(weekResult.metric).setBackground("white").setFontColor("black")
        weeklyDash.getRange(weekResult.curCounter,5,weekResult.target.length,weekResult.target[0].length).setValues(weekResult.target)
        weeklyDash.getRange(weekResult.curCounter,weeklyCol,weekResult.value.length,weekResult.value[0].length).setValues(weekResult.value)
        weeklyDash.getRange(weekResult.curCounter,2,weekResult.metric.length).setBackground(teamBg).setFontColor("black").setFontSize(10).setFontWeight("normal").setFontFamily("Poppins")
        weeklyDash.getRange("C"+weekResult.curCounter+":"+(weekResult.curCounter+weekResult.metric.length)).setBackground("white").setFontColor("black").setFontSize(10).setFontWeight("normal").setFontFamily("Poppins")
        weekResult.curCounter+=weekResult.metric.length
      }else{
        weekResult.curCounter++;
      }
      
      //paste result to Monthly
      if(monthResult.metric[0]!=null&&monthlyDash){
        for(var m=0;m<monthResult.metric.length;m++){
          var formula = configData[metricList.indexOf(monthResult.metric[m][1])]
          if(formula[1]!=[]){
            var metricDeArray = monthResult.metric.map(function(item){return item[1]})
            var formulaResult = formulaLoop(formula[1],metricDeArray,monthResult.value)
            if(formulaResult!=null){monthResult.value[m]=formulaResult}
          }else{monthResult.value[m]=monthResult.value[m]}
        }
        monthlyDash.getRange(2,monthlyCol,1,monthResult.days.length).setValues([monthResult.days])
        monthlyDash.getRange(monthResult.curCounter,2,monthResult.metric.length,monthResult.metric[0].length).setValues(monthResult.metric).setBackground("white").setFontColor("black")
        monthlyDash.getRange(monthResult.curCounter,5,monthResult.target.length,monthResult.target[0].length).setValues(monthResult.target)
        monthlyDash.getRange(monthResult.curCounter,monthlyCol,monthResult.value.length,monthResult.value[0].length).setValues(monthResult.value)
        monthlyDash.getRange(monthResult.curCounter,2,monthResult.metric.length).setBackground(teamBg).setFontColor("black").setFontSize(10).setFontWeight("normal").setFontFamily("Poppins")
        monthlyDash.getRange("C"+monthResult.curCounter+":"+(monthResult.curCounter+monthResult.metric.length)).setBackground("white").setFontColor("black").setFontSize(10).setFontWeight("normal").setFontFamily("Poppins")
        monthResult.curCounter+=monthResult.metric.length
      }else{
        monthResult.curCounter++;
      }
    
    }//end if applicable
  }//end team loop

  //RUN FORMAT FOR DASHBOARD
  // if(!divisionPart){
  //   if(dailyDash){runFormat(dailyDash)}
  //   if(weeklyDash){runFormat(weeklyDash)}
  //   if(monthlyDash){runFormat(monthlyDash)} 
  // }
  
}

//update new week --- use to update new number weekly
function updateWeekF(config,dailyDash,weeklyDash,monthlyDash,divisionPart){
  //Dashboard setup
  // var dailyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Daily Dashboard')
  // var weeklyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Weekly Dashboard')
  // var monthlyDash = SpreadsheetApp.openById(dashboardId).getSheetByName('Monthly Dashboard')
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config)
  var dailyCol = 7
  var weeklyCol = 7
  var monthlyCol = monthlyDash.getName()=='Monthly CO'||monthlyDash.getName()=='Monthly CS'||monthlyDash.getName()=='Monthly CC'||monthlyDash.getName()=='Monthly Dashboard'||monthlyDash.getName()=='Monthly' ? 9 : 7
  // var monthlyCol=7
  //Config info
  var configData = configSheet.getRange(startRowF,startColF).getDataRegion().getValues() //0=metric, 1=formula, 2=data file, 3=attributes
  var title = configSheet.getRange(startRowF,startColF,1,configSheet.getLastColumn()+1-startColF).getValues()[0]
  var dataValidationList = configSheet.getRange(opLRow,opLCol).getDataRegion().getValues()
  dataValidationList.shift()
  var nameL = dataValidationList.map(function(e){return e[0]})
  var urlL = dataValidationList.map(function(e){return SpreadsheetApp.openByUrl(e[1])})
  // nameL.shift(); urlL.shift()
  //Result setup  
  var dailyResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
  dailyResult.curCountry
  dailyResult.curCounter = divisionPart ? dailyDash.getRange(1,2,dailyDash.getLastRow()).getValues().map(function(r){return r[0]}).indexOf(divisionPart)+1 : 3
  dailyResult.days=[] 
  var weekResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
  weekResult.curCountry
  weekResult.curCounter = divisionPart ? weeklyDash.getRange(1,2,weeklyDash.getLastRow()).getValues().map(function(r){return r[0]}).indexOf(divisionPart)+1 : 4
  weekResult.days=[] 
  var monthResult={}//[team,metric,value] team=[country,team] metric=[applicable metrics ->destination sheet]
  monthResult.curCountry
  monthResult.curCounter = divisionPart ? monthlyDash.getRange(1,2,monthlyDash.getLastRow()).getValues().map(function(r){return r[0]}).indexOf(divisionPart)+1 : 4
  monthResult.days=[] 
  //setup length of dayRange,weekRange,monthRange
  var day=startReportPeriod
  while(day<=endReportPeriod){
    dailyResult.days.push(timeZone(day))
    if(weekResult.days.indexOf(timeZone(day,'startWeekDate'))==-1){ weekResult.days.push(timeZone(day,'startWeekDate'))}
    if(monthResult.days.indexOf(timeZone(day,'startMonthDate'))==-1){ monthResult.days.push(timeZone(day,'startMonthDate'))}
    day = new Date(day.getTime() + MILLIS_PER_DAY)
  }
  weekResult.days.shift()
  
  for (var i=0; i<title.length;i++){//loop team

      var partition = divisionPart ? title[i]=='Applicable' && configSheet.getRange(startRowF-2,startColF+i).getValue().split('|')[0].indexOf(divisionPart)>-1 : title[i]=='Applicable'

    if(partition){//if col is Applicable
    // if(title[i]=='Applicable'){
      var team = configSheet.getRange(startRowF-2,startColF+i).getValue().split('|')
      
      dailyResult.metric=[]
      dailyResult.value=[]
      weekResult.metric=[]
      weekResult.value=[]
      monthResult.metric=[]
      monthResult.value=[]
      
      if(dailyResult.curCountry!=team[0]){//if country not exists ->print country to sheets
        dailyResult.curCountry=team[0]
        weekResult.curCountry=team[0]
        monthResult.curCountry=team[0]
        dailyResult.curCounter++;
        weekResult.curCounter++;
        monthResult.curCounter++;
      }
      
      if(team[1]!=''){
        dailyResult.curCounter++;
        weekResult.curCounter++;
        monthResult.curCounter++;
      }
      
      var applicableList = configSheet.getRange(startRowF,startColF+i,configData.length,3).getValues() //[dash type, filters(indicator = ',')]
      
      //loop metrics
      for(var j=0;j<applicableList.length;j++){
        //split filter 
        var filters = applicableList[j][1].split('&')

        //Filter aggregation 'All'
        if(applicableList[j][0]=='All'){
          // Logger.log([team[0]+" "+team[1],configData[j][0]])
          //if no filters -> include metrics in the list || if filters exist -> push metrics info
          if(filters[0]==''){
            dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            var res=[]
            dailyResult.days.forEach(function(r){
              res.push('')
            })
            dailyResult.value.push(res)
            var res=[]
            weekResult.days.forEach(function(r){
              res.push('')
            })
            weekResult.value.push(res)
            var res=[]
            monthResult.days.forEach(function(r){
              res.push('')
            })
            monthResult.value.push(res)
            
          }else{
            //push metrics || configData[[0=metric, 1=formula, 2=data file, 3=attributes]]
            dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
            Logger.log([team[0]+" "+team[1],configData[j][0]])
            var data =[]
            filters.forEach(function(filter){
              
              var subFilter = filter.split(',')
              var subData = []
              subFilter[0].split(':')[1].split('|').forEach(function(filename){
                var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename)
                subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
              })
              //apply filters to data
              data = data.concat(filterCustom(subData,subFilter))
            })

            //grab data from database
            var curDat=[]
            curDat=attributeOperation(data,configData[j][3],[dailyResult.days,weekResult.days,monthResult.days],applicableList[j][0])

            //push val to result
            dailyResult.value.push(curDat[0])
            weekResult.value.push(curDat[1])
            monthResult.value.push(curDat[2])
          }
          
        }else {//Filter aggregation 'Daily,Weekly,Monthly'

          var dashType = applicableList[j][0].split(',')
          dashType.forEach(function(type){//loop dash type
            switch (type){
              case 'Daily':
              // Logger.log([team[0]+" "+team[1],configData[j][0]])
                if(filters[0]==''){
                  dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  
                  var res=[]
                  dailyResult.days.forEach(function(r){
                    res.push('')
                  })
                  dailyResult.value.push(res)
                }else{
                  dailyResult.metric.push([team[0]+" "+team[1],configData[j][0]])

                  var data =[]
                  filters.forEach(function(filter){
                    var subFilter = filter.split(',')
                    var subData = []
                    subFilter[0].split(':')[1].split('|').forEach(function(filename){
                      var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename.toString().replace('*','Daily'))
                      subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
                    })
                    //apply filters to data
                    data = data.concat(filterCustom(subData,subFilter))
                  })

                  //grab data from database
                  var curDat=[]
                  curDat=attributeOperation(data,configData[j][3],[dailyResult.days],type)
                  //push val to result
                  dailyResult.value.push(curDat[0])                  
                }
                break;
              case 'Weekly':
              // Logger.log([team[0]+" "+team[1],configData[j][0]])
                if(filters[0]==''){
                  weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  var res=[]
                  weekResult.days.forEach(function(r){
                    res.push('')
                  })  
                  weekResult.value.push(res)
                }else{
                  weekResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  var data =[]
                  filters.forEach(function(filter){
                    var subFilter = filter.split(',')
                    var subData = []
                    subFilter[0].split(':')[1].split('|').forEach(function(filename){
                      var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename.toString().replace('*','Weekly'))
                      subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
                    })
                    //apply filters to data
                    data = data.concat(filterCustom(subData,subFilter))
                  })

                  //grab data from database
                  var curDat=[]
                  curDat=attributeOperation(data,configData[j][3],[weekResult.days],type)

                  //push val to result
                  weekResult.value.push(curDat[0])                  
                }
                break;
              case 'Monthly':
              // Logger.log([team[0]+" "+team[1],configData[j][0]])
                if(filters[0]==''){
                  monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  var res=[]
                  monthResult.days.forEach(function(r){
                    res.push('')
                  })  
                  monthResult.value.push(res)
                }else{
                  monthResult.metric.push([team[0]+" "+team[1],configData[j][0]])
                  
                  //grab data from sheets
                  var data =[]
                  filters.forEach(function(filter){
                    var subFilter = filter.split(',')
                    var subData = []
                    subFilter[0].split(':')[1].split('|').forEach(function(filename){
                      var file = urlL[nameL.indexOf(configData[j][2])].getSheetByName(filename.toString().replace('*','Monthly'))
                      subData = subData.concat(file.getRange(1,1).getDataRegion().getValues())
                    })
                    //apply filters to data
                    data = data.concat(filterCustom(subData,subFilter))
                  })

                  //grab data from database
                  var curDat=[]
                  curDat=attributeOperation(data,configData[j][3],[monthResult.days],type)
                  
                  //push val to result
                  monthResult.value.push(curDat[0])                  
                }
                break;
            }
          })
          
        }// if daily,weekly,monthly
      }//end loop metrics

      //PASTE RESULT TO SHEET
      var metricList = configData.map(function(c){return c[0]})
      
      if(dailyResult.metric[0]!=null&&dailyDash){
        //run formula loop
        for(var m=0;m<dailyResult.metric.length;m++){
          var formula = configData[metricList.indexOf(dailyResult.metric[m][1])]
          if(formula[1]!=[]){
            var metricDeArray = dailyResult.metric.map(function(item){return item[1]})
            var formulaResult = formulaLoop(formula[1],metricDeArray,dailyResult.value)
            if(formulaResult!=null){dailyResult.value[m]=formulaResult}
          }else{dailyResult.value[m]=dailyResult.value[m]}
        }
        dailyDash.getRange(2,dailyCol,1,dailyResult.days.length).setValues([dailyResult.days]).setNumberFormat('MM/DD/YYYY')
        dailyDash.getRange(dailyResult.curCounter,dailyCol,dailyResult.value.length,dailyResult.value[0].length).setValues(dailyResult.value)
        dailyResult.curCounter+=dailyResult.metric.length
      }else{
        dailyResult.curCounter++;
      }
      
      //paste result to Weekly
      if(weekResult.metric[0]!=null&&weeklyDash){
        for(var m=0;m<weekResult.metric.length;m++){
          var formula = configData[metricList.indexOf(weekResult.metric[m][1])]
          if(formula[1]!=[]){
            var metricDeArray = weekResult.metric.map(function(item){return item[1]})
            var formulaResult = formulaLoop(formula[1],metricDeArray,weekResult.value)
            if(formulaResult!=null){weekResult.value[m]=formulaResult}
          }else{weekResult.value[m]=weekResult.value[m]}
        }
        weeklyDash.getRange(2,weeklyCol,1,weekResult.days.length).setValues([weekResult.days]).setNumberFormat('MM/DD/YYYY')
        weeklyDash.getRange(weekResult.curCounter,weeklyCol,weekResult.value.length,weekResult.value[0].length).setValues(weekResult.value)
        weekResult.curCounter+=weekResult.metric.length
      }else{
        weekResult.curCounter++;
      }
      
      //paste result to Monthly
      if(monthResult.metric[0]!=null&&monthlyDash){
        for(var m=0;m<monthResult.metric.length;m++){
          var formula = configData[metricList.indexOf(monthResult.metric[m][1])]
          if(formula[1]!=[]){
            var metricDeArray = monthResult.metric.map(function(item){return item[1]})
            var formulaResult = formulaLoop(formula[1],metricDeArray,monthResult.value)
            if(formulaResult!=null){monthResult.value[m]=formulaResult}
          }else{monthResult.value[m]=monthResult.value[m]}
        }
        monthlyDash.getRange(2,monthlyCol,1,monthResult.days.length).setValues([monthResult.days]).setNumberFormat('MMMM, YYYY')
        monthlyDash.getRange(monthResult.curCounter,monthlyCol,monthResult.value.length,monthResult.value[0].length).setValues(monthResult.value)
        monthResult.curCounter+=monthResult.metric.length
      }else{
        monthResult.curCounter++;
      }
    
    }//end if applicable
  }//end team loop
}