//TIME CONVERT
function timeZone(val,option){
  switch(option){
    default: return Utilities.formatDate(val, 'America/Winnipeg', 'MMMM dd, yyyy');
      break;
    case 1:
      var v = new Date( Utilities.formatDate(val, 'America/Winnipeg', "MMM d, yyyy") );
      // Get the week day where 1 = Monday and 7 = Sunday
      var dayOfWeek = ((v.getDay() + 6) % 7) + 1;
      // Locate the nearest thursday
      v.setDate(v.getDate() + (4 - dayOfWeek));
      var jan1 = new Date(v.getFullYear(), 0, 1);
      // Calculate the number of days in between the nearest thursday and januari first of 
      // the same year as the nearest thursday.
      var deltaDays = Math.floor( (v.getTime() - jan1.getTime()) / (86400 * 1000) )
      var weekNumber = 1 + Math.floor(deltaDays / 7);
      return weekNumber;
      break;
      
    case 'startWeekDate': //get start date of the week
      // Get the week day where 1 = Monday and 7 = Sunday
      var dayOfWeek = ((val.getDay() + 6) % 7) + 1;
      var startWeekDate = new Date(val.getTime() - MILLIS_PER_DAY*(dayOfWeek-1))
      return Utilities.formatDate(startWeekDate, 'America/Winnipeg', 'MM/dd/yyyy');
      break;
    case 'startMonthDate':
      var startMonthDate = new Date(val.getFullYear(), val.getMonth(), 1)
      return Utilities.formatDate(startMonthDate, 'America/Winnipeg', 'MM/dd/yyyy')
      break;
    case 2: return Utilities.formatDate(val, 'America/Winnipeg', 'M'); break;
    case 3: return Utilities.formatDate(val, 'America/Winnipeg', 'MM/dd/yyyy');break;
    case 4:return Utilities.formatDate(val, 'America/Winnipeg', 'MMMM, yyyy');break;
  }
}

//INDEX OF 2D ARRAY VERSION
/* Usage:
    2D_Array.indexOfArray2D(2D_value_array) 
   Returns:
    The index where 2D_value_array was found in 2D_Array or -1 if it is not found
*/
  Array.prototype.indexOfArray2D = function(array){
    var index=-1
    var i=0
    while(i<this.length){
      var j=0
      while (j<this[i].length){
        if(array[j]!=this[i][j]){
          break;
        }else{
          j++
        }
      }
      if(j==this[i].length){
        index=i
        break;
      }else{
        i++
      }
    }
    return index
  }
//operations 2D array
function calculateArrays(array1,operator,array2){
  var result=[]
  for(var i=0;i<array1.length;i++){
    var res=[]
    switch (operator)
    {
      case '+': 
        for(var j=0;j<array1[i].length;j++){
          var val1 = 0
          var val2 = 0
          array1[i][j]==''||typeof array1[i][j] !='number'?val1=0:val1=array1[i][j]
          array2[i][j]==''||typeof array2[i][j] !='number'?val2=0:val2=array2[i][j]
          res.push(val1+val2)
        }
        result.push(res)
        break;
      case '-': 
        for(var j=0;j<array1[i].length;j++){
          var val1 = 0
          var val2 = 0
          array1[i][j]==''||typeof array1[i][j] !='number'?val1=0:val1=array1[i][j]
          array2[i][j]==''||typeof array2[i][j] !='number'?val2=0:val2=array2[i][j]
          res.push(val1-val2)
        }
        result.push(res)
        break;
      case '*': 
        for(var j=0;j<array1[i].length;j++){
          var val1 = 0
          var val2 = 0
          array1[i][j]==''||typeof array1[i][j] !='number'?val1=0:val1=array1[i][j]
          array2[i][j]==''||typeof array2[i][j] !='number'?val2=0:val2=array2[i][j]
          res.push(val1*val2)
        }
        result.push(res)
        break;
      case '/': 
        for(var j=0;j<array1[i].length;j++){
          if(isFinite(array1[i][j]/array2[i][j])){res.push(array1[i][j]/array2[i][j])}else{res.push('')}
        }
        result.push(res)
        break;
    }
  }
  return result
}

//attribute operation
//attributeOperation(data,configData[j][3],dailyResult.days,weekResult.days,monthResult.days,applicableList[j][0])
function attributeOperation(data,formula,dayRange,dashType){
  if(formula.indexOf('(')==-1&&formula.indexOf(')')==-1){
    var curDat=[]
    var operator=''
    var lastIndex=0
    for(var m=0;m<=formula.length;m++){
      if(formula[m]=='/'||formula[m]=='-'||formula[m]=='*'||formula[m]=='+'||m==formula.length){
        if(operator==''){
          curDat=dataGrab(data,formula.substring(lastIndex,m),dayRange,dashType)
          operator=formula[m]
          lastIndex=m+1
        }else{
          var nextDat = dataGrab(data,formula.substring(lastIndex,m),dayRange,dashType)
          curDat=calculateArrays(curDat,operator,nextDat)
          operator=formula[m]
          lastIndex=m+1
        }
      }
    }
    return curDat
  }else{
    var formulaList=[]
    var lastIndex=0
    for(var m=0;m<=formula.length;m++){
      if(formula[m]=='('||formula[m]==')'){
        if(formula.substring(lastIndex,m)==''){lastIndex=m+1}else{
          formulaList.push(formula.substring(lastIndex,m))
        }
        lastIndex=m+1
      }
    }
    var finalResult=[]
    var tempOperation=''
    for(var i=0;i<formulaList.length;i++){

      if(formulaList[i]=='/'||formulaList[i]=='-'||formulaList[i]=='*'||formulaList[i]=='+'){
        tempOperation=formulaList[i]
      }else{
          var curDat=[]
          var operator=''
          var lastIndex=0
          for(var f=0;f<=formulaList[i].length;f++){
            if(formulaList[i][f]=='/'||formulaList[i][f]=='-'||formulaList[i][f]=='*'||formulaList[i][f]=='+'||f==formulaList[i].length){
              if(operator==''){
                curDat=dataGrab(data,formulaList[i].substring(lastIndex,f),dayRange,dashType)
                operator=formulaList[i][f]
                lastIndex=f+1
              }else{
                var nextDat = dataGrab(data,formulaList[i].substring(lastIndex,f),dayRange,dashType)
                curDat=calculateArrays(curDat,operator,nextDat)
                operator=formulaList[i][f]
                lastIndex=f+1
              }
            }
          }
          if(tempOperation!=''&&finalResult[0]!=null&&curDat[0]!=null){
            finalResult=calculateArrays(finalResult,tempOperation,curDat)
            tempOperation=''
          }else if(tempOperation==''){
            finalResult=curDat
          }
      }

    }
  return finalResult
  }
  
}

//grab data
function dataGrab(data,colName,dayRange,dashType){
  switch(dashType)
  {
    case 'All':
      var dRange = dayRange[0]
      var wRange = dayRange[1]
      var mRange = dayRange[2]
      var dailyVal =[]
      dRange.forEach(function(r){
        dailyVal.push(0)
      })
      var weeklyVal =[]
      wRange.forEach(function(r){
        weeklyVal.push(0)
      })
      var monthlyVal =[]
      mRange.forEach(function(r){
        monthlyVal.push(0)
      })
      var dayIndex = data[0].indexOf('Date')>-1 ? data[0].indexOf('Date') : data[0].indexOf('date')
      data.forEach(function(r){ //values of 30 days for metric j of team i
        if(r[dayIndex]>=startReportPeriod && r[dayIndex]<=endReportPeriod){
          //daily data
          dailyVal[dRange.indexOf(timeZone(r[dayIndex]))]+=r[data[0].indexOf(colName)]==''||typeof r[data[0].indexOf(colName)] !='number'? 0 : r[data[0].indexOf(colName)]
          //weekly data
          if(wRange.indexOf(timeZone(r[dayIndex],'startWeekDate'))>-1){
            weeklyVal[wRange.indexOf(timeZone(r[dayIndex],'startWeekDate'))]+= r[data[0].indexOf(colName)]==''||typeof r[data[0].indexOf(colName)] !='number'? 0 : r[data[0].indexOf(colName)]
          }            
          //monthly data
          if(mRange.indexOf(timeZone(r[dayIndex],'startMonthDate'))>-1){
            monthlyVal[mRange.indexOf(timeZone(r[dayIndex],'startMonthDate'))]+=r[data[0].indexOf(colName)]==''||typeof r[data[0].indexOf(colName)] !='number'? 0 : r[data[0].indexOf(colName)]
          }
        }
      })
      
      return [dailyVal,weeklyVal,monthlyVal]
      break;
    case 'Daily':
      return [aggregateDate(data,colName,dayRange[0])]
      break;
    case 'Weekly':
      return [aggregateDate(data,colName,dayRange[0],'startWeekDate')]
      break;
    case 'Monthly':
      return [aggregateDate(data,colName,dayRange[0],'startMonthDate')]
      break;
  }
  
}

//aggregateData(data,dayrange,''startMonthDate'')
function aggregateDate(data,colName,dayRange,startDateType){
  var val =[]
  dayRange.forEach(function(r){
    val.push(0)
  })
  var dayIndex = data[0].indexOf('Date')>-1 ? data[0].indexOf('Date') : data[0].indexOf('date')
  data.forEach(function(r){ //values of 30 days for metric j of team i
    if(r[dayIndex]>=startReportPeriod && r[dayIndex]<=endReportPeriod){
      //monthly data
      if(dayRange.indexOf(timeZone(r[dayIndex],startDateType))>-1){
        val[dayRange.indexOf(timeZone(r[dayIndex],startDateType))]+=r[data[0].indexOf(colName)]==''||typeof r[data[0].indexOf(colName)] !='number'? 0 : r[data[0].indexOf(colName)]
      }
    }
  })
  return val
}

//filter
function filterCustom(data,filters){
  var mapSource = SpreadsheetApp.openById(mapSourceId)
  
  var dataTitle = data[0]
  if(filters.length>=2){//if more filter add in beside sheet name
    for(var k=1;k<filters.length;k++){
      var colName = filters[k].split(':')[0]
      if(filters[k].split(':')[0]!='map'){//if there's no map file
      var subFilter=[]
        if(filters[k].indexOf('[')>-1){
          subFilter = filters[k].substring(filters[k].indexOf('[')+1,filters[k].indexOf(']')).split('|')
        }else{
          subFilter = filters[k].substring(filters[k].indexOf(":")+1,filters[k].length).split('|')
        }
        var subData = []
        subFilter.forEach(function(f){
          subData = subData.concat(data.filter(function(r){return r[dataTitle.indexOf(filters[k].split(':')[0])].toString().toLowerCase() == f.toString().toLowerCase()}))
        })
        data = subData.slice()
      }else{
        var string = filters[k].split(':')[1] //string = sheetname[mapCol=databaseCol]filterMapValue
        var mapSheet = mapSource.getSheetByName(string.substring(0,string.indexOf('['))).getDataRange().getValues() //sheetname
        var mapRule = string.substring(string.indexOf('[')+1,string.indexOf(']')).split('=') //substring inside []
        var mapFrom = dataTitle.indexOf(mapRule[1])//databaseCol
        var mapTo = mapSheet.map(function(r){return [r[0],r[mapSheet[0].indexOf(mapRule[0])].replace(", ",",").split(",")]})
                            .filter(function(r){return r[1][0]!=''}) //2D [resultFromMap,[databaseVal]] where [databaseVal] isn't empty
        //push a col for map result
        data.forEach(function(r){
          r.push('No tag')
          if(r[mapFrom]==''){
            r[r.length-1]='No tag'
          }else{
            var i=0
            var check=false
            while(check==false&&i<mapTo.length){
              var mapIndex = mapTo[i][1].indexOf(r[mapFrom])
              check = mapIndex>-1?true:false
              i++
            }
          r[r.length-1] = check==false?'Others':mapTo[i-1][0]
          }
        })
        
        var mapFilter = string.substring(string.indexOf(']')+1,string.length).split('|')
        var subData = []
        mapFilter.forEach(function(f){
          subData = subData.concat(data.filter(function(r){return r[r.length-1].toString().toLowerCase() == f.toString().toLowerCase()}))
        })
        data = subData.slice()
        var mapFilter = string.substring(string.indexOf(']')+1,string.length).split('|')
        var subData = []
        mapFilter.forEach(function(f){
          subData = subData.concat(data.filter(function(r){return r[r.length-1].toString().toLowerCase() == f.toString().toLowerCase()}))
        })
        data = subData.slice()
      }
    }
  }
  data.unshift(dataTitle)
  return data
}
            
//Formula Loop
function formulaLoop(formula,resultMetric,resultValue){
  if(formula.indexOf('(')==-1&&formula.indexOf(')')==-1){
    var curDat=[]
    var operator=''
    var lastIndex=0
    for(var f=0;f<=formula.length;f++){
      if(formula[f]=='/'||formula[f]=='-'||formula[f]=='*'||formula[f]=='+'||f==formula.length){
        if(resultMetric.indexOf(formula.substring(lastIndex,f))>-1){
          if(operator==''){
            curDat=[resultValue[resultMetric.indexOf(formula.substring(lastIndex,f))]]
            operator=formula[f]
            lastIndex=f+1
          }else{
            var nextDat = [resultValue[resultMetric.indexOf(formula.substring(lastIndex,f))]]
            curDat=calculateArrays(curDat,operator,nextDat)
            operator=formula[f]
            lastIndex=f+1
          }
        }else{
        //  operator==''?operator='':operator=formula[f]
          lastIndex=f+1
        }

      }//if operator
    }
    return curDat[0]
  }else{
    var formulaList=[]
    var lastIndex=0
    for(var m=0;m<=formula.length;m++){
      if(formula[m]=='('||formula[m]==')'){
        if(formula.substring(lastIndex,m)==''){lastIndex=m+1}else{
          formulaList.push(formula.substring(lastIndex,m))
        }
        lastIndex=m+1
      }
    }
    var finalResult=[]
    var tempOperation=''

    for(var i=0;i<formulaList.length;i++){

      if(formulaList[i]=='/'||formulaList[i]=='-'||formulaList[i]=='*'||formulaList[i]=='+'){
        tempOperation=formulaList[i]
      }else{
          var curDat=[]
          var operator=''
          var lastIndex=0
          for(var f=0;f<=formulaList[i].length;f++){
            if(formulaList[i][f]=='/'||formulaList[i][f]=='-'||formulaList[i][f]=='*'||formulaList[i][f]=='+'||f==formulaList[i].length){
              if(resultMetric.indexOf(formulaList[i].substring(lastIndex,f))>-1){
                if(operator==''){
                  curDat=[resultValue[resultMetric.indexOf(formulaList[i].substring(lastIndex,f))]]
                  operator=formulaList[i][f]
                  lastIndex=f+1
                }else{
                  var nextDat = [resultValue[resultMetric.indexOf(formulaList[i].substring(lastIndex,f))]]
                  curDat=calculateArrays(curDat,operator,nextDat)
                  operator=formulaList[i][f]
                  lastIndex=f+1
                }
              }else{
                // operator==''?operator='':operator=formula[f]
                lastIndex=f+1
              }

            }
          }
        if(tempOperation!=''&&finalResult[0]!=null&&curDat[0]!=null){
          finalResult=calculateArrays(finalResult,tempOperation,curDat)
          tempOperation=''
        }else if(tempOperation==''){
          finalResult=curDat
        }
      }

    }
  return finalResult[0]
  }
}
//REMOVE FORMAT ALL DASH
function removeFormat(dashboard){
  //get team list, metrics list, target list
  var teamList = dashboard.getRange(1,2,dashboard.getLastRow()).getValues().map(function(r){return r[0]})
  var metricList = dashboard.getRange(1,3,dashboard.getLastRow()).getValues().map(function(r){return r[0]})
  var targetList = dashboard.getRange(1,5,dashboard.getLastRow()).getValues().map(function(r){return r[0]})
  var curGroupR = 0
  dashboard.clearConditionalFormatRules()
  for(var rowNum=1;rowNum<=metricList.length;rowNum++){
    if(metricList[rowNum-1].indexOf("_")>-1 && curGroupR==0){
      dashboard.getRowGroup(rowNum,1).remove()
      curGroupR=rowNum
    }else if(metricList[rowNum-1].indexOf("_")==-1 && curGroupR>0){  //------if metrics.indexOf("-")==-1 && current group row >0 -> group rows -> set current group row=0
    curGroupR=0
    }
  }
  if(curGroupR>0){  //------if metrics.indexOf("-")==-1 && current group row >0 -> group rows -> set current group row=0
    dashboard.getRowGroupAt(rowNum,1).remove()
  }
}
//Reformat all dash
function runFormat(dashboard){
  //  runFormat()
  //get team list, metrics list, target list
  var teamList = dashboard.getRange(1,2,dashboard.getLastRow()).getValues().map(function(r){return r[0]})
  var metricList = dashboard.getRange(1,3,dashboard.getLastRow()).getValues().map(function(r){return r[0]})
  var targetList = dashboard.getRange(1,5,dashboard.getLastRow()).getValues().map(function(r){return r[0]})
  var curGroupR = 0
  var rules = dashboard.getConditionalFormatRules();

  //set format per rows by loop through rows
  for(var rowNum=1;rowNum<=metricList.length;rowNum++){
    //add note action name
      if(teamList[rowNum-1].indexOf("Customer")>-1&&metricList[rowNum-1]=="Live Actions SVL"){
        dashboard.getRange(rowNum,3).setNote("VERIFY_DUPLICATE_ORDERS"+"\n"+"VERIFY_FRAUD")
      }
      if(teamList[rowNum-1].indexOf("Restaurant")>-1&&metricList[rowNum-1]=="Live Actions SVL"){
        dashboard.getRange(rowNum,3).setNote("CALL_RESTAURANT_TO_PROVIDE_ASSISTANCE_WITH_ORDER")
      }
      if(teamList[rowNum-1].indexOf("Courier")>-1&&metricList[rowNum-1]=="Live Actions SVL"){
        dashboard.getRange(rowNum,3).setNote("ADDRESS_CHANGE"+"\n"+"CALL_COURIER_BEING_HELD_AT_CUSTOMER")
      }
    //conditional format
      if(targetList[rowNum-1]!=''&& typeof targetList[rowNum-1] =='number'){
        rules=conditionalFormatApply(dashboard,targetList[rowNum-1],"G"+rowNum+":"+rowNum,rules)
        if(metricList[rowNum-1].indexOf("Call SVL")>-1||metricList[rowNum-1].indexOf("Chat SVL")>-1||metricList[rowNum-1].indexOf("Actions SVL")>-1){
          if(metricList[rowNum-1]=="Inbound Call SVL 20s"){dashboard.getRange(rowNum,5).setValue(targetList[rowNum-1]*100+"%/20s")}
          else{dashboard.getRange(rowNum,5).setValue(targetList[rowNum-1]*100+"%/90s")}
        }
        if(metricList[rowNum-1]=="Facebook Tickets SVL"||metricList[rowNum-1]=="Escalation Tickets SVL"||metricList[rowNum-1]=="Retention Tickets SVL"){
          dashboard.getRange(rowNum,5).setValue(targetList[rowNum-1]*100+"%/1hr")
        }
        if(metricList[rowNum-1]=="Twitter Tickets SVL"){
          dashboard.getRange(rowNum,5).setValue(targetList[rowNum-1]*100+"%/30mins")
        }
        if(metricList[rowNum-1]=="Emails SVL"){
          dashboard.getRange(rowNum,5).setValue(targetList[rowNum-1]*100+"%/12hrs")
        }
      }
    //if metrics.indexOf("-")>-1 && current group row >0 -> set current group row
      if(metricList[rowNum-1].indexOf("_")>-1 && curGroupR==0){
      curGroupR=rowNum
      }else if(metricList[rowNum-1].indexOf("_")==-1 && curGroupR>0){  //------if metrics.indexOf("-")==-1 && current group row >0 -> group rows -> set current group row=0
      dashboard.getRange(curGroupR+":"+(rowNum-1)).shiftRowGroupDepth(1)
      curGroupR=0
      }
    //if metrics.indexOf("%") -> set number format to %
    if(metricList[rowNum-1]!=''&&rowNum>2){
        var rangeToHighlight = dashboard.getRange("G"+rowNum+":"+rowNum);
        var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenCellEmpty().setBackground("gray")
            .setRanges([rangeToHighlight])
            .build();
            rules.push(rule)
      if(metricList[rowNum-1].indexOf("Difference")>-1){
        dashboard.getRange("C"+rowNum+":"+rowNum).setNumberFormat("[color50]▲0% ;[color3]▼-0% ;[blue]▬0%")
      } else if(metricList[rowNum-1].indexOf("Propensity")>-1){
        dashboard.getRange("C"+rowNum+":"+rowNum).setNumberFormat("#0.00\%")  
      }else if(metricList[rowNum-1].indexOf("%")>-1 || metricList[rowNum-1].indexOf("SVL")>-1||metricList[rowNum-1].indexOf("CSAT")>-1||metricList[rowNum-1].indexOf("Rate")>-1||metricList[rowNum-1].indexOf("rate")>-1){
        dashboard.getRange("C"+rowNum+":"+rowNum).setNumberFormat("#0\%")
      }else if(metricList[rowNum-1].indexOf("Avg")>-1){
        dashboard.getRange("C"+rowNum+":"+rowNum).setNumberFormat("#0.00")
      }else{
        dashboard.getRange("C"+rowNum+":"+rowNum).setNumberFormat("#,###0.##")
      }
    }
  }// metric loop
  //1st col && last col+1 ->jeColor
  if(curGroupR>0){  //------if metrics.indexOf("-")==-1 && current group row >0 -> group rows -> set current group row=0
      dashboard.getRange(curGroupR+":"+(rowNum-1)).shiftRowGroupDepth(1)
  }
  dashboard.getRange("A:A").setBackground(jeColor)
  dashboard.getRange(1,dashboard.getLastColumn()+1,dashboard.getLastRow()).setBackground(jeColor)
  dashboard.getRange((dashboard.getLastRow()+1)+":"+(dashboard.getLastRow()+1)).setBackground(jeColor)
  // dashboard.collapseAllRowGroups()
  // var rule = SpreadsheetApp.newConditionalFormatRule()
  //   .whenCellEmpty().setBackground("gray")
  //   .setRanges([dashboard.getRange("G5:"+dashboard.getLastRow())])
  //   .build();
  // rules.push(rule);
  dashboard.setConditionalFormatRules(rules);
}

//add conditional format
function conditionalFormatApply(dashboard,target,range,rules){
  // var targetList = sheet.getRange("A1:A3").getValues()
  //    var rules = dashboard.getConditionalFormatRules();
    var rangeToHighlight = dashboard.getRange(range);
    // var rule = SpreadsheetApp.newConditionalFormatRule()
    //   .whenCellEmpty().setBackground("gray")
    //   .setRanges([rangeToHighlight])
    //   .build();
    // rules.push(rule);
       var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(target).setBackground("#b7e1cd")
      .setRanges([rangeToHighlight])
      .build();
    rules.push(rule);
      var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(.5).setBackground("#f4c7c3")
      .setRanges([rangeToHighlight])
      .build();
    rules.push(rule); 
       var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty().setBackground("#fce8b2")
      .setRanges([rangeToHighlight])
      .build();
    rules.push(rule);   
  // dashboard.setConditionalFormatRules(rules);
  return rules
}

//RUN QUERY
function runQuery(queryBody) {
  // Replace this value with the project ID listed in the Google
  // Cloud Platform project.
  var projectId = 'just-data-bq-users';
  var request = {
    query: queryBody,
    useLegacySql: false
  };
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job.
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  // Get all the rows of results.
  var rows = queryResults.rows;
  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken: queryResults.pageToken
    });
    rows = rows.concat(queryResults.rows);
  }
  if (rows)

  // Append the headers.
  var headers = queryResults.schema.fields.map(function(field) {
    return field.name;
    });

  // Append the results.
  var data = new Array(rows.length);
  for (var i = 0; i < rows.length; i++) {
    var cols = rows[i].f;
    data[i] = new Array(cols.length);
    for (var j = 0; j < cols.length; j++) {
      data[i][j] = cols[j].v;
     }
    }
  return data
}

//CONVERT NUMBER TO LETTER
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
//LIST SHEETNAMES OF THE FILE
/**
 * List out the sheets' names in the file
 *
 * @param {number} option The list type. (default) lists all sheet names in the spreadsheet file. 0 lists the current sheet name. 1 lists the current spreadsheet file
 * @param {string} contains_text  - [OPTIONAL] - Lists the sheet names that are contained nameContains
 * @return The list of sheet names
 * @customfunction
 */
function SHEETNAME(option,contains_text) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var thisSheet = sheet.getName();
  switch(option){
    //All Sheet Names in Spreadsheet
    default:
      var sheetList = [];
      ss.getSheets().forEach(function(val){
        if(contains_text){
          if(val.getName().indexOf(contains_text)>-1){
              sheetList.push(val.getName())
          }
        }else{
          sheetList.push(val.getName())
        }
      });
      return sheetList;
      break;
    //Current option Sheet Name
    case 0:return thisSheet;break;
    //The Spreadsheet File Name
    case 1:return ss.getName();break;
  }
}


//GET LINK FROM HYPERLINK FORMULA
function linkURL(formula){
  return formula.match(/=hyperlink\("([^"]+)"/i)[1]
}
//RENDER HTML PAGE
function render(file,argObject){
  var tmp = HtmlService.createTemplateFromFile(file)
  if(argObject){
    var keys = Object.keys(argObject);
    keys.forEach(function(key){
      tmp[key] = argObject[key];    
    })
    
  }
  return tmp.evaluate()
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function loadPartialHTML_(partial){
  const htmlServ = HtmlService.createTemplateFromFile(partial)
  return htmlServ.evaluate().getContent()
}
function loadTeamView(){
  return loadPartialHTML_("AddTeam")
}
//TEAM NAME VALIDATION
function TeamInfoValidate(info){
  var teamNameList = configSheet.getRange(startRowF-2,startColF,1,configSheet.getLastColumn()+1-startColF).getValues()[0]
  var newTeamName = info.country+'|'+info.name
  if(teamNameList.indexOf(newTeamName)==-1&&info.name!=''){return true}else{return false}
}