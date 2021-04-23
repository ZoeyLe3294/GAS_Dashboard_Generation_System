function runReasonContact(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet) {
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var mapSource = SpreadsheetApp.openById(mapSourceId)
  var colUsed = colUsed.split(',')
  //check if targetSheet exists in targetUrl
  if(target==null){
  var initSheet = 'Template'
    SpreadsheetApp.openByUrl(targetUrl).getSheetByName(initSheet).copyTo(SpreadsheetApp.openByUrl(targetUrl)).setName(targetSheet);
    target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  }
  var last = new Date(thisSunday.getTime() - MILLIS_PER_DAY*7)
  var result=[]
  //tenant_id	date	team	channel	direction	universal_code	total_contacts
  switch(option){
    case 'Zendesk':
    //tenant_id	date	team	channel	direction	universal_code	total_contacts
    //date ticket_group ticket_channel requester_role reason_for_contact tickets_created
      var dataTitle = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues()[0]
      var dataVal = source.getRange(2,1,source.getLastRow()-1,source.getLastColumn()).getValues()
      .filter(function(r){return r[dataTitle.indexOf('tenant_id')]==targetSheet||sourceSheet.split('_')[1]==targetSheet})
      // .filter(function(r){return r[dataTitle.indexOf('date')]>last &&(r[dataTitle.indexOf('tenant_id')]==targetSheet||sourceSheet.split('_')[1]==targetSheet)})
      var co_reason = mapSource.getSheetByName('CO_reasons').getDataRange().getValues()
      var cs_reason = mapSource.getSheetByName('CS_reasons').getDataRange().getValues()
      dataVal.forEach(function(item){
        var tenant_id = sourceSheet.split('_')[1]
        var day = item[colUsed[0]-1]
        var team = item[colUsed[1]-1]
        var channel = item[colUsed[2]-1]
        var direction = item[colUsed[3]-1]
        var total_contacts = item[colUsed[5]-1]
        var disposition_col = item[colUsed[4]-1]
        var mapSheet = team.indexOf('Onboarding')>-1||team.indexOf('OB')>-1 ? co_reason : cs_reason
        var mapCol = tenant_id.toString().toUpperCase()+' Zendesk'
        var mapTo = mapSheet.map(function(r){return [r[0],r[mapSheet[0].indexOf(mapCol)].replace(", ",",").split(",")]})
                  .filter(function(r){return r[1][0]!=''}) //2D [resultFromMap,[databaseVal]] where [databaseVal] isn't empty
        var universal_code = 'No tag' 
        if(disposition_col==''){
          universal_code = 'No tag'
        }else{
          var i=0
          var check=false
          while(check==false&&i<mapTo.length){
            var mapIndex = mapTo[i][1].indexOf(disposition_col)
            check = mapIndex>-1?true:false
            i++
          }
          universal_code=check==false?'Others':mapTo[i-1][0]
        }
        
        result.push([tenant_id,day,team,channel,direction,universal_code,total_contacts])
      })
    break;
    case 'NonLive Chat':
    //tenant_id	date	team	channel	direction	universal_code	total_contacts
    //tenant_id	date	tag_name	total_chat_sessions	agent_initiated_chat_sessions	courier_initiated_chat_sessions
      var dataTitle = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues()[0]
      var dataVal = source.getRange(2,1,source.getLastRow()-1,source.getLastColumn()).getValues()
      .filter(function(r){return r[dataTitle.indexOf('tenant_id')]==targetSheet||sourceSheet.split('_')[1]==targetSheet})
      // .filter(function(r){return r[dataTitle.indexOf('date')]>last && r[dataTitle.indexOf('tenant_id')]==targetSheet})
      var co_reason = mapSource.getSheetByName('CO_reasons').getDataRange().getValues()
      var cs_reason = mapSource.getSheetByName('CS_reasons').getDataRange().getValues()
      dataVal.forEach(function(item){
        var tenant_id = item[colUsed[0]-1]
        var day = item[colUsed[1]-1]
        var team = 'Courier Success'
        var channel = 'nonlive chat'
        var directionIn = 'Inbound'
        var directionOut = 'Outbound'
        var total_contacts_in = item[colUsed[5]-1]
        var total_contacts_out = item[colUsed[4]-1]
        var disposition_col = item[colUsed[2]-1]
        var mapSheet = team.indexOf('Onboarding')>-1||team.indexOf('OB')>-1 ? co_reason : cs_reason
        var mapCol = tenant_id.toString().toUpperCase()+' Chat'
        if(mapSheet[0].indexOf(mapCol)>-1){
          var mapTo = mapSheet.map(function(r){return [r[0],r[mapSheet[0].indexOf(mapCol)].replace(", ",",").split(",")]})
                      .filter(function(r){return r[1][0]!=''}) //2D [resultFromMap,[databaseVal]] where [databaseVal] isn't empty
          var universal_code = 'No tag' 
          if(disposition_col=='NA'){
            universal_code = 'No tag'
          }else{
            var i=0
            var check=false
            while(check==false&&i<mapTo.length){
              var mapIndex = mapTo[i][1].indexOf(disposition_col)
              check = mapIndex>-1?true:false
              i++
            }
            universal_code=check==false?'Others':mapTo[i-1][0]
          }
          
          result.push([tenant_id,day,team,channel,directionIn,universal_code,total_contacts_in])
          result.push([tenant_id,day,team,channel,directionOut,universal_code,total_contacts_out])
        }else{
          var universal_code = 'No mapping'
          result.push([tenant_id,day,team,channel,directionIn,universal_code,total_contacts_in])
          result.push([tenant_id,day,team,channel,directionOut,universal_code,total_contacts_out])
        }
      })
    break;
    case 'Live Chat':
    //tenant_id	date	team	channel	direction	universal_code	total_contacts
    //tenant_id	date	tag_name	total_chat_sessions	agent_initiated_chat_sessions	courier_initiated_chat_sessions
      var dataTitle = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues()[0]
      var dataVal = source.getRange(2,1,source.getLastRow()-1,source.getLastColumn()).getValues()
      .filter(function(r){return r[dataTitle.indexOf('tenant_id')]==targetSheet||sourceSheet.split('_')[1]==targetSheet})
      // .filter(function(r){return r[dataTitle.indexOf('date')]>last && r[dataTitle.indexOf('tenant_id')]==targetSheet})
      var cc_reason = mapSource.getSheetByName('CC_reasons').getDataRange().getValues();cc_reason.shift()
      dataVal.forEach(function(item){
        var tenant_id = item[colUsed[0]-1]
        var day = item[colUsed[1]-1]
        var team = 'Courier Care'
        var channel = 'live chat'
        var directionIn = 'Inbound'
        var directionOut = 'Outbound'
        var total_contacts_in = item[colUsed[5]-1]
        var total_contacts_out = item[colUsed[4]-1]
        var disposition_col = item[colUsed[2]-1]
        var mapSheet = cc_reason
        var mapCol = tenant_id.toString().toUpperCase()+' Chat'
        if(mapSheet[0].indexOf(mapCol)>-1){
          var mapTo = mapSheet.map(function(r){return [r[0],r[mapSheet[0].indexOf(mapCol)].replace(", ",",").split(",")]})
                      .filter(function(r){return r[1][0]!=''}) //2D [resultFromMap,[databaseVal]] where [databaseVal] isn't empty
          var universal_code = 'No tag' 
          if(disposition_col=='NA'){
            universal_code = 'No tag'
          }else{
            var i=0
            var check=false
            while(check==false&&i<mapTo.length){
              var mapIndex = mapTo[i][1].indexOf(disposition_col)
              check = mapIndex>-1?true:false
              i++
            }
            universal_code=check==false?'Others':mapTo[i-1][0]
          }
          
          result.push([tenant_id,day,team,channel,directionIn,universal_code,total_contacts_in])
          result.push([tenant_id,day,team,channel,directionOut,universal_code,total_contacts_out])
        }else{
          var universal_code = 'No mapping'
          result.push([tenant_id,day,team,channel,directionIn,universal_code,total_contacts_in])
          result.push([tenant_id,day,team,channel,directionOut,universal_code,total_contacts_out])
        }
      })
    break;
    case 'Talkdesk':
    //tenant_id	date	team	channel	direction	universal_code	total_contacts
    //tenant_id	date	team	call_type	disposition_code	total_contacts
      var dataTitle = source.getRange(1,1,source.getLastRow(),source.getLastColumn()).getValues()[0]
      var dataVal = source.getRange(2,1,source.getLastRow()-1,source.getLastColumn()).getValues()
      .filter(function(r){return r[dataTitle.indexOf('tenant_id')]==targetSheet||sourceSheet.split('_')[1]==targetSheet})
      // .filter(function(r){return r[dataTitle.indexOf('tenant_id')]==targetSheet && r[dataTitle.indexOf('date')] >= last})
      var co_reason = mapSource.getSheetByName('CO_reasons').getDataRange().getValues()
      var cs_reason = mapSource.getSheetByName('CS_reasons').getDataRange().getValues()
      dataVal.forEach(function(item){
        var tenant_id = item[colUsed[0]-1]
        var day = item[colUsed[1]-1]
        var team = item[colUsed[2]-1]
        var channel = 'talkdesk'
        var direction = item[colUsed[3]-1]
        var total_contacts = item[colUsed[5]-1]
        var disposition_col = item[colUsed[4]-1]
        var mapSheet = team.indexOf('Onboarding')>-1||team.indexOf('OB')>-1 ? co_reason : cs_reason
        var mapCol = tenant_id.toString().toUpperCase()+' Talkdesk'
        var mapTo = mapSheet.map(function(r){return [r[0],r[mapSheet[0].indexOf(mapCol)].replace(", ",",").split(",")]})
                  .filter(function(r){return r[1][0]!=''}) //2D [resultFromMap,[databaseVal]] where [databaseVal] isn't empty
        var universal_code = 'No tag' 
        if(disposition_col=='NA'){
          universal_code = 'No tag'
        }else{
          var i=0
          var check=false
          while(check==false&&i<mapTo.length){
            var mapIndex = mapTo[i][1].indexOf(disposition_col)
            check = mapIndex>-1?true:false
            i++
          }
          universal_code=check==false?'Others':mapTo[i-1][0]
        }
        
        result.push([tenant_id,day,team,channel,direction,universal_code,total_contacts])
      })
    break;
  }
  if(result.length>1){
    var resultVal = result.map(function(entry){return entry[entry.length-1]})
    var resultName = result.map(function(entry){
      entry[1] = timeZone(entry[1],3)
      var temp = []
      for (var i=0;i<entry.length-1;i++){temp.push(entry[i])}
      return temp
    })
    var finalOutputVal = []
    var finalOutput = []
    for (var i=0;i<resultName.length;i++){
      var index = finalOutput.indexOfArray2D(resultName[i])
      if(index==-1){
        finalOutput.push(resultName[i])
        finalOutputVal.push(resultVal[i])
      }else{
        finalOutputVal[index] += resultVal[i]
      }
    }
    for (var i=0;i<finalOutput.length;i++){
      finalOutput[i].push(finalOutputVal[i])
    }
    target.getRange(target.getRange(1,1).getDataRegion().getLastRow()+1,1,finalOutput.length,finalOutput[0].length).setValues(finalOutput)
  }
}

