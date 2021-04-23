function menuTaskUpdate(sourceUrl,sourceSheet,colUsed,option,targetUrl,targetSheet) {
  var source = SpreadsheetApp.openByUrl(sourceUrl).getSheetByName(sourceSheet)
  var target = SpreadsheetApp.openByUrl(targetUrl).getSheetByName(targetSheet)
  var colUsed = colUsed.split(',')
  var last = new Date(thisSunday.getTime() - MILLIS_PER_DAY*7)
  //tenant_id	date	team	channel	direction	universal_code	total_contacts
  switch(option){
    case 'Salesforce':
    //date	platform	case_type	svl_category	svl_subcategory	task_created	task_solved	task_within_24h	avg_close_time_hour
    //Date (Case Created)	Case Type/Priority	SVL Category	Tickets Created	Tickets Solved	Tickets On 24 hr SVL	Cases Ave Hrs to Close
    var dataVal = source.getRange(3,1,source.getLastRow()-1,source.getLastColumn()).getValues()
                  .filter(function(r){return r[0]!=""})
                  .map(function(r){return [r[colUsed[0]-1],'salesforce',r[colUsed[1]-1],r[colUsed[2]-1],r[colUsed[3]-1],r[colUsed[4]-1],r[colUsed[5]-1],r[colUsed[6]-1]]})
                  .filter(function(r){return timeZone(r[0],'startWeekDate')==timeZone(thisSunday,'startWeekDate')})
    break;
    case 'LiveOps':
    //date	platform	case_type	svl_category	svl_subcategory	task_created	task_solved	task_within_24h	avg_close_time_hour
    //Date (Ticket Created)	Case Type/Priority	restaurant_preferred_locale_language	SVL Category	Tasks Created	Tasks Solved	Tickets On 24 hr SVL	Cases Ave Hrs to Close
    var dataVal = source.getRange(3,1,source.getLastRow()-1,source.getLastColumn()).getValues()
                  .filter(function(r){return r[0]!=""})
                  .map(function(r){return [r[colUsed[0]-1],'liveops',r[colUsed[1]-1]+"_"+r[colUsed[2]-1],r[colUsed[3]-1],r[colUsed[4]-1],r[colUsed[5]-1],r[colUsed[6]-1],r[colUsed[7]-1]]})
                  .filter(function(r){return timeZone(r[0],'startWeekDate')==timeZone(thisSunday,'startWeekDate')})
    break;
    case 'Zendesk':
    //date	platform	case_type	svl_category	svl_subcategory	task_created	task_solved	task_within_24h	avg_close_time_hour
    //Date (Ticket Created)	Case Type/Priority	SVL Category	Tasks Created	Tasks Solved	Tickets On 24 hr SVL	Cases Ave Hrs to Close
    var dataVal = source.getRange(3,1,source.getLastRow()-1,source.getLastColumn()).getValues()
                  .filter(function(r){return r[0]!=""})
                  .map(function(r){return [r[colUsed[0]-1],'zendesk',r[colUsed[1]-1],r[colUsed[2]-1],r[colUsed[3]-1],r[colUsed[4]-1],r[colUsed[5]-1],r[colUsed[6]-1]]})
                  .filter(function(r){return timeZone(r[0],'startWeekDate')==timeZone(thisSunday,'startWeekDate')})
    break;
  }
  if(dataVal){
    target.getRange(target.getRange(1,1).getDataRegion().getLastRow()+1,1,dataVal.length,dataVal[0].length).setValues(dataVal)
  }
}

