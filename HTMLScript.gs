//GET SECTION ARRAY
function getSectionArray(setupSheetName){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(setupSheetName)
  var sectionRange = sheet.getRange(startRowF-2,startColF+5,1,sheet.getLastColumn()-(startColF+5)).getValues()[0]
  var sectionArray = sectionRange.map(function(r){return r.split("|")[0]})
  sectionArray = ArrayLib.unique(sectionArray)
  sectionArray.unshift("All")
  return sectionArray
}
function getSetupSheetDropDown(){
  var setupSheetList ={}
  setupSheetList.setupSheetListName = SHEETNAME(-1,'System Setup')
  setupSheetList.setupSheetSections = []
  setupSheetList.setupSheetListName.forEach(function(sheet){
    setupSheetList.setupSheetSections.push(getSectionArray(sheet))
  })
  return setupSheetList
}
function displaySidebar() {
  const htmlForSidebar = HtmlService.createTemplateFromFile("Sidebar")
  const htmlOutput = htmlForSidebar.evaluate()
  var ui = SpreadsheetApp.getUi()
  ui.showSidebar(htmlOutput)
}

//SIDEBAR FUNCTION

function userClicked(info){
//info ={}
//info.country='Canada'
//info.name='Courier Care'
var country = configSheet.getRange(2,startColF+1).getValue()
//configSheet.getRange(2,startColF+1).setValue(country+', '+info.country)
configSheet.insertColumnsAfter(configSheet.getLastColumn(), 3)
var lastCol=configSheet.getLastColumn()+3
configSheet.setColumnWidth(lastCol-2, 2)
configSheet.getRange(startRowF-2,lastCol-1).setValue(info.country+'|'+info.name)
configSheet.getRange(startRowF,lastCol-1,1,2).setValues([['Applicable','Filter']])
configSheet.getRange(startRowF+1,lastCol-1,configSheet.getLastRow()-startRowF-1,1).insertCheckboxes()
configSheet.getRange(1,startColF,1,lastCol-startColF+1).mergeAcross()
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
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mastersheet')
  var teamNameList = configSheet.getRange(startRowF-2,startColF,1,configSheet.getLastColumn()+1-startColF).getValues()[0]
  var newTeamName = info.country+'|'+info.name
  if(teamNameList.indexOf(newTeamName)==-1&&info.name!=''){return true}else{return false}
}
