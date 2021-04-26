//ON EDIT
// function onEdit(){
//   var s = SpreadsheetApp.getActiveSheet();

//   //Output file
//   if( s.getName() == 'SET UP') {
//     //Data Validation range
//     var dataValidationList = s.getRange(opLRow,opLCol).getDataRegion().getValues()
//     var nameL = dataValidationList.map(function(e){return e[0]})
//     var urlL = dataValidationList.map(function(e){return e[1]})
//     var optionL = dataValidationList.map(function(e){return e[2]})
//     nameL.shift(); urlL.shift();optionL.shift()
    
//     var pointer = s.getActiveCell()
//     //Create Data Validation when source file input exists
//     if(pointer.getColumn() == startColIp && pointer.getRow()>=startRowB) {
//       var nextCell = pointer.offset(0,startColOp-startColIp)
//       switch (pointer.getValue())
//       {
//         case '':
//           nextCell.getDataRegion(SpreadsheetApp.Dimension.COLUMNS).clearDataValidations().clear()
// //          .clearDataValidations() // Remove any existing data validation
// //          .clearContent(); // Clear the cell
//           break;
//         default:
//           nextCell
//           .clearDataValidations() // Remove any existing data validation
//           .clearContent(); // Clear the cell
//           var rule = SpreadsheetApp.newDataValidation().requireValueInRange(s.getRange(columnToLetter(opLCol)+startRowB+':'+columnToLetter(opLCol)),true).build();
//           nextCell.setDataValidation(rule);
//           break;
//       }
      
//     }
//     //Data Validation for output file
//     if( pointer.getColumn() == startColOp && pointer.getRow()>=startRowB) { 
//       var nextCell1 = pointer.offset(0,1)
//       var nextCell2 = pointer.offset(0,3)
//       switch (nameL.indexOf(pointer.getValue())){
//         case -1: 
//         nextCell1.getDataRegion(SpreadsheetApp.Dimension.COLUMNS).clearDataValidations().clear()
//         break;
//         default:
//         nextCell1.setValue(urlL[nameL.indexOf(pointer.getValue())]);
//         nextCell2
//           .clearDataValidations() // Remove any existing data validation
//           .clearContent(); // Clear the cell
//         var rule = SpreadsheetApp.newDataValidation().requireValueInList(optionL[nameL.indexOf(pointer.getValue())].split(','), true).build();
//         nextCell2.setDataValidation(rule);  
//         break;
//       }
//     }
//   }
// }

//ON OPEN
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Menu')
  .addItem('Run Dashboard Control Panel', 'displaySidebar')
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Frontend CO')
          .addItem('Initialize Update Metrics CO', 'runCOinit')
          .addItem('Update Metrics CO ca', 'runCOca')
          .addItem('Update Metrics CO uk', 'runCOuk')
          .addItem('Update Metrics CO au', 'runCOau')
          .addItem('Update Metrics CO ie', 'runCOie')
          .addItem('Update Metrics CO nz', 'runCOnz')
          .addItem('Update Metrics CO it', 'runCOit')
          .addItem('Finalize Metrics CO', 'runCOfinal')
  .addSeparator()
          .addItem('Update Content CO ca', 'updateCOca')
          .addItem('Update Content CO uk', 'updateCOuk')
          .addItem('Update Content CO au', 'updateCOau')
          .addItem('Update Content CO ie', 'updateCOie')
          .addItem('Update Content CO nz', 'updateCOnz')
          .addItem('Update Content CO it', 'updateCOit')
  .addSeparator()
          .addItem('Remove Format CO', 'removeFormatCO')
          .addItem('Apply Format CO', 'applyFormatCO'))
  // .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Backend')
  //         .addItem('Update WFM', 'updateWFM')
  //         .addItem('Update QA', 'updateQA')
  //         .addItem('Update Chats Live','updateChat')
  //         .addItem('Update Emails','updateEmail'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Frontend CS')
          .addItem('Initialize Update Metrics CS', 'runCSinit')
          .addItem('Update Metrics CS ca', 'runCSca')
          .addItem('Update Metrics CS uk', 'runCSuk')
          .addItem('Update Metrics CS au', 'runCSau')
          .addItem('Update Metrics CS ie', 'runCSie')
          .addItem('Update Metrics CS nz', 'runCSnz')
          .addItem('Update Metrics CS it', 'runCSit')
          .addItem('Finalize Metrics CS', 'runCSfinal')
  .addSeparator()
          .addItem('Update Content CS ca', 'updateCSca')
          .addItem('Update Content CS uk', 'updateCSuk')
          .addItem('Update Content CS au', 'updateCSau')
          .addItem('Update Content CS ie', 'updateCSie')
          .addItem('Update Content CS nz', 'updateCSnz')
          .addItem('Update Content CS it', 'updateCSit')
  .addSeparator()
          .addItem('Remove Format CS', 'removeFormatCS')
          .addItem('Apply Format CS', 'applyFormatCS'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Frontend CC')
          .addItem('Update Metrics CC', 'runCC')
          .addItem('Update Content CC', 'updateCC')
          .addItem('Remove Format CC', 'removeFormatCC')
          .addItem('Apply Format CC', 'applyFormatCC'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Frontend Reason CS')
          .addItem('Initialize Update Metrics Reason CS', 'runReasonCSinit')
          .addItem('Update Metrics Reason CS overall', 'runReasonCSoverall')
          .addItem('Update Metrics Reason CS ca', 'runReasonCSca')
          .addItem('Update Metrics Reason CS uk', 'runReasonCSuk')
          .addItem('Update Metrics Reason CS au', 'runReasonCSau')
          .addItem('Update Metrics Reason CS ie', 'runReasonCSie')
          .addItem('Update Metrics Reason CS nz', 'runReasonCSnz')
          .addItem('Update Metrics Reason CS it', 'runReasonCSit')
          .addItem('Finalize Metrics Reason CS', 'runReasonCSfinal')
  .addSeparator()
          .addItem('Update Content Reason CS overall', 'updateReasonCSoverall')
          .addItem('Update Content Reason CS ca', 'updateReasonCSca')
          .addItem('Update Content Reason CS uk', 'updateReasonCSuk')
          .addItem('Update Content Reason CS au', 'updateReasonCSau')
          .addItem('Update Content Reason CS ie', 'updateReasonCSie')
          .addItem('Update Content Reason CS nz', 'updateReasonCSnz')
          .addItem('Update Content Reason CS it', 'updateReasonCSit')
  .addSeparator()
          .addItem('Remove Format Reason CS', 'removeFormatReasonCS')
          .addItem('Apply Format Reason CS', 'applyFormatReasonCS')
          )
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Frontend Reason CO')
          .addItem('Initialize Update Metrics Reason CO', 'runReasonCOinit')
          .addItem('Update Metrics Reason CO overall', 'runReasonCOoverall')
          .addItem('Update Metrics Reason CO ca', 'runReasonCOca')
          .addItem('Update Metrics Reason CO uk', 'runReasonCOuk')
          .addItem('Update Metrics Reason CO au', 'runReasonCOau')
          .addItem('Update Metrics Reason CO ie', 'runReasonCOie')
          .addItem('Update Metrics Reason CO nz', 'runReasonCOnz')
          .addItem('Update Metrics Reason CO it', 'runReasonCOit')
          .addItem('Finalize Metrics Reason CO', 'runReasonCOfinal')
  .addSeparator()
          .addItem('Update Content Reason CO overall', 'updateReasonCOoverall')
          .addItem('Update Content Reason CO ca', 'updateReasonCOca')
          .addItem('Update Content Reason CO uk', 'updateReasonCOuk')
          .addItem('Update Content Reason CO au', 'updateReasonCOau')
          .addItem('Update Content Reason CO ie', 'updateReasonCOie')
          .addItem('Update Content Reason CO nz', 'updateReasonCOnz')
          .addItem('Update Content Reason CO it', 'updateReasonCOit')
  .addSeparator()
          .addItem('Remove Format Reason CO', 'removeFormatReasonCO')
          .addItem('Apply Format CO', 'applyFormatReasonCO')
          )
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Overall Propensity')
          .addItem('Update Metrics','runOverallPropensity')
      .addSeparator()
          .addItem('Update Content','updateOverallPropensity')
  )
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Update Overall Reasons')
          .addItem('Update Metrics','runOverallReason')
      .addSeparator()
          .addItem('Update Content','updateOverallReason')
  )
          .addToUi();
}

function updateWFM(){
  var wfmList = ['OV Forecast','Forecasted Hours','Adherence','Network Logistics Metrics for LC']
  wfmList.forEach(function(doc){updateBackEnd(doc)})
}
function updateQA(){
  var wfmList = ['QA Team Level']
  wfmList.forEach(function(doc){updateBackEnd(doc)})
}
function updateChat(){
  var wfmList = ['Chats Live']
  wfmList.forEach(function(doc){updateBackEnd(doc)})
}
function updateEmail(){
  var wfmList = ['Emails LS']
  wfmList.forEach(function(doc){updateBackEnd(doc)})
}

