//I created this function to be used with AppScript and Google Spreadsheet. 
//For the project I used it in, it was part of an expense traker/budget for an event planning agency
//https://www.youtube.com/playlist?list=PLv9Pf9aNgemviJmKkuOQyBd5uOmxMkcpB

function onEdit(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // this variable selects the sheet with data that will populate the dropdown lists
  var datass = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("tags");
 
  var activeCell = ss.getActiveCell();
  
  // when plugging this into another project, change the .getColumn to equal the correct corresponding number ie A = 1; P = 16
  if(activeCell.getColumn() == 1 && activeCell.getRow() > 1){ 
  
  // this clears content and data validations in child dropdown when parent dropdown is edited
    activeCell.offset(0,1).clearContent().clearDataValidations();
    
  // this variable populates the parent dropdown use data validation set range and hightlight the first row including blank columns
    var category = datass.getRange(1, 1, 1, datass.getLastColumn()).getValues();
    
  // this changes makes it so the index of the category matches the numeric value of the column  
    var categoryIndex = category[0].indexOf(activeCell.getValue()) + 1;   
    
    if(categoryIndex != 0) {
    
  // this grabs the tags from the selected category on parent drop down to be used as the options in child dropdown  
    var validationRange = datass.getRange(3, categoryIndex, datass.getLastRow());
  
  // this creates the child dropdown and offsets it from the parent cell by 0 rows and 1 column
    var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
    activeCell.offset(0, 1).setDataValidation(validationRule);

    }    
  }
}
