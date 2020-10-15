/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var mainMenu = ui.createMenu("HB Accounting");
  mainMenu.addItem("Generate Project", "generateSheet");
  mainMenu.addSeparator();
  mainMenu.addItem("Execute Function", "showDialog");
  mainMenu.addToUi();
}

function getHBSheet () {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HB");
}

const valueArray = [];

function generateSheet() {
  // Set Vars
  var sheet = SpreadsheetApp.getActiveSheet().setName('HB');
  var headers = ['Accounting Values']
  var initialData = [];
  var valueNamesCol = sheet.getRange('A2:A100');
  var valuesCol = sheet.getRange('B2:B100');
  
  // Styles
  sheet.getRange('A1').setValues([headers]).setFontWeight('bold');
  
  valueNamesCol.setFontWeight('bold')
  valueNamesCol.setBorder(true, true, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID)
  valuesCol.setBorder(true, false, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID)
  for (let i = 2; i < 101; i++){
    if (i % 2 === 0) {
      sheet.getRange('A'+ i).setBackground('#add8e6');
      sheet.getRange('B'+ i).setBackground('#add8e6');
    }
  }
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('functions')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Select Your Function');
}

function altman(){
  
  // Vars
  var HBsheet = getHBSheet();
  var ui = SpreadsheetApp.getUi();
  var currCell = HBsheet.getActiveCell()
  
  // Prompts
  var result = ui.prompt(
    'Let\'s get to know each other!',
    'Please enter your name:',
    ui.ButtonSet.OK);
  var text = result.getResponseText();
  
  currCell.setValue(text);
}

function execute(name) {
  switch (name) {
  case "altman":
    altman();
    break;
  default:
    break;
  }
}