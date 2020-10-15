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

const usedValueArray = [];

const storedValues = {};

function renderValues(){
  var HBsheet = getHBSheet();
  for (let i = 0; i < usedValueArray.length; i++){
    var cellNum = i + 2;
    HBsheet.getRange('A' + cellNum).setValue(usedValueArray[i]);
    HBsheet.getRange('B' + cellNum).setValue(storedValues[usedValueArray[i]]);
  }
  
  HBsheet.autoResizeColumns(1,2);
}

function updateValues(name, val){
  usedValueArray.push(name);
  storedValues[name] = val;
  renderValues();
}

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
  var altmanValues = ["Working Capital", 
                "Total Assets", 
                "Retained Earnings", 
                "EBIT", 
                "Market Value Of Equity", 
                "Total Liabilities", 
                "Sales"]
  var HBsheet = getHBSheet();
  var ui = SpreadsheetApp.getUi();
  var currCell = SpreadsheetApp.getActiveSpreadsheet().getSelection().getCurrentCell()
  
  altmanValues.forEach((val) => {
    Logger.log(usedValueArray);
    if (!usedValueArray.includes(val)){
        var result = ui.prompt(
        'We Need Some Values',
        'Please enter the ' + val + ':',
        ui.ButtonSet.OK);
        var num = result.getResponseText();
        updateValues(val, num);
    }
  });

  var A = parseInt(storedValues["Working Capital"]/storedValues["Total Assets"])
  var B = parseInt(storedValues["Retained Earnings"]/storedValues["Total Assets"])
  var C = parseInt(storedValues["EBIT"]/storedValues["Total Assets"])
  var D = parseInt(storedValues["Market Value Of Equity"]/storedValues["Total Liabilities"])
  var E = parseInt(storedValues["Sales"]/storedValues["Total Assets"])
  
  var ZScore = (1.2*A) + (1.4*B) + (3.3*C) + (0.6*D) + (0.99*E);
  
  currCell.setValue(ZScore).setFontWeight('bold').setBackground('#ffcccb');
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