//----------------------------------------------
//----------------Code by Rauxd-----------------
//----------------------------------------------

//----------------------------------------------
//Read me
//----------------------------------------------
/*This code is one dynamic menu(for multiopcion list), my project check only 1 column and aplicate in other column in the
same row, you are free the modification, i dont know resolve 1 problem of compatibility in the line 80, that's why my code
only verifies the column selected(my cases is x), but you can adapt my code in your project*/


//**********************************************
//make dynamic menu
//**********************************************
  function DynamicMenu(valueCellsX, menuCellsY) {
//--------------------------------
//Activation the sheets projects
//--------------------------------
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard"); //get activation in my dashboard
  var valueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Value") //activation my sheet value, if you not second sheet, delete this code (alone line)

//-------------------------------------------
//find value in the cell and de range cells
//-------------------------------------------
  var valueOptionsCase1 = valueSheet.getRange("E2:E4"); // Range containing case 1
  var optionsCase1 = valueOptionsCase1.getValues().flat(); // Get the values and flatten the array

  var valueOptionsCase2 = valueSheet.getRange("G2:G5"); // Range containing case 2
  var optionsCase2 = valueOptionsCase2.getValues().flat(); // Get the values and flatten the array

  //comprovating cells

  var cell = sheet.getRange(valueCellsX);
  var value = cell.getValue();
//----------------------------------------------------------
//Cases, if you need more, add more if, or change for swich
//----------------------------------------------------------
  //case 1

  if(value == "Ingreso"){  
  var cell = sheet.getRange(menuCellsY); 
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(optionsCase1, true)
    .build();
  cell.setDataValidation(rule);
  }

  //case 2

  else{
  var cell = sheet.getRange(menuCellsY); 
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(optionsCase2, true)
    .build();
  cell.setDataValidation(rule);
  }
  }
//**********************************************
//Listern of events modification
//**********************************************
function onEdit(e) {
//----------------------------------------------
//Variables
//----------------------------------------------
  var range = e.range; // capture range cell modified
  var notationNumberComp = range.getA1Notation(); // capture specific cell modified
  let notationNumber = DecomposeString(notationNumberComp); // convert cell in array 
  var letterSelector = "X"; //letter do search (you can change for letter in the colunm)
  let numMin = "8"; //set number min. (this variable no effects for compatibility problems)
  var totalElementsNN = notationNumber.length; //Contain at number total for element in the array

//----------------------------------------------
//Transformation elements for number int
//----------------------------------------------
  if(totalElementsNN == 2){
    var containNumber = notationNumber[1];
    }
  if(totalElementsNN == 3){
    var containNumber = notationNumber[1] + notationNumber [2];
    }
  if(totalElementsNN == 4){
    var containNumber = notationNumber[1] + notationNumber [2] + notationNumber[3];
    }
//--------------------------------------------------
//Conditional for activation a function DynamicMenu
//--------------------------------------------------
  if(letterSelector === notationNumber[0]){ // select de colunm X
      //how to fucking comparation the numbers?????????????????
      DynamicMenu(notationNumberComp, "Y"+containNumber);
      Debug(holi, pvto);
  }
   }
  
//**********************************************
//Descompose string to array
//**********************************************
function DecomposeString(str) {
  let characters = []; // Create an empty array to store characters

  for (let i = 0; i < str.length; i++) {
    characters.push(str.charAt(i)); // Add each character to the array
  }

  return characters; // Return the array of characters
  }

//**********************************************
//Debug
//**********************************************
function Debug(dev, dev2){
  var valueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Value")
  valueSheet.getRange("J3").setValue(dev); //propieties number 1, for debug in onEdit(e)
  valueSheet.getRange("K3").setValue(dev2); //propieties number 2,for debug in onEdit(e)
  //when you use a method onEdit(e), you loss debug, for a method work in background, you cant use logger.log 
  }
//**********************************************
//run
//**********************************************
function Run(){ //Is Debugger for functions
  DynamicMenu('X10', 'Y10')
 }