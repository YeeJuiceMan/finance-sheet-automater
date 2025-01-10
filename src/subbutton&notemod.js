function test(){
  //basic sheets var
    const ss = SpreadsheetApp.getActiveSpreadsheet(),
    s = SpreadsheetApp.getActiveSheet(),
    activeCell = ss.getActiveCell(),
    reference = activeCell.getA1Notation(),
    activeVal = activeCell.getValue(),
    refArr = reference.split(''),
    consoleSheet = ss.getSheetByName("Console"),
    usSheet = ss.getSheetByName("College Savings 3.0"),
    twSheet = ss.getSheetByName("College Savings 3.0 (TW)"),
    usSpecSheet = ss.getSheetByName("College Savings 3.0 Specifics"),
    usSpecSheetHideMenu = ss.getSheetByName("College Savings 3.0 Specifics Hide Menu"),

  //out var
    typeSheetOut = consoleSheet.getRange("B2:C3"),
    errorMsgOut = consoleSheet.getRange("B24"),
    checkOrResOut = consoleSheet.getRange("C5:C6"),
    needOrWantOrReimb = consoleSheet.getRange("C7:C8"),
    expenseType = consoleSheet.getRange("C9:C10"),
    amountOut = consoleSheet.getRange("C11:C12"),
    expenseNoteType = consoleSheet.getRange("C13:C14"),
    newExpenseNoteType = consoleSheet.getRange("C15:C16"),
    usDayVal = consoleSheet.getRange("C22"),
    twDayVal = consoleSheet.getRange("C24"),

  //in var
    typeSheetIn = consoleSheet.getRange("E2:F3"),
    errorMsgIn = consoleSheet.getRange("E18"),
    checkOrResIn = consoleSheet.getRange("F5:F6"),
    fixedOrNot = consoleSheet.getRange("F7:F8"),
    amountIn = consoleSheet.getRange("F9:F10"),
    incomeNoteType = consoleSheet.getRange("F11:F12"),
    newIncomeNoteType = consoleSheet.getRange("F13:F14"),

  //reimb var
    typeSheetReimb = consoleSheet.getRange("H2:I3"),
    errorMsgReimb = consoleSheet.getRange("H16"),
    year = consoleSheet.getRange("I5:I6"),
    month = consoleSheet.getRange("I7:I8"),
    nonReimbCell = consoleSheet.getRange("I11:I12"),
    specRow = consoleSheet.getRange("H4");
  
  //us normal col
    const resFixedCol = 4,
    resNonFixedCol = 5,
    resReimbIn = 6,
    checkInCol = 7,
    checkReimbIn = 8,
    resOutCol = 9,
    resReimbOut = 10,
    needTrans = 11,
    needFood = 12,
    needGroc = 13,
    needOn = 14,
    needItem = 15,
    needMisc = 16,
    wantTrans = 17,
    wantFood = 18,
    wantGroc = 19,
    wantOn = 20,
    wantItem = 21,
    wantMisc = 22,
    checkReimbOut = 23;

  //us spec col
    const resFixedColSpec = 4,
    resNonFixedColSpec = 9,
    resReimbInSpec = 14,
    checkInColSpec = 19,
    checkReimbInSpec = 24,
    resOutColSpec = 29,
    resReimbOutSPec = 34,
    needTransSpec = 40,
    needFoodSpec = 45,
    needGrocSpec = 50,
    needOnSpec = 55,
    needItemSpec = 60,
    needMiscSpec = 65,
    wantTransSpec = 70,
    wantFoodSpec = 75,
    wantGrocSpec = 80,
    wantOnSpec = 85,
    wantItemSpec = 90,
    wantMiscSpec = 95,
    checkReimbOutSpec = 100;

  var mo = 0, yr = 2025;
  var date = new Date(yr, mo);
  var specColNum = 4;

  // addEntryRow(today, 5, 76, usSpecSheet, usSpecSheetHideMenu);
  for (var normColNum = 4; normColNum < 24; normColNum++) {
    var reimbNo = false;

    if (normColNum == 10|| normColNum == 23) reimbNo = true;
    customNoteToSheets(date, usSheet, usSpecSheet, usSpecSheetHideMenu, mo + 32, normColNum, reimbNo, 5, specColNum, specColNum + 4);
    if (specColNum == 34) specColNum += 6;
    else specColNum += 5;
  }
  // subButtonAct(checkOrResOut, needOrWantOrReimb, expenseType, amountOut, expenseNoteType, newExpenseNoteType, usDayVal, usSheet, usSpecSheet, usSpecSheetHideMenu);
}


//Print all notes cost and exp type (given it isn't empty) to spec sheet
function customNoteToSheets(date, typeSheet, specSheet, hideSheet, addRow, addCol, reimbOrNot, monthEndRowsListCol, ccolWithBrokeDownCost, ccolWithReimbMark) {

  //other cols relative to first one (reimb is exception)
  var ccolWithExpTotCost = ccolWithBrokeDownCost + 1;
  var ccolWithExpTypeNames = ccolWithBrokeDownCost + 2;
  var ccolWithCardType = ccolWithBrokeDownCost + 3;

  var rangeArr = findSpecMonthRange(hideSheet, date, monthEndRowsListCol);
  var startRow = rangeArr[0];
  var lastRow = rangeArr[1];
  var totalMonthLen = rangeArr[2];
  var monthRow = rangeArr[3];

  //Print all notes cost and exp type (given it isn't empty)
  var notes = typeSheet.getRange(addRow, addCol).getNotes().toString().split("\n");
  var noteInd = 0;
  var sheetInd = startRow;

  //Finding real notes length and checking if it is greater than the max
  var findNoteLengthInd = 0;
  var actualNoteLength = 0;
  while (findNoteLengthInd < notes.length) {
    if (notes[findNoteLengthInd].length > 0) actualNoteLength++;
    findNoteLengthInd++;
  }

  //checks if enough space for note entries; extend if not
  if (actualNoteLength > totalMonthLen) {
    while (actualNoteLength > totalMonthLen) {
      addEntryRow(date, 5, 104, specSheet, hideSheet);
      lastRow = hideSheet.getRange(monthRow, monthEndRowsListCol).getValue();
      totalMonthLen = lastRow - startRow + 1;
      Logger.log("extended w/ month len " + totalMonthLen + " note len " + actualNoteLength);
    }
  }

  while (noteInd < notes.length && sheetInd <= lastRow) {
    Logger.log(notes[noteInd]);
    if (notes[noteInd].length > 0) { //line isn't empty (handling user typo error)
      var tempEntry = notes[noteInd].split(": ");

      //split the total cost & equation w/o tilde (if exists)

      var tempCostEntry = tempEntry[0].replace(")", "").replace("~", "").split(" (")
      Logger.log("total cost " + tempCostEntry);
      var tempFormulaEntry = tempCostEntry[0].replace(/@([A-Z]+\w*)/, "").trim();
      Logger.log("formula " + tempFormulaEntry);

      // find credit card info (if exists)
      var tempCardRegex = new RegExp("@([A-Z]+\w*)");
      var tempCardEntry = tempCardRegex.exec(tempCostEntry[0]);
      Logger.log("card " + tempCardEntry);

      //if note in reimb, split the tilde (if it exists)
      var tildeCheck = tempEntry[0].split(" ")[0];
      if (reimbOrNot == true) {
        if (tildeCheck == "~") {
          //tilde; hence not reimbed
          specSheet.getRange(sheetInd, ccolWithReimbMark).setValue(false);
        } 
        else {
          //not in reimb column or alr reimbed
          specSheet.getRange(sheetInd, ccolWithReimbMark).setValue(true);
        }
      }

      //put values in respective columns
      if (tempFormulaEntry != null) specSheet.getRange(sheetInd, ccolWithBrokeDownCost).setValue(tempFormulaEntry);

      //no formula exists (1 cost)
      if (tempCostEntry[1] == null && tempFormulaEntry != null) {
        specSheet.getRange(sheetInd, ccolWithExpTotCost).setValue(tempFormulaEntry)
      }
      else {
        //a formula exists
        specSheet.getRange(sheetInd, ccolWithExpTotCost).setValue(tempCostEntry[1])
      }
      specSheet.getRange(sheetInd, ccolWithExpTypeNames).setValue(tempEntry[1]);

      //card conditions
      if (tempCardEntry != null) {
        switch (tempCardEntry[0]) {
          case "@D":
            tempCardEntry = "DISCOVER";
            break;
          case "@B":
            tempCardEntry = "BILT";
            break;
          case "@C":
            tempCardEntry = "CHASE FU";
            break;
          case "@A":
            tempCardEntry = "CHASE AP";
            break;
          case "@C1":
            tempCardEntry = "CAPONE SAVOR";
            break;
          default:
            tempCardEntry = "N/A";
            break;
        }
      }
      else tempCardEntry = "N/A"

      specSheet.getRange(sheetInd, ccolWithCardType).setValue(tempCardEntry);
      noteInd++;
      sheetInd++;
    } 
    else {
      //skip empty lines
      noteInd++;
    }
  }
  return;
}


//note cols
const rowThatDropdownSheetStarts = 4, // for notes
colWithBrokeDownCost = 36, //for notes
colWithExpTotCost = 37, //for notes
colWithExpTypeNames = 38, //for notes
colWithReimbMark = 39, //for notes

//normal sheet cols
resFixedCol = 4,
resNonFixedCol = 5,
resReimbInCol = 6,
checkInCol = 7,
checkReimbInCol = 8,
resOutCol = 9,
resReimbOutCol = 10,
needStart = 11,
needEnd = 16,
wantStart = 17,
wantEnd = 22,
checkReimbOutCol = 23,

//spec sheet cols
resFixedColSpec = 3,
resNonFixedColSpec = 8,
resReimbInColSpec = 13,
checkInColSpec = 18,
checkReimbInColSpec = 23,
resOutColSpec = 28,
resReimbOutColSpec = 33,
needStartSpec = 39,
needEndSpec = 68,
wantStartSpec = 69,
wantEndSpec = 98,
checkReimbOutColSpec = 99;

//sheet vars
const spreadSheetConfig = {
  get spreadsheet() {
    delete this.spreadsheet;
    return (this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet()); //running under the assumption that this is a bound script
  },
};
const mainSpreadSheet = spreadSheetConfig.spreadsheet;
const consoleSheet = mainSpreadSheet.getSheetByName("Console"),
usSheet = mainSpreadSheet.getSheetByName("College Savings 3.0"),
twSheet = mainSpreadSheet.getSheetByName("College Savings 3.0 (TW)"),
usSpecSheet = mainSpreadSheet.getSheetByName("College Savings 3.0 Specifics"),
usSpecSheetHideMenu = mainSpreadSheet.getSheetByName("College Savings 3.0 Specifics Hide Menu"),
twSpecSheet = mainSpreadSheet.getSheetByName("College Savings 3.0 (TW) Specifics"),
twSpecSheetHideMenu = mainSpreadSheet.getSheetByName("College Savings 3.0 (TW) Specifics Hide Menu"),

//out var
typeSheetOut = consoleSheet.getRange("B2:C3"),
errorMsgOut = consoleSheet.getRange("B26"),
checkOrResOut = consoleSheet.getRange("C5:C6"),
needOrWantOrReimb = consoleSheet.getRange("C7:C8"),
expenseType = consoleSheet.getRange("C9:C10"),
amountOut = consoleSheet.getRange("C11:C12"),
expenseNoteType = consoleSheet.getRange("C13:C14"),
newExpenseNoteType = consoleSheet.getRange("C15:C16"),
outCreditType = consoleSheet.getRange("C17:C18"),
usDayVal = consoleSheet.getRange("C22"),
twDayVal = consoleSheet.getRange("C24"),

//in var
typeSheetIn = consoleSheet.getRange("E2:F3"),
errorMsgIn = consoleSheet.getRange("E18"),
checkOrResIn = consoleSheet.getRange("F5:F6"),
fixedOrNot = consoleSheet.getRange("F7:F8"),
amountIn = consoleSheet.getRange("F9:F10"),
incomeNoteType = consoleSheet.getRange("F11:F12"),
newIncomeNoteType = consoleSheet.getRange("F13:F14"),
inCreditType = consoleSheet.getRange("F15:F16"),

//reimb var
typeSheetReimb = consoleSheet.getRange("H2:I3"),
errorMsgReimb = consoleSheet.getRange("H16"),
reimbYear = consoleSheet.getRange("I5:I6"),
reimbMonth = consoleSheet.getRange("I7:I8"),
checkOrResReimb = consoleSheet.getRange("I9:I10"),
nonReimbCell = consoleSheet.getRange("I11:I12"),

//extras
usMonthEndRowListCol = 6,
usCategoryEndColListCol = 12,
twMonthEndRowListCol = 6;


function onEdit(e) {
  if (!e) {
    throw new Error(
      'Please do not run the onEdit(e) function in the script editor window. '
      + 'It runs automatically when you hand edit the spreadsheet. '
      + 'See https://stackoverflow.com/a/63851123/13045193.'
    );
  }
  try {
    onButtonTrigger(e);
  }
  catch (error) {
    errorMsgOut.setValue("Error: " + error);
    errorMsgOut.setBackground("#e06666");
    errorMsgIn.setValue("Error: " + error);
    errorMsgIn.setBackground("#e06666");
    errorMsgReimb.setValue("Error: " + error);
    errorMsgReimb.setBackground("#e06666");
}


function onButtonTrigger(e) {
  //basic event sheets var
    const activeCell = e.range,
    reference = activeCell.getA1Notation(),
    activeVal = activeCell.getValue(),
    activeSheetName = e.source.getActiveSheet().getName();

  //for spec hide menu
  if (reference[0] == "D" || reference[0] == "E"){ //hide month(s) buttons
    if (activeSheetName == usSpecSheetHideMenu.getName())
      entryHiding(activeCell, activeVal, usSpecSheetHideMenu, usSpecSheet, usMonthEndRowListCol, "row");
  }
  else if (reference[0] == "J" || reference[0] == "K"){ //hide category(s) buttons
    if (activeSheetName == usSpecSheetHideMenu.getName())
      entryHiding(activeCell, activeVal, usSpecSheetHideMenu, usSpecSheet, usCategoryEndColListCol, "col");
  }

  /*
  B20 = BUTTON RED (OUT)
  C20 = BUTTON GREEN (OUT)

  E18 = BUTTON RED (IN)
  F18 = BUTTON GREEN (IN)

  H14 = BUTTON RED (REIMB)
  I14 = BUTTON GREEN (REIMB)
  */
  //console buttons
  if (activeVal == true && activeSheetName == consoleSheet.getName()) {
    switch (reference){
      case "B20": //red out
        errorMsgOut.setValue("...");
        errorMsgOut.setBackground("#fbbc04");

        //if there's no input of money or the amountOut is 0 do nothing return error msg
        if (amountOut.getValue().length <= 0 || amountOut.getValue() == 0) {
          errorMsgOut.setValue("There are missing cost inputs!");
          errorMsgOut.setBackground("#e06666");
          activeCell.setValue(false);
          return;
        }

        if (typeSheetOut.getValue() == "US") {
          subButtonAct(usDayVal, usMonthEndRowListCol, usSheet, usSpecSheet, usSpecSheetHideMenu);
        } else if (typeSheetOut.getValue() == "TW") { //will be changed later
          subButtonAct(twDayVal, twMonthEndRowListCol, twSheet, twSpecSheet, twSpecSheetHideMenu);
        }

        errorMsgOut.setValue("Successfully added " + typeSheetOut.getValue() + " $" + amountOut.getValue() + ". Please input notes & press Green to continue.");
        errorMsgOut.setBackground("#f6b26b");
        activeCell.setValue(false);
        return;

      case "C20": //green out
        errorMsgOut.setValue("...");
        errorMsgOut.setBackground("#fbbc04");

        //if there's no input in exp note type & N/A in dropdown return error msg
        if (expenseNoteType.getValue() == "N/A" && newExpenseNoteType.getValue().length == 0) {
          errorMsgOut.setValue("There are missing note inputs!");
          errorMsgOut.setBackground("#e06666");
          activeCell.setValue(false);
          return;
        }

        if (typeSheetOut.getValue() == "US") {
          //setting up spec sheet modding (monthEndRowListCol is 5 in spec)
          subModSpecSheet(new Date(), usMonthEndRowListCol, usSpecSheet, usSpecSheetHideMenu);
        } else if (typeSheetOut.getValue() == "TW") { // will be changed later
          subModSpecSheet(new Date(), twMonthEndRowListCol, twSpecSheet, twSpecSheetHideMenu);
        }

        errorMsgOut.setValue("Specifics added to " + typeSheetOut.getValue() + ".");
        errorMsgOut.setBackground("#93c47d");
        activeCell.setValue(false);
        return;

      case "E18": //red in
        errorMsgIn.setValue("...");
        errorMsgIn.setBackground("#fbbc04");

        //if there's no input of money or the amountOut is 0 do nothing return error msg
        if (amountIn.getValue().length <= 0 || amountIn.getValue() == 0) {
          errorMsgIn.setValue("There are missing income inputs!");
          errorMsgIn.setBackground("#e06666");
          activeCell.setValue(false);
          return;
        }

        if (typeSheetIn.getValue() == "US") {
          addButtonAct(usMonthEndRowListCol, usSheet, usSpecSheet, usSpecSheetHideMenu);
        } else if (typeSheetIn.getValue() == "TW") {
          addButtonAct(twMonthEndRowListCol, twSheet, twSpecSheet, twSpecSheetHideMenu);
        }

        errorMsgIn.setValue("Successfully added " + typeSheetIn.getValue() + " $" + amountIn.getValue() + ". Please input notes & press Green to continue.");
        errorMsgIn.setBackground("#f6b26b");
        activeCell.setValue(false);
        return;

      case "F18": //green in
        activeCell.setValue(false);
        errorMsgIn.setValue("...");
        errorMsgIn.setBackground("#fbbc04");

        //if there's no input in exp note type & N/A in dropdown return error msg
        if (incomeNoteType.getValue() == "N/A" && newIncomeNoteType.getValue().length == 0) {
          errorMsgIn.setValue("There are missing note inputs!");
          errorMsgIn.setBackground("#e06666");
          activeCell.setValue(false);
          return;
        }

        if (typeSheetOut.getValue() == "US") {
          inNoteMod(checkOrResIn, fixedOrNot, amountIn, incomeNoteType, newIncomeNoteType, usSheet)
        } else if (typeSheetOut.getValue() == "TW") {
          inNoteMod(checkOrResIn, fixedOrNot, amountIn, incomeNoteType, newIncomeNoteType, twSheet)
        }

        errorMsgIn.setValue("Notes added to " + typeSheetOut.getValue() + ".");
        errorMsgIn.setBackground("#93c47d");
        activeCell.setValue(false);
        return;

      case "H14": //red reimb
        errorMsgReimb.setValue("...");
        errorMsgReimb.setBackground("#fbbc04");
        var needReimb;

        if (typeSheetReimb.getValue() == "US") {
          needReimb = checkReimb(usMonthEndRowListCol, usSpecSheet, usSpecSheetHideMenu)
        } else if (typeSheetReimb.getValue() == "TW") {
          needReimb = checkReimb(twMonthEndRowListCol, twSpecSheet, twSpecSheetHideMenu)
        }

        if (needReimb == true) {
          errorMsgReimb.setValue("There are items in need of reimb.");
          errorMsgReimb.setBackground("#f6b26b");
        } else {
          errorMsgReimb.setValue("There are no items to reimb.");
          errorMsgReimb.setBackground("#93c47d");
        }
        activeCell.setValue(false);
        return;

      case "I14": //green reimb
        errorMsgReimb.setValue("...");
        errorMsgReimb.setBackground("#fbbc04");

        if (typeSheetReimb.getValue() == "US") {
          alrReimbedNoteMod(year, month, nonReimbCell, usSheet);
        } else if (typeSheetReimb.getValue() == "TW") {
          alrReimbedNoteMod(year, month, nonReimbCell, twSheet);
        }

        errorMsgReimb.setValue("Item reimbed.");
        errorMsgReimb.setBackground("#93c47d");
        activeCell.setValue(false);
        return;

      default: //extra button conditions
        activeCell.setValue(false);
        return;
    }
  }
  return;
}

//----------button action----------//

//adds out val to chosen cell given parameters
function subButtonAct(dayVal, monthEndRowListCol, typeSheet, specSheet, hideSheet) {

  errorMsgOut.setValue("Finding rows...");

  let today = new Date();
  let addRow = findAddRow(typeSheet, today),
  addCol,
  addColSpec, //for dropdown list
  expenseTypeVal = expenseType.getValue(),
  needOrWantOrReimbVal = needOrWantOrReimb.getValue();
  if (needOrWantOrReimbVal == "REIMB") needOrWantOrReimbVal = "REIMB OUT";

  errorMsgOut.setValue("Finding columns...");

  //RES
  if (checkOrResOut.getValue() == "RES") {
    if (needOrWantOrReimb == "REIMB OUT") {
      addCol = findAddCol(typeSheet, expenseTypeVal, "REIMB OUT", "RES", "type");
      addColSpec = findAddCol(specSheet, expenseTypeVal, "REIMB OUT", "RES", "spec") + 3; //by default settles on date col
    }
    else {
      needOrWantOrReimb.setBackground("#999999");
      addCol = findAddCol(typeSheet, expenseTypeVal, "OUT", "RES", "type");
      addColSpec = findAddCol(specSheet, expenseTypeVal, "OUT", "RES", "spec") + 3; //by default settles on date col
    }
    expenseType.setBackground("#999999");
  }

  //CHECK
  else {
    needOrWantOrReimb.setBackground("#cccccc");
    if (needOrWantOrReimbVal == "REIMB OUT") {
      expenseType.setBackground("#999999");
    }
    else expenseType.setBackground("#cccccc");

    //find col of targetted cell given N/W/R & exp type
    addCol = findAddCol(typeSheet, expenseTypeVal, needOrWantOrReimbVal, "CHECK", "type");
    addColSpec = findAddCol(specSheet, expenseTypeVal, needOrWantOrReimbVal, "CHECK", "spec") + 3; //by default settles on date col

    //add daily val given it isn't reimb (daily expenses that is)
    var curDailyVal = dayVal.getValue();
    if (needOrWantOrReimbVal != "REIMB OUT") dayVal.setValue("=" + curDailyVal + "+" + amountOut.getValue());
  }

  errorMsgOut.setValue("Adding amount...");
  addMoney(addRow, addCol, amountOut.getValue(), typeSheet);

  //vars for dropdown
  errorMsgOut.setValue("Finding spec sheet month range...");
  let rangeArr = findSpecMonthRange(hideSheet, today, monthEndRowListCol);
  let addRowSpec = rangeArr[0],
  addRowSpecLen = rangeArr[2];
  errorMsgOut.setValue("Updating dropdown list...");
  let dropdownArr = specSheet.getRange(addRowSpec, addColSpec, addRowSpecLen, 1).getValues();
  dropdownArr.push("N/A"); //add N/A to dropdown list as by default it is not in the list

  //clear new expense type cell & revalidate expnotetype dropdown list
  expenseNoteType.setValue("N/A");
  newExpenseNoteType.clearContent();
  expenseNoteType.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(dropdownArr).build());

  return;
}

//adds in val to chosen cell given parameters
function addButtonAct(monthEndRowListCol, typeSheet, specSheet, hideSheet){

  errorMsgIn.setValue("Finding rows...");

  let today = new Date();
  let addRow = findAddRow(typeSheet, today),
  addCol,
  addColSpec, //for dropdown list
  fixedOrNotVal = fixedOrNot.getValue();

  errorMsgIn.setValue("Finding columns...");
  //CHECK
  if (checkOrResIn.getValue() == "CHECK") {
    addCol = findAddCol(typeSheet, null, "IN", "CHECK", "type");
    addColSpec = findAddCol(typeSheet, null, "IN", "CHECK", "spec") + 3;
    fixedOrNot.setBackground("#999999");
  }

  //RES
  else {
    fixedOrNot.setBackground("#cccccc");
    addCol = findAddCol(typeSheet, null, fixedOrNotVal, "RES", "type");
    addColSpec = findAddCol(typeSheet, null, fixedOrNotVal, "RES", "spec") + 3;
  }

  errorMsgIn.setValue("Adding amount...");
  addMoney(addRow, addCol, amountIn.getValue(), typeSheet); // adds amount to curr eqn

  //vars for dropdown
  errorMsgIn.setValue("Finding spec sheet month range...");
  let rangeArr = findSpecMonthRange(hideSheet, today, monthEndRowListCol);
  let addRowSpec = rangeArr[0],
  addRowSpecLen = rangeArr[2];
  errorMsgIn.setValue("Updating dropdown list...");
  let dropdownArr = specSheet.getRange(addRowSpec, addColSpec, addRowSpecLen, 1).getValues();
  dropdownArr.push("N/A"); //add N/A to dropdown list as by default it is not in the list

  //clear new income type cell & revalidate incomenotetype dropdown list
  incomeNoteType.setValue("N/A");
  newIncomeNoteType.clearContent();
  incomeNoteType.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(dropdownArr).build());

  return;
}


//checks reimb of specified year & date to see what isn't reimbed yet
function checkReimb(monthEndRowListCol, specSheet, hideSheet) {

  //month range
  errorMsgReimb.setValue("Finding month range...");
  let chosenDate = new Date(reimbYear.getValue(), reimbMonth.getValue() - 1);
  let rangeArr = findSpecMonthRange(hideSheet, chosenDate, monthEndRowListCol);
  let monthRowInd = rangeArr[0],
  monthEndRow = rangeArr[1];

  //find cols with expense type names & reimb mark
  errorMsgReimb.setValue("Finding columns...");
  let totCostColSpec = findAddCol(specSheet, null, "REIMB OUT", checkOrResReimb.getValue(), "spec") + 2; //expense type param ignored
  let expTypeColSpec = totCostColSpec + 1, //expense type param ignored
  reimbMarkColSpec = totCostColSpec + 3,

  //create array of non-reimbed items w/ N/A as default
  nonReimbArray = ["N/A"];

  //adds into array where only non-reimbed items exist w/ their respective costs
  errorMsgReimb.setValue("Finding non-reimb items...");
  while (monthRowInd <= monthEndRow) {
    if (!specSheet.getRange(monthRowInd, reimbMarkColSpec).getValue() && !specSheet.getRange(monthRowInd, totCostColSpec).isBlank()) {
      nonReimbArray.push(specSheet.getRange(monthRowInd, totCostColSpec).getValue() + ": " + specSheet.getRange(monthRowInd, expTypeColSpec).getValue());
    }
    monthRowInd++;
  }

  //revalidate nonReimbCell & nonReimbCostCell dropdown list
  errorMsgReimb.setValue("Updating dropdown list...");
  nonReimbCell.setValue("N/A");
  nonReimbCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(nonReimbArray, true).build());

  //check if there's anything to reimb
  if (nonReimbArray.length > 1) return true;
  return false;
}


//----------spec sheet mods----------//

//modifies the spec sheet for deductions in the spec sheet
function subModSpecSheet(date, monthEndRowListCol, specSheet, hideSheet) {

  //find add col in spec w/ some init vars
  let ccolWithDate,
  checkOrResOutVal = checkOrResOut.getValue(),
  expenseTypeVal = expenseType.getValue(),
  amountOutVal = amountOut.getValue(),
  needOrWantOrReimbVal = needOrWantOrReimb.getValue(),
  expenseNoteTypeVal = expenseNoteType.getValue(),
  newExpenseNoteTypeVal = newExpenseNoteType.getValue(),
  outCreditTypeVal = outCreditType.getValue();

  //for readability
  if (needOrWantOrReimbVal == "REIMB") needOrWantOrReimbVal = "REIMB OUT";

  //RES col finding
  if (checkOrResOutVal == "RES") {
    if (needOrWantOrReimbVal == "REIMB OUT") {
      ccolWithDate = findAddCol(specSheet, expenseTypeVal, needOrWantOrReimbVal, "RES", "spec"); //by default settles on date col
    }
    else { //not reimb out
      ccolWithDate = findAddCol(specSheet, expenseTypeVal, "OUT", "RES", "spec"); //by default settles on date col
    }
  }

  //CHECK col finding
  else {
    //find col of targeted cell given N/W/R & exp type
    ccolWithDate = findAddCol(specSheet, expenseTypeVal, needOrWantOrReimbVal, "CHECK", "spec"); //by default settles on date col
  }

  //other cols relative to found one (reimb is exception)
  let ccolWithBrokeDownCost = ccolWithDate + 1,
  ccolWithExpTotCost = ccolWithDate + 2,
  ccolWithExpTypeNames = ccolWithDate + 3,
  ccolWithCardType = ccolWithDate + 4,
  ccolWithReimbMark = ccolWithDate + 5; //may or may not be used

  //finding range of month in spec sheet to find target row
  let rangeArr = findSpecMonthRange(hideSheet, date, monthEndRowListCol);
  let startRow = rangeArr[0],
  lastRow = rangeArr[1],
  totalMonthLen = rangeArr[2],
  targetRow; //the row to add entry

  //checks if there is space in specific category to add entry; if not extend & set target row to last row
  if (!specSheet.getRange(lastRow, ccolWithBrokeDownCost).isBlank()) {
    addEntryRow(date, 5, 104, specSheet, hideSheet);
    lastRow++; //will only extend in 1 increments
    totalMonthLen++;
    targetRow = lastRow;
  }
  else targetRow = findFirstBlankRow(specSheet, startRow, lastRow, ccolWithBrokeDownCost); //first blank row set as target row

  //cell vars for readability
  let dateCell = specSheet.getRange(targetRow, ccolWithDate),
  brokeDownCostCell = specSheet.getRange(targetRow, ccolWithBrokeDownCost),
  totCostCell = specSheet.getRange(targetRow, ccolWithExpTotCost),
  expTypeCell = specSheet.getRange(targetRow, ccolWithExpTypeNames),
  creditCell = specSheet.getRange(targetRow, ccolWithCardType),
  reimbCell = specSheet.getRange(targetRow, ccolWithReimbMark);

  //note entry dne
  if (expenseNoteTypeVal == "N/A") {
    dateCell.setValue(date); //set date
    brokeDownCostCell.setValue(amountOutVal); //set cost
    totCostCell.setValue(amountOutVal); //set total cost the same as cost
    creditCell.setValue(outCreditTypeVal); //set credit type
    if (needOrWantOrReimbVal == "REIMB OUT") reimbCell.setValue(false); //if in reimb set default reimb to false (will set true by reimb button)
    expTypeCell.setValue(newExpenseNoteTypeVal); //set exp type as new
  }
  else { //possible existing expense type; reimb is assumed to be false (if it is in reimb to begin with)
    let newTargetRow = startRow;
    while (expenseNoteTypeVal != specSheet.getRange(newTargetRow, ccolWithExpTypeNames).getValue() && newTargetRow <= lastRow) newTargetRow++; //iterate to find the right row with the same exp type

    //re-find the range with the new target row for comparisons
    creditCell = specSheet.getRange(newTargetRow, ccolWithCardType),
    reimbCell = specSheet.getRange(newTargetRow, ccolWithReimbMark);

    //extra conditions if the reimb or the credit type differs from what's entered; add to original target row as new entry
    if ((needOrWantOrReimbVal == "REIMB OUT" && reimbCell.getValue()) || (creditCell.getValue() != outCreditTypeVal)) {
      //reset to prev target row
      creditCell = specSheet.getRange(targetRow, ccolWithCardType),
      reimbCell = specSheet.getRange(targetRow, ccolWithReimbMark);

      dateCell.setValue(date); //set date
      brokeDownCostCell.setValue(amountOutVal); //set cost
      totCostCell.setValue(amountOutVal); //set total cost the same as cost
      creditCell.setValue(outCreditTypeVal); //set credit type
      if (needOrWantOrReimbVal == "REIMB OUT") reimbCell.setValue(false); //if in reimb set default reimb to false (will set true by reimb button)
      expTypeCell.setValue(expenseNoteTypeVal); //set exp type as current as it is not N/A
    }
    else { //exact same entry with a new cost
      //update remaining cells to new target row
      dateCell = specSheet.getRange(newTargetRow, ccolWithDate),
      brokeDownCostCell = specSheet.getRange(newTargetRow, ccolWithBrokeDownCost),
      totCostCell = specSheet.getRange(newTargetRow, ccolWithExpTotCost);

      dateCell.setValue(date); //set date (force updates existing entry to recently modified date)
      brokeDownCostCell.setValue(brokeDownCostCell.getValue() + "+" + amountOutVal); //add onto existing formula
      totCostCell.setValue(totCostCell.getValue() + amountOutVal); //add onto existing total cost
    }
  }


  let dropdownArr = specSheet.getRange(startRow, ccolWithExpTypeNames, totalMonthLen, 1).getValues();
  dropdownArr.push("N/A"); //add N/A to dropdown list as by default it is not in the list
  expenseNoteType.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(dropdownArr).build()); //revalidate expnotetype dropdown list

  return;
}


//modifies the spec sheet for additions in the spec sheet
function addModSpecSheet(date, monthEndRowListCol, specSheet, hideSheet) {

  //find add col in spec w/ some init vars
  let ccolWithDate,
  checkOrResInVal = checkOrResIn.getValue(),
  fixedOrNotVal = fixedOrNot.getValue(),
  amountInVal = amountIn.getValue(),
  incomeNoteTypeVal = incomeNoteType.getValue(),
  newIncomeNoteTypeVal = newIncomeNoteType.getValue(),
  inCreditTypeVal = inCreditType.getValue();

  //CHECK
  if (checkOrResInVal == "CHECK") {
    ccolWithDate = findAddCol(specSheet, null, "IN", "CHECK", "spec");
  }

  //RES
  else {
    ccolWithDate = findAddCol(specSheet, null, fixedOrNotVal, "RES", "spec");
  }

  //other cols relative to found one (reimb is exception)
  let ccolWithBrokeDownCost = ccolWithDate + 1,
  ccolWithExpTotCost = ccolWithDate + 2,
  ccolWithInTypeNames = ccolWithDate + 3,
  ccolWithCardType = ccolWithDate + 4;

  //finding range of month in spec sheet to find target row
  let rangeArr = findSpecMonthRange(hideSheet, date, monthEndRowListCol);
  let startRow = rangeArr[0],
  lastRow = rangeArr[1],
  totalMonthLen = rangeArr[2],
  targetRow; //the row to add entry

  //checks if there is space in specific category to add entry; if not extend & set target row to last row
  if (!specSheet.getRange(lastRow, ccolWithBrokeDownCost).isBlank()) {
    addEntryRow(date, 5, 104, specSheet, hideSheet);
    lastRow++; //will only extend in 1 increments
    totalMonthLen++;
    targetRow = lastRow;
  }
  else targetRow = findFirstBlankRow(specSheet, startRow, lastRow, ccolWithBrokeDownCost); //first blank row set as target row

  //cell vars for readability
  let dateCell = specSheet.getRange(targetRow, ccolWithDate),
  brokeDownCostCell = specSheet.getRange(targetRow, ccolWithBrokeDownCost),
  totCostCell = specSheet.getRange(targetRow, ccolWithExpTotCost),
  inTypeCell = specSheet.getRange(targetRow, ccolWithInTypeNames),
  creditCell = specSheet.getRange(targetRow, ccolWithCardType);

  //note entry dne
  if (incomeNoteTypeVal == "N/A") {
    dateCell.setValue(date); //set date
    brokeDownCostCell.setValue(amountInVal); //set cost
    totCostCell.setValue(amountInVal); //set total cost the same as cost
    inTypeCell.setValue(newIncomeNoteTypeVal); //set exp type as new
    creditCell.setValue(inCreditTypeVal); //set credit type
  }
  else { //possible existing expense type
    let newTargetRow = startRow;
    while (incomeNoteTypeVal != specSheet.getRange(newTargetRow, ccolWithInTypeNames).getValue() && newTargetRow <= lastRow) newTargetRow++; //iterate to find the right row with the same exp type

    creditCell = specSheet.getRange(newTargetRow, ccolWithCardType); //re-find the range with the new target row for comparisons

    //extra conditions if the credit type differs from what's entered; add to original target row as new entry
    if (creditCell.getValue() != inCreditTypeVal) {

      creditCell = specSheet.getRange(targetRow, ccolWithCardType); //reset to prev target row

      dateCell.setValue(date); //set date
      brokeDownCostCell.setValue(amountInVal); //set cost
      totCostCell.setValue(amountInVal); //set total cost the same as cost
      inTypeCell.setValue(incomeNoteTypeVal); //set exp type as current as it is not N/A
      creditCell.setValue(inCreditTypeVal); //set credit type
    }

    //update remaining cells to new target row
    dateCell = specSheet.getRange(newTargetRow, ccolWithDate),
    brokeDownCostCell = specSheet.getRange(newTargetRow, ccolWithBrokeDownCost),
    totCostCell = specSheet.getRange(newTargetRow, ccolWithExpTotCost);

    dateCell.setValue(date); //set date (force updates existing entry to recently modified date)
    brokeDownCostCell.setValue(brokeDownCostCell.getValue() + "+" + amountInVal); //add onto existing formula
    totCostCell.setValue(totCostCell.getValue() + amountInVal); //add onto existing total cost
  }


  let dropdownArr = specSheet.getRange(startRow, ccolWithInTypeNames, totalMonthLen, 1).getValues();
  dropdownArr.push("N/A"); //add N/A to dropdown list as by default it is not in the list
  incomeNoteType.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(dropdownArr).build()); //revalidate innotetype dropdown list

  return;
}


//----------miscellaneous----------//


//Add amount to appropriate cell
function addMoney(addRow, addCol, amount, typeSheet){
  let curEq = typeSheet.getRange(addRow, addCol).getFormula();
  if (curEq == "=0") {
    typeSheet.getRange(addRow, addCol).setFormula(amount);
  }
  else {
    typeSheet.getRange(addRow, addCol).setFormula(curEq + "+" + amount);
  }
}


//find appropriate col given need/want and expense type & sheet type (specific or normal)
function findAddCol(sheet, expenseType, colCases, checkOrRes, typeOrSpec) {
  let addCol;
  switch (checkOrRes) {
    case "CHECK":
      switch (colCases) {
        case "NEED":
          if (typeOrSpec == "type") addCol = needWantLoop(needStart, needEnd, sheet, expenseType, typeOrSpec);
          else if (typeOrSpec == "spec") addCol = needWantLoop(needStartSpec, needEndSpec, sheet, expenseType, typeOrSpec);
          break;
        case "WANT":
          if (typeOrSpec == "type") addCol = needWantLoop(wantStart, wantEnd, sheet, expenseType, typeOrSpec);
          else if (typeOrSpec == "spec") addCol = needWantLoop(wantStartSpec, wantEndSpec, sheet, expenseType, typeOrSpec);
          break;
        case "REIMB IN":
          if (typeOrSpec == "type") addCol = checkReimbInCol;
          else if (typeOrSpec == "spec") addCol = checkReimbInColSpec;
          break;
        case "REIMB OUT":
          if (typeOrSpec == "type") addCol = checkReimbOutCol;
          else if (typeOrSpec == "spec") addCol = checkReimbOutColSpec;
          break;
        case "IN":
          if (typeOrSpec == "type") addCol = checkInCol;
          else if (typeOrSpec == "spec") addCol = checkInColSpec;
          break;
        default:
          addCol = -1;
      }
      break;
    case "RES":
      switch (colCases) {
        case "FIXED":
          if (typeOrSpec == "type") addCol = resFixedCol;
          else if (typeOrSpec == "spec") addCol = resFixedColSpec;
          break;
        case "NON-FIXED":
          if (typeOrSpec == "type") addCol = resNonFixedCol;
          else if (typeOrSpec == "spec") addCol = resNonFixedColSpec;
          break;
        case "REIMB IN":
          if (typeOrSpec == "type") addCol = resReimbInCol;
          else if (typeOrSpec == "spec") addCol = resReimbInColSpec;
          break;
        case "REIMB OUT":
          if (typeOrSpec == "type") addCol = resReimbOutCol;
          else if (typeOrSpec == "spec") addCol = resReimbOutColSpec;
          break;
        case "OUT":
          if (typeOrSpec == "type") addCol = resOutCol;
          else if (typeOrSpec == "spec") addCol = resOutColSpec;
          break;
        default:
          addCol = -1;
      }
      break;
    default:
      addCol = -1;
    }
  return addCol;
}


//find appropriate row given current month and year
function findAddRow(sheet, today) {
  let addRow,
  currYear = today.getFullYear();
  if (sheet.getName() == "College Savings 3.0") {
    if (currYear == 2022) { //2022 is special case (only 4 months)
      addRow = monthRowFinder(4, 7, today);
    }
    else { //finds normally
      let baseYear = 2023;
      let startRow = 8 + ((currYear - baseYear) * 12);
      addRow = monthRowFinder(startRow, startRow + 11, today);
    }
  }
  else if (sheet.getName() == "College Savings 3.0 (TW)") {
    let baseYear = 2023;
    let startRow = 4 + ((currYear - baseYear) * 12);
    addRow = monthRowFinder(startRow, startRow + 11, today);
  }
  return addRow;
}


//loop to find col num of expense type spec or type
function needWantLoop(start, end, sheet, expenseType, typeOrSpec) {
  if (typeOrSpec == "type") {
    for (let i = start; i <= end; i++) {
      if (sheet.getRange(3, i).getValue() == expenseType) {
        return i;
      }
    }
  } else if (typeOrSpec == "spec") {
    for (let i = start; i <= end; i+=5) {
      if (sheet.getRange(3, i).getValue() == expenseType) {
        return i;
      }
    }
  }
  return -1;
}


//adds rows to specific months and years for additional entries for spec sheets
function addEntryRow(today, monthEndRowsListCol, lastColWithData, sheet, hideSheet){
  //find row of month in hide menu
  let row = findAddRowForSpecHide(hideSheet, today);

  //find range of rows a month holds
  let lastRow = hideSheet.getRange(row, monthEndRowsListCol).getValue(),
  prevLastRow = hideSheet.getRange(row - 1, monthEndRowsListCol).getValue() + 1;

  //add row after last row of chosen month
  sheet.insertRowAfter(lastRow);

  //get A1 notation of first and last cell of month merged cell & newly created row's cell
  let prevCell = sheet.getRange(prevLastRow, 2).getA1Notation(),
  curCell = sheet.getRange(lastRow + 1, 2).getA1Notation();

  //for dec where the year cell needs to be extended
  if (today.getMonth() == 11){
    //set first month as sep if 2022, jan otherwise
    let tempDay;
    if (today.getFullYear() == 2022) tempDay = new Date(today.getFullYear(), 8);
    else tempDay = new Date(today.getFullYear(), 0);

    //find first row & new extended last row of year
    let tempRow = findAddRowForSpecHide(hideSheet, tempDay) - 1;
    let yearStartRow = hideSheet.getRange(tempRow, monthEndRowsListCol).getValue() + 1;

    //get A1 notation of respective cells and merge
    let yearStartCell = sheet.getRange(yearStartRow, 1).getA1Notation(),
    curYearCell = sheet.getRange(lastRow + 1, 1).getA1Notation();
    sheet.getRange(yearStartCell+":"+curYearCell).merge();
  }

  //merge curr month cell & new cell & increment last row from curr month to end
  sheet.getRange(prevCell+":"+curCell).merge();

  //increment all row values by 1 below the extended month
  let rangeToUpdate = hideSheet.getRange(row, monthEndRowsListCol, 58 - row, 1);
  let rowValues = rangeToUpdate.getValues();
  let updatedValues = rowValues.map(function(rows) {
    return [rows[0] + 1];
  });
  rangeToUpdate.setValues(updatedValues);

  //get cell range of all data in month
  let prevDataUpLeftCell = sheet.getRange(prevLastRow, 3).getA1Notation(),
  curDataDownRightCell = sheet.getRange(lastRow + 1, lastColWithData).getA1Notation();

  //redo borders in given cell range
  sheet.getRange(prevDataUpLeftCell+":"+curDataDownRightCell).setBorder(true, true, true, true, true, false, "black", null);
  return;
}


//find appropriate row given current month and year
function findAddRowForSpecHide(sheet, today) {
  let addRow,
  currYear = today.getFullYear();
  if (sheet.getName() == "College Savings 3.0 Specifics Hide Menu") {
    if (currYear == 2022) { //2022 is special case (only 4 months)
      addRow = monthRowFinder(6, 9, today);
    }
    else { //finds normally
      let baseYear = 2023;
      let startRow = 10 + ((currYear - baseYear) * 12);
      addRow = monthRowFinder(startRow, startRow + 11, today);
    }
  }
  return addRow;
}


//finds range of values and the length of the range; returns array with start, end, length, and row of month given date
function findSpecMonthRange(hideSheet, date, monthEndRowsListCol) {
  //find row of month in hide menu
  let monthRow = findAddRowForSpecHide(hideSheet, date);
  //find range of rows a month holds
  let lastRow = hideSheet.getRange(monthRow, monthEndRowsListCol).getValue();
  let startRow = hideSheet.getRange(monthRow - 1, monthEndRowsListCol).getValue() + 1;
  let totalMonthLen = lastRow - startRow + 1;
  return [startRow, lastRow, totalMonthLen, monthRow];
}


//hides certain rows or col entries based on pressed buttons
function entryHiding(activeCell, activeVal, hideSheet, targetSpecSheet, endRowOrColListCol, rowOrCol){
    //vars for readability
    let activeCellRange, //in the off chance the button is a merged button
    individualButtonCol = null, //in the off chance the button is a merged button
    activeCellRangeLastRow, //in the off chance the button is a merged button
    activeCellRangeRow,  //in the off chance the button is a merged button
    buttonRow,
    lastRowOrCol,
    prevLastRowOrCol,
    rowOrColRange;

    if (activeCell.isPartOfMerge()) { //if the button is part of a merged range, get the range & set rows accordingly
      activeCellRange = activeCell.getMergedRanges()[0]; //get the range of the clicked merged cell from the returned array
      activeCellRangeRow = activeCellRange.getRow(); //get the row of the merged
      activeCellRangeLastRow = activeCellRange.getLastRow(); //get the last row of the merged range

      lastRowOrCol = hideSheet.getRange(activeCellRangeLastRow, endRowOrColListCol).getValue(); //get the last row or col of the months or categories
      prevLastRowOrCol = hideSheet.getRange(activeCellRangeRow - 1, endRowOrColListCol).getValue() + 1; //get the row or col of the prev months or categories and add by 1
      individualButtonCol = endRowOrColListCol - 2;
    }
    else {
      buttonRow = activeCell.getRow(); //get the row of the button clicked
      lastRowOrCol = hideSheet.getRange(buttonRow, endRowOrColListCol).getValue(); //get the last row or col of the month or category
      prevLastRowOrCol = hideSheet.getRange(buttonRow - 1, endRowOrColListCol).getValue() + 1; //get the row or col of the prev month or category and add by 1
    }

    rowOrColRange = lastRowOrCol - prevLastRowOrCol + 1; //get the range/number of rows or cols to hide

    if (activeVal == true) {
      if (rowOrCol == "row") {
        targetSpecSheet.hideRows(prevLastRowOrCol, rowOrColRange);
      }
      else if (rowOrCol == "col")
        targetSpecSheet.hideColumns(prevLastRowOrCol, rowOrColRange);
    }
    else if (activeVal == false) {
      if (rowOrCol == "row")
        targetSpecSheet.showRows(prevLastRowOrCol, rowOrColRange);
      else if (rowOrCol == "col") {
        targetSpecSheet.showColumns(prevLastRowOrCol, rowOrColRange);
      }
    }

    if (individualButtonCol != null) { //if the button is part of a merged range, set the value of the individual button as the merged range value
      hideSheet.getRange(activeCellRangeRow, individualButtonCol, activeCellRangeLastRow - activeCellRangeRow + 1).setValue(activeVal);
    }
}

//loop to find row num of current month
function monthRowFinder(start, end, today) {
  let finalAddRow = start + today.getMonth() - (12 - (end - start + 1));
  if (finalAddRow < start || finalAddRow > end) return -1;
  return finalAddRow;
}

//finds first blank row in a given range of rows
function findFirstBlankRow(sheet, startRow, endRow, col) {
  while (startRow <= endRow) {
    let mid = Math.floor((startRow + endRow) / 2);
    if (sheet.getRange(mid, col).isBlank()) {
      if (sheet.getRange(mid - 1, col).isBlank()) { //if the row before is blank, set new end row
        endRow = mid - 1;
      } else {
        return mid; //first blank row found
      }
    } else {
      startRow = mid + 1; //if not blank, set new start row
    }
  }
  return -1; //no blank row found
}

//----------note modding; will be left here for the benefit of the reader----------//


//modifies specified out note w/ updated info
function outNoteMod(checkOrRes, needOrWantOrReimb, expenseType, amount, expenseNoteType, newExpenseNoteType, typeSheet) {

  var today = new Date();
  var sheetInd = rowThatDropdownSheetStarts + 1;

  var addCol;
  var addRow = findAddRow(typeSheet, today);

  if (checkOrRes.getValue() == "CHECK"){
    addCol = findAddCol(typeSheet, expenseType.getValue(), needOrWantOrReimb.getValue());
  }
  //RES
  else {
    addCol = resOutCol;
  }

  //finds row of chosen expense type in expTypeList & modifies targetted note accordingly
  if (expenseNoteType.getValue() != "N/A") {
    //Reset the notes of specified cell
    typeSheet.getRange(addRow, addCol).setNote("");
    while (sheetInd <= typeSheet.getLastRow() && typeSheet.getRange(sheetInd, 33).getValue() != "") {
      var finEq = typeSheet.getRange(sheetInd, colWithBrokeDownCost);
      var finCost = typeSheet.getRange(sheetInd, colWithExpTotCost);
      var expType = typeSheet.getRange(sheetInd, colWithExpTypeNames);
      var reimbOrNot = typeSheet.getRange(sheetInd, colWithReimbMark);

      if (expType.getValue() == expenseNoteType.getValue()) {
        //Finding final equation, final cost and modifying it into existing targetted note
        finEq.setValue(finEq.getValue() + "+" + amount.getValue());
        finCost.setValue(finCost.getValue() + amount.getValue());

        //additional tilde if in reimb and is false
        if (needOrWantOrReimb.getValue() == "REIMB" && reimbOrNot.getValue() == false) {
          typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + "~ " + finEq.getValue() + " (" + finCost.getValue() + "): " + expenseNoteType.getValue() + "\n");
        }
        else {
          typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + finEq.getValue() + " (" + finCost.getValue() + "): " + expenseNoteType.getValue() + "\n");
        }
      }
      else {
        if (finEq.getValue() == finCost.getValue()) {
          //prevent notes where the total cost and the sum is the same (ex: "50 (50): TEST")  
          if (needOrWantOrReimb.getValue() == "REIMB" && reimbOrNot.getValue() == false) {
            //in reimb and isnt paid bk
            typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + "~ " + finEq.getValue() + ": " + expType.getValue() + "\n");
          } 
          else {
            typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + finEq.getValue() + ": " + expType.getValue() + "\n");
          }
        } 
        else {
          //equation and total cost aren't the same string
          if (needOrWantOrReimb.getValue() == "REIMB" && reimbOrNot.getValue() == false) {
            //in reimb and isnt paid bk
            typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + "~ " + finEq.getValue() + " (" + finCost.getValue() + "): " + expType.getValue() + "\n");
          } 
          else {
            typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + finEq.getValue() + " (" + finCost.getValue() + "): " + expType.getValue() + "\n");
          }
        }
      }
      sheetInd++;
    }
  }
  else {
    //add new expense type entry into targetted notes
    if (needOrWantOrReimb.getValue() == "REIMB") {
      typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote().toString() + "~ " + amount.getValue() + ": " + newExpenseNoteType.getValue() + "\n");
    } else {
      typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote().toString() + amount.getValue() + ": " + newExpenseNoteType.getValue() + "\n");
    }
  }
  return;
}


//modifies specified in note w/ updated info
function inNoteMod(checkOrRes, fixedOrNot, amount, incomeNoteType, newIncomeNoteType, typeSheet){

  var today = new Date();
  var sheetInd = rowThatDropdownSheetStarts + 1;

  var addCol;
  var addRow = findAddRow(typeSheet, today);

  if (checkOrRes.getValue() == "RES"){
    if (fixedOrNot.getValue() == "FIXED") {
      addCol = resFixedCol;
    }
    //NON-FIXED
    else {
      addCol = resNonFixedCol;
    }
  }
  //CHECK
  else {
    addCol = checkInCol;
  }

  //finds row of chosen expense type in expTypeList & modifies targetted note accordingly
  if (incomeNoteType.getValue() != "N/A") {
    //Reset the notes of specified cell
    typeSheet.getRange(addRow, addCol).setNote("");
    while (sheetInd <= typeSheet.getLastRow() && typeSheet.getRange(sheetInd, colWithExpTypeNames).getValue() != "") {
      var finEq = typeSheet.getRange(sheetInd, colWithBrokeDownCost);
      var finInc = typeSheet.getRange(sheetInd, colWithExpTotCost);
      var incType = typeSheet.getRange(sheetInd, colWithExpTypeNames);

      if (incType.getValue() == incomeNoteType.getValue()) {
        //Finding final equation, final cost and modifying it into existing targetted note
        finEq.setValue(finEq.getValue() + "+" + amount.getValue());
        finInc.setValue(finInc.getValue() + amount.getValue());

        typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + finEq.getValue() + " (" + finInc.getValue() + "): " + incomeNoteType.getValue() + "\n");
      }
      else {
        if (finEq.getValue() == finInc.getValue()) {
          //prevent notes where the total cost and the sum is the same (ex: "50 (50): TEST")  
            typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + finEq.getValue() + ": " + incType.getValue() + "\n");
        } 
        else {
          //equation and total cost aren't the same string
            typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote() + finEq.getValue() + " (" + finInc.getValue() + "): " + incType.getValue() + "\n");
        }
      }
      sheetInd++;
    }
  }
  else {
    //add new expense type entry into targetted notes
      typeSheet.getRange(addRow, addCol).setNote(typeSheet.getRange(addRow, addCol).getNote().toString() + amount.getValue() + ": " + newIncomeNoteType.getValue() + "\n");
  }
  return;
}


//modifies specified reimb note w/ updated info
function alrReimbedNoteMod(year, month, nonReimbCell, typeSheet) {
  var monthRow = findAddRow(typeSheet, new Date(year.getValue(), month.getValue() - 1));
  var nonReimbEntry = nonReimbCell.getValue().split(": ");
  var sheetInd = rowThatDropdownSheetStarts + 1;
  var targetSheetInd;
  var foundReimbContent = false;

  //finds cell to reimb
  while (sheetInd <= typeSheet.getLastRow() && typeSheet.getRange(sheetInd, colWithExpTypeNames).getValue() != "" && foundReimbContent == false) {
    console.log(typeSheet.getRange(sheetInd, colWithExpTypeNames).getValue()+" "+nonReimbEntry[1]);
    if (typeSheet.getRange(sheetInd, colWithExpTypeNames).getValue() == nonReimbEntry[1]) {
      foundReimbContent = true;
      typeSheet.getRange(sheetInd, colWithReimbMark).setValue(true);
      targetSheetInd = sheetInd;
    }
    sheetInd++;
  }

  //Reset the notes of specified cell
  typeSheet.getRange(monthRow, checkReimbOutCol).setNote("");

  //adds notes bk into reimb out column
  sheetInd = rowThatDropdownSheetStarts + 1;
  while (sheetInd <= typeSheet.getLastRow() && typeSheet.getRange(sheetInd, colWithExpTypeNames).getValue() != "") {

    var finEq = typeSheet.getRange(sheetInd, colWithBrokeDownCost);
    var finCost = typeSheet.getRange(sheetInd, colWithExpTotCost);
    var expType = typeSheet.getRange(sheetInd, colWithExpTypeNames);
    var reimbOrNot = typeSheet.getRange(sheetInd, colWithReimbMark);

    if (finEq.getValue() == finCost.getValue()) {
      //prevent notes where the total cost and the sum is the same (ex: "50 (50): TEST")
      if (reimbOrNot.getValue() == false) {
        typeSheet.getRange(monthRow, checkReimbOutCol).setNote(typeSheet.getRange(monthRow, checkReimbOutCol).getNote() + "~ " + finEq.getValue() + ": " + expType.getValue() + "\n");
      } else {
        typeSheet.getRange(monthRow, checkReimbOutCol).setNote(typeSheet.getRange(monthRow, checkReimbOutCol).getNote() + finEq.getValue() + ": " + expType.getValue() + "\n");
      }
    } else {
      //equation and total cost aren't the same string
      if (reimbOrNot.getValue() == false) {
        typeSheet.getRange(monthRow, checkReimbOutCol).setNote(typeSheet.getRange(monthRow, checkReimbOutCol).getNote() + "~ " + finEq.getValue() + " (" + finCost.getValue() + "): " + expType.getValue() + "\n");
      } else {
        typeSheet.getRange(monthRow, checkReimbOutCol).setNote(typeSheet.getRange(monthRow, checkReimbInCol).getNote() + finEq.getValue() + " (" + finCost.getValue() + "): " + expType.getValue() + "\n");
      }
    }
    sheetInd++;
  }

  //add new reimbed entry into reimb in column
  var finEq = typeSheet.getRange(targetSheetInd, colWithBrokeDownCost);
  var finCost = typeSheet.getRange(targetSheetInd, colWithExpTotCost);
  var expType = typeSheet.getRange(targetSheetInd, colWithExpTypeNames);
  if (finEq.getValue() == finCost.getValue()) {
    //prevent notes where the total cost and the sum is the same (ex: "50 (50): TEST")  
    typeSheet.getRange(monthRow, checkReimbInCol).setNote(typeSheet.getRange(monthRow, checkReimbInCol).getNote() + finEq.getValue() + ": " + expType.getValue() + "\n");
  } else {
    //equation and total cost aren't the same string
    typeSheet.getRange(monthRow, checkReimbInCol).setNote(typeSheet.getRange(monthRow, checkReimbInCol).getNote() + finEq.getValue() + " (" + finCost.getValue() + "): " + expType.getValue() + "\n");
  }
  addMoney(monthRow, checkReimbInCol, Number(nonReimbEntry[0]), typeSheet);
  checkReimb(year, month, nonReimbCell, )
}


//Print all notes cost and exp type (given it isn't empty)
function noteToSheets(typeSheet, addRow, addCol, needOrWantOrReimb) {
  //clears expCost & expList for a new uninterrupted list (hidden)
  typeSheet.getRange(rowThatDropdownSheetStarts + 1, colWithBrokeDownCost, typeSheet.getLastRow() - rowThatDropdownSheetStarts, 4).clearContent();

  //Print all notes cost and exp type (given it isn't empty)
  var notes = typeSheet.getRange(addRow, addCol).getNotes().toString().split("\n");
  var noteInd = 0;
  var sheetInd = rowThatDropdownSheetStarts + 1;
  while (noteInd < notes.length && sheetInd <= typeSheet.getLastRow()) {
    if (notes[noteInd].length > 0) {
      //line isn't empty (handling user typo error)
      var tempEntry = notes[noteInd].split(": ");

      //split the total cost & equation w/o tilde (if exists)
      var tempCostEntry = tempEntry[0].replace(")", "").replace("~", "").split(" (");

      //if note in reimb, split the tilde (if it exists)
      var reimbedOrNot = tempEntry[0].split(" ");
      if (needOrWantOrReimb.getValue() == "REIMB" && reimbedOrNot[0] == "~") {
        //tilde; hence not reimbed
        typeSheet.getRange(sheetInd, colWithReimbMark).setValue(false);
      } 
      else {
        //not in reimb column or alr reimbed
        typeSheet.getRange(sheetInd, colWithReimbMark).setValue(true);
      }

      //put values in respective columns
      typeSheet.getRange(sheetInd, colWithBrokeDownCost).setValue(tempCostEntry[0]);

      //no formula exists (1 cost)
      if (tempCostEntry[1] == null) {
        typeSheet.getRange(sheetInd, colWithExpTotCost).setValue(tempCostEntry[0])
      } 
      else {
        //a formula exists
        typeSheet.getRange(sheetInd, colWithExpTotCost).setValue(tempCostEntry[1])
      }
      typeSheet.getRange(sheetInd, colWithExpTypeNames).setValue(tempEntry[1]);
      noteInd++;
      sheetInd++;
    } 
    else {
      //skip empty lines
      noteInd++;
    }
  }
  return;
}