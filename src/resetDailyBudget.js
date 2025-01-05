function resetDailyBudget() {
  var conSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IgGVDEgjKiO_6tKE7XQ7KtM6BFfMHtrEnh8wXgfTDHA/edit#gid=1961861177").getSheetByName("Console");

  if (conSheet.getRange("C20").getValue() > conSheet.getRange("C21").getValue()){
    conSheet.getRange("C21").setValue(conSheet.getRange("C20").getValue());
  }

  if (conSheet.getRange("C22").getValue() > conSheet.getRange("C23").getValue()){
    conSheet.getRange("C23").setValue(conSheet.getRange("C22").getValue());
  }

  conSheet.getRange("C20").setValue("=" + 0);
  conSheet.getRange("C22").setValue("=" + 0);
}
