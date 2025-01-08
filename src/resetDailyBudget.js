function resetDailyBudget() {
  var conSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IgGVDEgjKiO_6tKE7XQ7KtM6BFfMHtrEnh8wXgfTDHA/edit#gid=1961861177").getSheetByName("Console");

  if (conSheet.getRange("C22").getValue() > conSheet.getRange("C23").getValue()){
    conSheet.getRange("C22").setValue(conSheet.getRange("C23").getValue());
  }

  if (conSheet.getRange("C24").getValue() > conSheet.getRange("C25").getValue()){
    conSheet.getRange("C24").setValue(conSheet.getRange("C25").getValue());
  }

  conSheet.getRange("C22").setValue("=" + 0);
  conSheet.getRange("C24").setValue("=" + 0);
}
