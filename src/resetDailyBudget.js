function resetDailyBudget() {
  // other users cannot access this sheet; user must input their own URL
  var conSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IgGVDEgjKiO_6tKE7XQ7KtM6BFfMHtrEnh8wXgfTDHA/edit#gid=1961861177").getSheetByName("Console");

  if (conSheet.getRange("C22").getValue() > conSheet.getRange("C23").getValue()){
    conSheet.getRange("C22").setValue(conSheet.getRange("C23").getValue());
  }

  if (conSheet.getRange("C24").getValue() > conSheet.getRange("C25").getValue()){
    conSheet.getRange("C24").setValue(conSheet.getRange("C25").getValue());
  }

  conSheet.getRange("C22").setValue("=" + 0);
  conSheet.getRange("C24").setValue("=" + 0);

  conSheet.getRange("B26").setValue("...");
  conSheet.getRange("E20").setValue("...");
  conSheet.getRange("H16").setValue("...");
  conSheet.getRange("B26").setBackground("#93c47d");
  conSheet.getRange("E20").setBackground("#93c47d");
  conSheet.getRange("H16").setBackground("#93c47d");
}
