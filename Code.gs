function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Show Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Salary Comparator');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function eBuildDefault() {
  buildTable(40, '625.20', '480.33', 'gWeek');
}

function buildTable(hours, gNum, tNum, gross) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  sheet.clear();
  sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
  sheet.deleteRows(2, sheet.getMaxRows() - 1);
  
  while(sheet.getMaxColumns() < 2) {
    sheet.insertColumns(sheet.getMaxColumns());
  }
  
  while(sheet.getMaxRows() < 15) {
    sheet.insertRows(sheet.getMaxRows());
  }
  
  for(var a = 1; a <= sheet.getMaxColumns(); a++) {
    sheet.setColumnWidth(a, 200);
  }
  
  for(var a = 1; a <= sheet.getMaxRows(); a++) {
    sheet.setRowHeight(a, 21);
  }
  
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground('#D9D9D9');
  
  sheet.getRange(1, 1, 1, 2).merge().setValue("http://www.calculators.org/savings/wage-conversion.php#advanced").setBorder(true, true, true, true, true, true);
  
  sheet.getRange(2, 1, 1, 2).merge().setValue("Hours");
  sheet.getRange(3, 1, 1, 2).merge().setValue(hours).setBackground('#b6d7a8');
  sheet.getRange(2, 1, 2, 2).setBorder(true, true, true, true, false, false);
  
  sheet.getRange(4, 1).setValue("Gross Hourly");
  
  sheet.getRange(4, 2).setValue("Taxed Hourly");
  
  sheet.getRange(6, 1).setValue("Gross Daily");
  
  sheet.getRange(6, 2).setValue("Taxed Daily");
  
  sheet.getRange(8, 1).setValue("Gross Weekly");
  
  sheet.getRange(8, 2).setValue("Taxed Weekly");
  
  sheet.getRange(10, 1).setValue("Gross Monthly");
  
  sheet.getRange(10, 2).setValue("Taxed Monthly");
  
  sheet.getRange(12, 1).setValue("Gross Annual");
  
  sheet.getRange(12, 2).setValue("Taxed Annual");
  
  for(var a = 4; a <= 12; a+=2) {
    sheet.getRange(a, 1, 2, 1).setBorder(true, true, true, true, false, false);
    sheet.getRange(a, 2, 2, 1).setBorder(true, true, true, true, false, false);
    sheet.getRange((a + 1), 1, 1, 2).setNumberFormat('$#\,##0.00');
  }
  
  sheet.getRange(14, 1, 1, 2).merge().setValue("Percent Taxed");
  sheet.getRange(15, 1, 1, 2).merge().setValue("=(A9-B9)/A9").setNumberFormat('0.00%');
  sheet.getRange(14, 1, 2, 2).setBorder(true, true, true, true, false, false);
  
  switch (gross) {
    case "gWeek":
      sheet.getRange(9, 1).setValue(gNum).setBackground('#b6d7a8');
      sheet.getRange(9, 2).setValue(tNum).setBackground('#b6d7a8');
      sheet.getRange(5, 1).setValue('=A9/A3');
      sheet.getRange(5, 2).setValue('=B9/A3');
      sheet.getRange(7, 1).setValue('=A9/7');
      sheet.getRange(7, 2).setValue('=B9/7');
      sheet.getRange(11, 1).setValue('=A9*4');
      sheet.getRange(11, 2).setValue('=B9*4');
      sheet.getRange(13, 1).setValue('=A11*12');
      sheet.getRange(13, 2).setValue('=B11*12');
      break;
    case "gMonth":
      sheet.getRange(11, 1).setValue(gNum).setBackground('#b6d7a8');
      sheet.getRange(11, 2).setValue(tNum).setBackground('#b6d7a8');
      sheet.getRange(5, 1).setValue('=A9/A3');
      sheet.getRange(5, 2).setValue('=B9/A3');
      sheet.getRange(7, 1).setValue('=A9/7');
      sheet.getRange(7, 2).setValue('=B9/7');
      sheet.getRange(9, 1).setValue('=A11/4');
      sheet.getRange(9, 2).setValue('=B11/4');
      sheet.getRange(13, 1).setValue('=A11*12');
      sheet.getRange(13, 2).setValue('=B11*12');
      break;
    case "gYear":
      sheet.getRange(13, 1).setValue(gNum).setBackground('#b6d7a8');
      sheet.getRange(13, 2).setValue(tNum).setBackground('#b6d7a8');
      sheet.getRange(5, 1).setValue('=A9/A3');
      sheet.getRange(5, 2).setValue('=B9/A3');
      sheet.getRange(7, 1).setValue('=A9/7');
      sheet.getRange(7, 2).setValue('=B9/7');
      sheet.getRange(9, 1).setValue('=A11/4');
      sheet.getRange(9, 2).setValue('=B11/4');
      sheet.getRange(11, 1).setValue('=A13/12');
      sheet.getRange(11, 2).setValue('=B13/12');
      break;
  }
}
