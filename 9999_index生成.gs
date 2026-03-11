function buildIndex() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Index");
  const html = sh.getRange("M2").getDisplayValue();
  Logger.log(html);
}


function testIndexRows() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Index');
  Logger.log(sh.getLastRow());
}