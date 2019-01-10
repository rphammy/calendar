function Test() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S6').activate();
  spreadsheet.getRange('T6').activate();
  spreadsheet.getActiveRangeList().setShowHyperlink(true);
  spreadsheet.getCurrentCell().setFormula('=HYPERLINK("#rangeid=1693557053","Approve")');
};