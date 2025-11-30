function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range of the sheet and find the last row with data
  let usedRange = selectedSheet.getUsedRange();
  let lastRow = usedRange.getRowCount();

  // Set range O2 on selectedSheet with COUNTIF formula
  selectedSheet.getRange("O2").setFormulaLocal("=COUNTIF(F:F, \"@\")");

  // Auto fill range O2:O(lastRow)
  selectedSheet.getRange("O2").autoFill(selectedSheet.getRange("O2:O" + lastRow), ExcelScript.AutoFillType.fillDefault);

  // Set range P2 on selectedSheet with CONCATENATE formula
  selectedSheet.getRange("P2").setFormulaLocal("=C2&N2");

  // Auto fill range P2:P(lastRow)
  selectedSheet.getRange("P2").autoFill(selectedSheet.getRange("P2:P" + lastRow), ExcelScript.AutoFillType.fillDefault);

  // Paste values for the range O2:O(lastRow)
  selectedSheet.getRange("O2:O" + lastRow).copyFrom(selectedSheet.getRange("O2:O" + lastRow), ExcelScript.RangeCopyType.values, false, false);

  // Paste values for the range P2:P(lastRow)
  selectedSheet.getRange("P2:P" + lastRow).copyFrom(selectedSheet.getRange("P2:P" + lastRow), ExcelScript.RangeCopyType.values, false, false);

  // Set range Q2:AC2 on selectedSheet with VLOOKUP formulas
  selectedSheet.getRange("Q2:AC2").setFormulaLocal([[
    "=VLOOKUP(P2, 'Sheet1'!P:AC,3,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,4,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,5,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,6,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,7,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,8,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,9,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,10,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,11,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,12,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,13,0)", 
    "=VLOOKUP(P2, 'Sheet1'!P:AC,14,0)"
  ]]);

  // Auto fill range Q2:AC2 to Q(lastRow):AC(lastRow)
  selectedSheet.getRange("Q2:AC2").autoFill(selectedSheet.getRange("Q2:AC" + lastRow), ExcelScript.AutoFillType.fillDefault);

  // Paste values for range Q2:AC(lastRow)
  selectedSheet.getRange("Q2:AC" + lastRow).copyFrom(selectedSheet.getRange("Q2:AC" + lastRow), ExcelScript.RangeCopyType.values, false, false);

  // Paste to all cells on selectedSheet from all cells on selectedSheet (copy values)
  selectedSheet.getRange().copyFrom(selectedSheet.getRange(), ExcelScript.RangeCopyType.values, false, false);

  // Set format for range R:R (date format)
  selectedSheet.getRange("R:R").setNumberFormatLocal("dd-mm-yyyy");

  // Set format for range T:T (date format)
  selectedSheet.getRange("T:T").setNumberFormatLocal("dd-mm-yyyy");

  // Auto fit the columns of range M:M
  selectedSheet.getRange("M:M").getFormat().autofitColumns();

  // Create a dynamic range from A1 to the last row of column AC for borders
  let range = selectedSheet.getRange("A1:AC" + lastRow);

  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);

  // Auto fit the columns of range A1:AC(lastRow) on selectedSheet
  range.getFormat().autofitColumns();
}
I created this code using chatgpt for a report we do in my office , how to show it as a project I used and add in my resume

Initial commit – added Excel Office Script for automated reporting
