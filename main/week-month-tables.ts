function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Find the last row of data based on column C (where your dates/data are located)
  let lastRow = selectedSheet.getUsedRange().getRowCount();

  // Safety check: ensure at least header + 1 data row exists
  if (lastRow < 2) {
    lastRow = 2;
  }

  // Define dynamic ranges based on the detected last row
  let formulaRange = `G2:H${lastRow}`;
  let sourceTableRange = `F1:H${lastRow}`;

  // 1. Set formulas in the header and first data row
  selectedSheet.getRange("G1:H2").setFormulas([
    ["Week", "Month"],
    ["=WEEKNUM(C2)", "=MONTH(C2)"]
  ]);

  // 2. Auto fill formulas down to the last row dynamically
  selectedSheet.getRange("G2:H2").autoFill(`G2:H${lastRow}`, ExcelScript.AutoFillType.fillDefault);

  // 3. Add the first pivot table using the dynamic source range
  let newPivotTable = workbook.addPivotTable(
    "PivotTable3",
    selectedSheet.getRange(sourceTableRange),
    selectedSheet.getRange("J2")
  );

  // Configure PivotTable3 hierarchies
  newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Week"));
  newPivotTable.removeDataHierarchy(newPivotTable.getDataHierarchy("Sum of Week"));
  newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Week"));
  newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Diff"));

  // Set headers and formatting for PivotTable3 section
  selectedSheet.getRange("J1").setValue("Week");
  selectedSheet.getRange("J1").getFormat().getFont().setBold(true);

  // 4. Add the second pivot table using the dynamic source range
  let newPivotTable_1 = workbook.addPivotTable(
    "PivotTable4",
    selectedSheet.getRange(sourceTableRange),
    selectedSheet.getRange("M2")
  );

  // Configure PivotTable4 hierarchies
  newPivotTable_1.addDataHierarchy(newPivotTable_1.getHierarchy("Diff"));
  newPivotTable_1.addRowHierarchy(newPivotTable_1.getHierarchy("Month"));

  // Set header for PivotTable4 section
  selectedSheet.getRange("M1").setValue("Month");
  selectedSheet.getRange("M1").getFormat().getFont().setBold(true);
}
