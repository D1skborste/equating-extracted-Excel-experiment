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
	let sourceTableRange = `B1:I${lastRow}`;

	// 1. Set formulas in the header and first data row
	selectedSheet.getRange("G1:I2").setFormulas([
		["Felmarginal", "Week", "Month"],
		["=D2/E2-1", "=WEEKNUM(C2)", "=MONTH(C2)"]
	]);

	// 2. Auto fill formulas down to the last row dynamically
	selectedSheet.getRange("G2:I2").autoFill(`G2:I${lastRow}`, ExcelScript.AutoFillType.fillDefault);

	// 3. Apply percentage formatting with 2 decimal places to column G
	selectedSheet.getRange(`G2:G${lastRow}`).setNumberFormat("0.00%");


	// --- Pivot Table 1: Month 1 (e.g., January) ---
	let pivot1 = workbook.addPivotTable("PivotTable_Month1", sourceTableRange, selectedSheet.getRange("K2"));

	pivot1.addRowHierarchy(pivot1.getHierarchy("Month"));
	pivot1.addRowHierarchy(pivot1.getHierarchy("Namn"));
	pivot1.addColumnHierarchy(pivot1.getHierarchy("Week"));

	pivot1.addDataHierarchy(pivot1.getHierarchy("Personec"));
	pivot1.addDataHierarchy(pivot1.getHierarchy("EAM"));
	pivot1.addDataHierarchy(pivot1.getHierarchy("Diff (P-E)"));
	let felmarginalHierarchy = pivot1.addDataHierarchy(pivot1.getHierarchy("Felmarginal"));
	felmarginalHierarchy.setSummarizeBy(ExcelScript.AggregationFunction.average);
	
