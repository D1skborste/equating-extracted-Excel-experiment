function main(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getActiveWorksheet();

    // Month name lookup array
    const monthNames = [
        "", "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];

    // Find the last row of data based on used range
    let lastRow = selectedSheet.getUsedRange().getRowCount();

    if (lastRow < 2) {
        lastRow = 2;
    }

    let sourceTableRange = `B1:I${lastRow}`;

    // 1. Set formulas in the header and first data row
    selectedSheet.getRange("G1:I2").setFormulas([
        ["Felmarginal", "Week", "Month"],
        ["=D2/E2-1", "=WEEKNUM(C2)", "=MONTH(C2)"]
    ]);

    // 2. Auto fill formulas down to the last row dynamically
    selectedSheet.getRange("G2:I2").autoFill(`G2:I${lastRow}`, ExcelScript.AutoFillType.fillDefault);

    // 3. Apply percentage formatting
    selectedSheet.getRange(`G2:G${lastRow}`).setNumberFormat("0.00%");

    // 4. Extract unique months
    let monthValues = selectedSheet.getRange(`I2:I${lastRow}`).getValues();
    let uniqueMonths: string[] = [];

    for (let row of monthValues) {
        let monthVal = String(row[0]);
        if (monthVal && !uniqueMonths.includes(monthVal)) {
            uniqueMonths.push(monthVal);
        }
    }

    // Sort months numerically
    uniqueMonths.sort((a, b) => Number(a) - Number(b));

    // STEP A: Delete all existing Pivot Tables in the sheet first
    let pivotTables = selectedSheet.getPivotTables();
    for (let pivot of pivotTables) {
        pivot.delete();
    }

    // STEP B: Now safely clear Column K (headers and titles)
    selectedSheet.getRange("K:K").clear();

    // 5. Dynamically create and place Pivot Tables
    let currentStartRow = 2;
    let startColumn = "K";
    const blankSeparationRows = 4;

    uniqueMonths.forEach((month) => {
        let monthNum = Number(month);
        let monthName = monthNames[monthNum] || `Month ${month}`;
        let tableName = `PivotTable_Month_${month}`;

        // Add visual title row above the Pivot Table
        let titleCell = selectedSheet.getRange(`${startColumn}${currentStartRow}`);
        titleCell.setValue(`Report for ${monthName}`);
        titleCell.getFormat().getFont().setBold(true);

        // Place the pivot table 1 row below the header
        let destinationCell = selectedSheet.getRange(`${startColumn}${currentStartRow + 1}`);
        let pivot = workbook.addPivotTable(tableName, sourceTableRange, destinationCell);

        // Configure Pivot Table layout
        pivot.addRowHierarchy(pivot.getHierarchy("Namn"));
        pivot.addColumnHierarchy(pivot.getHierarchy("Week"));

        // Set layout to Tabular form
        pivot.getLayout().setLayoutType(ExcelScript.PivotLayoutType.tabular);

        pivot.addDataHierarchy(pivot.getHierarchy("Personec"));
        pivot.addDataHierarchy(pivot.getHierarchy("EAM"));
        pivot.addDataHierarchy(pivot.getHierarchy("Diff (P-E)"));

        let felmarginalHierarchy = pivot.addDataHierarchy(pivot.getHierarchy("Felmarginal"));
        felmarginalHierarchy.setSummarizeBy(ExcelScript.AggregationFunction.average);
        felmarginalHierarchy.setNumberFormat("0.00%");

        // Apply filter
        let monthFilter = pivot.addFilterHierarchy(pivot.getHierarchy("Month"));
        monthFilter.getFields()[0].applyFilter({
            manualFilter: {
                selectedItems: [month]
            }
        });

        // Dynamic Calculation: Read actual rendered height from layout
        let pivotRange = pivot.getLayout().getRange();
        let pivotRowCount = pivotRange.getRowCount();

        // Advance start position
        currentStartRow = currentStartRow + 1 + pivotRowCount + blankSeparationRows;
    });
}
