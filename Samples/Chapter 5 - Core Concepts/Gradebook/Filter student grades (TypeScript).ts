function filterStudentGrades() {
    Excel.run(async (context) => {
        // #1: Find the row that matches the name of the student:

        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let nameColumn = sheet.getRange("A:A");
        
        let studentName = $("#student-name").val();
        let matchingRowNum = context.workbook.functions.match(
            studentName, nameColumn, 0 /*exact match*/);
        matchingRowNum.load("value");
        
        await context.sync()


        // #2: Load the cell values (filtered to just the
        //     used range, to minimize data-transfer)

        let studentRow = sheet.getCell(matchingRowNum.value - 1, 0)
            .getEntireRow().getUsedRange().load("values");

        await context.sync();

        // Hide all cells except the header ones and the student row
        let cellB2AndOnward = sheet.getUsedRange()
            .getOffsetRange(1, 1).getResizedRange(-1, -1);
        cellB2AndOnward.rowHidden = true
        cellB2AndOnward.columnHidden = true;
        studentRow.rowHidden = false;

        // Turn the visiiblity back on for columns with low grades
        for (let c = 1; c < studentRow.values[0].length; c++) {
            if (studentRow.values[0][c] < 80) {
                studentRow.getColumn(c).columnHidden = false;
            }
        }

        studentRow.getCell(0, 0).select();

        await context.sync();

    }).catch(OfficeHelpers.Utilities.log);
}