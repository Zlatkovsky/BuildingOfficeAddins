function filterStudentGrades() {
    Excel.run(function(context) {
        // #1: Find the row that matches the name of the student:

        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var nameColumn = sheet.getRange("A:A");
        
        var studentName = $("#student-name").val();
        var matchingRowNum = context.workbook.functions.match(
            studentName, nameColumn, 0 /*exact match*/);
        matchingRowNum.load("value");

        var studentRow; // declared here for passing between "`.then`"s
        
        return context.sync()
            .then(function() {
                // #2: Load the cell values (filtered to just the
                //     used range, to minimize data-transfer)

                studentRow = sheet.getCell(matchingRowNum.value - 1, 0)
                    .getEntireRow().getUsedRange();
                studentRow.load("values");
            })
            .then(context.sync)
            .then(function() {
                // Hide all rows except header ones and the student row
                var cellB2AndOnward = sheet.getUsedRange()
                    .getOffsetRange(1, 1);
                cellB2AndOnward.rowHidden = true
                cellB2AndOnward.columnHidden = true;
                studentRow.rowHidden = false;

                for (var c = 0; c < studentRow.values[0].length; c++) {
                    if (studentRow.values[0][c] < 80) {
                        studentRow.getColumn(c).columnHidden = false;
                    }
                }

                studentRow.getCell(0, 0).select();
            })
            .then(context.sync);
    }).catch(OfficeHelpers.Utilities.log);
}