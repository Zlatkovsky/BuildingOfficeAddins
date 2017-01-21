function analyzePopulationGrowth() {
    Excel.run(async (context) => {
        // Create proxy objects to represent entities that are
        // in the actual workbook. More information on proxy objects 
        // will be presented in the very next section of this chapter.

        let originalTable = context.workbook.tables
            .getItem("PopulationTable");

        let nameColumn = originalTable.columns.getItem("City");
        let latestDataColumn = originalTable.columns.getItem(
            "7/1/2014 population estimate");
        let earliestDataColumn = originalTable.columns.getItem(
            "4/1/1990 census population");

        // Now, queue up a command to load the values for each of
        // the columns we want to read from.  Note that the actual 
        // fetching and returning of the values is deferred
        // until a "context.sync()".

        nameColumn.load("values");
        latestDataColumn.load("values");
        earliestDataColumn.load("values");

        await context.sync();


        // Create an in-memory data representation, using an
        // array with JSON objects representing each city.
        let citiesData: Array<{name: string, growth: number}> = [];

        // Start at i = 1 (that is, 2nd row of the table --
        // remember the 0-indexing) in order to skip the header.
        for (let i = 1; i < nameColumn.values.length; i++) {
            let name = nameColumn.values[i][0];

            // Note that because "values" is a 2D array
            // (even if, in this case, it's just a single 
            //  column), extract the 0th element of each row.
            let pop1990 = earliestDataColumn.values[i][0];
            let popLatest = latestDataColumn.values[i][0];

            // A couple of the cities don't have data for 1990,
            // so skip over those.
            if (isNaN(pop1990) || isNaN(popLatest)) {
                console.log('Skipping "' + name + '"');
            }

            let growth = popLatest - pop1990;
            citiesData.push({name: name, growth: growth});
        }

        let sorted = citiesData.sort((city1, city2) => {
            return city2.growth - city1.growth;
            // Note the opposite order from the usual
            // "first minus second" -- because want to sort in
            // descending order rather than ascending.
        });
        let top10 = sorted.slice(0, 10);

        // Now that we've computed the data, create a new worksheet
        // for the output, and write in the data:
        let sheetTitle = "Top 10 Growing Cities";
        let sheetHeaderTitle = "Population Growth 1990 - 2014"; 
        let tableCategories = ["Rank", "City", "Population Growth"];
        let outputSheet = context.workbook.worksheets.add(sheetTitle);

        let reportStartCell = outputSheet.getRange("B2");
        reportStartCell.values = [[sheetHeaderTitle]];
        reportStartCell.format.font.bold = true;
        reportStartCell.format.font.size = 14;
        reportStartCell.getResizedRange
            (0, tableCategories.length - 1).merge();

        let tableHeader = reportStartCell.getOffsetRange(2, 0)
            .getResizedRange(0, tableCategories.length - 1);
        tableHeader.values = [ tableCategories ];
        let table = outputSheet.tables.add(
            tableHeader, true /*hasHeaders*/);

        for (let i = 0; i < top10.length; i++) {
            let cityData = top10[i];
            table.rows.add(
                null /* null means "add to end" */,
                [[i + 1, cityData.name, cityData.growth]]);

            // Note: even though adding just a single row,
            // the API still expects a 2D array (for 
            // consistency and with Range.values)
        }

        // Auto-fit the column widths, and set uniform 
        // thousands-separator number-formatting on the
        // "Population" column of the table.
        table.getRange().getEntireColumn().format.autofitColumns();
        table.getDataBodyRange().getLastColumn()
            .numberFormat = [["#,##"]];


        // Finally, with the table in place, add a chart:
        let fullTableRange = table.getRange();

        // For the chart, no need to show the "Rank", so only use the
        //     column with the city's name, and expand it one column
        //     to the right to include the population data as well.
        let dataRangeForChart = fullTableRange
            .getColumn(1).getResizedRange(0, 1);

        let chart = outputSheet.charts.add(
            Excel.ChartType.columnClustered,
            dataRangeForChart,
            Excel.ChartSeriesBy.columns);

        chart.title.text = "Population Growth between 1990 and 2014";

        // Position the chart to start below the table, occupy
        // the full table width, and be 15 rows tall
        let chartPositionStart = fullTableRange
            .getLastRow().getOffsetRange(2, 0);
        chart.setPosition(chartPositionStart,
            chartPositionStart.getOffsetRange(14, 0));

        outputSheet.activate();

        await context.sync();

    }).catch((error) => {
        console.log(error);
        // Log additional information, if applicable:
        if (error instanceof OfficeExtension.Error) {
            console.log(error.debugInfo);
        }
    });
}
