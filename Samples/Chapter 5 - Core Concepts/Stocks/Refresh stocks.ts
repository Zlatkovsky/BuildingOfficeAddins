function refreshStocks() {
    Excel.run(async (context) => {
        console.log("#1: Obtain and loading worksheets");
        let sheets = context.workbook.worksheets.load("name");
        await context.sync();

        console.log("#2: Retrieve stock table & values. " +
            "Also load tables on previous sheet, if any");
        let latestSheet = sheets.items[sheets.items.length - 1];
        let latestTable = latestSheet.tables.getItemAt(0);
        let stocksRange = latestTable.getDataBodyRange()
            .getColumn(0).load("values");
        let previousTables = queueLoadPreviousSheetTables(sheets);
        await context.sync();

        console.log("#3: Issue web request to read stock data");
        let stockNames = stocksRange.values.map(rowData => rowData[0]);
        
        // With this names-list in hand, call the "getStockData"
        // function from an earlier chapter ("Promises Primer"").
        // The function returns a Promise, so await it:
        let stockData = await getStockData(stockNames);
        console.log("Received stock data", stockData);

        console.log("#4: Reload stock names, to ensure no " +
            "changes occurred while waiting on the web request");
        stocksRange.load("values");
        await context.sync();

        console.log("#5: Write in the updated stock prices, " + 
            "and load the new total (and previous, if any)");
        queueUpdateTable(latestTable, stocksRange, stockData);
        let latestTotalCell = latestTable
            .getTotalRowRange().getLastCell().load("values");
        let previousTotalCellIfAny =
            queueLoadPreviousTotalCellIfAny(previousTables);
        await context.sync();

        console.log("#6: If there was a previous 'Total' to " +
            "compare against, color the current one accordingly.");
        queueCellHighlightingIfAny(
            latestTotalCell, previousTotalCellIfAny);
        await context.sync();
        
        console.log(`Done! Data on sheet "${latestSheet.name}" ` +
            "has been refreshed");

    }).catch(OfficeHelpers.Utilities.log);


    // Helper functions

    function queueLoadPreviousSheetTables(
        sheets: Excel.WorksheetCollection
    ) {
        let sheetCount = sheets.items.length;
        if (sheetCount >= 2) {
            return sheets.items[sheetCount - 2].tables.load("name");
        }
        return null;
    }

    function queueUpdateTable(table: Excel.Table,
        stocksRange: Excel.Range, data: any
    ) {
        let pricesToWrite = stocksRange.values.map(row => {
            let stockName = row[0];
            let priceOrEmptyString = data[stockName];
            if (typeof priceOrEmptyString === "undefined") {
                priceOrEmptyString = "";
            }
            return [priceOrEmptyString];
        });

        let priceColumn = table.columns.getItem("Price")
            .getRange().getIntersection(stocksRange.getEntireRow());
        priceColumn.values = pricesToWrite;
    }

    function queueLoadPreviousTotalCellIfAny(
        tables: Excel.TableCollection
    ) {
        if (tables && tables.items.length === 1) {
            return tables.items[0].getTotalRowRange()
                .getLastCell().load("values");
        }
        return null;
    }

    function queueCellHighlightingIfAny(
        latestCell: Excel.Range, previousCell: Excel.Range
    ) {
        if (previousCell) {
            let isLatestGreater =
                latestCell.values[0][0] >
                previousCell.values[0][0];

            latestCell.format.fill.color =
                (isLatestGreater ? "#82E0AA" : "#EC7063");     
        } else {
            console.log("Skipped comparison with previous " +
                "total, as there doesn't appear to be one");
        }
    }
}


function getStockData(stockNamesList: string[]) {
    var quotedCommaSeparatedNames = stockNamesList
        .map(function(name) {
            return '"' + name + '"'
        })
        .join(",");

    var url = '//query.yahooapis.com/v1/public/yql';
    var data = 'q=' +
        encodeURIComponent(
            'select * from yahoo.finance.quotes ' +
            'where symbol in (' + quotedCommaSeparatedNames + ')'
        ) +
        "&env=http%3A%2F%2Fdatatables.org" +
        "%2Falltables.env&format=json";

    return new OfficeExtension.Promise((resolve, reject) => {
        $.ajax({
            url: url,
            dataType: 'json',
            data: data,
            timeout: 5000
        })
        .done((result) => {
            console.log("Web request succeeded");
            var stockDataDictionary = {};
            var stockDataArray = result.query.results.quote;
            stockDataArray.forEach(data => {
                var name = data['Symbol']
                var price = data['LastTradePriceOnly'];
                stockDataDictionary[name] = price;
            });

            resolve(stockDataDictionary);
        })
        .fail((error)=> {
            console.log("Web request failed:");
            console.log(url);

            reject(new Error(error.statusText));
        });
    });
}
