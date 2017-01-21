## A more complex `context.sync` example:

If the gradebook sample made sense, lets try something with even more `sync`-functions, and that feels even more like a "real-world" scenario.

**Scenario: Stock-tracker and net-worth calculator**

Imagine you have an Excel table, where you keep track of stocks that you own.  You keep multiple copies of the table, one in each sheet, with each sheet corresponding to a particular year (or month, or week... basically, snapshots in time). The last (rightmost) sheet is your *current year's* sheet -- which you want to update based on the current stock values fetched off of the internet. You also want to read the previous sheet's total-stock-worth and compare it with the latest sheet's total. Depending on whether the latest value is better or worse than the previous sheet's value, you will color the cell representing the total in either green or red.

![](http://buildingofficeaddins.com/wp-content/uploads/Stocks-Table.jpg)


As before, before writing the code, let's do a quick outline of the steps (and syncs) involved.

1. First, load the worksheet collection in order to obtain -- after a sync -- references to the last two worksheets. Note that since you can't load a collection with no properties on the children, you have to load at least one property. Let's load the `name` property on the sheets, since it is both short (won't transfer too much data over the wire), and useful for displaying messages in the UI(e.g., "*Data on sheet 'Stocks 2016' successfully refreshed*").  So: load and **sync**.
2. On the "latest" worksheet, get the first table on the worksheet (we'll assume that there is only one table per sheet -- though we could validate, if we really wanted to). Load the values of the leftmost column of the body of the table (the stock names). And again, **sync**.
3. Using the stock names obtained in the previous step, format and issue a web query to obtain the latest price data. This is an asynchronous operation, but not an Office-involved one.  So *no sync*.
4. Because the stocks web-server call may have taken a while -- and depending on how paranoid we are -- it may be best to re-load the table values, in case the user had rearranged or deleted stock names in the process[^not-battleship-excel]. Having re-issued the call for re-loading the stock names, **sync**. 
5. For each stock symbol, add the appropriate values to the price column. With the value-setting dispatched, queue up the loading of both this table's "Total", and the "Total" from the previous worksheet. (We are reading back the data, instead of computing the current total, so that if the "Total" involves formulas or some other Excel wizardry, that read it back correctly, just as it would show on the worksheet). **sync**.
6. Having compared the current and previous totals, assign a green (positive) or red (negative) background to the current total cell, and issue one final **sync** command to commit this action.


The TypeScript file in this folder will show how to accomplish this scenario -- and also one approach for breaking up a task into smaller subroutines.