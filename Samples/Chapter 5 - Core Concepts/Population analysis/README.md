## Canonical code sample: reading data and performing actions on the document

To set the context for the sort of code you'll be writing, let's take a very simple but canonical example of an automation task.  This particular example will use Excel, but the exact same concepts apply to any of the other applications (Word, OneNote) that have adopted the new host-specific/Office 2016+ API model.

**Scenario**:  Imagine I have some data on the population of the top cities in the United States, taken from "Top 50 Cities in the U.S. by Population and Rank" at <http://www.infoplease.com/ipa/a0763098.html>.  The data -- headers and all, just like I found it on the website -- describes the population over the last 20+ years.

Let's say the data is imported into Excel, into a table called "PopulationData". The table could just as easily have been a named range, or even just a selection -- but having it be a table makes it possible to address columns by name rather than index.  Tables are also very handy for end-users, as they can filter and sort them very easily. Here is a screenshot of a portion of the table:

![The population data, imported into an Excel table](http://buildingofficeaddins.com/wp-content/uploads/Core-Concepts-Original-Population-Table-Partial.jpg)


**Now, suppose my task is to find the top 10 cities that have experienced the most growth (in absolute numbers) since 1990.  How would I do that?**

***

The code to show you how to do this -- both with the TypeScript `async/await` syntax, and using plain ES5 JavaScript, can be found in in the following Script Lab snippets (see [import instructions](https://github.com/OfficeDev/script-lab/blob/master/README.md#import)):

1. (Recommended) TypeScript code, refactored to split the scenario into multiple subroutines: **fb8913d0a899c88e3ea82773a135dfd0**

2. TypeScript code, all in a single function: **aa81d73587f62f35e46ad6a904bb20df**

3. JavaScript code:  **98d04bc5293e027c84c8c03741698a94**
