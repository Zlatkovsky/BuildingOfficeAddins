## Canonical code sample: reading data and performing actions on the document

To set the context for the sort of code you'll be writing, let's take a very simple but canonical example of an automation task.  This particular example will use Excel, but the exact same concepts apply to any of the other applications (Word, OneNote) that have adopted the new host-specific/Office 2016+ API model.

**Scenario**:  Imagine I have some data on the population of the top cities in the United States, taken from "Top 50 Cities in the U.S. by Population and Rank" at <http://www.infoplease.com/ipa/a0763098.html>.  The data -- headers and all, just like I found it on the website -- describes the population over the last 20+ years.

Let's say the data is imported into Excel, into a table called "PopulationData". The table could just as easily have been a named range, or even just a selection -- but having it be a table makes it possible to address columns by name rather than index.  Tables are also very handy for end-users, as they can filter and sort them very easily. Here is a screenshot of a portion of the table:

![The population data, imported into an Excel table](http://buildingofficeaddins.com/wp-content/uploads/Core-Concepts-Original-Population-Table-Partial.jpg)


**Now, suppose my task is to find the top 10 cities that have experienced the most growth (in absolute numbers) since 1990.  How would I do that?**

***

The code in this folder shows you how to do just that -- both in TypeScript 2.1 (with `async/await`) and using plain ES5 JavaScript.
