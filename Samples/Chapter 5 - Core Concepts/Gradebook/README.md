### Real-world example of multiple `sync`-calls

**SCENARIO: A grade book tool**:

Imagine you are a math teacher, whose grade book is stored in Excel.  Each row represents a student, and each column represents an assignment.

![](http://buildingofficeaddins.com/wp-content/uploads/Gradebook.jpg)

You are preparing for a parent-teacher conference.  When you meet with each parent, **you want to be able to show them *just* the data for their child, and filter the view to just assignments where the student got a *less-than-80% grade*** (so that you can discuss areas for improvement).  Let's assume you'll be doing all the filtering in-place: you're not exporting the data to new sheets or anything like that, it's purely something that will be used for discussion with one set of parents, before being re-filtered for the next set of parents.


To give a concrete example:  for the three assignments in the image above (and let's assume there's a whole bunch more, omitted for brevity), when talking with Matthias D'Armon's parents, the teacher would only want to show columns A (to see the student's name) and columns B & D from the assignments category.  Column C, meanwhile, does *not* need to be shown, as it's the one assignment were Matthias got an over-80% grade (*might I say he had a "Eureka!" moment?*) -- and so this isn't a problem area that the teacher needs to discuss with Matthias' parents.

This scenario is reasonably similar to the population-analysis scenario -- but with the notable exception that here you only want to look at *one student's* data, which represents a small fraction of the data on the sheet.  I would argue that transferring a whole bunch of unneeded data is even worse than doing an extra sync, so let's see how we can do this task most efficiently, even if it means fudging a bit on the minimal-syncs principle for the sake of the avoiding-reading-copious-amounts-of-unneeded-data principle.

I think the most efficient breakdown of tasks is as follows:



**STEP 1: Use the left column to find the student name**.

First off, we need to find the student, so we can load just his/her data. This could be done by creating a Range object corresponding to `A:A`, trimming it down to just its used range (to reduce the one-million-plus cells into something that makes sense), and loading the cell values. However, in this particular case, we have even more efficient means: we can use an Excel function -- invoked from JS -- to do the lookup for us. The Excel function that fits this particular scenario is **=`MATCH(...)`**, which looks up a particular value and returns its row number. We'll assign the function result to a variable and load its "value" property.

> Side-note regarding invoking `context.workbook.functions.match(...)`:  As noted in the intro chapter, the one exception to the JavaScript API's 0-based indexing is when interacting with *numbers that are seen by the end-user*, or *parameters or function results of Excel formulas*. Hence, when we receive the result of `match(...)`, it will be 1-indexed, so we'll need to subtract 1 in order to interface correctly with the rest of the APIs.

Since we can't proceed with any further operations without knowing the row number (and it, in turn, can't be read without first doing a `sync`), perform a **sync**.



**STEP 2: Retrieve the appropriate row, and request to load values**

From the row number, retrieve the appropriate row, trim it down to just the used range (there is no sense in loading values for all 16,384 columns, when likely it's only a few dozen that are used), and load the cell values.

Again, without reading back the values, there is nothing further we can do.  So, **sync**.



**STEP 3: Do the processing & document-manipulation**

Having retrieved the values, hide any column where the score is equal or greater to 80%, since the teacher's goal is only to discuss problem areas in this meeting (we'll assume that complementing *good* grades will have been done separately). This is where we finally issue a bunch of *write* calls to the object model, where we're manipulating the document, instead of just reading from it. So, to *commit* the pending queue of changes, do a final **sync**.


This seems like a reasonable plan, so let's code it up.  The code in this folder shows you how to accomplish this scenario, both with TypeScript 2.1 (i.e., with `async/await`), and with just plain ES5 JavaScript.