## Book release notes

Updates to the book are free for all existing readers (which is what lean-publishing, and the concept of an evergreen book, is all about!).  Simply go to **<https://leanpub.com/user_dashboard/library>**, select the book, and download your newly-updated copy!

For any issues or topic requests, please file an issue on <http://buildingofficeaddins.com/issues>.

&nbsp;

### Version 1.6 (*August 26, 2017*) [261 pages]

* Major re-structuring of the book, splitting out the former chapters 5 & 6 -- which were bursting at the seams -- into a bunch of smaller chapters.  Also streamlined the rest of the book, moving topics that were less immediately-necessary (e.g., API versioning) further towards the back of the book, to make it faster to get started.
* As part of the getting-started chapter, added a section for my recommendations of "*The optimal dev environment*"
* Expanded the section on "*Handling errors*.
* Added runnable Script Lab snippets to "*Canonical code sample: reading data and performing actions on the document*" (including a refactored version that splits the task into multiple subroutines, and a plain ES5 JavaScript variant).

&nbsp;

### Version 1.5 (*August 2, 2017*) [249 pages]

* By popular demand, added a massively-detailed (15+ pages) and example-filled section on "*Using objects outside the "linear" `Excel.run` or `Word.run` flow (e.g., in a button-click callback, in a `setInterval`, etc.)*".  Its subsections include:
  * Re-hydrating an existing Request Context: the overall pattern, proper error-handling, `object.track`, and cleanup of tracked objects]
  * A common, and infuriatingly silent, mistake:  queueing up actions on the wrong request context
  * Resuming with multiple objects
  * Why can't we have a single global request context, and be one happy family?
* Addressed a couple of [small reader-reported issues](https://github.com/Zlatkovsky/BuildingOfficeAddins/milestone/8?closed=1).

&nbsp;

### Version 1.4 (*July 19, 2017*) [232 pages]

* Added a detailed section for how to check whether an object exists (e.g., whether there's a worksheet by a particular name, whether two ranges intersect, etc). The section also talks about the powerful but unusual "null object" pattern, used in methods and properties suffixed with `*OrNullObject`.
* Added a section with links to "*API documentation resources*".
* Addressed a couple of [small reader-reported issues](https://github.com/Zlatkovsky/BuildingOfficeAddins/milestone/7?closed=1).

&nbsp;

### Version 1.3 (*May 23, 2017*)
 
* Added an entire section devoted to the `PropertyNotLoaded` error.
* Added info about [Script Lab](https://aka.ms/scriptlab), a playground tool to easily try out code samples described in this book
* Addressed a variety of [small reader-reported issues](https://github.com/Zlatkovsky/BuildingOfficeAddins/milestone/5?closed=1).

&nbsp;

### Version 1.2  (*Feb 20, 2017*)

* Added a topic on the different flavors of Office 2016 / Office 365 -- and the practical implications for developers.
* Added a topic on API Versioning and Requirement Sets.
* Greatly expanded the "TypeScript-based Add-ins" topic, adding instructions for the updated Yeoman generator.
* Added a topic for attaching the debugger to Add-ins (breakpoints, DOM explorer, etc.)
* Added a link to the book's companion Twitter account.
* Addressed a number of other [reader-reported issues](https://github.com/Zlatkovsky/BuildingOfficeAddins/milestone/4?closed=1).

&nbsp;

### Version 1.1  (*Jan 22, 2017*)

* Re-pivoted the book around TypeScript and the `async/await` syntax. Moved the JS-specific content to a separate Appendix
* Added an information-packed JavaScript & TypeScript crash-course, tailored specifically at Office.js concepts, for those who are new to the world of JS/TS.
* With TypeScript now a first-class citizen of the book, added "*Getting started with building TypeScript-based Add-ins*" (section 3.2).  I expect to continue to expand this section in future releases.
* Added an in-depth explanation of the internal workings of the Office.js pipeline and proxy-object model.  See "*Implementation details, for those who want to know how it really works*" (section 5.5).
* Re-arranged, edited, and added to the content of the "*Office.js APIs: Core concepts*" chapter (chapter 5).
* Added links to downloadable code samples, for a few of the larger samples.
* Addressed a variety of smaller feedback items.

&nbsp;

### Updates are free for existing readers

Updates to the book are free for all existing readers (which is what lean-publishing, and the concept of an evergreen book, is all about!).  Simply go to <https://leanpub.com/user_dashboard/library>, select the book, and download your newly-updated copy!

If you haven't purchased the book yet, you can do so at <https://leanpub.com/buildingofficeaddins>.
