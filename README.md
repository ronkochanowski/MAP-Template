MAP-Template
============

Word and Excel macro enabled templates that make up a content management system (of sort) used to create a Ministry Action Plan.

Overview

Word.docm file contains the content of the report, used as a basis of functionality.  The content is adjusted to fit each situation that is encountered.  Once the content is modified, the Export_Excel.bas file is processed, which parses through the Word file looking for certain style elements and placing the style name and text into an array.  That array is then transposed into a new Excel file.  Other maintenance functionality is incorpurated as well.

The Excel.xlsm file is used to work at creating a timeline from the content elements from the Word file.  There's much more automation built into the Excel file.  After updating timing labels, a macro is used to format the timeline elements to make them presentable and organized.

A Change event is activated on the spreadsheet so that when a particular cell, or cells, are updated they are reformatted and sorted.

When this formatting and time setting is completed, the user runs another macro that will take the individual timelines that the user has been working with and converts them into a Master Timeline.  It also creates a thrid sheet that can be used to rack and assign the individual elements of the overall project.

Finally, when the timelines are finalized, a last macro is processed that transfers each of the individual timelines back to the appropriate location within the Word document; copies the Master Timeline and Project Tracker over to the Word document; then gives control back to the Word document.

-----------------------
Original code is very crude.  Needs a great deal of updating.  Two seperate tracts need to be taken:
  1. Fix code in both Word and Excel making extensive use of classes.
  2. Rewrite overall project as a database managed system with a web front-end
