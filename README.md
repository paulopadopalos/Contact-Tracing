# Contact-Tracing

This repository contains tools to assist contact tracing students through the timetable in Syllabus+.

## contact-tracing.vba

This is a VBA script for identifying clusters of infections. 

### How to use

- Copy and paste this script into the Excel code editor for an Excel workbook with two worksheets.
- The worksheets **must** be called `Data to Check` and `Results`.
- The list of student hostkeys you want to check must be in column `A` of `Data to Check`, starting at row `2`. This leaves the first row free for your own labelling.
- The start and end date of the window you want to check go in cells `E2` and `F2` respectively of `Data to Check`. These are treated as date and time objects, so it's best to explicit and include the time as well as the date. The `dd/mm/yyyy hh:mm` format is recommended.
- When you run the macro, the data entered so far will be used to populate `Results`. You'll get a grid, with one row per student and one column for activity studied by one or more students in the specified timeframe.
- Where a student is attending a particular activity, the intersecting row/column will be filled with an **X**. This then gives you a simple view of whether or not multiple students attended the same activity.
