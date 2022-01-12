# Contact-Tracing

This repository contains tools to assist contact tracing students through the timetable in Syllabus+.

## contact-tracing.vba

This is a VBA script for identifying clusters of infections.

Note that all of the tools here report on **Student Set** objects, *not* **Student** objects.

### How to use

- Copy and paste this script into the Excel code editor for an Excel workbook with two worksheets.
- The worksheets **must** be called `Data to Check` and `Results`.
- The list of student hostkeys you want to check must be in column `A` of `Data to Check`, starting at row `2`. This leaves the first row free for your own labelling.
- The start and end date of the window you want to check go in cells `E2` and `F2` respectively of `Data to Check`. These are treated as date and time objects, so it's best to explicit and include the time as well as the date. The `dd/mm/yyyy hh:mm` format is recommended.
- When you run the macro, the data entered so far will be used to populate `Results`. You'll get a grid, with one row per student and one column for activity studied by one or more students in the specified timeframe.
- Where a student is attending a particular activity, the intersecting row/column will be filled with an **X**. This then gives you a simple view of whether or not multiple students attended the same activity.


## Student Contact Tracing.rdl and Staff Contact Tracing.rdl

These are SSRS report files to be used with Scientia Report Manager. They should work "out of the box" but you may want to change the institution branding and boilerplate text.

### How to use

These reports can be run directly if you have an SSRS install, but they are also designed to be used with Scientia Report Manager. The report requires three parameters to be configured in Report Manager:

- `StartDateTime` which is a Date parameter.
- `EndDateTime` which is a Date parameter.
- `StudentSetId` which is a list of student set IDs.

**IMPORTANT** - the date parameters use midnight of the selected day, so for the end date you want to pick the day *after* the last day of interest - e.g. if you want to report on 1st to 6th January, select 1st to _7th_ January.

The following SQL will work for the `StudentSetId` parameter, and will let you pick ANY of the student sets in S+. 
`SELECT {HostKey: DISPLAYNAME }, {id : VALUE } FROM rdowner.V_STUDENTSET ORDER BY HostKey ASC`
