# Sheets Calendar Sync

Allows for making a calendar from a google spreadsheet

The created calendar will be named the same as the current sheet.

Heres the expected strucure of the sheet

Starting date* | -** | EndingDate | StartingTime* | EndingTime* | Title | Place* | Comments*
---------------|-----|------------|---------------|-------------|-------|--------|-----------

If no startingdate is defined program assumes all day event; unless the background color is white and `defaultTime != null`, in which case the time is default time.

Default values can be changed in the top of the program.
