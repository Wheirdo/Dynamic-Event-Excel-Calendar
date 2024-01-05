# Dynamic-Event-Excel-Calendar
It's an excel sheet that is a calendar that allows you to put in recurring events (Birthdays, Holidays, etc.) into the 'Schedule' sheet, and have them appear in the calendar automatically. You can use the arrow button[^1] located under 'Calendar Settings' in the January sheet to change the year. This will automatically move the events to their new day of the week, and increment birthdays and anniversarys to be correct with the new  year(e.g. 21st Birthday -> 22nd Birthday)

[^1]: A Spin Button to be precise

# How to Use

- To to the "Schedule" Tab
- Scroll to the bottom of the table, and find an empty row
- Enter the day in the first column
- Enter the month in the second column
- Enter the year in the third column
- Enter the name of the event in the fourth column
- For the fifth column, find the drop down menu in the bottom left corner of the cell, and select your option
 - Or, you can simply type "Birthday", "Anniversary", "Memorial", "Holiday", or "Floating Holiday"

And that's it! The event should now be on the calendar automatically.

# Where is Mardi Gras/Ash Wednesday/Good Friday/Easter?

The way the date of Easter is calculated is ... complicated. The very basic explanation is that the church wants Easter to line up with the Jewish holiday Passover. Passover is calculated using the Hebrew calendar, which is lunar. The Gregorian Calendar, the one that most of modern society uses, is a solar calendar. 
So the math for calculating the date of Easter is very complicated. And since Good Friday, Ash Wednesday, and Mardi Gras all use Easter for their date calculation, they also aren’t on the calendar.
This is something I'd like to get correctly, but just don't have the time for right now.


# References

The visual template for this calendar can be found [here](https://templates.office.com/en-us/any-year-calendar-1-month-per-tab-tm02930051) (note that it does require Mircosoft 365). This template included the spin button that changes the year, and the code that updates the location of the days.

The way I linked the month sheets to the schedule sheet was using a tutorial by Rubayed Razib Suprov, which you can find [here](https://www.exceldemy.com/how-to-create-a-schedule-in-excel-that-updates-automatically)
