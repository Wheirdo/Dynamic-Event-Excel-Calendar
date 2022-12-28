# Dynamic-Event-Excel-Calendar
It's an excel sheet that is a calendar that allows you to put in recurring events (Birthdays, Holidays, etc.) into the 'Schedule' sheet, and have them appear in the calendar automatically. You can use the arrow button[^1] located under 'Calendar Settings' in the January sheet to change the year. This will automatically move the events to their new day of the week, and increment birthdays and anniversarys to be correct with the new  year(e.g. 21st Birthday -> 22nd Birthday)

[^1]: A Spin Button to be precise

# How to Use

The Date column should be formated as follows: 

```
=DATE(CalendarYear,M,D)
```

Where M is the Month and D is the Day the event takes place. Don't change 'CalendarYear'; instead, put the year the event started in the Year column on the right.

The Event column gives you a preview of what the event will look like in the calendar. This is what the code looks like for a Birthday event

```
=TEXTJOIN("",FALSE,C3,"'s ",CalendarYear - D3,IF(RIGHT(CalendarYear - D3,1)="1","st",IF(RIGHT(CalendarYear - D3,1)="2","nd",IF(RIGHT(CalendarYear - D3,1)="3","rd","th")))," Birthday")
```

The Name column is the name of the event. For Birthdays, this column would simply contain the name of the person whose birthday it is.

The Year column is the year this event started. For Birthdays, it would be the year the person was born. This is used to calculate their age.

# References

The visual template for this calendar can be found [here](https://templates.office.com/en-us/any-year-calendar-1-month-per-tab-tm02930051) (note that it does require Mircosoft 365). This template included the spin button that changes the year, and the code that updates the location of the days.

The way I linked the month sheets to the schedule sheet was using a tutorial by Rubayed Razib Suprov, which you can find [here](https://www.exceldemy.com/how-to-create-a-schedule-in-excel-that-updates-automatically)
