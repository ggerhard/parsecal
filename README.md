# PARSECAL

A simple ruby script to read and parse a google calendar and create an Excel Timesheet.

## Prerequisites:
Set up a google calendar for your time tracking project and export the iCal URL. To do so, open the google calendar page, select "settings", "calendars" and select your calendar. On the calendars detail page, click the private iCal URL Button (see image below).

![ICAL URL](img/iCal.PNG)

Set the variable "*google_cal_url*" in parsecal.rb to this iCal URL (line 10).

The script requires the [icalendar](http://icalendar.rubyforge.org/) gem:

    sudo gem install icalendar

Also [axlsx](https://github.com/randym/axlsx) spreadsheet library is required:

    sudo gem install axlsx

(make sure zlib1g-dev is installed on linux)

(The ["spreadsheet"](https://github.com/zdavatz/spreadsheet) turned out quite disappointing for this purpose. Formulas did not work and there was very little documentation)
