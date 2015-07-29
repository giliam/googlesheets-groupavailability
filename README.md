# Google Spreadsheets Add-on - Group Availability

What is it for?
---

You are a group and you need to know who's available on a specific date ? You are a support team and you have to offer a non-stop support on a specific period of time ?

This addon could help. You only need a Google Spreadsheet shared with one unique sheet called Utils.

Initialization
---

In this sheet, you need to add the following lines:
* A1 -> the title: "Number of members"
* A2 -> the value: "6" (for example)
* A3 -> the title: "Name of the members"
* Row 4 -> the different names, one for each cell
* A5 -> the title: "Default values for the member"
* Row 6 -> the default value for the member corresponding on the 4th row
* A7 -> the emails to notifiy when nobody is ready
* Row 8 -> the values of the emails (< 100 emails)

Use
---

You can use the SupportMenu to show the Sidebar. In this sidebar, you can add a month of a certain year or add your availability for this day or the day after.


Switch language
----

In the code, you can change the name of the months. You only have to edit the first few lines with the months label.