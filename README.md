# sqlweb

![License](https://img.shields.io/badge/license-Apache%202.0-green)
![Version](https://img.shields.io/badge/version-1.02-yellow)
![Language](https://img.shields.io/badge/Language-ASP%20Classic-blue)
![IIS](https://img.shields.io/badge/IIS%20version-Any-blue)
![OS](https://img.shields.io/badge/OS-Any%20from%20Windows%202000-blue)

This is framework to create web site by using only SQL queries. 

If you have never created a front-end to databases or reports for the web before, then it's time to try it with this framework.

## Sample Page and Code

![Page](https://github.com/vku1/sqlweb/blob/main/sample_page.jpg)
![Code](https://github.com/vku1/sqlweb/blob/main/sample_code.jpg)

## Environment and Knowledges to start?

- You know what is the database, may write SQL queries and understand tables relations and keys
- Has PC/Server with operating system where Internet Information Server (IIS) may be installed
- Has any database to which you can connect using 32bit driver and connection string listed on [connectionstrings.com](https://www.connectionstrings.com)
- Any code editor with asp/sql syntax highlight

## What You get from the box?

- single page application (all reports/forms/filters in one page)
- unified style for html elements (5 color schemas - 1 style)
- vulnerability check (URL and forms check)
- navigation menu 
- data table
- data filters
- pagination
- totals for numbers
- columns sort
- export to excel
- data operations (add/insert one or multiply records, edit/update one record)
- database names/fields substitution
- debug

## Benefits

- The simpliest page needs only one variable which is also SQL query
- The hardest page with full possible functionality You may get using 15 variables (one variable is query or constraint or hybrid)

## Limitations

- Only 32bit odbc drivers supported
- You can't delete records
- All tables must have 1 autoincrement ID column (id name can be anything) for ADD and EDIT operations.
- Some odbc drivers support read only mode
- Special local symbols not supported by the default

## First configuration

Before you start, check if your database has 32-bit drivers. Without them, all further actions are meaningless. 

Install IIS for [Windows XP, 2000, 2003](http://www.shotdev.com/asp/asp-installation/install-iis-windows-xp-2000-2003/),[Windows 7,Vista,8,8.1, for Windows Server 2008, 2008 r2, 2012, 2012 r2](https://docs.microsoft.com/en-us/iis/application-frameworks/running-classic-asp-applications-on-iis-7-and-iis-8/classic-asp-not-installed-by-default-on-iis).  
[Windows 10](https://docs.microsoft.com/en-us/answers/questions/370931/does-iis-in-windows-10-support-an-asp-web-site.html).

Open IIS, find default application pool and set parameter "using 32 bit applications" to True.

Place [sqlsite.asp](https://github.com/vku1/sqlweb/blob/main/sqlsite.asp) and [global.asa](https://github.com/vku1/sqlweb/blob/main/global.asp) to iis directory c:\inetpub\wwwroot\.

Install 32 bit odbc driver for Your database. 
Check if it is present in odbc drivers list. 
For 32 bit OS go to Control Panel\All Control Panel Items\Administrative Tools\Data sources ODBC, 
for 64 bit OS run this file C:\WINDOWS\syswow64\odbcad32.exe

Try to make test odbc connection to your database using proper driver. If test is OK then 
visit [connectionstrings.com](https://www.connectionstrings.com) and try to find correct string and write it to the global.asa.

Open sqlsite.asp and change variables, listed below, to proper values you want. Before each variable there is short help attached directly in code.
- g_DefaultPageCode
- g_PortalName
- g_page_datasource
- g_columns_start_bracket
- g_columns_end_bracket
- g_DateFromTextToSQL
- g_DateTimeFromTextToSQL

## Debug

Errors checking is divided into 2 parts. First part can be accessed from Debug menu item while the application is running. 
And other errors you can get directly from iis logs in folder C:\inetpub\logs\LogFiles\W3SVC1 or W3SVC2 and so on. 
