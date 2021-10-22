# sqlweb
This is framework to create web site by using only SQL queries. 

If you have never created a front-end to databases or reports for the web before, then it's time to try it with this framework.

## Sample Page and Code
![Page](https://github.com/vku1/sqlweb/blob/main/sample_page.jpg)
![Code](https://github.com/vku1/sqlweb/blob/main/sample_code.jpg)

## Environment and Knowledges to start?
- You know what is the database, may write SQL queries and understand tables relations and keys
- Has PC/Server with operating system where Internet Information Server (IIS) may be installed
- Has any database to which you can connect using connection strings listed on [connectionstrings.com](https://www.connectionstrings.com)

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

## First configuration
- Your database proper connectionstring

variables:
- g_DefaultPageCode
- g_PortalName
- g_page_datasource
- g_columns_start_bracket
- g_columns_end_bracket
- g_DateFromTextToSQL
- g_DateTimeFromTextToSQL

