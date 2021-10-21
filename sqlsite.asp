<!DOCTYPE html>
<html>
<head>
	<title></title>
	<meta name="name" content="SQL WEB framework">
	<meta name="version" content="1.02">
	<meta name="description" content="Single page website based only on SQL queries and asp variables">
	<meta name="keywords" content="single page application, spa, sql, iis, asp, database frontend, front-end, web frontend">
	<meta name="author" content="vku1">
	<meta name="home page" content="https://github.com/vku1/sqlweb">
<% 

	' Adjusting Notepad++ to work with an ASP VBScript project. http://thewayofcoding.com/2016/11/adjusting-notepad-to-work-with-an-asp-vbscript-project/

	'/----- Page Visual Style: Font for all elements and Colors ----
	' You can use any installed in the system font.  
		dim font_ 
		font_ = "Courier" ' "Candara" ' '"Courier" "Verdana" "Arial Narrow" "Courier New" "Calibri" "Tahoma" 
			
		dim theme_color
		dim theme_color_menu_bg
		dim theme_color_menu_hover
		dim theme_color_font
		' to change visual elements colors unlock any line below by removing ' on left 	
		' theme_color = "#5D6D7E": theme_color_menu_bg = "#eee"   : theme_color_menu_hover = "#aaaaaa": theme_color_font = "#778877" : theme_color_table_child="#f2f2f2" ' dark gray
		 theme_color = "#818181": theme_color_menu_bg = "#FaFaFa": theme_color_menu_hover = "#cFcFcF": theme_color_font = "#666666" : theme_color_table_child="#f2f2f2" ' light gray
		' theme_color = "#5499C7": theme_color_menu_bg = "#D6EAF8": theme_color_menu_hover = "#85C1E9": theme_color_font = "#21618C" : theme_color_table_child="#EBF5FB"' blue
		' theme_color = "#52BE80": theme_color_menu_bg = "#A9DFBF": theme_color_menu_hover = "#27AE60": theme_color_font = "#196F3D" : theme_color_table_child="#E8F6F3" ' green
		' theme_color = "#935116": theme_color_menu_bg = "#FDF2E9": theme_color_menu_hover = "#FAE5D3": theme_color_font = "#784212" : theme_color_table_child="#f2f2f2" ' orange
	'\---------------------------------------	

%>
<style type="text/css">

body{
   background-image:url("body.jpg");
   background-size:cover;
   background-repeat:repeat;}

/* Fixed sidenav, full height */
.sidenav {
  height: 100%;  width: 200px;  position: fixed;  z-index: 1;  top: 0;  left: 0;  background-color: <%=theme_color_menu_bg%>;  overflow-x: hidden;  padding-top: 20px; font: 100% <%=font_%>;}

.sidenav img {
width: 35%;
padding: 6px 8px 10px 16px;}


.sidenav-logo {
width:90%;
padding: 0px 0px 10px 14px;
font-size: 20px;
font-weight: bold;}

.sidenav-GlobalObjects {
position: relative; bottom: 0;
width:90%;
padding: 0px 0px 10px 14px;
font-size: 16px;
font-weight: bold;}

/* Style the sidenav links and the dropdown button */
.sidenav a {
  padding: 6px 8px 6px 16px;
  text-decoration: none;
  color: <%=theme_color_font%>;
  display: block;
  border: none;
  background: none;
  width:100%;
  text-align: left;
  cursor: pointer;
  outline: none;}

.dropdown-btn {
  padding: 6px 8px 6px 16px;
  text-decoration: none;
  color: <%=theme_color_font%>;
  display: block;
  border: none;
  background: none;
  width:100%;
  text-align: left;
  cursor: pointer;
  outline: none;  
  
  font-weight:bold;
  font-size: 12px;
  font: 100% <%=font_%>;}

/* On mouse-over */
.sidenav a:hover, .dropdown-btn:hover {
  color: #212121; /*font color for menu items on hover*/
  background-color: <%=theme_color_menu_hover%>;}

/* Add an active class to the active dropdown button */
.active {
  background-color: #f1f1f1; 
  color: #010101;}

/* Dropdown container (hidden by default). Optional: add a lighter background color and some left padding to change the design of the dropdown content */
.dropdown-container {
  display: none;
  background-color: #f1f1f1; /*BG color of dropdown items*/
  padding-left: 12px;
  font-size: 12px;}

/* Optional: Style the caret down icon */
.fa-caret-down {
  float: right;
  padding-right: 8px;
  font-family; <%=font_%>;}

/* Main content */
.main {
  margin-left: 200px; /* Same as the width of the sidenav */
  padding: 0px 10px;}

table.DataTable {border-collapse: collapse;}
table.DataTable caption {background: white; color: <%=theme_color%>;border: 1px solid <%=theme_color%>; font-weight:bold;padding: 10px;font-size:14px;}
table.DataTable th {position: sticky; top: 0; z-index: 100; background-color: <%=theme_color%>;color: white; padding: 3px;} 

table.DataTable tr:nth-child(even) {background-color: <%=theme_color_table_child%>;} 
table.DataTable tr:hover {background-color: <%=theme_color_menu_hover%>;} 
table.DataTable td {border-bottom: 1px solid #ddd; padding: 6px; color:<%=theme_color_font%>; } 

table.tablefilter {table-layout: fixed; border-collapse: collapse;border: 1px solid <%=theme_color%>;}
table.tablefilter caption {background: <%=theme_color%>; color: ivory; font-weight:bold; padding: 6px 1px 4px 1px;font-size:14px;}
table.tablefilter td {padding: 6px;} 

body   { font-family: <%=font_%>;font-size:12px;}
input  { font-family: <%=font_%>; color: <%=theme_color_font%>;}
select { font-family: <%=font_%>;}
.main a:link,  .main a:visited, .main a:not([href]) {background-color: <%=theme_color%>;color: ivory; padding: 5px 7px;text-align: center;text-decoration: none;display: inline-block;}
.main a:hover, .main a:active {background-color: <%=theme_color_menu_hover%>;}

input[type=submit], input[type=reset] 
{
  background-color: <%=theme_color%>;
  border: none;
  color: ivory;
  padding: 5px 5px;
  text-decoration: none;
  margin: 4px 2px;s
  cursor: pointer;
  font: 400 12px <%=font_%>;}
input[type=submit]:hover, input[type=reset]:hover {background-color: <%=theme_color_menu_hover%>;} 

select {
    display: block;
    color: <%=theme_color_font%>;
    line-height: 1.3;
    width: auto;
    margin: 0;
    font: 400 12px <%=font_%>;
	border: 1px solid #aaa;
    border-radius: 2px;
    -moz-appearance: none;
    -webkit-appearance: none;
    --appearance: none;
    background-color: white;
    background-repeat: no-repeat, repeat;
    background-position: right .7em top 50%, 0 0;
    background-size: .65em auto, 100%;}
	
select::-ms-expand {display: none;}

select:hover {
    border-color: #888;}
select:focus {
    border-color: #aaa;
    box-shadow: 0 0 1px 1px <%=theme_color_font%>;
    box-shadow: 0 0 0 3px -moz-mac-focusring;
    color: #222; 
    outline: none;}
</style>

<%  ' -- timestamp generation for export menu item for filename ----
	DIM TMS
	TMS = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2) & "-" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
%>

<script>

function exportToExcel(elem) {
  var table = document.getElementById("DataTable");
  var html = table.outerHTML;
  var url = 'data:application/vnd.ms-excel,' + escape(html); 
  elem.setAttribute("href", url);
  elem.setAttribute("download", "Export<%=TMS%>.xls"); // Choose the file name
  return false;
}

function showNewsInfo() {
  var x = document.getElementById("MsgInfo");
  if (x.style.display === "none") {
    x.style.display = "block";
  } else {
    x.style.display = "none";
  }
}

function showDebugContent() {
  var x = document.getElementById("DebugInfo");
  if (x.style.display === "none") {
    x.style.display = "block";
  } else {
    x.style.display = "none";
  }
}

</script>
</head>
<body>
<%
''for each x in Request.ServerVariables
''  response.write(x & "  " & request.servervariables(x) & "<br>")
''next
''response.end
%>

<%

' /--- Maintenance block ----------------------------------
'  check visitor ip and make actions based on this, as example there is block which can be uncommented in case of maintenance on server
Dim g_clientIP  
	g_clientIP = Request.Servervariables("REMOTE_ADDR") 
	'if g_clientIP<>"localhost" and g_clientIP<>"192.168.1.1" then 
	'	response.write "Dear User, <br>Site temporary unavailable for maintenance purposes.<br>Have a nice day." & g_ClientIP
	'	response.end
	'end if
' \-------------------

' /-- Physical pagename of this file for links generation -------
Dim page_name ' this script physical filename like this sqlsite.asp
									
	page_name = func_getPageName()	
	
' /---- Page number ----------
'If the default page number is ommited, then we redirect script to proper page (by default p=1)
dim page
	page=get_page()
' \--------------												
' / --- Default page which will be displayed if ?p=XX is ommited in url, or if invalid ?p= is set	 
Dim g_DefaultPageCode	
	g_DefaultPageCode = "1"
' \ ------------------

' /---- Web Portal name ---
Dim g_PortalName
	g_PortalName = "San Francisco Purchasing Commodity Data"
' \------------------------

' / --- Site menu generation --------------
' B) Menu generation from text string
dim g_generate_menu_from_string ' YES/NO
' if g_generate_menu_from_string = "YES" then menu for project will be generated from string in variable g_MENU,
' if g_generate_menu_from_string = "NO"  then menu for project will be set manualy by You from block #MANUAL_MENU (find this string below). HTML knowledge needed.
	g_generate_menu_from_string="YES"

dim g_MENU ' global variable for menu. 
'  Menu Structure
' (Page_Name_without_submenu_items):Page_code;
' [Page_Name_with_submenu_items]:{First_submenu_item_name}:page_code_of_first_submenu:{Second_submenu_item_name}:page_code_of_second_submenu;
	g_MENU = ""
	g_MENU = g_MENU & "(Home):1;"
	g_MENU = g_MENU & "[Statistics]:1:{Years}:2:{Departments}:3:{Monthly by the department}:4:{Purchase orders};"
	g_MENU = g_MENU & "[Chinook]:5:{Artists}:6:{Albums};"
' \ ------------------

' /-----INPUT Fields -------------
' Functionality to transform databases field types to proper input field type. Not supported by all the browsers
' For more info: https://developer.mozilla.org/en-US/docs/Learn/Forms/HTML5_input_types
' Recognised: text field, numeric, date, time, datetime-local. 
dim g_use_html5_fields_for_input ' YES/NO '<input type='date|number|text|...'
	g_use_html5_fields_for_input = "YES" 
' \-------------------	

' /--- Maximum number of records on page --------------------	
' Pagination added automaticaly count of records in html table for viewing (pagination block added automaticaly). 
' For each separate page You can set its own count of visible records by using this variable in "page CASE block" 
dim g_page_records_count 
	g_page_records_count = 20 
' \-------------------------------

' /--- Connection string CODE stored in file global.asa for datasource --------------------	
' How it works? You open global.asa file. Find line Application("my_data_source_1") = "ConnectionStringHere"
' and change part between quotes "ConnectionStringHere" to the proper connection string.
' More information you can get on https://www.connectionstrings.com/
' SQL server Example: 
' Application("my_data_source_1")="Provider=SQLNCLI11;server=localhost\SQLEXPRESS;database=YOURDATABASENAME;uid=DB_username;pwd=DB_user_password"
dim g_page_datasource ' value of this variable is connection string to datasaource you want to use 
	g_page_datasource = "sqlweb"
' \-------------------------------

' /---- Brackets for columns with spaces ----------------
' These parameters depend on the type of Database that You use
' They mean the generally accepted characters in which the columns enclose, in the name of which there are spaces.
' For SQL Server if column contain space in it's name the column will looks like [my column name]. Where start bracket is "[", and end bracket is "]".
' For Oracle "my column name"
' For MySql `my column name`
dim g_columns_start_bracket
dim g_columns_end_bracket
	g_columns_start_bracket="[" ' [ for SQL Server , """" for Oracle,PostgreSQL,SQL Server,SQLite , "`" for MySQL,MariaDB
	g_columns_end_bracket="]"   ' ] for SQL Server , """" for Oracle,PostgreSQL,SQL Server,SQLite , "`" for MySQL,MariaDB
' \--------------------

' /---- HTML5 date posting and conversion from string YYYY-MM-DDTHH:MM to proper sql format based on special database engine rules ----------------
Dim g_DateFromTextToSQL   ' added for HTML5 universal date string value
	g_DateFromTextToSQL = "CAST('#DATE#' as Date)"
	
Dim g_DateTimeFromTextToSQL   ' added for HTML5 universal date string value
	g_DateTimeFromTextToSQL = "CAST('#DATE#:00' as DateTime)" ' Incorrect default browser datetime-local value is -> CAST('2021-06-02T08:51' as DateTime), Correct ->  CAST('2021-06-02T08:51:00' as DateTime)

' /---- Columns Beautifier ----------------
' As you know table columns names may have strange names in database, not friendly for end users.To convert these names to 
' good looking, set "g_use_columns_beautifier" parameter to YES and also add all transformations to variable "g_TableColumnsHeadersSubstitution" as pairs: [TableColumnName1;NameOfThisColumn1InReport;TableColumnName2;NameOfThisColumn2InReport;].
' "g_TableColumnsHeadersSubstitution" is text string separated by ";" delimiter, where N element is database table column code and N+1 element is N column Userfriendly name. With a step 2.
' Example g_TableColumnsHeadersSubstitution="id;Identifier;fname;First Name;lname;Last Name;" 
' All founded column names with a code "fname" will be transformed to "First Name". No changes in database will be made.
dim g_use_columns_beautifier ' YES/NO to use this functionality you need to fill variable "g_TableColumnsHeadersSubstitution=" with appropriate values
	g_use_columns_beautifier = "YES"
dim g_use_columns_beautifier_compact ' YES/NO g_use_columns_beautifier_compact variable replaces 'SPACE' symbols in the second parts of pairs in g_TableColumnsHeadersSubstitution with html element '<br>', which is new line analogue. Try it to see effect.
 	g_use_columns_beautifier_compact="YES" 
dim g_TableColumnsHeadersSubstitution
	g_TableColumnsHeadersSubstitution = "" _ 
	& "id;Identification Number;" _
	& "purch_dept;Purchasing department;" _
	& "fiscal_year;Fiscal year;" _
	& "po_count;Purchase orders count;"

Dim g_ColumnsSubstitutionKeyValue
g_ColumnsSubstitutionKeyValue = split(g_TableColumnsHeadersSubstitution,";")
	
' \--------------------	

' /---- Debug info.  ----------------
dim g_debug_flag ' if debug_="YES" then collect info for debugging to print to the same page. DO NOT activate in production! Create copy of page, and rename it in filesystem and in code, and then activate debug.
	g_debug_flag="YES" ' NO
dim g_debug_log ' All debug info will be collected in this variable 
	
' \--------------------	

' /---- Subtotals in table for numeric columns  -------
' This functionality will add SUM row under the table with data where all numeric values in columns will be summaryzed. 
' BUG or Trick:In some cases algorythm can find numbers in columns where the text data stored and make SUM for them.																									
dim g_ApplySubtotalsForNumericColumns
	g_ApplySubtotalsForNumericColumns="YES"
' \--------------------	
	
Dim g_OperationTypeInsertUpdate ' Global variable for sharing operation type on record Insert or Update																						   

' / -------- Page Global variables -------------
Dim g_Table_Caption_and_Info					 ' report or form Name
Dim g_Form_Info_Help							 ' use this to show users some additional information about form, description, comments, news or columns formats or other. To show content use main menu info item.
Dim g_SQL 										 ' sql select from database
Dim g_FilterDropdownsAllowed					 ' Filter enabled or not (YES/NO)
Dim g_FilterDropdownsColumns					 ' Example; select '%' as VendorName,'All vendors' as Vendor from dual union select VendorName,VendorName as Vendor from Vendors'
Dim g_FilterDatalistsColumns					 ' default type for Dropdown is <select><option> tags construction, but you can change it to datalist 																																	   
Dim g_FiltersDefaultValues						 ' select '' as VendorName,'' as Vendor from dual // dual is https://en.wikipedia.org/wiki/DUAL_table, for sql server the same will be -> select '' as VendorName,'' as Vendor -- https://en.wikipedia.org/wiki/DUAL_table
Dim g_TableColumnsSortingAllowed				 ' Allow Columns Sorting by click them (YES/NO)
Dim g_TableColumnsDefaultSorting				 ' Default sorting sql syntah may be very useful for default view in reports (example: "ColumnName1 ASC, ColumnName2 DESC")	
Dim g_TableRowsUpdateAllowed					 ' Allow Update operation on table (YES/NO)
Dim g_TableRowsInsertAllowed					 ' Allow Insert operation on table (YES/NO)
Dim g_DBTableForInsertUpdate					 ' For Insert/Update we need real database table name (it may be only one, unique table name)
Dim g_DBTableIdColumn							 ' For Update we need real database table id column name (it may be only one, unique column)
Dim g_DBTableFieldsListForInsertUpdate
Dim g_DBTableDropdownsForInsertUpdate            ' default type for Dropdown is <select><option> tags construction, but you can change it to datalist . Read func_GetFilterDropdownsIfExist info for this variable.
Dim g_DBTableDatalistsForInsertUpdate            ' change default tag construction from <select><option> to <input list><datalist><option> which support search in it. Very good for long lists.
Dim g_DBTableMultipleDropdownsFieldsForInsert    ' List of these values will be repeated N times while inserting rows
Dim g_TableUpdateInsertLayoutVerticalHorizontal	 ' For Operations Update and Insert data layout vertical or horisontal (V/H). For tables containing more than 10 columns, may be very useful	
' \ --------------------------------------------	


' /--------- Global Variables - Dropdowns in top of menu which values will be automatically applied to all filters and datatables 
' for each variable will be created session variable with value selected  by the user
' after selecting value, you will be redirected to main menu and value will be used in url string+in session variable; 
' at next steps value will be extracted from session variable 
Dim g_GlobalVariables
	'g_GlobalVariables       = func_CreateGlobalVariablesDD("VesselName;select '' vesselname union SELECT vesselname FROM vessels where IsEnabled='YES' order by vesselname") 
Dim g_GlobalVariablesValues
	g_GlobalVariablesValues = func_GetGlobalVariablesValues()
' \-------------------------------------------------------------------------------------------------------------------------------

SELECT CASE cstr(page)
	
	CASE "1" 
		
		g_Table_Caption_and_Info = "Yearly Statistics"
		g_Form_Info_Help = ""	
		
		g_SQL =         " select [fiscal_year],count([purchase_order]) po_count,'<a href=''sqlsite.asp?fiscal_year='+[fiscal_year]+'&purch_dept=%25&p=2''>...</a>' Info from ( "
        g_SQL = g_SQL & " select distinct CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ) [fiscal_year],[purchase_order]  from [test_sqlweb_db].[dbo].[purchasing_commodity] "
        g_SQL = g_SQL & " ) x group by fiscal_year  "
        g_SQL = g_SQL & "  "
        		
		g_FilterDropdownsAllowed = "NO"
		g_FilterDropdownsColumns = "select '%' as statusname, 'All statuses' as comment union select statusname,statuscomment from dbo.status"
		g_FilterDatalistsColumns = ""					   
		g_FiltersDefaultValues = "select '' statusname,'' statuscomment"
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "fiscal_year asc"
        
	    g_TableRowsUpdateAllowed = "NO" : g_TableRowsInsertAllowed  = "NO" 
		g_DBTableForInsertUpdate="dbo.purchasing_commodity"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate=""
		g_DBTableDropdownsForInsertUpdate = ""
		g_DBTableDatalistsForInsertUpdate = ""
		g_TableUpdateInsertLayoutVerticalHorizontal="H"
		
	case "2"
	
		g_Table_Caption_and_Info = "Yearly statistics by the departments"
		g_Form_Info_Help = ""
		
		g_SQL =         " select a.fiscal_year, a.purch_dept,b.deptname, a.po_count, a.encum_amount, '<a href=''sqlsite.asp?fiscal_year='+a.fiscal_year+'&purch_dept='+a.purch_dept+'&p=3''>...</a>' Info from ("
        g_SQL = g_SQL & " select [fiscal_year],[purch_dept], count([purchase_order]) po_count,sum([enc_amount]) [encum_amount] from ( "
        g_SQL = g_SQL & " select CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ) [fiscal_year],[purch_dept],[purchase_order],sum([encumbered_amount]) [enc_amount]   "
        g_SQL = g_SQL & " from [test_sqlweb_db].[dbo].[purchasing_commodity]  "
        g_SQL = g_SQL & " group by CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ),[purch_dept],[purchase_order] "
        g_SQL = g_SQL & " ) x group by [fiscal_year],[purch_dept] ) a inner join (select * from dbo.depts) b "
        g_SQL = g_SQL & " on a.purch_dept = b.deptcode "
        g_SQL = g_SQL & "  "
        		
		g_FilterDropdownsAllowed = "YES"
		g_FilterDropdownsColumns =  "select '%' as fiscal_year , 'All years' fiscal_year  union  select distinct CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ) fiscal_year,CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ) fiscal_year from [dbo].[purchasing_commodity];" _ 
								  & "select '%' as purch_dept , 'All depts' purch_dept  union  select distinct purch_dept,purch_dept + ' - ' + b.deptname from [dbo].[purchasing_commodity] a inner join (select * from dbo.depts) b on a.purch_dept = b.deptcode;"
			
		g_FilterDatalistsColumns = ""								
		g_FiltersDefaultValues = "select '' fiscal_year,'' purch_dept"
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "fiscal_year asc"
        
	    g_TableRowsUpdateAllowed = "NO" : g_TableRowsInsertAllowed  = "NO" 
		g_DBTableForInsertUpdate="dbo.purchasing_commodity"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate=""
		g_DBTableDropdownsForInsertUpdate = ""
		g_DBTableDatalistsForInsertUpdate = ""
		g_TableUpdateInsertLayoutVerticalHorizontal="V"
		
	CASE "3" 
		
		g_Table_Caption_and_Info = "Monthly statistics by the departments"
		g_Form_Info_Help = ""
		g_SQL =         " select a.fiscal_year,a.fiscal_month, a.purch_dept,b.deptname, a.po_count, a.encum_amount,'<a href=''sqlsite.asp?fiscal_year='+a.fiscal_year+'&purch_dept='+a.purch_dept+'&fiscal_month='+a.fiscal_month+'&p=4''>...</a>' purchase_lines from "
        g_SQL = g_SQL & " (select [fiscal_year],fiscal_month,[purch_dept], count([purchase_order]) po_count,sum([enc_amount]) [encum_amount] from ( "
        g_SQL = g_SQL & " select CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ) [fiscal_year], "
        g_SQL = g_SQL & " RIGHT('00' + CAST(DATEPART(mm, [post_date_orig]) AS varchar(2)), 2) fiscal_month, "
        g_SQL = g_SQL & " [purch_dept],[purchase_order],sum([encumbered_amount]) [enc_amount] from [test_sqlweb_db].[dbo].[purchasing_commodity]  "
        g_SQL = g_SQL & " group by CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ),RIGHT('00' + CAST(DATEPART(mm, [post_date_orig]) AS varchar(2)), 2), "
        g_SQL = g_SQL & " [purch_dept],[purchase_order] ) x group by [fiscal_year],fiscal_month,[purch_dept]  "
        g_SQL = g_SQL & " ) a inner join (select * from dbo.depts) b on a.purch_dept = b.deptcode "
        		
		g_FilterDropdownsAllowed = "YES"
		g_FilterDropdownsColumns =  "select '%' as fiscal_year , 'All years' fiscal_year  union    select distinct CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) )               fiscal_year, CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ) fiscal_year from [dbo].[purchasing_commodity];" _ 
								  & "select '%' as fiscal_month , 'All months' fiscal_month  union select distinct RIGHT('00' + CAST(DATEPART(mm, [post_date_orig]) AS varchar(2)), 2) fiscal_month,RIGHT('00' + CAST(DATEPART(mm, [post_date_orig]) AS varchar(2)), 2) fiscal_month from [dbo].[purchasing_commodity];" _
								  & "select '%' as purch_dept , 'All depts' purch_dept  union  select distinct purch_dept,purch_dept + ' - ' + b.deptname from [dbo].[purchasing_commodity] a inner join (select * from dbo.depts) b on a.purch_dept = b.deptcode;"
		g_FilterDatalistsColumns = ""	
		g_FiltersDefaultValues = "select '' fiscal_year,'' fiscal_month,'' purch_dept"
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "fiscal_year,fiscal_month,purch_dept asc"
        
	    g_TableRowsUpdateAllowed = "NO" : g_TableRowsInsertAllowed  = "NO" 
		g_DBTableForInsertUpdate="dbo.purchasing_commodity"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate=""
		g_DBTableDropdownsForInsertUpdate = ""
		g_DBTableDatalistsForInsertUpdate = ""
		g_TableUpdateInsertLayoutVerticalHorizontal="V"
	
	CASE "4" 
		
		g_Table_Caption_and_Info = "Purchase orders"
		g_Form_Info_Help = ""
		g_SQL =         " select CAST(DATEPART(yyyy, a.[post_date_orig]) AS varchar(4) ) [fiscal_year], "
        g_SQL = g_SQL & " RIGHT('00' + CAST(DATEPART(mm, a.[post_date_orig]) AS varchar(2)), 2) fiscal_month, a.[purch_dept],b.deptname, a.[purchase_order], "
        g_SQL = g_SQL & " [encumbered_amount],a.vendor_name,a.commodity_code, a.commodity_title from [test_sqlweb_db].[dbo].[purchasing_commodity] a  "
        g_SQL = g_SQL & " inner join (select * from dbo.depts) b on a.purch_dept = b.deptcode "
        		
		g_FilterDropdownsAllowed = "YES"
		g_FilterDropdownsColumns =  "select '%' as fiscal_year , 'All years' fiscal_year  union    select distinct CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) )               fiscal_year, CAST(DATEPART(yyyy, [post_date_orig]) AS varchar(4) ) fiscal_year from [dbo].[purchasing_commodity];" _ 
								  & "select '%' as fiscal_month , 'All months' fiscal_month  union select distinct RIGHT('00' + CAST(DATEPART(mm, [post_date_orig]) AS varchar(2)), 2) fiscal_month,RIGHT('00' + CAST(DATEPART(mm, [post_date_orig]) AS varchar(2)), 2) fiscal_month from [dbo].[purchasing_commodity];" _
								  & "select '%' as purch_dept , 'All depts' purch_dept  union  select distinct purch_dept,purch_dept + ' - ' + b.deptname from [dbo].[purchasing_commodity] a inner join (select * from dbo.depts) b on a.purch_dept = b.deptcode;" _
								  & "select '%' as commodity_code, 'All cc codes' commodity_title union select distinct commodity_code,commodity_title from [dbo].[purchasing_commodity];"
		g_FilterDatalistsColumns = "commodity_code;"	
		g_FiltersDefaultValues = "select '' fiscal_year,'' fiscal_month,'' purch_dept,'' commodity_code"
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "fiscal_year,fiscal_month,purch_dept asc"
        
	    g_TableRowsUpdateAllowed = "NO" : g_TableRowsInsertAllowed  = "NO" 
		g_DBTableForInsertUpdate="dbo.purchasing_commodity"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate=""
		g_DBTableDropdownsForInsertUpdate = ""
		g_DBTableDatalistsForInsertUpdate = ""
		g_TableUpdateInsertLayoutVerticalHorizontal="V"

	CASE "5" 
		
		g_page_datasource = "Chinook"
		g_PortalName = "Chinook"
		
		g_Table_Caption_and_Info = "Artist"
		g_Form_Info_Help = ""	
        g_SQL = " select artistid,name artist_name from artist "
        		
		g_FilterDropdownsAllowed = "NO"
		g_FilterDropdownsColumns = ""
		g_FilterDatalistsColumns = ""
		g_FiltersDefaultValues = ""
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "artist_name asc"
        
	    g_TableRowsUpdateAllowed = "YES" : g_TableRowsInsertAllowed  = "YES" 
		g_DBTableForInsertUpdate="artist"
		g_DBTableIdColumn="artistid"
		g_DBTableFieldsListForInsertUpdate="name"
		g_DBTableDropdownsForInsertUpdate = ""
		g_DBTableDatalistsForInsertUpdate = ""
		g_TableUpdateInsertLayoutVerticalHorizontal="V"

	CASE "6" 
		
		g_page_datasource = "Chinook"
		g_PortalName = "Chinook"
		
		g_Table_Caption_and_Info = " Albums " 
			
		g_Form_Info_Help = ""	
        g_SQL = " select albumid,a.artist_name,title album_title from album al inner join (select artistid,name artist_name from artist) a on al.artistid=a.artistid "
        		
		g_FilterDropdownsAllowed = "NO"
		g_FilterDropdownsColumns = ""
		g_FilterDatalistsColumns = ""
		g_FiltersDefaultValues = ""
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "album_title asc"
        
	    g_TableRowsUpdateAllowed = "YES" : g_TableRowsInsertAllowed  = "YES" 
		g_DBTableForInsertUpdate="album"
		g_DBTableIdColumn="albumid"
		g_DBTableFieldsListForInsertUpdate="title,artistid"
		g_DBTableDropdownsForInsertUpdate = "artistid;select null artistid,'' artist_name union select artistid,name artist_name from artist order by artist_name"
		g_DBTableDatalistsForInsertUpdate = ""
		g_DBTableMultipleDropdownsFieldsForInsert="artistid"
		g_TableUpdateInsertLayoutVerticalHorizontal="V"

    case "7"
	
		g_Table_Caption_and_Info = "Equipment Retirement Criterias"
		g_Form_Info_Help =                    "'One Rule' has 3 retirement criterias for any equipment: Running hours, Service life in month, Inspection status/color.<br>"
		g_Form_Info_Help = g_Form_Info_Help & "When any of the criteria is reached, the equipment will be written off/retired automatically.<br>"
		g_Form_Info_Help = g_Form_Info_Help & "We recommend to use rulename in format RHxxxxx-SMxx-RED, which mean RunningHoursxxxxx-ServiceMonthxx-Color."
		g_SQL =         " SELECT [Id],[rulename],[running_hours_greater_than],[service_life_inmonth_greater_than],[condition_color_is],[IsEnabled] FROM [dbo].[retirement_rules_criteria] "
        g_SQL = g_SQL & " "
        g_SQL = g_SQL & "  "
        g_SQL = g_SQL & "  "
        		
		g_FilterDropdownsAllowed = "NO"
		g_FilterDropdownsColumns = ""
		g_FiltersDefaultValues = ""
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "id asc"
        
	    g_TableRowsUpdateAllowed = "YES" : g_TableRowsInsertAllowed  = "YES" 
		g_DBTableForInsertUpdate="dbo.retirement_rules_criteria"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate="rulename,running_hours_greater_than,service_life_inmonth_greater_than,condition_color_is,IsEnabled"
		g_DBTableDropdownsForInsertUpdate = "IsEnabled;SELECT statusname IsEnabled,statuscomment FROM dbo.status;condition_color_is;SELECT [statusname] as condition_color_is,statuscomment FROM [dbo].[inspectionstatus]"
		g_TableUpdateInsertLayoutVerticalHorizontal="H"
	
	CASE "8"
	
		g_Table_Caption_and_Info = "Equipment Mooring Events"
		g_Form_Info_Help = ""	
		g_SQL =         " SELECT me.Id,e.vesselname,concat(e.eq_type,'/',e.winchuniquename,'/',e.vesselname,'/',e.serialnumber) Equipment,[portname],format([startmooring],'yyyy-MM-dd hh:m:ss') startmooring,format([endmooring],'yyyy-MM-dd hh:m:ss') endmooring,[mooringhours] FROM [WINCHES].[dbo].[mooring_events] me "
        g_SQL = g_SQL & " left join dbo.v_equipment e on e.id=me.equipment_id"
        g_SQL = g_SQL & "  "
        g_SQL = g_SQL & "  "
        		
		g_FilterDropdownsAllowed = "NO"
		g_FilterDropdownsColumns = ""
		g_FiltersDefaultValues = ""
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "e.id asc"
        
	    g_TableRowsUpdateAllowed = "YES" : g_TableRowsInsertAllowed  = "YES" 
		g_DBTableForInsertUpdate="dbo.mooring_events"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate="equipment_id,portname,startmooring,endmooring"
		g_DBTableDropdownsForInsertUpdate = "equipment_id;select id equipment_id,concat(eq_type,'/',winchuniquename,'/',vesselname,'/',serialnumber) Equipment,vesselname FROM dbo.v_equipment where IsEnabled='YES';portname;SELECT unlcode portname ,concat([unlcode],'/',[worldregion],'/',[country],'/',[province],'/',[city]) SeaPort  FROM [WINCHES].[dbo].[unlocode_ports] "
		g_DBTableDatalistsForInsertUpdate="portname"
		g_TableUpdateInsertLayoutVerticalHorizontal="V"
	
	case "9"
	
		g_Table_Caption_and_Info = "World Ports"
		g_Form_Info_Help = ""	
		g_SQL =         " SELECT [Id],[unlcode],[worldregion],[country],[province],[city],[coordinates] FROM [WINCHES].[dbo].[unlocode_ports] "
        g_SQL = g_SQL & " "
        g_SQL = g_SQL & "  "
        g_SQL = g_SQL & "  "
        		
		g_FilterDropdownsAllowed = "YES"
		g_FilterDropdownsColumns =  "select '%' as worldregion , 'All regions'    worldregion  union  select distinct worldregion, worldregion  from dbo.unlocode_ports;"  _
		                          & "select '%' as country     , 'All countries'  country      union  select distinct country,     country      from dbo.unlocode_ports;" _
								  & "select '%' as province    , 'All provinces'  province     union  select distinct province,    province     from dbo.unlocode_ports;" _
								  & "select '%' as city        , 'All cities'     city         union  select distinct city,        city         from dbo.unlocode_ports" 
								  
		g_FiltersDefaultValues = "select '' worldregion,'' country,'' province, '' city"
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "country asc"
        
	    g_TableRowsUpdateAllowed = "YES" : g_TableRowsInsertAllowed  = "YES" 
		g_DBTableForInsertUpdate="dbo.unlocode_ports"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate="unlcode,worldregion,country,province,city,coordinates"
		g_DBTableDropdownsForInsertUpdate = "worldregion;select distinct worldregion, worldregion  from dbo.unlocode_ports where worldregion<>'';country;select country,country countryname from (select distinct country  from dbo.unlocode_ports where country<>'') z order by country;province;select distinct province, province  from dbo.unlocode_ports where province<>'';city;select city,city cityname from (select distinct city  from dbo.unlocode_ports where city<>'') z order by city"
		g_DBTableDatalistsForInsertUpdate="province;city"
		g_TableUpdateInsertLayoutVerticalHorizontal="V"
	
	case "10"
			
		g_Table_Caption_and_Info = "Equipment Inspections"
		g_Form_Info_Help = ""	
		g_SQL =         " select ei.id, eq.equipment,ei.inspectioncomment,ei.inspectiondate,ei.inspectionstatus,ei.IsEnabled    from "
        g_SQL = g_SQL & " (SELECT Id,equipment_id,inspectiondate,inspectionstatus,inspectioncomment,IsEnabled FROM dbo.equipment_inspections) ei"
        g_SQL = g_SQL & " left join  (select id equipment_id,concat(eq_type,'/',winchuniquename,'/',vesselname,'/',serialnumber) Equipment FROM dbo.v_equipment) eq"
        g_SQL = g_SQL & " on ei.equipment_id = eq.equipment_id "
        		
		g_FilterDropdownsAllowed = "NO"
		g_FilterDropdownsColumns =  ""  _
		                          & ""
								  
		g_FiltersDefaultValues = ""
		
		g_TableColumnsSortingAllowed = "YES" 
		g_TableColumnsDefaultSorting = "id desc"
        
	    g_TableRowsUpdateAllowed = "YES" : g_TableRowsInsertAllowed  = "YES" 
		g_DBTableForInsertUpdate="dbo.equipment_inspections"
		g_DBTableIdColumn="id"
		g_DBTableFieldsListForInsertUpdate="inspectiondate,equipment_id,inspectionstatus,inspectioncomment,IsEnabled"
		g_DBTableDropdownsForInsertUpdate = "IsEnabled;SELECT statusname IsEnabled,statuscomment FROM dbo.status;equipment_id;select '' equipment_id,'' Equipment,'' vesselname union select id equipment_id,concat(eq_type,'/',winchuniquename,'/',vesselname,'/',serialnumber) Equipment,vesselname FROM dbo.v_equipment where IsEnabled='YES';inspectionstatus;SELECT statusname as inspectionstatus,statuscomment FROM dbo.inspectionstatus  where IsEnabled='YES'"
		g_DBTableDatalistsForInsertUpdate=""
		g_TableUpdateInsertLayoutVerticalHorizontal="V"
		
	case else
		response.redirect(page_name & "?p=" & g_DefaultPageCode) 
		
END SELECT


' / ---- Generate page block ------------------------------------------------------------------------------------------------------------------------------------------------------

	call debug_write ("Start page generator","")
	' /---  You can uncomment next line to write user activity log to database. May be used if you want to register all the details of visited page in database. But You need to investigate write_log function for more info)
	' write_log("Some message here")
	' \ --------

	' / --- Check for Vulnerable elements in url query string ----
	Dim var_QSvulnerabilities	
		var_QSvulnerabilities = func_CheckVulnerableElementsInQueryString()
		if var_QSvulnerabilities<>"" then 
			response.write var_QSvulnerabilities
			call debug_write ("Found Vulnerabilities in URL query string: " & var_QSvulnerabilities , "PRINT") ' If vulnerable elements in query string found, page will be terminated without rendering
		end if
	' \ ----------------------------------------------------------

	' / --- Menu block ---------------------
	dim var_MenuHTML
		var_MenuHTML = func_CreateMenuHTML()
		response.write var_MenuHTML
	' \ ------------------------------------

	' / --- Info/Help block -----------------
	dim var_InfoHelpHTML
		var_InfoHelpHTML = func_InfoHelpHTML()
		response.write var_InfoHelpHTML
	' \ -------------------------------------	

	response.write "<div class='main'>"					

	' Filter block		
	dim var_FiltersHTML
		var_FiltersHTML = func_CreateFiltersHTML(g_FilterDropdownsColumns)
		response.write var_FiltersHTML
		
	' Data table block	
	dim var_tableHTML
		var_tableHTML = func_CreateTableHTML() 
		response.write var_tableHTML 
					
	response.write "</div>"
					
	' Debug block
	call debug_write("End page generator","PRINT")

	response.end 
	
' \ ---- there the page generation ends -------------------------------------------------------------------------------------------------------------------------------------------

Sub debug_write (in_msg,in_termination_flag)
	if g_debug_flag="YES" then 
		g_debug_log=g_debug_log & in_msg & "<br>"
		if in_termination_flag="PRINT" then
			%>
			<div id="DebugInfo" style="display:none">
				<p style="margin-left: 200px; border:2px; border-style:solid; border-color:red; padding: 10px 10px 10px 10px;">
					Debug Info<br>		
					<%=g_debug_log%> 
				</p>
			</div>
			<%
			response.end
		end if
	end if
End Sub


Function func_InfoHelpHTML()

	' ### Message on page which will be opened after click on "Info/Help" menu item
	Dim show_news
		show_news = "YES"' / "NO"
		if show_news="YES" then	
			%>
			<div id="MsgInfo" style="display:none">
			<p style="margin-left: 200px; border:2px; border-style:solid; border-color:green; padding: 10px 10px 10px 10px;">
				<%=g_Table_Caption_and_Info%>
				<br>	
				<br><%=g_Form_Info_Help%>
			</p>
			</div>
			<%
			
		end if

End Function

Function func_CreateMenuHTML()
	
	if g_generate_menu_from_string="NO" then
	' #MANUAL_MENU block example
		%>
			<div class='sidenav'>
				<img src='your_image_url_here.jpg'>
				
				<a href="<%=page_name%>">Start Page</a>
				<button class='dropdown-btn'>Global Dictionaries<i class='fa fa-caret-down'></i></button>
					<div class='dropdown-container'>
						<a href='<%=page_name%>?p=1'>Record Status</a>
						<a href='<%=page_name%>?p=2'>Vessels</a>
						<a href='<%=page_name%>?p=7'>Ropes</a>
					</div>
				<button class='dropdown-btn'>Other Dictionaries<i class='fa fa-caret-down'></i></button>
					<div class='dropdown-container'>
						<a href='<%=page_name%>?p=3'>Operations</a>
						<a href='<%=page_name%>?p=4'>Locations</a>
						<a href='<%=page_name%>?p=5'>Damages</a>
						<a href='<%=page_name%>?p=6'>Inspection Results</a>
					</div>
				<button class='dropdown-btn'>Vessel Data<i class='fa fa-caret-down'></i></button>
					<div class='dropdown-container'>
						<a href='<%=page_name%>?p=8'>Vessel Ropes</a>
						<a href='<%=page_name%>?p=9'>Ropes Operations</a>
					</div>
				<%="<a id='downloadLink' onclick='exportToExcel(this)'>Export to excel</a>"%>
				<%="<a href='#' onclick='showNewsInfo()'>Info/Help</a>"%>			
			</div>
		<%
	' #MANUAL_MENU block end	
	else

		response.write func_GenerateMenu()

	end if

	' JS script below for menu items dropdowns
	%>
	<script>
				var dropdown = document.getElementsByClassName("dropdown-btn");
				var i;
				for (i = 0; i < dropdown.length; i++) {
				  dropdown[i].addEventListener("click", function() {
					this.classList.toggle("active");
					var dropdownContent = this.nextElementSibling;
					if (dropdownContent.style.display === "block") {
					  dropdownContent.style.display = "none";
					} else {
					  dropdownContent.style.display = "block";
					}
				  });
				}
			</script>
	<%
	
End Function


Function func_CreateTableHTML()  
	
	Dim page_ret_
	dim prc_ 
	' read previous stage parameters if they exists
	prc = CInt( NVL( request.querystring("prc") ,"1") )
	id_value = Request.QueryString("iv")
	'operation_type = Request.QueryString("op") ' i - insert (add new), e - edit row
	
	Select  Case Request.QueryString("op")
			Case "i"
				g_OperationTypeInsertUpdate = "INSERT"
			Case "e" 
				g_OperationTypeInsertUpdate = "UPDATE"
			Case else
				g_OperationTypeInsertUpdate = ""
	End Select

	action = Request.QueryString("a")          ' a - start action if it is initiated, mean button submit pressed after update or insert row
	call debug_write ("func_CreateTableHTML: id_value=" & id_value & " g_OperationTypeInsertUpdate=" & g_OperationTypeInsertUpdate & " action=" & action,"")
	
	if action="a" then

		if g_OperationTypeInsertUpdate="INSERT" and g_TableRowsInsertAllowed="YES" then 
			INSERT_SQL=func_CreateInsertUpdateStatementFromFormValues(g_DBTableForInsertUpdate,"") 
			page_ret_ = page_ret_ & execute_SCRIPT(INSERT_SQL)
		end if	
		if g_OperationTypeInsertUpdate="UPDATE" and g_TableRowsUpdateAllowed="YES" then 
			UPDATE_SQL=func_CreateInsertUpdateStatementFromFormValues(g_DBTableForInsertUpdate, func_CheckIfBracketsQuotesNeeded(g_DBTableIdColumn) & "=" & id_value) 
			page_ret_ = page_ret_ & execute_SCRIPT(UPDATE_SQL)
		end if	
		page_ret_ = page_ret_ &  "<a href='"& page_name &"?p=" & page & "&prc=" & prc & "'>Return to form</a><br>"
		func_CreateTableHTML = page_ret_: Exit Function
	
	else

		select case g_OperationTypeInsertUpdate
			   case "INSERT"
					if g_TableRowsInsertAllowed="YES" then
						page_ret_ = page_ret_ & add_rowRS("Add Record<br><br>"    & g_Table_Caption_and_Info,g_DBTableFieldsListForInsertUpdate,g_DBTableDropdownsForInsertUpdate)
					end if	

			   case "UPDATE"
					if g_TableRowsUpdateAllowed="YES" then
						page_ret_ = page_ret_ &  edit_rowRS("Edit Record<br><br>" & g_Table_Caption_and_Info,id_value,g_DBTableFieldsListForInsertUpdate,g_DBTableDropdownsForInsertUpdate)
					end if	
					
				case else
					page_ret_ = page_ret_ & get_htmlRS()
					
		end select

	end if
	func_CreateTableHTML = page_ret_  	
	
End function

Function func_CheckIfBracketsQuotesNeeded(in_field_for_sql)
	
	dim tmp_in_field_for_sql
	tmp_in_field_for_sql = trim(in_field_for_sql)
	if len(tmp_in_field_for_sql)>0 then
		
		if instr(tmp_in_field_for_sql," ")<>0 then
		
			if left(tmp_in_field_for_sql,1)<>g_columns_start_bracket then
				tmp_in_field_for_sql = g_columns_start_bracket & tmp_in_field_for_sql 
			end if
			
			if right(tmp_in_field_for_sql,1)<>g_columns_end_bracket then
				tmp_in_field_for_sql = tmp_in_field_for_sql & g_columns_end_bracket
			end if

		end if

	end if
	call debug_write ("func_CheckIfBracketsQuotesNeeded: in_value=" & in_field_for_sql & " out_value=" & tmp_in_field_for_sql,"")
	func_CheckIfBracketsQuotesNeeded = tmp_in_field_for_sql
	
End Function

Function func_AddFilterValuesToSQL(in_sql)

	Dim rs, cn
	dim ret_, res_

    Set rs = CreateObject("ADODB.Recordset")
	cn=Application(g_page_datasource)
	call debug_write ("func_AddFilterValuesToSQL: SQL: " & in_SQL,"")
	rs.open in_SQL, cn
        
	for i=0 to rs.fields.count-1
	    if len(request.querystring(rs.fields(i).name))<>0 then
			ret_=ret_ & func_CheckIfBracketsQuotesNeeded(rs.fields(i).name) & " like '" & request.querystring(rs.fields(i).name) & "' and "
		else
			if len(rs.fields(i).value)<>0 then
				ret_=ret_ & func_CheckIfBracketsQuotesNeeded(rs.fields(i).name) & " like '" & rs.fields(i).value & "' and "
			end if
		end if	
	next 
	if len(ret_)>0 and right(ret_,4)="and " then
	    ret_=" where " & mid(ret_,1,len(ret_)-4)
	end if
	
	rs.close
	set rs=nothing
	 
	func_AddFilterValuesToSQL = ret_
	
End Function

Function func_CreateFiltersHTML(g_FilterDropdownsColumns)

	if g_FilterDropdownsAllowed="YES" then
		if Request.QueryString("op")<>"i" and Request.QueryString("op")<>"e" and Request.QueryString("a")<>"a" then ' if mode edit or insert or button "submit data" pressed, we disable filter
		
			dim ret_
			dim dd  ' array of filter selects
			dim i

			dd=split(g_FilterDropdownsColumns,";")
			for i=0 to ubound(dd)
				if dd(i)<>"" then 
					ret_=ret_ & func_CreateFilterItemHTML(dd(i)) 
				end if	
			next 
			ret_="<form action='" & page_name & "' methos='post'><table class='tablefilter'><caption>Filters</caption>" & ret_ & "<tr><td><input type='hidden' name='p' value='" & get_page() & "'><input type='submit' value='Apply'></form></td><td></td></tr></table><br>"
		
		end if
	end if
	func_CreateFiltersHTML=ret_
	
End Function

Function func_CreateFilterItemHTML(in_SQL)

    ' incoming select needs to contain 2 columns
	' first column contain real value You need to have to filter 
	' second column contain visual good looking values for user-defined
    ' first column name will be used for filtering and applying to incoming select columns
	' but if there is no select statement we use one word field as search field 
	
	call debug_write ("func_CreateFilterItemHTML : in SQL =  " & in_SQL,"")
	'response.write in_sql
	'response.end
	if instr(in_SQL," ")<>0 then
		'if ucase(mid(in_SQL,1,6))="SELECT" then
			Dim rs, cn
			Dim ret_, res_

			Set rs = CreateObject("ADODB.Recordset")
				cn = Application(g_page_datasource)
			rs.open in_SQL, cn
				
				'#### GlobalVariablesFilter######
				dim filter_
				filter_ = func_GetGlobalFilter(rs)
				if filter_<>"" then rs.filter = filter_
			    '##########
			
			id_name    = rs.fields(0).name
			value_name = rs.fields(1).name
			
			' / - check if dropdown is <datalist or <select
			Dim dd_str, dd_type
			dd_str = split(g_FilterDatalistsColumns,";")
			for i=0 to ubound(dd_str)
				if ucase(dd_str(i))=ucase(id_name) then 
					dd_type = "DATALIST"
					exit for
				else
					dd_type = "SELECT"
				end if	
			next
			if dd_type = "" then dd_type="SELECT"
			' \ -
		   
			call debug_write ("func_CreateFilterItemHTML : dd_type =  " & dd_type,"")
			
			ret_ = "<tr><td>" & func_ReplaceTabColNameWithText(value_name) &  "</td><td>" 
			
			if dd_type = "SELECT" then
			
					ret_ = ret_ & "<select name='" & id_name & "'>" & vbcrlf
						
					do while not rs.eof
						if Request.QueryString(id_name)=rs.fields(0).value then 
							ret_ = ret_ & "<option value='" & rs.fields(0).value & "' selected>" & rs.fields(1).value & "</option>" & vbcrlf
						else
							ret_ = ret_ & "<option value='" & rs.fields(0).value & "'>" & rs.fields(1).value & "</option>" & vbcrlf
						end if
						rs.movenext
					loop
						
					rs.close
					set rs = nothing
					
					ret_=ret_ & "</select></td><tr>" & vbcrlf
			
			end if
			
			if dd_type = "DATALIST" then
			
					ret_ = ret_ & "<input list='" & id_name & "' name='" & id_name & "' type='text' value='" & Request.QueryString(id_name) & "'><datalist id='" & id_name & "'>" & vbcrlf
					
					do while not rs.eof
						ret_ = ret_ & "<option value='" & rs.fields(0).value & "'>" & rs.fields(1).value & "</option>" & vbcrlf
						rs.movenext
					loop
					rs.close
					set rs = nothing
					ret_=ret_ & "</datalist></td><tr>" & vbcrlf
			end if
			
		'end if
	else
		
		ret_ = ret_ & "<tr><td>" & func_ReplaceTabColNameWithText(in_SQL) &  "</td><td><input type='text' name='" & in_SQL & "' value='" & Request.QueryString(in_SQL) & "'>" & vbcrlf	
	end if
	
	func_CreateFilterItemHTML = ret_ 
	
End Function

Function func_GetGlobalFilter(in_rs)

		if g_GlobalVariablesValues<>"" then
				dim arr_ ,filter_ 
				dim field_name,field_value
				dim op, ft, i, ii
				
				arr_ = split(g_GlobalVariablesValues,";")
				for i=0 to in_rs.fields.count-1
					for ii=0 to ubound(arr_) step 2
						field_name  = lcase(arr_(ii))
						field_value =  session(arr_(ii))
						
						if lcase(in_rs.fields(i).name) = field_name  then
							
							if session(arr_(ii))<>"" then ' if session value is "" then ommit this in filter. 
								if rs_field_type(in_rs.fields(i).type)	=0 then ft="" else ft="'" 
									if field_value<>"%" then ' if used %, then we not limit data in filter

										if filter_="" then 
												filter_ = filter_ & lcase(arr_(ii)) & " = "  & ft & field_value & ft & " "
											else  
												filter_ = filter_ & " and " & lcase(arr_(ii)) & " = " & ft & field_value & ft & " "
										end if

									end if
								exit for
							end if
						end if	
					next	
				next
				if filter_<>"" then
					call debug_write ("func_GetGlobalFilter, Global filter applied: filter_ = " & filter_,"")
				end if	
			func_GetGlobalFilter = filter_
		end if
		
End Function

Function execute_SCRIPT(in_tsql)

    call debug_write ("Execute Script: " & in_tsql,"")
	on error resume next

	Dim cn,cns
	Dim msg_
    Set cn  = CreateObject("ADODB.Connection")
		cns = Application(g_page_datasource)
	cn.open cns
	cn.execute(in_tsql)
	cn.close
	set cn = nothing

	if err.number=0 then
		msg_ = "<br>Done without errors.<br> " '& in_tsql
	else
		msg_ = err.number & " " & err.description & " " '& in_tsql
	end if
	call debug_write(msg_,"")
	execute_SCRIPT = msg_
	
End Function

function func_CreateInsertUpdateStatementFromFormValues(in_table,where_statement)
    dim line_a
	dim line_b
	dim tmplt_a,tmplt_b, operations_count_, arr_
	dim vulnerability_result
	dim tmp_
	dim ret_
	dim multiple_
	
	'/ -------  Template creation  ------
    if g_OperationTypeInsertUpdate = "INSERT" then

			operations_count_="" ' minimal count of operations for insert is=1*1*1, but for multiple values  1*Xelements*Yelements 
			For i=1 to Request.Form.count
				Item=Request.Form.Key(i)
				val_=Request.Form.Item(i)
				
				x_="#" & i-1 & "#" ' count of elements in array for posted value in <select multiple
				
				'--/ ' check posted field name with a variable which contain list of all multiple  variables --
				if instr(g_DBTableMultipleDropdownsFieldsForInsert&",",mid(Item,1,len(Item)-1))<>0 then 
					multiple_=1 
					arr_=split(val_,", ")
				else 
					multiple_=0 
					arr_=split(val_,chr(0)) ' if posted values is not an array and is 1 element we use fake delimiter for multile items to always get array with 1 element and index=0 
				end if	
				'--\-------------------------------------------------------------------------------------------
				
				if ubound(arr_)<>-1 then 
					operations_count_=operations_count_ & ubound(arr_) & ","
					values_ = values_ & val_ & "|"
					
					fieldName = mid(Item,1,len(Item)-1)
					fieldType = mid(Item,len(Item),1)
					fieldValue = x_
					if fieldType="0" then fT="": if fieldValue="" then fieldValue="0" ' default value for numeric values
					if fieldType="1" then fT="'"
					if fieldType="2" then fT="": fieldValue=replace(g_DateFromTextToSQL,"#DATE#",fieldValue)    ' added for HTML5 universal date string value YYYY-MM-DD
					if fieldType="3" then fT="": fieldValue=replace(g_DateTimeFromTextToSQL,"#DATE#",fieldValue)' added for HTML5 universal date string value YYYY-MM-DDTHH:MM
									 
					tmplt_a = tmplt_a & func_CheckIfBracketsQuotesNeeded(fieldName) & ","
					tmplt_b = tmplt_b & fT & replace(fieldValue,fT,fT&fT) & fT & "," ' fieldValue -> replace(fieldValue,fT,fT&fT)  quote replaced with doublequotes
				
				else
					
					' if ubound(arr_) =-1 , this mean we have field but posted value in form was not filled and formaly it is equal to null, and we can ommit this value in insert statement
					'call debug_write ("While creating insert statement we ommit parameter " & Item & " because it value was NULL or not filled. It was by count in insert statement " & i & " element.","")
					operations_count_=operations_count_ & "0,"
					values_ = values_ & "NULL" & "|"
					fieldName = mid(Item,1,len(Item)-1)
					tmplt_a = tmplt_a & func_CheckIfBracketsQuotesNeeded(fieldName) & ","
					tmplt_b = tmplt_b & "NULL," ' fieldValue -> replace(fieldValue,fT,fT&fT)  quote replaced with doublequotes
				
				end if
				
			Next 
			if tmplt_a<>"" then tmplt_a = mid(tmplt_a,1,len(tmplt_a)-1)
			if tmplt_b<>"" then tmplt_b = mid(tmplt_b,1,len(tmplt_b)-1)
			
			if len(operations_count_)>0 then operations_count_=mid(operations_count_,1,len(operations_count_)-1): values_=mid(values_,1,len(values_)-1)
			call debug_write ("Template prepared for " & operations_count_ & " =operations in it: " & "Insert into " & in_table & " (" & tmplt_a & ") values (" & tmplt_b & ")" ,"")
			arr_=split(operations_count_,",")
			arr_val_ = split(values_,"|")
			
			call debug_write ( "{" & operations_count_ & "} {" & values_ & "} {" & tmplt_b & "}" ,"") ' : response.end
			
			res_ = tmplt_b' sozdajom 1 row i dalee ego budem zapolnjatj
			for i=0 to ubound(arr_)
				op_ = arr_(i)' skoljko budet operacij zamen, libo 1 libo bolshe
				
				'opv_=split(arr_val_(i),", ")' spisok znachenij v dannoj operacii, 1 libo bolshe
				
				if op_>0 then
					call debug_write (">0","")
					resl_=""
					opv_=split(arr_val_(i),", ")' spisok znachenij v dannoj operacii, 1 libo bolshe
					for ii=0 to op_
						res1_ = res1_ & replace(res_,"#" & i & "#",opv_(ii)) & vbcrlf '"<br>" 
					next
					res_=res1_ 
				end if
				
				if op_=0 then
					call debug_write ("=0","")	
					opv_=split(arr_val_(i),chr(0))' spisok znachenij v dannoj operacii, vsegda 1 
					res_ = replace(res_,"#" & i & "#",opv_(0)&"") & vbcrlf '"<br>"
				end if

				'if op_<0 then
				'	call debug_write ("=-1","")
				'	res_ = replace(res_,"#" & i & "#","NULL") & vbcrlf '"<br>"
				'end if
				call debug_write ("Counter i=" & i & " op_=" & op_ & " ubound of opv_=" & ubound(opv_) & " value res_=" & res_ ,"")
				out_=out_ & res_
			next
			
			' output array normalisation. deleting null lines and lines containing '#" signs which is abnormal 
			res_=""
			arr_=split(out_,vbcrlf)
			for i=0 to ubound(arr_)
				if instr(arr_(i),"#")<>0 or len(arr_(i))=0 then 
					'res_=res_ 
				else 
					res_=res_ & arr_(i) & vbcrlf ' Potential BUG if multiple values dropdown item contain ', ' or '#' signs
					ret_ = ret_ & "Insert into " & in_table & " (" & tmplt_a & ") values (" & arr_(i) & ");" & vbcrlf
					
				end if	
			next
			
	end if
	
	'\ ----------------------------------	
	
	if g_OperationTypeInsertUpdate = "UPDATE" then
			
			For i=1 to Request.Form.count
				Item=Request.Form.Key(i)
				val_=Request.Form.Item(i)
				'if instr(val_,", ")=0 then 
					call debug_write ("Posted Form field name={" & Item & "} value={" & val_ & "}" ,"")
				'else
				'	call debug_write ("Posted Form field name={" & Item & "} values={" & val_ & "} MULTIPLE POSTED VALUES DETECTED" ,"")
				'	if g_OperationTypeInsertUpdate = "UPDATE" then 
				'		call debug_write ("ERROR: For Update it's impossible to convert ONE row to MANY, because of that update operation terminated. Or other reason: value from posted form containing ', ' and 'this is detected like delimiter in multiple posted values which is impossible for update, recreated list of values in form to exclude this element ', ' in key values. ","PRINT")
				'	end if	
				'end if
				' vulnerability check
				tmp_ = func_VulnerableElementsCheck(val_)
				if tmp_<>"" then
					vulnerability_result = vulnerability_result & "HTML Form Key: <font color='green'>" & Item & "</font><br> Form Value: <font color='blue'>" & val_ & "</font><br> Vulnerable element: <br><font color='red'>" & tmp_ & "</font><br>"
				end if
				
				fieldName = mid(Item,1,len(Item)-1)
				fieldType = mid(Item,len(Item),1)
				fieldValue = val_
				if fieldType="0" then fT="": if fieldValue="" then fieldValue="0" ' default value for numeric values
				if fieldType="1" then fT="'"
				if fieldType="2" then fT="": if fieldValue<>"" then fieldValue=replace(g_DateFromTextToSQL,"#DATE#",fieldValue) else fieldValue="NULL"    ' added for HTML5 universal date string value YYYY-MM-DD
				if fieldType="3" then fT="": if fieldValue<>"" then fieldValue=replace(g_DateTimeFromTextToSQL,"#DATE#",fieldValue) else  fieldValue="NULL"' added for HTML5 universal date string value YYYY-MM-DDTHH:MM
				
				line_a = line_a & func_CheckIfBracketsQuotesNeeded(fieldName) & "=" & fT &  replace(fieldValue,fT,fT&fT) & fT & "," ' fieldValue -> replace(fieldValue,fT,fT&fT)  quote replced with doublequotes
				
			Next 
			
			if line_a<>"" then line_a = mid(line_a,1,len(line_a)-1)
			if line_b<>"" then line_b = mid(line_b,1,len(line_b)-1)

			if vulnerability_result<>"" then 
				call debug_write ("Table: " & in_table & "<br> WHERE Statement: " & where_statement & "<br> Reading type: " & g_OperationTypeInsertUpdate & "<br> Vulnerability list: <br>" & vulnerability_result,"")
				ret_="/**/"
			else
				ret_ = "Update " & in_table & " set " & line_a & " where " & where_statement & ";" & vbcrlf
			end if
			
	end if
	
	call debug_write ("func_CreateInsertUpdateStatementFromFormValues: SQL construction: " & replace(ret_,vbcrlf,"<br>"),"")
	func_CreateInsertUpdateStatementFromFormValues = ret_
	
end function


Function get_htmlRS()

'on error resume next ' '#####

	Dim rs, rc_null,cn
	dim res_
	dim table
	dim where_sql_
	dim i
	dim sSQL
	dim subtotals			 
	dim subtotals_values	
	dim subtotals_formula	
    
	dim prc 
	
	sSQL=g_SQL
	if g_FilterDropdownsAllowed="YES" then
		where_sql_ = func_AddFilterValuesToSQL(g_FiltersDefaultValues)
		if where_sql_ <> "" then
			sSQL="select * from (" & sSQL & ") x " & where_sql_ & " "
		end if
	end if
	
	if g_TableColumnsSortingAllowed="YES" then
		if request.querystring("s")<>"" then
            qs_s = func_MinimizeSortingQS(request.querystring("s"))		
			sSQL=sSQL & " order by " & qs_s '  #sorting / and \# ' this line is for tables content sorting by column
		else 
		    if g_TableColumnsDefaultSorting<>"" then sSQL=sSQL & " order by " & g_TableColumnsDefaultSorting ' if default sorting columns not null then apply them to view
		end if	
	end if
	
	if g_TableRowsInsertAllowed="YES" then
		table= table & "<a href='" & page_name & "?op=i&p=" & page & "' title='New record can be added if ID column generation on database level realyzed automatically or by trigger'>Add New Record</a><br><br>" 
	end if
	
	call debug_write("DataTable SQL: " & sSQL,"")
	
    Set rs = CreateObject("ADODB.Recordset")
	cn=Application(g_page_datasource)
	
	rs.PageSize = g_page_records_count ' kolichestvo zapisej na stranice VARIABLE IN (https://www.sitepoint.com/community/t/asp-recordset-paging/1026/3)
	rs_page_size = rs.PageSize
	rs.CacheSize = rs.PageSize
	RS.CursorLocation = 3  ' 3-adUseClient
	
	rs.open sSQL, cn , 0, 1, 1 'adOpenForwardOnly, adLockReadOnly, adCmdText ' https://www.w3schools.com/asp/met_rs_open.asp
	
	'if err.number<>0 then 
	'	response.write ssql
	'	response.end
	'end if
	'rs.open sSQL, cn , 3, 3, &H1 'adOpenStatic, adLockOptimistic, adCmdText ' https://www.w3schools.com/asp/met_rs_open.asp
    
			'#### GlobalVariablesFilter######
				dim filter_
				filter_ = func_GetGlobalFilter(rs)
				if filter_<>"" then rs.filter = filter_
			'##########
	if rs.eof or rs.bof then
		get_htmlRS=replace(table,"#PAGINATION_BLOCK#","") & "<br>No Records Found. Table has no records or filtering parameters didn't return results.<br>"
		exit function
	end if	
	
	prc = CInt( NVL( request.querystring("prc") ,"1") ) ' read current pagination from query string and if not found set to 1.
	rs.AbsolutePage = prc ' Current page of recordset , if not defined, then used value = 1
	
	CountOfPagesInRS = rs.PageCount
	min_page = prc - 3
	max_page = prc + 3
	if min_page<1 then min_page=1
	if max_page>CountOfPagesInRS then max_page=CountOfPagesInRS
	
	current_page = func_ModifyQS(Request.Servervariables("QUERY_STRING"),"prc")
	
	call debug_write("Page by RS : " & prc & " current_page=" & current_page ,"")
	
	for pp=min_page to max_page
			links_ = links_ & "<a href='" & request.Servervariables("SCRIPT_NAME") & "?" & current_page & "&prc=" & pp & "'>" & pp & "</a>&nbsp;"
	next 
	
	' pagination block on top of filter and table
	table = table & "<table class='DataTable' id='DataTable'><caption>" & g_Table_Caption_and_Info & "</caption><thead>" ' THEAD###
	
	if g_TableRowsUpdateAllowed="YES" then
		table=table & "<tr><th>Edit</th>"  
		else
		table=table & "<tr>"  
	end if	
	
	redim subtotals(rs.fields.count-1)				 
	redim subtotals_values(rs.fields.count-1)		
	redim subtotals_formula(rs.fields.count-1)      
	dim rec_on_page 
	
	dim qs_
	qs_ = func_ReduceSortingParametersInQS(QS())
	ID_=func_CheckIfBracketsQuotesNeeded(ucase(g_DBTableIdColumn))
	for i=0 to rs.fields.count-1
	  
	  subtotals(i)= rs_field_type(rs.fields(i).type) 
	  subtotals_values(i)=0.01-0.01
	  
	  if g_TableColumnsSortingAllowed="YES" then
		   table = table & "<th align='center'>" & "<a href='" & page_name & "?s=" & rs.fields(i).name & "&" & qs_ & "'>" & func_ReplaceTabColNameWithText(rs.fields(i).name) & "</a></th>" 
	  else
	  		table = table & "<th align='center'>" & func_ReplaceTabColNameWithText(rs.fields(i).name) & "</th>" 
	  end if
	  
	  if func_CheckIfBracketsQuotesNeeded(ucase(rs.fields(i).name))=ID_ then id_column=i 
	next 
	table = table & "</tr></thead><tbody>" & vbcrlf
	
	'############
	if g_TableRowsUpdateAllowed="YES" or g_TableRowsInsertAllowed="YES" then
		if id_column="" then 
			get_htmlRS="ID column not found but is necessary for records edition. Check Report settings syntax." 
			Exit Function
		end if	
	end if
	
	do while not rs.eof
	    rec_on_page=rec_on_page+1  ' added pagination
		if rec_on_page>rs_page_size then exit do ' added pagination
		
		if g_TableRowsUpdateAllowed="YES" then
		
			if rs_field_type(rs.fields(id_column).type)	= 0 then
				ft=""
			else
				ft="'"
			end if
			
			table=table & "<tr><td><a href='" & page_name & "?iv=" & ft & rs.fields(id_column).value & ft & "&op=e&p=" & page & "&prc=" & prc & "'>...</td>" '###VKU###
		else
			table=table & "<tr>" 
		end if
		for i=0 to rs.fields.count-1
				res_ = rs.fields(i).value
				ann = res_
				
				if len(res_)>0 then
					res_ = replace(res_,",","")
					res_ = replace(res_,"<font color='red'><b>","")
					res_ = replace(res_,"</b></font>","")
				end if
				
				if isnumeric(res_)=true then
					table=table & "<td align='right'>" & ann & "</td>" 
					subtotals(i)=subtotals(i)+1
					subtotals_values(i) = subtotals_values(i) + res_	  ' description of the bug 800a000d https://forums.adobe.com/thread/156021 - sql server datatype decimal can't be identified by vbscript
				    subtotals_formula(i) = subtotals_formula(i) & "+" & res_
				else
					' html 5 fields
					ft_inputtype = rs_field_db_type(rs.fields(i).type,rs.fields(i).name)
					if ft_inputtype = "date" then ' convert all the dates to one universal format YYYY/MM/DD 
						html_cell="" 
						html_cell=year(rs.fields(i).value) & "/"
						if month(rs.fields(i).value)<10 then html_cell = html_cell & "0" & month(rs.fields(i).value) & "/" else html_cell = html_cell & month(rs.fields(i).value) & "/"
						if day(rs.fields(i).value)<10 then html_cell = html_cell & "0" & day(rs.fields(i).value) else html_cell = html_cell & day(rs.fields(i).value)
						
					else
						html_cell = ann
					end if
					' datatable cells values
					table=table & "<td>" & html_cell & "</td>" 
					
				end if	
			next 
		table=table & "</tr>" & vbcrlf
		
		rs.movenext
	loop
	if g_TableRowsUpdateAllowed="YES"  then
			table = table & "<tr><td></td>"
		else
			table = table & "<tr>"		
	end if
	
	if g_ApplySubtotalsForNumericColumns="YES" then 	
		for i=0 to rs.fields.count-1
			if subtotals(i)<>0 and subtotals_values(i)<>0 and subtotals_values(i)<>"" then  
				if len(subtotals_formula(i))>1 and mid(subtotals_formula(i),1,1)="+" then subtotals_formula(i)=mid( subtotals_formula(i), 2 , len(subtotals_formula(i))-1 )
				table=table & "<td align='right' title='" & subtotals_formula(i) & "'><font color='#2471A3'><b>" & formatnumber(subtotals_values(i),2,0,0,0) & "</b></font></td>" 
			else 
				table=table & "<td></td>"
			end if	
		next						
	end if
	
	table = table & "</tr>"		
	
    table=table & "</tbody></table>"

	rs.close
	set rs=nothing
    
	' pagination block in below of filter and table
	table =table & "<br>Page: " & prc & "/" & CountOfPagesInRS & ": Select Page: " & links_ & "<br>"
	
	get_htmlRS = table 
	
End Function

Function func_ModifyQS(in_value , in_tag)

	'Function to remove "in_tag" and its value from query string - used for links generation

	dim val_
	dim ret_
	
	val_ = in_value
	
	if val_<>"" then
			dim tmp_
			dim out_tags
			tmp_ = split(val_,"&")
			for z=0 to ubound(tmp_)
				if mid( tmp_(z), 1 , len(in_tag) ) = in_tag then
					' do nothing
					else
					out_tags=out_tags & tmp_(z) & "&"
				end if
			next 
			
			if len(out_tags)>0 then 
				out_tags = mid(out_tags,1,len(out_tags)-1)
			end if
			
			ret_=out_tags
		else
			ret_=in_value
	end if
    
	'call debug_write(" In value: " & in_value & " in tag: " & in_tag & "<br>Out value: " & ret_,"")
	
	func_ModifyQS = ret_

End Function

Function NVL(in_value,null_replacement)

	if in_value & "" = "" then 
			NVL = null_replacement
		else
			NVL = in_value
	end if

End Function

Function func_ReduceSortingParametersInQS(in_qs)

	dim arr_
    dim tmp_qs
	dim out_qs
	
	arr_=split(in_qs,"&")
    tmp_qs = in_qs
	
	for i=0 to ubound(arr_)
        if arr_(i)<>"" then
			z=( len(tmp_qs) - len(replace(tmp_qs,arr_(i),"")) ) / len(arr_(i)) 		
			tmp_qs = replace(tmp_qs,arr_(i),"")	
            'response.write "i=" & i & "z=" & z & " " & arr_(i) & " " & tmp_qs & "<br>" 			
			if z<>0 then
				if z\2=z/2  then 
						out_qs = out_qs & arr_(i) & "&" & arr_(i) & "&" 
				else 
						out_qs = out_qs & arr_(i) & "&"
				end if	
			end if
		end if	
	next
    
	if len(out_qs)>0 then out_qs=mid(out_qs,1,len(out_qs)-1)
	func_ReduceSortingParametersInQS = out_qs
	
End Function

Function func_MinimizeSortingQS(in_qss)
	' Sorting parameters minimisation in query string (ASC and DESC policy)
	Dim arr_
	Dim ret_
	dim increment
	dim inc_counter

	arr_=split(in_qss,",")
	inc_counter = ""

	if ubound(arr_) > 0 then

		for i=0 to ubound(arr_)-1
			if arr_(i)<>"" then 
				increment=1 
				for x=i+1 to ubound(arr_)
					if trim(arr_(i))=trim(arr_(x)) then arr_(x)="":increment=increment+1
				next	
				inc_counter=inc_counter & increment & ","
			else
				inc_counter = inc_counter & "0,"
			end if	
		next

		dim count_
		count_=split(inc_counter,",")

		for i=0 to ubound(arr_)
			if arr_(i)<>"" then
				t=count_(i) ' count of sorting repeats
				if t="" then t=0
				if t<>0 then
					if  t/2=t\2 then word_=" DESC" else word_=" ASC"
					out_ = out_ & g_columns_start_bracket & trim(arr_(i)) & g_columns_end_bracket & word_ & ","
				end if
			end if
		next

		if out_<>"" then out_=mid(out_,1,len(out_)-1)
	else
		if arr_(0)<>"" then out_=g_columns_start_bracket & trim(arr_(0)) & g_columns_end_bracket & " ASC"
	end if
	
	func_MinimizeSortingQS=out_

End Function

Function add_rowRS(g_Table_Caption_and_Info,editable_cols,g_DBTableDropdownsForInsertUpdate)

	Dim rs, rc_null,cn
	dim res_, new_row
	dim table
	dim i
	dim rs_sql 
	dim f_arr_string
	dim f_arr_values
	dim prc
	
	ec= ucase(g_DBTableIdColumn & "," & editable_cols) 	            ' ######### g_DBTableIdColumn & editable columns
	rs_sql = "select " & ec & " from " & g_DBTableForInsertUpdate   ' ######### "select * from " & g_DBTableForInsertUpdate
	
    Set rs = CreateObject("ADODB.Recordset")
	cn=Application(g_page_datasource)
	rs.open rs_sql, cn
    
	ID_= func_CheckIfBracketsQuotesNeeded(ucase(g_DBTableIdColumn))
	
	for i=0 to rs.fields.count-1
	  ft = rs_field_type(rs.fields(i).type) 
	  ft_inputtype = rs_field_db_type(rs.fields(i).type,rs.fields(i).name)
	  arr_One_Value = func_ReplaceTabColNameWithText(rs.fields(i).name)
	  dd_ifexist = func_GetFilterDropdownsIfExist(g_DBTableDropdownsForInsertUpdate,rs.fields(i).name,"")
	  
		if func_CheckIfBracketsQuotesNeeded(ucase(rs.fields(i).name)) <> ID_ then 
			if dd_ifexist<>"" then
				arr_Two_Value = dd_ifexist
			else
				' html5 fields type support by browsers
				if g_use_html5_fields_for_input="YES" then 
					ft_inputtype=ft_inputtype
				else
					ft_inputtype="text"
				end if	
				arr_Two_Value = "<input type='" & ft_inputtype & "' name='" & rs.fields(i).name & ft & "' value=''>	"
			end if
		else 
			arr_Two_Value = "" 
		end if
		f_arr_string = f_arr_string & arr_One_Value & chr(0) & arr_Two_Value & chr(0)
	next 

	' Transform array values to table layout horizontal or vertical
			
	if len(f_arr_string)>0 then
		f_arr_values = split(f_arr_string,chr(0))
		
		if g_TableUpdateInsertLayoutVerticalHorizontal="V" then
				' vertical layout
				new_row = "<tr><th>Name</th><th>Value<th></tr>"  
				for i=0 to ubound(f_arr_values)-1 step 2
					new_row = new_row & "<tr><td>" & f_arr_values(i) & "</td><td>"	& f_arr_values(i+1) & "</td></tr>"
				next
		end if
		
		if g_TableUpdateInsertLayoutVerticalHorizontal="H" or g_TableUpdateInsertLayoutVerticalHorizontal=""  then
				' horizontal layout
				new_row = "<tr>"
				for i=0 to ubound(f_arr_values)-1 step 2
					new_row = new_row & "<th>" & f_arr_values(i) & "</th>"
				next
				new_row = new_row & "</tr><tr>"
				for i=0 to ubound(f_arr_values)-1 step 2
					new_row = new_row & "<td>" & f_arr_values(i+1) & "</td>"
				next
				new_row = new_row & "</tr>"
		end if
	end if
	prc = CInt( NVL( request.querystring("prc") ,"1") )'###VKU###
	table= "<br><form action='" & page_name & "?op=i&a=a&p=" & page & "&prc=" & prc & "' method='post'><table class='DataTable'><caption>" & g_Table_Caption_and_Info & "</caption>" & new_row & "</table><br>"
	table = table & "<input type='submit' value='Create Record'></form>"
	
	rs.close
	set rs=nothing
    
	add_rowRS = table ' return result
	
End Function


Function edit_rowRS(g_Table_Caption_and_Info,id_value,editable_cols,g_DBTableDropdownsForInsertUpdate)

	Dim rs, cn
	dim new_row
	dim table
	dim i
	dim ec
	dim rs_sql 
	dim f_arr_string
	dim f_arr_values
	dim prc 
	
	rs_sql = "select * from " & g_DBTableForInsertUpdate 
	ec="," & replace(replace(ucase(editable_cols),g_columns_start_bracket,""),g_columns_end_bracket,"") & ","
	rs_sql = rs_sql & " where " & func_CheckIfBracketsQuotesNeeded(g_DBTableIdColumn) & "=" & id_value
	
	if id_value ="" or instr(id_value,",")<>0 then 
		call debug_write("Abnormal Table ID for edition received : NULL or multiple values from query string. Vulnerable action from user.", "")
		exit function  ' prevent multiple id edition. iv=1&iv=2 not possible. 
	end if		
	
    Set rs = CreateObject("ADODB.Recordset")
	cn=Application(g_page_datasource)
	rs.open rs_sql, cn
	
	ID_=func_CheckIfBracketsQuotesNeeded(ucase(g_DBTableIdColumn))
	
	do while not rs.eof
		for i=0 to rs.fields.count-1
		    ft = rs_field_type(rs.fields(i).type)
			ft_inputtype = rs_field_db_type(rs.fields(i).type,rs.fields(i).name)
			
			select case ft_inputtype
			
				case  "date" 
					html_cell = func_DateTimeFormat("yyyy-mm-dd",rs.fields(i).value)
				case "datetime-local"
					html_cell = func_DateTimeFormat("yyyy-mm-ddThh:mi",rs.fields(i).value) '#### html5 datetime-local field has format YYYY-MM-DDTHH:MI
				case "time"
					html_cell = func_DateTimeFormat("hh:mi:ss",rs.fields(i).value)
				case else
					html_cell = rs.fields(i).value
			end select
			
			arr_One_Value = func_ReplaceTabColNameWithText(rs.fields(i).name) 
			
			dd_ifexist = func_GetFilterDropdownsIfExist(g_DBTableDropdownsForInsertUpdate,rs.fields(i).name,rs.fields(i).value)
			if func_CheckIfBracketsQuotesNeeded(ucase(rs.fields(i).name)) <> ID_ then 
				if instr(ec,"," & ucase(rs.fields(i).name) & ",")<>0 then 
					if dd_ifexist<>"" then
						arr_Two_Value= dd_ifexist 
					else
						' html5 fields type support by browsers
						if g_use_html5_fields_for_input="YES" then 
							ft_inputtype=ft_inputtype
						else
							ft_inputtype="text"
						end if	
						arr_Two_Value = "<input type='" & ft_inputtype & "' name='" & rs.fields(i).name & ft & "' value='" & html_cell & "' size='" & len(html_cell)+4 & "'>"  ' not tested on null values (len field if null)
					end if	
				else
					arr_Two_Value = html_cell 
				end if	
				
			else 
				arr_Two_Value = "" ' ignore value of ID
			end if
			f_arr_string = f_arr_string & arr_One_Value & chr(0) & arr_Two_Value & chr(0)
		next
		rs.movenext	
		'table = table & "</tr>"
		
	loop
	
	' Transform array values to table layout horizontal or vertical

	if len(f_arr_string)>0 then
		f_arr_values = split(f_arr_string,chr(0))
		
		if g_TableUpdateInsertLayoutVerticalHorizontal="V" then
				' vertical layout
				new_row = "<tr><th>Name</th><th>Value<th></tr>"  
				for i=0 to ubound(f_arr_values)-1 step 2
					new_row = new_row & "<tr><td>" & f_arr_values(i) & "</td><td>"	& f_arr_values(i+1) & "</td></tr>"
				next
		end if
		
		if g_TableUpdateInsertLayoutVerticalHorizontal="H" or g_TableUpdateInsertLayoutVerticalHorizontal=""  then
				' horizontal layout
				new_row = "<tr>"
				for i=0 to ubound(f_arr_values)-1 step 2
					new_row = new_row & "<th>" & f_arr_values(i) & "</th>"
				next
				new_row = new_row & "</tr><tr>"
				for i=0 to ubound(f_arr_values)-1 step 2
					new_row = new_row & "<td>" & f_arr_values(i+1) & "</td>"
				next
				new_row = new_row & "</tr>"
		end if
	end if

	prc = CInt( NVL( request.querystring("prc") ,"1") )'###VKU###
	table="<br><form action='" & page_name & "?op=e&a=a&iv=" & id_value & "&p=" & page & "&prc=" & prc & "' method='post'><table class='DataTable'><caption>" & g_Table_Caption_and_Info & "</caption>" & new_row & "</table><br>"
	table = table & "<input type='submit' value='Apply Changes'></form>"

	rs.close
	set rs=nothing
    
	edit_rowRS = table ' return result
	
End Function

Function func_ObjectIsPartOfList(in_array,in_field,in_delimiter)
	
	'g_DBTableMultipleDropdownsFieldsForInsert
	
	dim arr_,i
	dim out_
	out_ = "NO"
	arr_ = split(in_array,in_delimiter)
	for i = 0 to ubound(arr_)
		if ucase(arr_(i))=ucase(in_field) then 
			out_="YES"
			exit for
		end if	
	next
	func_ObjectIsPartOfList = out_
	
End Function																
			
Function func_GetFilterDropdownsIfExist(g_DBTableDropdownsForInsertUpdate,in_field,in_value)

    ' g_DBTableDropdownsForInsertUpdate field; select fieldid,description,globalfilterfield1,globalfilterfield2 from table 
	
	' first column contain real value You need to have to filter : first column name will be used for filtering and applying to incoming select columns
	' second column contain visual good looking values for user-defined
	' columns starting from 3, may be ommited or you can use them for values used in GlobalVariablesFilter
 
on error resume next
	
	Dim rs1, cn1, ret_, rec_count, multiple_, records_in_loop
	dim dd_fld,dd_sql, id_name
	dim dd_str, dd_found, dd_type
	dd_str = split(g_DBTableDropdownsForInsertUpdate,";")
	
	for i=0 to ubound(dd_str) step 2
		dd_fld = dd_str(i+0)' filter field name which will be linked to dropdownlist
		dd_sql = dd_str(i+1)' filter sql query :field name +field description + GlobalVariablesFilter If Needed(but may be ommited)
		if ucase(dd_fld)=ucase(in_field) then
			dd_found = dd_sql
			exit for
		end if
	next
	
	dd_str = split(g_DBTableDatalistsForInsertUpdate,";")
	for i=0 to ubound(dd_str)
		if ucase(dd_str(i))=ucase(in_field) then 
			dd_type = "DATALIST"
			exit for
		else
			dd_type = "SELECT"
		end if	
	next
	if dd_type="" then dd_type = "SELECT"						  
	
	if dd_found="" then exit function
	
    Set rs1 = CreateObject("ADODB.Recordset")
	cn1=Application(g_page_datasource)

	rs1.open dd_found, cn1
	
		'#### GlobalVariablesFilter######
		dim filter_
		filter_ = func_GetGlobalFilter(rs1)
		if filter_<>"" then rs1.filter = filter_
		'##########
				
    if err.number<>0 then 
		call debug_write ("Error occured in func_GetFilterDropdownsIfExist:" & dd_found , "")
		exit function    
	end if	
	
	
	id_name=rs1.fields(0).name
	ft = rs_field_type(rs1.fields(0).type)
	if id_name & "" ="" or ft="" then
		call debug_write ("func_GetFilterDropdownsIfExist: Problem found in SQL : " & dd_found & ". Column has no name but need it. Check and revert.","PRINT")
	end if	
	records_in_loop=0	

	 
	do while not rs1.eof
	    records_in_loop=records_in_loop+1
		if ucase(cstr(in_value))=ucase(cstr(rs1.fields(0).value)) and dd_type <> "DATALIST" then 
			ret_=ret_ & "<option value='" & rs1.fields(0).value & "' selected>" & rs1.fields(1).value & "</option>" & vbcrlf
			'call debug_write (dd_type & "===" & cstr(in_value) & "///" & cstr(rs1.fields(0).value) & "///" & cstr(rs1.fields(1).value) & "---MATCH"  , "" )
		else
			ret_=ret_ & "<option value='" & rs1.fields(0).value & "'>" & rs1.fields(1).value & "</option>" & vbcrlf
			'call debug_write (dd_type & "===" & cstr(in_value) & "///" & cstr(rs1.fields(0).value) & "///" & cstr(rs1.fields(1).value) & ""  , "" )
		end if
		rs1.movenext
	loop
		
	rs1.close
	set rs1=nothing
    
	' #######   Multiple values in operation ADD Records from Drop down list <select multiple> ' added 2021.08.11
	' if Records in multiple dropdown too much then we need to limit maximum available
	if func_ObjectIsPartOfList(g_DBTableMultipleDropdownsFieldsForInsert,in_field,",")="YES" and g_OperationTypeInsertUpdate="INSERT" then 
			multiple_=" multiple size='" & records_in_loop & "' title='To select multiple values from list push CTRL and then select values'" 
		else 
			multiple_=""
	end if		
	' #######
	
	if dd_type = "DATALIST" then 
		ret_="<input list='" & id_name & ft & "' name='" & id_name & ft & "' type='text' value='" & in_value & "'><datalist id='" & id_name & ft & "'>" & ret_ & "</datalist>" & vbcrlf
	else
		ret_="<select name='" & id_name & ft & "' id='uuu' " & multiple_ & ">" & ret_ & "</select>" & vbcrlf
	end if
		 
	func_GetFilterDropdownsIfExist = ret_ 
	
End Function




function rs_field_type(in_value)
' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/datatypeenum?view=sql-server-2017
	dim ret_
	select case in_value
		case 0 ret_="0"'No value adEmpty
		case 2 ret_="0"'A 2-byte signed integer. adSmallInt
		case 3 ret_="0"'A 4-byte signed integer. adInteger
		case 4 ret_="0"'A single-precision floating-point value. adSingle
		case 5 ret_="0"'A double-precision floating-point value. adDouble
		case 6 ret_="0"'A currency value adCurrency
		case 7 ret_="2"'The number of days since December 30, 1899 + the fraction of a day. adDate ' Date in MS Access
		case 8 ret_="1"'A null-terminated character string. adBSTR
		case 9 ret_="0"'A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO. adIDispatch
		case 10 ret_="0"'A 32-bit error code adError
		case 11 ret_="0"'A boolean value. adBoolean
		case 12 ret_="0"'An Automation Variant. Note: Currently not supported by ADO. adVariant
		case 13 ret_="0"'A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO. adIUnknown
		case 14 ret_="0"'An exact numeric value with a fixed precision and scale. adDecimal
		case 16 ret_="0"'A 1-byte signed integer. adTinyInt
		case 17 ret_="0"'A 1-byte unsigned integer. adUnsignedTinyInt
		case 18 ret_="0"'A 2-byte unsigned integer. adUnsignedSmallInt
		case 19 ret_="0"'A 4-byte unsigned integer. adUnsignedInt
		case 20 ret_="0"'An 8-byte signed integer. adBigInt
		case 21 ret_="0"'An 8-byte unsigned integer. adUnsignedBigInt
		case 64 ret_="0"'The number of 100-nanosecond intervals since January 1,1601 adFileTime
		case 72 ret_="1"'A globally unique identifier (GUID) adGUID
		case 128 ret_="1"'A binary value. adBinary
		case 129 ret_="1"'A string value. adChar
		case 130 ret_="1"'A null-terminated Unicode character string. adWChar
		case 131 ret_="0"'An exact numeric value with a fixed precision and scale. adNumeric
		case 132 ret_="1"'A user-defined variable. adUserDefined
		case 133 ret_="2"'A date value (yyyymmdd). adDBDate
		case 134 ret_="1"'A time value (hhmmss). adDBTime
		case 135 ret_="3"'A date/time stamp (yyyymmddhhmmss plus a fraction in billionths). adDBTimeStamp
		case 136 ret_="1"'A 4-byte chapter value that identifies rows in a child rowset adChapter
		case 138 ret_="0"'An Automation PROPVARIANT. adPropVariant
		case 139 ret_="0"'A numeric value (Parameter object only). adVarNumeric
		case 200 ret_="1"'A string value (Parameter object only). adVarChar
		case 201 ret_="1"'A long string value. adLongVarChar
		case 202 ret_="1"'A null-terminated Unicode character string. adVarWChar
		case 203 ret_="1"'A long null-terminated Unicode string value. adLongVarWChar
		case 204 ret_="1"'A binary value (Parameter object only). adVarBinary
		case 205 ret_="1"'A long binary value. adLongVarBinary
		case 0x2000 ret_="1"'A flag value combined with another data type constant. Indicates an array of that other data type. AdArray
	end select
	'response.write ret_ & " " & in_value & " " & in_name & "<br>"
	'qqqq = rs_field_db_type(in_value,in_name)
	rs_field_type=ret_
end function

function rs_field_db_type(in_value,in_name)
' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/datatypeenum?view=sql-server-2017
	dim ret_
	select case in_value
		case 0 ret_="number"'No value adEmpty
		case 2 ret_="number"'A 2-byte signed integer. adSmallInt
		case 3 ret_="number"'A 4-byte signed integer. adInteger
		case 4 ret_="number"'A single-precision floating-point value. adSingle
		case 5 ret_="number"'A double-precision floating-point value. adDouble
		case 6 ret_="number"'A currency value adCurrency
		case 7 ret_="date"'The number of days since December 30, 1899 + the fraction of a day. adDate ' Date in MS Access
		case 8 ret_="text"'A null-terminated character string. adBSTR
		case 9 ret_="number"'A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO. adIDispatch
		case 10 ret_="number"'A 32-bit error code adError
		case 11 ret_="number"'A boolean value. adBoolean
		case 12 ret_="number"'An Automation Variant. Note: Currently not supported by ADO. adVariant
		case 13 ret_="number"'A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO. adIUnknown
		case 14 ret_="number"'An exact numeric value with a fixed precision and scale. adDecimal
		case 16 ret_="number"'A 1-byte signed integer. adTinyInt
		case 17 ret_="number"'A 1-byte unsigned integer. adUnsignedTinyInt
		case 18 ret_="number"'A 2-byte unsigned integer. adUnsignedSmallInt
		case 19 ret_="number"'A 4-byte unsigned integer. adUnsignedInt
		case 20 ret_="number"'An 8-byte signed integer. adBigInt
		case 21 ret_="number"'An 8-byte unsigned integer. adUnsignedBigInt
		case 64 ret_="number"'The number of 100-nanosecond intervals since January 1,1601 adFileTime
		case 72 ret_="text"'A globally unique identifier (GUID) adGUID
		case 128 ret_="text"'A binary value. adBinary
		case 129 ret_="text"'A string value. adChar
		case 130 ret_="text"'A null-terminated Unicode character string. adWChar
		case 131 ret_="number"'An exact numeric value with a fixed precision and scale. adNumeric
		case 132 ret_="text"'A user-defined variable. adUserDefined
		case 133 ret_="date"'A date value (yyyymmdd). adDBDate
		case 134 ret_="time"'A time value (hhmmss). adDBTime
		case 135 ret_="datetime-local"'A date/time stamp (yyyymmddhhmmss plus a fraction in billionths). adDBTimeStamp
		case 136 ret_="text"'A 4-byte chapter value that identifies rows in a child rowset adChapter
		case 138 ret_="number"'An Automation PROPVARIANT. adPropVariant
		case 139 ret_="number"'A numeric value (Parameter object only). adVarNumeric
		case 200 ret_="text"'A string value (Parameter object only). adVarChar
		case 201 ret_="text"'A long string value. adLongVarChar
		case 202 ret_="text"'A null-terminated Unicode character string. adVarWChar
		case 203 ret_="text"'A long null-terminated Unicode string value. adLongVarWChar
		case 204 ret_="text"'A binary value (Parameter object only). adVarBinary
		case 205 ret_="text"'A long binary value. adLongVarBinary
		case 0x2000 ret_="text"'A flag value combined with another data type constant. Indicates an array of that other data type. AdArray
	end select

	rs_field_db_type=ret_

end function

Function Read_Rs(sSQL,delimiter,row_delimiter,include_header)

on error resume next
	Dim rs, rc_null,cn
	dim row_, res_

    Set rs = CreateObject("ADODB.Recordset")
	cn=Application(g_page_datasource)
	rs.open sSQL, cn
	if err <> "" then
		call debug_write ("Read_Rs: Error occured in SQL statement: " & sSQL , "")
	end if
    if rs.bof or rs.eof then
	    rc_null=1 
	end if   
    
	if rc_null<>1 then	   

		if include_header=1 then
			for i=0 to rs.fields.count-1
			  row_ = row_ & rs.fields(i).name & delimiter
			next 
			row_=mid(row_,1,len(row_)-len(delimiter))
		end if
		
		do while not rs.eof
			
			for i=0 to rs.fields.count-1
			    if isnull(rs.fields(i).value) then res_="" else res_ = rs.fields(i).value
				row_=row_ & res_ & delimiter 
			next 
			row_=mid(row_,1,len(row_)-len(delimiter)) & row_delimiter 
			rs.movenext
		loop
	end if
	row_=mid(row_,1,len(row_)-len(row_delimiter)) 
	rs.close
	set rs=nothing
    
	Read_Rs = row_ 
	
End Function


' Checked
Public sub write_log(in_msg)

' to use logging in application and messages generation You need to 
' 1. Create logging table in database 
' 2. example for MS SQL server (table name you can change and correct code below in str_ variabe "insert into log")
' CREATE TABLE dbo.log(
'	id int IDENTITY(1,1) NOT NULL, -- autoincrement id column
'	createdon datetime NULL,
'	visitorip nvarchar(20) NULL,
'	remotehost nvarchar(20) NULL,
'	httphost nvarchar(100) NULL,
'	scriptname nvarchar(10) NULL,
'	querystring nvarchar(max) NULL,
'	log_msg nvarchar(max) NULL
') 
'GO
'
'ALTER TABLE dbo.log ADD  DEFAULT (getdate()) FOR createdon
'GO 
' end of script 

	dim str_
	str_= "insert into log (visitorip,remotehost,httphost,scriptname,querystring,log_msg) values ('#1#','#2#','#3#','#4#','#5#','#6#')"
	str_ = replace(str_,"#1#",Request.Servervariables("REMOTE_ADDR"))
	str_ = replace(str_,"#2#",Request.Servervariables("REMOTE_HOST"))
	str_ = replace(str_,"#3#",Request.Servervariables("HTTP_HOST"))
	str_ = replace(str_,"#4#",Request.Servervariables("SCRIPT_NAME"))
	str_ = replace(str_,"#5#",Request.Servervariables("QUERY_STRING"))
	str_ = replace(str_,"#6#",replace(in_msg,"'","''"))
	execute_SCRIPT(str_)

End sub	
	
Function get_page()	
	dim page 
	page=request.querystring("p")
	if page="" then page=request.form("p")

	if request.querystring("p")="" then 
		page=g_DefaultPageCode
	end if
	get_page=page
End Function	

Function func_getPageName()
    dim tmp_,page_name_x
	tmp_   = request.servervariables("SCRIPT_NAME")
	page_name_x = instr(strreverse(tmp_),"/")-1
	if page_name_x=-1 then page_name_x = len(tmp_)
	'response.write right(tmp_,page_name_x):response.end
	func_getPageName   = right(tmp_,page_name_x)
end Function
	
Function QS()
	QS=Request.ServerVariables("QUERY_STRING") 
End Function

Function func_ReplaceTabColNameWithText(in_column)
	
	dim out_

	if g_use_columns_beautifier="YES" then 
	
		if len(in_column)=0 then
			func_ReplaceTabColNameWithText=""
			exit function
		end if
		
		if g_TableColumnsHeadersSubstitution="" then 
			func_ReplaceTabColNameWithText=in_column
			exit function
		end if
		
		'max id of substitution array always must have key and their pair value 
		max_key_id = ubound(g_ColumnsSubstitutionKeyValue)
		if max_key_id/2<>max_key_id\2 then 
			call debug_write("Substitution array problem. Must always have key and value pairs. Total count:" & max_key_id,"")
			exit function
		end if
		
		for i=0 to max_key_id step 2
			key_ = g_ColumnsSubstitutionKeyValue(i)
			if ucase(key_) = ucase(in_column) then out_ = g_ColumnsSubstitutionKeyValue(i+1): if g_use_columns_beautifier_compact="YES" then out_ = replace(out_," ","<br>")
		next 	

	end if
	if out_ ="" then out_ = in_column	
	func_ReplaceTabColNameWithText=out_
	
End Function


Function func_GenerateMenu()

	Dim ret_ ' collector of menu
	Dim f_MenuLevelOneTemplate
	Dim f_MenuLevelTwoTemplate
	
	f_MenuLevelOneTemplate = "<a href='" & page_name & "?p=#PAGE_CODE#'>#PAGE_NAME#</a>" & vbcrlf
	f_MenuLevelTwoTemplate = "<button class='dropdown-btn'>#PAGE_NAME#<i class='fa fa-caret-down'></i></button><div class='dropdown-container'>#S_ITEMS#</div>" & vbcrlf
	
	Dim f_MenuItemName
	Dim f_MenuItemPageCode ' page_code_
	Dim f_MenuButtonName 'button_name_
	Dim f_MenuButtonSubitems 'tmp_row_
	
	Dim i
	Dim w
	
	'on error resume next
	
	if g_MENU<>"" then ' \ 1
		if instr(g_MENU,"]")=0 then exit function
		Dim f_arr_level1
		Dim f_arr_level2 
		f_arr_level1 = Split(g_MENU,";")
		
		if ubound(f_arr_level1)>=0 then ' \ 2
			for i=0 to ubound(f_arr_level1)
				'call debug_write (f_arr_level1(i),"")
				
				if len(f_arr_level1(i))>2 then 	
				
					f_arr_level2 = Split(f_arr_level1(i),":")
					
					call debug_write ("func_GenerateMenu: Level 1 {" & f_arr_level1(i) & "}"      ,   "")
					
					if left(f_arr_level2(0),1)="(" then
						if right(f_arr_level2(0),1)=")" then
						   f_MenuItemName = replace(replace(f_arr_level2(0),"(",""),")","")			
						   f_MenuItemPageCode = f_arr_level2(1)
						   ret_=ret_ & replace(  replace(f_MenuLevelOneTemplate,"#PAGE_NAME#",f_MenuItemName)  ,"#PAGE_CODE#",f_MenuItemPageCode)
						end if
					end if

					if left(f_arr_level2(0),1)="[" then
						if right(f_arr_level2(0),1)="]" then
							f_MenuButtonName = replace(replace(f_arr_level2(0),"[",""),"]","")			
							if ubound(f_arr_level2)>0 then			
								f_MenuButtonSubitems=""
								for w=1 to ubound(f_arr_level2)-1 step 2
									'call debug_write (f_arr_level2(w),"")
									
									f_MenuItemPageCode = f_arr_level2(w)
									f_MenuItemName = replace(replace(f_arr_level2(w+1),"{",""),"}","")
									
									call debug_write ("func_GenerateMenu: Level 2 : Page {" & f_MenuItemPageCode & "} Name {" & f_MenuItemName & "}"      ,   "")
									
									f_MenuButtonSubitems = f_MenuButtonSubitems & replace(  replace(f_MenuLevelOneTemplate,"#PAGE_CODE#",f_MenuItemPageCode) ,"#PAGE_NAME#",f_MenuItemName) 
								next 
								ret_= ret_ & replace(   replace(f_MenuLevelTwoTemplate,"#PAGE_NAME#", f_MenuButtonName), "#S_ITEMS#",f_MenuButtonSubitems )
							end if 					
						end if
					end if
				end if			
			
			next
		end if ' / 2
		
		' if mode edit or insert or button "submit data" pressed, we disable Export To Excel on menu level
		if Request.QueryString("op")<>"i" and Request.QueryString("op")<>"e" and Request.QueryString("a")<>"a" then 
			ret_ = ret_ & "<a id='downloadLink' onclick='exportToExcel(this)'>Export to excel</a>" & vbcrlf
		end if
		ret_ = ret_ & "<a href='#' onclick='showNewsInfo()'>Info/Help</a>" & vbcrlf
		if g_debug_flag="YES" then
			ret_ = ret_ & "<a href='#' onclick='showDebugContent()'>Debug Log</a>" & vbcrlf
		end if
		'ret_ = "<div class='sidenav'><div class='sidenav-logo'>Mooring Lines</div>  " & vbcrlf & "<img src='ml.jpg'>" & vbcrlf & ret_ & "</div>"
		
		ret_ = "<div class='sidenav'><div class='sidenav-logo'>" & g_PortalName & "</div>  " & vbcrlf & g_GlobalVariables & vbcrlf & ret_ & "</div>"
	end if' / 1
	
	func_GenerateMenu = ret_
	
End Function

Function func_DateTimeFormat(in_mask, in_date)

	' input parameters check
	if in_date & "" = "" or in_mask="" then
		out_ ="" 
	else
		dim now_,yyyy,yy,mm,m,dd,d,hh,h,mi,ms,ss,s
		yyyy = year(in_date):yy = right(yyyy,2)
		mm = right("0" & month(in_date),2):m=month(in_date)
		dd = right("0" & day(in_date),2):d=day(in_date)
		hh = right("0" & hour(in_date),2):h=hour(in_date)
		mi = right("0" & minute(in_date),2):ms=minute(in_date) ' mi is month 01,02,03...12, and ms is short month 1,2,3...12
		ss = right("0" & second(in_date),2):s=second(in_date)
		out_ = replace(in_mask,"yyyy",yyyy):	
		out_ = replace(out_,"mm",mm)
		out_ = replace(out_,"dd",dd)
		out_ = replace(out_,"hh",hh)
		out_ = replace(out_,"mi",mi)
		out_ = replace(out_,"ss",ss)
		out_ = replace(out_,"yy",yy)
		out_ = replace(out_,"m",m)
		out_ = replace(out_,"d",d)
		out_ = replace(out_,"h",h)
		out_ = replace(out_,"ms",ms)
		out_ = replace(out_,"s",s)
	end if	
	
	func_DateTimeFormat = out_

end Function

Function func_CheckVulnerableElementsInQueryString()
	
	dim elements_
	dim key_
	dim value_
	dim cells_
    dim vulnerability_result
	dim req_qs
	
	req_qs = QS() 'Request.Servervariables("QUERY_STRING")
	
	call debug_write ("Query String = {" & req_qs & "}","")
	
	if instr(req_qs,"&")=0 then exit function
	if instr(req_qs,"=")=0 then exit function
	if len(req_qs)>1 then
       if left(req_qs,1)="&" then req_qs = mid(req_qs,2,len(req_qs)-1)
	else
		if len(req_qs>0) then call debug_write ("Abnormal Query String = {" & req_qs & "} detected. First element not equal to '&'.","")
		exit function
	end if	
	
	elements_=split(req_qs,"&")
	
	for i=0 to ubound(elements_)
        
		cells_=split(elements_(i),"=")

		select case ubound(cells_)
			Case 0
				key_ = cells_(0)
				value_ = ""
			Case 1
				key_ = cells_(0)
				value_ = cells_(1)
			Case else
				key_=""
				value_=""
		End Select 		
		call debug_write ("Loop through Query String = {" & req_qs & "}  " & i+1 & " pair = {" & elements_(i) & "} key = {" & key_ & "} value = {" & value_ & "}","")
		
		if key_="" and value_="" then

			vulnerability_result = vulnerability_result & "Query String " & req_qs & " Pair : <font color='green'>" & elements_(i) & "</font> has abnormal count of keys and values in it. Vulnerable elements: {<font color='red'>" & elements_(i) & "</font>}<br>"
				
		else
			tmp_ = func_VulnerableElementsCheck(key_)
			if tmp_<>"" then
				vulnerability_result = vulnerability_result & "Query String Key: <font color='green'>" & key_ & "</font> Value = <font color='green'>" & value_ & "</font> Vulnerable elements in key_: {<font color='red'>" & tmp_ & "</font>}<br>"
			end if
			tmp_ = func_VulnerableElementsCheck(value_)
			if tmp_<>"" then
				vulnerability_result = vulnerability_result & "Query String Key: <font color='green'>" & key_ & "</font> Value = <font color='green'>" & value_ & "</font> Vulnerable elements in value_: {<font color='red'>" & tmp_ & "</font>}<br>"
			end if
		end if
	next

	func_CheckVulnerableElementsInQueryString = vulnerability_result

End Function

Function func_VulnerableElementsCheck(in_statement)

	' vulnerable elements detector: return "" if not found 
	' reference
	'http://web.archive.org/web/20130401091931/http://www.comsecglobal.com/FrameWork/Upload/SQL_Smuggling.pdf
	'https://www.netsparker.com/blog/web-security/sql-injection-cheat-sheet/

	dim a(98),b(9)
	dim ret_
	dim tmp_
	
	tmp_ = replace(replace(replace(lcase(in_statement),vbcr,""),vblf,""),vbtab,"")
	b(0) = "select "
	b(1) = " from "
	b(2) = " union "
	b(3) = "delete from"
	b(4) = " and "
	b(5) = " or "
	b(6) = "%20or%20"
	b(7) = "select%20"
	b(8) = " insert into "
	b(9) = "1/0"
	for i=0 to ubound(b) 
	  if instr(lcase(in_statement),b(i))<>0 then ret_ = "b(" & i & ") " & ret_ & b(i) & " "
	next
	
	tmp_ = replace(replace(tmp_," ",""),"%20","")
	a(0) = "--"
	a(1) = ");"
	a(2) = "';"
	a(3) = "droptable"
	a(4) = "deletefrom"
	a(5) = "password"
	a(6) = "'admin'"
	a(7) = "drop/*"
	a(8) = "'true'"
	a(9) = "'false'"
	a(10) = ")='"
	a(11) = "char("
	a(12) = " + "
	a(13) = "unionselect"
	a(14) = vbcrlf
	a(15) = vblf
	a(16) = vbcr
	a(17) = "'='"
	a(18) = "null,"
	a(19) = "master."
	a(20) = "sys."
	a(21) = "/*"
	a(22) = "*/"
	a(23) = "'%'"
	a(24) = "'*'"
	a(25) = "','"
	a(26) = "'and"
	a(27) = "'or"
	a(28) = "values("
	a(29) = "'%25'" ' %
	a(30) = "%27" ' '
	a(31) = "%5C"
	a(32) = "%29" ' )
	a(33) = "%00" ' )
	a(34) = "%01" ' )
	a(35) = "%02" ' )
	a(36) = "%03" ' )
	a(37) = "%04" ' )
	a(38) = "%05" ' )
	a(39) = "%06" ' )
	a(40) = "%07" ' )
	a(41) = "%08" ' )
	a(42) = "%09" ' )
	a(43)="%0A"
	a(44) ="%0B"
	a(45) ="%0C"
	a(46) ="%0D"
	a(47) ="%0E"
	a(48) ="%0F"
	a(49) ="%10"
	a(50) ="%11"
	a(51) ="%12"
	a(52) ="%13"
	a(53) ="%14"
	a(54) ="%15"
	a(55) ="%16"
	a(56) ="%17"
	a(57) ="%18"
	a(58) ="%19"
	a(59) ="%1A"
	a(60) ="%1B"
	a(61) ="%1C"
	a(62) ="%1D"
	a(63) ="%1E"
	a(64) ="%1F"
	a(65) = "U+02BC"
	a(66) = "concat("
	a(67) = "ascii("
	a(68) = "unionall"
	a(69) = "hex("
	a(70) = "'or"
	a(71) = "'#"
	a(72) = "md5("
	a(73) = "execsp_"
	a(74) = "execmaster."
	a(75) = "sysmessages"
	a(76) = "sysservers"
	a(77) = "xp_reg"
	a(78) = "declare@"
	a(79) = "limit0,0"
	a(80) = "information_schema"
	a(81) = "waitfordelay"
	a(82) = "benchmark("
	a(83) = "sleep("
	a(84) = "select*"
	a(85) = "sha1("
	a(86) = "password("
	a(87) = "compress("
	a(88) = "row_count("
	a(89) = "schema("
	a(90) = "version("
	a(91) = "@@version"
	a(92) = "openrowset("
	a(93) = "load_file("
	a(94) = "utl_http."
	a(95) = "utl_inaddr."
	a(96) = "dbms_ldap."
	a(97) = "utl_inaddr."
	a(98) = "%3B"
	'a(31) = "%28" ' (
		
	for i=0 to ubound(a) 
	  if instr(lcase(tmp_),lcase(a(i)))<>0 then ret_ = "a(" & i & ") " & ret_ & a(i) & " "
	next
	
	func_VulnerableElementsCheck = ret_
	
End Function





Function func_CreateGlobalVariablesDD(in_data) ' in_name,in_sql

	Dim arr_, in_name, in_sql
	Dim rs1, cn1, ret_
	dim CurrentValue
	
	arr_ = split(in_data,";")
	
	for i=0 to ubound(arr_) step 2
	
			in_name = arr_(i)
			in_sql = arr_(i+1)
			
			'CurrentValue = request.querystring(in_name)
			'if CurrentValue<>"" and CurrentValue<>session(in_name) then session(in_name) = CurrentValue
			'if session(in_name)<>"" then CurrentValue = session(in_name) 
			
			if request.querystring(in_name).count>0 then ' if value was POSTed in form
				CurrentValue = request.querystring(in_name): session(in_name) = CurrentValue
				if CurrentValue<>session(in_name) then session(in_name) = CurrentValue
			else
				CurrentValue = session(in_name)
			end if
			
			Set rs1 = CreateObject("ADODB.Recordset")
			cn1=Application(g_page_datasource)

			rs1.open in_sql, cn1
			if err.number<>0 then 
				call debug_write ("Error occured in func_CreateGlobalVariablesDD: " & in_name , "")
				exit function    
			end if	
			'id_name=rs1.fields(0).name
			ft = rs_field_type(rs1.fields(0).type)
			ret_=ret_ & "<div class='sidenav-GlobalObjects'><form><select name='" & in_name & "' onchange='this.form.submit()'> " 
				
			do while not rs1.eof
				if cstr(CurrentValue)=cstr(rs1.fields(0).value) then 
					ret_=ret_ & "<option value='" & rs1.fields(0).value & "' selected>" & rs1.fields(0).value & "</option>" & vbcrlf
				else
					ret_=ret_ & "<option value='" & rs1.fields(0).value & "'>" & rs1.fields(0).value & "</option>" & vbcrlf
				end if
				rs1.movenext
			loop
				
			rs1.close
			set rs1=nothing
			
			ret_=ret_ & "</select></form></div>" & vbcrlf
		
	next
	
	func_CreateGlobalVariablesDD = ret_ 
	
End Function


Function func_GetGlobalVariablesValues()

	Dim ret_
	
		For Each x In session.contents
			if x<>"url" and x<>"key" then ret_ = ret_ & x & ";" & session.contents(x) & ";"
		Next
		
		if len(ret_)<>0 then ret_ = mid(ret_,1,len(ret_)-1)
		call debug_write("func_GetGlobalVariablesValues: " & ret_,"")
		
	func_GetGlobalVariablesValues = ret_
	
End Function
%>
</body>
</html>
