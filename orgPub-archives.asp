<!DOCTYPE html>
<html>
<head>
<link rel="stylesheet" href="/css/orgPub-archives.css">

<!--#include file="scripts/functions-orgPub-archives.asp"-->
<!--#include file="config-orgPub-archives.asp"-->

</head>
<body>
<%
dim files

'Create a list of all the files with the filetype extension in the path (defined in config.asp)
files = grabFiles ( path, filetype)

'Sort the files in descending order by date created
files = sortFilesDesc (files)


dim previousYear, datesDiff, previousFileDate, divCount, filename

previousFileDate = 0  	'0 means its the first file in the year
previousYear = 1		'Used to group files by year
divCount = 0			'keeps track of the number of columns

columnWidth = 100 / numColumns	'adjust width to reflect the number of columns

'This loop writes out all the HTML for displaying the files
for i=0 to ubound(files, 2) - 1

	'runs only once, and creates the first divs
	if i = 0 then
		%> 
		<div class="orgPub-archives">
			<div class="orgPub-archive-year" style="width: <%=columnWidth%>%">
				<h3><%=Year(files( 1, i))%></h3> 
				<ul>
		<%
		
		previousYear = Year(files( 1, i))
	
	'runs only when a new year has begun
	'closes all the divs from the previous year
	
	elseif previousYear <> Year(files( 1, i)) then
		%></ul><%
	
		'this loop will also create empty columns if no files are found for a specified year
		while previousYear > Year(files( 1, i)) 
			previousYear = previousYear - 1
			%></div><%
			divCount = divCount + 1		'this number represents the number of columns already created
			
			'checks to see if this is the final column,  
			'if true, a new row is started
			if divCount >= numColumns then 
				divCount = 0
				%>
					</div>
					<div class="orgPub-archives">
				<%
			end if
			
			'HTML code for a new column
			%>
			<div class="orgPub-archive-year" style="width: <%=columnWidth%>%">
				<h3><%=previousYear%></h3>
			<%
			
		wend

		%><ul><%
		
		'indicates that this file is the first in a new year
		previousFileDate = 0
		
	end if
	
	'Defines the margin-top px count, the margin increases if the file dates are further apart
	'gives a visual sense of time when file links are being displayed
	if ( previousFileDate <> 0 and previousYear = Year(files( 1, i))) then
		datesDiff = ( DatePart( "y", previousFileDate) - DatePart( "y", files( 1, i)) ) * distanceModifier
	else 
		datesDiff = 0
	end if

	'Creates the links to the files
	'Display text is the month and day the file was created
	%><li style="margin-top: <%=datesDiff%>px">
		<a href="<% response.write(path & filename) %>"><%=MonthName(Month(files( 1, i)))%>&nbsp;&nbsp;<%=Day(files( 1, i))%></a>
	</li>
	<%
	
	previousYear = Year(files( 1, i))
	previousFileDate = files( 1, i)
	
next

'Will run only if the for loop has been used at least once
if ( previousFileDate <> 0 ) then
  %></ul>
	</div>
	<%
end if

%>
</div>
</body>
</html> 