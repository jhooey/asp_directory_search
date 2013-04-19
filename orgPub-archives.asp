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


dim previousYear, datesDiff, previousFileDate, divCount

previousFileDate = 0
previousYear = 1
divCount = 0

columnWidth = 100 / numColumns

for i=0 to ubound(files, 2) - 1

	if i = 0 then
	
		%> 
		<div class="orgPub-archives">
			<div class="orgPub-archive-year" style="width: <%=columnWidth%>%">
				<h3><%=Year(files( 1, i))%></h3> 
				<ul>
		<%
		
		previousYear = Year(files( 1, i))
	
	elseif previousYear <> Year(files( 1, i)) then
		%></ul><%
	
		while previousYear > Year(files( 1, i)) 
			previousYear = previousYear - 1
			%></div><%
			divCount = divCount + 1
			
			if divCount >= numColumns then 
				divCount = 0
				%>
					</div>
					<div class="orgPub-archives">
				<%
			end if
			
			%>
			<div class="orgPub-archive-year" style="width: <%=columnWidth%>%">
				<h3><%=previousYear%></h3>
			<%
			
		wend

		%><ul><%
		
		previousFileDate = 0
		
	end if
	
	if ( previousFileDate <> 0 and previousYear = Year(files( 1, i))) then
		datesDiff = ( DatePart( "y", previousFileDate) - DatePart( "y", files( 1, i)) ) * distanceModifier
	else 
		datesDiff = 0
	end if
	
	
	%><li style="margin-top: <%=datesDiff%>px">
		<a href="<% response.write(path & files( 0, i)) %>"><%=MonthName(Month(files( 1, i)))%>&nbsp;&nbsp;<%=Day(files( 1, i))%></a>
	</li>
	<%
	
	previousYear = Year(files( 1, i))
	previousFileDate = files( 1, i)
	
next

if ( previousFileDate <> 0 ) then
  %></ul>
	</div>
	<%
end if

%>
</div>
</body>
</html> 