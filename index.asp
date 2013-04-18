<!DOCTYPE html>
<html>
<head>
<link rel="stylesheet" href="style.css">
</head>
<body>

<%
dim path, numColumns, distanceModifier

path=server.mappath("\") & "\SSC\Archivedcharts\"
distanceModifier = 1
numColumns = 5

'List files in an array'
dim strfile, i

dim files()

set fs = server.createobject("Scripting.FileSystemObject")
set archiveFolder = fs.getfolder(path)
set archivedFiles = archiveFolder.files			'list of files

i = 0
for each strfile in archivedFiles
	If InStr(strfile.name, ".htm") Then	
	

		
		redim preserve files( 2, i + 1)
				
		files( 0, i ) = strfile.name
		files( 1, i ) = strfile.DATECREATED
		
		i = i +1	
	End If
next

set fs = nothing
set archiveFolder = nothing
set archivedFiles = nothing
	
	

'Sort Arrays in order of Date Created
dim max, j, TempVar1, TempVar2
max = ubound(files, 2) - 1

For i=0 to max 
	For j=i+1 to max 
		If files( 1, i)<files( 1, j) then
			TempVar1=files( 1, i)
			TempVar2=files( 0, i)
			files( 1, i)=files( 1, j)
			files( 1, j)=TempVar1
			files( 0, i)=files( 0, j)
			files( 0, j)=TempVar2
		end if
	next 
next	

dim previousYear, datesDiff, previousFileDate, divCount

previousFileDate = 0
previousYear = 1
divCount = 0

columnWidth = 100 / numColumns
for i=0 to max 

	if i = 0 then
		%> 
	<div class="archives">
	<div class="archive_year" style="width: <%=columnWidth%>%">
	<h3><%=Year(files( 1, i))%></h3> 
	<ul>
		<%
		previousYear = Year(files( 1, i))
		
	elseif previousYear <> Year(files( 1, i)) then
		%>
	</ul>
		<%
	
		while previousYear > Year(files( 1, i)) 
			previousYear = previousYear - 1
			%>
	</div>
	<%
	divCount = divCount + 1
	
	if divCount >= numColumns then 
		divCount = 0
		%> </div><div class="archives"> <%
	end if
	
	
	%>
	<div class="archive_year" style="width: <%=columnWidth%>%"><h3><%=previousYear%></h3>
			<%
			
		wend

  %><ul>
  
		<%
		previousFileDate = 0
		
	end if
	
	if ( previousFileDate <> 0 and previousYear = Year(files( 1, i))) then
		
		datesDiff = ( DatePart( "y", previousFileDate) - DatePart( "y", files( 1, i)) ) * distanceModifier
	else 
		datesDiff = 0
	end if
	%>  <li style="margin-top: <%=datesDiff%>px">
				<a href="<% response.write(path & files( 0, i)) %>"><%=MonthName(Month(files( 1, i)))%>&nbsp;&nbsp;<%=Day(files( 1, i))%></a>
		</li><%
	
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