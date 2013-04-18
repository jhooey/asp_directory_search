<!DOCTYPE html>
<html>
<head>
<link rel="stylesheet" href="style.css">
</head>
<body>
<%
'List files in an array'
dim path, strfile
dim filename() 
dim creationDate()
dim i


path=server.mappath("\") & "\SSC\Archivedcharts\"

set fs = server.createobject("Scripting.FileSystemObject")
set archiveFolder = fs.getfolder(path)
set archivedFiles = archiveFolder.files			'list of files

i = 0
for each strfile in archivedFiles
	If InStr(strfile.name, ".htm") Then	
	
		redim preserve filename(i + 1)
		redim preserve creationDate(i + 1)
		
		filename(i) = strfile.name
		creationDate(i) = strfile.DATECREATED
		
		i = i +1	
	End If
next

set fs = nothing
set archiveFolder = nothing
set archivedFiles = nothing
	
	

'Sort Arrays in order of Date Created
dim max, j, TempVar1, TempVar2
max = ubound(creationDate) - 1

For i=0 to max 
	For j=i+1 to max 
		If creationDate(i)<creationDate(j) then
			TempVar1=creationDate(i)
			TempVar2=filename(i)
			creationDate(i)=creationDate(j)
			creationDate(j)=TempVar1
			filename(i)=filename(j)
			filename(j)=TempVar2
		end if
	next 
next	

dim previousYear, datesDiff, previousFileDate, distanceModifier, divCount

previousFileDate = 0
previousYear = 1
distanceModifier = 2
divCount = 0

for i=0 to max 

	if i = 0 then
		%> 
	<div class="archives">
	<div class="archive_year">
	<h3><%=Year(creationDate(i))%></h3> 
	<ul>
		<%
		previousYear = Year(creationDate(i))
		
	elseif previousYear <> Year(creationDate(i)) then
		%>
	</ul>
		<%
	
		while previousYear > Year(creationDate(i)) 
			previousYear = previousYear - 1
			%>
	</div>
	<%
	divCount = divCount + 1
	
	if divCount >= 4 then 
		divCount = 0
		%> </div><div class="archives"> <%
	end if
	
	
	%>
	<div class="archive_year"><h3><%=previousYear%></h3>
			<%
			
		wend

  %><ul>
  
		<%
		previousFileDate = 0
		
	end if
	
	if ( previousFileDate <> 0 and previousYear = Year(creationDate(i))) then
		
		datesDiff = ( DatePart( "y", previousFileDate) - DatePart( "y", creationDate(i)) ) * distanceModifier
	else 
		datesDiff = 0
	end if
	%>  <li style="margin-top: <%=datesDiff%>px">
				<a href="<% response.write(path & filename(i)) %>"><%=MonthName(Month(creationDate(i)))%>&nbsp;&nbsp;<%=Day(creationDate(i))%></a>
		</li><%
	
	previousYear = Year(creationDate(i))
	previousFileDate = creationDate(i)
	
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