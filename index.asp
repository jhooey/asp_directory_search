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

dim previousYear, datesDiff, previousFileDate, distanceModifier

previousFileDate = 0
previousYear = 1
distanceModifier = 1

for i=0 to max 

	if i = 0 then
		%> 
	<h3 class="archive_year"><%=Year(creationDate(i))%></h3> 
	<ol class="archives_list">
		<%
		previousYear = Year(creationDate(i))
		
	elseif previousYear <> Year(creationDate(i)) then
		%>
	</ol>
		<%
	
		while previousYear > Year(creationDate(i)) 
			previousYear = previousYear - 1
			%>
	<h3 class="archive_year"><%=previousYear%></h3>
			<%
		wend

		%>
	<ol class="archives_list">
		<%
		previousFileDate = 0
		
	end if
	
	if ( previousFileDate <> 0 and previousYear = Year(creationDate(i))) then
		
		datesDiff = ( DatePart( "y", previousFileDate) - DatePart( "y", creationDate(i)) ) * distanceModifier
	else 
		datesDiff = 0
	end if
	%>
		<li style="margin-top: <%=datesDiff%>px">
			<span class="date">
				<a href="<% response.write(path & filename(i)) %>"><%=MonthName(Month(creationDate(i)))%>&nbsp;&nbsp;<%=Day(creationDate(i))%></a>
			</span>
		</li>
	<%
	
	previousYear = Year(creationDate(i))
	previousFileDate = creationDate(i)
	
next

if ( previousFileDate <> 0 ) then
	%>
	</ol>
	<%
end if

%>
</body>
</html> 