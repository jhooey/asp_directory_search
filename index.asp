<!DOCTYPE html>
<html>
<head>
</head>
<body>
<%

'List files in an array'
dim path, strfile
dim filename() 
dim creationDate()
dim i

path=server.mappath("\") &"\SSC\Archivedcharts\"

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
	
dim currentYear
dim currentMonth
currentYear = 1
currentMonth = 20
missingMonthCheck = 12

For i=0 to max
	'Formatting code - will group all files into a single year'
	if currentYear <> Year(creationDate(i)) then
	
		if i <> 0 then
			%></div><%
		end if
		
		currentYear = Year(creationDate(i))	
		%>
			<h1 style="margin:10px 0 10px 0;"><%=currentYear%></h1>
			<div style="width:800px;padding:0 0 30px 0;">
		<%
		currentMonth = 20
	end if

	if currentMonth <> Month(creationDate(i)) then
		currentMonth = Month(creationDate(i))
		%>
			<div style="float:right; margin:0 11px 0 11px"><a href="<% response.write(path & filename(i)) %>"><%= MonthName(Month(creationDate(i))) %></a></div>
		<%	
	end if
next

%></div><%
%>
</body>
</html> 