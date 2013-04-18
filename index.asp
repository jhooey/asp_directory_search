<!DOCTYPE html>
<html>
<head>
<link rel="stylesheet" href="style.css">
</head>
<body>
<%

Function fillGaps ( missingMonthCheck, monthCreationDate ) 
	While (missingMonthCheck <> monthCreationDate And  missingMonthCheck > 0)
		%>
			<div class="calendarMonth"><%= MonthName(missingMonthCheck) %></div>
		<%
		
		missingMonthCheck = missingMonthCheck - 1
	wEnd
	
	fillGaps = missingMonthCheck
end function

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
	
dim currentYear, currentMonth
dim duplicateMonth

duplicateMonth = false

currentYear = 1
currentMonth = 20
missingMonthCheck = 12

For i=0 to max
	
	'if the previous value in the array is the same as the current one
	'set it up so that nothing will be written out
	if currentMonth = Month(creationDate(i)) then
		missingMonthCheck = missingMonthCheck + 1
		duplicateMonth = true
	end if
	
	if currentYear <> Year(creationDate(i)) and missingMonthCheck > 0 and i <> 0 then
		
		missingMonthCheck = fillGaps ( missingMonthCheck, Month(creationDate(i)) )
		
	'Formatting code - will group all files into a single year'
	end if 
	
	if currentYear <> Year(creationDate(i)) then
	
		if i <> 0 then
			%></div><%
		end if
		
		currentYear = Year(creationDate(i))	
		%>
			<h1><%=currentYear%></h1>
			<div class="calendarYear">
		<%
		currentMonth = 20
		missingMonthCheck = 12
	end if
	
	
	missingMonthCheck = fillGaps ( missingMonthCheck, Month(creationDate(i)) )

	
	
	if currentMonth <> Month(creationDate(i)) then
		currentMonth = Month(creationDate(i))
		%>
			<div class="calendarMonth"><a href="<% response.write(path & filename(i)) %>"><%= MonthName(Month(creationDate(i))) %></a></div>
		<%
		
		missingMonthCheck = currentMonth - 1
	end if
	
	'If this was a duplicate month, then prepare the missing check for the next pass
	if duplicateMonth then
		missingMonthCheck = missingMonthCheck - 1
		duplicateMonth = false
	end if
	
	
next
%></div><%
%>
</body>
</html> 