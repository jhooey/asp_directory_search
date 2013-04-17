<!DOCTYPE html>
<html>
<body>
<%

'List files in an array'
dim path, strfile
dim filename() 
dim creationDate()
dim i

path=server.mappath("\") &"./SSC/Archivedcharts"

set fs = server.createobject("Scripting.FileSystemObject")
set archiveFolder = fs.getfolder(path)
set archivedFiles = archiveFolder.files			'list of files

i = 0
for each strfile in archivedFiles
	redim preserve filename(i + 1)
	redim preserve creationDate(i + 1)
	
	If InStr(strfile.name, ".htm") Then	
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

max = ubound(creationDate)

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
	
For i=0 to max
%>
	<a href="./testfiles/<%= filename(i) %>"><%= filename(i) %></a> <%= MonthName(Month(creationDate(i))) %><br>
<%	
next

%>
</body>
</html> 