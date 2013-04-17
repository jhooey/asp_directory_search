<!DOCTYPE html>
<html>
<body>
<%

dim path, strfile

path=server.mappath("\") &"./SSC/Archivedcharts"

set fs = server.createobject("Scripting.FileSystemObject")
set archiveFolder = fs.getfolder(path)
set archivedFiles = archiveFolder.files			'list of files

for each strfile in archivedFiles
	If InStr(strfile.name, ".htm") Then	
%>
		<a href="./testfiles/<%= strfile.name %>"><%= strfile.name %></a> <%= MonthName(Month(strfile.DATECREATED)) %><br>
<%	End If
next

set fs = nothing
set archiveFolder = nothing
set archivedFiles = nothing
%>
</body>
</html> 