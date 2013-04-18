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
%>