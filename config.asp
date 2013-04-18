<%
dim numColumns, distanceModifier, path, filetype

path=server.mappath("\") & "\SSC\Archivedcharts\"
filetype = ".htm"

distanceModifier = 1 	'Controls the distance between listed files - margin: (# of days between files)* modifier
numColumns = 5  		'MAX columns is 7. for best results use 5
%>