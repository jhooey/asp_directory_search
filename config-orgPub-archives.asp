<%
dim numColumns, distanceModifier, path, filetype, department

filetype = ".htm"
department = "SSC"
path=server.mappath("\") & "\Archivedcharts\" & department & "\" 

distanceModifier = 1 	'Controls the distance between listed files - margin: (# of days between files)* modifier
numColumns = 5  		'MAX columns is 7. for best results use 5
%>