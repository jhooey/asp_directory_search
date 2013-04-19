<%
dim numColumns, distanceModifier, path, filetype, company

filetype = ".htm"
company = "SSC"
path=server.mappath("\") & "\" & company & "\Archivedcharts\"

distanceModifier = 1 	'Controls the distance between listed files - margin: (# of days between files)* modifier
numColumns = 5  		'MAX columns is 7. for best results use 5
%>