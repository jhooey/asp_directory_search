<!--#include file="templates/header-orgPub-Archives.asp"-->

<%
dim numColumns, distanceModifier, path, filetype, department

department = "SSC"		'name of the folder in archivedcharts

filetype = ".htm"		'for all files set to ".*"


path=server.mappath("\") & "\Archivedcharts\" & department & "\" 

distanceModifier = 1 	'Controls the distance between listed files - margin: (# of days between files)* modifier
numColumns = 5  		'MAX columns is 7. for best results use 5
%>

<!--#include virtual="./templates/content-orgPub-Archives.asp"-->
<!--#include virtual="./templates/footer-orgPub-Archives.asp"-->