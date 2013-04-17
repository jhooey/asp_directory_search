function ReplaceContent(){
document.write('<object id="OrgPublisherPluginX" name="PluginX"');
document.write('  classid="clsid:C3CBFE35-9BE8-11D1-B31B-006008948294"');
if ("https:" == document.location.protocol)
{
document.write('codebase="https://www.aquire.com/codebase101/OrgPubX.cab#Version=10,1,3027,1"');
}
else
{
document.write('codebase="http://www.aquire.com/codebase101/OrgPubX.cab#Version=10,1,3027,1"');
}
document.write('  align="middle" border="0" width="100%" height="100%">');
document.write('  <param name="_Version" value="65536">');
document.write('  <param name="_ExtentX" value="17110">');
document.write('  <param name="_ExtentY" value="8431">');
document.write('  <param name="_StockProps" value="0">');
document.write('<param name="ChartOCPFile" value="20130411_ssc_eng.ocp">');
 if (szSelectedBox.length > 0)
 {
document.write('  <PARAM name="SelectedBox" value="'+ szSelectedBox +'">');
}
document.write('  </OBJECT>');
}
