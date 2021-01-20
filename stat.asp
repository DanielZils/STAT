<%@ Language=JavaScript %>
<!--#include virtual="/pwserver/framework/ntworb.asp"-->
<!--#include virtual="/pwserver/framework/momasphelper.asp"-->
<% 
function GetRequestVariable(type)
{
	var strResult = Request.QueryString (type)();
	
	if (strResult + "" == "undefined")
	{
		strResult = Request.Form (type)();
		if (strResult + "" == "undefined")
			strResult = "";
	}
	
	return strResult;
}

var strSQL = GetRequestVariable("SQL");
   if (strSQL == "" )
   {
   	strSQL = "SELECT TOP 100 Service_T.Name AS Service, ServiceConsumer_T_1.Name AS CostCenter,\n ServiceProvider_T.Name AS Printer," + 
                      " ServiceProvider_t.Servername, ServiceConsumer_T.Name AS Username, \n ServiceUsage_T.UsageEnd AS UsageEnd, " + 
                      " ServiceUsage_T.Cardinality AS Anzahl, \n ServiceUsage_T.AmountPaid AS Betrag \n" +
                      "FROM ServiceUsage_T \n INNER JOIN " +
                      "ServiceProvider_T ON ServiceUsage_T.ServiceProvider = ServiceProvider_T.ID \n LEFT OUTER JOIN " +
                      "ServiceConsumer_T ServiceConsumer_T_1 ON ServiceUsage_T.ServConsProject = ServiceConsumer_T_1.ID \n INNER JOIN " +
                      "Service_T ON ServiceUsage_T.Service = Service_T.ID \n LEFT OUTER JOIN " +
                      "ServiceConsumer_T ON ServiceUsage_T.ServiceConsumer = ServiceConsumer_T.ID " +
		"\n ORDER BY ServiceUsage_T.UsageEnd DESC";

   }

// var DsPcDb = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=PbaIp;Password=ntwsqlpwd;Initial Catalog=DsPcDb;Data Source=FRASV030754\CANON;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 //EN">
<HTML>
<STYLE TYPE="text/css">
<!--
a  {
    font-family:Tahoma,sans-serif; text-align: right; font-size:8pt;
   }

tr {
    font-family:Tahoma,sans-serif; text-align: right font-size:10pt;
   }
tr.S1 {
    background-color:#DDDDDD
   }
tr.S0 {
    background-color:#FFFFFF
   }
tr.F1 {
    background-color:#FFCCCC
   }
tr.F0 {
    background-color:#FFDDDD
   }
tr.Spaltenkopf {
    font-weight:bold;background-color:#CCCCCC
   }
td
   { text-align:right; font-size:10pt}
td.Links
   { text-align:left; font-size:10pt}
td.Datum
   { text-align:left;font-size:8pt }
td.Titel {
    font-family:Tahoma,sans-serif; text-align: right; font-size:18pt;
   }

-->
</STYLE>
<HEAD>
<TITLE>Schnellstatistik</TITLE>
  <!--#include virtual="/pwserver/framework/encoding.htm"-->
<LINK REL="stylesheet" TYPE="text/css" HREF="/pwserver/style/mommaster.css">
<LINK REL="stylesheet" TYPE="text/css" HREF="/pwserver/style/mommasterV3.css">
<base target="_self">

<META NAME="author" CONTENT="dzils">
</HEAD>
<BODY style="margin-top:0px" onKeyUp='if(event.keyCode==83){document.getElementById("SQL").style.display="";}';>
<div id="divMain" class="lstDivMain" >
 
<table  width="100%" border="0" class="list" cellpadding="0" cellspacing="0" >
<tr>
<td colspan="7" style="background-color:#ffffff;background-image:url(/pwserver/images/ltr_header_redline_horiz.jpg);background-repeat:repeat-x;">
<div id="headline" class="listHeadline">
		<span class="editObjectType">Schnellstatistik - die letzten 100 Auftr&auml;ge</span>
	</div>
	<br><br>
</td>

</tr>	
</table>
</HEAD>

<form method="POST" action="<%= Request.ServerVariables("PATH_INFO") %>" id=frmMain name=frmMain>
	<DIV id="SQL" style="display:none">
		<textarea name = "SQL" cols="120" rows="3"><%=strSQL%></textarea><br>
		
	</DIV>
	<input type = "submit" value="Neu laden">
</form>

	
<%
  //try
  {
 	var rs = Server.CreateObject ("ADODB.Recordset");
	rs.CursorLocation = 3; // adUseClient
	rs.open(strSQL, DsPcDb);
	var toggle = 1;
	
	if (!(rs.EOF && rs.BOF))
	{
		%><TABLE width=100%><tr class='Spaltenkopf'><%
		for(var i=0;i < rs.Fields.Count;i++)
		{
			Response.Write("<td class='Links'>" + rs.Fields(i).Name + "</td>");
		}
		%><tr><%
		
		while(!rs.EOF)
		{
			toggle = (toggle == 0 ? 1 : 0);
			Response.Write("<tr class = 'S" + toggle + "'>")
			for(var i=0;i < rs.Fields.Count;i++)
			{
				Response.Write("<td class = 'Links'>" + rs.Fields(i) + "</td>");
			}
			
			Response.Write("</tr>");
			
			rs.Movenext();
		}
		
	}
	
  } // Ende Try
//  catch(e)
//  {
//  	Response.Write("Ungültige SQL Abfrage.")
//  }
  
%>

