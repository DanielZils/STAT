<%@ Language=JavaScript %>
<!--#include virtual="/pwserver/framework/ntworb.asp"-->
<!--#include virtual="/pwserver/framework/momasphelper.asp"-->

<% 
// 20150121 DZ added sum row to csv file
// 20141201 DZ added Duplex and Klick columns
// 20141128 DZ added default CC of printer

var CSVSUM = true;
Server.ScriptTimeout=180;
DsPcDb.commandTimeout = 180;
obj = Server.CreateObject("DsPcSrv.DbObjectProxy");

// Check for IPAddress column
var Result = DsPcDb.Execute("SELECT Col_Length('ServiceProvider_T','MgmtData_IPAddress') as MyLen");
Response.Write("Laenge:" + Result.Fields("MyLen"));
var bIP = true;
if (!(Result.Fields("MyLen") > 1))
{
try
{
	var wshshell = Server.CreateObject("Wscript.shell");
	var DBConn = Server.CreateObject("ADODB.Connection");
	try // the database connection string from uniFLOW reg entry
	{	conn= wshshell.regread("HKLM\\Software\\NT-Ware\\MOM\\Connectionstring");
	}	
	catch(e)
	{	conn=wshshell.regread("HKLM\\Software\\WOW6432node\\NT-Ware\\MOM\\Connectionstring");
	}
	DBConn.Connectionstring = conn;
	DBConn.Open();
	DBConn.Execute("USE DSPCDB; IF COL_LENGTH('ServiceProvider_T', 'MgmtData_IPAddress') IS NULL BEGIN ALTER TABLE ServiceProvider_T ADD MgmtData_IPAddress nvarchar(50); END");
	DBConn.Close();
	bIP = true;
}
catch(e)
{ bIP = false;
}
}
// end check new column


//SUMMENFELDER="Name," + (bIP ? "["+ PppString.GetItem(2451) +"]," : "") + "["+ PppString.GetItem(1308) +"], ["+ PppString.GetItem(6232) +"], ["+PppString.GetItem(1901) +"]";
SUMMENFELDER="Name," + (bIP ? "["+ PppString.GetItem(2451) +"]," : "") + "["+ PppString.GetItem(1308) +"], ["+ PppString.GetItem(6232) +"],["+PppString.GetItem(1901) +"]";

var SUMMENZIEL= ""

SUMMENZIEL = SUMMENZIEL + "Max(UsageEnd) AS 'Datum',"
SUMMENZIEL = SUMMENZIEL + "Max(Case WHEN ServiceCode IN (269287425) THEN ANZAHL ELSE 0 END) AS 'Zaehler A4',"
SUMMENZIEL = SUMMENZIEL + "Max(Case WHEN ServiceCode IN (269352961) THEN ANZAHL ELSE 0 END) AS 'Zaehler A4 Farbe',"
SUMMENZIEL = SUMMENZIEL + "Max(Case WHEN ServiceCode IN (269287426) THEN ANZAHL ELSE 0 END) AS 'Zaehler A3',"
SUMMENZIEL = SUMMENZIEL + "Max(Case WHEN ServiceCode IN (269352962) THEN ANZAHL ELSE 0 END) AS 'Zaehler A3 Farbe'"



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


function Right(str, n){
if (n <= 0) return "";
else if (n > String(str).length)
return str;
else {var iLen = String(str).length;
return String(str).substring(iLen, iLen - n);}
}

function GetCC(printerID)
{
	obj.LoadObject(printerID);
	var CCID = obj.GetProperty("DefaultCostCenter");
	if (CCID != "{00000000-0000-0000-0000-000000000000}")
	{	obj.LoadObject(CCID)
		return obj.GetProperty("Name");
	}
	else
		return "-"
		
}

// Übergabe aus Formular vorheriger aufruf, ansonsten Standardwerte

var Today = new Date();
var DateEnd = GetRequestVariable("DateEnd");
if (DateEnd=="")
		DateEnd = Right("0" + "1",2) + "." + Right("0" + (Today.getMonth()+1),2) +"." + Today.getYear();
Today.setMonth(Today.getMonth() - 1);
var DateStart = GetRequestVariable("DateStart");
if (DateStart=="")
		DateStart = Right("0" + "1",2) + "." + Right("0" + (Today.getMonth()+1),2) +"." + Today.getYear();

var Operator = GetRequestVariable("Operator");
if (Operator == "") Operator = "LIKE";

// Datumsfilter für Abfrage zusammenbauen
	strDatumFilter = "WHERE Serviceusage_t.UsageEnd < Convert(Datetime,'" + DateEnd + " 00:00:00',104) ";
	strDatumFilter = strDatumFilter + "AND Serviceusage_t.UsageEnd >= Convert(Datetime,'" + DateStart + " 00:00:00',104) AND ServiceCode IN (269352961,269287426,269352962,269287425) ";

// Wenn Filter gesetzt, dann Filterstring zusammenbauen
var strFilter = "";
if (GetRequestVariable("FilterFieldName") != "")
{
	strFilter = " WHERE [" + GetRequestVariable("FilterFieldName") + "] " + Operator + " ";
	if (Operator == "LIKE" || Operator =="NOT LIKE") {
		strFilter = strFilter + "'%" + GetRequestVariable("FilterString") + "%'"; }
	else {
		strFilter = strFilter + "'" + GetRequestVariable("FilterString") + "'"; }
}

// the file name:
	var theExportFileName = "Drucker_"+DateStart+"-"+DateEnd+"_" + Request.ServerVariables("LOGON_USER");
	theExportFileName = theExportFileName.replace ("\\","_");

// Die letzte SQL Abfrage behalten
var strSQL = GetRequestVariable("SQL");

// falls es keine letzte Abfrage gab (erster Aufruf)
   if (strSQL == "" )
   {
   	
	 	strSQL = "SELECT UsageEnd, Service_T.Name AS Service, Service_T.ServiceCode,  ServiceProvider_T.Name AS Name, ServiceProvider_T.MgmtData_Serial AS ["+ PppString.GetItem(1308) +"],"+ (!bIP ? "": "ISNULL(ServiceProvider_T.MgmtData_IPAddress,ServiceProvider_T.MgmtData_Hostname) AS ["+ PppString.GetItem(2451) +"],") +" ServiceProvider_T.MgmtData_Location AS ["+ PppString.GetItem(6232) +"], ServiceProvider_T.ID AS ["+ PppString.GetItem(1901) +"],"  + 
                      " CONVERT(varchar,ServiceUsage_T.UsageEnd,20) AS UsageEnd1, " + 
                      " ServiceUsage_T.Cardinality AS Anzahl " +
                      "FROM ServiceUsage_T  RIGHT JOIN " +
                      "ServiceProvider_T ON ServiceUsage_T.ServiceProvider = ServiceProvider_T.ID  RIGHT OUTER JOIN " +
                       "Service_T ON ServiceUsage_T.Service = Service_T.ID ";
                      
	//	"ORDER BY ServiceUsage_T.UsageEnd DESC";

   }

if (SUMMENFELDER != "")
{
  var strSQL2 = "SELECT "+ SUMMENFELDER + ", " + SUMMENZIEL + " FROM (" + strSQL + strDatumFilter + ") AS Tabelle " + strFilter + 
                " GROUP BY " + SUMMENFELDER + " ORDER BY " + SUMMENFELDER;
}  
else
{
  var strSQL2 = "SELECT * FROM (" + strSQL + strDatumFilter + ") AS Tabelle " + strFilter;
}
%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 //EN">
<HTML>
<STYLE TYPE="text/css">
<!--
a  {
    font-family:Tahoma,sans-serif; text-align: right;
   }

tr {
    font-family:Tahoma,sans-serif; text-align: right;
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
   { text-align:right; font-size:70%}
td.Links
   { text-align:left; font-size:70%}
td.Datum
   { text-align:left;font-size:70% }
td.Titel {
    font-family:Tahoma,sans-serif; text-align: left; font-size:18pt;
   }

-->
</STYLE>
<HEAD>
<TITLE><%=PppString.GetItem(3612)%></TITLE>
  <!--#include virtual="/pwserver/framework/encoding.htm"-->
<LINK REL="stylesheet" TYPE="text/css" HREF="/pwserver/style/mommaster.css">
<LINK REL="stylesheet" TYPE="text/css" HREF="/pwserver/style/mommasterV3.css">
<base target="_self">

<META NAME="author" CONTENT="dzils">
</HEAD>
<BODY style="margin-top:0px">
<div id="divMain" class="lstDivMain" >
 
<table  width="100%" border="0" class="list" cellpadding="0" cellspacing="0" >
<tr>
<td colspan="7" style="background-color:#ffffff;background-image:url(/pwserver/images/ltr_header_redline_horiz.jpg);background-repeat:repeat-x;">
<div id="headline" class="listHeadline">
		<span class="editObjectType"><%=PppString.GetItem(3612)%></span>
	</div>
	<br><br><br>
</td>

</tr>	
</table>
</HEAD>
<BODY>
<TABLE >
<tr><td>
<form method="POST" action="<%= Request.ServerVariables("PATH_INFO") %>" id=frmMain name=frmMain>
	<textarea style="display:none" name = "SQL" cols="120" rows="3"><%=strSQL%></textarea><br>

	<%=PppString.GetItem(11623)%>:<SELECT name=FilterFieldName id=FilterFieldName>
	<option value =""></option>
	</SELECT>
	<SELECT name=Operator>
		<OPTION value ="LIKE" <%=Operator=="LIKE" ? "SELECTED" : ""%>>LIKE</option>
		<OPTION value ="NOT LIKE" <%=Operator=="NOT LIKE" ? "SELECTED" : ""%>>NOT LIKE</option>
		<OPTION value ="=" <%=Operator=="=" ? "SELECTED" : ""%>>=</option>
		<OPTION value ="<>" <%=Operator=="<>" ? "SELECTED" : ""%>>!=</option>
	</SELECT>
	<input type = "text" value = "<%=GetRequestVariable("FilterString")%>" name = FilterString id =FilterString>
	<%=PppString.GetItem(7518)%>: <input type = "text" value = <%=DateStart%> name = DateStart id=DateStart size="10" maxlength="10">(00:00:00)
	<%=PppString.GetItem(7519)%>:<input type = "text" value = <%=DateEnd%> name = DateEnd id=DateEnd size="10" maxlength="10">(00:00:00)
	
	<input type = "submit" value="<%=PppString.GetItem(45032)%>">
	<a href="<%="/pwserver/xml/" + theExportFileName + ".csv"%>" target="NeuesFenster">CSV Download</a>
	<input type = "button" value = "SQL" onclick="ShowSQL();">
	
</form>
</td></TR>
</TABLE>
<script language='JavaScript'>
function ShowSQL() {	
	SQLFenster = window.open("","SQL", "scrollbars=1,width=500,height=255");
	SQLFenster.document.write("<html><%=strSQL2%></html>");
	SQLFenster.document.close();}
	
<%
  try
  {
 	var rs = Server.CreateObject ("ADODB.Recordset");
	rs.CursorLocation = 3; // adUseClient
	rs.ActiveConnection = DsPcDb;
	rs.activeConnection.CommandTimeout = 300;
	Response.Write(strSQL2);
	
	rs.open(strSQL2);
	var toggle = 1;
	
	// ### Mögliche Filter sind... ###
	
	for(var i=0;i < 2;i++)
		{
			Response.Write("document.forms['frmMain'].FilterFieldName.options[" + (i + 1) + "] = " +
			"new Option('" + rs.Fields(i).Name + "','" + rs.fields(i).Name + "');\r\n");			
			if (rs.Fields(i).Name == GetRequestVariable("FilterFieldName"))
			{ Response.Write("document.forms['frmMain'].FilterFieldName.selectedIndex =" + (i +1) + ";"); }
		}
%>
</script>
<%		
	var summe = new Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
	var bNoData = rs.EOF;
	var CSV = "";
	var LfdNr = 1;
	
	if (!(rs.EOF && rs.BOF))
	{
		%><TABLE width=100%><tr class='Spaltenkopf'><%
		CSV = CSV + "\"LfdNr\";";
		Response.Write("<td class = 'Links'>LfdNr</td>");
		for(var i=0;i < rs.Fields.Count;i++)
		{
			Response.Write("<td class='Links'>" + rs.Fields(i).Name + "</td>");			
			CSV = CSV + "\"" + rs.Fields(i).Name + "\";";
		}
		%></tr><%
		CSV = CSV + "\r\n";
		var MyValue = "";
		while(!rs.EOF)
		{
			CSV = CSV + "\"" + LfdNr + "\";";
			toggle = (toggle == 0 ? 1 : 0);
			Response.Write("<tr class = 'S" + toggle + "'>")
			Response.Write("<td class = 'Links'>" + LfdNr + "</td>");
			for(var i=0;i < rs.Fields.Count;i++)
			{
				if (rs.Fields(i).name == PppString.GetItem(1901))
				{
					MyValue = GetCC(rs.Fields(i));
					Response.Write("<td class = 'Links'>" + MyValue + "</td>");
					// summe[i] += rs.Fields(i);
					CSV = CSV + "\"" + MyValue + "\";";
					
				}
				else
				{
					Response.Write("<td class = 'Links'>" + rs.Fields(i) + "</td>");
					summe[i] += rs.Fields(i);
					CSV = CSV + "\"" + rs.Fields(i) + "\";";
				}
			}
			
			Response.Write("</tr>");
			CSV = CSV + "\r\n";
			LfdNr++;
			
			rs.Movenext();
		}
		
	}
	if(bNoData) 
	{
		Response.Write("<tr class='Spaltenkopf'>");
		Response.Write("<td class='Links'>Sorry no data.</td>");
		Response.Write("<td class='Links'></td>");
		Response.Write("</tr></table>");
	}
	else
	{
		Response.Write("<tr class='Spaltenkopf'>");
		Response.Write("<td class='Links'>"+PppString.GetItem(120004)+"</td>");
		if(bIP) Response.Write("<td class='Links'></td>");
		Response.Write("<td class='Links'></td>");
		Response.Write("<td class='Links'></td>");
		Response.Write("<td class='Links'></td>");
		Response.Write("<td class='Links'></td>");

		if (CSVSUM) CSV=CSV + "SUM;;;;;;";
		for (i=(bIP ? 5 : 4);i<rs.Fields.Count;i++)
		{
			Response.Write("<td class='Links'>" + summe[i] + "</td>");
			if (CSVSUM) CSV=CSV + summe[i] + ";";
		}
		Response.Write("</tr></table>");
	}
	
	var objFs= Server.CreateObject("Scripting.FileSystemObject");
	var strRedirectURL = "/pwserver/xml/" + theExportFileName +".CSV";
	var strFilePath = Server.MapPath(strRedirectURL);
	var objTextStream = objFs.CreateTextFile(strFilePath, true);
	objTextStream.Write(CSV);
	objTextStream.Close();

	if(GetRequestVariable("Format") == "XLS")
	{
		Response.Redirect("./statgetfile.asp");
	}
	
  } // Ende Try
  catch(e)
  {
  	Response.Write(e.Description);
  	//throw(e);
  }
  
%>

