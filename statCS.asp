<%@ Language=JavaScript %>
<!--#include virtual="/pwserver/framework/ntworb.asp"-->
<!--#include virtual="/pwserver/framework/momasphelper.asp"-->


<% 
Server.ScriptTimeout=180;
DsPcDb.commandTimeout = 180;

var COSTCENTERSPLIT = "";

//var SUMMENFELDER = "";
//var SUMMENZIEL = "";

SUMMENFELDER="[" + PppString.GetItem(2505) +"], [" + PppString.GetItem(30212) +"]";
SUMMENZIEL = "sum(Pages) AS [" + PppString.GetItem(51062) +"],sum(PagesColor) as [" + PppString.GetItem(14064) +"],sum(CostsSaved) AS [" + PppString.GetItem(1228) +"]";


//[" + PppString.GetItem(1228) +"]


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
if (n <= 0)
return "";
else if (n > String(str).length)
return str;
else {
var iLen = String(str).length;
return String(str).substring(iLen, iLen - n);
}
}
function runde(x, n) {
  if (n < 1 || n > 14) return false;
  var e = Math.pow(10, n);
  var k = (Math.round(x * e) / e).toString();
  if (k.indexOf('.') == -1) k += '.';
  k += e.toString().substring(1);
  return k.substring(0, k.indexOf('.') + n+1);
}

// Übergabe aus Formular vorheriger aufruf, ansonsten Standardwerte
var strFilter = "";

var Today = new Date();
var DateEnd = GetRequestVariable("DateEnd");
	if (DateEnd=="")
		DateEnd = Right("0" + "1",2) + "." + Right("0" + (Today.getMonth()+1),2) +"." + Today.getYear()
Today.setMonth(Today.getMonth() - 3);
var DateStart = GetRequestVariable("DateStart");
if (DateStart=="")
		DateStart = Right("0" + "1",2) + "." + Right("0" + (Today.getMonth()+1),2) +"." + Today.getYear()

var Operator = GetRequestVariable("Operator");
if (Operator == "") Operator = "LIKE";

// Datumsfilter für Abfrage zusammenbauen
	strDatumFilter = "WHERE UsageEnd < Convert(Datetime,'" + DateEnd + " 00:00:00',104) ";
	strDatumFilter = strDatumFilter + "AND UsageEnd >= Convert(Datetime,'" + DateStart + " 00:00:00',104) ";

// Wenn Filter gesetzt, dann Filterstring zusammenbauen
if (GetRequestVariable("FilterFieldName") != "")
{
	strFilter = " WHERE " + GetRequestVariable("FilterFieldName") + " " + Operator + " ";
	if (Operator == "LIKE" || Operator =="NOT LIKE") {
		strFilter = strFilter + "'%" + GetRequestVariable("FilterString") + "%'"; }
	else {
		strFilter = strFilter + "'" + GetRequestVariable("FilterString") + "'"; }
}

// the file name:
	var theExportFileName = "Costsavings_"+DateStart+"-"+DateEnd+"_" + Request.ServerVariables("LOGON_USER");
	theExportFileName = theExportFileName.replace ("\\","_");
	

// Die letzte SQL Abfrage behalten
var strSQL = GetRequestVariable("SQL");

// falls es keine letzte Abfrage gab (erster Aufruf)
   if (strSQL == "" )
   {
   	// 64436 = 30800 = Deleted (manual)
	// 131072 = 30801 = Deleted (maintenance)
	// 262144 = 30802 = Deleted (license exceeded)
	// 524288 = 30803 = Deleted (unknown user)
	// 1048576 = 30804 = Deleted (ACL restriction)
	// 2097152 = 30805 = Deleted (budget exceeded)
	// 4194304 = 30806 = Deleted (product avail. restriction)
	// > 30807 = Deleted (unknown reason)
	
	strSQL = "SELECT ServiceProvider_T.Name AS [" + PppString.GetItem(2505) +"]," + 
			 "CASE WHEN Type = 65536 THEN '" + PppString.GetItem(30800) + 
			 "' ELSE CASE WHEN Type = 131072 THEN '" + PppString.GetItem(30801) + 
			 "' ELSE CASE WHEN Type = 262144 THEN '" + PppString.GetItem(30802) + 
			 "' ELSE CASE WHEN Type = 524288 THEN '" + PppString.GetItem(30803) + 
			 "' ELSE CASE WHEN Type = 1048576 THEN '" + PppString.GetItem(30804) + 
			 "' ELSE CASE WHEN Type = 2097152 THEN '" + PppString.GetItem(30805) + 
			 "' ELSE CASE WHEN Type = 4194304 THEN '" + PppString.GetItem(30806) + 
			 "' ELSE  CAST (Type as nvarchar) END END END END END END END AS [" + PppString.GetItem(30212) +"],"+
			 "TYPE, Pages as Pages,  PagesColor as PagesColor, CostsSaved as CostsSaved " +
			 "FROM CostSavings_T Left JOIN ServiceProvider_T ON CostSavings_T.ServiceProvider = ServiceProvider_T.ID " +
			 "";

   }

if (SUMMENFELDER != "")
{
  var strSQL2 = "SELECT "+ SUMMENFELDER + ", " + SUMMENZIEL + " FROM (" + strSQL + strDatumFilter + " AND (Pages > 0 OR PagesColor > 0) AND Type > 1) AS Tabelle " + strFilter + 
                " GROUP BY " + SUMMENFELDER + " ORDER BY " + SUMMENFELDER;
}  
else
{
  var strSQL2 = "SELECT * FROM (" + strSQL + ") AS Tabelle1 " + strFilter ;
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
<TITLE><%=PppString.GetItem(30809)%></TITLE>
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
		<span class="editObjectType"><%=PppString.GetItem(30809)%></span>
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
	rs.CursorType = 0;    // forwardOnly
	rs.LockType = 1;     // readOnly
	rs.ActiveConnection = DsPcDb;
	rs.activeConnection.CommandTimeout = 300;
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
	var summe = new Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
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
		var MyValue;
		while(!rs.EOF)
		{
			
			CSV = CSV + "\"" + LfdNr + "\";";
			toggle = (toggle == 0 ? 1 : 0);
			Response.Write("<tr class = 'S" + toggle + "'>")
			Response.Write("<td class = 'Links'>" + LfdNr + "</td>");
			for(var i=0;i < rs.Fields.Count;i++)
			{
				if (rs.fields(i).Name == PppString.GetItem(1228)) 
					MyValue = runde(rs.fields(i),2);
				else
					MyValue = rs.fields(i);
					
				Response.Write("<td class = 'Links'>" + MyValue + "</td>");
				CSV = CSV + "\"" + MyValue + "\";";
				if (i> 1) summe[i] += rs.fields(i);
				
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
		Response.Write("<td class='Links'>Sorry, no data...</td>");
	}
	else
	{
		Response.Write("<tr class='Spaltenkopf'>");
		Response.Write("<td class='Links'>"+PppString.GetItem(120004)+"</td>");
		Response.Write("<td class='Links'></td>");
		Response.Write("<td class='Links'></td>");
		for (i=2;i<rs.Fields.Count;i++)
		{
			if (rs.fields(i).Name == PppString.GetItem(1228))
				Response.Write("<td class='Links'>" + runde(summe[i],2) + "</td>");
			else
				Response.Write("<td class='Links'>" + summe[i] + "</td>");
		}
	
	}
	Response.Write("</tr></table>");
	
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
  	Response.Write(e.Description)
  }
  
%>

