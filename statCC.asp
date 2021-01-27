<%@ Language=JavaScript %>
<!--#include virtual="/pwserver/framework/ntworb.asp"-->
<!--#include virtual="/pwserver/framework/momasphelper.asp"-->


<% 
Server.ScriptTimeout=180;
DsPcDb.commandTimeout = 180;

AMOUNTCOLUMNS = 2; // Preisspalten ganz rechts

// *** STRUKTUR=true: "Firma, Bereich, Kostenstelle", Ansonsten: "Kostenstelle, Abteilung" ***
STRUKTUR = false;

// *** Kostenstelle besteht aus mehreren Teilen: Trennzeichen hier angeben ***
if (STRUKTUR) {var COSTCENTERSPLIT = ".";}
else {var COSTCENTERSPLIT = "";} 

if (STRUKTUR) {var SUMMENFELDER="Firma, Bereich, ["+ PppString.GetItem(4676) +"]";}
else {SUMMENFELDER="["+ PppString.GetItem(4676) +"],["+ PppString.GetItem(7863) +"]";} 

var SUMMENZIEL= "Sum(Case WHEN ServiceCode IN 			(196609,196611,196613,196614,196619,196620,196621,196622,196623) THEN ANZAHL ELSE 0 END) AS 'Druck A4 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (196610,196612,196615,196617,196618) THEN ANZAHL ELSE 0 END) AS 'Druck A3 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (262145,262147,262149,262150,262155,262156,262157,262158,262159) THEN ANZAHL ELSE 0 END) AS 'Druck A4 C',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (262146,262148,262151,262153,262154) THEN ANZAHL ELSE 0 END) AS 'Druck A3 C',"

SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (65537,65539,65541,65542,65547,65548,65549,65550,65551) THEN ANZAHL ELSE 0 END) AS 'Kopie A4 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (65538,65540,65543,65545,65546) THEN ANZAHL ELSE 0 END) AS 'Kopie A3 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (131073,131075,131077,131078,131083,131084,131085,131086,131087) THEN ANZAHL ELSE 0 END) AS 'Kopie A4 C',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (131074,131076,131079,131081,131082) THEN ANZAHL ELSE 0 END) AS 'Kopie A3 C',"

/*
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (393217,393219,393221,393222,393227,393228,393229,393230,393221) THEN ANZAHL ELSE 0 END) AS 'Fax A4 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (393218,393220,393223,393225,393226) THEN ANZAHL ELSE 0 END) AS 'Fax A3 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (0) THEN ANZAHL ELSE 0 END) AS 'Fax A4 C',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (0) THEN ANZAHL ELSE 0 END) AS 'Fax A3 C',"
*/

/*
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (327681,327683,327685,327686,327686,327691,327692,327693,327694,327695) THEN ANZAHL ELSE 0 END) AS 'Scan A4 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (327682,327684,327687,327689,327690) THEN ANZAHL ELSE 0 END) AS 'Scan A3 SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (589825,589827,589829,589830,589835,589836,589837,589838,589839) THEN ANZAHL ELSE 0 END) AS 'Scan A4 C',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (589826,589828,589831,589833,589834) THEN ANZAHL ELSE 0 END) AS 'Scan A3 C',"
*/

SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode = 1048578 THEN ANZAHL ELSE 0 END) AS 'Duplex',"

// Klick SW
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (196610,196612,196615,196617,196618) THEN ANZAHL ELSE 0 END) * 2 + Sum(Case WHEN ServiceCode IN 			(196609,196611,196613,196614,196619,196620,196621,196622,196623) THEN ANZAHL ELSE 0 END) + Sum(Case WHEN ServiceCode IN (65537,65539,65541,65542,65547,65548,65549,65550,65551) THEN ANZAHL ELSE 0 END) + Sum(Case WHEN ServiceCode IN (65538,65540,65543,65545,65546) THEN ANZAHL ELSE 0 END) * 2 AS 'Klick SW',"

SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (262146,262148,262151,262153,262154) THEN ANZAHL ELSE 0 END) * 2 + Sum(Case WHEN ServiceCode IN 			(262145,262147,262149,262150,262155,262156,262157,262158,262159) THEN ANZAHL ELSE 0 END) + Sum(Case WHEN ServiceCode IN (131073,131075,131077,131078,131083,131084,131085,131086,131087) THEN ANZAHL ELSE 0 END) + Sum(Case WHEN ServiceCode IN (131074,131076,131079,131081,131082) THEN ANZAHL ELSE 0 END) * 2 AS 'Klick Color',"



SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (196609,196611,196613,196614,196619,196620,196621,196622,196623,196610,196612,196615,196617,196618,65537,65539,65541,65542,65547,65548,65549,65550,65551,65538,65540,65543,65545,65546) THEN BETRAG ELSE 0 END) AS 'Preis SW',"
SUMMENZIEL = SUMMENZIEL + "Sum(Case WHEN ServiceCode IN (262145,262147,262149,262150,262155,262156,262157,262158,262159,262146,262148,262151,262153,262154,131073,131075,131077,131078,131083,131084,131085,131086,131087,131074,131076,131079,131081,131082) THEN BETRAG ELSE 0 END) AS 'Preis C'"


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

function runde(x, n) {
  if (n < 1 || n > 14) return false;
  var e = Math.pow(10, n);
  var k = (Math.round(x * e) / e).toString();
  if (k.indexOf('.') == -1) k += '.';
  k += e.toString().substring(1);
  return k.substring(0, k.indexOf('.') + n+1).replace(".",",");
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

// Übergabe aus Formular vorheriger aufruf, ansonsten Standardwerte
var strFilter = "";

var Today = new Date();
var DateEnd = GetRequestVariable("DateEnd");
	if (DateEnd=="")
		DateEnd = Right("0" + "1",2) + "." + Right("0" + (Today.getMonth()+1),2) +"." + Today.getYear()
Today.setMonth(Today.getMonth() - 1);
var DateStart = GetRequestVariable("DateStart");
if (DateStart=="")
		DateStart = Right("0" + "1",2) + "." + Right("0" + (Today.getMonth()+1),2) +"." + Today.getYear()
var Operator = GetRequestVariable("Operator");
if (Operator == "") Operator = "LIKE";

// Datumsfilter für Abfrage zusammenbauen
	strDatumFilter = "WHERE Serviceusage_t.UsageEnd < Convert(Datetime,'" + DateEnd + " 00:00:00',104) ";
	strDatumFilter = strDatumFilter + "AND Serviceusage_t.UsageEnd >= Convert(Datetime,'" + DateStart + " 00:00:00',104) ";

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
	var theExportFileName = "KstStellen_"+DateStart+"-"+DateEnd+"_" + Request.ServerVariables("LOGON_USER");
	theExportFileName = theExportFileName.replace ("\\","_");
	

// Die letzte SQL Abfrage behalten
var strSQL = GetRequestVariable("SQL");

// falls es keine letzte Abfrage gab (erster Aufruf)
   if (strSQL == "" )
   {
   	if (COSTCENTERSPLIT != "") {
   	strKostenstelle = " CASE WHEN charindex('.',name)> 0 THEN left(name,charindex('.',name) - 1) ELSE '-' END " +
			"AS Firma, " +
			"CASE WHEN charindex('.',name)> 0 AND charindex('.',name,charindex('.',name)+1) > 0 " +
			"THEN substring(name,charindex('.',name)+1,charindex('.',name,charindex('.',name)+1) - charindex('.',name)-1) " +
			" ELSE '-' END AS Bereich, " +
			"CASE WHEN charindex('.',name)> 0 AND charindex('.',name,charindex('.',name)+1) > 0 " +
			"THEN substring(name,charindex('.',name,charindex('.',name)+1)+1,1000) ELSE name END AS ["+ PppString.GetItem(4676) +"] ";
	
	strKostenstelle = strKostenstelle.replace(/'\.'/g,"'" + COSTCENTERSPLIT + "'");
	
	}
	else
	{ strKostenstelle = "name AS ["+ PppString.GetItem(4676) +"]"; }
	
	strKostenstelle = strKostenstelle.replace(/name/g,"ServiceConsumer_T_1.Name") + ", ServiceConsumer_T_1.AddressOne AS ["+PppString.GetItem(7863)+"] ";
	
   	strSQL = "SELECT " + strKostenstelle + ", Service_T.Name AS Service, Service_T.ServiceCode,  ServiceProvider_T.Name AS Printer," + 
                      " ServiceConsumer_T.Name AS Username,  CONVERT(varchar,ServiceUsage_T.UsageEnd,20) AS UsageEnd, " + 
                      " ServiceUsage_T.Cardinality AS Anzahl,  ServiceUsage_T.AmountPaid AS Betrag " +
                      "FROM ServiceUsage_T  LEFT JOIN " +
                      "ServiceProvider_T ON ServiceUsage_T.ServiceProvider = ServiceProvider_T.ID  LEFT JOIN " +
                      "ServiceConsumer_T ServiceConsumer_T_1 ON ServiceUsage_T.ServConsProject = ServiceConsumer_T_1.ID  INNER JOIN " +
                      "Service_T ON ServiceUsage_T.Service = Service_T.ID  LEFT JOIN " +
                      "ServiceConsumer_T ON ServiceUsage_T.ServiceConsumer = ServiceConsumer_T.ID ";
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
<TITLE><%=PppString.GetItem(3610)%></TITLE>
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
		<span class="editObjectType"><%=PppString.GetItem(3610)%></span>
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
	rs.CursorType = 0; // forwardOnly
	rs.LockType = 1; // readOnly
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
	var summe = new Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
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
		
		while(!rs.EOF)
		{
			CSV = CSV + "\"" + LfdNr + "\";";
			toggle = (toggle == 0 ? 1 : 0);
			Response.Write("<tr class = 'S" + toggle + "'>")
			Response.Write("<td class = 'Links'>" + LfdNr + "</td>");
			for(var i=0;i < rs.Fields.Count;i++)
			{
				//if ((!STRUKTUR && i > 17) || (STRUKTUR && i > 18))
				if (i > rs.Fields.Count - AMOUNTCOLUMNS -1)
				{
					Response.Write("<td class = 'Links'>" + runde(rs.Fields(i),2) + "</td>");
					CSV = CSV + "\"" + runde(rs.Fields(i),2) + "\";";
					if (i> 1) summe[i] += rs.Fields(i);
				}
				else
				{
					
					Response.Write("<td class = 'Links'>" + rs.Fields(i) + "</td>");
					CSV = CSV + "\"" + rs.Fields(i) + "\";";
					//if (i> 2 && STRUKTUR) summe[i] += rs.Fields(i);
					if (i> 1) summe[i] += rs.Fields(i);
					//summe[i] += rs.Fields(i);
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
		Response.Write("<td class='Links'>Sorry, no data...</td>");
	}
	else
	{
		Response.Write("<tr class='Spaltenkopf'>");
		Response.Write("<td class='Links'>"+PppString.GetItem(120004)+"</td>");
		//Response.Write("<td class='Links'></td>");
		Response.Write("<td class='Links'></td>");
		if (STRUKTUR)
		{ Response.Write("<td></td>");
			Response.Write("<td></td>");
			for (i=3;i<rs.Fields.Count;i++)
			{
				if (i > rs.fields.count - AMOUNTCOLUMNS -1)
					Response.Write("<td class='Links'>" + runde(summe[i],2) + "</td>");
				else
					Response.Write("<td class='Links'>" + summe[i] + "</td>");
			}
		}
		else
		{
			Response.Write("<td></td>");
			for (i=2;i<rs.Fields.Count;i++)
			{
				if (i > rs.fields.count - AMOUNTCOLUMNS -1)
					Response.Write("<td class='Links'>" + runde(summe[i],2) + "</td>");
				else
					Response.Write("<td class='Links'>" + summe[i] + "</td>");
			}
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

