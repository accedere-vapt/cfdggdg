<%@ Language=VBScript %>
<%  Option Explicit
	On Error Resume Next
	Response.Expires = -1  

'****************************************************************************
' Application Name		: View Telephone Eligibilty Amount Master
' Author Name			: Janani Seshadri
' Date of Creation		: 10-June-2002
' Version Number		: 1.0
' Purpose				: This is the ASP for Telephone Eligibilty Amount Master View
' Remarks				: 
'**************************************************************************** 

%>
<!--#include file="../../Common/Check.asp" -->
<!--#include file="../../Common/Connection.asp" -->
<!-- #include file="../../Common/commonFunctions.asp"-->

<HTML>
<HEAD>
<TITLE>ECGC: Telephone Eligibilty Amount Master - View</TITLE>
<link rel="stylesheet" type="text/css" href="../../stylesheets/ecgc.css">
<SCRIPT language="JavaScript" src="../../includes/incSearch.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incValidationFunctions.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incCommonFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--

//-->
</SCRIPT>
</HEAD>
<BODY>

<%	
    
	'///////////////////////////////////////////////////////////////////////////
	' This part of the code gets executed when the form gets loaded
	'///////////////////////////////////////////////////////////////////////////
	
	Dim Conn, rs, cmd, param
	Dim SQL
	Dim strRankCd , strFinYr , strAmount
	
    set Conn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )
	Conn.Open astrConn	
	if Err.number <> 0 then
		 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	     Response.End  
	end if	 
			
    SQL ="select * from  ADM_TEL_AMT_ELIGIB_MST where upper(RANK_CD) = '" & ucase(Request.Form("txtRankId")) & "' and FINANCIAL_YEAR = '" & Request.Form("txtFinYr") & "'"
    rs.Open SQL,Conn
         	
	if not rs.EOF then
	       strRankCd = UCASE(Request.Form("txtRankId"))
	       strFinYr = Request.Form("txtFinYr")
	       strAmount = rs.Fields("ELIGIBLE_AMT")
	       rs.Close
		   Conn.close
		   set rs = nothing
		   set Conn = nothing
		   
	else 
	       rs.Close
		   Conn.close
	       set rs = nothing
		   set Conn = nothing
	       Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
	end if
	
	%>


<FORM name=frmMain method=post>

<TABLE border=0 width="100%">
   <TR>
    <TD class="clsHeader">Telephone Eligibilty Amount Master - View</TD>
   </TR>
</TABLE>
<BR>
<TABLE width="20%" class="clsTabHead" cellpadding=0 cellspacing=0 border=0 >
	<TR>
	<TH  class="clsActive" width="100%" >
		Telephone Eligibilty Amount 
	</TH>
	</TR>
</TABLE>

<DIV ID="divTab1" name="divTab1" class="clsTabBody">
<TABLE border=0 width="100%">
  
 <TR>
 <TD width="20%"></TD>
 <TD width="30%"></TD>
 <TD width="20%"></TD>
 <TD width="30%"></TD>
 </TR>
 
<tr>
	<td width="20%"><font class="clsLabel">Financial Year</font></td>
	<td width="30%" colSpan="3"><FONT class="clsText"><%=strFinYr%></FONT></TD>
	
</tr>


<tr>
	<td width="20%"><font class="clsLabel">Rank Code</font></td>
	<td width="30%" colSpan="3"><FONT class="clsText"><%=strRankCd%></FONT></TD>
		
</tr>


<tr>
  <td width="20%"><font class="clsLabel">Eligible Amount</font></td>
  <td width="30%" colSpan="3"><FONT class="clsText"><%=strAmount%></FONT></TD>
  
</tr>

<tr>
  <td colspan="4">&nbsp;</td>
</tr>

</TABLE>
</DIV>
<BR>
<CENTER>
<INPUT TYPE="button" name="btnCancel"  value="OK" onClick="history.go(-1)" class="clsButton" >
</CENTER>
</FORM>
</BODY>
</HTML>

