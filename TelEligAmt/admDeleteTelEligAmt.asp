<%@ Language=VBScript %>
<%  
Option Explicit
On Error Resume Next
Response.Expires = -1 %>

<!--#include file="../../Common/Check.asp" -->
<!--#include file="../../Common/Connection.asp" -->
<!-- #include file="../../Common/commonFunctions.asp"-->

<HTML>
<HEAD>
<TITLE>ECGC: Telephone Eligibilty Amount Master - Delete</TITLE>
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
   Dim Conn , rs
   Dim SQL
   Dim strRankCd , strFinYr , strAmount
   
   if Request.Form("btnSubmit") = "Delete" then
    	
		set Conn = Server.CreateObject( "ADODB.Connection" )
		Conn.Open astrConn	
		if Err.number <> 0 then
			 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-2&astrErrDescription=" & Err.description
		     Response.End  
		end if
		SQL ="Delete from  ADM_TEL_AMT_ELIGIB_MST where RANK_CD = '" & Request.Form("Param2") & "' and FINANCIAL_YEAR = '" & Request.Form("Param1") & "'"
		
		Conn.Execute SQL
       	Conn.close
       	set Conn = nothing
	    Response.Redirect "../../Common/Success.asp?aintCode=1&astrPage=../ADM/TelEligAmt/admIntTelEligAmt.asp?Param=D"
	    
	end if
    
	if Request.Form("btnSubmit") = "Next" then
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
		
	end if

	%>



<FORM name=frmMain method=post>

<TABLE border=0 width="100%">
   <TR>
    <TD class="clsHeader">Telephone Eligibilty Amount Master - Delete</TD>
   </TR>
</TABLE>
<BR>
<TABLE width="20%" class="clsTabHead" cellpadding=0 cellspacing=0 border=0 >
	<TR>
	<TH  class="clsActive"   width="100%" >
		Telephone Eligibilty Amount Master
	</TH>
	</TR>
</TABLE>

<DIV ID="divTab1" name="divTab1" class="clsTabBody">
<TABLE border=0 width="100%">
<input type="hidden" name="Param1" value="<%=strFinYr%>">
<input type="hidden" name="Param2" value="<%=strRankCd%>">
 <TR>
 <TD width="20%"></TD>
 <TD width="30%"></TD>
 <TD width="20%"></TD>
 <TD width="30%"></TD>
 </TR>
 
<TR>
   <TD><FONT class="clsLabel">Financial Year</FONT></TD>
   <td width="30%"><FONT class="clsText"><%=strFinYr%></FONT></TD>
   <TD width="20%"></TD>
   <TD width="30%"></TD>
</TR>
<tr>
	<td width="20%"><font class="clsLabel">Rank Code</font></td>
	<td width="30%" colSpan="3"><FONT class="clsText"><%=strRankCd%></FONT></TD>
		
</tr>
<tr>
  <td width="20%"><font class="clsLabel">Eligible Amount</font></td>
  <td width="30%"><FONT class="clsText"><%=strAmount%></FONT></TD>
  <td width="20%"></td>
  <td width="30%" ></TD>
  
</tr>

</TABLE>
</DIV>
<BR>
<CENTER>
<INPUT TYPE=submit name=btnSubmit value="Delete" class=clsButton>&nbsp;&nbsp;&nbsp;
<INPUT TYPE="button" name="btnCancel"  value="Cancel" onClick="history.go(-1)" class="clsButton" >
</CENTER>
</FORM>
</BODY>
</HTML>

