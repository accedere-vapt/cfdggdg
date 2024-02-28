<%@ Language=VBScript %>
<%  
Option Explicit
On Error Resume Next
Response.Expires = -1 

%>
<!--#include file="../../Common/Check.asp" -->
<!--#include file="../../Common/Connection.asp" -->
<!-- #include file="../../Common/commonFunctions.asp"-->

<HTML>
<HEAD>
<TITLE>ECGC: Telephone Eligibilty Amount Master - Modify</TITLE>
<link rel="stylesheet" type="text/css" href="../../stylesheets/ecgc.css">
<SCRIPT language="JavaScript" src="../../includes/incSearch.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incValidationFunctions.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incCommonFunctions.js"></SCRIPT>

<SCRIPT LANGUAGE=javascript>
<!--

//Validates the updated fields
function fnSubmit()
{
	var frmForm = document.frmMain;
	var txtTextbox;
	var selSelect;
	var strFieldName;
	var intMandatory;	
	var radRadio;
	
	
	/////////////////////////
	// Validate Amount
	/////////////////////////
	txtTextbox = frmForm.txtEligAmt;
	strFieldName = "Amount";
	strAmountField = "0";
	intMandatory = 1;
	
	// Should be valid
	if(gf_bValidateAmount(txtTextbox, intMandatory, strAmountField, strFieldName) == false)
	{
		return false;
	}
	
	
	return true;
}


//this function is executed when the form is loaded
function fnOnLoad()
{
	document.frmMain.txtFinYr.focus();
	
}



//-->
</SCRIPT>
</HEAD>
<BODY onLoad="fnOnLoad();">

<%
	'///////////////////////////////////////////////////////////////////////////////////
	' Declarations
	'///////////////////////////////////////////////////////////////////////////////////
	Dim  Conn, rs, cmd, param
	Dim  SQL
	Dim strRankCd , strFinYr , strAmount
	
	
	if Request.Form("btnSubmit") = "Modify" then
    	
		set Conn = Server.CreateObject( "ADODB.Connection" )
		Conn.Open astrConn	
		if Err.number <> 0 then
			 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-2&astrErrDescription=" & Err.description
		     Response.End  
		end if	
		SQL = "update ADM_TEL_AMT_ELIGIB_MST set ELIGIBLE_AMT='" & Request.Form("txtEligAmt") & "'"
		SQL = SQL & "  where FINANCIAL_YEAR='" & Request.Form("txtFinYr") & "' and RANK_CD = '" & Request.Form("txtRankId") & "'"
		Conn.Execute SQL
		Conn.close
		set Conn = nothing
		Response.Redirect "../../Common/Success.asp?aintCode=1&astrPage=../ADM/TelEligAmt/admIntTelEligAmt.asp?Param=M"
		
		
	end if
  
   	'////////////////////////////////////////////////////////////////////////////////////
   	' This part of code gets executed when the form gets loaded.
   	'///////////////////////////////////////////////////////////////////////////////////
   	
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

<FORM name="frmMain" method="post" >

<TABLE border=0 width="100%">
   <TR>
    <TD class="clsHeader">Telephone Eligibilty Amount Master - Modify</TD>
   </TR>
</TABLE>
<BR>
<TABLE width="20%" class="clsTabHead" cellpadding=0 cellspacing=0 border=0 >
	<TR>
	<TH  class="clsActive"   width="100%">
		Telephone Eligibilty Amount Master
	</TH>
	</TR>
</TABLE>

<DIV ID="divTab1" name="divTab1" class="clsTabBody">
<TABLE border=0 width="100%">
 <TR>
 <TD>
 </TD>
 </TR> 
  
 <tr>
	<td width="20%"><font class="clsLabel">Financial Year</font></td>
	<td width="30%" colSpan="3">
	<input type="text" class="clsSmall ; clsDisabled" name="txtFinYr" maxlength="9" value="<%=strFinYr%>">
	</td>
</tr>


<tr>
	<td width="20%"><font class="clsLabel">Rank Code</font></td>
	<td width="30%" colSpan="3">
	<input type="text" class="clsSmall ; clsDisabled" name="txtRankId" maxlength="10" value="<%=strRankCd%>">
	</td>  
	
</tr>


<tr>
  <td width="20%"><font class="clsLabel">Eligible Amount</font><FONT class="clsMandatory"> &nbsp;*</FONT></td>
  <td width="30%"><input class="clsSmall" name="txtEligAmt" maxlength="15" value="<%=strAmount%>"></td>
  <td width="20%"></td>
  <td width="30%"></td>
</tr>

<tr>
  <td colspan="4">&nbsp;</td>
</tr>
	
</TABLE>
</DIV>
<BR>
<CENTER>
<INPUT TYPE="submit" name="btnSubmit"  value="Modify" class ="clsButton"   onClick="return fnSubmit();" >&nbsp;&nbsp; &nbsp;
<INPUT TYPE="reset" name="btnClear" value="Reset" class="clsButton" >&nbsp; &nbsp; &nbsp;
<INPUT TYPE="button" name="btnCancel"  value="Cancel" onClick="history.go(-1)" class="clsButton" >
</CENTER>
</FORM>
</BODY>
</HTML>

