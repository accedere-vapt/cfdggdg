<%@ LANGUAGE=VBScript %>
<%
Option Explicit
On Error Resume Next
Response.Expires = -1

'****************************************************************************
' Application Name		: Add Telephone Eligibilty Amount Master
' Author Name			: Janani Seshadri
' Date of Creation		: 10-June-2002
' Version Number		: 1.0
' Purpose				: This is the ASP for Telephone Eligibilty Amount Master Add
' Remarks				: 
'**************************************************************************** 
%>
<!--#INCLUDE FILE="../../Common/Connection.asp"-->
<!-- #include file="../../common/check.asp"-->
<!-- #include file="../../common/commonfunctions.asp" -->
<!-- #include file="../common/ADMConstants.inc" -->
<%

if Request.Form("btnSubmit") = "Add" then
	 
	Dim Conn, rs, cmd, param
	Dim SQL
	Dim strSql				
	Dim strNextAssetNo	
	Dim rsAssetNo
	
		
			
    set Conn = Server.CreateObject( "ADODB.Connection" )
	Conn.Open astrConn	
		
	if Err.number <> 0 then
		 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	     Response.End  
	end if
		
	set rs = Server.CreateObject( "ADODB.Recordset" )
		
	SQL ="select FINANCIAL_YEAR,RANK_CD from ADM_TEL_AMT_ELIGIB_MST where FINANCIAL_YEAR = '" & Request.Form("txtFinYr") & "' and RANK_CD = '" & ucase(Request.Form("txtRankId"))& "'"
		
	rs.Open SQL,Conn
	if Err.number <> 0 then
			 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
	end if
	if not rs.EOF then
		rs.close 
		Conn.Close 
		set rs = nothing
		set Conn = nothing
		Response.Redirect "../../Common/Error.asp?aintCode=1&aintPage=-1&astrErrDescription=" & Err.description
	else
		rs.close 
		
		SQL = "SELECT RANK_ID,STATUS FROM HRD_RANK_MST where  RANK_ID = '" & Ucase(Request.Form("txtRankId")) & "'"                 
		rs.Open SQL,Conn
		if rs.EOF then
		   rs.close 
		   Conn.Close 
		   set rs = nothing
		   set Conn = nothing
		   Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
		else
			set rs = nothing  
			SQL ="Insert into ADM_TEL_AMT_ELIGIB_MST(FINANCIAL_YEAR,RANK_CD,ELIGIBLE_AMT) "
			SQL = SQL & "values( '" 
			SQL = SQL & Request.Form("txtFinYr")
			SQL = SQL & "','" & Ucase(Request.Form("txtRankId"))
			SQL = SQL & "','" & Request.Form("txtEligAmt") & "'"		
			SQL = SQL & ")"
		
			Conn.Execute SQL
			'Check Erron Number
			If Err.number <> 0 Then
				
				'Disconnect Database and nullify object before showing Error message
				Conn.Close
				Set Conn = Nothing
		
				'Show Error page
				Response.Redirect("../../common/Error.asp?aintPage=-1&aintCode=3004&astrErrDescription=" & Err.description)
		
			End If
					
	
			'Disconnect Database and nullify all objects before showing Success message
			Conn.Close
			Set Conn = Nothing	
	
			Response.Redirect "../../common/Success.asp?aintCode=1&astrMessage=Telephone Eligibilty Amount Details have been stored in the database.&astrPage=../ADM/TelEligAmt/admAddTelEligAmt.asp"
		end if
	end if	
end if
	
%>


<html>
<head>
<title></title>
<link rel="stylesheet" type="text/css" href="../../stylesheets/ecgc.css">
<script language="JavaScript" src="../../includes/incSearch.js">
</script>
<script language="JavaScript" src="../../includes/incCommonFunctions.js">
</script>
<script language="JavaScript" src="../../includes/incValidationFunctions.js">
</script>
<script language="JavaScript" src="../../includes/incForm.js">
</script>
<script LANGUAGE="javascript">
<!--



function fnOnLoad()
{
	var frmForm;
	frmForm = document.frmMain;
	
	
	frmForm.txtFinYr.focus();
}

// Invoked when Form is submitted
// If all fields are validated correctly, submits the form


// Validates Form fields
function lfnblnIsFormValid()
{
	var frmForm = document.frmMain;
	var txtTextbox;
	var selSelect;
	var strFieldName;
	var intMandatory;	
	var radRadio;
	
	
	/////////////////////////
	// Validate Financial Year
	/////////////////////////
	txtTextbox = frmForm.txtFinYr;
	strFieldName = "Financial Year";
	intMandatory = 1;
	
	// Should not be blank
	// Should be valid
	if (!(f_bValidateFiscalYear(txtTextbox,intMandatory,strFieldName)))
	{
		return false;
	}
	
	
	/////////////////////////
	// Validate Rank
	/////////////////////////
	txtTextbox = frmForm.txtRankId;
	strFieldName = "Rank";
	intMandatory = 1;
	
	// Should not be blank
	// Should be valid
	if(gf_bValidateCode(txtTextbox, intMandatory, strFieldName) == false)
	{
		return false;
	}
		
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

//-->
</script>
</head>
<body onLoad="fnOnLoad()">

<table border="0" width="100%">
   <tr>
    <td class="clsHeader">Telephone Eligibilty Amount Master - Add</td>
   </tr>
</table>

<form name="frmMain" Method="Post" onSubmit="return lfnblnIsFormValid();">
<!-- Start Tab Headings -->
<table width="20%" class="clsTabHead" cellpadding="0" cellspacing="0" border="0">
<tr>
	<th class="clsActive" width="100%">
	 Telephone Eligibilty Amount </th>
		
</tr>
</table>
<!-- End Tab Headings -->
<!-- Start of DIV (divTabMain) for Tab 1 -->
<div ID="divTabMain" name="divTabMain" class="clsTabBody">
<table border="0" width="100%">
	
<tr>
	<td width="20%"><font class="clsLabel">Financial Year</font><font class="clsMandatory">*</font></td>
	<td width="30%" colSpan="3">
	<input type="text" class="clsSmall" name="txtFinYr" maxlength="9">
	</td>
</tr>


<tr>
	<td width="20%"><font class="clsLabel">Rank Code</font><font class="clsMandatory">*</font></td>
	<td width="30%" colSpan="3">
	<input type="text" class="clsSmall" name="txtRankId" maxlength="10" value="<%=Request.QueryString("Rank_id")%>">
	<img class="clsSearch" onclick="fnOpenLookup('../../HR/Masters/Rank/hrLookupRankMstRankCode.asp?Rank_id='+document.frmMain.txtRankId.value)" alt="Search" src="../../images/search.gif" WIDTH="17" HEIGHT="12"></td>  
	
</tr>


<tr>
  <td width="20%"><font class="clsLabel">Eligible Amount</font><FONT class="clsMandatory"> &nbsp;*</FONT></td>
  <td width="30%"><input class="clsSmall" name="txtEligAmt" maxlength="15"></td>
  <td width="20%"></td>
  <td width="30%"></td>
</tr>

<tr>
  <td colspan="4">&nbsp;</td>
</tr>
	

</table>
</div>
<!-- End of DIV (divTabMain) for Tab 1 -->

<!-- Start of Form buttons -->
<br>
<center>
<INPUT TYPE="submit" name="btnSubmit"  value="Add" class ="clsButton" >&nbsp;&nbsp; &nbsp;
<INPUT TYPE="reset" name="btnClear" value="Clear" class="clsButton" onClick="fnOnLoad();">&nbsp; &nbsp; &nbsp;
</center>
<!-- End of Form buttons -->
</form>
</body>
</html>