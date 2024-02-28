<%@ Language=VBScript %>
<%  Response.Expires = -1 

'****************************************************************************
' Application Name		: Intermediate Telephone Eligibilty Amount Master
' Author Name			: Janani Seshadri
' Date of Creation		: 10-June-2002
' Version Number		: 1.0
' Purpose				: This is the ASP for Telephone Eligibilty Amount Master 
' Remarks				: 
'**************************************************************************** 

%>
<!--#include file="../../Common/Check.asp" -->
<!--#include file="../../Common/Connection.asp" -->
<!-- #include file="../../Common/commonFunctions.asp"-->

<html>
<head>
<title>ECGC: Telephone Eligibilty Amount Master</title>
<link rel="stylesheet" type="text/css" href="../../stylesheets/ecgc.css">
<script language="JavaScript" src="../../includes/incSearch.js"></script>
<script language="JavaScript" src="../../includes/incValidationFunctions.js"></script>
<script language="JavaScript" src="../../includes/incCommonFunctions.js"></script>
<script LANGUAGE="javascript">
<!--

//If no field is entered and next is clicked, the function displays the message.
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
		
	
	
	return true;
}



function fnOnLoad()
{
    	document.frmMain.txtFinYr.focus();
}


//-->
</script>
</head>
<body onLoad="fnOnLoad();">

<%
Dim strParam
strParam = Request.Querystring("Param")
Select Case strParam
		Case "M"
%>				
    <form name="frmMain" method="post" action="admModifyTelEligAmt.asp" onSubmit="return lfnblnIsFormValid();">
   <table border="0" width="100%">
   <tr>
    <td class="clsHeader">Telephone Eligibilty Amount Master - Modify</td>
   </tr>
   </table>
    				
				
<%				
		Case "D"
%>	
   <form name="frmMain" method="post" action="admDeleteTelEligAmt.asp" onSubmit="return lfnblnIsFormValid();">
   <table border="0" width="100%">
   <tr>
    <td class="clsHeader">Telephone Eligibilty Amount Master - Delete</td>
   </tr>
   </table>	
		
<%				
		Case "V"
%>
   <form name="frmMain" method="post" action="admViewTelEligAmt.asp" onSubmit="return lfnblnIsFormValid();">
   <table border="0" width="100%">
   <tr>
    <td class="clsHeader">Telephone Eligibilty Amount Master - View </td>
   </tr>
   </table>

			
<%		
		End Select
%>
<br>
<table width="20%" class="clsTabHead" cellpadding="0" cellspacing="0" border="0">
	<tr>
	<th class="clsActive" width="100%" onClick="fnTabNavigate(1)">
		Telephone Eligibilty Amount Master
	</th>
	</tr>
</table>

<div ID="divTab1" name="divTab1" class="clsTabBody">
<table border="0" width="100%">

<tr>
 <td width="20%"><font class="clsLabel">Financial Year</font><font class="clsMandatory">*</font></td>
 <td width="30%" colSpan="3">
 <input type="text" class="clsSmall" name="txtFinYr" maxlength="10" value="<%=Request.QueryString("fin_yr")%>">
 <img class="clsSearch" onclick="fnOpenLookup('../../Lookup/admLookupTelEligAmt.asp?fin_yr='+document.frmMain.txtFinYr.value + '&Rank_id='+ document.frmMain.txtRankId.value + '&txtbox_Cd=txtFinYr&txtbox_Name=txtRankId')" alt="Search" src="../../images/search.gif" WIDTH="17" HEIGHT="12"></td>  
</tr>


<tr>
 <td width="20%"><font class="clsLabel">Rank Code</font><font class="clsMandatory">*</font></td>
 <td width="30%" colSpan="3">
 <input type="text" class="clsSmall" name="txtRankId" maxlength="10" value="<%=Request.QueryString("Rank_id")%>">
 <img class="clsSearch" onclick="fnOpenLookup('../../HR/Masters/Rank/hrLookupRankMstRankCode.asp?Rank_id='+document.frmMain.txtRankId.value)" alt="Search" src="../../images/search.gif" WIDTH="17" HEIGHT="12"></td>  
</tr>

</table>
</div>
<br>
<center>
<input TYPE="submit" name="btnSubmit" value="Next" class="clsButton">&nbsp;&nbsp; &nbsp;
<input TYPE="reset" name="btnClear" value="Clear" class="clsButton" onClick="fnOnLoad();">&nbsp; &nbsp; &nbsp;

</center>
</form>
</body>
</html>

