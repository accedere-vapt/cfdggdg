
<%

' Included in NEIARptBrTrialBal.asp to show logical location code field
' in enabled and disabled mode depending on location of logged-in user.

Dim lstrLoginUserLogicalLocCd
		
' Get Employee Details
fnGetEmpDtls Session("sEmployeeNo")
strLoginUserLogicalLocCd = strBranchCd

%>	

<!-- Define variables for Logical Location Code lookup -->
<input TYPE="hidden" name="hdnCd1">
<input TYPE="hidden" name="hdnCd2">
<input TYPE="hidden" name="hdnName3">	

<%if strLoginUserLogicalLocCd <> "DAL2" then %>
	<td width="20%"><font class=clsLabel>Logical Location Code</font></td>
	<td width="30%">
		<input class="clsSmall ; clsDisabled"  maxlength="10" size="10" value="<%=ucase(strLoginUserLogicalLocCd)%>" readonly name=txtLogicalLocCd>
	</td>
<%else%>
	<td width="20%"><font class=clsLabel>Logical Location Code</font> <font class=clsMandatory>*</font></td>
	<td width= "30%">
		<input class="clsSmall" name="txtLogicalLocCd" maxlength="10" size="10" value="<%=ucase(strLoginUserLogicalLocCd)%>" >
		<img class="clsSearch" onclick="fnOpenLookup('../../lookup/lookupLogicalLocMst.asp?txtCd3='+document.frmMain.txtLogicalLocCd.value+'&Txtbox_Code3=txtLogicalLocCd&Txtbox_Name3=hdnName3&Txtbox_Code2=hdnCd2&Txtbox_Code1=hdnCd1')"
		alt="Search" src="../../images/search.gif" WIDTH="17" HEIGHT="12">
	</td>
<%end if%>
