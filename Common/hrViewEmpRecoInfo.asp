<%
'on error resume next
'CONNECT TO DATABASE
	
'Declarations
Dim dcEmp				'Holds Database connection
	
'Create Connection object
Set dcEmp = Server.CreateObject ("ADODB.Connection")

'Establish database connection
dcEmp.Open astrConn
	
'Check the Error Code
If Err.number <> 0 Then
	
	'Nulifying object before redirecting to Error.asp
	Set dcEmp = Nothing
		
	'Show Error page
	Response.Redirect("Error.asp?aintPage=-1&aintCode=1")
	
End If

'READ EMPLOYEE DETAILS
	
'Declarations
Dim rsEmpDetails			'Holds the details of the corresponding Bank Code
Dim sGetEmpDetailsQuery	'Holds the SQL query used to retrive the Bank details corresponding to Bank Code

'Checking NULL for Bank Code
If intEmpNo <> "" Then
	
	'Create Bank Details Record Set object
	Set rsEmpDetails = Server.CreateObject ("ADODB.Recordset")
	
	'Prepare SQL query string
	sGetEmpDetailsQuery = "select EMP_FIRSTNAME, EMP_MIDNAME, EMP_LASTNAME, LOC_CD, DESIG_CD, RANK_CD"
	sGetEmpDetailsQuery = sGetEmpDetailsQuery & " from HRD_EMP_MST where EMP_NO = '" & intEmpNo & "'"
			
	'Retriving Bank Details record set
	rsEmpDetails.Open sGetEmpDetailsQuery, dcEmp, adOpenStatic
		
	'Check Error Code
	If Err.number <> 0 Then
		'Disconnect databse & nulify objects created
		dcEmp.Close
		Set dcEmp = Nothing
		'Show Error page
		Response.Redirect("Error.asp?aintPage=-1&aintCode=1")
	End If

	'Check record set for end of file
	If rsEmpDetails.EOF Then
		'Close record set and disconnect database
		rsEmpDetails.Close
		dcEmp.Close
		'Nulifying all objects before redirecting the page to Error.asp
		Set rsEmpDetails = Nothing
		Set dcEmp = Nothing
		'Redirect to Error Page
		Response.Redirect("Error.asp?aintPage=-1&aintCode=2")
	Else
		strEmpName = rsEmpDetails("EMP_FIRSTNAME") & " " & rsEmpDetails("EMP_MIDNAME") & " " & rsEmpDetails("EMP_LASTNAME")
		strLocation = rsEmpDetails("LOC_CD")
		strRank = rsEmpDetails("RANK_CD")
	End If
	
	strRankName=fnGetNameFromCode("HRD_RANK_MST", "RANK_ID", "DESCRIPTION", strRank)

	if strRankName="0" then
		strRankName=""
	end if
	'Close record set
	rsEmpDetails.Close
	
	if strRecoType="LEAVE" then
			strGetEmpNo = "select a.LEV_STRT_DT LevStDt, a.LEV_END_DT LevEndDt, b.DESCRIPTION LevDesc from HRD_LEAVE_TXN a, HRD_LEAVE_TYPE b where a.LEV_APPLN_NUM='"&intLeaveAppNo&"' and a.LEAVE_CD=b.LEAVE_ID"
	
			rsEmpDetails.open strGetEmpNo, dcEmp
			'Check Error number
			If Err.number <> 0 Then
				'nullify all objects
				Set dcEmp = Nothing
				Set rsEmpDetails = Nothing
				'Show Error page
				Response.Redirect("../../common/Error.asp?aintPage=-1&aintCode=4")
			End If

			if not rsEmpDetails.EOF then
				strStartDt=fnDate(rsEmpDetails.fields("LevStDt"))
				strEndDt=fnDate(rsEmpDetails.fields("LevEndDt"))
				strLeaveDesc=rsEmpDetails.fields("LevDesc")
			end if		
	end if   

	if strRecoType="LOAN" then

			strGetLoanDtls = "select LOAN_TYPE, LOAN_AMT, LOAN_DT, APPLN_STAT from HRD_EMP_LOAN where LOAN_ID='"&intLoanAppNo&"'"
	
			rsEmpDetails.open strGetLoanDtls, dcEmp
			'Check Error number
			If Err.number <> 0 Then
				'nullify all objects
				Set dcEmp = Nothing
				Set rsEmpDetails = Nothing
				'Show Error page
				Response.Redirect("../../common/Error.asp?aintPage=-1&aintCode=4")
			End If

			if not rsEmpDetails.EOF then
				strLoanType = rsEmpDetails.fields("LOAN_TYPE")
				intAmt = rsEmpDetails.fields("LOAN_AMT")
				strLoanDt = fnDate(rsEmpDetails.fields("LOAN_DT"))
				strApplStat = rsEmpDetails.fields("APPLN_STAT")
			end if
			rsEmpDetails.close
	
			SQLDesc = "select DESCRIPTION from HRD_LOAN_MST where LOAN_TYPE='" & strLoanType & "'"
			rsEmpDetails.open SQLDesc, dcEmp
			'Check Error number
			If Err.number <> 0 Then
				'nullify all objects
				Set dcEmp = Nothing
				Set rsEmpDetails = Nothing
				'Show Error page
				Response.Redirect("../../common/Error.asp?aintPage=-1&aintCode=4")
			End If

			if not rsEmpDetails.EOF then
				strLoanTypeDesc = rsEmpDetails.fields("DESCRIPTION")
			else
				strLoanTypeDesc = "Not Available"
			end if
			rsEmpDetails.close
	
			select case strApplStat	
			case "I"
			strStatDesc = "Initiated"
			case "R"
			strStatDesc = "Rejected"
			case "A"
			strStatDesc = "Approved"
			end select
		
	end if


	if strRecoType="MEDICALCLAIM" then

			strGetClaimDtls = "select MEDICAL_CLAIM_TYPE, CLAIM_DT, CLAIMED_AMT, CLAIM_STAT from HRD_EMP_MED_CLAIM where SR_NO='"&intMedicalClaimAppNo&"'"
	
			rsEmpDetails.open strGetClaimDtls, dcEmp
			'Check Error number
			If Err.number <> 0 Then
				'nullify all objects
				Set dcEmp = Nothing
				Set rsEmpDetails = Nothing
				'Show Error page
				Response.Redirect("../../common/Error.asp?aintPage=-1&aintCode=4")
			End If

			if not rsEmpDetails.EOF then
				strMedicalClaimType = rsEmpDetails.fields("MEDICAL_CLAIM_TYPE")
				intAmt = rsEmpDetails.fields("CLAIMED_AMT")
				strMedicalClaimDt = fnDate(rsEmpDetails.fields("CLAIM_DT"))
				strApplStat = rsEmpDetails.fields("CLAIM_STAT")
			end if
			rsEmpDetails.close
	
			SQLDesc = "select CLAIM_TYPE_DESCRIPTION from HRD_MEDICAL_MST where MEDICAL_CLAIM_TYPE='" & strMedicalClaimType & "'"
			rsEmpDetails.open SQLDesc, dcEmp
			'Check Error number
			If Err.number <> 0 Then
				'nullify all objects
				Set dcEmp = Nothing
				Set rsEmpDetails = Nothing
				'Show Error page
				Response.Redirect("../../common/Error.asp?aintPage=-1&aintCode=4")
			End If

			if not rsEmpDetails.EOF then
				strMedicalClaimTypeDesc = rsEmpDetails.fields("CLAIM_TYPE_DESCRIPTION")
			else
				strMedicalClaimTypeDesc = "Not Available"
			end if
			rsEmpDetails.close
	
			select case strApplStat	
			case "I"
			strStatDesc = "Initiated"
			case "R"
			strStatDesc = "Rejected"
			case "A"
			strStatDesc = "Approved"
			end select
		
	end if

	if strRecoType="LTC" then
			strLtcFlg = left(strRecoType1,1)
			strGetEmpNo = "select BLOCK_YEAR, PLACE_OF_TRAVEL, LTC_TYPE, AMOUNT_CLAIMED, APPROVAL_STATUS from HRD_LTC_TXN where LTC_NO='"&intLTCNo&"' and LTC_FLAG='"&strLtcFlg&"'"
	
			rsEmpDetails.open strGetEmpNo, dcEmp
			'Check Error number
			If Err.number <> 0 Then
				'nullify all objects
				Set dcEmp = Nothing
				Set rsEmpDetails = Nothing
				'Show Error page
				Response.Redirect("../../common/Error.asp?aintPage=-1&aintCode=4")
			End If

			if not rsEmpDetails.EOF then
				strBLOCK_YEAR=rsEmpDetails.fields("BLOCK_YEAR")
				strPLACE_OF_TRAVEL=rsEmpDetails.fields("PLACE_OF_TRAVEL")
				strLTC_TYPE=rsEmpDetails.fields("LTC_TYPE")
				strAMOUNT_CLAIMED=rsEmpDetails.fields("AMOUNT_CLAIMED")
				strApplStat=rsEmpDetails.fields("APPROVAL_STATUS")
			end if		
			
			select case strApplStat	
			case "I"
			strStatDesc = "Initiated"
			case "R"
			strStatDesc = "Rejected"
			case "A"
			strStatDesc = "Approved"
			end select			
	end if   
		
	'Nulifying the objects
	Set dcEmp = Nothing
	Set rsEmpDetails = Nothing
End if	

if strRecoType="TADA" then
'	intAmountClm= rsReco.fields("ADV_AMT_CLAIMED")
'	intAmountApp= rsReco.fields("ADV_AMT_APPROVED")
'	strCityName=rsReco.fields("CITY_NAME")
'	intEmpNo=rsReco.fields("EMP_NO")
'	intTadaNo = rsReco.fields("IND_TOUR_ID")
'	strStatus = rsReco.fields("APPROVAL_STATUS")
	
	select case strStatus	
	case "I"
	strStatDesc = "Initiated"
	case "R"
	strStatDesc = "Rejected"
	case "A"
	strStatDesc = "Approved"
	end select				
end if	
%>

<HTML>
<HEAD>

<!--
'***************************************************************************
' Application Name		: View Employee Details for REcommendation purpose
' Author Name			: Riju
' Date of Creation		: 25-Mar-2002
' Version Number		: 1.10
' Purpose				: This is the page for viewing Employee Details
' Remarks				: 
'****************************************************************************
 -->

<TITLE>ECGC: Employee Info</TITLE>
<LINK rel="stylesheet" type="text/css" href="../../stylesheets/ecgc.css">
<SCRIPT language="JavaScript" src="../../includes/incCommonFunctions.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incValidationFunctions.js"></SCRIPT>

</HEAD>

<BODY>

<!-- Start of DIV (divBank) -->

<DIV ID="divBank" name="divBank" class="clsTabBody">
<TABLE border=0 width="100%" id=TABLE2>
	<TR>
		<TD width="20%"><FONT class="clsSectionLabel">Applicant Details</FONT></TD>
	</TR>
	
	<TR>
		<TD width="20%"><FONT class="clsLabel">Employee Number</FONT></TD>
		<TD width="30%"><%=intEmpNo%></TD>
		<TD width="20%"><FONT class="clsLabel">Name</FONT></TD>
		<TD width="30%"><%=strEmpName%></TD>
	</TR>
	<TR>	
		<td width="20%"><font class="clsLabel">Location</font></td>
		<td width="30%"><font class="clsText"><%=fnGetNameFromCode("LOGICAL_LOC_MST", "LOGICALLOC_CD", "DESCRIPTION", strLocation)%></font></td>
		<TD><FONT class="clsLabel">Rank</FONT></TD>
		<TD><%=strRankName%></TD>
	</TR> 

	<%if strRecoType="LEAVE" then%>
	<TR>
		<TD width="20%"><FONT class="clsLabel">Leave Type</FONT></TD>
		<TD width="30%"><%=strLeaveDesc%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Start Date</FONT></TD>
		<TD width="30%"><%=strStartDt%></TD>
		<TD><FONT class="clsLabel">End Date</FONT></TD>
		<TD><%=strEndDt%></TD>
	</TR>
	<%end if%>
	
	<%if strRecoType="ALLOWANCE" then%>
	<TR>
		<TD width="20%"><FONT class="clsLabel">Allowance Type</FONT></TD>
		<TD width="30%"><%=strAllType%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Amount</FONT></TD>
		<TD width="30%"><%=intAmount%> Rs.</TD>
		<%if strAllType="One Time" then%>
		<TD><FONT class="clsLabel">Date</FONT></TD>
		<TD><%=fnDate(txtForMonth)%></TD>
		<%end if%>
	</TR>
	<%end if%>
	
	<%if strRecoType="LOAN" then%>
	<TR>
		<TD width="20%"><FONT class="clsLabel">Loan ID</FONT></TD>
		<TD width="30%"><%=intLoanAppNo%></TD>
		<TD width="20%"><FONT class="clsLabel">Loan Description</FONT></TD>
		<TD width="30%"><%=strLoanTypeDesc%></TD>		
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Loan Amount</FONT></TD>
		<TD width="30%"><%=intAmt%></TD>
		<TD><FONT class="clsLabel">Loan Date</FONT></TD>
		<TD><%=strLoanDt%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Application Status</FONT></TD>
		<TD width="30%"><%=strStatDesc%></TD>
	</TR>	
	<%end if%>	

	<%if strRecoType="MEDICALCLAIM" then%>
	<TR>
		<TD width="20%"><FONT class="clsLabel">Claim No.</FONT></TD>
		<TD width="30%"><%=intMedicalClaimAppNo%></TD>
		<TD width="20%"><FONT class="clsLabel">Claim Type</FONT></TD>
		<TD width="30%"><%=strMedicalClaimTypeDesc%></TD>		
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Claimed Amount</FONT></TD>
		<TD width="30%"><%=intAmt%></TD>
		<TD><FONT class="clsLabel">Claim Date</FONT></TD>
		<TD><%=strMedicalClaimDt%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Claim Status</FONT></TD>
		<TD width="30%"><%=strStatDesc%></TD>
	</TR>	
	<%end if%>
	
	<%if strRecoType="LTC" then%>
	<TR>
		<TD width="20%"><FONT class="clsLabel">LTC No.</FONT></TD>
		<TD width="30%"><%=intLTCNo%></TD>
		<TD width="20%"><FONT class="clsLabel">LTC Type</FONT></TD>
		<TD width="30%"><%=strLTC_TYPE%></TD>		
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Claimed Amount</FONT></TD>
		<TD width="30%"><%=strAMOUNT_CLAIMED%></TD>
		<TD><FONT class="clsLabel">Block Year</FONT></TD>
		<TD><%=strBLOCK_YEAR%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Place of Travel</FONT></TD>
		<TD width="30%"><%=strPLACE_OF_TRAVEL%></TD>
		<TD width="20%"><FONT class="clsLabel">Claim Status</FONT></TD>
		<TD width="30%"><%=strStatDesc%></TD>
	</TR>	
	<%end if%>
	
	<%if strRecoType="TADA" then%>
	<TR>
		<TD width="20%"><FONT class="clsLabel">TA-DA No.</FONT></TD>
		<TD width="30%"><%=intTadaNo%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Claimed Amount</FONT></TD>
		<TD width="30%"><%=intAmountClm%></TD>
		<TD><FONT class="clsLabel">Approved Amount</FONT></TD>
		<TD><%=intAmountApp%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Purpose</FONT></TD>
		<%if strPurpose="TO" then%>
		<TD width="30%">Tour</TD>
		<%else%>
		<TD width="30%">Transfer</TD>
		<%end if%>
		
		<%if strTourType="I" then%>
		<TD width="20%"><FONT class="clsLabel">City</FONT></TD>
		<TD width="30%"><%=strDestn%></TD>
		<%else%>
		<TD width="20%"><FONT class="clsLabel">Country</FONT></TD>
		<TD width="30%"><%=strDestn%></TD>		
		<%end if%>
	</TR>	
	<%end if%>
</TABLE>

</DIV>
</BODY>
</HTML>