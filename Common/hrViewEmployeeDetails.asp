<%@ Language=VBScript %>
<%Response.Expires = -1%>
<%'On error Resume Next %>
<!-- #INCLUDE FILE="../../Common/Connection.asp"-->
<!-- #INCLUDE FILE="../../Common/CommonFunctions.asp" -->
<!-- #include file="../../common/check.asp"-->


<%
'on error resume next
'CONNECT TO DATABASE

'Declarations
Dim dcEmp				'Holds Database connection
Dim sBankCode			'Holds the Bank Code
	
'Read Bank Code
intEmpNo = Request.QueryString("Emp_No")

if intEmpNo="" or isNull(intEmpNo) then
Response.Redirect "../../Common/Error.asp?aintCode=999&aintPage=&astrErrDescription=" & Err.description
Response.End 
end if

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
	sGetEmpDetailsQuery = "select EMP_FIRSTNAME, EMP_MIDNAME, EMP_LASTNAME, LOC_CD, DESIG_CD,DIV_CD, DEPT_CD, RANK_CD"
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
		strEmpDiv = rsEmpDetails.Fields("DIV_CD")
	    strEmpDept = rsEmpDetails.Fields("DEPT_CD")
	    
	      
				
	 End If
		
	
	strDivName = fnGetNameFromCode("DIVISION_MST", "DIV_CD", "DIV_NAME", strEmpDiv )
	strDeptName = fnGetNameFromCode("DEPARTMENT_MST", "DEPT_CD", "DEPT_NAME", strEmpDept )

     if strDivName="0" then
		strDivName = ""
	end if
	if strDeptname="0" then
		strDeptName = ""
	end if
	strRankName=fnGetNameFromCode("HRD_RANK_MST", "RANK_ID", "DESCRIPTION", strRank)
	
	if (strRankDesc ="0" or strRankDesc ="-1") then
		strRankDesc=""
	end if
			

	'Close record set and disconnect database
	rsEmpDetails.Close
	'*********************************
		 SQL1 ="select * from  HRD_EMP_ADDRESS where EMP_NO = '" & intEmpNo & "'"
	        rsEmpDetails.Open SQL1,dcEmp
		    strEMP_NO = intEmpNo
		   ' strPER_ADDRESS1 =rsEmpDetails.Fields("PER_ADDRESS1")
		   ' strPER_ADDRESS2 = rsEmpDetails.Fields("PER_ADDRESS2")
		   ' strPER_ADDRESS3 = rsEmpDetails.Fields("PER_ADDRESS3")
		  '  strPER_CITY =rsEmpDetails.Fields("PER_CITY")
		  '  strPER_STATE = rsEmpDetails.Fields("PER_STATE")
		  '  strPER_COUNTRY = rsEmpDetails.Fields("PER_COUNTRY")
		  '  strPER_PIN_CODE = rsEmpDetails.Fields("PER_PIN_CODE")
		  '  strPER_PHONE_NO =rsEmpDetails.Fields("PER_PHONE_NO")
		  '  strPER_FAX_NO = rsEmpDetails.Fields("PER_FAX_NO")
		    strCURR_ADDRESS1 =rsEmpDetails.Fields("CURR_ADDRESS1")
		    strCURR_ADDRESS2 = rsEmpDetails.Fields("CURR_ADDRESS2")
		    strCURR_ADDRESS3 =rsEmpDetails.Fields("CURR_ADDRESS3")
		    strCURR_CITY = rsEmpDetails.Fields("CURR_CITY")
		    strCURR_STATE = rsEmpDetails.Fields("CURR_STATE")
		    strCURR_COUNTRY = rsEmpDetails.Fields("CURR_COUNTRY")
		    strCURR_PIN_CODE =rsEmpDetails.Fields("CURR_PIN_CODE")
		    strCURR_PHONE_NO = rsEmpDetails.Fields("CURR_PHONE_NO")
		    strCURR_FAX_NO = rsEmpDetails.Fields("CURR_FAX_NO")
   		'	strPER_MOBILE_NO = rsEmpDetails.Fields("PER_MOBILE_NO")
		'    strPER_EMAIL_ADD =rsEmpDetails.Fields("PER_EMAIL_ADD")
		    strCURR_MOBILE_NO = rsEmpDetails.Fields("CURR_MOBILE_NO")
		    strCURR_EMAIL_ADD = rsEmpDetails.Fields("CURR_EMAIL_ADD")	
		    
		    rsEmpDetails.Close
		    	 
		'**********************************************
	dcEmp.Close
   	'Nulifying the objects
	Set dcEmp = Nothing
	Set rsEmpDetails = Nothing
End if	
%>

<HTML>
<HEAD>

<!--
'***************************************************************************
' Application Name		: View Employee Details
' Author Name			: Riju
' Date of Creation		: 14-Mar-2002
' Version Number		: 1.10
' Purpose				: This is the page for viewing Employee Details
' Remarks				: 
'****************************************************************************
 -->

<TITLE>ECGC: Employee Details</TITLE>
<LINK rel="stylesheet" type="text/css" href="../../stylesheets/ecgc.css">
<SCRIPT language="JavaScript" src="../../includes/incCommonFunctions.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incValidationFunctions.js"></SCRIPT>

</HEAD>

<!--<BODY>-->

<!-- ********************************************************-->
<body topmargin="0" marginheight="0">

<br>
<center>
<table height="90%" width="90%" cellpadding="0" cellspacing="0" border="6" style="BORDER-BOTTOM-COLOR: gray; BORDER-BOTTOM-WIDTH: thick; BORDER-LEFT-COLOR: silver; BORDER-LEFT-WIDTH: thin; BORDER-RIGHT-COLOR: gray; BORDER-RIGHT-WIDTH: thick; BORDER-TOP-COLOR: silver; BORDER-TOP-WIDTH: thin;BACKGROUND-COLOR:white;" title="Employee Details">
<tr>
	<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr>
				<td>&nbsp;</td>
			</tr>
			<tr align="left">
				<td noWrap align="right">
				<img src="../../images/ecgc_small.gif" valign="center" WIDTH="74" HEIGHT="28">
				<!--					<FONT face=ARIAL size=3 align="right" color="black">ECGC of India Ltd.</FONT>					-->
				</td>
			</tr>
			<tr>
				<td>
					&nbsp;&nbsp;<font class="clsTextLabel" size="2"><%= strEmpName %></font>
				</td>
			</tr>
</table>
<TABLE border=0 width="100%" id=TABLE2>
	<TR>
	<TD width=40%>&nbsp;</TD>
	<TD width=60%>&nbsp;</TD>
	</TR>
	<TR align="left">
		<TD >&nbsp;&nbsp;<FONT class="clsLabel">Employee Number</FONT></TD>
		<TD ><%=intEmpNo%></TD>
	</TR>
	<TR align="left">	
		<TD >&nbsp;&nbsp;<FONT class="clsLabel">Location</FONT></TD>
		<TD ><%=fnGetNameFromCode("LOGICAL_LOC_MST", "LOGICALLOC_CD", "DESCRIPTION", strLocation) %></td>
	</TR> 
	<TR>
		<TD>&nbsp;&nbsp;<FONT class="clsLabel">Rank</FONT></TD>
		<TD><%=strRankName%></TD>
	</TR> 
	
	<TR>
	<td>&nbsp;&nbsp;<font class="clsLabel">Division</font></td>
    <td><%=strDivName%></td>
	</tr>
	<tr>
	<td>&nbsp;&nbsp;<font class="clsLabel">Department</font></td>
	<td ><%=strDeptName%></td>
	</tr>			
	
	
</table>
<TABLE border=0 width="100%" id=TABLE2>
				<tr>
					<td>
						<!--						<hr color="#20c0c9" size="1" noShade=false width="90%" align="left">						-->
						<img src="../../images/hr.gif" valign="center" WIDTH="280" HEIGHT="7">
					</td>
				</tr>
				
				<tr>
					<td>
						&nbsp;&nbsp;<font class="clsText"><i><%= strCURR_ADDRESS1 %>&nbsp;,&nbsp;<%= strCURR_ADDRESS2 %></i></font>
					</td>
				</tr>
												
				<tr>
					<td>
						&nbsp;&nbsp;<font class="clsText"><i><%= strCURR_ADDRESS3 %></i></font>
					</td>
				</tr>
				
				<tr>
					<td>
						&nbsp;&nbsp;<font class="clsText"><i><%= strCURR_CITY %>&nbsp;-&nbsp;<%= strCURR_PIN_CODE %></i></font>
					</td>
				</tr>
				
				<tr>
					<td>
						&nbsp;&nbsp;<font class="clsText"><i><%= strCURR_STATE %></i></font>
					</td>
				</tr>
				
				
				<tr>
					<td>
						&nbsp;&nbsp;<font class="clsText"><i><%= strCURR_COUNTRY %></i></font>
					</td>
				</tr>
				
				<!-- Show Phone and Fax -->
				
				<tr>
					<td>
						<table>
							<tr>
								<td>
								<img src="../../images/telephone.gif" align="absbottom" WIDTH="38" HEIGHT="32">
								</td>
								<td wrap>
									
									<table>
										<tr>
											<td><i>Phone: <%= strCURR_PHONE_NO %></i></td>
										</tr>
										
										<tr>
											<td><i>Fax: <%= strCURR_FAX_NO %></i></td>
										</tr>
										
										<tr>
											<td><i>E-mail: <%= strCURR_EMAIL_ADD %></i></td>
										</tr>
										
									</table>
								</td>
							</tr>
						</table>
					</td>	
				</tr>	
					
			</table>
		</td>
	</tr>
	</table>
</center>
</body>
</html>
<!--*************************************************************-->
<!--

<FORM name="frmMain" border=0 width="100%">

<TABLE border=0 width="100%">
	<TR>
		<TD class="clsHeader">Employee Details</TD>
	</TR>
</TABLE>



<TABLE width="20%" class="clsTabHead" cellpadding=0 cellspacing=0 border=0 >
	<TR>
		<TH name="general"   class="clsActive"   width="100%">Main</TH>
	</TR>
</TABLE>


<DIV ID="divBank" name="divBank" class="clsTabBody">
<TABLE border=0 width="100%" id=TABLE2>
	<TR>
		<TD width="20%"><FONT class="clsLabel">Employee Number</FONT></TD>
		<TD width="30%"><%=intEmpNo%></TD>
	</TR>

	<TR>
		<TD width="20%"><FONT class="clsLabel">Name</FONT></TD>
		<TD width="30%"><%=strEmpName%></TD>
	</TR>
	<TR>	
		<TD width="20%"><FONT class="clsLabel">Location</FONT></TD>
		<TD width="30%"><td width="30%"><font class="clsText"><%=fnGetNameFromCode("LOGICAL_LOC_MST", "LOGICALLOC_CD", "DESCRIPTION", strLocation) %></font></td>
	</TR> 
	<TR>	
		<TD><FONT class="clsLabel">Rank</FONT></TD>
		<TD><%=strRankName%></TD>
	</TR> 
	
	<TR>
  <TD colspan=4>&nbsp;</TD>
  </TR>
 <TR>
  <TD colSpan=4><FONT class=clsSectionLabel >Current Address</FONT></TD>
 </TR>
 
  <TR>
    <TD><FONT class="clsLabel"> Address 1</FONT></TD>
    <TD colspan=3><%=strCURR_ADDRESS1%></TD>
  </TR>
  
  <TR>
    <TD><FONT class="clsLabel"> Address 2</FONT></TD>
    <TD colspan=3><%=strCURR_ADDRESS2%></TD>
  </TR>
  
  <TR>
    <TD><FONT class="clsLabel"> Address 3</FONT></TD>
    <TD colspan=3><%=strCURR_ADDRESS3%></TD>  
  </TR>
    
  <TR>
    <TD ><FONT class="clsLabel">City</FONT></TD>
    <TD ><%=strCURR_CITY%></TD>
    <TD ><FONT class="clsLabel">State</FONT></TD>
    <TD><%=strCURR_STATE%></TD>
  </TR>
  
  <TR>
    <TD ><FONT class="clsLabel">Country </FONT></TD>
    <TD><%=strCURR_COUNTRY%> </TD>
    <TD ><FONT class="clsLabel">PIN Code</FONT></TD>
    <TD><%=strCURR_PIN_CODE%></TD>
  </TR>
  
  <TR>
    <TD ><FONT class="clsLabel">Phone No.</FONT></TD>
    <TD ><%=strCURR_PHONE_NO%></TD>
    <TD ><FONT class="clsLabel">Fax No.</FONT></TD>
    <TD><%=strCURR_FAX_NO%></TD>
  </TR>
  
  <TR>
    <TD ><FONT class="clsLabel">Mobile No.</FONT></TD>
    <TD ><%=strCURR_MOBILE_NO %></TD>
    <TD ><FONT class="clsLabel">E-mail Address</FONT></TD>
    <TD><%=strCURR_EMAIL_ADD %></TD>
   </TR>
  
  <TR>
   <TD colspan=4>&nbsp;</TD>
  </TR>
    
	
</TABLE>

</DIV>



<BR>

<CENTER>
	<INPUT TYPE="button" name="btnOk" value="OK" class="clsButton" onClick="window.close()">
</CENTER>

</FORM>
</BODY>
</HTML>
-->