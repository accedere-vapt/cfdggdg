<%@ Language=VBScript %>
<%Response.Expires = -1%>
<%
'Created by SHRADDHA PALANDE
%>
<!-- #include file="../../common/Check.asp"-->
<!-- #include file="../../common/Connection.asp"-->
<!-- #include file="../../common/CommonFunctions.asp"-->
<%
Set Con = server.CreateObject("ADODB.Connection")
Set Rs1 = server.CreateObject("ADODB.Recordset")
Set Rs2 = server.CreateObject("ADODB.Recordset")
Con.Open astrConn
IntEmpNo = request.querystring("IntEmpNo")
TxnNo = request.querystring("TxnNo")

	SQL = "select count(*) from adm_lease_pymt_schedule where emp_no='" & IntEmpNo & "' and status not in('Paid','Cancelled','Revoked','NotPaid')"
	rs2.open SQL,Con
		if cint(rs2.fields(0))>=1 then
			strmsgsuccess = "Decision is pending for previous schedule.kindly complete previous decision."
			Response.Redirect "../../../Common/Error.asp?aintCode=&aintPage=-1&astrmessage=" & strmsgsuccess
		end if

if request.form("btnSubmit") = "Revoke" then

	SQL = " Update ADM_LEASE_AGREEMENT_MST set CONTRACT_STATUS='Revoked', REVOKEDDATE=to_date('"&request.form("txtRevokeDt")&"','dd-mm-yyyy'),LAST_TXN_DATE = sysdate,MODIFIED_BY= '"&Session("sEmployeeNo")&"' where emp_no='"&IntEmpNo&"' and txn_no='"&TxnNo&"'  "
	Con.Execute SQL
	
	SQL = " update ADM_LEASE_PYMT_SCHEDULE set status='Revoked' where emp_no='"&IntEmpNo&"' and txn_no='"&TxnNo&"' and status='NotPaid' "
	Con.Execute SQL

	strSucMsg =  "<br>Transaction No. "& TxnNo& " revoked successfully.<br><br>"
	Response.Redirect "../../Common/success.asp?aintCode=1&astrPage=../ADM/Allowances/hrIntRentMasterSelfModifyView.asp?astrOpCode=RA&astrMessage=" & strSucMsg 
end if
%>
<HTML>
<HEAD>
<TITLE>ECGC: Lease Rent Reimbursement</TITLE> 
<LINK rel="stylesheet" type="text/css" href="../../stylesheets/ecgc.css">
<SCRIPT language="JavaScript" src="../../includes/incSearch.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incCommonFunctions.js"></SCRIPT>
<SCRIPT language="JavaScript" src="../../includes/incValidationFunctions.js"></SCRIPT>
<script language="vbscript" src="../../includes/temp.vbs"></script>
<script language="VBScript">
Function VBDINFunc(empNo)

	SQL = " select * from ADM_LEASE_AGREEMENT_MST where emp_no = '"&empNo&"' order by LAST_TXN_DATE desc "
	pCallDin= GetData("../../includes/",SQL)
	VBDINFunc= pCallDin
	
End Function

</script>
<SCRIPT LANGUAGE=javascript>
function fnOnload()
{
	 Quarter = VBDINFunc(document.getElementById("txtEmp").value)

}
function fnSubmit()
{
	//var Cnt = document.getElementById("hdnCnt").value;
	//
	//for(var i = 1;i<=Cnt;i++)
	//{
	//	var StrStatus = document.getElementById('SelStatus'+i).value;
	//	if(StrStatus == "NotPaid")
	//	{
	//		alert("Kindly change all NotPaid status to revoked.");
	//		return false;
	//	}
	//}
	if (!confirm("Are you sure you want to officially cancel upcomging payments for this transaction?"))
		{
			return false;
		}
	
		return true;

}
</SCRIPT>
</HEAD>
<BODY>
<TABLE border=0 width="100%">
	<TR>
		<TD class="clsHeader">Lease Rent Reimbursement</TD>
	</TR>
</TABLE>
<FORM name="frmMain" method="post" onSubmit="return fnSubmit();">
<input type = "hidden" name = "hdnEmpNo" name = "hdnEmpNo"  value = "<%=IntEmpNo%>">
<div>
<TABLE width="25%" class="clsTabHead" cellpadding=0 cellspacing=0 border=0 >
	<TR>
		<TH name="main" class="clsActive" width="100%" >Revoke Lease Rent Master</TH>
	</TR>
</TABLE>
<!-- #include file= "../../HR/Common/hrEmployeeHeaderDlts.asp"-->
<br>
<DIV name="divTab1" class="clsTabBody">
<table border=0 width="100%" cellspacing="0" cellpadding="0" id="tblMain">
	<TR>
		<th style="text-align:center"> Sr No.</TH>
		<th style="text-align:center"> Txn No.</TH>
		<th style="text-align:center"> Emp No.</TH>
		<th style="text-align:center"> Lease From Date</TH>
		<th style="text-align:center"> Lease To Date</TH>
		<th style="text-align:center"> Monthly Lease Amount</TH>
		<th style="text-align:center"> Status</TH>
     </TR>
<%
 Set Rs3 = server.CreateObject("ADODB.Recordset")
 SQL = " SELECT * FROM ADM_LEASE_PYMT_SCHEDULE where emp_no='"&IntEmpNo&"' and txn_no='"&TxnNo&"' and status='NotPaid'"
 Rs3.Open SQL,Con
	if not Rs3.eof then


 Set Rs = server.CreateObject("ADODB.Recordset")
 SQL = " SELECT * FROM ADM_LEASE_PYMT_SCHEDULE where emp_no='"&IntEmpNo&"' and txn_no='"&TxnNo&"' and status='Paid'"
 Rs.Open SQL,Con

 
%>

<%
	 i = 0
	If not Rs.Eof then
		While not Rs.Eof
		i = i +1
%>		
	<TR>
		<td align="center" width="10%"><%=i%></td>
		<td align="center" width="10%"><%=Rs(0)%></td>
		<td align="center" width="10%"><%=Rs(2)%></td>
		<td align="center" width="20%"><%=fndate(Rs(3))%></td>
		<td align="center" width="20%"><%=fndate(Rs(4))%></td>
		<td align="center" width="20%"><%=Rs(5)%></td>
		<td align="center" width="20%"><%=Rs(7)%>
    </TR>

<%	
		RevokeDt  = Rs(4)
		Rs.MoveNext
		
		wend
	else
%>
		<td align="Center" width="10%" Colspan="7"><Font Class="clsMandatory">--No paid records Found--</Font></td>
<%	
	End if
	
%>	
</table>
<br>
<br>
<table border=0 width="100%" cellspacing="0" cellpadding="0" id="tblMain">
<tr>
	<td colspan = "9"><FONT color = "red" size = "2px">Following pending payment will be revoked :</FONT><br><br></td>
</tr>
<%
 Set Rs2 = server.CreateObject("ADODB.Recordset")
 SQL = " SELECT * FROM ADM_LEASE_PYMT_SCHEDULE where emp_no='"&IntEmpNo&"' and txn_no='"&TxnNo&"' and status='NotPaid'"
 Rs2.Open SQL,Con
 RevokedDt  = Rs2(3)
%> 	
<%
	 i = 0
	If not Rs2.Eof then
		While not Rs2.Eof
		i = i +1
%>	
	<TR>
		<td align="center" width="10%"><%=i%></td>
		<td align="center" width="10%"><%=Rs2(0)%></td>
		<td align="center" width="10%"><%=Rs2(2)%></td>
		<td align="center" width="20%"><%=fndate(Rs2(3))%></td>
		<td align="center" width="20%"><%=fndate(Rs2(4))%></td>
		<td align="center" width="20%"><%=Rs2(5)%></td>
		<td align="center" width="20%">
		<select id = "SelStatus<%=i%>" name = "SelStatus">
		<option value = "Revoked">Revoked</option></select></td>
    </TR>

<%	
		Rs2.MoveNext
		wend
	else
%>
		<td align="Center" width="10%" Colspan="7"><Font Class="clsMandatory">--No NotPaid records Found--</Font></td>
<%	
	End if

%>	
<input type = "hidden" name = "hdnCnt" id = "hdnCnt" value = <%=i%>> 
</table>
<br>
<%if RevokeDt = "" then%>
	<FONT color = "red">Effective Revoke Date</FONT><FONT class="clsMandatory">&nbsp;*</FONT>
	<INPUT  name=txtRevokeDt class="clsDate" maxlength=10 ReadOnly value="<%=fndate(RevokedDt)%>">
	<% Response.write fnDisplayNewDatePicker("frmMain.txtRevokeDt",2) %> 
<%else%>
	<FONT color = "red">Effective Revoke Date</FONT><FONT class="clsMandatory">&nbsp;*</FONT>
	<INPUT  name=txtRevokeDt class="clsDate" maxlength=10 ReadOnly value="<%=fndate(RevokeDt)%>">
	<% Response.write fnDisplayNewDatePicker("frmMain.txtRevokeDt",2) %> 
<%end if%>
</div>
</div>
<br>
<br>
<center>
	<INPUT TYPE="submit" name="btnSubmit" value="Revoke" class="clsButton">&nbsp;&nbsp; &nbsp;
</center>
<%else%>
	<td align="Center" width="10%" Colspan="7"><Font Class="clsMandatory">--No record Found--</Font></td>
<%end if%>
</form>
</body>
</html>