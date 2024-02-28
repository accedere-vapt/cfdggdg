<%
	Set rs = Server.CreateObject ("ADODB.Recordset")
	Set cn = Server.CreateObject ("ADODB.Connection")
	if Request("txtEmpNo")<>"" then
	strSQL = "select DIV_CD, DEPT_CD,RANK_CD,EMP_FIRSTNAME,EMP_MIDNAME,EMP_LASTNAME,EMP_NO,EMP_STATUS,BRANCH_CD from HRD_EMP_MST where emp_no= UPPER('"& Request("txtEmpNo")&"')" 'Query to get division and department
	else
	strSQL = "select DIV_CD, DEPT_CD,RANK_CD,EMP_FIRSTNAME,EMP_MIDNAME,EMP_LASTNAME,EMP_NO,EMP_STATUS,BRANCH_CD from HRD_EMP_MST where emp_no= UPPER('"& intEmpNo &"')" 'Query to get division and department
	'Response.Write strSQL
	'Response.End 
	end if
	
	
	cn.Open astrConn	'Open the Connection to the Database
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if	 
	rs.Open strSQL , cn 
	if Err.number <> 0 then
		Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if	 
	if not rs.EOF then
     strEMP_NO = trim(rs.Fields("EMP_NO"))  	
	 strBRANCH_CD = trim(rs.Fields ("BRANCH_CD"))
	 strEmpDiv = trim(rs.Fields("DIV_CD"))
	 strSTATUS = fnDecodeLOV("HRD","EMP_STAT",trim(rs.Fields("EMP_STATUS"))) 
     strEmpDept = trim(rs.Fields("DEPT_CD"))
	 strEmpRankCd = trim(rs.Fields("RANK_CD"))
	 strEMP_FIRSTNAME =trim(rs.Fields("EMP_FIRSTNAME"))
	 strEMP_MIDNAME =trim(rs.Fields("EMP_MIDNAME")) 
	 strEMP_LASTNAME =trim(rs.Fields("EMP_LASTNAME")) 
	 strEMP_NAME=strEMP_FIRSTNAME &" "& strEMP_MIDNAME &" "& strEMP_LASTNAME
	else
	 strEmpDiv = ""
	 strEmpDept = ""
	end if
	strDivName = fnGetNameFromCode("DIVISION_MST", "DIV_CD", "DIV_NAME", strEmpDiv )
	strDeptName = fnGetNameFromCode("DEPARTMENT_MST", "DEPT_CD", "DEPT_NAME", strEmpDept )

	if strDivName="0" then
		strDivName = ""
	end if
	if strDeptname="0" then
		strDeptName = ""
	end if		

	strRankDesc = fnGetNameFromCode("HRD_RANK_MST", "RANK_ID", "DESCRIPTION", strEmpRankCd)
	
	'Response.Write strRankDesc
	'Response.End  
	
	if (strRankDesc ="0" or strRankDesc ="-1") then
		strRankDesc=""
	end if
%>
<div class="clsTabHeader">
		<table border="0" width="100%">
			<tr>
				<td width="20%"><font class="clsLabel">Employee No.</font></td>
			    <td width="30%"><font class="clsText"><%= strEMP_NO %></font></td>
				<td width="20%"><font class="clsLabel">Employee Name</font></td>
				<td width="30%"><font class="clsText"><%=strEMP_NAME%></font></td>
			</tr>
			
			<tr>
				<td width="20%"><font class="clsLabel">Status</font></td>
			    <td width="30%"><font class="clsText"><%= strSTATUS %></font></td>
				<td width="20%"><font class="clsLabel">Rank</font></td>
				<td width="30%"><font class="clsText"><%=strRankDesc%></font></td>
			</tr>

			<tr>
				<td width="20%"><font class="clsLabel">Office</font></td>
			    <td width="30%"><font class="clsText"><%=fnGetNameFromCode("BRANCH_MST", "BRANCH_CD", "BRANCH_NAME", strBRANCH_CD)%></font></td>
				<td width="20%"><font class="clsLabel">Division</font></td>
			    <td width="30%"><font class="clsText"><%=strDivName%></font></td>
			</tr>

			<tr>
				<td width="20%"><font class="clsLabel">Department</font></td>
				<td width="30%"><font class="clsText"><%=strDeptName%></font></td>
			</tr>			
		</table>
</div>


