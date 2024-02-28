<%
'**********************************************************
'Function name : fnGetDependentName
'Purpose       : This function returns the name of the dependents 
'Input         : The function should be called as 
'				 fnGetDependent("<Param1>")
'				 Param1 = employee no., 
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Sreehari
'Date          : 25-12-2001
'Reference	   : For implementation
'**********************************************************

Function fnGetDepName(strTableName, strCode1, strCode2, strDesc, strCodeVal)

	strSQL = "Select " & strDesc & " from " & strTableName & " where EMP_NO='"&Ucase(strCode1)&"' and upper("&strCode2&")='"&Ucase(strCodeVal)&"'"
	
	Set conCommon = Server.CreateObject ("ADODB.Connection")
	Set rsCommon = Server.CreateObject ("ADODB.Recordset")

	conCommon.Open astrConn
	rsCommon.Open strSQL, conCommon
		
	If rsCommon.EOF then
		fnGetDepName = ""
	Else
		fnGetDepName = rsCommon.Fields(0)
	End If
		
	rsCommon.Close 
	conCommon.Close 

	set rsCommon = nothing
	set conCommon = nothing

End Function


'**********************************************************
'Function name : fnGetDependent
'Purpose       : This function returns the name of the dependents of the employee and self name.
'Input         : The function should be called as 
'				 fnGetDependent("<Param1>")
'				 Param1 = employee no., 
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Amit Jain
'Date          : 25-12-2001
'Reference	   : For implementation
'**********************************************************
function fnGetDependent(intEmpNo)
	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	    set cn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
		cn.Open Connection
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		end if	 
		SQL1 = "Select PREFIX,EMP_FIRSTNAME , EMP_MIDNAME , EMP_LASTNAME from HRD_EMP_MST where EMP_NO = '" & intEmpNo & "'"
		rs.Open SQL1,cn
		
		strPREFIX = rs.Fields("PREFIX") 
		strEMP_FIRSTNAME = rs.Fields("EMP_FIRSTNAME") 
		strEMP_MIDNAME = rs.Fields("EMP_MIDNAME") 
		strEMP_LASTNAME = rs.Fields("EMP_LASTNAME") 
		strEMP_NAME = strPREFIX & " " & strEMP_FIRSTNAME & " " & strEMP_MIDNAME & " " & strEMP_LASTNAME 
		
		
		
        rs.Close 
        cn.Close 
    
    	cn.Open Connection
		SQL = "Select DEPENDENT_NAME,RELATION from HRD_EMP_DEPENDENT where EMP_NO = '" & intEmpNo & "' and DEPENDENT_TYPE = 'DEPENDENT'"
		rs.Open SQL,cn
			strReturn = "<OPTION selected value='" & strEMP_NAME &  "'>" & strEMP_NAME &  "</OPTION>"
			
			while not rs.EOF
		    		strReturn = strReturn & "<OPTION VALUE='"& rs("DEPENDENT_NAME")& "'>" &rs("DEPENDENT_NAME")& "</OPTION>"
					rs.MoveNext				
			wend
			rs.Close
			set rs = nothing
            cn.Close 
            set cn = nothing	
	fnGetDependent = strReturn
	
end function


'**********************************************************
'Function name : fnGetDependentName
'Purpose       : This function returns the name of the dependents of the employee .
'Input         : The function should be called as 
'				 fnGetDependent("<Param1>")
'				 Param1 = employee no., 
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Amit Jain
'Date          : 25-12-2001
'Reference	   : For implementation
'**********************************************************
function fnGetDependentName(intEmpNo)
	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	    set cn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
		cn.Open Connection
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		end if	 
		SQL1 = "Select EMP_FIRSTNAME from HRD_EMP_MST where EMP_NO = '" & intEmpNo & "'"
		rs.Open SQL1,cn
		strEMP_FIRSTNAME = rs.Fields("emp_firstname") 
        rs.Close 
        cn.Close 
    
    	cn.Open Connection
		SQL = "Select DEPENDENT_NAME, DEPENDENT_SR_NO from HRD_EMP_DEPENDENT where EMP_NO = '" & intEmpNo & "'"
		rs.Open SQL,cn
			
			while not rs.EOF
		    		strReturn = strReturn & "<OPTION VALUE="& rs("DEPENDENT_NAME")& ">" &rs("DEPENDENT_NAME")& "</OPTION>"
					rs.MoveNext				
			wend
			rs.Close
			set rs = nothing
            cn.Close 
            set cn = nothing	
	fnGetDependentName = strReturn
end function


'*************************************************
'This Function Returns Allowance Type
'Writeen By: SreeHari
'Date:27-12-2001
'*******************************************
'This function returns Allowanc Types
Function fnGetAllowType(pay_cd_cat)

	'Declarations
	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	
	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
	
	if Err.number <> 0 then
		 'Check if connection is opened properly
		 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if	 

	SQL1 = "Select pay_cd, pay_cd_desc from MIPRC_PAY_CODE where pay_cd_cat='" & pay_cd_cat &"'and LOAN_INT_ADV_FLAG ='O'and PAY_CD_TYPE IN ('P') "
    'SQL1 = "Select pay_cd, pay_cd_desc from MIPRC_PAY_CODE where pay_cd_cat='" & pay_cd_cat &"'and LOAN_INT_ADV_FLAG in ('O' , 'T')and PAY_CD_TYPE IN ('P') "
	
	rs.Open SQL1,cn
	
	if Err.number <> 0 or rs.EOF then
		 'Check if connection is opened properly
		 Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
	end if
	
	strReturn = ""
	
	While not rs.Eof
	
		strReturn = strReturn & "<option value='"&rs("pay_cd")&"'>"&rs("pay_cd_desc")&"</option>"
		
		rs.MoveNext
	Wend
	
	
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	
	fnGetAllowType = strReturn
End Function

'*************************************************
'This Function Returns Medical Claim Types
'Writeen By: SreeHari
'Date:30-12-2001
'*******************************************

'This function returns Allowanc Types
Function fnGetClaimTypes()

	'Declarations
	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	
	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
	
	if Err.number <> 0 then
		 'Check if connection is opened properly
		 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if	 

	SQL1 = "Select MEDICAL_CLAIM_TYPE, CLAIM_TYPE_DESCRIPTION from HRD_MEDICAL_MST where STATUS='A'order by PRIORITY Asc"
	
	rs.Open SQL1,cn
	
	if Err.number <> 0 or rs.EOF then
		 'Check if connection is opened properly
		 Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
	end if
	
	strReturn = ""
	
	While not rs.Eof
	
		strReturn = strReturn & "<option value='"&rs("MEDICAL_CLAIM_TYPE")&"'>"&rs("CLAIM_TYPE_DESCRIPTION")&"</option>"
		
		rs.MoveNext
	Wend
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	
	fnGetClaimTypes = strReturn
End Function


'**************************************************
'This Function Returns Allowance Description
'Writeen By:SreeHari
'Date:27-12-2001
'****************************************************


'This function returns Allowance Description 
Function fnGetAllowDesc(pay_cd)

	'Declarations
	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	
	
	
	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
	
	if Err.number <> 0 then
		 'Check if connection is opened properly
		 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if	 

	SQL1 = "Select pay_cd_desc from MIPRC_PAY_CODE where pay_cd='" & pay_cd &"'"
	rs.Open SQL1,cn
	
	if Err.number <> 0 or rs.EOF then
		 'Check if connection is opened properly
		 Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
	end if
	
	strReturn = rs("pay_cd_desc")

	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	
	fnGetAllowDesc = strReturn
End Function


'**********************************************************
'Function name : fnCheckHierarchy
'Purpose       : This function compares the heirarchy of two employees.
'Input         : The function should be called as 
'				 fnCheckHierarchy("<Param1>","<Param2>")
'				 Param1 = employee no. of authorizing person,Param2 = employee no. of applicant  
'Output        : The Function returns 1 if hierarchy for Param1 is greater than or equal to 
'				 Param2 and 0 otherwise.				 
'Author        : Riju Bhargava
'Date          : 27-12-2001
'Reference	   : For implementation
'**********************************************************
function fnCheckHierarchy(intEmpNo1,intEmpNo2)
	Dim cn, rs ,Connection
	Dim strAuthority,strApplicant,strHierAuth, strHierAppl,strReturn
	Dim strAuthRank,strApplRank,strAuthHier,strApplHier
	
	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )
	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
	if Err.number <> 0 then
		 'Check if connection is opened properly
		 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	     Response.End  
	end if
		
	'Check if the recommending authority is a competent one..higher heirarchy
    strAuthority = "select RANK_CD from HRD_EMP_MST where EMP_NO='" & intEmpNo1 &"'"
	
	''''Response.write strAuthority
	rs.Open strAuthority,cn
  	if Err.number <> 0 then
		 Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
		 Response.End  
	else
		 strAuthRank = trim(rs.Fields("RANK_CD"))
		 rs.Close
	end if
		
	strApplicant = "select RANK_CD from HRD_EMP_MST where EMP_NO='" & intEmpNo2 &"'"
	'''Response.write strApplicant
	rs.Open strApplicant,cn
  	if Err.number <> 0 then
		Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
		Response.End  
	else
	  	strApplRank = trim(rs.Fields("RANK_CD"))
		rs.Close
	end if
		
    strHierAuth = "select RANK_HIER from HRD_RANK_MST where RANK_ID='" & strAuthRank &"'"
	rs.Open strHierAuth,cn
  		if Err.number <> 0 then
			 Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		end if
		strAuthHier = cint(rs.Fields("RANK_HIER"))
		rs.Close
		
    strHierAppl = "select RANK_HIER from HRD_RANK_MST where RANK_ID='" & strApplRank &"'"
	rs.Open strHierAppl,cn
  		if Err.number <> 0 then
			 Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		end if
	if rs.Fields("RANK_HIER")="" then
		 Response.Redirect "../../Common/Error.asp?aintCode=1023&aintPage=-1&astrErrDescription=" & Err.description
	     Response.End  
	end if
	
	strApplHier = cint(rs.Fields("RANK_HIER"))
	rs.Close
	
	if 	(strAuthHier <  strApplHier) then
		strReturn = -1
	end if
	if 	(strAuthHier =  strApplHier) then
		strReturn = 0
	end if
	if 	(strAuthHier >  strApplHier) then
		strReturn = 1
	end if				 
					
	'end check authority
	
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	
	fnCheckHierarchy = strReturn

end function	


'**************************************************
'This Function Will check employee status and will give the error
'messages if the employee is suspended ,resigned,terminated or retiered

'written By : Amit Jain
'Date : 27/12/2001 

'example to call this function:fnGetEmployeeStatus intEmpNo,trim(rs.Fields("EMP_STATUS")),trim(rs.Fields("EMP_TYPE"))

'*************************************************

function fnGetEmployeeStatus(intEmpNo,empStatus,empType)

  	        if empType = "S" or empType= "I" then
 	          set rs = nothing
			   set cn = nothing
 	          Response.Redirect "../../Common/Error.asp?aintCode=1008&aintPage=-1&astrErrDescription=" & Err.description
 	       end if
  	        
  	        if  empStatus="SU" then
	            set rs = nothing
			    set cn = nothing
			    Response.Redirect "../../Common/Error.asp?aintCode=1004&aintPage=-1&astrErrDescription=" & Err.description
		    else
		    if  empStatus="RT" then
	            set rs = nothing
			    set cn = nothing
			    Response.Redirect "../../Common/Error.asp?aintCode=1002&aintPage=-1&astrErrDescription=" & Err.description
		    else
		    if  empStatus="TR" then
	            set rs = nothing
			    set cn = nothing
			    Response.Redirect "../../Common/Error.asp?aintCode=1003&aintPage=-1&astrErrDescription=" & Err.description
		    else
		    if  empStatus="RE" then
	            set rs = nothing
			    set cn = nothing
		    	Response.Redirect "../../Common/Error.asp?aintCode=1005&aintPage=-1&astrErrDescription=" & Err.description
		    end if
		    end if
		    end if
		    end if  
		     	
end function		 

'**********************************************************
'Function name : fnGetBlockyr
'Purpose       : This function returns the name of the dependents of the employee.
'Input         : The function should be called as 
'				 fnGetDependent("<Param1>")
'				 Param1 = employee no., 
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Amit Jain
'Date          : 25-12-2001
'Reference	   : For implementation
'**********************************************************
function fnGetBlockyr()

	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	    set cn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		
		Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
		cn.Open Connection
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		end if	 
		SQL = "Select BLOCK_YEAR from HRD_BLOCK_YEAR where ROWNUM<3 ORDER BY BLOCK_YEAR DESC "
		
		rs.Open SQL,cn
			while not rs.EOF
		    		strReturn = strReturn & "<OPTION VALUE="& rs("BLOCK_YEAR")& ">" &rs("BLOCK_YEAR")& "</OPTION>"
					rs.MoveNext				
			wend
			rs.Close
			set rs = nothing
            cn.Close 
            set cn = nothing	
	fnGetBlockyr = strReturn
end function



'**********************************************************
'Function name : fnExistEmployee
'Purpose       : 
'Input         : 
'				 
'				 
'Output        : 
'				 
'Author        : Amit Jain
'Date          : 18-01-2002
'Reference	   : 
'**********************************************************

'This function returns Allowance Description 
Function fnExistEmployee(intEmpNo)
     dim strReturn 
	 set cn = Server.CreateObject( "ADODB.Connection" )
	 set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
	
	if Err.number <> 0 then
		strReturn=-1  
	end if	 

	SQL1 = "Select EMP_NO , EMP_TYPE from HRD_EMP_MST where EMP_NO='" & intEmpNo &"'"
	
	rs.Open SQL1,cn
	
	if rs.EOF then
		strReturn=-2
	else 
	     strEMP_TYPE = trim(rs.Fields("EMP_TYPE"))
	if strEMP_TYPE = "S" or strEMP_TYPE = "I" then
	     strReturn=-3
    end if
	end if
		
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	fnExistEmployee = strReturn
End Function


'**********************************************************
'Function name : fnLtcCheck
'Purpose       : 
'Input         : 
'				 
'				 
'Output        : 
'				 
'Author        : Amit Jain
'Date          : 18-01-2002
'Reference	   : 
'**********************************************************

'This function returns Allowance Description 
Function fnLtcCheck(intEmpNo,strTravelMode,strTravelClass,strTrainType)

	dim strReturn 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
								
	if Err.number <> 0 then
	strReturn=-1  
	end if	 

	SQL1 = "Select RANK_CD from HRD_EMP_MST where EMP_NO='" & intEmpNo &"'"  
	
	rs.Open SQL1,cn        
	if rs.EOF then
        strReturn=-2
	else
	     strRANK = trim(rs.Fields("RANK_CD"))      	
	     rs.close	     
	     SQL2 = "Select LTC_ID from HRD_LTC_MST where RANK_ID='" & strRANK &"' and TRAVEL_MODE='" & strTravelMode &"' and TRAVEL_CLASS='" & strTravelClass &"'and TRAIN_TYPE='" & strTrainType &"'"  
	     
	     rs.Open SQL2	           	 
	     if rs.EOF then
	      strReturn=-3
	     end if
	     
	end if
		
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	
	fnLtcCheck = strReturn
End Function
'**************************************************
'This Function returns the dependent name based on employee no and dependent sr. no
'Writeen By:SreeHari
'Date:27-12-2001
'****************************************************
Function fnGetDependentName(intEmpNo, intDepNo)

	dim strReturn 

	'set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	'Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	'cn.Open Connection
								
	if Err.number <> 0 then
	strReturn=-1  
	end if	 

	SQL1 = "Select DEPENDENT_NAME from HRD_EMP_DEPENDENT where EMP_NO='" & intEmpNo & "' and DEPENDENT_SR_NO='" & intDepNo &"'"  

	rs.Open SQL1,cn        
	if rs.EOF then
        strReturn=-2
	else
		strReturn=rs("DEPENDENT_NAME")
	end if
		
	rs.Close
	'cn.Close
	Set rs = Nothing
	'Set cn = Nothing
	
	fnGetDependentName = strReturn
End Function

'**********************************************************
'Function name : fnCheckFrnEntitlement
'Purpose       : checks whether an employee is entitled for foreign tour 
'Input         : employee number
'Output        : 
'Author        : Riju
'Date          : 22-01-2002
'**********************************************************
Function fnCheckFrnEntitlement(intEmpNo)

	dim strReturn
	dim SQL1, SQL2 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rsMain = Server.CreateObject( "ADODB.Recordset" )
	set rsHier = Server.CreateObject( "ADODB.Recordset" )
	set rsHierLovMst = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
								
	if Err.number <> 0 then
		cn.Close
		Set cn = Nothing
		strReturn=-1  
		exit function
	end if	 
	
	SQL = "select  rank_hier from hrd_rank_mst where rank_id = 'GM'"
	rsHier.Open SQL, cn
	If Err.number <> 0 Then
	'Close connection
		rsHier.Close 
		set rsHier = nothing
		cn.Close
		Set cn = Nothing
		strReturn=-1
		exit function
	End If
	strRankHier = rsHier.Fields("rank_hier")
	rsHier.Close 
	set rsHier = nothing
	
	SQL1 = "Select rank_id,RANK_hier from hrd_rank_mst where rank_id = (select rank_id from HRD_EMP_MST where EMP_NO='" & intEmpNo &"')"  
	
	rsMain.Open SQL1,cn        
	If Err.number <> 0 Then
	'Close connection
		rsMain.Close 
		set rsMain = nothing
		cn.Close
		Set cn = Nothing
		strReturn=-1
		exit function
	End If

	if rsMain.EOF then
        rsMain.Close
        set rsMain = nothing
        cn.close
        set cn = nothing
        strReturn=-2
        exit function
        
	end if
	     strRANKHierMain = trim(rsMain.Fields("RANK_HIER"))      	
	     
	     rsMain.close	     
	     set rsMain = nothing
	    if cdbl(strRANKHierMain) >=  cdbl(strRankHier) then
			
				SQL2 = "Select LOV_VALUE from LOV_MST where rtrim(ltrim(LOV_DESC))='GM and above'"
		else
	     		SQL2 = "Select LOV_VALUE from LOV_MST where rtrim(ltrim(LOV_DESC))='All Other Officers'"
		end if
		     rsHierLovMst.Open SQL2,cn	           	 
			
			If Err.number <> 0 Then
				'Close connection
				rsHierLovMst.Close 
				set rsHierLovMst = nothing
				cn.Close
				Set cn = Nothing
				strReturn=-1
				exit function
			End If
			if rsHierLovMst.EOF then
				rsHierLovMst.Close 
				set rsHierLovMst = nothing
				cn.close
				set cn = nothing
				strReturn=-2
				exit function
			else
				strReturn= rsHierLovMst.Fields("LOV_VALUE")
				rsHierLovMst.Close 
				set rsHierLovMst = nothing
				cn.close
				set cn = nothing
			end if
	
		

	fnCheckFrnEntitlement = strReturn
End Function

'**********************************************************
'Function name : fnFrnTaDaExist
'Purpose       : Checks if the record for the same date exists and if the outfit claimed in last 3 years
'Input         : 		 
'Output        : 
'Author        : Riju Bhargava
'Date          : 22-01-2002
'Reference	   : 
'**********************************************************

Function fnFrnTaDaExist(intEmpNo,strDepDate ,strArrDate,strOutKit)

	dim strReturn
	dim SQL1, SQL2 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection						
	if Err.number <> 0 then
		strReturn=-1  
	end if	 

	'check that no other record for tour is already applied by the employee for any day in the given period...Riju

	 strComp1 = "to_date('" & strDepDate & "','dd/mm/yyyy') between DEP_DT and ARR_DT"
	 strComp2 = "to_date('" & strArrDate & "','dd/mm/yyyy')-1 between DEP_DT and ARR_DT"
	 strComp3 = "to_date('" & strDepDate & "','dd/mm/yyyy')< DEP_DT"
	 strComp4 = "to_date('" & strArrDate & "','dd/mm/yyyy')-1 >= DEP_DT"
	 strComp5 = "to_date('" & strDepDate & "','dd/mm/yyyy') < ARR_DT"
	 strComp6 = "to_date('" & strArrDate & "','dd/mm/yyyy')-1 > ARR_DT"

     SQL1 = "select * from HRD_TADA_FOR_TXN where EMP_NO ='"& intEmpNo
     SQL1 = SQL1 & "' and (" & strComp1
	 SQL1 = SQL1 & " or "& strComp2
	 SQL1 = SQL1 & " or ("& strComp3 & " and " & strComp4 & ")"
	 SQL1 = SQL1 & " or ("& strComp5 & " and " & strComp6 & "))"

	 rs.Open SQL1,cn
  	 if Err.number <> 0 then
		strReturn = -1
	 end if
  	 if not rs.EOF then
		 rs.Close
		 strReturn = -2
	 else
		 rs.Close
		 cn.Close
		 Set rs = Nothing
		 Set cn = Nothing
		 
		 if strOutKit="Y" then
			strReturn = strCheckOutfit(intEmpNo,strOutKit)
		 else
			strReturn = 1
		 end if				
	 end if
	 
	'finish check
	fnFrnTaDaExist = strReturn
End Function

'**********************************************************
'Function name : strCheckOutfit
'Purpose       : Checks if the outfit claimed in last 3 years
'Input         : 		 
'Output        : 
'Author        : Riju Bhargava
'Date          : 22-01-2002
'Reference	   : 
'**********************************************************
Function strCheckOutfit(intEmpNo,strOutKit)

	dim strReturn
	dim SQL1, SQL2 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection						
	if Err.number <> 0 then
		strReturn=-1  
	end if	 

	'check that no other record for tour is already applied by the employee for any day in the given period...Riju
	 SQL1 = "select to_char(sysdate,'yyyy') CurrYear from dual"
	 
	 rs.Open SQL1,cn
  	 if Err.number <> 0 then
  		 rs.close
		 strReturn = -1
	 else
		 intCurrYear = cint(rs.fields("CurrYear"))
		 rs.close
	 end if
	 
	 'previous years	 
	 intPrevYr1 = intCurrYear -1
	 intPrevYr2 = intCurrYear -2
	 
     SQL2 = "select OUTFIT_KIT_FLG from HRD_TADA_FOR_TXN where EMP_NO ='"& intEmpNo
     SQL2 = SQL2 & "' and to_char(APPR_OR_REJ_DT,'yyyy') in ('" & intCurrYear & "','" & intPrevYr1 & "','" & intPrevYr2 & "')" 
	 SQL2 = SQL2 & " and APPROVAL_STATUS='A' and OUTFIT_KIT_FLG = 'Y'"
	 
	 rs.Open SQL2,cn
  	 if Err.number <> 0 then
		 strReturn = -1
	     Response.End  
	 end if
  	 if not rs.EOF then
		 rs.Close
		 cn.Close
		 Set rs = Nothing
		 Set cn = Nothing
		 strReturn = -3
	 else
		 rs.Close
		 cn.Close
		 Set rs = Nothing
		 Set cn = Nothing		 
		 strReturn = 1
	 end if

strCheckOutfit = strReturn
	'finish check
End Function

'**********************************************************
'Function name : fnGetForeignDA
'Purpose       : Function to get the DA specified for foreign countries by Govt.
'Input         : 		 
'Output        : 
'Author        : Riju Bhargava
'Date          : 24-01-2002
'Reference	   : 
'**********************************************************
Function fnGetForeignDA(strCtry)
	dim strReturn
	dim SQL1, SQL2 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection						
	if Err.number <> 0 then
		strReturn=-1  
	end if	 

     SQL1 = "select DA_BY_GOV from HRD_TADA_FOR_MST where COUNTRY_CD ='"& strCtry & "'"
     
	 rs.Open SQL1,cn
  	 if Err.number <> 0 then
		strReturn = "-1"
	 end if
  	 if rs.EOF then
		 rs.Close
		 strReturn = "-2"
	 else
		 strReturn = rs.Fields("DA_BY_GOV")
		 rs.Close
		 cn.Close
		 Set rs = Nothing
		 Set cn = Nothing
	 end if
	 
	'finish check
	fnGetForeignDA = strReturn
End Function

'**********************************************************
'Function name : fnIndTaDaExist
'Purpose       : Function to check if the record for employee exist within the 
'				 given dates for tour.transfer in India.
'Input         : 		 
'Output        : 
'Author        : Riju Bhargava
'Date          : 24-01-2002
'Reference	   : 
'**********************************************************
Function fnIndTaDaExist(intEmpNo,strDepDate ,strArrDate)

	dim strReturn
	dim SQL1, SQL2 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection						
	if Err.number <> 0 then
		strReturn=-1  
	end if	 

	'check that no other record for tour is already applied by the employee for any day in the given period...Riju

	 strComp1 = "to_date('" & strDepDate & "','dd/mm/yyyy') between DEP_DT and ARR_DT"
	 strComp2 = "to_date('" & strArrDate & "','dd/mm/yyyy')-1 between DEP_DT and ARR_DT"
	 strComp3 = "to_date('" & strDepDate & "','dd/mm/yyyy')< DEP_DT"
	 strComp4 = "to_date('" & strArrDate & "','dd/mm/yyyy')-1 >= DEP_DT"
	 strComp5 = "to_date('" & strDepDate & "','dd/mm/yyyy') < ARR_DT"
	 strComp6 = "to_date('" & strArrDate & "','dd/mm/yyyy')-1 > ARR_DT"

     SQL1 = "select * from HRD_TADA_IND_MST where EMP_NO ='"& intEmpNo
     SQL1 = SQL1 & "' and (" & strComp1
	 SQL1 = SQL1 & " or "& strComp2
	 SQL1 = SQL1 & " or ("& strComp3 & " and " & strComp4 & ")"
	 SQL1 = SQL1 & " or ("& strComp5 & " and " & strComp6 & "))"

	 rs.Open SQL1,cn
  	 if Err.number <> 0 then
		strReturn = -1
	 end if
  	 if not rs.EOF then
		 rs.Close
		 strReturn = -2
	 else
		 strReturn = -3
		 rs.Close
		 cn.Close
		 Set rs = Nothing
		 Set cn = Nothing
	 end if
	 
	'finish check
	fnIndTaDaExist = strReturn
End Function

'**********************************************************
'Function name : fnDepIndTaDaExist
'Purpose       : Function to check if the record for dependent exist for
'				 given date for transfer in India.
'Input         : 		 
'Output        : 
'Author        : Riju Bhargava
'Date          : 25-01-2002
'Reference	   : 
'**********************************************************
Function fnDepIndTaDaExist(intEmpNo , intDepSrNo, strDepDate)

	dim strReturn
	dim SQL1, SQL2 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection						
	if Err.number <> 0 then
		strReturn=-1  
	end if	 

	'check that no other same dependent record is already applied by the employee for any day in the given period...Riju

     SQL1 = "select * from HRD_TADA_IND_MST where EMP_NO ='"& intEmpNo
     SQL1 = SQL1 & "' and DEPENDENT_SR_NO = '"& intDepSrNo
     SQL1 = SQL1 & "' and DEP_DT = to_date('" & strDepDate & "','dd/mm/yyyy')"

	 rs.Open SQL1,cn
  	 if Err.number <> 0 then
		strReturn = -1
	 end if
  	 if not rs.EOF then
		 rs.Close
		 strReturn = -2
	 else
		 strReturn = -3
		 rs.Close
		 cn.Close
		 Set rs = Nothing
		 Set cn = Nothing
	 end if
	 
	'finish check
	fnDepIndTaDaExist = strReturn

End Function
'**********************************************************
'Function name : fnGetIndCityRates
'Purpose       : Function to get the DA rates for Indian cities
'				 as per the rank applicable
'Input         : 		 
'Output        : 
'Author        : Riju Bhargava
'Date          : 24-01-2002
'Reference	   : 
'**********************************************************
Function fnGetIndCityRates(intEmpNo,strCityType)
	dim strReturn
	dim SQL1, SQL2 

	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
								
	if Err.number <> 0 then
		strReturn=-1  
	end if	 

	SQL1 = "Select RANK_CD from HRD_EMP_MST where EMP_NO='" & intEmpNo &"'"  
	
	rs.Open SQL1,cn        
	if rs.EOF then
        strReturn=-2
        rs.Close
	else
	     strRANK = trim(rs.Fields("RANK_CD"))      	
	     rs.close	     
	     SQL2 = "Select DA_AMT from HRD_TADA_CITY_MST where RANK_ID='" & strRANK & "' and CITY_CLASS='" & strCityType & "'"

		 rs.Open SQL2	           	 
		 if rs.EOF then
			strReturn=-3
		 else
			strReturn= rs.Fields("DA_AMT")
		 end if
	end if
		
	cn.Close
	Set rs = Nothing
	Set cn = Nothing

	fnGetIndCityRates = strReturn

End Function


'**********************************************************
'Function name : fnGetTourTransferDesc
'Purpose       : Function to get the tour/transfer description
'Input         : 		 
'Output        : 
'Author        : Riju Bhargava
'Date          : 25-01-2002
'Reference	   : 
'**********************************************************
Function fnGetTourTransferDesc(strPurpose,strTrfnType)

Dim strPurposeDesc

if strTrfnType="R" then
	strTrfndesc = "Requested"
elseif strTrfnType="P" then
	strTrfndesc = "Permanent"
elseif strTrfnType="T" then
	strTrfndesc = "Temporary"
end if

if strPurpose = "TO" then
	strPurposeDesc = "Tour"
else
	strPurposeDesc = strTrfndesc & " Transfer"
end if

fnGetTourTransferDesc = strPurposedesc

end function

'**********************************************************
'Function name : fnTravelEligibilityCheck
'Purpose       : check travel eligibility based on rank, travel details and the transfer type.
'Input         : 
'Output        : 
'Author        : Riju Bhargava
'Date          : 25-01-2002
'Reference	   : 
'**********************************************************

'This function returns Allowance Description 
Function fnTravelEligibilityCheck(intEmpNo,strTrfnType,strTravelMode,strTravelClass,strTrainType)

	dim strReturn 
	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
								
	if Err.number <> 0 then
		strReturn=-1  
	end if	 
	
	if strTrfnType="R" and strTravelMode="RAIL" then
		strTravelClass = "SLEEPER"
		strTrainType = "ORDINARY"
	end if
	
	SQL1 = "Select RANK_CD from HRD_EMP_MST where EMP_NO='" & intEmpNo &"'"  
	
	rs.Open SQL1,cn        
	if rs.EOF then
        strReturn=-2
	else
	     strRANK = trim(rs.Fields("RANK_CD"))      	
	     rs.close	     
	     SQL2 = "Select ELIGIBILITY_ID from HRD_TADA_TRAVEL_MODE_MST where RANK_ID='" & strRANK
	     SQL2 = SQL2 & "' and TRAVEL_MODE='" & strTravelMode &"' and TRAVEL_CLASS='" & strTravelClass
	     SQL2 = SQL2 &"'and TRAIN_TYPE='" & strTrainType &"'"  
	     
	     rs.Open SQL2	           	 
	     if rs.EOF then
	      strReturn=-3
	     end if
	     
	end if
		
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	
	fnTravelEligibilityCheck = strReturn
End Function

'**********************************************************
'Function name : fnCheckPack
'Purpose       : gets the packing charges allowed for the given rank.
'Input         : employee number
'Output        : packing charge value, -1, 0
'Author        : Riju Bhargava
'Date          : 30-01-2002
'Reference	   : 
'**********************************************************

'This function returns Allowance Description 
Function fnCheckPack(intEmpNo)

	dim strReturn 
	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
								
	if Err.number <> 0 then
		strReturn=-1  
	end if	 
	
	SQL1 = "Select RANK_CD from HRD_EMP_MST where EMP_NO='" & intEmpNo &"'"  
	
	rs.Open SQL1,cn        
	if rs.EOF then
        strReturn=-2
	else
	     strRANK = trim(rs.Fields("RANK_CD"))      	
	     rs.close	     
	     SQL2 = "Select PACK_CHARGE from HRD_TADA_TRANSPACK_MST where RANK_ID='" & strRANK & "'"
	     
	     rs.Open SQL2	           	 
	     if rs.EOF then
	      strReturn=-3
	     else
	      strReturn=rs.fields("PACK_CHARGE")
	     end if 
	end if
		
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
	
	fnCheckPack = strReturn
End Function

'**********************************************************
'Function name : fnGetKmCharges
'Purpose       : gets the road charges allowed for the given rank.
'Input         : employee number
'Output        : road charge per km value, -1, 0 in case of no
'				 specific charges. there the actual charges can be reumbursed.
'Author        : Riju Bhargava
'Date          : 30-01-2002
'Reference	   : 
'**********************************************************

'This function returns Allowance Description 
Function fnGetKmCharges(intEmpNo)

	dim strReturn 
	set cn = Server.CreateObject( "ADODB.Connection" )
	set rs = Server.CreateObject( "ADODB.Recordset" )

	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	cn.Open Connection
								
	if Err.number <> 0 then
		strReturn=-1  
	end if	 
	
	SQL1 = "Select RANK_CD from HRD_EMP_MST where EMP_NO='" & intEmpNo &"'"  
	
	rs.Open SQL1,cn        
	if rs.EOF then
        strReturn=-2
	else
	     strRANK = trim(rs.Fields("RANK_CD"))      	
	     rs.close	     

	     SQL2 = "Select ROAD_KM_EXPENSE from HRD_TADA_TRAVEL_MODE_MST where RANK_ID='" & strRANK & "' and TRAVEL_MODE='ROAD'"
	     
	     rs.Open SQL2
	     if rs.EOF then
	        strReturn=-3
	     else
		    strRoadExp = rs.fields("PACK_CHARGE")
			if strRoadExp = "" then
				strReturn = 0
			else
				strReturn = strRoadExp
			end if	
	     end if 
	end if
		
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing

	fnGetKmCharges = strReturn
End Function


'**********************************************************
'Function name : fnGetSlab
'Purpose       : This function returns the Basic pay values
'Input         : The function should be called as 
'				 fnGetDependent("<Param1>")
'				 Param1 = rank code, 
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Riju
'Date          : 14-03-2002
'Reference	   : 
'**********************************************************
function fnGetSlab(strRank,numBasicSalary,Level)

	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	    set cn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		
		Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
		cn.Open astrConn
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 if Level=5 then 
			 Response.Redirect "../../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		     elseif Level = 4 then
		     Response.Redirect "../../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		     end if  
		end if
	
		strSQL = "select SLAB_START_AMT, SLAB_END_AMT, INCREMENT_AMT from HRD_PAY_SCALE_MST where "
		strSQL = strSQL & "upper(RANK_CD)=UPPER('"&strRank&"') and PROCESS_FLAG='Y'"
	
		rs.Open strSQL,cn
		if Err.number <> 0 then
		     if Level=5 then 
			 Response.Redirect "../../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     elseif Level = 4 then
		     Response.Redirect "../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End
		     end if  
		end if	 
		
		if rs.EOF then
		     if Level=5 then 
			Response.Redirect "../../../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=Pay scale details for this rank do not exist"		
		     elseif Level = 4 then
		     Response.Redirect "../../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=Pay scale details for this rank do not exist"		
		     Response.End
		     end if
		end if
		         
		
		x=0
		
			if numBasicSalary = "" then
				strReturn="<OPTION value=''>---Select---</OPTION>"
			end if
		
		
		while not rs.EOF

			strSLAB_START_AMT = cdbl(rs.fields("SLAB_START_AMT"))
			strSLAB_END_AMT = cdbl(rs.fields("SLAB_END_AMT"))
			strINCREMENT_AMT = cdbl(rs.fields("INCREMENT_AMT"))

		
			intNumSlabs = ((strSLAB_END_AMT-strSLAB_START_AMT)/(strINCREMENT_AMT))
		
			intSlabCount = intNumSlabs
			intSlabStAmt = strSLAB_START_AMT
			intSlabIncr = strINCREMENT_AMT
		
			for j=0 to intSlabCount step 1
				intBasic = intSlabStAmt + (intSlabIncr * j)
			    
			    if trim(numBasicSalary) = "" then
					strReturn = strReturn & "<OPTION VALUE="& intBasic& ">" & intBasic & "</OPTION>"
				else
					if cdbl(numBasicSalary)=intBasic then
						strReturn = strReturn & "<OPTION VALUE="& intBasic & " selected>" &intBasic& "</OPTION>"
					else
						strReturn = strReturn & "<OPTION VALUE="& intBasic& ">" &intBasic& "</OPTION>"
					end if
				end if				
			next
			
		
		rs.movenext	
		wend
		
		rs.Close
		set rs = nothing
        cn.Close 
        set cn = nothing	
	fnGetSlab = strReturn
end function



'**********************************************************
'Function name : fnGetNextSlab
'Purpose       : This function returns the Basic pay values
'Input         : The function should be called as 
'				 fnGetDependent("<Param1>")
'				 Param1 = rank code, 
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Riju
'Date          : 14-03-2002
'Reference	   : 
'**********************************************************
function fnGetNextSlab(strRank,numBasicSalary,Level)

	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn
	    set cn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		
		Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
		cn.Open astrConn
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 if Level=5 then 
			 Response.Redirect "../../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		     elseif Level = 4 then
		     Response.Redirect "../../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		     end if  
		end if
			 

                    

'select RANK_ID from HRD_RANK_MST where RANK_HIER=(SELECT MIN(RANK_HIER) FROM HRD_RANK_MST  where rank_hier > (select rank_hier from hrd_rank_mst where rank_id = UPPER('"&strRank&"')) )
    	strSQL = "select SLAB_START_AMT, SLAB_END_AMT, INCREMENT_AMT from HRD_PAY_SCALE_MST where "
		strSQL = strSQL & "RANK_CD= (select UPPER(RANK_ID) from HRD_RANK_MST where RANK_HIER=(SELECT MIN(RANK_HIER) FROM HRD_RANK_MST  where rank_hier > (select rank_hier from hrd_rank_mst where rank_id = UPPER('"&strRank&"')) )) and PROCESS_FLAG='Y'"

		rs.Open strSQL,cn
		if Err.number <> 0 then
		     if Level=5 then 
			 Response.Redirect "../../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     elseif Level = 4 then
		     Response.Redirect "../../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End
		     end if  
		end if	 
		
		if rs.EOF then
		     if Level=5 then 
			Response.Redirect "../../../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=Pay scale details for this rank do not exist"		
		     elseif Level = 4 then
		     Response.Redirect "../../../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=Pay scale details for this rank do not exist"		
		     Response.End
		     end if
		end if
		         
		
		x=0
		
			if numBasicSalary = "" then
				strReturn="<OPTION value=''>---Select---</OPTION>"
			end if
		
		
		while not rs.EOF

			strSLAB_START_AMT = cdbl(rs.fields("SLAB_START_AMT"))
			strSLAB_END_AMT = cdbl(rs.fields("SLAB_END_AMT"))
			strINCREMENT_AMT = cdbl(rs.fields("INCREMENT_AMT"))

		
			intNumSlabs = ((strSLAB_END_AMT-strSLAB_START_AMT)/(strINCREMENT_AMT))
		
			intSlabCount = intNumSlabs
			intSlabStAmt = strSLAB_START_AMT
			intSlabIncr = strINCREMENT_AMT
		
			for j=0 to intSlabCount step 1
				intBasic = intSlabStAmt + (intSlabIncr * j)
			    
			    if trim(numBasicSalary) = "" then
					strReturn = strReturn & "<OPTION VALUE="& intBasic& ">" & intBasic & "</OPTION>"
				else
					if cdbl(numBasicSalary)=intBasic then
						strReturn = strReturn & "<OPTION VALUE="& intBasic & " selected>" &intBasic& "</OPTION>"
					else
						strReturn = strReturn & "<OPTION VALUE="& intBasic& ">" &intBasic& "</OPTION>"
					end if
				end if				
			next
			
		
		rs.movenext	
		wend
		
		rs.Close
		set rs = nothing
        cn.Close 
        set cn = nothing	
	fnGetNextSlab = strReturn
end function



'**********************************************************
'Function name : fnCheckRegionBranchCode
'Purpose       : This function validates the branch for a region 
'Input         :
'Output        :
'Author        : Riju
'Date          : 18-3-2002
'Reference	   : For implementation
'**********************************************************
function fnCheckRegionBranchCode(strRegionCd, strBranchCd)
	set objConnCheckCode = Server.CreateObject( "ADODB.Connection" )
	set objRsCheckCode = Server.CreateObject( "ADODB.Recordset" )
	objConnCheckCode.Open astrConn	
	if Err.number <> 0 then
		 Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if
		
	CheckCodeSQL = "select * from BRANCH_MST where UPPER(BRANCH_CD) = UPPER('" & strBranchCd & "') and UPPER(REGION_CD) = UPPER('" & strRegionCd & "')"

   	objRsCheckCode.Open CheckCodeSQL,objConnCheckCode
	if Err.number <> 0 then
		 Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if

	if objRsCheckCode.EOF then
		objRsCheckCode.Close
		objConnCheckCode.close
		set objRsCheckCode = nothing
		set objConnCheckCode = nothing
		fnCheckRegionBranchCode = "0"
	else
		objRsCheckCode.Close
		objConnCheckCode.close
		set objRsCheckCode = nothing
		set objConnCheckCode = nothing
		fnCheckRegionBranchCode = "1"
	end if
end function


'**********************************************************
'Function name : fnCheckBranchSectorCode
'Purpose       : This function validates the sector for a branch 
'Input         :
'Output        :
'Author        : Riju
'Date          : 18-3-2002
'Reference	   : For implementation
'**********************************************************
function fnCheckBranchSectorCode(strBranchCd,strSectorCd)
	set objConnCheckCode = Server.CreateObject( "ADODB.Connection" )
	set objRsCheckCode = Server.CreateObject( "ADODB.Recordset" )
	objConnCheckCode.Open astrConn	
	if Err.number <> 0 then
		 Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if
	
	objRsCheckCode.cursorlocation=3	
	
	CheckCodeSQL = "select * from SECTOR_MST where UPPER(SECTOR_CD) = UPPER('" & strSectorCd & "') and UPPER(BRANCH_CD) = UPPER('" & strBranchCd & "')"

   	objRsCheckCode.Open CheckCodeSQL,objConnCheckCode
	if Err.number <> 0 then
		 Response.Redirect "../../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
	end if

	if objRsCheckCode.EOF then
		objRsCheckCode.Close
		objConnCheckCode.close
		set objRsCheckCode = nothing
		set objConnCheckCode = nothing
		fnCheckBranchSectorCode = "0"
	else
		objRsCheckCode.Close
		objConnCheckCode.close
		set objRsCheckCode = nothing
		set objConnCheckCode = nothing
		fnCheckBranchSectorCode = "1"
	end if
end function

'**********************************************************
'Function name : fnGetDependent
'Purpose       : This function returns the name of the dependents of the employee and self name.
'Input         : The function should be called as 
'				 fnGetDependent("<Param1>")
'				 Param1 = employee no., 
'Output        : The Function returns the values to be populated in the Dropdown
'				 in between the <Option> tags.
'Author        : Amit Jain
'Date          : 25-12-2001
'Reference	   : For implementation
'**********************************************************
function fnGetDependentAndRelation(intEmpNo)
	Dim cn, rs ,Connection
	Dim SQL,SQL1, strReturn, strName, strRelation
	    set cn = Server.CreateObject( "ADODB.Connection" )
		set rs = Server.CreateObject( "ADODB.Recordset" )
		Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
		cn.Open Connection
		if Err.number <> 0 then
			 'Check if connection is opened properly
			 Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
		end if	 
		SQL1 = "Select EMP_FIRSTNAME , EMP_MIDNAME , EMP_LASTNAME from HRD_EMP_MST where EMP_NO = '" & intEmpNo & "'"
		rs.Open SQL1,cn
		
		'strEMP_FIRSTNAME = rs.Fields("PREFIX") 
		strEMP_FIRSTNAME = rs.Fields("EMP_FIRSTNAME") 
		strEMP_MIDNAME = rs.Fields("EMP_MIDNAME") 
		strEMP_LASTNAME = rs.Fields("EMP_LASTNAME") 
		strEMP_NAME = strEMP_FIRSTNAME & " " & strEMP_MIDNAME & " " & strEMP_LASTNAME 
		
		strRelation = "<OPTION selected value='Self'>Self</OPTION>"
		
        rs.Close 
        cn.Close 
    
    	cn.Open Connection
		SQL = "Select DEPENDENT_NAME,RELATION from HRD_EMP_DEPENDENT where EMP_NO = '" & intEmpNo & "' and DEPENDENT_TYPE = 'DEPENDENT'"
		rs.Open SQL,cn
			strName = "<OPTION selected value=" & strEMP_NAME &  ">" & strEMP_NAME &  "</OPTION>"
			
			while not rs.EOF
		    		strName = strName & "<OPTION VALUE='"& rs("DEPENDENT_NAME")& "'>" &rs("DEPENDENT_NAME")& "</OPTION>"
		    		strRelation = strRelation & "<OPTION VALUE="& rs("RELATION")& ">" &rs("RELATION")& "</OPTION>"
					rs.MoveNext				
			wend
			rs.Close
			set rs = nothing
            cn.Close 
            set cn = nothing
            
    strReturn = strName & "|" & strRelation
    	
	fnGetDependentAndRelation = strReturn
	
end function


'**************************************************
'This Function returns the pan Reason description based on action code and reason code
'Writeen By:amit jain
'Date:06-04-2001
'****************************************************
Function fnGetReasonDescription(strActionCode,strReasonCode,level)

    ' Response.Write strActionCode
    ' Response.Write strReasonCode
    ' Response.End 

	'dim strReturn 
	set con = Server.CreateObject( "ADODB.Connection" )
	set reco = Server.CreateObject( "ADODB.Recordset" )
	Connection = "Provider=MSDAORA.1;Data Source=ecgcprod;User ID=ecgc;Password=ecgc"	 
	con.Open Connection
	
    if Err.number <> 0 then
       if level=2 then
	         Response.Redirect "../../Common/Error.asp?aintCode=4&aintPage=-1&astrErrDescription=" & Err.description
		     Response.End  
	   end if
	end if	 
	SQL1 = "Select REASON_DESCRIPTION  from HRD_PAN_REASON_MST where ACTION_CD='" & strActionCode & "' and REASON_CD='" & strReasonCode &"'"  
	'Response.Write SQL1
	'Response.End 
	reco.Open SQL1,con        
	
	if reco.EOF then
	   if level = 2 then
	       Response.Redirect "../../Common/Error.asp?aintCode=2&aintPage=-1&astrErrDescription=Reason Description in reason master does not exist"		
		   Response.End  
	   end if
	   
	else
		if trim(reco("REASON_DESCRIPTION")) <> "" then
		strReturn=reco("REASON_DESCRIPTION")
		else
		strReturn=strReasonCode
		end if
	
	end if
		
	reco.Close
	con.Close
	Set reco = Nothing
	Set con = Nothing
	fnGetReasonDescription = strReturn
End Function

%>

